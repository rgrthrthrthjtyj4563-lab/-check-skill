from __future__ import annotations

import json
import os
import re
import subprocess
import sys
import tempfile
from pathlib import Path

import httpx
from fastapi import FastAPI, File, Form, HTTPException, UploadFile
from fastapi.responses import FileResponse, JSONResponse

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "scripts"))

from generate_answer_table import read_questions  # noqa: E402

DEEPSEEK_API_KEY = os.getenv("DEEPSEEK_API_KEY", "")
DEEPSEEK_BASE_URL = os.getenv("DEEPSEEK_BASE_URL", "https://api.deepseek.com").rstrip("/")
DEEPSEEK_MODEL = os.getenv("DEEPSEEK_MODEL", "deepseek-v4-pro")

app = FastAPI(title="Survey Self-Check Skill API")


def read_report_text(path: Path, max_chars: int = 30000) -> str:
    raw = path.read_bytes()
    for encoding in ("utf-8", "utf-8-sig", "gb18030"):
        try:
            return raw.decode(encoding)[:max_chars]
        except UnicodeDecodeError:
            continue
    return raw.decode("utf-8", errors="ignore")[:max_chars]


def strip_json_fence(text: str) -> str:
    text = text.strip()
    if text.startswith("```"):
        text = re.sub(r"^```(?:json)?\s*", "", text)
        text = re.sub(r"\s*```$", "", text)
    match = re.search(r"\{.*\}", text, flags=re.S)
    return match.group(0) if match else text


def question_payload(questionnaire_path: Path) -> list[dict]:
    _, questions, _ = read_questions(questionnaire_path)
    payload = []
    for question in questions:
        payload.append(
            {
                "id": f"Q{question.number}",
                "text": question.text,
                "options": list(question.options.keys()),
            }
        )
    return payload


async def generate_constraints(report_text: str, questions: list[dict]) -> dict:
    if not DEEPSEEK_API_KEY:
        raise HTTPException(status_code=500, detail="DEEPSEEK_API_KEY is not configured")

    schema = {
        "question_roles": {"Q1": "prerequisite"},
        "option_tags": {"Q1": {"A.example": ["positive"]}},
        "rules": [
            {
                "id": "R001",
                "type": "forbid_combination",
                "type_label": "体验与态度极性冲突",
                "if": {"question": "Q1", "option_tags_any": ["positive"]},
                "then_forbid": {"question": "Q2", "option_tags_any": ["negative"]},
                "repair": {
                    "target": "then_question",
                    "prefer_tags": ["positive"],
                    "fallback_tags": ["neutral"],
                    "description": "保留前置答案，修正后续答案。",
                },
                "severity": "high",
                "reason": "禁止逻辑冲突组合。",
            }
        ],
    }
    prompt = f"""
你是医药调研问卷逻辑约束专家。请基于报告和问卷题目/选项，输出可执行 constraints.json。

硬性要求：
1. 只输出 JSON，不要 Markdown，不要解释。
2. question_roles 覆盖所有题目，键使用 Q1、Q2 这种题号。
3. option_tags 尽量覆盖所有选项，键使用题号。
4. rules 只写强逻辑冲突规则，type 固定为 forbid_combination。
5. 角色限定：prerequisite, behavior, experience, attitude, price, channel, info, insurance, other。
6. 标签限定：positive, strong_positive, neutral, negative, no_experience, low_frequency, high_frequency, insurance_used, insurance_not_used, price_sensitive, price_not_sensitive。
7. 必须覆盖：未使用/未听说/不确定前提与后续评价冲突、正向体验与负向推荐冲突、负向体验与强正向推荐冲突、医保互斥、低频/不用药与强行为冲突、价格态度冲突。
8. 对“不确定、不知道、不清楚、无法评价、无法判断、其他”等模糊选项必须增加 fuzzy 标签。

输出 JSON schema 示例：
{json.dumps(schema, ensure_ascii=False)}

报告内容：
{report_text}

问卷题目和选项：
{json.dumps(questions, ensure_ascii=False)}
"""

    async with httpx.AsyncClient(timeout=120) as client:
        response = await client.post(
            f"{DEEPSEEK_BASE_URL}/chat/completions",
            headers={
                "Authorization": f"Bearer {DEEPSEEK_API_KEY}",
                "Content-Type": "application/json",
            },
            json={
                "model": DEEPSEEK_MODEL,
                "messages": [
                    {"role": "system", "content": "You output strict JSON only."},
                    {"role": "user", "content": prompt},
                ],
                "temperature": 0.1,
            },
        )
    if response.status_code >= 400:
        raise HTTPException(status_code=502, detail=response.text)

    content = response.json()["choices"][0]["message"]["content"]
    try:
        constraints = json.loads(strip_json_fence(content))
    except json.JSONDecodeError as exc:
        raise HTTPException(status_code=502, detail=f"DeepSeek returned invalid JSON: {exc}") from exc

    if not isinstance(constraints, dict):
        raise HTTPException(status_code=502, detail="DeepSeek constraints response is not an object")
    return constraints


@app.get("/health")
def health() -> dict:
    return {"ok": True, "model": DEEPSEEK_MODEL}


@app.post("/generate")
async def generate(
    report: UploadFile = File(...),
    questionnaire: UploadFile = File(...),
    sample_size: int = Form(...),
    max_fuzzy_rate: float = Form(0.2),
):
    if sample_size <= 0:
        raise HTTPException(status_code=400, detail="sample_size must be positive")
    if not 0 <= max_fuzzy_rate <= 1:
        raise HTTPException(status_code=400, detail="max_fuzzy_rate must be between 0 and 1")

    workdir = Path(tempfile.mkdtemp(prefix="survey_skill_"))
    report_path = workdir / (report.filename or "report.txt")
    questionnaire_path = workdir / (questionnaire.filename or "questionnaire.xlsx")
    output_dir = workdir / "output"
    constraints_path = workdir / "constraints.json"

    report_path.write_bytes(await report.read())
    questionnaire_path.write_bytes(await questionnaire.read())

    questions = question_payload(questionnaire_path)
    report_text = read_report_text(report_path)
    constraints = await generate_constraints(report_text, questions)
    constraints_path.write_text(json.dumps(constraints, ensure_ascii=False, indent=2), encoding="utf-8")

    proc = subprocess.run(
        [
            sys.executable,
            str(ROOT / "scripts" / "generate_answer_table.py"),
            "--questionnaire",
            str(questionnaire_path),
            "--prompt",
            f"请生成{sample_size}人答案选项表",
            "--constraints",
            str(constraints_path),
            "--max-fuzzy-rate",
            str(max_fuzzy_rate),
            "--output-dir",
            str(output_dir),
        ],
        text=True,
        capture_output=True,
        check=False,
    )
    if proc.returncode != 0:
        raise HTTPException(status_code=500, detail={"stdout": proc.stdout, "stderr": proc.stderr})

    result = json.loads(proc.stdout)
    output_path = Path(result["output"])
    return FileResponse(
        output_path,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        filename=output_path.name,
    )


@app.get("/")
def root() -> JSONResponse:
    return JSONResponse({"service": "survey-self-check", "usage": "POST /generate"})
