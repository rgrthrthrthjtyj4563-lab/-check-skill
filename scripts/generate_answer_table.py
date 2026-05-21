#!/usr/bin/env python3
"""Generate a single revised questionnaire answer-option wide table.

This script is intentionally conservative. It does not rewrite medical report
content. It treats an xlsx questionnaire detail file as source distribution,
removes metadata columns, samples answers by original option frequency, and
repairs rows that violate hard single-respondent logic rules. The final
artifact is one workbook with one sheet only.
"""

from __future__ import annotations

import argparse
import json
import random
import re
from collections import Counter, OrderedDict
from dataclasses import dataclass
from datetime import date, timedelta
from pathlib import Path
from typing import Iterable

from openpyxl import Workbook, load_workbook


META_COLUMNS = {
    "问卷名称",
    "品种",
    "被问卷人姓名",
    "问卷日期",
    "手机版本",
    "openid",
    "open id",
    "手机号",
    "手机号码",
    "地址",
    "备注",
    "受访者编号",
    "respondent_id",
    "respondent id",
}


NEGATIVE_PATTERNS = [
    "不用药",
    "几乎不用药",
    "未使用",
    "没使用",
    "不使用",
    "未听说",
    "不了解",
    "不清楚",
    "不确定",
    "无法评价",
    "无相关经历",
    "不适用",
]

USAGE_EVAL_QUESTION_PATTERNS = [
    "使用清咽滴丸",
    "购买清咽滴丸",
    "推荐清咽滴丸",
    "医保报销",
    "价格敏感度",
]

POSITIVE_OR_EVALUATIVE_PATTERNS = [
    "起效",
    "缓解",
    "效果",
    "便利",
    "推荐",
    "满意",
    "认可",
    "每次都使用医保",
    "偶尔使用医保",
    "当前价格合理",
    "促销活动时才购买",
    "疗效好",
]


@dataclass
class Question:
    number: int
    text: str
    options: OrderedDict[str, int]


def normalize_header(value: object) -> str:
    return str(value or "").strip()


def is_meta_column(header: str) -> bool:
    normalized = header.strip().lower()
    if normalized in META_COLUMNS:
        return True
    return any(token in normalized for token in ["openid", "open id", "手机号", "手机号码"])


def parse_sample_size(prompt: str) -> int | None:
    patterns = [
        r"(?:样本量\s*)?N\s*[=:：]\s*(\d+)",
        r"生成\s*(\d+)\s*(?:个)?人",
        r"(\d+)\s*个人",
        r"受访者\s*(\d+)\s*人?",
        r"(\d+)\s*份问卷",
        r"样本量\s*(\d+)",
        r"(\d+)\s*个样本",
    ]
    for pattern in patterns:
        match = re.search(pattern, prompt, flags=re.IGNORECASE)
        if match:
            return int(match.group(1))
    return None


def read_questions(path: Path) -> tuple[list[str], list[Question], dict[str, str]]:
    wb = load_workbook(path, read_only=True, data_only=True)
    ws = wb.active
    rows = ws.iter_rows(values_only=True)
    headers = [normalize_header(v) for v in next(rows)]
    meta_values: dict[str, str] = {}
    raw_rows = list(rows)

    if raw_rows:
        first = raw_rows[0]
        for idx, header in enumerate(headers):
            if header and is_meta_column(header) and idx < len(first) and first[idx] is not None:
                meta_values[header] = str(first[idx]).strip()

    meta_columns: list[str] = []
    questions: list[Question] = []
    for idx, header in enumerate(headers):
        if not header:
            continue
        if is_meta_column(header):
            meta_columns.append(header)
            continue
        counts: OrderedDict[str, int] = OrderedDict()
        for row in raw_rows:
            if idx >= len(row):
                continue
            value = row[idx]
            if value is None:
                continue
            option = str(value).strip()
            if not option:
                continue
            counts[option] = counts.get(option, 0) + 1
        if counts:
            questions.append(Question(len(questions) + 1, header, counts))
    return meta_columns, questions, meta_values


def weighted_choice(options: OrderedDict[str, int]) -> str:
    values = list(options.keys())
    weights = list(options.values())
    return random.choices(values, weights=weights, k=1)[0]


def contains_any(text: str, patterns: Iterable[str]) -> bool:
    return any(pattern in text for pattern in patterns)


def find_question(questions: list[Question], patterns: Iterable[str]) -> Question | None:
    for question in questions:
        if contains_any(question.text, patterns):
            return question
    return None


def row_conflicts(row: dict[str, str], questions: list[Question]) -> list[dict[str, str]]:
    conflicts: list[dict[str, str]] = []
    q_by_text = {q.text: q for q in questions}

    for q_text, answer in row.items():
        if contains_any(answer, ["不用药", "几乎不用药", "未使用", "不清楚", "不确定"]):
            for later_text, later_answer in row.items():
                later_q = q_by_text[later_text]
                current_q = q_by_text[q_text]
                if later_q.number <= current_q.number:
                    continue
                if contains_any(later_text, USAGE_EVAL_QUESTION_PATTERNS) and contains_any(
                    later_answer, POSITIVE_OR_EVALUATIVE_PATTERNS
                ):
                    conflicts.append(
                        {
                            "前置题号及选项": f"Q{current_q.number} {answer}",
                            "后续题号及选项": f"Q{later_q.number} {later_answer}",
                            "冲突类型": "前提不成立",
                            "冲突原因": "前置答案否定或弱化用药前提，后续答案需要明确使用或评价前提。",
                            "风险等级": "高",
                            "修改建议": "该受访者后续使用评价题应改为不确定/无法评价，或前置题改为存在用药行为的选项。",
                        }
                    )

    for q_text, answer in row.items():
        if "不使用医保" in answer or "全额自费" in answer:
            for other_text, other_answer in row.items():
                other_q = q_by_text[other_text]
                current_q = q_by_text[q_text]
                if other_text == q_text:
                    continue
                if "医保报销" in other_text and contains_any(other_answer, ["每次都使用医保", "偶尔使用医保"]):
                    conflicts.append(
                        {
                            "前置题号及选项": f"Q{current_q.number} {answer}",
                            "后续题号及选项": f"Q{other_q.number} {other_answer}",
                            "冲突类型": "医保使用互斥",
                            "冲突原因": "同一受访者不能同时选择不使用医保和使用医保报销。",
                            "风险等级": "高",
                            "修改建议": "保留一种医保使用状态。",
                        }
                    )
    return conflicts


def safer_option(question: Question) -> str:
    options = list(question.options.keys())
    for option in options:
        if contains_any(option, ["不确定", "无法评价", "不清楚"]):
            return option
    for option in options:
        if not contains_any(option, POSITIVE_OR_EVALUATIVE_PATTERNS):
            return option
    return options[0]


def repair_row(row: dict[str, str], questions: list[Question]) -> dict[str, str]:
    repaired = dict(row)
    for _ in range(8):
        conflicts = row_conflicts(repaired, questions)
        if not conflicts:
            return repaired
        for conflict in conflicts:
            match = re.match(r"Q(\d+)\s", conflict["后续题号及选项"])
            if not match:
                continue
            q_number = int(match.group(1))
            question = questions[q_number - 1]
            repaired[question.text] = safer_option(question)
    return repaired


def generate_answers(questions: list[Question], sample_size: int) -> tuple[list[dict[str, str]], list[dict[str, str]]]:
    answers: list[dict[str, str]] = []
    all_conflicts: list[dict[str, str]] = []
    for _ in range(sample_size):
        row = {question.text: weighted_choice(question.options) for question in questions}
        row = repair_row(row, questions)
        conflicts = row_conflicts(row, questions)
        all_conflicts.extend(conflicts)
        answers.append(row)
    return answers, all_conflicts


def write_outputs(
    output_dir: Path,
    source_path: Path,
    prompt: str,
    sample_size: int | None,
    meta_columns: list[str],
    questions: list[Question],
    meta_values: dict[str, str],
    answers: list[dict[str, str]],
    conflicts: list[dict[str, str]],
) -> dict[str, str]:
    output_dir.mkdir(parents=True, exist_ok=True)
    stem = source_path.stem
    output_path = output_dir / f"{stem}_修改后问卷_答案选项宽表.xlsx"

    if sample_size is None:
        return {"output": "", "error": "缺少生成人数，无法生成答案选项表。"}

    wb = Workbook()
    ws = wb.active
    ws.title = "修改后问卷及答案表"
    headers = ["问卷日期"]
    headers.extend(f"问题{question.number}：{question.text}" for question in questions)
    headers.extend(["手机版本", "OpenID"])
    ws.append(headers)

    start = date.today()
    for row in answers:
        values = [(start - timedelta(days=random.randint(0, 30))).isoformat()]
        values.extend(row[question.text] for question in questions)
        values.extend(["", ""])
        ws.append(values)

    wb.save(output_path)
    return {"output": str(output_path)}


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--questionnaire", required=True, type=Path)
    parser.add_argument("--prompt", required=True)
    parser.add_argument("--output-dir", required=True, type=Path)
    parser.add_argument("--seed", type=int, default=20240521)
    args = parser.parse_args()

    random.seed(args.seed)
    sample_size = parse_sample_size(args.prompt)
    meta_columns, questions, meta_values = read_questions(args.questionnaire)
    answers: list[dict[str, str]] = []
    conflicts: list[dict[str, str]] = []
    if sample_size is not None:
        answers, conflicts = generate_answers(questions, sample_size)

    outputs = write_outputs(
        args.output_dir,
        args.questionnaire,
        args.prompt,
        sample_size,
        meta_columns,
        questions,
        meta_values,
        answers,
        conflicts,
    )
    print(
        json.dumps(
            {
                "sample_size": sample_size,
                "effective_question_count": len(questions),
                "meta_columns": meta_columns,
                "conflict_count_after_generation": len(conflicts),
                **outputs,
            },
            ensure_ascii=False,
            indent=2,
        )
    )
    return 1 if sample_size is None or conflicts else 0


if __name__ == "__main__":
    raise SystemExit(main())
