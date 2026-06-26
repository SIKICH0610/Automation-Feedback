from __future__ import annotations

import re

from feedback_common import (
    StudentRow,
    additional_comment_for_local_message,
    join_naturally,
)
from feedback_quiz import strip_quiz_score_text


GENERAL_NOTE_CATEGORIES = {
    "proof": {
        "pattern": r"证明|proof|reason|理由|步骤|过程|严谨|rigor|完整|漏",
        "negative": r"证明.{0,8}(不|需要|错|漏|欠缺|不够)|proof.{0,16}(need|missing|unclear)|reason.{0,12}(需要|欠缺|unclear|missing)|步骤.{0,8}(漏|不完整|需要)|过程.{0,8}(不完整|需要)",
        "positive": r"严谨|完整|不错|很好|strong|good|rigor|concise",
        "zh_positive": "证明书写方面有不错的基础",
        "zh_needs": "证明书写和理由表达还可以继续加强",
        "en_positive": "has a solid foundation in proof writing",
        "en_needs": "can continue strengthening proof writing and reasoning",
    },
    "calculation": {
        "pattern": r"计算|算|accuracy|accurate|careless|detail|细节|不仔细",
        "negative": r"计算.{0,8}(不|错|需要|仔细)|细节.{0,8}(不|需要)|careless|calculation.{0,16}(need|error|mistake)",
        "positive": r"准确|不错|很好|good|accurate|strong",
        "zh_positive": "计算准确性整体不错",
        "zh_needs": "计算细节需要继续检查",
        "en_positive": "showed solid calculation accuracy",
        "en_needs": "should continue checking calculation details carefully",
    },
    "classroom_habits": {
        "pattern": r"课堂|专注|分心|提醒|参与|打扰|focus|participat|attention|habit|distract",
        "negative": r"分心|提醒|打扰|不专注|课堂.{0,8}(习惯|提醒|打扰)|distract|remind|attention.{0,12}(need|issue)",
        "positive": r"专注|积极|主动|认真|focused|active|participat",
        "zh_positive": "课堂状态和参与度整体不错",
        "zh_needs": "课堂专注和课堂习惯方面还可以继续调整",
        "en_positive": "showed steady focus and class participation",
        "en_needs": "can continue building stronger focus and classroom habits",
    },
    "work_habits": {
        "pattern": r"笔记|note|作图|画图|过程|步骤|homework|practice",
        "negative": r"笔记.{0,8}(不|需要|缺)|过程.{0,8}(不完整|需要)|步骤.{0,8}(漏|需要)|note.{0,16}(need|missing|incomplete)",
        "positive": r"工整|完整|认真|good|complete|neat",
        "zh_positive": "课堂记录和做题习惯整体比较稳定",
        "zh_needs": "课堂记录和做题过程还可以更加完整",
        "en_positive": "showed steady note-taking and work habits",
        "en_needs": "can make notes and written work more complete",
    },
}

def regular_teacher_note(student: StudentRow, *, include_quiz_remark: bool = True) -> str:
    remark = str(student.values.get("Remark for Student") or "").strip()
    if not remark:
        return ""
    if "quiz" in remark.lower() or "/8" in remark:
        if not include_quiz_remark:
            return ""
        return strip_quiz_score_text(remark)
    return remark

def parent_facing_note_points(remark: str, is_chinese: bool) -> list[str]:
    if not remark:
        return []

    points: list[str] = []
    for category in GENERAL_NOTE_CATEGORIES.values():
        if not re.search(category["pattern"], remark, flags=re.IGNORECASE):
            continue
        is_positive = bool(re.search(category["positive"], remark, flags=re.IGNORECASE))
        is_negative = bool(re.search(category["negative"], remark, flags=re.IGNORECASE))
        if is_chinese:
            points.append(category["zh_positive"] if is_positive and not is_negative else category["zh_needs"])
        else:
            points.append(category["en_positive"] if is_positive and not is_negative else category["en_needs"])
    return list(dict.fromkeys(points))

def parent_facing_note_summary(name: str, remark: str, is_chinese: bool) -> str:
    points = parent_facing_note_points(remark, is_chinese)
    if is_chinese:
        if points:
            return f"从课堂补充记录来看，{name} 在{'，'.join(points)}。"
        return f"老师的课堂补充记录也已经整理进反馈，后续会继续关注{name}的课堂状态和学习习惯。"

    if points:
        return f"Based on the teacher's classroom note, {name} " + ", and ".join(points) + "."
    return f"The teacher's additional classroom note has been included, and we will continue to monitor {name}'s class habits and learning progress."

def general_comment_paragraph(
    student: StudentRow,
    observations: list[str],
    is_chinese: bool,
    *,
    include_quiz_remark: bool = True,
) -> str:
    name = student.first_name or student.full_name
    remark = regular_teacher_note(student, include_quiz_remark=include_quiz_remark)
    additional_comment = additional_comment_for_local_message(student, is_chinese)

    if is_chinese:
        if remark:
            base = parent_facing_note_summary(name, remark, is_chinese)
            if additional_comment:
                return f"{base}{additional_comment}。"
            return base
        if observations:
            base = (
                f"{name} 今天整体表现稳定，"
                f"{join_naturally(observations, student.language)}。之后可以继续保持好的课堂习惯，"
                "同时在证明逻辑和细节检查上多练习。"
            )
            if additional_comment:
                return f"{base}{additional_comment}。"
            return base
        return f"{name} 今天的课堂表现已记录，之后可以继续保持稳定的学习节奏。"

    if observations:
        base = (
            f"{name} had a steady class today and "
            f"{join_naturally(observations, student.language)}. Continuing to build proof logic "
            "while checking details carefully will be helpful."
        )
        if additional_comment:
            return f"{base} {additional_comment}."
        return base
    if remark:
        base = parent_facing_note_summary(name, remark, is_chinese)
        if additional_comment:
            return f"{base} {additional_comment}."
        return base
    return f"{name}'s classroom notes have been recorded for today's lesson."
