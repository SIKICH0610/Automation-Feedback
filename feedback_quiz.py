from __future__ import annotations

import re

from feedback_common import (
    SECOND_QUIZ_COLUMNS,
    StudentRow,
    additional_comment_for_local_message,
    join_naturally,
    sentence_join_zh,
    value_from_any_column,
)

def quiz_score_from_remark(remark: str) -> float | None:
    match = re.search(r"(\d+(?:\.\d+)?)\s*/\s*8", remark)
    if not match:
        return None
    return float(match.group(1))

def numeric_score_from_text(text: str) -> float | None:
    match = re.search(r"(\d+(?:\.\d+)?)", text)
    if not match:
        return None
    return float(match.group(1))

def second_quiz_score_parts(text: str) -> tuple[float | None, float | None]:
    plus_match = re.search(
        r"(\d+(?:\.\d+)?)\s*(?:/\s*4)?\s*(?:\+|＋)\s*(\d+(?:\.\d+)?)",
        text,
    )
    if plus_match:
        return float(plus_match.group(1)), float(plus_match.group(2))

    fraction_match = re.search(r"(\d+(?:\.\d+)?)\s*/\s*4", text)
    if fraction_match:
        return float(fraction_match.group(1)), None

    score = numeric_score_from_text(text)
    return score, None

def second_quiz_total_score(text: str) -> float | None:
    base_score, bonus_score = second_quiz_score_parts(text)
    if base_score is None:
        return None
    return base_score + (bonus_score or 0)

def score_display_from_text(text: str, default_denominator: str = "/4") -> str:
    base_score, bonus_score = second_quiz_score_parts(text)
    if base_score is not None and bonus_score is not None:
        return f"{format_score(base_score)}{default_denominator} + {format_score(bonus_score)} bonus"

    fraction_match = re.search(r"(\d+(?:\.\d+)?)\s*/\s*(\d+(?:\.\d+)?)", text)
    if fraction_match:
        score = format_score(float(fraction_match.group(1)))
        denominator = format_score(float(fraction_match.group(2)))
        return f"{score}/{denominator}"

    score = numeric_score_from_text(text)
    if score is None:
        return text
    return f"{format_score(score)}{default_denominator}"

def second_quiz_score_display(text: str, is_chinese: bool) -> str:
    base_score, bonus_score = second_quiz_score_parts(text)
    if base_score is not None and bonus_score is not None:
        base = f"{format_score(base_score)}/4"
        if is_chinese:
            return f"{base}，bonus 为 {format_score(bonus_score)}"
        bonus_word = "point" if bonus_score == 1 else "points"
        return f"{base} plus {format_score(bonus_score)} bonus {bonus_word}"
    return score_display_from_text(text)

def format_score(score: float) -> str:
    if score.is_integer():
        return str(int(score))
    return str(score).rstrip("0").rstrip(".")

def strip_quiz_score_text(remark: str) -> str:
    text = re.sub(r"quiz\s*1?\s*[，,]?\s*", "", remark, flags=re.IGNORECASE)
    text = re.sub(r"满分\s*\d+(?:\.\d+)?\s*/\s*8[，,]?", "", text)
    text = re.sub(r"\d+(?:\.\d+)?\s*/\s*8[，,]?", "", text)
    return text.strip(" ，,。.")

def chinese_quiz_score_sentence(score: float) -> str:
    score_text = format_score(score)
    if score < 6.5:
        return (
            f"本次 quiz 得分为 {score_text}/8，低于本次班级平均分 6.9/8，"
            "建议家长最近多关注一下孩子对前半部分知识的掌握情况。"
        )
    if score >= 7.5:
        return f"本次 quiz 得分为 {score_text}/8，整体表现很不错，说明前半部分知识掌握比较扎实。"
    return f"本次 quiz 得分为 {score_text}/8，整体接近或达到本次班级平均水平。"

def english_quiz_score_sentence(score: float) -> str:
    score_text = format_score(score)
    if score < 6.5:
        return (
            f"The quiz score was {score_text}/8, which is below the class average of 6.9/8. "
            "It would be helpful to keep a closer eye on the student's current understanding."
        )
    if score >= 7.5:
        return (
            f"The quiz score was {score_text}/8, which is a strong result and shows solid understanding "
            "of the first-half topics."
        )
    return f"The quiz score was {score_text}/8, which is around the class average."

def second_quiz_text(student: StudentRow) -> str:
    return value_from_any_column(student, SECOND_QUIZ_COLUMNS)

def second_quiz_score_sentence(student: StudentRow, is_chinese: bool) -> str:
    text = second_quiz_text(student)
    if not text:
        return ""

    score = second_quiz_total_score(text)
    score_display = second_quiz_score_display(text, is_chinese)

    lower = text.lower()
    has_bonus = "bonus" in lower or "加" in text
    has_bonus_amount = "+" in text or "＋" in text
    without_bonus = "不加" in text or "without" in lower

    if is_chinese:
        if has_bonus and not without_bonus and not has_bonus_amount:
            base = f"第二次 quiz 的成绩为 {score_display}（包含 bonus）"
        else:
            base = f"第二次 quiz 的成绩为 {score_display}"

        if score is not None:
            if score >= 3.5:
                return f"{base}，整体表现不错，说明等腰三角形、角平分线、垂直平分线和勾股定理这部分掌握得比较稳定。"
            if score >= 3:
                return f"{base}，整体接近本次 quiz 的平均水平，后续可以继续加强证明书写的规范性，特别是使用 HL 时要明确直角三角形条件。"
            return f"{base}，建议接下来重点复习等腰三角形、角平分线、垂直平分线和勾股定理，并特别注意使用 HL 证明全等时要先明确直角三角形。"
        return f"第二次 quiz 的记录为：{text}。"

    if has_bonus and not without_bonus and not has_bonus_amount:
        base = f"The second quiz score was {score_display} with the bonus included"
    else:
        base = f"The second quiz score was {score_display}"

    if score is not None:
        if score >= 3.5:
            return f"{base}, which is a solid result and shows steady understanding of the recent triangle topics."
        if score >= 3:
            return f"{base}, which is around the class average. The next step is to keep proof writing precise, especially when using HL by clearly stating the right-triangle condition."
        return f"{base}. It would be helpful to review isosceles triangles, angle bisectors, perpendicular bisectors, and the Pythagorean Theorem, while paying close attention to the right-triangle condition when using HL."
    return f"The second quiz note is: {text}."

def chinese_observation_sentence(name: str, observations: list[str], language: str) -> str:
    if not observations:
        return ""
    return f"结合课堂表现来看，{name} {join_naturally(observations, language)}。"

def english_observation_sentence(name: str, observations: list[str], language: str) -> str:
    if not observations:
        return ""
    return f"During class, {name} {join_naturally(observations, language)}."

QUIZ_FOCUS_CATEGORIES = {
    "calculation": {
        "pattern": r"计算|算|accuracy|accurate|careless|detail|细节|不仔细",
        "negative": r"计算.{0,8}(不|错|需要|仔细)|细节.{0,8}(不|需要)|careless|calculation.{0,16}(need|error|mistake)",
        "positive": r"准确|不错|很好|good|accurate|strong|solid",
        "zh_positive": "计算准确性整体比较稳定",
        "zh_needs": "计算细节需要更加仔细",
        "en_positive": "maintaining solid calculation accuracy",
        "en_needs": "checking calculation details more carefully",
    },
    "proof_structure": {
        "pattern": r"证明|proof|reason|理由|步骤|过程|严谨|rigor|concise|完整|漏",
        "negative": r"证明.{0,8}(不|需要|错|漏|欠缺|不够)|proof.{0,16}(need|missing|unclear)|reason.{0,12}(需要|欠缺|unclear|missing)|步骤.{0,8}(漏|不完整|需要)|过程.{0,8}(不完整|需要)",
        "positive": r"严谨|完整|不错|很好|strong|good|rigor|concise|solid",
        "zh_positive": "证明书写方面已经有不错的基础",
        "zh_needs": "证明书写需要更加规范，步骤和理由要写完整",
        "en_positive": "continuing to build on a solid foundation in proof writing",
        "en_needs": "keeping proof writing organized, with complete steps and reasons",
    },
    "geometry_relationships": {
        "pattern": r"平行|角|全等|triangle|congruent|HL|SAS|SSA|angle|parallel",
        "negative": r"关系.{0,8}(不|需要|错|混)|定理.{0,8}(不|需要|错|混)|角.{0,8}(不|需要|错|混)|triangle.{0,16}(need|unclear|mistake)|theorem.{0,16}(need|unclear|mistake)",
        "positive": r"掌握|稳定|不错|很好|good|strong|solid|understand",
        "zh_positive": "图形关系和定理使用整体比较稳定",
        "zh_needs": "图形关系和定理使用还可以继续巩固",
        "en_positive": "showing steady use of diagram relationships and theorems",
        "en_needs": "continuing to review diagram relationships and theorem use",
    },
    "classroom_habits": {
        "pattern": r"课堂|专注|分心|提醒|参与|打扰|focus|participat|attention|habit|distract",
        "negative": r"分心|提醒|打扰|不专注|课堂.{0,8}(习惯|提醒|打扰)|distract|remind|attention.{0,12}(need|issue)",
        "positive": r"专注|积极|主动|认真|focused|active|participat",
        "zh_positive": "课堂专注和参与度整体不错",
        "zh_needs": "课堂专注和课堂习惯方面还可以继续调整",
        "en_positive": "maintaining steady focus and class participation",
        "en_needs": "continuing to build stronger focus and classroom habits",
    },
}


def focus_points_from_categories(text: str, language: str) -> list[str]:
    points: list[str] = []
    for category in QUIZ_FOCUS_CATEGORIES.values():
        if not re.search(category["pattern"], text, flags=re.IGNORECASE):
            continue
        is_positive = bool(re.search(category["positive"], text, flags=re.IGNORECASE))
        is_negative = bool(re.search(category["negative"], text, flags=re.IGNORECASE))
        tone = "positive" if is_positive and not is_negative else "needs"
        points.append(category[f"{language}_{tone}"])
    return list(dict.fromkeys(points))

def chinese_quiz_focus_sentences(remark: str) -> list[str]:
    return focus_points_from_categories(remark, "zh")

def english_quiz_focus_sentences(remark: str) -> list[str]:
    return focus_points_from_categories(remark, "en")

def quiz_comment_paragraph(
    student: StudentRow,
    observations: list[str],
    is_chinese: bool,
    *,
    include_observations: bool = False,
) -> str | None:
    remark = str(student.values.get("Remark for Student") or "").strip()
    second_quiz_sentence = second_quiz_score_sentence(student, is_chinese)
    if "quiz" not in remark.lower() and "/8" not in remark and not second_quiz_sentence:
        return None

    name = student.first_name or student.full_name
    score = quiz_score_from_remark(remark)
    additional_comment = additional_comment_for_local_message(student, is_chinese)

    if is_chinese:
        sentences: list[str] = []
        if score is not None:
            sentences.append(chinese_quiz_score_sentence(score))
        if second_quiz_sentence:
            sentences.append(second_quiz_sentence)

        focus = chinese_quiz_focus_sentences(remark)
        if focus:
            focus_text = sentence_join_zh(focus)
            if score is not None and score >= 7.5:
                sentences.append(f"{name} 这次表现比较扎实，后续可以继续保持，尤其是{focus_text}。")
            else:
                sentences.append(f"从这次 quiz 和课堂情况来看，{name} 接下来可以重点关注{focus_text}。")
        else:
            if remark:
                sentences.append(f"老师已经把{name}本次 quiz 相关的课堂记录整理进反馈，后续会继续结合错题类型和证明书写情况进行跟进。")

        if include_observations and observations:
            sentences.append(chinese_observation_sentence(name, observations, student.language))
        if additional_comment:
            sentences.append(additional_comment + "。")
        sentences.append("后续可以继续复习前半部分知识，并在计算准确性和证明步骤完整性上多检查。")
        return "".join(sentences)

    sentences = []
    if score is not None:
        sentences.append(english_quiz_score_sentence(score))
    if second_quiz_sentence:
        sentences.append(second_quiz_sentence)

    focus = english_quiz_focus_sentences(remark)
    if focus:
        focus_text = ", and ".join(focus)
        sentences.append(f"{name} can build from this by continuing to focus on {focus_text}.")
    elif remark:
        sentences.append(
            f"The teacher's quiz-related classroom note has been included, and we will keep following {name}'s error patterns and proof-writing habits."
        )

    if include_observations and observations:
        sentences.append(english_observation_sentence(name, observations, student.language))
    if additional_comment:
        sentences.append(additional_comment + ".")
    sentences.append("Going forward, reviewing the first-half topics while checking calculation details and proof structure carefully will help keep the foundation strong.")
    return " ".join(sentences)

