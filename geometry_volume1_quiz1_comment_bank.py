from __future__ import annotations

from dataclasses import dataclass
import re

from geometry_volume1_quiz2_comment_bank import build_geometry_volume1_quiz2_comment


@dataclass(frozen=True)
class QuizIssue:
    question: int
    title: str
    patterns: tuple[str, ...]
    chinese: str
    english: str


GEOMETRY_VOLUME1_QUIZ1_ISSUES: tuple[QuizIssue, ...] = (
    QuizIssue(
        question=4,
        title="Congruent triangle corresponding parts",
        patterns=(
            r"全等三角形.*对应",
            r"对应[边變变角]",
            r"corresponding.*congruent triangle",
            r"cpctc|corresponding parts",
        ),
        chinese=(
            "第 4 题主要需要注意全等三角形中的对应关系。做这类题时，孩子需要先看清楚两个三角形中哪些边或角是对应的，"
            "再使用对应边或对应角相等，避免把不对应的边角直接拿来用。"
        ),
        english=(
            "For question 4, the main point is the corresponding parts of congruent triangles. The student should first identify which sides or angles match between the two triangles before using them as equal."
        ),
    ),
    QuizIssue(
        question=5,
        title="Parallel line converse reasoning",
        patterns=(
            r"parallel lines are congruent",
            r"two lines are parallel",
            r"内错角.*平行",
            r"同位角.*平行",
            r"alternate interior.*parallel",
            r"corresponding angle.*parallel",
        ),
        chinese=(
            "第 5 题需要把平行线的判定理由写清楚。也就是说，要说明是因为哪一组同位角或内错角相等，所以可以推出两条直线平行，"
            "不能只写成比较笼统的结论。"
        ),
        english=(
            "For question 5, the student needs to state the converse reasoning for parallel lines more clearly. They should explain which pair of corresponding or alternate interior angles are congruent, and then conclude that the two lines are parallel."
        ),
    ),
    QuizIssue(
        question=9,
        title="Pythagorean theorem equation setup",
        patterns=(
            r"勾股定理.*方程",
            r"pythagorean.*equation",
            r"right triangle.*equation",
            r"a\^2\s*\+\s*b\^2",
        ),
        chinese=(
            "第 9 题需要加强用勾股定理建立方程的过程。孩子需要先准确判断直角边和斜边，再列出平方关系，"
            "解方程时也要注意计算细节。"
        ),
        english=(
            "For question 9, the student should keep practicing how to set up an equation with the Pythagorean Theorem. They need to identify the legs and hypotenuse first, then write the squared relationship carefully and check the calculation."
        ),
    ),
    QuizIssue(
        question=10,
        title="HL and invalid SSA congruence",
        patterns=(
            r"\bHL\b",
            r"\bSSA\b",
            r"ssa is not true",
            r"ssa.*not",
            r"斜边.*直角边",
            r"HL.*直角三角形",
        ),
        chinese=(
            "第 10 题需要继续区分三角形全等的判定条件。HL 只能用于直角三角形，并且需要斜边和一条直角边。"
            "SSA 不能作为三角形全等的判定方法，写证明时需要根据题目条件选择合适的 SSS、SAS、ASA、AAS 或 HL。"
        ),
        english=(
            "For question 10, the student needs to distinguish the triangle congruence criteria more carefully. HL only works for right triangles with the hypotenuse and one leg, while SSA is not a valid congruence reason."
        ),
    ),
)


def question_numbers_from_note(note: str) -> set[int]:
    stripped = re.sub(
        r"\b(?:physical\s*)?quiz\s*\d+\b|\b(?:first|second|third)\s+physical\s+quiz\b",
        " ",
        note.strip(),
        flags=re.IGNORECASE,
    ).strip(" -:：,，、;；")
    numbers: set[int] = set()

    for pattern in (
        r"\bq(?:uestion)?\s*(\d{1,2})\b",
        r"第\s*(\d{1,2})\s*题",
    ):
        numbers.update(
            int(value) for value in re.findall(pattern, stripped, flags=re.IGNORECASE)
        )

    if re.fullmatch(
        r"(?:\s*(?:q(?:uestion)?\s*)?\d{1,2}\s*[,，、;；-]?)+\s*",
        stripped,
        flags=re.IGNORECASE,
    ):
        numbers.update(int(value) for value in re.findall(r"\d{1,2}", stripped))

    return numbers


def match_geometry_volume1_quiz1_issues(note: str) -> list[QuizIssue]:
    number_matches = question_numbers_from_note(note)
    matches: list[QuizIssue] = []
    for issue in GEOMETRY_VOLUME1_QUIZ1_ISSUES:
        if issue.question in number_matches or any(
            re.search(pattern, note, flags=re.IGNORECASE) for pattern in issue.patterns
        ):
            matches.append(issue)
    return matches


def question_numbers_for_note(note: str) -> list[int]:
    return [issue.question for issue in match_geometry_volume1_quiz1_issues(note)]


def comment_sentences_for_note(note: str, *, language: str = "Chinese") -> list[str]:
    is_chinese = language.lower().startswith("chinese")
    return [
        issue.chinese if is_chinese else issue.english
        for issue in match_geometry_volume1_quiz1_issues(note)
    ]


def build_geometry_volume1_quiz1_comment(note: str, *, language: str = "Chinese") -> str:
    if re.search(r"\b(?:physical\s*)?quiz\s*2\b|\bsecond\s+physical\s+quiz\b", note, flags=re.IGNORECASE):
        routed_note = re.sub(
            r"\b(?:physical\s*)?quiz\s*2\b|\bsecond\s+physical\s+quiz\b|[:：-]",
            " ",
            note,
            flags=re.IGNORECASE,
        )
        return build_geometry_volume1_quiz2_comment(routed_note, language=language)

    sentences = comment_sentences_for_note(note, language=language)
    if not sentences:
        return ""
    if language.lower().startswith("chinese"):
        return "".join(sentences)
    return " ".join(sentences)


def main() -> None:
    import argparse

    parser = argparse.ArgumentParser(
        description="Match Geometry Volume 1 quiz 1 question notes to reusable comment sentences."
    )
    parser.add_argument(
        "note",
        help="Question numbers or teacher shorthand, such as '4,9,10', '勾股定理建立方程', or 'SSA is not true'.",
    )
    parser.add_argument("--language", default="Chinese", choices=["Chinese", "English"])
    args = parser.parse_args()

    matches = match_geometry_volume1_quiz1_issues(args.note)
    if not matches:
        print("No matching Geometry Volume 1 quiz 1 issue found.")
        return

    print("Matched questions:", ", ".join(str(issue.question) for issue in matches))
    print(build_geometry_volume1_quiz1_comment(args.note, language=args.language))


if __name__ == "__main__":
    main()
