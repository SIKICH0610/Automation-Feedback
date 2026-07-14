from __future__ import annotations

from dataclasses import dataclass
import re


CLASS_NAME = "AMC10 Intro"
QUIZ_NAME = "Quiz 1"


@dataclass(frozen=True)
class QuizIssue:
    question: int
    title: str
    patterns: tuple[str, ...]
    chinese: str
    english: str


AMC10_QUIZ1_ISSUES: tuple[QuizIssue, ...] = (
    QuizIssue(
        question=3,
        title="Exterior Angle Theorem equation setup",
        patterns=(
            r"exterior angle theorem",
            r"external angle theorem",
            r"unknown.*x",
            r"triangle interior angle sum",
            r"内角和.*方程",
            r"外角.*x",
        ),
        chinese=(
            "第 3 题需要用 Exterior Angle Theorem 先把未知量表示为 x，"
            "再通过三角形内角和建立方程并解出 x。做这类题时，关键是先把外角和不相邻内角的关系整理清楚，"
            "再把三角形内部的角度关系代入方程，避免直接猜角度。"
        ),
        english=(
            "For question 3, the student needs to use the Exterior Angle Theorem to express the unknown in terms of x, "
            "then use the triangle interior angle sum to set up an equation and solve for x. The key is to organize the exterior-angle relationship before writing the equation."
        ),
    ),
    QuizIssue(
        question=4,
        title="Integrated angle chasing in triangle DPC",
        patterns=(
            r"integrated thinking",
            r"triangle DPC",
            r"\bDPC\b",
            r"measurement of angle P",
            r"angle P",
            r"三角形DPC",
            r"DPC.*两个角",
        ),
        chinese=(
            "第 4 题需要用 integrated thinking 同时表示三角形 DPC 中的两个角，"
            "最后再用三角形内角和算出 angle P 的度数。这里不能只看单独一个角，"
            "而是要把图中相关角度放在一起追踪，找到两个角和 x 之间的关系后再统一计算。"
        ),
        english=(
            "For question 4, the student needs to use integrated thinking to express two angles in triangle DPC, "
            "then use the triangle interior angle sum to find the measure of angle P. This problem requires tracking the related angles together rather than solving from only one angle."
        ),
    ),
    QuizIssue(
        question=5,
        title="SSA is not a congruence criterion",
        patterns=(
            r"\bSSA\b",
            r"not.*congruence",
            r"cannot.*congruent",
            r"唯一.*不能.*全等",
            r"全等.*判定",
        ),
        chinese=(
            "第 5 题需要明确 SSA 是唯一一种不能得出三角形全等的判定条件，"
            "这是全等三角形中非常常错、也常考的知识点。之后复习时要把 SSS、SAS、ASA、AAS 和 HL "
            "这些可以证明全等的条件与 SSA 区分开。"
        ),
        english=(
            "For question 5, the student needs to remember that SSA is the one condition here that does not prove triangle congruence. "
            "This is a common and frequently tested mistake, so they should clearly distinguish SSA from valid criteria such as SSS, SAS, ASA, AAS, and HL."
        ),
    ),
    QuizIssue(
        question=6,
        title="Parallel line proof and corresponding angles",
        patterns=(
            r"parallel line",
            r"alternate interior",
            r"same[- ]?side interior",
            r"corresponding angle",
            r"DE.*BC",
            r"平行线.*判定",
            r"内错角",
            r"同旁内角",
            r"同位角",
        ),
        chinese=(
            "第 6 题主要考察平行线的判定与性质。第一问可以使用 alternate interior angles are congruent，"
            "或者 same-side interior angles are supplementary 来证明两条线平行。第二问需要明确目标是证明两个不相邻角相等，"
            "常见方法有三种：题目已知信息、平行线产生的三种特殊角关系、或全等三角形。"
            "这道题中应使用第二种方法，先证明 DE 平行于 BC，再通过平行线性质得到 corresponding angles are congruent。"
        ),
        english=(
            "For question 6, the main focus is the converse and properties of parallel lines. In the first part, the student can prove the lines are parallel by using alternate interior angles are congruent or same-side interior angles are supplementary. "
            "In the second part, the goal is to prove two non-adjacent angles are equal. The usual methods are given information, special angle relationships from parallel lines, or congruent triangles. Here, the student should first prove DE is parallel to BC, then use corresponding angles are congruent."
        ),
    ),
    QuizIssue(
        question=7,
        title="Triangle interior angle sum",
        patterns=(
            r"triangle interior angle sum",
            r"interior angle sum",
            r"ABD.*30",
            r"BEC.*122",
            r"DEC.*58",
            r"ECD.*32",
            r"BCD.*64",
            r"ABC.*56",
            r"三角形内角和",
        ),
        chinese=(
            "第 7 题主要考察 triangle interior angle sum，也就是第二讲的内容。"
            "第一问运用三角形内角和等于 180 度，可以直接得出 angle ABD 为 30 度。"
            "第二问先由 angle BEC = 122 度得到 angle DEC = 58 度，再运用三角形内角和得到 angle ECD = 32 度、"
            "angle BCD = 64 度，最后通过三角形内角和得出 angle ABC = 56 度。"
        ),
        english=(
            "For question 7, the main focus is the triangle interior angle sum from Lesson 2. In the first part, using the fact that a triangle's angles add to 180 degrees gives angle ABD as 30 degrees. "
            "In the second part, angle BEC = 122 degrees gives angle DEC = 58 degrees. Then the triangle interior angle sum gives angle ECD = 32 degrees and angle BCD = 64 degrees, and finally angle ABC = 56 degrees."
        ),
    ),
)


def question_numbers_from_note(note: str) -> set[int]:
    stripped = re.sub(
        r"\b(?:physical\s*)?quiz\s*\d+\b|\b(?:first|second|third)\s+physical\s+quiz\b|\bAMC10\b|\bIntro\b",
        " ",
        note.strip(),
        flags=re.IGNORECASE,
    ).strip(" -:,，。；;")
    numbers: set[int] = set()

    for pattern in (
        r"\bq(?:uestion)?\s*(\d{1,2})\b",
        r"第\s*(\d{1,2})\s*题",
    ):
        numbers.update(int(value) for value in re.findall(pattern, stripped, flags=re.IGNORECASE))

    if re.fullmatch(r"(?:\s*(?:q(?:uestion)?\s*)?\d{1,2}\s*[,，、;；]?)+\s*", stripped, flags=re.IGNORECASE):
        numbers.update(int(value) for value in re.findall(r"\d{1,2}", stripped))

    return numbers


def match_amc10_quiz1_issues(note: str) -> list[QuizIssue]:
    number_matches = question_numbers_from_note(note)
    matches: list[QuizIssue] = []
    for issue in AMC10_QUIZ1_ISSUES:
        if issue.question in number_matches or any(
            re.search(pattern, note, flags=re.IGNORECASE) for pattern in issue.patterns
        ):
            matches.append(issue)
    return matches


def question_numbers_for_note(note: str) -> list[int]:
    return [issue.question for issue in match_amc10_quiz1_issues(note)]


def comment_sentences_for_note(note: str, *, language: str = "Chinese") -> list[str]:
    is_chinese = language.lower().startswith("chinese")
    return [
        issue.chinese if is_chinese else issue.english
        for issue in match_amc10_quiz1_issues(note)
    ]


def build_amc10_quiz1_comment(note: str, *, language: str = "Chinese") -> str:
    sentences = comment_sentences_for_note(note, language=language)
    if not sentences:
        return ""
    if language.lower().startswith("chinese"):
        return "".join(sentences)
    return " ".join(sentences)


def main() -> None:
    import argparse

    parser = argparse.ArgumentParser(
        description="Match AMC10 Intro quiz 1 question notes to reusable comment sentences."
    )
    parser.add_argument(
        "note",
        help="Question numbers or teacher shorthand, such as '3,5,7', 'SSA', or 'Exterior Angle Theorem'.",
    )
    parser.add_argument("--language", default="Chinese", choices=["Chinese", "English"])
    args = parser.parse_args()

    matches = match_amc10_quiz1_issues(args.note)
    if not matches:
        print("No matching AMC10 Intro quiz 1 issue found.")
        return

    print("Matched questions:", ", ".join(str(issue.question) for issue in matches))
    print(build_amc10_quiz1_comment(args.note, language=args.language))


if __name__ == "__main__":
    main()
