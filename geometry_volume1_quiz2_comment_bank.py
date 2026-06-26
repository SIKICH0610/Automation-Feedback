from __future__ import annotations

from dataclasses import dataclass
import re


CLASS_NAME = "Geometry Volume 1"
QUIZ_NAME = "Quiz 2"


@dataclass(frozen=True)
class QuizIssue:
    question: int
    title: str
    patterns: tuple[str, ...]
    chinese: str
    english: str


GEOMETRY_VOLUME1_QUIZ2_ISSUES: tuple[QuizIssue, ...] = (
    QuizIssue(
        question=3,
        title="Triangle midsegment and parallelogram proof",
        patterns=(
            r"三角形.*中位线",
            r"中位线.*平行.*底边",
            r"证明.*平行四边形",
            r"对边.*平行.*相等",
            r"midsegment",
            r"opposite sides.*parallel.*equal",
        ),
        chinese=(
            "第 3 题需要更熟练地使用三角形中位线定理。三角形的中位线平行于底边，这个结论可以用来证明平行关系，"
            "进而证明平行四边形。做这类题时，也可以选择从对边平行且相等的角度入手，把证明思路写得更完整。"
        ),
        english=(
            "For question 3, the student should become more comfortable using the Triangle Midsegment Theorem. "
            "The midsegment is parallel to the third side, and this can help prove a parallelogram. "
            "Another valid approach is to show that one pair of opposite sides is both parallel and congruent."
        ),
    ),
    QuizIssue(
        question=5,
        title="Trapezoid area with auxiliary heights",
        patterns=(
            r"梯形.*面积",
            r"辅助线.*高",
            r"AE.*BF.*垂直",
            r"12\\sqrt\{?3\}?",
            r"trapezoid.*area",
            r"draw.*height",
            r"auxiliary.*height",
        ),
        chinese=(
            "第 5 题的关键是先找到梯形的高，因为要求梯形面积时不能只看上下底。这里需要作高为辅助线，使 AE 和 BF 垂直于 DC。"
            "这样可以看出 ABFE 是长方形，所以 EF 等于 AB。再结合两侧的直角三角形，可以得到 DE 和 CF 都等于 2，"
            "并通过勾股定理算出高 AE 和 BF 都是 2√3。最后代入梯形面积公式，面积为 12√3。"
        ),
        english=(
            "For question 5, the key is to find the height of the trapezoid first, because the area formula needs the two bases and the height. "
            "The student should draw auxiliary heights AE and BF perpendicular to DC. Then ABFE is a rectangle, so EF equals AB. "
            "Using the two right triangles on the sides, DE and CF are both 2, and the Pythagorean Theorem gives the height AE and BF as 2√3. "
            "Then the trapezoid area is 12√3."
        ),
    ),
    QuizIssue(
        question=6,
        title="Similarity ratios and transversal theorem",
        patterns=(
            r"相似.*比例",
            r"每一条边",
            r"Lesson\s*14",
            r"transversal",
            r"平行线.*比例.*直接",
            r"parallel.*proportion",
        ),
        chinese=(
            "第 6 题需要更明确相似带来的比例关系是针对对应边的。对于 Lesson 14 里的 theorem，也要注意 transversal 上可以得到对应比例，"
            "但在平行线本身上的线段比例不能直接当作对应比例来使用。做题时需要先判断哪些线段真正对应，再列比例。"
        ),
        english=(
            "For question 6, the student should be clearer that similarity gives ratios between corresponding sides. "
            "For the theorem from Lesson 14, the proportional relationships come from the transversals. "
            "The segments on the parallel lines themselves are not automatically corresponding ratios, so the student should identify the matching sides before setting up a proportion."
        ),
    ),
    QuizIssue(
        question=7,
        title="AA similarity and side proportions",
        patterns=(
            r"AA",
            r"角等.*相似",
            r"哪两个三角形相似",
            r"边.*比例",
            r"similar.*triangles",
            r"side.*proportion",
        ),
        chinese=(
            "第 7 题需要先通过角相等，也就是 AA，相似关系具体判断出哪两个三角形相似。确定三角形对应关系之后，"
            "再写出对应边的比例会更稳。如果一开始没有说明相似的是哪两个三角形，后面的比例关系就容易写错。"
        ),
        english=(
            "For question 7, the student should first use equal angles, or AA, to identify exactly which two triangles are similar. "
            "After the similar triangles and their correspondence are clear, the side proportion will be much easier to write correctly."
        ),
    ),
    QuizIssue(
        question=8,
        title="Right Triangle Altitude Theorem",
        patterns=(
            r"Right Triangle Altitude Theorem",
            r"Altitude Theorem",
            r"高.*定理",
            r"直角三角形.*高",
            r"model",
            r"geometric mean",
        ),
        chinese=(
            "第 8 题需要继续熟练掌握 Right Triangle Altitude Theorem 的证明模型和结论。做这类题时，孩子需要先看出直角三角形中作斜边上的高之后会形成三个相似三角形，"
            "再根据对应边关系写出需要的比例或几何平均关系。"
        ),
        english=(
            "For question 8, the student should keep practicing the proof model and conclusions of the Right Triangle Altitude Theorem. "
            "They need to recognize that the altitude to the hypotenuse creates three similar right triangles, then use the corresponding sides to set up the needed proportion or geometric mean relationship."
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


def match_geometry_volume1_quiz2_issues(note: str) -> list[QuizIssue]:
    number_matches = question_numbers_from_note(note)
    matches: list[QuizIssue] = []
    for issue in GEOMETRY_VOLUME1_QUIZ2_ISSUES:
        if issue.question in number_matches or any(
            re.search(pattern, note, flags=re.IGNORECASE) for pattern in issue.patterns
        ):
            matches.append(issue)
    return matches


def question_numbers_for_note(note: str) -> list[int]:
    return [issue.question for issue in match_geometry_volume1_quiz2_issues(note)]


def comment_sentences_for_note(note: str, *, language: str = "Chinese") -> list[str]:
    is_chinese = language.lower().startswith("chinese")
    return [
        issue.chinese if is_chinese else issue.english
        for issue in match_geometry_volume1_quiz2_issues(note)
    ]


def build_geometry_volume1_quiz2_comment(note: str, *, language: str = "Chinese") -> str:
    sentences = comment_sentences_for_note(note, language=language)
    if not sentences:
        return ""
    if language.lower().startswith("chinese"):
        return "".join(sentences)
    return " ".join(sentences)


def main() -> None:
    import argparse

    parser = argparse.ArgumentParser(
        description="Match Geometry Volume 1 quiz 2 question notes to reusable comment sentences."
    )
    parser.add_argument(
        "note",
        help="Question numbers or teacher shorthand, such as '3,5,8', '梯形面积', or 'Right Triangle Altitude Theorem'.",
    )
    parser.add_argument("--language", default="Chinese", choices=["Chinese", "English"])
    args = parser.parse_args()

    matches = match_geometry_volume1_quiz2_issues(args.note)
    if not matches:
        print("No matching Geometry Volume 1 quiz 2 issue found.")
        return

    print("Matched questions:", ", ".join(str(issue.question) for issue in matches))
    print(build_geometry_volume1_quiz2_comment(args.note, language=args.language))


if __name__ == "__main__":
    main()
