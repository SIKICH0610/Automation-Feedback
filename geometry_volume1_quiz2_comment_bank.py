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
        question=1,
        title="Triangle Inequality",
        patterns=(
            r"triangle inequality",
            r"三角形不等式",
            r"三条线段.*三角形",
            r"两边之和.*第三边",
        ),
        chinese=(
            "第 1 题需要先判断三条线段能否组成三角形。关键是使用 Triangle Inequality，"
            "任意两边之和都必须大于第三边，不能只检查其中一组。建议列出三组不等式并逐一确认，"
            "这样可以避免遗漏。"
        ),
        english=(
            "For question 1, the student should first decide whether the three segments can form a triangle. "
            "The key is the Triangle Inequality, which requires the sum of any two sides to be greater than "
            "the third side. Checking all three inequalities will help prevent missed conditions."
        ),
    ),
    QuizIssue(
        question=2,
        title="Trigonometric functions and the Pythagorean Theorem",
        patterns=(
            r"trigonometric function",
            r"\bsin\s*A\b",
            r"勾股定理.*sin",
            r"三角函数.*斜边",
            r"斜边长度",
        ),
        chinese=(
            "第 2 题需要更灵活地串联勾股定理和 trigonometric functions。先在已知直角三角形中"
            "利用勾股定理补出缺失边，再根据 opposite side 和 hypotenuse 写出 sin A，最后把同一个"
            "正弦比带入目标三角形求斜边。每一步都要先确认相对于角 A 的对边和斜边。"
        ),
        english=(
            "For question 2, the student should connect the Pythagorean Theorem and trigonometric functions "
            "more flexibly. First find the missing side in the known right triangle, then write sin A using "
            "the opposite side and hypotenuse, and finally use the same sine ratio to find the target "
            "hypotenuse. Identifying the sides relative to angle A is essential."
        ),
    ),
    QuizIssue(
        question=3,
        title="Triangle midsegment and parallelogram proof",
        patterns=(
            r"三角形.*中位线",
            r"中位线.*平行",
            r"证明.*平行四边形",
            r"CEF.*BED",
            r"midsegment.*parallelogram",
            r"opposite sides.*parallel.*congruent",
        ),
        chinese=(
            "第 3 题需要把平行四边形证明中的条件和结论连接得更完整。较快的方法之一是先用"
            "三角形中位线定理得到 DF 平行于 AC，再利用内错角关系推出 CF 平行于 AD，最后根据"
            "两组对边分别平行的定义证明四边形是平行四边形。另一种方法是先证明三角形 CEF "
            "全等于三角形 BED，从而得到 CF 等于 DB，再结合 CF 平行于 DB，利用一组对边平行且"
            "相等的判定方法得出平行四边形的结论。"
        ),
        english=(
            "For question 3, the student should connect each proof condition to the parallelogram conclusion "
            "clearly. One efficient method uses the Triangle Midsegment Theorem to show DF is parallel to AC, "
            "then uses alternate interior angles to show CF is parallel to AD, which proves a parallelogram "
            "by definition. Another method proves triangle CEF congruent to triangle BED, obtains CF equal "
            "to DB, and combines that equality with CF parallel to DB to use the one-pair-of-opposite-sides "
            "criterion."
        ),
    ),
    QuizIssue(
        question=4,
        title="Square and isosceles triangle angle relationships",
        patterns=(
            r"正方形.*等腰三角形",
            r"PCB",
            r"67\.5",
            r"22\.5",
            r"square.*isosceles",
        ),
        chinese=(
            "第 4 题需要综合使用正方形和等腰三角形的性质。先利用等腰三角形的底角关系求出"
            "角 PCB 为 67.5 度，再用角 PCB 减去角 ACB，得到目标角为 22.5 度。做题时应把每个"
            "角度的来源写清楚，避免只写计算结果。"
        ),
        english=(
            "For question 4, the student should combine the properties of a square and an isosceles triangle. "
            "First use the base-angle relationship to find angle PCB as 67.5 degrees, then subtract angle ACB "
            "from angle PCB to obtain 22.5 degrees. Each angle value should be supported by its geometric reason."
        ),
    ),
    QuizIssue(
        question=5,
        title="Trapezoid area and a 30-60-90 triangle",
        patterns=(
            r"梯形.*面积",
            r"梯形.*作高",
            r"1\s*[:：]\s*(?:sqrt|√)\s*3\s*[:：]\s*2",
            r"12\s*(?:sqrt|√)\s*3",
            r"trapezoid.*30-?60-?90",
            r"trapezoid.*area",
        ),
        chinese=(
            "第 5 题综合性较强，看到求梯形面积时，应先作高并寻找高的长度。作高后形成的"
            "直角三角形具有 1 比 √3 比 2 的边长关系，因此可以得到梯形的高为 2√3。最后把"
            "两条底边和高代入梯形面积公式，得到面积为 12√3。建议先标出辅助线和特殊三角形"
            "的边长关系，再进行面积计算。"
        ),
        english=(
            "Question 5 is more comprehensive. Since the problem asks for a trapezoid area, the first step "
            "is to draw the height. The resulting right triangle has the 1 to square root of 3 to 2 side "
            "ratio, so the height is 2 square root of 3. Substituting the two bases and this height into the "
            "trapezoid area formula gives 12 square root of 3."
        ),
    ),
    QuizIssue(
        question=6,
        title="A model and similarity ratios",
        patterns=(
            r"(?<!reverse )\bA\s*-?\s*model\b",
            r"(?<!反向)A\s*模型",
            r"AD\s*/\s*AB.*ED\s*/\s*DC",
            r"ED\s*/\s*DC.*AD\s*/\s*AB",
            r"底边.*ratio",
            r"底边.*比例",
        ),
        chinese=(
            "第 6 题是容易混淆的 A model 题。这个模型会提供三类不同的比例结论，使用时不能"
            "互相替换。本题要求的是三角形底边的 ratio，因此需要直接使用三角形相似得到 "
            "AD 比 AB 等于 ED 比 DC，最后得到所求 ratio 为 2 比 5。列比例前应先明确题目所求"
            "的是哪一组线段。"
        ),
        english=(
            "Question 6 uses the A model, whose three types of proportional conclusions should not be mixed. "
            "Because the question asks for the ratio on the triangle base, the student should use the similarity "
            "relationship AD over AB equals ED over DC, which gives the required ratio of 2 to 5. The requested "
            "segments should be identified before choosing a proportion."
        ),
    ),
    QuizIssue(
        question=7,
        title="Reverse A model and geometric mean relationship",
        patterns=(
            r"reverse\s*A\s*-?\s*model",
            r"反向\s*A\s*模型",
            r"ADC.*ACB.*相似",
            r"AC\s*\^?\s*2.*AD.*AB",
            r"AC²\s*=\s*AD",
        ),
        chinese=(
            "第 7 题考查 reverse A model 带来的特殊边关系。先根据题目条件证明三角形 ADC "
            "相似于三角形 ACB，再由对应边写出 AC 比 AB 等于 AD 比 AC。交叉相乘后得到 "
            "AC² 等于 AD 乘 AB，最后求出 AC 等于 2。这里最重要的是先写清两个相似三角形的"
            "对应顺序，再列出正确比例。"
        ),
        english=(
            "Question 7 focuses on the special side relationship from the reverse A model. First prove triangle "
            "ADC similar to triangle ACB, then use corresponding sides to write AC over AB equals AD over AC. "
            "Cross multiplication gives AC squared equals AD times AB, and therefore AC equals 2. Writing the "
            "triangles in matching order is important."
        ),
    ),
    QuizIssue(
        question=8,
        title="Right Triangle Altitude Theorem",
        patterns=(
            r"Right Triangle Altitude Theorem",
            r"直角三角形.*斜边.*高",
            r"同角.*余角",
            r"三个三角形.*相似",
            r"AN\s*\^?\s*2.*BN.*CN",
            r"AN²\s*=\s*BN",
        ),
        chinese=(
            "第 8 题考查 Right Triangle Altitude Theorem，也是 reverse model 的一种变体。"
            "可以利用同角的余角相等得到所需的角关系，再通过 AA 判断图中的三个三角形两两相似。"
            "由相似三角形的对应边比例可以得到 AN² 等于 BN 乘 CN，之后只需完成相应的代数计算。"
            "复习时应同时理解这个模型的角关系、相似关系和特殊边关系，而不是只记住公式。"
        ),
        english=(
            "Question 8 uses the Right Triangle Altitude Theorem, which is a variation of the reverse model. "
            "Angles complementary to the same angle are congruent, and AA similarity then shows that all three "
            "triangles are mutually similar. Their corresponding-side proportions give AN squared equals BN "
            "times CN, followed by a straightforward algebraic calculation. The student should understand the "
            "angle and similarity reasoning instead of memorizing only the formula."
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

    for start, end in re.findall(
        r"\b(\d{1,2})\s*(?:-|–|—|~|至|到)\s*(\d{1,2})\b",
        stripped,
    ):
        lower, upper = sorted((int(start), int(end)))
        numbers.update(range(lower, upper + 1))

    for pattern in (
        r"\bq(?:uestion)?\s*(\d{1,2})\b",
        r"第\s*(\d{1,2})\s*题",
    ):
        numbers.update(
            int(value) for value in re.findall(pattern, stripped, flags=re.IGNORECASE)
        )

    if re.fullmatch(
        r"(?:\s*(?:q(?:uestion)?\s*)?\d{1,2}\s*[,，、;；–—~至到-]?)+\s*",
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


def build_geometry_volume1_quiz2_comment(
    note: str,
    *,
    language: str = "Chinese",
) -> str:
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
        help="Question numbers or teacher shorthand, such as '1,3,8' or 'Triangle Inequality'.",
    )
    parser.add_argument("--language", default="Chinese", choices=["Chinese", "English"])
    args = parser.parse_args()

    matches = match_geometry_volume1_quiz2_issues(args.note)
    if not matches:
        print("No matching Geometry Volume 1 quiz 2 issue found.")
        return

    print("Matched questions", ", ".join(str(issue.question) for issue in matches))
    print(build_geometry_volume1_quiz2_comment(args.note, language=args.language))


if __name__ == "__main__":
    main()
