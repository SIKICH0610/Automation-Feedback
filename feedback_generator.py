from __future__ import annotations

import argparse
import csv
from dataclasses import dataclass
from pathlib import Path
from typing import Any

from openpyxl import load_workbook

from openai_api import DEFAULT_OPENAI_MODEL, create_response


PROJECT_DIR = Path(__file__).resolve().parent
DEFAULT_WORKBOOK = PROJECT_DIR / "Geo_TTh_Student_Script_fixed_rows_only.xlsx"
DEFAULT_SHEET = "Geo TTh"
ADDITIONAL_COMMENT_COLUMN = "Additional Comment"


EN_PHRASES = {
    "Note-taking": {
        "Excellent": "took excellent notes",
        "Good": "took good notes",
        "Needs Reminder": "needed reminders to keep notes complete",
    },
    "Listening & Focus": {
        "Very Focused": "stayed very focused",
        "Focused": "stayed focused",
        "Sometimes Distracted": "was sometimes distracted",
        "Sometimes Needs Reminder": "needed occasional reminders to stay focused",
        "Needs Support": "needed support staying focused",
    },
    "Participation": {
        "Very Active": "participated very actively",
        "Active": "participated actively",
        "Helpful": "helped classmates during practice",
        "Asks for Help": "asked for help when needed",
        "Average": "participated at an average level",
        "Needs Encouragement": "would benefit from more encouragement to participate",
    },
    "Practice Speed & Accuracy": {
        "Excellent": "worked through practice problems with excellent speed and accuracy",
        "High Accuracy": "showed high accuracy on practice problems",
        "Fast but Needs Care": "worked quickly and should continue checking details carefully",
        "Good but Rushed": "did well but sometimes rushed",
        "Good": "did well on practice problems",
        "Normal": "worked at a steady pace",
        "Slow but Thoughtful": "worked slowly but thoughtfully",
        "Needs Support": "needed support with practice problems",
    },
    "Proof Logic": {
        "Strong": "showed strong proof logic",
        "Good with Hint": "understood proof logic well with hints",
        "Needs Practice": "needs more practice with proof logic",
    },
    "Calculation": {
        "Good": "calculated accurately",
        "Needs More Care": "should take more care with calculations",
    },
}


ZH_PHRASES = {
    "Note-taking": {
        "Excellent": "笔记非常完整",
        "Good": "笔记比较完整",
        "Needs Reminder": "笔记方面需要提醒",
    },
    "Listening & Focus": {
        "Very Focused": "听课非常专注",
        "Focused": "听课比较专注",
        "Sometimes Distracted": "有时会分心",
        "Sometimes Needs Reminder": "偶尔需要提醒保持专注",
        "Needs Support": "专注度方面需要更多支持",
    },
    "Participation": {
        "Very Active": "课堂参与非常积极",
        "Active": "课堂参与积极",
        "Helpful": "会主动帮助同学",
        "Asks for Help": "遇到问题会主动寻求帮助",
        "Average": "课堂参与度一般",
        "Needs Encouragement": "需要更多鼓励来参与课堂",
    },
    "Practice Speed & Accuracy": {
        "Excellent": "练习速度和准确率都很好",
        "High Accuracy": "练习准确率较高",
        "Fast but Needs Care": "做题速度不错，但需要更仔细",
        "Good but Rushed": "练习表现不错，但有时略急",
        "Good": "练习完成情况不错",
        "Normal": "练习速度正常",
        "Slow but Thoughtful": "做题速度稍慢，但思考比较认真",
        "Needs Support": "练习题方面需要更多支持",
    },
    "Proof Logic": {
        "Strong": "证明逻辑较强",
        "Good with Hint": "在提示下能较好理解证明逻辑",
        "Needs Practice": "证明逻辑还需要更多练习",
    },
    "Calculation": {
        "Good": "计算比较准确",
        "Needs More Care": "计算方面需要更加仔细",
    },
}


OBSERVATION_FIELDS = [
    "Note-taking",
    "Listening & Focus",
    "Participation",
    "Practice Speed & Accuracy",
    "Proof Logic",
    "Calculation",
]


@dataclass
class StudentRow:
    excel_row: int
    values: dict[str, Any]

    @property
    def first_name(self) -> str:
        return str(self.values.get("First Name") or "").strip()

    @property
    def last_name(self) -> str:
        return str(self.values.get("Last Name") or "").strip()

    @property
    def full_name(self) -> str:
        return f"{self.first_name} {self.last_name}".strip()

    @property
    def language(self) -> str:
        return str(self.values.get("Parent Language") or "English").strip()

    @property
    def is_absent(self) -> bool:
        return str(self.values.get("Class Attendance") or "").strip().lower() == "absent"


def normalize_uid(value: Any) -> str:
    if value is None:
        return ""
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value).strip()


def load_student_row(workbook_path: Path, sheet_name: str, excel_row: int) -> tuple[Any, StudentRow]:
    workbook = load_workbook(workbook_path)
    if sheet_name not in workbook.sheetnames:
        available = ", ".join(workbook.sheetnames)
        raise ValueError(f"Sheet {sheet_name!r} not found. Available sheets: {available}")

    worksheet = workbook[sheet_name]
    headers = [worksheet.cell(1, col).value for col in range(1, worksheet.max_column + 1)]
    values = {
        str(header): worksheet.cell(excel_row, col).value
        for col, header in enumerate(headers, start=1)
        if header
    }

    if not values.get("First Name") and not values.get("Last Name"):
        raise ValueError(f"Row {excel_row} does not look like a student row.")

    return workbook, StudentRow(excel_row=excel_row, values=values)


def student_from_worksheet(worksheet: Any, headers: list[Any], excel_row: int) -> StudentRow | None:
    values = {
        str(header): worksheet.cell(excel_row, col).value
        for col, header in enumerate(headers, start=1)
        if header
    }

    if not values.get("First Name") and not values.get("Last Name"):
        return None
    return StudentRow(excel_row=excel_row, values=values)


def find_column(headers: list[Any], column_name: str) -> int:
    try:
        return headers.index(column_name) + 1
    except ValueError as exc:
        raise ValueError(f"The workbook does not have a {column_name!r} column.") from exc


def iter_student_rows(
    worksheet: Any,
    *,
    start_row: int = 2,
    end_row: int | None = None,
) -> list[StudentRow]:
    headers = [worksheet.cell(1, col).value for col in range(1, worksheet.max_column + 1)]
    final_row = end_row or worksheet.max_row
    students: list[StudentRow] = []
    for row_number in range(start_row, final_row + 1):
        student = student_from_worksheet(worksheet, headers, row_number)
        if student:
            students.append(student)
    return students


def phrase_for(field: str, value: Any, language: str) -> str | None:
    if value in (None, ""):
        return None
    text = str(value).strip()
    if text == "Not Observed":
        return None
    phrases = ZH_PHRASES if language.lower().startswith("chinese") else EN_PHRASES
    return phrases.get(field, {}).get(text, text)


def target_language(student: StudentRow) -> str:
    if student.language.lower().startswith("chinese"):
        return "Chinese"
    return "English"


def revise_remark_with_gpt(student: StudentRow, model: str) -> str:
    remark = str(student.values.get("Remark for Student") or "").strip()
    if not remark:
        return ""

    prompt = f"""
You are revising a teacher's private classroom note before it is saved back to the spreadsheet.

Keep the meaning and all important details.
The teacher note may be informal Chinese. Revise it into smooth, concise Chinese.
Do not add a greeting, parent-facing wording, homework, or information not in the note.
Output only the revised Chinese note.

Student: {student.full_name}
Original note:
{remark}
""".strip()
    return create_response(prompt, model=model)


def additional_comment_for_local_message(student: StudentRow, is_chinese: bool) -> str:
    additional_comment = str(student.values.get(ADDITIONAL_COMMENT_COLUMN) or "").strip()
    if not additional_comment:
        return ""
    if is_chinese:
        return f"另外，{additional_comment}"
    if additional_comment.isascii():
        return f"Additional note: {additional_comment}"
    return ""


def personal_feedback_with_gpt(
    student: StudentRow,
    *,
    observations: list[str],
    local_comment: str,
    homework: str | None,
    model: str,
) -> str:
    language = target_language(student)
    remark = str(student.values.get("Remark for Student") or "").strip()
    additional_comment = str(student.values.get(ADDITIONAL_COMMENT_COLUMN) or "").strip()
    homework_text = homework or ""

    prompt = f"""
You are writing the student-specific part of a parent class update.

Write in {language}.
Use a warm, natural teacher voice.
Do not mention "not observed".
Do not include attendance.
Do not include the class material paragraph.
Keep it concise and realistic, like a teacher wrote it after class.
If Additional Comment is provided, translate it if needed and add it naturally to the end of paragraph 2.

Output paragraph 2 as the personal comment.
If homework feedback is provided, add paragraph 3 as homework feedback.
If there is no homework feedback, output only paragraph 2.

Student: {student.full_name}
Teacher remark, usually in Chinese:
{remark}

Additional Comment:
{additional_comment or "None"}

Structured observations:
{join_naturally(observations, student.language) or "None"}

Local fallback draft:
{local_comment}

Homework feedback:
{homework_text or "None"}
""".strip()
    return create_response(prompt, model=model)


def class_review_paragraph(class_review: str, is_chinese: bool) -> str:
    class_review = class_review.strip()
    if class_review:
        return class_review

    if is_chinese:
        return "今天课堂主要围绕本节几何课的核心概念、例题讲解和课堂练习展开。"
    return (
        "In today's class, we reviewed the main geometry ideas for the lesson, "
        "worked through examples, and practiced applying the methods in class."
    )


def join_naturally(items: list[str], language: str) -> str:
    if not items:
        return ""
    if language.lower().startswith("chinese"):
        return "，".join(items)
    if len(items) == 1:
        return items[0]
    if len(items) == 2:
        return f"{items[0]} and {items[1]}"
    return ", ".join(items[:-1]) + f", and {items[-1]}"


def comment_paragraph(student: StudentRow, observations: list[str], is_chinese: bool) -> str:
    name = student.first_name or student.full_name
    remark = str(student.values.get("Remark for Student") or "").strip()
    additional_comment = additional_comment_for_local_message(student, is_chinese)

    if is_chinese:
        if remark:
            base = f"{name}今天课堂表现：{remark}。"
            if additional_comment:
                return f"{base}{additional_comment}。"
            return base
        if observations:
            base = (
                f"{name}今天整体表现稳定，"
                f"{join_naturally(observations, student.language)}。之后可以继续保持好的课堂习惯，"
                "同时在证明逻辑和细节检查上多练习。"
            )
            if additional_comment:
                return f"{base}{additional_comment}。"
            return base
        return f"{name}今天的课堂表现已记录，之后可以继续保持稳定的学习节奏。"

    if observations:
        base = (
            f"{name} had a steady class today. I noticed that {name} "
            f"{join_naturally(observations, student.language)}. The next step is to keep "
            "building proof logic while checking details carefully."
        )
        if additional_comment:
            return f"{base} {additional_comment}."
        return base
    if remark and remark.isascii():
        base = f"For {name}, my main classroom note is: {remark}"
        if additional_comment:
            return f"{base} {additional_comment}."
        return base
    return f"{name}'s classroom notes have been recorded for today's lesson."


def homework_paragraph(student: StudentRow, is_chinese: bool) -> str | None:
    homework = str(student.values.get("Homework Reflection") or "").strip()
    if not homework:
        return None
    if is_chinese:
        return f"作业反馈：{homework}"
    return f"Homework feedback: {homework}"


def generate_feedback(
    student: StudentRow,
    class_review: str = "",
    *,
    use_api: bool = False,
    model: str = DEFAULT_OPENAI_MODEL,
) -> str | None:
    if student.is_absent:
        return None

    language = student.language
    is_chinese = language.lower().startswith("chinese")
    observations = [
        phrase
        for field in OBSERVATION_FIELDS
        if (phrase := phrase_for(field, student.values.get(field), language))
    ]
    class_paragraph = class_review_paragraph(class_review, is_chinese)
    local_comment = comment_paragraph(student, observations, is_chinese)
    homework = homework_paragraph(student, is_chinese)

    if use_api:
        personal_section = personal_feedback_with_gpt(
            student,
            observations=observations,
            local_comment=local_comment,
            homework=homework,
            model=model,
        )
        return "\n\n".join([class_paragraph, personal_section.strip()])

    paragraphs = [class_paragraph, local_comment]
    if homework:
        paragraphs.append(homework)
    return "\n\n".join(paragraphs)


def write_feedback(
    workbook: Any,
    workbook_path: Path,
    sheet_name: str,
    excel_row: int,
    feedback: str,
) -> None:
    worksheet = workbook[sheet_name]
    headers = [worksheet.cell(1, col).value for col in range(1, worksheet.max_column + 1)]
    feedback_col = find_column(headers, "Feedback")

    worksheet.cell(excel_row, feedback_col).value = feedback
    workbook.save(workbook_path)


def write_column_value(
    workbook: Any,
    workbook_path: Path,
    sheet_name: str,
    excel_row: int,
    column_name: str,
    value: str,
) -> None:
    worksheet = workbook[sheet_name]
    headers = [worksheet.cell(1, col).value for col in range(1, worksheet.max_column + 1)]
    column = find_column(headers, column_name)

    worksheet.cell(excel_row, column).value = value
    workbook.save(workbook_path)


def set_column_value(
    worksheet: Any,
    headers: list[Any],
    excel_row: int,
    column_name: str,
    value: str,
) -> None:
    worksheet.cell(excel_row, find_column(headers, column_name)).value = value


def print_student_result(student: StudentRow, feedback: str | None) -> None:
    print("=" * 72)
    print(f"Row: {student.excel_row}")
    print(f"Student: {student.full_name}")
    print(f"UID: {normalize_uid(student.values.get('uid'))}")
    print()
    if feedback is None:
        print("Skipped: student was absent, so no feedback comment was generated.")
    else:
        print(feedback)
    print()


def export_review_csv(rows: list[dict[str, str]], output_path: Path) -> None:
    fieldnames = [
        "row",
        "uid",
        "student",
        "status",
        "revised_remark",
        "feedback",
    ]
    with output_path.open("w", newline="", encoding="utf-8-sig") as csv_file:
        writer = csv.DictWriter(csv_file, fieldnames=fieldnames)
        writer.writeheader()
        writer.writerows(rows)


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="Generate one parent-facing student feedback entry from the Excel tracker."
    )
    parser.add_argument("--workbook", type=Path, default=DEFAULT_WORKBOOK)
    parser.add_argument("--sheet", default=DEFAULT_SHEET)
    parser.add_argument(
        "--row",
        type=int,
        default=2,
        help="Excel row number to generate. Row 2 is the first student row.",
    )
    parser.add_argument(
        "--all",
        action="store_true",
        help="Generate feedback for every student row in the selected sheet.",
    )
    parser.add_argument(
        "--start-row",
        type=int,
        default=2,
        help="First row to use with --all. Defaults to 2.",
    )
    parser.add_argument(
        "--end-row",
        type=int,
        help="Last row to use with --all. Omit to continue through the sheet.",
    )
    parser.add_argument(
        "--write",
        action="store_true",
        help="Write generated text back to the Feedback column. Preview-only by default.",
    )
    parser.add_argument(
        "--review-csv",
        type=Path,
        help="Save generated previews to a CSV for review before or while writing to Excel.",
    )
    parser.add_argument(
        "--class-review",
        default="",
        help="What the class covered today. This becomes the first paragraph.",
    )
    parser.add_argument(
        "--class-review-file",
        type=Path,
        help="Optional text file containing the class review paragraph.",
    )
    parser.add_argument(
        "--use-api",
        action="store_true",
        help="Use the OpenAI API to polish the parent-facing personal comment.",
    )
    parser.add_argument("--model", default=DEFAULT_OPENAI_MODEL)
    parser.add_argument(
        "--revise-remark",
        action="store_true",
        help="Use the OpenAI API to revise the Remark for Student cell text first.",
    )
    parser.add_argument(
        "--write-revised-remark",
        action="store_true",
        help="Save the revised Remark for Student text back to the workbook.",
    )
    return parser


def main() -> None:
    args = build_parser().parse_args()
    class_review = args.class_review
    if args.class_review_file:
        class_review = args.class_review_file.read_text(encoding="utf-8").strip()

    workbook = load_workbook(args.workbook)
    if args.sheet not in workbook.sheetnames:
        available = ", ".join(workbook.sheetnames)
        raise ValueError(f"Sheet {args.sheet!r} not found. Available sheets: {available}")

    worksheet = workbook[args.sheet]
    headers = [worksheet.cell(1, col).value for col in range(1, worksheet.max_column + 1)]
    if args.all:
        students = iter_student_rows(
            worksheet,
            start_row=args.start_row,
            end_row=args.end_row,
        )
    else:
        student = student_from_worksheet(worksheet, headers, args.row)
        if not student:
            raise ValueError(f"Row {args.row} does not look like a student row.")
        students = [student]

    print(f"Sheet: {args.sheet}")
    print(f"Mode: {'all rows' if args.all else 'single row'}")
    print(f"Write feedback: {'yes' if args.write else 'no, preview only'}")
    print()

    generated = 0
    skipped = 0
    review_rows: list[dict[str, str]] = []
    for student in students:
        revised_remark = ""
        if args.revise_remark:
            revised_remark = revise_remark_with_gpt(student, args.model)
            if revised_remark:
                student.values["Remark for Student"] = revised_remark
                print("=" * 72)
                print(f"Row {student.excel_row} revised remark:")
                print(revised_remark)
                print()
                if args.write_revised_remark:
                    set_column_value(
                        worksheet,
                        headers,
                        student.excel_row,
                        "Remark for Student",
                        revised_remark,
                    )

        feedback = generate_feedback(
            student,
            class_review=class_review,
            use_api=args.use_api,
            model=args.model,
        )
        print_student_result(student, feedback)

        review_rows.append(
            {
                "row": str(student.excel_row),
                "uid": normalize_uid(student.values.get("uid")),
                "student": student.full_name,
                "status": "skipped_absent" if feedback is None else "generated",
                "revised_remark": revised_remark,
                "feedback": feedback or "",
            }
        )

        if feedback is None:
            skipped += 1
            continue

        generated += 1
        if args.write:
            set_column_value(worksheet, headers, student.excel_row, "Feedback", feedback)

    if args.write or args.write_revised_remark:
        workbook.save(args.workbook)
        print(f"Saved workbook: {args.workbook}")

    if args.review_csv:
        export_review_csv(review_rows, args.review_csv)
        print(f"Saved review CSV: {args.review_csv}")

    print(f"Done. Generated: {generated}. Skipped: {skipped}.")


if __name__ == "__main__":
    main()
