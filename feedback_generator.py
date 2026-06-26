from __future__ import annotations

import argparse
import csv
from pathlib import Path

from openpyxl import load_workbook

from openai_api import DEFAULT_OPENAI_MODEL
from feedback_common import (
    DEFAULT_SHEET,
    DEFAULT_WORKBOOK,
    FEEDBACK_TYPE_CHOICES,
    StudentRow,
    find_column,
    iter_student_rows,
    normalize_uid,
    revise_remark_with_gpt,
    set_column_value,
    student_from_worksheet,
    write_column_value,
    write_feedback,
    load_student_row,
)
from feedback_master import FeedbackGenerator, generate_feedback
from feedback_quiz import format_score, score_and_denominator_from_text


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
        "feedback_column",
        "feedback_type",
        "revised_remark",
        "feedback",
    ]
    with output_path.open("w", newline="", encoding="utf-8-sig") as csv_file:
        writer = csv.DictWriter(csv_file, fieldnames=fieldnames)
        writer.writeheader()
        writer.writerows(rows)


def ensure_column(worksheet, headers: list, column_name: str) -> int:
    if column_name in headers:
        return find_column(headers, column_name)
    next_column = len(headers) + 1
    worksheet.cell(1, next_column).value = column_name
    headers.append(column_name)
    return next_column


def calculate_average_from_score_column(worksheet, headers: list, score_column: str) -> tuple[str, int]:
    score_col = find_column(headers, score_column)
    scores: list[float] = []
    denominators: list[float] = []
    for row_number in range(2, worksheet.max_row + 1):
        text = str(worksheet.cell(row_number, score_col).value or "").strip()
        if not text:
            continue
        score, denominator = score_and_denominator_from_text(text)
        if score is None:
            continue
        scores.append(score)
        if denominator is not None:
            denominators.append(denominator)

    if not scores:
        raise ValueError(f"No numeric quiz scores found in column {score_column!r}.")

    average_score = sum(scores) / len(scores)
    denominator = denominators[0] if denominators else None
    average_text = (
        f"{format_score(round(average_score, 1))}/{format_score(denominator)}"
        if denominator is not None
        else format_score(round(average_score, 1))
    )
    return average_text, len(scores)


def build_parser(*, default_feedback_type: str = "comprehensive") -> argparse.ArgumentParser:
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
        help="Write generated text back to the selected feedback column. Preview-only by default.",
    )
    parser.add_argument(
        "--feedback-column",
        default="Feedback",
        help=(
            "Workbook column to write generated parent-facing messages into. "
            "Use 'Feedback' for comprehensive comments, 'Quiz1 Feedback' or 'Quiz2 Feedback' "
            "for quiz-specific comments, and 'Pre-quiz informing' only for shared announcements."
        ),
    )
    parser.add_argument(
        "--calculate-quiz-average",
        action="store_true",
        help="Calculate the class average from a quiz score column and write it to each row.",
    )
    parser.add_argument(
        "--quiz-score-column",
        default="",
        help="Score column used with --calculate-quiz-average, for example 'Quiz2 Score'.",
    )
    parser.add_argument(
        "--quiz-average-column",
        default="",
        help="Average column to write with --calculate-quiz-average, for example 'Quiz2 Average'.",
    )
    parser.add_argument(
        "--only-with-value-column",
        default="",
        help="Only generate rows where this workbook column has a non-empty value.",
    )
    parser.add_argument(
        "--clear-feedback-for-missing-value",
        action="store_true",
        help="When used with --only-with-value-column and --write, clear the feedback column for rows missing that value.",
    )
    parser.add_argument(
        "--feedback-type",
        choices=FEEDBACK_TYPE_CHOICES,
        default=default_feedback_type,
        help=(
            "general = course description + regular classroom feedback; "
            "quiz = course description + quiz feedback; "
            "comprehensive = course description + both quiz and regular feedback."
        ),
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
        "--class-review-file-zh",
        type=Path,
        help="Optional Chinese class review file for Chinese parent messages.",
    )
    parser.add_argument(
        "--class-review-file-en",
        type=Path,
        help="Optional English class review file for English parent messages.",
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


def main(*, default_feedback_type: str = "comprehensive") -> None:
    args = build_parser(default_feedback_type=default_feedback_type).parse_args()
    class_review = args.class_review
    if args.class_review_file:
        class_review = args.class_review_file.read_text(encoding="utf-8").strip()
    class_review_zh = (
        args.class_review_file_zh.read_text(encoding="utf-8").strip()
        if args.class_review_file_zh
        else class_review
    )
    class_review_en = (
        args.class_review_file_en.read_text(encoding="utf-8").strip()
        if args.class_review_file_en
        else class_review
    )

    workbook = load_workbook(args.workbook)
    if args.sheet not in workbook.sheetnames:
        available = ", ".join(workbook.sheetnames)
        raise ValueError(f"Sheet {args.sheet!r} not found. Available sheets: {available}")

    worksheet = workbook[args.sheet]
    headers = [worksheet.cell(1, col).value for col in range(1, worksheet.max_column + 1)]
    if args.calculate_quiz_average:
        if not args.quiz_score_column or not args.quiz_average_column:
            raise ValueError("--calculate-quiz-average requires --quiz-score-column and --quiz-average-column.")
        average_text, score_count = calculate_average_from_score_column(
            worksheet,
            headers,
            args.quiz_score_column,
        )
        average_col = ensure_column(worksheet, headers, args.quiz_average_column)
        for row_number in range(2, worksheet.max_row + 1):
            worksheet.cell(row_number, average_col).value = average_text
        print(
            f"Calculated {args.quiz_average_column}: {average_text} "
            f"from {score_count} scores in {args.quiz_score_column}"
        )
    if args.write:
        find_column(headers, args.feedback_column)
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

    if args.only_with_value_column:
        find_column(headers, args.only_with_value_column)
        all_students = students
        students = [
            student
            for student in all_students
            if str(student.values.get(args.only_with_value_column) or "").strip()
        ]
        if args.write and args.clear_feedback_for_missing_value:
            for student in all_students:
                if not str(student.values.get(args.only_with_value_column) or "").strip():
                    set_column_value(
                        worksheet,
                        headers,
                        student.excel_row,
                        args.feedback_column,
                        "",
                    )

    print(f"Sheet: {args.sheet}")
    print(f"Mode: {'all rows' if args.all else 'single row'}")
    print(f"Feedback type: {args.feedback_type}")
    print(f"Write feedback: {'yes' if args.write else 'no, preview only'}")
    print(f"Feedback column: {args.feedback_column}")
    if args.only_with_value_column:
        print(f"Only rows with value in: {args.only_with_value_column}")
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

        student_class_review = (
            class_review_zh if student.language.lower().startswith("chinese") else class_review_en
        )

        feedback = generate_feedback(
            student,
            class_review=student_class_review,
            use_api=args.use_api,
            model=args.model,
            feedback_type=args.feedback_type,
        )
        print_student_result(student, feedback)

        review_rows.append(
            {
                "row": str(student.excel_row),
                "uid": normalize_uid(student.values.get("uid")),
                "student": student.full_name,
                "status": "skipped_absent" if feedback is None else "generated",
                "feedback_column": args.feedback_column,
                "feedback_type": args.feedback_type,
                "revised_remark": revised_remark,
                "feedback": feedback or "",
            }
        )

        if feedback is None:
            skipped += 1
            continue

        generated += 1
        if args.write:
            set_column_value(worksheet, headers, student.excel_row, args.feedback_column, feedback)

    if args.write or args.write_revised_remark:
        workbook.save(args.workbook)
        print(f"Saved workbook: {args.workbook}")

    if args.review_csv:
        export_review_csv(review_rows, args.review_csv)
        print(f"Saved review CSV: {args.review_csv}")

    print(f"Done. Generated: {generated}. Skipped: {skipped}.")


if __name__ == "__main__":
    main()
