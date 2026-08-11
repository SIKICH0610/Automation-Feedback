from __future__ import annotations

import argparse
from pathlib import Path
from typing import Any

from openpyxl import load_workbook

from database_store import SQLiteFeedbackStore
from feedback_common import DEFAULT_WORKBOOK


PROJECT_DIR = Path(__file__).resolve().parent
DEFAULT_APP_DATA_DIR = PROJECT_DIR / "app_data"
DEFAULT_DATABASE = DEFAULT_APP_DATA_DIR / "feedback.db"
DEFAULT_EXPORT_DIR = PROJECT_DIR / "exports"
DEFAULT_ANNOUNCEMENT_DIR = PROJECT_DIR / "announcements"

REQUIRED_COLUMNS = (
    "classId",
    "className",
    "classTimeDescription",
    "subject",
    "firstName",
    "lastName",
    "学员id",
    "payStatus",
    "是否入班",
)


def normalize_id(value: Any) -> str:
    if value is None:
        return ""
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value).strip()


def read_enrollment_rows(xlsx_path: Path) -> list[dict[str, Any]]:
    workbook = load_workbook(xlsx_path, data_only=True)
    try:
        worksheet = workbook.worksheets[0]
        headers: dict[str, int] = {}
        for column in range(1, worksheet.max_column + 1):
            raw_header = worksheet.cell(1, column).value
            if raw_header is None or not str(raw_header).strip():
                continue
            headers[str(raw_header).strip()] = column

        missing = [name for name in REQUIRED_COLUMNS if name not in headers]
        if missing:
            raise ValueError(
                f"{xlsx_path.name} is missing expected column(s): {', '.join(missing)}"
            )

        rows: list[dict[str, Any]] = []
        for row_number in range(2, worksheet.max_row + 1):
            first_name = worksheet.cell(row_number, headers["firstName"]).value
            if first_name is None or not str(first_name).strip():
                continue
            rows.append(
                {name: worksheet.cell(row_number, column).value for name, column in headers.items()}
            )
        return rows
    finally:
        workbook.close()


def group_by_class(
    rows: list[dict[str, Any]],
    *,
    class_ids: set[str] | None,
) -> dict[str, dict[str, Any]]:
    groups: dict[str, dict[str, Any]] = {}
    seen_uids_by_class: dict[str, set[str]] = {}

    for row in rows:
        if str(row.get("payStatus") or "").strip() != "Paid":
            continue
        if str(row.get("是否入班") or "").strip() != "是":
            continue

        class_id = normalize_id(row.get("classId"))
        if not class_id:
            continue
        if class_ids and class_id not in class_ids:
            continue

        uid = normalize_id(row.get("学员id"))
        if not uid:
            continue

        seen_uids = seen_uids_by_class.setdefault(class_id, set())
        if uid in seen_uids:
            continue
        seen_uids.add(uid)

        group = groups.setdefault(
            class_id,
            {
                "external_class_id": class_id,
                "name": str(row.get("className") or "").strip(),
                "weekly_time": str(row.get("classTimeDescription") or "").strip(),
                "subject": str(row.get("subject") or "").strip(),
                "students": [],
            },
        )
        in_group_value = row.get("In group")
        group["students"].append(
            {
                "uid": uid,
                "first_name": str(row.get("firstName") or "").strip(),
                "last_name": str(row.get("lastName") or "").strip(),
                "in_group": bool(in_group_value) if in_group_value is not None else None,
            }
        )

    return groups


def print_plan(plan: dict[str, Any]) -> None:
    semester_state = "existing" if plan["semester_exists"] else "will be created"
    print(f"Semester: {plan['semester']!r} ({semester_state})")

    for class_plan in plan["classes"]:
        class_state = "existing class" if class_plan["class_exists"] else "NEW class"
        print()
        print(f"[{class_state}] {class_plan['name']} (classId={class_plan['external_class_id']})")
        print(f"    weekly time: {class_plan['weekly_time'] or '(none)'}")
        new_students = [s for s in class_plan["students"] if not s["already_present"]]
        already_present = [s for s in class_plan["students"] if s["already_present"]]
        print(f"    {len(new_students)} new student(s), {len(already_present)} already on roster")
        for student in new_students:
            print(f"      + {student['first_name']} {student['last_name']} (uid={student['uid']})")


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description=(
            "Import a platform enrollment export into the SQLite roster, grouped into a "
            "named semester block. Prints a preview by default; pass --commit to write."
        )
    )
    parser.add_argument("--source-file", type=Path, required=True, help="Enrollment export xlsx.")
    parser.add_argument(
        "--semester",
        required=True,
        help="Semester block name, e.g. '2026-27 School Year'. Created if it doesn't exist yet, "
        "reused otherwise.",
    )
    parser.add_argument(
        "--class-id",
        action="append",
        default=[],
        help="Only import this platform classId. Repeat for multiple classIds. Omit to import "
        "every paid, enrolled class found in the file.",
    )
    parser.add_argument(
        "--commit",
        action="store_true",
        help="Write the changes. Without this flag, only a preview is printed and nothing is saved.",
    )
    parser.add_argument("--database", type=Path, default=DEFAULT_DATABASE)
    parser.add_argument("--workbook", type=Path, default=DEFAULT_WORKBOOK)
    parser.add_argument("--app-data-dir", type=Path, default=DEFAULT_APP_DATA_DIR)
    parser.add_argument("--export-dir", type=Path, default=DEFAULT_EXPORT_DIR)
    parser.add_argument("--announcement-dir", type=Path, default=DEFAULT_ANNOUNCEMENT_DIR)
    return parser


def main() -> None:
    args = build_parser().parse_args()
    rows = read_enrollment_rows(args.source_file)
    class_ids = {normalize_id(value) for value in args.class_id} if args.class_id else None
    groups = group_by_class(rows, class_ids=class_ids)

    if not groups:
        print("No paid, enrolled rows matched the given filters. Nothing to import.")
        return

    store = SQLiteFeedbackStore(
        database_path=args.database,
        source_workbook_path=args.workbook,
        announcement_dir=args.announcement_dir,
        app_data_dir=args.app_data_dir,
        export_dir=args.export_dir,
    )
    plan = store.import_students(
        semester_name=args.semester,
        class_groups=list(groups.values()),
        commit=args.commit,
    )
    print_plan(plan)
    print()
    if args.commit:
        print("Committed to the database.")
        store.ensure_announcement_files()
    else:
        print("Dry run only. Nothing was written. Re-run with --commit to save these changes.")


if __name__ == "__main__":
    main()
