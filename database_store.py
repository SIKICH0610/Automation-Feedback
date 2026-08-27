from __future__ import annotations

from contextlib import contextmanager
from copy import copy
from datetime import date, datetime, time
import hashlib
import json
from pathlib import Path
import re
import shutil
import sqlite3
import threading
from typing import Any, Iterator
from uuid import uuid4

from openpyxl import Workbook, load_workbook
from openpyxl.utils import get_column_letter

from attachment_store import SQLiteAttachmentStore
from quiz_bank_store import SQLiteQuizBankStore


SCHEMA_VERSION = "1"
BOOLEAN_COLUMNS = {
    "Group Chat",
    "Before Class Informing",
    "Class Attendance",
    "Pre-quiz informing",
}
LONG_TEXT_COLUMNS = {
    "Remark for Student",
    "Additional Comment",
    "Homework Reflection",
    "Quiz Feedback",
    "Second Feeback",
    "Feedback",
    "Send Error",
    "Quiz1 Feedback",
    "Quiz1 Mistake",
    "Quiz2 Feedback",
    "Quiz2 Mistake",
}
COLUMN_OPTIONS = {
    "Group Chat": ["", "TRUE", "FALSE"],
    "Before Class Informing": ["", "TRUE", "FALSE"],
    "Pre-quiz informing": ["", "TRUE", "FALSE"],
    "Parent Language": ["", "Chinese", "English"],
    "Preferred Channel": ["", "wecom", "whatsapp"],
    "WhatsApp Target Type": ["", "group_search", "phone"],
    "Class Attendance": ["", "TRUE", "FALSE", "Present", "Absent"],
    "Note-taking": ["", "Excellent", "Good", "Needs Reminder", "Not Observed"],
    "Listening & Focus": [
        "",
        "Very Focused",
        "Focused",
        "Sometimes Distracted",
        "Sometimes Needs Reminder",
        "Needs Support",
        "Not Observed",
    ],
    "Participation": [
        "",
        "Very Active",
        "Active",
        "Helpful",
        "Asks for Help",
        "Average",
        "Needs Encouragement",
        "Not Observed",
    ],
    "Practice Speed & Accuracy": [
        "",
        "Excellent",
        "High Accuracy",
        "Fast but Needs Care",
        "Good but Rushed",
        "Good",
        "Normal",
        "Slow but Thoughtful",
        "Needs Support",
        "Not Observed",
    ],
    "Proof Logic": ["", "Strong", "Good with Hint", "Needs Practice", "Not Observed"],
    "Calculation": ["", "Good", "Needs More Care", "Not Observed"],
}
STUDENT_COLUMNS = {
    "First Name",
    "Last Name",
    "uid",
    "Group Chat",
    "Parent Language",
    "Preferred Channel",
    "WhatsApp Phone",
    "WhatsApp Search Key",
    "WhatsApp Target Type",
}
AUDIT_COLUMNS = {"Send Status", "Send Error", "Last Attempt"}

REPORT_COLUMNS = (
    "Name",
    "Student ID",
    "电话号码",
    "是否有群",
    "是否发开课提醒",
    "是否发课后反馈",
    "第一节课反馈",
    "第一次quiz反馈",
    "第二次quiz反馈",
)
# The three feedback columns hold whole paragraphs, so they get much wider cells.
REPORT_COLUMN_WIDTHS = (18, 13, 16, 11, 16, 16, 60, 60, 60)

# Sentinel meaning "the classes that predate semesters and have no semester_id" --
# the block the frontend labels "Other Classes". Distinct from None, which means
# every class regardless of semester.
UNASSIGNED_SEMESTER = "__unassigned__"


def _report_yes_no(value: Any) -> str:
    if isinstance(value, str):
        normalized = value.strip().lower()
        if normalized in {"true", "是"}:
            return "是"
        return "否"
    return "是" if bool(value) else "否"


def _report_row(values: dict[str, Any]) -> tuple[Any, ...]:
    name = f"{values.get('First Name') or ''} {values.get('Last Name') or ''}".strip()
    sent_after_class = str(values.get("Send Status") or "").strip().lower() == "pasted"
    return (
        name,
        values.get("uid") or "",
        values.get("WhatsApp Phone") or "",
        _report_yes_no(values.get("Group Chat")),
        _report_yes_no(values.get("Before Class Informing")),
        "是" if sent_after_class else "否",
        values.get("Feedback") or "",
        values.get("Quiz1 Feedback") or "",
        values.get("Quiz2 Feedback") or "",
    )


class StoreError(Exception):
    pass


def json_value(value: Any) -> Any:
    if isinstance(value, (datetime, date, time)):
        return value.isoformat()
    if value is None or isinstance(value, (str, int, float, bool)):
        return value
    return str(value)


def announcement_filename(sheet_name: str) -> str:
    stem = re.sub(r"[^A-Za-z0-9._-]+", "_", sheet_name).strip("._")
    if not stem:
        digest = hashlib.sha1(sheet_name.encode("utf-8")).hexdigest()[:8]
        stem = f"class_{digest}"
    return f"{stem}.txt"


INVALID_SHEET_CHARS = re.compile(r"[\\/*?:\[\]]")
MAX_SHEET_NAME_LENGTH = 31


def safe_sheet_name(raw_name: str, weekly_time: str, existing: set[str]) -> str:
    """Derive a valid, unique Excel worksheet title from a platform class name.

    Excel sheet titles cannot contain [ ] : * ? / \\ and are capped at 31 characters,
    but platform class names routinely violate both (e.g. "[In Person Cupertino DA]
    2026-27 School Year Geometry V1+V2 Honors Sun"). The boilerplate venue-bracket
    prefix and "YYYY-YY School Year" segment carry no distinguishing information, so
    they're stripped first; the weekly-time's leading day token is kept as an explicit
    suffix since it's usually what actually distinguishes sibling sections.
    """
    stripped = re.sub(r"^\[[^\]]*\]\s*", "", raw_name)
    stripped = re.sub(r"^\d{4}-\d{2}\s+School Year\s+", "", stripped)
    cleaned = INVALID_SHEET_CHARS.sub("", stripped)
    cleaned = re.sub(r"\s+", " ", cleaned).strip() or INVALID_SHEET_CHARS.sub("", raw_name).strip()

    day = INVALID_SHEET_CHARS.sub("", (weekly_time.split()[0] if weekly_time else "")).strip()
    base = cleaned
    if day and base.endswith(day):
        base = base[: -len(day)].rstrip()

    suffix = f" {day}" if day else ""
    budget = max(MAX_SHEET_NAME_LENGTH - len(suffix), 1)
    base = base[:budget].rstrip()
    candidate = f"{base}{suffix}".strip()[:MAX_SHEET_NAME_LENGTH] or (day or "Class")

    original, attempt = candidate, 2
    while candidate in existing:
        marker = f" ({attempt})"
        candidate = f"{original[: MAX_SHEET_NAME_LENGTH - len(marker)]}{marker}"
        attempt += 1
    return candidate


def column_group(header: str) -> str:
    lowered = header.lower().replace(" ", "")
    if header in STUDENT_COLUMNS:
        return "student"
    if header in AUDIT_COLUMNS:
        return "audit"
    if lowered.startswith("quiz1"):
        return "quiz1"
    if lowered.startswith("quiz2"):
        return "quiz2"
    return "general"


def coerce_cell_value(header: str, value: Any) -> Any:
    if value is None or value == "":
        return None
    if header == "uid":
        return str(value).strip()
    if header in BOOLEAN_COLUMNS and isinstance(value, str):
        normalized = value.strip().lower()
        if normalized == "true":
            return True
        if normalized == "false":
            return False
    return value


def header_cells(worksheet: Any) -> list[tuple[int, str]]:
    headers: list[tuple[int, str]] = []
    seen: set[str] = set()
    for column in range(1, worksheet.max_column + 1):
        raw_header = worksheet.cell(1, column).value
        if raw_header is None or not str(raw_header).strip():
            continue
        header = str(raw_header).strip()
        if header in seen:
            raise StoreError(f"Duplicate column header {header!r} in {worksheet.title!r}.")
        seen.add(header)
        headers.append((column, header))
    return headers


class SQLiteFeedbackStore:
    def __init__(
        self,
        *,
        database_path: Path,
        source_workbook_path: Path,
        announcement_dir: Path,
        app_data_dir: Path,
        export_dir: Path,
    ) -> None:
        self.database_path = database_path.resolve()
        self.source_workbook_path = source_workbook_path.resolve()
        self.announcement_dir = announcement_dir.resolve()
        self.app_data_dir = app_data_dir.resolve()
        self.export_dir = export_dir.resolve()
        self.template_path = self.app_data_dir / "workbook_template.xlsx"
        self.workbook_path = self.app_data_dir / "runtime_workbook.xlsx"
        self.lock = threading.RLock()

        self.app_data_dir.mkdir(parents=True, exist_ok=True)
        self.database_path.parent.mkdir(parents=True, exist_ok=True)
        self._initialize_database()
        self.quiz_banks = SQLiteQuizBankStore(
            self.database_path,
            lock=self.lock,
        )
        self.attachments = SQLiteAttachmentStore(
            self.database_path,
            lock=self.lock,
        )
        self.ensure_announcement_files()

    @contextmanager
    def _connect(self) -> Iterator[sqlite3.Connection]:
        connection = sqlite3.connect(self.database_path, timeout=30)
        connection.row_factory = sqlite3.Row
        connection.execute("PRAGMA foreign_keys = ON")
        try:
            yield connection
        except Exception:
            connection.rollback()
            raise
        else:
            connection.commit()
        finally:
            connection.close()

    def _initialize_database(self) -> None:
        new_database = not self.database_path.exists()
        with self._connect() as connection:
            connection.executescript(
                """
                CREATE TABLE IF NOT EXISTS meta (
                    key TEXT PRIMARY KEY,
                    value TEXT NOT NULL
                );

                CREATE TABLE IF NOT EXISTS semesters (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    name TEXT NOT NULL UNIQUE,
                    position INTEGER NOT NULL
                );

                CREATE TABLE IF NOT EXISTS classes (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    name TEXT NOT NULL UNIQUE,
                    position INTEGER NOT NULL
                );

                CREATE TABLE IF NOT EXISTS columns (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    class_id INTEGER NOT NULL,
                    name TEXT NOT NULL,
                    position INTEGER NOT NULL,
                    width REAL NOT NULL DEFAULT 13,
                    UNIQUE(class_id, name),
                    FOREIGN KEY(class_id) REFERENCES classes(id) ON DELETE CASCADE
                );

                CREATE TABLE IF NOT EXISTS students (
                    id TEXT PRIMARY KEY,
                    class_id INTEGER NOT NULL,
                    position INTEGER NOT NULL,
                    values_json TEXT NOT NULL,
                    created_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
                    updated_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
                    FOREIGN KEY(class_id) REFERENCES classes(id) ON DELETE CASCADE
                );

                CREATE INDEX IF NOT EXISTS students_class_position
                ON students(class_id, position);
                """
            )
            self._migrate_classes_table(connection)
            class_count = connection.execute("SELECT COUNT(*) FROM classes").fetchone()[0]
            if class_count == 0:
                if not self.source_workbook_path.exists():
                    raise FileNotFoundError(
                        "No SQLite data exists yet and the import workbook was not found at "
                        f"{self.source_workbook_path}"
                    )
                self._import_workbook(connection)
                self._create_blank_template()
            elif not self.template_path.exists():
                self._create_template_from_database(connection)

            connection.execute(
                "INSERT OR REPLACE INTO meta(key, value) VALUES('schema_version', ?)",
                (SCHEMA_VERSION,),
            )
            if new_database:
                connection.execute(
                    "INSERT OR REPLACE INTO meta(key, value) VALUES('import_source', ?)",
                    (str(self.source_workbook_path),),
                )

    @staticmethod
    def _table_columns(connection: sqlite3.Connection, table_name: str) -> set[str]:
        rows = connection.execute(f"PRAGMA table_info({table_name})").fetchall()
        return {str(row["name"]) for row in rows}

    def _migrate_classes_table(self, connection: sqlite3.Connection) -> None:
        existing = self._table_columns(connection, "classes")
        additions = {
            "semester_id": "INTEGER REFERENCES semesters(id)",
            "external_class_id": "TEXT",
            "weekly_time": "TEXT",
            "subject": "TEXT",
        }
        for column_name, definition in additions.items():
            if column_name not in existing:
                connection.execute(f"ALTER TABLE classes ADD COLUMN {column_name} {definition}")

    def _import_workbook(self, connection: sqlite3.Connection) -> None:
        workbook = load_workbook(self.source_workbook_path, data_only=False)
        try:
            for sheet_position, sheet_name in enumerate(workbook.sheetnames):
                worksheet = workbook[sheet_name]
                cursor = connection.execute(
                    "INSERT INTO classes(name, position) VALUES(?, ?)",
                    (sheet_name, sheet_position),
                )
                class_id = int(cursor.lastrowid)
                headers = header_cells(worksheet)
                for position, (column_index, header) in enumerate(headers):
                    width = worksheet.column_dimensions[get_column_letter(column_index)].width or 13
                    connection.execute(
                        "INSERT INTO columns(class_id, name, position, width) VALUES(?, ?, ?, ?)",
                        (class_id, header, position, float(width)),
                    )

                student_position = 0
                for row_number in range(2, worksheet.max_row + 1):
                    values = {
                        header: json_value(worksheet.cell(row_number, column_index).value)
                        for column_index, header in headers
                    }
                    if not any(value not in (None, "") for value in values.values()):
                        continue
                    connection.execute(
                        """
                        INSERT INTO students(id, class_id, position, values_json)
                        VALUES(?, ?, ?, ?)
                        """,
                        (
                            str(uuid4()),
                            class_id,
                            student_position,
                            json.dumps(values, ensure_ascii=False),
                        ),
                    )
                    student_position += 1
        finally:
            workbook.close()

    def _create_blank_template(self) -> None:
        workbook = load_workbook(self.source_workbook_path)
        try:
            for worksheet in workbook.worksheets:
                for row_number in range(2, worksheet.max_row + 1):
                    for column_number in range(1, worksheet.max_column + 1):
                        worksheet.cell(row_number, column_number).value = None
            temporary = self.template_path.with_suffix(".tmp.xlsx")
            workbook.save(temporary)
            temporary.replace(self.template_path)
        finally:
            workbook.close()

    def _create_template_from_database(self, connection: sqlite3.Connection) -> None:
        workbook = Workbook()
        try:
            class_rows = connection.execute(
                "SELECT id, name FROM classes ORDER BY position"
            ).fetchall()
            for index, class_row in enumerate(class_rows):
                worksheet = workbook.active if index == 0 else workbook.create_sheet()
                worksheet.title = str(class_row["name"])
                columns = self._columns(connection, int(class_row["id"]))
                for column_index, column in enumerate(columns, start=1):
                    worksheet.cell(1, column_index).value = str(column["name"])
                    worksheet.column_dimensions[get_column_letter(column_index)].width = float(
                        column["width"]
                    )
                worksheet.freeze_panes = "A2"
                if columns:
                    worksheet.auto_filter.ref = f"A1:{get_column_letter(len(columns))}1"

            temporary = self.template_path.with_suffix(".tmp.xlsx")
            workbook.save(temporary)
            temporary.replace(self.template_path)
        finally:
            workbook.close()

    def _class_row(self, connection: sqlite3.Connection, sheet_name: str) -> sqlite3.Row:
        row = connection.execute(
            "SELECT id, name, position FROM classes WHERE name = ?",
            (sheet_name,),
        ).fetchone()
        if row is None:
            raise StoreError(f"Sheet {sheet_name!r} was not found.")
        return row

    def sheet_names(self) -> list[str]:
        with self.lock, self._connect() as connection:
            rows = connection.execute(
                "SELECT name FROM classes ORDER BY position"
            ).fetchall()
            return [str(row["name"]) for row in rows]

    def list_semesters(self) -> list[dict[str, Any]]:
        with self.lock, self._connect() as connection:
            rows = connection.execute(
                "SELECT id, name FROM semesters ORDER BY position"
            ).fetchall()
            return [{"id": int(row["id"]), "name": str(row["name"])} for row in rows]

    def delete_semester(self, semester_name: str, *, commit: bool = False) -> dict[str, Any]:
        """Permanently delete a semester along with every class and student in it.

        Classes and students cascade-delete automatically (both reference their parent
        via ON DELETE CASCADE). With commit=False (the default), nothing is deleted and
        the same plan (which classes, how many students) is returned so callers can show
        it for confirmation first. A timestamped backup of the whole database file is
        made automatically before the delete runs.
        """
        with self.lock:
            with self._connect() as connection:
                semester_row = connection.execute(
                    "SELECT id FROM semesters WHERE name = ?", (semester_name,)
                ).fetchone()
                if semester_row is None:
                    raise StoreError(f"Semester {semester_name!r} was not found.")
                semester_id = int(semester_row["id"])
                class_rows = connection.execute(
                    "SELECT id, name FROM classes WHERE semester_id = ? ORDER BY position",
                    (semester_id,),
                ).fetchall()
                plan = {
                    "semester": semester_name,
                    "classes": [str(row["name"]) for row in class_rows],
                    "student_count": sum(
                        connection.execute(
                            "SELECT COUNT(*) FROM students WHERE class_id = ?", (int(row["id"]),)
                        ).fetchone()[0]
                        for row in class_rows
                    ),
                }
                if not commit:
                    connection.rollback()
                    return plan

            self._backup_database_file()

            with self._connect() as connection:
                connection.execute("DELETE FROM classes WHERE semester_id = ?", (semester_id,))
                connection.execute("DELETE FROM semesters WHERE id = ?", (semester_id,))

            return plan

    def sheet_groups(self) -> list[dict[str, Any]]:
        """Class sheet names grouped into semester blocks, in position order.

        Classes imported before semesters existed (semester_id is NULL) are grouped
        together under an "Other Classes" block, listed first.
        """
        with self.lock, self._connect() as connection:
            semester_names = {
                int(row["id"]): str(row["name"])
                for row in connection.execute("SELECT id, name FROM semesters").fetchall()
            }
            class_rows = connection.execute(
                """
                SELECT c.name, c.position, c.semester_id, COUNT(s.id) AS student_count
                FROM classes c
                LEFT JOIN students s ON s.class_id = c.id
                GROUP BY c.id
                ORDER BY c.position
                """
            ).fetchall()

        groups: list[dict[str, Any]] = []
        group_index: dict[int | None, int] = {}
        for row in class_rows:
            raw_semester_id = row["semester_id"]
            key = int(raw_semester_id) if raw_semester_id is not None else None
            if key not in group_index:
                group_index[key] = len(groups)
                label = semester_names.get(key, "") if key is not None else "Other Classes"
                groups.append(
                    {
                        "semester": label,
                        "is_semester": key is not None,
                        "sheets": [],
                        "student_counts": {},
                    }
                )
            group = groups[group_index[key]]
            sheet_name = str(row["name"])
            group["sheets"].append(sheet_name)
            group["student_counts"][sheet_name] = int(row["student_count"])
        return groups

    def _default_column_template(self, connection: sqlite3.Connection) -> list[tuple[str, float]]:
        row = connection.execute("SELECT id FROM classes ORDER BY position LIMIT 1").fetchone()
        if row is None:
            return [
                ("First Name", 13.0),
                ("Last Name", 13.0),
                ("uid", 13.0),
                ("Group Chat", 13.0),
                ("Parent Language", 13.0),
                ("Feedback", 40.0),
            ]
        return [(str(c["name"]), float(c["width"])) for c in self._columns(connection, int(row["id"]))]

    def _plan_import(
        self,
        connection: sqlite3.Connection,
        *,
        semester_name: str,
        class_groups: list[dict[str, Any]],
        overwrite: bool = False,
    ) -> dict[str, Any]:
        semester_row = connection.execute(
            "SELECT id FROM semesters WHERE name = ?", (semester_name,)
        ).fetchone()
        semester_exists = semester_row is not None
        plan: dict[str, Any] = {
            "semester": semester_name,
            "semester_exists": semester_exists,
            "overwrite": overwrite,
            "classes": [],
        }
        for group in class_groups:
            class_row = None
            if semester_exists:
                class_row = connection.execute(
                    "SELECT id FROM classes WHERE semester_id = ? AND external_class_id = ?",
                    (int(semester_row["id"]), group["external_class_id"]),
                ).fetchone()

            incoming_uids = {
                str(student["uid"]).strip() for student in group["students"] if str(student["uid"]).strip()
            }

            existing_uids: set[str] = set()
            students_to_remove: list[dict[str, str]] = []
            if class_row is not None:
                for student in self._students(connection, int(class_row["id"])):
                    values = json.loads(str(student["values_json"]))
                    uid = str(values.get("uid") or "").strip()
                    if uid:
                        existing_uids.add(uid)
                    if overwrite and uid and uid not in incoming_uids:
                        students_to_remove.append(
                            {
                                "uid": uid,
                                "first_name": str(values.get("First Name") or ""),
                                "last_name": str(values.get("Last Name") or ""),
                            }
                        )

            class_plan = {
                "name": group["name"],
                "external_class_id": group["external_class_id"],
                "weekly_time": group.get("weekly_time", ""),
                "subject": group.get("subject", ""),
                "class_exists": class_row is not None,
                "students": [
                    {
                        "uid": str(student["uid"]).strip(),
                        "first_name": student["first_name"],
                        "last_name": student["last_name"],
                        "already_present": str(student["uid"]).strip() in existing_uids,
                    }
                    for student in group["students"]
                ],
                "students_to_remove": students_to_remove,
            }
            plan["classes"].append(class_plan)
        return plan

    def _backup_database_file(self) -> Path:
        timestamp = datetime.now().strftime("%Y%m%d-%H%M%S")
        backup_path = self.database_path.with_name(
            f"{self.database_path.stem}.before-overwrite-{timestamp}{self.database_path.suffix}"
        )
        shutil.copy2(self.database_path, backup_path)
        return backup_path

    def import_students(
        self,
        *,
        semester_name: str,
        class_groups: list[dict[str, Any]],
        commit: bool = False,
        overwrite: bool = False,
    ) -> dict[str, Any]:
        """Import enrollment data grouped by platform class, under a named semester block.

        Each entry in class_groups is expected to look like:
        {
            "external_class_id": str, "name": str, "weekly_time": str, "subject": str,
            "students": [{"uid": str, "first_name": str, "last_name": str, "in_group": bool | None}, ...],
        }
        A semester is created if semester_name doesn't exist yet, otherwise reused. Within a
        semester, a class is matched by external_class_id and created if missing.

        With overwrite=False (the default), students are matched by uid within their class;
        existing students are left completely untouched and only new uids are inserted, so
        re-running an import (or importing one new class at a time) is safe.

        With overwrite=True, an existing class's roster is synced to exactly match the
        incoming group: a student already on the roster whose uid is not in the incoming
        group is permanently deleted, a student present in both gets First Name, Last Name,
        and Group Chat (when provided) refreshed from the incoming data, and a student only
        in the incoming group is inserted. Every other field for a retained student (feedback,
        quiz scores, teacher remarks, attendance notes, etc.) is left untouched -- overwrite
        only syncs roster membership and identity fields, not accumulated teaching data. A
        timestamped backup of the whole database file is made automatically before any
        overwrite delete runs.

        With commit=False (the default) nothing is written; the same plan that would be applied
        is returned so callers can preview it first.
        """
        with self.lock:
            with self._connect() as connection:
                plan = self._plan_import(
                    connection,
                    semester_name=semester_name,
                    class_groups=class_groups,
                    overwrite=overwrite,
                )
                if not commit:
                    connection.rollback()
                    return plan

            if overwrite:
                self._backup_database_file()

            with self._connect() as connection:
                semester_row = connection.execute(
                    "SELECT id FROM semesters WHERE name = ?", (semester_name,)
                ).fetchone()
                if semester_row is None:
                    position = connection.execute(
                        "SELECT COALESCE(MAX(position), -1) + 1 FROM semesters"
                    ).fetchone()[0]
                    cursor = connection.execute(
                        "INSERT INTO semesters(name, position) VALUES (?, ?)",
                        (semester_name, position),
                    )
                    semester_id = int(cursor.lastrowid)
                else:
                    semester_id = int(semester_row["id"])

                template_columns = self._default_column_template(connection)
                created_new_class = False
                existing_names = {
                    str(row["name"]) for row in connection.execute("SELECT name FROM classes").fetchall()
                }

                for group in class_groups:
                    class_row = connection.execute(
                        "SELECT id FROM classes WHERE semester_id = ? AND external_class_id = ?",
                        (semester_id, group["external_class_id"]),
                    ).fetchone()
                    is_new_class = class_row is None
                    if is_new_class:
                        created_new_class = True
                        position = connection.execute(
                            "SELECT COALESCE(MAX(position), -1) + 1 FROM classes"
                        ).fetchone()[0]
                        sheet_name = safe_sheet_name(
                            group["name"], group.get("weekly_time", ""), existing_names
                        )
                        existing_names.add(sheet_name)
                        cursor = connection.execute(
                            """
                            INSERT INTO classes(name, position, semester_id, external_class_id, weekly_time, subject)
                            VALUES (?, ?, ?, ?, ?, ?)
                            """,
                            (
                                sheet_name,
                                position,
                                semester_id,
                                group["external_class_id"],
                                group.get("weekly_time", ""),
                                group.get("subject", ""),
                            ),
                        )
                        class_id = int(cursor.lastrowid)
                        for column_position, (column_name, width) in enumerate(template_columns):
                            connection.execute(
                                "INSERT INTO columns(class_id, name, position, width) VALUES (?, ?, ?, ?)",
                                (class_id, column_name, column_position, width),
                            )
                    else:
                        class_id = int(class_row["id"])

                    incoming_uids = {
                        str(student["uid"]).strip()
                        for student in group["students"]
                        if str(student["uid"]).strip()
                    }

                    existing_by_uid: dict[str, tuple[str, dict[str, Any]]] = {}
                    for student in self._students(connection, class_id):
                        values = json.loads(str(student["values_json"]))
                        uid = str(values.get("uid") or "").strip()
                        if uid:
                            existing_by_uid[uid] = (str(student["id"]), values)

                    if overwrite and not is_new_class:
                        for uid in list(existing_by_uid):
                            if uid not in incoming_uids:
                                student_id, _ = existing_by_uid.pop(uid)
                                connection.execute("DELETE FROM students WHERE id = ?", (student_id,))

                    next_position = connection.execute(
                        "SELECT COALESCE(MAX(position), -1) + 1 FROM students WHERE class_id = ?",
                        (class_id,),
                    ).fetchone()[0]
                    for student in group["students"]:
                        uid = str(student["uid"]).strip()
                        if not uid:
                            continue
                        if uid in existing_by_uid:
                            if overwrite:
                                student_id, values = existing_by_uid[uid]
                                values["First Name"] = student["first_name"]
                                values["Last Name"] = student["last_name"]
                                if student.get("in_group") is not None:
                                    values["Group Chat"] = bool(student["in_group"])
                                connection.execute(
                                    """
                                    UPDATE students
                                    SET values_json = ?, updated_at = CURRENT_TIMESTAMP
                                    WHERE id = ?
                                    """,
                                    (json.dumps(values, ensure_ascii=False), student_id),
                                )
                            continue
                        new_values: dict[str, Any] = {
                            "First Name": student["first_name"],
                            "Last Name": student["last_name"],
                            "uid": uid,
                        }
                        if student.get("in_group") is not None:
                            new_values["Group Chat"] = bool(student["in_group"])
                        new_student_id = str(uuid4())
                        connection.execute(
                            "INSERT INTO students(id, class_id, position, values_json) VALUES (?, ?, ?, ?)",
                            (new_student_id, class_id, next_position, json.dumps(new_values, ensure_ascii=False)),
                        )
                        existing_by_uid[uid] = (new_student_id, new_values)
                        next_position += 1

                if created_new_class:
                    self._create_template_from_database(connection)

            return plan

    def announcement_path(self, sheet_name: str) -> Path:
        if sheet_name not in self.sheet_names():
            raise StoreError(f"Sheet {sheet_name!r} was not found.")
        return self.announcement_dir / announcement_filename(sheet_name)

    def ensure_announcement_files(self) -> None:
        self.announcement_dir.mkdir(parents=True, exist_ok=True)
        for sheet_name in self.sheet_names():
            path = self.announcement_dir / announcement_filename(sheet_name)
            path.touch(exist_ok=True)

    def read_announcement(self, sheet_name: str) -> tuple[str, Path]:
        path = self.announcement_path(sheet_name)
        path.parent.mkdir(parents=True, exist_ok=True)
        path.touch(exist_ok=True)
        return path.read_text(encoding="utf-8"), path

    def write_announcement(self, sheet_name: str, text: str) -> Path:
        path = self.announcement_path(sheet_name)
        path.parent.mkdir(parents=True, exist_ok=True)
        path.write_text(text.strip(), encoding="utf-8")
        return path

    @staticmethod
    def _columns(connection: sqlite3.Connection, class_id: int) -> list[sqlite3.Row]:
        return connection.execute(
            """
            SELECT name, position, width
            FROM columns
            WHERE class_id = ?
            ORDER BY position
            """,
            (class_id,),
        ).fetchall()

    @staticmethod
    def _students(connection: sqlite3.Connection, class_id: int) -> list[sqlite3.Row]:
        return connection.execute(
            """
            SELECT id, position, values_json
            FROM students
            WHERE class_id = ?
            ORDER BY position
            """,
            (class_id,),
        ).fetchall()

    def load_sheet(self, sheet_name: str) -> dict[str, Any]:
        with self.lock, self._connect() as connection:
            class_row = self._class_row(connection, sheet_name)
            class_id = int(class_row["id"])
            columns: list[dict[str, Any]] = []
            for column in self._columns(connection, class_id):
                header = str(column["name"])
                options = COLUMN_OPTIONS.get(header)
                columns.append(
                    {
                        "key": header,
                        "label": header,
                        "group": column_group(header),
                        "kind": "select" if options else (
                            "long_text" if header in LONG_TEXT_COLUMNS else "text"
                        ),
                        "options": options or [],
                        "width": min(max(round(float(column["width"]) * 7.2), 110), 360),
                    }
                )

            rows: list[dict[str, Any]] = []
            for position, student in enumerate(self._students(connection, class_id)):
                values = json.loads(str(student["values_json"]))
                rows.append(
                    {
                        "student_id": str(student["id"]),
                        "excel_row": position + 2,
                        "values": values,
                    }
                )

        announcement, announcement_path = self.read_announcement(sheet_name)
        return {
            "sheet": sheet_name,
            "columns": columns,
            "rows": rows,
            "announcement": announcement,
            "announcement_path": str(announcement_path),
            "attachments": self.list_attachments(sheet_name),
        }

    def search_students(self, query: str, limit: int = 50) -> list[dict[str, Any]]:
        normalized_query = query.strip().casefold()
        if not normalized_query:
            return []

        matches: list[tuple[tuple[int, int, int], dict[str, Any]]] = []
        with self.lock, self._connect() as connection:
            classes = connection.execute(
                "SELECT id, name, position FROM classes ORDER BY position"
            ).fetchall()
            for class_row in classes:
                class_id = int(class_row["id"])
                for student in self._students(connection, class_id):
                    values = json.loads(str(student["values_json"]))
                    first_name = str(values.get("First Name") or "").strip()
                    last_name = str(values.get("Last Name") or "").strip()
                    uid = str(values.get("uid") or "").strip()
                    full_name = " ".join(part for part in (first_name, last_name) if part)
                    searchable = (first_name, last_name, full_name, uid)
                    if not any(normalized_query in value.casefold() for value in searchable):
                        continue

                    exact_uid = 0 if uid.casefold() == normalized_query else 1
                    name_prefix = 0 if full_name.casefold().startswith(normalized_query) else 1
                    result = {
                        "student_id": str(student["id"]),
                        "sheet": str(class_row["name"]),
                        "excel_row": int(student["position"]) + 2,
                        "first_name": first_name,
                        "last_name": last_name,
                        "full_name": full_name,
                        "uid": uid,
                    }
                    matches.append(
                        (
                            (
                                exact_uid,
                                name_prefix,
                                int(class_row["position"]) * 10000
                                + int(student["position"]),
                            ),
                            result,
                        )
                    )

        matches.sort(key=lambda item: item[0])
        safe_limit = max(1, min(int(limit), 100))
        return [result for _, result in matches[:safe_limit]]

    def list_quiz_banks(self) -> list[dict[str, Any]]:
        return self.quiz_banks.list_banks()

    def list_attachments(self, sheet_name: str) -> list[dict[str, Any]]:
        return self.attachments.list_for_sheet(sheet_name)

    def add_attachment(
        self,
        sheet_name: str,
        *,
        filename: str,
        content_type: str,
        file_data: bytes,
    ) -> dict[str, Any]:
        return self.attachments.add(
            sheet_name,
            filename=filename,
            content_type=content_type,
            file_data=file_data,
        )

    def remove_attachment(self, sheet_name: str, attachment_id: str) -> None:
        self.attachments.remove(sheet_name, attachment_id)

    def get_attachment(self, attachment_id: str) -> dict[str, Any]:
        return self.attachments.get(attachment_id)

    def materialize_attachments(
        self,
        sheet_name: str,
        attachment_ids: list[str],
        target_dir: Path,
    ) -> list[Path]:
        return self.attachments.materialize(sheet_name, attachment_ids, target_dir)
    def load_quiz_bank(self, bank_id: str) -> dict[str, Any]:
        return self.quiz_banks.load_bank(bank_id)

    def save_quiz_bank(
        self,
        bank_id: str,
        *,
        class_name: str,
        quiz_name: str,
        display_name: str,
        entries: Any,
    ) -> dict[str, Any]:
        return self.quiz_banks.save_bank(
            bank_id,
            class_name=class_name,
            quiz_name=quiz_name,
            display_name=display_name,
            entries=entries,
        )

    def save_sheet(self, sheet_name: str, rows: list[dict[str, Any]]) -> dict[str, Any]:
        if not isinstance(rows, list):
            raise StoreError("Rows must be provided as a list.")

        with self.lock, self._connect() as connection:
            class_row = self._class_row(connection, sheet_name)
            class_id = int(class_row["id"])
            known_headers = {
                str(row["name"]) for row in self._columns(connection, class_id)
            }
            existing_ids = {
                str(row["id"]) for row in self._students(connection, class_id)
            }
            submitted_ids: set[str] = set()
            for item in rows:
                if not isinstance(item, dict) or not isinstance(item.get("values"), dict):
                    raise StoreError("Each row must contain a values object.")
                student_id = str(item.get("student_id") or "").strip()
                if not student_id:
                    continue
                if student_id in submitted_ids:
                    raise StoreError(f"Student record {student_id!r} was submitted twice.")
                if student_id not in existing_ids:
                    raise StoreError(f"Student record {student_id!r} was not found.")
                submitted_ids.add(student_id)

            missing_ids = existing_ids - submitted_ids
            if missing_ids:
                raise StoreError(
                    "The roster save was incomplete. Reload the class sheet and try again."
                )

            for position, item in enumerate(rows):
                if not isinstance(item, dict) or not isinstance(item.get("values"), dict):
                    raise StoreError("Each row must contain a values object.")
                values = item["values"]
                unknown_headers = set(values) - known_headers
                if unknown_headers:
                    raise StoreError(f"Unknown columns: {', '.join(sorted(unknown_headers))}")

                normalized_values = {
                    header: coerce_cell_value(header, values.get(header))
                    for header in known_headers
                }
                student_id = str(item.get("student_id") or "").strip()
                if student_id:
                    if student_id not in existing_ids:
                        raise StoreError(f"Student record {student_id!r} was not found.")
                    connection.execute(
                        """
                        UPDATE students
                        SET position = ?, values_json = ?, updated_at = CURRENT_TIMESTAMP
                        WHERE id = ? AND class_id = ?
                        """,
                        (
                            position,
                            json.dumps(normalized_values, ensure_ascii=False),
                            student_id,
                            class_id,
                        ),
                    )
                else:
                    student_id = str(uuid4())
                    connection.execute(
                        """
                        INSERT INTO students(id, class_id, position, values_json)
                        VALUES(?, ?, ?, ?)
                        """,
                        (
                            student_id,
                            class_id,
                            position,
                            json.dumps(normalized_values, ensure_ascii=False),
                        ),
                    )
                    existing_ids.add(student_id)

        return self.load_sheet(sheet_name)

    def delete_student(self, sheet_name: str, student_id: str) -> dict[str, Any]:
        student_id = str(student_id or "").strip()
        if not student_id:
            raise StoreError("A student record was not specified.")

        with self.lock, self._connect() as connection:
            class_row = self._class_row(connection, sheet_name)
            cursor = connection.execute(
                "DELETE FROM students WHERE id = ? AND class_id = ?",
                (student_id, int(class_row["id"])),
            )
            if cursor.rowcount == 0:
                raise StoreError(f"Student record {student_id!r} was not found.")

        return self.load_sheet(sheet_name)

    def delete_class(self, sheet_name: str) -> dict[str, Any]:
        # Resolved before the class row is deleted -- announcement_path() raises if the
        # sheet doesn't exist, and it isn't a DB row, so nothing else would clean it up.
        stale_announcement_path = self.announcement_path(sheet_name)

        with self.lock, self._connect() as connection:
            class_row = self._class_row(connection, sheet_name)
            connection.execute("DELETE FROM classes WHERE id = ?", (int(class_row["id"]),))
            # Deleting the class row cascades to its columns, students, and attachments
            # (all declared ON DELETE CASCADE), but the template workbook is a plain file
            # rebuilt only on template-missing or new-class-import, so it needs an
            # explicit refresh here or it would keep a stale sheet for the deleted class.
            self._create_template_from_database(connection)

        stale_announcement_path.unlink(missing_ok=True)

        sheets = self.sheet_names()
        return {
            "sheets": sheets,
            "sheet_groups": self.sheet_groups(),
            "default_sheet": sheets[0] if sheets else "",
        }

    @staticmethod
    def _copy_row_format(worksheet: Any, source_row: int, target_row: int) -> None:
        if source_row < 2:
            return
        if worksheet.row_dimensions[source_row].height is not None:
            worksheet.row_dimensions[target_row].height = worksheet.row_dimensions[source_row].height
        for column in range(1, worksheet.max_column + 1):
            source = worksheet.cell(source_row, column)
            target = worksheet.cell(target_row, column)
            if source.has_style:
                target._style = copy(source._style)
            target.alignment = copy(source.alignment)
            target.protection = copy(source.protection)
            target.number_format = source.number_format

    def _write_database_to_workbook(self, output_path: Path) -> dict[str, dict[str, int]]:
        workbook = load_workbook(self.template_path)
        row_maps: dict[str, dict[str, int]] = {}
        try:
            with self._connect() as connection:
                for sheet_name in self.sheet_names():
                    class_row = self._class_row(connection, sheet_name)
                    class_id = int(class_row["id"])
                    worksheet = workbook[sheet_name]
                    columns = self._columns(connection, class_id)
                    students = self._students(connection, class_id)

                    for column_index, column in enumerate(columns, start=1):
                        worksheet.cell(1, column_index).value = str(column["name"])

                    required_last_row = len(students) + 1
                    if required_last_row > worksheet.max_row:
                        source_row = 2 if worksheet.max_row >= 2 else 1
                        for target_row in range(worksheet.max_row + 1, required_last_row + 1):
                            self._copy_row_format(worksheet, source_row, target_row)

                    row_map: dict[str, int] = {}
                    for position, student in enumerate(students):
                        row_number = position + 2
                        row_map[str(student["id"])] = row_number
                        values = json.loads(str(student["values_json"]))
                        for column_index, column in enumerate(columns, start=1):
                            header = str(column["name"])
                            worksheet.cell(row_number, column_index).value = values.get(header)

                    clear_from = len(students) + 2
                    for row_number in range(clear_from, worksheet.max_row + 1):
                        for column_index in range(1, worksheet.max_column + 1):
                            worksheet.cell(row_number, column_index).value = None
                    row_maps[sheet_name] = row_map

            output_path.parent.mkdir(parents=True, exist_ok=True)
            temporary = output_path.with_suffix(".tmp.xlsx")
            workbook.save(temporary)
            temporary.replace(output_path)
            return row_maps
        finally:
            workbook.close()

    def prepare_runtime_workbook(self) -> dict[str, dict[str, int]]:
        with self.lock:
            return self._write_database_to_workbook(self.workbook_path)

    def sync_runtime_workbook(self) -> None:
        if not self.workbook_path.exists():
            return

        with self.lock:
            workbook = load_workbook(self.workbook_path, data_only=False)
            try:
                with self._connect() as connection:
                    for sheet_name in self.sheet_names():
                        if sheet_name not in workbook.sheetnames:
                            continue
                        class_row = self._class_row(connection, sheet_name)
                        class_id = int(class_row["id"])
                        worksheet = workbook[sheet_name]
                        headers = header_cells(worksheet)
                        existing_columns = {
                            str(row["name"]): row for row in self._columns(connection, class_id)
                        }
                        for position, (column_index, header) in enumerate(headers):
                            if header not in existing_columns:
                                width = (
                                    worksheet.column_dimensions[get_column_letter(column_index)].width
                                    or 13
                                )
                                connection.execute(
                                    """
                                    INSERT INTO columns(class_id, name, position, width)
                                    VALUES(?, ?, ?, ?)
                                    """,
                                    (class_id, header, position, float(width)),
                                )

                        students = self._students(connection, class_id)
                        for position, student in enumerate(students):
                            row_number = position + 2
                            values = {
                                header: json_value(worksheet.cell(row_number, column_index).value)
                                for column_index, header in headers
                            }
                            connection.execute(
                                """
                                UPDATE students
                                SET values_json = ?, updated_at = CURRENT_TIMESTAMP
                                WHERE id = ?
                                """,
                                (
                                    json.dumps(values, ensure_ascii=False),
                                    str(student["id"]),
                                ),
                            )
            finally:
                workbook.close()

    def remove_runtime_workbook(self) -> None:
        try:
            self.workbook_path.unlink(missing_ok=True)
        except OSError:
            pass

    def _scoped_class_rows(
        self,
        connection: sqlite3.Connection,
        semester_name: str | None,
    ) -> list[sqlite3.Row]:
        """Class rows for one semester block, in tab order.

        semester_name of None means every class (the whole database). The
        UNASSIGNED_SEMESTER sentinel means the legacy classes that predate semesters
        and therefore have no semester_id -- the group the frontend labels
        "Other Classes".
        """
        if semester_name is None:
            return connection.execute(
                "SELECT id, name FROM classes ORDER BY position"
            ).fetchall()

        if semester_name == UNASSIGNED_SEMESTER:
            return connection.execute(
                "SELECT id, name FROM classes WHERE semester_id IS NULL ORDER BY position"
            ).fetchall()

        semester_row = connection.execute(
            "SELECT id FROM semesters WHERE name = ?", (semester_name,)
        ).fetchone()
        if semester_row is None:
            raise StoreError(f"Semester {semester_name!r} was not found.")
        return connection.execute(
            "SELECT id, name FROM classes WHERE semester_id = ? ORDER BY position",
            (int(semester_row["id"]),),
        ).fetchall()

    @staticmethod
    def _export_filename(prefix: str, semester_name: str | None) -> str:
        if semester_name is None:
            return f"{prefix}_All_Classes.xlsx"
        label = "Other_Classes" if semester_name == UNASSIGNED_SEMESTER else semester_name
        safe_label = re.sub(r"[^A-Za-z0-9._-]+", "_", label).strip("._") or "Export"
        return f"{prefix}_{safe_label}.xlsx"

    def export_public_workbook(self, semester_name: str | None = None) -> Path:
        """Full round-trip copy of every internal column, one sheet per class.

        Scoped to a single semester block when semester_name is given, so the file
        holds exactly that semester's roster rather than every class ever imported.
        Built from the formatted template, then the out-of-scope sheets are dropped,
        which keeps each remaining sheet's column widths and header styling intact.
        """
        with self.lock:
            with self._connect() as connection:
                scoped_names = [str(row["name"]) for row in self._scoped_class_rows(connection, semester_name)]
            if not scoped_names:
                raise StoreError("That semester has no classes to export.")

            output_path = self.export_dir / self._export_filename("Student_Feedback_Export", semester_name)
            self._write_database_to_workbook(output_path)

            if semester_name is not None:
                workbook = load_workbook(output_path)
                try:
                    for sheet_name in list(workbook.sheetnames):
                        if sheet_name not in scoped_names:
                            del workbook[sheet_name]
                    temporary = output_path.with_suffix(".tmp.xlsx")
                    workbook.save(temporary)
                    temporary.replace(output_path)
                finally:
                    workbook.close()
            return output_path

    def export_summary_report(self, semester_name: str | None = None) -> Path:
        """A curated, parent-status-style report: one sheet per class, with only the
        columns staff actually review day to day, in plain-language Chinese headers.
        Distinct from export_public_workbook, which is a full raw round-trip copy of
        every internal column. Scoped to one semester block when semester_name is given.
        """
        workbook = Workbook()
        try:
            with self.lock, self._connect() as connection:
                class_rows = self._scoped_class_rows(connection, semester_name)
                if not class_rows:
                    raise StoreError("That semester has no classes to export.")
                for index, class_row in enumerate(class_rows):
                    worksheet = workbook.active if index == 0 else workbook.create_sheet()
                    worksheet.title = str(class_row["name"])
                    for column_index, header in enumerate(REPORT_COLUMNS, start=1):
                        worksheet.cell(1, column_index).value = header
                        worksheet.column_dimensions[get_column_letter(column_index)].width = (
                            REPORT_COLUMN_WIDTHS[column_index - 1]
                        )
                    worksheet.freeze_panes = "A2"

                    students = self._students(connection, int(class_row["id"]))
                    for row_index, student in enumerate(students, start=2):
                        values = json.loads(str(student["values_json"]))
                        for column_index, value in enumerate(_report_row(values), start=1):
                            worksheet.cell(row_index, column_index).value = value

            output_path = self.export_dir / self._export_filename("Student_Report_Export", semester_name)
            output_path.parent.mkdir(parents=True, exist_ok=True)
            temporary = output_path.with_suffix(".tmp.xlsx")
            workbook.save(temporary)
            temporary.replace(output_path)
            return output_path
        finally:
            workbook.close()



