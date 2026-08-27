from __future__ import annotations

import argparse
from copy import copy
from datetime import date, datetime, time
import hashlib
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
import json
import mimetypes
import os
from pathlib import Path
import re
import subprocess
import sys
from tempfile import TemporaryDirectory
import threading
from typing import Any
from urllib.parse import parse_qs, unquote, urlparse
import webbrowser

from openpyxl import load_workbook
from openpyxl.utils import get_column_letter

from attachment_store import AttachmentStoreError, MAX_ATTACHMENT_BYTES
from feedback_common import DEFAULT_WORKBOOK
from database_store import (
    SQLiteFeedbackStore,
    StoreError,
    announcement_filename,
)
from import_enrollment import group_by_class, read_enrollment_rows
from quiz_bank_store import QuizBankStoreError


PROJECT_DIR = Path(__file__).resolve().parent
STATIC_DIR = PROJECT_DIR / "frontend"
DEFAULT_ANNOUNCEMENT_DIR = PROJECT_DIR / "announcements"
MAX_REQUEST_BYTES = 8 * 1024 * 1024
MAX_IMPORT_BYTES = 25 * 1024 * 1024
ACTION_TIMEOUT_SECONDS = 30 * 60
APP_PROBE_TIMEOUT_SECONDS = 60

DEFAULT_APP_DATA_DIR = PROJECT_DIR / "app_data"
DEFAULT_DATABASE = DEFAULT_APP_DATA_DIR / "feedback.db"
DEFAULT_EXPORT_DIR = PROJECT_DIR / "exports"

FrontendError = StoreError
WorkbookStore = SQLiteFeedbackStore

class ActionRunner:
    def __init__(self, store: WorkbookStore) -> None:
        self.store = store

    @staticmethod
    def _row_spec(rows: Any) -> str:
        if not isinstance(rows, list) or not rows:
            raise FrontendError("Select at least one student row.")
        parsed: list[int] = []
        for value in rows:
            try:
                row_number = int(value)
            except (TypeError, ValueError) as exc:
                raise FrontendError(f"Invalid selected row {value!r}.") from exc
            if row_number < 2:
                raise FrontendError("Student rows must be row 2 or later.")
            parsed.append(row_number)
        return ",".join(str(row) for row in sorted(set(parsed)))

    @staticmethod
    def _attachment_ids(payload: dict[str, Any]) -> list[str]:
        raw_ids = payload.get("attachment_ids") or []
        if not isinstance(raw_ids, list):
            raise FrontendError("Attachment ids must be provided as a list.")
        return list(dict.fromkeys(str(value).strip() for value in raw_ids if str(value).strip()))

    def command_for(
        self,
        payload: dict[str, Any],
        *,
        attachment_paths: list[Path] | None = None,
    ) -> tuple[list[str], str]:
        action = str(payload.get("action") or "").strip()
        sheet_name = str(payload.get("sheet") or "").strip()
        if sheet_name not in self.store.sheet_names():
            raise FrontendError(f"Sheet {sheet_name!r} was not found.")
        row_spec = self._row_spec(payload.get("rows"))
        attachment_ids = self._attachment_ids(payload)
        if attachment_ids and not action.startswith("paste-"):
            raise FrontendError("Attachments can only be used with paste actions.")
        if attachment_ids and "," in row_spec:
            raise FrontendError(
                "Attachments stay open for review, so select one student for each attachment run."
            )
        announcement_path = self.store.announcement_path(sheet_name)
        # Paragraph 1 of a generated comment comes from the teacher's lesson recap, not
        # from the announcement. The announcement is separately pasted to whole classes,
        # so reusing it here put things like "next week is cancelled" at the top of every
        # student's feedback.
        class_review_args = ["--class-review", self.store.read_lesson_recap(sheet_name)]
        attachment_args = [
            item
            for path in (attachment_paths or [])
            for item in ("--attachment", str(path))
        ]

        common = [
            "--workbook",
            str(self.store.workbook_path),
            "--sheet",
            sheet_name,
            "--rows",
            row_spec,
        ]

        if action == "generate-comments":
            return [
                sys.executable,
                str(PROJECT_DIR / "feedback_generator.py"),
                *common,
                "--write",
                "--feedback-type",
                "general",
                "--feedback-column",
                "Feedback",
                *class_review_args,
            ], "Generated comments"

        if action == "paste-announcement":
            return [
                sys.executable,
                str(PROJECT_DIR / "paste_sender.py"),
                *common,
                "--action",
                "mass-notification",
                "--mass-message-file",
                str(announcement_path),
                "--mode",
                "paste-only",
                "--fallback-channel",
                *attachment_args,
            ], "Pasted announcement"

        if action == "paste-comments":
            return [
                sys.executable,
                str(PROJECT_DIR / "paste_sender.py"),
                *common,
                "--action",
                "comment",
                "--message-column",
                "Feedback",
                "--mode",
                "paste-only",
                *attachment_args,
            ], "Pasted comments"

        if action == "check-group-chat":
            return [
                sys.executable,
                str(PROJECT_DIR / "paste_sender.py"),
                *common,
                "--action",
                "check-group-chat",
                "--mode",
                "paste-only",
                "--channel",
                self._check_channel(payload),
            ], "Checked group chat status"

        if action in {"generate-quiz-feedback", "paste-quiz-feedback"}:
            quiz_number = str(payload.get("quiz_number") or "1")
            if quiz_number not in {"1", "2"}:
                raise FrontendError("Quiz number must be 1 or 2.")
            feedback_column = f"Quiz{quiz_number} Feedback"
            score_column = f"Quiz{quiz_number} Score"
            average_column = f"Quiz{quiz_number} Average"
            current_headers = {column["key"] for column in self.store.load_sheet(sheet_name)["columns"]}
            required = {feedback_column, score_column}
            if action == "generate-quiz-feedback":
                required.add(average_column)
            missing = sorted(required - current_headers)
            if missing:
                raise FrontendError(
                    f"The selected sheet is missing columns: {', '.join(missing)}"
                )

            if action == "generate-quiz-feedback":
                return [
                    sys.executable,
                    str(PROJECT_DIR / "feedback_generator.py"),
                    *common,
                    "--write",
                    "--feedback-type",
                    "quiz",
                    "--feedback-column",
                    feedback_column,
                    *class_review_args,
                    "--calculate-quiz-average",
                    "--quiz-score-column",
                    score_column,
                    "--quiz-average-column",
                    average_column,
                    "--only-with-value-column",
                    score_column,
                    "--clear-feedback-for-missing-value",
                ], f"Generated quiz {quiz_number} feedback"

            return [
                sys.executable,
                str(PROJECT_DIR / "paste_sender.py"),
                *common,
                "--action",
                "comment",
                "--message-column",
                feedback_column,
                "--feedback-type",
                "quiz",
                "--mode",
                "paste-only",
                *attachment_args,
            ], f"Pasted quiz {quiz_number} feedback"

        raise FrontendError(f"Unknown action {action!r}.")

    @staticmethod
    def _check_channel(payload: dict[str, Any]) -> str:
        channel = str(payload.get("channel") or "auto").strip().lower()
        if channel not in {"auto", "wecom", "whatsapp"}:
            raise FrontendError(f"Unknown channel {channel!r}.")
        return channel

    def _bulk_group_chat_sheets(self, payload: dict[str, Any]) -> list[str]:
        raw_sheets = payload.get("sheets")
        if not isinstance(raw_sheets, list) or not raw_sheets:
            raise FrontendError("Select at least one class to check.")
        known = set(self.store.sheet_names())
        sheets: list[str] = []
        for value in raw_sheets:
            sheet_name = str(value).strip()
            if sheet_name not in known:
                raise FrontendError(f"Sheet {sheet_name!r} was not found.")
            if sheet_name not in sheets:
                sheets.append(sheet_name)
        return sheets

    def _run_bulk_group_chat(
        self,
        payload: dict[str, Any],
        environment: dict[str, str],
    ) -> dict[str, Any]:
        """Run the group-chat check across whole class rosters, one class at a time.

        paste_sender.py takes a single --sheet, so this loops it per class rather than
        teaching it about multiple sheets. The runtime workbook is prepared once and
        synced back once at the end, so every class's results land in the database
        together even though each class ran as its own subprocess.
        """
        sheets = self._bulk_group_chat_sheets(payload)
        channel = self._check_channel(payload)

        # Probe the desktop apps once up front. Without this, an unavailable app makes
        # every single student re-attempt a launch that cannot succeed -- slow, and for
        # WhatsApp it opens a web.whatsapp.com browser tab per student. Since readiness
        # is established here, the per-class runs below use --no-auto-open so a long
        # unattended batch never tries to open anything on its own.
        #
        # The probe runs as its own process rather than importing pywinauto here: COM
        # misbehaves on the threaded HTTP server's request threads, where an in-process
        # probe was observed to hang instead of returning.
        try:
            probe = subprocess.run(
                [sys.executable, str(PROJECT_DIR / "paste_sender.py"), "--check-apps"],
                cwd=PROJECT_DIR,
                capture_output=True,
                text=True,
                encoding="utf-8",
                errors="replace",
                timeout=APP_PROBE_TIMEOUT_SECONDS,
                env=environment,
                check=False,
            )
            statuses = json.loads(probe.stdout.strip() or "{}")
        except (subprocess.TimeoutExpired, json.JSONDecodeError) as exc:
            raise FrontendError(
                "Could not determine whether WeCom or WhatsApp is open. "
                "Make sure at least one is running, then try again."
            ) from exc

        available = [key for key, status in statuses.items() if status.get("window_found")]
        # A forced channel needs that specific app; auto just needs something to work with.
        required = [channel] if channel != "auto" else list(statuses)
        if not any(key in available for key in required):
            details = "; ".join(
                f"{statuses[key].get('display_name', key)}: {statuses[key].get('message', 'unavailable')}"
                for key in required
                if key in statuses
            )
            wanted = (
                statuses.get(channel, {}).get("display_name", channel)
                if channel != "auto"
                else "Neither WeCom nor WhatsApp"
            )
            raise FrontendError(
                f"{wanted} has no open window, so nothing can be checked. "
                f"Open it and try again. ({details})"
            )

        sections: list[str] = [
            f"Channel: {channel} | available apps: "
            + ", ".join(statuses[key].get("display_name", key) for key in available)
        ]
        failed_sheets: list[str] = []
        checked_sheets = 0

        with self.store.lock:
            self.store.prepare_runtime_workbook()
            try:
                for sheet_name in sheets:
                    rows = [
                        int(row["excel_row"])
                        for row in self.store.load_sheet(sheet_name)["rows"]
                        if row.get("excel_row")
                    ]
                    if not rows:
                        sections.append(f"===== {sheet_name} =====\n(no students on this roster)")
                        continue

                    command = [
                        sys.executable,
                        str(PROJECT_DIR / "paste_sender.py"),
                        "--workbook",
                        str(self.store.workbook_path),
                        "--sheet",
                        sheet_name,
                        "--rows",
                        ",".join(str(row) for row in sorted(rows)),
                        "--action",
                        "check-group-chat",
                        "--mode",
                        "paste-only",
                        "--no-auto-open",
                        "--channel",
                        channel,
                    ]
                    try:
                        completed = subprocess.run(
                            command,
                            cwd=PROJECT_DIR,
                            capture_output=True,
                            text=True,
                            encoding="utf-8",
                            errors="replace",
                            timeout=ACTION_TIMEOUT_SECONDS,
                            env=environment,
                            check=False,
                        )
                    except subprocess.TimeoutExpired:
                        failed_sheets.append(sheet_name)
                        sections.append(
                            f"===== {sheet_name} =====\n"
                            f"Timed out after {ACTION_TIMEOUT_SECONDS // 60} minutes."
                        )
                        continue

                    checked_sheets += 1
                    if completed.returncode != 0:
                        failed_sheets.append(sheet_name)
                    body = "\n".join(
                        part
                        for part in (completed.stdout.strip(), completed.stderr.strip())
                        if part
                    )
                    sections.append(f"===== {sheet_name} =====\n{body}")
            finally:
                self.store.sync_runtime_workbook()
                self.store.remove_runtime_workbook()

        label = f"Checked group chats for {checked_sheets} class(es)"
        if failed_sheets:
            label += f"; {len(failed_sheets)} had errors"
        output = "\n\n".join(sections)
        if len(output) > 50000:
            output = output[-50000:]
        return {
            "ok": not failed_sheets,
            "label": label,
            "returncode": 1 if failed_sheets else 0,
            "output": output,
        }

    def run(self, payload: dict[str, Any]) -> dict[str, Any]:
        environment = os.environ.copy()
        environment["PYTHONUTF8"] = "1"
        environment["PYTHONDONTWRITEBYTECODE"] = "1"
        environment["FEEDBACK_DATABASE_PATH"] = str(self.store.database_path)

        if str(payload.get("action") or "").strip() == "check-group-chat-bulk":
            return self._run_bulk_group_chat(payload, environment)

        attachment_ids = self._attachment_ids(payload)
        with TemporaryDirectory(
            prefix="attachment_action_",
            dir=self.store.app_data_dir,
        ) as temporary_dir:
            attachment_paths = self.store.materialize_attachments(
                str(payload.get("sheet") or "").strip(),
                attachment_ids,
                Path(temporary_dir),
            )
            with self.store.lock:
                self.store.prepare_runtime_workbook()
                try:
                    command, label = self.command_for(
                        payload,
                        attachment_paths=attachment_paths,
                    )
                    try:
                        completed = subprocess.run(
                            command,
                            cwd=PROJECT_DIR,
                            capture_output=True,
                            text=True,
                            encoding="utf-8",
                            errors="replace",
                            timeout=ACTION_TIMEOUT_SECONDS,
                            env=environment,
                            check=False,
                        )
                    except subprocess.TimeoutExpired as exc:
                        raise FrontendError(
                            f"{label} timed out after {ACTION_TIMEOUT_SECONDS // 60} minutes."
                        ) from exc
                    finally:
                        self.store.sync_runtime_workbook()
                finally:
                    self.store.remove_runtime_workbook()

        stdout = completed.stdout.strip()
        stderr = completed.stderr.strip()
        output = "\n".join(part for part in (stdout, stderr) if part)
        if len(output) > 50000:
            output = output[-50000:]
        return {
            "ok": completed.returncode == 0,
            "label": label,
            "returncode": completed.returncode,
            "output": output,
        }


class FrontendHandler(BaseHTTPRequestHandler):
    store: WorkbookStore
    runner: ActionRunner
    server_version = "FeedbackFrontend/1.0"

    def log_message(self, format: str, *args: Any) -> None:
        print(f"[frontend] {self.address_string()} {format % args}")

    def _send_json(self, status: int, payload: dict[str, Any]) -> None:
        body = json.dumps(payload, ensure_ascii=False).encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(body)))
        self.send_header("Cache-Control", "no-store")
        self.end_headers()
        self.wfile.write(body)

    def _read_json(self) -> dict[str, Any]:
        try:
            length = int(self.headers.get("Content-Length", "0"))
        except ValueError as exc:
            raise FrontendError("Invalid request length.") from exc
        if length <= 0 or length > MAX_REQUEST_BYTES:
            raise FrontendError("Request body is empty or too large.")
        try:
            payload = json.loads(self.rfile.read(length).decode("utf-8"))
        except (UnicodeDecodeError, json.JSONDecodeError) as exc:
            raise FrontendError("Request body must be valid UTF-8 JSON.") from exc
        if not isinstance(payload, dict):
            raise FrontendError("Request body must be a JSON object.")
        return payload

    def _read_binary(self, max_bytes: int) -> bytes:
        try:
            length = int(self.headers.get("Content-Length", "0"))
        except ValueError as exc:
            raise FrontendError("Invalid request length.") from exc
        if length <= 0:
            raise FrontendError("The uploaded file is empty.")
        if length > max_bytes:
            raise FrontendError("Each attachment must be 100 MB or smaller.")
        return self.rfile.read(length)

    def _send_attachment(self, attachment: dict[str, Any]) -> None:
        content = attachment["file_data"]
        self.send_response(200)
        self.send_header("Content-Type", attachment["content_type"])
        self.send_header("Content-Length", str(len(content)))
        self.send_header("Content-Disposition", "inline")
        self.send_header("Cache-Control", "no-store")
        self.end_headers()
        self.wfile.write(content)

    def _send_download(self, file_path: Path) -> None:
        content = file_path.read_bytes()
        self.send_response(200)
        self.send_header(
            "Content-Type",
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
        self.send_header("Content-Length", str(len(content)))
        # _export_filename() already reduces the name to ASCII word characters, so it
        # needs no RFC 5987 escaping even when the semester name is non-Latin.
        self.send_header("Content-Disposition", f'attachment; filename="{file_path.name}"')
        self.send_header("Cache-Control", "no-store")
        self.end_headers()
        self.wfile.write(content)

    def _serve_static(self, request_path: str) -> None:
        relative = "index.html" if request_path in {"", "/"} else unquote(request_path.lstrip("/"))
        target = (STATIC_DIR / relative).resolve()
        try:
            target.relative_to(STATIC_DIR.resolve())
        except ValueError:
            self.send_error(404)
            return
        if not target.is_file():
            self.send_error(404)
            return
        content = target.read_bytes()
        content_type = mimetypes.guess_type(target.name)[0] or "application/octet-stream"
        if target.suffix in {".html", ".css", ".js"}:
            content_type += "; charset=utf-8"
        self.send_response(200)
        self.send_header("Content-Type", content_type)
        self.send_header("Content-Length", str(len(content)))
        self.send_header("Cache-Control", "no-cache")
        self.end_headers()
        self.wfile.write(content)

    def do_GET(self) -> None:
        parsed = urlparse(self.path)
        try:
            if parsed.path == "/api/bootstrap":
                sheets = self.store.sheet_names()
                self._send_json(
                    200,
                    {
                        "ok": True,
                        "database": str(self.store.database_path),
                        "import_source": str(self.store.source_workbook_path),
                        "data_source": "sqlite",
                        "sheets": sheets,
                        "sheet_groups": self.store.sheet_groups(),
                        "default_sheet": sheets[0] if sheets else "",
                        "quiz_banks": self.store.list_quiz_banks(),
                        "platform": sys.platform,
                        "paste_supported": os.name == "nt" or sys.platform == "darwin",
                    },
                )
                return
            if parsed.path == "/api/sheet":
                sheet_name = parse_qs(parsed.query).get("name", [""])[0]
                self._send_json(200, {"ok": True, "data": self.store.load_sheet(sheet_name)})
                return
            if parsed.path == "/api/search":
                query = parse_qs(parsed.query).get("q", [""])[0]
                self._send_json(
                    200,
                    {"ok": True, "results": self.store.search_students(query)},
                )
                return
            if parsed.path == "/api/semesters":
                self._send_json(
                    200,
                    {"ok": True, "semesters": self.store.list_semesters()},
                )
                return
            if parsed.path == "/api/quiz-banks":
                self._send_json(
                    200,
                    {"ok": True, "banks": self.store.list_quiz_banks()},
                )
                return
            if parsed.path == "/api/quiz-bank":
                bank_id = parse_qs(parsed.query).get("id", [""])[0]
                self._send_json(
                    200,
                    {"ok": True, "bank": self.store.load_quiz_bank(bank_id)},
                )
                return
            if parsed.path == "/api/attachment":
                attachment_id = parse_qs(parsed.query).get("id", [""])[0]
                self._send_attachment(self.store.get_attachment(attachment_id))
                return
            if parsed.path == "/api/export/download":
                query = parse_qs(parsed.query)
                export_type = (query.get("type", ["full"])[0] or "full").strip()
                # An absent semester param means "everything"; the frontend always sends
                # one so a download matches the semester block currently on screen.
                raw_semester = query.get("semester", [None])[0]
                semester_name = raw_semester.strip() if raw_semester else None
                if export_type == "report":
                    export_path = self.store.export_summary_report(semester_name)
                elif export_type == "full":
                    export_path = self.store.export_public_workbook(semester_name)
                else:
                    raise FrontendError(f"Unknown export type {export_type!r}.")
                self._send_download(export_path)
                return
            self._serve_static(parsed.path)
        except (AttachmentStoreError, FrontendError, QuizBankStoreError) as exc:
            self._send_json(400, {"ok": False, "error": str(exc)})
        except Exception as exc:
            self._send_json(500, {"ok": False, "error": f"{type(exc).__name__}: {exc}"})

    def do_POST(self) -> None:
        parsed = urlparse(self.path)
        try:
            if parsed.path == "/api/attachment/upload":
                query = parse_qs(parsed.query)
                sheet_name = query.get("sheet", [""])[0]
                filename = query.get("name", [""])[0]
                attachment = self.store.add_attachment(
                    sheet_name,
                    filename=filename,
                    content_type=self.headers.get("Content-Type", ""),
                    file_data=self._read_binary(MAX_ATTACHMENT_BYTES),
                )
                self._send_json(
                    200,
                    {
                        "ok": True,
                        "attachment": attachment,
                        "attachments": self.store.list_attachments(sheet_name),
                    },
                )
                return
            if parsed.path in {"/api/import/preview", "/api/import/commit"}:
                query = parse_qs(parsed.query)
                semester_name = (query.get("semester", [""])[0] or "").strip()
                if not semester_name:
                    raise FrontendError("Enter a semester name first.")
                class_id_raw = query.get("class_ids", [""])[0]
                class_ids = {value.strip() for value in class_id_raw.split(",") if value.strip()}
                overwrite = (query.get("overwrite", [""])[0] or "").strip().lower() in {"1", "true"}
                confirm_delete = (query.get("confirm_delete", [""])[0] or "").strip().lower() in {
                    "1",
                    "true",
                }

                file_bytes = self._read_binary(MAX_IMPORT_BYTES)
                with TemporaryDirectory(
                    prefix="enrollment_import_",
                    dir=self.store.app_data_dir,
                ) as temporary_dir:
                    xlsx_path = Path(temporary_dir) / "enrollment.xlsx"
                    xlsx_path.write_bytes(file_bytes)
                    try:
                        rows = read_enrollment_rows(xlsx_path)
                    except ValueError as exc:
                        raise FrontendError(str(exc)) from exc

                groups = group_by_class(rows, class_ids=class_ids or None)
                if not groups:
                    raise FrontendError(
                        "No paid, enrolled rows matched this file (or the selected classes)."
                    )

                commit = parsed.path == "/api/import/commit"
                with self.store.lock:
                    if commit and overwrite and not confirm_delete:
                        preview_plan = self.store.import_students(
                            semester_name=semester_name,
                            class_groups=list(groups.values()),
                            commit=False,
                            overwrite=True,
                        )
                        removal_count = sum(
                            len(class_plan.get("students_to_remove") or [])
                            for class_plan in preview_plan["classes"]
                        )
                        if removal_count:
                            self._send_json(
                                200,
                                {
                                    "ok": True,
                                    "plan": preview_plan,
                                    "committed": False,
                                    "requires_confirm_delete": True,
                                    "removal_count": removal_count,
                                },
                            )
                            return

                    plan = self.store.import_students(
                        semester_name=semester_name,
                        class_groups=list(groups.values()),
                        commit=commit,
                        overwrite=overwrite,
                    )
                    if commit:
                        self.store.ensure_announcement_files()
                self._send_json(200, {"ok": True, "plan": plan, "committed": commit})
                return
            payload = self._read_json()
            if parsed.path == "/api/sheet/save":
                data = self.store.save_sheet(str(payload.get("sheet") or ""), payload.get("rows"))
                self._send_json(200, {"ok": True, "data": data})
                return
            if parsed.path == "/api/semester/delete":
                semester_name = str(payload.get("semester") or "").strip()
                if not semester_name:
                    raise FrontendError("Semester name is required.")
                confirmed = bool(payload.get("confirm"))
                plan = self.store.delete_semester(semester_name, commit=confirmed)
                self._send_json(200, {"ok": True, "plan": plan, "deleted": confirmed})
                return
            if parsed.path == "/api/announcement/save":
                path = self.store.write_announcement(
                    str(payload.get("sheet") or ""),
                    str(payload.get("text") or ""),
                )
                self._send_json(200, {"ok": True, "path": str(path)})
                return
            if parsed.path == "/api/lesson-recap/save":
                recap = self.store.write_lesson_recap(
                    str(payload.get("sheet") or ""),
                    str(payload.get("text") or ""),
                )
                self._send_json(200, {"ok": True, "lesson_recap": recap})
                return
            if parsed.path == "/api/quiz-bank/save":
                bank = self.store.save_quiz_bank(
                    str(payload.get("bank_id") or ""),
                    class_name=str(payload.get("class_name") or ""),
                    quiz_name=str(payload.get("quiz_name") or ""),
                    display_name=str(payload.get("display_name") or ""),
                    entries=payload.get("entries"),
                )
                self._send_json(
                    200,
                    {
                        "ok": True,
                        "bank": bank,
                        "banks": self.store.list_quiz_banks(),
                    },
                )
                return
            if parsed.path == "/api/attachment/delete":
                sheet_name = str(payload.get("sheet") or "")
                self.store.remove_attachment(
                    sheet_name,
                    str(payload.get("attachment_id") or ""),
                )
                self._send_json(
                    200,
                    {
                        "ok": True,
                        "attachments": self.store.list_attachments(sheet_name),
                    },
                )
                return
            if parsed.path == "/api/student/delete":
                data = self.store.delete_student(
                    str(payload.get("sheet") or ""),
                    str(payload.get("student_id") or ""),
                )
                self._send_json(200, {"ok": True, "data": data})
                return
            if parsed.path == "/api/class/delete":
                result = self.store.delete_class(str(payload.get("sheet") or ""))
                self._send_json(200, {"ok": True, **result})
                return
            if parsed.path == "/api/action":
                result = self.runner.run(payload)
                self._send_json(200 if result["ok"] else 422, result)
                return
            if parsed.path in {"/api/export", "/api/export/report"}:
                # Writes the file into exports/ and reports its path, without streaming a
                # download. The frontend buttons use GET /api/export/download instead; this
                # stays for scripted/API use.
                raw_semester = payload.get("semester")
                semester_name = str(raw_semester).strip() if raw_semester else None
                if parsed.path == "/api/export":
                    export_path = self.store.export_public_workbook(semester_name)
                    label = "Excel export"
                else:
                    export_path = self.store.export_summary_report(semester_name)
                    label = "Report export"
                self._send_json(
                    200,
                    {
                        "ok": True,
                        "path": str(export_path),
                        "message": f"{label} created. Editing it will not change the database.",
                    },
                )
                return
            self._send_json(404, {"ok": False, "error": "API endpoint not found."})
        except (AttachmentStoreError, FrontendError, QuizBankStoreError) as exc:
            self._send_json(400, {"ok": False, "error": str(exc)})
        except Exception as exc:
            self._send_json(500, {"ok": False, "error": f"{type(exc).__name__}: {exc}"})


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="Run the local feedback automation frontend.")
    parser.add_argument(
        "--workbook",
        type=Path,
        default=DEFAULT_WORKBOOK,
        help="Workbook imported only when the SQLite database is first created.",
    )
    parser.add_argument("--database", type=Path, default=DEFAULT_DATABASE)
    parser.add_argument("--app-data-dir", type=Path, default=DEFAULT_APP_DATA_DIR)
    parser.add_argument("--export-dir", type=Path, default=DEFAULT_EXPORT_DIR)
    parser.add_argument("--announcement-dir", type=Path, default=DEFAULT_ANNOUNCEMENT_DIR)
    parser.add_argument("--host", default="127.0.0.1")
    parser.add_argument("--port", type=int, default=8765)
    parser.add_argument("--no-browser", action="store_true")
    return parser


def main() -> None:
    args = build_parser().parse_args()
    store = WorkbookStore(
        database_path=args.database,
        source_workbook_path=args.workbook,
        announcement_dir=args.announcement_dir,
        app_data_dir=args.app_data_dir,
        export_dir=args.export_dir,
    )
    FrontendHandler.store = store
    FrontendHandler.runner = ActionRunner(store)

    server = ThreadingHTTPServer((args.host, args.port), FrontendHandler)
    host, port = server.server_address[:2]
    url = f"http://{host}:{port}"
    print(f"Feedback frontend: {url}")
    print(f"Database: {store.database_path}")
    print(f"Initial import source: {store.source_workbook_path}")
    print("Press Ctrl+C to stop.")

    if not args.no_browser:
        threading.Timer(0.6, lambda: webbrowser.open(url)).start()

    try:
        server.serve_forever()
    except KeyboardInterrupt:
        print("\nStopping feedback frontend.")
    finally:
        server.server_close()


if __name__ == "__main__":
    main()


