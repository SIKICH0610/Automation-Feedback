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
import threading
from typing import Any
from urllib.parse import parse_qs, unquote, urlparse
import webbrowser

from openpyxl import load_workbook
from openpyxl.utils import get_column_letter

from feedback_common import DEFAULT_WORKBOOK
from database_store import (
    SQLiteFeedbackStore,
    StoreError,
    announcement_filename,
)


PROJECT_DIR = Path(__file__).resolve().parent
STATIC_DIR = PROJECT_DIR / "frontend"
DEFAULT_ANNOUNCEMENT_DIR = PROJECT_DIR / "announcements"
MAX_REQUEST_BYTES = 8 * 1024 * 1024
ACTION_TIMEOUT_SECONDS = 30 * 60

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

    def command_for(self, payload: dict[str, Any]) -> tuple[list[str], str]:
        action = str(payload.get("action") or "").strip()
        sheet_name = str(payload.get("sheet") or "").strip()
        if sheet_name not in self.store.sheet_names():
            raise FrontendError(f"Sheet {sheet_name!r} was not found.")
        row_spec = self._row_spec(payload.get("rows"))
        announcement_path = self.store.announcement_path(sheet_name)

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
                "--class-review-file",
                str(announcement_path),
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
            ], "Pasted comments"

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
                    "--class-review-file",
                    str(announcement_path),
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
            ], f"Pasted quiz {quiz_number} feedback"

        raise FrontendError(f"Unknown action {action!r}.")

    def run(self, payload: dict[str, Any]) -> dict[str, Any]:
        environment = os.environ.copy()
        environment["PYTHONUTF8"] = "1"
        environment["PYTHONDONTWRITEBYTECODE"] = "1"

        with self.store.lock:
            self.store.prepare_runtime_workbook()
            try:
                command, label = self.command_for(payload)
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
                        "default_sheet": sheets[0] if sheets else "",
                        "platform": sys.platform,
                        "paste_supported": os.name == "nt",
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
            self._serve_static(parsed.path)
        except FrontendError as exc:
            self._send_json(400, {"ok": False, "error": str(exc)})
        except Exception as exc:
            self._send_json(500, {"ok": False, "error": f"{type(exc).__name__}: {exc}"})

    def do_POST(self) -> None:
        parsed = urlparse(self.path)
        try:
            payload = self._read_json()
            if parsed.path == "/api/sheet/save":
                data = self.store.save_sheet(str(payload.get("sheet") or ""), payload.get("rows"))
                self._send_json(200, {"ok": True, "data": data})
                return
            if parsed.path == "/api/announcement/save":
                path = self.store.write_announcement(
                    str(payload.get("sheet") or ""),
                    str(payload.get("text") or ""),
                )
                self._send_json(200, {"ok": True, "path": str(path)})
                return
            if parsed.path == "/api/action":
                result = self.runner.run(payload)
                self._send_json(200 if result["ok"] else 422, result)
                return
            if parsed.path == "/api/export":
                export_path = self.store.export_public_workbook()
                self._send_json(
                    200,
                    {
                        "ok": True,
                        "path": str(export_path),
                        "message": "Excel export created. Editing it will not change the database.",
                    },
                )
                return
            self._send_json(404, {"ok": False, "error": "API endpoint not found."})
        except FrontendError as exc:
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


