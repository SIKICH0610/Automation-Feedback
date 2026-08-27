from __future__ import annotations

import argparse
import json
from datetime import datetime
from dataclasses import dataclass, replace
import os
from pathlib import Path
import re
import shutil
import subprocess
import sys
import time
from typing import Any
from urllib.parse import quote

from openpyxl import load_workbook

from feedback_common import (
    DEFAULT_SHEET,
    DEFAULT_WORKBOOK,
    FEEDBACK_TYPE_CHOICES,
    StudentRow,
    normalize_uid,
    student_from_worksheet,
)
from paste_comment import comment_payload_for_student
from paste_mass_notification import mass_notification_payload_for_student, resolve_mass_message
from paste_attachments import normalized_attachment_paths, stage_attachments


if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")
if hasattr(sys.stderr, "reconfigure"):
    sys.stderr.reconfigure(encoding="utf-8")


DEFAULT_WECOM_TITLE_RE = r".*(WeCom|企业微信|WXWork).*"
DEFAULT_WHATSAPP_TITLE_RE = r".*(WhatsApp|WhatsApp Web).*"
STATUS_COLUMN = "Send Status"
ERROR_COLUMN = "Send Error"
LAST_ATTEMPT_COLUMN = "Last Attempt"
PREFERRED_CHANNEL_COLUMN = "Preferred Channel"
WHATSAPP_PHONE_COLUMN = "WhatsApp Phone"
WHATSAPP_SEARCH_KEY_COLUMN = "WhatsApp Search Key"
WHATSAPP_TARGET_TYPE_COLUMN = "WhatsApp Target Type"
PASTE_ACTION_CHOICES = ("comment", "mass-notification", "check-group-chat")
GROUP_CHAT_COLUMN = "Group Chat"


@dataclass
class DesktopAppSpec:
    key: str
    display_name: str
    title_re: str
    process_names: tuple[str, ...]
    env_path_var: str
    exe_names: tuple[str, ...]
    common_paths: tuple[str, ...]
    uri: str | None = None
    mac_process_name: str | None = None


@dataclass
class DesktopAppStatus:
    key: str
    display_name: str
    dependency_ok: bool
    process_running: bool
    window_found: bool
    launch_attempted: bool = False
    launched: bool = False
    message: str = ""
    dependency_error: str = ""


@dataclass
class PasteJob:
    excel_row: int
    uid: str
    student_name: str
    parent_language: str
    language_explicit: bool
    channel: str
    channel_explicit: bool
    search_key: str
    expected_chat_name: str
    whatsapp_phone: str
    whatsapp_target_type: str
    feedback: str
    action: str = "comment"


@dataclass
class JobResult:
    status: str
    pasted: bool = False
    error: str = ""


@dataclass
class JobLoadError:
    row_number: int
    status: str
    error: str


def build_app_specs(
    *,
    wecom_title_re: str,
    whatsapp_title_re: str,
) -> dict[str, DesktopAppSpec]:
    local_app_data = os.environ.get("LOCALAPPDATA", "")
    program_files = os.environ.get("ProgramFiles", r"C:\Program Files")
    program_files_x86 = os.environ.get("ProgramFiles(x86)", r"C:\Program Files (x86)")

    return {
        "wecom": DesktopAppSpec(
            key="wecom",
            display_name="WeCom / 企业微信",
            title_re=wecom_title_re,
            process_names=("WXWork.exe", "WeCom.exe"),
            env_path_var="WECOM_EXE",
            exe_names=("WXWork.exe", "WeCom.exe"),
            common_paths=(
                rf"{program_files_x86}\Tencent\WeCom\WXWork.exe",
                rf"{program_files}\Tencent\WeCom\WXWork.exe",
                rf"{local_app_data}\Tencent\WXWork\WXWork.exe",
                rf"{local_app_data}\WXWork\WXWork.exe",
            ),
            mac_process_name="企业微信",
        ),
        "whatsapp": DesktopAppSpec(
            key="whatsapp",
            display_name="WhatsApp / WhatsApp Web",
            title_re=whatsapp_title_re,
            process_names=("WhatsApp.exe", "chrome.exe", "msedge.exe", "brave.exe", "firefox.exe"),
            env_path_var="WHATSAPP_EXE",
            exe_names=("WhatsApp.exe",),
            common_paths=(
                rf"{local_app_data}\WhatsApp\WhatsApp.exe",
                rf"{program_files}\WindowsApps\WhatsApp.exe",
            ),
            uri="https://web.whatsapp.com/",
            mac_process_name="WhatsApp",
        ),
    }


def header_values(worksheet: Any) -> list[Any]:
    return [worksheet.cell(1, col).value for col in range(1, worksheet.max_column + 1)]


def automation_dependency_status() -> tuple[bool, str]:
    if sys.platform == "darwin":
        try:
            import pyperclip  # noqa: F401
        except Exception as exc:
            return False, f"{type(exc).__name__}: {exc}"
        if shutil.which("osascript") is None:
            return False, "osascript was not found on this Mac."
        return True, ""

    try:
        import pywinauto  # noqa: F401
        import pyperclip  # noqa: F401
    except Exception as exc:
        return False, f"{type(exc).__name__}: {exc}"
    return True, ""


def process_is_running(process_names: tuple[str, ...], *, mac_process_name: str | None = None) -> bool:
    if sys.platform == "darwin":
        if not mac_process_name:
            return False
        try:
            result = subprocess.run(["pgrep", "-x", mac_process_name], capture_output=True, text=True, timeout=5)
        except Exception:
            return False
        return result.returncode == 0

    try:
        output = subprocess.check_output(
            ["tasklist"],
            text=True,
            encoding="utf-8",
            errors="ignore",
        )
    except Exception:
        return False

    output_lower = output.lower()
    return any(process_name.lower() in output_lower for process_name in process_names)


def app_window_found(title_re: str, *, mac_process_name: str | None = None) -> bool:
    if sys.platform == "darwin":
        if not mac_process_name or not process_is_running((), mac_process_name=mac_process_name):
            return False
        # The statements have to live inside run()'s body -- inlining them after a
        # `return` is a JavaScript syntax error, which silently made this report "no
        # window" for every app it was ever asked about.
        script = (
            "function run() {"
            '  const se = Application("System Events");'
            f"  const procs = se.processes.whose({{name: {json.dumps(mac_process_name)}}});"
            "  return procs.length > 0 && procs[0].windows.length > 0;"
            "}"
        )
        try:
            result = subprocess.run(
                ["osascript", "-l", "JavaScript", "-e", script],
                capture_output=True,
                text=True,
                timeout=10,
            )
        except Exception:
            return False
        return result.returncode == 0 and result.stdout.strip() == "true"

    return find_window_by_title_re(title_re) is not None


def find_window_by_title_re(title_re: str) -> Any | None:
    try:
        from pywinauto import Desktop
    except Exception:
        return None

    desktop = Desktop(backend="uia")
    pattern = re.compile(title_re)
    try:
        for window in desktop.windows():
            try:
                title = window.window_text().strip()
            except Exception:
                continue
            if title and pattern.search(title):
                return window
    except Exception:
        pass

    try:
        window = desktop.window(title_re=title_re)
        if window.exists(timeout=1):
            return window
    except Exception:
        return None
    return None


def app_status(spec: DesktopAppSpec) -> DesktopAppStatus:
    dependency_ok, dependency_error = automation_dependency_status()
    running = process_is_running(spec.process_names, mac_process_name=spec.mac_process_name)
    window_found = (
        app_window_found(spec.title_re, mac_process_name=spec.mac_process_name) if dependency_ok else False
    )

    if not dependency_ok:
        message = "desktop automation dependencies are not available"
    elif window_found:
        message = "window found"
    elif running:
        message = "process is running, but matching window was not found"
    else:
        message = "not running"

    return DesktopAppStatus(
        key=spec.key,
        display_name=spec.display_name,
        dependency_ok=dependency_ok,
        process_running=running,
        window_found=window_found,
        message=message,
        dependency_error=dependency_error,
    )


def launch_candidates(spec: DesktopAppSpec, configured_path: Path | None) -> list[str]:
    candidates: list[str] = []
    if configured_path:
        candidates.append(str(configured_path))

    env_path = os.environ.get(spec.env_path_var, "").strip()
    if env_path:
        candidates.append(env_path)

    for exe_name in spec.exe_names:
        found = shutil.which(exe_name)
        if found:
            candidates.append(found)

    candidates.extend(spec.common_paths)
    return candidates


def launch_app(spec: DesktopAppSpec, configured_path: Path | None = None) -> bool:
    if sys.platform == "darwin":
        if spec.mac_process_name:
            result = subprocess.run(["open", "-a", spec.mac_process_name], capture_output=True, text=True)
            if result.returncode == 0:
                return True
        if spec.uri:
            result = subprocess.run(["open", spec.uri], capture_output=True, text=True)
            return result.returncode == 0
        return False

    for candidate in launch_candidates(spec, configured_path):
        candidate_path = Path(candidate).expanduser()
        if not candidate_path.exists():
            continue
        subprocess.Popen([str(candidate_path)])
        return True

    if spec.uri:
        try:
            os.startfile(spec.uri)  # type: ignore[attr-defined]
            return True
        except OSError:
            return False

    return False


def ensure_app_available(
    spec: DesktopAppSpec,
    *,
    configured_path: Path | None = None,
    open_if_missing: bool = False,
    settle_seconds: float = 3,
) -> DesktopAppStatus:
    status = app_status(spec)
    if status.window_found or not open_if_missing:
        return status
    if not status.dependency_ok:
        return status

    launched = launch_app(spec, configured_path)
    status.launch_attempted = True
    status.launched = launched
    if not launched:
        status.message = (
            f"could not launch automatically; open {spec.display_name} manually or set "
            f"{spec.env_path_var} / pass an explicit exe path"
        )
        return status

    time.sleep(settle_seconds)
    refreshed = app_status(spec)
    refreshed.launch_attempted = True
    refreshed.launched = launched
    if not refreshed.window_found:
        refreshed.message = "launch attempted, but matching window was not found yet"
    return refreshed


def print_app_status(status: DesktopAppStatus) -> None:
    print(f"{status.display_name}: {status.message}")
    print(f"  dependency ok: {'yes' if status.dependency_ok else 'no'}")
    if status.dependency_error:
        print(f"  dependency error: {status.dependency_error}")
        print("  install command: py -3.11 -m pip install -r requirements.txt")
    print(f"  process running: {'yes' if status.process_running else 'no'}")
    print(f"  window found: {'yes' if status.window_found else 'no'}")
    if status.launch_attempted:
        print(f"  launch attempted: yes")
        print(f"  launch command accepted: {'yes' if status.launched else 'no'}")


def print_matching_window_titles(title_re: str) -> None:
    try:
        from pywinauto import Desktop
    except Exception as exc:
        print(f"Could not inspect windows: {type(exc).__name__}: {exc}")
        return

    pattern = re.compile(title_re)
    found = False
    for window in Desktop(backend="uia").windows():
        try:
            title = window.window_text().strip()
        except Exception:
            continue
        if title and pattern.search(title):
            found = True
            print(f"Matched window title: {title}")

    if not found:
        print(f"No top-level windows matched: {title_re}")


def value_for(student: StudentRow, column_name: str) -> str:
    return str(student.values.get(column_name) or "").strip()


def normalize_channel(value: str) -> str:
    text = value.strip().lower()
    if text in {"wechat", "wecom", "wxwork", "企业微信"}:
        return "wecom"
    if text in {"whatsapp", "whatapp", "wa"}:
        return "whatsapp"
    return ""


def choose_channel(student: StudentRow) -> str:
    configured_channel = normalize_channel(value_for(student, PREFERRED_CHANNEL_COLUMN))
    if configured_channel:
        return configured_channel

    language = value_for(student, "Parent Language").lower()
    if language.startswith("chinese"):
        return "wecom"
    return "whatsapp"


def normalize_whatsapp_phone(phone: str) -> str:
    return "".join(character for character in phone if character.isdigit())


def whatsapp_target_type(student: StudentRow) -> str:
    configured_type = value_for(student, WHATSAPP_TARGET_TYPE_COLUMN).strip().lower()
    if configured_type in {"phone", "direct", "direct_phone"}:
        return "phone"
    if configured_type in {"group", "group_search", "search"}:
        return "group_search"
    if value_for(student, WHATSAPP_PHONE_COLUMN):
        return "phone"
    return "group_search"


def build_expected_chat_name(student: StudentRow, uid: str) -> str:
    configured_name = value_for(student, "WeCom Group Name")
    if configured_name:
        return configured_name
    return f"{student.full_name} - {uid}".strip()


def build_search_key(student: StudentRow, channel: str, uid: str) -> str:
    if channel == "whatsapp":
        return value_for(student, WHATSAPP_SEARCH_KEY_COLUMN) or uid
    return uid


def clear_clipboard() -> None:
    try:
        import pyperclip

        pyperclip.copy("")
    except Exception:
        pass


def copy_text_to_clipboard(text: str, *, description: str) -> None:
    try:
        import pyperclip
    except ImportError as exc:
        raise RuntimeError("Install clipboard dependency first: python -m pip install pyperclip") from exc

    for attempt in range(3):
        pyperclip.copy(text)
        time.sleep(0.1)
        if pyperclip.paste() == text:
            return
        if attempt == 2:
            raise RuntimeError(f"Could not copy {description} into the clipboard.")


def timestamp() -> str:
    return datetime.now().isoformat(timespec="seconds")


def ensure_status_columns(workbook_path: Path, sheet_name: str) -> None:
    workbook = load_workbook(workbook_path)
    if sheet_name not in workbook.sheetnames:
        workbook.close()
        return

    worksheet = workbook[sheet_name]
    headers = header_values(worksheet)
    changed = False
    for column_name in (STATUS_COLUMN, ERROR_COLUMN, LAST_ATTEMPT_COLUMN):
        if column_name not in headers:
            worksheet.cell(1, worksheet.max_column + 1).value = column_name
            headers.append(column_name)
            changed = True

    if changed:
        workbook.save(workbook_path)
    workbook.close()


def set_status_value(worksheet: Any, headers: list[Any], row_number: int, column_name: str, value: str) -> None:
    if column_name not in headers:
        worksheet.cell(1, len(headers) + 1).value = column_name
        headers.append(column_name)
    worksheet.cell(row_number, headers.index(column_name) + 1).value = value


def write_job_status(
    workbook_path: Path,
    sheet_name: str,
    row_number: int,
    *,
    status: str,
    error: str = "",
) -> None:
    workbook = load_workbook(workbook_path)
    worksheet = workbook[sheet_name]
    headers = header_values(worksheet)
    set_status_value(worksheet, headers, row_number, STATUS_COLUMN, status)
    set_status_value(worksheet, headers, row_number, ERROR_COLUMN, error)
    set_status_value(worksheet, headers, row_number, LAST_ATTEMPT_COLUMN, timestamp())
    workbook.save(workbook_path)
    workbook.close()


def write_group_chat_status(
    workbook_path: Path,
    sheet_name: str,
    row_number: int,
    *,
    found: bool,
    parent_language: str | None = None,
) -> None:
    workbook = load_workbook(workbook_path)
    worksheet = workbook[sheet_name]
    headers = header_values(worksheet)
    set_status_value(worksheet, headers, row_number, GROUP_CHAT_COLUMN, found)
    if parent_language:
        set_status_value(worksheet, headers, row_number, "Parent Language", parent_language)
    workbook.save(workbook_path)
    workbook.close()


def payload_for_student(
    student: StudentRow,
    *,
    class_review: str,
    message_column: str = "Feedback",
    feedback_type: str,
    action: str,
    mass_message: str,
) -> str:
    if action == "mass-notification":
        return mass_notification_payload_for_student(student, mass_message=mass_message)
    if action == "check-group-chat":
        return ""
    return comment_payload_for_student(
        student,
        class_review=class_review,
        feedback_type=feedback_type,
        message_column=message_column,
    )


def build_paste_job(
    student: StudentRow,
    *,
    class_review: str,
    message_column: str = "Feedback",
    feedback_type: str,
    action: str,
    mass_message: str,
) -> PasteJob:
    uid = normalize_uid(student.values.get("uid"))
    if not uid:
        raise ValueError(f"Row {student.excel_row} does not have a uid.")

    channel = choose_channel(student)
    phone = normalize_whatsapp_phone(value_for(student, WHATSAPP_PHONE_COLUMN))
    target_type = whatsapp_target_type(student)

    if channel == "whatsapp" and target_type == "phone" and not phone:
        raise ValueError(
            f"Row {student.excel_row} uses WhatsApp phone mode but does not have {WHATSAPP_PHONE_COLUMN!r}."
        )

    return PasteJob(
        excel_row=student.excel_row,
        uid=uid,
        student_name=student.full_name,
        parent_language=value_for(student, "Parent Language") or "English",
        language_explicit=bool(value_for(student, "Parent Language")),
        channel=channel,
        channel_explicit=bool(normalize_channel(value_for(student, PREFERRED_CHANNEL_COLUMN))),
        search_key=build_search_key(student, channel, uid),
        expected_chat_name=build_expected_chat_name(student, uid),
        whatsapp_phone=phone,
        whatsapp_target_type=target_type,
        feedback=payload_for_student(
            student,
            class_review=class_review,
            message_column=message_column,
            feedback_type=feedback_type,
            action=action,
            mass_message=mass_message,
        ),
        action=action,
    )


def parse_row_numbers(row_spec: str) -> list[int]:
    rows: list[int] = []
    for part in row_spec.split(","):
        item = part.strip()
        if not item:
            continue
        if "-" in item:
            start_text, end_text = item.split("-", 1)
            start = int(start_text.strip())
            end = int(end_text.strip())
            if end < start:
                raise ValueError(f"Invalid row range {item!r}: end row is before start row.")
            rows.extend(range(start, end + 1))
        else:
            rows.append(int(item))

    if not rows:
        raise ValueError("--rows did not contain any row numbers.")
    return rows


def selected_row_numbers(args: argparse.Namespace) -> list[int]:
    if args.rows:
        return parse_row_numbers(args.rows)

    if args.start_row is not None or args.end_row is not None:
        start = args.start_row if args.start_row is not None else args.row
        end = args.end_row if args.end_row is not None else start
        if end < start:
            raise ValueError("--end-row cannot be before --start-row.")
        return list(range(start, end + 1))

    return [args.row]


def load_jobs(
    workbook_path: Path,
    *,
    sheet_name: str,
    row_numbers: list[int],
    class_review_zh: str,
    class_review_en: str,
    message_column: str = "Feedback",
    feedback_type: str,
    action: str,
    mass_message: str,
) -> list[PasteJob | JobLoadError]:
    workbook = load_workbook(workbook_path, read_only=True)
    if sheet_name not in workbook.sheetnames:
        available = ", ".join(workbook.sheetnames)
        raise ValueError(f"Sheet {sheet_name!r} not found. Available sheets: {available}")

    worksheet = workbook[sheet_name]
    headers = header_values(worksheet)
    jobs: list[PasteJob | JobLoadError] = []
    for row_number in row_numbers:
        student = student_from_worksheet(worksheet, headers, row_number)
        if not student:
            jobs.append(
                JobLoadError(
                    row_number=row_number,
                    status="needs_review",
                    error=f"Row {row_number} does not look like a student row.",
                )
            )
            continue

        try:
            student_class_review = (
                class_review_zh if student.language.lower().startswith("chinese") else class_review_en
            )
            jobs.append(
                build_paste_job(
                    student,
                    class_review=student_class_review,
                    message_column=message_column,
                    feedback_type=feedback_type,
                    action=action,
                    mass_message=mass_message,
                )
            )
        except ValueError as exc:
            status = "skipped_absent" if "absent" in str(exc).lower() else "needs_review"
            jobs.append(JobLoadError(row_number=row_number, status=status, error=str(exc)))

    workbook.close()
    return jobs


class WeComPasteRobot:
    def __init__(
        self,
        *,
        title_re: str = DEFAULT_WECOM_TITLE_RE,
        search_shortcut: str = "^f",
        settle_seconds: float = 0.8,
    ) -> None:
        try:
            from pywinauto import Desktop
            from pywinauto.keyboard import send_keys
        except ImportError as exc:
            raise RuntimeError(
                "Install desktop automation dependencies first: "
                "python -m pip install pywinauto pyperclip"
            ) from exc

        self.desktop = Desktop(backend="uia")
        self.send_keys = send_keys
        self.title_re = title_re
        self.search_shortcut = search_shortcut
        self.settle_seconds = settle_seconds
        self.window: Any | None = None

    def current_window(self) -> Any:
        if self.window is not None:
            try:
                if self.window.exists(timeout=0.5):
                    return self.window
            except Exception:
                self.window = None

        window = self.desktop.window(title_re=self.title_re)
        self.window = window
        return window

    def focus_window(self) -> Any:
        window = self.current_window()
        window.set_focus()
        return window

    def visible_text(self, *, ignored_edit_values: set[str] | None = None) -> str:
        window = self.current_window()
        ignored_edit_values = ignored_edit_values or set()
        texts: list[str] = []
        try:
            title = window.window_text().strip()
        except Exception:
            title = ""
        if title:
            texts.append(title)

        try:
            import win32gui

            foreground_title = win32gui.GetWindowText(win32gui.GetForegroundWindow()).strip()
        except Exception:
            foreground_title = ""
        if foreground_title and foreground_title not in texts:
            texts.append(f"foreground window: {foreground_title}")

        for element in window.descendants():
            try:
                text = element.window_text().strip()
                control_type = str(element.element_info.control_type or "")
                class_name = str(element.element_info.class_name or "")
            except Exception:
                continue
            is_edit = control_type == "Edit" or "Edit" in class_name
            if is_edit and text in ignored_edit_values:
                continue
            if text:
                texts.append(text)
        return "\n".join(texts)

    def search_chat(self, search_key: str) -> None:
        self.focus_window()
        self.send_keys(self.search_shortcut)
        time.sleep(0.3)
        self.send_keys("^a")
        self.send_keys("{BACKSPACE}")
        copy_text_to_clipboard(search_key, description="the WeCom search key")
        self.send_keys("^v")
        time.sleep(self.settle_seconds)

    def search_result_candidates(self, job: PasteJob) -> list[tuple[int, Any, str]]:
        window = self.current_window()
        name_parts = [part.lower() for part in job.student_name.split() if part]
        expected_chat_name = job.expected_chat_name.lower()
        elements: list[tuple[Any, str, Any]] = []
        candidates: list[tuple[int, Any, str]] = []

        for element in window.descendants():
            try:
                text = element.window_text().strip()
                rectangle = element.rectangle()
            except Exception:
                continue
            if not text or text == job.search_key:
                continue
            if rectangle.width() < 20 or rectangle.height() < 8:
                continue
            elements.append((element, text, rectangle))

        elements.sort(key=lambda item: (item[2].top, item[2].left))
        bands: list[list[tuple[Any, str, Any]]] = []
        for item in elements:
            rectangle = item[2]
            if not bands:
                bands.append([item])
                continue

            previous_bottom = max(existing[2].bottom for existing in bands[-1])
            if rectangle.top - previous_bottom > 24:
                bands.append([item])
            else:
                bands[-1].append(item)

        for band in bands:
            combined_text = "\n".join(item[1] for item in band)
            combined_lower = combined_text.lower()
            uid_ok = bool(job.uid and job.uid in combined_text)
            expected_ok = bool(expected_chat_name and expected_chat_name in combined_lower)
            name_hits = sum(1 for part in name_parts if part in combined_lower)

            has_identity_match = expected_ok or uid_ok
            if not has_identity_match:
                continue

            score = (120 if expected_ok else 0) + (100 if uid_ok else 0) + (30 * name_hits)

            def element_score(item: tuple[Any, str, Any]) -> int:
                _, text, rectangle = item
                lowered = text.lower()
                identity_score = (100 if job.uid and job.uid in text else 0) + (
                    30 * sum(1 for part in name_parts if part in lowered)
                )
                return identity_score + min(rectangle.width() * rectangle.height(), 5000)

            element, _, _ = max(band, key=element_score)
            candidates.append((score, element, combined_text.replace("\n", " | ")))

        candidates.sort(key=lambda candidate: candidate[0], reverse=True)
        return candidates

    def search_input_element(self, search_key: str) -> Any | None:
        window = self.current_window()
        for element in window.descendants():
            try:
                control_type = str(element.element_info.control_type or "")
                class_name = str(element.element_info.class_name or "")
                text = element.window_text().strip()
            except Exception:
                continue

            is_edit = control_type == "Edit" or "Edit" in class_name
            if is_edit and text == search_key:
                return element
        return None

    def search_input_contains(self, search_key: str) -> bool:
        return self.search_input_element(search_key) is not None

    def element_has_keyboard_focus(self, element: Any) -> bool:
        try:
            return bool(element.has_keyboard_focus())
        except Exception:
            pass
        try:
            return bool(element.element_info.has_keyboard_focus)
        except Exception:
            return False

    def search_is_active(self, job: PasteJob) -> bool:
        search_input = self.search_input_element(job.search_key)
        return search_input is not None and self.element_has_keyboard_focus(search_input)

    def clear_search_state(self, job: PasteJob) -> None:
        try:
            search_input = self.search_input_element(job.search_key)
            if search_input is not None:
                try:
                    search_input.set_focus()
                    time.sleep(0.1)
                except Exception:
                    pass
            else:
                self.focus_window()
                self.send_keys(self.search_shortcut)
                time.sleep(0.2)
            self.send_keys("^a")
            time.sleep(0.05)
            self.send_keys("{BACKSPACE}")
            time.sleep(0.1)
            self.send_keys("{ESC}")
            time.sleep(self.settle_seconds)
        except Exception:
            pass

    def open_best_result_with_uia(self, job: PasteJob) -> bool:
        for _, element, text in self.search_result_candidates(job):
            print(f"Trying UI Automation result: {text[:80]}")
            try:
                element.invoke()
                time.sleep(self.settle_seconds)
                if not self.search_is_active(job):
                    return True
            except Exception:
                pass

        return False

    def open_first_result_with_keyboard(self) -> None:
        self.send_keys("{DOWN}")
        time.sleep(0.15)
        self.send_keys("{ENTER}")
        time.sleep(self.settle_seconds)

    def wait_for_search_result_candidates(
        self,
        job: PasteJob,
        *,
        timeout: float = 3.0,
    ) -> list[tuple[int, Any, str]]:
        deadline = time.monotonic() + timeout
        while True:
            candidates = self.search_result_candidates(job)
            if candidates or time.monotonic() >= deadline:
                return candidates
            time.sleep(0.2)

    def print_search_result_candidates(self, job: PasteJob) -> None:
        self.search_chat(job.search_key)
        candidates = self.wait_for_search_result_candidates(job)
        if not candidates:
            print("No matching group-chat candidates found in the search results.")
            return
        print("Safe WeCom search candidates:")
        for score, element, text in candidates[:10]:
            print(f"- score={score}; text={text[:120]}")

    def open_chat_from_search(self, job: PasteJob, *, open_strategy: str = "enter-first") -> None:
        self.search_chat(job.search_key)

        if open_strategy == "manual-click":
            input("Click the correct WeCom search result/chat, then press Enter here to close search and verify. ")
            self.send_keys("{ESC}")
            time.sleep(self.settle_seconds)
            return

        candidates = self.search_result_candidates(job)
        if candidates:
            print(f"Matched WeCom search result: {candidates[0][2][:120]}")
        else:
            print("WeCom result was not exposed through UI Automation; opening with Enter.")
        if open_strategy == "keyboard":
            self.open_first_result_with_keyboard()
        elif open_strategy == "ui-control":
            if not self.open_best_result_with_uia(job):
                self.clear_search_state(job)
                raise LookupError(f"WeCom group chat could not be opened for uid {job.uid}.")
        else:
            print("Opening first WeCom search result with Enter.")
            self.send_keys("{ENTER}")
            time.sleep(self.settle_seconds)

        if self.search_is_active(job):
            print("WeCom search is still active; pressing Enter again.")
            self.send_keys("{ENTER}")
            time.sleep(self.settle_seconds)

        if self.search_is_active(job):
            self.send_keys("{ESC}")
            time.sleep(0.2)
            verified, _ = self.verify_chat(job)
            if not verified:
                self.clear_search_state(job)
                raise LookupError(f"WeCom group chat could not be opened for uid {job.uid}.")

    def verify_chat(self, job: PasteJob) -> tuple[bool, str]:
        text = self.visible_text(ignored_edit_values={job.search_key})
        uid_ok = job.uid in text
        name_parts = [part for part in job.student_name.split() if part]
        name_ok = any(part in text for part in name_parts)

        if uid_ok and name_ok:
            return True, "verified_uid_and_name"
        if uid_ok:
            return True, "verified_uid_only"
        if name_ok:
            return True, "verified_name_only"
        return False, "uid_and_name_not_visible_after_search"

    def print_visible_text_debug(self, *, limit: int = 2000) -> None:
        text = self.visible_text()
        print("Visible WeCom text sample:")
        print(text[:limit] if text else "(no visible text captured)")

    def focus_message_input(self) -> bool:
        window = self.current_window()
        candidates: list[tuple[int, Any]] = []

        for element in window.descendants():
            try:
                rectangle = element.rectangle()
                control_type = str(element.element_info.control_type or "")
                class_name = str(element.element_info.class_name or "")
                text = element.window_text().strip()
            except Exception:
                continue

            is_edit = control_type == "Edit" or "Edit" in class_name
            if not is_edit:
                continue
            if text and text == self.search_shortcut:
                continue
            if rectangle.width() < 100 or rectangle.height() < 20:
                continue

            # The message composer is usually the lowest sizeable edit control.
            candidates.append((rectangle.top, element))

        if candidates:
            _, element = max(candidates, key=lambda candidate: candidate[0])
            try:
                element.set_focus()
                time.sleep(0.2)
                return True
            except Exception:
                pass

        return False

    def paste_feedback(self, feedback: str) -> None:
        copy_text_to_clipboard(feedback, description="the WeCom feedback")
        print(f"Clipboard loaded with feedback ({len(feedback)} characters).")
        self.send_keys("^v")
        time.sleep(1.2)


class WhatsAppPasteRobot:
    def __init__(
        self,
        *,
        title_re: str = DEFAULT_WHATSAPP_TITLE_RE,
        search_shortcut: str = "^f",
        settle_seconds: float = 1.0,
    ) -> None:
        try:
            from pywinauto import Desktop
            from pywinauto.keyboard import send_keys
        except ImportError as exc:
            raise RuntimeError(
                "Install desktop automation dependencies first: "
                "python -m pip install pywinauto pyperclip"
            ) from exc

        self.desktop = Desktop(backend="uia")
        self.send_keys = send_keys
        self.title_re = title_re
        self.search_shortcut = search_shortcut
        self.settle_seconds = settle_seconds
        self.window: Any | None = None

    def current_window(self) -> Any:
        if self.window is not None:
            try:
                if self.window.exists(timeout=0.5):
                    return self.window
            except Exception:
                self.window = None

        window = find_window_by_title_re(self.title_re)
        if window is None:
            raise RuntimeError(f"Could not find WhatsApp window with title pattern: {self.title_re}")
        self.window = window
        return window

    def focus_window(self) -> Any:
        window = self.current_window()
        window.set_focus()
        return window

    def visible_text(self, *, ignored_edit_values: set[str] | None = None) -> str:
        window = self.current_window()
        ignored_edit_values = ignored_edit_values or set()
        texts: list[str] = []
        try:
            title = window.window_text().strip()
        except Exception:
            title = ""
        if title:
            texts.append(title)
        for element in window.descendants():
            try:
                text = element.window_text().strip()
                control_type = str(element.element_info.control_type or "")
                class_name = str(element.element_info.class_name or "")
            except Exception:
                continue
            is_edit = control_type == "Edit" or "Edit" in class_name
            if is_edit and text in ignored_edit_values:
                continue
            if text:
                texts.append(text)
        return "\n".join(texts)

    def open_phone_url(self, job: PasteJob) -> None:
        if not job.whatsapp_phone:
            raise RuntimeError("WhatsApp phone mode needs a WhatsApp Phone value.")

        url = f"https://wa.me/{job.whatsapp_phone}?text={quote(job.feedback)}"
        os.startfile(url)  # type: ignore[attr-defined]
        time.sleep(3)

    def search_chat(self, search_key: str) -> None:
        self.focus_window()
        self.send_keys("{ESC}")
        time.sleep(0.1)
        self.send_keys(self.search_shortcut)
        time.sleep(0.3)
        self.send_keys("^a")
        self.send_keys("{BACKSPACE}")
        copy_text_to_clipboard(search_key, description="the WhatsApp search key")
        self.send_keys("^v")
        time.sleep(self.settle_seconds)

    def search_input_element(self, search_key: str) -> Any | None:
        window = self.current_window()
        for element in window.descendants():
            try:
                control_type = str(element.element_info.control_type or "")
                class_name = str(element.element_info.class_name or "")
                text = element.window_text().strip()
            except Exception:
                continue

            is_edit = control_type == "Edit" or "Edit" in class_name
            if is_edit and text == search_key:
                return element
        return None

    def search_input_contains(self, search_key: str) -> bool:
        return self.search_input_element(search_key) is not None

    def element_has_keyboard_focus(self, element: Any) -> bool:
        try:
            return bool(element.has_keyboard_focus())
        except Exception:
            pass
        try:
            return bool(element.element_info.has_keyboard_focus)
        except Exception:
            return False

    def search_is_active(self, search_key: str) -> bool:
        search_input = self.search_input_element(search_key)
        return search_input is not None and self.element_has_keyboard_focus(search_input)

    def clear_search_state(self, search_key: str) -> None:
        try:
            search_input = self.search_input_element(search_key)
            if search_input is not None:
                try:
                    search_input.set_focus()
                    time.sleep(0.1)
                except Exception:
                    pass
            else:
                self.focus_window()
                self.send_keys(self.search_shortcut)
                time.sleep(0.2)
            self.send_keys("^a")
            time.sleep(0.05)
            self.send_keys("{BACKSPACE}")
            time.sleep(0.1)
            self.send_keys("{ESC}")
            time.sleep(self.settle_seconds)
        except Exception:
            pass

    def search_result_matches(self, job: PasteJob) -> bool:
        text = self.visible_text(ignored_edit_values={job.search_key}).lower()
        search_key = job.search_key.lower()
        uid_ok = bool(job.uid and job.uid.lower() in text)
        search_ok = bool(search_key and search_key in text)
        name_parts = [part.lower() for part in job.student_name.split() if part]
        name_ok = any(part in text for part in name_parts)
        return search_ok or (uid_ok and name_ok)

    def wait_for_search_result(self, job: PasteJob, *, timeout: float = 3.0) -> bool:
        deadline = time.monotonic() + timeout
        while True:
            if self.search_result_matches(job):
                return True
            if time.monotonic() >= deadline:
                return False
            time.sleep(0.2)

    def open_chat_from_search(self, job: PasteJob) -> None:
        self.search_chat(job.search_key)
        matched_result = self.search_result_matches(job)
        if matched_result:
            print("Opening matched WhatsApp search result with Enter.")
        else:
            print("WhatsApp result was not exposed through UI Automation; opening with Enter.")
        self.send_keys("{ENTER}")
        time.sleep(self.settle_seconds)
        if self.search_is_active(job.search_key):
            print("WhatsApp search is still active; pressing Enter again.")
            self.send_keys("{ENTER}")
            time.sleep(self.settle_seconds)

        if self.search_is_active(job.search_key):
            self.send_keys("{ESC}")
            time.sleep(0.2)
            verified, _ = self.verify_chat(job)
            if not verified:
                self.clear_search_state(job.search_key)
                raise LookupError(
                    f"WhatsApp chat could not be opened for search key {job.search_key!r}."
                )

    def close_search_overlay(self) -> None:
        try:
            self.send_keys("{ESC}")
            time.sleep(0.2)
        except Exception:
            pass

    def verify_chat(self, job: PasteJob) -> tuple[bool, str]:
        text = self.visible_text(ignored_edit_values={job.search_key})
        uid_ok = bool(job.uid and job.uid in text)
        search_ok = bool(job.search_key and job.search_key in text)
        name_parts = [part for part in job.student_name.split() if part]
        name_ok = any(part in text for part in name_parts)

        if uid_ok or search_ok:
            return True, "verified_search_key"
        if name_ok:
            return True, "verified_name_only"
        return False, "whatsapp_target_not_visible"

    def paste_feedback(self, feedback: str) -> None:
        copy_text_to_clipboard(feedback, description="the WhatsApp feedback")
        print(f"Clipboard loaded with feedback ({len(feedback)} characters).")
        self.send_keys("^v")
        time.sleep(1.5)


def print_job(job: PasteJob) -> None:
    print("=" * 72)
    print(f"Row: {job.excel_row}")
    print(f"UID: {job.uid}")
    print(f"Student: {job.student_name}")
    print(f"Parent language: {job.parent_language}")
    print(f"Channel: {job.channel}")
    print(f"Action: {job.action}")
    print(f"Search key: {job.search_key}")
    print(f"Expected chat name: {job.expected_chat_name}")
    if job.channel == "whatsapp":
        print(f"WhatsApp target type: {job.whatsapp_target_type}")
        print(f"WhatsApp phone: {job.whatsapp_phone or '(blank)'}")
    print()
    print(job.feedback)
    print()


def configured_exe_for_app(args: argparse.Namespace, app_key: str) -> Path | None:
    if app_key == "wecom":
        return args.wecom_exe
    if app_key == "whatsapp":
        return args.whatsapp_exe
    return None


def print_readiness(job: PasteJob, status: DesktopAppStatus) -> None:
    print_app_status(status)
    ready = status.dependency_ok and status.window_found
    if job.channel == "whatsapp" and job.whatsapp_target_type == "phone":
        print("Safe to run paste-only for this row: yes, phone URL mode does not need app window verification")
    else:
        print(f"Safe to run paste-only for this row: {'yes' if ready else 'no'}")


def build_parser(
    *,
    default_message_column: str = "Feedback",
    default_fallback_channel: bool = False,
) -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="Supervised paste helper for WeCom / 企业微信 and WhatsApp feedback messages. It never sends."
    )
    parser.add_argument("--workbook", type=Path, default=DEFAULT_WORKBOOK)
    parser.add_argument("--sheet", default=DEFAULT_SHEET)
    parser.add_argument("--row", type=int, default=2)
    parser.add_argument(
        "--rows",
        help="Comma-separated rows or ranges to process, such as 3,5 or 3-5.",
    )
    parser.add_argument(
        "--start-row",
        type=int,
        help="First row for a batch range. Use with --end-row.",
    )
    parser.add_argument(
        "--end-row",
        type=int,
        help="Last row for a batch range.",
    )
    parser.add_argument("--class-review", default="")
    parser.add_argument("--class-review-file", type=Path)
    parser.add_argument(
        "--message-column",
        default=default_message_column,
        help="Workbook column containing the exact text to paste. Default is Feedback.",
    )
    parser.add_argument(
        "--fallback-channel",
        action="store_true",
        default=default_fallback_channel,
        help="If an automatically selected channel fails, clear state and try the other app.",
    )
    parser.add_argument("--class-review-file-zh", type=Path)
    parser.add_argument("--class-review-file-en", type=Path)
    parser.add_argument(
        "--feedback-type",
        choices=FEEDBACK_TYPE_CHOICES,
        default="comprehensive",
        help="Feedback type to generate only when a row has no existing Feedback cell.",
    )
    parser.add_argument(
        "--action",
        choices=PASTE_ACTION_CHOICES,
        default="comment",
        help="comment = paste each row's Feedback; mass-notification = paste one shared message to each selected chat.",
    )
    parser.add_argument(
        "--mass-message",
        default="",
        help="Shared message used with --action mass-notification.",
    )
    parser.add_argument(
        "--mass-message-file",
        type=Path,
        help="Text file containing the shared message used with --action mass-notification.",
    )
    parser.add_argument(
        "--attachment",
        action="append",
        type=Path,
        default=[],
        help="Image or document to stage in the chat preview. Repeat for multiple files.",
    )
    parser.add_argument(
        "--status",
        action="store_true",
        help="Check workbook row, needed app, and desktop automation readiness without pasting.",
    )
    parser.add_argument(
        "--open-app",
        action="store_true",
        help="Open the app needed for this row if it is not already available, then stop unless a paste mode is selected.",
    )
    parser.add_argument(
        "--debug-search-results",
        action="store_true",
        help="Search WeCom / 企业微信 and print safe candidate UI elements without opening, pasting, or sending.",
    )
    parser.add_argument(
        "--channel",
        choices=("auto", "wecom", "whatsapp"),
        default="auto",
        help=(
            "For check-group-chat, force which app is searched. auto (default) uses each "
            "row's Preferred Channel / Parent Language and, when neither is set, tries "
            "WeCom then falls back to WhatsApp, filling in Parent Language from whichever "
            "matched. wecom / whatsapp search only that app and never write Parent "
            "Language, since forcing a channel proves nothing about which language a "
            "family uses."
        ),
    )
    parser.add_argument(
        "--check-apps",
        action="store_true",
        help="Print WeCom / WhatsApp window availability as JSON and exit. Touches no workbook.",
    )
    parser.add_argument(
        "--debug-window-titles",
        action="store_true",
        help="Print app window titles that match the selected channel without pasting.",
    )
    parser.add_argument(
        "--mode",
        choices=["dry-run", "paste-only"],
        default="dry-run",
        help="dry-run prints the job only; paste-only searches the selected app, verifies when possible, and pastes without sending.",
    )
    parser.add_argument(
        "--no-auto-open",
        action="store_true",
        help="In paste-only mode, do not try to launch the needed app automatically.",
    )
    parser.add_argument(
        "--auto-open-search-result",
        action="store_true",
        help="After searching WeCom / 企业微信, press Enter only.",
    )
    parser.add_argument(
        "--ui-control-result-open",
        action="store_true",
        help="After searching WeCom / 企业微信, open a matching result with UI Automation.",
    )
    parser.add_argument(
        "--keyboard-result-open",
        action="store_true",
        help="After searching WeCom / 企业微信, use Down then Enter to open the first result.",
    )
    parser.add_argument(
        "--manual-result-click",
        action="store_true",
        help="After searching WeCom / 企业微信, wait for you to click the search result.",
    )
    parser.add_argument(
        "--require-verification",
        action="store_true",
        help="Stop before pasting if the script cannot verify the WeCom chat by visible UI text.",
    )
    parser.add_argument("--wecom-title-re", default=DEFAULT_WECOM_TITLE_RE)
    parser.add_argument("--whatsapp-title-re", default=DEFAULT_WHATSAPP_TITLE_RE)
    parser.add_argument("--wecom-exe", type=Path, help="Optional explicit path to WXWork.exe / WeCom.exe.")
    parser.add_argument("--whatsapp-exe", type=Path, help="Optional explicit path to WhatsApp.exe.")
    parser.add_argument("--search-shortcut", default="^f")
    parser.add_argument(
        "--whatsapp-search-shortcut",
        default="^f",
        help="Keyboard shortcut used to open WhatsApp search. Default is Ctrl+F.",
    )
    parser.add_argument(
        "--no-status-write",
        action="store_true",
        help="Do not write Send Status, Send Error, or Last Attempt back to the workbook.",
    )
    return parser


def open_strategy_from_args(args: argparse.Namespace) -> str:
    if args.manual_result_click:
        return "manual-click"
    if args.auto_open_search_result:
        return "enter"
    if args.ui_control_result_open:
        return "ui-control"
    if args.keyboard_result_open:
        return "keyboard"
    return "enter-first"


def ensure_desktop_ready(
    app_spec: DesktopAppSpec,
    *,
    app_exe: Path | None,
    no_auto_open: bool,
) -> DesktopAppStatus:
    status = ensure_app_available(
        app_spec,
        configured_path=app_exe,
        open_if_missing=not no_auto_open,
    )
    print_app_status(status)
    if not status.dependency_ok:
        raise RuntimeError(
            "Desktop automation dependencies are not available. "
            "Run: python -m pip install -r requirements.txt"
        )
    if not status.window_found:
        raise RuntimeError("Needed app window is not available.")
    return status


# A real WeCom search result adds a result row (avatar + title + subtitle + timestamp)
# above the "Search for mobile number/email online: ..." suggestion that WeCom shows no
# matter what you search, match or not -- so the dropdown is reliably taller when there is
# a real match than when there isn't. That dropdown turns out to be a fixed pixel size
# that does NOT scale with the window (confirmed by resizing a real window and remeasuring:
# a match stayed exactly 309px tall whether the window was 948px or 650px tall), so neither
# a raw pixel threshold nor a fraction-of-window-height threshold is portable across
# different window sizes or, likely, different display DPI scaling on other computers.
# Comparing against a live-measured baseline instead of any hardcoded number sidesteps
# both problems: whatever the current window size or DPI is, it affects the baseline
# (guaranteed no-match) measurement and a real match's measurement the same way, so their
# ratio should hold steady even though the absolute pixel values wouldn't.
WECOM_MATCH_HEIGHT_RATIO = 1.3  # a match's dropdown must be at least this much taller than baseline
# WeCom's dropdown renders differently depending on whether the search text even looks like
# a phone number/uid, independent of whether anything matched (confirmed: a non-numeric
# baseline string measured 135px, a numeric non-matching uid measured 199px, at the same
# window size) -- so the baseline has to be numeric, the same shape as a real uid, or it
# isn't a fair comparison. "0000000" is a 7-digit number matching real uid length/format
# but far outside the range any of this org's real uids have used (which start with 7 or 8).
WECOM_BASELINE_SEARCH_KEY = "0000000"

# Timing for the dropdown measurement. A fixed settle wait long enough for the slowest
# case wastes that time on every single student, so instead the dropdown is polled until
# its measured height stops changing. Measured on a real window: the dropdown settles in
# roughly 0.2-0.3s, against the 0.8s fixed wait this replaces.
WECOM_POLL_INTERVAL_SECONDS = 0.07
WECOM_POLL_STABLE_READINGS = 3  # consecutive identical heights before trusting the value
WECOM_POLL_MAX_SECONDS = 2.0  # cap, so a hung dropdown cannot stall the whole batch

_wecom_baseline_height_cache: int | None = None
# The window box and the "search box is open and focused" state hold for a whole run, so
# they are established once instead of per student. _wecom_reset_search_session() drops
# them when a measurement looks wrong, so the next student re-establishes from scratch.
_wecom_session: dict[str, Any] = {"bbox": None, "search_open": False}


def _wecom_reset_search_session() -> None:
    _wecom_session["search_open"] = False


def _wecom_window_bbox(robot: "WeComPasteRobot") -> tuple[int, int, int, int]:
    if _wecom_session["bbox"] is None:
        rectangle = robot.current_window().rectangle()
        _wecom_session["bbox"] = (
            rectangle.left,
            rectangle.top,
            rectangle.right,
            rectangle.bottom,
        )
    return _wecom_session["bbox"]


def _ensure_wecom_search_open(robot: "WeComPasteRobot") -> None:
    """Put focus in WeCom's search box, once per run rather than once per student.

    After a measurement the box is only cleared, not closed, so it stays focused and the
    next student can type straight into it. Re-focusing and re-opening search for every
    student cost about 0.7s each with nothing to show for it.
    """
    if _wecom_session["search_open"]:
        return
    robot.focus_window()
    robot.send_keys(robot.search_shortcut)
    time.sleep(0.3)
    _clear_wecom_search_box(robot)
    time.sleep(0.15)
    _wecom_session["search_open"] = True


def _wecom_dropdown_height(robot: "WeComPasteRobot", search_key: str) -> int:
    """Type search_key into WeCom's already-focused, already-empty search box and measure
    how tall the resulting dropdown is, in pixels, by diffing screenshots taken before and
    after. Assumes the search box is already open and empty.

    This deliberately does not read any text on screen: UI Automation exposes nothing for
    WeCom (it draws its own UI rather than using real controls), and OCR text matching was
    tried and found unreliable -- WeCom always echoes the raw search term back in its own
    suggestion text regardless of whether anything matched, and the sidebar's many other
    real contacts can coincidentally contain a fragment of whatever name is being checked,
    so text-based checks produced false positives for uids that don't exist at all. Dropdown
    height doesn't depend on reading or matching any text, so it isn't exposed to either
    problem.
    """
    from PIL import ImageChops, ImageGrab

    bbox = _wecom_window_bbox(robot)
    before = ImageGrab.grab(bbox=bbox)
    copy_text_to_clipboard(search_key, description="the WeCom search key")
    robot.send_keys("^v")

    deadline = time.monotonic() + WECOM_POLL_MAX_SECONDS
    height = 0
    last_height: int | None = None
    stable = 0
    while True:
        time.sleep(WECOM_POLL_INTERVAL_SECONDS)
        after = ImageGrab.grab(bbox=bbox)
        box = ImageChops.difference(before.convert("RGB"), after.convert("RGB")).getbbox()
        height = (box[3] - box[1]) if box else 0

        # Height 0 means nothing on screen changed at all, so the dropdown has not drawn
        # yet; only a stable non-zero reading counts as settled.
        if height and height == last_height:
            stable += 1
            if stable >= WECOM_POLL_STABLE_READINGS - 1:
                break
        else:
            stable = 0
        last_height = height

        if time.monotonic() >= deadline:
            break
    return height


def _clear_wecom_search_box(robot: "WeComPasteRobot") -> None:
    robot.send_keys("^a")
    robot.send_keys("{BACKSPACE}")


def _wecom_no_match_baseline(robot: "WeComPasteRobot") -> int:
    """The dropdown height WeCom shows for a uid that certainly doesn't exist, measured
    once per process run and cached -- self-calibrates to whatever this run's actual
    window size and display DPI are, instead of assuming a fixed number will transfer
    from the machine this was developed on to whichever machine actually runs it.
    """
    global _wecom_baseline_height_cache
    if _wecom_baseline_height_cache is None:
        _wecom_baseline_height_cache = _wecom_dropdown_height(robot, WECOM_BASELINE_SEARCH_KEY)
        _clear_wecom_search_box(robot)
    return _wecom_baseline_height_cache


_desktop_ready_cache: set[str] = set()


def ensure_desktop_ready_once(
    app_spec: DesktopAppSpec,
    *,
    app_exe: Path | None,
    no_auto_open: bool,
) -> None:
    """ensure_desktop_ready, but only actually checked once per app per process.

    The readiness check enumerates every desktop window through UI Automation, which
    measured around 0.55s -- affordable once, but not once per student in a batch of
    dozens. An app that disappears mid-batch still surfaces: the check path notices a
    measurement where nothing on screen changed and re-establishes from scratch.
    """
    if app_spec.key in _desktop_ready_cache:
        return
    ensure_desktop_ready(app_spec, app_exe=app_exe, no_auto_open=no_auto_open)
    _desktop_ready_cache.add(app_spec.key)


def _run_wecom_job_mac(job: PasteJob, *, args: argparse.Namespace) -> JobResult:
    from wecom_mac import WeComAutomationError, WeComPasteRobotMac

    robot = WeComPasteRobotMac()
    name_parts = [part for part in job.student_name.split() if part]

    try:
        robot.focus_window()
    except WeComAutomationError as exc:
        return JobResult(status="needs_review", error=str(exc))

    try:
        robot.search_and_open_top_result(job.search_key)
    except LookupError as exc:
        robot.clear_search_state()
        return JobResult(status="not_found" if job.action == "check-group-chat" else "needs_review", error=str(exc))

    verified, reason = robot.verify_chat(
        uid=job.uid,
        name_parts=name_parts,
        expected_chat_name=job.expected_chat_name,
    )
    print(f"Verification: {reason}")

    if job.action == "check-group-chat":
        robot.clear_search_state()
        status = "verified" if verified else "not_found"
        print(f"Group chat check: {status}")
        return JobResult(status=status)

    if not verified:
        if args.require_verification:
            robot.clear_search_state()
            raise LookupError("WeCom chat could not be verified.")
        print("WARNING: Could not verify the chat automatically. Continuing because paste-only does not send.")

    if job.feedback:
        robot.paste_feedback(job.feedback)
    if args.attachments:
        print("Attachment paste is not yet supported on macOS; skipping attachments for this row.")
    robot.clear_search_state()
    print("Cleared WeCom search box for the next row.")
    return JobResult(status="pasted", pasted=True)


def run_wecom_job(
    job: PasteJob,
    *,
    args: argparse.Namespace,
    app_spec: DesktopAppSpec,
    app_exe: Path | None,
) -> JobResult:
    if sys.platform == "darwin":
        return _run_wecom_job_mac(job, args=args)

    if job.action == "check-group-chat":
        ensure_desktop_ready_once(app_spec, app_exe=app_exe, no_auto_open=args.no_auto_open)
    else:
        ensure_desktop_ready(app_spec, app_exe=app_exe, no_auto_open=args.no_auto_open)
    robot = WeComPasteRobot(
        title_re=args.wecom_title_re,
        search_shortcut=args.search_shortcut,
    )

    if job.action == "check-group-chat":
        # Always searches by uid only, never by name. Never opens a chat -- just measures
        # the search dropdown's height against a live-measured baseline; see
        # _wecom_dropdown_height and _wecom_no_match_baseline for why.
        _ensure_wecom_search_open(robot)
        baseline = _wecom_no_match_baseline(robot)
        height = _wecom_dropdown_height(robot, job.search_key)
        if height == 0:
            # Nothing on screen changed, so the keystrokes did not reach the search box --
            # focus was lost since the last student. Re-open search and measure again
            # rather than reporting a not_found that never actually got searched.
            print("WeCom search box lost focus; reopening and retrying.")
            _wecom_reset_search_session()
            _ensure_wecom_search_open(robot)
            height = _wecom_dropdown_height(robot, job.search_key)
        _clear_wecom_search_box(robot)
        threshold = baseline * WECOM_MATCH_HEIGHT_RATIO
        verified = height > threshold
        print(
            f"Verification: dropdown height {height}px "
            f"(baseline {baseline}px, threshold {threshold:.0f}px)"
        )
        status = "verified" if verified else "not_found"
        print(f"Group chat check: {status}")
        return JobResult(status=status)

    if args.debug_search_results:
        robot.print_search_result_candidates(job)
        return JobResult(status="checked")

    try:
        robot.open_chat_from_search(
            job,
            open_strategy=open_strategy_from_args(args),
        )
    except Exception:
        robot.clear_search_state(job)
        raise

    verified, reason = robot.verify_chat(job)
    print(f"Verification: {reason}")
    if not verified:
        robot.print_visible_text_debug()
        if args.require_verification:
            raise LookupError("WeCom chat could not be verified.")
        print("WARNING: Could not verify the chat automatically. Continuing because paste-only does not send.")

    if job.feedback:
        robot.paste_feedback(job.feedback)
    if args.attachments:
        stage_attachments(args.attachments, send_keys=robot.send_keys)
    else:
        robot.clear_search_state(job)
        print("Cleared WeCom search box for the next row.")
    return JobResult(status="pasted", pasted=True)


def run_whatsapp_job(
    job: PasteJob,
    *,
    args: argparse.Namespace,
    app_spec: DesktopAppSpec,
    app_exe: Path | None,
) -> JobResult:
    if sys.platform == "darwin":
        # WeCom automation was ported to macOS first since most rows route there;
        # WhatsApp's macOS port (search, verification) is still being worked out --
        # see whatsapp_mac.py for the in-progress version.
        return JobResult(
            status="needs_review",
            error="WhatsApp paste automation is not yet available on macOS. Send this row manually for now.",
        )

    if job.action == "check-group-chat":
        if job.whatsapp_target_type == "phone":
            return JobResult(
                status="needs_review",
                error="Phone-target WhatsApp rows cannot be checked by search; verify manually.",
            )
        if not job.search_key:
            return JobResult(
                status="needs_review",
                error="WhatsApp group_search needs WhatsApp Search Key or uid.",
            )
        ensure_desktop_ready_once(app_spec, app_exe=app_exe, no_auto_open=args.no_auto_open)
        robot = WhatsAppPasteRobot(
            title_re=args.whatsapp_title_re,
            search_shortcut=args.whatsapp_search_shortcut,
        )
        # WhatsApp Desktop/Web is a standard Electron app and (unlike WeCom) does expose
        # real UI Automation text, so this trusts verify_chat's uid/search-key/name check.
        # If this turns out to have the same false-positive risk WeCom's text check did
        # (a name fragment coincidentally matching something else on screen), it should
        # switch to the same dropdown-height approach used for WeCom instead.
        try:
            robot.open_chat_from_search(job)
        except LookupError as exc:
            robot.clear_search_state(job.search_key)
            print(f"Group chat check: not_found ({exc})")
            return JobResult(status="not_found")
        verified, reason = robot.verify_chat(job)
        print(f"Verification: {reason}")
        robot.clear_search_state(job.search_key)
        status = "verified" if verified else "not_found"
        print(f"Group chat check: {status}")
        return JobResult(status=status)

    if job.whatsapp_target_type == "phone":
        if not job.whatsapp_phone:
            raise LookupError("WhatsApp phone mode needs a WhatsApp Phone value.")
        url = f"https://wa.me/{job.whatsapp_phone}"
        if job.feedback:
            url += f"?text={quote(job.feedback)}"
        os.startfile(url)  # type: ignore[attr-defined]
        time.sleep(3)
        if args.attachments:
            ensure_desktop_ready(app_spec, app_exe=app_exe, no_auto_open=args.no_auto_open)
            robot = WhatsAppPasteRobot(
                title_re=args.whatsapp_title_re,
                search_shortcut=args.whatsapp_search_shortcut,
            )
            robot.focus_window()
            stage_attachments(args.attachments, send_keys=robot.send_keys)
        return JobResult(status="pasted", pasted=True)

    if not job.search_key:
        raise LookupError("WhatsApp group_search needs WhatsApp Search Key or uid.")

    ensure_desktop_ready(app_spec, app_exe=app_exe, no_auto_open=args.no_auto_open)
    robot = WhatsAppPasteRobot(
        title_re=args.whatsapp_title_re,
        search_shortcut=args.whatsapp_search_shortcut,
    )
    try:
        robot.open_chat_from_search(job)
    except Exception:
        robot.clear_search_state(job.search_key)
        raise

    verified, reason = robot.verify_chat(job)
    print(f"Verification: {reason}")
    if not verified:
        if args.require_verification:
            raise LookupError("WhatsApp chat could not be verified.")
        print("WARNING: Could not verify the WhatsApp chat automatically. Continuing because paste-only does not send.")

    if job.feedback:
        robot.paste_feedback(job.feedback)
    if args.attachments:
        stage_attachments(args.attachments, send_keys=robot.send_keys)
    else:
        robot.clear_search_state(job.search_key)
        print("Cleared WhatsApp search box for the next row.")

    return JobResult(status="pasted", pasted=True)


def run_job(
    job: PasteJob,
    *,
    args: argparse.Namespace,
    app_specs: dict[str, DesktopAppSpec],
    batch_mode: bool,
) -> JobResult:
    if job.action == "check-group-chat" and args.channel != "auto":
        # Forcing a channel overrides the row's own routing. The search key has to be
        # rebuilt too: a row routed to WhatsApp may carry a WhatsApp Search Key, which is
        # meaningless to WeCom, and vice versa.
        job = replace(
            job,
            channel=args.channel,
            channel_explicit=True,
            search_key=job.search_key if args.channel == "whatsapp" else job.uid,
            whatsapp_target_type="group_search",
        )

    print_job(job)

    if job.channel not in app_specs:
        return JobResult(status="needs_review", error=f"No desktop app setup is defined for channel {job.channel!r}.")

    app_spec = app_specs[job.channel]
    app_exe = configured_exe_for_app(args, job.channel)

    if args.status:
        status = app_status(app_spec)
        print_readiness(job, status)
        return JobResult(status="checked")

    if args.debug_window_titles:
        print_matching_window_titles(app_spec.title_re)
        return JobResult(status="checked")

    if args.open_app:
        status = ensure_app_available(
            app_spec,
            configured_path=app_exe,
            open_if_missing=True,
        )
        print_readiness(job, status)
        if args.mode == "dry-run":
            return JobResult(status="checked")

    if args.mode == "dry-run" and not args.debug_search_results:
        print("Dry run only. Nothing was pasted or sent.")
        return JobResult(status="dry_run")

    def attempt(current_job: PasteJob) -> JobResult:
        current_spec = app_specs[current_job.channel]
        current_exe = configured_exe_for_app(args, current_job.channel)
        clear_clipboard()
        try:
            if current_job.channel == "wecom":
                return run_wecom_job(
                    current_job,
                    args=args,
                    app_spec=current_spec,
                    app_exe=current_exe,
                )
            if current_job.channel == "whatsapp":
                return run_whatsapp_job(
                    current_job,
                    args=args,
                    app_spec=current_spec,
                    app_exe=current_exe,
                )
            return JobResult(
                status="needs_review",
                error=f"Unsupported channel {current_job.channel!r}.",
            )
        except LookupError as exc:
            return JobResult(status="needs_review", error=str(exc))
        except Exception as exc:
            return JobResult(status="failed", error=f"{type(exc).__name__}: {exc}")
        finally:
            time.sleep(0.2)
            clear_clipboard()

    if job.action == "check-group-chat" and not job.channel_explicit and not job.language_explicit:
        # Parent Language is unknown and nothing pins the channel. Most families here are
        # Chinese-speaking, so check WeCom first; only fall back to WhatsApp if that misses.
        wecom_result = attempt(replace(job, channel="wecom", search_key=job.uid))
        if wecom_result.status == "verified":
            return JobResult(status="verified_wecom_chinese")

        whatsapp_search_key = job.search_key if job.channel == "whatsapp" else job.uid
        whatsapp_result = attempt(
            replace(job, channel="whatsapp", search_key=whatsapp_search_key, whatsapp_target_type="group_search")
        )
        if whatsapp_result.status == "verified":
            return JobResult(status="verified_whatsapp_english")

        if wecom_result.status == "failed" and whatsapp_result.status == "failed":
            return JobResult(
                status="needs_review",
                error=(
                    f"wecom: {wecom_result.error or wecom_result.status}; "
                    f"whatsapp: {whatsapp_result.error or whatsapp_result.status}"
                ),
            )
        return JobResult(status="not_found")

    result = attempt(job)
    can_fallback = (
        args.fallback_channel
        and not job.channel_explicit
        and job.whatsapp_target_type != "phone"
        and not result.pasted
        and job.action != "check-group-chat"
    )
    if can_fallback:
        primary_error = result.error or result.status
        fallback_channel = "whatsapp" if job.channel == "wecom" else "wecom"
        print(
            f"Primary {job.channel} attempt did not paste. "
            f"Clearing state and trying {fallback_channel}."
        )
        fallback_job = replace(
            job,
            channel=fallback_channel,
            search_key=job.uid,
            whatsapp_phone="",
            whatsapp_target_type="group_search",
        )
        result = attempt(fallback_job)
        if not result.pasted:
            fallback_error = result.error or result.status
            result.error = (
                f"Primary {job.channel} attempt failed with {primary_error}. "
                f"Fallback {fallback_channel} attempt failed with {fallback_error}."
            )

    if result.pasted and job.action == "mass-notification":
        result.status = "mass_notification_pasted"

    if result.pasted:
        print("Pasted only. The script did not send the message.")
    elif result.error:
        print(f"{result.status}: {result.error}")
    return result


def main(
    *,
    default_message_column: str = "Feedback",
    default_fallback_channel: bool = False,
) -> None:
    args = build_parser(
        default_message_column=default_message_column,
        default_fallback_channel=default_fallback_channel,
    ).parse_args()

    if args.check_apps:
        # Readiness probe only: no workbook, no rows, no automation beyond looking for
        # the app windows. Callers (the frontend's bulk check) run this in its own
        # process because pywinauto/COM is unreliable inside a threaded HTTP server.
        specs = build_app_specs(
            wecom_title_re=args.wecom_title_re,
            whatsapp_title_re=args.whatsapp_title_re,
        )
        report = {
            key: {
                "display_name": status.display_name,
                "window_found": status.window_found,
                "process_running": status.process_running,
                "dependency_ok": status.dependency_ok,
                "message": status.message,
            }
            for key, status in ((k, app_status(spec)) for k, spec in specs.items())
        }
        print(json.dumps(report, ensure_ascii=False))
        return

    args.attachments = normalized_attachment_paths(args.attachment)
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
    mass_message = resolve_mass_message(
        mass_message=args.mass_message,
        mass_message_file=args.mass_message_file,
        fallback_text=class_review,
    )
    if args.action == "mass-notification" and not mass_message and not args.attachments:
        raise SystemExit(
            "Mass notification needs announcement text, at least one attachment, or both."
        )
    if args.attachments:
        print(
            f"Attachment mode: {len(args.attachments)} file(s) will be staged for manual review."
        )

    jobs = load_jobs(
        args.workbook,
        sheet_name=args.sheet,
        row_numbers=selected_row_numbers(args),
        class_review_zh=class_review_zh,
        class_review_en=class_review_en,
        message_column=args.message_column,
        feedback_type=args.feedback_type,
        action=args.action,
        mass_message=mass_message,
    )
    if args.attachments and len(jobs) != 1:
        raise SystemExit(
            "Attachment previews require manual review. Select exactly one row for each run."
        )

    app_specs = build_app_specs(
        wecom_title_re=args.wecom_title_re,
        whatsapp_title_re=args.whatsapp_title_re,
    )

    batch_mode = len(jobs) > 1
    should_write_status = args.mode == "paste-only" and not args.no_status_write
    if should_write_status:
        ensure_status_columns(args.workbook, args.sheet)

    processed = 0
    pasted = 0
    needs_review = 0
    skipped = 0
    failed = 0
    for index, item in enumerate(jobs, start=1):
        if batch_mode:
            print(f"Batch item {index}/{len(jobs)}")

        if isinstance(item, JobLoadError):
            print("=" * 72)
            print(f"Row: {item.row_number}")
            print(f"{item.status}: {item.error}")
            if should_write_status:
                write_job_status(
                    args.workbook,
                    args.sheet,
                    item.row_number,
                    status=item.status,
                    error=item.error,
                )
            if item.status == "skipped_absent":
                skipped += 1
            else:
                needs_review += 1
            clear_clipboard()
            continue

        result = run_job(item, args=args, app_specs=app_specs, batch_mode=batch_mode)
        processed += 1
        if result.pasted:
            pasted += 1
        elif result.status == "needs_review":
            needs_review += 1
        elif result.status == "skipped_absent":
            skipped += 1
        elif result.status == "failed":
            failed += 1

        if (
            should_write_status
            and item.action != "check-group-chat"
            and (result.pasted or result.status in {"needs_review", "skipped_absent", "failed"})
        ):
            write_job_status(
                args.workbook,
                args.sheet,
                item.excel_row,
                status=result.status,
                error=result.error,
            )

        group_chat_statuses = {
            "verified": True,
            "not_found": False,
            "verified_wecom_chinese": True,
            "verified_whatsapp_english": True,
        }
        inferred_language = {
            "verified_wecom_chinese": "Chinese",
            "verified_whatsapp_english": "English",
        }
        if (
            not args.no_status_write
            and item.action == "check-group-chat"
            and result.status in group_chat_statuses
        ):
            write_group_chat_status(
                args.workbook,
                args.sheet,
                item.excel_row,
                found=group_chat_statuses[result.status],
                parent_language=inferred_language.get(result.status),
            )

    if batch_mode:
        print("=" * 72)
        print(
            "Batch complete. "
            f"Processed: {processed}. Pasted: {pasted}. "
            f"Needs review: {needs_review}. Skipped: {skipped}. Failed: {failed}. Sent: 0."
        )


if __name__ == "__main__":
    main()
