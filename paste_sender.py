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
from paste_common import (
    DEFAULT_WECOM_TITLE_RE,
    DEFAULT_WHATSAPP_TITLE_RE,
    DesktopAppSpec,
    DesktopAppStatus,
    JobLoadError,
    JobResult,
    PasteJob,
    clear_clipboard,
    copy_text_to_clipboard,
    find_window_by_title_re,
)

# wecom_win / whatsapp_win reach for pywinauto as soon as a robot is constructed, so
# they are imported inside the Windows job runners rather than here: importing this
# module on a Mac must not require a Windows-only dependency.


if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")
if hasattr(sys.stderr, "reconfigure"):
    sys.stderr.reconfigure(encoding="utf-8")


STATUS_COLUMN = "Send Status"
ERROR_COLUMN = "Send Error"
LAST_ATTEMPT_COLUMN = "Last Attempt"
PREFERRED_CHANNEL_COLUMN = "Preferred Channel"
WHATSAPP_PHONE_COLUMN = "WhatsApp Phone"
WHATSAPP_SEARCH_KEY_COLUMN = "WhatsApp Search Key"
WHATSAPP_TARGET_TYPE_COLUMN = "WhatsApp Target Type"
PASTE_ACTION_CHOICES = ("comment", "mass-notification", "check-group-chat")
GROUP_CHAT_COLUMN = "Group Chat"












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


def _frozen_app_location_note() -> str:
    """Which .app this process actually is, for permission messages.

    Grants are per-copy: a teacher who runs the app straight out of the mounted
    DMG grants one copy in Settings while running another, and "still no
    permission" is unexplainable without showing the running path.
    """
    if not getattr(sys, "frozen", False):
        return ""
    exe = Path(sys.executable).resolve()
    app_path = next((str(parent) for parent in exe.parents if parent.suffix == ".app"), str(exe))
    if app_path.startswith("/Volumes/"):
        return (
            f"注意：现在运行的这份 App 在安装盘（DMG）里（{app_path}）。"
            "请先把它拖进「应用程序」文件夹，推出安装盘，再从「应用程序」里打开并授权。"
        )
    return f"（当前运行的 App：{app_path}。在系统设置里授权时，请确认选的就是这一份。）"


def mac_permission_hint(error_text: str) -> str:
    """Map a macOS TCC refusal to the exact setting the user must flip.

    Both errors surface as osascript stderr. They used to be swallowed into
    "window was not found", which pointed people at WeCom instead of at the
    missing grant -- a packaged app has its own TCC identity, so grants given to
    the dev Python do not carry over.
    """
    text = error_text or ""
    hint = ""
    if "-1743" in text or "Not authorized to send Apple events" in text:
        hint = (
            "这台电脑还没有允许本 App 控制 System Events（自动化权限）。"
            "请打开 系统设置 → 隐私与安全性 → 自动化，找到本 App，勾选 System Events，"
            "然后完全退出并重新打开本 App。"
            "如果列表里没有本 App，先点一次任意粘贴/检查按钮，屏幕会弹出询问框，点「允许」。"
        )
    elif "-25211" in text or "assistive access" in text:
        hint = (
            "这台电脑还没有给本 App 辅助功能权限。"
            "请打开 系统设置 → 隐私与安全性 → 辅助功能，把本 App 加入并打开开关，"
            "然后完全退出并重新打开本 App。"
        )
    if hint:
        note = _frozen_app_location_note()
        if note:
            hint += note
    return hint


def _mac_window_probe(mac_process_name: str) -> tuple[bool, str]:
    """(window found, permission error). Permission errors are not "no window"."""
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
            timeout=60,
        )
    except subprocess.TimeoutExpired:
        # The first automation call on a fresh machine blocks on the consent
        # dialog; timing out here almost always means it is sitting unanswered.
        return False, (
            "等待系统权限确认超时——屏幕上可能正弹着「想要控制 System Events」的询问框，"
            "请点「允许」后重试。" + _frozen_app_location_note()
        )
    except Exception as exc:
        return False, mac_permission_hint(str(exc))
    if result.returncode != 0:
        return False, mac_permission_hint(result.stderr)
    return result.stdout.strip() == "true", ""


def app_window_found(title_re: str, *, mac_process_name: str | None = None) -> bool:
    if sys.platform == "darwin":
        if not mac_process_name or not process_is_running((), mac_process_name=mac_process_name):
            return False
        found, _ = _mac_window_probe(mac_process_name)
        return found

    return find_window_by_title_re(title_re) is not None




def unminimize_app_window(spec: DesktopAppSpec) -> None:
    """Restore minimized windows, best effort.

    Activating an app does NOT bring a minimized window back on macOS (verified:
    AXMinimized stays true after `activate`), so a batch that minimized WeCom at
    the end would leave the next batch typing into nothing.
    """
    if sys.platform == "darwin":
        if not spec.mac_process_name:
            return
        script = (
            "function run() {"
            '  const se = Application("System Events");'
            f"  const procs = se.processes.whose({{name: {json.dumps(spec.mac_process_name)}}});"
            '  if (procs.length === 0) return "no-process";'
            "  const wins = procs[0].windows();"
            "  for (let i = 0; i < wins.length; i++) {"
            '    try { wins[i].attributes["AXMinimized"].value = false; } catch (e) {}'
            "  }"
            '  return "ok";'
            "}"
        )
        try:
            subprocess.run(
                ["osascript", "-l", "JavaScript", "-e", script],
                capture_output=True,
                text=True,
                timeout=15,
            )
        except Exception:
            pass
        return

    window = find_window_by_title_re(spec.title_re)
    if window is not None:
        try:
            window.restore()
        except Exception:
            pass


def minimize_app_window(spec: DesktopAppSpec) -> dict[str, Any]:
    """Minimize the app's windows, so a finished batch leaves the screen as it was."""
    if sys.platform == "darwin":
        if not spec.mac_process_name:
            return {"ok": False, "message": "no mac process name configured"}
        script = (
            "function run() {"
            '  const se = Application("System Events");'
            f"  const procs = se.processes.whose({{name: {json.dumps(spec.mac_process_name)}}});"
            '  if (procs.length === 0) return "no-process";'
            "  const wins = procs[0].windows();"
            "  let count = 0;"
            "  for (let i = 0; i < wins.length; i++) {"
            '    try { wins[i].attributes["AXMinimized"].value = true; count++; } catch (e) {}'
            "  }"
            '  return "minimized:" + count;'
            "}"
        )
        try:
            result = subprocess.run(
                ["osascript", "-l", "JavaScript", "-e", script],
                capture_output=True,
                text=True,
                timeout=15,
            )
        except Exception as exc:
            return {"ok": False, "message": str(exc)}
        if result.returncode != 0:
            hint = mac_permission_hint(result.stderr)
            return {"ok": False, "message": hint or result.stderr.strip()}
        return {"ok": True, "message": result.stdout.strip()}

    window = find_window_by_title_re(spec.title_re)
    if window is None:
        return {"ok": False, "message": "window not found"}
    try:
        window.minimize()
    except Exception as exc:
        return {"ok": False, "message": f"{type(exc).__name__}: {exc}"}
    return {"ok": True, "message": "minimized"}


def app_status(spec: DesktopAppSpec) -> DesktopAppStatus:
    dependency_ok, dependency_error = automation_dependency_status()
    running = process_is_running(spec.process_names, mac_process_name=spec.mac_process_name)
    permission_error = ""
    if not dependency_ok:
        window_found = False
    elif sys.platform == "darwin":
        window_found = False
        if spec.mac_process_name and running:
            window_found, permission_error = _mac_window_probe(spec.mac_process_name)
    else:
        window_found = app_window_found(spec.title_re, mac_process_name=spec.mac_process_name)

    if not dependency_ok:
        message = "desktop automation dependencies are not available"
    elif permission_error:
        message = permission_error
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
        permission_error=permission_error,
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
        "--prepare-app",
        choices=("wecom", "whatsapp"),
        help="Bring the app's window up (launching it if needed), print status JSON, and exit.",
    )
    parser.add_argument(
        "--minimize-app",
        choices=("wecom", "whatsapp"),
        help="Minimize the app's windows, print result JSON, and exit.",
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
    if status.permission_error:
        raise RuntimeError(status.permission_error)
    if not status.window_found:
        raise RuntimeError("Needed app window is not available.")
    return status


















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
        # A TCC refusal reads like an automation failure; point at the actual
        # setting instead of at WeCom.
        return JobResult(status="needs_review", error=mac_permission_hint(str(exc)) or str(exc))

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

    from wecom_win import (
        WECOM_MATCH_HEIGHT_RATIO,
        WeComPasteRobot,
        _clear_wecom_search_box,
        _ensure_wecom_search_open,
        _wecom_dropdown_height,
        _wecom_no_match_baseline,
        _wecom_reset_search_session,
    )

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

    from whatsapp_win import WhatsAppPasteRobot

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
                "permission_error": status.permission_error,
            }
            for key, status in ((k, app_status(spec)) for k, spec in specs.items())
        }
        print(json.dumps(report, ensure_ascii=False))
        return

    if args.prepare_app:
        # Bring the window up so a bulk run can start against an app that sits in
        # the background with its window closed.
        specs = build_app_specs(
            wecom_title_re=args.wecom_title_re,
            whatsapp_title_re=args.whatsapp_title_re,
        )
        spec = specs[args.prepare_app]
        status = app_status(spec)
        launched = False
        if status.dependency_ok and not status.window_found and not status.permission_error:
            launched = launch_app(spec, configured_exe_for_app(args, spec.key))
            for _ in range(10):
                time.sleep(1.5)
                status = app_status(spec)
                if status.window_found or status.permission_error:
                    break
        if status.window_found:
            # The window may exist but be sitting in the Dock (e.g. minimized by the
            # previous batch); AX-level presence does not mean it can take keystrokes.
            unminimize_app_window(spec)
        print(
            json.dumps(
                {
                    "window_found": status.window_found,
                    "launched": launched,
                    "message": status.message,
                    "permission_error": status.permission_error,
                },
                ensure_ascii=False,
            )
        )
        return

    if args.minimize_app:
        specs = build_app_specs(
            wecom_title_re=args.wecom_title_re,
            whatsapp_title_re=args.whatsapp_title_re,
        )
        print(json.dumps(minimize_app_window(specs[args.minimize_app]), ensure_ascii=False))
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
