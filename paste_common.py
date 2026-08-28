from __future__ import annotations

import re
import time
from dataclasses import dataclass
from typing import Any


"""Types and helpers shared by the sender and every platform's paste robot.

These live here rather than in paste_sender so the per-platform modules can import
them without importing the sender back, which would be a cycle.
"""


DEFAULT_WECOM_TITLE_RE = r".*(WeCom|企业微信|WXWork).*"


DEFAULT_WHATSAPP_TITLE_RE = r".*(WhatsApp|WhatsApp Web).*"


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
    # Set when the OS refused the automation call itself (macOS TCC): the app lacks
    # the Automation or Accessibility grant, which is different from "no window".
    permission_error: str = ""


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
