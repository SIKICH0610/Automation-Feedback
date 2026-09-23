from __future__ import annotations

from dataclasses import dataclass
from typing import Any


"""Types and helpers shared by paste_sender and the macOS paste robots.

These live here rather than in paste_sender so the robot modules can import
them without importing the sender back, which would be a cycle.
"""


@dataclass
class DesktopAppSpec:
    key: str
    display_name: str
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


def clear_clipboard() -> None:
    try:
        import pyperclip

        pyperclip.copy("")
    except Exception:
        pass
