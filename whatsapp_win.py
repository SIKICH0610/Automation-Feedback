from __future__ import annotations

import os
import time
from typing import Any
from urllib.parse import quote

from paste_common import (
    DEFAULT_WHATSAPP_TITLE_RE,
    PasteJob,
    copy_text_to_clipboard,
    find_window_by_title_re,
)


"""WhatsApp automation for Windows, driven through pywinauto's UI Automation.

whatsapp_mac.py is the in-progress macOS counterpart; it is not wired up yet, so
macOS rows come back needs_review instead.
"""


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
