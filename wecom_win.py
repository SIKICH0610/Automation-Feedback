from __future__ import annotations

import re
import time
from typing import Any

from paste_common import DEFAULT_WECOM_TITLE_RE, PasteJob, copy_text_to_clipboard


"""WeCom automation for Windows, driven through pywinauto's UI Automation.

The macOS half lives in wecom_mac.py and works completely differently: WeCom for
Windows draws its own UI, so nothing here can read text off the screen, while the Mac
build exposes a real accessibility tree. paste_sender picks between the two.
"""


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
