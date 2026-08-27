from __future__ import annotations

import json
import subprocess
import time
from typing import Any

import pyperclip


DEFAULT_PROCESS_NAME = "企业微信"


class WeComAutomationError(RuntimeError):
    pass


def _run_osascript(args: list[str], *, timeout: float = 15) -> str:
    try:
        result = subprocess.run(
            ["osascript", *args],
            capture_output=True,
            text=True,
            timeout=timeout,
        )
    except FileNotFoundError as exc:
        raise WeComAutomationError("osascript is not available on this system.") from exc
    except subprocess.TimeoutExpired as exc:
        raise WeComAutomationError("A WeCom automation step timed out.") from exc
    if result.returncode != 0:
        raise WeComAutomationError(result.stderr.strip() or "osascript failed.")
    return result.stdout.strip()


def _run_jxa(script: str, *, timeout: float = 15) -> str:
    return _run_osascript(["-l", "JavaScript", "-e", script], timeout=timeout)


def process_is_running(process_name: str = DEFAULT_PROCESS_NAME) -> bool:
    try:
        result = subprocess.run(["pgrep", "-x", process_name], capture_output=True, text=True, timeout=5)
    except Exception:
        return False
    return result.returncode == 0


def launch_app(app_name: str = DEFAULT_PROCESS_NAME) -> bool:
    try:
        result = subprocess.run(["open", "-a", app_name], capture_output=True, text=True, timeout=10)
    except Exception:
        return False
    return result.returncode == 0


# WeCom's Cmd+F search opens a separate floating AXDialog window. Its rows visibly
# highlight the correct match (confirmed via AXSelected) but do not respond to
# accessibility-level clicks or keyboard list-navigation aimed at the dialog itself.
# Pasting the key from the clipboard and then pressing Return does work, the same
# way a person using the app would -- the dialog closing on its own is not proof it
# worked, so this always gets verified afterward by checking the sidebar's own
# AXSelected row (see verify_chat).
_JXA_SEARCH_AND_OPEN_TOP_RESULT = """
function run() {
  const se = Application("System Events");
  const proc = se.processes["%(process)s"];

  se.keystroke("f", {using: "command down"});
  delay(0.3);
  se.keystroke("a", {using: "command down"});
  se.keystroke(String.fromCharCode(8));
  delay(0.2);
  se.keystroke("v", {using: "command down"});
  delay(1.5);

  const dialogsBefore = proc.windows.whose({subrole: "AXDialog"});
  if (dialogsBefore.length === 0) {
    se.keyCode(53); // Escape, in case the search box is left open with no results
    return JSON.stringify({ok: false, reason: "no_search_results"});
  }

  se.keyCode(36); // Return: open WeCom's own top search result
  delay(0.7);
  return JSON.stringify({ok: true});
}
"""

_JXA_CLEAR_SEARCH = """
function run() {
  const se = Application("System Events");
  const proc = se.processes["%(process)s"];
  se.keystroke("f", {using: "command down"});
  delay(0.2);
  se.keystroke("a", {using: "command down"});
  se.keystroke(String.fromCharCode(8));
  delay(0.1);
  se.keyCode(53); // Escape
  return "ok";
}
"""

_JXA_HAS_WINDOW = """
function run() {
  const se = Application("System Events");
  const procs = se.processes.whose({name: "%(process)s"});
  if (procs.length === 0) return "false";
  return procs[0].windows.length > 0 ? "true" : "false";
}
"""

# The conversation header isn't a reliable place to verify against: contacts show a
# simple name there, but a group without a custom announcement shows a generic
# "Group created by WeCom users..." system message instead, at the same position,
# with no group name in it at all. The sidebar row's own AXSelected state is a much
# more reliable signal of which chat is actually open (confirmed by cross-checking
# it against what was actually visible on screen).
_JXA_SELECTED_SIDEBAR_ROW = """
function run() {
  const se = Application("System Events");
  const proc = se.processes["%(process)s"];
  const win = proc.windows.whose({name: "WeCom"})[0];
  const sg = win.uiElements.whose({role: "AXSplitGroup"})[0];
  const inner = sg.uiElements.whose({role: "AXSplitGroup"})[0];
  const sa = inner.uiElements.whose({role: "AXScrollArea"})[0];
  const tbl = sa.uiElements.whose({role: "AXTable"})[0];
  const rows = tbl.uiElements();
  for (let i = 0; i < rows.length; i++) {
    let selected = false;
    try { selected = rows[i].attributes["AXSelected"].value(); } catch (e) {}
    if (!selected) continue;
    const cells = rows[i].uiElements();
    const texts = [];
    if (cells.length) {
      for (const el of cells[0].uiElements()) {
        try {
          const v = el.value();
          if (v) texts.push(String(v));
        } catch (e) {}
      }
    }
    return JSON.stringify({found: true, texts: texts});
  }
  return JSON.stringify({found: false, texts: []});
}
"""

# The compose text area (and its toolbar) only exist in the accessibility tree while a
# chat that can actually be replied to is open -- an official/system account like
# "WeCom Team" has none, which is how focus_compose_box below detects that case.
_JXA_FOCUS_COMPOSE_BOX = """
function run() {
  const se = Application("System Events");
  const proc = se.processes["%(process)s"];
  const win = proc.windows.whose({name: "WeCom"})[0];
  const sg = win.uiElements.whose({role: "AXSplitGroup"})[0];
  const inner = sg.uiElements.whose({role: "AXSplitGroup"})[0];
  const convoPane = inner.uiElements.whose({role: "AXSplitGroup"})[0];
  const level2 = convoPane.uiElements()[0];
  const areas = level2.uiElements.whose({role: "AXScrollArea"})();
  for (const area of areas) {
    for (const k of area.uiElements()) {
      let role = "";
      try { role = k.role(); } catch (e) { continue; }
      if (role === "AXTextArea") {
        k.attributes["AXFocused"].value = true;
        return "ok";
      }
    }
  }
  return "not_found";
}
"""


class WeComPasteRobotMac:
    """macOS equivalent of paste_sender.WeComPasteRobot, driven through the
    Accessibility API via osascript/JXA instead of pywinauto's Windows UI Automation.

    Unlike the Windows build (whose own comments note UI Automation exposes nothing
    useful for WeCom), the macOS WeCom client exposes real, readable text for both
    the chat list and the compose box -- so this verifies the opened chat by reading
    the sidebar's own selection state directly instead of the Windows
    implementation's pixel-based dropdown-height heuristic.
    """

    def __init__(self, *, process_name: str = DEFAULT_PROCESS_NAME, settle_seconds: float = 0.8) -> None:
        self.process_name = process_name
        self.settle_seconds = settle_seconds

    def focus_window(self) -> None:
        _run_osascript(["-e", f'tell application "{self.process_name}" to activate'])
        time.sleep(0.3)
        script = _JXA_HAS_WINDOW % {"process": self.process_name}
        if _run_jxa(script) != "true":
            raise WeComAutomationError(f"{self.process_name} window was not found.")

    def clear_search_state(self) -> None:
        try:
            script = _JXA_CLEAR_SEARCH % {"process": self.process_name}
            _run_jxa(script)
            time.sleep(self.settle_seconds)
        except WeComAutomationError:
            pass

    def search_and_open_top_result(self, search_key: str) -> None:
        self.focus_window()
        pyperclip.copy(search_key)
        script = _JXA_SEARCH_AND_OPEN_TOP_RESULT % {"process": self.process_name}
        result = json.loads(_run_jxa(script))
        if not result.get("ok"):
            raise LookupError(f"No WeCom search results for {search_key!r}.")

    def selected_sidebar_row_texts(self) -> list[str]:
        script = _JXA_SELECTED_SIDEBAR_ROW % {"process": self.process_name}
        result = json.loads(_run_jxa(script))
        return result.get("texts", [])

    def verify_chat(self, *, uid: str, name_parts: list[str], expected_chat_name: str) -> tuple[bool, str]:
        texts = self.selected_sidebar_row_texts()
        if not texts:
            return False, "no_chat_selected_in_sidebar"
        combined = " | ".join(texts)
        combined_lower = combined.lower()
        uid_ok = bool(uid) and uid in combined
        expected_ok = bool(expected_chat_name) and expected_chat_name.lower() in combined_lower
        name_ok = any(part and part in combined for part in name_parts)

        if expected_ok:
            return True, "verified_expected_name"
        if uid_ok and name_ok:
            return True, "verified_uid_and_name"
        if uid_ok or name_ok:
            return False, "selected_chat_partial_match_only"
        return False, "selected_chat_did_not_match"

    def focus_compose_box(self) -> bool:
        script = _JXA_FOCUS_COMPOSE_BOX % {"process": self.process_name}
        return _run_jxa(script) == "ok"

    def paste_feedback(self, feedback: str) -> None:
        if not self.focus_compose_box():
            raise WeComAutomationError(
                "Could not focus the WeCom message input box. The open chat may not "
                "support replies (for example, an official/system account)."
            )
        pyperclip.copy(feedback)
        time.sleep(0.2)
        _run_osascript(["-e", 'tell application "System Events" to keystroke "v" using command down'])
        time.sleep(0.8)
