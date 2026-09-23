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
# The wait for results is a poll, not a fixed delay: a fast machine continues the
# moment the results dialog appears (typically ~0.4s), while a slow machine gets up
# to 2.5s -- longer than the old fixed 1.5s -- before "no results" is concluded.
_JXA_SEARCH_AND_OPEN_TOP_RESULT = """
function run() {
  const se = Application("System Events");
  const proc = se.processes["%(process)s"];

  // Pre-flight: a stray search window left over from an earlier mishap (or the
  // person clicking around mid-batch) swallows every later Cmd+F, so anything
  // beyond the main window gets Escaped away before we type.
  let preflight = 0;
  while (proc.windows.length > 1 && preflight < 3) {
    se.keyCode(53); // Escape
    delay(0.25);
    preflight += 1;
  }

  se.keystroke("f", {using: "command down"});
  delay(0.3);
  se.keystroke("a", {using: "command down"});
  se.keystroke(String.fromCharCode(8));
  delay(0.2);
  se.keystroke("v", {using: "command down"});

  let waited = 0;
  while (waited < 2.5) {
    if (proc.windows.whose({subrole: "AXDialog"}).length > 0) break;
    delay(0.1);
    waited += 0.1;
  }
  if (proc.windows.whose({subrole: "AXDialog"}).length === 0) {
    se.keyCode(53); // Escape, in case the search box is left open with no results
    return JSON.stringify({ok: false, reason: "no_search_results"});
  }
  // The dialog pops before the real results are in it (it opens with the echo
  // suggestion row first). Pressing Return on that opens nothing, so the poll
  // only shaves the wait-for-dialog part; results still get a fixed settle.
  delay(0.8);

  se.keyCode(36); // Return: open WeCom's own top search result

  // A successful open closes the results dialog on its own. When nothing
  // matched, Return instead expands the palette into WeCom's full-page search
  // window (its own titled window with tabs), which then eats every later
  // Cmd+F and paste. So poll for the good outcome -- back to just the main
  // window -- and only after the full wait conclude "no result" and Escape
  // back out. The poll waits on success, never cuts it short, so slow
  // machines are safe.
  let settled = 0;
  while (settled < 1.5) {
    delay(0.15);
    settled += 0.15;
    if (proc.windows.length <= 1) break;
  }
  if (proc.windows.length > 1) {
    let esc = 0;
    while (proc.windows.length > 1 && esc < 3) {
      se.keyCode(53); // Escape closes the full-page search window
      delay(0.3);
      esc += 1;
    }
    return JSON.stringify({ok: false, reason: "no_search_results"});
  }
  delay(0.2);
  return JSON.stringify({ok: true});
}
"""

_JXA_CLEAR_SEARCH = """
function run() {
  const se = Application("System Events");
  const proc = se.processes["%(process)s"];
  let preflight = 0;
  while (proc.windows.length > 1 && preflight < 3) {
    se.keyCode(53); // Escape any stray search window before touching Cmd+F
    delay(0.25);
    preflight += 1;
  }
  se.keystroke("f", {using: "command down"});
  delay(0.2);
  se.keystroke("a", {using: "command down"});
  se.keystroke(String.fromCharCode(8));
  delay(0.1);
  se.keyCode(53); // Escape
  return "ok";
}
"""

_JXA_UNMINIMIZE = """
function run() {
  const se = Application("System Events");
  const procs = se.processes.whose({name: "%(process)s"});
  if (procs.length === 0) return "no-process";
  const wins = procs[0].windows();
  for (let i = 0; i < wins.length; i++) {
    try { wins[i].attributes["AXMinimized"].value = false; } catch (e) {}
  }
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

    def __init__(self, *, process_name: str = DEFAULT_PROCESS_NAME, settle_seconds: float = 0.4) -> None:
        self.process_name = process_name
        self.settle_seconds = settle_seconds

    def focus_window(self) -> None:
        _run_osascript(["-e", f'tell application "{self.process_name}" to activate'])
        time.sleep(0.3)
        # activate does not restore a minimized window (AXMinimized stays true), so a
        # batch that tidied WeCom into the Dock would otherwise leave the next run
        # typing into nothing.
        try:
            _run_jxa(_JXA_UNMINIMIZE % {"process": self.process_name})
        except WeComAutomationError:
            pass
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

    def refocus(self) -> None:
        """Cheap per-row focus: activate only, no sleeps, no unminimize.

        The full focus_window runs once per batch. This one exists so keystrokes
        cannot land in another app if the person switches away mid-batch --
        activate is a no-op when WeCom is already frontmost, and the search
        script's own leading delay covers activation latency.
        """
        _run_osascript(["-e", f'tell application "{self.process_name}" to activate'])

    def search_and_open_top_result(self, search_key: str) -> None:
        pyperclip.copy(search_key)
        script = _JXA_SEARCH_AND_OPEN_TOP_RESULT % {"process": self.process_name}
        result = json.loads(_run_jxa(script))
        if not result.get("ok"):
            raise LookupError(f"No WeCom search results for {search_key!r}.")

    def selected_sidebar_row_texts(self) -> list[str]:
        script = _JXA_SELECTED_SIDEBAR_ROW % {"process": self.process_name}
        result = json.loads(_run_jxa(script))
        return result.get("texts", [])

    def verify_chat(
        self,
        *,
        uid: str,
        name_parts: list[str],
        expected_chat_name: str,
        timeout: float = 2.5,
    ) -> tuple[bool, str]:
        """Poll the sidebar until the opened chat matches or the timeout passes.

        Polling replaces the old fixed post-Enter delay: a fast machine verifies on
        the first read, a slow one gets more total time than it ever did before.
        """
        deadline = time.monotonic() + timeout
        while True:
            verified, reason = self._verify_chat_once(
                uid=uid, name_parts=name_parts, expected_chat_name=expected_chat_name
            )
            if verified or time.monotonic() >= deadline:
                return verified, reason
            time.sleep(0.2)

    def _verify_chat_once(self, *, uid: str, name_parts: list[str], expected_chat_name: str) -> tuple[bool, str]:
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
        # The two partial cases are not equally risky: the search key IS the
        # uid, so a selected row carrying that uid is the group the search
        # opened (the name may simply not be part of the group's title), while
        # a name-only match smells like a same-named student's other group.
        if uid_ok:
            return False, "partial_uid_only"
        if name_ok:
            return False, "partial_name_only"
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
