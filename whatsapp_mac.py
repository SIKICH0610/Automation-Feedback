from __future__ import annotations

import json
import subprocess
import time

import pyperclip


DEFAULT_APP_NAME = "WhatsApp"


class WhatsAppAutomationError(RuntimeError):
    pass


def _run_osascript(args: list[str], *, timeout: float = 60) -> str:
    try:
        result = subprocess.run(
            ["osascript", *args],
            capture_output=True,
            text=True,
            timeout=timeout,
        )
    except FileNotFoundError as exc:
        raise WhatsAppAutomationError("osascript is not available on this system.") from exc
    except subprocess.TimeoutExpired as exc:
        raise WhatsAppAutomationError("A WhatsApp automation step timed out.") from exc
    if result.returncode != 0:
        raise WhatsAppAutomationError(result.stderr.strip() or "osascript failed.")
    return result.stdout.strip()


def _run_jxa(script: str, *, timeout: float = 60) -> str:
    return _run_osascript(["-l", "JavaScript", "-e", script], timeout=timeout)


def process_is_running(app_name: str = DEFAULT_APP_NAME) -> bool:
    try:
        result = subprocess.run(["pgrep", "-x", app_name], capture_output=True, text=True, timeout=5)
    except Exception:
        return False
    return result.returncode == 0


def launch_app(app_name: str = DEFAULT_APP_NAME) -> bool:
    try:
        result = subprocess.run(["open", "-a", app_name], capture_output=True, text=True, timeout=10)
    except Exception:
        return False
    return result.returncode == 0


# WhatsApp's search is inline in the sidebar (unlike WeCom's separate floating
# dialog): typing/pasting into it filters the same list into "Chats" / "Media" /
# "Messages" sections. The matching chat button is found by exact description text
# and clicked directly -- WhatsApp's chat rows are real AXButton elements, and
# (unlike WeCom's custom table rows) they respond correctly to accessibility clicks.
_JXA_SEARCH_AND_CLICK = """
function rawChildren(el) {
  try { return el.attributes["AXChildren"].value(); } catch (e) { return []; }
}
function findAndClick(el, depth, maxDepth, keyword, done) {
  if (depth > maxDepth || done.clicked) return;
  let role = "?", elDesc = "";
  try { role = el.role(); } catch (e) { return; }
  try { elDesc = String(el.description()); } catch (e) {}
  if (role === "AXButton" && elDesc === keyword) {
    el.click();
    done.clicked = true;
    return;
  }
  const kids = rawChildren(el);
  for (const k of kids) {
    findAndClick(k, depth + 1, maxDepth, keyword, done);
    if (done.clicked) return;
  }
}

function run() {
  const se = Application("System Events");
  const proc = se.processes["%(process)s"];
  const win = proc.windows()[0];

  se.keystroke("f", {using: "command down"});
  delay(0.3);
  se.keystroke("a", {using: "command down"});
  se.keystroke(String.fromCharCode(8));
  delay(0.2);
  se.keystroke("v", {using: "command down"});
  delay(1.2);

  const done = {clicked: false};
  findAndClick(win, 0, 14, %(chat_label_json)s, done);
  se.keyCode(53); // Escape, to close the search results afterward either way
  return JSON.stringify(done);
}
"""

# WhatsApp exposes an AXGroup whose description is literally "Messages in chat with
# <name>" for the currently open conversation -- a purpose-built, unambiguous
# accessibility label, more reliable than WeCom's header text (which can be a
# generic system message for groups without a custom announcement).
_JXA_OPEN_CHAT_LABEL = """
function rawChildren(el) {
  try { return el.attributes["AXChildren"].value(); } catch (e) { return []; }
}
function findOpenChatLabel(el, depth, maxDepth, results) {
  if (depth > maxDepth || results.length) return;
  let role = "?", elDesc = "";
  try { role = el.role(); } catch (e) { return; }
  try { elDesc = String(el.description()); } catch (e) {}
  if (role === "AXGroup" && elDesc.indexOf("Messages in chat with ") === 0) {
    results.push(elDesc.slice("Messages in chat with ".length));
    return;
  }
  let kids = rawChildren(el);
  if (role === "AXTable" || role === "AXRow" || role === "AXList" || role === "AXOutline") kids = kids.slice(0, 2);
  if (elDesc === "‎List of chats") kids = [];
  for (const k of kids) {
    findOpenChatLabel(k, depth + 1, maxDepth, results);
    if (results.length) return;
  }
}

function run() {
  const se = Application("System Events");
  const proc = se.processes["%(process)s"];
  const win = proc.windows()[0];
  const results = [];
  findOpenChatLabel(win, 0, 14, results);
  return JSON.stringify(results);
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


class WhatsAppPasteRobotMac:
    """macOS equivalent of paste_sender.WhatsAppPasteRobot, driven through the
    Accessibility API via osascript/JXA instead of pywinauto's Windows UI Automation.

    Unlike WeCom, WhatsApp for Mac's search is inline (not a separate floating
    dialog) and its chat rows are real AXButton elements that respond correctly to
    accessibility clicks -- and after opening a chat, its compose box is focused by
    default, so paste_feedback below pastes directly without an explicit focus step.
    """

    def __init__(self, *, process_name: str = DEFAULT_APP_NAME, settle_seconds: float = 0.8) -> None:
        self.process_name = process_name
        self.settle_seconds = settle_seconds

    def focus_window(self) -> None:
        _run_osascript(["-e", f'tell application "{self.process_name}" to activate'])
        time.sleep(0.3)
        script = _JXA_HAS_WINDOW % {"process": self.process_name}
        if _run_jxa(script) != "true":
            raise WhatsAppAutomationError(f"{self.process_name} window was not found.")

    def search_and_open(self, search_key: str, *, chat_label: str) -> None:
        """Search for search_key and click the "Chats" section result whose
        description exactly matches chat_label (the same text shown in the sidebar
        row, e.g. "Nikhil 7899198")."""
        self.focus_window()
        pyperclip.copy(search_key)
        script = _JXA_SEARCH_AND_CLICK % {
            "process": self.process_name,
            "chat_label_json": json.dumps(chat_label),
        }
        result = json.loads(_run_jxa(script))
        if not result.get("clicked"):
            raise LookupError(f"No WhatsApp chat named {chat_label!r} found for search key {search_key!r}.")
        time.sleep(self.settle_seconds)

    def open_chat_label(self) -> str:
        script = _JXA_OPEN_CHAT_LABEL % {"process": self.process_name}
        results = json.loads(_run_jxa(script))
        return results[0] if results else ""

    def verify_chat(self, *, name_parts: list[str], expected_chat_name: str) -> tuple[bool, str]:
        label = self.open_chat_label()
        if not label:
            return False, "no_chat_open"
        label_lower = label.lower()
        expected_ok = bool(expected_chat_name) and expected_chat_name.lower() in label_lower
        name_ok = any(part and part in label for part in name_parts)

        if expected_ok:
            return True, "verified_expected_name"
        if name_ok:
            return True, "verified_name"
        return False, "open_chat_did_not_match"

    def paste_feedback(self, feedback: str) -> None:
        # WhatsApp focuses the compose box by default after a chat is opened, so
        # this pastes directly rather than hunting for and focusing it explicitly
        # (which was tried and consistently failed to land the paste).
        _run_osascript(["-e", f'tell application "{self.process_name}" to activate'])
        time.sleep(0.2)
        pyperclip.copy(feedback)
        time.sleep(0.2)
        _run_osascript(["-e", 'tell application "System Events" to keystroke "v" using command down'])
        time.sleep(0.8)
