from __future__ import annotations

import json
import threading
import types
import unittest
from pathlib import Path

import frontend_server
from frontend_server import ActionRunner, FrontendError


class FakeStore:
    """Minimal stand-in for SQLiteFeedbackStore covering only what the bulk
    group-chat path touches, so the loop can be exercised without a database
    or a real WeCom / WhatsApp window."""

    def __init__(self, rosters: dict[str, int]) -> None:
        self._rosters = rosters
        self.lock = threading.RLock()
        self.workbook_path = Path("runtime.xlsx")
        self.database_path = Path("feedback.db")
        self.app_data_dir = Path(".")
        self.prepared = 0
        self.synced = 0
        self.removed = 0

    def sheet_names(self) -> list[str]:
        return list(self._rosters)

    def load_sheet(self, sheet_name: str) -> dict[str, object]:
        count = self._rosters[sheet_name]
        return {"rows": [{"excel_row": index} for index in range(2, count + 2)]}

    def prepare_runtime_workbook(self) -> None:
        self.prepared += 1

    def sync_runtime_workbook(self) -> None:
        self.synced += 1

    def remove_runtime_workbook(self) -> None:
        self.removed += 1


class BulkGroupChatTest(unittest.TestCase):
    def setUp(self) -> None:
        self.store = FakeStore({"Class A": 3, "Class B": 2, "Empty Class": 0})
        self.runner = ActionRunner(self.store)
        self.commands: list[list[str]] = []
        self._real_run = frontend_server.subprocess.run

    def tearDown(self) -> None:
        frontend_server.subprocess.run = self._real_run

    def _stub_subprocess(self, *, probe: dict, returncode_for=None) -> None:
        """Route --check-apps to a canned probe result and record every other command."""

        def fake_run(command, **kwargs):
            if "--check-apps" in command:
                return types.SimpleNamespace(
                    stdout=json.dumps(probe), stderr="", returncode=0
                )
            self.commands.append(list(command))
            sheet = command[command.index("--sheet") + 1]
            code = returncode_for(sheet) if returncode_for else 0
            return types.SimpleNamespace(
                stdout=f"ran {sheet}", stderr="", returncode=code
            )

        frontend_server.subprocess.run = fake_run

    BOTH_OPEN = {
        "wecom": {"display_name": "WeCom", "window_found": True, "message": "window found"},
        "whatsapp": {"display_name": "WhatsApp", "window_found": True, "message": "window found"},
    }
    NONE_OPEN = {
        "wecom": {"display_name": "WeCom", "window_found": False, "message": "not running"},
        "whatsapp": {"display_name": "WhatsApp", "window_found": False, "message": "not running"},
    }

    def test_runs_one_subprocess_per_selected_class(self) -> None:
        self._stub_subprocess(probe=self.BOTH_OPEN)
        result = self.runner.run(
            {"action": "check-group-chat-bulk", "sheets": ["Class A", "Class B"]}
        )

        self.assertTrue(result["ok"])
        self.assertEqual(len(self.commands), 2)
        sheets = [command[command.index("--sheet") + 1] for command in self.commands]
        self.assertEqual(sheets, ["Class A", "Class B"])

        rows = [command[command.index("--rows") + 1] for command in self.commands]
        self.assertEqual(rows, ["2,3,4", "2,3"])

        for command in self.commands:
            self.assertIn("--action", command)
            self.assertEqual(command[command.index("--action") + 1], "check-group-chat")
            # A long unattended batch must never try to launch apps on its own.
            self.assertIn("--no-auto-open", command)

        # The workbook is prepared and synced once for the whole batch, not per class.
        self.assertEqual(self.store.prepared, 1)
        self.assertEqual(self.store.synced, 1)
        self.assertEqual(self.store.removed, 1)

    def test_output_keeps_a_section_per_class(self) -> None:
        self._stub_subprocess(probe=self.BOTH_OPEN)
        result = self.runner.run(
            {"action": "check-group-chat-bulk", "sheets": ["Class A", "Class B"]}
        )
        self.assertIn("===== Class A =====", result["output"])
        self.assertIn("===== Class B =====", result["output"])
        self.assertIn("ran Class A", result["output"])
        self.assertIn("ran Class B", result["output"])

    def test_class_with_no_students_is_skipped_without_running(self) -> None:
        self._stub_subprocess(probe=self.BOTH_OPEN)
        result = self.runner.run(
            {"action": "check-group-chat-bulk", "sheets": ["Empty Class", "Class B"]}
        )
        sheets = [command[command.index("--sheet") + 1] for command in self.commands]
        self.assertEqual(sheets, ["Class B"])
        self.assertIn("no students on this roster", result["output"])

    def test_one_failing_class_does_not_stop_the_rest(self) -> None:
        self._stub_subprocess(
            probe=self.BOTH_OPEN,
            returncode_for=lambda sheet: 1 if sheet == "Class A" else 0,
        )
        result = self.runner.run(
            {"action": "check-group-chat-bulk", "sheets": ["Class A", "Class B"]}
        )
        self.assertFalse(result["ok"])
        self.assertIn("had errors", result["label"])
        # Class B still ran even though Class A failed first.
        sheets = [command[command.index("--sheet") + 1] for command in self.commands]
        self.assertEqual(sheets, ["Class A", "Class B"])
        # Results collected before the failure are still written back.
        self.assertEqual(self.store.synced, 1)

    def test_stops_before_any_class_when_no_app_is_open(self) -> None:
        self._stub_subprocess(probe=self.NONE_OPEN)
        with self.assertRaises(FrontendError) as caught:
            self.runner.run({"action": "check-group-chat-bulk", "sheets": ["Class A"]})
        self.assertIn("Neither WeCom nor WhatsApp", str(caught.exception))
        self.assertEqual(self.commands, [])
        # Nothing was prepared, so there is nothing to sync or clean up.
        self.assertEqual(self.store.prepared, 0)

    def test_rejects_empty_and_unknown_sheets(self) -> None:
        self._stub_subprocess(probe=self.BOTH_OPEN)
        with self.assertRaises(FrontendError):
            self.runner.run({"action": "check-group-chat-bulk", "sheets": []})
        with self.assertRaises(FrontendError):
            self.runner.run({"action": "check-group-chat-bulk", "sheets": ["No Such Class"]})

    ONLY_WECOM = {
        "wecom": {"display_name": "WeCom", "window_found": True, "message": "window found"},
        "whatsapp": {"display_name": "WhatsApp", "window_found": False, "message": "not running"},
    }

    def test_channel_is_passed_through_to_each_class(self) -> None:
        self._stub_subprocess(probe=self.BOTH_OPEN)
        self.runner.run(
            {
                "action": "check-group-chat-bulk",
                "sheets": ["Class A", "Class B"],
                "channel": "wecom",
            }
        )
        for command in self.commands:
            self.assertEqual(command[command.index("--channel") + 1], "wecom")

    def test_channel_defaults_to_auto(self) -> None:
        self._stub_subprocess(probe=self.BOTH_OPEN)
        self.runner.run({"action": "check-group-chat-bulk", "sheets": ["Class A"]})
        command = self.commands[0]
        self.assertEqual(command[command.index("--channel") + 1], "auto")

    def test_forced_channel_requires_that_specific_app(self) -> None:
        # WhatsApp is shut, so a WhatsApp-only run must refuse even though WeCom is open.
        self._stub_subprocess(probe=self.ONLY_WECOM)
        with self.assertRaises(FrontendError) as caught:
            self.runner.run(
                {
                    "action": "check-group-chat-bulk",
                    "sheets": ["Class A"],
                    "channel": "whatsapp",
                }
            )
        self.assertIn("WhatsApp", str(caught.exception))
        self.assertEqual(self.commands, [])

    def test_forced_channel_runs_when_only_that_app_is_open(self) -> None:
        self._stub_subprocess(probe=self.ONLY_WECOM)
        result = self.runner.run(
            {
                "action": "check-group-chat-bulk",
                "sheets": ["Class A"],
                "channel": "wecom",
            }
        )
        self.assertTrue(result["ok"])
        self.assertEqual(len(self.commands), 1)

    def test_auto_still_runs_when_only_one_app_is_open(self) -> None:
        self._stub_subprocess(probe=self.ONLY_WECOM)
        result = self.runner.run({"action": "check-group-chat-bulk", "sheets": ["Class A"]})
        self.assertTrue(result["ok"])

    def test_rejects_unknown_channel(self) -> None:
        self._stub_subprocess(probe=self.BOTH_OPEN)
        with self.assertRaises(FrontendError):
            self.runner.run(
                {
                    "action": "check-group-chat-bulk",
                    "sheets": ["Class A"],
                    "channel": "telegram",
                }
            )

    def test_duplicate_sheets_run_once(self) -> None:
        self._stub_subprocess(probe=self.BOTH_OPEN)
        self.runner.run(
            {"action": "check-group-chat-bulk", "sheets": ["Class A", "Class A"]}
        )
        self.assertEqual(len(self.commands), 1)


if __name__ == "__main__":
    unittest.main()
