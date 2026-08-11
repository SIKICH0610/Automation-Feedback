from __future__ import annotations

from types import SimpleNamespace
import unittest
from unittest.mock import Mock, patch

from paste_sender import PasteJob, WeComPasteRobot, WhatsAppPasteRobot


class FakeElement:
    def __init__(
        self,
        text: str,
        *,
        control_type: str = "Text",
        class_name: str = "",
        focused: bool = False,
        rectangle: tuple[int, int, int, int] = (0, 0, 240, 40),
    ) -> None:
        self._text = text
        self._focused = focused
        self.focus_count = 0
        self._rectangle = SimpleNamespace(
            left=rectangle[0],
            top=rectangle[1],
            right=rectangle[2],
            bottom=rectangle[3],
            width=lambda: rectangle[2] - rectangle[0],
            height=lambda: rectangle[3] - rectangle[1],
        )
        self.element_info = SimpleNamespace(
            control_type=control_type,
            class_name=class_name,
            has_keyboard_focus=focused,
        )

    def window_text(self) -> str:
        return self._text

    def has_keyboard_focus(self) -> bool:
        return self._focused

    def set_focus(self) -> None:
        self.focus_count += 1

    def rectangle(self):
        return self._rectangle


class FakeWindow(FakeElement):
    def __init__(self, title: str, elements: list[FakeElement]) -> None:
        super().__init__(title)
        self._elements = elements

    def exists(self, timeout: float = 0) -> bool:
        return True

    def descendants(self) -> list[FakeElement]:
        return self._elements


def paste_job(*, channel: str = "wecom") -> PasteJob:
    return PasteJob(
        excel_row=2,
        uid="7916327",
        student_name="Emma Zhong",
        parent_language="Chinese",
        channel=channel,
        channel_explicit=True,
        search_key="7916327",
        expected_chat_name="Emma Zhong - 7916327",
        whatsapp_phone="",
        whatsapp_target_type="group_search",
        feedback="A parent update",
    )


def bare_robot(robot_type: type[WeComPasteRobot] | type[WhatsAppPasteRobot], window: FakeWindow):
    robot = robot_type.__new__(robot_type)
    robot.window = window
    robot.desktop = SimpleNamespace(window=lambda **_: window)
    robot.title_re = ".*"
    robot.search_shortcut = "^f"
    robot.settle_seconds = 0
    robot.send_keys = Mock()
    return robot


class WindowsPasteRobotTest(unittest.TestCase):
    def test_visible_text_does_not_steal_focus_and_ignores_search_edit(self) -> None:
        search = FakeElement("7916327", control_type="Edit", focused=True)
        group = FakeElement("Geo TTh Emma Zhong 7916327")
        window = FakeWindow("WeCom", [search, group])
        robot = bare_robot(WeComPasteRobot, window)

        text = robot.visible_text(ignored_edit_values={"7916327"})

        self.assertIn("Emma Zhong 7916327", text)
        self.assertEqual(window.focus_count, 0)

    def test_both_searches_clear_old_text_before_pasting_new_key(self) -> None:
        for robot_type in (WeComPasteRobot, WhatsAppPasteRobot):
            with self.subTest(robot=robot_type.__name__):
                window = FakeWindow("Chat", [])
                robot = bare_robot(robot_type, window)
                with (
                    patch("paste_sender.copy_text_to_clipboard") as copy_text,
                    patch("paste_sender.time.sleep"),
                ):
                    robot.search_chat("7916327")

                keys = [call.args[0] for call in robot.send_keys.call_args_list]
                self.assertEqual(keys[-4:], ["^f", "^a", "{BACKSPACE}", "^v"])
                copy_text.assert_called_once()

    def test_text_paste_uses_the_active_chat_without_refocusing(self) -> None:
        for robot_type in (WeComPasteRobot, WhatsAppPasteRobot):
            with self.subTest(robot=robot_type.__name__):
                robot = bare_robot(robot_type, FakeWindow("Chat", []))
                robot.focus_window = Mock(side_effect=AssertionError("must not refocus"))
                with (
                    patch("paste_sender.copy_text_to_clipboard") as copy_text,
                    patch("paste_sender.time.sleep"),
                ):
                    robot.paste_feedback("A parent update")

                copy_text.assert_called_once()
                robot.send_keys.assert_called_once_with("^v")

    def test_wecom_press_enter_is_not_blocked_by_uia_detection(self) -> None:
        robot = bare_robot(WeComPasteRobot, FakeWindow("WeCom", []))
        robot.search_chat = Mock()
        robot.search_result_candidates = Mock(return_value=[])
        robot.search_is_active = Mock(return_value=False)

        with patch("paste_sender.time.sleep"):
            robot.open_chat_from_search(paste_job())

        robot.send_keys.assert_called_once_with("{ENTER}")

    def test_wecom_uid_match_does_not_require_a_fixed_group_name(self) -> None:
        custom_group = FakeElement("Teacher Custom Group 7916327")
        robot = bare_robot(WeComPasteRobot, FakeWindow("WeCom", [custom_group]))

        candidates = robot.search_result_candidates(paste_job())

        self.assertEqual(len(candidates), 1)
        self.assertIn("7916327", candidates[0][2])

    def test_whatsapp_press_enter_is_not_blocked_by_uia_detection(self) -> None:
        robot = bare_robot(WhatsAppPasteRobot, FakeWindow("WhatsApp", []))
        robot.search_chat = Mock()
        robot.search_result_matches = Mock(return_value=False)
        robot.search_is_active = Mock(return_value=False)

        with patch("paste_sender.time.sleep"):
            robot.open_chat_from_search(paste_job(channel="whatsapp"))

        robot.send_keys.assert_called_once_with("{ENTER}")

    def test_wecom_presses_enter_again_when_search_remains_active(self) -> None:
        robot = bare_robot(WeComPasteRobot, FakeWindow("WeCom", []))
        robot.search_chat = Mock()
        robot.wait_for_search_result_candidates = Mock(
            return_value=[(100, FakeElement("7916327"), "Custom Group 7916327")]
        )
        robot.search_is_active = Mock(side_effect=[True, False])

        with patch("paste_sender.time.sleep"):
            robot.open_chat_from_search(paste_job())

        keys = [call.args[0] for call in robot.send_keys.call_args_list]
        self.assertEqual(keys, ["{ENTER}", "{ENTER}"])

    def test_whatsapp_presses_enter_again_when_search_remains_active(self) -> None:
        robot = bare_robot(WhatsAppPasteRobot, FakeWindow("WhatsApp", []))
        robot.search_chat = Mock()
        robot.wait_for_search_result = Mock(return_value=True)
        robot.search_is_active = Mock(side_effect=[True, False])

        with patch("paste_sender.time.sleep"):
            robot.open_chat_from_search(paste_job(channel="whatsapp"))

        keys = [call.args[0] for call in robot.send_keys.call_args_list]
        self.assertEqual(keys, ["{ENTER}", "{ENTER}"])


if __name__ == "__main__":
    unittest.main()
