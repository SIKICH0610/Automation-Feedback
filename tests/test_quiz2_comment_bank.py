from __future__ import annotations

import unittest

from geometry_volume1_quiz2_comment_bank import (
    GEOMETRY_VOLUME1_QUIZ2_ISSUES,
    build_geometry_volume1_quiz2_comment,
    question_numbers_for_note,
)


class GeometryVolume1Quiz2CommentBankTest(unittest.TestCase):
    def test_all_provided_questions_are_available(self) -> None:
        self.assertEqual(
            [issue.question for issue in GEOMETRY_VOLUME1_QUIZ2_ISSUES],
            list(range(1, 9)),
        )
        self.assertEqual(question_numbers_for_note("1-8"), list(range(1, 9)))

    def test_question_keywords_match_independently(self) -> None:
        examples = {
            "Triangle Inequality": [1],
            "sin A": [2],
            "triangle midsegment parallelogram": [3],
            "square isosceles": [4],
            "trapezoid area": [5],
            "A model": [6],
            "reverse A model": [7],
            "Right Triangle Altitude Theorem": [8],
        }
        for note, expected in examples.items():
            with self.subTest(note=note):
                self.assertEqual(question_numbers_for_note(note), expected)

    def test_chinese_and_english_feedback_are_generated(self) -> None:
        chinese = build_geometry_volume1_quiz2_comment("1, 5, 8", language="Chinese")
        english = build_geometry_volume1_quiz2_comment("1, 5, 8", language="English")

        for question in (1, 5, 8):
            self.assertIn(f"第 {question} 题", chinese)
            self.assertIn(f"question {question}", english.lower())

        for forbidden in (":", ";", "：", "；"):
            self.assertNotIn(forbidden, chinese)
            self.assertNotIn(forbidden, english)

    def test_unprovided_question_nine_has_no_invented_feedback(self) -> None:
        self.assertEqual(question_numbers_for_note("9"), [])
        self.assertEqual(build_geometry_volume1_quiz2_comment("9"), "")


if __name__ == "__main__":
    unittest.main()
