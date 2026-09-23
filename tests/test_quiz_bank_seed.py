from __future__ import annotations

import tempfile
import unittest
from pathlib import Path

from quiz_bank_store import (
    AMC10_QUIZ1_BANK_ID,
    QUIZ1_BANK_ID,
    QUIZ2_BANK_ID,
    SQLiteQuizBankStore,
    _default_bank_seeds,
)


class QuizBankSeedTest(unittest.TestCase):
    """The banks live as CSV data files and seed the store on first run."""

    def setUp(self) -> None:
        self._tmp = tempfile.TemporaryDirectory()
        self.store = SQLiteQuizBankStore(Path(self._tmp.name) / "quiz.db")

    def tearDown(self) -> None:
        self._tmp.cleanup()

    def test_csv_seeds_cover_all_three_banks(self) -> None:
        seeds = {bank["id"]: bank for bank in _default_bank_seeds()}
        self.assertEqual(
            set(seeds),
            {QUIZ1_BANK_ID, QUIZ2_BANK_ID, AMC10_QUIZ1_BANK_ID},
        )
        self.assertEqual(len(seeds[QUIZ2_BANK_ID]["entries"]), 8)
        self.assertEqual(
            [entry["question"] for entry in seeds[QUIZ2_BANK_ID]["entries"]],
            list(range(1, 9)),
        )

    def test_seeded_store_builds_comments_by_question_number(self) -> None:
        chinese = self.store.build_comment(QUIZ2_BANK_ID, "1, 5, 8", language="Chinese")
        english = self.store.build_comment(QUIZ2_BANK_ID, "1, 5, 8", language="English")
        for question in (1, 5, 8):
            self.assertIn(f"第 {question} 题", chinese)
            self.assertIn(f"question {question}", english.lower())

    def test_amc10_bank_matches_by_question_number_only(self) -> None:
        comment = self.store.build_comment(AMC10_QUIZ1_BANK_ID, "第3题", language="Chinese")
        self.assertIn("第 3 题", comment)
        self.assertEqual(
            self.store.build_comment(AMC10_QUIZ1_BANK_ID, "没有题号的一句话", language="Chinese"),
            "",
        )

    def test_teacher_edits_survive_reseeding(self) -> None:
        bank = self.store.load_bank(QUIZ1_BANK_ID)
        entries = [dict(entry) for entry in bank["entries"]]
        entries[0]["chinese"] = "老师改过的解析。"
        self.store.save_bank(
            QUIZ1_BANK_ID,
            class_name=bank["class_name"],
            quiz_name=bank["quiz_name"],
            display_name=bank["display_name"],
            entries=entries,
        )
        reopened = SQLiteQuizBankStore(self.store.database_path)
        kept = reopened.load_bank(QUIZ1_BANK_ID)["entries"][0]
        self.assertEqual(kept["chinese"], "老师改过的解析。")


if __name__ == "__main__":
    unittest.main()
