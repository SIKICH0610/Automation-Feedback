from __future__ import annotations

import os
from pathlib import Path
from tempfile import TemporaryDirectory
import unittest
from unittest.mock import patch

from quiz_bank_store import (
    QUIZ1_BANK_ID,
    QUIZ2_BANK_ID,
    QuizBankStoreError,
    SQLiteQuizBankStore,
    runtime_quiz_bank_comment,
)


class QuizBankStoreTest(unittest.TestCase):
    def setUp(self) -> None:
        self.temp_dir = TemporaryDirectory(dir=Path(__file__).parent)
        self.database_path = Path(self.temp_dir.name) / "feedback.db"
        self.store = SQLiteQuizBankStore(self.database_path)

    def tearDown(self) -> None:
        self.temp_dir.cleanup()

    def test_default_banks_are_seeded_once(self) -> None:
        banks = self.store.list_banks()
        self.assertEqual(
            [bank["id"] for bank in banks],
            [QUIZ1_BANK_ID, QUIZ2_BANK_ID],
        )
        quiz2 = self.store.load_bank(QUIZ2_BANK_ID)
        self.assertEqual(
            [entry["question"] for entry in quiz2["entries"]],
            list(range(1, 9)),
        )

        self.store.save_bank(
            QUIZ2_BANK_ID,
            class_name=quiz2["class_name"],
            quiz_name=quiz2["quiz_name"],
            display_name=quiz2["display_name"],
            entries=[],
        )
        reopened = SQLiteQuizBankStore(self.database_path)
        self.assertEqual(reopened.load_bank(QUIZ2_BANK_ID)["entries"], [])

    def test_saved_feedback_is_used_by_runtime_generation(self) -> None:
        quiz2 = self.store.load_bank(QUIZ2_BANK_ID)
        quiz2["entries"][0]["chinese"] = "数据库中的第一题反馈。"
        quiz2["entries"][0]["english"] = "Question one from the database."
        quiz2["entries"][0]["patterns"] = ["custom triangle keyword"]
        saved = self.store.save_bank(
            QUIZ2_BANK_ID,
            class_name=quiz2["class_name"],
            quiz_name=quiz2["quiz_name"],
            display_name=quiz2["display_name"],
            entries=quiz2["entries"],
        )
        self.assertEqual(saved["entries"][0]["chinese"], "数据库中的第一题反馈。")

        with patch.dict(
            os.environ,
            {"FEEDBACK_DATABASE_PATH": str(self.database_path)},
            clear=False,
        ):
            self.assertEqual(
                runtime_quiz_bank_comment(
                    QUIZ2_BANK_ID,
                    "1",
                    language="Chinese",
                ),
                "数据库中的第一题反馈。",
            )
            self.assertEqual(
                runtime_quiz_bank_comment(
                    QUIZ2_BANK_ID,
                    "custom triangle keyword",
                    language="English",
                ),
                "Question one from the database.",
            )

    def test_duplicate_questions_and_invalid_patterns_are_rejected(self) -> None:
        quiz2 = self.store.load_bank(QUIZ2_BANK_ID)
        duplicate = [quiz2["entries"][0], dict(quiz2["entries"][0])]
        with self.assertRaises(QuizBankStoreError):
            self.store.save_bank(
                QUIZ2_BANK_ID,
                class_name=quiz2["class_name"],
                quiz_name=quiz2["quiz_name"],
                display_name=quiz2["display_name"],
                entries=duplicate,
            )

        invalid = [dict(quiz2["entries"][0], patterns=["("])]
        with self.assertRaises(QuizBankStoreError):
            self.store.save_bank(
                QUIZ2_BANK_ID,
                class_name=quiz2["class_name"],
                quiz_name=quiz2["quiz_name"],
                display_name=quiz2["display_name"],
                entries=invalid,
            )


if __name__ == "__main__":
    unittest.main()
