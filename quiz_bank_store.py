from __future__ import annotations

from contextlib import contextmanager
import csv
from functools import lru_cache
import json
import os
from pathlib import Path
import re
import sqlite3
import threading
from typing import Any, Iterator
from uuid import uuid4


QUIZ1_BANK_ID = "geometry-volume-1-quiz-1"
QUIZ2_BANK_ID = "geometry-volume-1-quiz-2"
AMC10_QUIZ1_BANK_ID = "amc10-intro-quiz-1"

# One CSV per bank. Adding a course means dropping in another CSV (columns:
# bank_id, class_name, quiz_name, display_name, question, title, chinese,
# english) -- no code. Files seed the sqlite store on first run only, keyed by
# bank_id with INSERT OR IGNORE, so entries the teacher edited in the app are
# never overwritten.
QUIZ_BANK_SEED_DIR = Path(__file__).resolve().parent / "data" / "quiz_banks"


class QuizBankStoreError(Exception):
    pass


def _default_bank_seeds() -> tuple[dict[str, Any], ...]:
    if not QUIZ_BANK_SEED_DIR.is_dir():
        return ()
    banks: dict[str, dict[str, Any]] = {}
    order: list[str] = []
    for path in sorted(QUIZ_BANK_SEED_DIR.glob("*.csv")):
        with open(path, newline="", encoding="utf-8-sig") as handle:
            for row in csv.DictReader(handle):
                bank_id = str(row.get("bank_id") or "").strip()
                if not bank_id:
                    continue
                bank = banks.get(bank_id)
                if bank is None:
                    bank = {
                        "id": bank_id,
                        "class_name": str(row.get("class_name") or "").strip(),
                        "quiz_name": str(row.get("quiz_name") or "").strip(),
                        "display_name": str(row.get("display_name") or bank_id).strip(),
                        "entries": [],
                    }
                    banks[bank_id] = bank
                    order.append(bank_id)
                question = str(row.get("question") or "").strip()
                if not question:
                    continue
                bank["entries"].append(
                    {
                        "question": int(question),
                        "title": str(row.get("title") or "").strip(),
                        "chinese": str(row.get("chinese") or "").strip(),
                        "english": str(row.get("english") or "").strip(),
                    }
                )
    for position, bank_id in enumerate(order):
        banks[bank_id]["position"] = position
    return tuple(banks[bank_id] for bank_id in order)


def question_numbers_from_note(note: str) -> set[int]:
    stripped = re.sub(
        r"\b(?:physical\s*)?quiz\s*\d+\b|\b(?:first|second|third)\s+physical\s+quiz\b",
        " ",
        note.strip(),
        flags=re.IGNORECASE,
    ).strip(" -:：,，、;；")
    numbers: set[int] = set()

    for start, end in re.findall(
        r"\b(\d{1,2})\s*(?:-|–|—|~|至|到)\s*(\d{1,2})\b",
        stripped,
    ):
        lower, upper = sorted((int(start), int(end)))
        numbers.update(range(lower, upper + 1))

    for pattern in (
        r"\bq(?:uestion)?\s*(\d{1,2})\b",
        r"第\s*(\d{1,2})\s*题",
    ):
        numbers.update(
            int(value) for value in re.findall(pattern, stripped, flags=re.IGNORECASE)
        )

    if re.fullmatch(
        r"(?:\s*(?:q(?:uestion)?\s*)?\d{1,2}\s*[,，、;；–—~至到-]?)+\s*",
        stripped,
        flags=re.IGNORECASE,
    ):
        numbers.update(int(value) for value in re.findall(r"\d{1,2}", stripped))

    return numbers


class SQLiteQuizBankStore:
    def __init__(
        self,
        database_path: Path,
        *,
        lock: threading.RLock | None = None,
        seed_defaults: bool = True,
    ) -> None:
        self.database_path = database_path.resolve()
        self.lock = lock or threading.RLock()
        self.database_path.parent.mkdir(parents=True, exist_ok=True)
        self._initialize(seed_defaults=seed_defaults)

    @contextmanager
    def _connect(self) -> Iterator[sqlite3.Connection]:
        connection = sqlite3.connect(self.database_path, timeout=30)
        connection.row_factory = sqlite3.Row
        connection.execute("PRAGMA foreign_keys = ON")
        try:
            yield connection
        except Exception:
            connection.rollback()
            raise
        else:
            connection.commit()
        finally:
            connection.close()

    @staticmethod
    def _drop_patterns_column(connection: sqlite3.Connection) -> None:
        """Keyword matching is gone: entries match by question number only, so the
        stored patterns column is dropped. Guarded so an older SQLite that cannot
        drop columns leaves it behind harmlessly (it has a default)."""
        columns = {row[1] for row in connection.execute("PRAGMA table_info(quiz_bank_entries)")}
        if "patterns_json" in columns:
            try:
                connection.execute("ALTER TABLE quiz_bank_entries DROP COLUMN patterns_json")
            except sqlite3.OperationalError:
                pass

    def _initialize(self, *, seed_defaults: bool) -> None:
        with self.lock, self._connect() as connection:
            connection.executescript(
                """
                CREATE TABLE IF NOT EXISTS quiz_banks (
                    id TEXT PRIMARY KEY,
                    class_name TEXT NOT NULL,
                    quiz_name TEXT NOT NULL,
                    display_name TEXT NOT NULL,
                    position INTEGER NOT NULL DEFAULT 0,
                    created_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
                    updated_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP
                );

                CREATE TABLE IF NOT EXISTS quiz_bank_entries (
                    id TEXT PRIMARY KEY,
                    bank_id TEXT NOT NULL,
                    question INTEGER NOT NULL,
                    title TEXT NOT NULL DEFAULT '',
                    chinese TEXT NOT NULL DEFAULT '',
                    english TEXT NOT NULL DEFAULT '',
                    position INTEGER NOT NULL DEFAULT 0,
                    created_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
                    updated_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
                    UNIQUE(bank_id, question),
                    FOREIGN KEY(bank_id) REFERENCES quiz_banks(id) ON DELETE CASCADE
                );

                CREATE INDEX IF NOT EXISTS quiz_bank_entries_order
                ON quiz_bank_entries(bank_id, position, question);
                """
            )
            self._drop_patterns_column(connection)
            if seed_defaults:
                self._seed_defaults(connection)

    @staticmethod
    def _seed_defaults(connection: sqlite3.Connection) -> None:
        for bank in _default_bank_seeds():
            cursor = connection.execute(
                """
                INSERT OR IGNORE INTO quiz_banks(
                    id, class_name, quiz_name, display_name, position
                )
                VALUES(?, ?, ?, ?, ?)
                """,
                (
                    bank["id"],
                    bank["class_name"],
                    bank["quiz_name"],
                    bank["display_name"],
                    bank["position"],
                ),
            )
            if cursor.rowcount != 1:
                continue
            for position, issue in enumerate(bank["entries"]):
                connection.execute(
                    """
                    INSERT INTO quiz_bank_entries(
                        id, bank_id, question, title,
                        chinese, english, position
                    )
                    VALUES(?, ?, ?, ?, ?, ?, ?)
                    """,
                    (
                        str(uuid4()),
                        bank["id"],
                        issue["question"],
                        issue["title"],
                        issue["chinese"],
                        issue["english"],
                        position,
                    ),
                )

    @staticmethod
    def _bank_row(connection: sqlite3.Connection, bank_id: str) -> sqlite3.Row:
        row = connection.execute(
            """
            SELECT id, class_name, quiz_name, display_name, position
            FROM quiz_banks
            WHERE id = ?
            """,
            (bank_id,),
        ).fetchone()
        if row is None:
            raise QuizBankStoreError(f"Quiz bank {bank_id!r} was not found.")
        return row

    @staticmethod
    def _entry_dict(row: sqlite3.Row) -> dict[str, Any]:
        return {
            "entry_id": str(row["id"]),
            "question": int(row["question"]),
            "title": str(row["title"]),
            "chinese": str(row["chinese"]),
            "english": str(row["english"]),
        }

    def list_banks(self) -> list[dict[str, Any]]:
        with self.lock, self._connect() as connection:
            rows = connection.execute(
                """
                SELECT b.id, b.class_name, b.quiz_name, b.display_name, b.position,
                       COUNT(e.id) AS entry_count
                FROM quiz_banks AS b
                LEFT JOIN quiz_bank_entries AS e ON e.bank_id = b.id
                GROUP BY b.id
                ORDER BY b.position, b.display_name
                """
            ).fetchall()
            return [
                {
                    "id": str(row["id"]),
                    "class_name": str(row["class_name"]),
                    "quiz_name": str(row["quiz_name"]),
                    "display_name": str(row["display_name"]),
                    "entry_count": int(row["entry_count"]),
                }
                for row in rows
            ]

    def load_bank(self, bank_id: str) -> dict[str, Any]:
        with self.lock, self._connect() as connection:
            bank = self._bank_row(connection, bank_id)
            entries = connection.execute(
                """
                SELECT id, question, title, chinese, english
                FROM quiz_bank_entries
                WHERE bank_id = ?
                ORDER BY position, question
                """,
                (bank_id,),
            ).fetchall()
            return {
                "id": str(bank["id"]),
                "class_name": str(bank["class_name"]),
                "quiz_name": str(bank["quiz_name"]),
                "display_name": str(bank["display_name"]),
                "entries": [self._entry_dict(row) for row in entries],
            }

    def save_bank(
        self,
        bank_id: str,
        *,
        class_name: str,
        quiz_name: str,
        display_name: str,
        entries: Any,
    ) -> dict[str, Any]:
        if not isinstance(entries, list):
            raise QuizBankStoreError("Quiz bank entries must be provided as a list.")

        normalized_entries: list[dict[str, Any]] = []
        seen_questions: set[int] = set()
        for raw_entry in entries:
            if not isinstance(raw_entry, dict):
                raise QuizBankStoreError("Each quiz bank entry must be an object.")
            try:
                question = int(raw_entry.get("question"))
            except (TypeError, ValueError) as exc:
                raise QuizBankStoreError("Every quiz bank entry needs a question number.") from exc
            if question < 1 or question > 999:
                raise QuizBankStoreError("Question numbers must be between 1 and 999.")
            if question in seen_questions:
                raise QuizBankStoreError(f"Question {question} appears more than once.")
            seen_questions.add(question)

            chinese = str(raw_entry.get("chinese") or "").strip()
            english = str(raw_entry.get("english") or "").strip()
            if not chinese and not english:
                raise QuizBankStoreError(
                    f"Question {question} needs Chinese or English feedback."
                )
            normalized_entries.append(
                {
                    "question": question,
                    "title": str(raw_entry.get("title") or "").strip(),
                    "chinese": chinese,
                    "english": english,
                }
            )

        normalized_entries.sort(key=lambda item: item["question"])
        safe_class_name = class_name.strip()
        safe_quiz_name = quiz_name.strip()
        safe_display_name = display_name.strip()
        if not safe_class_name or not safe_quiz_name or not safe_display_name:
            raise QuizBankStoreError("Quiz bank names cannot be blank.")

        with self.lock, self._connect() as connection:
            self._bank_row(connection, bank_id)
            connection.execute(
                """
                UPDATE quiz_banks
                SET class_name = ?, quiz_name = ?, display_name = ?,
                    updated_at = CURRENT_TIMESTAMP
                WHERE id = ?
                """,
                (safe_class_name, safe_quiz_name, safe_display_name, bank_id),
            )
            connection.execute(
                "DELETE FROM quiz_bank_entries WHERE bank_id = ?",
                (bank_id,),
            )
            for position, entry in enumerate(normalized_entries):
                connection.execute(
                    """
                    INSERT INTO quiz_bank_entries(
                        id, bank_id, question, title,
                        chinese, english, position
                    )
                    VALUES(?, ?, ?, ?, ?, ?, ?)
                    """,
                    (
                        str(uuid4()),
                        bank_id,
                        entry["question"],
                        entry["title"],
                        entry["chinese"],
                        entry["english"],
                        position,
                    ),
                )

        return self.load_bank(bank_id)

    def build_comment(
        self,
        bank_id: str,
        note: str,
        *,
        language: str = "Chinese",
    ) -> str:
        bank = self.load_bank(bank_id)
        question_numbers = question_numbers_from_note(note)
        is_chinese = language.lower().startswith("chinese")
        comments: list[str] = []
        for entry in bank["entries"]:
            if entry["question"] not in question_numbers:
                continue
            comment = entry["chinese"] if is_chinese else entry["english"]
            if comment:
                comments.append(comment)
        return "".join(comments) if is_chinese else " ".join(comments)


@lru_cache(maxsize=4)
def _runtime_store(database_path: str) -> SQLiteQuizBankStore:
    return SQLiteQuizBankStore(Path(database_path))


def runtime_quiz_bank_comment(
    bank_id: str,
    note: str,
    *,
    language: str,
) -> str | None:
    database_path = os.environ.get("FEEDBACK_DATABASE_PATH", "").strip()
    if not database_path:
        return None
    path = Path(database_path)
    if not path.exists():
        return None
    return _runtime_store(str(path.resolve())).build_comment(
        bank_id,
        note,
        language=language,
    )
