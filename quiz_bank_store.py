from __future__ import annotations

from contextlib import contextmanager
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


class QuizBankStoreError(Exception):
    pass


def _default_bank_seeds() -> tuple[dict[str, Any], ...]:
    from geometry_volume1_quiz1_comment_bank import GEOMETRY_VOLUME1_QUIZ1_ISSUES
    from geometry_volume1_quiz2_comment_bank import GEOMETRY_VOLUME1_QUIZ2_ISSUES

    return (
        {
            "id": QUIZ1_BANK_ID,
            "class_name": "Geometry Volume 1",
            "quiz_name": "Quiz 1",
            "display_name": "Geometry Volume 1 Quiz 1",
            "position": 0,
            "entries": GEOMETRY_VOLUME1_QUIZ1_ISSUES,
        },
        {
            "id": QUIZ2_BANK_ID,
            "class_name": "Geometry Volume 1",
            "quiz_name": "Quiz 2",
            "display_name": "Geometry Volume 1 Quiz 2",
            "position": 1,
            "entries": GEOMETRY_VOLUME1_QUIZ2_ISSUES,
        },
    )


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
                    patterns_json TEXT NOT NULL DEFAULT '[]',
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
                        id, bank_id, question, title, patterns_json,
                        chinese, english, position
                    )
                    VALUES(?, ?, ?, ?, ?, ?, ?, ?)
                    """,
                    (
                        str(uuid4()),
                        bank["id"],
                        issue.question,
                        issue.title,
                        json.dumps(list(issue.patterns), ensure_ascii=False),
                        issue.chinese,
                        issue.english,
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
            "patterns": list(json.loads(str(row["patterns_json"]))),
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
                SELECT id, question, title, patterns_json, chinese, english
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

    @staticmethod
    def _normalize_patterns(raw_patterns: Any, question: int) -> list[str]:
        if isinstance(raw_patterns, str):
            patterns = raw_patterns.splitlines()
        elif isinstance(raw_patterns, list):
            patterns = [str(value) for value in raw_patterns]
        else:
            raise QuizBankStoreError(
                f"Question {question} matching keywords must be a list."
            )

        normalized = list(dict.fromkeys(pattern.strip() for pattern in patterns if pattern.strip()))
        for pattern in normalized:
            try:
                re.compile(pattern, flags=re.IGNORECASE)
            except re.error as exc:
                raise QuizBankStoreError(
                    f"Question {question} has an invalid matching keyword pattern."
                ) from exc
        return normalized

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
                    "patterns": self._normalize_patterns(
                        raw_entry.get("patterns") or [],
                        question,
                    ),
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
                        id, bank_id, question, title, patterns_json,
                        chinese, english, position
                    )
                    VALUES(?, ?, ?, ?, ?, ?, ?, ?)
                    """,
                    (
                        str(uuid4()),
                        bank_id,
                        entry["question"],
                        entry["title"],
                        json.dumps(entry["patterns"], ensure_ascii=False),
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
            matched = entry["question"] in question_numbers
            if not matched:
                matched = any(
                    re.search(pattern, note, flags=re.IGNORECASE)
                    for pattern in entry["patterns"]
                )
            if not matched:
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
