from __future__ import annotations

from contextlib import contextmanager
import hashlib
import mimetypes
from pathlib import Path
import re
import sqlite3
import threading
from typing import Any, Iterator
from uuid import uuid4


MAX_ATTACHMENT_BYTES = 100 * 1024 * 1024
IMAGE_EXTENSIONS = {".bmp", ".gif", ".jpeg", ".jpg", ".png", ".webp"}
DOCUMENT_EXTENSIONS = {
    ".csv",
    ".doc",
    ".docx",
    ".pdf",
    ".ppt",
    ".pptx",
    ".txt",
    ".xls",
    ".xlsx",
}
ALLOWED_EXTENSIONS = IMAGE_EXTENSIONS | DOCUMENT_EXTENSIONS


class AttachmentStoreError(Exception):
    pass


def safe_filename(filename: str) -> str:
    name = Path(str(filename or "").replace("\\", "/")).name.strip()
    name = re.sub(r"[\x00-\x1f<>:\"/\\|?*]+", "_", name).strip(" .")
    if not name:
        raise AttachmentStoreError("The uploaded file needs a valid filename.")
    return name[:180]


def attachment_kind(filename: str) -> str:
    suffix = Path(filename).suffix.lower()
    if suffix in IMAGE_EXTENSIONS:
        return "image"
    if suffix in DOCUMENT_EXTENSIONS:
        return "document"
    allowed = ", ".join(sorted(ALLOWED_EXTENSIONS))
    raise AttachmentStoreError(
        f"File type {suffix or '(none)'} is not supported. Allowed types are {allowed}."
    )


class SQLiteAttachmentStore:
    def __init__(
        self,
        database_path: Path,
        *,
        lock: threading.RLock | None = None,
    ) -> None:
        self.database_path = Path(database_path).resolve()
        self.lock = lock or threading.RLock()
        self._initialize_database()

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

    def _initialize_database(self) -> None:
        with self.lock, self._connect() as connection:
            connection.executescript(
                """
                CREATE TABLE IF NOT EXISTS attachments (
                    id TEXT PRIMARY KEY,
                    class_id INTEGER NOT NULL,
                    original_name TEXT NOT NULL,
                    content_type TEXT NOT NULL,
                    kind TEXT NOT NULL,
                    size_bytes INTEGER NOT NULL,
                    sha256 TEXT NOT NULL,
                    file_data BLOB NOT NULL,
                    created_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
                    FOREIGN KEY(class_id) REFERENCES classes(id) ON DELETE CASCADE
                );

                CREATE INDEX IF NOT EXISTS attachments_class_created
                ON attachments(class_id, created_at, original_name);
                """
            )

    @staticmethod
    def _class_id(connection: sqlite3.Connection, sheet_name: str) -> int:
        row = connection.execute(
            "SELECT id FROM classes WHERE name = ?",
            (sheet_name,),
        ).fetchone()
        if row is None:
            raise AttachmentStoreError(f"Sheet {sheet_name!r} was not found.")
        return int(row["id"])

    @staticmethod
    def _metadata(row: sqlite3.Row) -> dict[str, Any]:
        return {
            "id": str(row["id"]),
            "name": str(row["original_name"]),
            "content_type": str(row["content_type"]),
            "kind": str(row["kind"]),
            "size": int(row["size_bytes"]),
            "created_at": str(row["created_at"]),
        }

    def list_for_sheet(self, sheet_name: str) -> list[dict[str, Any]]:
        with self.lock, self._connect() as connection:
            class_id = self._class_id(connection, sheet_name)
            rows = connection.execute(
                """
                SELECT id, original_name, content_type, kind, size_bytes, created_at
                FROM attachments
                WHERE class_id = ?
                ORDER BY created_at, original_name
                """,
                (class_id,),
            ).fetchall()
            return [self._metadata(row) for row in rows]

    def add(
        self,
        sheet_name: str,
        *,
        filename: str,
        content_type: str,
        file_data: bytes,
    ) -> dict[str, Any]:
        name = safe_filename(filename)
        kind = attachment_kind(name)
        if not file_data:
            raise AttachmentStoreError("The uploaded file is empty.")
        if len(file_data) > MAX_ATTACHMENT_BYTES:
            raise AttachmentStoreError("Each attachment must be 100 MB or smaller.")

        normalized_type = mimetypes.guess_type(name)[0] or "application/octet-stream"

        attachment_id = str(uuid4())
        digest = hashlib.sha256(file_data).hexdigest()
        with self.lock, self._connect() as connection:
            class_id = self._class_id(connection, sheet_name)
            connection.execute(
                """
                INSERT INTO attachments(
                    id, class_id, original_name, content_type, kind,
                    size_bytes, sha256, file_data
                )
                VALUES(?, ?, ?, ?, ?, ?, ?, ?)
                """,
                (
                    attachment_id,
                    class_id,
                    name,
                    normalized_type,
                    kind,
                    len(file_data),
                    digest,
                    sqlite3.Binary(file_data),
                ),
            )
            row = connection.execute(
                """
                SELECT id, original_name, content_type, kind, size_bytes, created_at
                FROM attachments
                WHERE id = ?
                """,
                (attachment_id,),
            ).fetchone()
            if row is None:
                raise AttachmentStoreError("The attachment could not be stored.")
            return self._metadata(row)

    def remove(self, sheet_name: str, attachment_id: str) -> None:
        with self.lock, self._connect() as connection:
            class_id = self._class_id(connection, sheet_name)
            cursor = connection.execute(
                "DELETE FROM attachments WHERE id = ? AND class_id = ?",
                (str(attachment_id), class_id),
            )
            if cursor.rowcount != 1:
                raise AttachmentStoreError("The selected attachment was not found.")

    def get(self, attachment_id: str) -> dict[str, Any]:
        with self.lock, self._connect() as connection:
            row = connection.execute(
                """
                SELECT id, original_name, content_type, kind, size_bytes,
                       created_at, file_data
                FROM attachments
                WHERE id = ?
                """,
                (str(attachment_id),),
            ).fetchone()
            if row is None:
                raise AttachmentStoreError("The selected attachment was not found.")
            result = self._metadata(row)
            result["file_data"] = bytes(row["file_data"])
            return result

    def materialize(
        self,
        sheet_name: str,
        attachment_ids: list[str],
        target_dir: Path,
    ) -> list[Path]:
        ordered_ids = list(dict.fromkeys(str(value) for value in attachment_ids if str(value)))
        if not ordered_ids:
            return []

        target_dir.mkdir(parents=True, exist_ok=True)
        paths: list[Path] = []
        with self.lock, self._connect() as connection:
            class_id = self._class_id(connection, sheet_name)
            for index, attachment_id in enumerate(ordered_ids, start=1):
                row = connection.execute(
                    """
                    SELECT original_name, file_data
                    FROM attachments
                    WHERE id = ? AND class_id = ?
                    """,
                    (attachment_id, class_id),
                ).fetchone()
                if row is None:
                    raise AttachmentStoreError(
                        "One of the selected attachments no longer exists for this class."
                    )
                path = target_dir / f"{index:02d}_{safe_filename(str(row['original_name']))}"
                path.write_bytes(bytes(row["file_data"]))
                paths.append(path.resolve())
        return paths
