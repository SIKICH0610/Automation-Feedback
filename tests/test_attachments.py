from __future__ import annotations

from pathlib import Path
import sqlite3
import struct
from tempfile import TemporaryDirectory
import unittest

from openpyxl import Workbook

from attachment_store import AttachmentStoreError, SQLiteAttachmentStore
from database_store import StoreError
from frontend_server import ActionRunner, WorkbookStore
from paste_attachments import build_dropfiles_payload


class AttachmentStoreTest(unittest.TestCase):
    def setUp(self) -> None:
        self.temp_dir = TemporaryDirectory(dir=Path(__file__).parent)
        self.root = Path(self.temp_dir.name)
        self.database_path = self.root / "feedback.db"
        connection = sqlite3.connect(self.database_path)
        try:
            connection.executescript(
                """
                CREATE TABLE classes (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    name TEXT NOT NULL UNIQUE,
                    position INTEGER NOT NULL
                );
                INSERT INTO classes(name, position) VALUES('Class A', 0);
                INSERT INTO classes(name, position) VALUES('Class B', 1);
                """
            )
        finally:
            connection.close()
        self.store = SQLiteAttachmentStore(self.database_path)

    def tearDown(self) -> None:
        self.temp_dir.cleanup()

    def test_file_is_stored_listed_materialized_and_removed(self) -> None:
        content = b"%PDF-1.4\nattachment test"
        saved = self.store.add(
            "Class A",
            filename="quiz review.pdf",
            content_type="application/pdf",
            file_data=content,
        )
        self.assertEqual(saved["name"], "quiz review.pdf")
        self.assertEqual(saved["kind"], "document")
        self.assertEqual(saved["size"], len(content))
        self.assertNotIn("file_data", saved)
        self.assertEqual(self.store.list_for_sheet("Class B"), [])

        materialized = self.store.materialize(
            "Class A",
            [saved["id"]],
            self.root / "runtime",
        )
        self.assertEqual(materialized[0].read_bytes(), content)
        self.assertTrue(materialized[0].name.endswith("quiz review.pdf"))

        self.store.remove("Class A", saved["id"])
        self.assertEqual(self.store.list_for_sheet("Class A"), [])

    def test_invalid_extension_and_wrong_class_are_rejected(self) -> None:
        with self.assertRaises(AttachmentStoreError):
            self.store.add(
                "Class A",
                filename="program.exe",
                content_type="application/octet-stream",
                file_data=b"not allowed",
            )

        saved = self.store.add(
            "Class A",
            filename="diagram.png",
            content_type="image/png",
            file_data=b"image bytes",
        )
        with self.assertRaises(AttachmentStoreError):
            self.store.materialize(
                "Class B",
                [saved["id"]],
                self.root / "wrong-class",
            )


class AttachmentPasteTest(unittest.TestCase):
    def test_dropfiles_payload_contains_absolute_unicode_file_paths(self) -> None:
        with TemporaryDirectory(dir=Path(__file__).parent) as temporary:
            path = Path(temporary) / "class diagram.png"
            path.write_bytes(b"png")
            payload = build_dropfiles_payload([path])

        offset, x, y, non_client, wide = struct.unpack("<IiiII", payload[:20])
        self.assertEqual((offset, x, y, non_client, wide), (20, 0, 0, 0, 1))
        decoded = payload[offset:].decode("utf-16le")
        self.assertEqual(decoded, f"{path.resolve()}\0\0")


class AttachmentActionTest(unittest.TestCase):
    def setUp(self) -> None:
        self.temp_dir = TemporaryDirectory(dir=Path(__file__).parent)
        root = Path(self.temp_dir.name)
        workbook_path = root / "students.xlsx"
        workbook = Workbook()
        sheet = workbook.active
        sheet.title = "Class A"
        sheet.append(["First Name", "Last Name", "uid", "Feedback"])
        sheet.append(["Ada", "Lovelace", "1001", "A message"])
        sheet.append(["Grace", "Hopper", "1002", "Another message"])
        workbook.save(workbook_path)
        workbook.close()
        self.store = WorkbookStore(
            database_path=root / "app_data" / "feedback.db",
            source_workbook_path=workbook_path,
            announcement_dir=root / "announcements",
            app_data_dir=root / "app_data",
            export_dir=root / "exports",
        )

    def tearDown(self) -> None:
        self.temp_dir.cleanup()

    def test_action_command_passes_materialized_files_to_sender(self) -> None:
        attachment = self.store.add_attachment(
            "Class A",
            filename="review.pdf",
            content_type="application/pdf",
            file_data=b"pdf",
        )
        runtime_dir = Path(self.temp_dir.name) / "action-files"
        paths = self.store.materialize_attachments(
            "Class A",
            [attachment["id"]],
            runtime_dir,
        )
        runner = ActionRunner(self.store)
        command, _ = runner.command_for(
            {
                "action": "paste-announcement",
                "sheet": "Class A",
                "rows": [2],
                "attachment_ids": [attachment["id"]],
            },
            attachment_paths=paths,
        )
        attachment_index = command.index("--attachment")
        self.assertEqual(command[attachment_index + 1], str(paths[0]))

    def test_attachment_action_rejects_multiple_recipients(self) -> None:
        runner = ActionRunner(self.store)
        with self.assertRaises(StoreError):
            runner.command_for(
                {
                    "action": "paste-announcement",
                    "sheet": "Class A",
                    "rows": [2, 3],
                    "attachment_ids": ["attachment-id"],
                }
            )


if __name__ == "__main__":
    unittest.main()
