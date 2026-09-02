from __future__ import annotations

import tempfile
import unittest
from pathlib import Path

from openpyxl import Workbook

from import_enrollment import group_by_class, read_enrollment_rows


def write_workbook(build) -> Path:
    workbook = Workbook()
    build(workbook)
    handle = tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False)
    workbook.save(handle.name)
    return Path(handle.name)


class ImportEnrollmentHeaderTest(unittest.TestCase):
    def test_exact_platform_format_still_imports(self) -> None:
        def build(workbook):
            sheet = workbook.active
            sheet.append(
                ["classId", "className", "classTimeDescription", "subject",
                 "firstName", "lastName", "学员id", "payStatus", "是否入班"]
            )
            sheet.append(["123", "Geo A", "Sat 10:00", "Math", "Amy", "Wang", "111", "Paid", "是"])

        rows = read_enrollment_rows(write_workbook(build))
        groups = group_by_class(rows, class_ids=None)
        self.assertEqual(list(groups), ["123"])
        self.assertEqual(groups["123"]["students"][0]["uid"], "111")
        self.assertEqual(groups["123"]["weekly_time"], "Sat 10:00")

    def test_renamed_reordered_headers_with_banner_row_and_extra_columns(self) -> None:
        def build(workbook):
            junk = workbook.active
            junk.title = "说明"
            junk.append(["这是平台导出的说明页，请勿修改"])

            sheet = workbook.create_sheet("学员明细")
            sheet.append(["学员报名明细表"])          # banner row
            sheet.append([])                          # blank row
            sheet.append(                             # header row 3: renamed, reordered, extras
                ["Pay Status", "Student_ID", "微信号", "Last Name", "First Name",
                 "Class ID", "班级名称", "上课时间", "新增列", "是否入班"]
            )
            sheet.append(["Paid", "222", "wx1", "Lin", "Bob", "456", "Geo B", "周六 10:00", "x", "是"])
            sheet.append(["paid", "333", "wx2", "Chen", "Cara", "456", "Geo B", "周六 10:00", "x", "是"])
            sheet.append(["Unpaid", "444", "wx3", "Wu", "Dan", "456", "Geo B", "周六 10:00", "x", "是"])
            sheet.append(["Paid", "555", "wx4", "Hu", "Eve", "456", "Geo B", "周六 10:00", "x", "否"])
            sheet.append(["Paid", "666", "wx5", "Li", "", "456", "Geo B", "周六 10:00", "x", "是"])

        rows = read_enrollment_rows(write_workbook(build))
        groups = group_by_class(rows, class_ids=None)
        self.assertEqual(list(groups), ["456"])
        uids = [s["uid"] for s in groups["456"]["students"]]
        # Bob (exact) and Cara (lowercase paid) import; Unpaid, 否, and the
        # blank-first-name row are filtered out.
        self.assertEqual(uids, ["222", "333"])
        self.assertEqual(groups["456"]["name"], "Geo B")
        # 上课时间 matched through its alias and came through intact.
        self.assertEqual(groups["456"]["weekly_time"], "周六 10:00")

    def test_missing_class_time_fails_loudly(self) -> None:
        def build(workbook):
            sheet = workbook.active
            sheet.append(["classId", "className", "firstName", "lastName", "学员id", "payStatus", "是否入班"])
            sheet.append(["1", "C", "A", "B", "9", "Paid", "是"])

        with self.assertRaises(ValueError) as caught:
            read_enrollment_rows(write_workbook(build))
        self.assertIn("classTimeDescription", str(caught.exception))

    def test_missing_required_column_error_lists_what_was_seen(self) -> None:
        def build(workbook):
            sheet = workbook.active
            sheet.append(["classId", "className", "firstName", "lastName", "payStatus", "是否入班"])
            sheet.append(["1", "C", "A", "B", "Paid", "是"])

        with self.assertRaises(ValueError) as caught:
            read_enrollment_rows(write_workbook(build))
        message = str(caught.exception)
        self.assertIn("学员id", message)
        self.assertIn("classId", message)  # the headers it DID see


if __name__ == "__main__":
    unittest.main()
