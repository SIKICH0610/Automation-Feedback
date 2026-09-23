from __future__ import annotations

import tempfile
import unittest
import zipfile
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

    def test_strict_ooxml_file_still_imports(self) -> None:
        # The platform exports Strict OOXML (purl.oclc.org namespaces), which
        # openpyxl silently loads as zero worksheets. Synthesize one by rewriting
        # a normal file's namespaces to the strict form.
        def build(workbook):
            sheet = workbook.active
            sheet.append(
                ["classId", "className", "classTimeDescription", "subject",
                 "firstName", "lastName", "学员id", "payStatus", "是否入班"]
            )
            sheet.append(["77", "Geo S", "Sun 09:00", "Math", "Zoe", "Xu", "888", "Paid", "是"])

        normal = write_workbook(build)
        strict = normal.with_name("strict.xlsx")
        with zipfile.ZipFile(normal) as source, zipfile.ZipFile(strict, "w") as target:
            for item in source.infolist():
                data = source.read(item.filename)
                if item.filename.endswith((".xml", ".rels")):
                    data = data.replace(
                        b"http://schemas.openxmlformats.org/spreadsheetml/2006/main",
                        b"http://purl.oclc.org/ooxml/spreadsheetml/main",
                    ).replace(
                        b"http://schemas.openxmlformats.org/officeDocument/2006/relationships",
                        b"http://purl.oclc.org/ooxml/officeDocument/relationships",
                    )
                target.writestr(item, data)

        groups = group_by_class(read_enrollment_rows(strict), class_ids=None)
        self.assertEqual(list(groups), ["77"])
        self.assertEqual(groups["77"]["students"][0]["uid"], "888")

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


class ImportEnrollmentFormatTest(unittest.TestCase):
    HEADER = "classId,className,classTimeDescription,subject,firstName,lastName,学员id,payStatus,是否入班"

    def _write(self, data: bytes, suffix: str) -> Path:
        handle = tempfile.NamedTemporaryFile(suffix=suffix, delete=False)
        handle.write(data)
        handle.close()
        return Path(handle.name)

    def test_csv_with_utf8_bom(self) -> None:
        text = self.HEADER + "\n11,Geo C,Sun 09:00,Math,Ann,Lee,901,Paid,是\n"
        path = self._write(b"\xef\xbb\xbf" + text.encode("utf-8"), ".csv")
        groups = group_by_class(read_enrollment_rows(path), class_ids=None)
        self.assertEqual(groups["11"]["students"][0]["uid"], "901")

    def test_csv_saved_by_chinese_excel_gbk(self) -> None:
        text = self.HEADER + "\n12,Geo D,Sun 09:00,Math,Bo,Han,902,Paid,是\n"
        path = self._write(text.encode("gb18030"), ".csv")
        groups = group_by_class(read_enrollment_rows(path), class_ids=None)
        self.assertEqual(groups["12"]["students"][0]["uid"], "902")

    def test_json_bare_list_with_alias_keys(self) -> None:
        payload = [
            {"Class ID": 13, "className": "Geo E", "上课时间": "Sat 1pm", "First Name": "Cy",
             "lastName": "Wu", "studentId": 903, "Pay Status": "paid", "是否入班": "是", "extra": 1},
            {"Class ID": 13, "className": "Geo E", "上课时间": "Sat 1pm", "First Name": "Di",
             "lastName": "Xu", "studentId": 904, "Pay Status": "Unpaid", "是否入班": "是"},
        ]
        import json as jsonlib
        path = self._write(jsonlib.dumps(payload, ensure_ascii=False).encode("utf-8"), ".json")
        groups = group_by_class(read_enrollment_rows(path), class_ids=None)
        uids = [s["uid"] for s in groups["13"]["students"]]
        self.assertEqual(uids, ["903"])  # Unpaid filtered out

    def test_json_wrapped_in_platform_envelope(self) -> None:
        import json as jsonlib
        payload = {"code": 0, "data": {"total": 1, "list": [
            {"classId": "14", "className": "Geo F", "classTimeDescription": "Sun",
             "firstName": "Ed", "lastName": "Yao", "学员id": "905", "payStatus": "Paid", "是否入班": "是"}
        ]}}
        path = self._write(jsonlib.dumps(payload, ensure_ascii=False).encode("utf-8"), ".json")
        groups = group_by_class(read_enrollment_rows(path), class_ids=None)
        self.assertEqual(groups["14"]["students"][0]["uid"], "905")

    def test_misnamed_file_is_sniffed_by_content(self) -> None:
        # xlsx bytes with a .csv name still import: format comes from content.
        def build(workbook):
            sheet = workbook.active
            sheet.append(["classId", "className", "classTimeDescription", "subject",
                          "firstName", "lastName", "学员id", "payStatus", "是否入班"])
            sheet.append(["15", "Geo G", "Fri", "Math", "Fay", "Zhu", "906", "Paid", "是"])

        real = write_workbook(build)
        misnamed = real.with_suffix(".csv")
        misnamed.write_bytes(real.read_bytes())
        groups = group_by_class(read_enrollment_rows(misnamed), class_ids=None)
        self.assertEqual(groups["15"]["students"][0]["uid"], "906")

    def test_legacy_xls_gets_a_clear_error(self) -> None:
        path = self._write(b"\xd0\xcf\x11\xe0junkjunk", ".xls")
        with self.assertRaises(ValueError) as caught:
            read_enrollment_rows(path)
        self.assertIn(".xlsx", str(caught.exception))


if __name__ == "__main__":
    unittest.main()
