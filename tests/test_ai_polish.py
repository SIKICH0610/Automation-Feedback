from __future__ import annotations

import unittest

from ai_polish import _looks_like_keywords, _strip_wrapping, fact_guard, tone_examples


class FakeStore:
    def __init__(self, comments):
        self._comments = comments

    def sheet_names(self):
        return ["A"]

    def load_sheet(self, name):
        return {"rows": [{"values": {"Additional Comment": c}} for c in self._comments]}


class DigitGuardTest(unittest.TestCase):
    def test_expansion_may_not_add_digits(self):
        self.assertEqual(fact_guard("student", "笔记认真", "表现稳定～"), "")
        self.assertIn("8", fact_guard("student", "笔记认真", "quiz 得了 8 分～"))

    def test_expansion_may_repeat_given_digits(self):
        self.assertEqual(fact_guard("recap", "例题3道", "围绕例题3道展开～"), "")

    def test_announce_must_keep_every_digit(self):
        source = "10月3日下午2点，提前15分钟"
        self.assertEqual(fact_guard("announce", source, "10月3日下午2点开始，请提前15分钟到"), "")
        self.assertIn("15", fact_guard("announce", source, "10月3日下午2点开始"))

    def test_announce_rejects_invented_temporal_gloss(self):
        source = "家长可以在3点45分来接"
        bad = "考试结束后，即3点45分来接孩子"
        self.assertIn("结束后", fact_guard("announce", source, bad))
        ok = "家长可以在下午3点45分来接孩子"
        self.assertEqual(fact_guard("announce", source, ok), "")

    def test_student_kind_ignores_lost_digits(self):
        # Expansion输入里的数字不强制出现（老师关键词可能含题号但成稿不必逐字带出）
        self.assertEqual(fact_guard("student", "第3题粗心", "做题有些粗心，可以放慢检查～"), "")


class ToneExamplesTest(unittest.TestCase):
    def test_filters_digits_and_length(self):
        store = FakeStore([
            "上课听得很专注，笔记记得完整，会主动提问～",       # good
            "quiz 得了 7/8，很不错",                              # digits -> excluded
            "短",                                                  # too short -> excluded
            "口吻示例" * 40,                                        # too long -> excluded
        ])
        examples = tone_examples(store, "笔记 专注")
        self.assertEqual(examples, ["上课听得很专注，笔记记得完整，会主动提问～"])

    def test_falls_back_when_no_local_corpus(self):
        self.assertTrue(len(tone_examples(FakeStore([]), "任意")) > 0)


class KeywordDetectionTest(unittest.TestCase):
    def test_short_bare_topics_are_keywords(self):
        self.assertTrue(_looks_like_keywords("等腰三角形 三线合一 课堂练习"))

    def test_finished_prose_is_not(self):
        self.assertFalse(_looks_like_keywords(
            "家长您好～我们现在已经完成了前两讲的学习，第三节课将开始全等三角形～"
        ))
        self.assertFalse(_looks_like_keywords("今天复习了圆的性质。"))


class StripWrappingTest(unittest.TestCase):
    def test_strips_labels_and_quotes(self):
        self.assertEqual(_strip_wrapping('润色后："大家好～"'), "大家好～")
        self.assertEqual(_strip_wrapping("扩写后： 内容～ "), "内容～")


if __name__ == "__main__":
    unittest.main()
