from __future__ import annotations

import unittest

from ai_polish import (
    _looks_like_keywords,
    _sanitize_draft,
    _strip_wrapping,
    attach_name,
    fact_guard,
    tone_examples,
)


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


class PluralGuardTest(unittest.TestCase):
    def test_sanitizer_downgrades_noun_plurals(self):
        self.assertEqual(_sanitize_draft("孩子们上课很认真，和同学们讨论积极～"), "孩子上课很认真，和同学讨论积极～")

    def test_sanitizer_naturalizes_child_references(self):
        self.assertEqual(_sanitize_draft("您孩子的表现非常认真，该生进步明显～"), "孩子的表现非常认真，孩子进步明显～")

    def test_sanitizer_removes_stock_courtesy(self):
        self.assertEqual(_sanitize_draft("请周四前提交作业。感谢您的配合！"), "请周四前提交作业。")

    def test_unfixable_group_address_is_rejected(self):
        self.assertIn("各位家长", fact_guard("announce", "带三角板", "各位家长请注意带三角板"))

    def test_singular_rewrite_passes(self):
        self.assertEqual(fact_guard("announce", "请同学们带三角板", "请帮孩子备好三角板～"), "")

    def test_sanitizer_degenders_pronouns(self):
        self.assertEqual(_sanitize_draft("她上课很认真，我每次点他都答得上来～"), "孩子上课很认真，我每次点孩子都答得上来～")

    def test_degender_keeps_real_compounds(self):
        self.assertEqual(_sanitize_draft("其他题目都做对了，做题不受他人影响。"), "其他题目都做对了，做题不受他人影响。")

    def test_degender_collapses_plural_and_stacked_pronouns(self):
        self.assertEqual(_sanitize_draft("他们的作业孩子他自己完成了。"), "孩子的作业孩子自己完成了。")

    def test_wave_only_closes_the_paragraph(self):
        self.assertEqual(_sanitize_draft("上课很认真～作业全对～"), "上课很认真。作业全对～")

    def test_wave_digit_ranges_survive(self):
        self.assertEqual(_sanitize_draft("下午3～4点上课，状态很好～"), "下午3～4点上课，状态很好～")

    def test_colons_and_dashes_demote_to_commas(self):
        self.assertEqual(
            _sanitize_draft("有个现象：偶尔走神——不过思路很好～"),
            "有个现象，偶尔走神，不过思路很好～",
        )

    def test_attach_name_merges_leading_role_noun(self):
        self.assertEqual(attach_name("Sunnie", "孩子上课很认真。"), "Sunnie上课很认真。")
        self.assertEqual(attach_name("Sunnie", "学生的作业完成得很好。"), "Sunnie的作业完成得很好。")
        self.assertEqual(attach_name("Sunnie", "这节课整体很稳。"), "Sunnie这节课整体很稳。")


class PolishCommentTest(unittest.TestCase):
    def test_only_middle_paragraph_is_rewritten(self):
        import ai_polish
        from unittest import mock

        comment = "家长您好～今天讲了平行线。\n\nSunnie原来的正文写得一般。\n\n请记得完成作业，谢谢配合。"
        stub = {"ok": True, "text": "孩子上课很认真。", "model": "m", "elapsed": 2.0}
        with mock.patch.object(ai_polish, "expand", return_value=stub):
            result = ai_polish.polish_comment(comment, "Sunnie", None)
        self.assertTrue(result["ok"])
        head, middle, tail = result["text"].split("\n\n")
        self.assertEqual(head, "家长您好～今天讲了平行线。")
        self.assertEqual(middle, "Sunnie上课很认真。")
        self.assertEqual(tail, "请记得完成作业，谢谢配合。")

    def test_two_paragraph_comment_is_refused(self):
        import ai_polish

        result = ai_polish.polish_comment("第一段\n\n第二段", "Sunnie", None)
        self.assertFalse(result["ok"])


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
