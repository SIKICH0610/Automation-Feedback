#!/usr/bin/env python3
"""Build the LoRA training set from the labeled corpus spreadsheet.

Reads ~/Downloads/语料汇总-labeling.xlsx, keeps rows the teacher marked 好,
applies the sentence fixes the teacher asked for in 备注, strips greetings and
student names, and writes:

    app_data/training/pairs_v1.jsonl    all approved 素材→评语 pairs (zh + en)
    app_data/training/singles_v1.jsonl  approved comments without 素材 (style bank)
    app_data/training/train.jsonl       MLX-LM chat format: authentic paragraphs
                                        behind a generic ask (majority, style
                                        objective) + 素材→评语 pairs (minority)
    app_data/training/valid.jsonl       held-out validation split (same mix)
    ~/Downloads/训练集v1-入选评语.xlsx    human-readable review copy

Label precedence per row: a 标注 typed into the sheet wins; LABEL_PATCH fills
rows whose cell is still empty (the first labeling round arrived via chat
before the sheet was saved); finally, unlabeled rows from the two handwritten
singles sections default to 好 — the teacher blanket-approved them on
2026-09-15 ("都是算好的") — minus the BULK_EXCLUDE rows that are not
parent-facing feedback at all. Re-run after more labeling:

    .venv/bin/python build_training_set.py
"""
from __future__ import annotations

import json
import random
import re
from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font

LABEL_XLSX = Path.home() / "Downloads" / "语料汇总-labeling.xlsx"
OUT_DIR = Path(__file__).resolve().parent / "app_data" / "training"
REVIEW_XLSX = Path.home() / "Downloads" / "训练集v1-入选评语.xlsx"
VALID_SIZE = 4
SEED = 42

# First labeling round (2026-09-15), transcribed from the teacher's pasted
# sheet; a non-empty 标注 cell in the file always overrides this.
_GOOD = ("好", "")
LABEL_PATCH: dict[int, tuple[str, str]] = {
    1: _GOOD,
    2: ("好", "有ai味"),
    3: ("不好", ""),
    4: ("待定", ""),
    5: ("不好", "有ai味"),
    6: ("不好", "有ai味"),
    7: _GOOD,
    8: _GOOD,
    9: ("不好", "不适合所有课程"),
    10: ("待定", ""),
    11: ("不好", "不适合所有课程"),
    12: ("待定", ""),
    13: ("待定", ""),
    14: ("好", "语气可以更友善"),
    15: ("好", "珍贵的多有点奇怪，可以改成“好习惯”"),
    16: ("好", "有ai味"),
    17: ("好", "有ai味"),
    18: _GOOD,
    19: _GOOD,
    20: _GOOD,
    21: _GOOD,
    22: _GOOD,
    23: ("好", "有ai味"),
    24: ("好", "有ai味，“这股劲我想帮她保住”很奇怪"),
    25: ("好", "有ai味，“所以安静不是没听进去，是性格。”很奇怪"),
    26: _GOOD,
    27: _GOOD,
    28: _GOOD,
    29: _GOOD,
    30: _GOOD,
    31: _GOOD,
    32: ("好", "有ai味"),
    33: _GOOD,
    34: ("好", "有ai味"),
    **{i: _GOOD for i in range(35, 49)},
}

# The teacher blanket-approved these handwritten sections; unlabeled rows from
# these sources count as 好 unless excluded below.
BULK_GOOD_SOURCES = ("student list/", "App数据库/Additional Comment")
BULK_EXCLUDE: dict[int, str] = {
    74: "缺勤询问，不是给家长的评语",
    108: "缺勤询问，不是给家长的评语",
    135: "句子截断（以冒号结尾）",
    146: "群发口吻（“孩子们”）且同时谈两个学生",
    149: "同时谈两个学生",
}
KEEP_FIRST_PARAGRAPH = {116}  # 一格里拼了多条消息，只保留第一条完整评语
# Point-to-point voice only: bulk-approved singles carrying a group address
# are dropped automatically (a hand-typed 好 in the sheet still wins).
GROUP_MARKERS = ("孩子们", "同学们", "学生们", "各位家长", "家长们")

# Voice attribution, confirmed by the teacher on 2026-09-15: all three source
# spreadsheets were written by OTHER teachers (student list is a mix), so the
# whole v1 corpus feeds the shared base model, not the teacher's personal
# adapter. Tags are per-table; merge tags later if two tables share an author.
VOICE_BY_SOURCE = (
    ("G5-Fall26/", "同事(G5表)"),
    ("Lesson Feedback/", "同事(LF表)"),
    ("student list/", "混合待分"),
    ("App数据库/Additional Comment", "同事(FallStudent表)"),
)


def _voice_for(source: str) -> str:
    for prefix, voice in VOICE_BY_SOURCE:
        if source.startswith(prefix):
            return voice
    return "未知"

# Sentence-level repairs the teacher requested in 备注, applied to the 评语
# before it enters the corpus. Keyed by 编号.
FIXUPS: dict[int, list[tuple[str, str]]] = {
    14: [
        ("，她有上了3D图形的课程不", "。"),
        ("，有任何问题随时找我～", "。"),
    ],
    15: [("这个习惯比多做对几道题珍贵得多", "这是非常好的习惯")],
    24: [("这股劲儿我想帮她保住", "希望她能把这股劲头保持下去")],
    25: [("所以安静不是没听进去，是性格。", "她只是性格偏安静，内容其实都听进去了。")],
    116: [("方便很多的做法", "方便很多的做法。")],
}
# Full-message rows: keep only the personal paragraph, from this substring on
# (the recap/greeting before it is the template's job, not the model's).
TRIM_TO: dict[int, str] = {
    14: "Joelle是第二节课才来的",
}

_GREETING = re.compile(r"^\s*(?:[A-Za-z][\w .'-]*)?家长您好[，,。！!～~\s]*")

# Two training objectives, style-first by the teacher's request (2026-09-15):
# the model should mainly absorb what AUTHENTIC feedback prose sounds like,
# not a keyword->phrase lookup table. Style samples (the majority) put the
# loss on real paragraphs behind a generic ask that carries no facts; pair
# samples (the minority) keep the "ground on the given facts" discipline the
# runtime task needs.
TRAIN_SYSTEM_PAIR = (
    "你是 Think Academy 的数学老师，正在给家长写孩子的课堂反馈。"
    "把老师的课堂速记/关键词改写成发给家长的评语：保留速记里的每个事实，不添加新事实；"
    "不写学生姓名和称呼，不写问候和结尾；这条消息只发给一位家长、只谈这一个孩子，"
    "用“孩子”做主语，绝不能出现“孩子们”“同学们”等群体称呼。"
)
TRAIN_SYSTEM_STYLE = (
    "你是 Think Academy 的数学老师，正在给家长写孩子的课堂反馈。"
    "用你平时的口吻写：只发给一位家长、只谈这一个孩子，用“孩子”做主语，"
    "绝不能出现“孩子们”“同学们”等群体称呼；不写学生姓名，不写问候和结尾。"
)
STYLE_MIN_CHARS = 14  # drop degenerate one-liners from training (corpus keeps them)
STYLE_VALID_SIZE = 8
# The singles corpus skews short (most are 20-60 char quick comments). Without
# length conditioning the v2 adapter learned "a proper answer is one short
# sentence" and started dropping facts / looping on long inputs, so each style
# sample now names its own length bucket and the model learns the authentic
# register PER LENGTH instead of averaging toward the shortest.
PAIR_UPWEIGHT = 2  # pairs appear twice in training to hold fact-grounding


def _style_ask(final: str) -> str:
    n = len(final)
    if n < 40:
        return "写一条发给家长的简短课堂短评，一两句话。"
    if n <= 90:
        return "写一条发给家长的课堂反馈短评，两三句话。"
    return "写一段完整的发给家长的课后反馈评语，三到五句话。"


def _chat_sample(entry: dict) -> dict:
    if entry["material"]:
        system = TRAIN_SYSTEM_PAIR
        user = f"课堂速记：{entry['material']}"
    else:
        system = TRAIN_SYSTEM_STYLE
        user = _style_ask(entry["final"])
    return {
        "messages": [
            {"role": "system", "content": system},
            {"role": "user", "content": user},
            {"role": "assistant", "content": entry["final"]},
        ]
    }


def _has_cjk(text: str) -> bool:
    return any("一" <= ch <= "鿿" for ch in text)


def _strip_name(text: str, student: str, lang: str) -> str:
    """Replace the student's name so the corpus teaches voice, not names."""
    result = _GREETING.sub("", text.strip())
    for token in (student or "").split():
        if len(token) < 3:
            continue  # too short to replace safely ("Q", "Me", "An")
        placeholder = "孩子" if lang == "zh" else "the student"
        # \b fails against CJK neighbours (they count as \w), so bound the
        # name against Latin letters only: "Katherine这节课" must still match.
        result = re.sub(
            rf"(?<![A-Za-z]){re.escape(token)}(?![A-Za-z])",
            placeholder,
            result,
            flags=re.IGNORECASE,
        )
    if lang == "zh":
        result = result.replace("孩子孩子", "孩子")
    else:
        # Capitalize placeholders that start a sentence.
        result = re.sub(
            r"(^|[.!?]\s+)the student", lambda m: m.group(1) + "The student", result
        )
    return result.strip()


def load_rows() -> list[dict]:
    wb = load_workbook(LABEL_XLSX, data_only=True)
    ws = wb.active
    rows = []
    for raw in ws.iter_rows(min_row=2, values_only=True):
        rid, rtype, material, final, _count, student, source, label, note = (
            (list(raw) + [None] * 9)[:9]
        )
        if rid is None or not final:
            continue
        rid = int(rid)
        src = str(source or "").strip()
        label = str(label).strip() if label else ""
        note = str(note).strip() if note else ""
        origin = "hand" if label else ""
        if not label and rid in LABEL_PATCH:
            label, note = LABEL_PATCH[rid]
            origin = "hand"
        if (
            not label
            and rid not in BULK_EXCLUDE
            and any(src.startswith(prefix) for prefix in BULK_GOOD_SOURCES)
        ):
            label, origin = "好", "bulk"
        rows.append(
            {
                "id": rid,
                "type": str(rtype or "").strip(),
                "material": str(material or "").strip(),
                "final": str(final or "").strip(),
                "student": str(student or "").strip(),
                "source": src,
                "label": label,
                "note": note,
                "origin": origin,
            }
        )
    return rows


def build() -> None:
    rows = load_rows()
    labeled = [r for r in rows if r["label"]]
    good = [r for r in labeled if r["label"] == "好"]
    dropped = [(r, BULK_EXCLUDE[r["id"]]) for r in rows if not r["label"] and r["id"] in BULK_EXCLUDE]

    pairs, singles = [], []
    for row in good:
        final = row["final"]
        if row["id"] in KEEP_FIRST_PARAGRAPH:
            final = next(block.strip() for block in final.split("\n") if block.strip())
        if row["id"] in TRIM_TO:
            anchor = TRIM_TO[row["id"]]
            if anchor in final:
                final = final[final.index(anchor):]
        for old, new in FIXUPS.get(row["id"], []):
            final = final.replace(old, new)
        if row["origin"] == "bulk" and any(marker in final for marker in GROUP_MARKERS):
            dropped.append((row, "群体称呼，自动剔除"))
            continue
        lang = "zh" if _has_cjk(final) else "en"
        entry = {
            "id": row["id"],
            "type": row["type"],
            "quality": "B" if row["note"] else "A",
            "lang": lang,
            "origin": "手标" if row["origin"] == "hand" else "批量",
            "voice": _voice_for(row["source"]),
            "source": row["source"],
            "material": _strip_name(row["material"], row["student"], lang),
            "final": _strip_name(final, row["student"], lang),
            "note": row["note"],
        }
        (pairs if entry["material"] else singles).append(entry)

    # The teacher sent some messages verbatim to several parents; keep one
    # copy of each such family so no phrasing is over-weighted.
    seen: dict[str, dict] = {}
    deduped = []
    for entry in singles:
        key = re.sub(r"\s+", " ", entry["final"]).strip().lower()
        if key in seen:
            seen[key]["dupes"] = seen[key].get("dupes", 1) + 1
            continue
        seen[key] = entry
        deduped.append(entry)
    collapsed = len(singles) - len(deduped)
    singles = deduped

    OUT_DIR.mkdir(parents=True, exist_ok=True)
    for name, data in (("pairs_v1.jsonl", pairs), ("singles_v1.jsonl", singles)):
        with open(OUT_DIR / name, "w", encoding="utf-8") as fh:
            for entry in data:
                fh.write(json.dumps(entry, ensure_ascii=False) + "\n")

    rng = random.Random(SEED)
    zh_pairs = [p for p in pairs if p["lang"] == "zh"]
    pair_pool = zh_pairs[:]
    rng.shuffle(pair_pool)
    pair_valid, pair_train = pair_pool[:VALID_SIZE], pair_pool[VALID_SIZE:]

    style_pool = [
        s for s in singles if s["lang"] == "zh" and len(s["final"]) >= STYLE_MIN_CHARS
    ]
    rng.shuffle(style_pool)
    style_valid, style_train = style_pool[:STYLE_VALID_SIZE], style_pool[STYLE_VALID_SIZE:]

    train = pair_train * PAIR_UPWEIGHT + style_train
    rng.shuffle(train)
    valid = pair_valid + style_valid
    for name, split in (("train.jsonl", train), ("valid.jsonl", valid)):
        with open(OUT_DIR / name, "w", encoding="utf-8") as fh:
            for entry in split:
                fh.write(json.dumps(_chat_sample(entry), ensure_ascii=False) + "\n")

    write_review(pairs, singles, dropped)

    zh_singles = sum(1 for s in singles if s["lang"] == "zh")
    voices: dict[str, int] = {}
    for entry in pairs + singles:
        voices[entry["voice"]] = voices.get(entry["voice"], 0) + 1
    print(f"sheet rows: {len(rows)}, labeled/approved: {len(labeled)}, 好: {len(good)}")
    print("voice breakdown:", ", ".join(f"{k} {v}" for k, v in sorted(voices.items())))
    print(f"pairs: {len(pairs)} (zh {len(zh_pairs)}, en {len(pairs) - len(zh_pairs)})")
    print(f"singles: {len(singles)} (zh {zh_singles}, en {len(singles) - zh_singles}), "
          f"collapsed dupes: {collapsed}, dropped: {len(dropped)}")
    print(f"train: {len(train)} (style {len(style_train)}, pair {len(pair_train)}), "
          f"valid: {len(valid)} (style {len(style_valid)}, pair {len(pair_valid)}) -> {OUT_DIR}")
    print(f"review copy -> {REVIEW_XLSX}")


def write_review(pairs: list[dict], singles: list[dict], dropped: list[tuple[dict, str]]) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "训练集v1"
    headers = ["原表编号", "类别", "质量", "语言", "素材（清洗后）", "评语（入选定稿）", "来源", "声音归属", "处理说明"]
    ws.append(headers)
    for cell in ws[1]:
        cell.font = Font(bold=True)

    def notes_for(entry: dict) -> str:
        notes = []
        if entry["id"] in FIXUPS or entry["id"] in TRIM_TO or entry["id"] in KEEP_FIRST_PARAGRAPH:
            notes.append("已按你的备注修改")
        if entry["origin"] == "批量":
            notes.append("批量入库（你确认整批算好）")
        if entry.get("dupes"):
            notes.append(f"同款消息共{entry['dupes']}条，仅保留这1条")
        if entry["note"]:
            notes.append(f"标注备注：{entry['note']}")
        return "；".join(notes)

    records = [("训练对", e) for e in pairs] + [("风格单条", e) for e in singles]
    for kind, entry in sorted(records, key=lambda item: item[1]["id"]):
        ws.append(
            [
                entry["id"],
                kind,
                entry["quality"],
                entry["lang"],
                entry["material"],
                entry["final"],
                entry["source"],
                entry["voice"],
                notes_for(entry),
            ]
        )
    for row, reason in sorted(dropped, key=lambda item: item[0]["id"]):
        ws.append(
            [
                row["id"],
                "剔除",
                "-",
                "-",
                row["material"],
                row["final"],
                row["source"],
                _voice_for(row["source"]),
                f"自动剔除：{reason}",
            ]
        )

    widths = {"A": 9, "B": 9, "C": 6, "D": 6, "E": 38, "F": 56, "G": 22, "H": 15, "I": 30}
    for col, width in widths.items():
        ws.column_dimensions[col].width = width
    wrap = Alignment(wrap_text=True, vertical="top")
    for row in ws.iter_rows(min_row=2):
        for cell in row:
            cell.alignment = wrap
    ws.freeze_panes = "A2"
    wb.save(REVIEW_XLSX)


if __name__ == "__main__":
    build()
