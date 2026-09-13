from __future__ import annotations

"""Local AI expansion/polish through Ollama.

Three kinds of help, all reviewed by the teacher before anything is sent:
- student: a few keywords about one student -> a short parent-facing paragraph
- recap:   topic keywords -> the one-line lesson recap
- announce: polish an already-written announcement without touching its facts

Design rules, learned from the August experiments:
- The model may only elaborate LANGUAGE, never facts. Digits are the enforcement
  point: any digit in the output that is not in the input fails the draft (and for
  announcements, losing a digit fails too). One retry at lower temperature, then
  give up with a clear message.
- Tone comes from the teacher's own past comments: a few of their real
  Additional Comment lines are picked from the local database as style examples
  at call time. Nothing is bundled; every machine learns from its own data.
- No Ollama, no button: availability is probed, never assumed.
"""

import json
import re
import subprocess
import sys
import time
import urllib.error
import urllib.request
from typing import Any

OLLAMA_BASE = "http://127.0.0.1:11434"
# Preference order per memory tier. 7B reads noticeably more natural; 3B keeps
# 8GB machines usable.
MODEL_LARGE = "qwen2.5:7b"
MODEL_SMALL = "qwen2.5:3b"
LARGE_MODEL_MIN_BYTES = 12 * 1024**3

CORPUS_FILENAME = "ai_corpus.jsonl"

_DIGITS = re.compile(r"\d+")

# Temporal-relation glosses the model loves to invent around times ("接孩子" gets
# glossed as 考试结束后 no matter how the prompt forbids it -- observed 6/6 runs on
# qwen2.5:7b). If the source didn't state the relation, a draft that adds one is
# rejected mechanically instead of argued with.
_INFERENTIAL_MARKERS = ("结束后", "结束前", "开始前", "开始后")


def _http_json(path: str, payload: dict[str, Any] | None = None, *, timeout: float) -> Any:
    url = f"{OLLAMA_BASE}{path}"
    if payload is None:
        request = urllib.request.Request(url)
    else:
        request = urllib.request.Request(
            url,
            data=json.dumps(payload).encode("utf-8"),
            headers={"Content-Type": "application/json"},
            method="POST",
        )
    with urllib.request.urlopen(request, timeout=timeout) as response:
        return json.loads(response.read())


def ollama_available() -> bool:
    try:
        _http_json("/api/version", timeout=1.5)
        return True
    except Exception:
        return False


def installed_models() -> set[str]:
    try:
        data = _http_json("/api/tags", timeout=3)
    except Exception:
        return set()
    return {str(item.get("name") or "") for item in data.get("models", [])}


def _memory_bytes() -> int:
    try:
        if sys.platform == "darwin":
            out = subprocess.run(
                ["sysctl", "-n", "hw.memsize"], capture_output=True, text=True, timeout=5
            )
            return int(out.stdout.strip())
        if sys.platform == "win32":
            import ctypes

            class MEMORYSTATUSEX(ctypes.Structure):
                _fields_ = [
                    ("dwLength", ctypes.c_ulong),
                    ("dwMemoryLoad", ctypes.c_ulong),
                    ("ullTotalPhys", ctypes.c_ulonglong),
                    ("ullAvailPhys", ctypes.c_ulonglong),
                    ("ullTotalPageFile", ctypes.c_ulonglong),
                    ("ullAvailPageFile", ctypes.c_ulonglong),
                    ("ullTotalVirtual", ctypes.c_ulonglong),
                    ("ullAvailVirtual", ctypes.c_ulonglong),
                    ("ullAvailExtendedVirtual", ctypes.c_ulonglong),
                ]

            status = MEMORYSTATUSEX()
            status.dwLength = ctypes.sizeof(MEMORYSTATUSEX)
            ctypes.windll.kernel32.GlobalMemoryStatusEx(ctypes.byref(status))
            return int(status.ullTotalPhys)
    except Exception:
        pass
    return 8 * 1024**3


def ai_status() -> dict[str, Any]:
    """What the frontend needs to decide whether to show the buttons."""
    if not ollama_available():
        return {
            "available": False,
            "model": None,
            "hint": "未检测到 Ollama。安装并启动后按钮会自动出现：brew install ollama && ollama serve",
        }
    models = installed_models()
    preferred = MODEL_LARGE if _memory_bytes() >= LARGE_MODEL_MIN_BYTES else MODEL_SMALL
    for candidate in (preferred, MODEL_SMALL, MODEL_LARGE):
        if candidate in models:
            return {"available": True, "model": candidate, "hint": ""}
    return {
        "available": False,
        "model": None,
        "hint": f"Ollama 已运行，但缺少模型。请执行：ollama pull {preferred}",
    }


# --- tone examples -------------------------------------------------------------

def tone_examples(store: Any, keywords: str, *, limit: int = 3) -> list[str]:
    """A few of the teacher's own parent-facing lines, as style examples.

    Pulled live from this machine's own database (Additional Comment column),
    never bundled: each teacher's app imitates that teacher. Names were already
    stripped when these were written; digit-bearing lines are skipped so the
    examples never invite the model to invent numbers.
    """
    candidates: list[str] = []
    try:
        for sheet_name in store.sheet_names():
            for row in store.load_sheet(sheet_name)["rows"]:
                text = str(row["values"].get("Additional Comment") or "").strip()
                if not text or _DIGITS.search(text):
                    continue
                if not 15 <= len(text) <= 120:
                    continue
                candidates.append(text)
    except Exception:
        candidates = []
    if not candidates:
        return _FALLBACK_EXAMPLES[:limit]

    unique = list(dict.fromkeys(candidates))
    key_chars = set(keywords)

    def overlap(example: str) -> int:
        return len(key_chars & set(example))

    unique.sort(key=overlap, reverse=True)
    return unique[:limit]


# Neutral, de-identified fallbacks for a fresh machine with no comments yet.
_FALLBACK_EXAMPLES = [
    "上课听得很专注，笔记记得完整，遇到不确定的地方会主动提问～",
    "做题速度不错，思路清楚，接下来在书写规范上再注意一点会更好～",
    "课堂参与很积极，愿意分享自己的解法，继续保持这个好习惯～",
]


# --- prompts -------------------------------------------------------------------

def _prompt_student(keywords: str, examples: list[str]) -> str:
    sample_lines = "\n".join(f"- {example}" for example in examples)
    return (
        "你是 Think Academy 的数学老师，正在给家长写孩子的课堂反馈。\n"
        "下面是你平时写反馈的语气示例：\n"
        f"{sample_lines}\n\n"
        "现在把这些关键词扩写成一段发给家长的话，2 到 3 句、60 到 90 个字：\n"
        f"关键词：{keywords}\n\n"
        "硬性要求：\n"
        "- 只能使用关键词里给出的事实，绝对不能编造成绩、名次或具体事件\n"
        "- 不要写学生姓名，不要任何称呼\n"
        "- 不要写“家长您好”，也不要写结尾问候（模板会自动加）\n"
        "- 语气亲切自然，多数句子以“～”结尾\n"
        "只输出扩写后的这段话。"
    )


def _prompt_recap(keywords: str) -> str:
    return (
        "你是 Think Academy 的数学老师，要写一句发给家长的本节课内容回顾。\n"
        f"知识点关键词：{keywords}\n\n"
        "要求：\n"
        "- 一句话，30 到 60 个字，句式类似“今天的课程主要围绕……展开”\n"
        "- 只能提到给出的知识点，不能添加其他内容\n"
        "- 不要写“家长您好”（模板会加），句尾用“～”\n"
        "只输出这句话。"
    )


def _prompt_announce(text: str) -> str:
    return (
        "润色下面这段发给家长的通知，让它更通顺、更有礼貌、分段更清楚。\n\n"
        "绝对不能更改、删除或添加任何事实、时间、日期、数字、地点或要求；"
        "只能调整语言表达和分段。\n\n"
        f"原文：\n{text}\n\n"
        "直接输出润色后的全文，不要任何解释。"
    )


# --- guards ----------------------------------------------------------------------

def fact_guard(kind: str, source: str, draft: str) -> str:
    """The factual firewall. Returns "" when the draft passes, else the reason.

    Expansion may never ADD digits the teacher didn't type; a polish of an
    announcement additionally may never LOSE one, nor introduce a temporal
    relation ("考试结束后...") the source never stated.
    """
    source_digits = _DIGITS.findall(source)
    draft_digits = _DIGITS.findall(draft)
    added = [d for d in set(draft_digits) if draft_digits.count(d) > source_digits.count(d)]
    if added:
        return f"草稿凭空出现了数字 {', '.join(sorted(added))}"
    if kind == "announce":
        missing = [d for d in set(source_digits) if source_digits.count(d) > draft_digits.count(d)]
        if missing:
            return f"草稿丢失了原文数字 {', '.join(sorted(missing))}"
        invented = [m for m in _INFERENTIAL_MARKERS if m in draft and m not in source]
        if invented:
            return f"草稿添加了原文没有的时间推断“{invented[0]}”"
    return ""


def _strip_wrapping(draft: str) -> str:
    text = draft.strip()
    # Models occasionally wrap output in quotes or a label despite instructions.
    text = re.sub(r"^(扩写[后的]*[:：]|润色[后的]*[:：])\s*", "", text)
    return text.strip(' "“”')


# --- main entry ------------------------------------------------------------------

def expand(kind: str, text: str, store: Any) -> dict[str, Any]:
    """keywords/draft -> parent-ready text, or a clear failure. Never raises."""
    source = str(text or "").strip()
    if not source:
        return {"ok": False, "error": "先在输入框里写几个关键词。"}
    status = ai_status()
    if not status["available"]:
        return {"ok": False, "error": status["hint"]}
    model = status["model"]

    if kind == "student":
        prompt = _prompt_student(source, tone_examples(store, source))
    elif kind == "recap":
        prompt = _prompt_recap(source)
    elif kind == "announce":
        prompt = _prompt_announce(source)
    else:
        return {"ok": False, "error": f"未知的扩写类型 {kind!r}。"}

    started = time.time()
    last_reason = ""
    temperatures = (0.1, 0.0) if kind == "announce" else (0.3, 0.1)
    for attempt, temperature in enumerate(temperatures):
        try:
            data = _http_json(
                "/api/generate",
                {
                    "model": model,
                    "prompt": prompt
                    + (
                        f"\n\n注意：上一稿因为“{last_reason}”被拒绝，请严格改正。"
                        if last_reason
                        else ""
                    ),
                    "stream": False,
                    "options": {"temperature": temperature},
                },
                timeout=120,
            )
        except urllib.error.URLError as exc:
            return {"ok": False, "error": f"本地模型调用失败：{exc.reason}"}
        except Exception as exc:
            return {"ok": False, "error": f"本地模型调用失败：{exc}"}

        draft = _strip_wrapping(str(data.get("response") or ""))
        if not draft:
            last_reason = "输出为空"
            continue
        reason = fact_guard(kind, source, draft)
        if not reason:
            return {
                "ok": True,
                "text": draft,
                "model": model,
                "elapsed": round(time.time() - started, 1),
                "attempts": attempt + 1,
            }
        last_reason = reason

    return {
        "ok": False,
        "error": f"AI 草稿两次都没通过安全校验（{last_reason}），请手动书写这段。",
    }
