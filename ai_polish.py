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

# Every message goes point-to-point to ONE family, so group address forms are
# banned outright in drafts -- prompting alone does not hold (learned repeatedly).
_PLURAL_MARKERS = ("家长们", "各位家长", "各位同学", "各位家长朋友")


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


_STATUS_CACHE: dict[str, Any] = {"at": 0.0, "value": None}


def ai_status_cached(ttl: float = 60.0) -> dict[str, Any]:
    now = time.time()
    if _STATUS_CACHE["value"] is None or now - _STATUS_CACHE["at"] > ttl:
        _STATUS_CACHE["value"] = ai_status()
        _STATUS_CACHE["at"] = now
    return _STATUS_CACHE["value"]


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

def _examples_from_rows(rows_texts: list[str], keywords: str, limit: int) -> list[str]:
    candidates = []
    for text in rows_texts:
        text = str(text or "").strip()
        if not text or _DIGITS.search(text):
            continue
        if not 15 <= len(text) <= 120:
            continue
        candidates.append(text)
    if not candidates:
        return _FALLBACK_EXAMPLES[:limit]
    unique = list(dict.fromkeys(candidates))
    key_chars = set(keywords)
    unique.sort(key=lambda example: len(key_chars & set(example)), reverse=True)
    return unique[:limit]


def _runtime_tone_texts() -> list[str]:
    """Additional Comment texts straight from the database named by
    FEEDBACK_DATABASE_PATH -- the worker subprocess has no store object."""
    import os
    import sqlite3

    path = os.environ.get("FEEDBACK_DATABASE_PATH", "").strip()
    if not path:
        return []
    try:
        connection = sqlite3.connect(path)
        try:
            rows = connection.execute("SELECT values_json FROM students").fetchall()
        finally:
            connection.close()
    except sqlite3.Error:
        return []
    texts = []
    for (raw,) in rows:
        try:
            texts.append(str(json.loads(raw).get("Additional Comment") or ""))
        except (json.JSONDecodeError, AttributeError):
            continue
    return texts


def tone_examples(store: Any, keywords: str, *, limit: int = 4) -> list[str]:
    """A few of the teacher's own parent-facing lines, as style examples.

    Pulled live from this machine's own database (Additional Comment column),
    never bundled: each teacher's app imitates that teacher. Names were already
    stripped when these were written; digit-bearing lines are skipped so the
    examples never invite the model to invent numbers.
    """
    texts: list[str] = []
    if store is not None:
        try:
            for sheet_name in store.sheet_names():
                for row in store.load_sheet(sheet_name)["rows"]:
                    texts.append(str(row["values"].get("Additional Comment") or ""))
        except Exception:
            texts = []
    else:
        texts = _runtime_tone_texts()
    return _examples_from_rows(texts, keywords, limit)


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
        "现在把这些关键词扩写成一段发给家长的话，3 到 4 句、90 到 130 个字：\n"
        f"关键词：{keywords}\n\n"
        "硬性要求：\n"
        "- 只能使用关键词里给出的事实，绝对不能编造成绩、名次或具体事件\n"
        "- 不要写学生姓名，不要任何称呼\n"
        "- 不要写“家长您好”，也不要写结尾问候（模板会自动加）\n"
        "- 语气亲切自然、有温度，像老师平时跟家长聊天；"
        "可以在事实基础上加一点具体的肯定和下一步的小建议，但不得引入新事实\n"
        "- 中间句子用句号，整段只在最后一句用“～”收尾\n"
        "- 不用冒号，不用破折号；用最普通、最常见的说法，别用书面腔和生僻表达\n"
        "- 这条消息只发给一位家长、只谈这一个孩子：绝不能出现“孩子们”“同学们”“各位”“大家”等群体称呼\n"
        "- 提到孩子时直接用“孩子”做主语，口语自然——不说“您孩子的表现非常认真”，要说“孩子上课很认真”；不用“该生”“表现出色”这类书面腔\n"
        "- 不写格言腔和抒情腔：不用“不是A，是B”式的总结句（如“安静不是没听进去，是性格”），"
        "不写老师的个人抒情（如“这股劲儿我想帮她保住”）；夸奖直接落在孩子的具体行为上\n"
        "- 通篇不用“他”“她”——性别不要猜：能省略代词就省略，需要主语或代词时一律写“孩子”\n"
        "只输出扩写后的这段话。"
    )


def _looks_like_keywords(text: str) -> bool:
    """Keywords vs. already-written prose decides expand vs. polish.

    A finished recap fed to the expander used to get lossily squeezed into the
    one-line template -- "第三节课将开始学全等" even came out as already taught.
    Prose gets a conservative polish instead.
    """
    stripped = text.strip()
    if "家长" in stripped:
        return False
    if any(mark in stripped for mark in "。！？～!?"):
        return False
    return len(stripped) <= 32


def _prompt_student_rewrite(material: str, examples: list[str]) -> str:
    sample_lines = "\n".join(f"- {example}" for example in examples)
    return (
        "你是 Think Academy 的数学老师。下面是你课上记的关于一个学生的速记，"
        "现在要把它改写成一段发给家长的话。\n"
        "你平时写给家长的语气示例：\n"
        f"{sample_lines}\n\n"
        f"课堂速记：{material}\n\n"
        "硬性要求：\n"
        "- 必须完全用自己的话重新表达，禁止照抄速记里的任何整句；"
        "速记里的错别字和内部说法（如“需要他……”）要转成家长易读的表达"
        "（如“接下来可以在……上多加练习～”）\n"
        "- 速记里的每个事实和建议都要保留，但绝不能添加速记里没有的事实\n"
        "- 3 到 4 句、90 到 130 个字\n"
        "- 不要写学生姓名，不要称呼，不要问候和结尾（模板会加）\n"
        "- 语气亲切自然；中间句子用句号，整段只在最后一句用“～”收尾\n"
        "- 不用冒号，不用破折号；用最普通、最常见的说法\n"
        "- 这条消息只发给一位家长、只谈这一个孩子：绝不能出现“孩子们”“同学们”“各位”“大家”等群体称呼\n"
        "- 提到孩子时直接用“孩子”做主语，口语自然——不说“您孩子的表现非常认真”，要说“孩子上课很认真”；不用“该生”“表现出色”这类书面腔\n"
        "- 不写格言腔和抒情腔：不用“不是A，是B”式的总结句（如“安静不是没听进去，是性格”），"
        "不写老师的个人抒情（如“这股劲儿我想帮她保住”）；夸奖直接落在孩子的具体行为上\n"
        "- 通篇不用“他”“她”——性别不要猜：能省略代词就省略，需要主语或代词时一律写“孩子”\n"
        "只输出改写后的这段话。"
    )


def _prompt_recap(keywords: str) -> str:
    return (
        "你是 Think Academy 的数学老师，要写一句发给家长的本节课内容回顾。\n"
        f"知识点关键词：{keywords}\n\n"
        "要求：\n"
        "- 一句话，30 到 60 个字，句式类似“今天的课程主要围绕……展开”\n"
        "- 只能提到给出的知识点，不能添加其他内容\n"
        "- 不要写“家长您好”（模板会加），句尾用“～”\n"
        "- 这条消息只发给一位家长、只谈这一个孩子：绝不能出现“孩子们”“同学们”“各位”“大家”等群体称呼\n"
        "只输出这句话。"
    )


def _prompt_recap_polish(text: str) -> str:
    return (
        "逐句润色下面这段老师已经写好的课程回顾。\n\n"
        "规则：\n"
        "- 保留全部信息点，一条都不能丢，也不能压缩合并\n"
        "- 时态绝对不能变：原文说“将开始”“下节课”“即将”的内容是还没学的，"
        "绝不能写成已经学过；已经学过的也不能写成将要学\n"
        "- 不得添加原文没有的内容或解释\n"
        "- 只调整用词和语气，让句子更通顺；不用冒号和破折号，“～”只放在结尾\n"
        "- 这条消息只发给一位家长、只谈这一个孩子：绝不能出现“孩子们”“同学们”“各位”“大家”等群体称呼\n""\n"
        f"原文：\n{text}\n\n"
        "直接输出润色后的全文，不要任何解释。"
    )


def _prompt_announce(text: str) -> str:
    return (
        "润色下面这段发给家长的通知，让它更通顺、更有礼貌、分段更清楚。\n\n"
        "绝对不能更改、删除或添加任何事实、时间、日期、数字、地点或要求；"
        "只能调整语言表达和分段。"
        "不用冒号和破折号（比如“时间：3点”要写成“时间是3点”）；“～”最多出现在段落结尾。\n\n"
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
    plural = [m for m in _PLURAL_MARKERS if m in draft]
    if plural:
        return f"草稿使用了群体称呼“{plural[0]}”（消息都是一对一发送的）"
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


# Stock officialese the model keeps producing no matter the prompt. These sentences
# carry no facts, so deleting them outright is safe.
_STOCK_COURTESY = re.compile(
    r"[^。！？!?～~\n]*(?:感谢您的配合|感谢配合|敬请谅解|特此通知|望周知)[^。！？!?～~\n]*[。！？!?～~]?"
)

# Point-to-point messages: plural nouns downgrade deterministically instead of
# burning a retry (the model writes 同学们 whenever the material mentions 同学).
_PLURAL_DOWNGRADES = (
    ("孩子们", "孩子"),
    ("同学们", "同学"),
    ("学生们", "学生"),
    # Stiff third-person references: the natural subject is just 孩子.
    ("您的孩子", "孩子"),
    ("您孩子", "孩子"),
    ("贵子女", "孩子"),
    ("该生", "孩子"),
)


# Parents must never see a guessed gender: the material rarely says 他/她, so
# the model invents one. Per the teacher (2026-09-15), gendered pronouns
# collapse to the role noun -- Chinese 孩子, mirroring English he/she -> the
# student. Guards keep 其他/他人/吉他 etc. intact; 她 has no such compounds
# once 她们 is handled.
def degender_zh(text: str) -> str:
    result = re.sub(r"[他她]们", "孩子", text)
    result = re.sub(r"(?<![其吉维排利])他(?![人日乡])", "孩子", result)
    result = result.replace("她", "孩子")
    while "孩子孩子" in result:
        result = result.replace("孩子孩子", "孩子")
    return result


# House punctuation for parent-facing drafts (the teacher's standing writing
# law): no colons and no dashes as punctuation, plainest common phrasing.
# Mechanical demotion to a comma is the deterministic backstop; the prompts do
# the real restructuring. Digit ranges (3—4点) survive.
def plain_punctuation(text: str) -> str:
    result = text.replace("：", "，")
    result = re.sub(r"(?<![0-9])[—–]+(?![0-9])", "，", result)
    return result.replace("，。", "。").replace("，，", "，")


# “～” belongs at the end of a paragraph, once; interior waves read as spam to
# the teacher. Digit ranges (下午3～4点) stay.
def limit_wave(text: str) -> str:
    lines = []
    for line in text.split("\n"):
        stripped = line.rstrip()
        core, tail = (stripped[:-1], "～") if stripped.endswith("～") else (stripped, "")
        core = re.sub(r"(?<![0-9０-９])～", "。", core)
        lines.append(core.replace("。。", "。") + tail)
    return "\n".join(lines)


def attach_name(name: str, text: str) -> str:
    """Join a student's name onto a draft that opens with a role noun.

    Drafts say 孩子/学生 by design and the template supplies the real name, so
    "Sunnie" + "孩子上课很认真" must read "Sunnie上课很认真", never "Sunnie 孩子…".
    """
    body = text.strip()
    for prefix in ("孩子", "学生", "该生"):
        if body.startswith(prefix):
            body = body[len(prefix):].lstrip("，, ")
            break
    if not body:
        return name
    joiner = "" if "一" <= body[0] <= "鿿" else " "
    return f"{name}{joiner}{body}"


def _sanitize_draft(draft: str) -> str:
    text = _STOCK_COURTESY.sub("", draft)
    for plural, singular in _PLURAL_DOWNGRADES:
        text = text.replace(plural, singular)
    text = degender_zh(text)
    # 模型偶尔把句号和波浪号叠在一起（“。～”）
    text = text.replace("。～", "～").replace("！～", "～")
    text = plain_punctuation(text)
    text = limit_wave(text)
    text = re.sub(r"\n{3,}", "\n\n", text)
    return text.strip()


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

    guard_kind = kind
    if kind == "student":
        examples = tone_examples(store, source)
        # Terse keywords get expanded; sentence-like shorthand gets a mandatory
        # rewrite -- fed prose, the model's instinct is to echo it back verbatim,
        # internal phrasing, typos and all.
        if _looks_like_keywords(source) or len(source) <= 45:
            prompt = _prompt_student(source, examples)
        else:
            prompt = _prompt_student_rewrite(source, examples)
    elif kind == "recap":
        if _looks_like_keywords(source):
            prompt = _prompt_recap(source)
        else:
            # Finished prose: polish under the strict announcement-grade guard
            # (no digit lost, no invented temporal gloss) instead of compressing.
            prompt = _prompt_recap_polish(source)
            guard_kind = "announce"
    elif kind == "announce":
        prompt = _prompt_announce(source)
    else:
        return {"ok": False, "error": f"未知的扩写类型 {kind!r}。"}

    started = time.time()
    last_reason = ""
    temperatures = (0.1, 0.0) if guard_kind == "announce" else (0.3, 0.1)
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

        draft = _sanitize_draft(_strip_wrapping(str(data.get("response") or "")))
        if not draft:
            last_reason = "输出为空"
            continue
        reason = fact_guard(guard_kind, source, draft)
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


def polish_comment(text: str, name: str, store: Any) -> dict[str, Any]:
    """Re-polish a generated comment in place.

    Only the personal middle paragraphs are rewritten. The templated
    greeting/recap head and the homework/closing tail carry standing
    instructions, so they must survive untouched.
    """
    comment = (text or "").strip()
    if not comment:
        return {"ok": False, "error": "这格还没有生成的评语。"}
    if not any("一" <= ch <= "鿿" for ch in comment):
        return {"ok": False, "error": "AI 润色目前只支持中文评语。"}
    paragraphs = [p for p in comment.split("\n\n")]
    if len(paragraphs) < 3:
        return {"ok": False, "error": "评语不是“开头、正文、结尾”的三段结构，直接重新生成更稳妥。"}
    clean_name = (name or "").strip()
    total_elapsed = 0.0
    model = ""
    middles = []
    for paragraph in paragraphs[1:-1]:
        body = paragraph.strip()
        had_name = bool(clean_name) and body.startswith(clean_name)
        if had_name:
            body = body[len(clean_name):].lstrip(" ，,")
        result = expand("student", body, store)
        if not result.get("ok"):
            return result
        total_elapsed += float(result.get("elapsed") or 0)
        model = result.get("model", model)
        middles.append(attach_name(clean_name, result["text"]) if had_name else result["text"])
    rebuilt = "\n\n".join([paragraphs[0], *middles, paragraphs[-1]])
    return {"ok": True, "text": rebuilt, "model": model, "elapsed": round(total_elapsed, 1)}
