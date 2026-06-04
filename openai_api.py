from __future__ import annotations

import base64
import json
import mimetypes
import os
from pathlib import Path
from typing import Any
from urllib import request
from urllib.error import HTTPError, URLError


DEFAULT_OPENAI_MODEL = "gpt-5.5"
OPENAI_RESPONSES_URL = "https://api.openai.com/v1/responses"


class OpenAIConfigError(RuntimeError):
    pass


def _api_key() -> str:
    key = os.environ.get("OPENAI_API_KEY", "").strip()
    if not key:
        raise OpenAIConfigError(
            "OPENAI_API_KEY is not set. Set it first, or run without --use-api."
        )
    return key


def _file_content(file_path: Path) -> dict[str, str]:
    mime_type = mimetypes.guess_type(file_path.name)[0] or "application/octet-stream"
    encoded = base64.b64encode(file_path.read_bytes()).decode("utf-8")
    return {
        "type": "input_file",
        "filename": file_path.name,
        "file_data": f"data:{mime_type};base64,{encoded}",
    }


def _extract_output_text(response: dict[str, Any]) -> str:
    if text := response.get("output_text"):
        return str(text).strip()

    chunks: list[str] = []
    for item in response.get("output", []):
        for content in item.get("content", []):
            if content.get("type") in {"output_text", "text"} and content.get("text"):
                chunks.append(str(content["text"]))
    return "\n".join(chunks).strip()


def create_response(
    prompt: str,
    *,
    model: str = DEFAULT_OPENAI_MODEL,
    file_path: Path | None = None,
) -> str:
    content: list[dict[str, str]] = [{"type": "input_text", "text": prompt}]
    if file_path:
        content.insert(0, _file_content(file_path))

    payload = {
        "model": model,
        "input": [
            {
                "role": "user",
                "content": content,
            }
        ],
    }
    data = json.dumps(payload).encode("utf-8")
    req = request.Request(
        OPENAI_RESPONSES_URL,
        data=data,
        headers={
            "Authorization": f"Bearer {_api_key()}",
            "Content-Type": "application/json",
        },
        method="POST",
    )

    try:
        with request.urlopen(req, timeout=120) as response:
            body = response.read().decode("utf-8")
    except HTTPError as exc:
        detail = exc.read().decode("utf-8", errors="replace")
        raise RuntimeError(f"OpenAI API request failed: {exc.code} {detail}") from exc
    except URLError as exc:
        raise RuntimeError(f"OpenAI API request failed: {exc.reason}") from exc

    output = _extract_output_text(json.loads(body))
    if not output:
        raise RuntimeError("OpenAI API response did not contain output text.")
    return output
