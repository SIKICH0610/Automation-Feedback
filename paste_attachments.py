from __future__ import annotations

import ctypes
from ctypes import wintypes
import os
from pathlib import Path
import struct
import time
from typing import Any, Iterable


CF_HDROP = 15
GMEM_MOVEABLE = 0x0002
DROPEFFECT_COPY = 1


def normalized_attachment_paths(paths: Iterable[Path | str]) -> list[Path]:
    normalized: list[Path] = []
    seen: set[str] = set()
    for value in paths:
        path = Path(value).expanduser().resolve()
        if not path.is_file():
            raise FileNotFoundError(f"Attachment was not found: {path}")
        key = os.path.normcase(str(path))
        if key not in seen:
            seen.add(key)
            normalized.append(path)
    return normalized


def build_dropfiles_payload(paths: Iterable[Path | str]) -> bytes:
    normalized = normalized_attachment_paths(paths)
    if not normalized:
        raise ValueError("Select at least one attachment.")
    file_list = "\0".join(str(path) for path in normalized) + "\0\0"
    header = struct.pack("<IiiII", 20, 0, 0, 0, 1)
    return header + file_list.encode("utf-16le")


def _global_memory(kernel32: Any, payload: bytes) -> int:
    handle = kernel32.GlobalAlloc(GMEM_MOVEABLE, len(payload))
    if not handle:
        raise OSError("Windows could not allocate clipboard memory.")
    pointer = kernel32.GlobalLock(handle)
    if not pointer:
        kernel32.GlobalFree(handle)
        raise OSError("Windows could not lock clipboard memory.")
    try:
        ctypes.memmove(pointer, payload, len(payload))
    finally:
        kernel32.GlobalUnlock(handle)
    return int(handle)


def set_clipboard_files(paths: Iterable[Path | str]) -> list[Path]:
    if os.name != "nt":
        raise RuntimeError("File attachment paste currently requires Windows.")

    normalized = normalized_attachment_paths(paths)
    payload = build_dropfiles_payload(normalized)
    user32 = ctypes.windll.user32
    kernel32 = ctypes.windll.kernel32

    kernel32.GlobalAlloc.argtypes = [wintypes.UINT, ctypes.c_size_t]
    kernel32.GlobalAlloc.restype = ctypes.c_void_p
    kernel32.GlobalLock.argtypes = [ctypes.c_void_p]
    kernel32.GlobalLock.restype = ctypes.c_void_p
    kernel32.GlobalUnlock.argtypes = [ctypes.c_void_p]
    kernel32.GlobalFree.argtypes = [ctypes.c_void_p]
    user32.OpenClipboard.argtypes = [wintypes.HWND]
    user32.OpenClipboard.restype = wintypes.BOOL
    user32.SetClipboardData.argtypes = [wintypes.UINT, ctypes.c_void_p]
    user32.SetClipboardData.restype = ctypes.c_void_p
    user32.RegisterClipboardFormatW.argtypes = [wintypes.LPCWSTR]
    user32.RegisterClipboardFormatW.restype = wintypes.UINT

    opened = False
    for _ in range(12):
        if user32.OpenClipboard(None):
            opened = True
            break
        time.sleep(0.05)
    if not opened:
        raise OSError("Windows clipboard is busy. Try the action again.")

    drop_handle = 0
    effect_handle = 0
    try:
        if not user32.EmptyClipboard():
            raise OSError("Windows could not clear the clipboard.")

        drop_handle = _global_memory(kernel32, payload)
        if not user32.SetClipboardData(CF_HDROP, drop_handle):
            raise OSError("Windows could not place the files on the clipboard.")
        drop_handle = 0

        preferred_effect = user32.RegisterClipboardFormatW("Preferred DropEffect")
        if preferred_effect:
            effect_handle = _global_memory(
                kernel32,
                struct.pack("<I", DROPEFFECT_COPY),
            )
            if user32.SetClipboardData(preferred_effect, effect_handle):
                effect_handle = 0
    finally:
        if drop_handle:
            kernel32.GlobalFree(drop_handle)
        if effect_handle:
            kernel32.GlobalFree(effect_handle)
        user32.CloseClipboard()

    return normalized


def stage_attachments(
    paths: Iterable[Path | str],
    *,
    send_keys: Any,
    settle_seconds: float = 2.0,
) -> list[Path]:
    normalized = set_clipboard_files(paths)
    print(
        "Clipboard loaded with attachment files: "
        + ", ".join(path.name for path in normalized)
    )
    time.sleep(0.2)
    send_keys("^v")
    time.sleep(settle_seconds)
    print("Attachments staged in the chat preview. The script did not send them.")
    return normalized
