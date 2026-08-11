from __future__ import annotations

import asyncio
import io
from typing import Any


def capture_window_text(window: Any) -> str:
    """Screenshot the given pywinauto window and OCR it, returning all recognized text.

    Some desktop apps (WeCom, for one) render their entire UI as custom-drawn graphics
    rather than real controls, so UI Automation exposes zero readable text no matter
    which backend is used. OCR reads the same pixels a person would see instead, which
    works regardless of how the app renders itself.
    """
    from PIL import ImageGrab

    rectangle = window.rectangle()
    image = ImageGrab.grab(bbox=(rectangle.left, rectangle.top, rectangle.right, rectangle.bottom))
    return asyncio.run(_ocr_image_async(image))


async def _ocr_image_async(image: Any) -> str:
    from winsdk.windows.graphics.imaging import BitmapDecoder
    from winsdk.windows.media.ocr import OcrEngine
    from winsdk.windows.storage.streams import DataWriter, InMemoryRandomAccessStream

    buffer = io.BytesIO()
    image.save(buffer, format="PNG")

    stream = InMemoryRandomAccessStream()
    writer = DataWriter(stream)
    writer.write_bytes(bytearray(buffer.getvalue()))
    await writer.store_async()
    await writer.flush_async()
    stream.seek(0)

    decoder = await BitmapDecoder.create_async(stream)
    bitmap = await decoder.get_software_bitmap_async()
    engine = OcrEngine.try_create_from_user_profile_languages()
    if engine is None:
        raise RuntimeError(
            "No OCR engine is available for this Windows user profile's languages. "
            "Add an OCR language pack in Windows Settings > Time & Language > Language."
        )
    result = await engine.recognize_async(bitmap)
    return result.text
