from __future__ import annotations

from pathlib import Path

from feedback_common import StudentRow


def resolve_mass_message(
    *,
    mass_message: str = "",
    mass_message_file: Path | None = None,
    fallback_text: str = "",
) -> str:
    if mass_message_file:
        return mass_message_file.read_text(encoding="utf-8").strip()
    if mass_message.strip():
        return mass_message.strip()
    return fallback_text.strip()


def mass_notification_payload_for_student(student: StudentRow, *, mass_message: str) -> str:
    if not mass_message.strip():
        raise ValueError(
            "Mass notification action needs --mass-message, --mass-message-file, "
            "or --class-review-file text."
        )
    return mass_message.strip()


def main() -> None:
    from paste_sender import main as sender_main

    sender_main(default_message_column="Pre-quiz informing", default_fallback_channel=True)


if __name__ == "__main__":
    main()
