from __future__ import annotations

from openai_api import DEFAULT_OPENAI_MODEL
from feedback_common import StudentRow
from feedback_master import generate_feedback


def comment_payload_for_student(
    student: StudentRow,
    *,
    class_review: str,
    use_api: bool,
    model: str = DEFAULT_OPENAI_MODEL,
    feedback_type: str = "comprehensive",
    message_column: str = "Feedback",
) -> str:
    existing_feedback = str(student.values.get(message_column) or "").strip()
    if existing_feedback:
        return existing_feedback

    if message_column != "Feedback":
        raise ValueError(
            f"Row {student.excel_row} does not have a message in column {message_column!r}."
        )

    feedback = generate_feedback(
        student,
        class_review=class_review,
        use_api=use_api,
        model=model,
        feedback_type=feedback_type,
    )
    if not feedback:
        raise ValueError(f"Row {student.excel_row} has no feedback because the student is absent.")
    return feedback


def main() -> None:
    from paste_sender import main as sender_main

    sender_main(default_message_column="Feedback")


if __name__ == "__main__":
    main()
