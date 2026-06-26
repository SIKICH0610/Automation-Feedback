<<<<<<< Updated upstream
from paste_sender import main


if __name__ == "__main__":
    main(default_message_column="Feedback")
=======
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
) -> str:
    existing_feedback = str(student.values.get("Feedback") or "").strip()
    if existing_feedback:
        return existing_feedback

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
>>>>>>> Stashed changes
