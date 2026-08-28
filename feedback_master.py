from __future__ import annotations

from dataclasses import dataclass

from feedback_common import (
    FEEDBACK_TYPE_CHOICES,
    OBSERVATION_FIELDS,
    StudentRow,
    closing_sentence,
    homework_note,
    class_review_paragraph,
    clean_parent_feedback_text,
    homework_paragraph,
    phrase_for,
)
from feedback_general import general_comment_paragraph
from feedback_quiz import quiz_comment_paragraph

def comprehensive_comment_paragraph(student: StudentRow, observations: list[str], is_chinese: bool) -> str:
    quiz_comment = quiz_comment_paragraph(
        student,
        observations,
        is_chinese,
        include_observations=False,
    )
    general_comment = general_comment_paragraph(
        student,
        observations,
        is_chinese,
        include_quiz_remark=False,
    )
    if quiz_comment and general_comment:
        return "\n\n".join([quiz_comment, general_comment])
    return quiz_comment or general_comment

@dataclass
class FeedbackGenerator:
    class_review: str = ""
    # "1" / "2" when the teacher picked a quiz in the UI; None lets the row's own
    # data decide, which is all the CLI and comprehensive feedback can do.
    quiz_number: str | None = None

    def is_chinese(self, student: StudentRow) -> bool:
        return student.language.lower().startswith("chinese")

    def observations_for_student(self, student: StudentRow) -> list[str]:
        return [
            phrase
            for field in OBSERVATION_FIELDS
            if (phrase := phrase_for(field, student.values.get(field), student.language))
        ]

    def class_paragraph(self, student: StudentRow, feedback_type: str = "comprehensive") -> str:
        kind = "quiz" if feedback_type == "quiz" else "lesson"
        return class_review_paragraph(self.class_review, self.is_chinese(student), kind=kind)

    def general_personal_paragraph(self, student: StudentRow) -> str:
        return general_comment_paragraph(
            student,
            self.observations_for_student(student),
            self.is_chinese(student),
        )

    def quiz_personal_paragraph(self, student: StudentRow) -> str:
        quiz_comment = quiz_comment_paragraph(
            student,
            self.observations_for_student(student),
            self.is_chinese(student),
            include_observations=False,
            quiz_number=self.quiz_number,
        )
        if quiz_comment:
            return quiz_comment

        name = student.first_name or student.full_name
        if self.is_chinese(student):
            return f"{name}本次 quiz 情况已记录，之后可以继续复习相关知识点，并注意证明书写的规范性。"
        return (
            f"{name}'s quiz record has been noted. Going forward, reviewing the related topics "
            "and keeping proof writing clear will be helpful."
        )

    def comprehensive_personal_paragraphs(self, student: StudentRow) -> list[str]:
        paragraphs = [
            self.quiz_personal_paragraph(student),
            general_comment_paragraph(
                student,
                self.observations_for_student(student),
                self.is_chinese(student),
                include_quiz_remark=False,
            ),
        ]
        return [paragraph for paragraph in paragraphs if paragraph.strip()]

    def personal_paragraphs(self, student: StudentRow, feedback_type: str) -> list[str]:
        if feedback_type == "general":
            return [self.general_personal_paragraph(student)]
        if feedback_type == "quiz":
            return [self.quiz_personal_paragraph(student)]
        if feedback_type == "comprehensive":
            return self.comprehensive_personal_paragraphs(student)
        raise ValueError(f"Unknown feedback type: {feedback_type!r}")

    def generate(self, student: StudentRow, feedback_type: str = "comprehensive") -> str | None:
        if student.is_absent:
            return None
        if feedback_type not in FEEDBACK_TYPE_CHOICES:
            raise ValueError(
                f"feedback_type must be one of {', '.join(FEEDBACK_TYPE_CHOICES)}; got {feedback_type!r}"
            )

        is_chinese = self.is_chinese(student)
        class_paragraph = self.class_paragraph(student, feedback_type)
        homework = homework_paragraph(student, is_chinese)
        personal_paragraphs = self.personal_paragraphs(student, feedback_type)

        # The standing homework note goes with the class-performance comment only. A
        # quiz-only message is about that quiz, and a parent who also gets the general
        # comment would otherwise read the same instructions twice.
        wants_homework_note = feedback_type in {"general", "comprehensive"}

        # Always exactly three paragraphs: greeting / the student's classroom
        # performance, nothing else / homework note + homework reflection + closing.
        joiner = "" if is_chinese else " "
        middle = joiner.join(part.strip() for part in personal_paragraphs if part and part.strip())
        tail_parts = [
            homework_note(is_chinese) if wants_homework_note else "",
            homework or "",
            closing_sentence(is_chinese),
        ]
        tail = joiner.join(part.strip() for part in tail_parts if part and part.strip())
        paragraphs = [class_paragraph, middle, tail] if middle else [class_paragraph, tail]
        return clean_parent_feedback_text("\n\n".join(paragraphs))

def generate_feedback(
    student: StudentRow,
    class_review: str = "",
    *,
    feedback_type: str = "comprehensive",
    quiz_number: str | None = None,
) -> str | None:
    return FeedbackGenerator(
        class_review=class_review,
        quiz_number=quiz_number,
    ).generate(student, feedback_type=feedback_type)
