from __future__ import annotations

from dataclasses import dataclass

from openai_api import DEFAULT_OPENAI_MODEL
from feedback_common import (
    FEEDBACK_TYPE_CHOICES,
    OBSERVATION_FIELDS,
    StudentRow,
    append_parent_closing,
    class_review_paragraph,
    clean_parent_feedback_text,
    homework_paragraph,
    personal_feedback_with_gpt,
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
    use_api: bool = False
    model: str = DEFAULT_OPENAI_MODEL

    def is_chinese(self, student: StudentRow) -> bool:
        return student.language.lower().startswith("chinese")

    def observations_for_student(self, student: StudentRow) -> list[str]:
        return [
            phrase
            for field in OBSERVATION_FIELDS
            if (phrase := phrase_for(field, student.values.get(field), student.language))
        ]

    def class_paragraph(self, student: StudentRow) -> str:
        return class_review_paragraph(self.class_review, self.is_chinese(student))

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
        class_paragraph = self.class_paragraph(student)
        homework = homework_paragraph(student, is_chinese)
        personal_paragraphs = self.personal_paragraphs(student, feedback_type)

        if self.use_api:
            local_comment = "\n\n".join(personal_paragraphs)
            personal_section = personal_feedback_with_gpt(
                student,
                observations=self.observations_for_student(student),
                local_comment=local_comment,
                homework=homework,
                model=self.model,
                feedback_type=feedback_type,
            )
            feedback = clean_parent_feedback_text("\n\n".join([class_paragraph, personal_section.strip()]))
            return append_parent_closing(feedback, is_chinese)

        paragraphs = [class_paragraph, *personal_paragraphs]
        if homework:
            paragraphs.append(homework)
        feedback = clean_parent_feedback_text("\n\n".join(paragraphs))
        return append_parent_closing(feedback, is_chinese)

def generate_feedback(
    student: StudentRow,
    class_review: str = "",
    *,
    use_api: bool = False,
    model: str = DEFAULT_OPENAI_MODEL,
    feedback_type: str = "comprehensive",
) -> str | None:
    return FeedbackGenerator(
        class_review=class_review,
        use_api=use_api,
        model=model,
    ).generate(student, feedback_type=feedback_type)
