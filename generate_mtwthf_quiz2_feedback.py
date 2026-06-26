from __future__ import annotations

import sys

from feedback_generator import main


# Inputs:
#   Quiz2 Score = score record for the corresponding quiz.
#   Quiz2 Mistake = missed questions for the corresponding quiz.
# Output:
#   Quiz2 Feedback = final parent-facing quiz message.
DEFAULT_ARGS = [
    "--sheet",
    "Geo MTWThF",
    "--all",
    "--write",
    "--feedback-column",
    "Quiz2 Feedback",
    "--calculate-quiz-average",
    "--quiz-score-column",
    "Quiz2 Score",
    "--quiz-average-column",
    "Quiz2 Average",
    "--only-with-value-column",
    "Quiz2 Score",
    "--clear-feedback-for-missing-value",
    "--class-review-file-zh",
    "class_review_mtwthf_zh.txt",
    "--class-review-file-en",
    "class_review_mtwthf_en.txt",
]


if __name__ == "__main__":
    if len(sys.argv) == 1:
        sys.argv.extend(DEFAULT_ARGS)
    main(default_feedback_type="quiz")
