from __future__ import annotations

import sys

from feedback_generator import main


# Geo MTWThF uses parent-facing Quiz1 Feedback for the first two physical quizzes.
DEFAULT_ARGS = [
    "--sheet",
    "Geo MTWThF",
    "--all",
    "--write",
    "--feedback-column",
    "Quiz1 Feedback",
    "--class-review-file-zh",
    "class_review_mtwthf_zh.txt",
    "--class-review-file-en",
    "class_review_mtwthf_en.txt",
]


if __name__ == "__main__":
    if len(sys.argv) == 1:
        sys.argv.extend(DEFAULT_ARGS)
    main(default_feedback_type="quiz")
