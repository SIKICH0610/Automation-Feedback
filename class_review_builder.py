from __future__ import annotations

import argparse
from pathlib import Path
from xml.etree import ElementTree
from zipfile import ZipFile

from openai_api import DEFAULT_OPENAI_MODEL, create_response


DEFAULT_OUTPUT = Path("class_review.txt")


def extract_pptx_text(file_path: Path) -> str:
    parts: list[str] = []
    namespace = "{http://schemas.openxmlformats.org/drawingml/2006/main}"

    with ZipFile(file_path) as archive:
        slide_names = sorted(
            name
            for name in archive.namelist()
            if name.startswith("ppt/slides/slide") and name.endswith(".xml")
        )
        for index, slide_name in enumerate(slide_names, start=1):
            root = ElementTree.fromstring(archive.read(slide_name))
            texts = [
                node.text.strip()
                for node in root.iter(f"{namespace}t")
                if node.text and node.text.strip()
            ]
            if texts:
                parts.append(f"Slide {index}: " + " ".join(texts))

    return "\n".join(parts).strip()


def read_source_text(file_path: Path) -> str:
    suffix = file_path.suffix.lower()
    if suffix in {".txt", ".md"}:
        return file_path.read_text(encoding="utf-8").strip()
    if suffix == ".pptx":
        return extract_pptx_text(file_path)
    raise ValueError(
        f"{suffix or 'this file type'} needs --use-api so GPT can read the file directly."
    )


def summarize_prompt(source_text: str, language: str) -> str:
    return f"""
You are helping a teacher write the class-material paragraph for a parent update.

Write one natural paragraph in {language}.
Mention the main topics, skills, and practice from today's class.
Do not add a greeting, student name, bullet list, homework feedback, or personal comments.
Output only the paragraph.

Class material text:
{source_text}
""".strip()


def summarize_file_prompt(language: str) -> str:
    return f"""
You are helping a teacher write the class-material paragraph for a parent update.

Read the attached class material or slides.
Write one natural paragraph in {language}.
Mention the main topics, skills, and practice from today's class.
Do not add a greeting, student name, bullet list, homework feedback, or personal comments.
Output only the paragraph.
""".strip()


def build_class_review(
    *,
    source_text: str = "",
    source_file: Path | None = None,
    language: str,
    use_api: bool,
    model: str,
) -> str:
    if source_file and use_api:
        return create_response(
            summarize_file_prompt(language),
            model=model,
            file_path=source_file,
        )

    if source_file:
        source_text = read_source_text(source_file)

    source_text = source_text.strip()
    if not source_text:
        raise ValueError("Provide --source-text or --source-file.")

    if use_api:
        return create_response(summarize_prompt(source_text, language), model=model)
    return source_text


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="Create the class review paragraph used as paragraph 1."
    )
    parser.add_argument("--source-text", default="", help="Teacher-written class review.")
    parser.add_argument("--source-file", type=Path, help="Text, pptx, pdf, or slide file.")
    parser.add_argument("--output", type=Path, default=DEFAULT_OUTPUT)
    parser.add_argument("--language", default="English")
    parser.add_argument("--use-api", action="store_true")
    parser.add_argument("--model", default=DEFAULT_OPENAI_MODEL)
    return parser


def main() -> None:
    args = build_parser().parse_args()
    review = build_class_review(
        source_text=args.source_text,
        source_file=args.source_file,
        language=args.language,
        use_api=args.use_api,
        model=args.model,
    )
    args.output.write_text(review.strip() + "\n", encoding="utf-8")
    print(f"Wrote class review to {args.output}")
    print()
    print(review)


if __name__ == "__main__":
    main()
