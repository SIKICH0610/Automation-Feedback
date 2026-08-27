from __future__ import annotations

import argparse
from pathlib import Path
from xml.etree import ElementTree
from zipfile import ZipFile



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
        f"{suffix or 'this file type'} cannot be read directly. Convert it to .txt or .pptx, "
        "or paste the recap into the Lesson recap box in the app."
    )


def build_class_review(
    *,
    source_text: str = "",
    source_file: Path | None = None,
) -> str:
    if source_file:
        source_text = read_source_text(source_file)

    source_text = source_text.strip()
    if not source_text:
        raise ValueError("Provide --source-text or --source-file.")
    return source_text


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="Create the class review paragraph used as paragraph 1."
    )
    parser.add_argument("--source-text", default="", help="Teacher-written class review.")
    parser.add_argument("--source-file", type=Path, help="Text, pptx, pdf, or slide file.")
    parser.add_argument("--output", type=Path, default=DEFAULT_OUTPUT)
    return parser


def main() -> None:
    args = build_parser().parse_args()
    review = build_class_review(
        source_text=args.source_text,
        source_file=args.source_file,
    )
    args.output.write_text(review.strip() + "\n", encoding="utf-8")
    print(f"Wrote class review to {args.output}")
    print()
    print(review)


if __name__ == "__main__":
    main()
