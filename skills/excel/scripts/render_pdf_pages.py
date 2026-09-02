#!/usr/bin/env python3
"""Render every page of a PDF to ordered PNG files with PyMuPDF."""

from __future__ import annotations

import argparse
import hashlib
import os
import sys
import tempfile
from pathlib import Path
from typing import Sequence

try:
    import pymupdf
except ImportError as exc:  # pragma: no cover - depends on caller environment
    raise SystemExit(
        "ERROR: PyMuPDF is required. Install it with: python -m pip install pymupdf"
    ) from exc


DEFAULT_DPI = 200


def configure_utf8_output() -> None:
    for stream in (sys.stdout, sys.stderr):
        reconfigure = getattr(stream, "reconfigure", None)
        if reconfigure is not None:
            reconfigure(encoding="utf-8", errors="replace")


def sha256_file(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as file_handle:
        for chunk in iter(lambda: file_handle.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def render_pages(
    source: Path,
    output_directory: Path,
    dpi: int,
    replace_existing: bool,
) -> list[Path]:
    """Render all PDF pages into a temporary directory, then promote them."""
    source_hash = sha256_file(source)
    document = None

    try:
        document = pymupdf.open(source)
        if not document.is_pdf:
            raise ValueError("The source is not a valid PDF document.")
        if document.needs_pass:
            raise ValueError("The PDF is password-protected and cannot be rendered.")
        page_count = document.page_count
        if page_count < 1:
            raise ValueError("The PDF contains no pages.")

        width = max(3, len(str(page_count)))
        output_directory.mkdir(parents=True, exist_ok=True)
        targets = [
            output_directory / f"page-{page_number:0{width}d}.png"
            for page_number in range(1, page_count + 1)
        ]
        existing = [path for path in targets if path.exists()]
        if existing and not replace_existing:
            names = ", ".join(path.name for path in existing)
            raise FileExistsError(
                f"Output file(s) already exist: {names}. "
                "Use --replace-existing to replace them."
            )

        with tempfile.TemporaryDirectory(
            prefix=".pdf-pages-",
            dir=output_directory,
        ) as temporary_directory:
            temporary_root = Path(temporary_directory)
            temporary_files: list[Path] = []
            for index, page in enumerate(document, start=1):
                temporary_path = (
                    temporary_root / f"page-{index:0{width}d}.png"
                )
                pixmap = page.get_pixmap(dpi=dpi, alpha=False)
                pixmap.save(temporary_path)
                if (
                    not temporary_path.is_file()
                    or temporary_path.stat().st_size == 0
                ):
                    raise RuntimeError(
                        f"Page {index} did not produce a non-empty PNG file."
                    )
                temporary_files.append(temporary_path)

            if sha256_file(source) != source_hash:
                raise RuntimeError(
                    "The source PDF changed during rendering; "
                    "generated images were not promoted."
                )

            for temporary_path, target in zip(temporary_files, targets):
                os.replace(temporary_path, target)

        return targets
    finally:
        if document is not None:
            document.close()


def parse_args(argv: Sequence[str] | None = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Render every PDF page to an ordered PNG using PyMuPDF."
    )
    parser.add_argument("pdf_path", type=Path, help="Path to an existing PDF")
    parser.add_argument(
        "--output-dir",
        type=Path,
        required=True,
        help="Directory for page-001.png, page-002.png, and so on",
    )
    parser.add_argument(
        "--dpi",
        type=int,
        default=DEFAULT_DPI,
        help=f"Rendering resolution in DPI (default: {DEFAULT_DPI})",
    )
    parser.add_argument(
        "--replace-existing",
        action="store_true",
        help="Replace predictable page PNG files that already exist",
    )
    args = parser.parse_args(argv)
    if args.dpi < 72:
        parser.error("--dpi must be at least 72")
    return args


def main(argv: Sequence[str] | None = None) -> int:
    configure_utf8_output()
    args = parse_args(argv)
    source: Path = args.pdf_path

    if not source.exists():
        print(f"ERROR: Source PDF not found: {source}", file=sys.stderr)
        return 2
    if not source.is_file():
        print(f"ERROR: Source path is not a file: {source}", file=sys.stderr)
        return 2
    if source.suffix.lower() != ".pdf":
        print(f"ERROR: Expected a .pdf source file: {source}", file=sys.stderr)
        return 2

    source = source.resolve()
    output_directory = args.output_dir.resolve()
    print(f"Source PDF: {source}")
    print(f"Output directory: {output_directory}")
    print(f"PyMuPDF version: {pymupdf.__version__}")
    print(f"Resolution: {args.dpi} DPI")

    try:
        generated = render_pages(
            source,
            output_directory,
            args.dpi,
            args.replace_existing,
        )
    except BaseException as exc:
        print(f"ERROR: PDF page rendering failed: {exc}", file=sys.stderr)
        return 2

    print(f"Page count: {len(generated)}")
    print("Generated PNG files:")
    for path in generated:
        print(f"  - {path}")
    print("SUCCESS: Every PDF page was rendered to PNG.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
