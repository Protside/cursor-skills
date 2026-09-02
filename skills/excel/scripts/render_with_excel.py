#!/usr/bin/env python3
"""Render an XLSX to PDF with Microsoft Excel without modifying the source."""

from __future__ import annotations

import argparse
import gc
import hashlib
import os
import shutil
import stat
import sys
import tempfile
import time
from pathlib import Path
from typing import Any, Sequence

try:
    import pythoncom
    import pywintypes
    import win32com.client
except ImportError as exc:  # pragma: no cover - depends on caller environment
    raise SystemExit(
        "ERROR: pywin32 is required. Install it with: python -m pip install pywin32"
    ) from exc


XL_CALCULATION_DONE = 0
XL_TYPE_PDF = 0
XL_QUALITY_STANDARD = 0
MSO_AUTOMATION_SECURITY_FORCE_DISABLE = 3


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


def wait_for_pending_calculation(
    application: Any,
    timeout_seconds: float,
) -> None:
    deadline = time.monotonic() + timeout_seconds
    while application.CalculationState != XL_CALCULATION_DONE:
        if time.monotonic() >= deadline:
            raise TimeoutError(
                f"Excel was still calculating after {timeout_seconds:g} seconds."
            )
        pythoncom.PumpWaitingMessages()
        time.sleep(0.2)


def format_com_error(error: BaseException) -> str:
    if isinstance(error, pywintypes.com_error):
        details = getattr(error, "excepinfo", None)
        if details and len(details) > 2 and details[2]:
            return f"{error}: {details[2]}"
    return str(error)


def render_pdf(
    source: Path,
    target: Path,
    timeout_seconds: float,
    replace_existing: bool,
) -> str:
    """Open a temporary source copy read-only and export it to a temporary PDF."""
    source_hash = sha256_file(source)
    output_directory = target.parent
    output_directory.mkdir(parents=True, exist_ok=True)

    if target.exists() and not replace_existing:
        raise FileExistsError(
            f"Output already exists: {target}. Use --replace-existing to replace it."
        )

    application = None
    workbook = None
    cleanup_errors: list[str] = []
    excel_version = "unknown"
    operation_error: BaseException | None = None

    with tempfile.TemporaryDirectory(
        prefix=".excel-render-",
        dir=output_directory,
    ) as temporary_directory:
        temporary_root = Path(temporary_directory)
        temporary_source = temporary_root / source.name
        temporary_pdf = temporary_root / target.name
        shutil.copy2(source, temporary_source)
        os.chmod(temporary_source, stat.S_IREAD | stat.S_IWRITE)

        pythoncom.CoInitialize()
        try:
            application = win32com.client.DispatchEx("Excel.Application")
            application.Visible = False
            application.DisplayAlerts = False
            application.ScreenUpdating = False
            application.EnableEvents = False
            application.AskToUpdateLinks = False
            application.AutomationSecurity = MSO_AUTOMATION_SECURITY_FORCE_DISABLE
            excel_version = str(application.Version)

            workbook = application.Workbooks.Open(
                str(temporary_source.resolve()),
                UpdateLinks=0,
                ReadOnly=True,
                IgnoreReadOnlyRecommended=True,
                AddToMru=False,
                Notify=False,
                Local=True,
            )
            wait_for_pending_calculation(application, timeout_seconds)
            workbook.ExportAsFixedFormat(
                Type=XL_TYPE_PDF,
                Filename=str(temporary_pdf.resolve()),
                Quality=XL_QUALITY_STANDARD,
                IncludeDocProperties=True,
                IgnorePrintAreas=False,
                OpenAfterPublish=False,
            )
            if not temporary_pdf.is_file() or temporary_pdf.stat().st_size == 0:
                raise RuntimeError("Excel did not produce a non-empty PDF output.")
        except BaseException as exc:
            operation_error = exc
        finally:
            if workbook is not None:
                try:
                    workbook.Close(SaveChanges=False)
                except BaseException as exc:
                    cleanup_errors.append(
                        f"Could not close the Excel workbook: {format_com_error(exc)}"
                    )
                workbook = None
            if application is not None:
                try:
                    application.Quit()
                except BaseException as exc:
                    cleanup_errors.append(
                        f"Could not quit Excel: {format_com_error(exc)}"
                    )
                application = None
            gc.collect()
            pythoncom.CoUninitialize()

        if operation_error is not None:
            detail = format_com_error(operation_error)
            if cleanup_errors:
                detail += " Cleanup errors: " + " | ".join(cleanup_errors)
            raise RuntimeError(detail) from operation_error
        if cleanup_errors:
            raise RuntimeError(" | ".join(cleanup_errors))
        if sha256_file(source) != source_hash:
            raise RuntimeError(
                "Source workbook changed during rendering; PDF was not promoted."
            )

        if target.exists():
            target.unlink()
        os.replace(temporary_pdf, target)

    return excel_version


def parse_args(argv: Sequence[str] | None = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description=(
            "Render an XLSX workbook to PDF with Microsoft Excel while preserving "
            "the source workbook."
        )
    )
    parser.add_argument("xlsx_path", type=Path, help="Path to an existing .xlsx file")
    parser.add_argument(
        "--output-dir",
        type=Path,
        required=True,
        help="Directory for the rendered PDF",
    )
    parser.add_argument(
        "--output-name",
        help="PDF filename (default: <source>-rendered.pdf)",
    )
    parser.add_argument(
        "--calculation-timeout",
        type=float,
        default=300.0,
        help="Maximum wait for pending Excel calculation in seconds (default: 300)",
    )
    parser.add_argument(
        "--replace-existing",
        action="store_true",
        help="Replace an existing PDF output; the source is always protected",
    )
    args = parser.parse_args(argv)
    if args.calculation_timeout <= 0:
        parser.error("--calculation-timeout must be greater than zero")
    return args


def main(argv: Sequence[str] | None = None) -> int:
    configure_utf8_output()
    args = parse_args(argv)
    source: Path = args.xlsx_path

    if sys.platform != "win32":
        print("ERROR: Microsoft Excel COM automation requires Windows.", file=sys.stderr)
        return 2
    if not source.exists():
        print(f"ERROR: Source workbook not found: {source}", file=sys.stderr)
        return 2
    if not source.is_file():
        print(f"ERROR: Source path is not a file: {source}", file=sys.stderr)
        return 2
    if source.suffix.lower() != ".xlsx":
        print(f"ERROR: Expected an .xlsx source workbook: {source}", file=sys.stderr)
        return 2

    source = source.resolve()
    output_directory = args.output_dir.resolve()
    output_name = args.output_name or f"{source.stem}-rendered.pdf"
    if Path(output_name).name != output_name or Path(output_name).suffix.lower() != ".pdf":
        print(
            "ERROR: --output-name must be a .pdf filename without directories.",
            file=sys.stderr,
        )
        return 2
    target = output_directory / output_name
    if target.resolve() == source:
        print(
            "ERROR: Refusing to use the source workbook as the PDF output path.",
            file=sys.stderr,
        )
        return 2

    print(f"Original workbook: {source}")
    print(f"Rendered PDF: {target}")
    try:
        excel_version = render_pdf(
            source,
            target,
            args.calculation_timeout,
            args.replace_existing,
        )
    except BaseException as exc:
        print(
            f"ERROR: Excel PDF export failed: {format_com_error(exc)}",
            file=sys.stderr,
        )
        return 2

    print(f"Excel version: {excel_version}")
    print(f"SUCCESS: Excel generated PDF: {target}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
