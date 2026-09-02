#!/usr/bin/env python3
"""Read-only structural inspection for Google Sheets via the official API."""

from __future__ import annotations

import argparse
import re
import sys
from typing import Any, Sequence
from urllib.parse import urlparse

from google_auth import GoogleAuthError, READONLY_SCOPES, get_credentials


SPREADSHEET_ID_PATTERN = re.compile(r"^[A-Za-z0-9_-]{10,}$")
URL_ID_PATTERN = re.compile(r"/spreadsheets/d/([A-Za-z0-9_-]+)")


class InspectionError(RuntimeError):
    """The read-only inspection could not be completed."""


def configure_utf8_output() -> None:
    for stream in (sys.stdout, sys.stderr):
        reconfigure = getattr(stream, "reconfigure", None)
        if reconfigure is not None:
            reconfigure(encoding="utf-8", errors="replace")


def extract_spreadsheet_id(source: str) -> str:
    value = source.strip()
    parsed = urlparse(value)
    if parsed.scheme or parsed.netloc:
        if parsed.scheme not in ("http", "https"):
            raise InspectionError("Google Sheets URL must use http or https.")
        match = URL_ID_PATTERN.search(parsed.path)
        if not match:
            raise InspectionError(
                "Could not extract a spreadsheet ID from the Google Sheets URL."
            )
        spreadsheet_id = match.group(1)
    else:
        spreadsheet_id = value

    if not SPREADSHEET_ID_PATTERN.fullmatch(spreadsheet_id):
        raise InspectionError("Spreadsheet ID contains invalid characters.")
    return spreadsheet_id


def quote_sheet_title(title: str) -> str:
    return "'" + title.replace("'", "''") + "'"


def column_name(index: int) -> str:
    if index < 1:
        raise ValueError("Column index must be positive.")
    result = ""
    while index:
        index, remainder = divmod(index - 1, 26)
        result = chr(65 + remainder) + result
    return result


def cell_address(row_index: int, column_index: int) -> str:
    return f"{column_name(column_index + 1)}{row_index + 1}"


def grid_range_to_a1(grid_range: dict[str, Any], title: str | None = None) -> str:
    start_row = int(grid_range.get("startRowIndex", 0))
    start_column = int(grid_range.get("startColumnIndex", 0))
    end_row = grid_range.get("endRowIndex")
    end_column = grid_range.get("endColumnIndex")

    start = f"{column_name(start_column + 1)}{start_row + 1}"
    if end_row is None and end_column is None:
        body = start + ":"
    elif end_row is None:
        body = f"{column_name(start_column + 1)}:{column_name(int(end_column))}"
    elif end_column is None:
        body = f"{start_row + 1}:{int(end_row)}"
    else:
        body = f"{start}:{column_name(int(end_column))}{int(end_row)}"
    return f"{quote_sheet_title(title)}!{body}" if title else body


def limited(items: Sequence[Any], maximum: int) -> tuple[Sequence[Any], int]:
    if maximum == 0 or len(items) <= maximum:
        return items, 0
    return items[:maximum], len(items) - maximum


def print_items(
    heading: str,
    items: Sequence[str],
    maximum: int,
    *,
    empty_text: str = "(none)",
) -> None:
    print(f"  {heading}:")
    shown, remaining = limited(items, maximum)
    if not shown:
        print(f"    {empty_text}")
        return
    for item in shown:
        print(f"    - {item}")
    if remaining:
        print(f"    ... {remaining} more (use --max-items 0 to show all)")


def detect_used_extent(values: Sequence[Sequence[Any]]) -> tuple[int, int]:
    used_rows = len(values)
    used_columns = max((len(row) for row in values), default=0)
    return used_rows, used_columns


def compact_format(format_payload: dict[str, Any]) -> str:
    selected: dict[str, Any] = {}
    for key in (
        "numberFormat",
        "backgroundColorStyle",
        "horizontalAlignment",
        "verticalAlignment",
        "wrapStrategy",
    ):
        if key in format_payload:
            selected[key] = format_payload[key]
    text_format = format_payload.get("textFormat")
    if text_format:
        selected["textFormat"] = {
            key: text_format[key]
            for key in ("fontFamily", "fontSize", "bold", "italic")
            if key in text_format
        }
    return repr(selected)


def inspect_grid_data(sheet_payload: dict[str, Any]) -> dict[str, list[str]]:
    sampled_cells: list[str] = []
    formulas: list[str] = []
    validations: list[str] = []
    checkboxes: list[str] = []
    formats: list[str] = []

    for grid_data in sheet_payload.get("data", []):
        start_row = int(grid_data.get("startRow", 0))
        start_column = int(grid_data.get("startColumn", 0))
        for row_offset, row_payload in enumerate(grid_data.get("rowData", [])):
            for column_offset, cell in enumerate(row_payload.get("values", [])):
                coordinate = cell_address(
                    start_row + row_offset,
                    start_column + column_offset,
                )
                entered = cell.get("userEnteredValue", {})
                formula = entered.get("formulaValue")
                formatted_value = cell.get("formattedValue")
                if formula is not None or formatted_value is not None:
                    value_text = repr(formatted_value)
                    if formula is not None:
                        value_text += f" | formula={formula}"
                    sampled_cells.append(f"{coordinate}: {value_text}")
                if formula is not None:
                    formulas.append(f"{coordinate}: {formula}")

                validation = cell.get("dataValidation")
                if validation:
                    condition = validation.get("condition", {})
                    condition_type = condition.get("type", "(unknown)")
                    values = [
                        item.get("userEnteredValue")
                        for item in condition.get("values", [])
                    ]
                    description = (
                        f"{coordinate}: type={condition_type}, "
                        f"values={values}, strict={validation.get('strict')}, "
                        f"showCustomUi={validation.get('showCustomUi')}"
                    )
                    validations.append(description)
                    if condition_type == "BOOLEAN":
                        checkboxes.append(description)

                user_format = cell.get("userEnteredFormat")
                if user_format:
                    formats.append(
                        f"{coordinate}: {compact_format(user_format)}"
                    )

    return {
        "sampled_cells": sampled_cells,
        "formulas": formulas,
        "validations": validations,
        "checkboxes": checkboxes,
        "formats": formats,
    }


def describe_conditional_rule(
    rule: dict[str, Any],
    title: str,
) -> str:
    ranges = [
        grid_range_to_a1(item, title)
        for item in rule.get("ranges", [])
    ]
    if "booleanRule" in rule:
        condition = rule["booleanRule"].get("condition", {})
        rule_type = f"boolean:{condition.get('type', '(unknown)')}"
    elif "gradientRule" in rule:
        rule_type = "gradient"
    else:
        rule_type = "unknown"
    return f"ranges={ranges}, type={rule_type}"


def inspect_spreadsheet(
    service: Any,
    spreadsheet_id: str,
    *,
    sample_rows: int,
    sample_columns: int,
    max_grid_cells: int,
    max_items: int,
) -> None:
    metadata = (
        service.spreadsheets()
        .get(
            spreadsheetId=spreadsheet_id,
            includeGridData=False,
        )
        .execute()
    )
    properties = metadata.get("properties", {})
    sheets = metadata.get("sheets", [])

    print("GOOGLE SHEETS READ-ONLY INSPECTION")
    print(f"Spreadsheet ID: {metadata.get('spreadsheetId', spreadsheet_id)}")
    print(f"Title: {properties.get('title', '(untitled)')}")
    print(f"Locale: {properties.get('locale', '(not reported)')}")
    print(f"Time zone: {properties.get('timeZone', '(not reported)')}")
    print(f"Sheet count: {len(sheets)}")

    sheet_titles_by_id = {
        item.get("properties", {}).get("sheetId"):
        item.get("properties", {}).get("title")
        for item in sheets
    }
    named_ranges = []
    for named_range in metadata.get("namedRanges", []):
        named_grid_range = named_range.get("range", {})
        named_range_a1 = grid_range_to_a1(
            named_grid_range,
            sheet_titles_by_id.get(named_grid_range.get("sheetId")),
        )
        named_ranges.append(
            f"{named_range.get('name', '(unnamed)')} | "
            f"id={named_range.get('namedRangeId', '(none)')} | "
            f"range={named_range_a1}"
        )
    print_items("Named ranges", named_ranges, max_items)

    used_extents: dict[str, tuple[int, int]] = {}
    inspection_ranges: list[str] = []
    range_notes: dict[str, str] = {}

    for sheet in sheets:
        sheet_properties = sheet.get("properties", {})
        title = sheet_properties.get("title", "")
        values_response = (
            service.spreadsheets()
            .values()
            .get(
                spreadsheetId=spreadsheet_id,
                range=quote_sheet_title(title),
                majorDimension="ROWS",
                valueRenderOption="FORMULA",
                dateTimeRenderOption="FORMATTED_STRING",
            )
            .execute()
        )
        values = values_response.get("values", [])
        used_rows, used_columns = detect_used_extent(values)
        used_extents[title] = (used_rows, used_columns)
        grid_properties = sheet_properties.get("gridProperties", {})
        available_rows = int(grid_properties.get("rowCount", sample_rows))
        available_columns = int(
            grid_properties.get("columnCount", sample_columns)
        )
        if used_rows and used_columns:
            candidate_rows = min(
                available_rows,
                max(used_rows, sample_rows),
            )
            candidate_columns = min(
                available_columns,
                max(used_columns, sample_columns),
            )
            if candidate_rows * candidate_columns <= max_grid_cells:
                inspect_rows = candidate_rows
                inspect_columns = candidate_columns
                range_notes[title] = (
                    "full values-based used range plus nearby blank cells"
                )
            elif used_rows * used_columns <= max_grid_cells:
                inspect_rows = used_rows
                inspect_columns = used_columns
                range_notes[title] = "full values-based used range"
            else:
                inspect_rows = min(used_rows, sample_rows)
                inspect_columns = min(used_columns, sample_columns)
                range_notes[title] = (
                    "sampled because used range exceeds --max-grid-cells"
                )
        else:
            inspect_rows = min(available_rows, sample_rows)
            inspect_columns = min(available_columns, sample_columns)
            range_notes[title] = "sampled because no value-based range was found"

        inspect_rows = max(inspect_rows, 1)
        inspect_columns = max(inspect_columns, 1)
        inspection_ranges.append(
            f"{quote_sheet_title(title)}!A1:"
            f"{column_name(inspect_columns)}{inspect_rows}"
        )

    grid_response = (
        service.spreadsheets()
        .get(
            spreadsheetId=spreadsheet_id,
            ranges=inspection_ranges,
            includeGridData=True,
        )
        .execute()
    )
    grid_sheets = {
        item.get("properties", {}).get("title"): item
        for item in grid_response.get("sheets", [])
    }

    for sheet in sheets:
        sheet_properties = sheet.get("properties", {})
        title = sheet_properties.get("title", "(untitled)")
        grid_properties = sheet_properties.get("gridProperties", {})
        used_rows, used_columns = used_extents.get(title, (0, 0))
        used_range = (
            f"A1:{column_name(used_columns)}{used_rows}"
            if used_rows and used_columns
            else "(no value-based used range detected)"
        )

        print()
        print(f"SHEET: {title}")
        print(f"  Sheet ID: {sheet_properties.get('sheetId')}")
        print(f"  Index: {sheet_properties.get('index')}")
        print(f"  Type: {sheet_properties.get('sheetType', 'GRID')}")
        print(f"  Hidden: {sheet_properties.get('hidden', False)}")
        print(
            "  Dimensions: "
            f"rows={grid_properties.get('rowCount', '(not reported)')}, "
            f"columns={grid_properties.get('columnCount', '(not reported)')}"
        )
        print(
            "  Frozen: "
            f"rows={grid_properties.get('frozenRowCount', 0)}, "
            f"columns={grid_properties.get('frozenColumnCount', 0)}"
        )
        print(f"  Used range: {used_range}")
        print(f"  Grid inspection: {range_notes.get(title)}")

        merges = [
            grid_range_to_a1(item, title)
            for item in sheet.get("merges", [])
        ]
        print_items("Merged ranges", merges, max_items)

        protected_ranges = []
        for protected in sheet.get("protectedRanges", []):
            protected_ranges.append(
                f"id={protected.get('protectedRangeId')}, "
                f"description={protected.get('description')!r}, "
                f"range={grid_range_to_a1(protected.get('range', {}), title)}, "
                f"warningOnly={protected.get('warningOnly', False)}"
            )
        print_items("Protected ranges", protected_ranges, max_items)

        basic_filter = sheet.get("basicFilter")
        filter_items = (
            [
                "range="
                + grid_range_to_a1(basic_filter.get("range", {}), title)
            ]
            if basic_filter
            else []
        )
        print_items("Basic filter", filter_items, max_items)

        filter_views = []
        for view in sheet.get("filterViews", []):
            filter_views.append(
                f"id={view.get('filterViewId')}, "
                f"title={view.get('title')!r}, "
                f"range={grid_range_to_a1(view.get('range', {}), title)}"
            )
        print_items("Filter views", filter_views, max_items)

        conditional_rules = [
            describe_conditional_rule(rule, title)
            for rule in sheet.get("conditionalFormats", [])
        ]
        print_items(
            "Conditional formatting rules",
            conditional_rules,
            max_items,
        )

        charts = []
        for chart in sheet.get("charts", []):
            charts.append(
                f"id={chart.get('chartId')}, "
                f"title={chart.get('spec', {}).get('title')!r}, "
                f"position={chart.get('position', {})}"
            )
        print_items("Charts / embedded charts", charts, max_items)

        slicers = [
            f"id={item.get('slicerId')}, position={item.get('position', {})}"
            for item in sheet.get("slicers", [])
        ]
        print_items("Slicers / embedded slicers", slicers, max_items)

        grid_details = inspect_grid_data(grid_sheets.get(title, {}))
        print_items(
            f"Sampled cell values ({len(grid_details['sampled_cells'])})",
            grid_details["sampled_cells"],
            max_items,
        )
        print_items(
            f"Formulas in inspected range ({len(grid_details['formulas'])})",
            grid_details["formulas"],
            max_items,
        )
        print_items(
            "Data validations in inspected range "
            f"({len(grid_details['validations'])})",
            grid_details["validations"],
            max_items,
        )
        print_items(
            f"Checkboxes in inspected range ({len(grid_details['checkboxes'])})",
            grid_details["checkboxes"],
            max_items,
        )
        print_items(
            "User-entered formatting in inspected range "
            f"({len(grid_details['formats'])})",
            grid_details["formats"],
            max_items,
        )

    print()
    print(
        "Read-only safety: only spreadsheets.get and values.get requests "
        "were used; no write, clear, update, or batchUpdate request was made."
    )
    print(
        "Inspection boundary: values-based used ranges omit cells that contain "
        "only formatting. Large grids may be sampled as explicitly reported."
    )
    print(
        "Embedded-object boundary: the Sheets API reports embedded charts and "
        "slicers, but does not comprehensively expose every object type."
    )


def parse_args(argv: Sequence[str] | None = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description=(
            "Inspect a Google Sheet read-only through the official Google "
            "Sheets API."
        )
    )
    parser.add_argument(
        "spreadsheet",
        help="Google Sheets URL or spreadsheet ID",
    )
    parser.add_argument(
        "--authorize",
        action="store_true",
        help="Allow first-time browser OAuth if no valid local token exists",
    )
    parser.add_argument(
        "--sample-rows",
        type=int,
        default=20,
        help="Rows to sample when full grid inspection is capped (default: 20)",
    )
    parser.add_argument(
        "--sample-columns",
        type=int,
        default=12,
        help="Columns to sample when full grid inspection is capped (default: 12)",
    )
    parser.add_argument(
        "--max-grid-cells",
        type=int,
        default=20000,
        help="Inspect a full used range up to this many cells (default: 20000)",
    )
    parser.add_argument(
        "--max-items",
        type=int,
        default=50,
        help="Maximum detailed entries per section; 0 shows all (default: 50)",
    )
    args = parser.parse_args(argv)
    for name in ("sample_rows", "sample_columns", "max_grid_cells"):
        if getattr(args, name) < 1:
            parser.error(f"--{name.replace('_', '-')} must be positive")
    if args.max_items < 0:
        parser.error("--max-items cannot be negative")
    return args


def main(argv: Sequence[str] | None = None) -> int:
    configure_utf8_output()
    args = parse_args(argv)

    try:
        spreadsheet_id = extract_spreadsheet_id(args.spreadsheet)
        credentials = get_credentials(
            READONLY_SCOPES,
            interactive=args.authorize,
        )
        try:
            from googleapiclient.discovery import build
            from googleapiclient.errors import HttpError
        except ImportError as exc:
            raise InspectionError(
                "Official Google API libraries are not installed. Install: "
                "google-api-python-client google-auth google-auth-oauthlib "
                "google-auth-httplib2"
            ) from exc

        service = build(
            "sheets",
            "v4",
            credentials=credentials,
            cache_discovery=False,
        )
        try:
            inspect_spreadsheet(
                service,
                spreadsheet_id,
                sample_rows=args.sample_rows,
                sample_columns=args.sample_columns,
                max_grid_cells=args.max_grid_cells,
                max_items=args.max_items,
            )
        except HttpError as exc:
            status = getattr(exc.resp, "status", "unknown")
            reason = exc._get_reason() or "Google Sheets API request failed"
            raise InspectionError(
                f"Google Sheets API error {status}: {reason}"
            ) from exc
        return 0
    except (GoogleAuthError, InspectionError) as exc:
        print(f"ERROR: {exc}", file=sys.stderr)
        return 2


if __name__ == "__main__":
    raise SystemExit(main())
