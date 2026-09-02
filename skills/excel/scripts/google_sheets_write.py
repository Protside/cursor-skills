#!/usr/bin/env python3
"""Reusable, verified Google Sheets read/write operations via the official API."""

from __future__ import annotations

import argparse
import json
import sys
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Sequence

from google_auth import GoogleAuthError, READWRITE_SCOPES, get_credentials

try:
    from googleapiclient.errors import HttpError
except ImportError:  # Allows CLI help before optional dependencies are installed.
    class HttpError(Exception):
        pass


class GoogleSheetsWriteError(RuntimeError):
    """A Sheets write or verification operation failed."""


def configure_utf8_output() -> None:
    for stream in (sys.stdout, sys.stderr):
        reconfigure = getattr(stream, "reconfigure", None)
        if reconfigure is not None:
            reconfigure(encoding="utf-8", errors="replace")


def grid_range(
    sheet_id: int,
    *,
    start_row: int | None = None,
    end_row: int | None = None,
    start_column: int | None = None,
    end_column: int | None = None,
) -> dict[str, int]:
    """Build a zero-based, end-exclusive GridRange."""
    result = {"sheetId": sheet_id}
    values = {
        "startRowIndex": start_row,
        "endRowIndex": end_row,
        "startColumnIndex": start_column,
        "endColumnIndex": end_column,
    }
    for key, value in values.items():
        if value is not None:
            if value < 0:
                raise ValueError(f"{key} cannot be negative")
            result[key] = value
    return result


def rgb(hex_color: str) -> dict[str, float]:
    """Convert six-digit RGB into a Sheets API color object."""
    value = hex_color.lstrip("#")
    if len(value) != 6:
        raise ValueError("Expected a six-digit RGB color.")
    try:
        red, green, blue = (
            int(value[index:index + 2], 16) / 255
            for index in (0, 2, 4)
        )
    except ValueError as exc:
        raise ValueError("Expected a six-digit RGB color.") from exc
    return {"red": red, "green": green, "blue": blue}


def solid_color(hex_color: str) -> dict[str, Any]:
    return {"rgbColor": rgb(hex_color)}


@dataclass
class VerifiedValueWrite:
    response: dict[str, Any]
    read_back: dict[str, Any]


class GoogleSheetsClient:
    """Small-scope Sheets v4 operations with explicit field masks."""

    def __init__(self, service: Any):
        self.service = service

    @classmethod
    def authorized(cls, *, interactive: bool = False) -> "GoogleSheetsClient":
        try:
            from googleapiclient.discovery import build
        except ImportError as exc:
            raise GoogleSheetsWriteError(
                "Official Google API libraries are not installed."
            ) from exc
        credentials = get_credentials(
            READWRITE_SCOPES,
            interactive=interactive,
        )
        service = build(
            "sheets",
            "v4",
            credentials=credentials,
            cache_discovery=False,
        )
        return cls(service)

    def close(self) -> None:
        close = getattr(self.service, "close", None)
        if close is not None:
            close()

    def create_spreadsheet(
        self,
        title: str,
        sheet_titles: Sequence[str],
        *,
        locale: str = "ru_RU",
        time_zone: str = "Europe/Moscow",
        rows: int = 1000,
        columns: int = 26,
    ) -> dict[str, Any]:
        if not title.strip():
            raise ValueError("Spreadsheet title cannot be empty.")
        if not sheet_titles or any(not item.strip() for item in sheet_titles):
            raise ValueError("At least one non-empty sheet title is required.")
        if len(set(sheet_titles)) != len(sheet_titles):
            raise ValueError("Sheet titles must be unique.")
        body = {
            "properties": {
                "title": title,
                "locale": locale,
                "timeZone": time_zone,
            },
            "sheets": [
                {
                    "properties": {
                        "title": sheet_title,
                        "index": index,
                        "gridProperties": {
                            "rowCount": rows,
                            "columnCount": columns,
                        },
                    }
                }
                for index, sheet_title in enumerate(sheet_titles)
            ],
        }
        return (
            self.service.spreadsheets()
            .create(
                body=body,
                fields=(
                    "spreadsheetId,spreadsheetUrl,properties(title,locale,"
                    "timeZone),sheets(properties(sheetId,title,index,"
                    "gridProperties))"
                ),
            )
            .execute()
        )

    def get_metadata(
        self,
        spreadsheet_id: str,
        *,
        include_grid_data: bool = False,
        ranges: Sequence[str] | None = None,
        fields: str | None = None,
    ) -> dict[str, Any]:
        arguments: dict[str, Any] = {
            "spreadsheetId": spreadsheet_id,
            "includeGridData": include_grid_data,
        }
        if ranges:
            arguments["ranges"] = list(ranges)
        if fields:
            arguments["fields"] = fields
        return (
            self.service.spreadsheets()
            .get(**arguments)
            .execute()
        )

    def read_range(
        self,
        spreadsheet_id: str,
        range_name: str,
        *,
        value_render_option: str = "FORMULA",
    ) -> dict[str, Any]:
        return (
            self.service.spreadsheets()
            .values()
            .get(
                spreadsheetId=spreadsheet_id,
                range=range_name,
                valueRenderOption=value_render_option,
                dateTimeRenderOption="FORMATTED_STRING",
            )
            .execute()
        )

    def read_ranges(
        self,
        spreadsheet_id: str,
        ranges: Sequence[str],
        *,
        value_render_option: str = "FORMULA",
    ) -> dict[str, Any]:
        return (
            self.service.spreadsheets()
            .values()
            .batchGet(
                spreadsheetId=spreadsheet_id,
                ranges=list(ranges),
                valueRenderOption=value_render_option,
                dateTimeRenderOption="FORMATTED_STRING",
            )
            .execute()
        )

    def write_values(
        self,
        spreadsheet_id: str,
        range_name: str,
        values: Sequence[Sequence[Any]],
        *,
        input_option: str = "USER_ENTERED",
        verify: bool = True,
    ) -> VerifiedValueWrite:
        response = (
            self.service.spreadsheets()
            .values()
            .update(
                spreadsheetId=spreadsheet_id,
                range=range_name,
                valueInputOption=input_option,
                includeValuesInResponse=False,
                body={"majorDimension": "ROWS", "values": list(values)},
            )
            .execute()
        )
        read_back = (
            self.read_range(spreadsheet_id, response["updatedRange"])
            if verify
            else {}
        )
        return VerifiedValueWrite(response=response, read_back=read_back)

    def batch_write_values(
        self,
        spreadsheet_id: str,
        data: Sequence[dict[str, Any]],
        *,
        input_option: str = "USER_ENTERED",
        verify_ranges: Sequence[str] | None = None,
    ) -> tuple[dict[str, Any], dict[str, Any]]:
        response = (
            self.service.spreadsheets()
            .values()
            .batchUpdate(
                spreadsheetId=spreadsheet_id,
                body={
                    "valueInputOption": input_option,
                    "includeValuesInResponse": False,
                    "data": list(data),
                },
            )
            .execute()
        )
        read_back = (
            self.read_ranges(spreadsheet_id, verify_ranges)
            if verify_ranges
            else {}
        )
        return response, read_back

    def append_rows(
        self,
        spreadsheet_id: str,
        range_name: str,
        rows: Sequence[Sequence[Any]],
        *,
        input_option: str = "USER_ENTERED",
        insert_option: str = "INSERT_ROWS",
        verify: bool = True,
    ) -> VerifiedValueWrite:
        response = (
            self.service.spreadsheets()
            .values()
            .append(
                spreadsheetId=spreadsheet_id,
                range=range_name,
                valueInputOption=input_option,
                insertDataOption=insert_option,
                body={"majorDimension": "ROWS", "values": list(rows)},
            )
            .execute()
        )
        updated_range = (
            response.get("updates", {}).get("updatedRange")
        )
        read_back = (
            self.read_range(spreadsheet_id, updated_range)
            if verify and updated_range
            else {}
        )
        return VerifiedValueWrite(response=response, read_back=read_back)

    def batch_update(
        self,
        spreadsheet_id: str,
        requests: Sequence[dict[str, Any]],
        *,
        include_spreadsheet_in_response: bool = False,
        response_ranges: Sequence[str] | None = None,
    ) -> dict[str, Any]:
        if not requests:
            raise ValueError("At least one batchUpdate request is required.")
        body: dict[str, Any] = {
            "requests": list(requests),
            "includeSpreadsheetInResponse": include_spreadsheet_in_response,
        }
        if response_ranges:
            body["responseRanges"] = list(response_ranges)
            body["responseIncludeGridData"] = False
        return (
            self.service.spreadsheets()
            .batchUpdate(
                spreadsheetId=spreadsheet_id,
                body=body,
            )
            .execute()
        )

    def add_sheet_request(
        self,
        title: str,
        *,
        rows: int = 1000,
        columns: int = 26,
        index: int | None = None,
    ) -> dict[str, Any]:
        properties: dict[str, Any] = {
            "title": title,
            "gridProperties": {
                "rowCount": rows,
                "columnCount": columns,
            },
        }
        if index is not None:
            properties["index"] = index
        return {"addSheet": {"properties": properties}}

    def delete_sheet_request(self, sheet_id: int) -> dict[str, Any]:
        return {"deleteSheet": {"sheetId": sheet_id}}

    def format_range_request(
        self,
        target_range: dict[str, int],
        user_entered_format: dict[str, Any],
        *,
        fields: str,
    ) -> dict[str, Any]:
        if not fields or not fields.startswith("userEnteredFormat"):
            raise ValueError(
                "Formatting requires an explicit userEnteredFormat field mask."
            )
        return {
            "repeatCell": {
                "range": target_range,
                "cell": {"userEnteredFormat": user_entered_format},
                "fields": fields,
            }
        }

    def resize_dimension_request(
        self,
        sheet_id: int,
        dimension: str,
        start_index: int,
        end_index: int,
        pixel_size: int,
    ) -> dict[str, Any]:
        normalized = dimension.upper()
        if normalized not in ("ROWS", "COLUMNS"):
            raise ValueError("Dimension must be ROWS or COLUMNS.")
        return {
            "updateDimensionProperties": {
                "range": {
                    "sheetId": sheet_id,
                    "dimension": normalized,
                    "startIndex": start_index,
                    "endIndex": end_index,
                },
                "properties": {"pixelSize": pixel_size},
                "fields": "pixelSize",
            }
        }

    def freeze_request(
        self,
        sheet_id: int,
        *,
        rows: int = 0,
        columns: int = 0,
    ) -> dict[str, Any]:
        return {
            "updateSheetProperties": {
                "properties": {
                    "sheetId": sheet_id,
                    "gridProperties": {
                        "frozenRowCount": rows,
                        "frozenColumnCount": columns,
                    },
                },
                "fields": (
                    "gridProperties.frozenRowCount,"
                    "gridProperties.frozenColumnCount"
                ),
            }
        }

    def dropdown_request(
        self,
        target_range: dict[str, int],
        source_range_a1: str,
        *,
        strict: bool = True,
        show_dropdown: bool = True,
    ) -> dict[str, Any]:
        return {
            "setDataValidation": {
                "range": target_range,
                "rule": {
                    "condition": {
                        "type": "ONE_OF_RANGE",
                        "values": [{"userEnteredValue": source_range_a1}],
                    },
                    "strict": strict,
                    "showCustomUi": show_dropdown,
                },
            }
        }

    def list_dropdown_request(
        self,
        target_range: dict[str, int],
        values: Sequence[str],
        *,
        strict: bool = True,
    ) -> dict[str, Any]:
        return {
            "setDataValidation": {
                "range": target_range,
                "rule": {
                    "condition": {
                        "type": "ONE_OF_LIST",
                        "values": [
                            {"userEnteredValue": value}
                            for value in values
                        ],
                    },
                    "strict": strict,
                    "showCustomUi": True,
                },
            }
        }

    def checkbox_request(
        self,
        target_range: dict[str, int],
        *,
        checked_value: str | None = None,
        unchecked_value: str | None = None,
    ) -> dict[str, Any]:
        values = []
        if checked_value is not None:
            values.append({"userEnteredValue": checked_value})
            values.append({"userEnteredValue": unchecked_value or ""})
        return {
            "setDataValidation": {
                "range": target_range,
                "rule": {
                    "condition": {"type": "BOOLEAN", "values": values},
                    "strict": True,
                    "showCustomUi": True,
                },
            }
        }

    def validation_request(
        self,
        target_range: dict[str, int],
        condition_type: str,
        values: Sequence[str] = (),
        *,
        strict: bool = True,
        input_message: str | None = None,
    ) -> dict[str, Any]:
        rule: dict[str, Any] = {
            "condition": {
                "type": condition_type,
                "values": [
                    {"userEnteredValue": value}
                    for value in values
                ],
            },
            "strict": strict,
        }
        if input_message:
            rule["inputMessage"] = input_message
        return {"setDataValidation": {"range": target_range, "rule": rule}}

    def conditional_format_request(
        self,
        ranges: Sequence[dict[str, int]],
        boolean_rule: dict[str, Any],
        *,
        index: int = 0,
    ) -> dict[str, Any]:
        return {
            "addConditionalFormatRule": {
                "rule": {
                    "ranges": list(ranges),
                    "booleanRule": boolean_rule,
                },
                "index": index,
            }
        }

    def protect_range_request(
        self,
        target_range: dict[str, int],
        description: str,
        *,
        warning_only: bool = True,
    ) -> dict[str, Any]:
        return {
            "addProtectedRange": {
                "protectedRange": {
                    "range": target_range,
                    "description": description,
                    "warningOnly": warning_only,
                }
            }
        }

    def add_chart_request(self, chart: dict[str, Any]) -> dict[str, Any]:
        if "spec" not in chart or "position" not in chart:
            raise ValueError("A chart requires spec and position.")
        return {"addChart": {"chart": chart}}


def load_json_argument(value: str) -> Any:
    candidate = Path(value)
    if candidate.is_file():
        return json.loads(candidate.read_text(encoding="utf-8"))
    return json.loads(value)


def parse_args(argv: Sequence[str] | None = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description=(
            "Safe Google Sheets read/write operations using only the official "
            "Sheets API."
        )
    )
    parser.add_argument(
        "--authorize",
        action="store_true",
        help="Allow explicit first-time browser authorization",
    )
    commands = parser.add_subparsers(dest="command", required=True)

    create = commands.add_parser("create", help="Create a spreadsheet")
    create.add_argument("title")
    create.add_argument("--sheets", nargs="+", required=True)

    read = commands.add_parser("read", help="Read one A1 range")
    read.add_argument("spreadsheet_id")
    read.add_argument("range")

    write = commands.add_parser("write", help="Write and read back one range")
    write.add_argument("spreadsheet_id")
    write.add_argument("range")
    write.add_argument(
        "values_json",
        help="JSON rows or path to a UTF-8 JSON file",
    )

    append = commands.add_parser("append", help="Append and read back rows")
    append.add_argument("spreadsheet_id")
    append.add_argument("range")
    append.add_argument(
        "values_json",
        help="JSON rows or path to a UTF-8 JSON file",
    )

    batch = commands.add_parser(
        "batch",
        help="Execute batchUpdate requests from JSON",
    )
    batch.add_argument("spreadsheet_id")
    batch.add_argument(
        "requests_json",
        help="JSON request list or path to a UTF-8 JSON file",
    )
    return parser.parse_args(argv)


def main(argv: Sequence[str] | None = None) -> int:
    configure_utf8_output()
    args = parse_args(argv)
    client: GoogleSheetsClient | None = None
    try:
        client = GoogleSheetsClient.authorized(interactive=args.authorize)
        if args.command == "create":
            response = client.create_spreadsheet(args.title, args.sheets)
            print(f"Spreadsheet ID: {response['spreadsheetId']}")
            print(f"Spreadsheet URL: {response['spreadsheetUrl']}")
        elif args.command == "read":
            print(json.dumps(
                client.read_range(args.spreadsheet_id, args.range),
                ensure_ascii=False,
                indent=2,
            ))
        elif args.command == "write":
            result = client.write_values(
                args.spreadsheet_id,
                args.range,
                load_json_argument(args.values_json),
            )
            print(json.dumps(
                {"response": result.response, "readBack": result.read_back},
                ensure_ascii=False,
                indent=2,
            ))
        elif args.command == "append":
            result = client.append_rows(
                args.spreadsheet_id,
                args.range,
                load_json_argument(args.values_json),
            )
            print(json.dumps(
                {"response": result.response, "readBack": result.read_back},
                ensure_ascii=False,
                indent=2,
            ))
        elif args.command == "batch":
            response = client.batch_update(
                args.spreadsheet_id,
                load_json_argument(args.requests_json),
            )
            print(json.dumps(response, ensure_ascii=False, indent=2))
        return 0
    except HttpError as exc:
        status = getattr(getattr(exc, "resp", None), "status", "unknown")
        reason_method = getattr(exc, "_get_reason", None)
        reason = (
            reason_method()
            if reason_method is not None
            else "Google Sheets API request failed"
        )
        print(f"ERROR: Google Sheets API error {status}: {reason}", file=sys.stderr)
        return 2
    except (GoogleAuthError, GoogleSheetsWriteError, ValueError) as exc:
        print(f"ERROR: {exc}", file=sys.stderr)
        return 2
    finally:
        if client is not None:
            client.close()


if __name__ == "__main__":
    raise SystemExit(main())
