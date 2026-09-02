#!/usr/bin/env python3
"""User-local OAuth 2.0 authentication for the official Google Sheets API."""

from __future__ import annotations

import argparse
import json
import os
import stat
import sys
import tempfile
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Sequence


SHEETS_READONLY_SCOPE = "https://www.googleapis.com/auth/spreadsheets.readonly"
SHEETS_READWRITE_SCOPE = "https://www.googleapis.com/auth/spreadsheets"
READONLY_SCOPES = (SHEETS_READONLY_SCOPE,)
READWRITE_SCOPES = (SHEETS_READWRITE_SCOPE,)


class GoogleAuthError(RuntimeError):
    """Authentication failed without exposing credential material."""


def configure_utf8_output() -> None:
    for stream in (sys.stdout, sys.stderr):
        reconfigure = getattr(stream, "reconfigure", None)
        if reconfigure is not None:
            reconfigure(encoding="utf-8", errors="replace")


def google_storage_directory() -> Path:
    local_app_data = os.environ.get("LOCALAPPDATA")
    if not local_app_data:
        raise GoogleAuthError(
            "LOCALAPPDATA is unavailable. A user-local Google credential "
            "directory cannot be determined safely."
        )
    return Path(local_app_data) / "CursorExcelSkill" / "google"


def oauth_client_path() -> Path:
    return google_storage_directory() / "oauth-client.json"


def token_path_for_scopes(scopes: Sequence[str]) -> Path:
    requested = set(scopes)
    if SHEETS_READWRITE_SCOPE in requested:
        filename = "token-sheets-readwrite.json"
    elif requested == {SHEETS_READONLY_SCOPE}:
        filename = "token-sheets-readonly.json"
    else:
        raise GoogleAuthError(
            "Unsupported OAuth scope set. Use the declared Sheets read-only "
            "or read/write scope."
        )
    return google_storage_directory() / filename


def load_google_dependencies() -> tuple[Any, Any, Any]:
    try:
        from google.auth.transport.requests import Request
        from google.oauth2.credentials import Credentials
        from google_auth_oauthlib.flow import InstalledAppFlow
    except ImportError as exc:
        raise GoogleAuthError(
            "Official Google libraries are not installed. Install: "
            "google-api-python-client google-auth google-auth-oauthlib "
            "google-auth-httplib2"
        ) from exc
    return Request, Credentials, InstalledAppFlow


def load_oauth_client_config(path: Path) -> dict[str, Any]:
    if not path.is_file():
        raise GoogleAuthError(
            f"OAuth client configuration is missing: {path}. "
            "Download a Desktop app OAuth client JSON from Google Cloud and "
            "save it at this user-local path."
        )
    try:
        payload = json.loads(path.read_text(encoding="utf-8"))
        installed = payload["installed"]
        for field in ("client_id", "client_secret", "token_uri"):
            if not installed.get(field):
                raise KeyError(field)
    except (OSError, ValueError, KeyError) as exc:
        raise GoogleAuthError(
            "OAuth client configuration is invalid or unreadable. "
            "Expected an OAuth 2.0 Desktop app JSON file."
        ) from exc
    return payload


def load_stored_credentials(
    token_path: Path,
    client_payload: dict[str, Any],
    requested_scopes: Sequence[str],
) -> Any | None:
    if not token_path.is_file():
        return None

    _, Credentials, _ = load_google_dependencies()
    try:
        payload = json.loads(token_path.read_text(encoding="utf-8"))
        stored_scopes = tuple(payload.get("scopes") or ())
        if not set(requested_scopes).issubset(stored_scopes):
            raise GoogleAuthError(
                "The stored OAuth token does not include the required scope. "
                "Authorize again using the appropriate token file."
            )
        expiry_text = payload.get("expiry")
        expiry = (
            datetime.fromisoformat(expiry_text.replace("Z", "+00:00"))
            if expiry_text
            else None
        )
        if expiry is not None and expiry.tzinfo is not None:
            expiry = (
                expiry.astimezone(timezone.utc)
                .replace(tzinfo=None)
            )
        installed = client_payload["installed"]
        return Credentials(
            token=payload.get("token"),
            refresh_token=payload.get("refresh_token"),
            token_uri=installed["token_uri"],
            client_id=installed["client_id"],
            client_secret=installed["client_secret"],
            scopes=stored_scopes,
            expiry=expiry,
        )
    except GoogleAuthError:
        raise
    except (OSError, ValueError, KeyError, TypeError) as exc:
        raise GoogleAuthError(
            "The stored OAuth token is invalid or unreadable. "
            "Remove it and authorize again."
        ) from exc


def save_token(
    path: Path,
    credentials: Any,
    requested_scopes: Sequence[str],
) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    expiry = credentials.expiry
    if expiry is not None and expiry.tzinfo is None:
        expiry = expiry.replace(tzinfo=timezone.utc)
    payload = {
        "token": credentials.token,
        "refresh_token": credentials.refresh_token,
        "scopes": list(credentials.scopes or requested_scopes),
        "expiry": expiry.isoformat() if expiry is not None else None,
    }

    temporary_path: Path | None = None
    try:
        with tempfile.NamedTemporaryFile(
            mode="w",
            encoding="utf-8",
            prefix=".google-token-",
            suffix=".json",
            dir=path.parent,
            delete=False,
        ) as handle:
            json.dump(payload, handle, ensure_ascii=False, indent=2)
            temporary_path = Path(handle.name)
        os.chmod(temporary_path, stat.S_IREAD | stat.S_IWRITE)
        os.replace(temporary_path, path)
    finally:
        if temporary_path is not None and temporary_path.exists():
            temporary_path.unlink()


def get_credentials(
    scopes: Sequence[str] = READONLY_SCOPES,
    *,
    interactive: bool = False,
) -> Any:
    """Return valid credentials, optionally running first-time browser OAuth."""
    Request, _, InstalledAppFlow = load_google_dependencies()
    client_path = oauth_client_path()
    token_path = token_path_for_scopes(scopes)
    client_payload = load_oauth_client_config(client_path)
    credentials = load_stored_credentials(token_path, client_payload, scopes)

    if credentials is not None and credentials.valid:
        return credentials

    if credentials is not None and credentials.expired and credentials.refresh_token:
        try:
            credentials.refresh(Request())
            save_token(token_path, credentials, scopes)
            return credentials
        except Exception as exc:
            raise GoogleAuthError(
                "OAuth token refresh failed. Run explicit browser "
                "authorization again."
            ) from exc

    if not interactive:
        raise GoogleAuthError(
            f"No valid OAuth token is available at {token_path}. "
            "Run google_auth.py --authorize-readonly for first-time browser "
            "authorization."
        )

    try:
        flow = InstalledAppFlow.from_client_secrets_file(
            str(client_path),
            scopes=list(scopes),
        )
        credentials = flow.run_local_server(
            host="localhost",
            port=0,
            open_browser=True,
            authorization_prompt_message=(
                "Complete Google authorization in the opened browser."
            ),
            success_message=(
                "Authorization completed. You may close this browser window."
            ),
        )
    except Exception as exc:
        raise GoogleAuthError(
            "Browser OAuth authorization did not complete successfully."
        ) from exc

    save_token(token_path, credentials, scopes)
    return credentials


def parse_args(argv: Sequence[str] | None = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description=(
            "Manage user-local OAuth for the official Google Sheets API. "
            "Secrets and tokens are never printed."
        )
    )
    action = parser.add_mutually_exclusive_group(required=True)
    action.add_argument(
        "--show-paths",
        action="store_true",
        help="Show proposed local credential paths without authenticating",
    )
    action.add_argument(
        "--authorize-readonly",
        action="store_true",
        help="Run browser OAuth for the Sheets read-only scope",
    )
    action.add_argument(
        "--authorize-readwrite",
        action="store_true",
        help="Run browser OAuth for the Sheets read/write scope",
    )
    return parser.parse_args(argv)


def main(argv: Sequence[str] | None = None) -> int:
    configure_utf8_output()
    args = parse_args(argv)

    try:
        if args.show_paths:
            print(f"OAuth client configuration: {oauth_client_path()}")
            print(
                "Read-only access/refresh token: "
                f"{token_path_for_scopes(READONLY_SCOPES)}"
            )
            print(
                "Future read/write token: "
                f"{token_path_for_scopes(READWRITE_SCOPES)}"
            )
            print(f"Current scope: {SHEETS_READONLY_SCOPE}")
            return 0

        scopes = (
            READWRITE_SCOPES
            if args.authorize_readwrite
            else READONLY_SCOPES
        )
        access_kind = (
            "read/write"
            if args.authorize_readwrite
            else "read-only"
        )
        get_credentials(scopes, interactive=True)
        print(
            f"SUCCESS: Google Sheets {access_kind} authorization is available. "
            "No token or secret value was printed."
        )
        return 0
    except GoogleAuthError as exc:
        print(f"ERROR: {exc}", file=sys.stderr)
        return 2


if __name__ == "__main__":
    raise SystemExit(main())
