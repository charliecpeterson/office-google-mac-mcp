"""OAuth flow, scopes, service-client builders, and active-document storage.

The only module that holds Google credentials — the analog of the Office MCP's
`bridge.py`. Everything else calls the Google APIs through clients built here.

Config (env-overridable) lives at `~/.config/google-mcp/`:
- `client_secret.json` (you provide; downloaded from Google Cloud Console)
- `token.json`         (cached after the first consent)
- `active.json`        (the active spreadsheet/document/presentation id per app)
"""

import json
import os
import re
from pathlib import Path

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build as build_service

_CONFIG_DIR = Path(os.environ.get("GOOGLE_MCP_CONFIG_DIR", "~/.config/google-mcp")).expanduser()
_CLIENT_SECRET = _CONFIG_DIR / "client_secret.json"
_TOKEN = _CONFIG_DIR / "token.json"
_ACTIVE = _CONFIG_DIR / "active.json"

# Request the union of scopes the suite may need, so one consent covers everything
# and adding Docs/Slides later doesn't force a re-consent.
SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/documents",
    "https://www.googleapis.com/auth/presentations",
    "https://www.googleapis.com/auth/drive.file",
]

_URL_ID = re.compile(r"/(?:spreadsheets|document|presentation)/d/([a-zA-Z0-9_-]+)")


def extract_id(url_or_id: str) -> str:
    """Pull the document ID out of a Google docs/sheets/slides URL; pass IDs through."""
    m = _URL_ID.search(url_or_id)
    return m.group(1) if m else url_or_id


def _credentials() -> Credentials:
    """Return valid Credentials, running the consent flow if needed (one-time)."""
    _CONFIG_DIR.mkdir(parents=True, exist_ok=True)
    creds = Credentials.from_authorized_user_file(str(_TOKEN), SCOPES) if _TOKEN.exists() else None
    if creds and creds.valid:
        return creds
    if creds and creds.expired and creds.refresh_token:
        creds.refresh(Request())
    else:
        if not _CLIENT_SECRET.exists():
            raise RuntimeError(
                f"missing OAuth client secret at {_CLIENT_SECRET}. See "
                "packages/google/PLAN.md for the one-time GCP setup."
            )
        flow = InstalledAppFlow.from_client_secrets_file(str(_CLIENT_SECRET), SCOPES)
        creds = flow.run_local_server(port=0)
    _TOKEN.write_text(creds.to_json())
    return creds


def sheets_service():
    return build_service("sheets", "v4", credentials=_credentials(), cache_discovery=False)


def docs_service():
    return build_service("docs", "v1", credentials=_credentials(), cache_discovery=False)


def slides_service():
    return build_service("slides", "v1", credentials=_credentials(), cache_discovery=False)


def drive_service():
    return build_service("drive", "v3", credentials=_credentials(), cache_discovery=False)


def _load_active() -> dict:
    if not _ACTIVE.exists():
        return {}
    try:
        return json.loads(_ACTIVE.read_text())
    except json.JSONDecodeError:
        return {}


def active_id(app: str) -> str | None:
    """The persisted active document ID for an app (sheets / docs / slides), or None."""
    return _load_active().get(app)


def set_active(app: str, url_or_id: str) -> str:
    """Persist the active document for an app. Accepts a Google URL or a raw ID.
    Returns the resolved ID."""
    _CONFIG_DIR.mkdir(parents=True, exist_ok=True)
    state = _load_active()
    doc_id = extract_id(url_or_id)
    state[app] = doc_id
    _ACTIVE.write_text(json.dumps(state, indent=2))
    return doc_id


def require_active(app: str) -> str:
    """Return the active doc id, or raise a clear error pointing to set_active."""
    doc_id = active_id(app)
    if not doc_id:
        raise RuntimeError(
            f"no active {app} document; call {app}_set_active(url) with the URL of the open doc first"
        )
    return doc_id
