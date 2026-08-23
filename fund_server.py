#!/usr/bin/env python3
"""Local server for Fund Selection Tool.

Serves the static app and stores shared draft files under Drafts.
"""

from __future__ import annotations

import json
import os
import re
import sqlite3
from datetime import datetime, timezone
from http import HTTPStatus
from http.server import SimpleHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.parse import parse_qs, unquote, urlencode, urlparse
from urllib.request import Request, urlopen

import pandas as pd


ROOT = Path(__file__).resolve().parent
DRAFTS_DIR = ROOT / "Drafts"
def drive_json_target(target_name: str, legacy_env_name: str, default: str = "") -> str:
    target_env_name = f"DRIVE_JSON_TARGET_{target_name.upper()}"
    return (
        os.environ.get(target_env_name, "").strip()
        or os.environ.get(legacy_env_name, "").strip()
        or default
    )


DRAFTS_DRIVE_FOLDER_ID = drive_json_target(
    "DRAFTS",
    "DRAFTS_DRIVE_FOLDER_ID",
    "1TwC7V8gpcDswftoweT-VcG89OnIVW5Ol",
)
FUND_SELECTION_LOGS_DRIVE_FOLDER_ID = drive_json_target(
    "FUND_SELECTION_LOGS",
    "FUND_SELECTION_LOGS_DRIVE_FOLDER_ID",
    "12ciJQq-dpBr-DpdnzXCOXqtW_ijctJN6",
)
FUND_SELECTION_LOGS_SHEET_ID = os.environ.get(
    "FUND_SELECTION_LOGS_SHEET_ID",
    "1fOdq3JSKTjRZLE8sQ62jn3OmmuKGuhDo2tUCLjy1zIg",
).strip()
FUND_SELECTION_LOGS_SHEET_HEADERS = [
    "quarter",
    "item_order",
    "item_id",
    "asset_class",
    "fund_type",
    "category",
    "role",
    "fund_code",
    "status",
    "reason",
    "tags",
    "data_as_of",
    "item_revision",
    "updated_by",
    "updated_at",
    "mention_id",
    "row_revision",
    "deleted",
]
DRAFT_API_WEB_APP_URL = os.environ.get(
    "DRAFT_API_WEB_APP_URL",
    "https://script.google.com/macros/s/AKfycbwtPbn93YVc3i8KVvPHCA28v0giukB7Ihe59TbLMcjwow0O1PcgaFjY39qa9KwWE3DJ6Q/exec",
).strip()
DRAFT_API_SECRET_KEY = os.environ.get("DRAFT_API_SECRET_KEY", "change-this-draft-api-key").strip()
FUND_MASTER_FILE = ROOT / "Python By Boss เพื่อดึงข้อมูล" / "fund_master_profiles.xlsx"
JSON_DRIVE_ROOT_FOLDER_ID = drive_json_target(
    "DATA",
    "JSON_DRIVE_ROOT_FOLDER_ID",
    "1vUWAU5qP0qiIHPa5C4TZUybVmEwqfl6W",
)
DEFAULT_QUARTER = os.environ.get("FUND_TOOL_DEFAULT_QUARTER", "2026-Q1")
FUND_OVERRIDES_FILE = ROOT / "Data" / "fund_overrides.json"
MASTER_ALLOCATIONS_FILE_NAME = "fund_master_allocations.json"
MASTER_ALLOCATIONS_FILE = ROOT / "Data" / MASTER_ALLOCATIONS_FILE_NAME
FIXED_INCOME_FACTORS_OVERRIDES_FILE = ROOT / "Data" / "fixed_income_factors_overrides.json"
FT_HISTORICAL_PRICES_DB = ROOT / "Data" / "ft_historical_prices" / "ft_historical_prices.sqlite"
MASTER_FUND_JSON_FILE = ROOT / "Data" / "AVP Master Fund ID - 2026-Q1.json"
MAX_BODY_BYTES = 2 * 1024 * 1024


def ft_symbol_base(value: str) -> str:
    return str(value or "").strip().split(":", 1)[0].upper()


def currency_code(value: str) -> str:
    text = str(value or "").strip().lower()
    mapping = {
        "euro": "EUR",
        "eur": "EUR",
        "us dollar": "USD",
        "usd": "USD",
        "u.s. dollar": "USD",
        "pound sterling": "GBP",
        "gbp": "GBP",
        "japanese yen": "JPY",
        "jpy": "JPY",
    }
    return mapping.get(text, str(value or "").strip().upper())


def load_master_fund_name_lookup() -> dict[str, str]:
    if not MASTER_FUND_JSON_FILE.exists():
        return {}
    try:
        data = json.loads(MASTER_FUND_JSON_FILE.read_text(encoding="utf-8"))
    except Exception:
        return {}
    if not isinstance(data, list) or not data:
        return {}
    headers = [str(item).strip() for item in data[0]]

    def idx(name: str) -> int:
        try:
            return headers.index(name)
        except ValueError:
            return -1

    name_idx = idx("Group/Investment")
    isin_idx = idx("ISIN")
    currency_idx = idx("Base Currency")
    lookup: dict[str, str] = {}
    if name_idx < 0 or isin_idx < 0:
        return lookup
    for row in data[1:]:
        if not isinstance(row, list):
            continue
        name = str(row[name_idx] if name_idx < len(row) else "").strip()
        isin = str(row[isin_idx] if isin_idx < len(row) else "").strip().upper()
        curr = currency_code(row[currency_idx] if currency_idx >= 0 and currency_idx < len(row) else "")
        if not name or not isin or isin == "-":
            continue
        lookup.setdefault(isin, name)
        if curr:
            lookup.setdefault(f"{isin}:{curr}", name)
    return lookup


def safe_slug(value: str, fallback: str = "draft") -> str:
    text = re.sub(r"\s+", "-", str(value or "").strip().lower())
    text = re.sub(r"[^a-z0-9ก-๙._-]+", "", text)
    text = text.strip(".-_")
    return text[:80] or fallback


def draft_path(draft_id: str) -> Path:
    clean_id = safe_slug(draft_id, "draft")
    return DRAFTS_DIR / f"{clean_id}.json"


def draft_file_name(draft_id: str) -> str:
    return f"{safe_slug(draft_id, 'draft')}.json"


def normalize_quarter(value: str) -> str:
    text = str(value or "").strip().upper()
    match = re.match(r"^(\d{4})-Q([1-4])$", text)
    if not match:
        raise ValueError("quarter must be in YYYY-QN format, e.g. 2026-Q1")
    return f"{match.group(1)}-Q{match.group(2)}"


def quarter_year(quarter: str) -> str:
    return normalize_quarter(quarter).split("-", 1)[0]


def json_override_folder_segments(quarter: str) -> list[str]:
    normalized = normalize_quarter(quarter)
    return [quarter_year(normalized), normalized, "overrides"]


def master_allocations_path(quarter: str = DEFAULT_QUARTER) -> Path:
    normalized = normalize_quarter(quarter)
    return ROOT / "Data" / quarter_year(normalized) / normalized / "overrides" / MASTER_ALLOCATIONS_FILE_NAME


def fund_selection_log_file_name(quarter: str) -> str:
    return f"Fund Selection Logs - {normalize_quarter(quarter)}.json"


def fund_selection_log_path(quarter: str) -> Path:
    return ROOT / "Data" / fund_selection_log_file_name(quarter)


def read_json_file(path: Path) -> dict:
    with path.open("r", encoding="utf-8") as fh:
        data = json.load(fh)
    return data if isinstance(data, dict) else {}


def call_draft_web_app(action: str, payload: dict | None = None, method: str = "GET") -> dict:
    if not DRAFT_API_WEB_APP_URL:
        raise RuntimeError("DRAFT_API_WEB_APP_URL is not configured")

    request_payload = {
        "key": DRAFT_API_SECRET_KEY,
        "action": action,
    }
    if payload:
        request_payload.update(payload)

    if method.upper() == "POST":
        req = Request(
            DRAFT_API_WEB_APP_URL,
            data=json.dumps(request_payload, ensure_ascii=False, separators=(",", ":")).encode("utf-8"),
            headers={"Content-Type": "text/plain;charset=utf-8"},
            method="POST",
        )
    else:
        params = {
            "key": DRAFT_API_SECRET_KEY,
            "action": action,
        }
        if payload:
            params["payload"] = json.dumps(payload, ensure_ascii=False, separators=(",", ":"))
        req = f"{DRAFT_API_WEB_APP_URL}?{urlencode(params)}"

    with urlopen(req, timeout=30) as resp:
        body = resp.read().decode("utf-8")
    data = json.loads(body)
    if isinstance(data, dict) and data.get("ok") is False:
        raise RuntimeError(data.get("error") or f"Draft API {action} failed")
    return data if isinstance(data, dict) else {"ok": False, "error": "Draft API returned invalid JSON"}


def upload_draft_to_drive(draft_id: str, draft: dict) -> dict:
    if not DRAFTS_DRIVE_FOLDER_ID:
        return {"ok": False, "skipped": True, "error": "DRAFTS_DRIVE_FOLDER_ID is not set"}

    try:
        from scripts.drive_json_store import upload_json_payload

        file_id = upload_json_payload(DRAFTS_DRIVE_FOLDER_ID, draft_file_name(draft_id), draft)
        return {
            "ok": True,
            "fileId": file_id,
            "folderId": DRAFTS_DRIVE_FOLDER_ID,
            "fileName": draft_file_name(draft_id),
        }
    except ModuleNotFoundError as exc:
        missing = exc.name or "required Google API package"
        return {
            "ok": False,
            "error": f"Drive upload requires Python package '{missing}'. Install scripts/requirements-json-export.txt.",
        }
    except Exception as exc:
        return {"ok": False, "error": str(exc)}


def delete_draft_from_drive(draft_id: str) -> dict:
    if not DRAFTS_DRIVE_FOLDER_ID:
        return {"ok": False, "skipped": True, "error": "DRAFTS_DRIVE_FOLDER_ID is not set"}

    try:
        from scripts.drive_json_store import delete_json_file

        removed = delete_json_file(DRAFTS_DRIVE_FOLDER_ID, draft_file_name(draft_id))
        return {
            "ok": True,
            "removed": removed,
            "folderId": DRAFTS_DRIVE_FOLDER_ID,
            "fileName": draft_file_name(draft_id),
        }
    except ModuleNotFoundError as exc:
        missing = exc.name or "required Google API package"
        return {
            "ok": False,
            "error": f"Drive delete requires Python package '{missing}'. Install scripts/requirements-json-export.txt.",
        }
    except Exception as exc:
        return {"ok": False, "error": str(exc)}


def read_fund_selection_log_from_drive(quarter: str) -> dict:
    if not FUND_SELECTION_LOGS_DRIVE_FOLDER_ID:
        return {"ok": False, "skipped": True, "error": "FUND_SELECTION_LOGS_DRIVE_FOLDER_ID is not set"}
    file_name = fund_selection_log_file_name(quarter)
    try:
        from scripts.drive_json_store import download_json_payload

        payload = download_json_payload(FUND_SELECTION_LOGS_DRIVE_FOLDER_ID, file_name)
        if payload is None:
            return {
                "ok": False,
                "notFound": True,
                "folderId": FUND_SELECTION_LOGS_DRIVE_FOLDER_ID,
                "fileName": file_name,
            }
        return {
            "ok": True,
            "payload": payload,
            "folderId": FUND_SELECTION_LOGS_DRIVE_FOLDER_ID,
            "fileName": file_name,
        }
    except ModuleNotFoundError as exc:
        missing = exc.name or "required Google API package"
        return {
            "ok": False,
            "error": f"Drive read requires Python package '{missing}'. Install scripts/requirements-json-export.txt.",
            "fileName": file_name,
        }
    except Exception as exc:
        return {"ok": False, "error": str(exc), "fileName": file_name}


def upload_fund_selection_log_to_drive(quarter: str, payload: dict) -> dict:
    if not FUND_SELECTION_LOGS_DRIVE_FOLDER_ID:
        return {"ok": False, "skipped": True, "error": "FUND_SELECTION_LOGS_DRIVE_FOLDER_ID is not set"}
    file_name = fund_selection_log_file_name(quarter)
    try:
        from scripts.drive_json_store import upload_json_payload

        file_id = upload_json_payload(FUND_SELECTION_LOGS_DRIVE_FOLDER_ID, file_name, payload)
        return {
            "ok": True,
            "fileId": file_id,
            "folderId": FUND_SELECTION_LOGS_DRIVE_FOLDER_ID,
            "fileName": file_name,
        }
    except ModuleNotFoundError as exc:
        missing = exc.name or "required Google API package"
        return {
            "ok": False,
            "error": f"Drive upload requires Python package '{missing}'. Install scripts/requirements-json-export.txt.",
            "fileName": file_name,
        }
    except Exception as exc:
        return {"ok": False, "error": str(exc), "fileName": file_name}


def upload_master_allocations_to_drive(quarter: str, payload: dict) -> dict:
    if not JSON_DRIVE_ROOT_FOLDER_ID:
        return {
            "ok": False,
            "skipped": True,
            "error": "DRIVE_JSON_TARGET_DATA or JSON_DRIVE_ROOT_FOLDER_ID is not set",
        }

    file_name = MASTER_ALLOCATIONS_FILE_NAME
    path_segments = json_override_folder_segments(quarter)
    try:
        from scripts.drive_json_store import upload_json_payload_to_path

        file_id, folder_id = upload_json_payload_to_path(
            JSON_DRIVE_ROOT_FOLDER_ID,
            path_segments,
            file_name,
            payload,
        )
        return {
            "ok": True,
            "fileId": file_id,
            "folderId": folder_id,
            "fileName": file_name,
            "path": "/".join([*path_segments, file_name]),
        }
    except ModuleNotFoundError as exc:
        missing = exc.name or "required Google API package"
        return {
            "ok": False,
            "error": f"Drive upload requires Python package '{missing}'. Install scripts/requirements-json-export.txt.",
            "fileName": file_name,
        }
    except Exception as exc:
        return {"ok": False, "error": str(exc), "fileName": file_name}


def spreadsheet_credentials_from_env():
    from google.oauth2 import service_account

    scopes = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive",
    ]
    raw_json = (
        os.environ.get("GOOGLE_SERVICE_ACCOUNT_JSON_EXPORT", "").strip()
        or os.environ.get("GOOGLE_SERVICE_ACCOUNT_JSON", "").strip()
    )
    if raw_json:
        return service_account.Credentials.from_service_account_info(
            json.loads(raw_json),
            scopes=scopes,
        )

    credentials_path = os.environ.get("GOOGLE_APPLICATION_CREDENTIALS", "").strip()
    if credentials_path:
        return service_account.Credentials.from_service_account_file(
            credentials_path,
            scopes=scopes,
        )

    raise RuntimeError(
        "Set GOOGLE_SERVICE_ACCOUNT_JSON_EXPORT, GOOGLE_SERVICE_ACCOUNT_JSON, "
        "or GOOGLE_APPLICATION_CREDENTIALS first."
    )


def spreadsheet_client():
    from googleapiclient.discovery import build

    return build(
        "sheets",
        "v4",
        credentials=spreadsheet_credentials_from_env(),
        cache_discovery=False,
    )


def sheet_column_name(index: int) -> str:
    name = ""
    while index:
        index, remainder = divmod(index - 1, 26)
        name = chr(65 + remainder) + name
    return name


def fund_selection_sheet_titles(sheets, spreadsheet_id: str) -> set[str]:
    response = (
        sheets.spreadsheets()
        .get(spreadsheetId=spreadsheet_id, fields="sheets(properties(title))")
        .execute()
    )
    return {
        row.get("properties", {}).get("title", "")
        for row in response.get("sheets", [])
    }


def ensure_fund_selection_sheet_tab(sheets, spreadsheet_id: str, quarter: str) -> bool:
    if quarter in fund_selection_sheet_titles(sheets, spreadsheet_id):
        return False
    sheets.spreadsheets().batchUpdate(
        spreadsheetId=spreadsheet_id,
        body={"requests": [{"addSheet": {"properties": {"title": quarter}}}]},
    ).execute()
    return True


def fund_selection_log_to_sheet_rows(log: dict, quarter: str) -> list[list[str]]:
    rows = [FUND_SELECTION_LOGS_SHEET_HEADERS]
    data_as_of = str(log.get("dataAsOf") or "").strip()
    updated_by = str(log.get("updatedBy") or "").strip()
    updated_at = str(log.get("updatedAt") or datetime.now(tz=timezone.utc).isoformat()).strip()

    for item_index, item in enumerate(log.get("items") or [], start=1):
        if not isinstance(item, dict):
            continue
        mentions = item.get("mentions") if isinstance(item.get("mentions"), list) else []
        if not mentions:
            mentions = [{}]
        for mention in mentions:
            mention = mention if isinstance(mention, dict) else {}
            rows.append(
                [
                    quarter,
                    str(item_index),
                    str(item.get("id") or "").strip(),
                    str(item.get("assetClass") or "").strip(),
                    str(item.get("fundType") or "").strip(),
                    str(item.get("category") or "").strip(),
                    str(mention.get("role") or "").strip(),
                    str(mention.get("fundCode") or "").strip().upper(),
                    str(mention.get("status") or "").strip(),
                    str(mention.get("reason") or "").strip(),
                    ", ".join(
                        str(tag or "").strip()
                        for tag in (mention.get("tags") if isinstance(mention.get("tags"), list) else [])
                        if str(tag or "").strip()
                    ),
                    data_as_of,
                    str(item.get("itemRevision") or "1").strip(),
                    str(mention.get("updatedBy") or item.get("updatedBy") or updated_by).strip(),
                    str(mention.get("updatedAt") or item.get("updatedAt") or updated_at).strip(),
                    str(mention.get("id") or "").strip(),
                    str(mention.get("rowRevision") or "1").strip(),
                    "FALSE",
                ]
            )
    return rows


def write_fund_selection_log_to_sheet(quarter: str, log: dict) -> dict:
    if not FUND_SELECTION_LOGS_SHEET_ID:
        return {"ok": False, "error": "FUND_SELECTION_LOGS_SHEET_ID is not set"}

    sheets = spreadsheet_client()
    created_tab = ensure_fund_selection_sheet_tab(sheets, FUND_SELECTION_LOGS_SHEET_ID, quarter)
    rows = fund_selection_log_to_sheet_rows(log, quarter)
    last_column = sheet_column_name(len(FUND_SELECTION_LOGS_SHEET_HEADERS))
    sheet_range = f"'{quarter}'!A:{last_column}"

    sheets.spreadsheets().values().clear(
        spreadsheetId=FUND_SELECTION_LOGS_SHEET_ID,
        range=sheet_range,
        body={},
    ).execute()
    response = (
        sheets.spreadsheets()
        .values()
        .update(
            spreadsheetId=FUND_SELECTION_LOGS_SHEET_ID,
            range=f"'{quarter}'!A1",
            valueInputOption="RAW",
            body={"values": rows},
        )
        .execute()
    )

    return {
        "ok": True,
        "spreadsheetId": FUND_SELECTION_LOGS_SHEET_ID,
        "quarter": quarter,
        "createdTab": created_tab,
        "updatedRows": response.get("updatedRows", len(rows)),
        "updatedColumns": response.get("updatedColumns", len(FUND_SELECTION_LOGS_SHEET_HEADERS)),
        "dataRows": max(0, len(rows) - 1),
    }


def normalize_override_key(value: str) -> str:
    return str(value or "").strip().upper()


def read_fund_overrides() -> dict:
    if not FUND_OVERRIDES_FILE.exists():
        return {"items": {}, "updatedAt": None}
    data = read_json_file(FUND_OVERRIDES_FILE)
    items = data.get("items") if isinstance(data.get("items"), dict) else {}
    return {
        "items": items,
        "updatedAt": data.get("updatedAt"),
    }


def write_fund_overrides(data: dict) -> None:
    FUND_OVERRIDES_FILE.parent.mkdir(parents=True, exist_ok=True)
    tmp_path = FUND_OVERRIDES_FILE.with_suffix(".tmp")
    with tmp_path.open("w", encoding="utf-8") as fh:
        json.dump(data, fh, ensure_ascii=False, indent=2)
        fh.write("\n")
    tmp_path.replace(FUND_OVERRIDES_FILE)


def read_master_allocations(quarter: str = DEFAULT_QUARTER) -> dict:
    path = master_allocations_path(quarter)
    source_path = path
    if not path.exists() and MASTER_ALLOCATIONS_FILE.exists():
        source_path = MASTER_ALLOCATIONS_FILE
    if not source_path.exists():
        return {"items": {}, "updatedAt": None, "sourcePath": path}
    data = read_json_file(source_path)
    items = data.get("items") if isinstance(data.get("items"), dict) else {}
    return {
        "items": items,
        "updatedAt": data.get("updatedAt"),
        "sourcePath": source_path,
    }


def write_master_allocations(data: dict, quarter: str = DEFAULT_QUARTER) -> Path:
    path = master_allocations_path(quarter)
    path.parent.mkdir(parents=True, exist_ok=True)
    tmp_path = path.with_suffix(".tmp")
    with tmp_path.open("w", encoding="utf-8") as fh:
        json.dump(data, fh, ensure_ascii=False, indent=2)
        fh.write("\n")
    tmp_path.replace(path)

    # Keep the legacy flat file as a local compatibility cache for older static flows.
    MASTER_ALLOCATIONS_FILE.parent.mkdir(parents=True, exist_ok=True)
    legacy_tmp_path = MASTER_ALLOCATIONS_FILE.with_suffix(".tmp")
    with legacy_tmp_path.open("w", encoding="utf-8") as fh:
        json.dump(data, fh, ensure_ascii=False, indent=2)
        fh.write("\n")
    legacy_tmp_path.replace(MASTER_ALLOCATIONS_FILE)
    return path


def read_fixed_income_factors_overrides() -> dict:
    if not FIXED_INCOME_FACTORS_OVERRIDES_FILE.exists():
        return {"items": {}, "updatedAt": None}
    data = read_json_file(FIXED_INCOME_FACTORS_OVERRIDES_FILE)
    items = data.get("items") if isinstance(data.get("items"), dict) else {}
    return {
        "items": items,
        "updatedAt": data.get("updatedAt"),
    }


def write_fixed_income_factors_overrides(data: dict) -> None:
    FIXED_INCOME_FACTORS_OVERRIDES_FILE.parent.mkdir(parents=True, exist_ok=True)
    tmp_path = FIXED_INCOME_FACTORS_OVERRIDES_FILE.with_suffix(".tmp")
    with tmp_path.open("w", encoding="utf-8") as fh:
        json.dump(data, fh, ensure_ascii=False, indent=2)
        fh.write("\n")
    tmp_path.replace(FIXED_INCOME_FACTORS_OVERRIDES_FILE)


class FundRequestHandler(SimpleHTTPRequestHandler):
    server_version = "FundSelectionServer/1.0"

    def __init__(self, *args, **kwargs):
        super().__init__(*args, directory=str(ROOT), **kwargs)

    def end_headers(self):
        self.send_header("Cache-Control", "no-store")
        super().end_headers()

    def send_json(self, status: int, payload: dict):
        body = json.dumps(payload, ensure_ascii=False, indent=2).encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    def read_request_json(self) -> dict:
        length = int(self.headers.get("Content-Length") or "0")
        if length > MAX_BODY_BYTES:
            raise ValueError("Request body is too large")
        raw = self.rfile.read(length)
        if not raw:
            return {}
        data = json.loads(raw.decode("utf-8"))
        if not isinstance(data, dict):
            raise ValueError("JSON body must be an object")
        return data

    def do_GET(self):
        parsed = urlparse(self.path)
        if parsed.path == "/api/draft-drive":
            return self.handle_draft_drive_get(parsed)
        if parsed.path == "/api/drafts":
            return self.handle_list_drafts()
        if parsed.path == "/api/fund-overrides":
            return self.handle_get_fund_overrides()
        if parsed.path == "/api/master-allocations":
            return self.handle_get_master_allocations(parsed)
        if parsed.path == "/api/fixed-income-factors-overrides":
            return self.handle_get_fixed_income_factors_overrides()
        if parsed.path == "/api/fund-selection-logs":
            return self.handle_get_fund_selection_log(parsed)
        if parsed.path == "/api/fund-master":
            return self.handle_fund_master(parsed)
        if parsed.path == "/api/dividend-check":
            return self.handle_dividend_check(parsed)
        if parsed.path == "/api/ft-historical-prices":
            return self.handle_get_ft_historical_prices(parsed)
        if parsed.path == "/api/ft-price-on-date":
            return self.handle_get_ft_price_on_date(parsed)
        return super().do_GET()

    def do_POST(self):
        parsed = urlparse(self.path)
        if parsed.path == "/api/draft-drive":
            return self.handle_draft_drive_save()
        if parsed.path == "/api/drafts":
            return self.handle_save_draft()
        if parsed.path == "/api/fund-overrides":
            return self.handle_save_fund_override()
        if parsed.path == "/api/master-allocations":
            return self.handle_save_master_allocation()
        if parsed.path == "/api/fixed-income-factors-overrides":
            return self.handle_save_fixed_income_factors_overrides()
        if parsed.path == "/api/fund-selection-logs":
            return self.handle_save_fund_selection_log()
        if parsed.path == "/api/fund-selection-logs-sheet":
            return self.handle_save_fund_selection_log_to_sheet()
        if parsed.path == "/api/ft-historical-prices":
            return self.handle_ft_historical_prices_sync()
        if parsed.path == "/api/ft-qualitative-data":
            return self.handle_ft_qualitative_data_sync()
        self.send_error(HTTPStatus.NOT_FOUND, "API endpoint not found")

    def do_DELETE(self):
        parsed = urlparse(self.path)
        if parsed.path.startswith("/api/draft-drive/"):
            draft_id = unquote(parsed.path.rsplit("/", 1)[-1])
            return self.handle_draft_drive_delete(draft_id, parsed)
        if parsed.path.startswith("/api/drafts/"):
            draft_id = unquote(parsed.path.rsplit("/", 1)[-1])
            return self.handle_delete_draft(draft_id)
        if parsed.path.startswith("/api/fund-overrides/"):
            fund_key = unquote(parsed.path.rsplit("/", 1)[-1])
            return self.handle_delete_fund_override(fund_key)
        if parsed.path.startswith("/api/master-allocations/"):
            fund_key = unquote(parsed.path.rsplit("/", 1)[-1])
            return self.handle_delete_master_allocation(fund_key)
        if parsed.path.startswith("/api/fixed-income-factors-overrides/"):
            fund_key = unquote(parsed.path.rsplit("/", 1)[-1])
            return self.handle_delete_fixed_income_factors_override(fund_key)
        self.send_error(HTTPStatus.NOT_FOUND, "API endpoint not found")

    def handle_list_drafts(self):
        DRAFTS_DIR.mkdir(parents=True, exist_ok=True)
        drafts = []
        for path in sorted(DRAFTS_DIR.glob("*.json"), reverse=True):
            try:
                drafts.append(read_json_file(path))
            except Exception:
                continue
        drafts.sort(key=lambda d: d.get("createdAt") or d.get("updatedAt") or "", reverse=True)
        self.send_json(HTTPStatus.OK, {"drafts": drafts})

    def handle_draft_drive_get(self, parsed):
        try:
            query = parse_qs(parsed.query)
            quarter = (query.get("quarter") or [""])[0].strip()
            data = call_draft_web_app("list", {"quarter": quarter} if quarter else None)
            self.send_json(HTTPStatus.OK, data)
        except Exception as exc:
            self.send_json(HTTPStatus.BAD_GATEWAY, {"ok": False, "error": str(exc)})

    def handle_draft_drive_save(self):
        try:
            draft = self.read_request_json()
            data = call_draft_web_app("save", {"draft": draft}, method="POST")
            self.send_json(HTTPStatus.OK, data)
        except Exception as exc:
            self.send_json(HTTPStatus.BAD_GATEWAY, {"ok": False, "error": str(exc)})

    def handle_draft_drive_delete(self, draft_id: str, parsed):
        try:
            query = parse_qs(parsed.query)
            quarter = (query.get("quarter") or [""])[0].strip()
            data = call_draft_web_app("delete", {"id": draft_id, "quarter": quarter}, method="POST")
            self.send_json(HTTPStatus.OK, data)
        except Exception as exc:
            self.send_json(HTTPStatus.BAD_GATEWAY, {"ok": False, "error": str(exc)})

    def handle_get_fund_overrides(self):
        data = read_fund_overrides()
        self.send_json(
            HTTPStatus.OK,
            {
                "ok": True,
                "items": data["items"],
                "updatedAt": data["updatedAt"],
                "source": str(FUND_OVERRIDES_FILE.relative_to(ROOT)),
            },
        )

    def handle_get_master_allocations(self, parsed):
        query = parse_qs(parsed.query)
        quarter = normalize_quarter((query.get("quarter") or [""])[0] or DEFAULT_QUARTER)
        data = read_master_allocations(quarter)
        source_path = data.get("sourcePath") or master_allocations_path(quarter)
        self.send_json(
            HTTPStatus.OK,
            {
                "ok": True,
                "items": data["items"],
                "updatedAt": data["updatedAt"],
                "quarter": quarter,
                "source": str(source_path.relative_to(ROOT)),
            },
        )

    def handle_get_fixed_income_factors_overrides(self):
        data = read_fixed_income_factors_overrides()
        self.send_json(
            HTTPStatus.OK,
            {
                "ok": True,
                "items": data["items"],
                "updatedAt": data["updatedAt"],
                "source": str(FIXED_INCOME_FACTORS_OVERRIDES_FILE.relative_to(ROOT)),
            },
        )

    def handle_get_fund_selection_log(self, parsed):
        try:
            query = parse_qs(parsed.query)
            quarter = normalize_quarter((query.get("quarter") or [""])[0] or "2026-Q1")
            path = fund_selection_log_path(quarter)
            drive_result = read_fund_selection_log_from_drive(quarter)

            if drive_result.get("ok") and isinstance(drive_result.get("payload"), dict):
                payload = drive_result["payload"]
                source = f"Google Drive: {drive_result.get('fileName')}"
                local_warning = None
                try:
                    path.parent.mkdir(parents=True, exist_ok=True)
                    tmp_path = path.with_suffix(".tmp")
                    with tmp_path.open("w", encoding="utf-8") as fh:
                        json.dump(payload, fh, ensure_ascii=False, indent=2)
                        fh.write("\n")
                    tmp_path.replace(path)
                except Exception as exc:
                    local_warning = f"Drive loaded, but local cache write failed: {exc}"
                return self.send_json(
                    HTTPStatus.OK,
                    {
                        "ok": True,
                        "log": payload,
                        "source": source,
                        "path": str(path.relative_to(ROOT)),
                        "drive": drive_result,
                        "warning": local_warning,
                    },
                )

            if path.exists():
                payload = read_json_file(path)
                return self.send_json(
                    HTTPStatus.OK,
                    {
                        "ok": True,
                        "log": payload,
                        "source": f"Local JSON: {path.name}",
                        "path": str(path.relative_to(ROOT)),
                        "drive": drive_result,
                        "warning": None if drive_result.get("notFound") else drive_result.get("error"),
                    },
                )

            return self.send_json(
                HTTPStatus.OK,
                {
                    "ok": True,
                    "log": None,
                    "source": "New file",
                    "path": str(path.relative_to(ROOT)),
                    "drive": drive_result,
                    "warning": None if drive_result.get("notFound") else drive_result.get("error"),
                },
            )
        except Exception as exc:
            self.send_json(HTTPStatus.BAD_REQUEST, {"ok": False, "error": str(exc)})

    def handle_save_fund_selection_log(self):
        try:
            payload = self.read_request_json()
            log = payload.get("log") if isinstance(payload.get("log"), dict) else payload
            quarter = normalize_quarter(log.get("quarter") or payload.get("quarter") or "2026-Q1")
            path = fund_selection_log_path(quarter)
            base_revision = payload.get("baseRevision")
            current = read_json_file(path) if path.exists() else {}
            current_revision = current.get("revision")
            if (
                base_revision is not None
                and current_revision is not None
                and int(base_revision) != int(current_revision)
            ):
                return self.send_json(
                    HTTPStatus.CONFLICT,
                    {
                        "ok": False,
                        "error": "revision conflict",
                        "currentRevision": current_revision,
                        "currentLog": current,
                    },
                )

            now = datetime.now(tz=timezone.utc).isoformat()
            next_revision = int(current_revision or log.get("revision") or 0) + 1
            log = {
                **log,
                "schemaVersion": int(log.get("schemaVersion") or 1),
                "quarter": quarter,
                "revision": next_revision,
                "updatedAt": now,
                "updatedBy": str(log.get("updatedBy") or payload.get("updatedBy") or "").strip(),
            }
            if not log.get("createdAt"):
                log["createdAt"] = current.get("createdAt") or now

            path.parent.mkdir(parents=True, exist_ok=True)
            tmp_path = path.with_suffix(".tmp")
            with tmp_path.open("w", encoding="utf-8") as fh:
                json.dump(log, fh, ensure_ascii=False, indent=2)
                fh.write("\n")
            tmp_path.replace(path)

            drive_result = upload_fund_selection_log_to_drive(quarter, log)
            self.send_json(
                HTTPStatus.OK,
                {
                    "ok": True,
                    "log": log,
                    "path": str(path.relative_to(ROOT)),
                    "drive": drive_result,
                    "driveUploaded": bool(drive_result.get("ok")),
                    "driveFileId": drive_result.get("fileId"),
                    "driveFolderId": drive_result.get("folderId") or FUND_SELECTION_LOGS_DRIVE_FOLDER_ID,
                    "warning": None if drive_result.get("ok") else f"บันทึกลงไฟล์ในเครื่องแล้ว แต่ยัง sync เข้า Google Drive ไม่ได้: {drive_result.get('error')}",
                },
            )
        except ValueError as exc:
            self.send_json(HTTPStatus.BAD_REQUEST, {"ok": False, "error": str(exc)})
        except Exception as exc:
            self.send_json(HTTPStatus.BAD_REQUEST, {"ok": False, "error": str(exc)})

    def handle_save_fund_selection_log_to_sheet(self):
        try:
            payload = self.read_request_json()
            log = payload.get("log") if isinstance(payload.get("log"), dict) else payload
            quarter = normalize_quarter(log.get("quarter") or payload.get("quarter") or "2026-Q1")
            now = datetime.now(tz=timezone.utc).isoformat()
            log = {
                **log,
                "schemaVersion": int(log.get("schemaVersion") or 1),
                "quarter": quarter,
                "updatedAt": now,
                "updatedBy": str(log.get("updatedBy") or payload.get("updatedBy") or "").strip(),
            }
            sheet_result = write_fund_selection_log_to_sheet(quarter, log)
            self.send_json(
                HTTPStatus.OK if sheet_result.get("ok") else HTTPStatus.BAD_GATEWAY,
                {
                    "ok": bool(sheet_result.get("ok")),
                    "log": log,
                    "sheet": sheet_result,
                    "warning": None if sheet_result.get("ok") else sheet_result.get("error"),
                },
            )
        except ValueError as exc:
            self.send_json(HTTPStatus.BAD_REQUEST, {"ok": False, "error": str(exc)})
        except ModuleNotFoundError as exc:
            missing = exc.name or "required Google API package"
            self.send_json(
                HTTPStatus.BAD_GATEWAY,
                {
                    "ok": False,
                    "error": f"Google Sheet save requires Python package '{missing}'. Install scripts/requirements-json-export.txt.",
                },
            )
        except Exception as exc:
            self.send_json(HTTPStatus.BAD_REQUEST, {"ok": False, "error": str(exc)})

    def handle_save_fund_override(self):
        try:
            payload = self.read_request_json()
            fund = payload.get("fund") if isinstance(payload.get("fund"), dict) else payload
            code = normalize_override_key(fund.get("code") or fund.get("key"))
            if not code:
                return self.send_json(HTTPStatus.BAD_REQUEST, {"ok": False, "error": "fund code is required"})

            allowed_fields = {
                "code",
                "name",
                "category",
                "type",
                "dividend",
                "style",
                "masterId",
                "masterName",
                "assetHouse",
                "note",
                "status",
                "mode",
            }
            now = datetime.now(tz=timezone.utc).isoformat()
            cleaned = {
                key: str(value).strip()
                for key, value in fund.items()
                if key in allowed_fields and value is not None
            }
            cleaned["code"] = code
            cleaned["key"] = code
            cleaned["updatedAt"] = now

            data = read_fund_overrides()
            items = data["items"]
            existing = items.get(code, {})
            created_at = existing.get("createdAt") or now
            items[code] = {
                **existing,
                **cleaned,
                "createdAt": created_at,
                "updatedAt": now,
            }
            data["updatedAt"] = now
            write_fund_overrides(data)
            self.send_json(HTTPStatus.OK, {"ok": True, "fund": items[code], "source": str(FUND_OVERRIDES_FILE.relative_to(ROOT))})
        except Exception as exc:
            self.send_json(HTTPStatus.BAD_REQUEST, {"ok": False, "error": str(exc)})

    def handle_delete_fund_override(self, fund_key: str):
        key = normalize_override_key(fund_key)
        data = read_fund_overrides()
        removed = data["items"].pop(key, None)
        data["updatedAt"] = datetime.now(tz=timezone.utc).isoformat()
        write_fund_overrides(data)
        self.send_json(HTTPStatus.OK, {"ok": True, "removed": bool(removed), "key": key})

    def handle_save_fixed_income_factors_overrides(self):
        try:
            payload = self.read_request_json()
            raw_items = payload.get("items") if isinstance(payload.get("items"), dict) else None
            replace_items = bool(payload.get("replace"))
            if raw_items is None:
                raw_item = payload.get("item") if isinstance(payload.get("item"), dict) else payload
                key = normalize_override_key(raw_item.get("code") or raw_item.get("key"))
                if not key:
                    return self.send_json(HTTPStatus.BAD_REQUEST, {"ok": False, "error": "fund code is required"})
                raw_items = {key: raw_item}

            allowed_fields = {
                "cash",
                "bond",
                "sector2",
                "foreign",
                "aaa",
                "aa",
                "a",
                "bbb",
                "duration",
                "ytm",
                "holdings",
                "top10",
                "fundSize",
                "maxDd3y",
                "sd3y",
                "maxDd5y",
                "sd5y",
                "note",
            }
            now = datetime.now(tz=timezone.utc).isoformat()
            data = read_fixed_income_factors_overrides()
            changed = {}
            for raw_key, raw_item in raw_items.items():
                if not isinstance(raw_item, dict):
                    continue
                code = normalize_override_key(raw_item.get("code") or raw_item.get("key") or raw_key)
                if not code:
                    continue
                cleaned = {
                    field: str(raw_item.get(field) or "").strip()
                    for field in allowed_fields
                    if field in raw_item
                }
                if replace_items and not cleaned:
                    removed = data["items"].pop(code, None)
                    if removed:
                        changed[code] = {"key": code, "code": code, "removed": True}
                    continue
                existing = data["items"].get(code, {})
                created_at = existing.get("createdAt") or now
                if replace_items:
                    data["items"][code] = {
                        "key": code,
                        "code": code,
                        **cleaned,
                        "createdAt": created_at,
                        "updatedAt": now,
                    }
                else:
                    data["items"][code] = {
                        **existing,
                        **cleaned,
                        "key": code,
                        "code": code,
                        "createdAt": created_at,
                        "updatedAt": now,
                    }
                changed[code] = data["items"][code]

            if not changed:
                return self.send_json(HTTPStatus.BAD_REQUEST, {"ok": False, "error": "no valid override items"})

            data["updatedAt"] = now
            write_fixed_income_factors_overrides(data)
            self.send_json(
                HTTPStatus.OK,
                {
                    "ok": True,
                    "items": changed,
                    "updatedAt": data["updatedAt"],
                    "source": str(FIXED_INCOME_FACTORS_OVERRIDES_FILE.relative_to(ROOT)),
                },
            )
        except Exception as exc:
            self.send_json(HTTPStatus.BAD_REQUEST, {"ok": False, "error": str(exc)})

    def handle_delete_fixed_income_factors_override(self, fund_key: str):
        key = normalize_override_key(fund_key)
        data = read_fixed_income_factors_overrides()
        removed = data["items"].pop(key, None)
        data["updatedAt"] = datetime.now(tz=timezone.utc).isoformat()
        write_fixed_income_factors_overrides(data)
        self.send_json(HTTPStatus.OK, {"ok": True, "removed": bool(removed), "key": key})

    def handle_save_master_allocation(self):
        try:
            payload = self.read_request_json()
            quarter = normalize_quarter(payload.get("quarter") or DEFAULT_QUARTER)
            thai_fund_code = normalize_override_key(payload.get("thaiFundCode") or payload.get("key"))
            if not thai_fund_code:
                return self.send_json(HTTPStatus.BAD_REQUEST, {"ok": False, "error": "thaiFundCode is required"})

            raw_allocations = payload.get("allocations")
            if not isinstance(raw_allocations, list) or not raw_allocations:
                return self.send_json(HTTPStatus.BAD_REQUEST, {"ok": False, "error": "at least one allocation is required"})

            allocations = []
            for idx, item in enumerate(raw_allocations, start=1):
                if not isinstance(item, dict):
                    continue
                master_id = str(item.get("masterId") or "").strip()
                master_name = str(item.get("masterName") or "").strip()
                try:
                    weight = float(item.get("weight"))
                except (TypeError, ValueError):
                    return self.send_json(HTTPStatus.BAD_REQUEST, {"ok": False, "error": f"allocation #{idx} has invalid weight"})
                if not master_id and not master_name:
                    return self.send_json(HTTPStatus.BAD_REQUEST, {"ok": False, "error": f"allocation #{idx} requires master fund"})
                if weight <= 0 or weight > 100:
                    return self.send_json(HTTPStatus.BAD_REQUEST, {"ok": False, "error": f"allocation #{idx} weight must be 0-100"})
                raw_returns = item.get("returns") if isinstance(item.get("returns"), dict) else {}
                returns = {}
                for key in ("r3m", "r6m", "rytd", "r1y", "r3y", "r5y", "r10y"):
                    value = str(raw_returns.get(key) or "").strip()
                    if value:
                        returns[key] = value
                allocations.append({
                    "masterId": master_id,
                    "masterName": master_name,
                    "baseCurrency": str(item.get("baseCurrency") or "").strip(),
                    "weight": round(weight, 4),
                    "returns": returns,
                })

            total_weight = round(sum(item["weight"] for item in allocations), 4)
            if abs(total_weight - 100) > 0.01:
                return self.send_json(HTTPStatus.BAD_REQUEST, {"ok": False, "error": f"total weight must be 100%, currently {total_weight:g}%"})

            now = datetime.now(tz=timezone.utc).isoformat()
            data = read_master_allocations(quarter)
            existing = data["items"].get(thai_fund_code, {})
            data["items"][thai_fund_code] = {
                "key": thai_fund_code,
                "thaiFundCode": thai_fund_code,
                "thaiFundName": str(payload.get("thaiFundName") or "").strip(),
                "allocations": allocations,
                "note": str(payload.get("note") or "").strip(),
                "sourceDate": str(payload.get("sourceDate") or "").strip(),
                "status": str(payload.get("status") or "Active").strip(),
                "createdAt": existing.get("createdAt") or now,
                "updatedAt": now,
            }
            data["updatedAt"] = now
            data.pop("sourcePath", None)
            path = write_master_allocations(data, quarter)
            drive_result = upload_master_allocations_to_drive(quarter, data)
            self.send_json(
                HTTPStatus.OK,
                {
                    "ok": True,
                    "item": data["items"][thai_fund_code],
                    "quarter": quarter,
                    "source": str(path.relative_to(ROOT)),
                    "drive": drive_result,
                    "driveUploaded": bool(drive_result.get("ok")),
                    "warning": None if drive_result.get("ok") else f"บันทึกลงไฟล์ในเครื่องแล้ว แต่ยัง sync เข้า Google Drive ไม่ได้: {drive_result.get('error')}",
                },
            )
        except Exception as exc:
            self.send_json(HTTPStatus.BAD_REQUEST, {"ok": False, "error": str(exc)})

    def handle_delete_master_allocation(self, fund_key: str):
        query = parse_qs(urlparse(self.path).query)
        quarter = normalize_quarter((query.get("quarter") or [""])[0] or DEFAULT_QUARTER)
        key = normalize_override_key(fund_key)
        data = read_master_allocations(quarter)
        removed = data["items"].pop(key, None)
        data["updatedAt"] = datetime.now(tz=timezone.utc).isoformat()
        data.pop("sourcePath", None)
        path = write_master_allocations(data, quarter)
        drive_result = upload_master_allocations_to_drive(quarter, data)
        self.send_json(
            HTTPStatus.OK,
            {
                "ok": True,
                "removed": bool(removed),
                "key": key,
                "quarter": quarter,
                "source": str(path.relative_to(ROOT)),
                "drive": drive_result,
                "driveUploaded": bool(drive_result.get("ok")),
                "warning": None if drive_result.get("ok") else f"ลบจากไฟล์ในเครื่องแล้ว แต่ยัง sync เข้า Google Drive ไม่ได้: {drive_result.get('error')}",
            },
        )

    def handle_dividend_check(self, parsed):
        query = parse_qs(parsed.query)
        proj_id = (query.get("proj_id") or [""])[0].strip().upper()
        if not proj_id:
            return self.send_json(HTTPStatus.BAD_REQUEST, {"ok": False, "error": "proj_id is required"})
        try:
            from dividend_live_check_lib import summarize_dividend_check

            result = summarize_dividend_check(proj_id)
            return self.send_json(HTTPStatus.OK, {"ok": True, "result": result})
        except ModuleNotFoundError as exc:
            if exc.name == "requests":
                return self.send_json(
                    HTTPStatus.SERVICE_UNAVAILABLE,
                    {
                        "ok": False,
                        "error": "Dividend check requires the Python package 'requests'. The main server is still available.",
                    },
                )
            raise
        except Exception as exc:
            return self.send_json(HTTPStatus.BAD_GATEWAY, {"ok": False, "error": str(exc)})

    def handle_fund_master(self, parsed):
        query = parse_qs(parsed.query)
        q = (query.get("q") or [""])[0].strip().lower()
        limit = int((query.get("limit") or ["300"])[0] or "300")
        limit = max(1, min(limit, 2000))

        if not FUND_MASTER_FILE.exists():
            return self.send_json(
                HTTPStatus.NOT_FOUND,
                {"ok": False, "error": f"Fund master file not found: {FUND_MASTER_FILE.name}"},
            )

        try:
            df = pd.read_excel(FUND_MASTER_FILE)
            keep_cols = [
                c
                for c in [
                    "proj_id",
                    "proj_name_th",
                    "proj_name_en",
                    "proj_abbr_name",
                    "fund_status",
                    "unique_id",
                    "comp_name_th",
                    "comp_name_en",
                    "fund_class_name",
                    "fund_class_detail",
                    "policy_desc",
                    "management_style",
                    "last_upd_date",
                ]
                if c in df.columns
            ]
            df = df[keep_cols].copy()

            for col in df.columns:
                df[col] = df[col].where(pd.notna(df[col]), None)

            if q:
                mask = pd.Series(False, index=df.index)
                for col in ["proj_id", "proj_name_th", "proj_name_en", "proj_abbr_name", "comp_name_th", "fund_class_name"]:
                    if col in df.columns:
                        mask = mask | df[col].astype(str).str.lower().str.contains(q, na=False)
                df = df[mask]

            total = len(df)
            rows = df.head(limit).to_dict(orient="records")
            return self.send_json(
                HTTPStatus.OK,
                {
                    "ok": True,
                    "total": total,
                    "returned": len(rows),
                    "limit": limit,
                    "rows": rows,
                    "source": str(FUND_MASTER_FILE.relative_to(ROOT)),
                },
            )
        except Exception as exc:
            return self.send_json(HTTPStatus.BAD_GATEWAY, {"ok": False, "error": str(exc)})

    def handle_ft_historical_prices_sync(self):
        try:
            payload = self.read_request_json()
            symbol = str(payload.get("symbol") or "").strip()
            symbols_raw = str(payload.get("symbols") or "").strip()
            all_symbols_from_db_raw = payload.get("allSymbolsFromDb", payload.get("all_symbols_from_db", False))
            all_symbols_from_db = str(all_symbols_from_db_raw).strip().lower() in {"1", "true", "yes", "on"}
            continue_on_error_raw = payload.get("continueOnError", payload.get("continue_on_error", True))
            continue_on_error = str(continue_on_error_raw).strip().lower() not in {"0", "false", "no", "off"}
            try:
                sleep_seconds = max(0.0, float(payload.get("sleepSeconds", payload.get("sleep_seconds", 0)) or 0))
            except (TypeError, ValueError):
                sleep_seconds = 0.0
            ft_url = str(payload.get("url") or "").strip()
            product_match = re.search(r"/data/(etfs|funds)/tearsheet/", ft_url)
            product_path = product_match.group(1) if product_match else None
            start_date = str(payload.get("startDate") or "").strip()
            end_date = str(payload.get("endDate") or "").strip()
            if not symbol and ft_url:
                parsed_url = urlparse(ft_url)
                symbol = (parse_qs(parsed_url.query).get("s") or [""])[0].strip()
            if not symbol and not symbols_raw and not all_symbols_from_db:
                return self.send_json(HTTPStatus.BAD_REQUEST, {"ok": False, "error": "FT symbol is required"})
            if not re.match(r"^\d{4}-\d{2}-\d{2}$", start_date):
                return self.send_json(HTTPStatus.BAD_REQUEST, {"ok": False, "error": "startDate must be YYYY-MM-DD"})
            if not re.match(r"^\d{4}-\d{2}-\d{2}$", end_date):
                return self.send_json(HTTPStatus.BAD_REQUEST, {"ok": False, "error": "endDate must be YYYY-MM-DD"})
            if start_date > end_date:
                return self.send_json(HTTPStatus.BAD_REQUEST, {"ok": False, "error": "startDate must be before endDate"})

            from scripts.ft_historical_prices_store import (
                list_symbols_from_db,
                parse_symbol_list,
                sync_historical_prices,
                sync_historical_symbols,
            )

            batch_symbols = parse_symbol_list(symbols_raw)
            if all_symbols_from_db:
                seen = {item.upper() for item in batch_symbols}
                for item in list_symbols_from_db(FT_HISTORICAL_PRICES_DB):
                    if item.upper() not in seen:
                        batch_symbols.append(item)
                        seen.add(item.upper())
                if not batch_symbols:
                    return self.send_json(HTTPStatus.BAD_REQUEST, {"ok": False, "error": "No FT symbols found in the local database"})

            if batch_symbols:
                result = sync_historical_symbols(
                    batch_symbols,
                    start_date,
                    end_date,
                    db_path=FT_HISTORICAL_PRICES_DB,
                    continue_on_error=continue_on_error,
                    sleep_seconds=sleep_seconds,
                )
            else:
                result = sync_historical_prices(symbol, start_date, end_date, product_path=product_path)
            for key in ("csvPath", "profileCsvPath", "sqlitePath", "rawDir"):
                if result.get(key):
                    try:
                        result[key] = str(Path(result[key]).relative_to(ROOT))
                    except ValueError:
                        pass
            for item in result.get("ranges", []):
                if item.get("rawPath"):
                    try:
                        item["rawPath"] = str(Path(item["rawPath"]).relative_to(ROOT))
                    except ValueError:
                        pass
            self.send_json(HTTPStatus.OK, {"ok": True, "result": result})
        except Exception as exc:
            self.send_json(HTTPStatus.BAD_GATEWAY, {"ok": False, "error": str(exc)})

    def handle_ft_qualitative_data_sync(self):
        try:
            payload = self.read_request_json()
            symbol = str(payload.get("symbol") or "").strip()
            symbols_raw = str(payload.get("symbols") or "").strip()
            all_symbols_from_db_raw = payload.get("allSymbolsFromDb", payload.get("all_symbols_from_db", False))
            all_symbols_from_db = str(all_symbols_from_db_raw).strip().lower() in {"1", "true", "yes", "on"}
            continue_on_error_raw = payload.get("continueOnError", payload.get("continue_on_error", True))
            continue_on_error = str(continue_on_error_raw).strip().lower() not in {"0", "false", "no", "off"}
            try:
                sleep_seconds = max(0.0, float(payload.get("sleepSeconds", payload.get("sleep_seconds", 0)) or 0))
            except (TypeError, ValueError):
                sleep_seconds = 0.0
            ft_url = str(payload.get("url") or "").strip()
            product_match = re.search(r"/data/(etfs|funds)/tearsheet/", ft_url)
            product_path = product_match.group(1) if product_match else None
            if not symbol and ft_url:
                parsed_url = urlparse(ft_url)
                symbol = (parse_qs(parsed_url.query).get("s") or [""])[0].strip()
            if not symbol and not symbols_raw and not all_symbols_from_db:
                return self.send_json(HTTPStatus.BAD_REQUEST, {"ok": False, "error": "FT symbol is required"})

            from scripts.ft_historical_prices_store import (
                list_symbols_from_db,
                parse_symbol_list,
                sync_qualitative_data,
                sync_qualitative_symbols,
            )

            batch_symbols = parse_symbol_list(symbols_raw)
            if all_symbols_from_db:
                seen = {item.upper() for item in batch_symbols}
                for item in list_symbols_from_db(FT_HISTORICAL_PRICES_DB):
                    if item.upper() not in seen:
                        batch_symbols.append(item)
                        seen.add(item.upper())
                if not batch_symbols:
                    return self.send_json(HTTPStatus.BAD_REQUEST, {"ok": False, "error": "No FT symbols found in the local database"})

            if batch_symbols:
                result = sync_qualitative_symbols(
                    batch_symbols,
                    db_path=FT_HISTORICAL_PRICES_DB,
                    continue_on_error=continue_on_error,
                    sleep_seconds=sleep_seconds,
                )
                for item in result.get("results", []):
                    for key in ("profileCsvPath", "riskCsvPath", "holdingsCsvPath", "sqlitePath", "snapshotDir"):
                        if item.get(key):
                            try:
                                item[key] = str(Path(item[key]).relative_to(ROOT))
                            except ValueError:
                                pass
            else:
                result = sync_qualitative_data(symbol, db_path=FT_HISTORICAL_PRICES_DB, product_path=product_path)
                for key in ("profileCsvPath", "riskCsvPath", "holdingsCsvPath", "sqlitePath", "snapshotDir"):
                    if result.get(key):
                        try:
                            result[key] = str(Path(result[key]).relative_to(ROOT))
                        except ValueError:
                            pass
            self.send_json(HTTPStatus.OK, {"ok": True, "result": result})
        except Exception as exc:
            self.send_json(HTTPStatus.BAD_GATEWAY, {"ok": False, "error": str(exc)})

    def handle_get_ft_historical_prices(self, parsed):
        try:
            if not FT_HISTORICAL_PRICES_DB.exists():
                return self.send_json(
                    HTTPStatus.OK,
                    {"ok": True, "symbols": [], "rows": [], "stats": [], "source": str(FT_HISTORICAL_PRICES_DB.relative_to(ROOT))},
                )
            from scripts.ft_historical_prices_store import init_db

            init_db(FT_HISTORICAL_PRICES_DB)
            query = parse_qs(parsed.query)
            symbol = (query.get("symbol") or [""])[0].strip()
            limit_raw = (query.get("limit") or ["120"])[0]
            try:
                limit = max(1, min(int(limit_raw), 5000))
            except ValueError:
                limit = 120

            name_lookup = load_master_fund_name_lookup()

            with sqlite3.connect(FT_HISTORICAL_PRICES_DB) as conn:
                conn.row_factory = sqlite3.Row
                ft_name_lookup = {
                    str(row["symbol"]).strip().upper(): str(row["value"] or "").strip()
                    for row in conn.execute(
                        """
                        select symbol, value
                        from ft_profile_investment
                        where section = 'metadata' and field = 'FT display name' and value is not null and value <> ''
                        """
                    )
                }
                symbols = [
                    dict(row)
                    for row in conn.execute(
                        """
                        with symbol_keys as (
                            select symbol from ft_historical_prices
                            union
                            select symbol from ft_profile_investment
                            union
                            select symbol from ft_risk_measures
                            union
                            select symbol from ft_top_holdings
                        ),
                        price_symbols as (
                            select
                                symbol,
                                max(ft_issue_id) as ft_issue_id,
                                count(*) as rowCount,
                                min(price_date) as startDate,
                                max(price_date) as endDate
                            from ft_historical_prices
                            group by symbol
                        ),
                        profile_symbols as (
                            select
                                symbol,
                                max(ft_issue_id) as ft_issue_id,
                                count(*) as profileRowCount
                            from ft_profile_investment
                            group by symbol
                        ),
                        risk_symbols as (
                            select
                                symbol,
                                max(ft_issue_id) as ft_issue_id,
                                count(*) as riskRowCount
                            from ft_risk_measures
                            group by symbol
                        ),
                        holdings_symbols as (
                            select
                                symbol,
                                max(ft_issue_id) as ft_issue_id,
                                count(*) as holdingsRowCount
                            from ft_top_holdings
                            group by symbol
                        ),
                        profile_isins as (
                            select symbol, max(value) as isin
                            from ft_profile_investment
                            where field = 'ISIN' and value is not null and value <> ''
                            group by symbol
                        )
                        select
                            k.symbol as symbol,
                            coalesce(p.ft_issue_id, q.ft_issue_id, r.ft_issue_id, h.ft_issue_id) as ftIssueId,
                            coalesce(i.isin, '') as isin,
                            coalesce(p.rowCount, 0) as rowCount,
                            coalesce(q.profileRowCount, 0) as profileRowCount,
                            coalesce(r.riskRowCount, 0) as riskRowCount,
                            coalesce(h.holdingsRowCount, 0) as holdingsRowCount,
                            coalesce(p.startDate, '') as startDate,
                            coalesce(p.endDate, '') as endDate
                        from symbol_keys k
                        left join price_symbols p on upper(p.symbol) = upper(k.symbol)
                        left join profile_symbols q on upper(q.symbol) = upper(k.symbol)
                        left join risk_symbols r on upper(r.symbol) = upper(k.symbol)
                        left join holdings_symbols h on upper(h.symbol) = upper(k.symbol)
                        left join profile_isins i on upper(i.symbol) = upper(k.symbol)
                        order by symbol
                        """
                    )
                ]
                for item in symbols:
                    item_symbol = str(item.get("symbol") or "").strip()
                    display_name = (
                        ft_name_lookup.get(item_symbol.upper())
                        or name_lookup.get(item_symbol.upper())
                        or name_lookup.get(ft_symbol_base(item_symbol))
                        or ""
                    )
                    item["displayName"] = display_name
                    item["displayLabel"] = f"{item_symbol} — {display_name}" if display_name else item_symbol
                requested_symbol = symbol
                selected_symbol = ""
                if requested_symbol:
                    requested_upper = requested_symbol.upper()
                    for item in symbols:
                        candidate = str(item.get("symbol") or "").strip()
                        candidate_upper = candidate.upper()
                        if (
                            candidate_upper == requested_upper
                            or candidate_upper.startswith(f"{requested_upper}:")
                            or ft_symbol_base(candidate_upper) == requested_upper
                        ):
                            selected_symbol = candidate
                            break
                    selected_symbol = selected_symbol or requested_symbol
                else:
                    selected_symbol = symbols[0]["symbol"] if symbols else ""
                selected_display_name = (
                    ft_name_lookup.get(selected_symbol.upper())
                    or name_lookup.get(selected_symbol.upper())
                    or name_lookup.get(ft_symbol_base(selected_symbol))
                    or ""
                )
                rows = []
                profile_rows = []
                risk_rows = []
                holdings_rows = []
                stats = []
                if selected_symbol:
                    profile_rows = [
                        dict(row)
                        for row in conn.execute(
                            """
                            select
                                section,
                                field,
                                value,
                                source,
                                fetched_at as fetchedAt
                            from ft_profile_investment
                            where symbol = ?
                            order by
                                case section when 'profile' then 1 when 'investment' then 2 else 3 end,
                                field
                            """,
                            (selected_symbol,),
                        )
                    ]
                    risk_rows = [
                        dict(row)
                        for row in conn.execute(
                            """
                            select
                                period,
                                metric,
                                fund_value as fundValue,
                                category_average as categoryAverage,
                                benchmark_used as benchmarkUsed,
                                as_of_date as asOfDate,
                                source,
                                fetched_at as fetchedAt
                            from ft_risk_measures
                            where upper(symbol) = upper(?)
                            order by
                                case period when '1 year' then 1 when '3 year' then 2 when '5 years' then 3 else 4 end,
                                metric
                            """,
                            (selected_symbol,),
                        )
                    ]
                    holdings_rows = [
                        dict(row)
                        for row in conn.execute(
                            """
                            select
                                rank,
                                holding_name as holdingName,
                                holding_symbol as holdingSymbol,
                                one_year_change as oneYearChange,
                                portfolio_weight as portfolioWeight,
                                long_allocation as longAllocation,
                                top10_portfolio_percent as top10PortfolioPercent,
                                as_of_date as asOfDate,
                                source,
                                fetched_at as fetchedAt
                            from ft_top_holdings
                            where upper(symbol) = upper(?)
                            order by rank
                            """,
                            (selected_symbol,),
                        )
                    ]
                    rows = [
                        dict(row)
                        for row in conn.execute(
                            """
                            select symbol, price_date as date, open, high, low, close, volume, source, fetched_at as fetchedAt
                            from ft_historical_prices
                            where symbol = ?
                            order by price_date desc
                            limit ?
                            """,
                            (selected_symbol, limit),
                        )
                    ]
                    ordered = [
                        dict(row)
                        for row in conn.execute(
                            """
                            select price_date as date, close, volume
                            from ft_historical_prices
                            where symbol = ? and close is not null
                            order by price_date asc
                            """,
                            (selected_symbol,),
                        )
                    ]
                    returns = []
                    prev = None
                    for row in ordered:
                        close = row.get("close")
                        if prev and close:
                            returns.append((close / prev) - 1)
                        if close:
                            prev = close
                    first_close = ordered[0]["close"] if ordered else None
                    last_close = ordered[-1]["close"] if ordered else None
                    positive = [value for value in returns if value > 0]
                    negative = [value for value in returns if value < 0]
                    stats.append(
                        {
                            "symbol": selected_symbol,
                            "displayName": selected_display_name,
                            "displayLabel": f"{selected_symbol} — {selected_display_name}" if selected_display_name else selected_symbol,
                            "rowCount": len(ordered),
                            "startDate": ordered[0]["date"] if ordered else "",
                            "endDate": ordered[-1]["date"] if ordered else "",
                            "firstClose": first_close,
                            "lastClose": last_close,
                            "cumulativeReturn": ((last_close / first_close) - 1) if first_close and last_close else None,
                            "dailyReturnCount": len(returns),
                            "upDays": len(positive),
                            "downDays": len(negative),
                            "avgDailyReturn": (sum(returns) / len(returns)) if returns else None,
                            "avgUpDayReturn": (sum(positive) / len(positive)) if positive else None,
                            "avgDownDayReturn": (sum(negative) / len(negative)) if negative else None,
                        }
                    )
            self.send_json(
                HTTPStatus.OK,
                {
                    "ok": True,
                    "symbols": symbols,
                    "selectedSymbol": selected_symbol,
                    "selectedDisplayName": selected_display_name,
                    "selectedDisplayLabel": f"{selected_symbol} — {selected_display_name}" if selected_display_name else selected_symbol,
                    "rows": rows,
                    "profile": profile_rows if selected_symbol else [],
                    "risk": risk_rows if selected_symbol else [],
                    "holdings": holdings_rows if selected_symbol else [],
                    "stats": stats,
                    "source": str(FT_HISTORICAL_PRICES_DB.relative_to(ROOT)),
                },
            )
        except Exception as exc:
            self.send_json(HTTPStatus.BAD_GATEWAY, {"ok": False, "error": str(exc)})

    def handle_get_ft_price_on_date(self, parsed):
        try:
            query = parse_qs(parsed.query)
            symbol = (query.get("symbol") or [""])[0].strip()
            date_value = (query.get("date") or [""])[0].strip()
            if not symbol:
                return self.send_json(HTTPStatus.BAD_REQUEST, {"ok": False, "error": "symbol is required"})
            if not re.match(r"^\d{4}-\d{2}-\d{2}$", date_value):
                return self.send_json(HTTPStatus.BAD_REQUEST, {"ok": False, "error": "date must be YYYY-MM-DD"})
            if not FT_HISTORICAL_PRICES_DB.exists():
                return self.send_json(HTTPStatus.NOT_FOUND, {"ok": False, "error": "FT historical prices database not found"})
            from scripts.ft_historical_prices_store import init_db

            init_db(FT_HISTORICAL_PRICES_DB)
            with sqlite3.connect(FT_HISTORICAL_PRICES_DB) as conn:
                conn.row_factory = sqlite3.Row
                row = conn.execute(
                    """
                    select
                        symbol,
                        price_date as date,
                        open,
                        high,
                        low,
                        close,
                        volume,
                        source,
                        fetched_at as fetchedAt
                    from ft_historical_prices
                    where symbol = ? and price_date <= ?
                    order by price_date desc
                    limit 1
                    """,
                    (symbol, date_value),
                ).fetchone()
            if not row:
                return self.send_json(
                    HTTPStatus.OK,
                    {"ok": True, "price": None, "symbol": symbol, "requestedDate": date_value},
                )
            price = dict(row)
            price["requestedDate"] = date_value
            price["isExactDate"] = price["date"] == date_value
            self.send_json(HTTPStatus.OK, {"ok": True, "price": price})
        except Exception as exc:
            self.send_json(HTTPStatus.BAD_GATEWAY, {"ok": False, "error": str(exc)})

    def handle_save_draft(self):
        try:
            draft = self.read_request_json()
            draft_id = safe_slug(draft.get("id") or str(int(datetime.now(tz=timezone.utc).timestamp() * 1000)))
            draft["id"] = draft_id
            draft["updatedAt"] = datetime.now(tz=timezone.utc).isoformat()
            DRAFTS_DIR.mkdir(parents=True, exist_ok=True)
            path = draft_path(draft_id)
            tmp_path = path.with_suffix(".tmp")
            with tmp_path.open("w", encoding="utf-8") as fh:
                json.dump(draft, fh, ensure_ascii=False, indent=2)
                fh.write("\n")
            tmp_path.replace(path)
            drive_result = upload_draft_to_drive(draft_id, draft)
            self.send_json(
                HTTPStatus.OK,
                {
                    "ok": True,
                    "draft": draft,
                    "path": str(path.relative_to(ROOT)),
                    "drive": drive_result,
                    "driveUploaded": bool(drive_result.get("ok")),
                    "driveFileId": drive_result.get("fileId"),
                    "driveFolderId": drive_result.get("folderId") or DRAFTS_DRIVE_FOLDER_ID,
                    "warning": None if drive_result.get("ok") else f"Saved locally, but Drive sync failed: {drive_result.get('error')}",
                },
            )
        except Exception as exc:
            self.send_json(HTTPStatus.BAD_REQUEST, {"ok": False, "error": str(exc)})

    def handle_delete_draft(self, draft_id: str):
        path = draft_path(draft_id)
        if path.exists():
            path.unlink()
        drive_result = delete_draft_from_drive(draft_id)
        self.send_json(
            HTTPStatus.OK,
            {
                "ok": True,
                "drive": drive_result,
                "warning": None if drive_result.get("ok") else f"Deleted locally, but Drive sync failed: {drive_result.get('error')}",
            },
        )


def main():
    port = 8080
    server = ThreadingHTTPServer(("", port), FundRequestHandler)
    print("============================================")
    print(" Fund Selection Tool - Local Server")
    print("============================================")
    print("")
    print(f"Starting server at http://localhost:{port}")
    print("Draft files will be saved under Drafts/")
    print(f"Draft Drive folder: {DRAFTS_DRIVE_FOLDER_ID or '(not configured)'}")
    print(f"Fund Selection Logs Drive folder: {FUND_SELECTION_LOGS_DRIVE_FOLDER_ID or '(not configured)'}")
    print("Press Ctrl+C to stop")
    print("")
    server.serve_forever()


if __name__ == "__main__":
    main()
