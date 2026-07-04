#!/usr/bin/env python3
"""Export Fund Selection Logs Google Sheet rows to a quarter JSON file on Drive.

The sheet is treated as the editing/database surface: one row per fund mention.
For each requested quarter, this script creates the tab when missing, ensures the
header row exists, then groups rows into the JSON shape used by the web app.
"""

from __future__ import annotations

import argparse
import json
import os
import re
import sys
from datetime import datetime, timezone
from pathlib import Path
from typing import Any

from drive_json_store import upload_json_payload


SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]

DEFAULT_SHEET_ID = "1fOdq3JSKTjRZLE8sQ62jn3OmmuKGuhDo2tUCLjy1zIg"
DEFAULT_DRIVE_FOLDER_ID = "12ciJQq-dpBr-DpdnzXCOXqtW_ijctJN6"

HEADERS = [
    "quarter",
    "item_order",
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
]

ROLE_MAP = {
    "mainchoice": "mainChoice",
    "ตัวเลือกหลัก": "mainChoice",
    "secondarychoice": "secondaryChoice",
    "ตัวเลือกรอง": "secondaryChoice",
    "additionalnote": "additionalNote",
    "ความเห็นเพิ่มเติม": "additionalNote",
    "notselected": "notSelected",
    "ไม่ถูกคัดเลือก": "notSelected",
}


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Export Fund Selection Logs sheet rows to Drive JSON."
    )
    parser.add_argument("--quarter", required=True, help="Quarter tab, e.g. 2026-Q1.")
    parser.add_argument(
        "--sheet-id",
        default=os.environ.get("FUND_SELECTION_LOGS_SHEET_ID") or DEFAULT_SHEET_ID,
        help="Google Sheet ID containing Fund Selection Logs tabs.",
    )
    parser.add_argument(
        "--drive-folder-id",
        default=os.environ.get("DRIVE_JSON_TARGET_FUND_SELECTION_LOGS") or DEFAULT_DRIVE_FOLDER_ID,
        help="Drive folder ID that receives Fund Selection Logs - {quarter}.json.",
    )
    parser.add_argument(
        "--dry-run",
        action="store_true",
        help="Validate and print the JSON payload without uploading to Drive.",
    )
    return parser.parse_args()


def validate_quarter(value: str) -> str:
    quarter = value.strip().upper()
    if not re.fullmatch(r"\d{4}-Q[1-4]", quarter):
        raise ValueError("quarter must look like 2026-Q1")
    return quarter


def credentials_from_env() -> Any:
    from google.oauth2 import service_account

    raw_json = os.environ.get("GOOGLE_SERVICE_ACCOUNT_JSON", "").strip()
    if raw_json:
        return service_account.Credentials.from_service_account_info(
            json.loads(raw_json),
            scopes=SCOPES,
        )

    credentials_path = os.environ.get("GOOGLE_APPLICATION_CREDENTIALS", "").strip()
    if credentials_path:
        return service_account.Credentials.from_service_account_file(
            credentials_path,
            scopes=SCOPES,
        )

    raise RuntimeError("Set GOOGLE_SERVICE_ACCOUNT_JSON or GOOGLE_APPLICATION_CREDENTIALS first.")


def sheet_titles(sheets: Any, spreadsheet_id: str) -> set[str]:
    response = (
        sheets.spreadsheets()
        .get(spreadsheetId=spreadsheet_id, fields="sheets(properties(title))")
        .execute()
    )
    return {
        row.get("properties", {}).get("title", "")
        for row in response.get("sheets", [])
    }


def ensure_tab(sheets: Any, spreadsheet_id: str, quarter: str) -> None:
    if quarter in sheet_titles(sheets, spreadsheet_id):
        return
    sheets.spreadsheets().batchUpdate(
        spreadsheetId=spreadsheet_id,
        body={"requests": [{"addSheet": {"properties": {"title": quarter}}}]},
    ).execute()


def read_values(sheets: Any, spreadsheet_id: str, quarter: str) -> list[list[Any]]:
    response = (
        sheets.spreadsheets()
        .values()
        .get(
            spreadsheetId=spreadsheet_id,
            range=f"'{quarter}'",
            valueRenderOption="UNFORMATTED_VALUE",
            dateTimeRenderOption="FORMATTED_STRING",
        )
        .execute()
    )
    return response.get("values", [])


def ensure_headers(sheets: Any, spreadsheet_id: str, quarter: str, values: list[list[Any]]) -> list[str]:
    current = [str(value).strip() for value in (values[0] if values else [])]
    if not current:
        next_headers = HEADERS
    else:
        lowered = {header.lower() for header in current}
        next_headers = [*current, *[header for header in HEADERS if header.lower() not in lowered]]

    if next_headers != current:
        end_col = column_name(len(next_headers))
        sheets.spreadsheets().values().update(
            spreadsheetId=spreadsheet_id,
            range=f"'{quarter}'!A1:{end_col}1",
            valueInputOption="RAW",
            body={"values": [next_headers]},
        ).execute()
    return next_headers


def column_name(index: int) -> str:
    name = ""
    while index:
        index, remainder = divmod(index - 1, 26)
        name = chr(65 + remainder) + name
    return name


def clean(value: Any) -> str:
    return str(value if value is not None else "").strip()


def slug(value: str, fallback: str) -> str:
    out = re.sub(r"\s+", "-", value.strip().lower())
    out = re.sub(r"[^a-z0-9ก-๙._-]+", "", out)
    return out.strip("-")[:80] or fallback


def normalize_role(value: str) -> str:
    key = value.strip().replace(" ", "").lower()
    return ROLE_MAP.get(key, value.strip() or "mainChoice")


def split_tags(value: str) -> list[str]:
    return [
        part.strip()
        for part in re.split(r"[,|]", value or "")
        if part.strip()
    ]


def rows_to_records(headers: list[str], values: list[list[Any]]) -> list[dict[str, str]]:
    normalized = [header.strip().lower() for header in headers]
    records: list[dict[str, str]] = []
    for row in values[1:]:
        if not any(clean(cell) for cell in row):
            continue
        record: dict[str, str] = {}
        for idx, header in enumerate(normalized):
            record[header] = clean(row[idx] if idx < len(row) else "")
        records.append(record)
    return records


def build_payload(quarter: str, records: list[dict[str, str]], drive_folder_id: str) -> dict[str, Any]:
    generated_at = datetime.now(timezone.utc).isoformat()
    grouped: dict[tuple[str, str, str], dict[str, Any]] = {}

    for index, record in enumerate(records, start=1):
        row_quarter = clean(record.get("quarter")) or quarter
        if row_quarter.upper() != quarter:
            continue

        asset_class = clean(record.get("asset_class"))
        fund_type = clean(record.get("fund_type"))
        category = clean(record.get("category"))
        if not asset_class and not fund_type:
            continue

        key = (asset_class, fund_type, category)
        if key not in grouped:
            item_id = f"{quarter}-{slug(asset_class or f'section-{index}', 'section')}-{slug(fund_type or 'general', 'general')}"
            grouped[key] = {
                "id": item_id,
                "assetClass": asset_class,
                "fundType": fund_type,
                "category": category,
                "itemRevision": int(clean(record.get("item_revision")) or "1"),
                "updatedAt": clean(record.get("updated_at")) or generated_at,
                "updatedBy": clean(record.get("updated_by")),
                "itemOrder": int(float(clean(record.get("item_order")) or index)),
                "mentions": [],
            }

        fund_code = clean(record.get("fund_code")).upper()
        reason = clean(record.get("reason"))
        tags = split_tags(clean(record.get("tags")))
        if not fund_code and not reason and not tags:
            continue

        mention_id = clean(record.get("mention_id")) or f"mention-{index}"
        grouped[key]["mentions"].append(
            {
                "id": mention_id,
                "fundCode": fund_code,
                "role": normalize_role(clean(record.get("role"))),
                "status": clean(record.get("status")),
                "sentiment": "neutral",
                "reason": reason,
                "tags": tags,
                "updatedAt": clean(record.get("updated_at")) or generated_at,
                "updatedBy": clean(record.get("updated_by")),
            }
        )

    items = sorted(grouped.values(), key=lambda item: (item.get("itemOrder", 0), item["assetClass"], item["fundType"]))
    for item in items:
        item.pop("itemOrder", None)

    data_as_of = next((clean(row.get("data_as_of")) for row in records if clean(row.get("data_as_of"))), "")
    return {
        "schemaVersion": 1,
        "quarter": quarter,
        "title": f"Fund Selection Logs {quarter}",
        "revision": 1,
        "dataAsOf": data_as_of,
        "createdAt": generated_at,
        "updatedAt": generated_at,
        "updatedBy": "github-actions",
        "driveFolderId": drive_folder_id,
        "items": items,
    }


def main() -> int:
    args = parse_args()
    quarter = validate_quarter(args.quarter)

    from googleapiclient.discovery import build

    credentials = credentials_from_env()
    sheets = build("sheets", "v4", credentials=credentials, cache_discovery=False)

    ensure_tab(sheets, args.sheet_id, quarter)
    values = read_values(sheets, args.sheet_id, quarter)
    headers = ensure_headers(sheets, args.sheet_id, quarter, values)
    if not values:
        values = [headers]
    elif headers != values[0]:
        values = read_values(sheets, args.sheet_id, quarter)

    records = rows_to_records(headers, values)
    payload = build_payload(quarter, records, args.drive_folder_id)
    file_name = f"Fund Selection Logs - {quarter}.json"

    if args.dry_run:
        print(json.dumps(payload, ensure_ascii=False, indent=2))
        return 0

    file_id = upload_json_payload(args.drive_folder_id, file_name, payload)
    mention_count = sum(len(item.get("mentions", [])) for item in payload["items"])
    print(
        f"Uploaded {file_name}: {len(payload['items']):,} items, "
        f"{mention_count:,} mentions, file id {file_id}"
    )
    return 0


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except Exception as exc:
        print(f"ERROR: {exc}", file=sys.stderr)
        raise SystemExit(1)
