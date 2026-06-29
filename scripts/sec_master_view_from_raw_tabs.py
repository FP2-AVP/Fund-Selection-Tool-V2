#!/usr/bin/env python3
"""Build SEC master_view from raw endpoint tabs already stored in Google Sheets."""

from __future__ import annotations

import argparse
import json
import os
from typing import Any

SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]
SAFE_CELL_CHARS = 49_000


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Build SEC master_view from existing raw Google Sheet tabs.")
    parser.add_argument(
        "--spreadsheet-id",
        default=os.environ.get("SEC_MASTER_VIEW_SPREADSHEET_ID", ""),
        help="Google Sheet ID. Can also be set by SEC_MASTER_VIEW_SPREADSHEET_ID.",
    )
    parser.add_argument("--output-tab", default="master_view", help="Target tab name for master_view.")
    return parser.parse_args()


def credentials_from_env() -> Any:
    from google.oauth2 import service_account

    raw_json = os.environ.get("GOOGLE_SERVICE_ACCOUNT_JSON", "").strip()
    if raw_json:
        info = json.loads(raw_json)
        return service_account.Credentials.from_service_account_info(info, scopes=SCOPES)

    credentials_path = os.environ.get("GOOGLE_APPLICATION_CREDENTIALS", "").strip()
    if credentials_path:
        return service_account.Credentials.from_service_account_file(credentials_path, scopes=SCOPES)

    raise RuntimeError("Set GOOGLE_SERVICE_ACCOUNT_JSON or GOOGLE_APPLICATION_CREDENTIALS first.")


def normalized_text(value: Any) -> str:
    if value is None:
        return ""
    return str(value).strip()


def normalized_key(value: Any) -> str:
    return normalized_text(value).casefold()


def numeric_text(value: Any) -> str:
    text = normalized_text(value)
    if text.endswith(".0"):
        return text[:-2]
    return text


def row_date(row: dict[str, Any]) -> str:
    return normalized_text(
        row.get("last_upd_date")
        or row.get("as_of_date")
        or row.get("start_date")
        or row.get("nav_date")
        or row.get("period")
    )


def row_proj(row: dict[str, Any]) -> str:
    return normalized_text(row.get("proj_id"))


def row_class(row: dict[str, Any]) -> str:
    return normalized_text(row.get("fund_class_name"))


def proj_class_key(row: dict[str, Any]) -> tuple[str, str]:
    return row_proj(row), row_class(row)


def truncate_cell_value(value: Any) -> Any:
    if not isinstance(value, str) or len(value) <= SAFE_CELL_CHARS:
        return value
    return (
        value[:SAFE_CELL_CHARS]
        + f"...[truncated {len(value) - SAFE_CELL_CHARS} chars to fit Google Sheets cell limit]"
    )


def read_tab_rows(sheets: Any, spreadsheet_id: str, tab_name: str) -> list[dict[str, Any]]:
    result = sheets.spreadsheets().values().get(
        spreadsheetId=spreadsheet_id,
        range=f"'{tab_name}'",
    ).execute()
    values = result.get("values", [])
    if not values:
        print(f"Read {tab_name}: 0 rows")
        return []

    headers = [normalized_text(value) for value in values[0]]
    rows: list[dict[str, Any]] = []
    for values_row in values[1:]:
        row = {
            header: values_row[index] if index < len(values_row) else ""
            for index, header in enumerate(headers)
            if header
        }
        if any(normalized_text(value) for value in row.values()):
            rows.append(row)
    print(f"Read {tab_name}: {len(rows)} rows")
    return rows


def latest_row(rows: list[dict[str, Any]]) -> dict[str, Any]:
    return sorted(rows, key=row_date, reverse=True)[0] if rows else {}


def latest_lookup_by_key(rows: list[dict[str, Any]], key_fn) -> dict[Any, dict[str, Any]]:
    grouped: dict[Any, list[dict[str, Any]]] = {}
    for row in rows:
        key = key_fn(row)
        if key:
            grouped.setdefault(key, []).append(row)
    return {key: latest_row(items) for key, items in grouped.items()}


def fee_management_lookup(rows: list[dict[str, Any]]) -> dict[tuple[str, str], dict[str, Any]]:
    filtered = [
        row for row in rows
        if normalized_key(row.get("fee_type_desc")) == "management fee"
    ]
    return latest_lookup_by_key(filtered, proj_class_key)


def top1_lookup(rows: list[dict[str, Any]]) -> dict[str, dict[str, Any]]:
    filtered = [
        row for row in rows
        if numeric_text(row.get("asset_seq")) == "1"
    ]
    return latest_lookup_by_key(filtered, row_proj)


def performance_1y_lookup(rows: list[dict[str, Any]]) -> dict[tuple[str, str], dict[str, Any]]:
    filtered: list[dict[str, Any]] = []
    for row in rows:
        ref = normalized_key(row.get("reference_period")).replace(" ", "")
        if ref in {"1y", "1yr", "1year"}:
            filtered.append(row)
    return latest_lookup_by_key(filtered, proj_class_key)


def benchmarks_lookup(rows: list[dict[str, Any]]) -> dict[str, str]:
    grouped: dict[str, list[str]] = {}
    seen_by_proj: dict[str, set[str]] = {}
    for row in sorted(rows, key=lambda item: (row_proj(item), numeric_text(item.get("group_seq")), row_date(item))):
        proj_id = row_proj(row)
        benchmark = normalized_text(row.get("benchmark"))
        if not proj_id or not benchmark:
            continue
        seen = seen_by_proj.setdefault(proj_id, set())
        if benchmark in seen:
            continue
        seen.add(benchmark)
        grouped.setdefault(proj_id, []).append(benchmark)
    return {proj_id: " | ".join(values) for proj_id, values in grouped.items()}


def ensure_sheet(sheets: Any, spreadsheet_id: str, tab_name: str) -> int:
    metadata = sheets.spreadsheets().get(spreadsheetId=spreadsheet_id).execute()
    existing = {
        sheet["properties"]["title"]: sheet["properties"]["sheetId"]
        for sheet in metadata.get("sheets", [])
    }
    if tab_name in existing:
        return existing[tab_name]
    result = sheets.spreadsheets().batchUpdate(
        spreadsheetId=spreadsheet_id,
        body={"requests": [{"addSheet": {"properties": {"title": tab_name}}}]},
    ).execute()
    return result["replies"][0]["addSheet"]["properties"]["sheetId"]


def resize_sheet_grid(
    sheets: Any,
    spreadsheet_id: str,
    sheet_id: int,
    row_count: int,
    column_count: int,
) -> None:
    sheets.spreadsheets().batchUpdate(
        spreadsheetId=spreadsheet_id,
        body={
            "requests": [
                {
                    "updateSheetProperties": {
                        "properties": {
                            "sheetId": sheet_id,
                            "gridProperties": {
                                "rowCount": max(row_count, 1),
                                "columnCount": max(column_count, 1),
                            },
                        },
                        "fields": "gridProperties(rowCount,columnCount)",
                    }
                }
            ]
        },
    ).execute()


def write_values_to_sheet(sheets: Any, spreadsheet_id: str, tab_name: str, values: list[list[Any]]) -> None:
    sheet_id = ensure_sheet(sheets, spreadsheet_id, tab_name)
    row_count = len(values) if values else 1
    column_count = max((len(row) for row in values), default=1)
    resize_sheet_grid(sheets, spreadsheet_id, sheet_id, row_count, column_count)
    sheets.spreadsheets().values().clear(
        spreadsheetId=spreadsheet_id,
        range=f"'{tab_name}'",
        body={},
    ).execute()
    if not values:
        return
    for start in range(0, len(values), 1000):
        chunk = [
            [truncate_cell_value(cell) for cell in row]
            for row in values[start : start + 1000]
        ]
        sheets.spreadsheets().values().update(
            spreadsheetId=spreadsheet_id,
            range=f"'{tab_name}'!A{start + 1}",
            valueInputOption="RAW",
            body={"values": chunk},
        ).execute()
        print(f"Wrote {tab_name} rows {start + 1} - {start + len(chunk)}")


def build_master_rows(raw_tabs: dict[str, list[dict[str, Any]]]) -> list[dict[str, Any]]:
    profiles = raw_tabs["02_profiles"]
    management_fees = fee_management_lookup(raw_tabs["14_fees"])
    top1_holdings = top1_lookup(raw_tabs["17_top5_holdings"])
    performance_1y = performance_1y_lookup(raw_tabs["15_performance"])
    benchmarks = benchmarks_lookup(raw_tabs["08_benchmarks"])

    master_rows: list[dict[str, Any]] = []
    for profile in profiles:
        proj_id = row_proj(profile)
        fund_class_name = row_class(profile)
        key = (proj_id, fund_class_name)
        fee = management_fees.get(key, {})
        holding = top1_holdings.get(proj_id, {})
        performance = performance_1y.get(key, {})

        master_rows.append(
            {
                "proj_id": proj_id,
                "fund_class_name": fund_class_name,
                "fund_name": normalized_text(
                    profile.get("proj_abbr_name")
                    or profile.get("proj_name_th")
                    or profile.get("proj_name_en")
                ),
                "fund_status": normalized_text(profile.get("fund_status")),
                "last_upd_date": normalized_text(profile.get("last_upd_date")),
                "management_fee_actual": normalized_text(fee.get("actual_value")),
                "top1_holding_name": normalized_text(holding.get("asset_name")),
                "top1_holding_ratio": normalized_text(holding.get("asset_ratio")),
                "performance_1y": normalized_text(performance.get("performance_value")),
                "benchmarks": benchmarks.get(proj_id, ""),
            }
        )
    return master_rows


def rows_to_values(rows: list[dict[str, Any]], headers: list[str]) -> list[list[Any]]:
    return [headers] + [
        [row.get(header, "") for header in headers]
        for row in rows
    ]


def main() -> int:
    args = parse_args()
    if not args.spreadsheet_id.strip():
        raise RuntimeError("Missing --spreadsheet-id or SEC_MASTER_VIEW_SPREADSHEET_ID.")

    from googleapiclient.discovery import build

    sheets = build("sheets", "v4", credentials=credentials_from_env(), cache_discovery=False)
    spreadsheet_id = args.spreadsheet_id.strip()
    required_tabs = [
        "02_profiles",
        "08_benchmarks",
        "14_fees",
        "15_performance",
        "17_top5_holdings",
    ]
    raw_tabs = {
        tab_name: read_tab_rows(sheets, spreadsheet_id, tab_name)
        for tab_name in required_tabs
    }
    rows = build_master_rows(raw_tabs)
    headers = [
        "proj_id",
        "fund_class_name",
        "fund_name",
        "fund_status",
        "last_upd_date",
        "management_fee_actual",
        "top1_holding_name",
        "top1_holding_ratio",
        "performance_1y",
        "benchmarks",
    ]
    write_values_to_sheet(sheets, spreadsheet_id, args.output_tab.strip(), rows_to_values(rows, headers))
    print(f"Master view rows: {len(rows)}")
    return 0


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except Exception as exc:
        print(f"ERROR: {exc}")
        raise SystemExit(1)
