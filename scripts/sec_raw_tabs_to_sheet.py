#!/usr/bin/env python3
"""Fetch SEC fund API datasets and write raw endpoint tabs to Google Sheets."""

from __future__ import annotations

import argparse
import json
import os
from pathlib import Path
from typing import Any

from fetch_sec_data import (
    DATASET_FILES,
    ENDPOINTS,
    PROFILE_COLUMNS,
    PROJECT_ID_DATASETS,
    add_if_present,
    dataset_params,
    fetch_dataset_for_registered_proj_ids,
    fetch_profiles_for_proj_ids,
    fetch_registered_profiles,
    fetch_sec_all_pages,
    get_api_key,
    registered_proj_id_rows,
    registered_proj_ids,
    requested_proj_ids,
    should_continue_on_error,
    transform_dataset_rows,
)

SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Write raw SEC endpoint tabs to Google Sheets.")
    parser.add_argument(
        "--spreadsheet-id",
        default=os.environ.get("SEC_MASTER_VIEW_SPREADSHEET_ID", ""),
        help="Google Sheet ID to write to. Can also be set by SEC_MASTER_VIEW_SPREADSHEET_ID.",
    )
    parser.add_argument(
        "--datasets",
        default="profiles",
        help="Comma-separated dataset keys. Use all for every SEC dataset.",
    )
    parser.add_argument("--proj-id", default="", help="Optional single SEC project ID.")
    parser.add_argument(
        "--proj-ids",
        default="",
        help="Optional project IDs separated by comma, whitespace, or new lines.",
    )
    parser.add_argument("--fund-status", default="Registered", help="Profiles fund_status filter.")
    parser.add_argument(
        "--registered-max-funds",
        type=int,
        default=0,
        help="Limit registered proj_id values for testing. Use 0 for no limit.",
    )
    parser.add_argument("--fund-class-name", default="", help="Optional SEC fund_class_name filter.")
    parser.add_argument("--latest", default="true", help="Use latest factsheet data for supported endpoints.")
    parser.add_argument("--start-date", default="", help="Factsheet start_date when latest=false.")
    parser.add_argument("--end-date", default="", help="Factsheet end_date when latest=false.")
    parser.add_argument("--start-period", default="", help="Outstanding data start period, YYYYMM.")
    parser.add_argument("--end-period", default="", help="Outstanding data end period, YYYYMM.")
    parser.add_argument("--start-nav-date", default="", help="NAV start date, YYYY-MM-DD.")
    parser.add_argument("--end-nav-date", default="", help="NAV end date, YYYY-MM-DD.")
    parser.add_argument("--page-size", type=int, default=100, help="SEC API page_size.")
    parser.add_argument(
        "--max-pages",
        type=int,
        default=0,
        help="Maximum pages per endpoint. Use 0 for no limit.",
    )
    parser.add_argument(
        "--continue-on-error",
        choices=["true", "false"],
        default="true",
        help="Continue when endpoint returns an error.",
    )
    parser.add_argument(
        "--status-tab-name",
        default="sec_endpoint_status",
        help="Tab name for endpoint status summary.",
    )
    return parser.parse_args()


def selected_datasets(raw: str) -> list[str]:
    if raw.strip().lower() == "all":
        return list(ENDPOINTS.keys())
    keys = [item.strip() for item in raw.split(",") if item.strip()]
    unknown = [key for key in keys if key not in ENDPOINTS]
    if unknown:
        raise ValueError(f"Unknown datasets: {', '.join(unknown)}")
    if "profiles" not in keys:
        keys.insert(0, "profiles")
    return list(dict.fromkeys(keys))


def dataset_tab_name(dataset: str) -> str:
    file_name = DATASET_FILES.get(dataset, f"{dataset}.csv")
    return Path(file_name).stem


def credentials_from_env() -> Any:
    from google.oauth2 import service_account

    raw_json = os.environ.get("GOOGLE_SERVICE_ACCOUNT_JSON", "").strip()
    if raw_json:
        info = json.loads(raw_json)
        return service_account.Credentials.from_service_account_info(info, scopes=SCOPES)

    credentials_path = os.environ.get("GOOGLE_APPLICATION_CREDENTIALS", "").strip()
    if credentials_path:
        return service_account.Credentials.from_service_account_file(
            credentials_path,
            scopes=SCOPES,
        )

    raise RuntimeError("Set GOOGLE_SERVICE_ACCOUNT_JSON or GOOGLE_APPLICATION_CREDENTIALS first.")


def flatten_value(value: Any) -> Any:
    if value is None:
        return ""
    if isinstance(value, (str, int, float, bool)):
        return value
    return json.dumps(value, ensure_ascii=False, separators=(",", ":"))


def ordered_headers(rows: list[dict[str, Any]], preferred: list[str] | None = None) -> list[str]:
    headers: list[str] = []
    seen: set[str] = set()
    for key in preferred or []:
        if key not in seen:
            seen.add(key)
            headers.append(key)
    for row in rows:
        for key in row.keys():
            if key not in seen:
                seen.add(key)
                headers.append(key)
    return headers or (preferred or ["message"])


def values_from_rows(
    rows: list[dict[str, Any]],
    preferred_headers: list[str] | None = None,
) -> list[list[Any]]:
    headers = ordered_headers(rows, preferred_headers)
    if not rows:
        return [headers]
    return [headers] + [
        [flatten_value(row.get(header, "")) for header in headers]
        for row in rows
    ]


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


def write_values_to_sheet(
    sheets: Any,
    spreadsheet_id: str,
    tab_name: str,
    values: list[list[Any]],
) -> None:
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
        chunk = values[start : start + 1000]
        sheets.spreadsheets().values().update(
            spreadsheetId=spreadsheet_id,
            range=f"'{tab_name}'!A{start + 1}",
            valueInputOption="RAW",
            body={"values": chunk},
        ).execute()
        print(f"Wrote {tab_name} rows {start + 1} - {start + len(chunk)}")


def fetch_dataset_rows(
    dataset: str,
    api_key: str,
    args: argparse.Namespace,
    project_ids: list[str],
    profile_rows: list[dict[str, Any]],
    errors: list[dict[str, Any]],
) -> tuple[list[dict[str, Any]], list[str] | None]:
    if dataset == "profiles":
        return profile_rows, PROFILE_COLUMNS

    endpoint = ENDPOINTS[dataset]
    try:
        if dataset in PROJECT_ID_DATASETS and project_ids:
            raw_rows = fetch_dataset_for_registered_proj_ids(
                dataset,
                endpoint,
                project_ids,
                api_key,
                args,
                errors,
            )
        else:
            params = dataset_params(dataset, args)
            raw_rows = fetch_sec_all_pages(
                endpoint,
                api_key,
                query_params=params,
                max_pages=args.max_pages,
            )
        return transform_dataset_rows(dataset, raw_rows)
    except Exception as exc:
        errors.append(
            {
                "dataset": dataset,
                "endpoint": endpoint,
                "proj_id": "",
                "params": json.dumps(dataset_params(dataset, args), ensure_ascii=False, separators=(",", ":")),
                "error": str(exc),
            }
        )
        print(f"ERROR fetching {dataset}: {exc}")
        if not should_continue_on_error(args):
            raise
        rows, headers = transform_dataset_rows(dataset, [])
        return rows, headers


def build_status_rows(
    datasets: list[str],
    row_counts: dict[str, int],
    project_ids: list[str],
    errors: list[dict[str, Any]],
) -> list[dict[str, Any]]:
    error_counts: dict[str, int] = {}
    for error in errors:
        dataset = str(error.get("dataset") or "")
        if dataset:
            error_counts[dataset] = error_counts.get(dataset, 0) + 1

    return [
        {
            "dataset": dataset,
            "tab_name": dataset_tab_name(dataset),
            "row_count": row_counts.get(dataset, 0),
            "project_count": len(project_ids),
            "error_count": error_counts.get(dataset, 0),
            "status": "ready" if row_counts.get(dataset, 0) else "empty",
        }
        for dataset in datasets
    ]


def main() -> int:
    args = parse_args()
    if not args.spreadsheet_id.strip():
        raise RuntimeError("Missing --spreadsheet-id or SEC_MASTER_VIEW_SPREADSHEET_ID.")

    api_key = get_api_key()
    datasets = selected_datasets(args.datasets)
    errors: list[dict[str, Any]] = []
    explicit_project_ids = requested_proj_ids(args)

    if explicit_project_ids:
        profile_rows = fetch_profiles_for_proj_ids(explicit_project_ids, api_key, args, errors)
        project_ids = explicit_project_ids
    else:
        profile_rows = fetch_registered_profiles(api_key, args)
        project_ids = registered_proj_ids(profile_rows, args.registered_max_funds)
        if args.registered_max_funds and project_ids:
            project_id_set = set(project_ids)
            profile_rows = [
                row for row in profile_rows
                if str(row.get("proj_id") or "").strip() in project_id_set
            ]

    print(f"Selected datasets: {', '.join(datasets)}")
    print(f"Profiles rows: {len(profile_rows)}")
    print(f"Project IDs: {len(project_ids)}")

    from googleapiclient.discovery import build

    sheets = build("sheets", "v4", credentials=credentials_from_env(), cache_discovery=False)
    row_counts: dict[str, int] = {}

    for dataset in datasets:
        rows, preferred_headers = fetch_dataset_rows(
            dataset,
            api_key,
            args,
            project_ids,
            profile_rows,
            errors,
        )
        row_counts[dataset] = len(rows)
        tab_name = dataset_tab_name(dataset)
        values = values_from_rows(rows, preferred_headers)
        write_values_to_sheet(sheets, args.spreadsheet_id.strip(), tab_name, values)
        print(f"Dataset {dataset}: {len(rows)} rows -> {tab_name}")

    if project_ids:
        write_values_to_sheet(
            sheets,
            args.spreadsheet_id.strip(),
            "00_registered_proj_ids",
            values_from_rows(registered_proj_id_rows(project_ids), ["proj_id"]),
        )

    status_rows = build_status_rows(datasets, row_counts, project_ids, errors)
    write_values_to_sheet(
        sheets,
        args.spreadsheet_id.strip(),
        args.status_tab_name,
        values_from_rows(
            status_rows,
            ["dataset", "tab_name", "row_count", "project_count", "error_count", "status"],
        ),
    )

    if errors:
        write_values_to_sheet(
            sheets,
            args.spreadsheet_id.strip(),
            "sec_fetch_errors",
            values_from_rows(
                errors,
                ["dataset", "endpoint", "proj_id", "params", "error"],
            ),
        )

    return 0


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except Exception as exc:
        print(f"ERROR: {exc}")
        raise SystemExit(1)
