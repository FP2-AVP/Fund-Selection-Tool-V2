#!/usr/bin/env python3
"""Export Fund List Tool Google Sheets tabs to JSON files on Google Drive.

This script is intended for GitHub Actions or another scheduled runner.

Example:
  python scripts/export_sheets_to_drive_json.py --quarter 2026-Q3

Authentication:
  Set one of these environment variables:
    GOOGLE_SERVICE_ACCOUNT_JSON_EXPORT='{"type":"service_account",...}'
    GOOGLE_SERVICE_ACCOUNT_JSON='{"type":"service_account",...}'  # legacy fallback
    GOOGLE_APPLICATION_CREDENTIALS=/path/to/service-account.json

If DRIVE_JSON_TARGET_DATA or JSON_DRIVE_ROOT_FOLDER_ID is provided, the script
creates the year/quarter folders inside that folder. Otherwise it creates/finds
"Fund List Tool JSON" in the service account's My Drive.
"""

from __future__ import annotations

import argparse
import json
import os
import re
import sys
import tempfile
from dataclasses import dataclass
from datetime import datetime, timezone
from pathlib import Path
from typing import Any

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets.readonly",
    "https://www.googleapis.com/auth/drive",
]

JSON_ROOT_NAME = "Fund List Tool JSON"


@dataclass(frozen=True)
class Dataset:
    key: str
    sheet_id: str
    output_file: str


DATASETS: dict[str, Dataset] = {
    "sec-api": Dataset(
        key="sec-api",
        sheet_id="16agx9pl9adtMh-U7MCbgnIncBxpciCvFgsdurH6Ob8w",
        output_file="Data For SEC API.json",
    ),
    "fund-key-performance": Dataset(
        key="fund-key-performance",
        sheet_id="1s-0ciSOB2Tj0C9azeMXyd1zZxljOg8I5QilI0FgjdW4",
        output_file="Fund Key Performance AVP.json",
    ),
    "thai-quality": Dataset(
        key="thai-quality",
        sheet_id="1m1rSyJAel9atGMrmeRSwgYWa9wgc4gi7-3cp4Yvc8GM",
        output_file="AVP Thai Fund for Quality.json",
    ),
    "master-fund": Dataset(
        key="master-fund",
        sheet_id="10Bsu4w7CluWdOWYIbi1K6OWoZlVXTSE_ixVl13rWBig",
        output_file="AVP Master Fund ID.json",
    ),
}


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Export quarter tabs from Google Sheets to Drive JSON files."
    )
    parser.add_argument(
        "--output-dir",
        default="",
        help="Optional directory for a local copy, arranged as YEAR/QUARTER/base/*.json.",
    )
    parser.add_argument(
        "--quarter",
        required=True,
        help="Quarter tab to export, e.g. 2026-Q3.",
    )
    parser.add_argument(
        "--dataset",
        choices=["all", *DATASETS.keys()],
        default="all",
        help="Dataset to export. Default: all.",
    )
    parser.add_argument(
        "--root-folder-id",
        default=(
            os.environ.get("DRIVE_JSON_TARGET_DATA", "").strip()
            or os.environ.get("JSON_DRIVE_ROOT_FOLDER_ID", "").strip()
        ),
        help="Optional Google Drive folder ID for the JSON root.",
    )
    parser.add_argument(
        "--dry-run",
        action="store_true",
        help="Print target paths without writing files to Google Drive.",
    )
    return parser.parse_args()


def validate_quarter(quarter: str) -> str:
    quarter = quarter.strip().upper()
    if not re.fullmatch(r"\d{4}-Q[1-4]", quarter):
        raise ValueError("quarter must look like 2026-Q1")
    return quarter


def quarter_year(quarter: str) -> str:
    return quarter.split("-", 1)[0]


def credentials_from_env() -> Any:
    from google.oauth2 import service_account

    raw_json = (
        os.environ.get("GOOGLE_SERVICE_ACCOUNT_JSON_EXPORT", "").strip()
        or os.environ.get("GOOGLE_SERVICE_ACCOUNT_JSON", "").strip()
    )
    if raw_json:
        info = json.loads(raw_json)
        return service_account.Credentials.from_service_account_info(info, scopes=SCOPES)

    credentials_path = os.environ.get("GOOGLE_APPLICATION_CREDENTIALS", "").strip()
    if credentials_path:
        return service_account.Credentials.from_service_account_file(
            credentials_path,
            scopes=SCOPES,
        )

    raise RuntimeError(
        "Set GOOGLE_SERVICE_ACCOUNT_JSON_EXPORT, GOOGLE_SERVICE_ACCOUNT_JSON, "
        "or GOOGLE_APPLICATION_CREDENTIALS first."
    )


def q(value: str) -> str:
    return value.replace("\\", "\\\\").replace("'", "\\'")


def find_child_folder(drive: Any, parent_id: str | None, name: str) -> str | None:
    parent_clause = f" and '{q(parent_id)}' in parents" if parent_id else ""
    response = (
        drive.files()
        .list(
            q=(
                "mimeType='application/vnd.google-apps.folder'"
                f" and name='{q(name)}'"
                " and trashed=false"
                f"{parent_clause}"
            ),
            fields="files(id,name)",
            pageSize=10,
            supportsAllDrives=True,
            includeItemsFromAllDrives=True,
        )
        .execute()
    )
    files = response.get("files", [])
    return files[0]["id"] if files else None


def create_folder(drive: Any, parent_id: str | None, name: str) -> str:
    metadata: dict[str, Any] = {
        "name": name,
        "mimeType": "application/vnd.google-apps.folder",
    }
    if parent_id:
        metadata["parents"] = [parent_id]
    folder = (
        drive.files()
        .create(
            body=metadata,
            fields="id",
            supportsAllDrives=True,
        )
        .execute()
    )
    return folder["id"]


def ensure_folder(drive: Any, parent_id: str | None, name: str) -> str:
    existing = find_child_folder(drive, parent_id, name)
    return existing or create_folder(drive, parent_id, name)


def ensure_json_folders(drive: Any, root_folder_id: str, quarter: str) -> tuple[str, str]:
    year_id = ensure_folder(drive, root_folder_id, quarter_year(quarter))
    quarter_id = ensure_folder(drive, year_id, quarter)
    base_id = ensure_folder(drive, quarter_id, "base")
    overrides_id = ensure_folder(drive, quarter_id, "overrides")
    return base_id, overrides_id


def read_sheet_values(sheets: Any, sheet_id: str, tab_name: str) -> list[list[Any]]:
    response = (
        sheets.spreadsheets()
        .values()
        .get(
            spreadsheetId=sheet_id,
            range=f"'{tab_name}'",
            valueRenderOption="UNFORMATTED_VALUE",
            dateTimeRenderOption="FORMATTED_STRING",
        )
        .execute()
    )
    return response.get("values", [])


def find_file_in_folder(drive: Any, folder_id: str, name: str) -> str | None:
    response = (
        drive.files()
        .list(
            q=f"name='{q(name)}' and '{q(folder_id)}' in parents and trashed=false",
            fields="files(id,name)",
            pageSize=10,
            supportsAllDrives=True,
            includeItemsFromAllDrives=True,
        )
        .execute()
    )
    files = response.get("files", [])
    return files[0]["id"] if files else None


def upload_json(drive: Any, folder_id: str, file_name: str, rows: list[list[Any]]) -> str:
    from googleapiclient.http import MediaFileUpload

    payload = rows
    with tempfile.NamedTemporaryFile("w", encoding="utf-8", suffix=".json", delete=False) as fh:
        json.dump(payload, fh, ensure_ascii=False, separators=(",", ":"))
        fh.write("\n")
        tmp_path = Path(fh.name)

    try:
        media = MediaFileUpload(str(tmp_path), mimetype="application/json", resumable=False)
        existing_id = find_file_in_folder(drive, folder_id, file_name)
        metadata = {
            "name": file_name,
            "mimeType": "application/json",
        }
        if existing_id:
            result = (
                drive.files()
                .update(
                    fileId=existing_id,
                    body=metadata,
                    media_body=media,
                    fields="id",
                    supportsAllDrives=True,
                )
                .execute()
            )
        else:
            metadata["parents"] = [folder_id]
            result = (
                drive.files()
                .create(
                    body=metadata,
                    media_body=media,
                    fields="id",
                    supportsAllDrives=True,
                )
                .execute()
            )
        return result["id"]
    finally:
        tmp_path.unlink(missing_ok=True)


def write_local_json(output_dir: str, quarter: str, file_name: str, rows: list[list[Any]]) -> Path | None:
    if not output_dir:
        return None
    path = Path(output_dir) / quarter_year(quarter) / quarter / "base" / file_name
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(
        json.dumps(rows, ensure_ascii=False, separators=(",", ":")) + "\n",
        encoding="utf-8",
    )
    return path


def main() -> int:
    args = parse_args()
    quarter = validate_quarter(args.quarter)
    selected = DATASETS.values() if args.dataset == "all" else [DATASETS[args.dataset]]

    print(f"Quarter: {quarter}")
    print(f"Generated at: {datetime.now(timezone.utc).isoformat()}")

    if args.dry_run:
        for dataset in selected:
            print(
                f"[dry-run] {JSON_ROOT_NAME}/{quarter_year(quarter)}/{quarter}/base/"
                f"{dataset.output_file}"
            )
        print(
            f"[dry-run] {JSON_ROOT_NAME}/{quarter_year(quarter)}/{quarter}/overrides/"
        )
        return 0

    from googleapiclient.discovery import build

    credentials = credentials_from_env()
    sheets = build("sheets", "v4", credentials=credentials, cache_discovery=False)
    drive = build("drive", "v3", credentials=credentials, cache_discovery=False)

    root_id = args.root_folder_id or ensure_folder(drive, None, JSON_ROOT_NAME)
    base_folder_id, overrides_folder_id = ensure_json_folders(drive, root_id, quarter)

    print(f"Root folder id: {root_id}")
    print(f"Base folder id: {base_folder_id}")
    print(f"Overrides folder id: {overrides_folder_id}")

    for dataset in selected:
        rows = read_sheet_values(sheets, dataset.sheet_id, quarter)
        if not rows:
            raise RuntimeError(f"No data found for {dataset.key} tab {quarter}")
        file_id = upload_json(drive, base_folder_id, dataset.output_file, rows)
        local_path = write_local_json(args.output_dir, quarter, dataset.output_file, rows)
        print(
            f"Uploaded {dataset.output_file}: {len(rows):,} rows, "
            f"{max((len(row) for row in rows), default=0):,} columns, file id {file_id}"
        )
        if local_path:
            print(f"Wrote local copy: {local_path}")

    return 0


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except Exception as exc:
        print(f"ERROR: {exc}", file=sys.stderr)
        raise SystemExit(1)
