#!/usr/bin/env python3
"""Download dividend-history input JSON files from the Drive JSON store."""

from __future__ import annotations

import argparse
import json
import os
from pathlib import Path
from typing import Any

from drive_json_store import download_json_payload, drive_client, resolve_folder_path


PROJECT_ROOT = Path(__file__).resolve().parents[1]
DEFAULT_JSON_ROOT_FOLDER_ID = "1vUWAU5qP0qiIHPa5C4TZUybVmEwqfl6W"
INPUT_FILES = [
    ("Data For SEC API.json", "Data For SEC API - {quarter}.json"),
    ("Fund Key Performance AVP.json", "Fund Key Performance AVP - {quarter}.json"),
]


def quarter_year(quarter: str) -> str:
    return str(quarter or "").split("-")[0] or "2026"


def write_json(path: Path, payload: Any) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(payload, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--quarter", default="2026-Q1")
    parser.add_argument(
        "--root-folder-id",
        default=os.environ.get("DRIVE_JSON_ROOT_FOLDER_ID", "").strip() or DEFAULT_JSON_ROOT_FOLDER_ID,
    )
    args = parser.parse_args()

    folder_id = resolve_folder_path(
        drive_client(),
        args.root_folder_id,
        [quarter_year(args.quarter), args.quarter, "base"],
    )

    results = []
    for drive_name, local_template in INPUT_FILES:
        payload = download_json_payload(folder_id, drive_name)
        if payload is None:
            raise FileNotFoundError(f"{drive_name} not found in Drive JSON base folder")
        local_path = PROJECT_ROOT / "Data" / local_template.format(quarter=args.quarter)
        write_json(local_path, payload)
        rows = len(payload) if isinstance(payload, list) else 0
        results.append({"driveFile": drive_name, "localPath": str(local_path), "rows": rows})

    print(json.dumps({"ok": True, "quarter": args.quarter, "folderId": folder_id, "files": results}, ensure_ascii=False, indent=2))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
