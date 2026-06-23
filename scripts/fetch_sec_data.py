#!/usr/bin/env python3
"""Fetch SEC fund API samples and export CSV files.

This script is designed for GitHub Actions first-pass testing. It reads the
SEC API key from SEC_API_KEY, fetches a few useful endpoints, and writes one
CSV file per dataset into an output directory.
"""

from __future__ import annotations

import argparse
import csv
import json
import os
import sys
from pathlib import Path
from typing import Any

import requests


BASE_URL = "https://api.sec.or.th"


ENDPOINTS = {
    "amcs": "/v2/fund/general-info/amcs",
    "profiles": "/v2/fund/general-info/profiles",
    "factsheet_urls": "/v2/fund/factsheet/urls",
    "top5_holdings": "/v2/fund/factsheet/top5-holdings",
    "nav_daily": "/v2/fund/daily-info/nav",
    "dividend_history": "/v2/fund/daily-info/dividend-history",
}


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Fetch SEC fund data and export CSV files.")
    parser.add_argument("--output-dir", default="sec_output", help="Directory for CSV output.")
    parser.add_argument("--proj-id", default="M0000_2552", help="Project ID for fund-specific endpoints.")
    parser.add_argument("--start-nav-date", default="2023-10-01", help="NAV start date, YYYY-MM-DD.")
    parser.add_argument("--end-nav-date", default="2023-10-31", help="NAV end date, YYYY-MM-DD.")
    parser.add_argument("--page-size", type=int, default=100, help="API page_size value.")
    parser.add_argument(
        "--max-pages",
        type=int,
        default=3,
        help="Maximum pages per endpoint for this test run. Use 0 for no limit.",
    )
    parser.add_argument(
        "--datasets",
        default="amcs,profiles,factsheet_urls,top5_holdings,nav_daily,dividend_history",
        help="Comma-separated dataset keys or all.",
    )
    return parser.parse_args()


def get_api_key() -> str:
    api_key = os.environ.get("SEC_API_KEY", "").strip()
    if not api_key:
        raise RuntimeError("Missing SEC_API_KEY environment variable.")
    return api_key


def fetch_sec_data(
    endpoint_path: str,
    api_key: str,
    query_params: dict[str, Any] | None = None,
    timeout: int = 60,
) -> dict[str, Any]:
    headers = {
        "Content-Type": "application/json",
        "Cache-Control": "no-cache",
        "Ocp-Apim-Subscription-Key": api_key,
    }
    params = {"next_cursor": ""}
    if query_params:
        params.update(query_params)

    response = requests.get(
        f"{BASE_URL}{endpoint_path}",
        headers=headers,
        params=params,
        timeout=timeout,
    )

    if response.status_code != 200:
        raise RuntimeError(
            f"SEC API error {response.status_code} for {endpoint_path}: {response.text[:1000]}"
        )

    return response.json()


def unwrap_items(data: Any) -> list[dict[str, Any]]:
    if isinstance(data, list):
        return [item if isinstance(item, dict) else {"value": item} for item in data]

    if not isinstance(data, dict):
        return [{"value": data}]

    items = data.get("items")
    if isinstance(items, list):
        return [item if isinstance(item, dict) else {"value": item} for item in items]
    if isinstance(items, dict):
        nested_rows: list[dict[str, Any]] = []
        for key, value in items.items():
            if isinstance(value, list):
                for item in value:
                    row = item if isinstance(item, dict) else {"value": item}
                    nested_rows.append({"items_key": key, **row})
            elif isinstance(value, dict):
                nested_rows.append({"items_key": key, **value})
            else:
                nested_rows.append({"items_key": key, "value": value})
        return nested_rows

    row = {
        key: value
        for key, value in data.items()
        if key not in {"next_cursor", "items"}
    }
    return [row] if row else []


def get_next_cursor(data: Any) -> str:
    if not isinstance(data, dict):
        return ""
    return str(data.get("next_cursor") or "")


def fetch_sec_all_pages(
    endpoint_path: str,
    api_key: str,
    query_params: dict[str, Any] | None = None,
    max_pages: int = 3,
) -> list[dict[str, Any]]:
    rows: list[dict[str, Any]] = []
    next_cursor = ""
    page = 0

    while True:
        page += 1
        params = {"next_cursor": next_cursor}
        if query_params:
            params.update(query_params)

        data = fetch_sec_data(endpoint_path, api_key, params)
        page_rows = unwrap_items(data)
        rows.extend(page_rows)
        print(f"Fetched {endpoint_path} page {page}: {len(page_rows)} rows")

        next_cursor = get_next_cursor(data)
        if not next_cursor:
            break
        if max_pages and page >= max_pages:
            print(f"Stopped {endpoint_path} at max_pages={max_pages}")
            break

    return rows


def flatten_value(value: Any) -> Any:
    if value is None:
        return ""
    if isinstance(value, (str, int, float, bool)):
        return value
    return json.dumps(value, ensure_ascii=False, separators=(",", ":"))


def collect_headers(rows: list[dict[str, Any]]) -> list[str]:
    headers: list[str] = []
    seen: set[str] = set()
    for row in rows:
        for key in row.keys():
            if key not in seen:
                seen.add(key)
                headers.append(key)
    return headers


def write_csv(path: Path, rows: list[dict[str, Any]]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    headers = collect_headers(rows)
    if not headers:
        headers = ["message"]
        rows = [{"message": "No rows returned"}]

    with path.open("w", encoding="utf-8-sig", newline="") as fh:
        writer = csv.DictWriter(fh, fieldnames=headers, extrasaction="ignore")
        writer.writeheader()
        for row in rows:
            writer.writerow({key: flatten_value(row.get(key, "")) for key in headers})

    print(f"Wrote {path}: {len(rows)} rows")


def selected_datasets(raw: str) -> list[str]:
    if raw.strip().lower() == "all":
        return list(ENDPOINTS.keys())
    keys = [item.strip() for item in raw.split(",") if item.strip()]
    unknown = [key for key in keys if key not in ENDPOINTS]
    if unknown:
        raise ValueError(f"Unknown datasets: {', '.join(unknown)}")
    return keys


def dataset_params(dataset: str, args: argparse.Namespace) -> dict[str, Any]:
    if dataset == "profiles":
        return {"page_size": args.page_size}
    if dataset in {"factsheet_urls", "top5_holdings", "dividend_history"}:
        return {"proj_id": args.proj_id, "page_size": args.page_size}
    if dataset == "nav_daily":
        return {
            "proj_id": args.proj_id,
            "start_nav_date": args.start_nav_date,
            "end_nav_date": args.end_nav_date,
            "page_size": args.page_size,
        }
    return {"page_size": args.page_size}


def main() -> int:
    args = parse_args()
    api_key = get_api_key()
    output_dir = Path(args.output_dir)

    for dataset in selected_datasets(args.datasets):
        endpoint = ENDPOINTS[dataset]
        params = dataset_params(dataset, args)
        rows = fetch_sec_all_pages(
            endpoint,
            api_key,
            query_params=params,
            max_pages=args.max_pages,
        )
        write_csv(output_dir / f"{dataset}.csv", rows)

    return 0


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except Exception as exc:
        print(f"ERROR: {exc}", file=sys.stderr)
        raise SystemExit(1)
