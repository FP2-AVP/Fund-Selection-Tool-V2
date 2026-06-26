#!/usr/bin/env python3
"""Build a SEC fund master view and write it directly to Google Sheets.

The script is designed for GitHub Actions. It uses profiles as the base table,
then appends selected SEC endpoint summaries as columns joined by proj_id and,
when available, fund_class_name.
"""

from __future__ import annotations

import argparse
import json
import os
import tempfile
from pathlib import Path
from typing import Any, Callable

from fetch_sec_data import (
    ENDPOINTS,
    PROJECT_ID_DATASETS,
    add_if_present,
    dataset_params,
    fetch_dataset_for_registered_proj_ids,
    fetch_registered_profiles,
    fetch_sec_all_pages,
    get_api_key,
    registered_proj_ids,
    should_continue_on_error,
    transform_dataset_rows,
)

SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]
CONFIG_PATH = Path(__file__).with_name("sec_master_view_config.json")


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Write SEC master_view to Google Sheets.")
    parser.add_argument(
        "--spreadsheet-id",
        default=os.environ.get("SEC_MASTER_VIEW_SPREADSHEET_ID", ""),
        help="Google Sheet ID to write to. Can also be set by SEC_MASTER_VIEW_SPREADSHEET_ID.",
    )
    parser.add_argument("--tab-name", default="", help="Target tab name. Default from config.")
    parser.add_argument(
        "--dataset-preset",
        default="master_core",
        choices=["master_core", "all_master", "daily", "weekly"],
        help="Dataset preset from sec_master_view_config.json.",
    )
    parser.add_argument(
        "--datasets",
        default="",
        help="Optional comma-separated dataset keys. Overrides --dataset-preset.",
    )
    parser.add_argument("--proj-id", default="", help="Optional single SEC project ID for testing.")
    parser.add_argument("--fund-status", default="Registered", help="Profiles fund_status filter.")
    parser.add_argument(
        "--registered-max-funds",
        type=int,
        default=0,
        help="Limit registered funds for testing. Use 0 for no limit.",
    )
    parser.add_argument("--fund-class-name", default="", help="Optional fund_class_name filter.")
    parser.add_argument("--latest", default="true", help="Use latest factsheet data.")
    parser.add_argument("--start-date", default="", help="Factsheet start_date when latest=false.")
    parser.add_argument("--end-date", default="", help="Factsheet end_date when latest=false.")
    parser.add_argument("--start-period", default="", help="Outstanding start period, YYYYMM.")
    parser.add_argument("--end-period", default="", help="Outstanding end period, YYYYMM.")
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
    return parser.parse_args()


def load_config() -> dict[str, Any]:
    with CONFIG_PATH.open("r", encoding="utf-8") as fh:
        return json.load(fh)


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


def selected_datasets(args: argparse.Namespace, config: dict[str, Any]) -> list[str]:
    raw = args.datasets.strip()
    if raw:
        datasets = [item.strip() for item in raw.split(",") if item.strip()]
    else:
        datasets = list(config["presets"][args.dataset_preset])
    if "profiles" not in datasets:
        datasets.insert(0, "profiles")
    unknown = [dataset for dataset in datasets if dataset not in ENDPOINTS]
    if unknown:
        raise ValueError(f"Unknown datasets: {', '.join(unknown)}")
    return list(dict.fromkeys(datasets))


def normalized_text(value: Any) -> str:
    if value is None:
        return ""
    if isinstance(value, (dict, list)):
        return json.dumps(value, ensure_ascii=False, separators=(",", ":"))
    return str(value).strip()


def compact_join(values: list[Any], sep: str = " | ", limit: int = 6) -> str:
    out: list[str] = []
    seen: set[str] = set()
    for value in values:
        text = normalized_text(value)
        if not text or text in seen:
            continue
        seen.add(text)
        out.append(text)
        if limit and len(out) >= limit:
            break
    return sep.join(out)


def pair_join(row: dict[str, Any], keys: list[str]) -> str:
    parts = [normalized_text(row.get(key)) for key in keys]
    return " ".join(part for part in parts if part)


def row_date(row: dict[str, Any]) -> str:
    return normalized_text(
        row.get("last_upd_date")
        or row.get("as_of_date")
        or row.get("start_date")
        or row.get("nav_date")
        or row.get("period")
    )


def by_proj(rows: list[dict[str, Any]]) -> dict[str, list[dict[str, Any]]]:
    grouped: dict[str, list[dict[str, Any]]] = {}
    for row in rows:
        proj_id = normalized_text(row.get("proj_id"))
        if proj_id:
            grouped.setdefault(proj_id, []).append(row)
    return grouped


def latest_by_proj(rows: list[dict[str, Any]]) -> dict[str, dict[str, Any]]:
    grouped = by_proj(rows)
    return {
        proj_id: sorted(project_rows, key=row_date, reverse=True)[0]
        for proj_id, project_rows in grouped.items()
        if project_rows
    }


def latest_by_unique_id(rows: list[dict[str, Any]]) -> dict[str, dict[str, Any]]:
    lookup: dict[str, dict[str, Any]] = {}
    for row in rows:
        unique_id = normalized_text(row.get("unique_id"))
        if not unique_id:
            continue
        current = lookup.get(unique_id)
        if current is None or row_date(row) >= row_date(current):
            lookup[unique_id] = row
    return lookup


def add_column(headers: list[str], row: dict[str, Any], name: str, value: Any) -> None:
    if name not in headers:
        headers.append(name)
    row[name] = normalized_text(value)


SummaryFn = Callable[[list[str], dict[str, Any], dict[str, list[dict[str, Any]]]], None]


def summarize_amcs(headers: list[str], row: dict[str, Any], rows_by_dataset: dict[str, list[dict[str, Any]]]) -> None:
    lookup = latest_by_unique_id(rows_by_dataset.get("amcs", []))
    source = lookup.get(normalized_text(row.get("unique_id")), {})
    add_column(headers, row, "amc_name_th", source.get("comp_name_th"))
    add_column(headers, row, "amc_name_en", source.get("comp_name_en"))


def summarize_specifications(headers: list[str], row: dict[str, Any], rows_by_dataset: dict[str, list[dict[str, Any]]]) -> None:
    grouped = by_proj(rows_by_dataset.get("specifications", []))
    values = [item.get("spec_desc") for item in grouped.get(normalized_text(row.get("proj_id")), [])]
    add_column(headers, row, "specifications", compact_join(values))


def summarize_mutual_fund_fees(headers: list[str], row: dict[str, Any], rows_by_dataset: dict[str, list[dict[str, Any]]]) -> None:
    grouped = by_proj(rows_by_dataset.get("mutual_fund_fees", []))
    values = [
        pair_join(item, ["fee_type_desc", "rate", "rate_unit"])
        for item in grouped.get(normalized_text(row.get("proj_id")), [])
    ]
    add_column(headers, row, "mutual_fund_fees", compact_join(values))


def summarize_involve_parties(headers: list[str], row: dict[str, Any], rows_by_dataset: dict[str, list[dict[str, Any]]]) -> None:
    grouped = by_proj(rows_by_dataset.get("involve_parties", []))
    values = [
        pair_join(item, ["entity_type_label", "entity_name_th"])
        for item in grouped.get(normalized_text(row.get("proj_id")), [])
    ]
    add_column(headers, row, "involve_parties", compact_join(values))


def summarize_factsheet_urls(headers: list[str], row: dict[str, Any], rows_by_dataset: dict[str, list[dict[str, Any]]]) -> None:
    latest = latest_by_proj(rows_by_dataset.get("factsheet_urls", [])).get(normalized_text(row.get("proj_id")), {})
    add_column(headers, row, "factsheet_url_latest", latest.get("pdf_factsheet") or latest.get("amc_url_factsheet"))
    add_column(headers, row, "factsheet_as_of_date", latest.get("as_of_date"))


def summarize_ipos(headers: list[str], row: dict[str, Any], rows_by_dataset: dict[str, list[dict[str, Any]]]) -> None:
    latest = latest_by_proj(rows_by_dataset.get("ipos", [])).get(normalized_text(row.get("proj_id")), {})
    add_column(headers, row, "ipo_period", pair_join(latest, ["first_sell_start_date", "first_sell_end_date"]))


def summarize_benchmarks(headers: list[str], row: dict[str, Any], rows_by_dataset: dict[str, list[dict[str, Any]]]) -> None:
    grouped = by_proj(rows_by_dataset.get("benchmarks", []))
    values = [item.get("benchmark") for item in grouped.get(normalized_text(row.get("proj_id")), [])]
    add_column(headers, row, "benchmark_latest", compact_join(values))


def summarize_minimums(headers: list[str], row: dict[str, Any], rows_by_dataset: dict[str, list[dict[str, Any]]]) -> None:
    latest = latest_by_proj(rows_by_dataset.get("subscription_redemption_minimums", [])).get(normalized_text(row.get("proj_id")), {})
    add_column(headers, row, "minimum_sub", latest.get("minimum_sub"))
    add_column(headers, row, "minimum_redempt", latest.get("minimum_redempt"))
    add_column(headers, row, "lowbal_val", latest.get("lowbal_val"))


def summarize_periods(headers: list[str], row: dict[str, Any], rows_by_dataset: dict[str, list[dict[str, Any]]]) -> None:
    grouped = by_proj(rows_by_dataset.get("subscription_redemption_periods", []))
    values = [
        pair_join(item, ["type", "period", "settlement_period"])
        for item in grouped.get(normalized_text(row.get("proj_id")), [])
    ]
    add_column(headers, row, "subscription_redemption_periods", compact_join(values))


def summarize_risk(headers: list[str], row: dict[str, Any], rows_by_dataset: dict[str, list[dict[str, Any]]]) -> None:
    latest = latest_by_proj(rows_by_dataset.get("risk_spectrum", [])).get(normalized_text(row.get("proj_id")), {})
    add_column(headers, row, "risk_spectrum", latest.get("risk_spectrum"))
    add_column(headers, row, "risk_spectrum_desc", latest.get("risk_spectrum_desc"))


def summarize_statistics(headers: list[str], row: dict[str, Any], rows_by_dataset: dict[str, list[dict[str, Any]]]) -> None:
    latest = latest_by_proj(rows_by_dataset.get("statistics", [])).get(normalized_text(row.get("proj_id")), {})
    for column in ["maximum_drawdown", "sharpe_ratio", "beta", "alpha", "tracking_error", "yield_to_maturity"]:
        add_column(headers, row, column, latest.get(column))


def summarize_dividend_policy(headers: list[str], row: dict[str, Any], rows_by_dataset: dict[str, list[dict[str, Any]]]) -> None:
    latest = latest_by_proj(rows_by_dataset.get("dividend_policy", [])).get(normalized_text(row.get("proj_id")), {})
    add_column(headers, row, "dividend_policy", latest.get("dividend_policy"))


def summarize_fees(headers: list[str], row: dict[str, Any], rows_by_dataset: dict[str, list[dict[str, Any]]]) -> None:
    grouped = by_proj(rows_by_dataset.get("fees", []))
    values = [
        pair_join(item, ["fee_type_desc", "actual_value", "rate"])
        for item in grouped.get(normalized_text(row.get("proj_id")), [])
    ]
    add_column(headers, row, "factsheet_fees", compact_join(values))


def summarize_performance(headers: list[str], row: dict[str, Any], rows_by_dataset: dict[str, list[dict[str, Any]]]) -> None:
    grouped = by_proj(rows_by_dataset.get("performance", []))
    values = [
        pair_join(item, ["reference_period", "performance_value"])
        for item in grouped.get(normalized_text(row.get("proj_id")), [])
    ]
    add_column(headers, row, "performance_latest", compact_join(values, limit=10))


def summarize_asset_allocation(headers: list[str], row: dict[str, Any], rows_by_dataset: dict[str, list[dict[str, Any]]]) -> None:
    grouped = by_proj(rows_by_dataset.get("asset_allocation", []))
    values = [
        pair_join(item, ["asset_name", "asset_ratio"])
        for item in grouped.get(normalized_text(row.get("proj_id")), [])
    ]
    add_column(headers, row, "asset_allocation", compact_join(values, limit=8))


def summarize_top5(headers: list[str], row: dict[str, Any], rows_by_dataset: dict[str, list[dict[str, Any]]]) -> None:
    grouped = by_proj(rows_by_dataset.get("top5_holdings", []))
    values = [
        pair_join(item, ["asset_name", "asset_ratio"])
        for item in grouped.get(normalized_text(row.get("proj_id")), [])
    ]
    add_column(headers, row, "top5_holdings", compact_join(values, limit=5))


def summarize_outstanding_portfolio(headers: list[str], row: dict[str, Any], rows_by_dataset: dict[str, list[dict[str, Any]]]) -> None:
    project_rows = by_proj(rows_by_dataset.get("outstanding_portfolio", [])).get(normalized_text(row.get("proj_id")), [])
    add_column(headers, row, "outstanding_portfolio_rows", len(project_rows) if project_rows else "")
    add_column(headers, row, "outstanding_portfolio_latest_period", compact_join([item.get("period") for item in project_rows], limit=1))


def summarize_portfolio_asset_type(headers: list[str], row: dict[str, Any], rows_by_dataset: dict[str, list[dict[str, Any]]]) -> None:
    grouped = by_proj(rows_by_dataset.get("portfolio_asset_type", []))
    values = [
        pair_join(item, ["period", "assetliab_desc", "percent_nav"])
        for item in grouped.get(normalized_text(row.get("proj_id")), [])
    ]
    add_column(headers, row, "portfolio_asset_type", compact_join(values, limit=8))


def summarize_nav_daily(headers: list[str], row: dict[str, Any], rows_by_dataset: dict[str, list[dict[str, Any]]]) -> None:
    latest = latest_by_proj(rows_by_dataset.get("nav_daily", [])).get(normalized_text(row.get("proj_id")), {})
    add_column(headers, row, "nav_latest_date", latest.get("nav_date"))
    add_column(headers, row, "nav_latest", latest.get("last_val"))


def summarize_dividend_history(headers: list[str], row: dict[str, Any], rows_by_dataset: dict[str, list[dict[str, Any]]]) -> None:
    latest = latest_by_proj(rows_by_dataset.get("dividend_history", [])).get(normalized_text(row.get("proj_id")), {})
    add_column(headers, row, "dividend_latest_date", latest.get("dividend_date"))
    add_column(headers, row, "dividend_latest_value", latest.get("dividend_value"))


SUMMARY_FUNCTIONS: dict[str, SummaryFn] = {
    "amcs": summarize_amcs,
    "specifications": summarize_specifications,
    "mutual_fund_fees": summarize_mutual_fund_fees,
    "involve_parties": summarize_involve_parties,
    "factsheet_urls": summarize_factsheet_urls,
    "ipos": summarize_ipos,
    "benchmarks": summarize_benchmarks,
    "subscription_redemption_minimums": summarize_minimums,
    "subscription_redemption_periods": summarize_periods,
    "risk_spectrum": summarize_risk,
    "statistics": summarize_statistics,
    "dividend_policy": summarize_dividend_policy,
    "fees": summarize_fees,
    "performance": summarize_performance,
    "asset_allocation": summarize_asset_allocation,
    "top5_holdings": summarize_top5,
    "outstanding_portfolio": summarize_outstanding_portfolio,
    "portfolio_asset_type": summarize_portfolio_asset_type,
    "nav_daily": summarize_nav_daily,
    "dividend_history": summarize_dividend_history,
}


def fetch_dataset(
    dataset: str,
    api_key: str,
    args: argparse.Namespace,
    registered_ids: list[str],
    errors: list[dict[str, Any]],
) -> list[dict[str, Any]]:
    if dataset == "profiles":
        return fetch_registered_profiles(api_key, args)

    endpoint = ENDPOINTS[dataset]
    if dataset in PROJECT_ID_DATASETS and registered_ids and not getattr(args, "proj_id", ""):
        raw_rows = fetch_dataset_for_registered_proj_ids(
            dataset,
            endpoint,
            registered_ids,
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
    rows, _ = transform_dataset_rows(dataset, raw_rows)
    return rows


def fetch_profile_rows(api_key: str, args: argparse.Namespace) -> list[dict[str, Any]]:
    if not args.proj_id.strip():
        return fetch_registered_profiles(api_key, args)

    params = {
        "page_size": args.page_size,
        "project_info": args.proj_id.strip(),
    }
    add_if_present(params, "fund_status", args.fund_status)
    raw_rows = fetch_sec_all_pages(
        ENDPOINTS["profiles"],
        api_key,
        query_params=params,
        max_pages=args.max_pages,
    )
    rows, _ = transform_dataset_rows("profiles", raw_rows)
    return rows


def build_master_view(
    rows_by_dataset: dict[str, list[dict[str, Any]]],
    datasets: list[str],
    base_columns: list[str],
) -> tuple[list[str], list[dict[str, Any]]]:
    headers = list(base_columns)
    output_rows: list[dict[str, Any]] = []

    for profile in rows_by_dataset.get("profiles", []):
        row = {
            "proj_id": normalized_text(profile.get("proj_id")),
            "fund_class_name": normalized_text(profile.get("fund_class_name")),
            "fund_name": normalized_text(
                profile.get("proj_abbr_name")
                or profile.get("proj_name_th")
                or profile.get("proj_name_en")
            ),
            "status": normalized_text(profile.get("fund_status")),
            "last_upd_date": normalized_text(profile.get("last_upd_date")),
            "unique_id": normalized_text(profile.get("unique_id")),
        }

        if "unique_id" not in headers:
            headers.append("unique_id")

        for dataset in datasets:
            if dataset == "profiles":
                continue
            summary_fn = SUMMARY_FUNCTIONS.get(dataset)
            if summary_fn:
                summary_fn(headers, row, rows_by_dataset)
        output_rows.append(row)

    return headers, output_rows


def values_from_rows(headers: list[str], rows: list[dict[str, Any]]) -> list[list[Any]]:
    return [headers] + [[row.get(header, "") for header in headers] for row in rows]


def ensure_sheet(sheets: Any, spreadsheet_id: str, tab_name: str) -> None:
    metadata = sheets.spreadsheets().get(spreadsheetId=spreadsheet_id).execute()
    existing = {sheet["properties"]["title"] for sheet in metadata.get("sheets", [])}
    if tab_name in existing:
        return
    sheets.spreadsheets().batchUpdate(
        spreadsheetId=spreadsheet_id,
        body={
            "requests": [
                {
                    "addSheet": {
                        "properties": {
                            "title": tab_name,
                        }
                    }
                }
            ]
        },
    ).execute()


def write_values_to_sheet(sheets: Any, spreadsheet_id: str, tab_name: str, values: list[list[Any]]) -> None:
    ensure_sheet(sheets, spreadsheet_id, tab_name)
    sheets.spreadsheets().values().clear(
        spreadsheetId=spreadsheet_id,
        range=f"'{tab_name}'",
        body={},
    ).execute()
    if not values:
        return

    for start in range(0, len(values), 1000):
        chunk = values[start : start + 1000]
        range_name = f"'{tab_name}'!A{start + 1}"
        sheets.spreadsheets().values().update(
            spreadsheetId=spreadsheet_id,
            range=range_name,
            valueInputOption="RAW",
            body={"values": chunk},
        ).execute()
        print(f"Wrote Google Sheet rows {start + 1} - {start + len(chunk)}")


def main() -> int:
    args = parse_args()
    if not args.spreadsheet_id.strip():
        raise RuntimeError("Missing --spreadsheet-id or SEC_MASTER_VIEW_SPREADSHEET_ID.")

    config = load_config()
    tab_name = args.tab_name.strip() or config["sheet"]["default_tab_name"]
    datasets = selected_datasets(args, config)
    api_key = get_api_key()
    errors: list[dict[str, Any]] = []

    profile_rows = fetch_profile_rows(api_key, args)
    registered_ids = registered_proj_ids(profile_rows, args.registered_max_funds)
    if args.registered_max_funds and registered_ids:
        registered_set = set(registered_ids)
        profile_rows = [
            row for row in profile_rows if normalized_text(row.get("proj_id")) in registered_set
        ]
    rows_by_dataset: dict[str, list[dict[str, Any]]] = {"profiles": profile_rows}
    print(f"Profiles rows: {len(profile_rows)}")
    print(f"Registered proj_id values: {len(registered_ids)}")
    print(f"Selected datasets: {', '.join(datasets)}")

    for dataset in datasets:
        if dataset == "profiles":
            continue
        try:
            rows_by_dataset[dataset] = fetch_dataset(dataset, api_key, args, registered_ids, errors)
            print(f"Dataset {dataset}: {len(rows_by_dataset[dataset])} rows")
        except Exception as exc:
            errors.append({"dataset": dataset, "error": str(exc)})
            print(f"ERROR fetching {dataset}: {exc}")
            if not should_continue_on_error(args):
                raise
            rows_by_dataset[dataset] = []

    headers, rows = build_master_view(
        rows_by_dataset,
        datasets,
        base_columns=config["base_columns"],
    )
    print(f"Master view rows: {len(rows)}")
    print(f"Master view columns: {len(headers)}")

    from googleapiclient.discovery import build

    credentials = credentials_from_env()
    sheets = build("sheets", "v4", credentials=credentials, cache_discovery=False)
    write_values_to_sheet(
        sheets=sheets,
        spreadsheet_id=args.spreadsheet_id.strip(),
        tab_name=tab_name,
        values=values_from_rows(headers, rows),
    )

    if errors:
        error_values = values_from_rows(["dataset", "error"], errors)
        write_values_to_sheet(
            sheets=sheets,
            spreadsheet_id=args.spreadsheet_id.strip(),
            tab_name=f"{tab_name}_errors",
            values=error_values,
        )

    return 0


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except Exception as exc:
        print(f"ERROR: {exc}")
        raise SystemExit(1)
