#!/usr/bin/env python3
"""Fetch SEC fund API datasets and export CSV files.

This script is designed for GitHub Actions first-pass testing. It reads the
SEC API key from SEC_API_KEY, fetches SEC fund endpoints, and writes one CSV
file per dataset into an output directory.
"""

from __future__ import annotations

import argparse
import csv
import json
import os
import re
import sys
from datetime import date
from pathlib import Path
from typing import Any

import requests


BASE_URL = "https://api.sec.or.th"


ENDPOINTS = {
    "amcs": "/v2/fund/general-info/amcs",
    "profiles": "/v2/fund/general-info/profiles",
    "specifications": "/v2/fund/general-info/specifications",
    "mutual_fund_fees": "/v2/fund/general-info/mutual-fund-fees",
    "involve_parties": "/v2/fund/general-info/involve-parties",
    "factsheet_urls": "/v2/fund/factsheet/urls",
    "ipos": "/v2/fund/factsheet/ipos",
    "benchmarks": "/v2/fund/factsheet/benchmarks",
    "subscription_redemption_minimums": "/v2/fund/factsheet/subscription-redemption-minimums",
    "subscription_redemption_periods": "/v2/fund/factsheet/subscription-redemption-periods",
    "risk_spectrum": "/v2/fund/factsheet/risk-spectrum",
    "statistics": "/v2/fund/factsheet/statistics",
    "dividend_policy": "/v2/fund/factsheet/dividend-policy",
    "fees": "/v2/fund/factsheet/fees",
    "performance": "/v2/fund/factsheet/performance",
    "asset_allocation": "/v2/fund/factsheet/asset-allocation",
    "top5_holdings": "/v2/fund/factsheet/top5-holdings",
    "outstanding_portfolio": "/v2/fund/outstanding/portfolio",
    "portfolio_asset_type": "/v2/fund/outstanding/portfolio-asset-type",
    "nav_daily": "/v2/fund/daily-info/nav",
    "dividend_history": "/v2/fund/daily-info/dividend-history",
}

PROJECT_ID_DATASETS = {
    "mutual_fund_fees",
    "involve_parties",
    "factsheet_urls",
    "ipos",
    "benchmarks",
    "subscription_redemption_minimums",
    "subscription_redemption_periods",
    "risk_spectrum",
    "statistics",
    "dividend_policy",
    "fees",
    "performance",
    "asset_allocation",
    "top5_holdings",
    "outstanding_portfolio",
    "portfolio_asset_type",
    "nav_daily",
    "dividend_history",
}

DATASET_FILES = {
    "amcs": "01_amcs.csv",
    "profiles": "02_profiles.csv",
    "specifications": "03_specifications.csv",
    "mutual_fund_fees": "04_mutual_fund_fees.csv",
    "involve_parties": "05_involve_parties.csv",
    "factsheet_urls": "06_factsheet_urls.csv",
    "ipos": "07_ipos.csv",
    "benchmarks": "08_benchmarks.csv",
    "subscription_redemption_minimums": "09_subscription_redemption_minimums.csv",
    "subscription_redemption_periods": "10_subscription_redemption_periods.csv",
    "risk_spectrum": "11_risk_spectrum.csv",
    "statistics": "12_statistics.csv",
    "dividend_policy": "13_dividend_policy.csv",
    "fees": "14_fees.csv",
    "performance": "15_performance.csv",
    "asset_allocation": "16_asset_allocation.csv",
    "top5_holdings": "17_top5_holdings.csv",
    "outstanding_portfolio": "18_outstanding_portfolio.csv",
    "portfolio_asset_type": "19_portfolio_asset_type.csv",
    "nav_daily": "20_nav_daily.csv",
    "dividend_history": "21_dividend_history.csv",
}

PROFILE_COLUMNS = [
    "unique_id",
    "comp_name_th",
    "comp_name_en",
    "proj_id",
    "regis_id",
    "proj_name_th",
    "proj_name_en",
    "proj_abbr_name",
    "fund_status",
    "fund_status_label",
    "init_date",
    "regis_date",
    "cancel_date",
    "invest_country_flag",
    "invest_country_flag_label",
    "proj_retail_type",
    "proj_retail_type_label",
    "proj_term_flag",
    "proj_term_flag_label",
    "proj_term_year",
    "proj_term_month",
    "proj_term_day",
    "policy_desc",
    "investment_policy_desc",
    "management_style",
    "management_style_label",
    "feederfund_master_fund",
    "feederfund_country",
    "exchange_rate_protection_policy",
    "fund_class_name",
    "fund_class_detail",
    "fund_class_description",
    "fund_class_tax_incentive_type",
    "fund_class_isin_code",
    "last_upd_date",
]

FUND_STATUS_LABELS = {
    "Registered": "จดทะเบียน",
    "IPO": "เสนอขายหน่วยลงทุนครั้งแรก",
    "Expired": "หมดเวลาเสนอขาย",
    "Canceled": "เลิกโครงการ",
    "Liquidated": "จดทะเบียนเลิก",
}

INVEST_COUNTRY_FLAG_LABELS = {
    "1": "เน้นลงทุนต่างประเทศ",
    "2": "ลงทุนในต่างประเทศบางส่วน",
    "3": "ไม่มีความเสี่ยงต่างประเทศ",
    "4": "มีความเสี่ยงทั้งในและต่างประเทศ",
}

PROJ_RETAIL_TYPE_LABELS = {
    "A": "กองทุนรวมที่เสนอขายเฉพาะผู้ลงทุนที่มิใช่รายย่อย",
    "B": "กองทุนรวมที่เสนอขายเฉพาะผู้มีเงินลงทุนสูง",
    "F": "กองทุนรวมเสริมสภาพคล่องเพื่อลดความเสี่ยงของการระดมทุนในตลาดตราสารหนี้ภาคเอกชน",
    "G": "กองทุนรวมพิเศษเพื่อตอบสนองนโยบายภาครัฐ",
    "H": "กองทุนรวมที่เสนอขายผู้ลงทุนที่มิใช่รายย่อยและผู้มีเงินลงทุนสูง",
    "N": "กองทุนเพื่อผู้ลงทุนสถาบัน",
    "R": "กองทุนเพื่อผู้ลงทุนทั่วไป",
    "V": "กองทุนรวมเพื่อผู้ลงทุนที่เป็นกองทุนสำรองเลี้ยงชีพ",
    "X": "กองทุนรวมที่เสนอขายผู้ลงทุนสถาบันและผู้ลงทุนรายใหญ่พิเศษ",
}

PROJ_TERM_FLAG_LABELS = {
    "Y": "กำหนดอายุโครงการ",
    "N": "ไม่กำหนดอายุโครงการ",
}

MANAGEMENT_STYLE_LABELS = {
    "AM": "active management",
    "AN": "กองทุนไทยตามกองทุนหลัก และกองทุนหลัก active management",
    "PM": "passive management / index tracking",
    "PN": "กองทุนไทยตามกองทุนหลัก และกองทุนหลัก passive management",
    "IM": "inverse management",
    "IN": "กองทุนไทยตามกองทุนหลัก และกองทุนหลัก inverse management",
    "LM": "leveraged management",
    "LN": "กองทุนไทยตามกองทุนหลัก และกองทุนหลัก leveraged management",
    "BH": "buy-and-hold",
    "SM": "index tracking และบางโอกาสอาจสร้างผลตอบแทนสูงกว่าดัชนี",
    "OT": "อื่น ๆ",
}

SPECIFICATION_COLUMNS = [
    "proj_id",
    "fund_class_name",
    "spec_code",
    "spec_desc",
    "last_upd_date",
]

MUTUAL_FUND_FEE_COLUMNS = [
    "proj_id",
    "fund_class_name",
    "fee_type_desc",
    "rate",
    "rate_unit",
    "fee_other_desc",
    "last_upd_date",
]

INVOLVE_PARTY_COLUMNS = [
    "proj_id",
    "entity_type",
    "entity_type_label",
    "entity_name_th",
    "entity_name_en",
    "address",
    "last_upd_date",
]

ENTITY_TYPE_LABELS = {
    "A": "ผู้สอบบัญชี",
    "U": "ผู้จัดจำหน่าย",
    "S": "ผู้สนับสนุนการขายและรับซื้อคืน",
    "R": "นายทะเบียนหน่วยลงทุน",
    "V": "ผู้ดูแลผลประโยชน์",
    "M": "ที่ปรึกษาการลงทุน",
    "O": "ผู้รับมอบหมายงานด้านการจัดการลงทุน",
    "P": "ผู้ลงทุนรายใหญ่",
    "K": "ผู้ดูแลสภาพคล่อง",
    "N": "ที่ปรึกษาทางการเงิน",
    "F": "ผู้จัดการกองทุน",
}

FACTSHEET_URL_COLUMNS = [
    "proj_id",
    "fund_class_name",
    "prospectus_type",
    "prospectus_type_label",
    "amc_url_factsheet",
    "pdf_factsheet",
    "as_of_date",
    "last_upd_date",
]

FACTSHEET_IPO_COLUMNS = [
    "proj_id",
    "start_date",
    "end_date",
    "prospectus_type",
    "prospectus_type_label",
    "first_sell_start_date",
    "first_sell_end_date",
    "last_upd_date",
]

BENCHMARK_COLUMNS = [
    "proj_id",
    "start_date",
    "end_date",
    "prospectus_type",
    "prospectus_type_label",
    "group_seq",
    "benchmark",
    "remark",
    "last_upd_date",
]

SUBSCRIPTION_REDEMPTION_MINIMUM_COLUMNS = [
    "proj_id",
    "fund_class_name",
    "start_date",
    "end_date",
    "prospectus_type",
    "prospectus_type_label",
    "minimum_sub_ipo",
    "minimum_sub_ipo_cur",
    "minimum_sub",
    "minimum_sub_cur",
    "minimum_sub_unit",
    "minimum_redempt",
    "minimum_redempt_cur",
    "minimum_redempt_unit",
    "lowbal_val",
    "lowbal_val_cur",
    "lowbal_unit",
    "last_upd_date",
]

SUBSCRIPTION_REDEMPTION_PERIOD_COLUMNS = [
    "proj_id",
    "fund_class_name",
    "start_date",
    "end_date",
    "prospectus_type",
    "prospectus_type_label",
    "type",
    "period",
    "redemp_period_oth",
    "settlement_period",
    "last_upd_date",
]

RISK_SPECTRUM_COLUMNS = [
    "proj_id",
    "start_date",
    "end_date",
    "prospectus_type",
    "prospectus_type_label",
    "risk_spectrum",
    "risk_spectrum_desc",
    "last_upd_date",
]

STATISTICS_COLUMNS = [
    "proj_id",
    "fund_class_name",
    "start_date",
    "end_date",
    "prospectus_type",
    "prospectus_type_label",
    "portfolio_turnover_ratio",
    "recovering_period",
    "portfolio_duration_period",
    "maximum_drawdown",
    "sharpe_ratio",
    "beta",
    "alpha",
    "fx_hedging",
    "tracking_error",
    "yield_to_maturity",
    "last_upd_date",
]

DIVIDEND_POLICY_COLUMNS = [
    "proj_id",
    "fund_class_name",
    "start_date",
    "end_date",
    "prospectus_type",
    "prospectus_type_label",
    "dividend_policy",
    "last_upd_date",
]

FACTSHEET_FEE_COLUMNS = [
    "proj_id",
    "fund_class_name",
    "start_date",
    "end_date",
    "prospectus_type",
    "prospectus_type_label",
    "fee_type_desc",
    "rate",
    "actual_value",
    "fee_other_desc",
    "last_upd_date",
]

PERFORMANCE_COLUMNS = [
    "proj_id",
    "fund_class_name",
    "start_date",
    "end_date",
    "prospectus_type",
    "prospectus_type_label",
    "performance_type_desc",
    "reference_period",
    "performance_value",
    "last_upd_date",
]

ASSET_ALLOCATION_COLUMNS = [
    "proj_id",
    "start_date",
    "end_date",
    "prospectus_type",
    "prospectus_type_label",
    "asset_seq",
    "asset_name",
    "asset_ratio",
    "last_upd_date",
]

TOP5_HOLDING_COLUMNS = [
    "proj_id",
    "start_date",
    "end_date",
    "prospectus_type",
    "prospectus_type_label",
    "asset_seq",
    "asset_name",
    "asset_ratio",
    "last_upd_date",
]

OUTSTANDING_PORTFOLIO_COLUMNS = [
    "proj_id",
    "period",
    "as_of_date",
    "assetliab_code",
    "assetliab_desc",
    "issue_code",
    "isin_code",
    "issuer",
    "market_value",
    "percent_nav",
    "last_upd_date",
]

PORTFOLIO_ASSET_TYPE_COLUMNS = [
    "proj_id",
    "period",
    "assetliab_code",
    "assetliab_desc",
    "market_value",
    "percent_nav",
]

NAV_DAILY_COLUMNS = [
    "proj_id",
    "nav_date",
    "fund_class_name",
    "net_asset",
    "last_val",
    "unique_id",
    "sell_price",
    "buy_price",
    "sell_swap_price",
    "buy_swap_price",
    "last_upd_date",
]

DIVIDEND_HISTORY_COLUMNS = [
    "unique_id",
    "proj_id",
    "class_abbr_name",
    "book_close_date",
    "dividend_date",
    "dividend_value",
    "last_upd_date",
]

PROSPECTUS_TYPE_LABELS = {
    "IPO": "ส่งเมื่อยื่นขอจัดตั้งกองทุน",
    "Monthly": "ส่งรายเดือน",
    "SignificantFactsheet": "ส่งเมื่อมีการเปลี่ยนแปลงข้อมูลอย่างมีนัยสำคัญ",
}


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Fetch SEC fund data and export CSV files.")
    parser.add_argument("--output-dir", default="sec_output", help="Directory for CSV output.")
    parser.add_argument("--proj-id", default="", help="Optional project ID for fund-specific endpoints.")
    parser.add_argument(
        "--proj-ids",
        default="",
        help="Optional project IDs separated by comma, whitespace, or new lines.",
    )
    parser.add_argument("--fund-status", default="Registered", help="Optional fund_status filter for profiles.")
    parser.add_argument(
        "--registered-only",
        choices=["true", "false"],
        default="false",
        help="Fetch profiles first and use only registered proj_id values for project-scoped datasets.",
    )
    parser.add_argument(
        "--registered-max-funds",
        type=int,
        default=0,
        help="Limit registered proj_id values for testing. Use 0 for no limit.",
    )
    parser.add_argument("--fund-class-name", default="", help="Optional SEC fund_class_name filter.")
    parser.add_argument("--start-date", default="", help="Factsheet start_date filter, YYYY-MM-DD.")
    parser.add_argument("--end-date", default="", help="Factsheet end_date filter, YYYY-MM-DD.")
    parser.add_argument("--latest", default="true", help="Use latest factsheet data for supported endpoints.")
    parser.add_argument("--start-period", default="", help="Outstanding data start period, YYYYMM.")
    parser.add_argument("--end-period", default="", help="Outstanding data end period, YYYYMM.")
    parser.add_argument("--start-nav-date", default="2023-10-01", help="NAV start date, YYYY-MM-DD.")
    parser.add_argument("--end-nav-date", default="2023-10-31", help="NAV end date, YYYY-MM-DD.")
    parser.add_argument("--page-size", type=int, default=100, help="API page_size value.")
    parser.add_argument(
        "--output-layout",
        choices=["flat", "latest-snapshot"],
        default="flat",
        help="Use flat files or write latest/ and snapshots/YYYY-MM-DD/ copies.",
    )
    parser.add_argument("--snapshot-date", default="", help="Snapshot folder date, YYYY-MM-DD.")
    parser.add_argument(
        "--output-group",
        default="",
        help="Optional output group folder, such as master, daily, or weekly.",
    )
    parser.add_argument(
        "--max-pages",
        type=int,
        default=3,
        help="Maximum pages per endpoint for this test run. Use 0 for no limit.",
    )
    parser.add_argument(
        "--datasets",
        default="all",
        help="Comma-separated dataset keys or all.",
    )
    parser.add_argument(
        "--write-curated",
        choices=["true", "false"],
        default="true",
        help="Write combined curated CSV files when possible.",
    )
    parser.add_argument(
        "--continue-on-error",
        choices=["true", "false"],
        default="true",
        help="Continue to the next dataset when an SEC endpoint returns an error.",
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


def should_continue_on_error(args: argparse.Namespace) -> bool:
    return str(args.continue_on_error).strip().lower() in {"1", "true", "yes", "y"}


def should_use_registered_only(args: argparse.Namespace) -> bool:
    return str(args.registered_only).strip().lower() in {"1", "true", "yes", "y"}


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


def ordered_headers(rows: list[dict[str, Any]], preferred: list[str] | None = None) -> list[str]:
    discovered = collect_headers(rows)
    if not preferred:
        return discovered

    preferred_set = set(preferred)
    return [header for header in preferred if header in discovered or header in preferred_set] + [
        header for header in discovered if header not in preferred_set
    ]


def write_csv(
    path: Path,
    rows: list[dict[str, Any]],
    preferred_headers: list[str] | None = None,
) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    headers = ordered_headers(rows, preferred_headers)
    if not headers:
        headers = ["message"]
        rows = [{"message": "No rows returned"}]

    with path.open("w", encoding="utf-8-sig", newline="") as fh:
        writer = csv.DictWriter(fh, fieldnames=headers, extrasaction="ignore")
        writer.writeheader()
        for row in rows:
            writer.writerow({key: flatten_value(row.get(key, "")) for key in headers})

    print(f"Wrote {path}: {len(rows)} rows")


def label_from_map(value: Any, labels: dict[str, str]) -> str:
    text = str(value or "").strip()
    return labels.get(text, "")


def enrich_profile_row(row: dict[str, Any]) -> dict[str, Any]:
    enriched = dict(row)
    enriched["fund_status_label"] = label_from_map(row.get("fund_status"), FUND_STATUS_LABELS)
    enriched["invest_country_flag_label"] = label_from_map(
        row.get("invest_country_flag"),
        INVEST_COUNTRY_FLAG_LABELS,
    )
    enriched["proj_retail_type_label"] = label_from_map(
        row.get("proj_retail_type"),
        PROJ_RETAIL_TYPE_LABELS,
    )
    enriched["proj_term_flag_label"] = label_from_map(
        row.get("proj_term_flag"),
        PROJ_TERM_FLAG_LABELS,
    )
    enriched["management_style_label"] = label_from_map(
        row.get("management_style"),
        MANAGEMENT_STYLE_LABELS,
    )
    return enriched


def enrich_involve_party_row(row: dict[str, Any]) -> dict[str, Any]:
    enriched = dict(row)
    enriched["entity_type_label"] = label_from_map(row.get("entity_type"), ENTITY_TYPE_LABELS)
    return enriched


def enrich_prospectus_type_row(row: dict[str, Any]) -> dict[str, Any]:
    enriched = dict(row)
    enriched["prospectus_type_label"] = label_from_map(
        row.get("prospectus_type"),
        PROSPECTUS_TYPE_LABELS,
    )
    return enriched


def transform_dataset_rows(dataset: str, rows: list[dict[str, Any]]) -> tuple[list[dict[str, Any]], list[str] | None]:
    if dataset == "profiles":
        return [enrich_profile_row(row) for row in rows], PROFILE_COLUMNS
    if dataset == "specifications":
        return rows, SPECIFICATION_COLUMNS
    if dataset == "mutual_fund_fees":
        return rows, MUTUAL_FUND_FEE_COLUMNS
    if dataset == "involve_parties":
        return [enrich_involve_party_row(row) for row in rows], INVOLVE_PARTY_COLUMNS
    if dataset == "factsheet_urls":
        return [enrich_prospectus_type_row(row) for row in rows], FACTSHEET_URL_COLUMNS
    if dataset == "ipos":
        return [enrich_prospectus_type_row(row) for row in rows], FACTSHEET_IPO_COLUMNS
    if dataset == "benchmarks":
        return [enrich_prospectus_type_row(row) for row in rows], BENCHMARK_COLUMNS
    if dataset == "subscription_redemption_minimums":
        return [enrich_prospectus_type_row(row) for row in rows], SUBSCRIPTION_REDEMPTION_MINIMUM_COLUMNS
    if dataset == "subscription_redemption_periods":
        return [enrich_prospectus_type_row(row) for row in rows], SUBSCRIPTION_REDEMPTION_PERIOD_COLUMNS
    if dataset == "risk_spectrum":
        return [enrich_prospectus_type_row(row) for row in rows], RISK_SPECTRUM_COLUMNS
    if dataset == "statistics":
        return [enrich_prospectus_type_row(row) for row in rows], STATISTICS_COLUMNS
    if dataset == "dividend_policy":
        return [enrich_prospectus_type_row(row) for row in rows], DIVIDEND_POLICY_COLUMNS
    if dataset == "fees":
        return [enrich_prospectus_type_row(row) for row in rows], FACTSHEET_FEE_COLUMNS
    if dataset == "performance":
        return [enrich_prospectus_type_row(row) for row in rows], PERFORMANCE_COLUMNS
    if dataset == "asset_allocation":
        return [enrich_prospectus_type_row(row) for row in rows], ASSET_ALLOCATION_COLUMNS
    if dataset == "top5_holdings":
        return [enrich_prospectus_type_row(row) for row in rows], TOP5_HOLDING_COLUMNS
    if dataset == "outstanding_portfolio":
        return rows, OUTSTANDING_PORTFOLIO_COLUMNS
    if dataset == "portfolio_asset_type":
        return rows, PORTFOLIO_ASSET_TYPE_COLUMNS
    if dataset == "nav_daily":
        return rows, NAV_DAILY_COLUMNS
    if dataset == "dividend_history":
        return rows, DIVIDEND_HISTORY_COLUMNS
    return rows, None


def selected_datasets(raw: str) -> list[str]:
    if raw.strip().lower() == "all":
        return list(ENDPOINTS.keys())
    keys = [item.strip() for item in raw.split(",") if item.strip()]
    unknown = [key for key in keys if key not in ENDPOINTS]
    if unknown:
        raise ValueError(f"Unknown datasets: {', '.join(unknown)}")
    return keys


def parse_project_ids(raw: str) -> list[str]:
    values: list[str] = []
    seen: set[str] = set()
    for item in re.split(r"[\s,;]+", str(raw or "").strip()):
        proj_id = item.strip()
        if not proj_id or proj_id in seen:
            continue
        seen.add(proj_id)
        values.append(proj_id)
    return values


def requested_proj_ids(args: argparse.Namespace) -> list[str]:
    raw_values = []
    if getattr(args, "proj_id", ""):
        raw_values.append(args.proj_id)
    if getattr(args, "proj_ids", ""):
        raw_values.append(args.proj_ids)
    return parse_project_ids(" ".join(raw_values))


def add_if_present(params: dict[str, Any], key: str, value: Any) -> None:
    if value is None:
        return
    text = str(value).strip()
    if text:
        params[key] = text


def factsheet_params(args: argparse.Namespace, include_fund_class: bool = True) -> dict[str, Any]:
    params: dict[str, Any] = {"page_size": args.page_size}
    add_if_present(params, "proj_id", args.proj_id)
    if include_fund_class:
        add_if_present(params, "fund_class_name", getattr(args, "fund_class_name", ""))

    latest = str(args.latest).strip().lower()
    if latest in {"1", "true", "yes", "y"}:
        params["latest"] = "true"
    else:
        add_if_present(params, "start_date", args.start_date)
        add_if_present(params, "end_date", args.end_date)
    return params


def project_params(args: argparse.Namespace, include_fund_class: bool = False) -> dict[str, Any]:
    params: dict[str, Any] = {"page_size": args.page_size}
    add_if_present(params, "proj_id", args.proj_id)
    if include_fund_class:
        add_if_present(params, "fund_class_name", getattr(args, "fund_class_name", ""))
    return params


def dataset_params(dataset: str, args: argparse.Namespace) -> dict[str, Any]:
    if dataset == "profiles":
        params = {"page_size": args.page_size}
        add_if_present(params, "fund_status", args.fund_status)
        return params
    if dataset == "mutual_fund_fees":
        return project_params(args, include_fund_class=True)
    if dataset == "factsheet_urls":
        return project_params(args, include_fund_class=True)
    if dataset in {"involve_parties", "dividend_history"}:
        return project_params(args)
    if dataset in {
        "ipos",
        "benchmarks",
        "subscription_redemption_minimums",
        "subscription_redemption_periods",
        "risk_spectrum",
        "statistics",
        "dividend_policy",
        "fees",
        "performance",
        "top5_holdings",
    }:
        return factsheet_params(args)
    if dataset == "asset_allocation":
        return factsheet_params(args, include_fund_class=False)
    if dataset == "outstanding_portfolio":
        params = {"page_size": args.page_size}
        add_if_present(params, "proj_id", args.proj_id)
        add_if_present(params, "start_period", args.start_period)
        add_if_present(params, "end_period", args.end_period)
        return params
    if dataset == "portfolio_asset_type":
        params = {"page_size": args.page_size}
        add_if_present(params, "proj_id", args.proj_id)
        add_if_present(params, "start_period", args.start_period)
        add_if_present(params, "end_period", args.end_period)
        return params
    if dataset == "nav_daily":
        params = {
            "start_nav_date": args.start_nav_date,
            "end_nav_date": args.end_nav_date,
            "page_size": args.page_size,
        }
        add_if_present(params, "proj_id", args.proj_id)
        add_if_present(params, "fund_class_name", args.fund_class_name)
        return params
    return {"page_size": args.page_size}


def fetch_registered_profiles(
    api_key: str,
    args: argparse.Namespace,
) -> list[dict[str, Any]]:
    params = {"page_size": args.page_size}
    add_if_present(params, "fund_status", args.fund_status)
    rows = fetch_sec_all_pages(
        ENDPOINTS["profiles"],
        api_key,
        query_params=params,
        max_pages=args.max_pages,
    )
    rows, _ = transform_dataset_rows("profiles", rows)
    return rows


def fetch_profiles_for_proj_ids(
    proj_ids: list[str],
    api_key: str,
    args: argparse.Namespace,
    errors: list[dict[str, Any]] | None = None,
) -> list[dict[str, Any]]:
    rows: list[dict[str, Any]] = []
    total = len(proj_ids)
    for index, proj_id in enumerate(proj_ids, start=1):
        params = {"page_size": args.page_size, "project_info": proj_id}
        add_if_present(params, "fund_status", args.fund_status)
        try:
            project_rows = fetch_sec_all_pages(
                ENDPOINTS["profiles"],
                api_key,
                query_params=params,
                max_pages=args.max_pages,
            )
            rows.extend(project_rows)
            print(f"Fetched profiles for proj_id {index}/{total}: {proj_id}")
        except Exception as exc:
            if errors is not None:
                errors.append(
                    {
                        "dataset": "profiles",
                        "endpoint": ENDPOINTS["profiles"],
                        "proj_id": proj_id,
                        "params": json.dumps(params, ensure_ascii=False, separators=(",", ":")),
                        "error": str(exc),
                    }
                )
            print(f"ERROR fetching profiles for proj_id={proj_id}: {exc}", file=sys.stderr)
            if not should_continue_on_error(args):
                raise
    rows, _ = transform_dataset_rows("profiles", rows)
    return rows


def registered_proj_ids(profile_rows: list[dict[str, Any]], limit: int = 0) -> list[str]:
    proj_ids: list[str] = []
    seen: set[str] = set()
    for row in profile_rows:
        proj_id = str(row.get("proj_id") or "").strip()
        if not proj_id or proj_id in seen:
            continue
        seen.add(proj_id)
        proj_ids.append(proj_id)
        if limit and len(proj_ids) >= limit:
            break
    return proj_ids


def registered_proj_id_rows(proj_ids: list[str]) -> list[dict[str, str]]:
    return [{"proj_id": proj_id} for proj_id in proj_ids]


def fetch_dataset_for_registered_proj_ids(
    dataset: str,
    endpoint: str,
    proj_ids: list[str],
    api_key: str,
    args: argparse.Namespace,
    errors: list[dict[str, Any]],
) -> list[dict[str, Any]]:
    rows: list[dict[str, Any]] = []
    total = len(proj_ids)
    for index, proj_id in enumerate(proj_ids, start=1):
        params = dataset_params(dataset, args)
        params["proj_id"] = proj_id
        try:
            project_rows = fetch_sec_all_pages(
                endpoint,
                api_key,
                query_params=params,
                max_pages=args.max_pages,
            )
            rows.extend(project_rows)
            print(f"Fetched {dataset} for registered proj_id {index}/{total}: {proj_id}")
        except Exception as exc:
            errors.append(
                {
                    "dataset": dataset,
                    "endpoint": endpoint,
                    "proj_id": proj_id,
                    "params": json.dumps(params, ensure_ascii=False, separators=(",", ":")),
                    "error": str(exc),
                }
            )
            print(f"ERROR fetching {dataset} for proj_id={proj_id}: {exc}", file=sys.stderr)
            if not should_continue_on_error(args):
                raise
    return rows


def output_base(output_dir: Path, args: argparse.Namespace) -> Path:
    group = args.output_group.strip()
    return output_dir / group if group else output_dir


def output_paths(output_dir: Path, file_name: str, args: argparse.Namespace) -> list[Path]:
    base_dir = output_base(output_dir, args)
    if args.output_layout == "flat":
        return [base_dir / file_name]

    snapshot_date = args.snapshot_date.strip() or date.today().isoformat()
    group = args.output_group.strip()
    if group:
        return [
            output_dir / "latest" / group / file_name,
            output_dir / "snapshots" / group / snapshot_date / file_name,
        ]
    return [
        output_dir / "latest" / file_name,
        output_dir / "snapshots" / snapshot_date / file_name,
    ]


def curated_paths(output_dir: Path, file_name: str, args: argparse.Namespace) -> list[Path]:
    base_dir = output_base(output_dir, args)
    if args.output_layout == "flat":
        return [base_dir / "curated" / file_name]

    snapshot_date = args.snapshot_date.strip() or date.today().isoformat()
    group = args.output_group.strip()
    if group:
        return [
            output_dir / "latest" / group / "curated" / file_name,
            output_dir / "snapshots" / group / snapshot_date / "curated" / file_name,
        ]
    return [
        output_dir / "latest" / "curated" / file_name,
        output_dir / "snapshots" / snapshot_date / "curated" / file_name,
    ]


def row_key(row: dict[str, Any], include_fund_class: bool = True) -> tuple[str, str]:
    proj_id = str(row.get("proj_id") or "").strip()
    fund_class = str(row.get("fund_class_name") or "").strip() if include_fund_class else ""
    return proj_id, fund_class


def row_date_score(row: dict[str, Any]) -> str:
    return str(
        row.get("last_upd_date")
        or row.get("as_of_date")
        or row.get("start_date")
        or row.get("nav_date")
        or ""
    )


def latest_lookup(
    rows: list[dict[str, Any]],
    include_fund_class: bool = True,
) -> dict[tuple[str, str], dict[str, Any]]:
    lookup: dict[tuple[str, str], dict[str, Any]] = {}
    for row in rows:
        key = row_key(row, include_fund_class=include_fund_class)
        if not key[0]:
            continue
        current = lookup.get(key)
        if current is None or row_date_score(row) >= row_date_score(current):
            lookup[key] = row
    return lookup


def add_prefixed_fields(
    target: dict[str, Any],
    source: dict[str, Any] | None,
    prefix: str,
    skip: set[str] | None = None,
) -> None:
    if not source:
        return
    skip = skip or set()
    for key, value in source.items():
        if key in skip:
            continue
        target[f"{prefix}{key}"] = value


def build_fund_core(rows_by_dataset: dict[str, list[dict[str, Any]]]) -> list[dict[str, Any]]:
    profiles = rows_by_dataset.get("profiles", [])
    urls_by_class = latest_lookup(rows_by_dataset.get("factsheet_urls", []))
    risk_by_fund = latest_lookup(rows_by_dataset.get("risk_spectrum", []), include_fund_class=False)
    stats_by_class = latest_lookup(rows_by_dataset.get("statistics", []))
    dividend_by_class = latest_lookup(rows_by_dataset.get("dividend_policy", []))

    rows: list[dict[str, Any]] = []
    for profile in profiles:
        row = dict(profile)
        class_key = row_key(profile)
        fund_key = row_key(profile, include_fund_class=False)
        skip = {"proj_id", "fund_class_name"}
        add_prefixed_fields(row, urls_by_class.get(class_key), "factsheet_", skip=skip)
        add_prefixed_fields(row, risk_by_fund.get(fund_key), "risk_", skip={"proj_id"})
        add_prefixed_fields(row, stats_by_class.get(class_key), "stats_", skip=skip)
        add_prefixed_fields(row, dividend_by_class.get(class_key), "dividend_policy_", skip=skip)
        rows.append(row)
    return rows


def with_source(rows: list[dict[str, Any]], source: str) -> list[dict[str, Any]]:
    return [{"source_dataset": source, **row} for row in rows]


def build_curated_outputs(
    rows_by_dataset: dict[str, list[dict[str, Any]]],
) -> dict[str, tuple[list[dict[str, Any]], list[str] | None]]:
    manifest = []
    for dataset, file_name in DATASET_FILES.items():
        manifest.append(
            {
                "dataset": dataset,
                "raw_file": file_name,
                "row_count": len(rows_by_dataset.get(dataset, [])),
            }
        )

    fund_fees = (
        with_source(rows_by_dataset.get("mutual_fund_fees", []), "mutual_fund_fees")
        + with_source(rows_by_dataset.get("fees", []), "fees")
    )
    trading_terms = (
        with_source(
            rows_by_dataset.get("subscription_redemption_minimums", []),
            "subscription_redemption_minimums",
        )
        + with_source(
            rows_by_dataset.get("subscription_redemption_periods", []),
            "subscription_redemption_periods",
        )
    )
    holdings = (
        with_source(rows_by_dataset.get("asset_allocation", []), "asset_allocation")
        + with_source(rows_by_dataset.get("top5_holdings", []), "top5_holdings")
        + with_source(rows_by_dataset.get("outstanding_portfolio", []), "outstanding_portfolio")
        + with_source(rows_by_dataset.get("portfolio_asset_type", []), "portfolio_asset_type")
    )

    return {
        "00_sec_export_manifest.csv": (manifest, ["dataset", "raw_file", "row_count"]),
        "01_fund_core.csv": (build_fund_core(rows_by_dataset), None),
        "02_fund_fees.csv": (fund_fees, None),
        "03_fund_trading_terms.csv": (trading_terms, None),
        "04_fund_performance.csv": (rows_by_dataset.get("performance", []), PERFORMANCE_COLUMNS),
        "05_fund_holdings.csv": (holdings, None),
        "06_fund_nav_daily.csv": (rows_by_dataset.get("nav_daily", []), NAV_DAILY_COLUMNS),
        "07_fund_dividend_history_raw.csv": (
            rows_by_dataset.get("dividend_history", []),
            DIVIDEND_HISTORY_COLUMNS,
        ),
    }


def error_log_paths(output_dir: Path, args: argparse.Namespace) -> list[Path]:
    return output_paths(output_dir, "sec_fetch_errors.csv", args)


def should_write_curated(args: argparse.Namespace) -> bool:
    return str(args.write_curated).strip().lower() in {"1", "true", "yes", "y"}


def write_curated_outputs(
    output_dir: Path,
    rows_by_dataset: dict[str, list[dict[str, Any]]],
    args: argparse.Namespace,
) -> None:
    for file_name, (rows, headers) in build_curated_outputs(rows_by_dataset).items():
        if file_name != "00_sec_export_manifest.csv" and not rows:
            continue
        for path in curated_paths(output_dir, file_name, args):
            write_csv(path, rows, preferred_headers=headers)


def main() -> int:
    args = parse_args()
    api_key = get_api_key()
    output_dir = Path(args.output_dir)
    rows_by_dataset: dict[str, list[dict[str, Any]]] = {}
    errors: list[dict[str, Any]] = []
    datasets = selected_datasets(args.datasets)
    explicit_proj_ids = requested_proj_ids(args)
    registered_profile_rows: list[dict[str, Any]] = []
    registered_ids: list[str] = []

    if explicit_proj_ids:
        registered_ids = explicit_proj_ids
        if "profiles" in datasets:
            registered_profile_rows = fetch_profiles_for_proj_ids(
                explicit_proj_ids,
                api_key,
                args,
                errors,
            )
            print(f"Explicit profiles: {len(registered_profile_rows)} rows")
            for path in output_paths(output_dir, "00_registered_proj_ids.csv", args):
                write_csv(
                    path,
                    registered_proj_id_rows(explicit_proj_ids),
                    preferred_headers=["proj_id"],
                )
    elif should_use_registered_only(args) and not args.proj_id.strip():
        needs_registered_ids = any(dataset in PROJECT_ID_DATASETS for dataset in datasets)
        if needs_registered_ids or "profiles" in datasets:
            registered_profile_rows = fetch_registered_profiles(api_key, args)
            registered_ids = registered_proj_ids(
                registered_profile_rows,
                limit=args.registered_max_funds,
            )
            print(f"Registered profiles: {len(registered_profile_rows)} rows")
            print(f"Registered proj_id values: {len(registered_ids)}")
            for path in output_paths(output_dir, "00_registered_proj_ids.csv", args):
                write_csv(
                    path,
                    registered_proj_id_rows(registered_ids),
                    preferred_headers=["proj_id"],
                )

    for dataset in datasets:
        endpoint = ENDPOINTS[dataset]
        preferred_headers: list[str] | None = None
        if dataset == "profiles" and registered_profile_rows:
            rows = registered_profile_rows
            preferred_headers = PROFILE_COLUMNS
        elif (
            (explicit_proj_ids or should_use_registered_only(args))
            and not (args.proj_id.strip() and not explicit_proj_ids)
            and dataset in PROJECT_ID_DATASETS
            and registered_ids
        ):
            rows = fetch_dataset_for_registered_proj_ids(
                dataset,
                endpoint,
                registered_ids,
                api_key,
                args,
                errors,
            )
            rows, preferred_headers = transform_dataset_rows(dataset, rows)
        else:
            params = dataset_params(dataset, args)
            try:
                rows = fetch_sec_all_pages(
                    endpoint,
                    api_key,
                    query_params=params,
                    max_pages=args.max_pages,
                )
            except Exception as exc:
                errors.append(
                    {
                        "dataset": dataset,
                        "endpoint": endpoint,
                        "proj_id": "",
                        "params": json.dumps(params, ensure_ascii=False, separators=(",", ":")),
                        "error": str(exc),
                    }
                )
                print(f"ERROR fetching {dataset}: {exc}", file=sys.stderr)
                if should_continue_on_error(args):
                    rows = []
                else:
                    raise
            rows, preferred_headers = transform_dataset_rows(dataset, rows)
        rows_by_dataset[dataset] = rows
        file_name = DATASET_FILES.get(dataset, f"{dataset}.csv")
        for path in output_paths(output_dir, file_name, args):
            write_csv(path, rows, preferred_headers=preferred_headers)

    if errors:
        for path in error_log_paths(output_dir, args):
            write_csv(
                path,
                errors,
                preferred_headers=["dataset", "endpoint", "proj_id", "params", "error"],
            )

    if should_write_curated(args):
        write_curated_outputs(output_dir, rows_by_dataset, args)

    return 0


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except Exception as exc:
        print(f"ERROR: {exc}", file=sys.stderr)
        raise SystemExit(1)
