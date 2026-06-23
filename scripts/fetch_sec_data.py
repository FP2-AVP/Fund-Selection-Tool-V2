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
    parser.add_argument("--proj-id", default="M0000_2552", help="Project ID for fund-specific endpoints.")
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


def add_if_present(params: dict[str, Any], key: str, value: Any) -> None:
    if value is None:
        return
    text = str(value).strip()
    if text:
        params[key] = text


def factsheet_params(args: argparse.Namespace, include_fund_class: bool = True) -> dict[str, Any]:
    params: dict[str, Any] = {"proj_id": args.proj_id, "page_size": args.page_size}
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
    params: dict[str, Any] = {"proj_id": args.proj_id, "page_size": args.page_size}
    if include_fund_class:
        add_if_present(params, "fund_class_name", getattr(args, "fund_class_name", ""))
    return params


def dataset_params(dataset: str, args: argparse.Namespace) -> dict[str, Any]:
    if dataset == "profiles":
        return {"page_size": args.page_size}
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
        params = {"proj_id": args.proj_id, "page_size": args.page_size}
        add_if_present(params, "start_period", args.start_period)
        add_if_present(params, "end_period", args.end_period)
        return params
    if dataset == "portfolio_asset_type":
        params = {"proj_id": args.proj_id, "page_size": args.page_size}
        add_if_present(params, "start_period", args.start_period)
        add_if_present(params, "end_period", args.end_period)
        return params
    if dataset == "nav_daily":
        return {
            "proj_id": args.proj_id,
            "start_nav_date": args.start_nav_date,
            "end_nav_date": args.end_nav_date,
            "page_size": args.page_size,
        }
    return {"page_size": args.page_size}


def output_paths(output_dir: Path, file_name: str, args: argparse.Namespace) -> list[Path]:
    if args.output_layout == "flat":
        return [output_dir / file_name]

    snapshot_date = args.snapshot_date.strip() or date.today().isoformat()
    return [
        output_dir / "latest" / file_name,
        output_dir / "snapshots" / snapshot_date / file_name,
    ]


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
        rows, preferred_headers = transform_dataset_rows(dataset, rows)
        file_name = DATASET_FILES.get(dataset, f"{dataset}.csv")
        for path in output_paths(output_dir, file_name, args):
            write_csv(path, rows, preferred_headers=preferred_headers)

    return 0


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except Exception as exc:
        print(f"ERROR: {exc}", file=sys.stderr)
        raise SystemExit(1)
