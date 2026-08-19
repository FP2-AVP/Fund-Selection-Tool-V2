#!/usr/bin/env python3
"""Build SEC data_preparation and master_view rows from raw endpoint tabs in Google Sheets."""

from __future__ import annotations

import argparse
import json
import os
import re
import time
from typing import Any

SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]
SAFE_CELL_CHARS = 49_000
SHEET_WRITE_CHUNK_ROWS = 100
PROFILE_ALL_COLUMNS = [
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
MASTER_PROFILE_COLUMNS = [
    "comp_name_th",
    "proj_id",
    "regis_id",
    "fund_class_name",
    "proj_name_th",
    "proj_name_en",
    "proj_abbr_name",
    "fund_status",
    "invest_country_flag_label",
    "investment_policy_desc",
    "management_style",
    "feederfund_master_fund",
    "feederfund_country",
    "exchange_rate_protection_policy",
    "fund_class_detail",
    "fund_class_description",
    "fund_class_tax_incentive_type",
    "fund_class_isin_code",
]
SUMMARY_COLUMNS = [
    "pdf_factsheet",
    "amc_url_factsheet",
    "ipo_start_date",
    "ipo_end_date",
    "benchmarks",
    "minimum_sub_ipo",
    "minimum_sub",
    "type_settlement_period",
    "settlement_period",
    "risk_spectrum",
    "risk_spectrum_desc",
    "maximum_drawdown",
    "recovering_period",
    "fx_hedging",
    "portfolio_turnover_ratio",
    "portfolio_duration_period",
    "yield_to_maturity",
    "sharpe_ratio",
    "beta",
    "alpha",
    "statistics_last_upd_date",
    "dividend_policy",
    "rate_fee_type_desc_nav",
    "actual_value_fee_type_desc_nav",
    "rate_fee_type_desc_buy",
    "actual_value_fee_type_desc_buy",
    "fees_last_upd_date",
    "ผลตอบแทนกองทุนรวมแบบปักหมุด",
    "ผลตอบแทนตัวชี้วัดแบบปักหมุด",
    "ค่าเฉลี่ยในกลุ่มเดียวกันแบบปักหมุด",
    "ความผันผวนของกองทุนรวมแบบปักหมุด",
    "ความผันผวนของตัวชี้วัดแบบปักหมุด",
    "ผลตอบแทนกองทุนรวม 5 ปีย้อนหลัง",
    "ความผันผวนของกองทุนรวม 5 ปีย้อนหลัง",
    "ผลตอบแทนตัวชี้วัด 5 ปีย้อนหลัง",
    "ความผันผวนของตัวชี้วัด 5 ปีย้อนหลัง",
    "performance_last_upd_date",
    "asset_allocation",
    "asset_allocation_last_upd_date",
    "top5_holdings",
    "top5_holdings_last_upd_date",
]
DATA_PREPARATION_HEADERS = PROFILE_ALL_COLUMNS + SUMMARY_COLUMNS
MASTER_HEADERS = MASTER_PROFILE_COLUMNS + SUMMARY_COLUMNS


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Build SEC data_preparation and master_view from existing raw Google Sheet tabs.")
    parser.add_argument(
        "--spreadsheet-id",
        default=os.environ.get("SEC_MASTER_VIEW_SPREADSHEET_ID", ""),
        help="Destination Google Sheet ID or URL. Can also be set by SEC_MASTER_VIEW_SPREADSHEET_ID.",
    )
    parser.add_argument(
        "--source-spreadsheet-id",
        default=os.environ.get("SEC_MASTER_VIEW_SPREADSHEET_ID", ""),
        help="Source raw-tabs Google Sheet ID or URL. Defaults to SEC_MASTER_VIEW_SPREADSHEET_ID.",
    )
    parser.add_argument("--output-tab", default="data_preparation", help="Target tab name for data_preparation.")
    parser.add_argument("--master-output-tab", default="master_view", help="Target tab name for master_view.")
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


def spreadsheet_id_from_value(value: str) -> str:
    text = normalized_text(value)
    match = re.search(r"/spreadsheets/d/([a-zA-Z0-9_-]+)", text)
    if match:
        return match.group(1)
    return text


def normalized_key(value: Any) -> str:
    return re.sub(r"\s+", " ", normalized_text(value)).casefold()


def numeric_text(value: Any) -> str:
    text = normalized_text(value)
    if text.endswith(".0"):
        return text[:-2]
    return text


def numeric_sort_value(value: Any) -> tuple[int, str]:
    text = numeric_text(value)
    try:
        return int(float(text)), text
    except ValueError:
        return 999_999, text


def row_date(row: dict[str, Any]) -> str:
    return normalized_text(
        row.get("last_upd_date")
        or row.get("as_of_date")
        or row.get("start_date")
        or row.get("nav_date")
        or row.get("period")
    )


def row_effective_date(row: dict[str, Any]) -> str:
    """Date/period represented by the record, independent of revision time."""
    return normalized_text(
        row.get("end_date")
        or row.get("as_of_date")
        or row.get("period")
        or row.get("nav_date")
        or row.get("dividend_date")
        or row.get("start_date")
        or row.get("last_upd_date")
    )


def row_revision_date(row: dict[str, Any]) -> str:
    return normalized_text(row.get("last_upd_date"))


def snapshot_signature(row: dict[str, Any]) -> tuple[str, ...]:
    """Fields that identify one SEC snapshot/revision of an effective period."""
    return tuple(normalized_text(row.get(field)) for field in [
        "start_date",
        "end_date",
        "as_of_date",
        "period",
        "nav_date",
        "dividend_date",
        "prospectus_type",
        "last_upd_date",
    ])


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


def execute_with_retry(request: Any, label: str, retries: int = 5) -> Any:
    last_error: Exception | None = None
    for attempt in range(1, retries + 1):
        try:
            return request.execute()
        except Exception as exc:
            last_error = exc
            if attempt >= retries:
                break
            delay = min(60, 2 ** attempt)
            print(f"Retry Google Sheets {label} {attempt}/{retries - 1} after {delay}s: {exc}")
            time.sleep(delay)
    raise RuntimeError(f"Google Sheets {label} failed after {retries} attempts: {last_error}")


def read_tab_rows(sheets: Any, spreadsheet_id: str, tab_name: str) -> list[dict[str, Any]]:
    result = execute_with_retry(
        sheets.spreadsheets().values().get(
            spreadsheetId=spreadsheet_id,
            range=f"'{tab_name}'",
        ),
        f"read {tab_name}",
    )
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
    return max(rows, key=lambda row: (row_effective_date(row), row_revision_date(row))) if rows else {}


def latest_lookup_by_key(rows: list[dict[str, Any]], key_fn) -> dict[Any, dict[str, Any]]:
    grouped: dict[Any, list[dict[str, Any]]] = {}
    for row in rows:
        key = key_fn(row)
        if key:
            grouped.setdefault(key, []).append(row)
    return {key: latest_row(items) for key, items in grouped.items()}


def latest_group_lookup_by_key(
    rows: list[dict[str, Any]],
    key_fn,
    item_key_fn=None,
    prefer_completeness: bool = False,
) -> dict[Any, list[dict[str, Any]]]:
    grouped: dict[Any, list[dict[str, Any]]] = {}
    for row in rows:
        key = key_fn(row)
        if key:
            grouped.setdefault(key, []).append(row)

    out: dict[Any, list[dict[str, Any]]] = {}
    for key, items in grouped.items():
        latest_effective_date = max(row_effective_date(item) for item in items)
        effective_items = [
            item for item in items
            if row_effective_date(item) == latest_effective_date
        ]
        snapshots: dict[tuple[str, ...], list[dict[str, Any]]] = {}
        for item in effective_items:
            snapshots.setdefault(snapshot_signature(item), []).append(item)

        def completeness(snapshot_items: list[dict[str, Any]]) -> int:
            if item_key_fn is not None:
                return len({item_key_fn(item) for item in snapshot_items})
            return len({
                tuple(sorted((field, normalized_text(value)) for field, value in item.items()))
                for item in snapshot_items
            })

        # Normally the newest revision wins inside the latest effective period.
        # Completeness is considered first only for datasets with a known
        # expected sequence (currently Top 5 holdings).
        selected = max(
            snapshots.values(),
            key=lambda snapshot_items: (
                completeness(snapshot_items) if prefer_completeness else 0,
                max(row_revision_date(item) for item in snapshot_items),
            ),
        )
        out[key] = selected
    return out


def copy_fields(row: dict[str, Any], fields: list[str]) -> dict[str, Any]:
    return {field: normalized_text(row.get(field)) for field in fields}


def join_parts(parts: list[str]) -> str:
    return " | ".join(part for part in parts if normalized_text(part))


def benchmarks_lookup(rows: list[dict[str, Any]]) -> dict[str, str]:
    grouped: dict[str, list[str]] = {}
    rows_by_proj = latest_group_lookup_by_key(
        rows,
        row_proj,
        lambda row: numeric_text(row.get("group_seq")) or normalized_text(row.get("benchmark")),
    )
    for proj_id, project_rows in rows_by_proj.items():
        seen: set[str] = set()
        for row in sorted(project_rows, key=lambda item: numeric_sort_value(item.get("group_seq"))):
            benchmark = normalized_text(row.get("benchmark"))
            if not benchmark:
                continue
            group_seq = numeric_text(row.get("group_seq"))
            text = f"{group_seq}.{benchmark}" if group_seq else benchmark
            if text in seen:
                continue
            seen.add(text)
            grouped.setdefault(proj_id, []).append(text)
    return {proj_id: " | ".join(values) for proj_id, values in grouped.items()}


def latest_field_lookup(rows: list[dict[str, Any]], key_fn, fields: list[str]) -> dict[Any, dict[str, Any]]:
    return {
        key: copy_fields(row, fields)
        for key, row in latest_lookup_by_key(rows, key_fn).items()
    }


def fee_type_matches(row: dict[str, Any], aliases: list[str]) -> bool:
    desc = normalized_key(row.get("fee_type_desc"))
    return any(normalized_key(alias) == desc for alias in aliases)


def format_fee_group(rows: list[dict[str, Any]], specs: list[tuple[str, list[str]]], value_field: str) -> str:
    parts: list[str] = []
    for label, aliases in specs:
        row = latest_row([item for item in rows if fee_type_matches(item, [label, *aliases])])
        value = normalized_text(row.get(value_field))
        if value:
            parts.append(f"{label} : {value}")
    return join_parts(parts)


def fees_lookup(rows: list[dict[str, Any]]) -> dict[tuple[str, str], dict[str, str]]:
    nav_specs = [
        ("ค่าธรรมเนียมการจัดการ (Management Fee)", ["Management Fee"]),
        ("ค่าธรรมเนียมและค่าใช้จ่ายรวมทั้งหมด (Total Fee and Expense)", ["Total Fee and Expense"]),
    ]
    buy_specs = [
        ("ค่าธรรมเนียมการขายหน่วยลงทุน (Front-end Fee)", ["Front-end Fee"]),
        ("ค่าธรรมเนียมการรับซื้อคืนหน่วยลงทุน (Back-end Fee)", ["Back-end Fee"]),
        ("ค่าธรรมเนียมการสับเปลี่ยนหน่วยลงทุนเข้า (Switching In)", ["Switching In", "SWITCHING IN"]),
        ("ค่าธรรมเนียมการสับเปลี่ยนหน่วยลงทุนออก (Switching Out)", ["Switching Out", "SWITCHING OUT"]),
        ("ค่าธรรมเนียมการโอนหน่วยลงทุน (Transfer Fee)", ["Transfer Fee"]),
    ]
    out: dict[tuple[str, str], dict[str, str]] = {}
    for key, fee_rows in latest_group_lookup_by_key(
        rows,
        proj_class_key,
        lambda row: normalized_key(row.get("fee_type_desc")),
    ).items():
        out[key] = {
            "rate_fee_type_desc_nav": format_fee_group(fee_rows, nav_specs, "rate"),
            "actual_value_fee_type_desc_nav": format_fee_group(fee_rows, nav_specs, "actual_value"),
            "rate_fee_type_desc_buy": format_fee_group(fee_rows, buy_specs, "rate"),
            "actual_value_fee_type_desc_buy": format_fee_group(fee_rows, buy_specs, "actual_value"),
            "fees_last_upd_date": row_date(latest_row(fee_rows)),
        }
    return out


PINNED_PERIODS = [
    ("year to date", {"year to date", "ytd"}),
    ("3 months", {"3 months", "3 month", "3m"}),
    ("6 months", {"6 months", "6 month", "6m"}),
    ("1 year", {"1 year", "1y", "1 yr", "1yr"}),
    ("3 year", {"3 year", "3 years", "3y", "3 yr", "3yr"}),
    ("5 year", {"5 year", "5 years", "5y", "5 yr", "5yr"}),
    ("10 year", {"10 year", "10 years", "10y", "10 yr", "10yr"}),
    ("inception date", {"inception date", "since inception", "inception"}),
]
PINNED_PERFORMANCE_COLUMNS = [
    "ผลตอบแทนกองทุนรวมแบบปักหมุด",
    "ผลตอบแทนตัวชี้วัดแบบปักหมุด",
    "ค่าเฉลี่ยในกลุ่มเดียวกันแบบปักหมุด",
    "ความผันผวนของกองทุนรวมแบบปักหมุด",
    "ความผันผวนของตัวชี้วัดแบบปักหมุด",
]
ANNUAL_PERFORMANCE_COLUMNS = [
    "ผลตอบแทนกองทุนรวม 5 ปีย้อนหลัง",
    "ความผันผวนของกองทุนรวม 5 ปีย้อนหลัง",
    "ผลตอบแทนตัวชี้วัด 5 ปีย้อนหลัง",
    "ความผันผวนของตัวชี้วัด 5 ปีย้อนหลัง",
]


def performance_type_matches(value: Any, output_column: str) -> bool:
    desc = normalized_key(value)
    target = output_column.replace("แบบปักหมุด", "").replace(" 5 ปีย้อนหลัง", "")
    return normalized_key(target) == desc or normalized_key(target) in desc


def pinned_period_label(value: Any) -> str:
    ref = normalized_key(value)
    compact = ref.replace(" ", "")
    for label, aliases in PINNED_PERIODS:
        if ref in aliases or compact in {alias.replace(" ", "") for alias in aliases}:
            return label
    return ""


def format_pinned_performance(rows: list[dict[str, Any]], output_column: str) -> str:
    type_rows = [row for row in rows if performance_type_matches(row.get("performance_type_desc"), output_column)]
    parts: list[str] = []
    for label, _aliases in PINNED_PERIODS:
        period_rows = [row for row in type_rows if pinned_period_label(row.get("reference_period")) == label]
        row = latest_row(period_rows)
        value = normalized_text(row.get("performance_value"))
        if value:
            parts.append(f"{label} : {value}")
    return join_parts(parts)


def format_annual_performance(rows: list[dict[str, Any]], output_column: str) -> str:
    type_rows = [row for row in rows if performance_type_matches(row.get("performance_type_desc"), output_column)]
    by_year: dict[int, dict[str, Any]] = {}
    for row in type_rows:
        ref = numeric_text(row.get("reference_period"))
        if not re.fullmatch(r"\d{4}", ref):
            continue
        year = int(ref)
        current = by_year.get(year)
        if current is None or row_date(row) >= row_date(current):
            by_year[year] = row
    return join_parts([
        f"{year} : {normalized_text(by_year[year].get('performance_value'))}"
        for year in sorted(by_year)
        if normalized_text(by_year[year].get("performance_value"))
    ])


def performance_lookup(rows: list[dict[str, Any]]) -> dict[tuple[str, str], dict[str, str]]:
    out: dict[tuple[str, str], dict[str, str]] = {}
    for key, perf_rows in latest_group_lookup_by_key(
        rows,
        proj_class_key,
        lambda row: (
            normalized_key(row.get("performance_type_desc")),
            normalized_key(row.get("reference_period")),
        ),
    ).items():
        values = {
            column: format_pinned_performance(perf_rows, column)
            for column in PINNED_PERFORMANCE_COLUMNS
        }
        values.update({
            column: format_annual_performance(perf_rows, column)
            for column in ANNUAL_PERFORMANCE_COLUMNS
        })
        values["performance_last_upd_date"] = row_date(latest_row(perf_rows))
        out[key] = values
    return out


def seq_asset_lookup(
    rows: list[dict[str, Any]],
    output_column: str,
    date_column: str,
    prefer_completeness: bool = False,
) -> dict[str, dict[str, str]]:
    out: dict[str, dict[str, str]] = {}
    for proj_id, project_rows in latest_group_lookup_by_key(
        rows,
        row_proj,
        lambda row: numeric_text(row.get("asset_seq")) or normalized_text(row.get("asset_name")),
        prefer_completeness=prefer_completeness,
    ).items():
        parts: list[str] = []
        seen: set[str] = set()
        for row in sorted(project_rows, key=lambda item: numeric_sort_value(item.get("asset_seq"))):
            seq = numeric_text(row.get("asset_seq"))
            name = normalized_text(row.get("asset_name"))
            ratio = normalized_text(row.get("asset_ratio"))
            if not name:
                continue
            text = f"{seq}.{name} : {ratio}" if seq else f"{name} : {ratio}"
            if text in seen:
                continue
            seen.add(text)
            parts.append(text)
        out[proj_id] = {
            output_column: join_parts(parts),
            date_column: row_date(latest_row(project_rows)),
        }
    return out


def ensure_sheet(sheets: Any, spreadsheet_id: str, tab_name: str) -> int:
    metadata = execute_with_retry(
        sheets.spreadsheets().get(spreadsheetId=spreadsheet_id),
        "get spreadsheet metadata",
    )
    existing = {
        sheet["properties"]["title"]: sheet["properties"]["sheetId"]
        for sheet in metadata.get("sheets", [])
    }
    if tab_name in existing:
        return existing[tab_name]
    result = execute_with_retry(
        sheets.spreadsheets().batchUpdate(
            spreadsheetId=spreadsheet_id,
            body={"requests": [{"addSheet": {"properties": {"title": tab_name}}}]},
        ),
        f"add sheet {tab_name}",
    )
    return result["replies"][0]["addSheet"]["properties"]["sheetId"]


def resize_sheet_grid(
    sheets: Any,
    spreadsheet_id: str,
    sheet_id: int,
    row_count: int,
    column_count: int,
) -> None:
    execute_with_retry(
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
        ),
        "resize grid",
    )


def write_values_to_sheet(sheets: Any, spreadsheet_id: str, tab_name: str, values: list[list[Any]]) -> None:
    sheet_id = ensure_sheet(sheets, spreadsheet_id, tab_name)
    row_count = len(values) if values else 1
    column_count = max((len(row) for row in values), default=1)
    resize_sheet_grid(sheets, spreadsheet_id, sheet_id, row_count, column_count)
    execute_with_retry(
        sheets.spreadsheets().values().clear(
            spreadsheetId=spreadsheet_id,
            range=f"'{tab_name}'",
            body={},
        ),
        f"clear {tab_name}",
    )
    if not values:
        return
    for start in range(0, len(values), SHEET_WRITE_CHUNK_ROWS):
        chunk = [
            [truncate_cell_value(cell) for cell in row]
            for row in values[start : start + SHEET_WRITE_CHUNK_ROWS]
        ]
        execute_with_retry(
            sheets.spreadsheets().values().update(
                spreadsheetId=spreadsheet_id,
                range=f"'{tab_name}'!A{start + 1}",
                valueInputOption="RAW",
                body={"values": chunk},
            ),
            f"write {tab_name} rows {start + 1}-{start + len(chunk)}",
        )
        print(f"Wrote {tab_name} rows {start + 1} - {start + len(chunk)}")


def build_data_preparation_rows(raw_tabs: dict[str, list[dict[str, Any]]]) -> list[dict[str, Any]]:
    profiles = list(latest_lookup_by_key(raw_tabs["02_profiles"], proj_class_key).values())
    factsheet_urls = latest_field_lookup(
        raw_tabs["06_factsheet_urls"],
        proj_class_key,
        ["pdf_factsheet", "amc_url_factsheet"],
    )
    ipos = latest_field_lookup(raw_tabs["07_ipos"], row_proj, ["first_sell_start_date", "first_sell_end_date", "start_date", "end_date"])
    benchmarks = benchmarks_lookup(raw_tabs["08_benchmarks"])
    minimums = latest_field_lookup(
        raw_tabs["09_subscription_redemption_minimums"],
        proj_class_key,
        ["minimum_sub_ipo", "minimum_sub"],
    )
    periods = latest_field_lookup(
        raw_tabs["10_subscription_redemption_periods"],
        proj_class_key,
        ["type", "settlement_period"],
    )
    risks = latest_field_lookup(
        raw_tabs["11_risk_spectrum"],
        row_proj,
        ["risk_spectrum", "risk_spectrum_desc"],
    )
    statistics = latest_field_lookup(
        raw_tabs["12_statistics"],
        proj_class_key,
        [
            "maximum_drawdown",
            "recovering_period",
            "fx_hedging",
            "portfolio_turnover_ratio",
            "portfolio_duration_period",
            "yield_to_maturity",
            "sharpe_ratio",
            "beta",
            "alpha",
            "last_upd_date",
        ],
    )
    dividend_policies = latest_field_lookup(
        raw_tabs["13_dividend_policy"],
        proj_class_key,
        ["dividend_policy"],
    )
    fees = fees_lookup(raw_tabs["14_fees"])
    performance = performance_lookup(raw_tabs["15_performance"])
    asset_allocations = seq_asset_lookup(
        raw_tabs["16_asset_allocation"],
        "asset_allocation",
        "asset_allocation_last_upd_date",
    )
    top5_holdings = seq_asset_lookup(
        raw_tabs["17_top5_holdings"],
        "top5_holdings",
        "top5_holdings_last_upd_date",
        prefer_completeness=True,
    )

    data_preparation_rows: list[dict[str, Any]] = []
    for profile in profiles:
        proj_id = row_proj(profile)
        fund_class_name = row_class(profile)
        key = (proj_id, fund_class_name)
        statistics_row = statistics.get(key, {})
        period = periods.get(key, {})

        row = {column: normalized_text(profile.get(column)) for column in PROFILE_ALL_COLUMNS}
        row.update(factsheet_urls.get(key, {}))
        ipo = ipos.get(proj_id, {})
        row["ipo_start_date"] = normalized_text(ipo.get("first_sell_start_date") or ipo.get("start_date"))
        row["ipo_end_date"] = normalized_text(ipo.get("first_sell_end_date") or ipo.get("end_date"))
        row["benchmarks"] = benchmarks.get(proj_id, "")
        row.update(minimums.get(key, {}))
        row["type_settlement_period"] = normalized_text(period.get("type"))
        row["settlement_period"] = normalized_text(period.get("settlement_period"))
        row.update(risks.get(proj_id, {}))
        for column in [
            "maximum_drawdown",
            "recovering_period",
            "fx_hedging",
            "portfolio_turnover_ratio",
            "portfolio_duration_period",
            "yield_to_maturity",
            "sharpe_ratio",
            "beta",
            "alpha",
        ]:
            row[column] = normalized_text(statistics_row.get(column))
        row["statistics_last_upd_date"] = normalized_text(statistics_row.get("last_upd_date"))
        row.update(dividend_policies.get(key, {}))
        row.update(fees.get(key, {}))
        row.update(performance.get(key, {}))
        row.update(asset_allocations.get(proj_id, {}))
        row.update(top5_holdings.get(proj_id, {}))

        data_preparation_rows.append(row)
    return data_preparation_rows


def build_master_rows(data_preparation_rows: list[dict[str, Any]]) -> list[dict[str, Any]]:
    return [
        {column: normalized_text(row.get(column)) for column in MASTER_HEADERS}
        for row in data_preparation_rows
    ]


def rows_to_values(rows: list[dict[str, Any]], headers: list[str]) -> list[list[Any]]:
    return [headers] + [
        [row.get(header, "") for header in headers]
        for row in rows
    ]


def main() -> int:
    args = parse_args()
    if not args.spreadsheet_id.strip():
        raise RuntimeError("Missing --spreadsheet-id or SEC_MASTER_VIEW_SPREADSHEET_ID.")
    if not args.source_spreadsheet_id.strip():
        raise RuntimeError("Missing --source-spreadsheet-id or SEC_MASTER_VIEW_SPREADSHEET_ID.")

    from googleapiclient.discovery import build

    sheets = build("sheets", "v4", credentials=credentials_from_env(), cache_discovery=False)
    source_spreadsheet_id = spreadsheet_id_from_value(args.source_spreadsheet_id)
    destination_spreadsheet_id = spreadsheet_id_from_value(args.spreadsheet_id)
    required_tabs = [
        "02_profiles",
        "06_factsheet_urls",
        "07_ipos",
        "08_benchmarks",
        "09_subscription_redemption_minimums",
        "10_subscription_redemption_periods",
        "11_risk_spectrum",
        "12_statistics",
        "13_dividend_policy",
        "14_fees",
        "15_performance",
        "16_asset_allocation",
        "17_top5_holdings",
    ]
    raw_tabs = {
        tab_name: read_tab_rows(sheets, source_spreadsheet_id, tab_name)
        for tab_name in required_tabs
    }
    data_preparation_rows = build_data_preparation_rows(raw_tabs)
    master_rows = build_master_rows(data_preparation_rows)
    output_tab = args.output_tab.strip() or "data_preparation"
    master_output_tab = args.master_output_tab.strip() or "master_view"
    write_values_to_sheet(
        sheets,
        destination_spreadsheet_id,
        output_tab,
        rows_to_values(data_preparation_rows, DATA_PREPARATION_HEADERS),
    )
    if master_output_tab != output_tab:
        write_values_to_sheet(
            sheets,
            destination_spreadsheet_id,
            master_output_tab,
            rows_to_values(master_rows, MASTER_HEADERS),
        )
    print(f"data_preparation rows: {len(data_preparation_rows)}")
    print(f"master_view rows: {len(master_rows)}")
    return 0


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except Exception as exc:
        print(f"ERROR: {exc}")
        raise SystemExit(1)
