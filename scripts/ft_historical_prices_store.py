#!/usr/bin/env python3
"""Fetch FT Markets historical prices into Drive-friendly local storage."""

from __future__ import annotations

import argparse
import csv
import html
import json
import re
import sqlite3
import urllib.error
import urllib.parse
import urllib.request
from dataclasses import dataclass
from datetime import date, datetime, timezone
from html.parser import HTMLParser
from pathlib import Path
from typing import Any


PROJECT_ROOT = Path(__file__).resolve().parents[1]
DEFAULT_OUTPUT_ROOT = PROJECT_ROOT / "Data" / "ft_historical_prices"
DEFAULT_DB_PATH = DEFAULT_OUTPUT_ROOT / "ft_historical_prices.sqlite"
FT_MARKETS_BASE_URL = "https://markets.ft.com/data"
FT_PRODUCT_PATHS = ("etfs", "funds")
FT_AJAX_URL = "https://markets.ft.com/data/equities/ajax/get-historical-prices"


@dataclass(frozen=True)
class PriceRow:
    symbol: str
    ft_issue_id: str
    price_date: str
    open: float | None
    high: float | None
    low: float | None
    close: float | None
    volume: int | None
    source: str
    fetched_at: str


class HistoricalRowsParser(HTMLParser):
    def __init__(self) -> None:
        super().__init__()
        self.rows: list[list[str]] = []
        self._current_row: list[str] | None = None
        self._current_cell: list[str] | None = None
        self._capture_span = False

    def handle_starttag(self, tag: str, attrs: list[tuple[str, str | None]]) -> None:
        attrs_dict = dict(attrs)
        if tag == "tr":
            self._current_row = []
        elif tag == "td" and self._current_row is not None:
            self._current_cell = []
            self._capture_span = False
        elif tag == "span" and self._current_cell is not None:
            classes = attrs_dict.get("class", "")
            if "mod-ui-hide-small-below" in classes:
                self._capture_span = True

    def handle_data(self, data: str) -> None:
        if self._current_cell is None:
            return
        if self._capture_span or not self._current_cell:
            text = data.strip()
            if text:
                self._current_cell.append(text)

    def handle_endtag(self, tag: str) -> None:
        if tag == "span" and self._current_cell is not None:
            self._capture_span = False
        elif tag == "td" and self._current_cell is not None:
            cell = " ".join(self._current_cell).strip()
            if self._current_row is not None:
                self._current_row.append(html.unescape(cell))
            self._current_cell = None
        elif tag == "tr" and self._current_row is not None:
            if len(self._current_row) == 6:
                self.rows.append(self._current_row)
            self._current_row = None


class ProfileInvestmentParser(HTMLParser):
    def __init__(self) -> None:
        super().__init__()
        self.rows: list[dict[str, str]] = []
        self._in_app = False
        self._app_depth = 0
        self._current_section = ""
        self._current_row: dict[str, list[str]] | None = None
        self._current_cell: str | None = None

    def handle_starttag(self, tag: str, attrs: list[tuple[str, str | None]]) -> None:
        attrs_dict = dict(attrs)
        classes = attrs_dict.get("class", "") or ""
        if tag == "div" and "mod-profile-and-investment-app" in classes.split():
            self._in_app = True
            self._app_depth = 1
            return
        if not self._in_app:
            return
        self._app_depth += 1
        if tag == "table":
            if "table--profile" in classes:
                self._current_section = "profile"
            elif "table--invest" in classes:
                self._current_section = "investment"
        elif tag == "tr":
            self._current_row = {"label": [], "value": []}
        elif tag in {"th", "td"} and self._current_row is not None:
            self._current_cell = "label" if tag == "th" else "value"
        elif tag == "br" and self._current_cell and self._current_row is not None:
            self._current_row[self._current_cell].append("\n")

    def handle_data(self, data: str) -> None:
        if not self._in_app or not self._current_cell or self._current_row is None:
            return
        text = data.strip()
        if text:
            self._current_row[self._current_cell].append(text)

    def handle_endtag(self, tag: str) -> None:
        if not self._in_app:
            return
        if tag in {"th", "td"}:
            self._current_cell = None
        elif tag == "tr" and self._current_row is not None:
            label = normalize_whitespace(" ".join(self._current_row["label"]))
            value = normalize_multiline(" ".join(self._current_row["value"]))
            if label:
                self.rows.append({"section": self._current_section, "field": label, "value": value})
            self._current_row = None
        elif tag == "table":
            self._current_section = ""
        self._app_depth -= 1
        if self._app_depth <= 0:
            self._in_app = False


def utc_now() -> str:
    return datetime.now(timezone.utc).replace(microsecond=0).isoformat().replace("+00:00", "Z")


def normalize_whitespace(value: str) -> str:
    return re.sub(r"\s+", " ", value or "").strip()


def normalize_multiline(value: str) -> str:
    text = re.sub(r"[ \t]*\n[ \t]*", "\n", value or "")
    text = re.sub(r"[ \t]{2,}", " ", text)
    text = re.sub(r"\n{2,}", "\n", text)
    return text.strip()


def extract_ft_display_name(page: str, symbol: str) -> str:
    title_match = re.search(r"<title>\s*(.*?)\s*</title>", page, re.IGNORECASE | re.DOTALL)
    if title_match:
        title = html.unescape(normalize_whitespace(re.sub(r"<[^>]+>", " ", title_match.group(1))))
        for suffix in (f", {symbol} summary - FT.com", f" {symbol} summary - FT.com", " summary - FT.com", " - FT.com"):
            if title.endswith(suffix):
                title = title[: -len(suffix)].strip()
                break
        if title:
            return title
    desc_match = re.search(
        r'<meta\s+name=["\']description["\']\s+content=["\'](.*?)["\']',
        page,
        re.IGNORECASE | re.DOTALL,
    )
    if desc_match:
        desc = html.unescape(normalize_whitespace(desc_match.group(1)))
        match = re.match(r"Latest\s+(.+?)\s+\(" + re.escape(symbol) + r"\)", desc)
        if match:
            return match.group(1).strip()
    return ""


def slug_symbol(symbol: str) -> str:
    return re.sub(r"[^A-Za-z0-9]+", "_", symbol).strip("_")


def parse_number(value: str) -> float | None:
    text = value.replace(",", "").strip()
    if not text or text == "-":
        return None
    return float(text)


def parse_int(value: str) -> int | None:
    text = value.replace(",", "").strip()
    if not text or text == "-":
        return None
    return int(float(text))


def parse_ft_date(value: str) -> str:
    parsed = datetime.strptime(value, "%A, %B %d, %Y")
    return parsed.date().isoformat()


def parse_short_ft_date(value: str) -> str:
    text = normalize_whitespace(value).rstrip(".")
    parsed = datetime.strptime(text, "%b %d %Y")
    return parsed.date().isoformat()


def request_text(url: str, params: dict[str, str] | None = None, referer: str | None = None) -> str:
    full_url = url
    if params:
        full_url = f"{url}?{urllib.parse.urlencode(params)}"
    request = urllib.request.Request(
        full_url,
        headers={
            "User-Agent": "Mozilla/5.0",
            "Accept": "text/html,application/json",
        },
    )
    if referer:
        request.add_header("Referer", referer)
    with urllib.request.urlopen(request, timeout=30) as response:
        return response.read().decode("utf-8")


def ft_tearsheet_url(section: str, product_path: str = "etfs") -> str:
    product = product_path if product_path in FT_PRODUCT_PATHS else "etfs"
    return f"{FT_MARKETS_BASE_URL}/{product}/tearsheet/{section}"


def resolve_issue_id(symbol: str, product_path: str | None = None) -> tuple[str, str]:
    candidates = [product_path] if product_path in FT_PRODUCT_PATHS else list(FT_PRODUCT_PATHS)
    tried_urls: list[str] = []
    for product in candidates:
        historical_url = ft_tearsheet_url("historical", product)
        page_url = f"{historical_url}?{urllib.parse.urlencode({'s': symbol})}"
        tried_urls.append(page_url)
        try:
            page = request_text(historical_url, {"s": symbol})
        except urllib.error.HTTPError:
            continue
        match = re.search(r"data-mod-config=\"[^\"\n]*&quot;symbol&quot;:&quot;(\d+)&quot;", page)
        if match:
            return match.group(1), product
    raise RuntimeError(f"Could not find FT issue id for {symbol}; tried {', '.join(tried_urls)}")


def clean_html_text(value: str) -> str:
    text = re.sub(r"<br\s*/?>", "\n", value or "", flags=re.IGNORECASE)
    text = re.sub(r"<[^>]+>", " ", text)
    return html.unescape(normalize_multiline(text))


def parse_as_of_date(page: str) -> str:
    match = re.search(r"As of\s+([A-Z][a-z]{2}\s+\d{1,2}\s+\d{4})\.?", page)
    if not match:
        return ""
    try:
        return parse_short_ft_date(match.group(1))
    except ValueError:
        return ""


def parse_risk_measures(symbol: str, ft_issue_id: str, page: str, fetched_at: str) -> list[dict[str, str]]:
    as_of_date = parse_as_of_date(page)
    page_text = clean_html_text(page)
    benchmark_match = re.search(
        r"Benchmark used:\s*(.*?)(?:\s+Fund\s+Category average|\s+Alpha|\s+Beta|\s+R squared|\s+Sharpe ratio|\s+Standard deviation|\s+As of\b)",
        page_text,
        re.IGNORECASE,
    )
    benchmark_used = benchmark_match.group(1).strip() if benchmark_match else ""
    period_labels = {"1y": "1 year", "3y": "3 year", "5y": "5 years"}
    rows: list[dict[str, str]] = []
    row_pattern = re.compile(
        r"<tr>\s*<td[^>]*>(?P<metric>.*?)</td>\s*<td[^>]*>(?P<fund>.*?)</td>\s*<td[^>]*>(?P<category>.*?)</td>\s*</tr>",
        re.IGNORECASE | re.DOTALL,
    )
    panel_starts = [
        (match.start(), match.group("period").lower())
        for match in re.finditer(r'<div role="tabpanel" id="modriskmeasures(?P<period>[^"]+)-panel"', page, re.IGNORECASE)
    ]
    for idx, (start_idx, period_key) in enumerate(panel_starts):
        end_idx = panel_starts[idx + 1][0] if idx + 1 < len(panel_starts) else page.find('<footer class="mod-module__footer"', start_idx)
        if end_idx < 0:
            end_idx = len(page)
        panel_html = page[start_idx:end_idx]
        period = period_labels.get(period_key, period_key)
        for row_match in row_pattern.finditer(panel_html):
            metric = clean_html_text(row_match.group("metric"))
            if not metric:
                continue
            rows.append(
                {
                    "symbol": symbol,
                    "ft_issue_id": ft_issue_id,
                    "period": period,
                    "metric": metric,
                    "fund_value": clean_html_text(row_match.group("fund")),
                    "category_average": clean_html_text(row_match.group("category")),
                    "benchmark_used": benchmark_used,
                    "as_of_date": as_of_date,
                    "source": "FT Markets",
                    "fetched_at": fetched_at,
                }
            )
    return rows


def parse_top_holdings(symbol: str, ft_issue_id: str, page: str, fetched_at: str) -> list[dict[str, str]]:
    as_of_date = ""
    as_of_match = re.search(r"as of\s+([A-Z][a-z]{2}\s+\d{1,2}\s+\d{4})", page, re.IGNORECASE)
    if as_of_match:
        try:
            as_of_date = parse_short_ft_date(as_of_match.group(1))
        except ValueError:
            as_of_date = ""

    module_match = re.search(
        r'<div data-f2-app-id="mod-top-ten".*?(?=<div class="o-ads|\s*</section>|\s*<script|\Z)',
        page,
        re.IGNORECASE | re.DOTALL,
    )
    if not module_match:
        return []
    module_html = module_match.group(0)

    total_match = re.search(
        r"Per cent of portfolio in top 10 holdings:\s*<strong>(.*?)</strong>",
        module_html,
        re.IGNORECASE | re.DOTALL,
    )
    top10_portfolio_percent = clean_html_text(total_match.group(1)) if total_match else ""

    row_pattern = re.compile(r"<tr>\s*(?P<cells>(?:<td\b.*?</td>\s*){4})\s*</tr>", re.IGNORECASE | re.DOTALL)
    cell_pattern = re.compile(r"<td\b[^>]*>(.*?)</td>", re.IGNORECASE | re.DOTALL)
    rows: list[dict[str, str]] = []
    for row_match in row_pattern.finditer(module_html):
        cells = cell_pattern.findall(row_match.group("cells"))
        if len(cells) < 4:
            continue
        holding_cell = cells[0]
        holding_name = clean_html_text(re.sub(r"<span\b.*?</span>", " ", holding_cell, flags=re.IGNORECASE | re.DOTALL))
        holding_symbol_match = re.search(r'<span[^>]*class="[^"]*disclaimer[^"]*"[^>]*>(.*?)</span>', holding_cell, re.IGNORECASE | re.DOTALL)
        holding_symbol = clean_html_text(holding_symbol_match.group(1)) if holding_symbol_match else ""
        one_year_change = clean_html_text(cells[1])
        portfolio_weight = clean_html_text(cells[2])
        long_allocation_width = ""
        width_match = re.search(r"width:\s*([^;]+);", cells[3], re.IGNORECASE)
        if width_match:
            long_allocation_width = normalize_whitespace(width_match.group(1))
        if not holding_name or holding_name in {"--", "Category average"}:
            continue
        rows.append(
            {
                "symbol": symbol,
                "ft_issue_id": ft_issue_id,
                "rank": str(len(rows) + 1),
                "holding_name": holding_name,
                "holding_symbol": holding_symbol,
                "one_year_change": one_year_change,
                "portfolio_weight": portfolio_weight,
                "long_allocation": long_allocation_width,
                "top10_portfolio_percent": top10_portfolio_percent,
                "as_of_date": as_of_date,
                "source": "FT Markets",
                "fetched_at": fetched_at,
            }
        )
    return rows


def parse_rows(symbol: str, ft_issue_id: str, response_data: dict[str, Any], fetched_at: str) -> list[PriceRow]:
    parser = HistoricalRowsParser()
    parser.feed(response_data.get("html") or "")
    rows: list[PriceRow] = []
    for raw in parser.rows:
        rows.append(
            PriceRow(
                symbol=symbol,
                ft_issue_id=ft_issue_id,
                price_date=parse_ft_date(raw[0]),
                open=parse_number(raw[1]),
                high=parse_number(raw[2]),
                low=parse_number(raw[3]),
                close=parse_number(raw[4]),
                volume=parse_int(raw[5]),
                source="FT Markets",
                fetched_at=fetched_at,
            )
        )
    return rows


def fetch_range(symbol: str, ft_issue_id: str, start_date: str, end_date: str, product_path: str = "etfs") -> tuple[dict[str, Any], list[PriceRow]]:
    fetched_at = utc_now()
    page_url = f"{ft_tearsheet_url('historical', product_path)}?{urllib.parse.urlencode({'s': symbol})}"
    raw = request_text(
        FT_AJAX_URL,
        {"startDate": start_date.replace("-", "/"), "endDate": end_date.replace("-", "/"), "symbol": ft_issue_id},
        referer=page_url,
    )
    data = json.loads(raw)
    return data, parse_rows(symbol, ft_issue_id, data, fetched_at)


def fetch_profile_investment(symbol: str, ft_issue_id: str, product_path: str = "etfs") -> tuple[dict[str, Any], list[dict[str, str]]]:
    fetched_at = utc_now()
    page = request_text(ft_tearsheet_url("summary", product_path), {"s": symbol})
    display_name = extract_ft_display_name(page, symbol)
    as_of_date = parse_as_of_date(page)
    parser = ProfileInvestmentParser()
    parser.feed(page)
    rows = [
        {
            "symbol": symbol,
            "ft_issue_id": ft_issue_id,
            "section": row["section"],
            "field": row["field"],
            "value": row["value"],
            "source": "FT Markets",
            "fetched_at": fetched_at,
        }
        for row in parser.rows
    ]
    if display_name:
        rows.insert(
            0,
            {
                "symbol": symbol,
                "ft_issue_id": ft_issue_id,
                "section": "metadata",
                "field": "FT display name",
                "value": display_name,
                "source": "FT Markets",
                "fetched_at": fetched_at,
            },
        )
    raw = {
        "symbol": symbol,
        "ft_issue_id": ft_issue_id,
        "fetched_at": fetched_at,
        "display_name": display_name,
        "as_of_date": as_of_date,
        "profile_investment": rows,
    }
    return raw, rows


def fetch_risk_measures(symbol: str, ft_issue_id: str, product_path: str = "etfs") -> tuple[dict[str, Any], list[dict[str, str]]]:
    fetched_at = utc_now()
    page = request_text(ft_tearsheet_url("risk", product_path), {"s": symbol})
    display_name = extract_ft_display_name(page, symbol)
    rows = parse_risk_measures(symbol, ft_issue_id, page, fetched_at)
    raw = {
        "symbol": symbol,
        "ft_issue_id": ft_issue_id,
        "fetched_at": fetched_at,
        "display_name": display_name,
        "as_of_date": rows[0]["as_of_date"] if rows else "",
        "risk_measures": rows,
    }
    return raw, rows


def fetch_top_holdings(symbol: str, ft_issue_id: str, product_path: str = "etfs") -> tuple[dict[str, Any], list[dict[str, str]]]:
    fetched_at = utc_now()
    page = request_text(ft_tearsheet_url("holdings", product_path), {"s": symbol})
    display_name = extract_ft_display_name(page, symbol)
    rows = parse_top_holdings(symbol, ft_issue_id, page, fetched_at)
    raw = {
        "symbol": symbol,
        "ft_issue_id": ft_issue_id,
        "fetched_at": fetched_at,
        "display_name": display_name,
        "as_of_date": rows[0]["as_of_date"] if rows else "",
        "top_holdings": rows,
    }
    return raw, rows


def yearly_ranges(start_date: date, end_date: date) -> list[tuple[str, str]]:
    ranges: list[tuple[str, str]] = []
    cursor = start_date
    while cursor <= end_date:
        next_year = date(cursor.year + 1, cursor.month, cursor.day)
        chunk_end = min(next_year, end_date)
        ranges.append((cursor.isoformat(), chunk_end.isoformat()))
        cursor = date.fromordinal(chunk_end.toordinal() + 1)
    return ranges


def init_db(path: Path) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with sqlite3.connect(path) as conn:
        conn.execute(
            """
            create table if not exists ft_historical_prices (
                symbol text not null,
                ft_issue_id text not null,
                price_date text not null,
                open real,
                high real,
                low real,
                close real,
                volume integer,
                source text not null,
                fetched_at text not null,
                primary key (symbol, price_date)
            )
            """
        )
        conn.execute("create index if not exists idx_ft_prices_date on ft_historical_prices(price_date)")
        conn.execute(
            """
            create table if not exists ft_profile_investment (
                symbol text not null,
                ft_issue_id text not null,
                section text not null,
                field text not null,
                value text,
                source text not null,
                fetched_at text not null,
                primary key (symbol, section, field)
            )
            """
        )
        conn.execute("create index if not exists idx_ft_profile_symbol on ft_profile_investment(symbol)")
        conn.execute(
            """
            create table if not exists ft_risk_measures (
                symbol text not null,
                ft_issue_id text not null,
                period text not null,
                metric text not null,
                fund_value text,
                category_average text,
                benchmark_used text,
                as_of_date text,
                source text not null,
                fetched_at text not null,
                primary key (symbol, period, metric)
            )
            """
        )
        conn.execute("create index if not exists idx_ft_risk_symbol on ft_risk_measures(symbol)")
        conn.execute(
            """
            create table if not exists ft_top_holdings (
                symbol text not null,
                ft_issue_id text not null,
                rank integer not null,
                holding_name text,
                holding_symbol text,
                one_year_change text,
                portfolio_weight text,
                long_allocation text,
                top10_portfolio_percent text,
                as_of_date text,
                source text not null,
                fetched_at text not null,
                primary key (symbol, rank)
            )
            """
        )
        conn.execute("create index if not exists idx_ft_holdings_symbol on ft_top_holdings(symbol)")


def save_rows(db_path: Path, rows: list[PriceRow]) -> None:
    init_db(db_path)
    with sqlite3.connect(db_path) as conn:
        conn.executemany(
            """
            insert into ft_historical_prices (
                symbol, ft_issue_id, price_date, open, high, low, close, volume, source, fetched_at
            ) values (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            on conflict(symbol, price_date) do update set
                ft_issue_id = excluded.ft_issue_id,
                open = excluded.open,
                high = excluded.high,
                low = excluded.low,
                close = excluded.close,
                volume = excluded.volume,
                source = excluded.source,
                fetched_at = excluded.fetched_at
            """,
            [
                (
                    row.symbol,
                    row.ft_issue_id,
                    row.price_date,
                    row.open,
                    row.high,
                    row.low,
                    row.close,
                    row.volume,
                    row.source,
                    row.fetched_at,
                )
                for row in rows
            ],
        )


def save_profile_rows(db_path: Path, rows: list[dict[str, str]]) -> None:
    init_db(db_path)
    with sqlite3.connect(db_path) as conn:
        symbols = sorted({row["symbol"] for row in rows})
        conn.executemany("delete from ft_profile_investment where symbol = ?", [(symbol,) for symbol in symbols])
        conn.executemany(
            """
            insert into ft_profile_investment (
                symbol, ft_issue_id, section, field, value, source, fetched_at
            ) values (?, ?, ?, ?, ?, ?, ?)
            on conflict(symbol, section, field) do update set
                ft_issue_id = excluded.ft_issue_id,
                value = excluded.value,
                source = excluded.source,
                fetched_at = excluded.fetched_at
            """,
            [
                (
                    row["symbol"],
                    row["ft_issue_id"],
                    row["section"],
                    row["field"],
                    row["value"],
                    row["source"],
                    row["fetched_at"],
                )
                for row in rows
            ],
        )


def save_risk_rows(db_path: Path, rows: list[dict[str, str]]) -> None:
    init_db(db_path)
    if not rows:
        return
    with sqlite3.connect(db_path) as conn:
        symbols = sorted({row["symbol"] for row in rows})
        conn.executemany("delete from ft_risk_measures where symbol = ?", [(symbol,) for symbol in symbols])
        conn.executemany(
            """
            insert into ft_risk_measures (
                symbol, ft_issue_id, period, metric, fund_value, category_average,
                benchmark_used, as_of_date, source, fetched_at
            ) values (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            on conflict(symbol, period, metric) do update set
                ft_issue_id = excluded.ft_issue_id,
                fund_value = excluded.fund_value,
                category_average = excluded.category_average,
                benchmark_used = excluded.benchmark_used,
                as_of_date = excluded.as_of_date,
                source = excluded.source,
                fetched_at = excluded.fetched_at
            """,
            [
                (
                    row["symbol"],
                    row["ft_issue_id"],
                    row["period"],
                    row["metric"],
                    row["fund_value"],
                    row["category_average"],
                    row["benchmark_used"],
                    row["as_of_date"],
                    row["source"],
                    row["fetched_at"],
                )
                for row in rows
            ],
        )


def save_top_holdings_rows(db_path: Path, rows: list[dict[str, str]]) -> None:
    init_db(db_path)
    if not rows:
        return
    with sqlite3.connect(db_path) as conn:
        symbols = sorted({row["symbol"] for row in rows})
        conn.executemany("delete from ft_top_holdings where symbol = ?", [(symbol,) for symbol in symbols])
        conn.executemany(
            """
            insert into ft_top_holdings (
                symbol, ft_issue_id, rank, holding_name, holding_symbol,
                one_year_change, portfolio_weight, long_allocation,
                top10_portfolio_percent, as_of_date, source, fetched_at
            ) values (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            on conflict(symbol, rank) do update set
                ft_issue_id = excluded.ft_issue_id,
                holding_name = excluded.holding_name,
                holding_symbol = excluded.holding_symbol,
                one_year_change = excluded.one_year_change,
                portfolio_weight = excluded.portfolio_weight,
                long_allocation = excluded.long_allocation,
                top10_portfolio_percent = excluded.top10_portfolio_percent,
                as_of_date = excluded.as_of_date,
                source = excluded.source,
                fetched_at = excluded.fetched_at
            """,
            [
                (
                    row["symbol"],
                    row["ft_issue_id"],
                    int(row["rank"]),
                    row["holding_name"],
                    row["holding_symbol"],
                    row["one_year_change"],
                    row["portfolio_weight"],
                    row["long_allocation"],
                    row["top10_portfolio_percent"],
                    row["as_of_date"],
                    row["source"],
                    row["fetched_at"],
                )
                for row in rows
            ],
        )


def write_csv(path: Path, rows: list[PriceRow]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    ordered = sorted(rows, key=lambda row: row.price_date)
    with path.open("w", encoding="utf-8", newline="") as handle:
        writer = csv.DictWriter(
            handle,
            fieldnames=["symbol", "date", "open", "high", "low", "close", "volume", "source", "fetched_at"],
        )
        writer.writeheader()
        for row in ordered:
            writer.writerow(
                {
                    "symbol": row.symbol,
                    "date": row.price_date,
                    "open": row.open,
                    "high": row.high,
                    "low": row.low,
                    "close": row.close,
                    "volume": row.volume,
                    "source": row.source,
                    "fetched_at": row.fetched_at,
                }
            )


def write_profile_csv(path: Path, rows: list[dict[str, str]]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("w", encoding="utf-8", newline="") as handle:
        writer = csv.DictWriter(
            handle,
            fieldnames=["symbol", "section", "field", "value", "source", "fetched_at"],
        )
        writer.writeheader()
        for row in rows:
            writer.writerow(
                {
                    "symbol": row["symbol"],
                    "section": row["section"],
                    "field": row["field"],
                    "value": row["value"],
                    "source": row["source"],
                    "fetched_at": row["fetched_at"],
                }
            )


def write_risk_csv(path: Path, rows: list[dict[str, str]]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("w", encoding="utf-8", newline="") as handle:
        writer = csv.DictWriter(
            handle,
            fieldnames=[
                "symbol",
                "period",
                "metric",
                "fund_value",
                "category_average",
                "benchmark_used",
                "as_of_date",
                "source",
                "fetched_at",
            ],
        )
        writer.writeheader()
        for row in rows:
            writer.writerow({key: row.get(key, "") for key in writer.fieldnames})


def write_top_holdings_csv(path: Path, rows: list[dict[str, str]]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("w", encoding="utf-8", newline="") as handle:
        writer = csv.DictWriter(
            handle,
            fieldnames=[
                "symbol",
                "rank",
                "holding_name",
                "holding_symbol",
                "one_year_change",
                "portfolio_weight",
                "long_allocation",
                "top10_portfolio_percent",
                "as_of_date",
                "source",
                "fetched_at",
            ],
        )
        writer.writeheader()
        for row in rows:
            writer.writerow({key: row.get(key, "") for key in writer.fieldnames})


def write_raw(path: Path, data: dict[str, Any]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")


def row_dict(row: sqlite3.Row) -> dict[str, Any]:
    return {key: row[key] for key in row.keys()}


def export_database_json(db_path: Path = DEFAULT_DB_PATH, output_path: Path | None = None) -> dict[str, Any]:
    init_db(db_path)
    output_path = output_path or (DEFAULT_OUTPUT_ROOT / "exports" / "ft_historical_prices_database.json")
    output_path.parent.mkdir(parents=True, exist_ok=True)
    with sqlite3.connect(db_path) as conn:
        conn.row_factory = sqlite3.Row
        symbols = [
            row_dict(row)
            for row in conn.execute(
                """
                with symbol_keys as (
                    select symbol from ft_historical_prices
                    union
                    select symbol from ft_profile_investment
                    union
                    select symbol from ft_risk_measures
                    union
                    select symbol from ft_top_holdings
                ),
                price_symbols as (
                    select
                        symbol,
                        max(ft_issue_id) as ftIssueId,
                        count(*) as rowCount,
                        min(price_date) as startDate,
                        max(price_date) as endDate
                    from ft_historical_prices
                    group by symbol
                ),
                profile_symbols as (
                    select
                        symbol,
                        max(ft_issue_id) as ftIssueId,
                        count(*) as profileRowCount
                    from ft_profile_investment
                    group by symbol
                ),
                risk_symbols as (
                    select
                        symbol,
                        max(ft_issue_id) as ftIssueId,
                        count(*) as riskRowCount
                    from ft_risk_measures
                    group by symbol
                ),
                holdings_symbols as (
                    select
                        symbol,
                        max(ft_issue_id) as ftIssueId,
                        count(*) as holdingsRowCount
                    from ft_top_holdings
                    group by symbol
                ),
                display_names as (
                    select symbol, max(value) as displayName
                    from ft_profile_investment
                    where section = 'metadata' and field = 'FT display name' and value is not null and value <> ''
                    group by symbol
                )
                select
                    k.symbol,
                    coalesce(p.ftIssueId, q.ftIssueId, r.ftIssueId, h.ftIssueId, '') as ftIssueId,
                    coalesce(d.displayName, '') as displayName,
                    coalesce(p.rowCount, 0) as rowCount,
                    coalesce(q.profileRowCount, 0) as profileRowCount,
                    coalesce(r.riskRowCount, 0) as riskRowCount,
                    coalesce(h.holdingsRowCount, 0) as holdingsRowCount,
                    coalesce(p.startDate, '') as startDate,
                    coalesce(p.endDate, '') as endDate
                from symbol_keys k
                left join price_symbols p on upper(p.symbol) = upper(k.symbol)
                left join profile_symbols q on upper(q.symbol) = upper(k.symbol)
                left join risk_symbols r on upper(r.symbol) = upper(k.symbol)
                left join holdings_symbols h on upper(h.symbol) = upper(k.symbol)
                left join display_names d on upper(d.symbol) = upper(k.symbol)
                order by k.symbol
                """
            )
        ]
        for item in symbols:
            item["displayLabel"] = f"{item['symbol']} - {item['displayName']}" if item.get("displayName") else item.get("symbol", "")

        prices = [
            row_dict(row)
            for row in conn.execute(
                """
                select
                    symbol,
                    price_date as date,
                    open,
                    high,
                    low,
                    close,
                    volume,
                    source,
                    fetched_at as fetchedAt
                from ft_historical_prices
                order by symbol, price_date desc
                """
            )
        ]
        profile = [
            row_dict(row)
            for row in conn.execute(
                """
                select
                    symbol,
                    section,
                    field,
                    value,
                    source,
                    fetched_at as fetchedAt
                from ft_profile_investment
                order by symbol,
                    case section when 'metadata' then 0 when 'profile' then 1 when 'investment' then 2 else 3 end,
                    field
                """
            )
        ]
        risk = [
            row_dict(row)
            for row in conn.execute(
                """
                select
                    symbol,
                    period,
                    metric,
                    fund_value as fundValue,
                    category_average as categoryAverage,
                    benchmark_used as benchmarkUsed,
                    as_of_date as asOfDate,
                    source,
                    fetched_at as fetchedAt
                from ft_risk_measures
                order by symbol,
                    case period when '1 year' then 1 when '3 year' then 2 when '5 years' then 3 else 4 end,
                    metric
                """
            )
        ]
        holdings = [
            row_dict(row)
            for row in conn.execute(
                """
                select
                    symbol,
                    rank,
                    holding_name as holdingName,
                    holding_symbol as holdingSymbol,
                    one_year_change as oneYearChange,
                    portfolio_weight as portfolioWeight,
                    long_allocation as longAllocation,
                    top10_portfolio_percent as top10PortfolioPercent,
                    as_of_date as asOfDate,
                    source,
                    fetched_at as fetchedAt
                from ft_top_holdings
                order by symbol, rank
                """
            )
        ]

    selected_symbol = symbols[0]["symbol"] if symbols else ""
    rows = [row for row in prices if row.get("symbol") == selected_symbol]
    symbol_export_dir = output_path.parent / "symbols"
    symbol_export_dir.mkdir(parents=True, exist_ok=True)
    symbol_files: list[dict[str, Any]] = []
    for symbol_item in symbols:
        symbol = symbol_item.get("symbol", "")
        slug = slug_symbol(symbol)
        symbol_prices = [row for row in prices if row.get("symbol") == symbol]
        symbol_profile = [row for row in profile if row.get("symbol") == symbol]
        symbol_risk = [row for row in risk if row.get("symbol") == symbol]
        symbol_holdings = [row for row in holdings if row.get("symbol") == symbol]
        symbol_payload = {
            "generated_at": utc_now(),
            "source_database": str(db_path),
            "symbol": symbol,
            "symbolSlug": slug,
            "metadata": symbol_item,
            "counts": {
                "price_rows": len(symbol_prices),
                "profile_rows": len(symbol_profile),
                "risk_rows": len(symbol_risk),
                "holding_rows": len(symbol_holdings),
            },
            "prices": symbol_prices,
            "profile": symbol_profile,
            "risk": symbol_risk,
            "holdings": symbol_holdings,
            "source": f"Google Drive JSON: exports/symbols/{slug}.json",
        }
        symbol_path = symbol_export_dir / f"{slug}.json"
        symbol_path.write_text(json.dumps(symbol_payload, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
        symbol_files.append({
            "symbol": symbol,
            "symbolSlug": slug,
            "fileName": f"{slug}.json",
            "relativePath": f"exports/symbols/{slug}.json",
            "counts": symbol_payload["counts"],
        })

    payload = {
        "generated_at": utc_now(),
        "source_database": str(db_path),
        "counts": {
            "symbols": len(symbols),
            "price_rows": len(prices),
            "profile_rows": len(profile),
            "risk_rows": len(risk),
            "holding_rows": len(holdings),
        },
        "symbols": symbols,
        "symbolFiles": symbol_files,
        "selectedSymbol": selected_symbol,
        "selectedDisplayName": next((item.get("displayName", "") for item in symbols if item.get("symbol") == selected_symbol), ""),
        "selectedDisplayLabel": next((item.get("displayLabel", "") for item in symbols if item.get("symbol") == selected_symbol), selected_symbol),
        "rows": rows,
        "prices": prices,
        "profile": profile,
        "risk": risk,
        "holdings": holdings,
        "stats": [],
        "source": "Google Drive JSON: ft_historical_prices_database.json",
    }
    output_path.write_text(json.dumps(payload, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    return {"output_path": str(output_path), **payload["counts"]}


def sync_historical_prices(
    symbol: str,
    start_date: str,
    end_date: str,
    output_root: Path = DEFAULT_OUTPUT_ROOT,
    db_path: Path = DEFAULT_DB_PATH,
    product_path: str | None = None,
) -> dict[str, Any]:
    symbol_slug = slug_symbol(symbol)
    ft_issue_id, resolved_product_path = resolve_issue_id(symbol, product_path)
    start = date.fromisoformat(start_date)
    end = date.fromisoformat(end_date)
    all_rows: list[PriceRow] = []
    range_summaries: list[dict[str, Any]] = []

    for range_start, range_end in yearly_ranges(start, end):
        data, rows = fetch_range(symbol, ft_issue_id, range_start, range_end, resolved_product_path)
        all_rows.extend(rows)
        range_summaries.append(
            {
                "startDate": range_start,
                "endDate": range_end,
                "rows": len(rows),
            }
        )

    deduped = {row.price_date: row for row in all_rows}
    final_rows = list(deduped.values())
    save_rows(db_path, final_rows)
    csv_path = output_root / "prices" / f"{symbol_slug}.csv"
    write_csv(csv_path, final_rows)

    return {
        "symbol": symbol,
        "symbolSlug": symbol_slug,
        "ftIssueId": ft_issue_id,
        "productPath": resolved_product_path,
        "startDate": start_date,
        "endDate": end_date,
        "totalRows": len(final_rows),
        "ranges": range_summaries,
        "csvPath": str(csv_path),
        "sqlitePath": str(db_path),
    }


def sync_qualitative_data(
    symbol: str,
    output_root: Path = DEFAULT_OUTPUT_ROOT,
    db_path: Path = DEFAULT_DB_PATH,
    product_path: str | None = None,
) -> dict[str, Any]:
    symbol_slug = slug_symbol(symbol)
    ft_issue_id, resolved_product_path = resolve_issue_id(symbol, product_path)

    profile_raw, profile_rows = fetch_profile_investment(symbol, ft_issue_id, resolved_product_path)
    risk_raw, risk_rows = fetch_risk_measures(symbol, ft_issue_id, resolved_product_path)
    holdings_raw, holdings_rows = fetch_top_holdings(symbol, ft_issue_id, resolved_product_path)
    as_of_date = (
        holdings_raw.get("as_of_date")
        or risk_raw.get("as_of_date")
        or profile_raw.get("as_of_date")
        or utc_now()[:10]
    )
    snapshot_dir = output_root / "qualitative" / f"as_of_{as_of_date}" / symbol_slug

    profile_raw_path = snapshot_dir / "summary_profile_investment.raw.json"
    write_raw(profile_raw_path, profile_raw)
    save_profile_rows(db_path, profile_rows)
    profile_csv_path = snapshot_dir / "profile_investment.csv"
    write_profile_csv(profile_csv_path, profile_rows)

    risk_raw_path = snapshot_dir / "risk.raw.json"
    write_raw(risk_raw_path, risk_raw)
    save_risk_rows(db_path, risk_rows)
    risk_csv_path = snapshot_dir / "risk_measures.csv"
    write_risk_csv(risk_csv_path, risk_rows)

    holdings_raw_path = snapshot_dir / "holdings.raw.json"
    write_raw(holdings_raw_path, holdings_raw)
    save_top_holdings_rows(db_path, holdings_rows)
    holdings_csv_path = snapshot_dir / "holdings.csv"
    write_top_holdings_csv(holdings_csv_path, holdings_rows)

    return {
        "symbol": symbol,
        "symbolSlug": symbol_slug,
        "ftIssueId": ft_issue_id,
        "productPath": resolved_product_path,
        "displayName": profile_raw.get("display_name") or risk_raw.get("display_name") or "",
        "profileRows": len(profile_rows),
        "riskRows": len(risk_rows),
        "holdingsRows": len(holdings_rows),
        "asOfDate": as_of_date,
        "profileAsOfDate": profile_raw.get("as_of_date") or "",
        "riskAsOfDate": risk_raw.get("as_of_date") or "",
        "holdingsAsOfDate": holdings_raw.get("as_of_date") or "",
        "profileCsvPath": str(profile_csv_path),
        "riskCsvPath": str(risk_csv_path),
        "holdingsCsvPath": str(holdings_csv_path),
        "sqlitePath": str(db_path),
        "snapshotDir": str(snapshot_dir),
    }


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--symbol", default="IXN:PCQ:USD")
    parser.add_argument("--start-date")
    parser.add_argument("--end-date")
    parser.add_argument("--output-root", type=Path, default=DEFAULT_OUTPUT_ROOT)
    parser.add_argument("--db-path", type=Path, default=DEFAULT_DB_PATH)
    parser.add_argument("--qualitative-only", action="store_true", help="Fetch summary/profile and risk data without historical prices")
    parser.add_argument("--export-json", action="store_true", help="Export the SQLite database to a Drive-friendly JSON file")
    parser.add_argument(
        "--output",
        type=Path,
        default=DEFAULT_OUTPUT_ROOT / "exports" / "ft_historical_prices_database.json",
        help="Output path for --export-json",
    )
    args = parser.parse_args()

    if args.export_json:
        result = export_database_json(args.db_path, args.output)
        print(json.dumps(result, ensure_ascii=False, indent=2))
        return 0

    if args.qualitative_only:
        result = sync_qualitative_data(args.symbol, args.output_root, args.db_path)
        print(f"resolved_issue_id={result['ftIssueId']}")
        print(f"display_name={result.get('displayName') or ''}")
        print(f"profile_csv={result['profileCsvPath']}")
        print(f"risk_csv={result['riskCsvPath']}")
        print(f"sqlite={result['sqlitePath']}")
        print(f"risk_rows={result['riskRows']}")
        print(f"holdings_rows={result['holdingsRows']}")
        print(f"risk_as_of_date={result.get('riskAsOfDate') or ''}")
        print(f"holdings_as_of_date={result.get('holdingsAsOfDate') or ''}")
        return 0

    if not args.start_date or not args.end_date:
        parser.error("--start-date and --end-date are required unless --qualitative-only is used")

    result = sync_historical_prices(args.symbol, args.start_date, args.end_date, args.output_root, args.db_path)
    for range_summary in result["ranges"]:
        print(f"{args.symbol} {range_summary['startDate']}..{range_summary['endDate']}: {range_summary['rows']} rows")
    print(f"resolved_issue_id={result['ftIssueId']}")
    print(f"csv={result['csvPath']}")
    print(f"sqlite={result['sqlitePath']}")
    print(f"total_rows={result['totalRows']}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
