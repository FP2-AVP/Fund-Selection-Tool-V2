#!/usr/bin/env python3
"""Manage weekly SEC dividend-history storage in SQLite.

This script keeps the dividend-history workflow intentionally small:

1. Initialize/upgrade the SQLite tables used by dividend history.
2. Import tracked funds from the AVP performance list and map them to SEC proj_id.
3. Fetch SEC dividend-history rows and insert only new records.
"""

from __future__ import annotations

import argparse
import hashlib
import json
import os
import shutil
import sqlite3
import sys
import time
import urllib.error
import urllib.parse
import urllib.request
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import datetime, timezone
from pathlib import Path
from typing import Any


PROJECT_ROOT = Path(__file__).resolve().parents[1]
DEFAULT_DB_PATH = PROJECT_ROOT / "Data" / "sec_fund_store.db"
DEFAULT_SEC_MASTER_PATH = PROJECT_ROOT / "Data" / "Data For SEC API - 2026-Q1.json"
DEFAULT_AVP_PERFORMANCE_PATH = PROJECT_ROOT / "Data" / "Fund Key Performance AVP - 2026-Q1.json"
DEFAULT_RAW_DIR = PROJECT_ROOT / "Data" / "raw_snapshots" / "dividend_history"
DEFAULT_BACKUP_DIR = PROJECT_ROOT / "Data" / "backups"
DEFAULT_SYSTEM_ROOT = PROJECT_ROOT / "Dividend Database"
DEFAULT_EXPORT_DIR = DEFAULT_SYSTEM_ROOT / "exports"
DEFAULT_DRIVE_FOLDER_ID = "1SBJshoukb5O9hdTawTtexVQzo4oLfW3u"
SEC_DIVIDEND_HISTORY_URL = "https://api.sec.or.th/v2/fund/daily-info/dividend-history"
TRACKING_LABELS = {"dividend", "redemption"}


SCHEMA_SQL = """
pragma foreign_keys = on;

create table if not exists tracked_funds (
    id integer primary key autoincrement,
    proj_id text not null unique,
    fund_code text not null,
    fund_name text,
    asset_house text,
    tracking_type text not null,
    avp_fund_type text,
    sec_proj_name_th text,
    sec_proj_name_en text,
    sec_fund_status text,
    sec_unique_id text,
    sec_fund_class_name text,
    source_file text,
    source_updated_at text,
    is_active integer not null default 1,
    created_at text not null default current_timestamp,
    updated_at text not null default current_timestamp
);

create table if not exists tracked_fund_classes (
    id integer primary key autoincrement,
    fund_code text not null unique,
    proj_id text not null,
    fund_name text,
    asset_house text,
    tracking_type text not null,
    avp_fund_type text,
    sec_proj_name_th text,
    sec_proj_name_en text,
    sec_fund_status text,
    sec_unique_id text,
    sec_fund_class_name text,
    source_file text,
    source_updated_at text,
    is_active integer not null default 1,
    created_at text not null default current_timestamp,
    updated_at text not null default current_timestamp
);

create table if not exists fund_dividend_history (
    id integer primary key autoincrement,
    proj_id text not null,
    unique_id text not null default '',
    class_abbr_name text not null default '',
    book_close_date text not null default '',
    dividend_date text not null,
    dividend_value real not null,
    last_upd_date text,
    payload_json text not null,
    payload_hash text not null,
    first_seen_at text not null default current_timestamp,
    last_seen_at text not null default current_timestamp,
    sync_run_id integer,
    unique (
        proj_id,
        unique_id,
        class_abbr_name,
        book_close_date,
        dividend_date,
        dividend_value
    )
);

create table if not exists finnomena_dividend_history (
    id integer primary key autoincrement,
    proj_id text not null,
    fund_code text not null,
    fund_name text,
    asset_house text,
    finnomena_fund_id text not null,
    xd_date text not null,
    pay_date text not null,
    dividend_value real not null,
    payload_json text not null,
    payload_hash text not null,
    first_seen_at text not null default current_timestamp,
    last_seen_at text not null default current_timestamp,
    sync_run_id integer,
    unique (
        proj_id,
        fund_code,
        finnomena_fund_id,
        xd_date,
        pay_date,
        dividend_value
    )
);

create index if not exists idx_tracked_funds_code on tracked_funds(fund_code);
create index if not exists idx_tracked_funds_type on tracked_funds(tracking_type);
create index if not exists idx_tracked_fund_classes_proj on tracked_fund_classes(proj_id);
create index if not exists idx_tracked_fund_classes_type on tracked_fund_classes(tracking_type);
create index if not exists idx_dividend_history_proj on fund_dividend_history(proj_id);
create index if not exists idx_dividend_history_date on fund_dividend_history(dividend_date);
create index if not exists idx_dividend_history_last_upd on fund_dividend_history(last_upd_date);
create index if not exists idx_finnomena_dividend_proj on finnomena_dividend_history(proj_id);
create index if not exists idx_finnomena_dividend_fund_code on finnomena_dividend_history(fund_code);
create index if not exists idx_finnomena_dividend_pay_date on finnomena_dividend_history(pay_date);

create table if not exists dividend_sync_items (
    id integer primary key autoincrement,
    sync_run_id integer,
    proj_id text not null,
    fund_code text,
    status text not null,
    fetched_rows integer not null default 0,
    inserted_rows integer not null default 0,
    updated_rows integer not null default 0,
    skipped_rows integer not null default 0,
    error_message text,
    fetched_at text not null default current_timestamp
);

create table if not exists finnomena_sync_items (
    id integer primary key autoincrement,
    sync_run_id integer,
    proj_id text not null,
    fund_code text not null,
    finnomena_fund_id text,
    status text not null,
    fetched_rows integer not null default 0,
    inserted_rows integer not null default 0,
    updated_rows integer not null default 0,
    skipped_rows integer not null default 0,
    error_message text,
    fetched_at text not null default current_timestamp
);
"""


def utc_now() -> str:
    return datetime.now(timezone.utc).replace(microsecond=0).isoformat().replace("+00:00", "Z")


def normalized_text(value: Any) -> str:
    if value is None:
        return ""
    return str(value).strip()


def load_dotenv(path: Path = PROJECT_ROOT / ".env") -> None:
    if not path.exists():
        return
    for line in path.read_text(encoding="utf-8").splitlines():
        stripped = line.strip()
        if not stripped or stripped.startswith("#") or "=" not in stripped:
            continue
        key, value = stripped.split("=", 1)
        key = key.strip()
        value = value.strip().strip('"').strip("'")
        if key and key not in os.environ:
            os.environ[key] = value


def compact_date(value: Any) -> str:
    text = normalized_text(value)
    if not text:
        return ""
    return text[:10]


def payload_hash(payload: dict[str, Any]) -> str:
    raw = json.dumps(payload, ensure_ascii=False, sort_keys=True, separators=(",", ":"))
    return hashlib.sha256(raw.encode("utf-8")).hexdigest()


def read_sheet_json(path: Path) -> tuple[list[str], list[dict[str, Any]]]:
    data = json.loads(path.read_text(encoding="utf-8"))
    if not data:
        return [], []
    headers = [normalized_text(value) for value in data[0]]
    rows = []
    for raw_row in data[1:]:
        row = {
            header: raw_row[index] if index < len(raw_row) else ""
            for index, header in enumerate(headers)
            if header
        }
        rows.append(row)
    return headers, rows


def connect(db_path: Path) -> sqlite3.Connection:
    db_path.parent.mkdir(parents=True, exist_ok=True)
    conn = sqlite3.connect(db_path)
    conn.row_factory = sqlite3.Row
    conn.execute("pragma foreign_keys = on")
    return conn


def init_db(db_path: Path) -> None:
    with connect(db_path) as conn:
        conn.executescript(SCHEMA_SQL)
        conn.commit()


def ensure_sync_runs_table(conn: sqlite3.Connection) -> None:
    conn.execute(
        """
        create table if not exists sync_runs (
            id integer primary key autoincrement,
            run_type text not null,
            status text not null default 'running',
            started_at text not null default current_timestamp,
            finished_at text,
            notes text
        )
        """
    )


def build_sec_lookup(sec_master_path: Path) -> dict[str, dict[str, Any]]:
    _headers, rows = read_sheet_json(sec_master_path)
    lookup: dict[str, dict[str, Any]] = {}
    for row in rows:
        keys = [
            normalized_text(row.get("proj_abbr_name")),
            normalized_text(row.get("fund_class_name")),
        ]
        for key in keys:
            if key and key.casefold() not in lookup:
                lookup[key.casefold()] = row
    return lookup


def import_tracked_funds(db_path: Path, sec_master_path: Path, avp_performance_path: Path) -> dict[str, int]:
    init_db(db_path)
    sec_lookup = build_sec_lookup(sec_master_path)
    _headers, avp_rows = read_sheet_json(avp_performance_path)
    stats = {
        "source_rows": 0,
        "matched_fund_codes": 0,
        "unmatched_fund_codes": 0,
        "upserted_fund_codes": 0,
        "tracked_proj_ids": 0,
    }

    with connect(db_path) as conn:
        ensure_sync_runs_table(conn)
        for row in avp_rows:
            tracking_type = normalized_text(row.get("Dividend"))
            if tracking_type.casefold() not in TRACKING_LABELS:
                continue
            stats["source_rows"] += 1
            fund_code = normalized_text(row.get("Fund Code"))
            if not fund_code:
                stats["unmatched_fund_codes"] += 1
                continue

            sec_row = sec_lookup.get(fund_code.casefold())
            if not sec_row:
                stats["unmatched_fund_codes"] += 1
                continue

            stats["matched_fund_codes"] += 1
            proj_id = normalized_text(sec_row.get("proj_id"))
            common_values = (
                proj_id,
                fund_code,
                normalized_text(row.get("Name")),
                normalized_text(row.get("Asset House")),
                tracking_type,
                normalized_text(row.get("Fund Type")),
                normalized_text(sec_row.get("proj_name_th")),
                normalized_text(sec_row.get("proj_name_en")),
                normalized_text(sec_row.get("fund_status")),
                normalized_text(sec_row.get("unique_id")),
                normalized_text(sec_row.get("fund_class_name")),
                str(avp_performance_path.relative_to(PROJECT_ROOT)),
                normalized_text(sec_row.get("statistics_last_upd_date")),
            )
            conn.execute(
                """
                insert into tracked_funds (
                    proj_id, fund_code, fund_name, asset_house, tracking_type,
                    avp_fund_type, sec_proj_name_th, sec_proj_name_en,
                    sec_fund_status, sec_unique_id, sec_fund_class_name,
                    source_file, source_updated_at, is_active, updated_at
                ) values (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, 1, current_timestamp)
                on conflict(proj_id) do update set
                    fund_code = excluded.fund_code,
                    fund_name = excluded.fund_name,
                    asset_house = excluded.asset_house,
                    tracking_type = excluded.tracking_type,
                    avp_fund_type = excluded.avp_fund_type,
                    sec_proj_name_th = excluded.sec_proj_name_th,
                    sec_proj_name_en = excluded.sec_proj_name_en,
                    sec_fund_status = excluded.sec_fund_status,
                    sec_unique_id = excluded.sec_unique_id,
                    sec_fund_class_name = excluded.sec_fund_class_name,
                    source_file = excluded.source_file,
                    source_updated_at = excluded.source_updated_at,
                    is_active = 1,
                    updated_at = current_timestamp
                """,
                common_values,
            )
            conn.execute(
                """
                insert into tracked_fund_classes (
                    proj_id, fund_code, fund_name, asset_house, tracking_type,
                    avp_fund_type, sec_proj_name_th, sec_proj_name_en,
                    sec_fund_status, sec_unique_id, sec_fund_class_name,
                    source_file, source_updated_at, is_active, updated_at
                ) values (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, 1, current_timestamp)
                on conflict(fund_code) do update set
                    proj_id = excluded.proj_id,
                    fund_name = excluded.fund_name,
                    asset_house = excluded.asset_house,
                    tracking_type = excluded.tracking_type,
                    avp_fund_type = excluded.avp_fund_type,
                    sec_proj_name_th = excluded.sec_proj_name_th,
                    sec_proj_name_en = excluded.sec_proj_name_en,
                    sec_fund_status = excluded.sec_fund_status,
                    sec_unique_id = excluded.sec_unique_id,
                    sec_fund_class_name = excluded.sec_fund_class_name,
                    source_file = excluded.source_file,
                    source_updated_at = excluded.source_updated_at,
                    is_active = 1,
                    updated_at = current_timestamp
                """,
                common_values,
            )
            stats["upserted_fund_codes"] += 1
        conn.commit()
        stats["tracked_proj_ids"] = conn.execute(
            "select count(distinct proj_id) as c from tracked_fund_classes where is_active = 1"
        ).fetchone()["c"]
    return stats


def api_key_from_args(args: argparse.Namespace) -> str:
    api_key = normalized_text(args.api_key) or normalized_text(os.environ.get("SEC_API_KEY"))
    if not api_key:
        raise RuntimeError("Set SEC_API_KEY or pass --api-key before syncing SEC API data.")
    return api_key


def fetch_dividend_history(api_key: str, proj_id: str, page_size: int) -> tuple[list[dict[str, Any]], dict[str, Any]]:
    headers = {
        "Content-Type": "application/json",
        "Cache-Control": "no-cache",
        "Ocp-Apim-Subscription-Key": api_key,
    }
    next_cursor = ""
    all_items: list[dict[str, Any]] = []
    pages: list[dict[str, Any]] = []

    while True:
        params = {"proj_id": proj_id, "page_size": page_size}
        if next_cursor:
            params["next_cursor"] = next_cursor
        url = SEC_DIVIDEND_HISTORY_URL + "?" + urllib.parse.urlencode(params)
        request = urllib.request.Request(url, headers=headers)
        with urllib.request.urlopen(request, timeout=45) as response:
            payload = json.loads(response.read().decode("utf-8"))
        items = payload.get("items", [])
        if not isinstance(items, list):
            items = []
        all_items.extend(items)
        pages.append(
            {
                "message": payload.get("message"),
                "page_size": payload.get("page_size"),
                "next_cursor": payload.get("next_cursor") or "",
                "items": items,
            }
        )
        next_cursor = normalized_text(payload.get("next_cursor"))
        if not next_cursor:
            break
        time.sleep(0.1)

    return all_items, {"proj_id": proj_id, "fetched_at": utc_now(), "pages": pages}


def save_raw_snapshot(raw_dir: Path, proj_id: str, payload: dict[str, Any]) -> Path:
    today = datetime.now(timezone.utc).date().isoformat()
    out_dir = raw_dir / today
    out_dir.mkdir(parents=True, exist_ok=True)
    out_path = out_dir / f"{proj_id}.json"
    out_path.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
    return out_path


def upsert_dividend_rows(
    conn: sqlite3.Connection,
    rows: list[dict[str, Any]],
    sync_run_id: int | None,
) -> tuple[int, int, int]:
    inserted = 0
    updated = 0
    skipped = 0
    for row in rows:
        proj_id = normalized_text(row.get("proj_id"))
        dividend_date = normalized_text(row.get("dividend_date"))
        dividend_value_raw = row.get("dividend_value")
        if not proj_id or not dividend_date or dividend_value_raw in {"", None}:
            skipped += 1
            continue

        row_hash = payload_hash(row)
        payload_json = json.dumps(row, ensure_ascii=False, sort_keys=True)
        key_values = (
            proj_id,
            normalized_text(row.get("unique_id")),
            normalized_text(row.get("class_abbr_name")),
            normalized_text(row.get("book_close_date")),
            dividend_date,
            float(dividend_value_raw),
        )
        existing = conn.execute(
            """
            select id, payload_hash, coalesce(last_upd_date, '') as last_upd_date
            from fund_dividend_history
            where proj_id = ?
              and unique_id = ?
              and class_abbr_name = ?
              and book_close_date = ?
              and dividend_date = ?
              and dividend_value = ?
            """,
            key_values,
        ).fetchone()

        if not existing:
            conn.execute(
                """
                insert into fund_dividend_history (
                    proj_id, unique_id, class_abbr_name, book_close_date,
                    dividend_date, dividend_value, last_upd_date,
                    payload_json, payload_hash, sync_run_id
                ) values (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                """,
                (
                    *key_values,
                    normalized_text(row.get("last_upd_date")),
                    payload_json,
                    row_hash,
                    sync_run_id,
                ),
            )
            inserted += 1
            continue

        incoming_last_upd = normalized_text(row.get("last_upd_date"))
        if existing["payload_hash"] == row_hash and incoming_last_upd <= existing["last_upd_date"]:
            conn.execute(
                "update fund_dividend_history set last_seen_at = current_timestamp, sync_run_id = ? where id = ?",
                (sync_run_id, existing["id"]),
            )
            skipped += 1
            continue

        conn.execute(
            """
            update fund_dividend_history
            set last_seen_at = current_timestamp,
                last_upd_date = ?,
                payload_json = ?,
                payload_hash = ?,
                sync_run_id = ?
            where id = ?
            """,
            (
                max(incoming_last_upd, existing["last_upd_date"]),
                payload_json,
                row_hash,
                sync_run_id,
                existing["id"],
            ),
        )
        updated += 1
    return inserted, updated, skipped


def upsert_finnomena_dividend_rows(
    conn: sqlite3.Connection,
    resolved: dict[str, Any],
    rows: list[dict[str, Any]],
    sync_run_id: int | None,
) -> tuple[int, int, int]:
    inserted = 0
    updated = 0
    skipped = 0
    for row in rows:
        xd_date = compact_date(row.get("xd_date"))
        pay_date = compact_date(row.get("pay_date"))
        dividend_value_raw = row.get("value")
        if not xd_date or not pay_date or dividend_value_raw in {"", None}:
            skipped += 1
            continue

        normalized_row = {
            "proj_id": normalized_text(resolved.get("proj_id")),
            "fund_code": normalized_text(resolved.get("fund_code")),
            "fund_name": normalized_text(resolved.get("fund_name")),
            "asset_house": normalized_text(resolved.get("asset_house")),
            "finnomena_fund_id": normalized_text(resolved.get("finnomena_fund_id")),
            "xd_date": xd_date,
            "pay_date": pay_date,
            "dividend_value": float(dividend_value_raw),
            "raw": row,
        }
        row_hash = payload_hash(normalized_row)
        payload_json = json.dumps(normalized_row, ensure_ascii=False, sort_keys=True)
        key_values = (
            normalized_row["proj_id"],
            normalized_row["fund_code"],
            normalized_row["finnomena_fund_id"],
            normalized_row["xd_date"],
            normalized_row["pay_date"],
            normalized_row["dividend_value"],
        )
        existing = conn.execute(
            """
            select id, payload_hash
            from finnomena_dividend_history
            where proj_id = ?
              and fund_code = ?
              and finnomena_fund_id = ?
              and xd_date = ?
              and pay_date = ?
              and dividend_value = ?
            """,
            key_values,
        ).fetchone()

        if not existing:
            conn.execute(
                """
                insert into finnomena_dividend_history (
                    proj_id, fund_code, fund_name, asset_house, finnomena_fund_id,
                    xd_date, pay_date, dividend_value, payload_json, payload_hash,
                    sync_run_id
                ) values (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                """,
                (
                    normalized_row["proj_id"],
                    normalized_row["fund_code"],
                    normalized_row["fund_name"],
                    normalized_row["asset_house"],
                    normalized_row["finnomena_fund_id"],
                    normalized_row["xd_date"],
                    normalized_row["pay_date"],
                    normalized_row["dividend_value"],
                    payload_json,
                    row_hash,
                    sync_run_id,
                ),
            )
            inserted += 1
            continue

        if existing["payload_hash"] == row_hash:
            conn.execute(
                "update finnomena_dividend_history set last_seen_at = current_timestamp, sync_run_id = ? where id = ?",
                (sync_run_id, existing["id"]),
            )
            skipped += 1
            continue

        conn.execute(
            """
            update finnomena_dividend_history
            set fund_name = ?,
                asset_house = ?,
                payload_json = ?,
                payload_hash = ?,
                last_seen_at = current_timestamp,
                sync_run_id = ?
            where id = ?
            """,
            (
                normalized_row["fund_name"],
                normalized_row["asset_house"],
                payload_json,
                row_hash,
                sync_run_id,
                existing["id"],
            ),
        )
        updated += 1
    return inserted, updated, skipped


def create_sync_run(conn: sqlite3.Connection, run_type: str, notes: str) -> int:
    ensure_sync_runs_table(conn)
    row = conn.execute(
        "insert into sync_runs (run_type, status, notes) values (?, 'running', ?) returning id",
        (run_type, notes),
    ).fetchone()
    return int(row["id"])


def finish_sync_run(conn: sqlite3.Connection, sync_run_id: int, status: str, notes: str) -> None:
    conn.execute(
        """
        update sync_runs
        set status = ?, finished_at = current_timestamp, notes = ?
        where id = ?
        """,
        (status, notes, sync_run_id),
    )


def selected_tracked_funds(conn: sqlite3.Connection, proj_id: str = "", limit: int = 0) -> list[sqlite3.Row]:
    params: list[Any] = []
    where = "where is_active = 1"
    if proj_id:
        where += " and proj_id = ?"
        params.append(proj_id)
    sql = f"""
        select proj_id, min(fund_code) as fund_code, min(fund_name) as fund_name,
               group_concat(distinct tracking_type) as tracking_type
        from tracked_fund_classes
        {where}
        group by proj_id
        order by fund_code
    """
    if limit:
        sql += " limit ?"
        params.append(limit)
    return list(conn.execute(sql, params).fetchall())


def sync_dividend_history(args: argparse.Namespace) -> dict[str, int]:
    api_key = api_key_from_args(args)
    init_db(args.db_path)
    args.raw_dir.mkdir(parents=True, exist_ok=True)
    stats = {
        "funds": 0,
        "fetched_rows": 0,
        "inserted_rows": 0,
        "updated_rows": 0,
        "skipped_rows": 0,
        "error_funds": 0,
    }

    with connect(args.db_path) as conn:
        sync_run_id = create_sync_run(conn, "dividend_history", f"proj_id={args.proj_id or 'all'}")
        conn.commit()
        funds = selected_tracked_funds(conn, args.proj_id, args.limit)
        if not funds:
            finish_sync_run(conn, sync_run_id, "failed", "No tracked funds found. Run import-tracked-funds first.")
            conn.commit()
            raise RuntimeError("No tracked funds found. Run import-tracked-funds first.")

        for fund in funds:
            proj_id = fund["proj_id"]
            fund_code = fund["fund_code"]
            try:
                items, raw_payload = fetch_dividend_history(api_key, proj_id, args.page_size)
                save_raw_snapshot(args.raw_dir, proj_id, raw_payload)
                inserted, updated, skipped = upsert_dividend_rows(conn, items, sync_run_id)
                status = "success"
                error_message = ""
                stats["funds"] += 1
                stats["fetched_rows"] += len(items)
                stats["inserted_rows"] += inserted
                stats["updated_rows"] += updated
                stats["skipped_rows"] += skipped
            except (urllib.error.URLError, TimeoutError, json.JSONDecodeError, RuntimeError) as exc:
                status = "error"
                error_message = str(exc)
                inserted = updated = skipped = 0
                items = []
                stats["error_funds"] += 1

            conn.execute(
                """
                insert into dividend_sync_items (
                    sync_run_id, proj_id, fund_code, status, fetched_rows,
                    inserted_rows, updated_rows, skipped_rows, error_message
                ) values (?, ?, ?, ?, ?, ?, ?, ?, ?)
                """,
                (
                    sync_run_id,
                    proj_id,
                    fund_code,
                    status,
                    len(items),
                    inserted,
                    updated,
                    skipped,
                    error_message,
                ),
            )
            conn.commit()
            if args.sleep_seconds:
                time.sleep(args.sleep_seconds)

        final_status = "success" if stats["error_funds"] == 0 else "partial"
        finish_sync_run(conn, sync_run_id, final_status, json.dumps(stats, ensure_ascii=False, sort_keys=True))
        conn.commit()
    return stats


def selected_finnomena_funds(db_path: Path, limit: int = 0) -> list[dict[str, Any]]:
    from fetch_finnomena_dividend_by_proj_id import build_proj_id_mapping

    mapping = build_proj_id_mapping()
    with connect(db_path) as conn:
        tracked_codes = {
            normalized_text(row["fund_code"]).upper()
            for row in conn.execute("select fund_code from tracked_fund_classes where is_active = 1")
        }
    items = [item for group in mapping.values() for item in group]
    if tracked_codes:
        items = [
            item for item in items
            if normalized_text(item.get("fund_code")).upper() in tracked_codes
        ]
    items.sort(key=lambda item: (normalized_text(item.get("fund_code")), normalized_text(item.get("proj_id"))))
    if limit:
        items = items[:limit]
    return items


def fetch_one_finnomena_dividend(resolved: dict[str, Any]) -> tuple[dict[str, Any], list[dict[str, Any]], str]:
    from fetch_finnomena_dividend_by_proj_id import fetch_finnomena_dividend_history

    payload = fetch_finnomena_dividend_history(normalized_text(resolved.get("finnomena_fund_id")))
    dividends = ((payload or {}).get("data") or {}).get("dividends") or []
    if not isinstance(dividends, list):
        dividends = []
    return resolved, dividends, ""


def sync_finnomena_dividend_history(args: argparse.Namespace) -> dict[str, int]:
    init_db(args.db_path)
    stats = {
        "funds": 0,
        "fetched_rows": 0,
        "inserted_rows": 0,
        "updated_rows": 0,
        "skipped_rows": 0,
        "error_funds": 0,
    }
    items = selected_finnomena_funds(args.db_path, args.limit)
    with connect(args.db_path) as conn:
        sync_run_id = create_sync_run(conn, "finnomena_dividend_history", f"funds={len(items)}")
        conn.commit()
        with ThreadPoolExecutor(max_workers=max(1, args.workers)) as executor:
            future_map = {
                executor.submit(fetch_one_finnomena_dividend, resolved): resolved
                for resolved in items
            }
            for future in as_completed(future_map):
                resolved = future_map[future]
                try:
                    _resolved, dividends, _error = future.result()
                    inserted, updated, skipped = upsert_finnomena_dividend_rows(conn, resolved, dividends, sync_run_id)
                    status = "success" if dividends else "no_data"
                    error_message = ""
                    stats["funds"] += 1
                    stats["fetched_rows"] += len(dividends)
                    stats["inserted_rows"] += inserted
                    stats["updated_rows"] += updated
                    stats["skipped_rows"] += skipped
                except Exception as exc:
                    dividends = []
                    inserted = updated = skipped = 0
                    status = "error"
                    error_message = str(exc)
                    stats["error_funds"] += 1

                conn.execute(
                    """
                    insert into finnomena_sync_items (
                        sync_run_id, proj_id, fund_code, finnomena_fund_id, status,
                        fetched_rows, inserted_rows, updated_rows, skipped_rows, error_message
                    ) values (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                    """,
                    (
                        sync_run_id,
                        normalized_text(resolved.get("proj_id")),
                        normalized_text(resolved.get("fund_code")),
                        normalized_text(resolved.get("finnomena_fund_id")),
                        status,
                        len(dividends),
                        inserted,
                        updated,
                        skipped,
                        error_message,
                    ),
                )
                conn.commit()

        final_status = "success" if stats["error_funds"] == 0 else "partial"
        finish_sync_run(conn, sync_run_id, final_status, json.dumps(stats, ensure_ascii=False, sort_keys=True))
        conn.commit()
    return stats


def backup_db(db_path: Path, backup_dir: Path) -> Path:
    backup_dir.mkdir(parents=True, exist_ok=True)
    stamp = datetime.now().strftime("%Y%m%d-%H%M%S")
    out_path = backup_dir / f"{db_path.stem}-{stamp}{db_path.suffix}"
    shutil.copy2(db_path, out_path)
    return out_path


def setup_system_folder(system_root: Path, db_path: Path) -> dict[str, str]:
    folders = {
        "root": system_root,
        "database": system_root / "database",
        "exports": system_root / "exports",
        "raw_sec": system_root / "raw" / "sec",
        "raw_finnomena": system_root / "raw" / "finnomena",
        "logs": system_root / "logs",
        "docs": system_root / "docs",
    }
    for folder in folders.values():
        folder.mkdir(parents=True, exist_ok=True)

    db_pointer = folders["database"] / "README.md"
    db_pointer.write_text(
        "\n".join(
            [
                "# Dividend Database",
                "",
                "SQLite working database:",
                f"`{db_path}`",
                "",
                "This folder keeps database pointers, exports, raw API snapshots, logs, and docs for the dividend-history system.",
            ]
        )
        + "\n",
        encoding="utf-8",
    )

    data_dictionary = folders["docs"] / "data_dictionary.md"
    if not data_dictionary.exists():
        data_dictionary.write_text(
            "\n".join(
                [
                    "# Dividend History Data Dictionary",
                    "",
                    "## Main Tables",
                    "",
                    "- `tracked_fund_classes`: fund codes/classes to track from AVP Dividend/Redemption universe.",
                    "- `tracked_funds`: distinct SEC `proj_id` values used for SEC API sync.",
                    "- `fund_dividend_history`: normalized SEC dividend-history rows.",
                    "- `dividend_sync_items`: per-fund sync status for each run.",
                    "- `sync_runs`: top-level sync logs.",
                    "",
                    "## Dividend History Fields",
                    "",
                    "- `proj_id`: SEC project id.",
                    "- `fund_code`: local/AVP fund code when it can be mapped.",
                    "- `unique_id`: SEC class unique id.",
                    "- `class_abbr_name`: SEC class abbreviation from dividend-history.",
                    "- `book_close_date`: book close / XD date.",
                    "- `dividend_date`: dividend payment date.",
                    "- `dividend_value`: dividend amount per unit.",
                    "- `source`: current export source. SEC rows use `SEC`; Finnomena rows will be added in the next phase.",
                ]
            )
            + "\n",
            encoding="utf-8",
        )

    return {name: str(path) for name, path in folders.items()}


def row_dict(row: sqlite3.Row) -> dict[str, Any]:
    return {key: row[key] for key in row.keys()}


def export_dividend_database_json(db_path: Path, output_path: Path) -> dict[str, Any]:
    init_db(db_path)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    with connect(db_path) as conn:
        fund_summary = [
            row_dict(row)
            for row in conn.execute(
                """
                select
                    tfc.fund_code,
                    tfc.proj_id,
                    tfc.fund_name,
                    tfc.asset_house,
                    tfc.tracking_type,
                    tfc.avp_fund_type,
                    tfc.sec_proj_name_th,
                    tfc.sec_proj_name_en,
                    tfc.sec_fund_class_name,
                    tfc.sec_unique_id,
                    count(fdh.id) as dividend_rows,
                    min(substr(fdh.dividend_date, 1, 10)) as first_dividend_date,
                    max(substr(fdh.dividend_date, 1, 10)) as latest_dividend_date
                from tracked_fund_classes tfc
                left join fund_dividend_history fdh
                  on fdh.proj_id = tfc.proj_id
                 and (
                    fdh.class_abbr_name = tfc.fund_code
                    or tfc.sec_fund_class_name = 'main'
                 )
                where tfc.is_active = 1
                group by tfc.fund_code
                order by tfc.fund_code
                """
            )
        ]

        dividend_history = [
            row_dict(row)
            for row in conn.execute(
                """
                with proj_lookup as (
                    select proj_id, min(fund_code) as fallback_fund_code, min(fund_name) as fallback_fund_name
                    from tracked_fund_classes
                    where is_active = 1
                    group by proj_id
                )
                select
                    fdh.proj_id,
                    coalesce(tfc.fund_code, proj_lookup.fallback_fund_code, fdh.class_abbr_name) as fund_code,
                    coalesce(tfc.fund_name, proj_lookup.fallback_fund_name, '') as fund_name,
                    coalesce(tfc.asset_house, '') as asset_house,
                    fdh.unique_id,
                    fdh.class_abbr_name,
                    substr(fdh.book_close_date, 1, 10) as book_close_date,
                    substr(fdh.dividend_date, 1, 10) as dividend_date,
                    fdh.dividend_value,
                    fdh.last_upd_date,
                    'SEC' as source,
                    fdh.first_seen_at,
                    fdh.last_seen_at
                from fund_dividend_history fdh
                left join tracked_fund_classes tfc
                  on tfc.proj_id = fdh.proj_id
                 and tfc.fund_code = fdh.class_abbr_name
                left join proj_lookup
                  on proj_lookup.proj_id = fdh.proj_id
                order by fdh.proj_id, fdh.dividend_date desc, fdh.book_close_date desc
                """
            )
        ]

        finnomena_history = [
            row_dict(row)
            for row in conn.execute(
                """
                select
                    proj_id,
                    fund_code,
                    fund_name,
                    asset_house,
                    '' as unique_id,
                    fund_code as class_abbr_name,
                    xd_date as book_close_date,
                    pay_date as dividend_date,
                    dividend_value,
                    '' as last_upd_date,
                    'Finnomena' as source,
                    finnomena_fund_id,
                    first_seen_at,
                    last_seen_at
                from finnomena_dividend_history
                order by proj_id, pay_date desc, xd_date desc
                """
            )
        ]
        dividend_history.extend(finnomena_history)
        dividend_history.sort(
            key=lambda row: (
                normalized_text(row.get("proj_id")),
                normalized_text(row.get("dividend_date")),
                normalized_text(row.get("book_close_date")),
                normalized_text(row.get("source")),
            ),
            reverse=True,
        )

        sync_runs = [
            row_dict(row)
            for row in conn.execute(
                """
                select id, run_type, status, started_at, finished_at, notes
                from sync_runs
                order by id desc
                limit 20
                """
            )
        ]

    payload = {
        "generated_at": utc_now(),
        "source_database": str(db_path),
        "counts": {
            "funds": len(fund_summary),
            "sec_dividend_history_rows": len(dividend_history) - len(finnomena_history),
            "finnomena_dividend_history_rows": len(finnomena_history),
            "dividend_history_rows": len(dividend_history),
            "sync_runs_included": len(sync_runs),
        },
        "funds": fund_summary,
        "dividend_history": dividend_history,
        "recent_sync_runs": sync_runs,
    }
    output_path.write_text(json.dumps(payload, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    return {"output_path": str(output_path), **payload["counts"]}


def upload_dividend_database_json(db_path: Path, output_path: Path, drive_folder_id: str, file_name: str) -> dict[str, Any]:
    load_dotenv()
    export_result = export_dividend_database_json(db_path, output_path)
    payload = json.loads(output_path.read_text(encoding="utf-8"))
    from drive_json_store import upload_json_payload

    file_id = upload_json_payload(drive_folder_id, file_name, payload)
    return {
        **export_result,
        "drive_folder_id": drive_folder_id,
        "drive_file_id": file_id,
        "drive_file_name": file_name,
        "drive_file_url": f"https://drive.google.com/file/d/{file_id}/view",
    }


def print_summary(db_path: Path) -> None:
    with connect(db_path) as conn:
        tables = [
            "tracked_fund_classes",
            "tracked_funds",
            "fund_dividend_history",
            "finnomena_dividend_history",
            "dividend_sync_items",
            "finnomena_sync_items",
            "sync_runs",
        ]
        for table in tables:
            count = conn.execute(f"select count(*) as c from {table}").fetchone()["c"]
            print(f"{table:28} {count:>8}")
        distinct_proj_ids = conn.execute(
            "select count(distinct proj_id) as c from tracked_fund_classes where is_active = 1"
        ).fetchone()["c"]
        print(f"{'distinct sync proj_ids':28} {distinct_proj_ids:>8}")
        rows = conn.execute(
            """
            select min(tfc.fund_code) as fund_code, tfc.proj_id, count(fdh.id) as dividend_rows,
                   max(substr(fdh.dividend_date, 1, 10)) as latest_dividend_date
            from tracked_fund_classes tfc
            left join fund_dividend_history fdh on fdh.proj_id = tfc.proj_id
            group by tfc.proj_id
            having dividend_rows > 0
            order by latest_dividend_date desc, fund_code
            limit 10
            """
        ).fetchall()
        if rows:
            print("\nLatest funds with dividend rows:")
            for row in rows:
                print(
                    f"{row['fund_code']:18} {row['proj_id']:12} "
                    f"rows={row['dividend_rows']:>3} latest={row['latest_dividend_date']}"
                )


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Store SEC dividend history in SQLite.")
    parser.add_argument("--db-path", type=Path, default=DEFAULT_DB_PATH)
    subparsers = parser.add_subparsers(dest="command", required=True)

    setup_parser = subparsers.add_parser("setup-system", help="Create the dividend database folder structure.")
    setup_parser.add_argument("--system-root", type=Path, default=DEFAULT_SYSTEM_ROOT)

    subparsers.add_parser("init-db", help="Create or upgrade dividend-history tables.")

    import_parser = subparsers.add_parser("import-tracked-funds", help="Import dividend/redemption funds into tracked_funds.")
    import_parser.add_argument("--sec-master", type=Path, default=DEFAULT_SEC_MASTER_PATH)
    import_parser.add_argument("--avp-performance", type=Path, default=DEFAULT_AVP_PERFORMANCE_PATH)

    sync_parser = subparsers.add_parser("sync", help="Fetch SEC dividend-history rows for tracked funds.")
    sync_parser.add_argument("--api-key", default="")
    sync_parser.add_argument("--proj-id", default="", help="Limit sync to one SEC proj_id.")
    sync_parser.add_argument("--limit", type=int, default=0, help="Limit number of tracked funds to sync.")
    sync_parser.add_argument("--page-size", type=int, default=100)
    sync_parser.add_argument("--sleep-seconds", type=float, default=0.1)
    sync_parser.add_argument("--raw-dir", type=Path, default=DEFAULT_RAW_DIR)

    finnomena_sync_parser = subparsers.add_parser("sync-finnomena", help="Fetch Finnomena dividend-history rows for mapped funds.")
    finnomena_sync_parser.add_argument("--limit", type=int, default=0, help="Limit number of mapped fund codes to sync.")
    finnomena_sync_parser.add_argument("--workers", type=int, default=12)

    backup_parser = subparsers.add_parser("backup-db", help="Copy the SQLite database to Data/backups.")
    backup_parser.add_argument("--backup-dir", type=Path, default=DEFAULT_BACKUP_DIR)

    export_json_parser = subparsers.add_parser("export-json", help="Export the dividend database to JSON.")
    export_json_parser.add_argument(
        "--output",
        type=Path,
        default=DEFAULT_EXPORT_DIR / "dividend_history_database.json",
    )

    upload_json_parser = subparsers.add_parser("upload-json", help="Export JSON and upload it to the Dividend History Drive folder.")
    upload_json_parser.add_argument(
        "--output",
        type=Path,
        default=DEFAULT_EXPORT_DIR / "dividend_history_database.json",
    )
    upload_json_parser.add_argument(
        "--drive-folder-id",
        default=os.environ.get("DIVIDEND_HISTORY_DRIVE_FOLDER_ID", "").strip() or DEFAULT_DRIVE_FOLDER_ID,
    )
    upload_json_parser.add_argument("--file-name", default="dividend_history_database.json")

    subparsers.add_parser("summary", help="Print dividend-history table counts.")
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    if args.command == "init-db":
        init_db(args.db_path)
        print(f"Initialized dividend-history tables in {args.db_path}")
        return 0
    if args.command == "setup-system":
        init_db(args.db_path)
        folders = setup_system_folder(args.system_root, args.db_path)
        print(json.dumps(folders, ensure_ascii=False, indent=2))
        return 0
    if args.command == "import-tracked-funds":
        stats = import_tracked_funds(args.db_path, args.sec_master, args.avp_performance)
        print(json.dumps(stats, ensure_ascii=False, indent=2))
        return 0
    if args.command == "sync":
        stats = sync_dividend_history(args)
        print(json.dumps(stats, ensure_ascii=False, indent=2))
        return 0
    if args.command == "sync-finnomena":
        stats = sync_finnomena_dividend_history(args)
        print(json.dumps(stats, ensure_ascii=False, indent=2))
        return 0
    if args.command == "backup-db":
        out_path = backup_db(args.db_path, args.backup_dir)
        print(f"Backed up {args.db_path} to {out_path}")
        return 0
    if args.command == "export-json":
        result = export_dividend_database_json(args.db_path, args.output)
        print(json.dumps(result, ensure_ascii=False, indent=2))
        return 0
    if args.command == "upload-json":
        result = upload_dividend_database_json(args.db_path, args.output, args.drive_folder_id, args.file_name)
        print(json.dumps(result, ensure_ascii=False, indent=2))
        return 0
    if args.command == "summary":
        init_db(args.db_path)
        print_summary(args.db_path)
        return 0
    raise AssertionError(f"Unhandled command: {args.command}")


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except Exception as exc:
        print(f"ERROR: {exc}", file=sys.stderr)
        raise SystemExit(1)
