#!/usr/bin/env python3
"""Restore the FT working SQLite database from R2, rebuilding it from shards when needed."""

from __future__ import annotations

import argparse
import gzip
import json
import os
import sqlite3
from pathlib import Path
from typing import Any

from ft_historical_prices_store import init_db


DEFAULT_OUTPUT = Path("Data/ft_historical_prices/ft_historical_prices.sqlite")
DEFAULT_PREFIX = "Data For FT.com"


def r2_client():
    import boto3

    account_id = os.environ["CLOUDFLARE_ACCOUNT_ID"].strip()
    return boto3.client(
        "s3",
        endpoint_url=f"https://{account_id}.r2.cloudflarestorage.com",
        region_name="auto",
        aws_access_key_id=os.environ["CLOUDFLARE_R2_ACCESS_KEY_ID"].strip(),
        aws_secret_access_key=os.environ["CLOUDFLARE_R2_SECRET_ACCESS_KEY"].strip(),
    )


def list_keys(client, bucket: str, prefix: str) -> list[str]:
    keys: list[str] = []
    token = None
    while True:
        args = {"Bucket": bucket, "Prefix": prefix}
        if token:
            args["ContinuationToken"] = token
        page = client.list_objects_v2(**args)
        keys.extend(item["Key"] for item in page.get("Contents", []))
        if not page.get("IsTruncated"):
            return keys
        token = page.get("NextContinuationToken")


def get_bytes(client, bucket: str, key: str) -> bytes:
    return client.get_object(Bucket=bucket, Key=key)["Body"].read()


def get_json(client, bucket: str, key: str) -> dict[str, Any]:
    body = get_bytes(client, bucket, key)
    if key.endswith(".gz"):
        body = gzip.decompress(body)
    return json.loads(body.decode("utf-8"))


def symbol_count(path: Path) -> int:
    if not path.is_file() or path.stat().st_size == 0:
        return 0
    with sqlite3.connect(path) as conn:
        tables = {row[0] for row in conn.execute("select name from sqlite_master where type='table'")}
        available = [
            table for table in (
                "ft_historical_prices", "ft_profile_investment", "ft_performance_measures",
                "ft_risk_measures", "ft_top_holdings",
            ) if table in tables
        ]
        if not available:
            return 0
        unions = " union ".join(f"select symbol from {table}" for table in available)
        return int(conn.execute(f"select count(*) from ({unions})").fetchone()[0])


def insert_dicts(conn: sqlite3.Connection, table: str, rows: list[dict[str, Any]], columns: list[str]) -> int:
    clean_rows = []
    for row in rows:
        if not isinstance(row, dict):
            continue
        clean_rows.append(tuple(row.get(column) for column in columns))
    if not clean_rows:
        return 0
    placeholders = ",".join("?" for _ in columns)
    conn.executemany(
        f"insert or replace into {table} ({','.join(columns)}) values ({placeholders})",
        clean_rows,
    )
    return len(clean_rows)


def rebuild_from_shards(client, bucket: str, prefix: str, output: Path) -> dict[str, int]:
    symbol_prefix = f"{prefix.rstrip('/')}/symbols/"
    keys = list_keys(client, bucket, symbol_prefix)
    price_keys = sorted(key for key in keys if "/prices/" in key and key.endswith(".json.gz"))
    qualitative_keys = sorted(key for key in keys if key.endswith("/qualitative/latest.json"))
    metadata_keys = sorted(key for key in keys if key.endswith("/metadata.json"))
    if not price_keys and not qualitative_keys and not metadata_keys:
        raise RuntimeError(f"No FT symbol objects found under {symbol_prefix}")

    output.parent.mkdir(parents=True, exist_ok=True)
    if output.exists():
        output.unlink()
    init_db(output)
    counts = {"prices": 0, "profile": 0, "performance": 0, "risk": 0, "holdings": 0}
    with sqlite3.connect(output) as conn:
        for key in price_keys:
            payload = get_json(client, bucket, key)
            rows = []
            for row in payload.get("rows", []):
                rows.append({
                    "symbol": row.get("symbol") or payload.get("symbol"),
                    "ft_issue_id": row.get("ft_issue_id") or row.get("ftIssueId") or "",
                    "price_date": row.get("date") or row.get("price_date"),
                    "open": row.get("open"), "high": row.get("high"), "low": row.get("low"),
                    "close": row.get("close"), "volume": row.get("volume"),
                    "source": row.get("source") or "FT Markets",
                    "fetched_at": row.get("fetchedAt") or row.get("fetched_at") or "",
                })
            counts["prices"] += insert_dicts(conn, "ft_historical_prices", rows, [
                "symbol", "ft_issue_id", "price_date", "open", "high", "low", "close", "volume", "source", "fetched_at",
            ])

        for key in qualitative_keys:
            payload = get_json(client, bucket, key)
            counts["profile"] += insert_dicts(conn, "ft_profile_investment", payload.get("profile", []), [
                "symbol", "ft_issue_id", "section", "field", "value", "source", "fetched_at",
            ])
            counts["performance"] += insert_dicts(conn, "ft_performance_measures", payload.get("performance", []), [
                "symbol", "ft_issue_id", "period", "series", "value", "as_of_date", "source", "fetched_at",
            ])
            counts["risk"] += insert_dicts(conn, "ft_risk_measures", payload.get("risk", []), [
                "symbol", "ft_issue_id", "period", "metric", "fund_value", "category_average",
                "benchmark_used", "as_of_date", "source", "fetched_at",
            ])
            counts["holdings"] += insert_dicts(conn, "ft_top_holdings", payload.get("holdings", []), [
                "symbol", "ft_issue_id", "rank", "holding_name", "holding_symbol", "one_year_change",
                "portfolio_weight", "long_allocation", "top10_portfolio_percent", "as_of_date", "source", "fetched_at",
            ])

        # Preserve symbols that only have metadata and no price/qualitative rows.
        known = {str(row[0]).upper() for row in conn.execute(
            "select symbol from ft_historical_prices union select symbol from ft_profile_investment"
        )}
        for key in metadata_keys:
            metadata = get_json(client, bucket, key)
            symbol = str(metadata.get("symbol") or "").strip()
            if not symbol or symbol.upper() in known:
                continue
            conn.execute(
                "insert or replace into ft_profile_investment values (?,?,?,?,?,?,?)",
                (symbol, metadata.get("ftIssueId") or "", "metadata", "FT display name",
                 metadata.get("displayName") or symbol, "Cloudflare R2 recovery", metadata.get("generatedAt") or ""),
            )
            counts["profile"] += 1
    counts["symbols"] = symbol_count(output)
    if counts["symbols"] < 1:
        raise RuntimeError("R2 shard recovery produced an empty FT database")
    return counts


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--output", type=Path, default=DEFAULT_OUTPUT)
    parser.add_argument("--prefix", default=DEFAULT_PREFIX)
    parser.add_argument("--force-rebuild", action="store_true")
    args = parser.parse_args()
    client = r2_client()
    bucket = os.environ["CLOUDFLARE_R2_BUCKET"].strip()
    database_key = f"{args.prefix.rstrip('/')}/database/ft_historical_prices.sqlite"
    args.output.parent.mkdir(parents=True, exist_ok=True)

    if not args.force_rebuild:
        try:
            args.output.write_bytes(get_bytes(client, bucket, database_key))
            count = symbol_count(args.output)
            if count:
                print(json.dumps({"source": database_key, "symbols": count}, indent=2))
                return 0
        except Exception as error:
            print(f"R2 working SQLite unavailable; rebuilding from shards: {error}")

    counts = rebuild_from_shards(client, bucket, args.prefix, args.output)
    print(json.dumps({"source": f"{args.prefix}/symbols/", **counts}, indent=2))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
