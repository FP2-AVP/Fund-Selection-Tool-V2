#!/usr/bin/env python3
"""Export the FT SQLite database as sharded R2 objects and optionally upload them."""

from __future__ import annotations

import argparse
import gzip
import hashlib
import json
import os
import shutil
import sqlite3
from collections import defaultdict
from datetime import datetime, timezone
from pathlib import Path
from typing import Any


PROJECT_ROOT = Path(__file__).resolve().parents[1]
DEFAULT_DB_PATH = PROJECT_ROOT / "Data" / "ft_historical_prices" / "ft_historical_prices.sqlite"
DEFAULT_BUILD_ROOT = PROJECT_ROOT / "Data" / "ft_historical_prices" / "r2_export"
DEFAULT_R2_PREFIX = "Data For FT.com"


def utc_now() -> str:
    return datetime.now(timezone.utc).replace(microsecond=0).isoformat().replace("+00:00", "Z")


def symbol_slug(symbol: str) -> str:
    import re

    return re.sub(r"[^A-Za-z0-9]+", "_", symbol).strip("_").upper()


def row_dict(row: sqlite3.Row) -> dict[str, Any]:
    return {key: row[key] for key in row.keys()}


def write_json(path: Path, payload: Any) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(payload, ensure_ascii=False, separators=(",", ":")) + "\n", encoding="utf-8")


def write_gzip_json(path: Path, payload: Any) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with gzip.open(path, "wt", encoding="utf-8", compresslevel=6) as handle:
        json.dump(payload, handle, ensure_ascii=False, separators=(",", ":"))
        handle.write("\n")


def rows_for_symbol(conn: sqlite3.Connection, table: str, symbol: str) -> list[dict[str, Any]]:
    order_by = {
        "ft_profile_investment": "section, field",
        "ft_performance_measures": "period, series",
        "ft_risk_measures": "period, metric",
        "ft_top_holdings": "rank",
    }[table]
    return [row_dict(row) for row in conn.execute(f"select * from {table} where upper(symbol)=upper(?) order by {order_by}", (symbol,))]


def qualitative_as_of(rows: list[dict[str, Any]], generated_at: str) -> str:
    dates = [str(row.get("as_of_date") or "") for row in rows if str(row.get("as_of_date") or "")]
    return max(dates) if dates else generated_at[:10]


def export_objects(db_path: Path, build_root: Path) -> dict[str, Any]:
    if not db_path.is_file():
        raise FileNotFoundError(f"FT SQLite database not found: {db_path}")
    if build_root.exists():
        shutil.rmtree(build_root)
    build_root.mkdir(parents=True, exist_ok=True)
    generated_at = utc_now()

    with sqlite3.connect(db_path) as conn:
        conn.row_factory = sqlite3.Row
        symbols = [
            str(row[0])
            for row in conn.execute(
                """
                select symbol from ft_historical_prices
                union select symbol from ft_profile_investment
                union select symbol from ft_performance_measures
                union select symbol from ft_risk_measures
                union select symbol from ft_top_holdings
                order by symbol collate nocase
                """
            )
        ]
        index_items: list[dict[str, Any]] = []
        for symbol in symbols:
            slug = symbol_slug(symbol)
            symbol_root = build_root / "symbols" / slug
            price_rows = [
                row_dict(row)
                for row in conn.execute(
                    """
                    select symbol, ft_issue_id, price_date as date, open, high, low, close, volume,
                           source, fetched_at as fetchedAt
                    from ft_historical_prices
                    where upper(symbol)=upper(?) order by price_date
                    """,
                    (symbol,),
                )
            ]
            prices_by_year: dict[str, list[dict[str, Any]]] = defaultdict(list)
            for row in price_rows:
                prices_by_year[str(row.get("date") or "")[:4]].append(row)
            price_files = []
            for year, year_rows in sorted(prices_by_year.items()):
                if not year:
                    continue
                relative_path = f"symbols/{slug}/prices/{year}.json.gz"
                write_gzip_json(build_root / relative_path, {
                    "generatedAt": generated_at,
                    "symbol": symbol,
                    "year": year,
                    "rows": year_rows,
                })
                price_files.append({"year": year, "path": relative_path, "rows": len(year_rows)})

            profile = rows_for_symbol(conn, "ft_profile_investment", symbol)
            performance = rows_for_symbol(conn, "ft_performance_measures", symbol)
            risk = rows_for_symbol(conn, "ft_risk_measures", symbol)
            holdings = rows_for_symbol(conn, "ft_top_holdings", symbol)
            qualitative_rows = profile + performance + risk + holdings
            as_of_date = qualitative_as_of(qualitative_rows, generated_at)
            qualitative = {
                "generatedAt": generated_at,
                "symbol": symbol,
                "symbolSlug": slug,
                "asOfDate": as_of_date,
                "counts": {
                    "profile": len(profile),
                    "performance": len(performance),
                    "risk": len(risk),
                    "holdings": len(holdings),
                },
                "profile": profile,
                "performance": performance,
                "risk": risk,
                "holdings": holdings,
            }
            if qualitative_rows:
                write_json(symbol_root / "qualitative" / "latest.json", qualitative)
                write_json(symbol_root / "qualitative" / "snapshots" / f"{as_of_date}.json", qualitative)

            display_name = next(
                (str(row.get("value") or "") for row in profile if row.get("section") == "metadata" and row.get("field") == "FT display name"),
                "",
            )
            isin = next((str(row.get("value") or "") for row in profile if row.get("field") == "ISIN"), "")
            issue_id = next((str(row.get("ft_issue_id") or "") for row in price_rows + qualitative_rows if row.get("ft_issue_id")), "")
            metadata = {
                "generatedAt": generated_at,
                "symbol": symbol,
                "symbolSlug": slug,
                "ftIssueId": issue_id,
                "displayName": display_name,
                "isin": isin,
                "priceStart": price_rows[0]["date"] if price_rows else "",
                "priceEnd": price_rows[-1]["date"] if price_rows else "",
                "priceRows": len(price_rows),
                "priceFiles": price_files,
                "qualitativeAsOf": as_of_date if qualitative_rows else "",
                "qualitativeCounts": qualitative["counts"],
            }
            write_json(symbol_root / "metadata.json", metadata)
            index_items.append(metadata)

    index_payload = {
        "schemaVersion": 1,
        "generatedAt": generated_at,
        "source": "Cloudflare R2: Data For FT.com",
        "counts": {
            "symbols": len(index_items),
            "priceRows": sum(int(item["priceRows"]) for item in index_items),
        },
        "symbols": index_items,
    }
    write_json(build_root / "index.json", index_payload)
    return index_payload


def upload_objects(build_root: Path, prefix: str) -> dict[str, Any]:
    account_id = os.environ["CLOUDFLARE_ACCOUNT_ID"].strip()
    bucket = os.environ["CLOUDFLARE_R2_BUCKET"].strip()
    import boto3

    client = boto3.client(
        "s3",
        endpoint_url=f"https://{account_id}.r2.cloudflarestorage.com",
        region_name="auto",
        aws_access_key_id=os.environ["CLOUDFLARE_R2_ACCESS_KEY_ID"].strip(),
        aws_secret_access_key=os.environ["CLOUDFLARE_R2_SECRET_ACCESS_KEY"].strip(),
    )
    uploaded = []
    for path in sorted(item for item in build_root.rglob("*") if item.is_file()):
        relative = path.relative_to(build_root).as_posix()
        key = f"{prefix.strip('/')}/{relative}"
        body = path.read_bytes()
        args = {
            "Bucket": bucket,
            "Key": key,
            "Body": body,
            "ContentType": "application/json; charset=utf-8",
            "Metadata": {"sha256": hashlib.sha256(body).hexdigest()},
        }
        if path.suffix == ".gz":
            args["ContentEncoding"] = "gzip"
        if relative == "index.json":
            args["CacheControl"] = "no-cache"
        client.put_object(**args)
        uploaded.append({"key": key, "bytes": len(body)})
        print(f"Uploaded s3://{bucket}/{key} ({len(body):,} bytes)")
    return {"bucket": bucket, "prefix": prefix, "uploaded": len(uploaded), "objects": uploaded}


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--db-path", type=Path, default=DEFAULT_DB_PATH)
    parser.add_argument("--build-root", type=Path, default=DEFAULT_BUILD_ROOT)
    parser.add_argument("--prefix", default=DEFAULT_R2_PREFIX)
    parser.add_argument("--upload", action="store_true")
    args = parser.parse_args()
    result = export_objects(args.db_path, args.build_root)
    summary: dict[str, Any] = {
        "buildRoot": str(args.build_root),
        "symbols": result["counts"]["symbols"],
        "priceRows": result["counts"]["priceRows"],
    }
    if args.upload:
        summary["upload"] = upload_objects(args.build_root, args.prefix)
    print(json.dumps(summary, ensure_ascii=False, indent=2))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
