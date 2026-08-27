#!/usr/bin/env python3
"""Upload one quarter's base JSON files to R2 and update manifest.json."""

from __future__ import annotations

import argparse
import hashlib
import json
import os
import re
from datetime import datetime, timezone
from pathlib import Path

REQUIRED_FILES = {
    "secApi": "Data For SEC API.json",
    "fundKeyPerformance": "Fund Key Performance AVP.json",
    "thaiQuality": "AVP Thai Fund for Quality.json",
    "masterFund": "AVP Master Fund ID.json",
}


def args_parser() -> argparse.Namespace:
    parser = argparse.ArgumentParser()
    parser.add_argument("--quarter", required=True)
    parser.add_argument("--source-dir", required=True)
    return parser.parse_args()


def main() -> None:
    args = args_parser()
    quarter = args.quarter.strip().upper()
    if not re.fullmatch(r"\d{4}-Q[1-4]", quarter):
        raise ValueError("quarter must look like 2026-Q3")
    year = quarter.split("-", 1)[0]
    base_dir = Path(args.source_dir) / year / quarter / "base"
    missing = [name for name in REQUIRED_FILES.values() if not (base_dir / name).is_file()]
    if missing:
        raise RuntimeError(f"Missing required base JSON: {', '.join(missing)}")

    account_id = os.environ["CLOUDFLARE_ACCOUNT_ID"].strip()
    bucket = os.environ["CLOUDFLARE_R2_BUCKET"].strip()
    endpoint = f"https://{account_id}.r2.cloudflarestorage.com"

    import boto3
    from botocore.exceptions import ClientError

    client = boto3.client(
        "s3",
        endpoint_url=endpoint,
        region_name="auto",
        aws_access_key_id=os.environ["CLOUDFLARE_R2_ACCESS_KEY_ID"].strip(),
        aws_secret_access_key=os.environ["CLOUDFLARE_R2_SECRET_ACCESS_KEY"].strip(),
    )

    now = datetime.now(timezone.utc).replace(microsecond=0).isoformat().replace("+00:00", "Z")
    file_meta: dict[str, dict[str, object]] = {}
    for dataset_key, name in REQUIRED_FILES.items():
        path = base_dir / name
        body = path.read_bytes()
        object_key = f"{year}/{quarter}/base/{name}"
        client.put_object(Bucket=bucket, Key=object_key, Body=body, ContentType="application/json")
        file_meta[dataset_key] = {
            "name": name,
            "path": object_key,
            "bytes": len(body),
            "sha256": hashlib.sha256(body).hexdigest(),
        }
        print(f"Uploaded s3://{bucket}/{object_key} ({len(body):,} bytes)")

    try:
        response = client.get_object(Bucket=bucket, Key="manifest.json")
        manifest = json.loads(response["Body"].read().decode("utf-8"))
    except ClientError as exc:
        if exc.response.get("Error", {}).get("Code") not in {"NoSuchKey", "404"}:
            raise
        manifest = {}

    quarters = manifest.get("quarters") if isinstance(manifest.get("quarters"), dict) else {}
    quarters[quarter] = {"year": year, "basePath": f"{year}/{quarter}/base/", "files": file_meta, "updatedAt": now}
    ready = sorted(quarters, key=lambda value: (int(value[:4]), int(value[-1])), reverse=True)
    manifest = {"version": 1, "updatedAt": now, "readyQuarters": ready, "quarters": quarters}
    payload = (json.dumps(manifest, ensure_ascii=False, separators=(",", ":")) + "\n").encode("utf-8")
    client.put_object(Bucket=bucket, Key="manifest.json", Body=payload, ContentType="application/json", CacheControl="no-cache")
    print(f"Updated manifest.json: {', '.join(ready)}")


if __name__ == "__main__":
    main()
