#!/usr/bin/env python3
"""Upload dividend_history_database.json to the Apps Script Drive web app."""

from __future__ import annotations

import argparse
import json
import urllib.error
import urllib.request
from pathlib import Path
from typing import Any


def post_json(url: str, payload: dict[str, Any]) -> dict[str, Any]:
    data = json.dumps(payload, ensure_ascii=False).encode("utf-8")
    request = urllib.request.Request(
        url,
        data=data,
        headers={"Content-Type": "text/plain;charset=utf-8"},
        method="POST",
    )
    try:
        with urllib.request.urlopen(request, timeout=120) as response:
            text = response.read().decode("utf-8")
    except urllib.error.HTTPError as exc:
        text = exc.read().decode("utf-8", errors="replace")
        raise RuntimeError(f"Apps Script upload failed ({exc.code}): {text}") from exc
    result = json.loads(text)
    if result.get("ok") is False:
        raise RuntimeError(result.get("error") or result.get("message") or "Apps Script upload failed")
    return result


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--url", required=True)
    parser.add_argument("--key", default="")
    parser.add_argument("--input", type=Path, required=True)
    parser.add_argument("--folder-id", required=True)
    parser.add_argument("--file-name", default="dividend_history_database.json")
    args = parser.parse_args()

    database = json.loads(args.input.read_text(encoding="utf-8"))
    result = post_json(args.url, {
        "action": "uploadDatabase",
        "key": args.key,
        "folderId": args.folder_id,
        "fileName": args.file_name,
        "database": database,
    })
    print(json.dumps(result, ensure_ascii=False, indent=2))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
