#!/usr/bin/env python3
"""Download one FT historical artifact from the Apps Script Drive web app."""

from __future__ import annotations

import argparse
import base64
import json
import urllib.error
import urllib.parse
import urllib.request
from pathlib import Path
from typing import Any


def get_json(url: str, params: dict[str, Any]) -> dict[str, Any]:
    full_url = f"{url}?{urllib.parse.urlencode(params)}"
    request = urllib.request.Request(full_url, method="GET")
    try:
        with urllib.request.urlopen(request, timeout=180) as response:
            text = response.read().decode("utf-8")
    except urllib.error.HTTPError as exc:
        text = exc.read().decode("utf-8", errors="replace")
        raise RuntimeError(f"Apps Script download failed ({exc.code}): {text}") from exc
    return json.loads(text)


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--url", required=True)
    parser.add_argument("--key", default="")
    parser.add_argument("--folder-id", required=True)
    parser.add_argument("--relative-path", required=True)
    parser.add_argument("--output", type=Path, required=True)
    parser.add_argument("--optional", action="store_true", help="Exit successfully if the remote file is not found.")
    args = parser.parse_args()

    result = get_json(args.url, {
        "action": "downloadFile",
        "key": args.key,
        "folderId": args.folder_id,
        "relativePath": args.relative_path,
    })
    if result.get("notFound"):
        if args.optional:
            print(json.dumps(result, ensure_ascii=False, indent=2))
            return 0
        raise FileNotFoundError(result.get("error") or args.relative_path)
    if result.get("ok") is False:
        raise RuntimeError(result.get("error") or result.get("message") or "Apps Script download failed")

    content = base64.b64decode(result.get("contentBase64") or "")
    args.output.parent.mkdir(parents=True, exist_ok=True)
    args.output.write_bytes(content)
    print(json.dumps({
        "ok": True,
        "output": str(args.output),
        "relativePath": args.relative_path,
        "bytes": len(content),
        "drive": result.get("drive") or {},
    }, ensure_ascii=False, indent=2))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
