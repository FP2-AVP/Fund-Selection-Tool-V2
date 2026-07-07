#!/usr/bin/env python3
"""Upload FT historical SQLite/CSV/raw artifacts to the Apps Script Drive web app."""

from __future__ import annotations

import argparse
import base64
import json
import mimetypes
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
        with urllib.request.urlopen(request, timeout=180) as response:
            text = response.read().decode("utf-8")
    except urllib.error.HTTPError as exc:
        text = exc.read().decode("utf-8", errors="replace")
        raise RuntimeError(f"Apps Script file upload failed ({exc.code}): {text}") from exc
    result = json.loads(text)
    if result.get("ok") is False:
        raise RuntimeError(result.get("error") or result.get("message") or "Apps Script file upload failed")
    return result


def iter_files(root: Path) -> list[Path]:
    return sorted(path for path in root.rglob("*") if path.is_file())


def guess_mime_type(path: Path) -> str:
    if path.suffix == ".sqlite":
        return "application/vnd.sqlite3"
    guessed, _ = mimetypes.guess_type(path.name)
    return guessed or "application/octet-stream"


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--url", required=True)
    parser.add_argument("--key", default="")
    parser.add_argument("--root", type=Path, default=Path("Data/ft_historical_prices"))
    parser.add_argument("--folder-id", required=True)
    parser.add_argument("--prefix", default="", help="Optional Drive relative path prefix.")
    args = parser.parse_args()

    if not args.root.exists():
      raise FileNotFoundError(f"Artifact root not found: {args.root}")

    uploaded: list[dict[str, Any]] = []
    prefix = "/".join(part for part in args.prefix.replace("\\", "/").split("/") if part)
    for path in iter_files(args.root):
        relative = path.relative_to(args.root).as_posix()
        relative_path = f"{prefix}/{relative}" if prefix else relative
        payload = {
            "action": "uploadFile",
            "key": args.key,
            "folderId": args.folder_id,
            "relativePath": relative_path,
            "fileName": path.name,
            "mimeType": guess_mime_type(path),
            "contentBase64": base64.b64encode(path.read_bytes()).decode("ascii"),
        }
        result = post_json(args.url, payload)
        drive = result.get("drive") or {}
        uploaded.append({
            "relativePath": relative_path,
            "fileId": drive.get("fileId", ""),
            "size": path.stat().st_size,
        })
        print(f"uploaded {relative_path} ({path.stat().st_size} bytes)")

    print(json.dumps({"ok": True, "uploaded": uploaded, "count": len(uploaded)}, ensure_ascii=False, indent=2))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
