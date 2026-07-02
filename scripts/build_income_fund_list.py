from __future__ import annotations

import json
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
QUARTER = "2026-Q1"
SEC_FILE = ROOT / "Data" / f"Data For SEC API - {QUARTER}.json"
KEY_PERFORMANCE_FILE = ROOT / "Data" / f"Fund Key Performance AVP - {QUARTER}.json"
OUTPUT_FILE = ROOT / "Data" / f"Income Fund - {QUARTER}.json"


def normalize(value: object) -> str:
    return " ".join(str(value or "").replace("\xa0", " ").split()).strip()


def column_index(headers: list[object], candidates: list[str]) -> int:
    normalized_headers = [normalize(header).lower() for header in headers]
    for candidate in candidates:
        wanted = normalize(candidate).lower()
        if wanted in normalized_headers:
            return normalized_headers.index(wanted)
    return -1


def value_at(row: list[object], index: int) -> str:
    if index < 0 or index >= len(row):
        return ""
    return normalize(row[index])


def load_rows(path: Path) -> list[list[object]]:
    rows = json.loads(path.read_text(encoding="utf-8"))
    if not isinstance(rows, list) or not rows:
        raise ValueError(f"Invalid rows in {path}")
    return rows


def build_income_rows() -> list[dict[str, str]]:
    sec_rows = load_rows(SEC_FILE)
    key_rows = load_rows(KEY_PERFORMANCE_FILE)

    sec_headers = sec_rows[0]
    sec_idx = {
        "name_th": column_index(sec_headers, ["proj_name_th"]),
        "name_en": column_index(sec_headers, ["proj_name_en"]),
        "class_name": column_index(sec_headers, ["fund_class_name"]),
        "abbr": column_index(sec_headers, ["proj_abbr_name"]),
        "policy": column_index(sec_headers, ["dividend_policy"]),
    }

    key_headers = key_rows[0]
    key_idx = {
        "name": column_index(key_headers, ["Name", "Fund Name", "FundName"]),
        "code": column_index(key_headers, ["Fund Code", "Code"]),
        "policy": column_index(key_headers, ["Dividend", "Div"]),
    }

    items: list[dict[str, str]] = []
    for row in sec_rows[1:]:
        if value_at(row, sec_idx["policy"]).upper() != "Y":
            continue
        base_name = (
            value_at(row, sec_idx["name_th"])
            or value_at(row, sec_idx["name_en"])
            or value_at(row, sec_idx["abbr"])
            or value_at(row, sec_idx["class_name"])
        )
        if not base_name:
            continue
        class_name = value_at(row, sec_idx["class_name"])
        abbr = value_at(row, sec_idx["abbr"])
        class_suffix = (
            f" ({class_name})"
            if class_name and class_name.lower() != "main" and class_name != abbr
            else ""
        )
        items.append({
            "name": f"{base_name}{class_suffix}",
            "policy": "Dividend",
            "source": "Data For SEC API",
        })

    for row in key_rows[1:]:
        policy = value_at(row, key_idx["policy"])
        if policy not in {"Dividend", "Redemption"}:
            continue
        name = value_at(row, key_idx["name"]) or value_at(row, key_idx["code"])
        if not name:
            continue
        items.append({
            "name": name,
            "policy": policy,
            "source": "Fund Key Performance AVP",
        })

    seen: set[tuple[str, str]] = set()
    result: list[dict[str, str]] = []
    for item in sorted(items, key=lambda row: (row["name"], row["policy"])):
        key = (normalize(item["name"]).upper(), item["policy"])
        if key in seen:
            continue
        seen.add(key)
        result.append(item)
    return result


def main() -> None:
    items = build_income_rows()
    payload = {
        "quarter": QUARTER,
        "sourceFiles": [
            str(SEC_FILE.relative_to(ROOT)),
            str(KEY_PERFORMANCE_FILE.relative_to(ROOT)),
        ],
        "headers": ["ชื่อกองทุน", "นโยบาย"],
        "items": items,
    }
    OUTPUT_FILE.write_text(
        json.dumps(payload, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )
    print(f"Wrote {len(items):,} rows to {OUTPUT_FILE.relative_to(ROOT)}")


if __name__ == "__main__":
    main()
