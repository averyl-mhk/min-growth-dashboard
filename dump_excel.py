"""dump_excel.py — flatten data.json into a multi-sheet Excel for inspection.

Run from the repo root:
    python dump_excel.py                       # writes data.xlsx with every month
    python dump_excel.py --month "Feb '26"     # only one month
    python dump_excel.py --out review.xlsx     # custom output path
"""

import argparse
import json
import sys
from pathlib import Path

try:
    import pandas as pd
except ImportError:
    sys.exit("Missing dependency: pip install pandas openpyxl")


SCRIPT_DIR = Path(__file__).parent
DATA_FILE = SCRIPT_DIR / "data.json"


def flatten_monthly(monthly: dict) -> pd.DataFrame:
    rows = []
    for month, m in monthly.items():
        platforms = m.get("platforms", {})
        pos = m.get("pos") or {}
        row = {
            "month": month,
            "totalSpend": m.get("spend"),
            "totalAttRev": m.get("attRev"),
            "posTotal": pos.get("total"),
            "posAmazon": pos.get("Amazon"),
            "posShopify": pos.get("Shopify"),
            "posFlipkart": pos.get("Flipkart"),
        }
        for plat, vals in platforms.items():
            row[f"{plat} spend"] = vals.get("spend")
            row[f"{plat} attRev"] = vals.get("attRev")
            row[f"{plat} acos"] = vals.get("acos")
            row[f"{plat} pending"] = vals.get("pending")
        rows.append(row)
    return pd.DataFrame(rows)


def flatten_channel_detail(channel_detail: dict) -> pd.DataFrame:
    rows = []
    for month, buckets in channel_detail.items():
        for bucket_key in ("awareness", "coreSales", "retargeting"):
            for r in buckets.get(bucket_key, []):
                rows.append({"month": month, "bucket": bucket_key, **r})
            sub = buckets.get(f"{bucket_key}Subtotal")
            if sub:
                rows.append({"month": month, "bucket": bucket_key, "campaignGroup": "SUBTOTAL", **sub})
    return pd.DataFrame(rows)


def flatten_ranked(section: dict) -> pd.DataFrame:
    rows = []
    for month, entries in section.items():
        for r in entries:
            rows.append({"month": month, **r})
    return pd.DataFrame(rows)


def flatten_bucket_summary(bucket_summary: dict) -> pd.DataFrame:
    rows = []
    for month, buckets in bucket_summary.items():
        for bucket_key, b in buckets.items():
            row = {"month": month, "bucket": bucket_key, **{k: v for k, v in b.items() if k != "channels"}}
            row["channels"] = ", ".join(b.get("channels", []))
            rows.append(row)
    return pd.DataFrame(rows)


def filter_month(data: dict, month: str) -> dict:
    """Return a copy of data with all month-keyed sections restricted to the given month."""
    keys_with_month = ("notices", "monthly", "channelDetail", "topPerformers", "needsAttention", "bucketSummary")
    out = dict(data)
    for k in keys_with_month:
        if k in out and isinstance(out[k], dict) and month in out[k]:
            out[k] = {month: out[k][month]}
        elif k in out:
            out[k] = {}
    out["months"] = [month] if month in data.get("months", []) else []
    return out


def main():
    parser = argparse.ArgumentParser(description="Flatten data.json into a multi-sheet Excel")
    parser.add_argument("--month", help="Only include this month, e.g. \"Feb '26\". Default: all months.")
    parser.add_argument("--out", default="data.xlsx", help="Output filename (default: data.xlsx)")
    args = parser.parse_args()

    if not DATA_FILE.exists():
        sys.exit(f"ERROR: {DATA_FILE} not found")
    with open(DATA_FILE, encoding="utf-8") as f:
        data = json.load(f)

    if args.month:
        if args.month not in data.get("months", []):
            sys.exit(f"ERROR: month '{args.month}' not in data.json. Available: {data.get('months', [])}")
        data = filter_month(data, args.month)

    out_path = SCRIPT_DIR / args.out
    with pd.ExcelWriter(out_path, engine="openpyxl") as writer:
        pd.DataFrame([data.get("_meta", {})]).to_excel(writer, sheet_name="_meta", index=False)
        flatten_monthly(data.get("monthly", {})).to_excel(writer, sheet_name="monthly", index=False)
        flatten_channel_detail(data.get("channelDetail", {})).to_excel(writer, sheet_name="channelDetail", index=False)
        flatten_ranked(data.get("topPerformers", {})).to_excel(writer, sheet_name="topPerformers", index=False)
        flatten_ranked(data.get("needsAttention", {})).to_excel(writer, sheet_name="needsAttention", index=False)
        flatten_bucket_summary(data.get("bucketSummary", {})).to_excel(writer, sheet_name="bucketSummary", index=False)
        pd.DataFrame([data.get("benchmarks", {})]).to_excel(writer, sheet_name="benchmarks", index=False)
        notices_df = pd.DataFrame([{"month": m, "notice": n} for m, n in data.get("notices", {}).items()])
        notices_df.to_excel(writer, sheet_name="notices", index=False)

    print(f"Wrote {out_path}")


if __name__ == "__main__":
    main()
