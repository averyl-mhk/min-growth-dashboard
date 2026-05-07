"""processor.py — MIN Growth Marketing Dashboard data processor.

Reads monthly raw exports from raw_exports/YYYY_MM/ and writes data.json.
Run: python processor.py --month 2026_03
"""

import argparse
import json
import sys
from collections import Counter
from datetime import datetime
from fnmatch import fnmatch
from pathlib import Path

try:
    import pandas as pd
except ImportError:
    sys.exit("Missing dependency: pip install pandas openpyxl")

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")


SCRIPT_DIR = Path(__file__).parent
RAW_EXPORTS = SCRIPT_DIR / "raw_exports"
CONTEXT_DIR = SCRIPT_DIR / "context"
OUTPUT_FILE = SCRIPT_DIR / "data.json"
TARGETING_REF = CONTEXT_DIR / "GMD_TargetingType_Reference.xlsx"

DEFAULT_BENCHMARKS = {
    "prospectingCTR": 0.015,
    "retargetingCTR": 0.02,
    "salesCVR": 0.008,
    "amazonACOSTarget": 0.25,
    "roasHigh": 4,
    "roasMid": 2,
    "acosGood": 0.25,
    "acosMid": 0.50,
}

BENCHMARKS_NOTE = (
    "Edit these directly in data.json to change thresholds. The processor NEVER "
    "overwrites an existing benchmarks block — it only writes defaults if the key is absent."
)

AMAZON_BUCKETS = {"SP": "Core Sales", "SB": "Awareness", "SB2": "Awareness", "SD": "Retargeting"}
META_BUCKETS = {"Prospecting": "Awareness", "Remarketing": "Retargeting", "Sales Traffic": "Core Sales"}
FLIPKART_BUCKETS = {"PLA": "Core Sales", "SP": "Core Sales", "SELLER_PCA": "Retargeting", "PCA": "Retargeting"}

BUCKET_KEYS = {"Awareness": "awareness", "Core Sales": "coreSales", "Retargeting": "retargeting"}


# ─── Small helpers ────────────────────────────────────────────────────────────

def to_float(val):
    """Safe float. Returns None for NaN, '-', formula strings, or unparseable text."""
    if val is None:
        return None
    try:
        if pd.isna(val):
            return None
    except (TypeError, ValueError):
        pass
    if isinstance(val, str):
        v = val.strip().replace(",", "")
        if not v or v == "-" or v.startswith("="):
            return None
        try:
            return float(v)
        except ValueError:
            return None
    try:
        return float(val)
    except (TypeError, ValueError):
        return None


def safe_div(num, denom):
    if denom and denom > 0:
        return num / denom
    return None


def round_or_none(v, n=2):
    return round(v, n) if v is not None else None


def read_table(path: Path, **kwargs) -> pd.DataFrame:
    if path.suffix.lower() == ".csv":
        return pd.read_csv(path, **kwargs)
    return pd.read_excel(path, **kwargs)


# ─── File discovery (case-insensitive glob) ───────────────────────────────────

def find_file(folder: Path, includes, excludes=()):
    for p in sorted(folder.iterdir()):
        if not p.is_file():
            continue
        name = p.name.lower()
        if any(fnmatch(name, g.lower()) for g in excludes):
            continue
        if any(fnmatch(name, g.lower()) for g in includes):
            return p
    return None


def find_amazon_campaigns(folder):
    return find_file(folder, ["*mazon*"], ["*search*", "*targeting*"])


def find_sp_search(folder):
    return find_file(folder, ["*sponsored_products*", "*sp*search*"])


def find_sb_search(folder):
    return find_file(folder, ["*sponsored_brands*"])


def find_sd_targeting(folder):
    return find_file(folder, ["*sponsored_display*", "*sd*target*"])


def find_meta(folder):
    return find_file(folder, ["*eta*"], ["*mazon*", "*lipkart*"])


def find_flipkart(folder):
    return find_file(folder, ["*lipkart*"])


def find_pos(folder):
    return find_file(folder, ["*pos*"])


# ─── Keyword classification (Amazon search terms) ─────────────────────────────

def load_targeting_reference():
    if not TARGETING_REF.exists():
        return {"branded": [], "competition": []}
    df = pd.read_excel(TARGETING_REF, header=0)
    competition = [str(v).strip().lower() for v in df.iloc[:, 0].dropna() if str(v).strip()]
    branded = [str(v).strip().lower() for v in df.iloc[:, 1].dropna() if str(v).strip()]
    return {"branded": branded, "competition": competition}


def classify_search_term(term: str, ref: dict) -> str:
    t = str(term).lower()
    for kw in ref["branded"]:
        if kw and kw in t:
            return "Branded"
    for kw in ref["competition"]:
        if kw and kw in t:
            return "Competition"
    return "Generic"


def compute_campaign_keyword_types(search_dfs, ref):
    """Returns {campaign_name: dominant_keyword_type} based on spend share."""
    classified = []
    for df in search_dfs:
        if df is None or df.empty:
            continue
        if not {"Campaign Name", "Customer Search Term", "Spend"}.issubset(df.columns):
            continue
        sub = df[["Campaign Name", "Customer Search Term", "Spend"]].copy()
        sub["Spend"] = pd.to_numeric(sub["Spend"], errors="coerce").fillna(0)
        sub = sub.dropna(subset=["Customer Search Term"])
        sub["kw_type"] = sub["Customer Search Term"].apply(lambda t: classify_search_term(t, ref))
        classified.append(sub)
    if not classified:
        return {}
    combined = pd.concat(classified, ignore_index=True)
    spend_by = combined.groupby(["Campaign Name", "kw_type"])["Spend"].sum().reset_index()
    out = {}
    for campaign, sub in spend_by.groupby("Campaign Name"):
        out[campaign] = sub.loc[sub["Spend"].idxmax(), "kw_type"]
    return out


# ─── Platform loaders ─────────────────────────────────────────────────────────

def _amazon_col(df, *candidates):
    for c in candidates:
        if c in df.columns:
            return c
    return None


def load_amazon_campaigns(path: Path, kw_map: dict):
    df = read_table(path)
    spend_col = _amazon_col(df, "Total cost (converted)", "Total cost")
    rev_col = _amazon_col(df, "Sales (converted)", "Sales")

    rows = []
    unknown_types = Counter()
    for _, r in df.iterrows():
        ad_type = str(r.get("Type", "") or "").strip().upper()
        bucket = AMAZON_BUCKETS.get(ad_type)
        if not bucket:
            if ad_type:
                unknown_types[ad_type] += 1
            continue
        targeting = str(r.get("Targeting", "") or "").strip()
        sub_type = targeting.title() if targeting else None
        campaign = str(r.get("Campaign name", "") or "").strip()

        spend = to_float(r.get(spend_col)) if spend_col else None
        att_rev = to_float(r.get(rev_col)) if rev_col else None
        impressions = to_float(r.get("Impressions"))
        clicks = to_float(r.get("Clicks"))
        acos = to_float(r.get("ACOS"))

        rows.append({
            "platform": "Amazon",
            "ad_type": ad_type,
            "ad_sub_type": sub_type,
            "campaign_name": campaign,
            "bucket": bucket,
            "spend": spend or 0.0,
            "att_rev": att_rev or 0.0,
            "impressions": impressions or 0.0,
            "clicks": clicks or 0.0,
            "acos": acos,
            "ctr": None,
            "keyword_type": kw_map.get(campaign),
            "source_file": path.name,
        })

    # If ACOS was exported as percent (any value > 1), normalize all to decimal
    acos_vals = [r["acos"] for r in rows if r["acos"] is not None]
    if acos_vals and max(acos_vals) > 1:
        for r in rows:
            if r["acos"] is not None:
                r["acos"] = r["acos"] / 100

    for r in rows:
        if r["impressions"] > 0:
            r["ctr"] = r["clicks"] / r["impressions"]

    if unknown_types:
        print(f"  WARNING: Amazon rows skipped — unmapped Type values: {dict(unknown_types)}")
        print(f"  Add these to AMAZON_BUCKETS in processor.py if they should be counted.")

    return rows


def load_meta(path: Path):
    # Row 0 is "Added, Added, ..." junk; row 1 is the real header
    df = pd.read_excel(path, header=1)
    df = df.loc[:, ~df.columns.duplicated()]

    rows = []
    for _, r in df.iterrows():
        keyword_type = str(r.get("Keyword Type", "") or "").strip()
        bucket = META_BUCKETS.get(keyword_type)
        if not bucket:
            continue
        targeting_type = str(r.get("Targeting Type", "") or "").strip() or None

        rev = r.get("Revenue")
        att_rev = 0.0 if (rev is None or pd.isna(rev)) else (to_float(rev) or 0.0)

        rows.append({
            "platform": "Meta",
            "ad_type": targeting_type,
            "ad_sub_type": keyword_type or None,
            "campaign_name": str(r.get("Campaign name", "") or "").strip(),
            "bucket": bucket,
            "spend": to_float(r.get("Amount spent (INR)")) or 0.0,
            "att_rev": att_rev,
            "impressions": to_float(r.get("Impressions")) or 0.0,
            "clicks": to_float(r.get("Link clicks")) or 0.0,
            "acos": None,
            "ctr": None,
            "keyword_type": None,
            "source_file": path.name,
        })

    for r in rows:
        if r["impressions"] > 0:
            r["ctr"] = r["clicks"] / r["impressions"]

    return rows


def load_flipkart(path: Path):
    peek = read_table(path, header=None, nrows=2)
    first = peek.iloc[0, 0]
    is_junk = pd.isna(first) or (isinstance(first, str) and first.strip().lower() == "calculated")
    header_row = 1 if is_junk else 0
    df = read_table(path, header=header_row)

    rows = []
    for _, r in df.iterrows():
        ct = str(r.get("campaign_type", "") or "").strip().upper()
        bucket = FLIPKART_BUCKETS.get(ct)
        if not bucket:
            continue

        rows.append({
            "platform": "Flipkart",
            "ad_type": ct or None,
            "ad_sub_type": None,
            "campaign_name": str(r.get("Campaign Name", "") or "").strip(),
            "bucket": bucket,
            "spend": to_float(r.get("Ad Spend")) or 0.0,
            "att_rev": to_float(r.get("Total Revenue (Rs.)")) or 0.0,
            "impressions": to_float(r.get("Views")) or 0.0,
            "clicks": to_float(r.get("SUM(clicks)")) or 0.0,
            "acos": None,
            "ctr": None,
            "keyword_type": None,
            "source_file": path.name,
        })

    for r in rows:
        if r["impressions"] > 0:
            r["ctr"] = r["clicks"] / r["impressions"]

    return rows


def load_pos(path: Path):
    df = pd.read_excel(path)
    if df.shape[1] < 2:
        return None
    name_col, sales_col = df.columns[0], df.columns[1]

    pos = {}
    for _, r in df.iterrows():
        name = r[name_col]
        if name is None or pd.isna(name):
            continue
        name = str(name).strip()
        if not name:
            continue
        val = to_float(r[sales_col])
        if val is None:
            continue
        pos[name] = val

    if not pos:
        return None
    pos["total"] = sum(pos.values())
    return pos


# ─── Aggregation ──────────────────────────────────────────────────────────────

def agg_platforms(rows):
    by_plat = {}
    for r in rows:
        by_plat.setdefault(r["platform"], []).append(r)

    out = {}
    for plat, label in (("Amazon", "Amazon AMS"), ("Meta", "Meta"), ("Flipkart", "Flipkart")):
        sub = by_plat.get(plat, [])
        if not sub:
            out[label] = {"spend": 0, "attRev": 0, "acos": None, "pending": True}
            continue
        spend = sum(r["spend"] for r in sub)
        att_rev = sum(r["att_rev"] for r in sub)
        acos = round(spend / att_rev, 4) if (plat == "Amazon" and att_rev > 0) else None
        out[label] = {
            "spend": round(spend, 2),
            "attRev": round(att_rev, 2),
            "acos": acos,
            "pending": False,
        }

    out["Google Ads"] = {"spend": 0, "attRev": 0, "acos": None, "pending": True}
    return out


def _group_key(r):
    if r["platform"] == "Amazon":
        return ("Amazon", r["ad_type"], r["ad_sub_type"])
    if r["platform"] == "Meta":
        return ("Meta", r["ad_type"], r["ad_sub_type"])
    if r["platform"] == "Flipkart":
        return ("Flipkart", r["ad_type"], None)
    return None


def _group_label(platform, t1, t2):
    if platform == "Amazon":
        return " ".join(x for x in (t1, t2) if x)
    if platform == "Meta":
        return " + ".join(x for x in (t1, t2) if x)
    return t1 or platform


def agg_channel_detail(rows):
    detail = {"awareness": [], "coreSales": [], "retargeting": []}

    groups = {}
    for r in rows:
        k = _group_key(r)
        if k is None:
            continue
        groups.setdefault((r["bucket"], k), []).append(r)

    for (bucket, k), grows in groups.items():
        platform, t1, t2 = k
        spend = sum(r["spend"] for r in grows)
        att_rev = sum(r["att_rev"] for r in grows)
        impressions = sum(r["impressions"] for r in grows)
        clicks = sum(r["clicks"] for r in grows)
        roas = safe_div(att_rev, spend)
        acos = safe_div(spend, att_rev) if platform == "Amazon" else None
        ctr = safe_div(clicks, impressions)
        ctr_display = f"{ctr * 100:.2f}%" if ctr is not None else None

        targeting = None
        if platform == "Amazon":
            kw_types = [r.get("keyword_type") for r in grows if r.get("keyword_type")]
            if kw_types:
                targeting = Counter(kw_types).most_common(1)[0][0]

        sources = sorted({r["source_file"] for r in grows})

        row = {
            "platform": platform,
            "adType": t1,
            "adSubType": t2,
            "targeting": targeting,
            "campaignGroup": _group_label(platform, t1, t2),
            "spend": round(spend, 2),
            "attRev": round(att_rev, 2),
            "roas": round_or_none(roas, 2),
            "acos": round_or_none(acos, 4),
            "impressions": int(impressions),
            "clicks": int(clicks),
            "ctr": round_or_none(ctr, 4),
            "ctrDisplay": ctr_display,
            "ctrStatus": "na",
            "ctrCvr": ctr_display,
            "source": ", ".join(sources),
        }
        key = BUCKET_KEYS.get(bucket)
        if key:
            detail[key].append(row)

    for key in detail:
        detail[key].sort(key=lambda r: r["spend"], reverse=True)

    for bucket, key in BUCKET_KEYS.items():
        sub = [r for r in rows if r["bucket"] == bucket]
        spend = sum(r["spend"] for r in sub)
        att_rev = sum(r["att_rev"] for r in sub)
        roas = safe_div(att_rev, spend)
        amazon_present = any(r["platform"] == "Amazon" for r in sub)
        acos = safe_div(spend, att_rev) if amazon_present else None
        detail[f"{key}Subtotal"] = {
            "spend": round(spend, 2),
            "attRev": round(att_rev, 2),
            "roas": round_or_none(roas, 2),
            "acos": round_or_none(acos, 4),
            "note": "",
        }

    return detail


def agg_top_bottom(rows, n=5):
    def roas_of(r):
        return r["att_rev"] / r["spend"] if r["spend"] > 0 else 0

    top_pool = [r for r in rows if r["spend"] >= 5000 and r["att_rev"] > 0]
    top_pool.sort(key=roas_of, reverse=True)

    bottom_pool = [r for r in rows if r["spend"] >= 10000 and r["bucket"] != "Awareness"]
    bottom_pool.sort(key=roas_of)

    def to_out(r, rank):
        return {
            "rank": rank,
            "platform": r["platform"],
            "campaignName": r["campaign_name"],
            "campaignFull": r["campaign_name"],
            "bucket": r["bucket"],
            "spend": round(r["spend"], 2),
            "roas": round(roas_of(r), 2),
            "acos": round_or_none(r["acos"], 4) if r["platform"] == "Amazon" else None,
            "source": r["source_file"],
        }

    top = [to_out(r, i + 1) for i, r in enumerate(top_pool[:n])]
    bottom = [to_out(r, i + 1) for i, r in enumerate(bottom_pool[:n])]
    return top, bottom


def agg_bucket_summary(rows):
    out = {}
    for bucket, key in BUCKET_KEYS.items():
        sub = [r for r in rows if r["bucket"] == bucket]
        spend = sum(r["spend"] for r in sub)
        att_rev = sum(r["att_rev"] for r in sub)
        roas = safe_div(att_rev, spend) or 0
        chans = {}
        for r in sub:
            chans[r["platform"]] = chans.get(r["platform"], 0) + r["spend"]
        out[key] = {
            "totalSpend": round(spend, 2),
            "roas": round(roas, 2),
            "channels": [{"platform": p, "spend": round(s, 2)} for p, s in chans.items()],
        }
    return out


# ─── Main ─────────────────────────────────────────────────────────────────────

def process_month(month_str: str):
    folder = RAW_EXPORTS / month_str
    if not folder.exists():
        sys.exit(f"ERROR: Folder not found: {folder}")

    year, month = month_str.split("_")
    month_label = datetime(int(year), int(month), 1).strftime("%b '%y")

    print(f"Processing {month_str} → {month_label}")
    print(f"Folder: {folder}\n")

    missing = []
    source_files = []

    # Search term files (for Amazon keyword classification)
    ref = load_targeting_reference()
    sp_path = find_sp_search(folder)
    sb_path = find_sb_search(folder)
    search_dfs = []
    if sp_path:
        search_dfs.append(pd.read_excel(sp_path))
        source_files.append(sp_path.name)
    if sb_path:
        search_dfs.append(pd.read_excel(sb_path))
        source_files.append(sb_path.name)
    kw_map = compute_campaign_keyword_types(search_dfs, ref) if search_dfs else {}

    # SD targeting (recorded in source files; not aggregated separately — SD totals come from the campaign report)
    sd_path = find_sd_targeting(folder)
    if sd_path:
        source_files.append(sd_path.name)

    all_rows = []

    amazon_path = find_amazon_campaigns(folder)
    if amazon_path:
        all_rows.extend(load_amazon_campaigns(amazon_path, kw_map))
        source_files.append(amazon_path.name)
    else:
        missing.append("Amazon_Campaigns")

    meta_path = find_meta(folder)
    if meta_path:
        all_rows.extend(load_meta(meta_path))
        source_files.append(meta_path.name)
    else:
        missing.append("Meta_Campaigns")

    flipkart_path = find_flipkart(folder)
    if flipkart_path:
        all_rows.extend(load_flipkart(flipkart_path))
        source_files.append(flipkart_path.name)
    else:
        missing.append("Flipkart_Campaigns")

    pos_path = find_pos(folder)
    pos = load_pos(pos_path) if pos_path else None
    if pos_path:
        source_files.append(pos_path.name)
    else:
        missing.append("POS")

    # When search term files are missing, set keyword_type to null on every Amazon row
    if not search_dfs:
        for r in all_rows:
            if r["platform"] == "Amazon":
                r["keyword_type"] = None

    platforms = agg_platforms(all_rows)
    total_spend = sum(p["spend"] for p in platforms.values() if not p.get("pending"))
    total_att_rev = sum(p["attRev"] for p in platforms.values() if not p.get("pending"))
    channel_detail = agg_channel_detail(all_rows)
    top, bottom = agg_top_bottom(all_rows)
    bucket_summary = agg_bucket_summary(all_rows)

    if OUTPUT_FILE.exists():
        with open(OUTPUT_FILE, encoding="utf-8") as f:
            data = json.load(f)
    else:
        data = {}

    data.setdefault("_meta", {})
    data.setdefault("months", [])
    data.setdefault("notices", {})
    data.setdefault("monthly", {})
    data.setdefault("channelDetail", {})
    data.setdefault("topPerformers", {})
    data.setdefault("needsAttention", {})
    data.setdefault("bucketSummary", {})
    if "benchmarks" not in data:
        data["benchmarks"] = DEFAULT_BENCHMARKS.copy()
    data["_benchmarks_note"] = BENCHMARKS_NOTE

    data["_meta"].update({
        "dashboard": "MIN Growth Marketing Dashboard",
        "marketingArm": "Meyer India (MIN)",
        "lastUpdated": month_str,
        "currency": "INR",
        "generatedBy": "processor.py",
        "sourceFiles": source_files,
        "missing": missing,
    })

    if month_label not in data["months"]:
        data["months"].append(month_label)
        data["months"].sort(key=lambda m: datetime.strptime(m, "%b '%y"))

    notice = f"{month_label} — Generated from raw exports on {datetime.today().strftime('%Y-%m-%d')}."
    if missing:
        notice += f" Missing: {', '.join(missing)}."
    data["notices"][month_label] = notice

    data["monthly"][month_label] = {
        "spend": round(total_spend, 2),
        "attRev": round(total_att_rev, 2),
        "pos": pos,
        "platforms": platforms,
    }
    data["channelDetail"][month_label] = channel_detail
    data["topPerformers"][month_label] = top
    data["needsAttention"][month_label] = bottom
    data["bucketSummary"][month_label] = bucket_summary

    with open(OUTPUT_FILE, "w", encoding="utf-8") as f:
        json.dump(data, f, indent=2, ensure_ascii=False)

    _print_summary(month_label, platforms, total_spend, total_att_rev, pos, missing, source_files)


def _print_summary(month_label, platforms, total_spend, total_att_rev, pos, missing, source_files):
    print(f"✓ {month_label} written to data.json")
    plat_strs = []
    for plat, label in (("Amazon", "Amazon AMS"), ("Flipkart", "Flipkart"), ("Meta", "Meta")):
        p = platforms.get(label, {})
        if p.get("pending"):
            plat_strs.append(f"{plat} MISSING")
        else:
            plat_strs.append(f"{plat} ₹{p['spend'] / 100000:.1f}L spend")
    print(f"  Platforms:    {' | '.join(plat_strs)}")
    print(f"  Total spend:  ₹{total_spend:,.0f}")
    print(f"  Total attRev: ₹{total_att_rev:,.0f}")
    roas = safe_div(total_att_rev, total_spend)
    print(f"  ROAS:         {roas:.2f}×" if roas else "  ROAS:         N/A")
    print(f"  POS:          {'₹' + format(pos['total'], ',.0f') if pos else 'MISSING'}")
    print(f"  Missing:      {', '.join(missing) if missing else 'none'}")
    print(f"  Source files: {source_files}")


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="MIN dashboard data processor")
    parser.add_argument("--month", required=True, help="Month folder name, e.g. 2026_03")
    args = parser.parse_args()
    process_month(args.month)
