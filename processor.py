"""
GMD_DataProcessor_v1.py
MIN Growth Marketing Dashboard — Raw Export Processor

Reads raw platform exports for a given month and writes GMD_Dashboard_MIN_data.json.
Every number in the JSON is derived mechanically from the source files — no manual editing.

Usage:
    python GMD_DataProcessor_v1.py --month 2026_03

Requirements:
    pip install pandas openpyxl

Input files (place in raw_exports/YYYY_MM/):
    Amazon_SP_Campaigns.xlsx    — SP campaign-level report from Amazon AMS
    Amazon_SB_Campaigns.xlsx    — SB campaign-level report
    Amazon_SD_Campaigns.xlsx    — SD campaign-level report
    Amazon_SP_SearchTerms.xlsx  — SP search term report (for keyword classification)
    Amazon_SB_SearchTerms.xlsx  — SB search term report
    Meta_Campaigns.xlsx         — Meta Ads Manager campaign-level export
    Flipkart_Campaigns.xlsx     — Flipkart Seller Hub campaign-level export
    POS_Manual.xlsx             — Manual POS table (Platform | Total Sales)

Reference files (in context/):
    GMD_TargetingType_Reference.xlsx — Branded / Competition / Generic keyword lists

Output:
    outputs/GMD_Dashboard_MIN_data.json  (overwrites)
"""

import argparse
import json
import os
import sys
from datetime import datetime
from pathlib import Path

try:
    import pandas as pd
except ImportError:
    sys.exit("Missing dependency: pip install pandas openpyxl")


# ─── Paths ────────────────────────────────────────────────────────────────────

SCRIPT_DIR   = Path(__file__).parent
PROJECT_ROOT = SCRIPT_DIR.parent          # growth-marketing-dashboard/
CONTEXT_DIR  = PROJECT_ROOT / "context"
OUTPUT_FILE   = PROJECT_ROOT / "data.json"
TARGETING_REF = CONTEXT_DIR / "GMD_TargetingType_Reference.xlsx"


# ─── Benchmarks (edit here if targets change) ────────────────────────────────

BENCHMARKS = {
    "prospectingCTR":   0.015,
    "retargetingCTR":   0.02,
    "salesCVR":         0.008,
    "amazonACOSTarget": 0.25,
    "roasHigh": 4,
    "roasMid":  2,
    "acosGood": 0.25,
    "acosMid":  0.50,
}


# ─── Campaign categorisation ──────────────────────────────────────────────────
#
# Buckets are assigned by ad type (SP / SB / SD) and objective:
#   Awareness   → Amazon SB (Sponsored Brands)
#   Core Sales  → Amazon SP (Sponsored Products), Flipkart PLA, Meta Sales Traffic
#   Retargeting → Amazon SD (Sponsored Display), Meta Remarketing, Flipkart PCA
#
# Keyword targeting type (Branded / Competition / Generic) is derived from the
# search term reports using the reference file in context/.

CAMPAIGN_BUCKET_RULES = {
    # Amazon
    "SP": "Core Sales",
    "SB": "Awareness",
    "SD": "Retargeting",
    # Meta — determined by campaign name keywords below
    # Flipkart
    "PLA": "Core Sales",
    "PCA": "Retargeting",
}

META_RETARGETING_KEYWORDS = ["remarketing", "retarget", "re-target", "catalogue"]
META_AWARENESS_KEYWORDS   = ["prospecting", "awareness", "instream", "video"]
# Anything else in Meta → Core Sales


# ─── Column name maps ─────────────────────────────────────────────────────────
#
# TODO: Confirm exact column names once Akash's files arrive.
# Update the values below to match the actual header row in each export.
# Keys are the canonical names used in this script; values are what appears in the file.

AMAZON_COLS = {
    # TODO: update these after inspecting the actual export
    "campaign_name":  "Campaign Name",      # e.g. "SP | Cast Iron All | Auto"
    "ad_type":        "Ad Type",            # SP / SB / SD
    "targeting_type": "Targeting Type",     # Auto / Manual
    "impressions":    "Impressions",
    "clicks":         "Clicks",
    "spend":          "Spend",              # in ₹
    "attributed_rev": "14 Day Total Sales", # attributed revenue in ₹ — confirm exact name
    "orders":         "14 Day Total Orders",
    "acos":           "ACOS",               # decimal (0.25) or percent (25%) — check
}

META_COLS = {
    # TODO: update these after inspecting the actual export
    "campaign_name":  "Campaign name",
    "objective":      "Objective",
    "spend":          "Amount spent (INR)",
    "attributed_rev": "Purchases conversion value",
    "impressions":    "Impressions",
    "clicks":         "Link clicks",
    "ctr":            "CTR (link click-through rate)",
}

FLIPKART_COLS = {
    # TODO: update these after inspecting the actual export
    "campaign_name":  "Campaign Name",
    "ad_type":        "Type",          # PLA / PCA
    "targeting_type": "Targeting",     # Auto / Manual
    "spend":          "Spend",
    "attributed_rev": "Revenue",
    "impressions":    "Impressions",
    "clicks":         "Clicks",
}

POS_COLS = {
    # Simple manual table: two columns
    "platform": "Platform",    # Amazon / Shopify / Flipkart
    "sales":    "Total Sales",
}


# ─── Keyword classification ───────────────────────────────────────────────────

def load_targeting_reference():
    """
    Load branded/competition keyword lists from the reference file.
    Returns dict: { 'branded': [...], 'competition': [...] }
    Anything not matching either list = Generic.
    """
    if not TARGETING_REF.exists():
        print(f"  WARNING: Targeting reference not found at {TARGETING_REF}")
        print("  Keyword classification will be skipped — all search terms labelled Generic.")
        return {"branded": [], "competition": []}

    df = pd.read_excel(TARGETING_REF, header=0)
    # Columns: Competition | Branded | Generic (Generic is just a rule, not a list)
    competition = [str(v).strip().lower() for v in df.iloc[:, 0].dropna() if str(v).strip()]
    branded     = [str(v).strip().lower() for v in df.iloc[:, 1].dropna() if str(v).strip()]

    # Remove the header row values if they snuck in
    competition = [v for v in competition if v != "competition"]
    branded     = [v for v in branded     if v not in ("branded", "generic")]

    print(f"  Targeting reference loaded: {len(competition)} competition terms, {len(branded)} branded terms")
    return {"branded": branded, "competition": competition}


def classify_search_term(term: str, ref: dict) -> str:
    """
    Classify a single Amazon search term as Branded / Competition / Generic.
    Matching is case-insensitive substring check.
    """
    term_lower = term.lower()
    for kw in ref["branded"]:
        if kw in term_lower:
            return "Branded"
    for kw in ref["competition"]:
        if kw in term_lower:
            return "Competition"
    return "Generic"


# ─── Platform loaders ─────────────────────────────────────────────────────────

def load_amazon(folder: Path, ad_type: str, ref: dict) -> pd.DataFrame:
    """
    Load an Amazon campaign-level export (SP, SB, or SD).
    Returns a cleaned DataFrame with canonical column names.
    """
    filename = f"Amazon_{ad_type}_Campaigns.xlsx"
    filepath = folder / filename

    if not filepath.exists():
        print(f"  MISSING: {filename} — skipping {ad_type}")
        return pd.DataFrame()

    print(f"  Loading {filename}...")
    df = pd.read_excel(filepath)

    # TODO: validate that expected columns exist, print helpful error if not
    # Once real files arrive, add: assert AMAZON_COLS['spend'] in df.columns, f"Column '{AMAZON_COLS['spend']}' not found. Available: {list(df.columns)}"

    rename = {v: k for k, v in AMAZON_COLS.items() if v in df.columns}
    df = df.rename(columns=rename)

    # Normalise ACOS: convert percent to decimal if needed
    if "acos" in df.columns:
        if df["acos"].dropna().max() > 1:
            df["acos"] = df["acos"] / 100

    df["platform"]    = "Amazon"
    df["bucket"]      = CAMPAIGN_BUCKET_RULES.get(ad_type, "Unknown")
    df["source_file"] = filename

    print(f"    {len(df)} rows loaded")
    return df


def load_meta(folder: Path) -> pd.DataFrame:
    """Load Meta Ads Manager campaign-level export."""
    filename = "Meta_Campaigns.xlsx"
    filepath = folder / filename

    if not filepath.exists():
        print(f"  MISSING: {filename} — skipping Meta")
        return pd.DataFrame()

    print(f"  Loading {filename}...")
    df = pd.read_excel(filepath)

    rename = {v: k for k, v in META_COLS.items() if v in df.columns}
    df = df.rename(columns=rename)

    # Classify into bucket by campaign name
    def meta_bucket(name):
        name_lower = str(name).lower()
        if any(kw in name_lower for kw in META_RETARGETING_KEYWORDS):
            return "Retargeting"
        if any(kw in name_lower for kw in META_AWARENESS_KEYWORDS):
            return "Awareness"
        return "Core Sales"

    df["platform"]    = "Meta"
    df["bucket"]      = df["campaign_name"].apply(meta_bucket) if "campaign_name" in df.columns else "Unknown"
    df["source_file"] = filename

    print(f"    {len(df)} rows loaded")
    return df


def load_flipkart(folder: Path) -> pd.DataFrame:
    """Load Flipkart Seller Hub campaign-level export."""
    filename = "Flipkart_Campaigns.xlsx"
    filepath = folder / filename

    if not filepath.exists():
        print(f"  MISSING: {filename} — skipping Flipkart")
        return pd.DataFrame()

    print(f"  Loading {filename}...")
    df = pd.read_excel(filepath)

    rename = {v: k for k, v in FLIPKART_COLS.items() if v in df.columns}
    df = df.rename(columns=rename)

    df["platform"]    = "Flipkart"
    df["bucket"]      = df["ad_type"].map(CAMPAIGN_BUCKET_RULES).fillna("Unknown") if "ad_type" in df.columns else "Unknown"
    df["source_file"] = filename

    print(f"    {len(df)} rows loaded")
    return df


def load_pos(folder: Path) -> dict:
    """Load manual POS table. Returns dict: { 'Amazon': 0, 'Shopify': 0, 'Flipkart': 0, 'total': 0 }"""
    filename = "POS_Manual.xlsx"
    filepath = folder / filename

    if not filepath.exists():
        print(f"  MISSING: {filename} — POS will be null for this month")
        return None

    print(f"  Loading {filename}...")
    df = pd.read_excel(filepath)
    rename = {v: k for k, v in POS_COLS.items() if v in df.columns}
    df = df.rename(columns=rename)

    pos = {}
    for _, row in df.iterrows():
        pos[str(row["platform"]).strip()] = float(row["sales"])

    pos["total"] = sum(pos.values())
    print(f"    POS loaded: {pos}")
    return pos


# ─── Aggregation helpers ──────────────────────────────────────────────────────

def safe_div(num, denom, default=None):
    try:
        if denom and denom > 0:
            return round(num / denom, 4)
    except Exception:
        pass
    return default


def agg_platform_summary(all_df: pd.DataFrame) -> dict:
    """
    Aggregate spend, attRev, and ACOS by platform.
    Returns dict matching the 'platforms' key structure in the JSON.
    """
    platforms = {}
    platform_map = {
        "Amazon": "Amazon AMS",
        "Meta":   "Meta",
        "Flipkart": "Flipkart",
    }

    for plat, label in platform_map.items():
        sub = all_df[all_df["platform"] == plat] if not all_df.empty else pd.DataFrame()
        if sub.empty:
            platforms[label] = {"spend": 0, "attRev": 0, "acos": None, "pending": False}
            continue

        spend    = float(sub["spend"].sum())            if "spend"          in sub.columns else 0
        att_rev  = float(sub["attributed_rev"].sum())   if "attributed_rev" in sub.columns else 0

        # ACOS: only meaningful for Amazon
        acos = None
        if plat == "Amazon" and "acos" in sub.columns:
            # weighted ACOS = total spend / total attributed revenue
            acos = round(safe_div(spend, att_rev) or 0, 4) if att_rev > 0 else None

        platforms[label] = {
            "spend":   round(spend, 2),
            "attRev":  round(att_rev, 2),
            "acos":    acos,
            "pending": False,
        }

    # Google Ads — always pending until data is available
    platforms["Google Ads"] = {"spend": 0, "attRev": 0, "acos": None, "pending": True}

    return platforms


def agg_channel_detail(all_df: pd.DataFrame) -> dict:
    """
    Build the channelDetail structure: awareness / coreSales / retargeting rows.
    Each row = one campaign (or campaign group if deduplication is needed).
    TODO: once real files arrive, decide whether to keep per-campaign rows or group by ad_type+targeting.
    """
    if all_df.empty:
        return {}

    def make_rows(bucket: str) -> list:
        sub = all_df[all_df["bucket"] == bucket].copy()
        rows = []
        for _, r in sub.iterrows():
            spend     = float(r.get("spend", 0) or 0)
            att_rev   = float(r.get("attributed_rev", 0) or 0) if r.get("attributed_rev") else None
            roas      = safe_div(att_rev, spend) if att_rev else None
            acos      = float(r.get("acos", 0)) if r.get("acos") else None
            ctr       = float(r.get("ctr", 0))  if r.get("ctr")  else None

            rows.append({
                "platform":    str(r.get("platform", "")),
                "adType":      str(r.get("ad_type", "")),
                "adSubType":   str(r.get("targeting_type", "")) if r.get("targeting_type") else None,
                "targeting":   str(r.get("keyword_type", ""))   if r.get("keyword_type")   else None,
                "campaignName":str(r.get("campaign_name", "")),
                "source":      str(r.get("source_file", "")),
                "spend":       round(spend, 2),
                "attRev":      round(att_rev, 2) if att_rev else None,
                "roas":        round(roas, 2)    if roas    else None,
                "acos":        round(acos, 4)    if acos    else None,
                "ctr":         round(ctr, 4)     if ctr     else None,
                "ctrDisplay":  None,   # will be formatted in dashboard JS
                "ctrStatus":   "na",
                "ctrCvr":      None,
            })
        return rows

    detail = {}
    for bucket, key in [("Awareness", "awareness"), ("Core Sales", "coreSales"), ("Retargeting", "retargeting")]:
        rows = make_rows(bucket)
        sub = all_df[all_df["bucket"] == bucket]
        total_spend  = float(sub["spend"].sum())           if "spend"          in sub.columns else 0
        total_attrev = float(sub["attributed_rev"].sum())  if "attributed_rev" in sub.columns else 0
        roas         = safe_div(total_attrev, total_spend)
        acos         = safe_div(total_spend, total_attrev) if total_attrev else None

        detail[key] = rows
        detail[f"{key}Subtotal"] = {
            "spend":  round(total_spend, 2),
            "attRev": round(total_attrev, 2),
            "roas":   round(roas, 2) if roas else None,
            "acos":   round(acos, 4) if acos else None,
            "note":   "",
        }

    return detail


def agg_top_bottom(all_df: pd.DataFrame, n: int = 5) -> tuple:
    """
    Returns (top_performers, needs_attention) — each a list of n campaign dicts.
    Ranked by ROAS, minimum spend threshold applied.
    """
    if all_df.empty or "spend" not in all_df.columns or "attributed_rev" not in all_df.columns:
        return [], []

    df = all_df.copy()
    df["spend"]         = pd.to_numeric(df["spend"], errors="coerce").fillna(0)
    df["attributed_rev"]= pd.to_numeric(df["attributed_rev"], errors="coerce").fillna(0)
    df["roas"]          = df.apply(lambda r: safe_div(r["attributed_rev"], r["spend"]) or 0, axis=1)
    df["acos"]          = pd.to_numeric(df.get("acos"), errors="coerce") if "acos" in df.columns else None

    # Top performers: min spend ₹5,000, any bucket, ranked by ROAS desc
    top = (df[df["spend"] >= 5000]
           .sort_values("roas", ascending=False)
           .head(n))

    # Needs attention: min spend ₹10,000, exclude awareness, ranked by ROAS asc
    bottom = (df[(df["spend"] >= 10000) & (df["bucket"] != "Awareness")]
              .sort_values("roas", ascending=True)
              .head(n))

    def row_to_dict(r, rank):
        return {
            "rank":         rank,
            "platform":     str(r.get("platform", "")),
            "campaignName": str(r.get("campaign_name", "")),
            "campaignFull": str(r.get("campaign_name", "")),
            "bucket":       str(r.get("bucket", "")),
            "spend":        round(float(r.get("spend", 0)), 2),
            "roas":         round(float(r.get("roas", 0)), 2),
            "acos":         round(float(r["acos"]), 4) if r.get("acos") and not pd.isna(r["acos"]) else None,
            "source":       str(r.get("source_file", "")),
        }

    return (
        [row_to_dict(r, i + 1) for i, (_, r) in enumerate(top.iterrows())],
        [row_to_dict(r, i + 1) for i, (_, r) in enumerate(bottom.iterrows())],
    )


# ─── Main ─────────────────────────────────────────────────────────────────────

def process_month(month_str: str):
    """
    Main entry point. Reads all files for the given month, builds the JSON, writes it out.
    month_str format: YYYY_MM  e.g. "2026_03"
    """
    folder = PROJECT_ROOT / "raw_exports" / month_str
    if not folder.exists():
        sys.exit(f"ERROR: Folder not found: {folder}\nCreate it and drop the export files in.")

    # Parse month label for display  e.g. "2026_03" → "Mar '26"
    year, month = month_str.split("_")
    month_label = datetime(int(year), int(month), 1).strftime("%b '%y")

    print(f"\n{'='*60}")
    print(f"Processing: {month_str}  →  {month_label}")
    print(f"Folder: {folder}")
    print(f"{'='*60}\n")

    # 1. Load targeting reference
    print("Loading targeting reference...")
    ref = load_targeting_reference()

    # 2. Load all platform files
    print("\nLoading platform exports...")
    amazon_sp = load_amazon(folder, "SP", ref)
    amazon_sb = load_amazon(folder, "SB", ref)
    amazon_sd = load_amazon(folder, "SD", ref)
    meta      = load_meta(folder)
    flipkart  = load_flipkart(folder)
    pos       = load_pos(folder)

    # 3. Combine into one DataFrame
    frames = [df for df in [amazon_sp, amazon_sb, amazon_sd, meta, flipkart] if not df.empty]
    if not frames:
        sys.exit("ERROR: No data files loaded. Check that files are in the folder and named correctly.")

    all_df = pd.concat(frames, ignore_index=True)
    print(f"\nTotal rows across all platforms: {len(all_df)}")

    # 4. Aggregate
    print("\nAggregating...")
    platforms = agg_platform_summary(all_df)
    total_spend  = sum(p["spend"]  for p in platforms.values() if not p.get("pending"))
    total_attrev = sum(p["attRev"] for p in platforms.values() if not p.get("pending"))

    channel_detail = agg_channel_detail(all_df)
    tops, bottoms  = agg_top_bottom(all_df)

    # 5. Load existing JSON to preserve other months
    print("\nUpdating JSON...")
    if OUTPUT_FILE.exists():
        with open(OUTPUT_FILE) as f:
            data = json.load(f)
    else:
        data = {
            "_meta": {}, "months": [], "notices": {}, "monthly": {},
            "benchmarks": BENCHMARKS,
            "channelDetail": {}, "topPerformers": {}, "needsAttention": {}, "bucketSummary": {},
        }

    # 6. Add / overwrite this month
    source_files = []
    for df in frames:
        if "source_file" in df.columns:
            source_files.extend(df["source_file"].unique().tolist())

    if month_label not in data["months"]:
        data["months"].append(month_label)
        data["months"].sort()  # keep chronological

    data["notices"][month_label] = f"{month_label} — Generated from raw platform exports on {datetime.today().strftime('%Y-%m-%d')}."

    data["monthly"][month_label] = {
        "spend":     round(total_spend, 2),
        "attRev":    round(total_attrev, 2),
        "pos":       pos,
        "platforms": platforms,
    }

    data["channelDetail"][month_label]  = channel_detail
    data["topPerformers"][month_label]  = tops
    data["needsAttention"][month_label] = bottoms
    data["bucketSummary"][month_label]  = {}  # TODO: populate once CTR/CVR range logic is confirmed

    data["_meta"]["lastUpdated"] = month_str
    data["_meta"]["generatedBy"] = "GMD_DataProcessor_v1.py"
    data["_meta"]["sourceFiles"] = list(set(data["_meta"].get("sourceFiles", []) + source_files))

    # 7. Write JSON
    with open(OUTPUT_FILE, "w") as f:
        json.dump(data, f, indent=2, ensure_ascii=False)

    print(f"\n✓ JSON written to: {OUTPUT_FILE}")
    print(f"\nSummary for {month_label}:")
    print(f"  Total spend:     ₹{total_spend:,.0f}")
    print(f"  Total att. rev:  ₹{total_attrev:,.0f}")
    print(f"  Overall ROAS:    {safe_div(total_attrev, total_spend):.2f}×" if total_spend else "  Overall ROAS: N/A")
    if pos:
        print(f"  Total POS:       ₹{pos.get('total', 0):,.0f}")
    print(f"\n  Source files:    {source_files}")
    print(f"\nNEXT STEPS:")
    print(f"  1. Open GMD_Dashboard_MIN_data.json — verify numbers against your raw exports")
    print(f"  2. Open GMD_Dashboard_MIN_v4.html   — check the dashboard looks right")
    print(f"  3. If column names were wrong, update the _COLS dicts at the top of this script and re-run")


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="MIN dashboard data processor")
    parser.add_argument("--month", required=True, help="Month folder name, e.g. 2026_03")
    args = parser.parse_args()
    process_month(args.month)
