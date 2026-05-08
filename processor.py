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

AMAZON_BUCKETS = {"SP": "Core Sales", "SB": "Awareness", "SBV": "Awareness", "SD": "Retargeting"}
# Campaign report Type values that get normalized into a canonical ad_type
AMAZON_TYPE_NORMALIZE = {"SB2": "SBV"}  # SB2 (Sponsored Brands video) is treated as SBV
SP_AUTO_TARGETS = {"loose-match", "close-match", "complements", "substitutes"}

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
    # *sp*search* alone would match "Sponsored_Brands_Search..." since "sponsored" starts with "sp",
    # so the SB and SD files must be explicitly excluded.
    return find_file(folder, ["*sponsored_products*", "*sp*search*"], excludes=["*brands*", "*display*"])


def find_sb_search(folder):
    return find_file(folder, ["*sponsored_brands*", "*sb*search*"], excludes=["*products*", "*display*"])


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


def _sp_auto_or_manual(targeting_value):
    """SP search term Targeting → 'Auto' if value is one of Amazon's auto-targeting tokens, else 'Manual'."""
    if targeting_value is None:
        return "Manual"
    try:
        if pd.isna(targeting_value):
            return "Manual"
    except (TypeError, ValueError):
        pass
    return "Auto" if str(targeting_value).strip().lower() in SP_AUTO_TARGETS else "Manual"


def compute_keyword_ratios(sp_df, sb_df, ref):
    """Build per-campaign spend and revenue ratios per keyword type from search term reports.

    Returns: {campaign_name: {auto_manual: {kw_type: (spend_ratio, rev_ratio)}}}
    Ratios are normalized so they sum to 1.0 across kw_types within each
    (campaign_name, auto_manual) bucket. Campaigns with zero search-term spend
    are excluded — the caller falls back to 100% Generic for those.
    """
    sources = []
    if sp_df is not None and not sp_df.empty:
        sources.append((sp_df, "7 Day Total Sales (₹)", _sp_auto_or_manual))
    if sb_df is not None and not sb_df.empty:
        sources.append((sb_df, "14 Day Total Sales (₹)", lambda _: "Manual"))

    ratios = {}
    for df, sales_col, am_fn in sources:
        if not {"Campaign Name", "Customer Search Term", "Spend"}.issubset(df.columns):
            continue
        sub = df.copy()
        sub["_campaign"] = sub["Campaign Name"].astype(str).str.strip()
        sub["_kw"] = sub["Customer Search Term"].apply(
            lambda t: classify_search_term(t, ref) if pd.notna(t) else "Generic"
        )
        targeting_series = sub["Targeting"] if "Targeting" in sub.columns else pd.Series([None] * len(sub))
        sub["_am"] = targeting_series.apply(am_fn)
        sub["_spend"] = pd.to_numeric(sub["Spend"], errors="coerce").fillna(0)
        sub["_rev"] = (
            pd.to_numeric(sub[sales_col], errors="coerce").fillna(0) if sales_col in sub.columns else 0
        )

        agg = sub.groupby(["_campaign", "_am", "_kw"]).agg(spend=("_spend", "sum"), rev=("_rev", "sum"))
        for (campaign, am), grp in agg.groupby(level=[0, 1]):
            total_spend = float(grp["spend"].sum())
            total_rev = float(grp["rev"].sum())
            if total_spend <= 0:
                continue
            kw_to_ratio = {}
            for kw in grp.index.get_level_values("_kw"):
                kw_spend = float(grp.loc[(campaign, am, kw), "spend"])
                kw_rev = float(grp.loc[(campaign, am, kw), "rev"])
                spend_ratio = kw_spend / total_spend
                rev_ratio = (kw_rev / total_rev) if total_rev > 0 else spend_ratio
                kw_to_ratio[kw] = (spend_ratio, rev_ratio)
            ratios.setdefault(campaign, {})[am] = kw_to_ratio
    return ratios


def expand_amazon_by_keyword_type(rows, ratios):
    """Split Amazon SP/SB/SBV campaign rows into per-keyword-type sub-rows.

    Spend, revenue, impressions, and clicks are scaled by the search-term-derived
    spend/revenue ratios, so per-campaign totals (and therefore platform/bucket
    totals) reconcile exactly with the campaign report. Campaigns missing from
    the search term reports get a single fallback row at 100% Generic.

    SD rows pass through with keyword_type=None (renders as "—" in the dashboard).
    Non-Amazon rows pass through unchanged.
    """
    splittable = {"SP", "SB", "SBV"}
    expanded = []
    fallback = []

    for r in rows:
        if r["platform"] != "Amazon" or r["ad_type"] not in splittable:
            new_r = dict(r)
            new_r.setdefault("keyword_type", None)
            expanded.append(new_r)
            continue

        campaign = r["campaign_name"].strip()
        am = r.get("ad_sub_type") or "Manual"
        c_ratios = ratios.get(campaign, {}).get(am)
        if not c_ratios:
            fallback.append((campaign, am, r["ad_type"]))
            new_r = dict(r)
            new_r["keyword_type"] = "Generic"
            expanded.append(new_r)
            continue

        for kw_type, (s_ratio, r_ratio) in c_ratios.items():
            new_r = dict(r)
            new_r["spend"] = r["spend"] * s_ratio
            new_r["att_rev"] = r["att_rev"] * r_ratio
            new_r["impressions"] = r["impressions"] * s_ratio
            new_r["clicks"] = r["clicks"] * s_ratio
            new_r["keyword_type"] = kw_type
            expanded.append(new_r)

    return expanded, fallback


# ─── Platform loaders ─────────────────────────────────────────────────────────

def _amazon_col(df, *candidates):
    for c in candidates:
        if c in df.columns:
            return c
    return None


def load_amazon_campaigns(path: Path):
    df = read_table(path)
    spend_col = _amazon_col(df, "Total cost (converted)", "Total cost")
    rev_col = _amazon_col(df, "Sales (converted)", "Sales")

    rows = []
    unknown_types = Counter()
    for _, r in df.iterrows():
        raw_type = str(r.get("Type", "") or "").strip().upper()
        ad_type = AMAZON_TYPE_NORMALIZE.get(raw_type, raw_type)
        bucket = AMAZON_BUCKETS.get(ad_type)
        if not bucket:
            if raw_type:
                unknown_types[raw_type] += 1
            continue
        targeting = str(r.get("Targeting", "") or "").strip().upper()
        if targeting == "AUTOMATIC":
            sub_type = "Auto"
        elif targeting == "MANUAL":
            sub_type = "Manual"
        elif ad_type in ("SB", "SBV"):
            sub_type = "Manual"  # Sponsored Brands has no auto targeting
        else:
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
            "keyword_type": None,
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


AD_TYPE_ORDER = {"SP": 0, "SB": 1, "SBV": 2, "SD": 3}
SUB_TYPE_ORDER = {"Auto": 0, "Manual": 1}
KW_TYPE_ORDER = {"Branded": 0, "Competition": 1, "Generic": 2, None: 3}


def _group_key(r):
    """4-tuple group key. Amazon includes keyword_type; other platforms use None for that slot."""
    kw = r.get("keyword_type") if r["platform"] == "Amazon" else None
    if r["platform"] == "Amazon":
        return ("Amazon", r["ad_type"], r["ad_sub_type"], kw)
    if r["platform"] == "Meta":
        return ("Meta", r["ad_type"], r["ad_sub_type"], None)
    if r["platform"] == "Flipkart":
        return ("Flipkart", r["ad_type"], None, None)
    return None


def _group_label(platform, t1, t2, kw):
    if platform == "Amazon":
        return " ".join(x for x in (t1, t2, kw) if x)
    if platform == "Meta":
        return " + ".join(x for x in (t1, t2) if x)
    return t1 or platform


def _channel_sort_key(row):
    """Spec ordering: ad_type → SP, SB, SBV, SD; auto_manual → Auto, Manual; kw_type → Branded, Competition, Generic, None."""
    return (
        AD_TYPE_ORDER.get(row.get("adType"), 99),
        SUB_TYPE_ORDER.get(row.get("adSubType"), 99),
        KW_TYPE_ORDER.get(row.get("targeting"), 99),
        -row.get("spend", 0),  # spend desc within ties (e.g. Meta/Flipkart with no kw_type)
    )


def agg_channel_detail(rows):
    detail = {"awareness": [], "coreSales": [], "retargeting": []}

    groups = {}
    for r in rows:
        k = _group_key(r)
        if k is None:
            continue
        groups.setdefault((r["bucket"], k), []).append(r)

    for (bucket, k), grows in groups.items():
        platform, t1, t2, kw = k
        spend = sum(r["spend"] for r in grows)
        att_rev = sum(r["att_rev"] for r in grows)
        impressions = sum(r["impressions"] for r in grows)
        clicks = sum(r["clicks"] for r in grows)
        roas = safe_div(att_rev, spend)
        acos = safe_div(spend, att_rev) if platform == "Amazon" else None
        ctr = safe_div(clicks, impressions)
        ctr_display = f"{ctr * 100:.2f}%" if ctr is not None else None

        sources = sorted({r["source_file"] for r in grows})

        row = {
            "platform": platform,
            "adType": t1,
            "adSubType": t2,
            "targeting": kw,
            "campaignGroup": _group_label(platform, t1, t2, kw),
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
        bkey = BUCKET_KEYS.get(bucket)
        if bkey:
            detail[bkey].append(row)

    for key in detail:
        detail[key].sort(key=_channel_sort_key)

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


def _channel_label(r):
    if r["platform"] == "Amazon":
        return f"Amazon {r['ad_type']}" if r["ad_type"] else "Amazon"
    if r["platform"] == "Meta":
        return f"Meta {r['ad_sub_type']}" if r["ad_sub_type"] else "Meta"
    if r["platform"] == "Flipkart":
        return f"Flipkart {r['ad_type']}" if r["ad_type"] else "Flipkart"
    return r["platform"]


def agg_bucket_summary(rows, benchmarks):
    """Bucket-level cards on the Channel Detail view.

    Dashboard expects channels as a string array, plus CTR/CVR range/status/note
    fields. CVR is left as 'n/a' because purchase counts aren't loaded.
    """
    bucket_ctr_benchmark = {
        "Awareness": "prospectingCTR",
        "Core Sales": "prospectingCTR",
        "Retargeting": "retargetingCTR",
    }

    out = {}
    for bucket, key in BUCKET_KEYS.items():
        sub = [r for r in rows if r["bucket"] == bucket]
        spend = sum(r["spend"] for r in sub)
        att_rev = sum(r["att_rev"] for r in sub)
        roas = safe_div(att_rev, spend) or 0

        channels = sorted({_channel_label(r) for r in sub if r["spend"] > 0})

        ctrs = [r["ctr"] for r in sub if r.get("ctr") is not None and r["impressions"] > 0]
        total_imp = sum(r["impressions"] for r in sub)
        total_clicks = sum(r["clicks"] for r in sub)

        if ctrs and total_imp > 0:
            ctr_min, ctr_max = min(ctrs), max(ctrs)
            ctr_range = f"{ctr_min * 100:.2f}% – {ctr_max * 100:.2f}%"
            avg_ctr = total_clicks / total_imp
            bench = benchmarks.get(bucket_ctr_benchmark[bucket], 0)
            ctr_status = "ok" if avg_ctr >= bench else "fail"
            ctr_note = f"Avg {avg_ctr * 100:.2f}% vs bench {bench * 100:.1f}%"
        else:
            ctr_range = "—"
            ctr_status = "na"
            ctr_note = "No impressions"

        # CVR requires purchase counts, which aren't loaded today.
        cvr_display = "n/a"
        cvr_range = "n/a"
        cvr_status = "na"
        cvr_note = "Purchase data not loaded"

        out[key] = {
            "totalSpend": round(spend, 2),
            "roas": round(roas, 2),
            "channels": channels,
            "ctrRange": ctr_range,
            "ctrStatus": ctr_status,
            "ctrNote": ctr_note,
            "cvrDisplay": cvr_display,
            "cvrRange": cvr_range,
            "cvrStatus": cvr_status,
            "cvrNote": cvr_note,
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

    # Search term files — used to derive per-keyword-type spend/revenue ratios per campaign.
    ref = load_targeting_reference()
    sp_path = find_sp_search(folder)
    sb_path = find_sb_search(folder)
    sp_df = pd.read_excel(sp_path) if sp_path else None
    sb_df = pd.read_excel(sb_path) if sb_path else None
    if sp_path:
        source_files.append(sp_path.name)
    if sb_path:
        source_files.append(sb_path.name)
    keyword_ratios = compute_keyword_ratios(sp_df, sb_df, ref)

    # SD targeting (recorded in source files; not aggregated separately — SD totals come from the campaign report)
    sd_path = find_sd_targeting(folder)
    if sd_path:
        source_files.append(sd_path.name)

    all_rows = []

    amazon_path = find_amazon_campaigns(folder)
    if amazon_path:
        all_rows.extend(load_amazon_campaigns(amazon_path))
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

    if OUTPUT_FILE.exists():
        with open(OUTPUT_FILE, encoding="utf-8") as f:
            data = json.load(f)
    else:
        data = {}

    active_benchmarks = data.get("benchmarks") or DEFAULT_BENCHMARKS

    # Expand Amazon SP/SB/SBV campaigns into per-keyword-type rows for the channel detail
    # view only. Spend/revenue stay rebased to campaign-report totals via the search-term
    # ratios, so platform/bucket/monthly totals reconcile exactly. Top performers and
    # bucket summary still operate on the original campaign-level rows.
    detail_rows, fallback_campaigns = expand_amazon_by_keyword_type(all_rows, keyword_ratios)

    platforms = agg_platforms(all_rows)
    total_spend = sum(p["spend"] for p in platforms.values() if not p.get("pending"))
    total_att_rev = sum(p["attRev"] for p in platforms.values() if not p.get("pending"))
    channel_detail = agg_channel_detail(detail_rows)
    top, bottom = agg_top_bottom(all_rows)
    bucket_summary = agg_bucket_summary(all_rows, active_benchmarks)

    _print_reconciliation(all_rows, detail_rows, fallback_campaigns)

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


def _print_reconciliation(campaign_rows, detail_rows, fallback_campaigns):
    """Verify keyword-type-split rows reconcile to campaign-level totals per Amazon ad_type."""
    amazon_campaign = [r for r in campaign_rows if r["platform"] == "Amazon"]
    amazon_detail = [r for r in detail_rows if r["platform"] == "Amazon"]
    if not amazon_campaign:
        return

    print("\nReconciliation (campaign report ↔ channel detail):")
    by_type_campaign = {}
    by_type_detail = {}
    for r in amazon_campaign:
        by_type_campaign.setdefault(r["ad_type"], [0.0, 0.0])
        by_type_campaign[r["ad_type"]][0] += r["spend"]
        by_type_campaign[r["ad_type"]][1] += r["att_rev"]
    for r in amazon_detail:
        by_type_detail.setdefault(r["ad_type"], [0.0, 0.0])
        by_type_detail[r["ad_type"]][0] += r["spend"]
        by_type_detail[r["ad_type"]][1] += r["att_rev"]

    for ad_type in sorted(by_type_campaign, key=lambda t: AD_TYPE_ORDER.get(t, 99)):
        cs, cr = by_type_campaign[ad_type]
        ds, dr = by_type_detail.get(ad_type, [0, 0])
        print(
            f"  {ad_type:4s} spend ₹{cs:>12,.2f} → ₹{ds:>12,.2f} (Δ₹{ds - cs:+.2f})  "
            f"rev ₹{cr:>12,.2f} → ₹{dr:>12,.2f} (Δ₹{dr - cr:+.2f})"
        )

    if fallback_campaigns:
        print(f"\n  Fallback (100% Generic) applied to {len(fallback_campaigns)} campaign(s):")
        for campaign, am, ad_type in fallback_campaigns[:10]:
            print(f"    - [{ad_type} {am}] {campaign}")
        if len(fallback_campaigns) > 10:
            print(f"    ... and {len(fallback_campaigns) - 10} more")
    print()


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
