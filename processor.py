"""processor.py — MIN Growth Marketing Dashboard data processor.

Reads monthly raw exports from raw_exports/YYYY_MM/ and writes data.json.
Run: python processor.py --month 2026_03
"""

import argparse
import json
import sys
from collections import Counter, defaultdict
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
    """Safe float. Handles currency-formatted strings (₹1,75,216.16), NaN, '-', formula strings."""
    if val is None:
        return None
    try:
        if pd.isna(val):
            return None
    except (TypeError, ValueError):
        pass
    if isinstance(val, str):
        v = val.strip()
        if not v or v == "-" or v.startswith("="):
            return None
        # Keep digits, decimal point, and a leading minus. Strip anything else
        # (currency symbols, thousands separators, trailing whitespace, % signs).
        negative = v.lstrip().startswith("-")
        cleaned = "".join(c for c in v if c.isdigit() or c == ".")
        if not cleaned:
            return None
        try:
            f = float(cleaned)
            return -f if negative else f
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
            fallback.append((campaign, am, r["ad_type"], r["spend"]))
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
    # Some Meta exports have an "Added, Added, ..." junk row 0 with the real header on row 1;
    # other exports have the real header on row 0. Detect by checking row 0 col 0.
    peek = pd.read_excel(path, header=None, nrows=1)
    first = peek.iloc[0, 0]
    is_junk = isinstance(first, str) and first.strip().lower() == "added"
    df = pd.read_excel(path, header=1 if is_junk else 0)
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

def process_month(month_str: str, export_brief_flag: bool = False):
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

    if export_brief_flag:
        export_brief(month_str, month_label, data, all_rows, detail_rows, pos, missing, source_files, top, bottom)


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
        nonzero = [c for c in fallback_campaigns if c[3] > 0]
        total_fb_spend = sum(c[3] for c in fallback_campaigns)
        print(
            f"\n  Fallback (100% Generic) applied to {len(fallback_campaigns)} campaign(s), "
            f"total spend ₹{total_fb_spend:,.0f}"
        )
        if not nonzero:
            print("    (all fallback campaigns have ₹0 spend — paused/inactive, no impact on totals)")
        else:
            print(f"    {len(nonzero)} fallback campaign(s) with non-zero spend:")
            for campaign, am, ad_type, spend in sorted(nonzero, key=lambda c: -c[3])[:10]:
                print(f"      - ₹{spend:>10,.0f} [{ad_type} {am}] {campaign}")
            if len(nonzero) > 10:
                print(f"      ... and {len(nonzero) - 10} more")
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


# ─── AI brief export ──────────────────────────────────────────────────────────

def export_brief(month_str, month_label, data, all_rows, detail_rows, pos, missing, source_files, top, bottom):
    """Write a formatted text brief to ai_brief_YYYY_MM.txt for pasting into Claude."""

    lines = []

    def h(title):
        lines.append(f"\n{'=' * 60}")
        lines.append(f"  {title}")
        lines.append(f"{'=' * 60}")

    def sub(title):
        lines.append(f"\n  [{title}]")

    # ── Header ──────────────────────────────────────────────────────────────────
    lines.append("MIN GROWTH MARKETING DASHBOARD — AI ANALYSIS BRIEF")
    lines.append(f"Month:     {month_label}")
    lines.append(f"Generated: {datetime.today().strftime('%Y-%m-%d')}")
    lines.append(f"Files:     {', '.join(source_files)}")
    if missing:
        lines.append(f"MISSING:   {', '.join(missing)}")
    lines.append("")
    lines.append("The system context (brand background, guiding principles, benchmarks)")
    lines.append("is saved in this Project's instructions — do not re-paste it.")
    lines.append("Analyse the data below and return recommendations as JSON (format at end).")

    # ── POS ─────────────────────────────────────────────────────────────────────
    h("POS — ACTUAL PLATFORM SALES")
    if pos:
        for k, v in pos.items():
            if k != "total":
                lines.append(f"  {k:<22} ₹{v:>12,.0f}")
        lines.append(f"  {'TOTAL':<22} ₹{pos.get('total', 0):>12,.0f}")
    else:
        lines.append("  POS data not available this month.")

    # ── MER ─────────────────────────────────────────────────────────────────────
    h("MARKETING EFFICIENCY RATIO (MER = Total POS ÷ Total Spend)")
    total_spend = sum(r["spend"] for r in all_rows)
    total_att_rev = sum(r["att_rev"] for r in all_rows)

    prior_months = [m for m in data.get("months", []) if m != month_label]
    prior_label = prior_months[-1] if prior_months else None
    prior_monthly = data.get("monthly", {}).get(prior_label) if prior_label else None

    if pos and pos.get("total") and total_spend > 0:
        mer = pos["total"] / total_spend
        lines.append(f"  Total POS:    ₹{pos['total']:>12,.0f}")
        lines.append(f"  Total Spend:  ₹{total_spend:>12,.0f}")
        lines.append(f"  MER:          {mer:.2f}×")
        if prior_monthly:
            prior_pos_block = prior_monthly.get("pos") or {}
            prior_total_pos = prior_pos_block.get("total") if prior_pos_block else None
            prior_sp = prior_monthly.get("spend", 0)
            if prior_total_pos and prior_sp > 0:
                prior_mer = prior_total_pos / prior_sp
                lines.append(f"  Prior MER ({prior_label}): {prior_mer:.2f}×  (Δ {mer - prior_mer:+.2f}×)")
    else:
        roas_proxy = total_att_rev / total_spend if total_spend > 0 else None
        lines.append("  POS not available — using attributed revenue as proxy.")
        lines.append(f"  Total attRev: ₹{total_att_rev:>12,.0f}")
        lines.append(f"  Total Spend:  ₹{total_spend:>12,.0f}")
        lines.append(f"  Proxy MER:    {roas_proxy:.2f}×" if roas_proxy else "  Proxy MER:   N/A")

    if pos and pos.get("total"):
        shopify_val = next((v for k, v in pos.items() if "shopify" in k.lower()), None)
        if shopify_val:
            pct = shopify_val / pos["total"] * 100
            lines.append(f"  Shopify % of POS: {pct:.1f}%")
            if prior_monthly and prior_monthly.get("pos"):
                prior_pos_b = prior_monthly["pos"] or {}
                prior_sh = next((v for k, v in prior_pos_b.items() if "shopify" in k.lower()), None)
                prior_tot = prior_pos_b.get("total")
                if prior_sh and prior_tot:
                    prior_pct = prior_sh / prior_tot * 100
                    lines.append(f"  Prior ({prior_label}) Shopify %: {prior_pct:.1f}%  (Δ {pct - prior_pct:+.1f}pp)")

    # ── Platform summary ────────────────────────────────────────────────────────
    h("PLATFORM SUMMARY")
    lines.append(f"  {'Platform':<15} {'Spend':>12} {'AttRev':>12} {'ROAS':>8} {'ACOS':>8}")
    lines.append(f"  {'-'*15} {'-'*12} {'-'*12} {'-'*8} {'-'*8}")
    for plat, label in [("Amazon", "Amazon AMS"), ("Meta", "Meta"), ("Flipkart", "Flipkart")]:
        sub_rows = [r for r in all_rows if r["platform"] == plat]
        if not sub_rows:
            lines.append(f"  {label:<15} {'MISSING':>12}")
            continue
        sp = sum(r["spend"] for r in sub_rows)
        rev = sum(r["att_rev"] for r in sub_rows)
        roas = rev / sp if sp > 0 else 0
        acos = sp / rev if (plat == "Amazon" and rev > 0) else None
        lines.append(
            f"  {label:<15} ₹{sp:>10,.0f} ₹{rev:>10,.0f} {roas:>7.2f}× "
            f"{'—':>8}" if acos is None else
            f"  {label:<15} ₹{sp:>10,.0f} ₹{rev:>10,.0f} {roas:>7.2f}× {acos*100:>7.1f}%"
        )
    lines.append(f"  {'-'*15} {'-'*12} {'-'*12} {'-'*8} {'-'*8}")
    overall_roas = total_att_rev / total_spend if total_spend > 0 else 0
    lines.append(f"  {'TOTAL':<15} ₹{total_spend:>10,.0f} ₹{total_att_rev:>10,.0f} {overall_roas:>7.2f}×")

    if prior_label and prior_monthly:
        prior_sp = prior_monthly.get("spend", 0)
        prior_rev = prior_monthly.get("attRev", 0)
        prior_roas = prior_rev / prior_sp if prior_sp > 0 else 0
        sp_delta = (total_spend - prior_sp) / prior_sp * 100 if prior_sp > 0 else None
        rev_delta = (total_att_rev - prior_rev) / prior_rev * 100 if prior_rev > 0 else None
        lines.append(f"\n  MoM vs {prior_label}:")
        lines.append(f"    Spend:  ₹{prior_sp:>10,.0f} → ₹{total_spend:>10,.0f}  ({sp_delta:+.1f}%)" if sp_delta is not None else f"    Spend: N/A")
        lines.append(f"    AttRev: ₹{prior_rev:>10,.0f} → ₹{total_att_rev:>10,.0f}  ({rev_delta:+.1f}%)" if rev_delta is not None else f"    AttRev: N/A")
        lines.append(f"    ROAS:   {prior_roas:.2f}× → {overall_roas:.2f}×  (Δ {overall_roas - prior_roas:+.2f}×)")

    # ── Bucket summary ──────────────────────────────────────────────────────────
    h("BUCKET SUMMARY")
    lines.append(f"  {'Bucket':<14} {'Spend':>12} {'AttRev':>12} {'ROAS':>8} {'AmazonACOS':>12}  Channels")
    lines.append(f"  {'-'*14} {'-'*12} {'-'*12} {'-'*8} {'-'*12}  {'-'*30}")
    for bucket in ["Awareness", "Core Sales", "Retargeting"]:
        sub_rows = [r for r in all_rows if r["bucket"] == bucket]
        sp = sum(r["spend"] for r in sub_rows)
        rev = sum(r["att_rev"] for r in sub_rows)
        roas = rev / sp if sp > 0 else 0
        amz = [r for r in sub_rows if r["platform"] == "Amazon"]
        amz_sp = sum(r["spend"] for r in amz)
        amz_rev = sum(r["att_rev"] for r in amz)
        acos_str = f"{amz_sp/amz_rev*100:.1f}%" if amz_rev > 0 else "—"
        channels = ", ".join(sorted({r["platform"] for r in sub_rows if r["spend"] > 0}))
        lines.append(f"  {bucket:<14} ₹{sp:>10,.0f} ₹{rev:>10,.0f} {roas:>7.2f}× {acos_str:>12}  {channels}")

    # ── Amazon keyword type breakdown ───────────────────────────────────────────
    h("AMAZON KEYWORD TYPE BREAKDOWN (from search term data)")
    amz_detail = [r for r in detail_rows if r["platform"] == "Amazon" and r.get("keyword_type")]
    if amz_detail:
        total_amz_sp = sum(r["spend"] for r in amz_detail)
        kw_agg = defaultdict(lambda: {"spend": 0.0, "rev": 0.0, "types": set()})
        for r in amz_detail:
            kw = r["keyword_type"]
            kw_agg[kw]["spend"] += r["spend"]
            kw_agg[kw]["rev"] += r["att_rev"]
            kw_agg[kw]["types"].add(r["ad_type"])
        lines.append(f"  {'KwType':<14} {'AdTypes':<10} {'Spend':>12} {'AttRev':>12} {'ROAS':>8} {'Spend%':>8}")
        lines.append(f"  {'-'*14} {'-'*10} {'-'*12} {'-'*12} {'-'*8} {'-'*8}")
        for kw in ["Branded", "Competition", "Generic"]:
            if kw not in kw_agg:
                continue
            d = kw_agg[kw]
            sp = d["spend"]
            rev = d["rev"]
            roas = rev / sp if sp > 0 else 0
            pct = sp / total_amz_sp * 100 if total_amz_sp > 0 else 0
            types = ",".join(sorted(d["types"]))
            lines.append(f"  {kw:<14} {types:<10} ₹{sp:>10,.0f} ₹{rev:>10,.0f} {roas:>7.2f}× {pct:>7.1f}%")
    else:
        lines.append("  Search term files not available for this month — keyword type breakdown not possible.")

    # ── Channel detail ──────────────────────────────────────────────────────────
    h("CHANNEL DETAIL — ALL CAMPAIGN GROUPS")
    hdr = f"  {'Platform':<10} {'AdType':<6} {'Sub':<8} {'KwType':<14} {'Spend':>10} {'AttRev':>10} {'ROAS':>7} {'ACOS':>8} {'Impressions':>12} {'Clicks':>8} {'CTR':>7}"
    sep = f"  {'-'*10} {'-'*6} {'-'*8} {'-'*14} {'-'*10} {'-'*10} {'-'*7} {'-'*8} {'-'*12} {'-'*8} {'-'*7}"
    lines.append(hdr)
    lines.append(sep)

    for bucket in ["Awareness", "Core Sales", "Retargeting"]:
        bucket_rows = [r for r in detail_rows if r["bucket"] == bucket]
        groups = {}
        for r in bucket_rows:
            k = _group_key(r)
            if k:
                groups.setdefault(k, []).append(r)
        lines.append(f"\n  [{bucket.upper()}]")
        sorted_groups = sorted(
            groups.items(),
            key=lambda x: _channel_sort_key({
                "adType": x[0][1], "adSubType": x[0][2],
                "targeting": x[0][3], "spend": sum(r["spend"] for r in x[1])
            })
        )
        for k, grows in sorted_groups:
            platform, t1, t2, kw = k
            sp = sum(r["spend"] for r in grows)
            rev = sum(r["att_rev"] for r in grows)
            imp = sum(r["impressions"] for r in grows)
            clk = sum(r["clicks"] for r in grows)
            roas = rev / sp if sp > 0 else 0
            acos = sp / rev if (platform == "Amazon" and rev > 0) else None
            ctr = clk / imp if imp > 0 else None
            lines.append(
                f"  {platform:<10} {t1 or '—':<6} {t2 or '—':<8} {kw or '—':<14} "
                f"₹{sp:>8,.0f} ₹{rev:>8,.0f} {roas:>6.2f}× "
                f"{'—' if acos is None else f'{acos*100:.1f}%':>8} "
                f"{int(imp):>12,} {int(clk):>8,} "
                f"{'—' if ctr is None else f'{ctr*100:.2f}%':>7}"
            )

    # ── Top performers ──────────────────────────────────────────────────────────
    h("TOP PERFORMERS (min ₹5,000 spend, ROAS descending)")
    if top:
        for r in top:
            acos_str = f"{r['acos']*100:.1f}%" if r.get("acos") is not None else "—"
            lines.append(f"  #{r['rank']} [{r['platform']}] [{r['bucket']}] ROAS {r['roas']:.2f}× | ACOS {acos_str} | Spend ₹{r['spend']:,.0f}")
            lines.append(f"     {r['campaignName']}")
    else:
        lines.append("  No campaigns meet the minimum spend threshold.")

    # ── Needs attention ─────────────────────────────────────────────────────────
    h("NEEDS ATTENTION (min ₹10,000 spend, non-Awareness, ROAS ascending)")
    if bottom:
        for r in bottom:
            acos_str = f"{r['acos']*100:.1f}%" if r.get("acos") is not None else "—"
            lines.append(f"  #{r['rank']} [{r['platform']}] [{r['bucket']}] ROAS {r['roas']:.2f}× | ACOS {acos_str} | Spend ₹{r['spend']:,.0f}")
            lines.append(f"     {r['campaignName']}")
    else:
        lines.append("  No campaigns meet the minimum spend threshold.")

    # ── All Amazon campaigns ────────────────────────────────────────────────────
    h("ALL AMAZON CAMPAIGNS (spend > ₹0, sorted by spend desc)")
    amz_rows = sorted([r for r in all_rows if r["platform"] == "Amazon"], key=lambda r: -r["spend"])
    lines.append(f"  {'AdType':<6} {'Sub':<8} {'Spend':>10} {'AttRev':>10} {'ROAS':>7} {'ACOS':>8} {'Impr':>10} {'Clicks':>7} {'CTR':>7}")
    lines.append(f"  {'-'*6} {'-'*8} {'-'*10} {'-'*10} {'-'*7} {'-'*8} {'-'*10} {'-'*7} {'-'*7}")
    for r in amz_rows:
        sp, rev = r["spend"], r["att_rev"]
        roas = rev / sp if sp > 0 else 0
        acos = r["acos"]
        ctr = r.get("ctr")
        lines.append(
            f"  {r['ad_type']:<6} {r.get('ad_sub_type') or '—':<8} "
            f"₹{sp:>8,.0f} ₹{rev:>8,.0f} {roas:>6.2f}× "
            f"{'—' if acos is None else f'{acos*100:.1f}%':>8} "
            f"{int(r['impressions']):>10,} {int(r['clicks']):>7,} "
            f"{'—' if ctr is None else f'{ctr*100:.2f}%':>7}"
        )
        lines.append(f"     {r['campaign_name']}")

    # ── All Meta campaigns ──────────────────────────────────────────────────────
    meta_rows = sorted([r for r in all_rows if r["platform"] == "Meta"], key=lambda r: -r["spend"])
    if meta_rows:
        h("ALL META CAMPAIGNS")
        lines.append(f"  {'TargetingType':<16} {'KeywordType':<16} {'Spend':>10} {'AttRev':>10} {'ROAS':>7} {'Impr':>10} {'Clicks':>7} {'CTR':>7}")
        lines.append(f"  {'-'*16} {'-'*16} {'-'*10} {'-'*10} {'-'*7} {'-'*10} {'-'*7} {'-'*7}")
        for r in meta_rows:
            sp, rev = r["spend"], r["att_rev"]
            roas = rev / sp if sp > 0 else 0
            ctr = r.get("ctr")
            lines.append(
                f"  {r.get('ad_type') or '—':<16} {r.get('ad_sub_type') or '—':<16} "
                f"₹{sp:>8,.0f} ₹{rev:>8,.0f} {roas:>6.2f}× "
                f"{int(r['impressions']):>10,} {int(r['clicks']):>7,} "
                f"{'—' if ctr is None else f'{ctr*100:.2f}%':>7}"
            )
            lines.append(f"     {r['campaign_name']}")

    # ── All Flipkart campaigns ──────────────────────────────────────────────────
    fk_rows = sorted([r for r in all_rows if r["platform"] == "Flipkart"], key=lambda r: -r["spend"])
    if fk_rows:
        h("ALL FLIPKART CAMPAIGNS")
        lines.append(f"  {'Type':<12} {'Bucket':<14} {'Spend':>10} {'AttRev':>10} {'ROAS':>7} {'Impr':>10} {'Clicks':>7} {'CTR':>7}")
        lines.append(f"  {'-'*12} {'-'*14} {'-'*10} {'-'*10} {'-'*7} {'-'*10} {'-'*7} {'-'*7}")
        for r in fk_rows:
            sp, rev = r["spend"], r["att_rev"]
            roas = rev / sp if sp > 0 else 0
            ctr = r.get("ctr")
            lines.append(
                f"  {r.get('ad_type') or '—':<12} {r['bucket']:<14} "
                f"₹{sp:>8,.0f} ₹{rev:>8,.0f} {roas:>6.2f}× "
                f"{int(r['impressions']):>10,} {int(r['clicks']):>7,} "
                f"{'—' if ctr is None else f'{ctr*100:.2f}%':>7}"
            )
            lines.append(f"     {r['campaign_name']}")

    # ── Output format instructions ──────────────────────────────────────────────
    h("YOUR OUTPUT — JSON ONLY, NO PROSE BEFORE OR AFTER")
    lines.append("""
Return ONLY a valid JSON object matching this schema exactly.

{
  "generatedDate": "YYYY-MM-DD",
  "month": "<month label, e.g. Mar '26>",
  "summary": "<2-3 sentences: what happened, overall health, most important signal>",
  "counts": {
    "critical": <int>,
    "warnings": <int>,
    "opportunities": <int>,
    "insights": <int>
  },
  "recommendations": [
    {
      "n": <int 1–12>,
      "priority": "<critical | warning | opportunity | insight>",
      "title": "<one-line title naming the specific campaign or issue>",
      "platform": "<Amazon | Meta | Flipkart | All>",
      "bucket": "<Core Sales | Awareness | Retargeting | null>",
      "whatDataShows": "<specific numbers from the data — cite exact campaign name and figures>",
      "whyItMatters": "<brand equity / MER / organic share frame — not just this month's ROAS>",
      "exactAction": "<direct instruction: which campaign, what change, by how much, what to do first>",
      "watchNextMonth": "<what number should move, in which direction, by roughly how much>",
      "keyMetrics": [
        {"label": "<metric name>", "value": "<formatted value e.g. 84% or 0.6x or Rs 41,200>"}
      ],
      "estimatedImpact": "<optional one-liner e.g. Est. savings Rs 12-15K/mo — omit if not quantifiable>"
    }
  ],
  "dataQualityFlags": [
    "<any reporting artefact, zero-revenue anomaly, column mismatch, or attribution gap>"
  ],
  "strategicGap": "<one paragraph: the thing not visible in this month's data but implied by trends>"
}

After generating, save Claude's response to a .json file (e.g. recs_2026_03.json),
then run: python ingest_recs.py --month 2026_03 --file recs_2026_03.json
""")

    brief_path = SCRIPT_DIR / f"ai_brief_{month_str}.txt"
    with open(brief_path, "w", encoding="utf-8") as f:
        f.write("\n".join(lines))

    print(f"\n✓ AI brief written to ai_brief_{month_str}.txt")
    print(f"  1. Open your MIN Dashboard Claude Project")
    print(f"  2. Start a new chat and paste the full contents of that file")
    print(f"  3. Save Claude's JSON response to e.g. recs_{month_str}.json")
    print(f"  4. Run: python ingest_recs.py --month {month_str} --file recs_{month_str}.json")
    return brief_path


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="MIN dashboard data processor")
    parser.add_argument("--month", required=True, help="Month folder name, e.g. 2026_03")
    parser.add_argument(
        "--export-brief",
        action="store_true",
        help="After processing, write ai_brief_YYYY_MM.txt for pasting into Claude"
    )
    args = parser.parse_args()
    process_month(args.month, export_brief_flag=args.export_brief)
