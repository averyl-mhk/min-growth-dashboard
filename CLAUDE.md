# MIN Growth Marketing Dashboard — Processor Script

## Your task

Write `processor.py` — a Python script that reads raw marketing export files for a given month and writes `data.json`, which the dashboard at `index.html` reads to render.

Run it like this:
```
python processor.py --month 2026_03
```

The script must be **fully deterministic** — every number in `data.json` must be traceable to a specific row in a specific source file. No hardcoded values, no filling in gaps from memory.

---

## File structure

```
growth-marketing-dashboard/       ← repo root (run script from here)
  processor.py                    ← the script you are writing
  data.json                       ← output (overwritten each run)
  index.html                      ← dashboard (do not touch)
  raw_exports/
    2026_02/                      ← drop files in here by month
    2026_03/
    2026_04/
  context/
    GMD_TargetingType_Reference.xlsx   ← keyword classification reference
```

---

## Input files (per month folder)

Files have inconsistent names. Use **case-insensitive glob** to find them.

| File | Glob pattern | Format |
|------|-------------|--------|
| Amazon campaigns | `*mazon*` but NOT `*Search*` and NOT `*Targeting*` | CSV or XLSX |
| SP search terms | `*Sponsored_Products*` or `*SP*Search*` | XLSX |
| SB search terms | `*Sponsored_Brands*` | XLSX |
| SD targeting | `*Sponsored_Display*` or `*SD*Target*` | XLSX |
| Meta campaigns | `*eta*` (not amazon/flipkart) | XLSX |
| Flipkart campaigns | `*lipkart*` | CSV or XLSX |
| POS (manual) | `*POS*` or `*pos*` | XLSX |

If a file is not found for a given month, skip it gracefully and note it in the output `_meta.missing` array. Never crash on a missing optional file.

---

## Confirmed column names (verified from actual files)

### Amazon campaign report (CSV or XLSX)
```
Campaign name | Type | Targeting | Portfolio name | Impressions | Clicks | CTR |
Total cost (converted) | CPC (converted) | Purchases | Sales (converted) | ACOS | ROAS
```
**IMPORTANT:** The `(converted)` suffix was dropped in some exports. Handle both:
- Spend: `Total cost (converted)` OR `Total cost`
- Revenue: `Sales (converted)` OR `Sales`

Column meanings:
- `Type` = `SP` / `SB` / `SD`
- `Targeting` = `AUTOMATIC` / `MANUAL`
- `ACOS` = **decimal** (0.394 = 39.4%) — do NOT divide by 100

### SP search term report (XLSX)
```
Start Date | End Date | Portfolio name | Currency | Campaign Name | Ad Group Name |
Country | Targeting | Match Type | Customer Search Term | Impressions | Clicks |
Click-Through Rate (CTR) | Cost Per Click (CPC) | Spend | 7 Day Total Sales (₹) |
Total Advertising Cost of Sales (ACOS) | Total Return on Advertising Spend (ROAS) | ...
```
Key columns: `Campaign Name`, `Customer Search Term`, `Spend`, `7 Day Total Sales (₹)`

### SB search term report (XLSX)
```
Start Date | End Date | Portfolio name | Currency | Campaign Name | Ad Group Name |
Targeting | Match Type | Customer Search Term | Cost Type | Impressions | Viewable Impressions |
Clicks | Click-Through Rate (CTR) | Spend | ... | 14 Day Total Sales (₹) | ...
```
Key columns: `Campaign Name`, `Customer Search Term`, `Spend`, `14 Day Total Sales (₹)`

### SD targeting report (XLSX)
```
Start Date | End Date | Currency | Campaign Name | Portfolio name | Cost Type |
Ad Group Name | Targeting | Bid Optimisation | Impressions | Viewable Impressions |
Clicks | Click-Through Rate (CTR) | 14 Day Detail Page Views (DPV) | Spend | ... |
Total Advertising Cost of Sales (ACOS) | Total Return on Advertising Spend (ROAS) |
14 Day Total Orders (#) | 14 Day Total Units (#) | 14 Day Total Sales (₹) | ...
```
Key columns: `Campaign Name`, `Spend`, `14 Day Total Sales (₹)`, `Total Advertising Cost of Sales (ACOS)`

### Meta campaign report (XLSX)
**CRITICAL:** Row 0 contains `Added, Added, Added, ...` — **skip it**. Row 1 is the real header.

```
Targeting Type | Keyword Type | Advantage Plus | Campaign name | Ad set name |
Objective | Media type | ... | Amount spent (INR) | Impressions | Reach | ... |
CTR (link click-through rate) | CPC (cost per link click) | ... | Purchases |
Revenue | Purchase ROAS (return on ad spend) | Link clicks | ...
```
Key columns: `Targeting Type`, `Keyword Type`, `Amount spent (INR)`, `Revenue`, `Impressions`, `CTR (link click-through rate)`

- `Targeting Type`: `Carousel` / `Instream` / `Catalogue` / `Media Mix`
- `Keyword Type`: `Prospecting` / `Remarketing` / `Sales Traffic`
- `Revenue`: may be `None` for awareness campaigns (no purchase attribution) — treat as 0

### Flipkart campaign report (CSV or XLSX)
**CRITICAL header row varies by export:**
- Some files have a junk row 0 (all `None` / `Calculated`) — **detect and skip it** (if row 0 column 0 is None or "Calculated", skip to row 1 for headers)
- Other files have headers on row 0 directly

```
Campaign Name | campaign_type | Ad Spend | Views | SUM(clicks) |
Total converted units | Total Revenue (Rs.) | ROI | Click Through Rate | Conversion Rate
```
- `campaign_type` values vary: `PLA`, `SP`, `SELLER_PCA`, `PCA` — map to buckets (see below)
- `ROI` and `ACOS` columns may contain formula strings (`=+C4/G4`) or `-` — treat as null

### POS file (XLSX)
Simple two-column table. Header row has `Name` and a sales column (label varies, e.g. `Total Sale for Feb`).
```
Name       | Total Sale for Feb
Shopify    | 8,53,000
Amazon     | 64,26,000
Flipkart   | 2,74,000
```
- Platform names may have trailing spaces — strip them
- Sales values are **Indian-format strings with commas** (e.g. `8,53,000` = 853000) — remove all commas and convert to float

### Keyword classification reference (`context/GMD_TargetingType_Reference.xlsx`)
Three columns, **no header handling needed** — row 0 is the header:
```
Competition          | Branded           | Generic
Stahl                | Myer              | Rest all the Searches will be considered as Generic
Milton               | Mayer             |
Wonderchef           | Meyer             |
agaro                | circulon          |
bergner              |                   |
... (62 competition terms total)
```
- Column 0 = competitor brand keywords
- Column 1 = Meyer's own brand variants (Myer, Mayer, Meyer, circulon)
- Column 2 = rule only, no list needed
- Classification is **case-insensitive substring match** on the search term:
  1. If term contains any Branded keyword → `"Branded"`
  2. Else if term contains any Competition keyword → `"Competition"`
  3. Else → `"Generic"`

---

## Business logic

### Campaign bucketing

| Platform | Signal | Bucket |
|----------|--------|--------|
| Amazon | `Type` = `SP` | `Core Sales` |
| Amazon | `Type` = `SB` | `Awareness` |
| Amazon | `Type` = `SD` | `Retargeting` |
| Meta | `Keyword Type` = `Prospecting` | `Awareness` |
| Meta | `Keyword Type` = `Remarketing` | `Retargeting` |
| Meta | `Keyword Type` = `Sales Traffic` | `Core Sales` |
| Flipkart | `campaign_type` in (`PLA`, `SP`) | `Core Sales` |
| Flipkart | `campaign_type` in (`SELLER_PCA`, `PCA`) | `Retargeting` |

### Keyword type classification (Amazon only)

Use the search term reports to determine the dominant keyword type for each campaign:
1. For every row in the SP/SB search term reports, classify `Customer Search Term` as `Branded` / `Competition` / `Generic` using the reference file
2. Group by `Campaign Name`, sum `Spend` per keyword type
3. The type with the highest spend share becomes that campaign's `keyword_type`
4. If search term data is unavailable for a month, set `keyword_type` to `null`

### Channel detail aggregation

Group campaign rows for the channel detail view. Do NOT show one row per individual campaign — aggregate into groups:

**Amazon:** group by (`Type`, `Targeting`) → e.g. `SP + AUTOMATIC`, `SP + MANUAL`, `SB + MANUAL`, `SD + (various)`
**Meta:** group by (`Targeting Type`, `Keyword Type`) → e.g. `Carousel + Prospecting`, `Media Mix + Remarketing`
**Flipkart:** group by `campaign_type`

For each group compute: total spend, total attributed revenue, ROAS (revenue/spend), ACOS (spend/revenue for Amazon groups only), total impressions, total clicks, CTR (clicks/impressions).

### Top performers / Needs attention

- **Top performers:** campaigns ranked by ROAS descending, minimum spend ₹5,000, any bucket
- **Needs attention:** campaigns ranked by ROAS ascending, minimum spend ₹10,000, exclude `Awareness` bucket (low ROAS is expected there)
- Use individual campaign rows (not aggregated groups) for these tables
- Limit to 5 entries each

---

## Output: `data.json` schema

Write to `data.json` at the **repo root** (not in `outputs/`). The file accumulates months — running the script for March adds March without removing February.

```json
{
  "_meta": {
    "dashboard": "MIN Growth Marketing Dashboard",
    "marketingArm": "Meyer India (MIN)",
    "lastUpdated": "2026_03",
    "currency": "INR",
    "generatedBy": "processor.py",
    "sourceFiles": ["Campaign Amazon Mar (1).csv", "Flipkart (1).csv", "..."],
    "missing": ["Meta_Campaigns", "POS"]
  },

  "months": ["Feb '26", "Mar '26"],

  "notices": {
    "Feb '26": "February 2026 — ...",
    "Mar '26": "March 2026 — Generated from raw exports on 2026-05-07. Missing: Meta, POS."
  },

  "monthly": {
    "Mar '26": {
      "spend": 1234567.89,
      "attRev": 9876543.21,
      "pos": null,
      "platforms": {
        "Amazon AMS": { "spend": 1000000, "attRev": 8000000, "acos": 0.125, "pending": false },
        "Meta":       { "spend": 0, "attRev": 0, "acos": null, "pending": true },
        "Google Ads": { "spend": 0, "attRev": 0, "acos": null, "pending": true },
        "Flipkart":   { "spend": 234567, "attRev": 1876543, "acos": null, "pending": false }
      }
    }
  },

  "benchmarks": {
    "prospectingCTR": 0.015,
    "retargetingCTR": 0.02,
    "salesCVR": 0.008,
    "amazonACOSTarget": 0.25,
    "roasHigh": 4,
    "roasMid": 2,
    "acosGood": 0.25,
    "acosMid": 0.50
  },

  "_benchmarks_note": "Edit these directly in data.json to change thresholds. The processor NEVER overwrites an existing benchmarks block — it only writes defaults if the key is absent.",

  "channelDetail": {
    "Mar '26": {
      "awareness": [
        {
          "platform": "Amazon",
          "adType": "SB",
          "adSubType": "Manual",
          "targeting": null,
          "campaignGroup": "SB Manual",
          "spend": 168791,
          "attRev": 479284,
          "roas": 2.84,
          "acos": 0.352,
          "impressions": 500000,
          "clicks": 3600,
          "ctr": 0.0072,
          "ctrDisplay": "0.72%",
          "ctrStatus": "na",
          "ctrCvr": "0.72% (raw avg)",
          "source": "Campaign Amazon Mar (1).csv"
        }
      ],
      "awarenessSubtotal": {
        "spend": 168791,
        "attRev": 479284,
        "roas": 2.84,
        "acos": 0.352,
        "note": ""
      },
      "coreSales": [ ... ],
      "coreSalesSubtotal": { ... },
      "retargeting": [ ... ],
      "retargetingSubtotal": { ... }
    }
  },

  "topPerformers": {
    "Mar '26": [
      {
        "rank": 1,
        "platform": "Amazon",
        "campaignName": "SP | Cast Iron All | Auto",
        "campaignFull": "SP | Cast Iron All | Auto",
        "bucket": "Core Sales",
        "spend": 6976,
        "roas": 9.00,
        "acos": 0.111,
        "source": "Campaign Amazon Mar (1).csv"
      }
    ]
  },

  "needsAttention": {
    "Mar '26": [ ... ]
  },

  "bucketSummary": {
    "Mar '26": {
      "awareness":   { "totalSpend": 0, "roas": 0, "channels": [] },
      "coreSales":   { "totalSpend": 0, "roas": 0, "channels": [] },
      "retargeting": { "totalSpend": 0, "roas": 0, "channels": [] }
    }
  }
}
```

**Month label format:** convert `2026_03` → `Mar '26` using `datetime(2026, 3, 1).strftime("%b '%y")`

**Preserving existing months:** load `data.json` before writing, merge the new month in, write back. Never delete months already in the file.

**Preserving benchmarks:** if `data.json` already contains a `benchmarks` key, keep it exactly as-is. Only write the default benchmarks if the key does not exist. This lets the user edit benchmarks directly in `data.json` without the processor overwriting their changes.

---

## Edge cases to handle

1. **ACOS as decimal vs percent:** Amazon exports use decimal (0.394). If any value > 1, divide by 100.
2. **Formula strings in Flipkart:** `ROI` column may contain `=+C4/G4` or `-` — catch with `try/except float()`, set to `null`.
3. **Indian number format in POS:** `8,53,000` → strip all commas → `853000`. Do not use `locale`.
4. **Meta Revenue = None for awareness:** expected — set attributed revenue to 0.
5. **Column name variants:** always try both `Total cost (converted)` and `Total cost` for Amazon spend. Whichever exists, use it.
6. **Flipkart junk header row:** if the first row's first cell is `None` or `"Calculated"`, treat row 1 as the header instead.
7. **Missing files:** if a file type is not found, set that platform's data to `pending: true` in the output and add the platform name to `_meta.missing`. Do not crash.
8. **Duplicate column names in Meta:** the Meta export has `Ad set name` twice — pandas will rename the second to `Ad set name.1`. Ignore the duplicate.
9. **Search term files not available:** if the search term files are missing for a month, set `keyword_type: null` on all channel detail rows for that month. Do not fail.
10. **ACOS for non-Amazon platforms:** only compute ACOS for Amazon rows. Set to `null` for Meta and Flipkart in all output objects.

---

## Dependencies

```
pandas
openpyxl
```

Install with: `pip install pandas openpyxl`

---

## Verification output

After writing `data.json`, print a summary to stdout:
```
✓ Mar '26 written to data.json
  Platforms:    Amazon ₹X.XL spend | Flipkart ₹X.XL spend | Meta MISSING
  Total spend:  ₹X,XX,XXX
  Total attRev: ₹X,XX,XXX
  ROAS:         X.XX×
  POS:          MISSING
  Missing:      Meta_Campaigns, POS
  Source files: [list of filenames used]
```
