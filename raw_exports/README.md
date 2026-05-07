# Raw Exports — Drop Files Here

This folder feeds the data processing script (`GMD_DataProcessor_v1.py`).
Drop Akash's monthly exports into a subfolder named `YYYY_MM` (e.g. `2026_03` for March).

---

## Folder structure

```
raw_exports/
  2026_02/                          ← already processed (Feb baseline)
  2026_03/                          ← drop March files here
    Amazon_SP_Campaigns.xlsx        Amazon Sponsored Products campaign-level report
    Amazon_SB_Campaigns.xlsx        Amazon Sponsored Brands campaign-level report
    Amazon_SD_Campaigns.xlsx        Amazon Sponsored Display campaign-level report
    Amazon_SP_SearchTerms.xlsx      Amazon SP search term report (for keyword classification)
    Amazon_SB_SearchTerms.xlsx      Amazon SB search term report (for keyword classification)
    Meta_Campaigns.xlsx             Meta Ads Manager campaign-level export
    Flipkart_Campaigns.xlsx         Flipkart Seller Hub campaign-level export
    POS_Manual.xlsx                 Point-of-sale totals by channel (Amazon / Shopify / Flipkart)
```

---

## File notes

**Amazon campaign exports** — export at campaign level from Amazon AMS console.
Expected columns (exact names may vary slightly): Campaign Name, Ad Type, Targeting Type,
Impressions, Clicks, Spend, Sales (14-day), Orders, ACOS.

**Amazon search term reports** — needed to classify keywords as Branded / Competition / Generic.
Reference list is in `context/GMD_TargetingType_Reference.xlsx`.

**Meta export** — campaign-level from Meta Ads Manager.
Expected columns: Campaign Name, Objective, Amount Spent, Purchases Value, Impressions, Clicks, CTR.

**Flipkart export** — campaign-level from Flipkart Seller Hub.
Expected columns: Campaign Name, Type, Targeting, Spend, Revenue, Impressions, Clicks.

**POS_Manual.xlsx** — a simple two-column table: Platform | Total Sales (₹).
Platforms: Amazon, Shopify, Flipkart. Fill in manually each month from your sales dashboards.

---

## Running the processor

```bash
python GMD_DataProcessor_v1.py --month 2026_03
```

Output: `outputs/GMD_Dashboard_MIN_data.json` (overwrites previous version)
Then open `outputs/GMD_Dashboard_MIN_v4.html` to see the updated dashboard.
