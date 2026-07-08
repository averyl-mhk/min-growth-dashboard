# MIN Dashboard — Monthly Update Process

## What you need before starting
- The raw export files from the MIN team (see file checklist below)
- Your terminal open in the dashboard folder
- The MIN Dashboard Claude Project open in a browser tab

---

## Step 1 — Drop the raw files into the right folder

Create a new folder under `raw_exports/` named `YYYY_MM` (e.g. `2026_05` for May).

```
growth-marketing-dashboard/
  raw_exports/
    2026_05/       ← create this, drop all files here
```

**File checklist** — you need these from the MIN team each month:

| What | How to recognise it |
|------|-------------------|
| Amazon campaign report | filename contains "Amazon"/"mazon" OR "Campaign", and does NOT contain "Search", "Targeting", "Sponsored", or another platform name (so "Campaign May Dashboard.csv" works too) |
| Amazon SP search terms | filename contains "Sponsored_Products" or "SP" and "Search" |
| Amazon SB search terms | filename contains "Sponsored_Brands" |
| Amazon SD targeting | filename contains "Sponsored_Display" or "SD" and "Target" |
| Meta campaigns | filename contains "Meta" or "eta" |
| Google Ads campaigns | filename contains "Google" (the 2 preamble rows are skipped automatically) |
| Flipkart campaigns | filename contains "Flipkart" or "lipkart" |
| POS file (= DTC sales) | filename contains "POS"/"pos". This is the DTC sales channel; the Shopify row is the DTC line used for MER. Optional/manual — often arrives late. |

The processor finds files by pattern — exact names don't matter as long as they match the above.

---

## Step 2 — Run the processor

In PowerShell, from the dashboard folder:

```powershell
python processor.py --month 2026_05
```

Replace `2026_05` with the actual month.

**What to check in the output:**
- Reconciliation numbers should show `Δ₹0.00` (or very close) for all ad types — this confirms spend and revenue totals match exactly
- "Missing:" line should ideally be empty. If files are missing, check your folder and re-run
- The platform line should list **Google Ads** (not "MISSING") when the Google file is included
- ROAS, spend, and POS numbers should look roughly right
- A **POS / DTC check** block prints at the end: if the POS file is present it shows the DTC (Shopify) line and MER (total POS ÷ spend); if absent it flags POS as "NOT YET PROVIDED — pending" so you know to add it later and re-run

If something looks wrong, fix the raw files and re-run before moving on.

---

## Step 3 — Generate the AI brief

```powershell
python processor.py --month 2026_05 --export-brief
```

This creates `ai_brief_2026_05.txt` in the dashboard folder. It contains all the processed data pre-formatted for Claude to analyse.

---

## Step 4 — Get AI recommendations from Claude

1. Open your **MIN Dashboard Claude Project** in Claude.ai
2. Start a **new chat** (your system prompt is already saved in the Project — don't re-paste it)
3. Open `ai_brief_2026_05.txt`, copy the entire contents, and paste it into the chat
4. Claude will return a JSON object with 8–12 recommendations
5. Copy the entire JSON response and save it as `recs_2026_05.json` in the dashboard folder

---

## Step 5 — Ingest the recommendations

```powershell
python ingest_recs.py --month 2026_05 --file recs_2026_05.json
```

You should see a confirmation like:
```
✓ AI recommendations for May '26 written to data.json
  10 recommendations  |  4 data quality flags
```

If you get a validation error, the JSON Claude returned may have extra text around it — open `recs_2026_05.json`, delete anything before the first `{` and after the last `}`, and re-run.

---

## Step 6 — Push to GitHub

```powershell
git add data.json
git commit -m "Add May '26 data and AI recommendations"
git push
```

GitHub Pages updates within ~60 seconds. Then share the dashboard link with the team:
**https://averyl-mhk.github.io/min-growth-dashboard/**

---

## Quick reference — all commands in order

```powershell
# 1. Process the raw data
python processor.py --month 2026_05

# 2. Generate the AI brief
python processor.py --month 2026_05 --export-brief

# (Paste ai_brief_2026_05.txt into Claude Project, save response as recs_2026_05.json)

# 3. Ingest recommendations
python ingest_recs.py --month 2026_05 --file recs_2026_05.json

# 4. Push
git add data.json
git commit -m "Add May '26 data and AI recommendations"
git push
```

---

## Troubleshooting

**"Folder not found" error on processor**
→ Check that you created `raw_exports/2026_05/` and the month string matches exactly.

**"Missing: Amazon_Campaigns" (or similar) in processor output**
→ The file wasn't found. Check the filename contains the right keyword (see Step 1 table). Rename if needed and re-run.

**Reconciliation shows a large Δ (not near zero)**
→ The search term report and campaign report are out of sync — likely different date ranges. Check that all files cover the same month.

**ingest_recs.py says "Invalid JSON"**
→ Claude's response had prose before or after the JSON. Open the `.json` file, delete everything before `{` and after `}`, save, re-run.

**"already exist — use --force to overwrite"**
→ You've already ingested this month. To replace with a new version: add `--force` to the ingest command.

**Dashboard not updating after push**
→ Wait 60 seconds and hard-refresh the browser (Ctrl+Shift+R).

**git says "Unable to create .git/index.lock: File exists" or "Operation not permitted"**
→ OneDrive left a stale git lock (the `.git` folder is inside OneDrive). In PowerShell from the dashboard folder, run `Remove-Item .git\index.lock`, then re-run your git commands. If it says the file is in use, close GitHub Desktop / VS Code's Git panel first. Pausing OneDrive sync while you commit also avoids this.

**Committing more than data.json**
→ Normally only `data.json` changes each month. If you also changed the dashboard code (`index.html`) or the processor (`processor.py`), add those files to the `git add` too — e.g. `git add data.json index.html processor.py`.
