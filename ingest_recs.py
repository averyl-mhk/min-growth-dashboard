"""ingest_recs.py — Write Claude's AI recommendations into data.json.

Usage:
  python ingest_recs.py --month 2026_03 --file recs_2026_03.json

Workflow:
  1. Run:  python processor.py --month 2026_03 --export-brief
  2. Open your MIN Dashboard Claude Project and paste ai_brief_2026_03.txt
  3. Save Claude's JSON response to a file, e.g. recs_2026_03.json
  4. Run:  python ingest_recs.py --month 2026_03 --file recs_2026_03.json

The recommendations are written to data.json under aiRecommendations["Mar '26"].
Re-running without --force will refuse to overwrite an existing entry.
"""

import argparse
import json
import sys
from datetime import datetime
from pathlib import Path

SCRIPT_DIR = Path(__file__).parent
OUTPUT_FILE = SCRIPT_DIR / "data.json"

REQUIRED_KEYS = {"recommendations", "summary", "counts"}


def validate(recs: dict, month_label: str) -> list[str]:
    """Return a list of validation error strings (empty = valid)."""
    errors = []
    missing = REQUIRED_KEYS - set(recs.keys())
    if missing:
        errors.append(f"Missing required keys: {', '.join(sorted(missing))}")
    if not isinstance(recs.get("recommendations"), list):
        errors.append("'recommendations' must be a JSON array")
    else:
        for i, rec in enumerate(recs["recommendations"], 1):
            for field in ("n", "priority", "title", "platform", "whatDataShows", "exactAction"):
                if field not in rec:
                    errors.append(f"Recommendation #{i} is missing field '{field}'")
            priority = rec.get("priority", "")
            if priority not in ("critical", "warning", "opportunity", "insight", ""):
                errors.append(f"Recommendation #{i} has unknown priority '{priority}'")
    counts = recs.get("counts")
    if counts and not isinstance(counts, dict):
        errors.append("'counts' must be a JSON object")
    return errors


def main():
    parser = argparse.ArgumentParser(
        description="Ingest Claude's AI recommendations JSON into data.json",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=__doc__,
    )
    parser.add_argument("--month", required=True, help="Month folder name, e.g. 2026_03")
    parser.add_argument("--file", required=True, help="Path to JSON file with Claude's recommendations")
    parser.add_argument(
        "--force",
        action="store_true",
        help="Overwrite existing recommendations for this month",
    )
    args = parser.parse_args()

    # Parse month label
    try:
        year, month = args.month.split("_")
        month_label = datetime(int(year), int(month), 1).strftime("%b '%y")
    except (ValueError, TypeError):
        sys.exit(f"ERROR: --month must be in YYYY_MM format, got '{args.month}'")

    # Load recommendations file
    recs_path = Path(args.file)
    if not recs_path.exists():
        sys.exit(f"ERROR: File not found: {recs_path}")

    with open(recs_path, encoding="utf-8") as f:
        raw = f.read().strip()

    # Strip markdown code fences if Claude wrapped the JSON in ```json ... ```
    if raw.startswith("```"):
        lines = raw.splitlines()
        raw = "\n".join(
            line for line in lines if not line.strip().startswith("```")
        ).strip()

    try:
        recs = json.loads(raw)
    except json.JSONDecodeError as e:
        sys.exit(
            f"ERROR: Could not parse JSON from {recs_path}.\n"
            f"  Make sure the file contains only the JSON object Claude returned.\n"
            f"  Detail: {e}"
        )

    # Validate
    errors = validate(recs, month_label)
    if errors:
        print("ERROR: Recommendations JSON failed validation:")
        for err in errors:
            print(f"  • {err}")
        sys.exit(1)

    # Load data.json
    if not OUTPUT_FILE.exists():
        sys.exit(
            f"ERROR: {OUTPUT_FILE} not found.\n"
            f"  Run processor.py --month {args.month} first to create it."
        )

    with open(OUTPUT_FILE, encoding="utf-8") as f:
        data = json.load(f)

    # Check for existing entry
    data.setdefault("aiRecommendations", {})
    if month_label in data["aiRecommendations"] and not args.force:
        n_existing = len(data["aiRecommendations"][month_label].get("recommendations", []))
        print(f"⚠  AI recommendations for {month_label} already exist in data.json ({n_existing} recs).")
        print(f"   Use --force to overwrite them.")
        sys.exit(0)

    # Stamp with ingest date if not already present
    if "generatedDate" not in recs:
        recs["generatedDate"] = datetime.today().strftime("%Y-%m-%d")
    recs["month"] = month_label

    # Write to a temp file first, verify it's valid JSON, then replace
    data["aiRecommendations"][month_label] = recs

    tmp_file = OUTPUT_FILE.with_suffix(".tmp")
    try:
        with open(tmp_file, "w", encoding="utf-8") as f:
            json.dump(data, f, indent=2, ensure_ascii=False)

        # Verify the temp file round-trips cleanly before replacing
        with open(tmp_file, encoding="utf-8") as f:
            json.load(f)

        tmp_file.replace(OUTPUT_FILE)

    except Exception as e:
        tmp_file.unlink(missing_ok=True)
        sys.exit(f"ERROR: Failed to write data.json safely: {e}\n  Original file was not modified.")

    n_recs = len(recs.get("recommendations", []))
    n_flags = len(recs.get("dataQualityFlags", []))
    counts = recs.get("counts", {})
    print(f"✓ AI recommendations for {month_label} written to data.json")
    print(f"  {n_recs} recommendations  |  {n_flags} data quality flags")
    if counts:
        parts = [f"{v} {k}" for k, v in counts.items() if v]
        print(f"  Breakdown: {', '.join(parts)}")
    print(f"  Summary: {recs.get('summary', '')[:120]}")


if __name__ == "__main__":
    main()
