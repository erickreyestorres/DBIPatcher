"""DBI Translation Pipeline.

Usage:
    python -m src.main sync        — sync ua.csv into dictionary.xlsx
    python -m src.main translate   — translate missing cells via AI
    python -m src.main validate    — validate all translations
    python -m src.main export      — export per-language CSVs
    python -m src.main build       — build .bin files from CSVs
    python -m src.main all         — run full pipeline
"""

from __future__ import annotations

import argparse
import csv
import json
import subprocess
import sys
import time
from pathlib import Path

import openpyxl

from src.core.text_utils import tokenize, detokenize, normalize_tokens_out
from src.core.validator import validate
from src.core.ai_client import translate_batch, refine, init_session

ROOT = Path(__file__).resolve().parent.parent
DATA_DIR = ROOT / "data"
DICT_PATH = DATA_DIR / "dictionary.xlsx"
UA_CSV = DATA_DIR / "ua.csv"
LANG_JSON = DATA_DIR / "languages.json"
OUTPUT_DIR = ROOT / "output"
BUILD_SCRIPT = ROOT / "scripts" / "build_translation_bin.py"

SHEET_NAME = "Translations"
META_SHEET = "Metadata"


BLOCK_JSON = DATA_DIR / "blocks.json"


# ── helpers ──────────────────────────────────────────────────────────

def load_languages() -> dict[str, str]:
    with LANG_JSON.open("r", encoding="utf-8") as f:
        return json.load(f)


def load_blocks() -> dict[str, list[str]]:
    if not BLOCK_JSON.exists():
        return {}
    with BLOCK_JSON.open("r", encoding="utf-8") as f:
        return json.load(f)


def open_or_create_workbook() -> openpyxl.Workbook:
    if DICT_PATH.exists():
        return openpyxl.load_workbook(DICT_PATH)
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = SHEET_NAME
    meta = wb.create_sheet(META_SHEET)
    meta["A1"] = "version"
    meta["B1"] = "0.0.0"
    meta["A2"] = "updated"
    meta["B2"] = ""
    return wb


def get_version(wb: openpyxl.Workbook) -> str:
    meta = wb[META_SHEET]
    return str(meta["B1"].value or "0.0.0")


def bump_version(wb: openpyxl.Workbook) -> str:
    meta = wb[META_SHEET]
    ver = str(meta["B1"].value or "0.0.0")
    parts = ver.split(".")
    parts[-1] = str(int(parts[-1]) + 1)
    new_ver = ".".join(parts)
    meta["B1"] = new_ver
    from datetime import datetime, timezone
    meta["B2"] = datetime.now(timezone.utc).isoformat()
    return new_ver


def save_workbook(wb: openpyxl.Workbook) -> None:
    DICT_PATH.parent.mkdir(parents=True, exist_ok=True)
    max_retries = 10
    for attempt in range(max_retries):
        try:
            wb.save(DICT_PATH)
            return
        except PermissionError:
            if attempt < max_retries - 1:
                print(f"\n  [!] File '{DICT_PATH.name}' is locked (likely open in Excel).")
                print(f"      PLEASE CLOSE IT! Retrying in 3s... ({attempt + 1}/{max_retries})")
                time.sleep(3)
            else:
                print(f"\n  [CRITICAL] Could not save '{DICT_PATH.name}' after {max_retries} attempts.")
                raise


# ── sync ─────────────────────────────────────────────────────────────

def cmd_sync() -> None:
    """Read ua.csv, tokenize, add missing rows to dictionary.xlsx."""
    langs = load_languages()
    wb = open_or_create_workbook()
    ws = wb[SHEET_NAME]

    # Ensure header row
    expected_cols = ["Original"] + list(langs.keys())
    if ws.max_row == 0 or ws.cell(1, 1).value is None:
        for ci, col in enumerate(expected_cols, 1):
            ws.cell(1, ci, col)

    # Read current header to build col map
    header = [ws.cell(1, c).value for c in range(1, ws.max_column + 1)]

    # Add missing language columns
    for lang_code in langs:
        if lang_code not in header:
            idx = len(header) + 1
            ws.cell(1, idx, lang_code)
            header.append(lang_code)

    col_map = {h: i + 1 for i, h in enumerate(header) if h}

    # Collect existing originals
    existing: set[str] = set()
    for row in range(2, ws.max_row + 1):
        val = ws.cell(row, col_map["Original"]).value
        if val:
            existing.add(val)

    # Read ua.csv and insert missing
    added = 0
    with UA_CSV.open("r", encoding="utf-8", newline="") as f:
        reader = csv.DictReader(f)
        for rec in reader:
            ru_raw = rec.get("original", "") or ""
            if not ru_raw:
                continue
            ru_tok = tokenize(ru_raw)
            if ru_tok in existing:
                continue
            next_row = ws.max_row + 1
            ws.cell(next_row, col_map["Original"], ru_tok)
            existing.add(ru_tok)
            added += 1

    ver = bump_version(wb)
    save_workbook(wb)
    print(f"Sync done. Added: {added}, Version: {ver}")


# ── translate ────────────────────────────────────────────────────────

def cmd_translate() -> None:
    """Find empty cells and translate via AI with continuous chat."""
    MAX_REFINE_ATTEMPTS = 3

    langs = load_languages()
    wb = open_or_create_workbook()
    ws = wb[SHEET_NAME]

    header = [ws.cell(1, c).value for c in range(1, ws.max_column + 1)]
    col_map = {h: i + 1 for i, h in enumerate(header) if h}

    lang_codes = [lc for lc in langs if lc in col_map and lc != "ru"]

    # ── Scan: count rows that need translation ───────────────────────
    rows_to_translate = []
    for row in range(2, ws.max_row + 1):
        original = ws.cell(row, col_map["Original"]).value
        if not original:
            continue
        missing = []
        for lc in lang_codes:
            cell_val = ws.cell(row, col_map[lc]).value
            if not cell_val or not str(cell_val).strip():
                missing.append(lc)
        if missing:
            rows_to_translate.append((row, original, missing))

    total_rows = len(rows_to_translate)
    total_all = ws.max_row - 1  # excluding header

    if total_rows == 0:
        print(f"  Nothing to translate. All {total_all} rows are complete.")
        return

    # ── Phase 1: INIT ────────────────────────────────────────────────
    print()
    print("=" * 60)
    print("  DBI TRANSLATOR")
    print("=" * 60)
    print(f"  Languages : {len(lang_codes)} ({', '.join(lang_codes)})")
    print(f"  Total rows: {total_all}")
    print(f"  To translate: {total_rows}")
    print("-" * 60)

    init_session()

    print("-" * 60)
    print(f"  Starting translation...")
    print("=" * 60)
    print()

    # ── Phase 2: TRANSLATE ───────────────────────────────────────────
    total_translated = 0
    total_failed = 0

    for idx, (row, original, missing) in enumerate(rows_to_translate, 1):
        print(f"  [Row {row} | {idx}/{total_rows}] {original[:50]}  -> {len(missing)} langs")

        try:
            results = translate_batch(original, missing, row_id=row)
        except Exception as e:
            print(f"  [Row {row} | {idx}/{total_rows}] ERROR: {e}")
            total_failed += len(missing)
            time.sleep(2)
            continue

        # Normalize tokens in AI output
        for lc in list(results.keys()):
            results[lc] = normalize_tokens_out(results.get(lc, ""))

        # ── Validate + Refine loop ───────────────────────────────────
        for attempt in range(MAX_REFINE_ATTEMPTS):
            errors = []
            for lc in missing:
                translation = results.get(lc, "")
                if not translation:
                    errors.append((lc, "Translation is empty"))
                    continue
                ok, msg = validate(original, translation, lc)
                if not ok:
                    errors.append((lc, msg))

            if not errors:
                break

            if attempt < MAX_REFINE_ATTEMPTS - 1:
                error_lines = "\n".join(f"- {lc}: {msg}" for lc, msg in errors)
                correction = (
                    f"The following translations have errors:\n"
                    f"{error_lines}\n\n"
                    f"Please fix them. Source text: \"{original}\""
                )
                print(f"    [Row {row} | {idx}/{total_rows}] Refine #{attempt + 2}: {len(errors)} errors")
                try:
                    refined = refine(correction, missing, row_id=row)
                    for lc in list(refined.keys()):
                        refined[lc] = normalize_tokens_out(refined.get(lc, ""))
                    results.update(refined)
                except Exception as e:
                    print(f"    [Row {row} | {idx}/{total_rows}] Refine error: {e}")
                    break
                time.sleep(0.3)

        # ── Write results ────────────────────────────────────────────
        row_ok = 0
        row_fail = 0
        for lc in missing:
            translation = results.get(lc, "")
            if translation and translation.strip():
                ok, msg = validate(original, translation, lc)
                if not ok and "English preservation" in msg:
                    translation = original
                ws.cell(row, col_map[lc], translation)
                total_translated += 1
                row_ok += 1
            else:
                total_failed += 1
                row_fail += 1

        save_workbook(wb)

        status = "OK" if row_fail == 0 else f"OK:{row_ok} FAIL:{row_fail}"
        print(f"    [Row {row} | {idx}/{total_rows}] Saved. {status}")

        time.sleep(0.5)

    # ── Phase 3: SUMMARY ─────────────────────────────────────────────
    print()
    print("=" * 60)
    ver = bump_version(wb)
    save_workbook(wb)
    print(f"  DONE! Translated: {total_translated}, Failed: {total_failed}")
    print(f"  Version: {ver}")
    print("=" * 60)




# ── validate ─────────────────────────────────────────────────────────

def cmd_validate() -> None:
    """Validate all translations in the dictionary."""
    langs = load_languages()
    wb = open_or_create_workbook()
    ws = wb[SHEET_NAME]

    header = [ws.cell(1, c).value for c in range(1, ws.max_column + 1)]
    col_map = {h: i + 1 for i, h in enumerate(header) if h}

    errors = 0
    checked = 0

    for row in range(2, ws.max_row + 1):
        original = ws.cell(row, col_map["Original"]).value
        if not original:
            continue
        for lc in langs:
            if lc not in col_map or lc == "ru":
                continue
            translation = ws.cell(row, col_map[lc]).value
            if not translation:
                continue
            ok, msg = validate(original, str(translation), lc)
            checked += 1
            if not ok:
                errors += 1
                print(f"  [Row {row}][{lc}] {msg}")

    print(f"Validation done. Checked: {checked}, Errors: {errors}")


# ── export ───────────────────────────────────────────────────────────

def cmd_export() -> None:
    """Export per-language CSV files (original, translation) for the bin builder."""
    langs = load_languages()
    wb = open_or_create_workbook()
    ws = wb[SHEET_NAME]

    header = [ws.cell(1, c).value for c in range(1, ws.max_column + 1)]
    col_map = {h: i + 1 for i, h in enumerate(header) if h}

    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

    for lc in langs:
        if lc not in col_map or lc == "ru":
            continue
        csv_path = OUTPUT_DIR / f"{lc}.csv"
        count = 0
        with csv_path.open("w", encoding="utf-8", newline="") as f:
            writer = csv.writer(f)
            writer.writerow(["original", "translation"])
            for row in range(2, ws.max_row + 1):
                original = ws.cell(row, col_map["Original"]).value
                translation = ws.cell(row, col_map[lc]).value
                if not original or not translation:
                    continue
                writer.writerow([detokenize(original), detokenize(str(translation))])
                count += 1
        print(f"  {lc}.csv: {count} entries")

    print("Export done.")


# ── build ────────────────────────────────────────────────────────────

def cmd_build() -> None:
    """Run build_translation_bin.py for each exported CSV."""
    langs = load_languages()
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

    for lc in langs:
        if lc == "ru":
            continue
        csv_path = OUTPUT_DIR / f"{lc}.csv"
        if not csv_path.exists():
            print(f"  Skip {lc}: no CSV")
            continue
        bin_path = OUTPUT_DIR / f"translation_{lc}.bin"
        cmd = [sys.executable, str(BUILD_SCRIPT), str(csv_path), "-o", str(bin_path)]
        print(f"  Building {bin_path.name}...")
        result = subprocess.run(cmd, capture_output=True, text=True)
        if result.returncode != 0:
            print(f"  [!] Error building {lc}: {result.stderr}")
        else:
            print(f"  OK: {bin_path.name}")

    print("Build done.")


# ── align ────────────────────────────────────────────────────────────

def cmd_align() -> None:
    """Align colons in blocks defined in data/blocks.json."""
    blocks = load_blocks()
    if not blocks:
        print("No blocks defined in blocks.json.")
        return

    langs = load_languages()
    wb = open_or_create_workbook()
    ws = wb[SHEET_NAME]
    header = [ws.cell(1, c).value for c in range(1, ws.max_column + 1)]
    col_map = {h: i + 1 for i, h in enumerate(header) if h}

    # Invert blocks for quick lookup: { "OriginalText": "BlockID" }
    row_to_block = {}
    for bid, texts in blocks.items():
        for t in texts:
            row_to_block[t] = bid

    # Collect all rows that belong to blocks
    # grouped_data[block_id][lang_code] = list of (row_idx, current_val)
    grouped_data = {}

    for row in range(2, ws.max_row + 1):
        original = ws.cell(row, col_map["Original"]).value
        if not original or original not in row_to_block:
            continue
        
        bid = row_to_block[original]
        if bid not in grouped_data:
            grouped_data[bid] = {}

        for lc in langs:
            if lc not in col_map or lc == "ru":
                continue
            if lc not in grouped_data[bid]:
                grouped_data[bid][lc] = []
            
            val = ws.cell(row, col_map[lc]).value
            if val:
                grouped_data[bid][lc].append((row, str(val)))

    if not grouped_data:
        print("No rows found matching blocks.json.")
        return

    affected_count = 0
    for bid, lang_dict in grouped_data.items():
        for lc, rows in lang_dict.items():
            if not rows:
                continue

            # Find max length before first colon
            # Note: we use detokenize to get real character count (approx)
            # but usually labels don't have many tokens.
            max_prefix = 0
            row_parts = []

            for row_idx, val in rows:
                if ":" in val:
                    prefix, suffix = val.split(":", 1)
                    # Detokenize to count visual length accurately
                    visual_prefix = detokenize(prefix.strip())
                    max_prefix = max(max_prefix, len(visual_prefix))
                    row_parts.append((row_idx, prefix.strip(), suffix))
                else:
                    # No colon? Skip alignment for this row but keep it in list
                    row_parts.append((row_idx, None, val))

            # Apply padding
            for row_idx, prefix, suffix in row_parts:
                if prefix is not None:
                    # Pad prefix back. We need to handle tokens.
                    # Simplest: pad detokenized prefix, then re-tokenize?
                    # But usually prefixes are just text. 
                    # Let's pad and re-join.
                    padded_prefix = detokenize(prefix).ljust(max_prefix)
                    # We re-tokenize just in case, though padding spaces shouldn't affect tokens
                    new_val = f"{tokenize(padded_prefix)}:{suffix}"
                    if ws.cell(row_idx, col_map[lc]).value != new_val:
                        ws.cell(row_idx, col_map[lc]).value = new_val
                        affected_count += 1

    if affected_count > 0:
        ver = bump_version(wb)
        save_workbook(wb)
        print(f"Alignment done. Adjusted {affected_count} cells. Version: {ver}")
    else:
        print("Alignment done. No changes needed.")


# ── clear ────────────────────────────────────────────────────────────

def cmd_clear(lang_code: str) -> None:
    """Clear all translations for a specific language in the dictionary."""
    if not DICT_PATH.exists():
        print("No dictionary found.")
        return

    wb = openpyxl.load_workbook(DICT_PATH)
    if SHEET_NAME not in wb.sheetnames:
        print(f"Sheet {SHEET_NAME} not found.")
        return

    ws = wb[SHEET_NAME]
    header = [ws.cell(1, c).value for c in range(1, ws.max_column + 1)]
    col_map = {h: i + 1 for i, h in enumerate(header) if h}

    if lang_code not in col_map:
        print(f"Language '{lang_code}' not found in dictionary columns.")
        return

    col_idx = col_map[lang_code]
    cleared = 0
    for row in range(2, ws.max_row + 1):
        if ws.cell(row, col_idx).value:
            ws.cell(row, col_idx).value = None
            cleared += 1

    ver = bump_version(wb)
    save_workbook(wb)
    print(f"Cleared {cleared} entries for '{lang_code}'. Version: {ver}")


# ── main ─────────────────────────────────────────────────────────────

COMMANDS = {
    "sync": cmd_sync,
    "translate": cmd_translate,
    "validate": cmd_validate,
    "align": cmd_align,
    "export": cmd_export,
    "build": cmd_build,
    "clear": cmd_clear,
}


def cmd_all() -> None:
    for name in ("sync", "translate", "align", "validate", "export", "build"):
        print(f"\n{'='*60}\n  STEP: {name}\n{'='*60}")
        COMMANDS[name]()


def main() -> int:
    parser = argparse.ArgumentParser(description="DBI Translation Pipeline")
    parser.add_argument("command", choices=[*COMMANDS.keys(), "all"], help="Pipeline step to run")
    parser.add_argument("lang", nargs="?", help="Language code (required for 'clear')")
    args = parser.parse_args()

    if args.command == "all":
        cmd_all()
    elif args.command == "clear":
        if not args.lang:
            print("Error: 'clear' command requires a language code (e.g., 'ua')")
            sys.exit(1)
        cmd_clear(args.lang)
    else:
        COMMANDS[args.command]()
    return 0


if __name__ == "__main__":
    sys.exit(main())
