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

from src.core.text_utils import tokenize, detokenize, normalize_tokens_out, visual_length
from src.core.validator import validate
from src.core.ai_client import translate_batch, refine, init_session, init_session_shadok, translate_shadok_block

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
SHADOK_JSON = DATA_DIR / "shadok.json"


def load_shadok_config() -> dict | None:
    """Load shadok.json config. Returns None if not found."""
    if not SHADOK_JSON.exists():
        return None
    with SHADOK_JSON.open("r", encoding="utf-8") as f:
        return json.load(f)

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

    # ── Phase 0: SHADOK BLOCK ────────────────────────────────────────
    shadok_config = load_shadok_config()
    shadok_row_set = set()  # rows to exclude from regular translation

    if shadok_config:
        shadok_lines = shadok_config["lines"]
        max_line_len = shadok_config["max_line_length"]

        # Build lookup and populate shadok_row_set (always, even if translated)
        shadok_rows = []  # list of (order_index, row_idx, original_text)
        excel_lookup = {}
        for row in range(2, ws.max_row + 1):
            val = ws.cell(row, col_map["Original"]).value
            if val:
                excel_lookup[str(val).strip()] = row

        for i, line in enumerate(shadok_lines):
            if line in excel_lookup:
                shadok_rows.append((i, excel_lookup[line], line))
                shadok_row_set.add(excel_lookup[line])

        # Check translated flag
        if shadok_config.get("translated", False):
            print(f"  [SHADOK] Already translated (flag in shadok.json). Skipping {len(shadok_row_set)} rows.")
        else:
            # Check which shadok rows need translation
            shadok_missing_langs = set()
            for _, row_idx, _ in shadok_rows:
                for lc in lang_codes:
                    cell_val = ws.cell(row_idx, col_map[lc]).value
                    if not cell_val or not str(cell_val).strip():
                        shadok_missing_langs.add(lc)

            if shadok_missing_langs:
                shadok_missing_langs = sorted(shadok_missing_langs)
                print()
                print("=" * 60)
                print("  SHADOK BLOCK TRANSLATION")
                print("=" * 60)
                print(f"  Lines: {len(shadok_rows)}, Missing langs: {len(shadok_missing_langs)}")
                print(f"  Max line length: {max_line_len}")
                print("-" * 60)

                # Init special shadok session
                init_session_shadok()

                # Join all lines into one text
                full_text = "\n".join(line for _, _, line in shadok_rows)

                try:
                    results = translate_shadok_block(full_text, shadok_missing_langs, max_line_len)

                    # Split results and write to dictionary
                    for lc in shadok_missing_langs:
                        translated_text = results.get(lc, "")
                        if not translated_text:
                            print(f"  [SHADOK][{lc}] Empty translation, skipping.")
                            continue

                        # Split translated text into lines
                        trans_lines = translated_text.split("\n")

                        # Enforce max line count
                        if len(trans_lines) > len(shadok_rows):
                            trans_lines = trans_lines[:len(shadok_rows)]

                        # Write each translated line to its corresponding row
                        for line_idx, (order_idx, row_idx, _orig) in enumerate(shadok_rows):
                            if line_idx < len(trans_lines):
                                line = trans_lines[line_idx]
                                # Enforce max line length
                                if len(line) > max_line_len:
                                    line = line[:max_line_len]
                                ws.cell(row_idx, col_map[lc], line)
                            else:
                                # Fewer translated lines than original — leave empty
                                ws.cell(row_idx, col_map[lc], "")

                        print(f"  [SHADOK][{lc}] Written {min(len(trans_lines), len(shadok_rows))} lines.")

                    save_workbook(wb)
                    print("  [SHADOK] Saved to dictionary.")

                    # Set translated flag in shadok.json
                    shadok_config["translated"] = True
                    with SHADOK_JSON.open("w", encoding="utf-8") as f:
                        json.dump(shadok_config, f, ensure_ascii=False, indent=2)
                    print("  [SHADOK] Flag 'translated' set to true in shadok.json.")

                except Exception as e:
                    print(f"  [SHADOK] ERROR: {e}")

                print("=" * 60)
                print()

    # ── Scan: count rows that need translation (excluding shadoks) ───
    rows_to_translate = []
    for row in range(2, ws.max_row + 1):
        if row in shadok_row_set:
            continue  # skip shadok rows
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
    print(f"  To translate: {total_rows} (excl. {len(shadok_row_set)} shadok rows)")
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

        # Skip AI translation for strings without Cyrillic — copy as-is
        import re as _re
        if not _re.search(r'[а-яА-ЯёЁіІїЇєЄґҐ]', original):
            for lc in missing:
                ws.cell(row, col_map[lc], original)
                total_translated += 1
            save_workbook(wb)
            print(f"    [Row {row} | {idx}/{total_rows}] No Cyrillic — copied as-is.")
            continue

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
            if not translation or not translation.strip():
                total_failed += 1
                row_fail += 1
                continue
            
            ok, msg = validate(original, translation, lc)
            if ok:
                ws.cell(row, col_map[lc], translation)
                total_translated += 1
                row_ok += 1
            elif "English preservation" in msg:
                # ASCII-only text should stay as original
                ws.cell(row, col_map[lc], original)
                total_translated += 1
                row_ok += 1
            else:
                # Invalid translation — do NOT write to Excel
                print(f"    [Row {row}][{lc}] SKIPPED (invalid): {msg}")
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
    """Validate all translations in the dictionary using advanced Validator."""
    from src.core.validator import Validator
    langs = load_languages()
    wb = open_or_create_workbook()
    ws = wb[SHEET_NAME]

    header = [ws.cell(1, c).value for c in range(1, ws.max_column + 1)]
    col_map = {h: i + 1 for i, h in enumerate(header) if h}

    # Pass BLOCK_JSON to validator for regex-aware checks
    validator = Validator(str(BLOCK_JSON) if BLOCK_JSON.exists() else None)

    # Build shadok exclusion set
    shadok_row_set = set()
    shadok_config = load_shadok_config()
    if shadok_config:
        excel_lookup = {}
        for row in range(2, ws.max_row + 1):
            val = ws.cell(row, col_map["Original"]).value
            if val:
                excel_lookup[str(val).strip()] = row
        for line in shadok_config["lines"]:
            if line in excel_lookup:
                shadok_row_set.add(excel_lookup[line])
        if shadok_row_set:
            print(f"  Excluding {len(shadok_row_set)} shadok rows from validation.")

    errors = 0
    checked = 0

    print(f"Starting validation for {len([l for l in langs if l != 'ru'])} languages...")

    for row in range(2, ws.max_row + 1):
        if row in shadok_row_set:
            continue  # skip shadok rows
        original = ws.cell(row, col_map["Original"]).value
        if not original:
            continue
        original_str = str(original)
        
        for lc in langs:
            if lc not in col_map or lc == "ru":
                continue
            translation = ws.cell(row, col_map[lc]).value
            if not translation:
                continue
            
            checked += 1
            row_errors = validator.validate_row(original_str, str(translation))
            if row_errors:
                errors += len(row_errors)
                print(f"  [Row {row}][{lc}] Error(s):")
                for err in row_errors:
                    print(f"    - {err}")

    print(f"\n--- Translation checks: {checked} checked, {errors} issues ---")

    # Phase 2: Regex block validation (blocks.json patterns vs Original column)
    if validator.compiled_blocks:
        print("\nRunning regex block validation (blocks.json)...")
        originals = {}
        for row in range(2, ws.max_row + 1):
            val = ws.cell(row, col_map["Original"]).value
            if val:
                originals[row] = str(val).strip()

        block_errors = validator.validate_blocks(originals)
        if block_errors:
            errors += len(block_errors)
            for err in block_errors:
                print(f"  {err}")
            print(f"--- Block regex checks: {len(block_errors)} issues ---")
        else:
            print("--- Block regex checks: all OK ---")

    print(f"\nValidation complete. Total issues: {errors}")


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


def cmd_align() -> None:
    """Align colons in blocks defined in data/blocks.json (regex-based)."""
    import re as _re

    blocks = load_blocks()
    if not blocks:
        print("No blocks defined in blocks.json.")
        return

    langs = load_languages()
    wb = open_or_create_workbook()
    ws = wb[SHEET_NAME]
    header = [ws.cell(1, c).value for c in range(1, ws.max_column + 1)]
    col_map = {h: i + 1 for i, h in enumerate(header) if h}

    # Cache all Original values for faster regex matching
    originals = {}  # row -> original_str
    for row in range(2, ws.max_row + 1):
        val = ws.cell(row, col_map["Original"]).value
        if val:
            originals[row] = str(val).strip()

    # Build grouped_data[block_id] = list of matched row indices
    grouped_data = {}  # block_id -> [row_idx, ...]

    for bid, patterns in blocks.items():
        matched_rows = []
        for pattern in patterns:
            try:
                regex = _re.compile(pattern, _re.IGNORECASE)
            except _re.error as e:
                print(f"  [WARN] Invalid regex in block {bid}: {pattern} -> {e}")
                continue

            found = False
            for row, orig_val in originals.items():
                if regex.search(orig_val):
                    matched_rows.append(row)
                    found = True
                    break

            if not found:
                print(f"  [WARN] Pattern not found for block {bid}: {pattern[:50]}...")

        if matched_rows:
            grouped_data[bid] = matched_rows

    if not grouped_data:
        print("No rows found matching blocks.json patterns.")
        return

    # Language columns to process (all except 'ru' and 'Original')
    lang_cols = {lc: col_map[lc] for lc in langs if lc in col_map and lc != "ru"}
    # Also always process Original column
    lang_cols["Original"] = col_map["Original"]

    affected_count = 0

    for bid, matched_rows in grouped_data.items():
        print(f"  Aligning block: {bid} ({len(matched_rows)} rows)")

        for lc, col_idx in lang_cols.items():
            # 1. Collect prefix lengths for this language in this block
            max_prefix_len = 0
            row_parts = []

            for row_idx in matched_rows:
                val = ws.cell(row_idx, col_idx).value
                if not val:
                    continue
                val_str = str(val).strip()

                if ":" in val_str:
                    prefix, suffix = val_str.split(":", 1)
                    clean_prefix = prefix.rstrip()  # remove trailing spaces before colon
                    visual_len = visual_length(clean_prefix)
                    max_prefix_len = max(max_prefix_len, visual_len)
                    row_parts.append((row_idx, clean_prefix, suffix))
                else:
                    row_parts.append((row_idx, None, val_str))

            if max_prefix_len == 0:
                continue

            # 2. Apply padding so all colons line up
            target_len = max_prefix_len + 1  # at least 1 space before colon

            for row_idx, prefix, suffix in row_parts:
                if prefix is None:
                    continue

                current_visual_len = visual_length(prefix)
                padding = target_len - current_visual_len
                if padding < 1:
                    padding = 1

                new_val = f"{prefix}{' ' * padding}:{suffix}"
                old_val = ws.cell(row_idx, col_idx).value
                if old_val != new_val:
                    ws.cell(row_idx, col_idx).value = new_val
                    affected_count += 1

    if affected_count > 0:
        ver = bump_version(wb)
        save_workbook(wb)
        print(f"\nAlignment done. Adjusted {affected_count} cells. Version: {ver}")
    else:
        print("\nAlignment done. No changes needed.")



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

    # Clear log at every start
    log_path = ROOT / "logs" / "ai_proxy.log"
    log_path.parent.mkdir(parents=True, exist_ok=True)
    from datetime import datetime, timezone
    with open(log_path, "w", encoding="utf-8") as f:
        f.write(f"--- SESSION: {args.command} | {datetime.now(timezone.utc).isoformat()} ---\n")

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
