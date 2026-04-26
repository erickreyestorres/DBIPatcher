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
import re
import subprocess
import sys
import time
from pathlib import Path

import openpyxl

from src.core.text_utils import tokenize, detokenize, normalize_tokens_out, visual_length, normalize_fullwidth
from src.core.validator import validate
from src.core.ai_client import translate_batch, refine, init_session, init_session_shadok, translate_shadok_block

ROOT = Path(__file__).resolve().parent.parent
DATA_DIR = ROOT / "data"
DICT_PATH = DATA_DIR / "dictionary.xlsx"
UA_CSV = DATA_DIR / "ua.csv"
LANG_JSON = DATA_DIR / "languages.json"
OUTPUT_DIR = ROOT / "output"
TRANSLATIONS_DIR = ROOT / "translations"
BUILD_SCRIPT = ROOT / "scripts" / "build_translation_bin.py"

SHEET_NAME = "Translations"
META_SHEET = "Metadata"

BLOCK_JSON = DATA_DIR / "blocks.json"
SHADOK_JSON = DATA_DIR / "shadok.json"

def find_dbi_version_row(ws, col_map) -> int:
    """Find the row that likely contains the pure 3-4 digit version number."""
    import re as _re
    for row in range(2, min(50, ws.max_row + 1)):
        val = str(ws.cell(row, col_map["Original"]).value or "").strip()
        if _re.match(r"^\d{3,4}$", val):
            return row
    return 9  # fallback


def load_shadok_config() -> dict | None:
    """Load shadok.json config. Returns None if not found."""
    if not SHADOK_JSON.exists():
        return None
    with SHADOK_JSON.open("r", encoding="utf-8") as f:
        return json.load(f)


def get_nro_version() -> str | None:
    """Extract DBI version from NRO filename (e.g. DBI.892.ru_patched.nro -> '892')."""
    import re as _re
    for nro_file in ROOT.glob("DBI.*.nro"):
        match = _re.search(r'DBI\.(\d+)\.', nro_file.name)
        if match:
            return match.group(1)
    return None

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

    # Remove columns not in languages.json (except 'Original')
    valid_cols = {"Original"} | set(langs.keys())
    cols_to_remove = []
    for ci, col_name in enumerate(header):
        if col_name and col_name not in valid_cols:
            cols_to_remove.append((ci + 1, col_name))  # 1-indexed
    
    if cols_to_remove:
        # Delete in reverse order to preserve indices
        for col_idx, col_name in sorted(cols_to_remove, reverse=True):
            ws.delete_cols(col_idx)
            print(f"  Removed column '{col_name}' (not in languages.json)")
        # Re-read header after deletion
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
            
    # CRITICAL: Since Shadok translations replace 'Original' with their satellite text,
    # the original text from ua.csv will falsely appear as missing.
    # We must explicitly track original Shadok texts and match them dynamically.
    shadok_origs_stripped = set()
    shadok_config = load_shadok_config()
    if shadok_config:
        for item in shadok_config.get("mapping", []):
            if item["new"] in existing:
                # Keep track of the original Shadok texts (stripped)
                shadok_origs_stripped.add(item["orig"].strip())

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
            
            # Formatting-agnostic check for Shadok strings (which often have leading spaces)
            if ru_tok.strip() in shadok_origs_stripped:
                continue

            next_row = ws.max_row + 1
            ws.cell(next_row, col_map["Original"], ru_tok)
            existing.add(ru_tok)
            added += 1

    # Update DBI version from NRO filename
    nro_ver = get_nro_version()
    if nro_ver:
        version_row = find_dbi_version_row(ws, col_map)
        current_ver = ws.cell(version_row, col_map["Original"]).value
        if str(current_ver) != nro_ver:
            print(f"  Updating DBI version at row {version_row}: {current_ver} -> {nro_ver}")
            # Write version to Original and ALL language columns
            for col_idx in col_map.values():
                ws.cell(version_row, col_idx, nro_ver)
        else:
            print(f"  DBI version already current: {nro_ver}")
    else:
        print("  WARNING: No DBI.*.nro file found, version not updated.")

    ver = bump_version(wb)
    save_workbook(wb)
    print(f"Sync done. Added: {added}, Version: {ver}")


# ── translate ────────────────────────────────────────────────────────

def wrap_text(text: str, max_chars: int, lang_code: str) -> list[str]:
    """Wrap text into lines, with special handling for CJK width."""
    # Heuristic for wider CJK characters as requested by user
    effective_max = max_chars
    # Japanese (jp), Korean (kr), Chinese (zh/zhcn/zhtw)
    if lang_code.lower() in ["zhcn", "zhtw", "jp", "kr", "zh"]:
        # "1.5 times smaller"
        effective_max = int(max_chars / 1.5)
        
    import textwrap
    # replace_whitespace=True converts all tabs/newlines into spaces before wrapping
    lines = textwrap.wrap(text, width=effective_max, break_long_words=True, replace_whitespace=True)
    return lines

def cmd_translate() -> None:
    """Find empty cells and translate via AI with continuous chat."""
    MAX_REFINE_ATTEMPTS = 3

    langs = load_languages()
    wb = open_or_create_workbook()
    ws = wb[SHEET_NAME]

    force_all = "--force" in sys.argv or "-f" in sys.argv

    header = [ws.cell(1, c).value for c in range(1, ws.max_column + 1)]
    col_map = {h: i + 1 for i, h in enumerate(header) if h}

    lang_codes = [lc for lc in langs if lc in col_map and lc != "ru"]

    # ── Phase -1: CLEANUP DUPLICATE ROWS ───────────────────────────────
    # Rows are considered duplicates if their "Original" column is exactly the same
    best_rows = {}
    rows_to_delete = []
    
    for row in range(2, ws.max_row + 1):
        original_val = str(ws.cell(row, col_map["Original"]).value or "")
        key = original_val.strip()
        if not key:
            continue
            
        # Count translated languages
        non_empty = 0
        for lc in lang_codes:
            if str(ws.cell(row, col_map[lc]).value or "").strip():
                non_empty += 1
                
        if key in best_rows:
            best_idx, best_count = best_rows[key]
            if non_empty > best_count:
                # The current row is better (more translations)
                rows_to_delete.append(best_idx)
                best_rows[key] = (row, non_empty)
            else:
                # The previous row was better or equal, so we delete the current row
                rows_to_delete.append(row)
        else:
            best_rows[key] = (row, non_empty)
            
    if rows_to_delete:
        # Sort in reverse to delete from bottom to top so indices don't shift
        for row in sorted(rows_to_delete, reverse=True):
            ws.delete_rows(row)
        print(f"  [CLEANUP] Deleted {len(rows_to_delete)} exact duplicate rows.")
        save_workbook(wb)

    # ── Phase 0: SHADOK BLOCK ────────────────────────────────────────
    shadok_config = load_shadok_config()
    shadok_row_set = set()  # rows to exclude from regular translation

    if shadok_config:
        shadok_mapping = shadok_config.get("mapping", [])
        max_line_len = shadok_config.get("max_line_length", 39)
        config_changed = False

        # Resolve row numbers by searching for original lines
        excel_lookup = {}
        for row in range(2, ws.max_row + 1):
            val = ws.cell(row, col_map["Original"]).value
            if val:
                excel_lookup[str(val).strip()] = row

        shadok_rows = []
        shadok_order_idx = 0
        # First, find Excel rows for each Shadok block
        for item in shadok_config.get("mapping", []):
            orig_text = item["orig"]
            new_text = item["new"]
            row_idx = excel_lookup.get(orig_text.strip()) or excel_lookup.get(new_text.strip())
            
            if row_idx:
                # Capture the exact leading whitespace from the actual Excel string before overwriting it
                actual_val = str(ws.cell(row_idx, col_map["Original"]).value or "")
                leading_spaces_count = len(actual_val) - len(actual_val.lstrip(' '))
                
                padded_new_text = (" " * leading_spaces_count) + new_text.lstrip(' ')
                # SYNC: Update Original and RU in Excel to the NEW satellite text (with proper padding)
                ws.cell(row_idx, col_map["Original"], padded_new_text)
                if "ru" in col_map:
                    ws.cell(row_idx, col_map["ru"], padded_new_text)
                # Ensure the original text falls into the dictionary mapping but pass the padding forward
                shadok_rows.append((shadok_order_idx, row_idx, leading_spaces_count))
                shadok_order_idx += 1
                shadok_row_set.add(row_idx)
            else:
                print(f"  [SHADOK] WARNING: original line not found in Excel: {orig_text[:40]}...")

        if not shadok_rows:
            print("  [SHADOK] No Shadok lines found in Excel! Skipping.")
        else:
            # Progress tracking (translated_langs)
            translated_langs = set(shadok_config.get("translated_langs", []))
            config_changed = True

            # CLEANUP: If any language in translated_langs has suspiciously few lines in the Excel,
            # remove it from the set so it gets re-translated.
            valid_translated_langs = []
            for lc in sorted(list(translated_langs)):
                if lc not in col_map: continue
                col = col_map[lc]
                non_empty_count = 0
                for _, row_idx, _ in shadok_rows:
                    val = ws.cell(row_idx, col).value
                    if val and str(val).strip():
                        non_empty_count += 1
                # Flexible threshold: CJK languages are more compact
                threshold = 10 if lc in ["zhcn", "zhtw", "jp", "kr"] else 20
                if non_empty_count >= threshold:
                    valid_translated_langs.append(lc)
                else:
                    print(f"  [SHADOK] CLEANUP: Language '{lc}' has only {non_empty_count}/{len(shadok_rows)} lines. Resetting.")
            
            translated_langs = set(valid_translated_langs)
            shadok_config["translated_langs"] = valid_translated_langs
            # Save cleanup results
            with SHADOK_JSON.open("w", encoding="utf-8") as f:
                json.dump(shadok_config, f, ensure_ascii=False, indent=2)

            if not force_all and all(lc in translated_langs for lc in lang_codes):
                print(f"  [SHADOK] All languages already translated. Skipping.")
            else:
                if force_all:
                    langs_to_translate = sorted(lang_codes)
                    translated_langs = set()
                    print(f"  [SHADOK] Force re-translate all {len(langs_to_translate)} languages.")
                    # Clear all shadok cells before re-translating
                    for _, row_idx, _ in shadok_rows:
                        for lc in langs_to_translate:
                            if lc in col_map:
                                ws.cell(row_idx, col_map[lc], "")
                    save_workbook(wb)
                    print(f"  [SHADOK] Cleared {len(shadok_rows)} × {len(langs_to_translate)} cells.")
                else:
                    # Filter out languages actually present in translated_langs
                    langs_to_translate = sorted([lc for lc in lang_codes if lc not in translated_langs])
                    print(f"  [SHADOK] {len(langs_to_translate)} languages remaining: {', '.join(langs_to_translate)}")

                if langs_to_translate:
                    print()
                    print("=" * 60)
                    print("  SHADOK BLOCK TRANSLATION")
                    print("=" * 60)
                    print(f"  Lines: {len(shadok_rows)}, Langs to translate: {len(langs_to_translate)}")
                    print(f"  Max line length: {max_line_len}")
                    print("-" * 60)

                    init_session_shadok()
                    # Use mapped orig text from config directly, NOT from excel since excel holds the satellite now
                    # BUT wait. Where does full_text come from?
                    # It comes from shadok.json! I'll map it natively from the config items
                    mapped_origs = [item["new"].strip() for item in shadok_config.get("mapping", [])]
                    full_text = "\n".join(mapped_origs)

                    try:
                        SHADOK_BATCH_SIZE = 1 
                        MAX_PASSES = 10
                        
                        for pass_idx in range(MAX_PASSES):
                            # Re-calculate remaining languages in every pass
                            remaining_langs = sorted([lc for lc in lang_codes if lc not in translated_langs])
                            if not remaining_langs:
                                break
                                
                            print(f"  [SHADOK] Pass {pass_idx + 1}/{MAX_PASSES}. Remaining: {len(remaining_langs)}")
                            total_batches = len(remaining_langs)

                            for batch_idx, lc in enumerate(remaining_langs):
                                print(f"  [SHADOK] Pass {pass_idx+1}: {lc} ({batch_idx+1}/{total_batches})")

                                try:
                                    batch_results = translate_shadok_block(full_text, [lc], max_line_len)
                                    translated_block = batch_results.get(lc, "").strip()
                                    if not translated_block:
                                        print(f"  [SHADOK][{lc}] Empty translation, skipping.")
                                        continue

                                    # Logical wrapping: handle indented lines as mandatory break points
                                    current_text = translated_block
                                    for i in range(len(shadok_rows)):
                                        order_idx, row_idx, lead_spaces = shadok_rows[i]
                                        
                                        if not current_text.strip():
                                            ws.cell(row_idx, col_map[lc], "")
                                            continue
                                            
                                        # If lead_spaces > 0, we treat this as a start of a new part (like a signature).
                                        # But we also have a limit for the line length.
                                        wrapped = wrap_text(current_text, max_line_len, lc)
                                        if not wrapped:
                                            ws.cell(row_idx, col_map[lc], "")
                                            continue
                                            
                                        line_val = wrapped[0]
                                        padded_val = (" " * lead_spaces) + line_val.lstrip()
                                        
                                        # Clip just in case
                                        if len(padded_val) > max_line_len:
                                            padded_val = padded_val[:max_line_len]
                                            
                                        ws.cell(row_idx, col_map[lc], padded_val)
                                        
                                        # Mark text as consumed
                                        consumed_len = len(line_val)
                                        current_text = current_text[consumed_len:].lstrip()
                                        
                                        # Mandatory Break Logic:
                                        # If the NEXT row in shadok_rows has leading spaces, we must NOT 
                                        # put any more text into the current logical sequence.
                                        if i + 1 < len(shadok_rows):
                                            next_lead_spaces = shadok_rows[i+1][2]
                                            if next_lead_spaces > 0:
                                                # The rest of 'current_text' will start from the next cell 
                                                # because it has indentation.
                                                pass


                                    translated_langs.add(lc)
                                    print(f"  [SHADOK][{lc}] SUCCESS: Written to Excel.")

                                    # Save progress after each successful language
                                    save_workbook(wb)
                                    shadok_config["translated_langs"] = sorted(list(translated_langs))
                                    with SHADOK_JSON.open("w", encoding="utf-8") as f:
                                        json.dump(shadok_config, f, ensure_ascii=False, indent=2)
                                    
                                    time.sleep(0.8)

                                except Exception as batch_error:
                                    print(f"  [SHADOK][{lc}] FAILED: {batch_error}")
                                    continue

                        if not any(lc not in translated_langs for lc in lang_codes):
                            print("  [SHADOK] All languages complete!")
                        else:
                            print(f"  [SHADOK] Finished {MAX_PASSES} passes. Some languages might still be missing.")

                    except Exception as e:
                        print(f"  [SHADOK] Fatal error: {e}")

                    print("=" * 60)
                    print()
                else:
                    print(f"  [SHADOK] No languages to translate. Skipping.")

        # Save updated config at the end to be sure
        with SHADOK_JSON.open("w", encoding="utf-8") as f:
            json.dump(shadok_config, f, ensure_ascii=False, indent=2)


    # ── Scan: count rows that need translation (excluding shadoks) ───
    rows_to_translate = []
    
    # Safety net: collect all known Shadok strings to ensure no duplicates slip into the general loop
    shadok_all_strings = set()
    for item in shadok_config.get("mapping", []):
        shadok_all_strings.add(item["orig"].strip())
        shadok_all_strings.add(item["new"].strip())

    for row in range(2, ws.max_row + 1):
        if row in shadok_row_set:
            continue  # skip strictly mapped shadok rows
            
        actual_val = str(ws.cell(row, col_map["Original"]).value or "")
        if actual_val.strip() in shadok_all_strings:
            # Skip any duplicate or stray Shadok rows so they don't break the general pipeline
            continue
            
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

        # Skip AI translation if string has no Cyrillic characters — copy as-is
        cyrillic_count = len(re.findall(r'[а-яА-ЯёЁіІїЇєЄґҐ]', original))
        if cyrillic_count < 2:
            for lc in missing:
                ws.cell(row, col_map[lc], original)
                total_translated += 1
            save_workbook(wb)
            print(f"    [Row {row} | {idx}/{total_rows}] Cyrillic count {cyrillic_count} <= 3 — copied as-is.")
            continue

        try:
            results = translate_batch(original, missing, row_id=row)
        except Exception as e:
            print(f"  [Row {row} | {idx}/{total_rows}] ERROR: {e}")
            total_failed += len(missing)
            time.sleep(2)
            continue

        # Normalize tokens and full-width chars in AI output
        for lc in list(results.keys()):
            results[lc] = normalize_fullwidth(normalize_tokens_out(results.get(lc, "")))

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
                        refined[lc] = normalize_fullwidth(normalize_tokens_out(refined.get(lc, "")))
                    results.update(refined)
                except Exception as e:
                    print(f"    [Row {row} | {idx}/{total_rows}] Refine error: {e}")
                    break
                time.sleep(0.3)

        # ── Write results ────────────────────────────────────────────
        row_ok = 0
        row_fail = 0
        failed_langs = []
        for lc in missing:
            translation = results.get(lc, "")
            if not translation or not translation.strip():
                failed_langs.append(lc)
                continue
            
            ok, msg = validate(original, translation, lc)
            if ok:
                ws.cell(row, col_map[lc], translation)
                total_translated += 1
                row_ok += 1
            elif "English preservation" in msg:
                ws.cell(row, col_map[lc], original)
                total_translated += 1
                row_ok += 1
            else:
                failed_langs.append(lc)

        # ── Retry failed languages with error context ─────────────────
        MAX_RETRY_ROUNDS = 3
        for retry_round in range(MAX_RETRY_ROUNDS):
            if not failed_langs:
                break

            # Build error context for the AI
            error_details = []
            for lc in failed_langs:
                translation = results.get(lc, "")
                _, msg = validate(original, translation, lc) if translation else (False, "empty")
                error_details.append(f"- {lc}: {msg}")

            error_context = (
                f"Retry round {retry_round + 1}. The following translations for "
                f"\"{original}\" have validation errors:\n"
                + "\n".join(error_details)
                + "\n\nPlease fix these issues. "
                f"Source text: \"{original}\""
            )

            print(f"    [Row {row}] Retry {retry_round + 1}/{MAX_RETRY_ROUNDS}: {len(failed_langs)} langs ({', '.join(failed_langs)})")

            try:
                retry_results = refine(error_context, failed_langs, row_id=row)
                still_failed = []
                for lc in failed_langs:
                    translation = normalize_fullwidth(normalize_tokens_out(retry_results.get(lc, "")))
                    if not translation or not translation.strip():
                        still_failed.append(lc)
                        continue
                    ok, msg = validate(original, translation, lc)
                    if ok:
                        ws.cell(row, col_map[lc], translation)
                        total_translated += 1
                        row_ok += 1
                    else:
                        results[lc] = translation  # update for next round's error context
                        still_failed.append(lc)
                failed_langs = still_failed
            except Exception as e:
                print(f"    [Row {row}] Retry error: {e}")
                break
            time.sleep(0.3)

        # Mark remaining failed langs
        for lc in failed_langs:
            print(f"    [Row {row}][{lc}] SKIPPED (invalid after {MAX_RETRY_ROUNDS} retries)")
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

    # Build shadok exclusion set (use cached rows if available)
    shadok_row_set = set()
    shadok_config = load_shadok_config()
    if shadok_config:
        if "rows" in shadok_config and shadok_config["rows"]:
            shadok_row_set = {r for r in shadok_config["rows"] if r is not None}
        else:
            excel_lookup = {}
            for row in range(2, ws.max_row + 1):
                val = ws.cell(row, col_map["Original"]).value
                if val:
                    excel_lookup[str(val).strip()] = row

            for item in shadok_config.get("mapping", []):
                orig_text = item["orig"].strip()
                new_text = item["new"].strip()
                r = excel_lookup.get(orig_text) or excel_lookup.get(new_text)
                if r:
                    shadok_row_set.add(r)
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

    TRANSLATIONS_DIR.mkdir(parents=True, exist_ok=True)

    missing_total = 0
    
    for lc in langs:
        if lc not in col_map or lc == "ru":
            continue
        csv_path = TRANSLATIONS_DIR / f"{lc}.csv"
        count = 0
        missing_count = 0
        with csv_path.open("w", encoding="utf-8", newline="") as f:
            writer = csv.writer(f)
            writer.writerow(["original", "translation"])
            for row in range(2, ws.max_row + 1):
                original = ws.cell(row, col_map["Original"]).value
                if not original:
                    continue
                    
                translation = ws.cell(row, col_map[lc]).value
                
                if not translation or not str(translation).strip():
                    english_fallback = ws.cell(row, col_map["en"]).value if "en" in col_map else None
                    translation = english_fallback if english_fallback and str(english_fallback).strip() else original
                    missing_count += 1
                    missing_total += 1
                    print(f"  [WARNING] Row {row} missing translation for '{lc}'. Fallback to: {translation[:30] + '...' if len(str(translation)) > 30 else translation}")

                writer.writerow([detokenize(str(original)), detokenize(str(translation))])
                count += 1
                
        print(f"  {lc}.csv: {count} entries" + (f" ({missing_count} missing translations filled with fallback)" if missing_count else ""))

    if missing_total > 0:
        print(f"\n[ALERT] Export finished with {missing_total} missing translations!")
        print("Run `python -m src.main translate` to translate missing lines.")
    else:
        print("\nExport done.")

    print("\n[BUILD] Auto-building binaries...")
    cmd_build()


# ── build ────────────────────────────────────────────────────────────

def cmd_build() -> None:
    """Run build_translation_bin.py for each exported CSV."""
    langs = load_languages()
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

    for lc in langs:
        if lc == "ru":
            continue
        csv_path = TRANSLATIONS_DIR / f"{lc}.csv"
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
                regex = re.compile(pattern, re.IGNORECASE)
            except re.error as e:
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

def cmd_deploy() -> None:
    """Commit, push and create a GitHub release with assets."""
    print("\n" + "="*60)
    print("  STEP: deploy")
    print("="*60)

    # 1. Get versions
    try:
        dbi_ver = get_nro_version() or "unknown"
        wb = open_or_create_workbook()
        patcher_ver = get_version(wb)
    except Exception as e:
        print(f"  [ERROR] Failed to get version: {e}")
        return

    # 2. Git operations
    print("  [GIT] Staging changes and pushing...")
    try:
        subprocess.run(["git", "add", "."], check=True)
        # Check if there are changes to commit
        status = subprocess.run(["git", "status", "--porcelain"], capture_output=True, text=True).stdout
        if status.strip():
            subprocess.run(["git", "commit", "-m", f"chore: deploy DBI {dbi_ver} localization (v{patcher_ver})"], check=True)
            subprocess.run(["git", "push", "origin", "master"], check=True)
            print("  [GIT] Changes pushed successfully.")
        else:
            print("  [GIT] No changes to commit.")
    except subprocess.CalledProcessError as e:
        print(f"  [ERROR] Git operation failed: {e}")
        return

    # 3. Prepare release body
    langs_list = """*   **BE** — Belarusian
*   **DE** — German
*   **EN** — English (US)
*   **ENGB** — English (UK)
*   **ES** — Spanish (Spain)
*   **ES419** — Spanish (Latin America)
*   **ET** — Estonian
*   **FR** — French
*   **FRCA** — French (Canada)
*   **IT** — Italian
*   **JP** — Japanese
*   **KK** — Kazakh
*   **KR** — Korean
*   **LT** — Lithuanian
*   **LV** — Latvian
*   **NL** — Dutch
*   **PL** — Polish
*   **PT** — Portuguese (Portugal)
*   **PTBR** — Portuguese (Brazil)
*   **UA** — Ukrainian
*   **ZHCN** — Chinese (Simplified)
*   **ZHTW** — Chinese (Traditional)"""

    release_body = f"""![GitHub release (tag)](https://img.shields.io/github/downloads/rashevskyv/DBIPatcher/{dbi_ver}/total)

{langs_list}

***

### ⚠️ IMPORTANT: Installation Instruction
To use the translation on your Nintendo Switch:
1. Download the `translation_XX.bin` file for your language from the assets below.
2. **Rename the file to `translation.bin`** (it must be exactly this name).
3. Place `translation.bin` in the same folder where your **DBI.nro** is located.

***
Check the changelog at [https://github.com/rashevskyv/dbi/releases/](https://github.com/rashevskyv/dbi/releases/).
"""
    
    body_path = Path("scratch/release_body.md")
    body_path.parent.mkdir(exist_ok=True)
    body_path.write_text(release_body, encoding="utf-8")

    # 4. GitHub Release
    print(f"  [GH] Preparing release {dbi_ver}...")
    
    # Pre-delete existing release/tag to simulate overwrite (compatible with older gh)
    subprocess.run(["gh", "release", "delete", dbi_ver, "--yes"], capture_output=True)
    subprocess.run(["git", "push", "origin", ":refs/tags/" + dbi_ver], capture_output=True)

    assets = [str(a) for a in Path("output").glob("*.bin")]
    nro_assets = [str(n) for n in Path(".").glob(f"DBI.{dbi_ver}*.nro")]
    
    cmd = [
        "gh", "release", "create", dbi_ver,
        "--title", f"DBI {dbi_ver} Localization",
        "--notes-file", str(body_path)
    ]
    cmd.extend(assets)
    cmd.extend(nro_assets)

    try:
        subprocess.run(cmd, check=True)
        print(f"  [GH] Release {dbi_ver} created successfully!")
    except subprocess.CalledProcessError as e:
        print(f"  [ERROR] GitHub release failed: {e}")


COMMANDS = {
    "sync": cmd_sync,
    "translate": cmd_translate,
    "validate": cmd_validate,
    "align": cmd_align,
    "export": cmd_export,
    "build": cmd_build,
    "clear": cmd_clear,
    "deploy": cmd_deploy,
}


def cmd_all() -> None:
    for name in ("sync", "translate", "align", "validate", "export", "build", "deploy"):
        print(f"\n{'='*60}\n  STEP: {name}\n{'='*60}")
        COMMANDS[name]()


def main() -> int:
    parser = argparse.ArgumentParser(description="DBI Translation Pipeline")
    parser.add_argument("command", choices=[*COMMANDS.keys(), "all"], help="Pipeline step to run")
    parser.add_argument("lang", nargs="?", help="Language code (required for 'clear')")
    parser.add_argument("-f", "--force", action="store_true", help="Force re-translate all strings")
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
