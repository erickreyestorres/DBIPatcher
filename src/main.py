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
from src.core.ai_client import translate_batch

ROOT = Path(__file__).resolve().parent.parent
DATA_DIR = ROOT / "data"
DICT_PATH = DATA_DIR / "dictionary.xlsx"
UA_CSV = DATA_DIR / "ua.csv"
LANG_JSON = DATA_DIR / "languages.json"
OUTPUT_DIR = ROOT / "output"
BUILD_SCRIPT = ROOT / "scripts" / "build_translation_bin.py"

SHEET_NAME = "Translations"
META_SHEET = "Metadata"


# ── helpers ──────────────────────────────────────────────────────────

def load_languages() -> dict[str, str]:
    with LANG_JSON.open("r", encoding="utf-8") as f:
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
    wb.save(DICT_PATH)


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
            ua_raw = rec.get("translation", "") or ""
            if not ru_raw:
                continue
            ru_tok = tokenize(ru_raw)
            if ru_tok in existing:
                # Update UA if empty
                for row in range(2, ws.max_row + 1):
                    if ws.cell(row, col_map["Original"]).value == ru_tok:
                        if not ws.cell(row, col_map.get("ua", 0)).value and ua_raw:
                            ws.cell(row, col_map["ua"], tokenize(ua_raw))
                        break
                continue
            next_row = ws.max_row + 1
            ws.cell(next_row, col_map["Original"], ru_tok)
            if ua_raw and "ua" in col_map:
                ws.cell(next_row, col_map["ua"], tokenize(ua_raw))
            existing.add(ru_tok)
            added += 1

    ver = bump_version(wb)
    save_workbook(wb)
    print(f"Sync done. Added: {added}, Version: {ver}")


# ── translate ────────────────────────────────────────────────────────

def cmd_translate() -> None:
    """Find empty cells and translate via AI."""
    langs = load_languages()
    wb = open_or_create_workbook()
    ws = wb[SHEET_NAME]

    header = [ws.cell(1, c).value for c in range(1, ws.max_column + 1)]
    col_map = {h: i + 1 for i, h in enumerate(header) if h}

    lang_codes = [lc for lc in langs if lc in col_map and lc != "ru"]
    total_translated = 0
    total_failed = 0

    for row in range(2, ws.max_row + 1):
        original = ws.cell(row, col_map["Original"]).value
        if not original:
            continue

        # Find which languages are missing for this row
        missing = []
        for lc in lang_codes:
            cell_val = ws.cell(row, col_map[lc]).value
            if not cell_val or not str(cell_val).strip():
                missing.append(lc)

        if not missing:
            continue

        print(f"  [{row}] Translating: {original[:60]}... -> {missing}")
        try:
            results = translate_batch(original, missing)
        except Exception as e:
            print(f"  [!] AI error: {e}")
            total_failed += len(missing)
            time.sleep(2)
            continue

        for lc in missing:
            translation = results.get(lc, "")
            if translation:
                translation = normalize_tokens_out(translation)
                ws.cell(row, col_map[lc], translation)
                total_translated += 1
            else:
                total_failed += 1

        time.sleep(0.5)  # Rate limit

    ver = bump_version(wb)
    save_workbook(wb)
    print(f"Translate done. OK: {total_translated}, Failed: {total_failed}, Version: {ver}")


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


# ── main ─────────────────────────────────────────────────────────────

COMMANDS = {
    "sync": cmd_sync,
    "translate": cmd_translate,
    "validate": cmd_validate,
    "export": cmd_export,
    "build": cmd_build,
}


def cmd_all() -> None:
    for name in ("sync", "translate", "validate", "export", "build"):
        print(f"\n{'='*60}\n  STEP: {name}\n{'='*60}")
        COMMANDS[name]()


def main() -> int:
    parser = argparse.ArgumentParser(description="DBI Translation Pipeline")
    parser.add_argument("command", choices=[*COMMANDS.keys(), "all"], help="Pipeline step to run")
    args = parser.parse_args()

    if args.command == "all":
        cmd_all()
    else:
        COMMANDS[args.command]()
    return 0


if __name__ == "__main__":
    sys.exit(main())
