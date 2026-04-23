#!/usr/bin/env python3
"""Build a DBI translation.bin file from a CSV.

  Header (0x20 bytes):
    [+0x00]  u64  magic       "DBITRNS\\0"
    [+0x08]  u32  version     (= 2)
    [+0x0C]  u32  entry_count
    [+0x10]  u32  entry_size
    [+0x14]  u32  ru_data_off (= 4)
    [+0x18]  u32  tr_data_off
    [+0x1C]  u32  reserved    (= 0)

  Entry (entry_size bytes, repeated entry_count times):
    [+0x00]          u16   ru_len
    [+0x02]          u16   tr_len
    [+ru_data_off]   bytes ru_utf8  (null-padded)
    [+tr_data_off]   bytes tr_utf8  (null-padded)
"""

from __future__ import annotations

import argparse
import csv
import re
import struct
import sys
from pathlib import Path


TABLE_MAGIC = b"DBITRNS\x00"
TABLE_VERSION = 2
TABLE_HEADER_SIZE = 0x20
RU_DATA_OFF = 4


def _unescape(s: str) -> str:
    """Expand \\n, \\r, \\t, \\xHH escape sequences in a CSV cell."""
    s = s.replace("\\n", "\n").replace("\\r", "\r").replace("\\t", "\t")
    s = re.sub(r"\\x([0-9a-fA-F]{2})", lambda m: chr(int(m.group(1), 16)), s)
    return s


def load_pairs_csv(csv_path: Path) -> list[tuple[bytes, bytes]]:
    """Load unique (ru_utf8, tr_utf8) pairs from a CSV file.

    Skips rows where:
      - original or translation is empty
      - translation equals original (identity / nothing to do)
      - the Russian side has already been seen (dedup)
    """
    pairs: list[tuple[bytes, bytes]] = []
    seen: set[bytes] = set()
    skipped_empty = 0
    skipped_identity = 0
    skipped_dupes = 0

    with csv_path.open("r", encoding="utf-8", newline="") as f:
        reader = csv.DictReader(f)
        if not reader.fieldnames or "original" not in reader.fieldnames or "translation" not in reader.fieldnames:
            raise SystemExit(
                f"{csv_path}: CSV must have 'original' and 'translation' columns. "
                f"Got: {reader.fieldnames}"
            )

        for row in reader:
            ru_text = _unescape(row.get("original", "") or "")
            tr_text = _unescape(row.get("translation", "") or "")
            if not ru_text or not tr_text:
                skipped_empty += 1
                continue
            ru_bytes = ru_text.encode("utf-8")
            tr_bytes = tr_text.encode("utf-8")
            if ru_bytes == tr_bytes:
                skipped_identity += 1
                continue
            if ru_bytes in seen:
                skipped_dupes += 1
                continue
            seen.add(ru_bytes)
            pairs.append((ru_bytes, tr_bytes))

    print(f"Loaded pairs           : {len(pairs)}")
    print(f"Skipped (empty cell)   : {skipped_empty}")
    print(f"Skipped (identity)     : {skipped_identity}")
    print(f"Skipped (duplicate ru) : {skipped_dupes}")
    return pairs


def build_table(pairs: list[tuple[bytes, bytes]]) -> tuple[bytes, int, int]:
    """Pack (ru, tr) pairs into translation.bin bytes.

    Returns (table_bytes, entry_size, tr_data_off).
    """
    if not pairs:
        raise SystemExit("No translation pairs to build.")

    max_ru = max(len(ru) for ru, _ in pairs)
    max_tr = max(len(tr) for _, tr in pairs)

    # Align Russian and translation data regions to 4 bytes so subsequent
    # entries remain naturally aligned.
    max_ru_aligned = (max_ru + 3) & ~3
    max_tr_aligned = (max_tr + 3) & ~3

    ru_data_off = RU_DATA_OFF
    tr_data_off = ru_data_off + max_ru_aligned
    entry_size = tr_data_off + max_tr_aligned

    if max_ru > 0xFFFF or max_tr > 0xFFFF:
        raise SystemExit(
            f"String too long for u16 length field: max_ru={max_ru}, max_tr={max_tr}"
        )

    header = bytearray(TABLE_HEADER_SIZE)
    header[0:8] = TABLE_MAGIC
    struct.pack_into(
        "<IIIIII",
        header, 0x08,
        TABLE_VERSION,
        len(pairs),
        entry_size,
        ru_data_off,
        tr_data_off,
        0,
    )

    buf = bytearray(header)
    for ru, tr in pairs:
        entry = bytearray(entry_size)
        struct.pack_into("<HH", entry, 0, len(ru), len(tr))
        entry[ru_data_off : ru_data_off + len(ru)] = ru
        entry[tr_data_off : tr_data_off + len(tr)] = tr
        buf.extend(entry)

    return bytes(buf), entry_size, tr_data_off


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("csv", type=Path, help="Input CSV with 'original,translation' columns")
    parser.add_argument(
        "-o", "--output",
        type=Path,
        default=Path("translation.bin"),
        help="Output path for translation.bin (default: ./translation.bin)",
    )
    args = parser.parse_args()

    if not args.csv.exists():
        print(f"CSV not found: {args.csv}", file=sys.stderr)
        return 2

    print(f"Source CSV : {args.csv}")
    pairs = load_pairs_csv(args.csv)

    table_bytes, entry_size, tr_data_off = build_table(pairs)

    args.output.parent.mkdir(parents=True, exist_ok=True)
    args.output.write_bytes(table_bytes)

    print()
    print(f"Wrote      : {args.output}")
    print(f"Entries    : {len(pairs)}")
    print(f"Entry size : {entry_size} bytes")
    print(f"ru_data_off: {RU_DATA_OFF}")
    print(f"tr_data_off: {tr_data_off}")
    print(f"Total size : {len(table_bytes)} bytes")
    print()
    print("Copy to SD card: sdmc:/switch/DBI/translation.bin")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
