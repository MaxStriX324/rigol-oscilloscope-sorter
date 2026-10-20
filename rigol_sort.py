# -*- coding: utf-8 -*-
"""
Rigol DHO sorter for PNG/STP/CSV/BIN files.

Run this script from the Rigol shared folder or use run_rigol_sort.bat.
It groups RigolDS*.png/.stp/.csv/.bin files, archives them by date, converts
CSV/BIN waveform data to XLSX, and assigns a global measurement number N000001.
"""

from __future__ import annotations

import argparse
import csv
import math
import os
import re
import shutil
import struct
import sys
import time
import traceback
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Dict, Iterable, Iterator, List, Optional, Sequence, Tuple

print("rigol_sort.py: Python started, loading modules...", flush=True)
try:
    from openpyxl import Workbook
except ImportError:
    print("ERROR: Python module 'openpyxl' is not installed.", flush=True)
    print("Install it with: py -3 -m pip install openpyxl", flush=True)
    raise
print("rigol_sort.py: openpyxl loaded.", flush=True)


# -------------------- USER SETTINGS --------------------

SETTINGS_DIR_NAME = "Seting"       # Left as on the oscilloscope/share screenshot.
LOGS_DIR_NAME = "_logs"
FAILED_DIR_NAME = "_failed"

SUPPORTED_EXTENSIONS = {".png", ".stp", ".csv", ".bin"}

# Files younger than this are skipped to avoid touching a file while Rigol writes it.
MIN_FILE_AGE_SECONDS = 60.0
STABILITY_DELAY_SECONDS = 1.0

# Excel limits. Excel supports 1,048,576 rows including the header row.
EXCEL_MAX_ROWS = 1_048_576
ROWS_PER_XLSX_PART = 500_000
MAX_ROWS_FOR_SPLIT_XLSX = 3_000_000
PREVIEW_DECIMATED_ROWS = 100_000
PREVIEW_MINMAX_BLOCKS = 5_000

# If True, BIN is converted to XLSX even when CSV exists in the same group.
# In normal Rigol operation you said CSV and BIN are not saved simultaneously,
# so False avoids duplicate XLSX files if that ever happens.
CONVERT_BIN_WHEN_CSV_EXISTS = False

# -------------------------------------------------------

RIGOL_STEM_RE = re.compile(r"^(RigolDS\d+)(?:\(\d+\))?$", re.IGNORECASE)
ARCHIVE_NAME_RE = re.compile(r"_N(\d{1,12})_", re.IGNORECASE)


@dataclass
class FileGroup:
    stem: str
    files: Dict[str, Path]  # extension -> path


@dataclass
class CsvInfo:
    headers: List[str]
    data_rows: int
    has_time_column: bool
    t0: Optional[float]
    tinc: Optional[float]
    delimiter: str


@dataclass
class BinData:
    headers: List[str]
    rows: int
    channels: Dict[str, Sequence[float]]
    time_values: Optional[Sequence[float]]
    span: Optional[float]
    dt: Optional[float]


def now_str() -> str:
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def log(root: Path, message: str) -> None:
    logs = root / LOGS_DIR_NAME
    logs.mkdir(exist_ok=True)
    line = f"[{now_str()}] {message}"
    print(line, flush=True)
    with open(logs / "sort_log.txt", "a", encoding="utf-8") as f:
        f.write(line + "\n")


def get_file_time(path: Path) -> datetime:
    """Creation time on Windows; metadata-change time on Unix. Good enough for Windows/Samba use."""
    ts = path.stat().st_ctime
    return datetime.fromtimestamp(ts)


def safe_move(src: Path, dst: Path, dry_run: bool = False) -> Path:
    dst.parent.mkdir(parents=True, exist_ok=True)
    final = unique_path(dst)
    if dry_run:
        print(f"DRY-RUN move: {src} -> {final}")
        return final
    shutil.move(str(src), str(final))
    return final


def unique_path(path: Path) -> Path:
    if not path.exists():
        return path
    base = path.with_suffix("")
    suffix = path.suffix
    for i in range(1, 10_000):
        candidate = Path(f"{base}_{i:02d}{suffix}")
        if not candidate.exists():
            return candidate
    raise RuntimeError(f"Cannot create unique file name for {path}")


def acquire_lock(root: Path):
    logs = root / LOGS_DIR_NAME
    logs.mkdir(exist_ok=True)
    lock_path = logs / "rigol_sort.lock"
    try:
        fd = os.open(str(lock_path), os.O_CREAT | os.O_EXCL | os.O_RDWR)
        os.write(fd, f"pid={os.getpid()} time={now_str()}\n".encode("utf-8"))
        return fd, lock_path
    except FileExistsError:
        raise RuntimeError(
            f"Lock file exists: {lock_path}. Another sorter may be running. "
            f"If not, delete this file manually."
        )


def release_lock(lock) -> None:
    if not lock:
        return
    fd, lock_path = lock
    try:
        os.close(fd)
    finally:
        try:
            lock_path.unlink()
        except FileNotFoundError:
            pass


def collect_groups(root: Path) -> List[FileGroup]:
    groups: Dict[str, Dict[str, Path]] = {}
    for path in root.iterdir():
        if not path.is_file():
            continue
        ext = path.suffix.lower()
        if ext not in SUPPORTED_EXTENSIONS:
            continue
        m = RIGOL_STEM_RE.match(path.stem)
        if not m:
            continue
        key = path.stem
        groups.setdefault(key, {})[ext] = path
    return [FileGroup(stem=k, files=v) for k, v in sorted(groups.items())]


def ready_groups(groups: List[FileGroup], force: bool = False, min_age_seconds: float = MIN_FILE_AGE_SECONDS) -> List[FileGroup]:
    if force:
        return groups

    all_paths = [p for g in groups for p in g.files.values()]
    if not all_paths:
        return []

    now = time.time()
    size_before = {}
    for p in all_paths:
        try:
            st = p.stat()
        except FileNotFoundError:
            continue
        if now - st.st_mtime < min_age_seconds:
            continue
        size_before[p] = st.st_size

    time.sleep(STABILITY_DELAY_SECONDS)

    ready = []
    for group in groups:
        ok = True
        for p in group.files.values():
            if p not in size_before:
                ok = False
                break
            try:
                if p.stat().st_size != size_before[p]:
                    ok = False
                    break
            except FileNotFoundError:
                ok = False
                break
        if ok:
            ready.append(group)
    return ready


def group_time(group: FileGroup) -> datetime:
    # Data file has priority; STP/PNG may have slightly different timestamps.
    for ext in (".csv", ".bin", ".png", ".stp"):
        p = group.files.get(ext)
        if p:
            return get_file_time(p)
    return datetime.now()


def scan_max_archive_number(root: Path, full_rescan: bool = False) -> int:
    """Find last global N number.

    Fast path: _logs/rigol_counter.txt and _logs/sort_index.csv.
    Full recursive filename scan is used only if needed or requested, because SMB shares
    can be slow when walking a large archive.
    """
    max_n = 0
    logs = root / LOGS_DIR_NAME
    counter_file = logs / "rigol_counter.txt"

    if counter_file.exists():
        try:
            max_n = max(max_n, int(counter_file.read_text(encoding="utf-8").strip()))
        except Exception:
            pass

    index_file = logs / "sort_index.csv"
    if index_file.exists():
        try:
            with open(index_file, "r", encoding="utf-8-sig", errors="replace") as f:
                for line in f:
                    m = re.search(r"N(\d{1,12})", line, re.IGNORECASE)
                    if m:
                        max_n = max(max_n, int(m.group(1)))
        except Exception:
            pass

    # If counter/index exists, do not crawl the whole SMB archive on every run.
    if max_n > 0 and not full_rescan:
        return max_n

    print("Counter not found; scanning archive names for Nxxxxxx...", flush=True)
    scanned_dirs = 0
    for dirpath, dirnames, filenames in os.walk(root):
        rel_parts = set(Path(dirpath).relative_to(root).parts) if Path(dirpath) != root else set()
        if FAILED_DIR_NAME in rel_parts or LOGS_DIR_NAME in rel_parts:
            dirnames[:] = []
            continue
        # Reduce SMB walk: only enter archive date folders and Seting at the root level.
        if Path(dirpath) == root:
            keep = []
            for d in dirnames:
                if d == SETTINGS_DIR_NAME or re.match(r"^\d{4}_\d{2}_\d{2}$", d):
                    keep.append(d)
            dirnames[:] = keep
        scanned_dirs += 1
        if scanned_dirs % 100 == 0:
            print(f"  scanned folders: {scanned_dirs}, current max N={max_n:06d}", flush=True)
        for name in filenames:
            m = ARCHIVE_NAME_RE.search(name)
            if m:
                try:
                    max_n = max(max_n, int(m.group(1)))
                except ValueError:
                    pass
    return max_n

def write_counter(root: Path, value: int) -> None:
    logs = root / LOGS_DIR_NAME
    logs.mkdir(exist_ok=True)
    (logs / "rigol_counter.txt").write_text(str(value), encoding="utf-8")


def make_base_name(dt: datetime, number: int) -> str:
    return f"D{dt:%Y_%m_%d}_N{number:06d}_T{dt:%H_%M_%S}"


def date_folder_name(dt: datetime) -> str:
    return f"{dt:%Y_%m_%d}"


def parse_number(text: str) -> Optional[float]:
    s = str(text).strip().strip('"').strip("'")
    if not s:
        return None
    s = s.replace(" ", "").replace("\u00a0", "")
    # Remove units accidentally appended after a number only when float() fails later.
    if "," in s and "." in s:
        # Last separator is treated as decimal separator.
        if s.rfind(",") > s.rfind("."):
            s = s.replace(".", "").replace(",", ".")
        else:
            s = s.replace(",", "")
    elif "," in s:
        s = s.replace(",", ".")
    try:
        return float(s)
    except ValueError:
        m = re.search(r"[-+]?\d+(?:[\.,]\d+)?(?:[eE][-+]?\d+)?", s)
        if not m:
            return None
        return parse_number(m.group(0))


def detect_delimiter(sample: str) -> str:
    first = ""
    for line in sample.splitlines():
        if line.strip():
            first = line
            break
    if not first:
        return ","
    candidates = [",", ";", "\t"]
    counts = {c: first.count(c) for c in candidates}
    # Semicolon is preferred when decimal comma is used.
    best = max(candidates, key=lambda c: counts[c])
    return best if counts[best] > 0 else ","


def open_text_csv(path: Path):
    # utf-8-sig handles BOM; errors replace avoids death on odd symbols.
    return open(path, "r", newline="", encoding="utf-8-sig", errors="replace")


def analyze_csv(path: Path) -> CsvInfo:
    with open(path, "r", encoding="utf-8-sig", errors="replace") as f:
        sample = f.read(4096)
    delimiter = detect_delimiter(sample)

    with open_text_csv(path) as f:
        reader = csv.reader(f, delimiter=delimiter)
        raw_header = next(reader, [])

        t0 = None
        tinc = None
        headers: List[str] = []
        has_time = False

        for token in raw_header:
            t = token.strip().strip('"')
            low = t.lower().replace(" ", "")
            if low.startswith("t0="):
                t0 = parse_number(t.split("=", 1)[1])
                continue
            if low.startswith("tinc="):
                tinc = parse_number(t.split("=", 1)[1])
                continue
            if low in {"time(s)", "time", "x"} or low.startswith("time"):
                headers.append("Time(s)")
                has_time = True
                continue
            if re.match(r"^ch\d+v?$", low, re.IGNORECASE):
                name = t.upper()
                if not name.endswith("V"):
                    name += "V"
                headers.append(name)
                continue
            if t:
                # Keep unknown columns, but avoid metadata garbage.
                if "=" not in t:
                    headers.append(t)

        if not headers:
            # Fallback for headerless files.
            first_data = next(reader, [])
            n = len(first_data)
            headers = [f"COL{i+1}" for i in range(n)]
            data_rows = 1 + sum(1 for _ in reader)
            return CsvInfo(headers, data_rows, False, t0, tinc, delimiter)

        data_rows = 0
        last_report = time.time()
        for row in reader:
            if not row or not any(cell.strip() for cell in row):
                continue
            data_rows += 1
            if data_rows % 200000 == 0 or (time.time() - last_report) > 10:
                print(f"  CSV analyze: {path.name}: {data_rows:,} rows counted", flush=True)
                last_report = time.time()

    if not has_time and t0 is not None and tinc is not None:
        headers = ["Time(s)"] + headers
        has_time = True

    return CsvInfo(headers, data_rows, has_time, t0, tinc, delimiter)


def iter_csv_rows(path: Path, info: CsvInfo) -> Iterator[List[Optional[float]]]:
    with open_text_csv(path) as f:
        reader = csv.reader(f, delimiter=info.delimiter)
        raw_header = next(reader, [])

        # Number of physical data columns expected from file excluding generated Time(s).
        generated_time = info.t0 is not None and info.tinc is not None and raw_header and not any(
            token.strip().lower().replace(" ", "").startswith("time") for token in raw_header
        )
        physical_cols = len(info.headers) - (1 if generated_time else 0)

        idx = 0
        for row in reader:
            if not row or not any(cell.strip() for cell in row):
                continue
            values = [parse_number(c) for c in row[:physical_cols]]
            if generated_time:
                t = info.t0 + idx * info.tinc  # type: ignore[operator]
                out = [t] + values
            else:
                out = values[: len(info.headers)]
            if len(out) < len(info.headers):
                out.extend([None] * (len(info.headers) - len(out)))
            yield out
            idx += 1


def setup_sheet(ws, headers: Sequence[str]) -> None:
    ws.freeze_panes = "A2"
    for idx, header in enumerate(headers, start=1):
        width = 14
        if header.lower().startswith("time"):
            width = 18
        ws.column_dimensions[chr(64 + idx) if idx <= 26 else None].width = width if idx <= 26 else width


def save_workbook(wb: Workbook, path: Path, dry_run: bool = False) -> Path:
    final = unique_path(path)
    if dry_run:
        print(f"DRY-RUN save xlsx: {final}")
        return final
    final.parent.mkdir(parents=True, exist_ok=True)
    wb.save(final)
    return final


def write_full_xlsx(headers: Sequence[str], rows: Iterable[Sequence[Optional[float]]], out_path: Path, dry_run: bool = False) -> Path:
    wb = Workbook(write_only=True)
    ws = wb.create_sheet("Data")
    ws.append(list(headers))
    row_count = 0
    last_report = time.time()
    for row in rows:
        ws.append(list(row))
        row_count += 1
        if row_count % 100000 == 0 or (time.time() - last_report) > 10:
            print(f"  XLSX write: {out_path.name}: {row_count:,} rows", flush=True)
            last_report = time.time()
    return save_workbook(wb, out_path, dry_run=dry_run)


def write_split_xlsx(headers: Sequence[str], rows: Iterable[Sequence[Optional[float]]], out_base: Path, dry_run: bool = False) -> List[Path]:
    outputs: List[Path] = []
    part = 1
    wb: Optional[Workbook] = None
    ws = None
    count_in_part = 0

    def new_book():
        nwb = Workbook(write_only=True)
        nws = nwb.create_sheet("Data")
        nws.append(list(headers))
        return nwb, nws

    wb, ws = new_book()
    for row in rows:
        if count_in_part >= ROWS_PER_XLSX_PART:
            part_path = out_base.with_name(f"{out_base.stem}_P{part:03d}.xlsx")
            outputs.append(save_workbook(wb, part_path, dry_run=dry_run))  # type: ignore[arg-type]
            part += 1
            wb, ws = new_book()
            count_in_part = 0
        ws.append(list(row))  # type: ignore[union-attr]
        count_in_part += 1
        if count_in_part % 100000 == 0:
            print(f"  XLSX split write: part {part:03d}: {count_in_part:,} rows", flush=True)

    if count_in_part > 0:
        part_path = out_base.with_name(f"{out_base.stem}_P{part:03d}.xlsx")
        outputs.append(save_workbook(wb, part_path, dry_run=dry_run))  # type: ignore[arg-type]
    return outputs


def write_preview_xlsx_csv(path: Path, info: CsvInfo, out_path: Path, dry_run: bool = False) -> Path:
    total = max(info.data_rows, 1)
    dec_step = max(1, math.ceil(total / PREVIEW_DECIMATED_ROWS))
    block_size = max(1, math.ceil(total / PREVIEW_MINMAX_BLOCKS))

    wb = Workbook(write_only=True)
    ws_dec = wb.create_sheet("Preview_decimated")
    ws_minmax = wb.create_sheet("MinMax_blocks")
    ws_info = wb.create_sheet("Info")

    ws_dec.append(info.headers)
    mm_headers = ["Block", "Start_index", "End_index"]
    if info.headers and info.headers[0].lower().startswith("time"):
        mm_headers += ["Time_start(s)", "Time_end(s)"]
        channel_headers = info.headers[1:]
        channel_offset = 1
    else:
        channel_headers = info.headers
        channel_offset = 0
    for h in channel_headers:
        mm_headers += [f"{h}_min", f"{h}_max"]
    ws_minmax.append(mm_headers)

    mins = [None] * len(channel_headers)
    maxs = [None] * len(channel_headers)
    block_start_idx = 0
    block_start_time = None
    last_time = None
    block_no = 1

    last_report = time.time()
    for idx, row in enumerate(iter_csv_rows(path, info)):
        if idx and (idx % 200000 == 0 or (time.time() - last_report) > 10):
            print(f"  Preview build: {path.name}: {idx:,}/{info.data_rows:,} rows", flush=True)
            last_report = time.time()
        if idx % dec_step == 0:
            ws_dec.append(list(row))

        if channel_offset:
            if block_start_time is None:
                block_start_time = row[0]
            last_time = row[0]
        for c, value in enumerate(row[channel_offset:channel_offset + len(channel_headers)]):
            if value is None:
                continue
            if mins[c] is None or value < mins[c]:
                mins[c] = value
            if maxs[c] is None or value > maxs[c]:
                maxs[c] = value

        if (idx + 1) % block_size == 0:
            mm_row = [block_no, block_start_idx, idx]
            if channel_offset:
                mm_row += [block_start_time, last_time]
            for mn, mx in zip(mins, maxs):
                mm_row += [mn, mx]
            ws_minmax.append(mm_row)
            block_no += 1
            block_start_idx = idx + 1
            block_start_time = None
            mins = [None] * len(channel_headers)
            maxs = [None] * len(channel_headers)

    if any(v is not None for v in mins + maxs):
        mm_row = [block_no, block_start_idx, info.data_rows - 1]
        if channel_offset:
            mm_row += [block_start_time, last_time]
        for mn, mx in zip(mins, maxs):
            mm_row += [mn, mx]
        ws_minmax.append(mm_row)

    ws_info.append(["Parameter", "Value"])
    ws_info.append(["Source", path.name])
    ws_info.append(["Rows", info.data_rows])
    ws_info.append(["Delimiter", repr(info.delimiter)])
    ws_info.append(["t0", info.t0])
    ws_info.append(["tInc", info.tinc])
    ws_info.append(["Decimation step", dec_step])
    ws_info.append(["MinMax block size", block_size])
    ws_info.append(["Note", "Full source CSV is archived unchanged. Preview is reduced for Excel usability."])

    return save_workbook(wb, out_path, dry_run=dry_run)


def convert_csv_to_xlsx(path: Path, out_dir: Path, base_name: str, dry_run: bool = False) -> Tuple[List[Path], CsvInfo]:
    print(f"CSV: analyzing {path.name} ({path.stat().st_size / (1024*1024):.1f} MB)...", flush=True)
    info = analyze_csv(path)
    print(f"CSV: {path.name}: rows={info.data_rows:,}, columns={len(info.headers)}, delimiter={repr(info.delimiter)}", flush=True)
    out_base = out_dir / f"{base_name}.xlsx"

    if info.data_rows + 1 <= EXCEL_MAX_ROWS and info.data_rows <= ROWS_PER_XLSX_PART:
        print("CSV mode: full XLSX", flush=True)
        outputs = [write_full_xlsx(info.headers, iter_csv_rows(path, info), out_base, dry_run=dry_run)]
    elif info.data_rows <= MAX_ROWS_FOR_SPLIT_XLSX:
        print("CSV mode: split XLSX parts", flush=True)
        outputs = write_split_xlsx(info.headers, iter_csv_rows(path, info), out_base, dry_run=dry_run)
    else:
        print("CSV mode: large file, preview XLSX + full CSV archive", flush=True)
        preview = out_dir / f"{base_name}_preview.xlsx"
        outputs = [write_preview_xlsx_csv(path, info, preview, dry_run=dry_run)]

    info_txt = out_dir / f"{base_name}_info.txt"
    if not dry_run:
        info_txt.write_text(
            "Rigol CSV conversion info\n"
            f"source={path.name}\n"
            f"rows={info.data_rows}\n"
            f"headers={', '.join(info.headers)}\n"
            f"delimiter={repr(info.delimiter)}\n"
            f"t0={info.t0}\n"
            f"tInc={info.tinc}\n",
            encoding="utf-8",
        )
    else:
        print(f"DRY-RUN save info: {info_txt}")
    return outputs, info


def parse_rigol_bin(path: Path) -> BinData:
    data = path.read_bytes()
    if len(data) < 32 or data[:4] != b"RG03":
        raise ValueError("Unsupported BIN format: expected RG03 header")
    channel_count = struct.unpack_from("<I", data, 12)[0]
    if not (1 <= channel_count <= 8):
        raise ValueError(f"Suspicious channel count in BIN: {channel_count}")

    offset = 16
    channels: Dict[str, Sequence[float]] = {}
    rows = None
    span = None

    for _ in range(channel_count):
        if offset + 156 > len(data):
            raise ValueError("BIN channel block is truncated")
        header_size = struct.unpack_from("<I", data, offset)[0]
        points = struct.unpack_from("<I", data, offset + 12)[0]
        block_span = struct.unpack_from("<f", data, offset + 20)[0]
        name_bytes = data[offset + 112: offset + 132]
        name = name_bytes.split(b"\x00", 1)[0].decode("ascii", errors="ignore").strip()
        if not name:
            name = f"CH{len(channels)+1}"
        if not name.upper().endswith("V"):
            name = name.upper() + "V"

        data_size = struct.unpack_from("<I", data, offset + 148)[0]
        data_offset = offset + 156
        expected_size = points * 4
        if data_size < expected_size:
            raise ValueError(f"BIN data block too small for {name}: {data_size} < {expected_size}")
        if data_offset + expected_size > len(data):
            raise ValueError(f"BIN data block truncated for {name}")

        fmt = f"<{points}f"
        values = struct.unpack_from(fmt, data, data_offset)
        channels[name] = values
        rows = points if rows is None else min(rows, points)
        if span is None and block_span > 0:
            span = float(block_span)

        # In tested DHO814 files next channel starts immediately after data_size bytes.
        # header_size is 140; data actually begins at offset + 156.
        offset = data_offset + data_size

    if rows is None:
        rows = 0

    time_values = None
    dt = None
    if span and rows > 0:
        dt = span / rows
        t0 = -span / 2.0 + dt / 2.0
        time_values = [t0 + i * dt for i in range(rows)]
        headers = ["Time(s)"] + list(channels.keys())
    else:
        headers = ["Sample"] + list(channels.keys())

    return BinData(headers=headers, rows=rows, channels=channels, time_values=time_values, span=span, dt=dt)


def iter_bin_rows(bin_data: BinData) -> Iterator[List[Optional[float]]]:
    channel_names = list(bin_data.channels.keys())
    for i in range(bin_data.rows):
        if bin_data.time_values is not None:
            row = [bin_data.time_values[i]]
        else:
            row = [i]
        for name in channel_names:
            vals = bin_data.channels[name]
            row.append(vals[i] if i < len(vals) else None)
        yield row


def write_preview_xlsx_bin(bin_data: BinData, out_path: Path, dry_run: bool = False) -> Path:
    total = max(bin_data.rows, 1)
    dec_step = max(1, math.ceil(total / PREVIEW_DECIMATED_ROWS))
    block_size = max(1, math.ceil(total / PREVIEW_MINMAX_BLOCKS))
    headers = bin_data.headers
    channel_names = list(bin_data.channels.keys())
    channel_offset = 1

    wb = Workbook(write_only=True)
    ws_dec = wb.create_sheet("Preview_decimated")
    ws_minmax = wb.create_sheet("MinMax_blocks")
    ws_info = wb.create_sheet("Info")

    ws_dec.append(headers)
    first_col = "Time(s)" if bin_data.time_values is not None else "Sample"
    mm_headers = ["Block", "Start_index", "End_index", f"{first_col}_start", f"{first_col}_end"]
    for h in channel_names:
        mm_headers += [f"{h}_min", f"{h}_max"]
    ws_minmax.append(mm_headers)

    mins = [None] * len(channel_names)
    maxs = [None] * len(channel_names)
    block_start_idx = 0
    block_start_x = None
    last_x = None
    block_no = 1

    for idx, row in enumerate(iter_bin_rows(bin_data)):
        if idx % dec_step == 0:
            ws_dec.append(row)
        x = row[0]
        if block_start_x is None:
            block_start_x = x
        last_x = x
        for c, value in enumerate(row[channel_offset:]):
            if value is None:
                continue
            if mins[c] is None or value < mins[c]:
                mins[c] = value
            if maxs[c] is None or value > maxs[c]:
                maxs[c] = value
        if (idx + 1) % block_size == 0:
            mm_row = [block_no, block_start_idx, idx, block_start_x, last_x]
            for mn, mx in zip(mins, maxs):
                mm_row += [mn, mx]
            ws_minmax.append(mm_row)
            block_no += 1
            block_start_idx = idx + 1
            block_start_x = None
            mins = [None] * len(channel_names)
            maxs = [None] * len(channel_names)

    if any(v is not None for v in mins + maxs):
        mm_row = [block_no, block_start_idx, bin_data.rows - 1, block_start_x, last_x]
        for mn, mx in zip(mins, maxs):
            mm_row += [mn, mx]
        ws_minmax.append(mm_row)

    ws_info.append(["Parameter", "Value"])
    ws_info.append(["Rows", bin_data.rows])
    ws_info.append(["Headers", ", ".join(headers)])
    ws_info.append(["span", bin_data.span])
    ws_info.append(["dt", bin_data.dt])
    ws_info.append(["Decimation step", dec_step])
    ws_info.append(["MinMax block size", block_size])
    return save_workbook(wb, out_path, dry_run=dry_run)


def convert_bin_to_xlsx(path: Path, out_dir: Path, base_name: str, dry_run: bool = False) -> Tuple[List[Path], BinData]:
    print(f"BIN: parsing {path.name} ({path.stat().st_size / 1024:.1f} KB)...", flush=True)
    bin_data = parse_rigol_bin(path)
    print(f"BIN: {path.name}: rows={bin_data.rows:,}, channels={len(bin_data.channels)}, dt={bin_data.dt}", flush=True)
    out_base = out_dir / f"{base_name}.xlsx"
    if bin_data.rows + 1 <= EXCEL_MAX_ROWS and bin_data.rows <= ROWS_PER_XLSX_PART:
        outputs = [write_full_xlsx(bin_data.headers, iter_bin_rows(bin_data), out_base, dry_run=dry_run)]
    elif bin_data.rows <= MAX_ROWS_FOR_SPLIT_XLSX:
        outputs = write_split_xlsx(bin_data.headers, iter_bin_rows(bin_data), out_base, dry_run=dry_run)
    else:
        outputs = [write_preview_xlsx_bin(bin_data, out_dir / f"{base_name}_preview.xlsx", dry_run=dry_run)]

    info_txt = out_dir / f"{base_name}_bin_info.txt"
    if not dry_run:
        info_txt.write_text(
            "Rigol BIN conversion info\n"
            f"source={path.name}\n"
            f"rows={bin_data.rows}\n"
            f"headers={', '.join(bin_data.headers)}\n"
            f"span={bin_data.span}\n"
            f"dt={bin_data.dt}\n",
            encoding="utf-8",
        )
    else:
        print(f"DRY-RUN save info: {info_txt}")
    return outputs, bin_data


def append_index(root: Path, row: Dict[str, str], dry_run: bool = False) -> None:
    logs = root / LOGS_DIR_NAME
    logs.mkdir(exist_ok=True)
    path = logs / "sort_index.csv"
    fields = [
        "archive_number", "archive_base", "source_stem", "date", "time", "status",
        "data_type", "points", "channels", "outputs", "note",
    ]
    if dry_run:
        print("DRY-RUN index row:", row)
        return
    exists = path.exists()
    with open(path, "a", newline="", encoding="utf-8-sig") as f:
        writer = csv.DictWriter(f, fieldnames=fields, delimiter=";")
        if not exists:
            writer.writeheader()
        writer.writerow({k: row.get(k, "") for k in fields})


def move_failed(root: Path, group: FileGroup, reason: str) -> None:
    failed_root = root / FAILED_DIR_NAME / f"{group.stem}_{datetime.now():%Y_%m_%d_%H_%M_%S}"
    failed_root.mkdir(parents=True, exist_ok=True)
    for p in group.files.values():
        if p.exists():
            try:
                shutil.move(str(p), str(failed_root / p.name))
            except Exception:
                pass
    (failed_root / "error.txt").write_text(reason, encoding="utf-8")


def process_group(root: Path, group: FileGroup, number: int, dry_run: bool = False) -> None:
    dt = group_time(group)
    folder = root / date_folder_name(dt)
    settings_folder = root / SETTINGS_DIR_NAME / date_folder_name(dt)
    base = make_base_name(dt, number)
    outputs: List[Path] = []
    data_type = ""
    points = ""
    channels = ""
    note = ""

    log(root, f"Processing {group.stem} -> {base}")

    try:
        folder.mkdir(parents=True, exist_ok=True)
        settings_folder.mkdir(parents=True, exist_ok=True)

        # Convert data before moving source files. If conversion fails, source remains intact.
        if ".csv" in group.files:
            csv_outputs, info = convert_csv_to_xlsx(group.files[".csv"], folder, base, dry_run=dry_run)
            outputs.extend(csv_outputs)
            data_type = "CSV"
            points = str(info.data_rows)
            channels = ",".join([h for h in info.headers if h.upper().startswith("CH")])
            if info.data_rows > MAX_ROWS_FOR_SPLIT_XLSX:
                note = "large CSV: preview XLSX created; full CSV archived"

        if ".bin" in group.files and (CONVERT_BIN_WHEN_CSV_EXISTS or ".csv" not in group.files):
            suffix_base = base if ".csv" not in group.files else base + "_bin"
            bin_outputs, bin_info = convert_bin_to_xlsx(group.files[".bin"], folder, suffix_base, dry_run=dry_run)
            outputs.extend(bin_outputs)
            data_type = "BIN" if not data_type else data_type + "+BIN"
            points = str(bin_info.rows) if not points else points
            channels = ",".join([h for h in bin_info.headers if h.upper().startswith("CH")]) or channels
            if bin_info.rows > MAX_ROWS_FOR_SPLIT_XLSX:
                note = "large BIN: preview XLSX created; full BIN archived"

        # Move original files.
        for ext, src in sorted(group.files.items()):
            if ext == ".stp":
                dst = settings_folder / f"{base}{ext}"
            else:
                dst = folder / f"{base}{ext}"
            outputs.append(safe_move(src, dst, dry_run=dry_run))

        append_index(root, {
            "archive_number": f"N{number:06d}",
            "archive_base": base,
            "source_stem": group.stem,
            "date": f"{dt:%Y_%m_%d}",
            "time": f"{dt:%H_%M_%S}",
            "status": "OK",
            "data_type": data_type,
            "points": points,
            "channels": channels,
            "outputs": " | ".join(str(p.relative_to(root)) if p.is_absolute() else str(p) for p in outputs),
            "note": note,
        }, dry_run=dry_run)

    except Exception as e:
        err = traceback.format_exc()
        log(root, f"ERROR processing {group.stem}: {e}")
        append_index(root, {
            "archive_number": f"N{number:06d}",
            "archive_base": base,
            "source_stem": group.stem,
            "date": f"{dt:%Y_%m_%d}",
            "time": f"{dt:%H_%M_%S}",
            "status": "ERROR",
            "note": str(e),
        }, dry_run=dry_run)
        if not dry_run:
            move_failed(root, group, err)
        raise


def sort_once(root: Path, force: bool = False, dry_run: bool = False,
              rescan_counter: bool = False, min_age_seconds: float = MIN_FILE_AGE_SECONDS,
              verbose_empty: bool = True, continue_on_error: bool = False) -> tuple[int, int]:
    """Process eligible captures once. Returns (processed_groups, failed_groups).

    The watch loop and one-shot mode share the same grouping and processing logic.
    A group becomes eligible only when all currently present files are old enough
    and their sizes remain stable through the final check.
    """
    groups = collect_groups(root)
    if not groups:
        if verbose_empty:
            log(root, "No RigolDS*.png/.stp/.csv/.bin files in root folder.")
        return 0, 0

    groups = ready_groups(groups, force=force, min_age_seconds=min_age_seconds)
    groups = sorted(groups, key=lambda g: (group_time(g), g.stem.lower()))
    if not groups:
        if verbose_empty:
            log(root, "No eligible files: waiting for age/stability checks.")
        return 0, 0

    max_n = scan_max_archive_number(root, full_rescan=rescan_counter)
    next_n = max_n + 1
    processed = 0
    failed = 0
    for group in groups:
        try:
            process_group(root, group, next_n, dry_run=dry_run)
        except Exception as exc:
            failed += 1
            if not continue_on_error:
                raise
            log(root, f"Group failed ({group.stem}): {exc}")
        else:
            if not dry_run:
                write_counter(root, next_n)
            processed += 1
        # Keep numbering monotonic within each batch even if a group failed.
        next_n += 1
    log(root, f"Batch complete. Processed: {processed}; failed: {failed}")
    return processed, failed


def main() -> int:
    parser = argparse.ArgumentParser(description="Sort Rigol DHO PNG/STP/CSV/BIN files and convert data to XLSX.")
    parser.add_argument("--root", default=None, help="Rigol folder path. Default: current folder.")
    parser.add_argument("--dry-run", action="store_true", help="Show actions without moving/converting files.")
    parser.add_argument("--force", action="store_true", help="Process files without age/stability checks (use cautiously).")
    parser.add_argument("--rescan-counter", action="store_true", help="Recover last N from archived filenames. Slower on SMB.")
    parser.add_argument("--watch", action="store_true", help="Continuously scan the SMB folder and archive captures when ready.")
    parser.add_argument("--min-age-seconds", type=float, default=MIN_FILE_AGE_SECONDS,
                        help="Seconds since last modification of EVERY file in a group (default 60).")
    parser.add_argument("--poll-seconds", type=float, default=5.0,
                        help="Scan interval in watch mode (default 5 seconds).")
    args = parser.parse_args()
    if args.min_age_seconds < 0 or not 0.2 <= args.poll_seconds <= 3600:
        parser.error("--min-age-seconds must be >= 0; --poll-seconds must be 0.2..3600")

    # Preserve the existing Windows/UNC path handling from v1.
    if args.root:
        root_text = str(args.root).strip().strip('"')
        root_text = root_text.replace(':"', ':\\')
        if len(root_text) == 2 and root_text[1] == ':':
            root_text += '\\'
        root = Path(root_text)
    else:
        root = Path.cwd()
    if not root.exists() or not root.is_dir():
        print(f"ERROR: root folder does not exist: {root}", flush=True)
        return 2

    lock = None
    try:
        lock = acquire_lock(root)
        log(root, f"Rigol sorter started. Root: {root}; age={args.min_age_seconds:g}s; "
                  f"watch={args.watch}; poll={args.poll_seconds:g}s")
        if args.watch and args.force:
            parser.error("--force and --watch must not be used together")
        if args.watch and args.dry_run:
            parser.error("--dry-run and --watch must not be used together")

        if not args.watch:
            _, failed = sort_once(root, force=args.force, dry_run=args.dry_run,
                                  rescan_counter=args.rescan_counter,
                                  min_age_seconds=args.min_age_seconds)
            return 1 if failed else 0

        log(root, "Watch mode active. Press Ctrl+C to stop.")
        while True:
            try:
                sort_once(root, min_age_seconds=args.min_age_seconds,
                          verbose_empty=False, continue_on_error=True,
                          rescan_counter=args.rescan_counter)
            except OSError as exc:
                # Temporary SMB failures should not kill the watch loop.
                log(root, f"SMB/IO error: {exc}. Will retry.")
            time.sleep(args.poll_seconds)
    except KeyboardInterrupt:
        log(root, "Stopped by user.")
        return 0
    except Exception as exc:
        print("\nERROR:", exc, flush=True)
        print("See _logs\\sort_log.txt for details.", flush=True)
        return 1
    finally:
        release_lock(lock)


if __name__ == "__main__":
    raise SystemExit(main())
