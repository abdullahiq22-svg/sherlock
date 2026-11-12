# -*- coding: utf-8 -*-
"""
Engineering Drawing & Document Control System
=============================================
A local, single-file drawing register for MEP / construction projects.

    streamlit run app.py

Features
--------
* Multi-format indexer (PDF, DWG, DXF, Excel, Word, CSV) with incremental,
  multi-threaded scanning and automatic clean-up of deleted files.
* Smart title-block parsing for PDFs: drawing number, title, revision, date.
* ISO 19650-style revision ranking (A, B ... P01, P02 ... 00, 01 ... C01).
* Current / Superseded status per drawing number, duplicate detection
  (content fingerprint) and revision-conflict detection.
* Token search ("GF LTG 02" finds "GS-E-LTG-GF-02"), folder / type / revision
  / status filters, pagination, card and table views.
* One-click Excel / CSV drawing-register export.
"""
from __future__ import annotations

import hashlib
import inspect
import io
import os
import platform
import re
import sqlite3
import subprocess
import time
from concurrent.futures import ThreadPoolExecutor, as_completed
from contextlib import closing
from datetime import datetime

import numpy as np
import pandas as pd
import streamlit as st

try:  # PyMuPDF >= 1.24 exposes the new name, older versions only "fitz"
    import pymupdf as fitz
except ImportError:  # pragma: no cover
    import fitz  # type: ignore

try:
    import docx  # python-docx
except ImportError:  # pragma: no cover
    docx = None

try:
    import openpyxl
except ImportError:  # pragma: no cover
    openpyxl = None


# ==========================================
# 0. Configuration
# ==========================================
APP_TITLE = "Engineering Drawing & Document Control System"
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DB_FILE = os.path.join(BASE_DIR, "project_drawings_db.sqlite")

SUPPORTED_EXT = {
    ".pdf", ".dwg", ".dxf", ".xlsx", ".xls", ".xlsm", ".csv", ".docx", ".doc",
}
SKIP_DIR_NAMES = {
    ".git", ".svn", "__pycache__", "$recycle.bin", "system volume information",
    "node_modules", ".backup", "backup", "autosave",
}
MAX_WORKERS = max(2, min(8, os.cpu_count() or 4))
FINGERPRINT_BYTES = 1024 * 1024          # hash the first 1 MB + size
DXF_MAX_BYTES = 40 * 1024 * 1024         # do not parse DXF files bigger than this
NO_DWG = "N/A"
DEFAULT_REV = "00"


def _wide_kwargs(func) -> dict:
    """Full-width kwargs that work on both old and new Streamlit versions."""
    params = inspect.signature(func).parameters
    if "width" in params:
        return {"width": "stretch"}
    if "use_container_width" in params:
        return {"use_container_width": True}
    return {}


WIDE_BTN = _wide_kwargs(st.button)
WIDE_DF = _wide_kwargs(st.dataframe)
WIDE_DL = _wide_kwargs(st.download_button)


# ==========================================
# 1. Database
# ==========================================
SCHEMA_COLUMNS = {
    "file_name": "TEXT",
    "file_path": "TEXT",
    "file_type": "TEXT",
    "dwg_no": "TEXT",
    "title": "TEXT",
    "revision": "TEXT",
    "issue_date": "TEXT",
    "folder_name": "TEXT",
    "root_folder": "TEXT",
    "ref_dwgs": "TEXT",
    "file_size": "INTEGER",
    "last_modified": "REAL",
    "fingerprint": "TEXT",
    "meta_source": "TEXT",
    "indexed_at": "TEXT",
}


def get_db_connection() -> sqlite3.Connection:
    conn = sqlite3.connect(DB_FILE, timeout=30.0)
    conn.execute("PRAGMA journal_mode=WAL;")
    conn.execute("PRAGMA synchronous=NORMAL;")
    return conn


def init_db() -> None:
    with closing(get_db_connection()) as conn:
        conn.execute(
            """
            CREATE TABLE IF NOT EXISTS documents (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                file_name TEXT,
                file_path TEXT UNIQUE,
                file_type TEXT,
                dwg_no TEXT,
                title TEXT,
                revision TEXT,
                issue_date TEXT,
                folder_name TEXT,
                root_folder TEXT,
                ref_dwgs TEXT,
                file_size INTEGER,
                last_modified REAL,
                fingerprint TEXT,
                meta_source TEXT,
                indexed_at TEXT
            )
            """
        )
        # Migrate databases created by older versions of this app.
        existing = {row[1] for row in conn.execute("PRAGMA table_info(documents)")}
        for col, col_type in SCHEMA_COLUMNS.items():
            if col not in existing:
                conn.execute(f"ALTER TABLE documents ADD COLUMN {col} {col_type}")
        conn.execute("CREATE INDEX IF NOT EXISTS idx_dwg_no ON documents(dwg_no);")
        conn.execute("CREATE INDEX IF NOT EXISTS idx_file_type ON documents(file_type);")
        conn.execute("CREATE INDEX IF NOT EXISTS idx_root ON documents(root_folder);")
        conn.execute(
            "CREATE TABLE IF NOT EXISTS settings (key TEXT PRIMARY KEY, value TEXT)"
        )
        conn.commit()


def clear_db() -> None:
    with closing(get_db_connection()) as conn:
        conn.execute("DELETE FROM documents")
        conn.commit()
        conn.execute("VACUUM")


def get_setting(key: str, default: str = "") -> str:
    with closing(get_db_connection()) as conn:
        row = conn.execute("SELECT value FROM settings WHERE key=?", (key,)).fetchone()
    return row[0] if row else default


def set_setting(key: str, value: str) -> None:
    with closing(get_db_connection()) as conn:
        conn.execute(
            "INSERT INTO settings(key, value) VALUES(?, ?) "
            "ON CONFLICT(key) DO UPDATE SET value=excluded.value",
            (key, value),
        )
        conn.commit()


# ==========================================
# 2. Engineering regex & parsers
# ==========================================
MONTHS = {m: i + 1 for i, m in enumerate(
    ["JAN", "FEB", "MAR", "APR", "MAY", "JUN", "JUL", "AUG", "SEP", "OCT", "NOV", "DEC"]
)}

# Drawing number: letters prefix, then 2..7 more segments, at least one digit.
# Segments that are revision markers (REV-03) are not absorbed into the number.
DWG_REGEX = re.compile(
    r"(?<![A-Z0-9])"
    r"(?!(?:REV|REVISION|ISSUE|DATE|SHEET|PAGE|JAN|FEB|MAR|APR|MAY|JUN|JUL|AUG|SEP|OCT|NOV|DEC)[-_])"
    r"(?=[A-Z0-9_-]*\d)"
    r"([A-Z]{1,6}(?:[-_](?!REV(?:ISION)?(?:[-_.]|$))[A-Z0-9]{1,10}){2,7})"
    r"(?![A-Z0-9])"
)

_REV_TOKEN = r"([A-Z]?\d{1,3}[A-Z]?|[A-Z])"

# Inside PDF / DXF text:  "REV: 03", "Revision A", "Issue 2", "REV. P01"
REV_LABEL_REGEX = re.compile(
    r"\b(?:REV(?:ISION)?\.?|ISSUE)\s*(?:NO\.?)?\s*[:.#-]?\s*" + _REV_TOKEN + r"\b",
    re.IGNORECASE,
)

# Inside file names: "_Rev03", "-REV.A", "(Rev 2)", "Revision-P01", "-R03" (only at end)
REV_FILENAME_REGEX = re.compile(
    r"(?:^|[-_\s(\[.])(?:REV(?:ISION)?|ISSUE)[-_.\s:]*" + _REV_TOKEN + r"(?=[-_\s)\].]|$)"
    r"|(?:^|[-_\s(\[.])R[-_.]?(\d{1,2}[A-Z]?)(?=[\s)\].]|$)",
    re.IGNORECASE,
)

DATE_REGEX = re.compile(
    r"(?<!\d)("
    r"\d{4}[-/.]\d{1,2}[-/.]\d{1,2}"                                     # 2024-05-01
    r"|\d{1,2}[-/.]\d{1,2}[-/.]\d{2,4}"                                  # 01/05/2024, 1.5.24
    r"|\d{1,2}[-\s]?(?:JAN|FEB|MAR|APR|MAY|JUN|JUL|AUG|SEP|OCT|NOV|DEC)[A-Z]*[-\s,.]*\d{2,4}"  # 01-MAY-2024
    r"|(?:JAN|FEB|MAR|APR|MAY|JUN|JUL|AUG|SEP|OCT|NOV|DEC)[A-Z]*[-\s]+\d{1,2},?[-\s]+\d{4}"    # May 1, 2024
    r"|20\d{2}(?:0[1-9]|1[0-2])(?:0[1-9]|[12]\d|3[01])"                   # 20240501
    r")(?!\d)",
    re.IGNORECASE,
)

TITLE_LABEL_REGEX = re.compile(
    r"^\s*(?:DRAWING\s+TITLE|DRG\.?\s*TITLE|DWG\.?\s*TITLE|SHEET\s+TITLE|TITLE)\s*[:.\-]?\s*(.*)$",
    re.IGNORECASE,
)
LABEL_LINE_REGEX = re.compile(
    r"^\s*(?:DRAWING|DWG|DRG|SHEET|SCALE|DATE|REV|REVISION|PROJECT|CLIENT|OWNER|CONSULTANT|"
    r"CONTRACTOR|SUB-?CONTRACTOR|CHECKED|DRAWN|APPROVED|DESIGNED|STATUS|SIZE|JOB|FILE|"
    r"DISCIPLINE|STAGE|NOTES?|LEGEND|KEY\s*PLAN|NORTH)\b",
    re.IGNORECASE,
)
DWG_LABEL_REGEX = re.compile(
    r"(?:DRAWING\s*(?:NO|NUMBER|#)\.?|DWG\.?\s*(?:NO|NUMBER|#)\.?|DRG\.?\s*(?:NO|NUMBER|#)\.?|"
    r"SHEET\s*(?:NO|NUMBER|#)\.?)[\s:.-]*([A-Z0-9][A-Z0-9_./-]{3,30})",
    re.IGNORECASE,
)


def normalize_date(raw: str | None) -> str | None:
    """Return an ISO date (YYYY-MM-DD) or None when the text is not a valid date."""
    if not raw:
        return None
    s = raw.strip().upper()
    y = m = d = None

    mm = re.fullmatch(r"(\d{4})[-/.](\d{1,2})[-/.](\d{1,2})", s)
    if mm:
        y, m, d = int(mm.group(1)), int(mm.group(2)), int(mm.group(3))

    if y is None:
        mm = re.fullmatch(r"(\d{1,2})[-/.](\d{1,2})[-/.](\d{2,4})", s)
        if mm:
            a, b, yy = int(mm.group(1)), int(mm.group(2)), int(mm.group(3))
            y = yy + 2000 if yy < 100 else yy
            # Day-first by default (DD/MM/YYYY); switch when only month-first is valid.
            if a > 12 and b <= 12:
                d, m = a, b
            elif a <= 12 and b > 12:
                m, d = a, b
            else:
                d, m = a, b

    if y is None:
        mm = re.fullmatch(r"(\d{1,2})[-\s]?([A-Z]{3})[A-Z]*[-\s,.]*(\d{2,4})", s)
        if mm and mm.group(2) in MONTHS:
            d, m, yy = int(mm.group(1)), MONTHS[mm.group(2)], int(mm.group(3))
            y = yy + 2000 if yy < 100 else yy

    if y is None:
        mm = re.fullmatch(r"([A-Z]{3})[A-Z]*[-\s]+(\d{1,2}),?[-\s]+(\d{4})", s)
        if mm and mm.group(1) in MONTHS:
            m, d, y = MONTHS[mm.group(1)], int(mm.group(2)), int(mm.group(3))

    if y is None:
        mm = re.fullmatch(r"(20\d{2})(\d{2})(\d{2})", s)
        if mm:
            y, m, d = int(mm.group(1)), int(mm.group(2)), int(mm.group(3))

    if y is None:
        return None
    try:
        dt = datetime(y, m, d)
    except (TypeError, ValueError):
        return None
    if not (1990 <= dt.year <= datetime.now().year + 2):
        return None
    return dt.strftime("%Y-%m-%d")


def latest_date(texts: list[str]) -> str | None:
    """Pick the most recent valid date found in the given text snippets."""
    dates = []
    for text in texts:
        for raw in DATE_REGEX.findall(text):
            iso = normalize_date(raw)
            if iso:
                dates.append(iso)
    return max(dates) if dates else None


def revision_sort_rank(rev_val) -> tuple[int, int, str]:
    """
    Rank revisions the way document controllers expect:
        unknown (-) < letters (A, B, C) < P01/P02 (preliminary) < 00, 01, 02 < C01 (contract)
    Returns a tuple usable as a sort key: (tier, number, suffix).
    """
    if rev_val is None:
        return (0, -1, "")
    r = str(rev_val).strip().upper()
    if r in ("", "-", "N/A", "NONE", "NAN"):
        return (0, -1, "")
    m = re.fullmatch(r"([A-Z]?)(\d{1,3})([A-Z]?)", r)
    if m:
        prefix, num, suffix = m.groups()
        if prefix == "":
            tier = 3
        elif prefix == "C":
            tier = 4
        else:
            tier = 2
        return (tier, int(num), suffix)
    if re.fullmatch(r"[A-Z]{1,2}", r):
        return (1, 0, r)
    return (1, 0, r)


def highest_revision(candidates: list[str]) -> str | None:
    cleaned = [c.strip().upper() for c in candidates if c and c.strip()]
    if not cleaned:
        return None
    return max(cleaned, key=revision_sort_rank)


def parse_filename_metadata(filename: str) -> dict:
    """Extract drawing number / revision / date from a file name only."""
    stem = os.path.splitext(filename)[0]
    upper = stem.upper()
    meta = {"dwg_no": NO_DWG, "revision": None, "issue_date": None}

    dwg_match = DWG_REGEX.search(upper)
    dwg_span = dwg_match.span(1) if dwg_match else None
    if dwg_match:
        meta["dwg_no"] = dwg_match.group(1).strip("-_.")

    rev_match = REV_FILENAME_REGEX.search(stem)
    if rev_match:
        meta["revision"] = (rev_match.group(1) or rev_match.group(2)).upper()
        # "GS-E-LTG-02-R03": the rev segment was absorbed into the drawing number -> trim it.
        if dwg_span and dwg_span[0] <= rev_match.start() < dwg_span[1] <= rev_match.end():
            trimmed = upper[dwg_span[0]:rev_match.start()].strip("-_.")
            if DWG_REGEX.fullmatch(trimmed):
                meta["dwg_no"] = trimmed

    meta["issue_date"] = latest_date([stem])
    return meta


def _guess_title(tb_text: str) -> str | None:
    lines = [ln.strip() for ln in tb_text.splitlines()]
    for i, line in enumerate(lines):
        m = TITLE_LABEL_REGEX.match(line)
        if not m:
            continue
        parts: list[str] = []
        rest = m.group(1).strip(" :.-")
        if rest and not LABEL_LINE_REGEX.match(rest):
            parts.append(rest)
        j = i + 1
        while len(parts) < 3 and j < len(lines):
            cand = lines[j]
            j += 1
            if not cand:
                if parts:
                    break
                continue
            if LABEL_LINE_REGEX.match(cand) or DWG_REGEX.fullmatch(cand.upper()):
                break
            parts.append(cand)
        title = re.sub(r"\s+", " ", " ".join(parts)).strip()
        if 3 <= len(title) <= 160:
            return title
    return None


def extract_metadata_from_pdf(pdf_path: str) -> dict:
    """Read the first page's title-block region and return whatever metadata it contains."""
    result = {"dwg_no": None, "title": None, "revision": None, "issue_date": None}
    with fitz.open(pdf_path) as doc:
        if doc.needs_pass or doc.page_count == 0:
            return result

        pdf_title = (doc.metadata or {}).get("title") or ""
        pdf_title = pdf_title.strip()
        if pdf_title and not re.fullmatch(r"(?i)(model|layout\d*|untitled|drawing\d*|.*\.dwg)", pdf_title):
            result["title"] = pdf_title[:160]

        page = doc[0]
        blocks = page.get_text("blocks") or []
        w, h = page.rect.width, page.rect.height

        tb_blocks, all_blocks = [], []
        for b in blocks:
            if len(b) >= 5 and isinstance(b[4], str):
                text = b[4].strip()
                if not text:
                    continue
                all_blocks.append(text)
                x0, y0 = b[0], b[1]
                # Title blocks sit in the bottom strip or the right-hand strip of the sheet.
                if y0 >= h * 0.55 or x0 >= w * 0.60:
                    tb_blocks.append(text)

        if not all_blocks:          # scanned / raster PDF - nothing to read
            return result

        tb_text = "\n".join(tb_blocks) if tb_blocks else "\n".join(all_blocks)
        full_text = "\n".join(all_blocks)

        label = DWG_LABEL_REGEX.search(tb_text)
        if label and any(ch.isdigit() for ch in label.group(1)):
            result["dwg_no"] = label.group(1).upper().strip("-_./")
        else:
            matches = DWG_REGEX.findall(tb_text.upper())
            if matches:
                # The drawing number is usually repeated / last in the block; prefer the most common.
                result["dwg_no"] = max(reversed(matches), key=matches.count)

        guessed = _guess_title(tb_text)
        if guessed:
            result["title"] = guessed

        result["revision"] = highest_revision(REV_LABEL_REGEX.findall(tb_text))
        result["issue_date"] = latest_date([tb_text]) or latest_date([full_text])
    return result


def extract_dwgs_from_docx(docx_path: str) -> set[str]:
    found: set[str] = set()
    if docx is None:
        return found
    document = docx.Document(docx_path)
    for p in document.paragraphs:
        found.update(DWG_REGEX.findall(p.text.upper()))
    for table in document.tables:
        for row in table.rows:
            for cell in row.cells:
                found.update(DWG_REGEX.findall(cell.text.upper()))
    return found


def extract_dwgs_from_excel(excel_path: str) -> set[str]:
    found: set[str] = set()
    if openpyxl is None:
        return found
    wb = openpyxl.load_workbook(excel_path, read_only=True, data_only=True)
    try:
        for sheet in wb.sheetnames[:5]:
            ws = wb[sheet]
            for row in ws.iter_rows(max_row=500, max_col=30, values_only=True):
                for val in row:
                    if val is not None:
                        found.update(DWG_REGEX.findall(str(val).upper()))
    finally:
        wb.close()
    return found


def extract_text_from_dxf(dxf_path: str) -> str:
    """Collect TEXT / MTEXT string values (group code 1) from an ASCII DXF file."""
    if os.path.getsize(dxf_path) > DXF_MAX_BYTES:
        return ""
    strings: list[str] = []
    with open(dxf_path, "r", encoding="utf-8", errors="ignore") as fh:
        prev = None
        for line in fh:
            if prev is not None and prev.strip() == "1":
                val = line.strip()
                if val and not val.startswith("AcDb"):
                    strings.append(val)
                    if len(strings) >= 5000:
                        break
            prev = line
    return "\n".join(strings)


def quick_fingerprint(path: str, size: int) -> str:
    h = hashlib.md5()
    h.update(str(size).encode())
    with open(path, "rb") as fh:
        h.update(fh.read(FINGERPRINT_BYTES))
    return h.hexdigest()


# ==========================================
# 3. Project folder scanner
# ==========================================
def iter_project_files(root_folder: str):
    """Yield (path, name, ext, mtime, size) for every supported file under root_folder."""
    stack = [root_folder]
    while stack:
        current = stack.pop()
        try:
            with os.scandir(current) as it:
                for entry in it:
                    try:
                        if entry.is_dir(follow_symlinks=False):
                            name = entry.name.lower()
                            if name not in SKIP_DIR_NAMES and not name.startswith("."):
                                stack.append(entry.path)
                        elif entry.is_file(follow_symlinks=False):
                            name = entry.name
                            if name.startswith(("~$", ".")):
                                continue
                            ext = os.path.splitext(name)[1].lower()
                            if ext in SUPPORTED_EXT:
                                stt = entry.stat()
                                yield entry.path, name, ext, stt.st_mtime, stt.st_size
                    except OSError:
                        continue
        except OSError:
            continue


def _extract_file(job: tuple) -> tuple:
    """Worker: build one database row for a file. Runs in a thread pool."""
    path, name, ext, mtime, size, root = job
    ftype = ext.lstrip(".").upper()
    stem = os.path.splitext(name)[0]

    meta = parse_filename_metadata(name)
    dwg_no = meta["dwg_no"]
    revision = meta["revision"]
    issue_date = meta["issue_date"]
    title = stem
    ref_dwgs = None
    source = "filename"

    try:
        if ftype == "PDF":
            pdf = extract_metadata_from_pdf(path)
            if dwg_no == NO_DWG and pdf["dwg_no"]:
                dwg_no, source = pdf["dwg_no"], "pdf-text"
            if pdf["title"]:
                title = pdf["title"]
            if revision is None and pdf["revision"]:
                revision = pdf["revision"]
            if issue_date is None and pdf["issue_date"]:
                issue_date = pdf["issue_date"]
        elif ftype == "DXF":
            text = extract_text_from_dxf(path).upper()
            if text:
                if dwg_no == NO_DWG:
                    label = DWG_LABEL_REGEX.search(text)
                    if label and any(ch.isdigit() for ch in label.group(1)):
                        dwg_no, source = label.group(1).strip("-_./"), "dxf-text"
                    else:
                        found = DWG_REGEX.findall(text)
                        if found:
                            dwg_no, source = max(reversed(found), key=found.count), "dxf-text"
                if revision is None:
                    revision = highest_revision(REV_LABEL_REGEX.findall(text))
                if issue_date is None:
                    issue_date = latest_date([text])
        elif ftype == "DOCX":
            refs = extract_dwgs_from_docx(path)
            if refs:
                ref_dwgs = ", ".join(sorted(refs)[:12])
        elif ftype in ("XLSX", "XLSM"):
            refs = extract_dwgs_from_excel(path)
            if refs:
                ref_dwgs = ", ".join(sorted(refs)[:12])
    except Exception:
        # Corrupt / locked / password-protected file: keep the filename-based metadata.
        source = source + " (content unreadable)"

    try:
        fingerprint = quick_fingerprint(path, size)
    except OSError:
        fingerprint = None

    rel_folder = os.path.relpath(os.path.dirname(path), root)
    folder_name = os.path.basename(root.rstrip("\\/")) if rel_folder == "." else rel_folder.replace("\\", "/")

    return (
        name, path, ftype, dwg_no, title, revision or DEFAULT_REV, issue_date or NO_DWG,
        folder_name, root, ref_dwgs, size, mtime, fingerprint, source,
        datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
    )


_UPSERT_SQL = """
    INSERT INTO documents
        (file_name, file_path, file_type, dwg_no, title, revision, issue_date, folder_name,
         root_folder, ref_dwgs, file_size, last_modified, fingerprint, meta_source, indexed_at)
    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
    ON CONFLICT(file_path) DO UPDATE SET
        file_name=excluded.file_name,
        file_type=excluded.file_type,
        dwg_no=excluded.dwg_no,
        title=excluded.title,
        revision=excluded.revision,
        issue_date=excluded.issue_date,
        folder_name=excluded.folder_name,
        root_folder=excluded.root_folder,
        ref_dwgs=excluded.ref_dwgs,
        file_size=excluded.file_size,
        last_modified=excluded.last_modified,
        fingerprint=excluded.fingerprint,
        meta_source=excluded.meta_source,
        indexed_at=excluded.indexed_at
"""


def scan_project_folder(root_folder: str, deep: bool = False, purge_missing: bool = True,
                        progress_bar=None, status_text=None) -> dict:
    """
    Index every supported file under root_folder.
    Incremental: unchanged files (same mtime + size) are skipped unless deep=True.
    """
    root = os.path.normpath(root_folder)
    t0 = time.time()
    stats = {"total": 0, "updated": 0, "skipped": 0, "removed": 0, "errors": 0, "seconds": 0.0}

    if status_text is not None:
        status_text.caption("Discovering files ...")
    files = list(iter_project_files(root))
    stats["total"] = len(files)

    with closing(get_db_connection()) as conn:
        existing = {
            p: (m, s) for p, m, s in conn.execute(
                "SELECT file_path, last_modified, file_size FROM documents"
            )
        }
        jobs, seen = [], set()
        for path, name, ext, mtime, size in files:
            norm = os.path.normpath(path)
            seen.add(norm)
            prev = existing.get(norm)
            if not deep and prev and prev[0] == mtime and prev[1] == size:
                stats["skipped"] += 1
                continue
            jobs.append((norm, name, ext, mtime, size, root))

        done, batch = 0, []
        if jobs:
            with ThreadPoolExecutor(max_workers=MAX_WORKERS) as pool:
                future_map = {pool.submit(_extract_file, job): job for job in jobs}
                for fut in as_completed(future_map):
                    done += 1
                    try:
                        batch.append(fut.result())
                    except Exception:
                        stats["errors"] += 1
                    if progress_bar is not None:
                        progress_bar.progress(done / len(jobs))
                    if status_text is not None and (done % 10 == 0 or done == len(jobs)):
                        status_text.caption(f"Indexing {done:,}/{len(jobs):,}: `{future_map[fut][1]}`")
                    if len(batch) >= 200:
                        conn.executemany(_UPSERT_SQL, batch)
                        conn.commit()
                        batch = []
            if batch:
                conn.executemany(_UPSERT_SQL, batch)
        stats["updated"] = len(jobs) - stats["errors"]

        if purge_missing:
            stale = []
            rows = conn.execute(
                "SELECT file_path FROM documents WHERE root_folder=? OR root_folder IS NULL", (root,)
            ).fetchall()
            for (p,) in rows:
                if p in seen:
                    continue
                inside_root = os.path.normcase(p).startswith(os.path.normcase(root + os.sep))
                if inside_root and not os.path.exists(p):
                    stale.append((p,))
            if stale:
                conn.executemany("DELETE FROM documents WHERE file_path=?", stale)
            stats["removed"] = len(stale)
        conn.commit()

    stats["seconds"] = round(time.time() - t0, 1)
    return stats


# ==========================================
# 4. Data loading & classification
# ==========================================
def _norm_search(text) -> str:
    return re.sub(r"[\s\-_./\\,()\[\]]+", " ", str(text).upper()).strip()


@st.cache_data(show_spinner=False)
def load_documents(version: int) -> pd.DataFrame:
    """Load the register and compute status / duplicate flags. Cached per data version."""
    with closing(get_db_connection()) as conn:
        df = pd.read_sql_query("SELECT * FROM documents", conn)
    if df.empty:
        return df

    for col in ("dwg_no", "title", "revision", "issue_date", "folder_name", "ref_dwgs",
                "fingerprint", "root_folder", "meta_source"):
        df[col] = df[col].fillna("").astype(str)
    df["file_size"] = pd.to_numeric(df["file_size"], errors="coerce").fillna(0)

    ranks = df["revision"].map(revision_sort_rank)
    df["rev_tier"] = [r[0] for r in ranks]
    df["rev_num"] = [r[1] for r in ranks]
    df["rev_suffix"] = [r[2] for r in ranks]
    df["rev_key"] = [f"{t}{n + 1:05d}{sfx}" for t, n, sfx in ranks]   # sortable text key
    df["has_dwg_no"] = ~df["dwg_no"].isin([NO_DWG, "-", ""])
    df["modified"] = pd.to_datetime(df["last_modified"], unit="s", errors="coerce").dt.floor("min")
    df["size_mb"] = (df["file_size"] / 1_048_576).round(2)
    df["project"] = df["root_folder"].map(lambda p: os.path.basename(p.rstrip("\\/")) or p)

    df = df.sort_values(
        by=["has_dwg_no", "dwg_no", "rev_tier", "rev_num", "rev_suffix", "last_modified"],
        ascending=[False, True, False, False, False, False],
    ).reset_index(drop=True)

    # Latest revision per drawing number + file type (across all folders). Every file that carries
    # the highest revision of its group is "Current", so identical copies are never marked superseded.
    mask = df["has_dwg_no"]
    group_max = df.loc[mask].groupby(["dwg_no", "file_type"])["rev_key"].transform("max")
    df["is_latest"] = True
    df.loc[mask, "is_latest"] = df.loc[mask, "rev_key"] == group_max
    df["status"] = np.where(~mask, "Unclassified", np.where(df["is_latest"], "Current", "Superseded"))

    # Content duplicates: identical fingerprint stored under several paths.
    fp = df["fingerprint"]
    df["is_duplicate"] = fp.ne("") & fp.duplicated(keep=False)

    # Revision conflicts: same drawing no. + revision + type, but different content.
    key_cols = ["dwg_no", "revision", "file_type"]
    conflict = (
        df[mask].groupby(key_cols)["fingerprint"].transform("nunique") > 1
    )
    df["is_conflict"] = False
    df.loc[conflict[conflict].index, "is_conflict"] = True

    df["_search"] = (
        df["file_name"] + " " + df["dwg_no"] + " " + df["title"] + " " + df["folder_name"]
        + " " + df["ref_dwgs"] + " REV " + df["revision"]
    ).map(_norm_search)
    return df


def apply_search(df: pd.DataFrame, query: str) -> pd.DataFrame:
    tokens = _norm_search(query).split()
    if not tokens:
        return df
    mask = pd.Series(True, index=df.index)
    for tok in tokens:
        mask &= df["_search"].str.contains(re.escape(tok), regex=True, na=False)
    return df[mask]


def build_excel_register(df: pd.DataFrame) -> bytes:
    cols = ["status", "dwg_no", "revision", "title", "file_type", "issue_date", "project",
            "folder_name", "file_name", "modified", "size_mb", "ref_dwgs", "file_path"]
    out = df[cols].copy()
    out["modified"] = out["modified"].dt.strftime("%Y-%m-%d %H:%M")
    out.columns = ["Status", "Drawing No", "Rev", "Title", "Type", "Issue Date", "Project",
                   "Folder", "File Name", "Modified", "Size (MB)", "Referenced Drawings", "Path"]
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        out.to_excel(writer, index=False, sheet_name="Drawing Register")
        ws = writer.sheets["Drawing Register"]
        ws.freeze_panes = "A2"
        ws.auto_filter.ref = ws.dimensions
        widths = [12, 24, 7, 45, 7, 12, 18, 30, 40, 17, 10, 30, 70]
        for i, wdt in enumerate(widths, start=1):
            ws.column_dimensions[openpyxl.utils.get_column_letter(i)].width = wdt
    return buf.getvalue()


# ==========================================
# 5. Cross-platform file & folder launchers
# ==========================================
def open_file(path: str) -> None:
    norm_path = os.path.normpath(path)
    if not os.path.exists(norm_path):
        st.toast(f"File not found: {norm_path}", icon="❌")
        return
    try:
        system = platform.system()
        if system == "Windows":
            os.startfile(norm_path)  # type: ignore[attr-defined]
        elif system == "Darwin":
            subprocess.Popen(["open", norm_path])
        else:
            subprocess.Popen(["xdg-open", norm_path])
        st.toast(f"Opening {os.path.basename(norm_path)}", icon="📂")
    except Exception as exc:  # pragma: no cover
        st.toast(f"Could not open file: {exc}", icon="❌")


def open_folder(path: str) -> None:
    norm_path = os.path.normpath(path)
    folder = os.path.dirname(norm_path)
    if not os.path.isdir(folder):
        st.toast(f"Folder not found: {folder}", icon="❌")
        return
    try:
        system = platform.system()
        if system == "Windows":
            subprocess.Popen(["explorer", "/select,", norm_path])
        elif system == "Darwin":
            subprocess.Popen(["open", "-R", norm_path])
        else:
            subprocess.Popen(["xdg-open", folder])
    except Exception as exc:  # pragma: no cover
        st.toast(f"Could not open folder: {exc}", icon="❌")


# ==========================================
# 6. Streamlit UI
# ==========================================
TYPE_ICON = {
    "PDF": "📕", "DWG": "📐", "DXF": "📐", "XLSX": "📊", "XLS": "📊", "XLSM": "📊",
    "CSV": "📊", "DOCX": "📝", "DOC": "📝",
}
STATUS_BADGE = {"Current": "🟢 Current", "Superseded": "🟠 Superseded", "Unclassified": "⚪ Unclassified"}


def bump_data_version() -> None:
    st.session_state["data_version"] = st.session_state.get("data_version", 0) + 1
    load_documents.clear()


def main() -> None:
    st.set_page_config(page_title="Drawing Control System", page_icon="🏗️", layout="wide")
    init_db()
    st.session_state.setdefault("data_version", 0)
    st.session_state.setdefault("page", 1)
    st.session_state.setdefault("filter_sig", None)

    st.title("🏗️ " + APP_TITLE)

    # ---------- Sidebar ----------
    with st.sidebar:
        st.header("⚙️ Scanner & Indexer")
        raw_folder_path = st.text_input(
            "Project root folder:",
            value=st.session_state.get("root_input", get_setting("last_root", "")),
            placeholder=r"C:\Users\<you>\Desktop\GS project",
            help="Paste the folder path from Windows Explorer (quotes are removed automatically).",
            key="root_input",
        )
        folder_path = os.path.expandvars(os.path.expanduser(raw_folder_path.strip().strip('"').strip("'")))

        deep_scan = st.toggle("Deep scan (re-read unchanged files)", value=False,
                              help="Use after upgrading the app or when title-block data looks wrong.")
        purge_missing = st.toggle("Remove entries whose files were deleted", value=True)

        if st.button("🔍 Scan & Index", type="primary", **WIDE_BTN):
            if folder_path and os.path.isdir(folder_path):
                set_setting("last_root", folder_path)
                prog_bar = st.progress(0.0)
                stat_txt = st.empty()
                with st.spinner("Scanning project folder ..."):
                    stats = scan_project_folder(folder_path, deep=deep_scan, purge_missing=purge_missing,
                                                progress_bar=prog_bar, status_text=stat_txt)
                stat_txt.empty()
                prog_bar.empty()
                st.session_state["last_scan"] = stats
                st.session_state["page"] = 1
                bump_data_version()
                st.rerun()
            else:
                st.error("Invalid folder path. Please verify the directory exists and is accessible.")

        if "last_scan" in st.session_state:
            s = st.session_state["last_scan"]
            st.success(
                f"Scan finished in {s['seconds']}s\n\n"
                f"- Found: **{s['total']:,}** files\n"
                f"- Added / updated: **{s['updated']:,}**\n"
                f"- Unchanged: **{s['skipped']:,}**\n"
                f"- Removed (deleted on disk): **{s['removed']:,}**\n"
                + (f"- Errors: **{s['errors']:,}**" if s["errors"] else "")
            )

        st.divider()
        st.subheader("🛠️ Display")
        items_per_page = st.selectbox("Rows per page:", options=[25, 50, 100, 250, 500], index=1)
        view_mode = st.radio("View mode:", options=["Cards (with Open buttons)", "Table (fast scrolling)"])
        st.caption("Open buttons work when this app runs on the same PC as the browser.")

        st.divider()
        st.subheader("🗄️ Database")
        st.caption(f"`{DB_FILE}`")
        confirm_reset = st.checkbox("I understand this deletes the whole index")
        if st.button("🗑️ Reset database", disabled=not confirm_reset, **WIDE_BTN):
            clear_db()
            st.session_state.pop("last_scan", None)
            st.session_state["page"] = 1
            bump_data_version()
            st.warning("Database cleared.")
            st.rerun()

    # ---------- Data ----------
    df = load_documents(st.session_state["data_version"])
    if df.empty:
        st.info("No files indexed yet. Enter the project root folder in the sidebar and click **Scan & Index**.")
        return

    # ---------- Metrics ----------
    m1, m2, m3, m4, m5, m6 = st.columns(6)
    m1.metric("Indexed files", f"{len(df):,}")
    m2.metric("Drawing numbers", f"{df.loc[df['has_dwg_no'], 'dwg_no'].nunique():,}")
    m3.metric("PDF", f"{(df['file_type'] == 'PDF').sum():,}")
    m4.metric("DWG / DXF", f"{df['file_type'].isin(['DWG', 'DXF']).sum():,}")
    m5.metric("Excel / Word", f"{df['file_type'].isin(['XLSX', 'XLS', 'XLSM', 'CSV', 'DOCX', 'DOC']).sum():,}")
    m6.metric("Superseded", f"{(df['status'] == 'Superseded').sum():,}",
              help="Older revisions of a drawing number that has a newer revision indexed.")

    st.divider()

    # ---------- Filters ----------
    c1, c2, c3, c4, c5, c6 = st.columns([2.6, 1.4, 1.7, 1.1, 1.1, 1.3])
    with c1:
        search_query = st.text_input("🔎 Search (e.g. 'GF Lighting', 'E LTG 02', 'Rev 03'):", key="search")
    with c2:
        projects = sorted(df["project"].unique().tolist())
        selected_projects = st.multiselect("🏢 Project:", options=projects, placeholder="All projects")
    with c3:
        folders = sorted(df["folder_name"].unique().tolist())
        selected_folders = st.multiselect("📍 Location / folder:", options=folders, placeholder="All folders")
    with c4:
        types = sorted(df["file_type"].unique().tolist())
        selected_types = st.multiselect("📄 Type:", options=types, placeholder="All types")
    with c5:
        revs = sorted(df["revision"].unique().tolist(), key=revision_sort_rank, reverse=True)
        selected_revs = st.multiselect("🔢 Revision:", options=revs, placeholder="All revs")
    with c6:
        status_choice = st.selectbox("✅ Status:", options=["All", "Current (latest rev)", "Superseded", "Unclassified"],
                                     key="status")

    s1, s2, s3 = st.columns([1.6, 1.2, 3])
    with s1:
        sort_by = st.selectbox("Sort by:", options=["Drawing No", "Revision (high → low)", "Last modified (newest)",
                                                    "Issue date (newest)", "File name", "Folder"])
    with s2:
        extra = st.multiselect("Flags:", options=["Duplicates only", "Revision conflicts only", "Has issue date"],
                               placeholder="No extra filter")

    filtered = df
    if selected_projects:
        filtered = filtered[filtered["project"].isin(selected_projects)]
    if selected_folders:
        filtered = filtered[filtered["folder_name"].isin(selected_folders)]
    if selected_types:
        filtered = filtered[filtered["file_type"].isin(selected_types)]
    if selected_revs:
        filtered = filtered[filtered["revision"].isin(selected_revs)]
    if status_choice == "Current (latest rev)":
        filtered = filtered[filtered["status"] == "Current"]
    elif status_choice == "Superseded":
        filtered = filtered[filtered["status"] == "Superseded"]
    elif status_choice == "Unclassified":
        filtered = filtered[filtered["status"] == "Unclassified"]
    if "Duplicates only" in extra:
        filtered = filtered[filtered["is_duplicate"]]
    if "Revision conflicts only" in extra:
        filtered = filtered[filtered["is_conflict"]]
    if "Has issue date" in extra:
        filtered = filtered[filtered["issue_date"] != NO_DWG]
    filtered = apply_search(filtered, search_query)

    sort_map = {
        "Drawing No": (["has_dwg_no", "dwg_no", "rev_tier", "rev_num", "rev_suffix"], [False, True, False, False, False]),
        "Revision (high → low)": (["rev_tier", "rev_num", "rev_suffix", "dwg_no"], [False, False, False, True]),
        "Last modified (newest)": (["last_modified"], [False]),
        "Issue date (newest)": (["issue_date"], [False]),
        "File name": (["file_name"], [True]),
        "Folder": (["folder_name", "file_name"], [True, True]),
    }
    by, asc = sort_map[sort_by]
    filtered = filtered.sort_values(by=by, ascending=asc, kind="stable")

    # Reset pagination whenever the filter set changes.
    sig = (search_query, tuple(selected_projects), tuple(selected_folders), tuple(selected_types),
           tuple(selected_revs), status_choice, tuple(extra), sort_by, items_per_page)
    if st.session_state["filter_sig"] != sig:
        st.session_state["filter_sig"] = sig
        st.session_state["page"] = 1

    tab_register, tab_issues, tab_stats = st.tabs(["📋 Register", "⚠️ Issues", "📊 Statistics"])

    # ---------- Register tab ----------
    with tab_register:
        total = len(filtered)
        total_pages = max(1, -(-total // items_per_page))
        st.session_state["page"] = min(max(1, st.session_state["page"]), total_pages)

        h1, h2, h3, h4 = st.columns([3, 1.4, 1.4, 1.2])
        h1.markdown(f"**{total:,}** of {len(df):,} files match &nbsp;·&nbsp; page "
                    f"{st.session_state['page']} / {total_pages}")
        if total:
            h2.download_button("⬇️ Excel register", data=build_excel_register(filtered),
                               file_name=f"drawing_register_{datetime.now():%Y%m%d_%H%M}.xlsx",
                               mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                               **WIDE_DL)
            csv_cols = ["status", "dwg_no", "revision", "title", "file_type", "issue_date", "folder_name",
                        "file_name", "file_path"]
            h3.download_button("⬇️ CSV", data=filtered[csv_cols].to_csv(index=False).encode("utf-8-sig"),
                               file_name=f"drawing_register_{datetime.now():%Y%m%d_%H%M}.csv", mime="text/csv",
                               **WIDE_DL)
        with h4:
            new_page = st.number_input("Page", min_value=1, max_value=total_pages,
                                       value=st.session_state["page"], step=1, label_visibility="collapsed")
            st.session_state["page"] = int(new_page)

        start = (st.session_state["page"] - 1) * items_per_page
        page_df = filtered.iloc[start:start + items_per_page]

        if total == 0:
            st.warning("No files match the current filters.")
        elif view_mode.startswith("Table"):
            show_cols = ["status", "dwg_no", "revision", "title", "file_type", "issue_date", "project",
                         "folder_name", "file_name", "modified", "size_mb", "is_duplicate", "is_conflict",
                         "file_path"]
            st.dataframe(
                page_df[show_cols],
                hide_index=True,
                column_config={
                    "status": st.column_config.TextColumn("Status", width="small"),
                    "dwg_no": st.column_config.TextColumn("Drawing No", width="medium"),
                    "revision": st.column_config.TextColumn("Rev", width="small"),
                    "title": st.column_config.TextColumn("Title", width="large"),
                    "file_type": st.column_config.TextColumn("Type", width="small"),
                    "issue_date": st.column_config.TextColumn("Issue date", width="small"),
                    "project": st.column_config.TextColumn("Project", width="small"),
                    "folder_name": st.column_config.TextColumn("Folder", width="medium"),
                    "file_name": st.column_config.TextColumn("File", width="medium"),
                    "modified": st.column_config.DatetimeColumn("Modified", format="YYYY-MM-DD HH:mm"),
                    "size_mb": st.column_config.NumberColumn("MB", format="%.2f"),
                    "is_duplicate": st.column_config.CheckboxColumn("Dup"),
                    "is_conflict": st.column_config.CheckboxColumn("Conflict"),
                    "file_path": st.column_config.TextColumn("Path", width="large"),
                },
                height=min(900, 42 + 35 * len(page_df)),
                **WIDE_DF,
            )
        else:
            for row in page_df.itertuples(index=False):
                with st.container(border=True):
                    a, b, c, d, e, f, g = st.columns([3.2, 1.6, 0.7, 1.1, 1.1, 2.2, 1.6])
                    icon = TYPE_ICON.get(row.file_type, "📄")
                    a.markdown(f"{icon} **{row.title}**")
                    a.caption(row.file_name)
                    b.markdown(f"**{row.dwg_no}**")
                    if row.ref_dwgs:
                        b.caption(f"refs: {row.ref_dwgs}")
                    c.markdown(f"Rev **{row.revision}**")
                    d.markdown(STATUS_BADGE.get(row.status, row.status))
                    flags = []
                    if row.is_duplicate:
                        flags.append("🔁 duplicate")
                    if row.is_conflict:
                        flags.append("⚠️ conflict")
                    if flags:
                        d.caption(" · ".join(flags))
                    e.markdown(f"📅 {row.issue_date}")
                    e.caption(f"{row.size_mb} MB")
                    f.markdown(f"📍 {row.folder_name}")
                    f.caption(f"modified {row.modified:%Y-%m-%d %H:%M}" if pd.notna(row.modified) else "")
                    with g:
                        b1, b2 = st.columns(2)
                        if b1.button("📂 Open", key=f"open_{row.id}", help=row.file_path, **WIDE_BTN):
                            open_file(row.file_path)
                        if b2.button("🗂️ Folder", key=f"folder_{row.id}", help=os.path.dirname(row.file_path),
                                     **WIDE_BTN):
                            open_folder(row.file_path)
                        with st.popover("Path", **WIDE_BTN):
                            st.code(row.file_path, language=None)

        if total_pages > 1:
            p1, p2, p3 = st.columns([1, 6, 1])
            if p1.button("◀ Previous", disabled=st.session_state["page"] <= 1, **WIDE_BTN):
                st.session_state["page"] -= 1
                st.rerun()
            if p3.button("Next ▶", disabled=st.session_state["page"] >= total_pages, **WIDE_BTN):
                st.session_state["page"] += 1
                st.rerun()

    # ---------- Issues tab ----------
    with tab_issues:
        no_dwg = df[~df["has_dwg_no"]]
        dups = df[df["is_duplicate"]].sort_values(["fingerprint", "file_path"])
        conflicts = df[df["is_conflict"]].sort_values(["dwg_no", "revision", "file_path"])
        i1, i2, i3 = st.columns(3)
        i1.metric("Files without drawing number", f"{len(no_dwg):,}")
        i2.metric("Duplicate copies (same content)", f"{len(dups):,}")
        i3.metric("Revision conflicts", f"{len(conflicts):,}",
                  help="Same drawing number and revision, but the file contents differ.")

        issue_cols = ["dwg_no", "revision", "file_type", "file_name", "folder_name", "modified", "size_mb", "file_path"]
        with st.expander(f"🔁 Duplicate copies ({len(dups):,})", expanded=bool(len(dups))):
            if len(dups):
                st.dataframe(dups[["fingerprint"] + issue_cols], hide_index=True, **WIDE_DF)
            else:
                st.caption("No duplicate files found.")
        with st.expander(f"⚠️ Revision conflicts ({len(conflicts):,})", expanded=bool(len(conflicts))):
            if len(conflicts):
                st.dataframe(conflicts[issue_cols], hide_index=True, **WIDE_DF)
            else:
                st.caption("No conflicts found.")
        with st.expander(f"⚪ Files without a drawing number ({len(no_dwg):,})"):
            if len(no_dwg):
                st.dataframe(no_dwg[["file_type", "file_name", "folder_name", "meta_source", "file_path"]],
                             hide_index=True, **WIDE_DF)
            else:
                st.caption("Every indexed file has a drawing number.")

    # ---------- Statistics tab ----------
    with tab_stats:
        g1, g2 = st.columns(2)
        with g1:
            st.markdown("**Files per folder (top 25)**")
            st.bar_chart(df["folder_name"].value_counts().head(25), horizontal=True)
        with g2:
            st.markdown("**Files per type**")
            st.bar_chart(df["file_type"].value_counts())
            st.markdown("**Status**")
            st.bar_chart(df["status"].value_counts())
        st.markdown("**Revision distribution (drawings with a number)**")
        rev_counts = df.loc[df["has_dwg_no"], "revision"].value_counts()
        rev_counts = rev_counts.reindex(sorted(rev_counts.index, key=revision_sort_rank))
        st.bar_chart(rev_counts)
        st.markdown("**Latest indexing activity**")
        st.dataframe(
            df.sort_values("indexed_at", ascending=False).head(15)[
                ["indexed_at", "dwg_no", "revision", "file_name", "folder_name", "meta_source"]
            ],
            hide_index=True, **WIDE_DF,
        )


if __name__ == "__main__":
    main()
