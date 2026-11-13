# -*- coding: utf-8 -*-
"""
Drawing & Document Control  (local EDMS for MEP / construction projects)
=======================================================================
Single-file Streamlit application.

    streamlit run app.py

What it does
------------
* Indexes a project folder (PDF, DWG, DXF, Excel, Word, CSV): incremental, multi-threaded.
* Reads PDF title blocks: drawing number, title, revision, date.
* Folder tree navigation, automatic zone / discipline classification.
* Current / Superseded status, version history per drawing, visual revision compare.
* Built-in drawing viewer and previews (no need to leave the app).
* Workflow: CDE state (WIP / Shared / Published / Archived), approval code, notes, favourites.
* Transmittals with numbering, Excel transmittal sheet and file export.
* Quality checks: duplicates, revision conflicts, naming standard, missing CAD / PDF pairs.
* Activity log (audit trail), Excel / CSV drawing register.
"""
from __future__ import annotations

import base64
import getpass
import hashlib
import html
import inspect
import io
import os
import platform
import re
import shutil
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
APP_TITLE = "Drawing & Document Control"
APP_VERSION = "3.0"
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DB_FILE = os.path.join(BASE_DIR, "project_drawings_db.sqlite")
CACHE_DIR = os.path.join(BASE_DIR, ".dc_cache")
PRIMARY_COLOR = "#1565D8"

UNSET = "—"
CDE_STATES = [UNSET, "WIP", "Shared", "Published", "Archived"]
APPROVAL_CODES = [UNSET, "A · Approved", "B · Approved as noted", "C · Revise & resubmit",
                  "D · Rejected", "For information"]
TRANSMITTAL_PURPOSES = ["For Approval", "For Review & Comments", "For Construction",
                        "For Information", "As-Built", "For Tender"]
DRAWING_TYPES = ("PDF", "DWG", "DXF")

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
WIDE_IMG = _wide_kwargs(st.image)


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
        # User-entered workflow data. Keyed by drawing number + revision (shared by the PDF and
        # the DWG of the same revision) or by file path for files without a drawing number.
        conn.execute(
            """CREATE TABLE IF NOT EXISTS doc_meta (
                   meta_key TEXT PRIMARY KEY, cde_state TEXT, approval_code TEXT, notes TEXT,
                   favorite INTEGER DEFAULT 0, updated_at TEXT)"""
        )
        conn.execute(
            """CREATE TABLE IF NOT EXISTS activity_log (
                   id INTEGER PRIMARY KEY AUTOINCREMENT, ts TEXT, user TEXT, action TEXT,
                   target TEXT, details TEXT)"""
        )
        conn.execute(
            """CREATE TABLE IF NOT EXISTS transmittals (
                   id INTEGER PRIMARY KEY AUTOINCREMENT, number TEXT UNIQUE, tdate TEXT, to_party TEXT,
                   from_party TEXT, subject TEXT, purpose TEXT, remarks TEXT, created_at TEXT)"""
        )
        conn.execute(
            """CREATE TABLE IF NOT EXISTS transmittal_items (
                   id INTEGER PRIMARY KEY AUTOINCREMENT, transmittal_id INTEGER, file_path TEXT,
                   file_name TEXT, dwg_no TEXT, revision TEXT, title TEXT, file_type TEXT)"""
        )
        conn.commit()


def clear_db(everything: bool = False) -> None:
    """Delete the file index. With everything=True also workflow data, transmittals and the log."""
    with closing(get_db_connection()) as conn:
        conn.execute("DELETE FROM documents")
        if everything:
            for table in ("doc_meta", "activity_log", "transmittals", "transmittal_items"):
                conn.execute(f"DELETE FROM {table}")
        conn.commit()
        conn.execute("VACUUM")


def current_user() -> str:
    try:
        return getpass.getuser()
    except Exception:  # pragma: no cover
        return "user"


def log_event(action: str, target: str = "", details: str = "") -> None:
    """Append one line to the audit trail. Never raises."""
    try:
        with closing(get_db_connection()) as conn:
            conn.execute(
                "INSERT INTO activity_log(ts, user, action, target, details) VALUES(?, ?, ?, ?, ?)",
                (datetime.now().strftime("%Y-%m-%d %H:%M:%S"), current_user(), action, target, details),
            )
            conn.commit()
    except sqlite3.Error:  # pragma: no cover
        pass


def save_meta(meta_key: str, field: str, value) -> None:
    assert field in ("cde_state", "approval_code", "notes", "favorite")
    with closing(get_db_connection()) as conn:
        conn.execute(
            f"INSERT INTO doc_meta(meta_key, {field}, updated_at) VALUES(?, ?, ?) "
            f"ON CONFLICT(meta_key) DO UPDATE SET {field}=excluded.{field}, updated_at=excluded.updated_at",
            (meta_key, value, datetime.now().strftime("%Y-%m-%d %H:%M:%S")),
        )
        conn.commit()


def load_activity(limit: int = 500) -> pd.DataFrame:
    with closing(get_db_connection()) as conn:
        return pd.read_sql_query(
            "SELECT ts, user, action, target, details FROM activity_log ORDER BY id DESC LIMIT ?",
            conn, params=(limit,),
        )


def next_transmittal_number() -> str:
    with closing(get_db_connection()) as conn:
        rows = conn.execute("SELECT number FROM transmittals").fetchall()
    highest = 0
    for (num,) in rows:
        m = re.search(r"(\d+)\s*$", num or "")
        if m:
            highest = max(highest, int(m.group(1)))
    return f"TR-{highest + 1:04d}"


def create_transmittal(header: dict, items: pd.DataFrame) -> int:
    with closing(get_db_connection()) as conn:
        cur = conn.execute(
            "INSERT INTO transmittals(number, tdate, to_party, from_party, subject, purpose, remarks, created_at) "
            "VALUES(?, ?, ?, ?, ?, ?, ?, ?)",
            (header["number"], header["tdate"], header["to_party"], header["from_party"], header["subject"],
             header["purpose"], header["remarks"], datetime.now().strftime("%Y-%m-%d %H:%M:%S")),
        )
        tid = cur.lastrowid
        conn.executemany(
            "INSERT INTO transmittal_items(transmittal_id, file_path, file_name, dwg_no, revision, title, file_type) "
            "VALUES(?, ?, ?, ?, ?, ?, ?)",
            [(tid, r.file_path, r.file_name, r.dwg_no, r.revision, r.title, r.file_type)
             for r in items.itertuples(index=False)],
        )
        conn.commit()
    return tid


def load_transmittals() -> pd.DataFrame:
    with closing(get_db_connection()) as conn:
        return pd.read_sql_query(
            "SELECT t.*, (SELECT COUNT(*) FROM transmittal_items i WHERE i.transmittal_id = t.id) AS items "
            "FROM transmittals t ORDER BY t.id DESC", conn)


def load_transmittal_items(transmittal_id: int) -> pd.DataFrame:
    with closing(get_db_connection()) as conn:
        return pd.read_sql_query(
            "SELECT file_path, file_name, dwg_no, revision, title, file_type FROM transmittal_items "
            "WHERE transmittal_id=? ORDER BY dwg_no, file_name", conn, params=(int(transmittal_id),))


def delete_transmittal(transmittal_id: int) -> None:
    with closing(get_db_connection()) as conn:
        conn.execute("DELETE FROM transmittal_items WHERE transmittal_id=?", (int(transmittal_id),))
        conn.execute("DELETE FROM transmittals WHERE id=?", (int(transmittal_id),))
        conn.commit()


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
    result = {"dwg_no": None, "title": None, "revision": None, "issue_date": None,
              "dwg_labelled": False, "text": ""}
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
        result["text"] = full_text.upper()

        label = DWG_LABEL_REGEX.search(tb_text)
        if label and any(ch.isdigit() for ch in label.group(1)):
            result["dwg_no"] = label.group(1).upper().strip("-_./")
            result["dwg_labelled"] = True
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


# ---------- Zone (area) & discipline auto-classification ----------
ZONE_PATTERNS = [
    ("Site / External", r"\bSITE\b|EXTERNAL|LANDSCAPE|\bEXT\b|\bINFRA"),
    ("Roof", r"\bROOF\b|\bRF\b|\bRFL\b"),
    ("Basement", r"BASEMENT|\bBSMT\b|\bB[1-4]\b|\bB0[1-4]\b|\bLB[1-4]\b|\bB[1-4]F\b|\bBF\b|\bLEVEL\s*-\s*[1-4]\b"),
    ("Ground Floor", r"GROUND|\bGF\b|\bG\.F\.?|\bGFL\b|\bL00\b|\bLG\b|\bLEVEL\s*0\b|\bLVL\s*0\b"),
    ("Mezzanine", r"MEZZ|\bMZ\b"),
    ("First Floor", r"\bFIRST\b|\b1ST\b|\bFF\b|\bL01\b|\bF1\b|\bLEVEL\s*0?1\b|\bLVL\s*0?1\b"),
    ("Second Floor", r"\bSECOND\b|\b2ND\b|\bL02\b|\bF2\b|\bLEVEL\s*0?2\b|\bLVL\s*0?2\b"),
    ("Third Floor", r"\bTHIRD\b|\b3RD\b|\bL03\b|\bF3\b|\bLEVEL\s*0?3\b|\bLVL\s*0?3\b"),
    ("Typical Floor", r"TYPICAL|\bTYP\b"),
]
ZONE_PATTERNS = [(name, re.compile(pat)) for name, pat in ZONE_PATTERNS]
GENERIC_LEVEL_REGEX = re.compile(r"\b(\d{1,2})(?:ST|ND|RD|TH)\s*FLOOR\b|\bL(\d{2})\b|\bLEVEL\s*(\d{1,2})\b|\bLVL\s*(\d{1,2})\b")

DISCIPLINE_PATTERNS = [
    ("Lighting", r"LIGHTING|\bLTG\b|\bLIGHT\b|LUMINAIRE|\bLGT\b|\bLT\b"),
    ("Power", r"\bPOWER\b|\bPWR\b|SOCKET|\bSP\b|SMALL\s*POWER|\bPWS\b"),
    ("Fire Alarm", r"FIRE\s*ALARM|\bFA\b|\bFAS\b|DETECTION|SMOKE"),
    ("Fire Fighting", r"FIRE\s*FIGHT|\bFFS\b|SPRINKLER|HYDRANT|\bFP\b|FIRE\s*PROTECTION|\bFM200\b|SUPPRESSION"),
    ("HVAC", r"HVAC|\bDUCT|CHILLED|\bCHW\b|VENTILATION|\bVRF\b|\bFCU\b|\bAHU\b|AIR\s*COND|\bAC\b|MECHANICAL|\bMECH\b"),
    ("Plumbing / Drainage", r"PLUMB|DRAIN|SANITARY|WATER\s*SUPPLY|\bPLB\b|\bDWS\b|\bSWR\b|SEWER|\bDR\b|\bPL\b"),
    ("ELV / Low Current", r"\bELV\b|LOW\s*CURRENT|CCTV|\bDATA\b|TELECOM|\bTEL\b|ACCESS\s*CONTROL|\bBMS\b|PUBLIC\s*ADDRESS|\bPA\b|\bLC\b|STRUCTURED\s*CABL|\bICT\b|\bAV\b|MATV|IPTV"),
    ("Earthing / Lightning", r"EARTH|LIGHTNING|GROUNDING|\bLPS\b|\bERT\b"),
    ("Single Line / Schematics", r"\bSLD\b|SINGLE\s*LINE|SCHEMATIC|RISER|LOAD\s*SCHEDULE|PANEL\s*SCHEDULE|\bDB\s*SCHEDULE"),
    ("Containment / Cable Tray", r"CABLE\s*TRAY|TRAY|CONTAINMENT|TRUNK|CONDUIT|\bCT\b"),
    ("Electrical (general)", r"ELECTRICAL|\bELEC\b|\bELE\b|\bEL\b|\bE-\w"),
    ("Architectural", r"ARCHITECT|\bARCH\b|\bAR-\w"),
    ("Structural", r"STRUCT|\bSTR\b|\bST-\w"),
    ("Civil", r"\bCIVIL\b|\bCV\b|\bC-\w"),
]
DISCIPLINE_PATTERNS = [(name, re.compile(pat)) for name, pat in DISCIPLINE_PATTERNS]


def _classify(patterns, texts: list[str], generic=None) -> str | None:
    """Return the first pattern name found; texts are tried in priority order."""
    for text in texts:
        if not text:
            continue
        t = text.upper()
        for name, rx in patterns:
            if rx.search(t):
                return name
        if generic is not None:
            m = generic.search(t)
            if m:
                n = next(g for g in m.groups() if g)
                return f"Level {int(n):02d}"
    return None


def classify_zone(rel_dir: str, file_name: str, title: str, dwg_no: str) -> str:
    # Deepest folder first (the user's own filing), then file name, drawing number, title.
    parts = [p for p in rel_dir.replace("\\", "/").split("/") if p and p != "."]
    texts = list(reversed(parts)) + [os.path.splitext(file_name)[0], dwg_no, title]
    return _classify(ZONE_PATTERNS, texts, GENERIC_LEVEL_REGEX) or "Unassigned"


def classify_discipline(rel_dir: str, file_name: str, title: str, dwg_no: str) -> str:
    parts = [p for p in rel_dir.replace("\\", "/").split("/") if p and p != "."]
    texts = list(reversed(parts)) + [os.path.splitext(file_name)[0], dwg_no, title]
    return _classify(DISCIPLINE_PATTERNS, texts) or "Unassigned"


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
            elif pdf["dwg_labelled"] and pdf["dwg_no"] != dwg_no and dwg_no not in pdf["text"]:
                # Descriptive file name ("HVAC_B1_layout_final"): the number printed next to the
                # "DWG NO" label in the title block is the real one.
                dwg_no, source = pdf["dwg_no"], "pdf-label"
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


CDE_FOLDER_HINTS = [
    ("Archived", re.compile(r"SUPERSED|\bOLD\b|ARCHIVE|\bVOID\b|CANCEL|OBSOLETE")),
    ("Published", re.compile(r"AS[\s_-]?BUILT|\bIFC\b|\bAFC\b|FOR\s*CONSTRUCTION|APPROVED")),
    ("Shared", re.compile(r"SHOP\s*DRAW|SUBMITTAL|SUBMISSION|FOR\s*APPROVAL|FOR\s*REVIEW|\bSHARED\b")),
    ("WIP", re.compile(r"\bWIP\b|DRAFT|WORKING|IN\s*PROGRESS")),
]


def auto_cde_state(tree_path: str, status: str) -> str:
    """Best guess of the CDE state from the folder names; the user can always override it."""
    upper = tree_path.upper()
    for state, rx in CDE_FOLDER_HINTS:
        if rx.search(upper):
            return state
    return "Archived" if status == "Superseded" else UNSET


def naming_regex_from_example(example: str):
    """
    Build a naming-standard check from one correct drawing number, e.g. 'GS-E-LTG-GF-02':
    same number of parts, same separators, digits where the example has digits.
    """
    ex = (example or "").strip().upper()
    if not ex:
        return None
    out = []
    for part in re.split(r"([-_.\s])", ex):
        if not part:
            continue
        if re.fullmatch(r"[-_.\s]", part):
            out.append(re.escape(part))
        elif part.isdigit():
            out.append(rf"\d{{{len(part)},}}")
        elif part.isalpha():
            out.append(r"[A-Z][A-Z0-9]*")
        else:
            out.append(r"[A-Z0-9]+")
    return re.compile("".join(out))


@st.cache_data(show_spinner=False)
def load_documents(version: int, naming_example: str = "") -> pd.DataFrame:
    """Load the register and compute status / duplicate flags. Cached per data version."""
    with closing(get_db_connection()) as conn:
        df = pd.read_sql_query("SELECT * FROM documents", conn)
        meta = pd.read_sql_query("SELECT * FROM doc_meta", conn)
        trans = pd.read_sql_query(
            "SELECT i.file_path, t.number FROM transmittal_items i JOIN transmittals t ON t.id = i.transmittal_id "
            "ORDER BY t.id", conn)
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

    def _rel_dir(row):
        try:
            rel = os.path.relpath(os.path.dirname(row["file_path"]), row["root_folder"])
        except ValueError:
            rel = "."
        return "" if rel == "." else rel.replace("\\", "/")

    df["rel_dir"] = df.apply(_rel_dir, axis=1)
    df["tree_path"] = np.where(df["rel_dir"] == "", df["project"], df["project"] + "/" + df["rel_dir"])
    df["zone"] = [classify_zone(r, f, t, d) for r, f, t, d in
                  zip(df["rel_dir"], df["file_name"], df["title"], df["dwg_no"])]
    df["discipline"] = [classify_discipline(r, f, t, d) for r, f, t, d in
                        zip(df["rel_dir"], df["file_name"], df["title"], df["dwg_no"])]

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

    # ----- preview source: the PDF itself, or the PDF of the same drawing (same revision first) -----
    pdfs = df[(df["file_type"] == "PDF") & mask]           # already sorted: highest revision first
    by_rev = dict(zip(zip(pdfs["dwg_no"], pdfs["revision"]), pdfs["file_path"]))
    by_dwg = dict(zip(pdfs["dwg_no"][::-1], pdfs["file_path"][::-1]))   # reversed -> first (latest) wins
    df["preview_path"] = [
        path if ftype == "PDF" else (by_rev.get((dwg, rev)) or by_dwg.get(dwg) or "") if has else ""
        for path, ftype, dwg, rev, has in zip(df["file_path"], df["file_type"], df["dwg_no"], df["revision"], mask)
    ]

    # ----- workflow data entered by the user -----
    df["meta_key"] = np.where(mask, df["dwg_no"] + "|" + df["revision"], "path|" + df["file_path"])
    meta = meta.set_index("meta_key") if not meta.empty else meta
    def _meta(col, default):
        if meta.empty:
            return pd.Series(default, index=df.index)
        return df["meta_key"].map(meta[col]).where(lambda x: x.notna(), default)
    user_cde = _meta("cde_state", "").astype(str)
    auto_cde = [auto_cde_state(tp, stt) for tp, stt in zip(df["tree_path"], df["status"])]
    df["cde_auto"] = ~user_cde.isin(CDE_STATES[1:])
    df["cde_state"] = np.where(df["cde_auto"], auto_cde, user_cde)
    approval = _meta("approval_code", UNSET).astype(str)
    df["approval_code"] = np.where(approval.isin(APPROVAL_CODES), approval, UNSET)
    df["notes"] = _meta("notes", "").astype(str).replace({"None": "", "nan": ""})
    df["favorite"] = pd.to_numeric(_meta("favorite", 0), errors="coerce").fillna(0).astype(int).astype(bool)

    if trans.empty:
        df["transmittals"] = ""
    else:
        per_file = trans.groupby("file_path")["number"].agg(lambda x: ", ".join(dict.fromkeys(x)))
        df["transmittals"] = df["file_path"].map(per_file).fillna("")

    # ----- quality checks -----
    naming_rx = naming_regex_from_example(naming_example)
    if naming_rx is None:
        df["naming_ok"] = True
    else:
        df["naming_ok"] = [(not has) or bool(naming_rx.fullmatch(d)) for has, d in zip(mask, df["dwg_no"])]

    # A current drawing should exist both as PDF and as CAD source (DWG / DXF).
    drawings = df[mask & df["file_type"].isin(DRAWING_TYPES) & (df["status"] == "Current")]
    types_per_dwg = drawings.groupby("dwg_no")["file_type"].agg(set)
    def _pair(row_dwg, row_type, is_current):
        if not is_current or row_type not in DRAWING_TYPES:
            return ""
        types = types_per_dwg.get(row_dwg, set())
        if row_type == "PDF" and not (types & {"DWG", "DXF"}):
            return "No CAD source"
        if row_type in ("DWG", "DXF") and "PDF" not in types:
            return "No PDF"
        return ""
    df["pair_issue"] = [_pair(d, t, stt == "Current" and has)
                        for d, t, stt, has in zip(df["dwg_no"], df["file_type"], df["status"], mask)]

    df["_search"] = (
        df["file_name"] + " " + df["dwg_no"] + " " + df["title"] + " " + df["folder_name"]
        + " " + df["ref_dwgs"] + " REV " + df["revision"] + " " + df["zone"] + " " + df["discipline"]
        + " " + df["notes"] + " " + df["transmittals"]
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
    cols = ["status", "dwg_no", "revision", "title", "zone", "discipline", "cde_state", "approval_code",
            "file_type", "issue_date", "transmittals", "notes", "project", "folder_name", "file_name",
            "modified", "size_mb", "ref_dwgs", "file_path"]
    out = df[cols].copy()
    out["modified"] = out["modified"].dt.strftime("%Y-%m-%d %H:%M")
    out.columns = ["Status", "Drawing No", "Rev", "Title", "Zone", "Discipline", "CDE State", "Approval",
                   "Type", "Issue Date", "Transmittals", "Notes", "Project", "Folder", "File Name",
                   "Modified", "Size (MB)", "Referenced Drawings", "Path"]
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        out.to_excel(writer, index=False, sheet_name="Drawing Register")
        ws = writer.sheets["Drawing Register"]
        ws.freeze_panes = "A2"
        ws.auto_filter.ref = ws.dimensions
        widths = [12, 24, 7, 45, 14, 18, 11, 20, 7, 12, 16, 30, 18, 30, 40, 17, 10, 30, 70]
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
        log_event("open file", os.path.basename(norm_path), norm_path)
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
# 6. Previews, viewer and revision compare
# ==========================================
def _cache_file(key: str) -> str:
    return os.path.join(CACHE_DIR, hashlib.md5(key.encode("utf-8", "ignore")).hexdigest() + ".png")


@st.cache_data(show_spinner=False, max_entries=400)
def page_image(path: str, mtime: float, page_no: int = 0, max_px: int = 900) -> bytes | None:
    """PNG of one PDF page (longest side = max_px). Small sizes are also cached on disk."""
    use_disk = max_px <= 1000
    cache_path = _cache_file(f"{path}|{mtime}|{page_no}|{max_px}")
    if use_disk:
        try:
            if os.path.exists(cache_path):
                with open(cache_path, "rb") as fh:
                    return fh.read()
        except OSError:
            pass
    try:
        with fitz.open(path) as doc:
            if doc.needs_pass or page_no >= doc.page_count:
                return None
            page = doc[page_no]
            zoom = max_px / max(page.rect.width, page.rect.height, 1)
            png = page.get_pixmap(matrix=fitz.Matrix(zoom, zoom), alpha=False).tobytes("png")
    except Exception:
        return None
    if use_disk:
        try:
            os.makedirs(CACHE_DIR, exist_ok=True)
            with open(cache_path, "wb") as fh:
                fh.write(png)
        except OSError:
            pass
    return png


@st.cache_data(show_spinner=False, max_entries=2000)
def pdf_page_count(path: str, mtime: float) -> int:
    try:
        with fitz.open(path) as doc:
            return 0 if doc.needs_pass else doc.page_count
    except Exception:
        return 0


def safe_mtime(path: str) -> float | None:
    try:
        return os.path.getmtime(path)
    except OSError:
        return None


def _gray_page(path: str, page_no: int, width_px: int, height_px: int | None = None) -> np.ndarray:
    """Render a page to a grayscale array. With height_px the page is stretched to that exact size."""
    with fitz.open(path) as doc:
        page = doc[min(page_no, doc.page_count - 1)]
        if height_px is None:
            zoom = width_px / max(page.rect.width, page.rect.height, 1)
            matrix = fitz.Matrix(zoom, zoom)
        else:
            matrix = fitz.Matrix(width_px / max(page.rect.width, 1), height_px / max(page.rect.height, 1))
        pix = page.get_pixmap(matrix=matrix, colorspace=fitz.csGRAY, alpha=False)
        arr = np.frombuffer(pix.samples, dtype=np.uint8).reshape(pix.height, pix.stride)[:, :pix.width]
        return arr.copy()


def _dilate(mask: np.ndarray) -> np.ndarray:
    """3x3 binary dilation with plain numpy (absorbs 1-pixel anti-aliasing shifts)."""
    out = mask.copy()
    out[1:, :] |= mask[:-1, :]
    out[:-1, :] |= mask[1:, :]
    out[:, 1:] |= mask[:, :-1]
    out[:, :-1] |= mask[:, 1:]
    out[1:, 1:] |= mask[:-1, :-1]
    out[:-1, :-1] |= mask[1:, 1:]
    out[1:, :-1] |= mask[:-1, 1:]
    out[:-1, 1:] |= mask[1:, :-1]
    return out


@st.cache_data(show_spinner=False, max_entries=12)
def compare_pages(old_path: str, old_mtime: float, new_path: str, new_mtime: float,
                  page_old: int = 0, page_new: int = 0, max_px: int = 2200) -> tuple[bytes, float]:
    """
    Overlay two revisions of a drawing: grey = unchanged, red = only in the old revision (removed),
    blue = only in the new revision (added). Returns (png, percent of linework that changed).
    """
    new = _gray_page(new_path, page_new, max_px)
    h, w = new.shape
    old_raw = _gray_page(old_path, page_old, w, h)
    old = np.full((h, w), 255, np.uint8)
    hh, ww = min(h, old_raw.shape[0]), min(w, old_raw.shape[1])
    old[:hh, :ww] = old_raw[:hh, :ww]

    ink_old, ink_new = old < 175, new < 175
    removed = ink_old & ~_dilate(ink_new)
    added = ink_new & ~_dilate(ink_old)

    out = np.full((h, w, 3), 255, np.uint8)
    out[ink_old | ink_new] = (168, 173, 181)
    out[_dilate(removed)] = (214, 40, 40)
    out[_dilate(added)] = (21, 101, 216)
    total = max(1, int((ink_old | ink_new).sum()))
    changed = int(removed.sum()) + int(added.sum())
    png = fitz.Pixmap(fitz.csRGB, w, h, out.tobytes(), False).tobytes("png")
    return png, round(100.0 * changed / total, 2)


# ==========================================
# 7. UI building blocks
# ==========================================
TYPE_ICON = {
    "PDF": "📕", "DWG": "📐", "DXF": "📐", "XLSX": "📊", "XLS": "📊", "XLSM": "📊",
    "CSV": "📊", "DOCX": "📝", "DOC": "📝",
}
PILL_COLORS = {
    "Current": ("#E3F6E8", "#17753A"), "Superseded": ("#FDEBD3", "#A85B00"),
    "Unclassified": ("#ECEFF3", "#566070"),
    "WIP": ("#FFF3C4", "#8A6100"), "Shared": ("#DCEBFF", "#1552B0"),
    "Published": ("#E3F6E8", "#17753A"), "Archived": ("#ECEFF3", "#566070"),
    "A": ("#E3F6E8", "#17753A"), "B": ("#D5F5F0", "#0B6B5E"), "C": ("#FDEBD3", "#A85B00"),
    "D": ("#FDE2E1", "#B42318"), "F": ("#DCEBFF", "#1552B0"),
    "rev": ("#E8E7FF", "#3B34B5"), "type": ("#EEF1F5", "#3A4656"), "warn": ("#FDE2E1", "#B42318"),
    "info": ("#EEF1F5", "#3A4656"),
}

APP_CSS = """
<style>
.block-container {padding: 1.1rem 2.2rem 3rem 2.2rem; max-width: 100%;}
header[data-testid="stHeader"] {background: transparent;}
.stAppDeployButton {display: none;}

/* ---- sidebar: brand, navigation and folder tree ---- */
section[data-testid="stSidebar"] .stButton button {
    justify-content: flex-start; text-align: left; padding: 0.2rem 0.55rem; min-height: 1.9rem; border: none;
}
section[data-testid="stSidebar"] .stButton button > div,
div[class*="st-key-dcfolder_"] button > div,
div[class*="st-key-dcrow_"] button[kind="tertiary"] > div,
div[class*="st-key-dctile_"] button[kind="tertiary"] > div {
    justify-content: flex-start; width: 100%; text-align: left;
}
section[data-testid="stSidebar"] .stButton button p {
    white-space: nowrap; overflow: hidden; text-overflow: ellipsis; font-size: 0.88rem;
}
section[data-testid="stSidebar"] [data-testid="stVerticalBlock"] {gap: 0.15rem;}
p.dc-brand {font-size: 1.15rem; font-weight: 700; line-height: 1.25;}
p.dc-brand small {display: block; font-weight: 400; font-size: 0.72rem; opacity: 0.65;}
p.dc-section {font-size: 0.68rem; font-weight: 700; letter-spacing: 0.08em; opacity: 0.55;
              padding: 0.7rem 0 0 0.3rem; line-height: 1.2;}

/* ---- pills / meta text ---- */
.dc-pill {display: inline-block; padding: 1px 9px; border-radius: 999px; font-size: 0.72rem; font-weight: 600;
          line-height: 1.55; white-space: nowrap; margin: 1px 4px 1px 0;}
.dc-meta {opacity: 0.7; font-size: 0.82rem;}
.dc-dwg {font-weight: 600; font-size: 0.9rem;}

/* ---- file rows, tiles, details panel ---- */
div[class*="st-key-dcrow_"] {padding: 0.35rem 0.8rem !important;}
div[class*="st-key-dcrow_"] [data-testid="stVerticalBlock"] {gap: 0.1rem;}
div[class*="st-key-dcrow_"] button[kind="tertiary"], div[class*="st-key-dctile_"] button[kind="tertiary"] {
    text-align: left; justify-content: flex-start; font-weight: 600; min-height: 1.6rem; padding: 0;}
div[class*="st-key-dcrow_"] button[kind="tertiary"] p, div[class*="st-key-dctile_"] button[kind="tertiary"] p {
    font-weight: 600; text-align: left;}
div[class*="st-key-dctile_"] {padding: 0.6rem !important;}
div[class*="st-key-dctile_"] img {border-radius: 6px; border: 1px solid rgba(128,128,128,0.25); background: #fff;}
.dc-thumb-empty {height: 150px; display: flex; align-items: center; justify-content: center; font-size: 2.6rem;
                 border-radius: 6px; border: 1px dashed rgba(128,128,128,0.4); opacity: 0.75;}
.st-key-dc_details {position: sticky; top: 3.6rem; max-height: calc(100vh - 4.6rem); overflow-y: auto;}
.st-key-dc_details img {border: 1px solid rgba(128,128,128,0.3); border-radius: 6px; background: #fff;
                        max-height: 36vh; object-fit: contain;}
table.dc-props {width: 100%; border-collapse: collapse; font-size: 0.84rem;}
table.dc-props td {padding: 3px 4px; border-bottom: 1px solid rgba(128,128,128,0.18); vertical-align: top;
                   word-break: break-word;}
table.dc-props td:first-child {opacity: 0.65; white-space: nowrap; width: 34%;}
div[class*="st-key-dcfolder_"] button {justify-content: flex-start; text-align: left;}
[data-testid="stMetricValue"] {font-size: 1.6rem;}
</style>
"""


def inject_style(css: str) -> None:
    if hasattr(st, "html"):
        st.html(css)
    else:  # pragma: no cover
        st.markdown(css, unsafe_allow_html=True)


def md_escape(text) -> str:
    """Escape markdown control characters so file names render literally."""
    return re.sub(r"([\\`*_{}\[\]<>#|~$:])", r"\\\1", str(text))


def pill(text: str, kind: str) -> str:
    bg, fg = PILL_COLORS.get(kind, PILL_COLORS["info"])
    return f'<span class="dc-pill" style="background:{bg};color:{fg}">{html.escape(str(text))}</span>'


def approval_kind(code: str) -> str:
    return "F" if code.startswith("For") else code[:1]


def doc_pills(row, with_type: bool = False) -> str:
    parts = []
    if with_type:
        parts.append(pill(row.file_type, "type"))
    parts.append(pill(f"Rev {row.revision}", "rev"))
    parts.append(pill(row.status, row.status))
    if row.cde_state != UNSET:
        parts.append(pill(row.cde_state, row.cde_state))
    if row.approval_code != UNSET:
        label = "For info" if row.approval_code.startswith("For") else f"Code {row.approval_code[:1]}"
        parts.append(pill(label, approval_kind(row.approval_code)))
    if row.favorite:
        parts.append(pill("★", "info"))
    if row.is_conflict:
        parts.append(pill("Conflict", "warn"))
    if row.is_duplicate:
        parts.append(pill("Duplicate", "info"))
    if not row.naming_ok:
        parts.append(pill("Naming", "warn"))
    return "".join(parts)


def human_size(num_bytes) -> str:
    try:
        n = float(num_bytes)
    except (TypeError, ValueError):
        return ""
    for unit in ("B", "KB", "MB", "GB"):
        if n < 1024 or unit == "GB":
            return f"{n:.0f} {unit}" if unit in ("B", "KB") else f"{n:.1f} {unit}"
        n /= 1024
    return ""


def bump_data_version() -> None:
    st.session_state["data_version"] = st.session_state.get("data_version", 0) + 1
    load_documents.clear()


def select_doc(doc_id) -> None:
    st.session_state["selected_id"] = int(doc_id)


def clear_selection() -> None:
    st.session_state["selected_id"] = None


def add_to_basket(ids) -> None:
    basket = st.session_state.setdefault("basket", [])
    added = 0
    for i in ids:
        if int(i) not in basket:
            basket.append(int(i))
            added += 1
    st.toast(f"{added} file(s) added to the transmittal basket ({len(basket)} total)", icon="📨")


def go_to(page: str) -> None:
    st.session_state["page"] = page
    st.session_state["global_search"] = ""


def rows_per_page() -> int:
    try:
        return max(10, int(get_setting("rows_per_page", "50")))
    except ValueError:
        return 50


def scroll_image(png: bytes, zoom_pct: int, height_vh: int = 70) -> None:
    b64 = base64.b64encode(png).decode()
    body = (f'<div style="overflow:auto;height:{height_vh}vh;border:1px solid rgba(128,128,128,0.35);'
            f'border-radius:8px;background:#fff"><img src="data:image/png;base64,{b64}" '
            f'style="width:{zoom_pct}%;max-width:none;display:block"></div>')
    if hasattr(st, "html"):
        st.html(body)
    else:  # pragma: no cover
        st.markdown(body, unsafe_allow_html=True)


@st.dialog("Drawing viewer", width="large")
def viewer_dialog(path: str, title: str) -> None:
    mtime = safe_mtime(path)
    pages = pdf_page_count(path, mtime) if mtime else 0
    if not pages:
        st.error("This file cannot be displayed (missing, encrypted or not a PDF).")
        return
    st.markdown(f"**{md_escape(title)}**  \n<span class='dc-meta'>{html.escape(os.path.basename(path))}</span>",
                unsafe_allow_html=True)
    c1, c2, c3 = st.columns([1.2, 3, 1.6], vertical_alignment="bottom")
    page_no = int(c1.number_input("Page", min_value=1, max_value=pages, value=1, step=1,
                                  help=f"{pages} page(s)")) if pages > 1 else 1
    zoom = c2.select_slider("Zoom", options=[100, 150, 200, 300, 400], value=100, format_func=lambda v: f"{v}%")
    if c3.button("📂 Open in default app", **WIDE_BTN):
        open_file(path)
    png = page_image(path, mtime, page_no - 1, 1900 if zoom <= 150 else 3400)
    if png:
        scroll_image(png, zoom)
    else:
        st.error("Could not render this page.")


@st.dialog("Compare revisions", width="large")
def compare_dialog(old_path: str, old_label: str, new_path: str, new_label: str) -> None:
    old_m, new_m = safe_mtime(old_path), safe_mtime(new_path)
    if old_m is None or new_m is None:
        st.error("One of the two files no longer exists on disk. Re-scan the project.")
        return
    n_old, n_new = pdf_page_count(old_path, old_m), pdf_page_count(new_path, new_m)
    if not n_old or not n_new:
        st.error("Both revisions must be readable PDF files.")
        return
    st.markdown(
        f"{pill(old_label + ' · removed', 'D')} {pill(new_label + ' · added', 'F')} {pill('unchanged', 'info')}",
        unsafe_allow_html=True)
    c1, c2 = st.columns([1.2, 3], vertical_alignment="bottom")
    pages = min(n_old, n_new)
    page_no = int(c1.number_input("Page", min_value=1, max_value=pages, value=1, step=1)) if pages > 1 else 1
    zoom = c2.select_slider("Zoom", options=[100, 150, 200, 300, 400], value=100, format_func=lambda v: f"{v}%")
    with st.spinner("Comparing the two revisions ..."):
        try:
            png, pct = compare_pages(old_path, old_m, new_path, new_m, page_no - 1, page_no - 1)
        except Exception as exc:
            st.error(f"Comparison failed: {exc}")
            return
    if pct == 0:
        st.success("No differences found on this page.")
    else:
        st.caption(f"About {pct}% of the linework on this page differs between the two revisions.")
    scroll_image(png, zoom)
    log_event("compare revisions", os.path.basename(new_path), f"{old_label} vs {new_label}")


def paginate(df_slice: pd.DataFrame, per_page: int, state_key: str) -> pd.DataFrame:
    """Pager widget; keeps its page number in session_state[state_key]. Returns the current page."""
    total = len(df_slice)
    total_pages = max(1, -(-total // per_page))
    page = min(max(1, st.session_state.get(state_key, 1)), total_pages)
    if total_pages > 1:
        p1, p2, p3, p4 = st.columns([1, 3, 1.3, 1], vertical_alignment="center")
        if p1.button("◀ Prev", key=f"{state_key}_prev", disabled=page <= 1, **WIDE_BTN):
            page -= 1
        p2.caption(f"{total:,} files · page {page} of {total_pages}")
        page = int(p3.number_input("Page", min_value=1, max_value=total_pages, value=page, step=1,
                                   key=f"{state_key}_num_{page}", label_visibility="collapsed"))
        if p4.button("Next ▶", key=f"{state_key}_next", disabled=page >= total_pages, **WIDE_BTN):
            page += 1
    page = min(max(1, page), total_pages)
    if page != st.session_state.get(state_key, 1):
        st.session_state[state_key] = page
        if total_pages > 1:
            st.rerun()
    st.session_state[state_key] = page
    start = (page - 1) * per_page
    return df_slice.iloc[start:start + per_page]


# ---------- the three listing styles ----------
def render_rows(page_df: pd.DataFrame, key_prefix: str, compact: bool) -> None:
    selected = st.session_state.get("selected_id")
    for row in page_df.itertuples(index=False):
        icon = TYPE_ICON.get(row.file_type, "📄")
        box_key = f"dcrow_{key_prefix}_{row.id}"
        if selected == row.id:
            inject_style(f"<style>.st-key-{box_key}{{border-color:{PRIMARY_COLOR} !important;"
                         f"box-shadow:0 0 0 1px {PRIMARY_COLOR};}}</style>")
        with st.container(border=True, key=box_key):
            if compact:
                c1, c2 = st.columns([5, 1.5], vertical_alignment="center")
                c1.button(f"{icon} {md_escape(row.title)}", key=f"{key_prefix}_sel_{row.id}", type="tertiary",
                          on_click=select_doc, args=(row.id,), help="Show preview and details")
                c1.markdown(f'<span class="dc-dwg">{html.escape(row.dwg_no)}</span> &nbsp;{doc_pills(row)}',
                            unsafe_allow_html=True)
                if c2.button("Open", key=f"{key_prefix}_open_{row.id}", help=row.file_path, **WIDE_BTN):
                    open_file(row.file_path)
            else:
                c1, c2, c3, c4, c5 = st.columns([4.3, 2.4, 3.2, 1.3, 1.2], vertical_alignment="center")
                c1.button(f"{icon} {md_escape(row.title)}", key=f"{key_prefix}_sel_{row.id}", type="tertiary",
                          on_click=select_doc, args=(row.id,), help="Show preview and details")
                c1.markdown(f'<span class="dc-meta">{html.escape(row.file_name)}</span>', unsafe_allow_html=True)
                c2.markdown(f'<span class="dc-dwg">{html.escape(row.dwg_no)}</span><br>'
                            f'<span class="dc-meta">{html.escape(row.zone)} · {html.escape(row.discipline)}</span>',
                            unsafe_allow_html=True)
                c3.markdown(doc_pills(row), unsafe_allow_html=True)
                when = row.issue_date if row.issue_date != NO_DWG else (
                    f"{row.modified:%Y-%m-%d}" if pd.notna(row.modified) else "")
                c4.markdown(f'<span class="dc-meta">📅 {html.escape(when)}<br>{human_size(row.file_size)}</span>',
                            unsafe_allow_html=True)
                if c5.button("📂 Open", key=f"{key_prefix}_open_{row.id}", help=row.file_path, **WIDE_BTN):
                    open_file(row.file_path)


def render_grid(page_df: pd.DataFrame, key_prefix: str, compact: bool) -> None:
    per_row = 3 if compact else 5
    rows = list(page_df.itertuples(index=False))
    selected = st.session_state.get("selected_id")
    for start in range(0, len(rows), per_row):
        cols = st.columns(per_row)
        for col, row in zip(cols, rows[start:start + per_row]):
            box_key = f"dctile_{key_prefix}_{row.id}"
            if selected == row.id:
                inject_style(f"<style>.st-key-{box_key}{{border-color:{PRIMARY_COLOR} !important;"
                             f"box-shadow:0 0 0 1px {PRIMARY_COLOR};}}</style>")
            with col, st.container(border=True, key=box_key):
                png = None
                if row.preview_path:
                    mtime = safe_mtime(row.preview_path)
                    if mtime:
                        png = page_image(row.preview_path, mtime, 0, 420)
                if png:
                    st.image(png, **WIDE_IMG)
                else:
                    st.markdown(f'<div class="dc-thumb-empty">{TYPE_ICON.get(row.file_type, "📄")}</div>',
                                unsafe_allow_html=True)
                st.button(md_escape(row.title), key=f"{key_prefix}_sel_{row.id}", type="tertiary",
                          on_click=select_doc, args=(row.id,), help="Show preview and details")
                st.markdown(f'<span class="dc-dwg">{html.escape(row.dwg_no)}</span><br>{doc_pills(row, True)}',
                            unsafe_allow_html=True)
                if st.button("📂 Open", key=f"{key_prefix}_open_{row.id}", help=row.file_path, **WIDE_BTN):
                    open_file(row.file_path)


TABLE_COLUMNS = ["dwg_no", "revision", "title", "status", "cde_state", "approval_code", "zone", "discipline",
                 "file_type", "issue_date", "transmittals", "folder_name", "file_name", "modified", "size_mb"]


def table_column_config() -> dict:
    cc = st.column_config
    return {
        "dwg_no": cc.TextColumn("Drawing No", width="medium", pinned=True),
        "revision": cc.TextColumn("Rev", width="small"),
        "title": cc.TextColumn("Title", width="large"),
        "status": cc.TextColumn("Status", width="small"),
        "cde_state": cc.TextColumn("CDE", width="small"),
        "approval_code": cc.TextColumn("Approval", width="small"),
        "zone": cc.TextColumn("Zone", width="small"),
        "discipline": cc.TextColumn("Discipline", width="small"),
        "file_type": cc.TextColumn("Type", width="small"),
        "issue_date": cc.TextColumn("Issue date", width="small"),
        "transmittals": cc.TextColumn("Transmittals", width="small"),
        "folder_name": cc.TextColumn("Folder", width="medium"),
        "file_name": cc.TextColumn("File", width="medium"),
        "modified": cc.DatetimeColumn("Modified", format="YYYY-MM-DD HH:mm"),
        "size_mb": cc.NumberColumn("MB", format="%.2f"),
    }


def render_table(page_df: pd.DataFrame, key_prefix: str) -> None:
    sig = hashlib.md5(",".join(map(str, page_df["id"].tolist())).encode()).hexdigest()[:8]
    event = st.dataframe(
        page_df[TABLE_COLUMNS], hide_index=True, column_config=table_column_config(),
        on_select="rerun", selection_mode="multi-row", key=f"tbl_{key_prefix}_{sig}",
        height=min(720, 42 + 35 * len(page_df)), **WIDE_DF,
    )
    picked = list(event.selection.rows) if event and event.selection else []
    if picked:
        ids = page_df.iloc[picked]["id"].tolist()
        a, b, c, d = st.columns([1.6, 1.4, 1.8, 3], vertical_alignment="center")
        a.caption(f"{len(ids)} selected")
        if b.button("ℹ️ Details", key=f"{key_prefix}_tbl_det_{sig}", **WIDE_BTN):
            select_doc(ids[0])
            st.rerun()
        if c.button("📨 Add to transmittal", key=f"{key_prefix}_tbl_add_{sig}", **WIDE_BTN):
            add_to_basket(ids)
    else:
        st.caption("Tick rows to see details or add them to a transmittal.")


def render_listing(df_slice: pd.DataFrame, key_prefix: str) -> None:
    """One page of files in the view chosen in the toolbar (List / Grid / Table)."""
    if df_slice.empty:
        st.caption("No files here.")
        return
    view = st.session_state.get("view_mode") or "List"
    compact = st.session_state.get("selected_id") is not None
    per_page = rows_per_page() if view != "Grid" else 20
    page_df = paginate(df_slice, per_page, f"{key_prefix}_page")
    if view == "Table":
        render_table(page_df, key_prefix)
    elif view == "Grid":
        with st.spinner("Preparing previews ..."):
            render_grid(page_df, key_prefix, compact)
    else:
        render_rows(page_df, key_prefix, compact)


def render_grouped(df_slice: pd.DataFrame, group_cols: list[str], key_prefix: str) -> None:
    """Files split into one expander per group (zone, discipline, folder ...)."""
    if df_slice.empty:
        st.caption("No files here.")
        return
    first, rest = group_cols[0], group_cols[1:]
    keys = sorted(df_slice[first].unique().tolist(), key=lambda k: (k in ("Unassigned", UNSET), str(k)))
    for i, key in enumerate(keys):
        sub = df_slice[df_slice[first] == key]
        with st.expander(f"**{md_escape(key)}**  ·  {len(sub):,} files", expanded=len(keys) <= 6):
            if rest:
                render_grouped(sub, rest, f"{key_prefix}_{i}")
            else:
                render_listing(sub, f"{key_prefix}_{i}")


def view_toolbar(container) -> None:
    container.segmented_control("View", options=["List", "Grid", "Table"], key="view_mode",
                                label_visibility="collapsed", help="How files are displayed")


# ---------- details panel ----------
def _save_field(meta_key: str, field: str, widget_key: str, target: str) -> None:
    value = st.session_state.get(widget_key)
    if value is None or value == UNSET:
        value = ""
    save_meta(meta_key, field, value)
    log_event(f"set {field.replace('_', ' ')}", target, str(value) or "(cleared)")
    bump_data_version()


def _toggle_favorite(meta_key: str, current: bool, target: str) -> None:
    save_meta(meta_key, "favorite", 0 if current else 1)
    log_event("favourite", target, "removed" if current else "added")
    bump_data_version()


def render_details(df: pd.DataFrame, row) -> None:
    icon = TYPE_ICON.get(row.file_type, "📄")
    target = row.dwg_no if row.has_dwg_no else row.file_name
    with st.container(border=True, key="dc_details"):
        h1, h2 = st.columns([8, 1], vertical_alignment="top")
        h1.markdown(f"**{icon} {md_escape(row.title)}**  \n<span class='dc-meta'>{html.escape(row.file_name)}</span>",
                    unsafe_allow_html=True)
        h2.button("✕", key="dc_close", on_click=clear_selection, type="tertiary", help="Close details")
        st.markdown(doc_pills(row, True), unsafe_allow_html=True)

        png = None
        if row.preview_path:
            mtime = safe_mtime(row.preview_path)
            if mtime:
                png = page_image(row.preview_path, mtime, 0, 900)
        if png:
            st.image(png, **WIDE_IMG)
            if row.preview_path != row.file_path:
                st.caption(f"Preview taken from the PDF: {os.path.basename(row.preview_path)}")
        else:
            reason = "No PDF of this drawing is indexed." if row.file_type != "PDF" else "This PDF cannot be rendered."
            st.markdown(f'<div class="dc-thumb-empty">{icon}</div><span class="dc-meta">{reason}</span>',
                        unsafe_allow_html=True)

        b1, b2, b3, b4 = st.columns(4)
        if b1.button("📂 Open", key="dc_open", help=row.file_path, type="primary", **WIDE_BTN):
            open_file(row.file_path)
        if b2.button("🗂️ Folder", key="dc_folder", help="Show the file in its folder", **WIDE_BTN):
            open_folder(row.file_path)
        if b3.button("🔍 View", key="dc_view", disabled=png is None, help="Open the built-in viewer", **WIDE_BTN):
            viewer_dialog(row.preview_path, row.title)
        b4.button("★ Saved" if row.favorite else "☆ Save", key="dc_fav", on_click=_toggle_favorite,
                  args=(row.meta_key, bool(row.favorite), target), help="Add to / remove from favourites", **WIDE_BTN)
        if st.button("📨 Add to transmittal basket", key="dc_basket", **WIDE_BTN):
            add_to_basket([row.id])

        versions = df[df["dwg_no"] == row.dwg_no] if row.has_dwg_no else df[df["id"] == row.id]
        tab_info, tab_flow, tab_ver = st.tabs(["Details", "Workflow", f"Versions ({len(versions)})"])

        with tab_info:
            props = [
                ("Drawing No", row.dwg_no), ("Revision", row.revision), ("Issue date", row.issue_date),
                ("Zone", row.zone), ("Discipline", row.discipline), ("Type", row.file_type),
                ("Size", human_size(row.file_size)),
                ("Modified", f"{row.modified:%Y-%m-%d %H:%M}" if pd.notna(row.modified) else ""),
                ("Folder", row.folder_name), ("Transmittals", row.transmittals or UNSET),
            ]
            if row.ref_dwgs:
                props.append(("References", row.ref_dwgs))
            if row.pair_issue:
                props.append(("Check", row.pair_issue))
            cells = "".join(f"<tr><td>{html.escape(k)}</td><td>{html.escape(str(v))}</td></tr>" for k, v in props)
            st.markdown(f'<table class="dc-props">{cells}</table>', unsafe_allow_html=True)
            st.code(row.file_path, language=None, wrap_lines=True)

        with tab_flow:
            cde_key, app_key, note_key = f"cde_{row.id}", f"appr_{row.id}", f"notes_{row.id}"
            st.session_state[cde_key] = row.cde_state
            st.session_state[app_key] = row.approval_code
            st.session_state[note_key] = row.notes
            st.selectbox("CDE state", CDE_STATES, key=cde_key, on_change=_save_field,
                         args=(row.meta_key, "cde_state", cde_key, target),
                         help="ISO 19650: WIP = work in progress, Shared = issued for review, "
                              "Published = approved for use, Archived = superseded.")
            if row.cde_auto and row.cde_state != UNSET:
                st.caption("Detected automatically from the folder name. Choose a value to override it.")
            st.selectbox("Approval code", APPROVAL_CODES, key=app_key, on_change=_save_field,
                         args=(row.meta_key, "approval_code", app_key, target),
                         help="Consultant / client reply to the submittal.")
            st.text_area("Notes", key=note_key, on_change=_save_field, height=90,
                         args=(row.meta_key, "notes", note_key, target),
                         placeholder="Comments, consultant remarks, site instructions ...")
            if row.has_dwg_no:
                st.caption("Workflow data is shared by all files of this drawing revision (PDF and CAD).")

        with tab_ver:
            for v in versions.itertuples(index=False):
                c1, c2 = st.columns([4, 1.4], vertical_alignment="center")
                marker = "➤ " if v.id == row.id else ""
                when = v.issue_date if v.issue_date != NO_DWG else (
                    f"{v.modified:%Y-%m-%d}" if pd.notna(v.modified) else "")
                c1.markdown(f"{marker}{pill('Rev ' + v.revision, 'rev')}{pill(v.file_type, 'type')}"
                            f"{pill(v.status, v.status)}<br><span class='dc-meta'>{html.escape(when)} · "
                            f"{html.escape(v.folder_name)}</span>", unsafe_allow_html=True)
                if v.id == row.id:
                    c2.caption("shown")
                else:
                    c2.button("Select", key=f"ver_sel_{v.id}", on_click=select_doc, args=(v.id,), **WIDE_BTN)

            others = versions[(versions["file_type"] == "PDF") & (versions["file_path"] != row.preview_path)
                              & (versions["fingerprint"] != row.fingerprint)]
            if row.preview_path and not others.empty:
                st.divider()
                labels = {int(o.id): f"Rev {o.revision} · {o.folder_name}" for o in others.itertuples(index=False)}
                other_id = st.selectbox("Compare with revision:", options=list(labels), key=f"cmp_{row.id}",
                                        format_func=lambda i: labels[i])
                if st.button("🔀 Compare the two revisions", key="dc_compare", **WIDE_BTN):
                    other = others[others["id"] == other_id].iloc[0]
                    base_rows = versions[versions["file_path"] == row.preview_path]
                    base = base_rows.iloc[0] if not base_rows.empty else None
                    base_key = base["rev_key"] if base is not None else row.rev_key
                    base_rev = base["revision"] if base is not None else row.revision
                    if other["rev_key"] <= base_key:
                        compare_dialog(other["file_path"], f"Rev {other['revision']}",
                                       row.preview_path, f"Rev {base_rev}")
                    else:
                        compare_dialog(row.preview_path, f"Rev {base_rev}",
                                       other["file_path"], f"Rev {other['revision']}")
            elif len(versions) > 1:
                st.caption("Comparison needs two different PDF revisions of this drawing.")


def with_details(df: pd.DataFrame, body) -> None:
    """Render `body()`; when a file is selected, show the details panel next to it."""
    selected_id = st.session_state.get("selected_id")
    match = df[df["id"] == selected_id] if selected_id is not None else df.iloc[0:0]
    if match.empty:
        st.session_state["selected_id"] = None
        body()
        return
    left, right = st.columns([1.55, 1], gap="medium")
    with left:
        body()
    with right:
        render_details(df, next(match.itertuples(index=False)))


# ==========================================
# 8. Transmittal documents
# ==========================================
def _safe_name(text: str) -> str:
    return re.sub(r'[\\/:*?"<>|]+', "_", str(text)).strip() or "transmittal"


def build_transmittal_excel(header: dict, items: pd.DataFrame) -> bytes:
    """Printable transmittal sheet (A4 landscape) listing every document that was issued."""
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Transmittal"
    thin = Side(style="thin", color="9AA3AF")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    ws.merge_cells("A1:F1")
    ws["A1"] = "DOCUMENT TRANSMITTAL"
    ws["A1"].font = Font(bold=True, size=16)
    ws["A1"].alignment = Alignment(horizontal="center")

    info = [("Transmittal No.", header.get("number", "")), ("Date", header.get("tdate", "")),
            ("To", header.get("to_party", "")), ("From", header.get("from_party", "")),
            ("Subject", header.get("subject", "")), ("Purpose", header.get("purpose", "")),
            ("Remarks", header.get("remarks", ""))]
    r = 3
    for label, value in info:
        ws.cell(r, 1, label).font = Font(bold=True)
        ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=6)
        ws.cell(r, 2, value).alignment = Alignment(wrap_text=True, vertical="top")
        for c in range(1, 7):
            ws.cell(r, c).border = border
        r += 1

    r += 1
    heads = ["No.", "Drawing / Document No.", "Title", "Rev", "Type", "File name"]
    for c, head in enumerate(heads, start=1):
        cell = ws.cell(r, c, head)
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="1565D8")
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = border
    for i, it in enumerate(items.itertuples(index=False), start=1):
        r += 1
        values = [i, it.dwg_no, it.title, it.revision, it.file_type, it.file_name]
        for c, val in enumerate(values, start=1):
            cell = ws.cell(r, c, val)
            cell.border = border
            cell.alignment = Alignment(wrap_text=True, vertical="top",
                                       horizontal="center" if c in (1, 4, 5) else "left")
    r += 2
    ws.cell(r, 1, f"Total documents: {len(items)}").font = Font(bold=True)
    r += 2
    ws.cell(r, 1, "Issued by:")
    ws.cell(r, 3, "Signature: ____________________     Date: ______________")
    r += 2
    ws.cell(r, 1, "Received by:")
    ws.cell(r, 3, "Signature: ____________________     Date: ______________")

    for col, width in zip("ABCDEF", [16, 28, 48, 7, 8, 44]):
        ws.column_dimensions[col].width = width
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


def export_transmittal_files(header: dict, items: pd.DataFrame, dest: str) -> tuple[int, list[str]]:
    """Copy the issued files and the transmittal sheet into `dest`. Returns (copied, missing names)."""
    os.makedirs(dest, exist_ok=True)
    copied, missing = 0, []
    for it in items.itertuples(index=False):
        if not os.path.exists(it.file_path):
            missing.append(it.file_name)
            continue
        target = os.path.join(dest, it.file_name)
        if os.path.exists(target) and os.path.getsize(target) != os.path.getsize(it.file_path):
            stem, ext = os.path.splitext(it.file_name)
            target = os.path.join(dest, f"{stem} ({copied + 1}){ext}")
        shutil.copy2(it.file_path, target)
        copied += 1
    sheet = os.path.join(dest, f"{_safe_name(header['number'])} Transmittal.xlsx")
    with open(sheet, "wb") as fh:
        fh.write(build_transmittal_excel(header, items))
    return copied, missing


# ==========================================
# 9. Pages
# ==========================================
PAGES = [("Files", "📂"), ("Register", "📋"), ("Transmittals", "📨"), ("Issues", "⚠️"),
         ("Dashboard", "📊"), ("Activity", "🕘"), ("Settings", "⚙️")]


def apply_theme() -> None:
    """Blue, clean theme. Written to .streamlit/config.toml once and applied at runtime."""
    theme = {"primaryColor": PRIMARY_COLOR, "backgroundColor": "#FFFFFF",
             "secondaryBackgroundColor": "#F3F5F9", "textColor": "#1B2430"}
    try:
        cfg_dir = os.path.join(BASE_DIR, ".streamlit")
        cfg = os.path.join(cfg_dir, "config.toml")
        if not os.path.exists(cfg):
            os.makedirs(cfg_dir, exist_ok=True)
            with open(cfg, "w", encoding="utf-8") as fh:
                fh.write("[theme]\n" + "".join(f'{k} = "{v}"\n' for k, v in theme.items())
                         + "\n[browser]\ngatherUsageStats = false\n")
    except OSError:
        pass
    try:
        cfg_api = st._config  # noqa: SLF001 - runtime theme switch
        if cfg_api.get_option("theme.primaryColor") is None:
            for key, value in theme.items():
                cfg_api.set_option(f"theme.{key}", value)
            if not st.session_state.get("_theme_applied"):
                st.session_state["_theme_applied"] = True
                st.rerun()
    except Exception:
        pass


def browse_to(path: str) -> None:
    """Callback: open a folder in the Files page and expand its ancestors in the tree."""
    ss = st.session_state
    opened = ss.setdefault("tree_open", set())
    parts = [p for p in path.split("/") if p]
    for i in range(1, len(parts) + 1):
        opened.add("/".join(parts[:i]))
    ss["browse_path"] = path
    ss["page"] = "Files"
    ss["global_search"] = ""
    ss["selected_id"] = None


def tree_click(path: str) -> None:
    """Callback for the sidebar tree: select the folder, toggle it when clicked again."""
    ss = st.session_state
    opened = ss.setdefault("tree_open", set())
    if ss.get("browse_path") == path and ss.get("page") == "Files" and path in opened:
        opened.discard(path)
        return
    browse_to(path)


def build_tree(counts: dict) -> dict:
    root = {"children": {}, "count": 0}
    for path, n in counts.items():
        node = root
        node["count"] += int(n)
        for part in path.split("/"):
            node = node["children"].setdefault(part, {"children": {}, "count": 0})
            node["count"] += int(n)
    return root


def render_tree(children: dict, parent: str, depth: int) -> None:
    ss = st.session_state
    opened = ss.setdefault("tree_open", set())
    for name in sorted(children, key=str.lower):
        node = children[name]
        path = f"{parent}/{name}" if parent else name
        has_kids = bool(node["children"])
        is_open = path in opened
        arrow = ("▾" if is_open else "▸") if has_kids else "•"
        label = f"{'　' * depth}{arrow} 📁 {md_escape(name)} · {node['count']:,}"
        active = ss.get("browse_path") == path and ss.get("page") == "Files"
        st.button(label, key=f"tree_{path}", on_click=tree_click, args=(path,),
                  type="primary" if active else "tertiary", **WIDE_BTN)
        if has_kids and is_open:
            render_tree(node["children"], path, depth + 1)


def run_scan(folder_path: str, deep: bool, purge: bool) -> None:
    if not folder_path or not os.path.isdir(folder_path):
        st.error("Invalid folder path. Please verify the directory exists and is accessible.")
        return
    set_setting("last_root", folder_path)
    prog_bar = st.progress(0.0)
    stat_txt = st.empty()
    with st.spinner("Scanning project folder ..."):
        stats = scan_project_folder(folder_path, deep=deep, purge_missing=purge,
                                    progress_bar=prog_bar, status_text=stat_txt)
    stat_txt.empty()
    prog_bar.empty()
    stats["when"] = datetime.now().strftime("%Y-%m-%d %H:%M")
    st.session_state["last_scan"] = stats
    log_event("scan", folder_path, f"found {stats['total']}, updated {stats['updated']}, "
                                   f"unchanged {stats['skipped']}, removed {stats['removed']}")
    bump_data_version()
    st.rerun()


def render_scanner(key_prefix: str) -> None:
    raw = st.text_input("Project root folder", value=get_setting("last_root", ""),
                        placeholder=r"C:\Users\<you>\Desktop\GS project", key=f"{key_prefix}_root",
                        help="Paste the folder path from Windows Explorer (quotes are removed automatically).")
    folder_path = os.path.expandvars(os.path.expanduser(raw.strip().strip('"').strip("'")))
    c1, c2 = st.columns(2)
    deep = c1.toggle("Deep scan (re-read unchanged files)", value=False, key=f"{key_prefix}_deep",
                     help="Use after upgrading the app or when title-block data looks wrong.")
    purge = c2.toggle("Remove entries whose files were deleted", value=True, key=f"{key_prefix}_purge")
    if st.button("🔍 Scan & index", type="primary", key=f"{key_prefix}_scan"):
        run_scan(folder_path, deep, purge)
    if "last_scan" in st.session_state:
        s = st.session_state["last_scan"]
        st.success(f"Last scan {s.get('when', '')} · {s['seconds']}s · found **{s['total']:,}** · "
                   f"added/updated **{s['updated']:,}** · unchanged **{s['skipped']:,}** · "
                   f"removed **{s['removed']:,}**" + (f" · errors **{s['errors']:,}**" if s["errors"] else ""))


def render_sidebar(df: pd.DataFrame) -> None:
    ss = st.session_state
    with st.sidebar:
        st.markdown(f'<p class="dc-brand">🏗️ {APP_TITLE}<small>Local EDMS · v{APP_VERSION}</small></p>',
                    unsafe_allow_html=True)
        st.markdown('<p class="dc-section">WORKSPACE</p>', unsafe_allow_html=True)
        basket = len(ss.get("basket", []))
        for name, icon in PAGES:
            label = f"{icon} {name}"
            if name == "Transmittals" and basket:
                label += f" · {basket}"
            st.button(label, key=f"nav_{name}", on_click=go_to, args=(name,),
                      type="primary" if ss.get("page") == name and not ss.get("global_search") else "tertiary",
                      **WIDE_BTN)

        if not df.empty:
            st.markdown('<p class="dc-section">PROJECT FOLDERS</p>', unsafe_allow_html=True)
            tree = build_tree(df["tree_path"].value_counts().to_dict())
            if len(tree["children"]) == 1 and not ss.get("tree_open"):
                ss.setdefault("tree_open", set()).add(next(iter(tree["children"])))
            render_tree(tree["children"], "", 0)

            st.markdown('<p class="dc-section">INDEX</p>', unsafe_allow_html=True)
            last_root = get_setting("last_root", "")
            if st.button("🔄 Rescan project", key="quick_rescan", disabled=not last_root, type="tertiary",
                         help=f"Index new and changed files in {last_root}", **WIDE_BTN):
                run_scan(last_root, False, True)
            if "last_scan" in ss:
                s = ss["last_scan"]
                st.caption(f"Last scan {s.get('when', '')}: {s['updated']:,} updated, {s['removed']:,} removed")


def render_breadcrumb(current: str, single_root: bool = False) -> None:
    crumbs = [c for c in current.split("/") if c]
    widths = [1.2] + [max(1.5, min(4.0, 0.9 + len(c) * 0.16)) for c in crumbs]
    cols = st.columns(widths + [max(0.5, 13 - sum(widths))])
    home = crumbs[0] if single_root and crumbs else ""
    cols[0].button("🏠 Home", key="bc_home", disabled=current == home, on_click=browse_to, args=(home,), **WIDE_BTN)
    for i, crumb in enumerate(crumbs):
        target = "/".join(crumbs[: i + 1])
        cols[i + 1].button(f"📁 {md_escape(crumb)}", key=f"bc_{i}", disabled=target == current,
                           on_click=browse_to, args=(target,), **WIDE_BTN)


GROUP_OPTIONS = {"Folder": ["folder_name"], "Zone": ["zone"], "Discipline": ["discipline"],
                 "Zone → Discipline": ["zone", "discipline"], "No grouping": None}


def page_files(df: pd.DataFrame) -> None:
    ss = st.session_state
    current = ss.get("browse_path", "")
    all_paths = df["tree_path"].unique().tolist()
    if current and not any(p == current or p.startswith(current + "/") for p in all_paths):
        current = ss["browse_path"] = ""
    roots = {p.split("/")[0] for p in all_paths}
    if not current and len(roots) == 1:                 # single project: skip the empty home level
        current = ss["browse_path"] = next(iter(roots))
    render_breadcrumb(current, single_root=len(roots) == 1)

    prefix = current + "/" if current else ""
    in_scope = df[(df["tree_path"] == current) | df["tree_path"].str.startswith(prefix)] if current else df
    children: dict[str, int] = {}
    for tp, n in in_scope["tree_path"].value_counts().items():
        remainder = tp[len(prefix):] if current else tp
        if remainder and tp != current:
            child = remainder.split("/")[0]
            children[child] = children.get(child, 0) + int(n)
    direct = in_scope[in_scope["tree_path"] == current] if current else df.iloc[0:0]

    t1, t2, t3, t4, t5 = st.columns([2.1, 1.4, 1.6, 1.9, 2.2], vertical_alignment="center")
    view_toolbar(t1)
    latest_only = t2.toggle("Latest only", value=True, key="files_latest",
                            help="Hide superseded revisions.")
    include_sub = t3.toggle("Sub-folders", value=False, key="files_incl",
                            help="Show every file under this folder, grouped.")
    group_mode = t4.selectbox("Group by", options=list(GROUP_OPTIONS), key="files_group",
                              label_visibility="collapsed", disabled=not include_sub,
                              help="How files are grouped when sub-folders are included")
    if current and not in_scope.empty:
        # tree path = project name + relative folders -> absolute folder on disk
        abs_folder = os.path.join(in_scope.iloc[0]["root_folder"], *current.split("/")[1:])
        if t5.button("🗂️ Open in Explorer", key="open_browse_folder", **WIDE_BTN):
            open_folder(os.path.join(abs_folder, "_"))

    if children:
        st.markdown('<p class="dc-section">FOLDERS</p>', unsafe_allow_html=True)
        names = sorted(children.items(), key=lambda kv: kv[0].lower())
        for start in range(0, len(names), 4):
            grid = st.columns(4)
            for i, (name, n) in enumerate(names[start:start + 4]):
                grid[i].button(f"📁 {md_escape(name)}  ·  {n:,}", key=f"dcfolder_{start + i}",
                               on_click=browse_to, args=(prefix + name,), **WIDE_BTN)

    files = in_scope if include_sub else direct
    if latest_only:
        files = files[files["status"] != "Superseded"]
    files = files.sort_values(["dwg_no", "rev_tier", "rev_num", "rev_suffix", "file_name"],
                              ascending=[True, False, False, False, True])

    if files.empty and children and not include_sub:
        st.caption("Choose a folder above, or switch on **Sub-folders** to list everything below this level.")
        with_details(df, lambda: None)
        return

    h1, h2 = st.columns([5, 1.8], vertical_alignment="center")
    hidden = (len(in_scope if include_sub else direct) - len(files)) if latest_only else 0
    h1.markdown(f'<p class="dc-section">FILES · {len(files):,}'
                + (f" &nbsp;({hidden:,} superseded hidden)" if hidden else "") + "</p>", unsafe_allow_html=True)
    if len(files) and h2.button("📨 Add all to transmittal", key="files_add_all", **WIDE_BTN):
        add_to_basket(files["id"].tolist())

    group_cols = GROUP_OPTIONS[group_mode] if include_sub else None

    def body() -> None:
        if group_cols:
            render_grouped(files, group_cols, "files")
        else:
            render_listing(files, "files")

    with_details(df, body)


def page_search(df: pd.DataFrame, query: str) -> None:
    results = apply_search(df, query)
    t1, t2, t3 = st.columns([2.1, 1.6, 5], vertical_alignment="center")
    view_toolbar(t1)
    if t2.toggle("Latest only", value=False, key="search_latest", help="Hide superseded revisions."):
        results = results[results["status"] != "Superseded"]
    t3.markdown(f"**{len(results):,}** result(s) for “{html.escape(query)}”", unsafe_allow_html=True)
    with_details(df, lambda: render_listing(results, "search"))


def page_register(df: pd.DataFrame) -> None:
    c1, c2, c3, c4 = st.columns([1.6, 1.6, 1.4, 1.4])
    zones = sorted(df["zone"].unique().tolist(), key=lambda z: (z == "Unassigned", z))
    selected_zones = c1.multiselect("Zone / floor", options=zones, placeholder="All zones", key="f_zone")
    discs = sorted(df["discipline"].unique().tolist(), key=lambda z: (z == "Unassigned", z))
    selected_discs = c2.multiselect("Discipline", options=discs, placeholder="All disciplines", key="f_disc")
    status_choice = c3.selectbox("Status", options=["Current (latest rev)", "All", "Superseded", "Unclassified"],
                                 key="status")
    selected_cde = c4.multiselect("CDE state", options=CDE_STATES, placeholder="Any", key="f_cde")

    c5, c6, c7, c8, c9 = st.columns([1.6, 1.2, 1.1, 1.5, 1.6])
    selected_appr = c5.multiselect("Approval", options=APPROVAL_CODES, placeholder="Any", key="f_appr")
    types = sorted(df["file_type"].unique().tolist())
    selected_types = c6.multiselect("Type", options=types, placeholder="All", key="f_type")
    revs = sorted(df["revision"].unique().tolist(), key=revision_sort_rank, reverse=True)
    selected_revs = c7.multiselect("Revision", options=revs, placeholder="All", key="f_rev")
    extra = c8.multiselect("Flags", options=["Favourites", "Duplicates", "Revision conflicts", "Naming issues",
                                             "Not transmitted yet"], placeholder="None", key="f_flags")
    group_by = c9.selectbox("Group by", options=["No grouping", "Zone → Discipline", "Zone", "Discipline", "Folder"],
                            key="group_by")

    filtered = df
    if selected_zones:
        filtered = filtered[filtered["zone"].isin(selected_zones)]
    if selected_discs:
        filtered = filtered[filtered["discipline"].isin(selected_discs)]
    if status_choice == "Current (latest rev)":
        filtered = filtered[filtered["status"] != "Superseded"]
    elif status_choice in ("Superseded", "Unclassified"):
        filtered = filtered[filtered["status"] == status_choice]
    if selected_cde:
        filtered = filtered[filtered["cde_state"].isin(selected_cde)]
    if selected_appr:
        filtered = filtered[filtered["approval_code"].isin(selected_appr)]
    if selected_types:
        filtered = filtered[filtered["file_type"].isin(selected_types)]
    if selected_revs:
        filtered = filtered[filtered["revision"].isin(selected_revs)]
    if "Favourites" in extra:
        filtered = filtered[filtered["favorite"]]
    if "Duplicates" in extra:
        filtered = filtered[filtered["is_duplicate"]]
    if "Revision conflicts" in extra:
        filtered = filtered[filtered["is_conflict"]]
    if "Naming issues" in extra:
        filtered = filtered[~filtered["naming_ok"]]
    if "Not transmitted yet" in extra:
        filtered = filtered[filtered["transmittals"] == ""]
    filtered = filtered.sort_values(["has_dwg_no", "dwg_no", "rev_tier", "rev_num", "rev_suffix"],
                                    ascending=[False, True, False, False, False], kind="stable")

    h0, h1, h2, h3, h4 = st.columns([2.1, 2.3, 1.5, 1.0, 1.7], vertical_alignment="center")
    view_toolbar(h0)
    h1.markdown(f"**{len(filtered):,}** of {len(df):,} files")
    if len(filtered):
        stamp = f"{datetime.now():%Y%m%d_%H%M}"
        h2.download_button("⬇️ Excel register", data=build_excel_register(filtered),
                           file_name=f"drawing_register_{stamp}.xlsx", key="dl_xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", **WIDE_DL)
        csv_cols = ["status", "dwg_no", "revision", "title", "zone", "discipline", "cde_state", "approval_code",
                    "file_type", "issue_date", "transmittals", "folder_name", "file_name", "file_path"]
        h3.download_button("⬇️ CSV", data=filtered[csv_cols].to_csv(index=False).encode("utf-8-sig"),
                           file_name=f"drawing_register_{stamp}.csv", mime="text/csv", key="dl_csv", **WIDE_DL)
        if h4.button("📨 Add all to transmittal", key="reg_add_all", **WIDE_BTN):
            add_to_basket(filtered["id"].tolist())

    group_cols = {"Zone → Discipline": ["zone", "discipline"], "Zone": ["zone"], "Discipline": ["discipline"],
                  "Folder": ["folder_name"], "No grouping": None}[group_by]

    def body() -> None:
        if filtered.empty:
            st.warning("No files match the current filters.")
        elif group_cols:
            render_grouped(filtered, group_cols, "reg")
        else:
            render_listing(filtered, "reg")

    with_details(df, body)


def page_transmittals(df: pd.DataFrame) -> None:
    ss = st.session_state
    if "_tr_tab_next" in ss:                      # set before the widget is created
        ss["tr_tab"] = ss.pop("_tr_tab_next")
    ss.setdefault("tr_tab", "New transmittal")
    tab = st.segmented_control("Section", options=["New transmittal", "History"], key="tr_tab",
                               label_visibility="collapsed") or "New transmittal"

    if tab == "New transmittal":
        basket_ids = ss.get("basket", [])
        basket_df = df[df["id"].isin(basket_ids)] if not df.empty else df
        if basket_df.empty:
            st.info("The transmittal basket is empty. Open **Files** or **Register**, pick the drawings and "
                    "click **Add to transmittal**.")
            return
        st.markdown(f"**{len(basket_df):,} document(s) in the basket**")
        event = st.dataframe(
            basket_df[["dwg_no", "revision", "title", "file_type", "status", "cde_state", "file_name"]],
            hide_index=True, on_select="rerun", selection_mode="multi-row", key=f"basket_tbl_{len(basket_df)}",
            column_config={"dwg_no": "Drawing No", "revision": "Rev", "title": "Title", "file_type": "Type",
                           "status": "Status", "cde_state": "CDE", "file_name": "File"},
            height=min(420, 42 + 35 * len(basket_df)), **WIDE_DF)
        picked = list(event.selection.rows) if event and event.selection else []
        b1, b2, _ = st.columns([1.6, 1.4, 5])
        if b1.button("Remove selected", disabled=not picked, key="basket_remove", **WIDE_BTN):
            drop = set(basket_df.iloc[picked]["id"].tolist())
            ss["basket"] = [i for i in basket_ids if i not in drop]
            st.rerun()
        if b2.button("Clear basket", key="basket_clear", **WIDE_BTN):
            ss["basket"] = []
            st.rerun()
        superseded = int((basket_df["status"] == "Superseded").sum())
        if superseded:
            st.warning(f"{superseded} document(s) in the basket are superseded revisions. "
                       "Check that you are issuing the latest revision.")

        with st.form("transmittal_form"):
            f1, f2, f3 = st.columns([1.2, 1.2, 2])
            number = f1.text_input("Transmittal No.", value=next_transmittal_number())
            tdate = f2.date_input("Date", value=datetime.now().date())
            purpose = f3.selectbox("Purpose of issue", options=TRANSMITTAL_PURPOSES)
            f4, f5 = st.columns(2)
            to_party = f4.text_input("To (company / person)", placeholder="e.g. Consultant - Electrical Dept.")
            from_party = f5.text_input("From", value=get_setting("from_party", ""))
            subject = st.text_input("Subject", placeholder="e.g. Basement lighting shop drawings")
            remarks = st.text_area("Remarks", height=70)
            mark_shared = st.checkbox("Set the CDE state of these documents to **Shared**", value=True,
                                      help="Applied only to documents that are not Published or Archived.")
            submitted = st.form_submit_button("📨 Create transmittal", type="primary")
        if submitted:
            number = number.strip()
            existing = set(load_transmittals()["number"].tolist())
            if not number:
                st.error("The transmittal number is required.")
            elif number in existing:
                st.error(f"Transmittal {number} already exists. Choose another number.")
            else:
                header = {"number": number, "tdate": str(tdate), "to_party": to_party.strip(),
                          "from_party": from_party.strip(), "subject": subject.strip(), "purpose": purpose,
                          "remarks": remarks.strip()}
                create_transmittal(header, basket_df)
                set_setting("from_party", header["from_party"])
                if mark_shared:
                    for key in basket_df.loc[~basket_df["cde_state"].isin(["Published", "Archived"]),
                                             "meta_key"].unique():
                        save_meta(key, "cde_state", "Shared")
                log_event("create transmittal", number, f"{len(basket_df)} document(s) to {header['to_party']}")
                ss["basket"] = []
                ss["_tr_tab_next"] = "History"
                ss["tr_created"] = number
                bump_data_version()
                st.rerun()
        return

    # ----- history -----
    transmittals = load_transmittals()
    if transmittals.empty:
        st.info("No transmittals yet.")
        return
    if ss.get("tr_created"):
        st.success(f"Transmittal {ss.pop('tr_created')} created.")
    st.dataframe(
        transmittals[["number", "tdate", "to_party", "subject", "purpose", "items"]], hide_index=True,
        column_config={"number": "No.", "tdate": "Date", "to_party": "To", "subject": "Subject",
                       "purpose": "Purpose", "items": "Docs"},
        height=min(300, 42 + 35 * len(transmittals)), **WIDE_DF)

    numbers = transmittals["number"].tolist()
    chosen = st.selectbox("Open transmittal", options=numbers, key="tr_open")
    head = transmittals[transmittals["number"] == chosen].iloc[0]
    items = load_transmittal_items(head["id"])
    header = {k: (head[k] or "") for k in ("number", "tdate", "to_party", "from_party", "subject", "purpose", "remarks")}

    with st.container(border=True):
        st.markdown(f"#### {html.escape(header['number'])} &nbsp; {pill(header['purpose'], 'Shared')}",
                    unsafe_allow_html=True)
        st.markdown(f"**Date:** {header['tdate']} &nbsp;·&nbsp; **To:** {md_escape(header['to_party'] or UNSET)} "
                    f"&nbsp;·&nbsp; **From:** {md_escape(header['from_party'] or UNSET)}  \n"
                    f"**Subject:** {md_escape(header['subject'] or UNSET)}"
                    + (f"  \n**Remarks:** {md_escape(header['remarks'])}" if header["remarks"] else ""))
        st.dataframe(items[["dwg_no", "revision", "title", "file_type", "file_name"]], hide_index=True,
                     column_config={"dwg_no": "Drawing No", "revision": "Rev", "title": "Title",
                                    "file_type": "Type", "file_name": "File"},
                     height=min(360, 42 + 35 * len(items)), **WIDE_DF)

        d1, d2 = st.columns([1.5, 4], vertical_alignment="bottom")
        d1.download_button("⬇️ Transmittal sheet (Excel)", data=build_transmittal_excel(header, items),
                           file_name=f"{_safe_name(header['number'])} Transmittal.xlsx", key="tr_dl",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", **WIDE_DL)
        default_dest = os.path.join(BASE_DIR, "Transmittals", _safe_name(header["number"]))
        dest = d2.text_input("Export folder (files + transmittal sheet)", value=default_dest,
                             key=f"tr_dest_{head['id']}")
        e1, e2, e3 = st.columns([1.5, 1.5, 3])
        if e1.button("📦 Copy files to folder", key="tr_export", **WIDE_BTN):
            try:
                copied, missing = export_transmittal_files(header, items, dest.strip().strip('"'))
                log_event("export transmittal", header["number"], f"{copied} file(s) to {dest}")
                st.success(f"{copied} file(s) and the transmittal sheet were copied to: {dest}")
                if missing:
                    st.warning("Not found on disk: " + ", ".join(missing))
            except OSError as exc:
                st.error(f"Export failed: {exc}")
        if e2.button("🗂️ Open export folder", key="tr_open_dest", **WIDE_BTN):
            open_folder(os.path.join(dest.strip().strip('"'), "_"))
        with e3.popover("🗑️ Delete this transmittal"):
            st.write(f"Delete {header['number']} from the history? Exported files are not touched.")
            if st.button("Yes, delete it", key="tr_delete", type="primary"):
                delete_transmittal(head["id"])
                log_event("delete transmittal", header["number"])
                bump_data_version()
                st.rerun()


def page_issues(df: pd.DataFrame) -> None:
    no_dwg = df[~df["has_dwg_no"]]
    dups = df[df["is_duplicate"]].sort_values(["fingerprint", "file_path"])
    conflicts = df[df["is_conflict"]].sort_values(["dwg_no", "revision", "file_path"])
    naming = df[~df["naming_ok"]]
    pairs = df[df["pair_issue"] != ""]
    unassigned = df[(df["zone"] == "Unassigned") | (df["discipline"] == "Unassigned")]

    m = st.columns(6)
    m[0].metric("Revision conflicts", f"{len(conflicts):,}", help="Same number and revision, different content.")
    m[1].metric("Naming issues", f"{len(naming):,}", help="Drawing numbers that do not follow the naming standard.")
    m[2].metric("Missing CAD / PDF", f"{len(pairs):,}", help="Current drawings without a CAD source or without a PDF.")
    m[3].metric("Duplicate copies", f"{len(dups):,}")
    m[4].metric("No drawing number", f"{len(no_dwg):,}")
    m[5].metric("Zone / discipline unknown", f"{len(unassigned):,}")

    cols = ["dwg_no", "revision", "file_type", "file_name", "folder_name", "modified", "file_path"]
    with st.expander(f"⚠️ Revision conflicts ({len(conflicts):,})", expanded=bool(len(conflicts))):
        st.caption("Two files claim the same drawing number and revision but their contents differ. "
                   "One of them needs a new revision number.")
        if len(conflicts):
            st.dataframe(conflicts[cols], hide_index=True, **WIDE_DF)
    with st.expander(f"🔤 Naming standard ({len(naming):,})", expanded=bool(len(naming))):
        example = get_setting("naming_example", "")
        if not example:
            st.caption("No naming standard is set. Enter one correct drawing number in **Settings** "
                       "to check all numbers against it.")
        else:
            st.caption(f"Checked against the pattern of `{example}` (same parts, same separators).")
            if len(naming):
                st.dataframe(naming[cols], hide_index=True, **WIDE_DF)
    with st.expander(f"📐 Missing CAD source or PDF ({len(pairs):,})"):
        st.caption("Every current drawing should exist as a PDF and as its CAD source (DWG / DXF).")
        if len(pairs):
            st.dataframe(pairs[["pair_issue"] + cols], hide_index=True,
                         column_config={"pair_issue": "Problem"}, **WIDE_DF)
    with st.expander(f"🔁 Duplicate copies ({len(dups):,})"):
        st.caption("Identical files stored in more than one place.")
        if len(dups):
            st.dataframe(dups[cols], hide_index=True, **WIDE_DF)
    with st.expander(f"⚪ Files without a drawing number ({len(no_dwg):,})"):
        if len(no_dwg):
            st.dataframe(no_dwg[["file_type", "file_name", "folder_name", "meta_source", "file_path"]],
                         hide_index=True, **WIDE_DF)
    with st.expander(f"🏷️ Zone or discipline not detected ({len(unassigned):,})"):
        st.caption("Detected from folder names, file names and titles. Rename the folder "
                   "(e.g. 'Basement', 'GF', 'Lighting') and re-scan to classify these.")
        if len(unassigned):
            st.dataframe(unassigned[["zone", "discipline", "file_name", "folder_name", "file_path"]],
                         hide_index=True, **WIDE_DF)


def page_dashboard(df: pd.DataFrame) -> None:
    # One row per current drawing revision: the PDF and the DWG of a drawing count once.
    cur = df[df["status"] == "Current"].drop_duplicates("meta_key")
    drawings = cur[cur["file_type"].isin(DRAWING_TYPES)].drop_duplicates("dwg_no")
    k = st.columns(7)
    k[0].metric("Indexed files", f"{len(df):,}")
    k[1].metric("Drawing numbers", f"{df.loc[df['has_dwg_no'], 'dwg_no'].nunique():,}")
    k[2].metric("Superseded files", f"{(df['status'] == 'Superseded').sum():,}")
    k[3].metric("Shared (in review)", f"{(cur['cde_state'] == 'Shared').sum():,}")
    k[4].metric("Published", f"{(cur['cde_state'] == 'Published').sum():,}")
    k[5].metric("Approved (A / B)", f"{cur['approval_code'].str[:1].isin(['A', 'B']).sum():,}")
    k[6].metric("To resubmit (C / D)", f"{cur['approval_code'].str[:1].isin(['C', 'D']).sum():,}")

    g1, g2 = st.columns(2)
    with g1:
        st.markdown("**Current drawings per zone**")
        st.bar_chart(drawings["zone"].value_counts(), horizontal=True, color=PRIMARY_COLOR)
        st.markdown("**CDE state (current drawings)**")
        st.bar_chart(cur["cde_state"].replace(UNSET, "Not set").value_counts(), horizontal=True,
                     color=PRIMARY_COLOR)
    with g2:
        st.markdown("**Current drawings per discipline**")
        st.bar_chart(drawings["discipline"].value_counts(), horizontal=True, color=PRIMARY_COLOR)
        st.markdown("**Approval codes (current drawings)**")
        st.bar_chart(cur["approval_code"].replace(UNSET, "No reply yet").value_counts(), horizontal=True,
                     color=PRIMARY_COLOR)

    st.markdown("**Zone × discipline matrix (current drawing numbers)**")
    if not drawings.empty:
        st.dataframe(pd.crosstab(drawings["zone"], drawings["discipline"]), **WIDE_DF)

    st.markdown("**Recently modified files**")
    recent = df.sort_values("last_modified", ascending=False).head(15)
    st.dataframe(recent[["modified", "dwg_no", "revision", "title", "status", "folder_name"]], hide_index=True,
                 column_config={"modified": st.column_config.DatetimeColumn("Modified", format="YYYY-MM-DD HH:mm"),
                                "dwg_no": "Drawing No", "revision": "Rev", "title": "Title", "status": "Status",
                                "folder_name": "Folder"}, **WIDE_DF)


def page_activity() -> None:
    log = load_activity(2000)
    if log.empty:
        st.info("No activity recorded yet. Opening files, changing states, scans and transmittals are logged here.")
        return
    c1, c2 = st.columns([3, 1], vertical_alignment="bottom")
    actions = c1.multiselect("Action", options=sorted(log["action"].unique().tolist()), placeholder="All actions")
    if actions:
        log = log[log["action"].isin(actions)]
    c2.download_button("⬇️ Export log (CSV)", data=log.to_csv(index=False).encode("utf-8-sig"),
                       file_name=f"activity_log_{datetime.now():%Y%m%d_%H%M}.csv", mime="text/csv", **WIDE_DL)
    st.dataframe(log, hide_index=True, height=620,
                 column_config={"ts": "Time", "user": "User", "action": "Action", "target": "Document",
                                "details": st.column_config.TextColumn("Details", width="large")}, **WIDE_DF)


def page_settings(df: pd.DataFrame) -> None:
    st.markdown("#### Project index")
    render_scanner("settings")

    st.divider()
    st.markdown("#### Naming standard")
    st.caption("Type one correct drawing number. Every indexed number must have the same parts and separators; "
               "the rest appear under **Issues**.")
    n1, n2 = st.columns([3, 1], vertical_alignment="bottom")
    example = n1.text_input("Example of a correct drawing number", value=get_setting("naming_example", ""),
                            placeholder="GS-E-LTG-GF-02", key="naming_input")
    if n2.button("Save", key="naming_save", **WIDE_BTN):
        set_setting("naming_example", example.strip().upper())
        log_event("set naming standard", example.strip().upper())
        bump_data_version()
        st.rerun()

    st.divider()
    st.markdown("#### Display")
    options = [25, 50, 100, 200]
    current = rows_per_page()
    chosen = st.selectbox("Files per page (List and Table views)", options=options,
                          index=options.index(current) if current in options else 1, key="rows_setting")
    if chosen != current:
        set_setting("rows_per_page", str(chosen))
        st.rerun()

    st.divider()
    st.markdown("#### Data")
    st.caption(f"Database: `{DB_FILE}`")
    st.caption(f"Preview cache: `{CACHE_DIR}`")
    d1, d2, d3 = st.columns(3)
    with d1:
        if st.button("🧹 Clear preview cache", key="cache_clear", **WIDE_BTN):
            shutil.rmtree(CACHE_DIR, ignore_errors=True)
            page_image.clear()
            st.success("Preview cache cleared.")
    with d2.popover("🗑️ Reset file index"):
        st.write("Removes the list of indexed files. Workflow data, transmittals and the activity log are kept. "
                 "Your project files are never touched.")
        if st.button("Reset index", key="reset_index", type="primary"):
            clear_db(False)
            log_event("reset index")
            st.session_state.pop("last_scan", None)
            st.session_state["browse_path"] = ""
            bump_data_version()
            st.rerun()
    with d3.popover("⛔ Delete everything"):
        st.write("Deletes the index **and** all workflow data, transmittals and the activity log.")
        if st.button("Delete all data", key="reset_all", type="primary"):
            clear_db(True)
            st.session_state.pop("last_scan", None)
            st.session_state["browse_path"] = ""
            st.session_state["basket"] = []
            bump_data_version()
            st.rerun()

    st.divider()
    st.caption(f"{APP_TITLE} v{APP_VERSION} · runs locally on this computer · signed in as {current_user()}")


def main() -> None:
    st.set_page_config(page_title=APP_TITLE, page_icon="🏗️", layout="wide", initial_sidebar_state="expanded")
    init_db()
    apply_theme()
    inject_style(APP_CSS)

    ss = st.session_state
    ss.setdefault("data_version", 0)
    ss.setdefault("page", "Files")
    ss.setdefault("selected_id", None)
    ss.setdefault("basket", [])
    ss.setdefault("browse_path", "")
    ss.setdefault("tree_open", set())
    ss.setdefault("view_mode", "List")
    ss.setdefault("global_search", "")

    df = load_documents(ss["data_version"], get_setting("naming_example", ""))
    render_sidebar(df)
    page = ss["page"]

    if df.empty and page not in ("Settings", "Activity", "Transmittals"):
        st.title("🏗️ " + APP_TITLE)
        st.markdown("Index your project folder to browse drawings by area and discipline, preview them, "
                    "track revisions and issue transmittals.")
        with st.container(border=True):
            render_scanner("welcome")
        return

    top1, top2, _ = st.columns([6, 1.5, 0.35], vertical_alignment="center")
    query = top1.text_input("Search", key="global_search", label_visibility="collapsed",
                            placeholder="🔎  Search all drawings: number, title, folder, revision, notes ...")
    basket = len(ss.get("basket", []))
    top2.button(f"📨 Basket · {basket}", key="top_basket", on_click=go_to, args=("Transmittals",),
                help="Documents selected for the next transmittal", **WIDE_BTN)

    if query.strip() and not df.empty:
        page_search(df, query.strip())
    elif page == "Files":
        page_files(df)
    elif page == "Register":
        page_register(df)
    elif page == "Transmittals":
        page_transmittals(df)
    elif page == "Issues":
        page_issues(df)
    elif page == "Dashboard":
        page_dashboard(df)
    elif page == "Activity":
        page_activity()
    else:
        page_settings(df)


if __name__ == "__main__":
    main()
