#!/usr/bin/env python3
"""Monthly run for the BH (warranty-rate) dashboard.

What this script does (run by GitHub Actions once a month):
  1. Reads monthly sales (Ban) + warranty (BH) .xlsx files from Google Drive
     (folder "Ti le BH DVKT - 2026"), or from a local directory with
     --local-dir (same layout, for testing).
  2. Merges group assignments from the review Google Sheet ("Bang tra" tab
     and assigned rows in "Ma moi cho gan").
  3. Fills blank col A (product group) in sales files from the product-code
     mapping, and writes the filled files back (Drive upload, or copies
     under <local-dir>/_filled/ in local mode).
  4. New product codes (no group in the mapping) are appended to the
     "Ma moi cho gan" tab and the run STOPS WITHOUT writing
     dashboard-data.json (locked user rule: the user assigns groups and
     re-runs; the dashboard must never silently drop unmapped rows).
  5. When there are no new codes, aggregates and writes dashboard-data.json.

Locked rules honored:
  - 7 groups: MT, Apple, DTDD, Samsung, Smartwatch, MTB, May In.
  - Sales counted by ROWS; returned rows still count as sold.
  - BH is chung (total BH, no FRT/internal split in the display).
  - Brand tables count ALL rows incl. group Loai; group tables exclude Loai.
  - Warranty month from col D "Ngay nhan (80)", never col E "Thang".
  - Sales date col B: Excel serial, datetime, or apostrophe text date.
  - Unreadable dates / missing codes are counted separately, never guessed.
  - New codes are NEVER auto-guessed a group.
  - A month enters the dashboard ONLY when BOTH its Ban and BH files exist
    (keeps Sep 2025 sales-only data out of the tables).
  - Brand names are stripped of leading/trailing whitespace before any
    aggregation ("EDRA " -> "EDRA").
  - Product codes are normalized (strip one leading apostrophe, strip one
    leading "00") before any mapping lookup.

Env vars:
  GOOGLE_SERVICE_ACCOUNT_JSON  service-account JSON string (Drive + Sheets)
  DRIVE_FOLDER_ID              Drive folder id (empty = built-in default)
  REVIEW_SHEET_ID              review Google Sheet id (empty = skip review steps)

An empty or whitespace-only DRIVE_FOLDER_ID / REVIEW_SHEET_ID counts as
UNSET (GitHub sends "" for an unset repo variable).
"""
from __future__ import annotations

import argparse
import csv
import io
import json
import os
import re
import sys
import tempfile
from collections import Counter, defaultdict
from datetime import date, datetime, timedelta
from pathlib import Path

import openpyxl

GROUPS_7 = ["MT", "Apple", "ĐTDĐ", "Samsung", "Smartwatch", "MTB", "Máy In"]
GROUP_SET = set(GROUPS_7)
VALID_MAPPING_GROUPS = set(GROUPS_7) | {"Loại"}

DEFAULT_DRIVE_FOLDER_ID = "12DZVAw45iAMEdDC4ks-7bDKNbCBYatz1"
REVIEW_TAB_LOOKUP = "Bảng tra"
REVIEW_TAB_NEW = "Mã mới chờ gán"
MAPPING_HEADER = ["code", "group", "brand", "name", "source"]

SALES_FILE_RE = re.compile(r"^Ban-(\d{4})-(\d{2})\.xlsx$")
WARRANTY_FILE_RE = re.compile(r"^BH-(\d{4})-(\d{2})\.xlsx$")

# Sales subfolder name: "Bán" on Drive, "Ban" in the local split copy.
# Both are accepted everywhere; the name actually found is reused for
# downloads and for the local _filled/ copies.
SALES_SUBDIRS = ("Bán", "Ban")
WARRANTY_SUBDIRS = ("BH",)

EXCEL_EPOCH = datetime(1899, 12, 30)
XLSX_MIMETYPE = ("application/vnd.openxmlformats-officedocument."
                 "spreadsheetml.sheet")
DRIVE_FOLDER_MIMETYPE = "application/vnd.google-apps.folder"


def env_or_none(name):
    """Env value, with empty/whitespace-only treated as UNSET.

    GitHub sends an empty string for an unset repo variable, so "" must
    fall back to the default / skip, never be used as a real id.
    """
    v = os.environ.get(name)
    if v is None:
        return None
    v = v.strip()
    return v if v else None


# ---- locked rules: groups, dates, codes (same as build_dashboard_agg.py) ----

def sales_group(raw_group, brand):
    g = raw_group
    if g in ("Máy tính", "MT"):
        return "MT"
    if g in (0.0, "0.0", None, "") and brand == "HP":
        return "MT"
    if g == "Smartưatch":
        return "Smartwatch"
    if g in ("Máy In", "Máy in"):
        return "Máy In"
    if g in ("ÐTDÐ", "ĐTDĐ"):
        return "ĐTDĐ"
    if g in ("Apple", "Samsung", "Smartwatch", "MTB"):
        return g
    return None  # Loai, 0.0 khac HP, trong...


def warranty_group(raw):
    if raw == "Máy tính":
        return "MT"
    if raw in ("ÐTDÐ", "ĐTDĐ"):
        return "ĐTDĐ"
    return raw  # Apple, Samsung, Smartwatch, MTB, Máy In


def month_from_excel_serial(v):
    if not isinstance(v, (int, float)):
        return "KHONG-RO"
    try:
        d = datetime(1899, 12, 30) + timedelta(days=int(v))
    except Exception:
        return "KHONG-RO"
    return f"M{d.month:02d}.{str(d.year)[2:]}"


def parse_vn_date(s):
    """'01/10/2025 07:54 AM' hoac '28/09/2025 12:00 AM' -> date, None neu khong doc duoc."""
    if not s or not isinstance(s, str):
        return None
    s = s.strip().lstrip("'").strip()
    for fmt in ("%d/%m/%Y %I:%M %p", "%d/%m/%Y %H:%M", "%d/%m/%Y", "%Y-%m-%d"):
        try:
            return datetime.strptime(s, fmt).date()
        except ValueError:
            continue
    return None


def month_from_sales_date(v):
    """Ngay ban cot B: so serial Excel, datetime, hoac chu 'DD/MM/YYYY...' co dau '."""
    if isinstance(v, (int, float)):
        return month_from_excel_serial(v)
    if isinstance(v, datetime):
        return f"M{v.month:02d}.{str(v.year)[2:]}"
    if isinstance(v, date):
        return f"M{v.month:02d}.{str(v.year)[2:]}"
    d = parse_vn_date(v)
    if d is None:
        return "KHONG-RO"
    return f"M{d.month:02d}.{str(d.year)[2:]}"


def parse_any_date(v):
    """Warranty dates: normally text 'DD/MM/YYYY ...', but also accept
    datetime/date/Excel-serial values in case a file stores them so."""
    if isinstance(v, datetime):
        return v.date()
    if isinstance(v, date):
        return v
    if isinstance(v, (int, float)):
        try:
            return (EXCEL_EPOCH + timedelta(days=int(v))).date()
        except (OverflowError, ValueError):
            return None
    return parse_vn_date(v)


def normalize_code(v):
    """Normalize a product code BEFORE any mapping lookup, identically
    everywhere: whole-number floats -> plain integers (never scientific
    notation), strip one leading apostrophe, strip exactly one leading 00."""
    if v is None:
        return None
    if isinstance(v, float):
        s = str(int(v)) if v.is_integer() else str(v)
    else:
        s = str(v)
    s = s.strip()
    if s.startswith("'"):
        s = s[1:]
    s = s.strip()
    if s.startswith("00"):
        s = s[2:]
    return s if s else None


def clean_brand(v):
    """Brand names are stripped before any aggregation ("EDRA " -> "EDRA",
    "Goly " -> "Goly", "Cyber " -> "Cyber"), so no duplicate brand rows."""
    if v is None:
        return None
    s = str(v).strip()
    return s if s else None


def month_key(year, month):
    return f"M{month:02d}.{str(year)[2:]}"


def month_sort_key(mk):
    # "M10.25" -> (2025, 10); assumes years 2000-2099.
    return (2000 + int(mk[4:6]), int(mk[1:3]))


def is_blank_group(v):
    """True only when col A is really empty. 0.0 / "0.0" are VALUES handled
    by sales_group (HP -> MT), not blanks."""
    if v is None:
        return True
    return isinstance(v, str) and v.strip() == ""


# ---- mapping.csv ----

def load_mapping(path):
    """Returns (mapping, ) where mapping is code -> {group, brand, name, source}."""
    mapping = {}
    if not path.exists():
        print(f"WARNING: mapping file not found: {path} (starting empty)")
        return mapping
    with open(path, newline="", encoding="utf-8-sig") as f:
        reader = csv.DictReader(f)
        for row in reader:
            code = normalize_code(row.get("code"))
            group = (row.get("group") or "").strip()
            if not code:
                continue
            if group not in VALID_MAPPING_GROUPS:
                print(f"WARNING: mapping row {code}: invalid group "
                      f"{group!r}, skipped")
                continue
            mapping[code] = {
                "group": group,
                "brand": (row.get("brand") or "").strip(),
                "name": (row.get("name") or "").strip(),
                "source": (row.get("source") or "").strip(),
            }
    return mapping


def save_mapping(path, mapping):
    path.parent.mkdir(parents=True, exist_ok=True)
    with open(path, "w", newline="", encoding="utf-8") as f:
        w = csv.writer(f)
        w.writerow(MAPPING_HEADER)
        for code in sorted(mapping):
            e = mapping[code]
            w.writerow([code, e["group"], e.get("brand", ""),
                        e.get("name", ""), e.get("source", "")])


# ---- file discovery ----

def discover_local(root):
    """Walk <root>/Bán|Ban/<year>/<month>/Ban-YYYY-MM.xlsx and the BH side.

    Returns (found, subdirs) where found is
    {"sales": {(y, m): Path}, "warranty": {(y, m): Path}} and subdirs is
    the subfolder name actually found per kind. Files that do not match
    the month-folder layout (e.g. the review file at the folder root)
    are never even visited.
    """
    root = Path(root)
    found = {"sales": {}, "warranty": {}}
    subdirs = {}
    for kind, candidates, pattern in (("sales", SALES_SUBDIRS, SALES_FILE_RE),
                                      ("warranty", WARRANTY_SUBDIRS,
                                       WARRANTY_FILE_RE)):
        for subdir in candidates:
            base = root / subdir
            if not base.is_dir():
                continue
            subdirs.setdefault(kind, subdir)
            for year_dir in sorted(base.iterdir()):
                if not year_dir.is_dir() or not year_dir.name.isdigit():
                    continue
                for month_dir in sorted(year_dir.iterdir()):
                    if not month_dir.is_dir() or not month_dir.name.isdigit():
                        continue
                    for f in sorted(month_dir.iterdir()):
                        m = pattern.match(f.name)
                        if not m or not f.is_file():
                            continue
                        y, mo = int(m.group(1)), int(m.group(2))
                        if (y, mo) != (int(year_dir.name), int(month_dir.name)):
                            print(f"WARNING: {f} name/path month mismatch, "
                                  f"skipped")
                            continue
                        if (y, mo) in found[kind]:
                            print(f"WARNING: duplicate month file for "
                                  f"{y}-{mo:02d}: {f} (keeping first)")
                            continue
                        found[kind][(y, mo)] = f
        if kind not in subdirs:
            print(f"WARNING: no {'/'.join(candidates)} subfolder under {root}")
    return found, subdirs


def drive_list_children(drive, folder_id):
    out, page = [], None
    while True:
        resp = drive.files().list(
            q=f"'{folder_id}' in parents and trashed = false",
            fields="nextPageToken, files(id,name,mimeType)",
            pageSize=1000, pageToken=page).execute()
        out.extend(resp.get("files", []))
        page = resp.get("nextPageToken")
        if not page:
            break
    return out


def drive_find_folder(drive, parent_id, name):
    for child in drive_list_children(drive, parent_id):
        if child["name"] == name and child["mimeType"] == DRIVE_FOLDER_MIMETYPE:
            return child["id"]
    return None


def discover_drive_ids(drive, folder_id):
    """Walk the Drive folder tree (no hardcoded file ids).

    Returns (found, subdirs) where found is
    {"sales": {(y, m): {"id", "name"}}, "warranty": {...}} and subdirs is
    the subfolder name actually found per kind.
    """
    found = {"sales": {}, "warranty": {}}
    subdirs = {}
    for kind, candidates, pattern in (("sales", SALES_SUBDIRS, SALES_FILE_RE),
                                      ("warranty", WARRANTY_SUBDIRS,
                                       WARRANTY_FILE_RE)):
        for subdir in candidates:
            sub_id = drive_find_folder(drive, folder_id, subdir)
            if not sub_id:
                continue
            subdirs.setdefault(kind, subdir)
            for year in drive_list_children(drive, sub_id):
                if year["mimeType"] != DRIVE_FOLDER_MIMETYPE or not year["name"].isdigit():
                    continue
                for month in drive_list_children(drive, year["id"]):
                    if month["mimeType"] != DRIVE_FOLDER_MIMETYPE or not month["name"].isdigit():
                        continue
                    for f in drive_list_children(drive, month["id"]):
                        m = pattern.match(f["name"])
                        if not m or f["mimeType"] == DRIVE_FOLDER_MIMETYPE:
                            continue
                        y, mo = int(m.group(1)), int(m.group(2))
                        if (y, mo) != (int(year["name"]), int(month["name"])):
                            print(f"WARNING: Drive file {f['name']} name/path "
                                  f"month mismatch, skipped")
                            continue
                        if (y, mo) in found[kind]:
                            print(f"WARNING: duplicate Drive month file for "
                                  f"{y}-{mo:02d}: {f['name']} (keeping first)")
                            continue
                        found[kind][(y, mo)] = {"id": f["id"], "name": f["name"]}
        if kind not in subdirs:
            print(f"WARNING: no {'/'.join(candidates)} folder under Drive "
                  f"folder {folder_id}")
    return found, subdirs


def download_drive_files(drive, file_ids, subdirs, tmp_root):
    """Download every discovered file into tmp_root, same layout as Drive."""
    from googleapiclient.http import MediaIoBaseDownload
    for kind in ("sales", "warranty"):
        subdir = subdirs.get(kind)
        if subdir is None:
            continue
        for (y, m), meta in sorted(file_ids[kind].items()):
            dest = tmp_root / subdir / str(y) / f"{m:02d}" / meta["name"]
            dest.parent.mkdir(parents=True, exist_ok=True)
            request = drive.files().get_media(fileId=meta["id"])
            with open(dest, "wb") as fh:
                downloader = MediaIoBaseDownload(fh, request)
                done = False
                while not done:
                    _, done = downloader.next_chunk()
            print(f"  downloaded {subdir}/{y}/{m:02d}/{meta['name']}")


def upload_drive_file(drive, file_id, local_path):
    from googleapiclient.http import MediaFileUpload
    media = MediaFileUpload(str(local_path), mimetype=XLSX_MIMETYPE,
                            resumable=True)
    drive.files().update(fileId=file_id, media_body=media).execute()


# ---- aggregation ----

class Stats:
    def __init__(self):
        self.sales_rows_total = 0
        self.warranty_rows_total = 0
        self.sales_no_date = 0
        self.warranty_no_date = 0
        self.sales_no_code = 0
        self.warranty_no_code = 0
        # Dated rows outside the included months: counted, never in tables.
        self.sales_excluded_rows = 0
        self.warranty_excluded_rows = 0
        self.unmapped_rows = 0
        self.col_a_filled = 0
        self.files_filled = 0
        self.sales_mg = defaultdict(Counter)   # month -> group -> sold
        self.sales_mb = defaultdict(Counter)   # month -> brand -> sold
        self.w_mg = defaultdict(lambda: defaultdict(Counter))
        self.w_mb = defaultdict(Counter)       # month -> brand -> bh
        self.brand_groups = defaultdict(set)   # brand -> set of groups
        self.new_codes = {}  # code -> {"name", "brand", "units"}


def process_sales_file(path, included, mapping, stats):
    """Stream one sales file (read_only). Returns [(sheet_row, group)] col-A
    fills for this file.

    Column indices (same as the original file): A group (0), B date (1),
    AE brand (30), AI code (34), AJ name (35). Return status col K is read
    by the reference logic but returns still count as sold, so it needs no
    action here.
    """
    fills = []
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    ws = wb[wb.sheetnames[0]]
    for idx, row in enumerate(ws.iter_rows(values_only=True)):
        if idx == 0:  # header
            continue
        g_raw = row[0] if len(row) > 0 else None
        ngay = row[1] if len(row) > 1 else None
        brand = clean_brand(row[30]) if len(row) > 30 else None
        code = normalize_code(row[34]) if len(row) > 34 else None
        name = row[35] if len(row) > 35 else None
        name = str(name).strip() if name is not None else ""
        stats.sales_rows_total += 1
        if code is None:
            stats.sales_no_code += 1
        # --- group: col A, else mapping lookup for blank col A ---
        group = None
        blank_a = is_blank_group(g_raw)
        if blank_a:
            if code is not None and code in mapping:
                mapped = mapping[code]["group"]
                fills.append((idx + 1, mapped))
                stats.col_a_filled += 1
                if mapped in GROUP_SET:
                    group = mapped
            # Unknown code -> handled below, only if the row enters the
            # dashboard (a row outside the tables cannot block it).
        else:
            group = sales_group(g_raw, brand)
        # --- month from col B; unreadable -> counted separately, skipped ---
        month = month_from_sales_date(ngay)
        if month == "KHONG-RO":
            stats.sales_no_date += 1
            continue
        if month not in included:
            stats.sales_excluded_rows += 1
            continue
        # --- new-code detection (dashboard rows only, never auto-guessed) ---
        if blank_a and code is not None and code not in mapping:
            stats.unmapped_rows += 1
            entry = stats.new_codes.get(code)
            if entry is None:
                stats.new_codes[code] = {"name": name, "brand": brand or "",
                                         "units": 1}
            else:
                entry["units"] += 1
        # --- aggregate: brand counts ALL rows, groups only the 7 ---
        if brand is not None:
            stats.sales_mb[month][brand] += 1
        if group is not None:
            stats.sales_mg[month][group] += 1
            if brand is not None:
                stats.brand_groups[brand].add(group)
    wb.close()
    return fills


def process_warranty_file(path, included, stats):
    """Stream one warranty file (read_only).

    Column indices: D Ngay nhan (3), H code (7), I name (8), J group (9),
    K brand (10), L San pham BH (11), M Loai BH (12), O Ngay mua (14),
    Q Huong xu ly (16). Month comes from col D, never col E.
    """
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    ws = wb[wb.sheetnames[0]]
    for idx, row in enumerate(ws.iter_rows(values_only=True)):
        if idx == 0:  # header
            continue
        ngay_nhan = parse_any_date(row[3]) if len(row) > 3 else None
        code = normalize_code(row[7]) if len(row) > 7 else None
        grp_raw = row[9] if len(row) > 9 else None
        brand = clean_brand(row[10]) if len(row) > 10 else None
        sp_bh = row[11] if len(row) > 11 else None
        huong = row[16] if len(row) > 16 else None
        ngay_mua = parse_any_date(row[14]) if len(row) > 14 else None
        stats.warranty_rows_total += 1
        if code is None:
            stats.warranty_no_code += 1
        if ngay_nhan is None:
            stats.warranty_no_date += 1
            continue
        month = month_key(ngay_nhan.year, ngay_nhan.month)
        if month not in included:
            stats.warranty_excluded_rows += 1
            continue
        group = warranty_group(grp_raw)
        if brand is not None:
            stats.w_mb[month][brand] += 1
        if group in GROUP_SET:
            c = stats.w_mg[month][group]
            c["total"] += 1
            if huong == "FRT đổi máy cho Khách":
                c["frtX"] += 1
            elif huong == "Hãng đổi máy cho Khách":
                c["brandX"] += 1
            if sp_bh == "Hỏng phụ kiện đi kèm":
                c["acc"] += 1
            if ngay_nhan is not None and ngay_mua is not None:
                c["earlyKnown"] += 1
                if (ngay_nhan - ngay_mua).days <= 30:
                    c["early"] += 1
            if brand is not None:
                stats.brand_groups[brand].add(group)
    wb.close()


def apply_fills(src_path, fills, dest_path):
    """Load a sales file (normal mode), set the col-A cells, save to dest."""
    wb = openpyxl.load_workbook(src_path)
    ws = wb[wb.sheetnames[0]]
    for sheet_row, group in fills:
        ws.cell(row=sheet_row, column=1, value=group)
    dest_path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(dest_path)
    wb.close()


def build_output(stats, months_sorted):
    data = []
    for m in months_sorted:
        for g in GROUPS_7:
            c = stats.w_mg[m][g]
            data.append({
                "m": m, "g": g,
                "sold": stats.sales_mg[m][g],
                "bh": c["total"],
                "frtX": c["frtX"],
                "brandX": c["brandX"],
                "acc": c["acc"],
                "early": c["early"],
                "earlyKnown": c["earlyKnown"],
            })
    brand_month = {}
    for m in months_sorted:
        brands = set(stats.sales_mb[m]) | set(stats.w_mb[m])
        brand_month[m] = [[b, stats.sales_mb[m][b], stats.w_mb[m][b]]
                          for b in sorted(brands)]
    brand_groups = {}
    for b in sorted(stats.brand_groups):
        groups = [g for g in GROUPS_7 if g in stats.brand_groups[b]]
        if groups:  # brands with no group in either file are omitted
            brand_groups[b] = groups
    return {
        "meta": {
            "generated": datetime.now().isoformat(timespec="seconds"),
            "months": months_sorted,
            "sales_rows_total": stats.sales_rows_total,
            "warranty_total": stats.warranty_rows_total,
        },
        "data": data,
        "brand_groups": brand_groups,
        "brand_month": brand_month,
    }


# ---- Google services + review sheet ----

def build_services(sa_json_str):
    """Drive v3 + Sheets v4 from a service-account JSON string.

    Imported lazily so --local-dir / --self-test work without google libs.
    """
    from google.oauth2.service_account import Credentials
    from googleapiclient.discovery import build
    info = json.loads(sa_json_str)
    creds = Credentials.from_service_account_info(
        info, scopes=["https://www.googleapis.com/auth/drive",
                      "https://www.googleapis.com/auth/spreadsheets"])
    return (build("drive", "v3", credentials=creds),
            build("sheets", "v4", credentials=creds))


def read_tab_values(sheets, sheet_id, tab, last_col):
    resp = sheets.spreadsheets().values().get(
        spreadsheetId=sheet_id, range=f"'{tab}'!A:{last_col}").execute()
    return resp.get("values", [])


def write_tab_values(sheets, sheet_id, tab, last_col, rows):
    rng = f"'{tab}'!A:{last_col}"
    sheets.spreadsheets().values().clear(
        spreadsheetId=sheet_id, range=rng, body={}).execute()
    if rows:
        sheets.spreadsheets().values().update(
            spreadsheetId=sheet_id, range=rng, valueInputOption="RAW",
            body={"values": rows}).execute()


def merge_lookup_tab(sheets, sheet_id, mapping):
    """(a) Read "Bang tra" -> merge into mapping (sheet wins on conflict).
    Returns True if the mapping changed."""
    rows = read_tab_values(sheets, sheet_id, REVIEW_TAB_LOOKUP, "D")
    changed = False
    for r in rows[1:]:  # skip header
        if not r:
            continue
        code = normalize_code(r[0]) if len(r) > 0 else None
        group = r[1].strip() if len(r) > 1 else ""
        if not code:
            continue
        if group not in VALID_MAPPING_GROUPS:
            print(f"WARNING: '{REVIEW_TAB_LOOKUP}' row {code}: invalid group "
                  f"{group!r}, skipped")
            continue
        brand = r[2].strip() if len(r) > 2 else ""
        name = r[3].strip() if len(r) > 3 else ""
        old = mapping.get(code)
        if (old is None or old["group"] != group
                or old.get("brand") != brand or old.get("name") != name):
            mapping[code] = {"group": group, "brand": brand, "name": name,
                             "source": "sheet"}
            changed = True
    print(f"Review sheet: '{REVIEW_TAB_LOOKUP}' merged ({len(rows) - 1} rows)")
    return changed


def process_new_codes_tab(sheets, sheet_id, mapping):
    """(b) Read "Ma moi cho gan": rows with a filled, valid Nhom are added
    to the mapping and REMOVED from the tab. Returns (changed, codes left
    in the tab)."""
    rows = read_tab_values(sheets, sheet_id, REVIEW_TAB_NEW, "E")
    if not rows:
        return False, set()
    header = rows[0]
    kept = [header]
    remaining = set()
    changed = False
    removed = 0
    for r in rows[1:]:
        if not r or all(not (c or "").strip() for c in r):
            continue
        code = normalize_code(r[0]) if len(r) > 0 else None
        group = r[4].strip() if len(r) > 4 else ""
        if code and group:
            if group not in VALID_MAPPING_GROUPS:
                print(f"WARNING: '{REVIEW_TAB_NEW}' row {code}: invalid "
                      f"group {group!r}, left in tab")
                kept.append(r)
                remaining.add(code)
                continue
            name = r[1].strip() if len(r) > 1 else ""
            brand = r[2].strip() if len(r) > 2 else ""
            mapping[code] = {"group": group, "brand": brand, "name": name,
                             "source": "review"}
            changed = True
            removed += 1
        else:
            kept.append(r)
            if code:
                remaining.add(code)
    if removed:
        write_tab_values(sheets, sheet_id, REVIEW_TAB_NEW, "E", kept)
        print(f"Review sheet: {removed} assigned code(s) moved from "
              f"'{REVIEW_TAB_NEW}' into the mapping")
    return changed, remaining


def append_new_codes(sheets, sheet_id, new_codes, mapping, already_in_tab):
    """(c) Append newly found codes to "Ma moi cho gan" (code, name, brand,
    units sold). Skips codes already in the tab or already in the mapping.
    Returns the number appended."""
    rows = []
    for code in sorted(new_codes):
        if code in mapping or code in already_in_tab:
            continue
        e = new_codes[code]
        rows.append([code, e["name"], e["brand"], e["units"], ""])
    if not rows:
        return 0
    sheets.spreadsheets().values().append(
        spreadsheetId=sheet_id, range=f"'{REVIEW_TAB_NEW}'!A:E",
        valueInputOption="RAW", insertDataOption="INSERT_ROWS",
        body={"values": rows}).execute()
    return len(rows)


def rewrite_lookup_tab(sheets, sheet_id, mapping):
    """(d) Rewrite "Bang tra" fully from the merged mapping."""
    rows = [["Mã sản phẩm", "Nhóm", "Hãng", "Tên sản phẩm"]]
    for code in sorted(mapping):
        e = mapping[code]
        rows.append([code, e["group"], e.get("brand", ""), e.get("name", "")])
    write_tab_values(sheets, sheet_id, REVIEW_TAB_LOOKUP, "D", rows)
    print(f"Review sheet: '{REVIEW_TAB_LOOKUP}' rewritten "
          f"({len(mapping)} codes)")


# ---- self-test ----

def run_self_test():
    results = []

    def check(name, actual, expected):
        ok = actual == expected
        results.append(ok)
        print(f"  {'PASS' if ok else 'FAIL'}: {name}: "
              f"got {actual!r}, expected {expected!r}")

    print("Self-test: date parsing + code normalization + group merges")
    # Serial 45903 = 2025-09-03 by the Excel epoch (45931 = 2025-10-01),
    # so it belongs to M09.25. (An earlier brief draft wrongly said M10.25.)
    check("month_from_sales_date(45903) = 2025-09-03",
          month_from_sales_date(45903), "M09.25")
    check("month_from_sales_date(45931) = 2025-10-01",
          month_from_sales_date(45931), "M10.25")
    check("month_from_sales_date(45954) = 2025-10-24 (real Oct file value)",
          month_from_sales_date(45954), "M10.25")
    check("month_from_sales_date(\"'28/09/2025 12:00 AM\")",
          month_from_sales_date("'28/09/2025 12:00 AM"), "M09.25")
    check("month_from_sales_date(datetime(2026,1,15,10,30))",
          month_from_sales_date(datetime(2026, 1, 15, 10, 30)), "M01.26")
    check("month_from_sales_date(None)",
          month_from_sales_date(None), "KHONG-RO")
    check("month_from_sales_date('')",
          month_from_sales_date(""), "KHONG-RO")
    check("month_from_sales_date('garbage')",
          month_from_sales_date("garbage"), "KHONG-RO")
    check("parse_vn_date(\"'28/09/2025 12:00 AM\")",
          parse_vn_date("'28/09/2025 12:00 AM"), date(2025, 9, 28))
    check("parse_vn_date(None)", parse_vn_date(None), None)
    check("parse_vn_date('xx')", parse_vn_date("xx"), None)
    check("normalize_code(921800.0)", normalize_code(921800.0), "921800")
    check("normalize_code('001KG75B6HG00757')",
          normalize_code("001KG75B6HG00757"), "1KG75B6HG00757")
    check("normalize_code(\"'00123\")", normalize_code("'00123"), "123")
    check("normalize_code(None)", normalize_code(None), None)
    check("normalize_code('00') -> None (empty after strip)",
          normalize_code("00"), None)
    check("sales_group('Máy tính', 'X')", sales_group("Máy tính", "X"), "MT")
    check("sales_group(0.0, 'HP')", sales_group(0.0, "HP"), "MT")
    check("sales_group(0.0, 'Dell')", sales_group(0.0, "Dell"), None)
    check("sales_group('Smartưatch', 'X')",
          sales_group("Smartưatch", "X"), "Smartwatch")
    check("sales_group('Loại', 'X')", sales_group("Loại", "X"), None)
    check("warranty_group('Máy tính')", warranty_group("Máy tính"), "MT")
    check("clean_brand('EDRA ')", clean_brand("EDRA "), "EDRA")
    check("clean_brand('Cyber ')", clean_brand("Cyber "), "Cyber")
    check("clean_brand(None)", clean_brand(None), None)
    n_pass = sum(results)
    print(f"Self-test: {n_pass}/{len(results)} passed")
    return 0 if n_pass == len(results) else 1


# ---- main ----

def main(argv=None):
    ap = argparse.ArgumentParser(
        description="Monthly run for the BH dashboard: Drive/local monthly "
                    "files -> dashboard-data.json")
    ap.add_argument("--local-dir",
                    help="read monthly files from a local directory (same "
                         "Ban/<year>/<month>/ + BH/... layout) instead of "
                         "Drive; filled copies go to <local-dir>/_filled/")
    ap.add_argument("--mapping", help="mapping.csv path "
                    "(default: <repo>/data/mapping.csv)")
    ap.add_argument("--out", help="dashboard-data.json path "
                    "(default: <repo>/dashboard-data.json)")
    ap.add_argument("--drive-folder-id",
                    help="Drive folder id (overrides env DRIVE_FOLDER_ID)")
    ap.add_argument("--self-test", action="store_true",
                    help="run the date/code/group unit checks and exit")
    args = ap.parse_args(argv)
    if args.self_test:
        return run_self_test()

    # Repo root is the folder above scripts/ (deploy/): ../../data/mapping.csv
    # and ../../dashboard-data.json relative to this script.
    repo_root = Path(__file__).resolve().parents[1]
    mapping_path = Path(args.mapping) if args.mapping \
        else repo_root / "data" / "mapping.csv"
    out_path = Path(args.out) if args.out \
        else repo_root / "dashboard-data.json"

    # Empty/whitespace-only env values count as UNSET.
    drive_folder_id = ((args.drive_folder_id or "").strip()
                       or env_or_none("DRIVE_FOLDER_ID")
                       or DEFAULT_DRIVE_FOLDER_ID)
    review_sheet_id = env_or_none("REVIEW_SHEET_ID")
    sa_json = env_or_none("GOOGLE_SERVICE_ACCOUNT_JSON")
    local_mode = args.local_dir is not None

    mapping = load_mapping(mapping_path)
    print(f"Mapping loaded: {len(mapping)} codes from {mapping_path}")
    mapping_changed = False

    drive = sheets = None
    use_review = False
    if review_sheet_id:
        if not sa_json:
            print("WARNING: REVIEW_SHEET_ID is set but "
                  "GOOGLE_SERVICE_ACCOUNT_JSON is not; skipping review-sheet "
                  "steps")
        else:
            drive, sheets = build_services(sa_json)
            use_review = True
    else:
        print("WARNING: REVIEW_SHEET_ID is not set; skipping review-sheet "
              "steps")

    remaining_tab_codes = set()
    if use_review:
        if merge_lookup_tab(sheets, review_sheet_id, mapping):
            mapping_changed = True
        changed, remaining_tab_codes = process_new_codes_tab(
            sheets, review_sheet_id, mapping)
        mapping_changed = mapping_changed or changed

    tmpdir = None
    try:
        file_ids = {}
        if local_mode:
            local_root = Path(args.local_dir)
            found, _subdirs = discover_local(local_root)
            print(f"Local mode: {len(found['sales'])} sales files, "
                  f"{len(found['warranty'])} warranty files under {local_root}")
        else:
            if not sa_json:
                print("ERROR: GOOGLE_SERVICE_ACCOUNT_JSON is not set; cannot "
                      "read Drive. Use --local-dir for local testing.",
                      file=sys.stderr)
                return 2
            if drive is None:
                drive, sheets = build_services(sa_json)
            tmpdir = tempfile.TemporaryDirectory(prefix="bh-monthly-")
            tmp_root = Path(tmpdir.name)
            file_ids, drive_subdirs = discover_drive_ids(drive,
                                                         drive_folder_id)
            print(f"Drive mode: {len(file_ids['sales'])} sales files, "
                  f"{len(file_ids['warranty'])} warranty files in folder "
                  f"{drive_folder_id}")
            download_drive_files(drive, file_ids, drive_subdirs, tmp_root)
            found, _subdirs = discover_local(tmp_root)

        sales_months = {month_key(y, m) for (y, m) in found["sales"]}
        warranty_months = {month_key(y, m) for (y, m) in found["warranty"]}
        included = sorted(sales_months & warranty_months, key=month_sort_key)
        included_set = set(included)
        print(f"Months with BOTH files ({len(included)}): "
              f"{', '.join(included)}")
        for m in sorted(sales_months - warranty_months, key=month_sort_key):
            print(f"  month {m} has sales but no warranty file -> rows "
                  f"counted in CHECK only, not in tables")
        for m in sorted(warranty_months - sales_months, key=month_sort_key):
            print(f"  month {m} has warranty but no sales file -> rows "
                  f"counted in CHECK only, not in tables")

        stats = Stats()
        for (y, m) in sorted(found["sales"]):
            path = found["sales"][(y, m)]
            fills = process_sales_file(path, included_set, mapping, stats)
            if fills:
                stats.files_filled += 1
                if local_mode:
                    # Mirror the original layout under _filled/; the
                    # originals are never touched.
                    dest = (Path(args.local_dir) / "_filled"
                            / path.relative_to(Path(args.local_dir)))
                    apply_fills(path, fills, dest)
                    print(f"  {path.name}: filled {len(fills)} col-A cells "
                          f"-> {dest}")
                else:
                    filled_tmp = Path(tmpdir.name) / f"filled-{path.name}"
                    apply_fills(path, fills, filled_tmp)
                    upload_drive_file(
                        drive, file_ids["sales"][(y, m)]["id"], filled_tmp)
                    print(f"  {path.name}: filled {len(fills)} col-A cells, "
                          f"uploaded back to Drive")
        for (y, m) in sorted(found["warranty"]):
            process_warranty_file(found["warranty"][(y, m)], included_set,
                                  stats)
    finally:
        if tmpdir is not None:
            tmpdir.cleanup()

    # --- review sheet: append new codes, rewrite lookup tab ---
    if use_review:
        n_appended = append_new_codes(sheets, review_sheet_id,
                                      stats.new_codes, mapping,
                                      remaining_tab_codes)
        if n_appended:
            print(f"Review sheet: appended {n_appended} new code(s) to "
                  f"'{REVIEW_TAB_NEW}'")
        rewrite_lookup_tab(sheets, review_sheet_id, mapping)

    if mapping_changed:
        save_mapping(mapping_path, mapping)
        print(f"Mapping saved: {len(mapping)} codes -> {mapping_path}")

    # --- CHECK line ---
    sales_7 = sum(sum(stats.sales_mg[m].values()) for m in included)
    warr_7 = sum(sum(stats.w_mg[m][g]["total"] for g in GROUPS_7)
                 for m in included)
    check = {
        "sales_rows_total": stats.sales_rows_total,
        "warranty_total": stats.warranty_rows_total,
        "months": included,
        "sales_7groups_total": sales_7,
        "warranty_7groups_total": warr_7,
        "no_code_rows": stats.sales_no_code,
        "no_code_rows_warranty": stats.warranty_no_code,
        "no_date_rows_sales": stats.sales_no_date,
        "no_date_rows_warranty": stats.warranty_no_date,
        "sales_excluded_rows": stats.sales_excluded_rows,
        "warranty_excluded_rows": stats.warranty_excluded_rows,
        "new_codes_found": len(stats.new_codes),
        "col_a_filled": stats.col_a_filled,
        "files_filled": stats.files_filled,
        "unmapped_rows": stats.unmapped_rows,
    }
    print("CHECK: " + json.dumps(check, ensure_ascii=False))

    # --- locked stop rule: new codes -> NO dashboard write ---
    if stats.new_codes:
        print(f"PENDING_REVIEW: {len(stats.new_codes)} new codes need group "
              f"assignment in the review sheet; dashboard-data.json NOT "
              f"updated")
        for code in sorted(stats.new_codes):
            e = stats.new_codes[code]
            print(f"  new code: {code} | {e['brand']} | {e['name']} | "
                  f"sold {e['units']}")
        return 0

    out = build_output(stats, included)
    out_path.parent.mkdir(parents=True, exist_ok=True)
    with open(out_path, "w", encoding="utf-8") as f:
        json.dump(out, f, ensure_ascii=False, indent=1)
    print(f"Saved: {out_path} ({len(out['data'])} month x group rows, "
          f"{len(out['brand_groups'])} brands with groups)")
    return 0


if __name__ == "__main__":
    sys.exit(main())
