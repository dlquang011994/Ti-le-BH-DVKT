"""Build product-code -> group mapping table for BH dashboard.

Seed from warranty file (standard), then add sales-only codes.
Writes deploy/data/mapping.csv with header: code,group,brand,name,source
"""
import csv
import json
import time
from collections import Counter, defaultdict

import openpyxl
from pyxlsb import open_workbook

SALES_PATH = "/home/hatch/workspace/yearly-report/data/Data-ban-2026.xlsb"
WARRANTY_PATH = "/home/hatch/workspace/user/files/Data_ba_o_ha_nh_2026__1.xlsx"
OUT_CSV = "/home/hatch/workspace/yearly-report/deploy/data/mapping.csv"
OUT_REPORT = "/home/hatch/workspace/yearly-report/deploy/data/mapping_report.json"
SHEETS = ["M10.25-M03.26", "M04.26-M09.26"]

GROUPS_7 = ["MT", "Apple", "ĐTDĐ", "Samsung", "Smartwatch", "MTB", "Máy In"]


def normalize_code(v):
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
    return None


def warranty_group(raw):
    if raw == "Máy tính":
        return "MT"
    if raw in ("ÐTDÐ", "ĐTDĐ"):
        return "ĐTDĐ"
    return raw


def sales_row_group(raw_group, brand):
    """Resolved mapping group for one sales row: 7-group, 'Loại', or None (no group)."""
    g = sales_group(raw_group, brand)
    if g is not None:
        return g
    if raw_group == "Loại":
        return "Loại"
    return None


def warranty_row_group(raw):
    """Resolved mapping group for one warranty row: 7-group, 'Loại', or None (no group)."""
    if raw in (None, ""):
        return None
    g = warranty_group(raw)
    if g in (None, ""):
        return None
    return g


def cellv(row, idx):
    if idx >= len(row) or row[idx] is None:
        return None
    return row[idx].v


def main():
    t0 = time.time()

    # ---------- WARRANTY (standard) ----------
    wb = openpyxl.load_workbook(WARRANTY_PATH, read_only=True, data_only=True)
    ws = wb[wb.sheetnames[0]]
    w_rows = 0
    w_empty_code_rows = 0
    w_codes = {}  # code -> {"groups": Counter, "brand": first, "name": first, "rows": n, "raw_groups": Counter}
    for r in ws.iter_rows(values_only=True):
        if w_rows == 0:  # header
            w_rows += 1
            continue
        w_rows += 1
        code = normalize_code(r[7])
        if code is None:
            w_empty_code_rows += 1
            continue
        grp = warranty_row_group(r[9])
        brand = r[10]
        name = r[8]
        ent = w_codes.get(code)
        if ent is None:
            ent = {"groups": Counter(), "brand": brand, "name": name,
                   "rows": 0, "raw_groups": Counter()}
            w_codes[code] = ent
        ent["rows"] += 1
        ent["raw_groups"][repr(r[9])] += 1
        if grp is not None:
            ent["groups"][grp] += 1
    wb.close()
    w_data_rows = w_rows - 1
    print(f"Warranty: {w_data_rows} data rows, {len(w_codes)} distinct codes, "
          f"{w_empty_code_rows} rows with empty code ({time.time()-t0:.0f}s)", flush=True)

    # Resolve warranty codes (majority if disagreeing)
    w_final = {}  # code -> group
    w_conflicts = []  # (code, {group: count})
    w_no_group_codes = []
    for code, ent in w_codes.items():
        if not ent["groups"]:
            w_no_group_codes.append(code)
            continue
        if len(ent["groups"]) > 1:
            w_conflicts.append((code, dict(ent["groups"])))
        w_final[code] = ent["groups"].most_common(1)[0][0]
    print(f"Warranty resolved: {len(w_final)} codes with group, "
          f"{len(w_no_group_codes)} codes with no group, "
          f"{len(w_conflicts)} codes with disagreeing groups", flush=True)

    # ---------- SALES ----------
    s_rows = 0
    s_empty_code_rows = 0
    s_codes = {}
    with open_workbook(SALES_PATH) as wb:
        for name in SHEETS:
            with wb.get_sheet(name) as sh:
                for i, row in enumerate(sh.rows()):
                    if i == 0:
                        continue
                    s_rows += 1
                    code = normalize_code(cellv(row, 34))
                    if code is None:
                        s_empty_code_rows += 1
                        continue
                    grp = sales_row_group(cellv(row, 0), cellv(row, 30))
                    brand = cellv(row, 30)
                    nm = cellv(row, 35)
                    ent = s_codes.get(code)
                    if ent is None:
                        ent = {"groups": Counter(), "brand": brand, "name": nm,
                               "rows": 0, "raw_groups": Counter()}
                        s_codes[code] = ent
                    ent["rows"] += 1
                    ent["raw_groups"][repr(cellv(row, 0))] += 1
                    if grp is not None:
                        ent["groups"][grp] += 1
            print(f"Sales sheet {name} done ({time.time()-t0:.0f}s)", flush=True)
    print(f"Sales: {s_rows} data rows, {len(s_codes)} distinct codes, "
          f"{s_empty_code_rows} rows with empty code ({time.time()-t0:.0f}s)", flush=True)

    s_final = {}
    s_conflicts = []
    s_no_group_codes = []
    for code, ent in s_codes.items():
        if not ent["groups"]:
            s_no_group_codes.append(code)
            continue
        if len(ent["groups"]) > 1:
            s_conflicts.append((code, dict(ent["groups"])))
        s_final[code] = ent["groups"].most_common(1)[0][0]
    print(f"Sales resolved: {len(s_final)} codes with group, "
          f"{len(s_no_group_codes)} codes with no group, "
          f"{len(s_conflicts)} codes with disagreeing groups", flush=True)

    # ---------- MERGE: warranty wins ----------
    both_codes = set(w_codes) & set(s_codes)
    both_with_group = [c for c in both_codes if c in w_final and c in s_final]
    ws_conflicts = []  # (code, warranty_group, sales_group)
    for c in both_with_group:
        if w_final[c] != s_final[c]:
            ws_conflicts.append((c, w_final[c], s_final[c]))

    mapping = {}  # code -> (group, brand, name, source)
    for code, grp in w_final.items():
        ent = w_codes[code]
        mapping[code] = (grp, ent["brand"], ent["name"], "BH")
    ban_added = 0
    for code, grp in s_final.items():
        if code in mapping:
            continue
        ent = s_codes[code]
        mapping[code] = (grp, ent["brand"], ent["name"], "Ban")
        ban_added += 1

    # ---------- WRITE CSV ----------
    import os
    os.makedirs(os.path.dirname(OUT_CSV), exist_ok=True)
    with open(OUT_CSV, "w", encoding="utf-8", newline="") as f:
        w = csv.writer(f)
        w.writerow(["code", "group", "brand", "name", "source"])
        for code in sorted(mapping):
            grp, brand, nm, src = mapping[code]
            w.writerow([code, grp, brand if brand is not None else "",
                        nm if nm is not None else "", src])

    group_counts = Counter(g for g, _, _, _ in mapping.values())
    report = {
        "warranty_rows": w_data_rows,
        "warranty_distinct_codes": len(w_codes),
        "warranty_empty_code_rows": w_empty_code_rows,
        "warranty_codes_with_group": len(w_final),
        "warranty_codes_no_group": len(w_no_group_codes),
        "warranty_codes_no_group_list": sorted(w_no_group_codes)[:50],
        "warranty_disagreeing_codes": len(w_conflicts),
        "warranty_disagreeing_examples": [[c, g] for c, g in sorted(w_conflicts)[:20]],
        "sales_rows": s_rows,
        "sales_distinct_codes": len(s_codes),
        "sales_empty_code_rows": s_empty_code_rows,
        "sales_codes_with_group": len(s_final),
        "sales_codes_no_group": len(s_no_group_codes),
        "sales_codes_no_group_list": sorted(s_no_group_codes)[:50],
        "sales_disagreeing_codes": len(s_conflicts),
        "sales_disagreeing_examples": [[c, g] for c, g in sorted(s_conflicts)[:20]],
        "codes_in_both_files": len(both_codes),
        "codes_in_both_with_group": len(both_with_group),
        "warranty_vs_sales_conflicts": len(ws_conflicts),
        "warranty_vs_sales_conflict_list": sorted(ws_conflicts)[:50],
        "mapping_total_rows": len(mapping),
        "mapping_from_BH": sum(1 for v in mapping.values() if v[3] == "BH"),
        "mapping_from_Ban": sum(1 for v in mapping.values() if v[3] == "Ban"),
        "mapping_group_counts": dict(group_counts),
        "elapsed_seconds": round(time.time() - t0, 1),
    }
    with open(OUT_REPORT, "w", encoding="utf-8") as f:
        json.dump(report, f, ensure_ascii=False, indent=1)

    print(json.dumps(report, ensure_ascii=False, indent=1), flush=True)
    print("Wrote:", OUT_CSV, flush=True)


if __name__ == "__main__":
    main()
