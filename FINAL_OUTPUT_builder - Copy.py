#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
FINAL_OUTPUT_builder.py
=======================
Builds the monthly RCPA + Coverage output from:
  - KPI master sheet
  - Division-wise DCR (SlotMAX) files
  - COMEX AIL hierarchy file

RCPA Coverage definition
------------------------
  RCPA Coverage (%) = (Doctors with valid Rx / Total DR Visited) x 100

A doctor is "valid for Rx" when ALL required brands for that division have
at least one visit row where Rx/Month > 0  (strictly greater than zero).
Rx/Month = 0 means no prescription was written and does NOT count.

Division 42 special rule
------------------------
For Div-42, ALL required brands must appear with Rx > 0 in a SINGLE visit
row (i.e. one DCR entry must capture all four brands simultaneously).

CHANGELOG
---------
2026-04-09  v2.0 - Fixes applied after UAT findings:
  * BUG FIX: rx >= 0 -> rx > 0  (Rx=0 is NOT a valid prescription)
  * BUG FIX: Division 34 added to DCR_FILES & ALLOWED_DIVISIONS
  * BUG FIX: Same rx>0 fix applied inside row_has_all_required (Div-42)
  * CLEANUP: Removed duplicate print statement in DCR loading loop
  * Added TODO comment for Div-34 brand set - verify with business owner
"""

import pandas as pd
import numpy as np
import re

# ============================================================
# CONFIG  - update file names here each month
# ============================================================

KPI_FILE  = "KPI - Mar 2026.xlsx"
KPI_SHEET = "final_KPI_TBM"

# Division-wise DCR files
# Key   = Division code string (must match KPI Division column exactly)
# Value = DCR file name
DCR_FILES = {
    "28": "DCR_RAW_STANDARDIZED_4div_2026-03-01_2026-03-31_Div28.xlsx",
    "30": "DCR_RAW_STANDARDIZED_4div_2026-03-01_2026-03-31_Div30.xlsx",
    "34": "DCR_RAW_STANDARDIZED_4div_2026-03-01_2026-03-31_Div34.xlsx",
    "35": "DCR_RAW_STANDARDIZED_4div_2026-03-01_2026-03-31_Div35.xlsx",
    "42": "DCR_RAW_STANDARDIZED_4div_2026-03-01_2026-03-31_Div42.xlsx",
}

COMEX_FILE  = "Comex_AIL.xlsx"
COMEX_SHEET = "AIL"

OUTPUT_FILE = "FINAL_OUTPUT.xlsx"

# Divisions to include in the final output.
# Must stay in sync with DCR_FILES keys above.
ALLOWED_DIVISIONS = {"28", "30", "34", "35", "42"}

# Required brands per division for RCPA validity.
# A doctor is counted only when ALL listed brands have Rx > 0.
# TODO [Div-34]: Confirm the exact 4 required brands with the business owner.
#   Current assumption: DUPHASTON, ESTRABET, NOVELON, FEMOSTON
#   (top-4 by DCR frequency in the Mar-2026 file)
DIVISION_BRAND_MAP = {
    "28": ["CREON",         "HEPTRAL SAME", "VONEFI",     "ROWASA"],
    "30": ["UDILIV",        "COLOSPA",      "FLORACHAMP", "EZYBIXY"],
    "34": ["DUPHASTON",     "ESTRABET",     "NOVELON",    "FEMOSTON"],  # <- VERIFY with business owner
    "35": ["CREMAFFIN PLUS","DIGERAFT PLUS","CREMAFFIN",  "LIBRAX"],
    "42": ["GANATON",       "GANATON TOTAL","DUPHALAC",   "ELDICET"],
}

FINAL_COLS = [
    "Division",
    "ZBM CODE", "ZBM NAME",
    "ABM CODE", "ABM NAME",
    "Employee Code", "Full Name",
    "Territory Headquarter", "Abbott Designation",
    "DOJ", "Territory",
    "Last Submitted DCR Date",
    "Status",
    "Total Dr Total", "Total DR Visited",
    "Total Coverage",
    "Number of doctors with Rx entered",
    "RCPA Coverage",
]

# ============================================================
# HELPERS
# ============================================================

def clean_cols(df):
    df = df.copy()
    df.columns = df.columns.astype(str).str.strip()
    return df

def norm_code(x):
    """Normalise employee / hierarchy codes: upper-case, strip whitespace."""
    if pd.isna(x):
        return ""
    return re.sub(r"\s+", "", str(x).upper().strip())

def norm_brand(x):
    """Normalise brand names: keep only alphanumeric chars, upper-case."""
    return re.sub(r"[^A-Z0-9]", "", str(x).upper())

_ZERO_WIDTH = "\u200b\u200c\u200d\ufeff"

def norm_doc_id(x):
    """Normalise doctor identifiers (Assignment / Account codes)."""
    s = "" if pd.isna(x) else str(x)
    s = s.translate({ord(ch): None for ch in _ZERO_WIDTH})
    s = s.replace("\xa0", " ")
    return re.sub(r"\s+", "", s.strip().upper())

def compute_total_coverage(df):
    visited = pd.to_numeric(df["Total DR Visited"], errors="coerce").fillna(0)
    total   = pd.to_numeric(
        df.get("Total DR Total", df.get("Total Dr Total", 0)),
        errors="coerce"
    ).fillna(0)
    return np.where(total > 0, (visited / total) * 100, 0).round(2)

# ============================================================
# 1. READ KPI (MASTER)
# ============================================================

print("[INFO] Reading KPI master...")
kpi = clean_cols(pd.read_excel(KPI_FILE, sheet_name=KPI_SHEET, engine="openpyxl"))

kpi["Division"] = kpi["Division"].astype(str).str.strip()
kpi = kpi[kpi["Division"].isin(ALLOWED_DIVISIONS)].copy()

kpi["Employee Code"] = kpi["Employee Code"].map(norm_code)
kpi["_emp_key"]      = kpi["Employee Code"]

kpi["Total DR Visited"] = pd.to_numeric(kpi["Total DR Visited"], errors="coerce").fillna(0)
kpi["Total DR Total"]   = pd.to_numeric(
    kpi.get("Total DR Total", kpi.get("Total Dr Total", 0)),
    errors="coerce"
).fillna(0)

kpi["Total Coverage"] = compute_total_coverage(kpi)

print(f"  KPI rows loaded  : {len(kpi)}")
print(f"  Divisions present: {sorted(kpi['Division'].unique())}")

# ============================================================
# 2. READ COMEX (HIERARCHY)
# ============================================================

print("[INFO] Reading COMEX AIL for hierarchy...")
comex = clean_cols(pd.read_excel(COMEX_FILE, sheet_name=COMEX_SHEET, engine="openpyxl"))

for col in ["DIVISION", "EMPLOYEE_CODE", "EHIER_CD", "PAR_EHIER_CD",
            "EMPLOYEE_NAME", "PAR_EMPLOYEE_NAME"]:
    if col not in comex.columns:
        comex[col] = ""

comex["DIVISION"]      = comex["DIVISION"].astype(str).str.strip()
comex["EMPLOYEE_CODE"] = comex["EMPLOYEE_CODE"].map(norm_code)
comex["EHIER_CD"]      = comex["EHIER_CD"].map(norm_code)
comex["PAR_EHIER_CD"]  = comex["PAR_EHIER_CD"].map(norm_code)

comex = comex[comex["DIVISION"].isin(ALLOWED_DIVISIONS)].copy()

node_by_ehier = (
    comex.drop_duplicates(subset=["DIVISION", "EHIER_CD"])
    .set_index("EHIER_CD")[["DIVISION", "EMPLOYEE_CODE", "EMPLOYEE_NAME", "PAR_EHIER_CD"]]
    .to_dict(orient="index")
)

emp_to_ehier = {}
for _, r in comex.iterrows():
    if r["EMPLOYEE_CODE"]:
        emp_to_ehier[(r["DIVISION"], r["EMPLOYEE_CODE"])] = r["EHIER_CD"]

def climb_to(prefix, start):
    """Walk up the COMEX hierarchy until we reach a node whose EHIER_CD
    starts with *prefix* (e.g. 'IA' for ABM, 'RG' for ZBM)."""
    seen, cur = set(), start
    while cur and cur in node_by_ehier and cur not in seen:
        seen.add(cur)
        row = node_by_ehier[cur]
        if cur.startswith(prefix):
            return cur, row.get("EMPLOYEE_NAME", "")
        cur = row.get("PAR_EHIER_CD", "")
    return "", ""

hier_rows = []
for _, r in kpi[["Division", "_emp_key"]].drop_duplicates().iterrows():
    div, emp = r["Division"], r["_emp_key"]
    start = emp_to_ehier.get((div, emp), "")
    abm_c, abm_n = climb_to("IA", start) if start else ("", "")
    zbm_c, zbm_n = climb_to("RG", start) if start else ("", "")
    hier_rows.append({
        "Division": div, "_emp_key": emp,
        "ABM CODE": abm_c, "ABM NAME": abm_n,
        "ZBM CODE": zbm_c, "ZBM NAME": zbm_n,
    })

hierarchy = pd.DataFrame(hier_rows)

# ============================================================
# 3. READ DCR (MULTIPLE FILES)
# ============================================================

print("[INFO] Reading DCR SlotMAX (division-wise files)...")

dcr_frames = []
for div, fname in DCR_FILES.items():
    print(f"  - Division {div}: {fname}")
    df = clean_cols(pd.read_excel(fname, engine="openpyxl"))
    df["Division"] = div          # enforce consistent string key
    dcr_frames.append(df)

dcr = pd.concat(dcr_frames, ignore_index=True)

dcr["_emp_key"] = dcr["User: Alias"].map(norm_code)

brand_cols = [c for c in dcr.columns if c.startswith("Brand")]
rx_cols    = [c for c in dcr.columns if c.startswith("Rx/Month")]

dcr[brand_cols] = dcr[brand_cols].fillna("").astype(str)
dcr[rx_cols]    = dcr[rx_cols].apply(pd.to_numeric, errors="coerce")

dcr["Assignment"] = dcr["Assignment"].apply(norm_doc_id)

_CUST_COL = next(
    (c for c in ["Account: Customer Code", "Customer Code", "Account"]
     if c in dcr.columns),
    None,
)
dcr["_acc_code"] = dcr[_CUST_COL].map(norm_code) if _CUST_COL else ""

_ID_COL = "Account ID_18" if "Account ID_18" in dcr.columns else None
dcr["_acc_id"] = dcr[_ID_COL].map(norm_code) if _ID_COL else ""

# Doctor key priority: Account ID_18 > Customer Code > Assignment name
dcr["_doctor_key"] = np.where(
    dcr["_acc_id"]   != "", dcr["_acc_id"],
    np.where(dcr["_acc_code"] != "", dcr["_acc_code"], dcr["Assignment"])
)

# ============================================================
# 4. COMPUTE VALID DOCTORS (RCPA LOGIC)
# ============================================================

def row_has_all_required(row, required_norm):
    """
    Div-42 only: check that a SINGLE visit row covers ALL required brands
    each with Rx/Month > 0.
    NOTE: uses rx > 0 (not >= 0) - Rx=0 is not a valid prescription.
    """
    for rb in required_norm:
        ok = False
        for b, r in zip(brand_cols, rx_cols):
            # FIX: > 0 (was >= 0)
            if norm_brand(row[b]) == rb and pd.notna(row[r]) and row[r] > 0:
                ok = True
                break
        if not ok:
            return False
    return True


records = []

for (div, emp, doc), grp in dcr.groupby(["Division", "_emp_key", "_doctor_key"]):

    # Build brand -> first-seen Rx mapping for this (employee, doctor) pair
    brand_rx = {}
    for _, row in grp.iterrows():
        for b, r in zip(brand_cols, rx_cols):
            br = norm_brand(row[b])
            if br and br not in brand_rx:
                brand_rx[br] = row[r]

    required = [norm_brand(x) for x in DIVISION_BRAND_MAP.get(div, [])]

    # FIX: rx > 0  (was rx >= 0 - Rx=0 does NOT represent a written prescription)
    valid = all(
        br in brand_rx and pd.notna(brand_rx[br]) and brand_rx[br] > 0
        for br in required
    )

    # Div-42 extra constraint: all brands must appear in one single visit row
    if valid and div == "42":
        valid = any(row_has_all_required(r, required) for _, r in grp.iterrows())

    if valid:
        records.append({"Division": div, "_emp_key": emp, "_doctor_key": doc})


valid_doctors = (
    pd.DataFrame(records).drop_duplicates()
    if records
    else pd.DataFrame(columns=["Division", "_emp_key", "_doctor_key"])
)

rx_cnt = (
    valid_doctors.groupby(["Division", "_emp_key"])["_doctor_key"]
    .nunique()
    .reset_index(name="Number of doctors with Rx entered")
)

# ============================================================
# 5. MERGE & FINAL OUTPUT
# ============================================================

out = kpi.merge(hierarchy, on=["Division", "_emp_key"], how="left")
out = out.merge(rx_cnt,    on=["Division", "_emp_key"], how="left")

out["Number of doctors with Rx entered"] = (
    out["Number of doctors with Rx entered"].fillna(0).astype(int)
)

# Employees with zero field visits cannot have Rx entries
out.loc[out["Total DR Visited"] == 0, "Number of doctors with Rx entered"] = 0

out["RCPA Coverage"] = np.clip(
    np.where(
        out["Total DR Visited"] > 0,
        (out["Number of doctors with Rx entered"] / out["Total DR Visited"]) * 100,
        0,
    ).round(2),
    0, 100,
)

final = out.copy()
final["Total Dr Total"] = final["Total DR Total"]

for c in FINAL_COLS:
    if c not in final.columns:
        final[c] = ""

final = final[FINAL_COLS]
final.to_excel(OUTPUT_FILE, index=False)

# ============================================================
# SUMMARY
# ============================================================

print("\n[OK] FINAL_OUTPUT.xlsx written successfully")
print(f"  Total rows       : {len(final)}")
for div in sorted(final["Division"].astype(str).unique()):
    sub = final[final["Division"].astype(str) == div]
    avg_rcpa = sub["RCPA Coverage"].mean()
    print(f"  Division {div:>3}  ->  {len(sub):4d} employees  |  avg RCPA {avg_rcpa:.1f}%")
