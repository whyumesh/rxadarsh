#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import pandas as pd
import numpy as np
import re

# ============================================================
# CONFIG
# ============================================================

KPI_FILE = "KPI - Mar 2026.xlsx"
KPI_SHEET = "final_KPI_TBM"

DCR_FILES = {
    "28": "DCR_RAW_STANDARDIZED_4div_2026-03-01_2026-03-31_Div28.xlsx",
    "30": "DCR_RAW_STANDARDIZED_4div_2026-03-01_2026-03-31_Div30.xlsx",
    "35": "DCR_RAW_STANDARDIZED_4div_2026-03-01_2026-03-31_Div35.xlsx",
    "42": "DCR_RAW_STANDARDIZED_4div_2026-03-01_2026-03-31_Div42.xlsx",
}

# DCR_SHEET = "Sheet1"
COMEX_FILE = "Comex_AIL.xlsx"
COMEX_SHEET = "AIL"
OUTPUT_FILE = "FINAL_OUTPUT.xlsx"

ALLOWED_DIVISIONS = {"28", "30", "35", "42"}

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
    "RCPA Coverage"
]

# ============================================================
# HELPERS
# ============================================================

def clean_cols(df):
    df = df.copy()
    df.columns = df.columns.astype(str).str.strip()
    return df

def norm_code(x):
    if pd.isna(x):
        return ""
    return re.sub(r"\s+", "", str(x).upper().strip())

def norm_brand(x):
    if pd.isna(x):
        return ""
    return re.sub(r"[^A-Z0-9]", "", str(x).upper())

ZERO_WIDTH_CHARS = "\u200b\u200c\u200d\ufeff"

def norm_doc_id(x):
    s = "" if pd.isna(x) else str(x)
    s = s.translate({ord(ch): None for ch in ZERO_WIDTH_CHARS})
    s = s.replace("\xa0", " ")
    s = re.sub(r"\s+", "", s)
    return s.upper()

def compute_total_coverage(df):
    visited = pd.to_numeric(df["Total DR Visited"], errors="coerce").fillna(0)
    total = pd.to_numeric(
        df.get("Total DR Total", df.get("Total Dr Total", 0)),
        errors="coerce"
    ).fillna(0)
    return np.where(total > 0, (visited / total) * 100, 0).round(2)

# ============================================================
# READ KPI
# ============================================================

print("[INFO] Reading KPI file...")
kpi = clean_cols(pd.read_excel(KPI_FILE, sheet_name=KPI_SHEET, engine="openpyxl"))

kpi["Division"] = kpi["Division"].astype(str).str.strip()
kpi = kpi[kpi["Division"].isin(ALLOWED_DIVISIONS)].copy()

kpi["Employee Code"] = kpi["Employee Code"].map(norm_code)
kpi["_emp_key"] = kpi["Employee Code"]

kpi["Total DR Visited"] = pd.to_numeric(kpi["Total DR Visited"], errors="coerce").fillna(0)
kpi["Total DR Total"] = pd.to_numeric(
    kpi.get("Total DR Total", kpi.get("Total Dr Total", 0)),
    errors="coerce"
).fillna(0)

kpi["Total Coverage"] = compute_total_coverage(kpi)

# ============================================================
# READ COMEX (Hierarchy)
# ============================================================

print("[INFO] Reading COMEX file...")
comex = clean_cols(pd.read_excel(COMEX_FILE, sheet_name=COMEX_SHEET, engine="openpyxl"))

for c in [
    "DIVISION", "EMPLOYEE_CODE",
    "EHIER_CD", "PAR_EHIER_CD",
    "EMPLOYEE_NAME", "PAR_EMPLOYEE_NAME"
]:
    if c not in comex.columns:
        comex[c] = ""

comex["DIVISION"] = comex["DIVISION"].astype(str).str.strip()
comex["EMPLOYEE_CODE"] = comex["EMPLOYEE_CODE"].map(norm_code)
comex["EHIER_CD"] = comex["EHIER_CD"].map(norm_code)
comex["PAR_EHIER_CD"] = comex["PAR_EHIER_CD"].map(norm_code)

comex = comex[comex["DIVISION"].isin(ALLOWED_DIVISIONS)].copy()

node_by_ehier = (
    comex.drop_duplicates(subset=["DIVISION", "EHIER_CD"])
         .set_index("EHIER_CD")[
             ["DIVISION", "EMPLOYEE_CODE", "EMPLOYEE_NAME", "PAR_EHIER_CD"]
         ]
         .to_dict(orient="index")
)

emp_to_ehier = {}
for _, r in comex.iterrows():
    if r["EMPLOYEE_CODE"]:
        emp_to_ehier[(r["DIVISION"], r["EMPLOYEE_CODE"])] = r["EHIER_CD"]

def climb_to(prefix, start):
    seen = set()
    cur = start
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
        "Division": div,
        "_emp_key": emp,
        "ABM CODE": abm_c,
        "ABM NAME": abm_n,
        "ZBM CODE": zbm_c,
        "ZBM NAME": zbm_n,
    })

hierarchy = pd.DataFrame(hier_rows)

# ============================================================
# READ DCR FILES
# ============================================================

print("[INFO] Reading DCR files...")
dcr_frames = []

for div, fname in DCR_FILES.items():
    print(f"  Loading Division {div}...")

    # Read Excel file and auto-pick first sheet
    xl = pd.ExcelFile(fname, engine="openpyxl")
    first_sheet = xl.sheet_names[0]

    df = clean_cols(pd.read_excel(xl, sheet_name=first_sheet))
    df["Division"] = div

    dcr_frames.append(df)

dcr = pd.concat(dcr_frames, ignore_index=True)

dcr["_emp_key"] = dcr["User: Alias"].map(norm_code)

brand_cols = [c for c in dcr.columns if c.startswith("Brand")]
rx_cols = [c for c in dcr.columns if c.startswith("Rx/Month")]

dcr[brand_cols] = dcr[brand_cols].fillna("").astype(str)
dcr[rx_cols] = dcr[rx_cols].apply(pd.to_numeric, errors="coerce")

dcr["Assignment"] = dcr["Assignment"].apply(norm_doc_id)

ACCOUNT_COL = next(
    (c for c in ["Account ID_18", "Account: Customer Code", "Customer Code"] if c in dcr.columns),
    None
)

dcr["_doctor_key"] = np.where(
    dcr["Assignment"].astype(str).str.strip() != "",
    dcr["Assignment"].astype(str).str.strip(),
    dcr[ACCOUNT_COL].map(norm_code) if ACCOUNT_COL else ""
)


# ============================================================
# ✅ CORRECT RX DOCTOR COUNT LOGIC (FIXED)
# ============================================================

records = []

for (div, emp, doc), grp in dcr.groupby(["Division", "_emp_key", "_doctor_key"]):
    has_valid_rx = False

    for _, row in grp.iterrows():
        for b, r in zip(brand_cols, rx_cols):
            if norm_brand(row[b]) and pd.notna(row[r]) and row[r] > 0:
                has_valid_rx = True
                break
        if has_valid_rx:
            break

    if has_valid_rx:
        records.append({
            "Division": div,
            "_emp_key": emp,
            "_doctor_key": doc
        })

valid_doctors = pd.DataFrame(records).drop_duplicates()

rx_cnt = (
    valid_doctors
    .groupby(["Division", "_emp_key"])["_doctor_key"]
    .nunique()
    .reset_index(name="Number of doctors with Rx entered")
)

# ============================================================
# FINAL MERGE & OUTPUT
# ============================================================

out = kpi.merge(hierarchy, on=["Division", "_emp_key"], how="left")
out = out.merge(rx_cnt, on=["Division", "_emp_key"], how="left")

out["Number of doctors with Rx entered"] = (
    out["Number of doctors with Rx entered"].fillna(0).astype(int)
)

out.loc[out["Total DR Visited"] == 0, "Number of doctors with Rx entered"] = 0

out["RCPA Coverage"] = np.clip(
    np.where(
        out["Total DR Visited"] > 0,
        (out["Number of doctors with Rx entered"] / out["Total DR Visited"]) * 100,
        0
    ).round(2),
    0, 100
)

final = out.copy()
final["Total Dr Total"] = final["Total DR Total"]

for c in FINAL_COLS:
    if c not in final.columns:
        final[c] = ""

final = final[FINAL_COLS]
final.to_excel(OUTPUT_FILE, index=False)

print("\n[OK] FINAL_OUTPUT.xlsx generated successfully")