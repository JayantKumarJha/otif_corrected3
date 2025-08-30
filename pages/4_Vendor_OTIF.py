# Colab OTIF report script (auto-detect .xls/.xlsx, normalize columns, same logic as before)
# Run this cell in Google Colab. When prompted, upload your Excel file.

# Install helpful packages (quiet)
!pip install --quiet openpyxl xlrd reportlab

import io
import os
import re
import sys
from datetime import datetime
import pandas as pd
import numpy as np

# Reportlab imports for PDF
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
from reportlab.lib.units import mm

# ---------- CONFIG / MAPPINGS ----------
# Placeholder item-category mapping (extend with your real map)
df1 = pd.DataFrame({
    'Material Code': ['4AO005', '1DAT04S', '1DCT01', '2AE06', '2CC02', '4BT021G', '2AB01-C', '4BT008G', '4BT011G'],
    'Item Category': ['Seal', 'Vial', 'Vial', 'Ampoule', 'Ampoule', 'Seal', 'Ampoule', 'Seal', 'Seal']
})

# Canonical column names used by the processing logic
COL_MAT_TYPE = 'Mat Type'
COL_MATERIAL_CODE = 'Material Code'
COL_MATERIAL_NAME = 'Material Name'
COL_UOM = 'UOM'
COL_PO_DT = 'P.O. Dt.'        # canonical P.O. date (with dot)
COL_PO_NO = 'P. O. No.'      # canonical PO number
COL_SUPPLIER = 'Supplier'
COL_PO_QTY = 'PO Qty.'       # canonical PO qty
COL_GNR_DT = 'GNR Dt.'       # canonical GRN date
COL_INWARD_QTY = 'Inward Qty.'
COL_ITEM_CAT = 'Item Category'  # optional

REQUIRED_COLS = [
    COL_MAT_TYPE, COL_MATERIAL_CODE, COL_MATERIAL_NAME, COL_UOM,
    COL_PO_DT, COL_PO_NO, COL_SUPPLIER, COL_PO_QTY, COL_GNR_DT, COL_INWARD_QTY
]

DEFAULT_RULES = {"RM": 30, "SPM": 15, "TPM": 15}
DEFAULT_PPM_LT = 30

PPM_CATEGORY_MAP = {
    7:  ['Vial', 'Rubber Stopper', 'Rubber', 'Stopper', 'Seal', 'Cap', 'Collar', 'Inner Cap', 'Outer Cap'],
    12: ['Ampoule', 'Amp'],
    90: ['Pfs Syringe', 'Plunger Stopper', 'Plunger', 'U plug', 'U-plug'],
    15: ['Al Tube', 'Plastic Bottle', 'Plastic Nozzle', 'Nozzle'],
}

# Colab-specific behavior: default lead time assigned to unknown Mat Types (change if needed)
DEFAULT_UNKNOWN_LEAD_TIME = 30   # days assigned if Mat Type is unknown
# Optional override dict (set before uploading if you wish)
# Example: custom_lead_times = {"CUSTOMTYPE": 20}
custom_lead_times = {}

# ---------- UTILITIES: robust Excel reading & column normalization ----------
def read_excel_auto(path_or_buffer):
    """
    Read an excel file (path or file-like) while auto-selecting the engine based on file extension.
    Works with .xls (xlrd) and .xlsx/.xlsm (openpyxl). Falls back if necessary.
    """
    # If a path-like string was passed, we can inspect the extension
    engine = None
    fname = None

    # If the argument is a path string
    if isinstance(path_or_buffer, str):
        fname = path_or_buffer
        ext = os.path.splitext(fname)[1].lower()
        if ext == '.xls':
            engine = 'xlrd'
        elif ext in ('.xlsx', '.xlsm', '.xltx', '.xltm'):
            engine = 'openpyxl'
        else:
            # unknown extension: try openpyxl then xlrd
            engine = None

        if engine:
            try:
                df = pd.read_excel(fname, engine=engine)
                print(f"Read '{fname}' using engine='{engine}'.")
                return df
            except Exception as e:
                print(f"Failed reading with engine='{engine}': {e}. Will attempt fallbacks.")
                # fallthrough to generic attempts

    # If argument is file-like (e.g., BytesIO), try openpyxl first, then xlrd
    attempts = []
    try:
        # prefer openpyxl for modern xlsx files
        df = pd.read_excel(path_or_buffer, engine='openpyxl')
        print("Read Excel using engine='openpyxl'.")
        return df
    except Exception as e_open:
        attempts.append(("openpyxl", str(e_open)))
        try:
            df = pd.read_excel(path_or_buffer, engine='xlrd')
            print("Read Excel using engine='xlrd'.")
            return df
        except Exception as e_xl:
            attempts.append(("xlrd", str(e_xl)))
            # If both fail, surface a helpful message
            msg = "Failed to read Excel file. Attempts:\n"
            for eng, err in attempts:
                msg += f" - engine={eng}: {err}\n"
            raise ValueError(msg)

def standardize_column_names(df: pd.DataFrame) -> pd.DataFrame:
    """
    Robustly map many variants of column names to canonical names used by the script.
    Preserves any columns not recognized (they remain with near-original name, cleaned).
    """
    def map_one(col):
        orig = str(col).strip()
        # clean spaces and NBSP
        s = re.sub(r'\u00A0', ' ', orig)
        s = re.sub(r'\s+', ' ', s).strip()
        key = re.sub(r'[\s\.\-_/()]', '', s).lower()

        # mapping dictionary based on common variants
        if key in ('mattype','materialtype'):
            return COL_MAT_TYPE
        if key in ('materialcode','matcode','material_code'):
            return COL_MATERIAL_CODE
        if key in ('materialname','itemname','material'):
            return COL_MATERIAL_NAME
        if key == 'uom':
            return COL_UOM
        if key in ('podt','podat','podate','podatet','podt'):
            return COL_PO_DT
        if key in ('pono','ponumber','po'):
            return COL_PO_NO
        if key in ('supplier','suppliername'):
            return COL_SUPPLIER
        if key in ('poqty','poqty','purchaseorderqty','quantityordered'):
            return COL_PO_QTY
        if key in ('gnrdt','grndt','grndate','grn','grndate'):
            return COL_GNR_DT
        if key in ('inwardqty','inwardquantity','receivedqty','receivedquantity'):
            return COL_INWARD_QTY
        if key in ('itemcategory','itemcat','category'):
            return COL_ITEM_CAT
        # else return the cleaned original (preserve punctuation if user wants)
        return orig

    new_cols = [map_one(c) for c in df.columns]
    df = df.copy()
    df.columns = new_cols
    return df

# ---------- HELPERS: original logic, but using canonical column variables ----------
def compute_lead_time_for_row(row: pd.Series, rules: dict):
    mat_type = str(row.get(COL_MAT_TYPE, "")).strip().upper()
    if mat_type in rules:
        return rules[mat_type]
    if mat_type == "PPM":
        item_cat = str(row.get(COL_ITEM_CAT, "") or "").strip()
        if item_cat:
            low = item_cat.lower()
            for lt, cats in PPM_CATEGORY_MAP.items():
                if low in [c.lower() for c in cats]:
                    return lt
        return DEFAULT_PPM_LT
    return np.nan

def ensure_types_and_drop_nulls(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df[COL_PO_DT] = pd.to_datetime(df[COL_PO_DT], errors="coerce", dayfirst=True)
    df[COL_GNR_DT] = pd.to_datetime(df[COL_GNR_DT], errors="coerce", dayfirst=True)
    df[COL_PO_QTY] = pd.to_numeric(df[COL_PO_QTY], errors="coerce")
    df[COL_INWARD_QTY] = pd.to_numeric(df[COL_INWARD_QTY], errors="coerce")
    df = df.dropna(subset=[COL_PO_DT, COL_GNR_DT, COL_PO_QTY, COL_INWARD_QTY]).copy()
    return df

def merge_item_category(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    if COL_ITEM_CAT not in df.columns:
        df = df.merge(df1, on=COL_MATERIAL_CODE, how="left")
    else:
        df = df.merge(df1, on=COL_MATERIAL_CODE, how="left", suffixes=("", "_map"))
        if "Item Category_map" in df.columns:
            df[COL_ITEM_CAT] = df[COL_ITEM_CAT].fillna(df["Item Category_map"])
            df.drop(columns=["Item Category_map"], inplace=True, errors="ignore")
    df[COL_ITEM_CAT] = df.get(COL_ITEM_CAT, "").fillna("")
    return df

def compute_po_level_metrics(df: pd.DataFrame):
    # Product-level fulfillment (sum duplicates)
    df_po_item = (
        df.groupby([COL_PO_NO, COL_MATERIAL_CODE], as_index=False)
          .agg({COL_PO_QTY: "sum", COL_INWARD_QTY: "sum"})
    )
    df_po_item["Fulfilled"] = (df_po_item[COL_INWARD_QTY] >= 0.95 * df_po_item[COL_PO_QTY]).astype(int)

    # PO-level In-Full
    df_po_status = (
        df_po_item.groupby(COL_PO_NO)["Fulfilled"]
                  .min()
                  .reset_index()
                  .rename(columns={"Fulfilled": "PO_Fulfilled"})
    )
    df_line = df.merge(df_po_status, on=COL_PO_NO, how="left")

    # PO-level On-Time
    def po_ontime(group: pd.DataFrame) -> int:
        due_dates = group[COL_PO_DT] + pd.to_timedelta(group["Lead Time"], unit="D")
        return int((group[COL_GNR_DT] <= due_dates).all())

    po_ontime_df = df_line.groupby(COL_PO_NO).apply(po_ontime).reset_index(name="OnTime")
    df_line = df_line.merge(po_ontime_df, on=COL_PO_NO, how="left")

    # Collapse to one-row-per-PO (keep other PO-level columns)
    df_po = df_line.drop_duplicates(subset=[COL_PO_NO]).copy()
    df_po["OTIF"] = (df_po["PO_Fulfilled"].astype(int) * df_po["OnTime"].astype(int)).astype(int)

    # Use LAST GNR date per PO for bucketing
    po_last_grn = (
        df.groupby(COL_PO_NO, as_index=False)[COL_GNR_DT]
          .max()
          .rename(columns={COL_GNR_DT: "PO_GNR_Dt"})
    )
    df_po = df_po.merge(po_last_grn, on=COL_PO_NO, how="left")
    df_po[COL_GNR_DT] = pd.to_datetime(df_po["PO_GNR_Dt"])
    df_po.drop(columns=["PO_GNR_Dt"], inplace=True)

    # Add Year/Month parts
    df_po["Year"] = df_po[COL_GNR_DT].dt.year
    df_po["MonthNum"] = df_po[COL_GNR_DT].dt.month
    df_po["Month"] = df_po[COL_GNR_DT].dt.strftime("%b")

    return df_line, df_po

def generate_failed_orders_pdf_colab(breaches_df: pd.DataFrame, vendor_stats: pd.DataFrame, year: int, output_path: str):
    buf = io.BytesIO()
    width, height = A4
    c = canvas.Canvas(buf, pagesize=A4)
    margin_x = 20 * mm
    y = height - 20 * mm
    line_height = 8 * mm

    c.setFont("Helvetica-Bold", 16)
    c.drawString(margin_x, y, f"ALL OTIF FAILED ORDERS — {year}")
    y -= 12 * mm

    vendor_order = vendor_stats[vendor_stats["OTIF_Failures"]>0].sort_values("OTIF_Failures", ascending=False)

    vendor_idx = 1
    for _, vr in vendor_order.iterrows():
        vendor = vr["Supplier"]
        failures = int(vr["OTIF_Failures"])
        total_orders = int(vr["Total_Orders"])
        otif_pct = float(vr["Vendor_OTIF_pct"])
        contrib_pct = float(vr["Total_Contribution_pct"])

        vendor_group = breaches_df[breaches_df["Supplier"] == vendor].sort_values(COL_GNR_DT, ascending=False)

        if y < 40 * mm:
            c.showPage()
            y = height - 20 * mm
            c.setFont("Helvetica-Bold", 16)
            c.drawString(margin_x, y, f"ALL OTIF FAILED ORDERS — {year} (cont.)")
            y -= 12 * mm

        c.setFont("Helvetica-Bold", 12)
        header = f"{vendor_idx}. {vendor}   (Failures: {failures})   OTIF: {otif_pct:.1f}%   Contribution: {contrib_pct:.1f}%   Total Orders: {total_orders}"
        c.drawString(margin_x, y, header)
        y -= line_height

        c.setFont("Helvetica", 10)
        for _, orow in vendor_group.iterrows():
            ord_date = orow.get(COL_GNR_DT)
            date_str = "" if pd.isna(ord_date) else pd.to_datetime(ord_date).strftime("%d-%m-%Y")
            po_no = str(orow.get(COL_PO_NO, ""))
            line_text = f"    {date_str}    {po_no}"
            c.drawString(margin_x + 6 * mm, y, line_text)
            y -= (6 * mm)
            if y < 25 * mm:
                c.showPage()
                y = height - 20 * mm
                c.setFont("Helvetica", 10)
        y -= 4 * mm
        vendor_idx += 1

    c.save()
    buf.seek(0)
    with open(output_path, "wb") as f:
        f.write(buf.read())
    print(f"PDF written to: {output_path}")

# ---------- RUN: Upload file & process ----------
from google.colab import files
uploaded = files.upload()
if not uploaded:
    raise SystemExit("No file uploaded.")
# take the first uploaded file
file_name = list(uploaded.keys())[0]
print("Uploaded:", file_name)

# Read Excel with auto-detection
try:
    df_raw = read_excel_auto(file_name)
except Exception as e:
    # Provide the detailed exception so user can inspect
    raise RuntimeError(f"Error reading Excel file: {e}")

# Normalize column names to canonical names (accepts dotted or non-dotted variants)
df_raw = standardize_column_names(df_raw)
missing = [c for c in REQUIRED_COLS if c not in df_raw.columns]
if missing:
    raise ValueError(f"Input file missing required columns (after normalization): {missing}. Columns found: {list(df_raw.columns)}")

# Convert dtypes & drop nulls (will use canonical names)
df_raw = ensure_types_and_drop_nulls(df_raw)
if df_raw.empty:
    raise ValueError("No usable rows after type coercion / dropping nulls.")

# Merge Item Category mapping (if Item Category not present, merge from df1)
df = merge_item_category(df_raw)

# Lead-time rules: known defaults + user overrides (custom_lead_times)
lead_time_rules = DEFAULT_RULES.copy()
# apply overrides from custom_lead_times dict (if user configured)
if custom_lead_times:
    for k, v in custom_lead_times.items():
        lead_time_rules[str(k).strip().upper()] = int(v)

# Detect Mat Types present and auto-assign default lead times for unknown Mat Types
mat_types_in_data = set(df[COL_MAT_TYPE].dropna().astype(str).str.strip().str.upper().unique().tolist())
known_keys = set(lead_time_rules.keys()) | set(["PPM"])
unknowns = sorted(mat_types_in_data - known_keys)

if unknowns:
    print("Found Mat Types with no specified lead time. Assigning default lead time (days) to them:")
    for u in unknowns:
        print(f"  - {u} -> {DEFAULT_UNKNOWN_LEAD_TIME} days")
        lead_time_rules[u] = DEFAULT_UNKNOWN_LEAD_TIME

# Compute Lead Time for each row
df["Lead Time"] = df.apply(lambda row: compute_lead_time_for_row(row, lead_time_rules), axis=1)

# Final fallback — if any rows still have NaN lead time, fill with default and notify
if df["Lead Time"].isna().any():
    print("Warning: Some rows still lack Lead Time; filling with DEFAULT_UNKNOWN_LEAD_TIME.")
    print(df.loc[df["Lead Time"].isna(), [COL_MAT_TYPE]].drop_duplicates().head(10).to_string(index=False))
    df["Lead Time"] = df["Lead Time"].fillna(DEFAULT_UNKNOWN_LEAD_TIME)

# PO-level metrics
df_line, df_po = compute_po_level_metrics(df)

years = sorted(df_po["Year"].dropna().unique().astype(int).tolist())
if not years:
    raise ValueError("No valid years found in processed data.")
selected_year = years[-1]  # choose most recent by default
print("Selected year:", selected_year)

po_year = df_po[df_po["Year"] == selected_year].copy()

# Monthly summary
monthly = (
    po_year.groupby(["MonthNum", "Month"], as_index=False)
           .agg(
               Avg_OTIF=("OTIF", "mean"),
               Avg_OnTime=("OnTime", "mean"),
               Avg_InFull=("PO_Fulfilled", "mean"),
               Total_Orders=(COL_PO_NO, "count")
           )
           .sort_values("MonthNum")
)

# Vendor stats for the selected year
total_orders_year = po_year.shape[0]
vendor_stats = (
    po_year.groupby(COL_SUPPLIER, dropna=False)
           .agg(Total_Orders=(COL_PO_NO, "count"),
                OTIF_Failures=("OTIF", lambda x: int((x==0).sum())),
                OTIF_Success=("OTIF", lambda x: int((x==1).sum())))
           .reset_index()
)
vendor_stats[COL_SUPPLIER] = vendor_stats[COL_SUPPLIER].fillna("Unknown Supplier")
vendor_stats["Vendor_OTIF_pct"] = vendor_stats["OTIF_Success"] / vendor_stats["Total_Orders"] * 100
vendor_stats["Total_Contribution_pct"] = vendor_stats["Total_Orders"] / (total_orders_year if total_orders_year>0 else 1) * 100

# Top 10 vendors by failures
top10 = vendor_stats[vendor_stats["OTIF_Failures"]>0].sort_values("OTIF_Failures", ascending=False).head(10)
if top10.empty:
    print("No OTIF failures found in selected year.")
else:
    # Format and display top10
    display_top10 = top10[[COL_SUPPLIER, "OTIF_Failures", "Vendor_OTIF_pct", "Total_Contribution_pct", "Total_Orders"]].copy()
    display_top10["Vendor_OTIF_pct"] = display_top10["Vendor_OTIF_pct"].map(lambda v: f"{v:.1f}%")
    display_top10["Total_Contribution_pct"] = display_top10["Total_Contribution_pct"].map(lambda v: f"{v:.1f}%")
    print("\nTop vendors (failures, OTIF%, contribution%):")
    print(display_top10.to_string(index=False))

# Save CSVs
po_csv = f"po_level_{selected_year}.csv"
monthly_csv = f"monthly_otif_{selected_year}.csv"
po_year.to_csv(po_csv, index=False)
monthly.to_csv(monthly_csv, index=False)
print(f"Saved: {po_csv}, {monthly_csv}")

# Generate PDF for failures (grouped by supplier)
breaches = po_year[po_year["OTIF"] == 0].copy()
if breaches.shape[0] == 0:
    print("No OTIF breaches in selected year.")
else:
    out_pdf = f"OTIF_failed_orders_{selected_year}.pdf"
    generate_failed_orders_pdf_colab(breaches[[COL_SUPPLIER, COL_GNR_DT, COL_PO_NO]], vendor_stats, selected_year, out_pdf)
    # Offer file for download in Colab:
    from google.colab import files as gfiles
    gfiles.download(out_pdf)

print("Completed processing.")
