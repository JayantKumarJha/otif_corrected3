# pages/4_Vendor_OTIF.py
# Streamlit OTIF vendor analysis page
# Save to pages/4_Vendor_OTIF.py in your multipage Streamlit repo.

import io
import os
import re
from datetime import datetime
import pandas as pd
import numpy as np
import streamlit as st
import plotly.express as px

# Try to import reportlab for PDF generation
try:
    from reportlab.lib.pagesizes import A4
    from reportlab.pdfgen import canvas
    from reportlab.lib.units import mm
    reportlab_available = True
except Exception:
    reportlab_available = False

st.set_page_config(page_title="Vendor OTIF Analysis", page_icon="📦", layout="wide")
st.title("📦 OTIF (On-Time In-Full) Analysis — Order Level")
st.caption(
    "Upload procurement/GRN Excel → we'll clean it, apply lead-time rules, compute PO-level In-Full, On-Time, and OTIF. "
    "Use the Mat Type checkboxes in the left sidebar. Month bucketing uses the PO's last GRN date."
)

# ------------------------- CONSTANTS (canonical dotted columns) -------------------------
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

# --------------------------- HARD-CODED ITEM CATEGORY MAP (placeholder) ---------------------------
# Replace/extend this mapping with your real Material Code → Item Category mapping as needed.
df1 = pd.DataFrame({'Material Code': ['4AO005', '1DAT04S', '1DCT01', '2AE06', '2CC02', '4BT021G', '2AB01-C', '4BT008G', '4BT008G', '4BT008G', '4BT011G', '4BT015G', '4BT015G', '4BT019G', '4BT031G', '4BT035G', '4BT036G', '4BT036G', '4BT038G', '4BT050G', '4BT050G', '1FCM01', '2CD09-C', '5DAGB01', '1DCM03', '1FCM03', '1HCM03', '1HCM03', '1HCM03', '1HCM03', '1HCM03', '1HCM03', '1HCM03', '2CD08', '3BG02R', '2AB02', '1ICM01Y', '2AB01-C', '2AB01-C', '2CA07-C', '3CG02', '3CG02', '4CT045G', '2AD01-C', '12OT07', '2CA07-P', '4BO040G', '1FCM03', '4BT017G', '4BT01G', '4BT043G', '1BAM01', '1BAM01', '1BAM01', '1BAM01', '2CC09-C', '9PFSA01-B', '9PFSA03-A', '9PFSA03-A', '4BT036G', '4BT036G', '1DCT01R', '4BO015G', '4BO015G', '1DCT01', '4BO018G', '2CB01-C', '1BAM01', '1DCT01R', '3BG02R', '4BO015G', '4BO015G', '1BCM03', '1DAT04S', '1FAM01', '3BG02', '4BO040G', '4BO040G', '4BT013G', '4BT051G', '1DAT01L', '2AB02', '2AB02', '2AB02', '2CC09-C', '4BO049G', '4BO049G', '4BT003G', '4BT010G', '4BT050G', '2AA01-C', '1DAM01', '1DCT01R', '1FCT01R', '2CA23-C', '2CA23-C', '4BT025G', '1FAM01', '3BS02', '3BS02', '4BO040G', '4BO040G', '7PBA01-A', '7PBA01-A', '7PBB01-B', '7PBB01-B', '7PBB01-C', '4BT011G', '1DAT01L', '1DCT01', '2AA01', '3BG02-U', '4BT025G', '4BT025G', '4BT049G', '4BT049G', '4BT015G', '1FCM03', '1HCM03', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '4BT011G', '4BT035G', '4BT036G', '4BT036G', '1HCM01', '4BT019G', '1CCM03', '1FCM03', '1FCM03', '1FCM03', '1FCM03', '4BT006G', '5DAGB01', '4BT008G', '4BT008G', '4BT008G', '4BT008G', '1DCT01', '2CE02', '2CE02', '2CE02', '1HCM03', '1HCM03', '1HCM03', '1HCM03', '1HCM03', '4BT022G', '4BT031G', '4BT035G', '4BT036G', '4BT036G', '4BT036G', '4BT036G', '1HCM01', '1BCM03', '1ICM01', '4BO043G', '4BT034G', '4BT034G', '4BT034G', '2AB07', '4BT008G', '1DAM01', '4AO002', '4BT025G', '1ACT01', '1DCT01R', '1FCT01R', '3AG02', '3BG02R', '4BO015G', '1BCM03', '1DAT01L', '1HCM01', '2AB02', '2AB02', '2AB02', '4BO018G', '4BO01G', '4BO01G', '4BO049', '4BT025G', '2AD02', '4BO049G', '4BO049G', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '4BT008G', '4BT019G', '1FCM01', '1FCM01', '1FCM01', '1FCM01', '2AD02', '2AD02', '2AA02', '3BG02', '1DAT01L', '7PBA01-A', '7PBB01-B', '7PBB01-C', '1BCT01S', '1HCM01', '3BG02', '3BG02', '3BG02', '3BG02', '4BO01G', '4BO049G', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '4BT050G', '1FCM01', '4BO044G', '5DAGB01', '4BT063G', '4BT008G', '1BAM01', '1HAM01', '2AE06', '2CC09-C', '4BO015G', '4BT015G', '4BT050G', '2CA07', '4BT013G', '4BT013G', '4BT038G', '1BCT01S', '1BCT01S', '2AA02', '4BO01', '1ECM03', '4BT006G', '1BCM03', '2CC09-C', '2CE04', '4BT034G', '4BT034G', '1DAM01', '1FCM01', '2CC09', '2CC09', '4BO031G', '4BO044G', '1CCM03', '1CCM03', '2AE02', '1DCT01R', '1FCT01R', '2AC07', '2CB01-C', '4BO015G', '1DCM03', '1DCM03', '1DCM03', '1DAM03', '1ICM01', '2CD08', '2CD08', '4BO01G', '4BO044G', '4BT025G', '1ECM03', '4BT008G', '4BT011G', '4BT045G', '1BCT01S', '1BCT01S', '1BCT01S', '1DAT01L', '2AB07', '4BT021G', '4BT02G', '7PBB04-A', '7PBB04-B', '7PBB04-C', '2CD09-C', '1HCM03', '4BT036G', '1DAM01', '1DAT04S', '1HCM01', '2AC07', '2CE02', '4BO015G', '4BO031G', '4BO044G', '4BT021G', '4BT021G', '4BT021G', '1DAM03', '4BO049G', '4BT025G', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '1ACT01', '3AG04', '3AG04', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '4AT003G', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '4BT021G', '4BT008G', '4BT011G', '1FCM01', '2AD02', '4BT015G', '5DAGB01', '9PFSA01-B', '9PFSA01-B', '9PFSA03-A', '9PFSA03-A', '1DCM01', '2CD08', '4BT024G', '1DCM03', '4BT036G', '4BT063G', '2CB01-C', '2CD09-C', '2CE04', '4BT008G', '4BT011G', '4BT036G', '4BT038G', '4BT050G', '1BCM03', '4BT011G', '1DCM03', '5CAGB01', '6C02', '2AB01-C', '2AD02', '2AA01-C', '4BO015G', '1CCM03', '4BT008G', '4BT008G', '4BT022G', '4BT015G', '4AO005G', '4BT034G', '4BT016G', '1BAM01', '1ECT01R', '1HCM01', '2AA01-C', '2CC09', '4BO044G', '4BT015G', '4BT021G', '7PBC01-A', '4BT006G', '1BCM03', '1BCM03', '1BCM03', '1BCM03', '1BCM03', '1BCM03', '1DAT04S', '1DAT04S', '1FCM01', '2CE10-C', '4BO044G', '4BT011G', '4BT033G', '1BAM01', '1DCT01', '4BT02G', '7PBA01-A', '7PBB01-B', '7PBB01-C', '1CCM03', '1DCM03', '1FCM01', '1FCM03', '1FCM03', '1FCM03', '1FCM03', '1FCM03', '1FCM03', '1FCM03', '1FCM03', '1HCM01', '1HCM01', '1HCM03', '1HCM03', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '1ACT01', '1BCM03', '1BCM03', '1BCM03', '1BCM03', '1BCM03', '1BCM03', '1BCT01S', '1BCT01S', '1BCT01S', '1BCT01S', '1BCT01S', '1BCT01S', '1BCT01S', '1BCT01S', '1BCT01S', '1DAM01', '1DCM03', '1FCM01', '1HCM01', '1HCM01', '1ICM01', '2AA01-C', '2AB01-C', '2AB02', '2AB02', '2AD02', '2AD02', '2AD02', '2AD02', '2AE02', '2CA07-C', '2CB01-C', '2CC09-C', '2CD08', '2CD08', '2CD09-C', '2CD09-C', '2CD09-C', '2CD09-C', '2CE02', '2CE02', '2CE10-C', '3AG02', '3AG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3CG02', '3CG02', '3CG02', '3CG02', '4BT02G', '4BT02G', '1FCT02', '1HCM03', '3BG06', '3BG06', '3BG06', '3BG06', '3BG06', '3BG06', '4BT015G', '4BT065', '4BT065', '4BT066', '4BT066', '4BT021G', '4BT011G', '4BT013G', '4BT033G', '4BT033G', '4BT033G', '4BT008G', '4BT008G', '2AB07', '2AE06', '2CD09-C', '2CD09-C', '2CD09-C', '2CD09-C', '4BT033G', '4BT033G', '4BT033G', '1FCM01', '4BO044G', '4BT006G', '4BT008G', '4BT008G', '1DAM03', '1DCT01R', '1FCT01R', '4BO015G', '4BO015G', '4BT034G', '4BT034G', '4BT034G', '7PBB04-A', '7PBB04-B', '7PBB04-C', '4BT019G', '4BT02G', '4BT02G', '7PBB04-A', '7PBB04-B', '7PBB04-C', '4BT003G', '4BT011G', '4BT013G', '4BT019G', '4BT032G', '4BT051G', '1DCM01', '4BO015G', '4BO015G', '4BO01G', '4BO01G', '4BO01G', '4BO044G', '4BT011G', '4BT011G', '4BT015G', '4BT015G', '4BT021G', '4BT021G', '4BT021G', '4BT025G', '4BT025G', '4BT025G', '4BT02G', '4BT02G', '4BT033G', '4BT033G', '4BT034G', '4BT034G', '4BT034G', '1BAM01', '4BT034G', '4BT034G', '4BT034G', '2CA07-C', '1FAM01', '4BO040G', '4BO040G', '4BT008G', '4BT008G', '4BT011G', '2AB02', '2AB02', '2AB02', '2AB02', '5DAGB01', '7PBB01-A', '4BT009G', '2CC09', '9PFSA01-B', '9PFSA01-B', '9PFSA03-A', '1DCM03', '2AD02', '2AD02', '2CE10-C', '3BG05', '4BT019G', '1BCM03', '1HAM01', '4BO018G', '4BO055G', '7PBB01-A', '7PBB01-B', '7PBB01-C', '4BT015G', '3BG04', '4BT008G', '4BT008G', '4BT008G', '9PFSA06-A', '9PFSA06-B', '9PFSA06-C', '9PFSA06-D', '1HCM03', '1HCM03', '4BT011G', '4BT011G', '4BT013G', '1DAM01', '1IAM01', '1ICM01Y', '2AB07', '2CD07', '3CG02', '3CG02', '3CG02', '3CG02', '4BO058G', '4CT002G', '4CT002G', '5CAGB01', '5DAGB01', '6C02', '7PBB04-A', '7PBB04-B', '7PBB04-C', '8CP06', '9PL01', '2AD02', '2CE02', '4BT021G', '4BT021G', '4BT021G', '4BT022G', '1BCT01S', '1BCT01S', '1BCT01S', '1BCT01S', '1DCT01', '1HCM01', '2AA01-C', '2AB02', '2AB02', '2AD01-C', '4AO043G', '1DCM03', '4BT015G', '4BT015G', '1DAT01L', '1DCM03', '2AB02', '2AB02', '4AO001G', '1FCM03', '4BT008G', '4BT008G', '4BT008G', '4BT036G', '1FCM01', '2AE06', '2AE06', '1CCM03', '4BT063G', '1FAM01', '1BAM01', '2AA01-C', '1DCM03', '1BCM03', '1HCM01', '2AA01-C', '2AB01-C', '2AD02', '2AD02', '2CA02', '2CB01-C', '2CD09-C', '4BT034G', '7PBB02-A', '7PBB02-B', '7PBB02-C', '1ACT01', '1ACT01', '3AG02', '3AG02', '3AG02', '3AG02', '4AO005G', '4AO005G', '1FAT01', '2CE08-C', '4BT008G', '7PBB04-A', '7PBB04-B', '2AA01-C', '2AB01-C', '2CB01-C', '2CD09-C', '1BAT01', '2CB01-C', '4BT063G', '1DCT01R', '1FCT01R', '2AD01-C', '3BG02', '1FCM03', '1FCM03', '1FCM03', '1FCM03', '1FCM03', '1FCM03', '1FCM03', '1HCM03', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '1BCT01S', '1BCT01S', '1BCT01S', '1BCT01S', '1BCT01S', '1BCT01S', '1BCT01S', '4BT02G', '4BT02G', '4BT02G', '4BT02G', '4BT02G', '4BT02G', '4BT02G', '1FCM01', '1FCM01', '1FCM01', '1FCM03', '2AD02', '2AD02', '2CA07-C', '2CA27', '2CC02', '4AO001G', '4BO01G', '4BO01G', '4BT011G', '4BT025G', '1BAM01', '4BO034G', '4BO034G', '4BT068G', '2CC09-C', '2CD02', '2CD02', '2CD02', '7PBA01-A', '7PBB01-B', '7PBB01-C', '1DAM03', '1DCM01', '4BT008G', '4BT008G', '4BT008G', '4BT011G', '4BT013G', '4BT015G', '4BT022G', '4BT022G', '4BT022G', '4BT036G', '4BT036G', '4BT036G', '4BT050G', '4BT063G', '2CA07-C', '2CB01-C', '1DCM03', '2AE02', '1FCM03', '4AO005G', '4AO005G', '7PBB01-B', '7PBB01-C', '1FCM03', '1HCM03', '4BT008G', '4BT011G', '4BT011G', '4BT011G', '4BT013G', '4BT015G', '4BT019G', '4BT043G', '1BAM01', '4BT011G', '4BT036G', '4BT050G', '1DCM03', '1DCM03', '1DCM03', '1FCM03', '1FCM03', '1BCM03', '1DAM03', '2CE02', '2CE02', '2CE02', '2CE02', '4BT005G', '4BT034G', '4BT034G', '4BT034G', '1FCM01', '4BO044G', '4BT011G', '4BT050G', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '1FAM01', '1FAM01', '1FCM01', '2AD02', '2AD02', '2CA07', '2CE10-C', '3BS02', '2CE02', '2CE02', '2CE02', '2CE02', '4BT006G', '1DCT01', '2AB02', '2AD02', '2AD02', '2AD02', '3AC02', '3BG04', '4BT034G', '4BT034G', '4BT034G', '2CB01', '4BT008G', '4BT008G', '4BT063G', '2AB01-C', '5DAGB01', '5DAGB01', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '4BT008G', '4BT008G', '4BT016G', '4BT036G', '1FCM03', '1FCM03', '1HCM03', '3BG07', '3BG07', '1BAM01', '2CC09-C', '2CE02', '2CE02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '1DAM03', '1DCT01', '2AE06', '4AO001G', '1DCM03', '1DCM03', '1DCM03', '1DCM03', '1DCM03', '1HAM01', '2AA01-C', '2CA07-C', '4BO055G', '1DAT01L', '4BO049G', '1ACT01', '1BCM03', '4AO002G', '4BT034G', '4BT034G', '4BT034G', '4BT034G', '1BCT01', '1DCM03', '2CC09', '2CE08-C', '4BT011G', '4BT011G', '4BT049G', '4BT003G', '4BT032G', '4BT050G', '1BAM01', '1DAM01', '2CD02', '2CD02-LE', '4BO031G', '7PBB04-A', '7PBB04-B', '4BT048G', '1DAM03', '2CC09-C', '2CD02-LE', '2CD02-LE', '2CE08-C', '4AO043G', '4AO043G', '7PBB02-A', '7PBB02-B', '7PBB02-C', '4BT017G', '4BT017G', '4BT01G', '4BT022G', '4BT022G', '4BT022G', '1DCM01', '2AB07', '2CE10-C', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '7PBB04-A', '7PBB04-B', '1BAM01', '2CC09', '2CD04', '4BT024G', '4BT02G', '7PBA01-A', '7PBB01-B', '7PBB01-B', '7PBB01-C', '4BT008G', '4BT008G', '4BT008G', '4BT021G', '2AB01-C', '2CD02', '2CD02', '1DCM03', '12OT11', '12OT11', '12OT11', '12OT11', '12OT11', '12OT11', '2CA07-C', '2CB01-C', '2CD02', '2CD02', '2CD02', '2CD02', '4BO036', '1HAM01', '2AB01-C', '2CE10', '4BO055G', '4BO01G', '4BO01G', '4BT021G', '4BT021G', '4BT021G', '4BT021G', '4BT021G', '4BT021G', '4BT021G', '4BT015G', '4BT015G', '4BT015G', '4BT022G', '4BT022G', '2CC09-C', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '2CB01-C', '1DCT01-PB', '1JCT01-PB', '2CB34-PB', '3BG02-PB', '3BG02-PB', '3BG02-PB', '3BG02-PB', '3BG02-PB', '3BG02-PB', '3BG02-PB', '3BG02-PB', '3BG02-PB', '3BG02-PB', '4BO036-PB', '4BO043-PB', '4BO061-PB', '4BO063-PB', '4BT011G', '4BT019G', '4BT019G', '2AB07', '2AE02', '4BT006G', '4BT009G', '4BT009G', '4BT012G', '1DCT01R', '1BAM01', '4BO034G', '4BT025G', '2CC09-C', '1DCM03', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '4BT008G', '4BT015G', '4BT050G', '4BT036G', '4BT036G', '4BT036G', '4BT025G', '4BT025G', '4BT025G', '4BT025G', '2CD02', '4BT022G', '2AB02', '2AB02', '2AD02', '2AD02', '4BT021G', '3AG02', '1FAM01', '2CA07-C', '4BO015G', '4BO015G', '2AA01-C', '1DCM03', '1DCM03', '1DCM03', '1DCM03', '1FCM03', '4BO013G', '4BT022G', '4BT022G', '2CD02', '2CD02', '2CD02', '4BT036G', '4BT036G', '7PBA01-A', '7PBB01-B', '7PBB01-C', '2AD08', '3BS02', '1FCM01', '4BO044G', '3BG08', '3BG08', '5DAGB01', '6C01', '6C02', '1DCM03', '1DCM03', '1HCM03', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '4BT015G', '1DAM03', '1DAM03', '2CA23-C', '7PBB04-A', '7PBB04-B', '1BCT01-G', '1DCT01-G', '3BG04-G', '4BT019-G', '5CAGB01', '4BT036G', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '1FCM01', '2AA01-C', '4BO044G', '1DCM03', '1DCM03', '1DCM03', '1DCM03', '1DCM03', '1DCM03', '1DCM03', '1FCM03', '1DCM01', '1DCM03', '4BT008G', '4BT008G', '4BT008G', '4BT011G', '4BT015G', '4BT015G', '2CA07', '2CC02', '2CC09', '9PFSA01-B', '9PFSA03-A', '9PFSA03-A', '9PFSA03-A', '2AA02', '4BT011G', '4BT013G', '4BT063G', '2CD08-EP', '4BO044G', '1ICM01-G', '2CC09-C', '1BCT01', '4BO018G', '7PBA01-A', '7PBB01-B', '7PBB01-C', '7PBB01-C', '1BCM03', '2AA01-C', '4BT034G', '4BT034G', '4BT034G', '4BT021G', '1CCM03', '2AC07', '1BCM03', '1BCM03', '1DAM03', '4BT034G', '1BCM03', '2CC09', '2CC09', '1HCM03', '4BT008G', '4BT008G', '2CE02', '2CE02', '2CE02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '4BT021G', '4BT021G', '4BT021G', '4BT021G', '4BT021G', '7PBB04-A', '7PBB04-A', '7PBB04-A', '7PBB04-A', '7PBB04-A', '7PBB04-A', '7PBB04-A', '7PBB04-B', '7PBB04-B', '7PBB04-B', '7PBB04-B', '7PBB04-B', '7PBB04-B', '7PBB04-B', '1DCM03', '1DCM03', '4BT006G', '4BT008G', '4BT008G', '4BT02G', '1BCT01S', '1BCT01S', '1BCT01S', '1BCT01S', '2CE02', '4BT02G', '1DCM03', '1DCM03', '4BT011G', '4BT011G', '4BT011G', '4BT050G', '2CA07-C', '4AO004G', '4BT006G', '4BT006G', '1DAM01', '1DCM03', '1FCM03', '1HAM01', '1HCM01', '1ICM01', '1ICM01', '2CB01', '2CC09-C', '4BO015G', '4BO055G', '4BT011G', '4BT025G', '4BT025G', '1DCM04', '9PFSA01-B', '9PFSA01-B', '9PFSA03-A', '1FAM01', '1FAM01', '2AD01-C', '2AD01-C', '4BO040G', '4BO036-PB', '4BO043-PB', '4BO061-PB', '4BO063-PB', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '4BT008G', '4BT008G', '4BT008G', '4BT008G', '4BT008G', '4BT008G', '4BT008G', '4BT008G', '2CE10', '2CE10', '4BT015G', '2AA02', '4BO031G', '4BT034G', '4BT034G', '4BT034G', '4BT034G', '1BCM03', '1DCT01R', '1ECT01R', '2AD01-C', '2AD01-C', '7PBB01-A', '7PBB01-B', '7PBB01-C', '4AO004', '1DCM01', '1FAM01', '4BT021G', '4BT021G', '4BT021G', '4BT021G', '4BT024G', '1FCM03', '4BT013G', '4BT013G', '4BT013G', '4BT015G', '1BAM01', '4BT02G', '4BT02G', '1BCM03', '1DCM03', '1DCM03', '1ACT01', '3AG02', '3AG02', '3AG02', '3AG02', '4AO001G', '4AO002G', '4BT006G', '6C02', '2CA07', '1BCT01X', '2CE10', '2CE10', '1IAM01', '2AE02', '4BO058G', '7PBB02-A', '7PBB02-B', '7PBB02-C', '4BT008G', '1BCM03', '1BCM03', '1DAT01S', '3BG02', '3BG02', '3BG02', '3BG02', '4BT025G', '4BT025G', '4BT034G', '1HCM03', '4BT015G', '1BAM01', '2AA01-C', '2AD01-C', '2AD01-C', '2CA07-C', '1BCM03', '1BCM03', '1BCM03', '1BCM03', '2CB01-C', '2CB01-C', '2CD08', '4BT011G', '2AA02', '2AD02', '2AD02', '2AD02', '8CP06', '9PFSA01-B', '9PFSA03-A', '9PL01', '1DCT01', '2CC09', '1FCM01', '1FCM01', '1BAM01', '2CC09-C', '2CC09-C', '2CC09-C', '4BT021G', '4BT021G', '1DAT01L', '1DCM03', '4BT021G', '4BT021G', '4BT021G', '5DAGB01', '4BT017G', '1BCT01S', '1BCT01S', '1BCT01S', '1BCT01S', '1BCT01S', '1BCT01S', '1BCT01S', '1FCM03', '4BT02G', '4BT02G', '4BT02G', '4BT02G', '4BT02G', '7PBB02-A', '7PBB02-B', '7PBB02-C', '2AB02', '2AB02', '1DAM01', '1DAM01', '4BT034G', '4BT034G', '2AD01-C', '1DCM03', '1DCM03', '1DCT01R', '1FCT01R', '2CE10', '2CE10', '3BG04', '4BO015G', '1BCM03', '1BCM03', '1DCM03', '1DCM03', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '4BT011G', '4BT011G', '4BT011G', '4BT011G', '4BT011G', '4BT011G', '4BT011G', '4BT011G', '4BT033G', '4BT033G', '4BT033G', '4BT033G', '4BT033G', '4BT033G', '4BT033G', '2AB05', '2AB05', '1FCM01', '1FCM03', '1FCM03', '1FCM03', '1HCM03', '1HCM03', '1HCM03', '1HCM03', '1HCM03', '2AA01-C', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '4BT033G', '4BT02G', '4BT048G', '2AD01-C', '2CC02', '7PBB02-A', '7PBB02-A', '7PBB02-A', '7PBB02-B', '7PBB02-B', '7PBB02-B', '7PBB02-C', '7PBB02-C', '4BT011G', '2AB35', '2AB35', '2CA07', '2CA07', '2CB01-C', '2CD04', '1DAM03', '4BO040G', '2CA07', '3AC02', '4AO044G', '4BT036G', '1FAM01', '2AB07', '5CAGB01', '2AA01-C', '3BG10', '4BT065G', '4BT066G', '1BAM01', '1DAT01L', '1DCT01R', '2AB01-C', '2AB01-C', '2AD02', '1DCM03', '1ECM03', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '4BT017G', '2CD02', '2CD02', '2CD02', '1CCM03', '10CTC01', '1FCT01R', '1FCT01R', '2AA02', '4BT008G', '2AD02', '4BT050G', '4BT034G', '1BAM01', '2AB01-C', '4BT015G', '7PBA01-A', '7PBB01-B', '7PBB01-C', '1HCM03', '1HCM03', '4BT008G', '4BT008G', '4BT008G', '4BT008G', '4BT036G', '1CCM03', '2CD02', '4BT011G', '4BT013G', '4BT050G', '2CD08', '1FCM01', '4AO001G', '4BO015G', '4BO044G', '5CAGB01', '6C01', '7PBA01-A', '7PBB01-B', '7PBB01-C', '8CP06', '9PL01', '6C02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '1BCM03', '1DAT01L', '2AA01-C', '4BT033G', '2CC09', '1DCM03', '2AD04-K', '4BT015G', '4BT035G', '1FCM01', '1FCM01', '1FCM01', '4BO044G', '1DCT01', '1DCT01R', '1BCT01X', '4AT002G', '4BT02G', '1BCT01S', '1FCM01', '2CD02', '9PFSA01-B', '9PFSA01-B', '9PFSA03-A', '9PFSA03-A', '3BG02', '3BG02', '3BG02', '9PCCC01-NA', '9PCCC02-NA', '1FCM03', '4BT063G', '7PBA01-A', '7PBB01-B', '7PBB01-C', '1DCT01R', '1FCT01R', '2CB01-C', '4BO015G', '1DAM03', '4BO031G', '1HCM03', '4BT015G', '4BT036G', '4BT034G', '1FAM01', '9PCCC01-NA', '9PCCC02-NA'], 'Item Category': ['Seal', 'Vial', 'Vial', 'Ampoule', 'Ampoule', 'Seal', 'Ampoule', 'Seal', 'Seal', 'Seal', 'Seal', 'Seal', 'Seal', 'Seal', 'Seal', 'Seal', 'Seal', 'Seal', 'Seal', 'Seal', 'Seal', 'Vial', 'Ampoule', 'Cap', 'Vial', 'Vial', 'Vial', 'Vial', 'Vial', 'Vial', 'Vial', 'Vial', 'Vial', 'Ampoule', 'Rubber Stopper', 'Ampoule', 'Vial', 'Ampoule', 'Ampoule', 'Ampoule', 'Rubber Stopper', 'Rubber Stopper', 'Seal', 'Ampoule', 'Al Tube', 'Ampoule', 'Seal', 'Vial', 'Seal', 'Seal', 'Seal', 'Vial', 'Vial', 'Vial', 'Vial', 'Ampoule', 'Plunger Stopper', 'PFS Syringe', 'PFS Syringe', 'Seal', 'Seal', 'Vial', 'Seal', 'Seal', 'Vial', 'Seal', 'Ampoule', 'Vial', 'Vial', 'Rubber Stopper', 'Seal', 'Seal', 'Vial', 'Vial', 'Vial', 'Rubber Stopper', 'Seal', 'Seal', 'Seal', 'Seal', 'Vial', 'Ampoule', 'Ampoule', 'Ampoule', 'Ampoule', 'Seal', 'Seal', 'Seal', 'Seal', 'Seal', 'Ampoule', 'Vial', 'Vial', 'Vial', 'Ampoule', 'Ampoule', 'Seal', 'Vial', 'Rubber Stopper', 'Rubber Stopper', 'Seal', 'Seal', 'Plastic Bottles', 'Plastic Bottles', 'Cap', 'Cap', 'Plastic Nozzle', 'Seal', 'Vial', 'Vial', 'Ampoule', 'Rubber Stopper', 'Seal', 'Seal', 'Seal', 'Seal', 'Seal', 'Vial', 'Vial', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Seal', 'Seal', 'Seal', 'Seal', 'Vial', 'Seal', 'Vial', 'Vial', 'Vial', 'Vial', 'Vial', 'Seal', 'Cap', 'Seal', 'Seal', 'Seal', 'Seal', 'Vial', 'Ampoule', 'Ampoule', 'Ampoule', 'Vial', 'Vial', 'Vial', 'Vial', 'Vial', 'Seal', 'Seal', 'Seal', 'Seal', 'Seal', 'Seal', 'Seal', 'Vial', 'Vial', 'Vial', 'Seal', 'Seal', 'Seal', 'Seal', 'Ampoule', 'Seal', 'Vial', 'Seal', 'Seal', 'Vial', 'Vial', 'Vial', 'Rubber Stopper', 'Rubber Stopper', 'Seal', 'Vial', 'Vial', 'Vial', 'Ampoule', 'Ampoule', 'Ampoule', 'Seal', 'Seal', 'Seal', 'Seal', 'Seal', 'Ampoule', 'Seal', 'Seal', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Seal', 'Seal', 'Vial', 'Vial', 'Vial', 'Vial', 'Ampoule', 'Ampoule', 'Ampoule', 'Rubber Stopper', 'Vial', 'Plastic Bottles', 'Cap', 'Plastic Nozzle', 'Vial', 'Vial', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Seal', 'Seal', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Seal', 'Vial', 'Seal', 'Cap', 'Seal', 'Seal', 'Vial', 'Vial', 'Ampoule', 'Ampoule', 'Seal', 'Seal', 'Seal', 'Ampoule', 'Seal', 'Seal', 'Seal', 'Vial', 'Vial', 'Ampoule', 'Seal', 'Vial', 'Seal', 'Vial', 'Ampoule', 'Ampoule', 'Seal', 'Seal', 'Vial', 'Vial', 'Ampoule', 'Ampoule', 'Seal', 'Seal', 'Vial', 'Vial', 'Ampoule', 'Vial', 'Vial', 'Ampoule', 'Ampoule', 'Seal', 'Vial', 'Vial', 'Vial', 'Vial', 'Vial', 'Ampoule', 'Ampoule', 'Seal', 'Seal', 'Seal', 'Vial', 'Seal', 'Seal', 'Seal', 'Vial', 'Vial', 'Vial', 'Vial', 'Ampoule', 'Seal', 'Seal', 'Plastic Bottles', 'Cap', 'Outer cap', 'Ampoule', 'Vial', 'Seal', 'Vial', 'Vial', 'Vial', 'Ampoule', 'Ampoule', 'Seal', 'Seal', 'Seal', 'Seal', 'Seal', 'Seal', 'Vial', 'Seal', 'Seal', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Vial', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Seal', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Seal', 'Seal', 'Seal', 'Vial', 'Ampoule', 'Seal', 'Cap', 'Plunger Stopper', 'Plunger Stopper', 'PFS Syringe', 'PFS Syringe', 'Vial', 'Ampoule', 'Seal', 'Vial', 'Seal', 'Seal', 'Ampoule', 'Ampoule', 'Ampoule', 'Seal', 'Seal', 'Seal', 'Seal', 'Seal', 'Vial', 'Seal', 'Vial', 'Cap', 'Collar', 'Ampoule', 'Ampoule', 'Ampoule', 'Seal', 'Vial', 'Seal', 'Seal', 'Seal', 'Seal', 'Seal', 'Seal', 'Seal', 'Vial', 'Vial', 'Vial', 'Ampoule', 'Ampoule', 'Seal', 'Seal', 'Seal', 'Plastic Bottles', 'Seal', 'Vial', 'Vial', 'Vial', 'Vial', 'Vial', 'Vial', 'Vial', 'Vial', 'Vial', 'Ampoule', 'Seal', 'Seal', 'Seal', 'Vial', 'Vial', 'Seal', 'Plastic Bottles', 'Cap', 'Plastic Nozzle', 'Vial', 'Vial', 'Vial', 'Vial', 'Vial', 'Vial', 'Vial', 'Vial', 'Vial', 'Vial', 'Vial', 'Vial', 'Vial', 'Vial', 'Vial', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Vial', 'Vial', 'Vial', 'Vial', 'Vial', 'Vial', 'Vial', 'Vial', 'Vial', 'Vial', 'Vial', 'Vial', 'Vial', 'Vial', 'Vial', 'Vial', 'Vial', 'Vial', 'Vial', 'Vial', 'Vial', 'Vial', 'Ampoule', 'Ampoule', 'Ampoule', 'Ampoule', 'Ampoule', 'Ampoule', 'Ampoule', 'Ampoule', 'Ampoule', 'Ampoule', 'Ampoule', 'Ampoule', 'Ampoule', 'Ampoule', 'Ampoule', 'Ampoule', 'Ampoule', 'Ampoule', 'Ampoule', 'Ampoule', 'Ampoule', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Seal', 'Seal', 'Vial', 'Vial', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Seal', 'Seal', 'Seal', 'Seal', 'Seal', 'Seal', 'Seal', 'Seal', 'Seal', 'Seal', 'Seal', 'Seal', 'Seal', 'Ampoule', 'Ampoule', 'Ampoule', 'Ampoule', 'Ampoule', 'Ampoule', 'Seal', 'Seal', 'Seal', 'Vial', 'Seal', 'Seal', 'Seal', 'Seal', 'Vial', 'Vial', 'Vial', 'Seal', 'Seal', 'Seal', 'Seal', 'Seal', 'Plastic Bottles', 'Cap', 'Outer cap', 'Seal', 'Seal', 'Seal', 'Plastic Bottles', 'Cap', 'Outer cap', 'Seal', 'Seal', 'Seal', 'Seal', 'Seal', 'Seal', 'Vial', 'Seal', 'Seal', 'Seal', 'Seal', 'Seal', 'Seal', 'Seal', 'Seal', 'Seal', 'Seal', 'Seal', 'Seal', 'Seal', 'Seal', 'Seal', 'Seal', 'Seal', 'Seal', 'Seal', 'Seal', 'Seal', 'Seal', 'Seal', 'Vial', 'Seal', 'Seal', 'Seal', 'Ampoule', 'Vial', 'Seal', 'Seal', 'Seal', 'Seal', 'Seal', 'Ampoule', 'Ampoule', 'Ampoule', 'Ampoule', 'Cap', 'Plastic Bottles', 'Seal', 'Ampoule', 'Plunger Stopper', 'Plunger Stopper', 'PFS Syringe', 'Vial', 'Ampoule', 'Ampoule', 'Ampoule', 'Rubber Stopper', 'Seal', 'Vial', 'Vial', 'Seal', 'Seal', 'Plastic Bottles', 'Cap', 'Plastic Nozzle', 'Seal', 'Rubber Stopper', 'Seal', 'Seal', 'Seal', 'Plunger Rod', 'Plunger Rod', 'Plunger Rod', 'Plunger Rod', 'Vial', 'Vial', 'Seal', 'Seal', 'Seal', 'Vial', 'Vial', 'Vial', 'Ampoule', 'Ampoule', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Seal', 'Seal', 'Seal', 'Cap', 'Cap', 'Collar', 'Plastic Bottles', 'Cap', 'Plunger Rod', 'Cap', 'U plug', 'Ampoule', 'Ampoule', 'Seal', 'Seal', 'Seal', 'Seal', 'Vial', 'Vial', 'Vial', 'Vial', 'Vial', 'Vial', 'Ampoule', 'Ampoule', 'Ampoule', 'Ampoule', 'Seal', 'Vial', 'Seal', 'Seal', 'Vial', 'Vial', 'Ampoule', 'Ampoule', 'Seal', 'Vial', 'Seal', 'Seal', 'Seal', 'Seal', 'Vial', 'Ampoule', 'Ampoule', 'Vial', 'Seal', 'Vial', 'Vial', 'Ampoule', 'Vial', 'Vial', 'Vial', 'Ampoule', 'Ampoule', 'Ampoule', 'Ampoule', 'Ampoule', 'Ampoule', 'Ampoule', 'Seal', 'Plastic Bottles', 'Cap', 'Plastic Nozzle', 'Vial', 'Vial', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Seal', 'Seal', 'Vial', 'Ampoule', 'Seal', 'Plastic Bottles', 'Cap', 'Ampoule', 'Ampoule', 'Ampoule', 'Ampoule', 'Vial', 'Ampoule', 'Seal', 'Vial', 'Vial', 'Ampoule', 'Rubber Stopper', 'Vial', 'Vial', 'Vial', 'Vial', 'Vial', 'Vial', 'Vial', 'Vial', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Vial', 'Vial', 'Vial', 'Vial', 'Vial', 'Vial', 'Vial', 'Seal', 'Seal', 'Seal', 'Seal', 'Seal', 'Seal', 'Seal', 'Vial', 'Vial', 'Vial', 'Vial', 'Ampoule', 'Ampoule', 'Ampoule', 'Ampoule', 'Ampoule', 'Seal', 'Seal', 'Seal', 'Seal', 'Seal', 'Vial', 'Seal', 'Seal', 'Seal', 'Ampoule', 'Ampoule', 'Ampoule', 'Ampoule', 'Plastic Bottles', 'Cap', 'Plastic Nozzle', 'Vial', 'Vial', 'Seal', 'Seal', 'Seal', 'Seal', 'Seal', 'Seal', 'Seal', 'Seal', 'Seal', 'Seal', 'Seal', 'Seal', 'Seal', 'Seal', 'Ampoule', 'Ampoule', 'Vial', 'Ampoule', 'Vial', 'Seal', 'Seal', 'Cap', 'Plastic Nozzle', 'Vial', 'Vial', 'Seal', 'Seal', 'Seal', 'Seal', 'Seal', 'Seal', 'Seal', 'Seal', 'Vial', 'Seal', 'Seal', 'Seal', 'Vial', 'Vial', 'Vial', 'Vial', 'Vial', 'Vial', 'Vial', 'Ampoule', 'Ampoule', 'Ampoule', 'Ampoule', 'Seal', 'Seal', 'Seal', 'Seal', 'Vial', 'Seal', 'Seal', 'Seal', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Vial', 'Vial', 'Vial', 'Ampoule', 'Ampoule', 'Ampoule', 'Ampoule', 'Rubber Stopper', 'Ampoule', 'Ampoule', 'Ampoule', 'Ampoule', 'Seal', 'Vial', 'Ampoule', 'Ampoule', 'Ampoule', 'Ampoule', 'Rubber Stopper', 'Rubber Stopper', 'Seal', 'Seal', 'Seal', 'Ampoule', 'Seal', 'Seal', 'Seal', 'Ampoule', 'Cap', 'Cap', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Seal', 'Seal', 'Seal', 'Seal', 'Vial', 'Vial', 'Vial', 'Rubber Stopper', 'Rubber Stopper', 'Vial', 'Ampoule', 'Ampoule', 'Ampoule', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Vial', 'Vial', 'Ampoule', 'Seal', 'Vial', 'Vial', 'Vial', 'Vial', 'Vial', 'Vial', 'Ampoule', 'Ampoule', 'Seal', 'Vial', 'Seal', 'Vial', 'Vial', 'Seal', 'Seal', 'Seal', 'Seal', 'Seal', 'Vial', 'Vial', 'Ampoule', 'Ampoule', 'Seal', 'Seal', 'Seal', 'Seal', 'Seal', 'Seal', 'Vial', 'Vial', 'Ampoule', 'Ampoule', 'Seal', 'Plastic Bottles', 'Cap', 'Seal', 'Vial', 'Ampoule', 'Ampoule', 'Ampoule', 'Ampoule', 'Seal', 'Seal', 'Plastic Bottles', 'Cap', 'Plastic Nozzle', 'Seal', 'Seal', 'Seal', 'Seal', 'Seal', 'Seal', 'Vial', 'Ampoule', 'Ampoule', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Plastic Bottles', 'Cap', 'Vial', 'Ampoule', 'Ampoule', 'Seal', 'Seal', 'Plastic Bottles', 'Cap', 'Cap', 'Plastic Nozzle', 'Seal', 'Seal', 'Seal', 'Seal', 'Ampoule', 'Ampoule', 'Ampoule', 'Vial', 'Al Tube', 'Al Tube', 'Al Tube', 'Al Tube', 'Al Tube', 'Al Tube', 'Ampoule', 'Ampoule', 'Ampoule', 'Ampoule', 'Ampoule', 'Ampoule', 'Seal', 'Vial', 'Ampoule', 'Ampoule', 'Seal', 'Seal', 'Seal', 'Seal', 'Seal', 'Seal', 'Seal', 'Seal', 'Seal', 'Seal', 'Seal', 'Seal', 'Seal', 'Seal', 'Seal', 'Ampoule', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Ampoule', 'Vial', 'Vial', 'Ampoule', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Seal', 'Seal', 'Seal', 'Seal', 'Seal', 'Seal', 'Seal', 'Ampoule', 'Ampoule', 'Seal', 'Seal', 'Seal', 'Seal', 'Vial', 'Vial', 'Seal', 'Seal', 'Ampoule', 'Vial', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Seal', 'Seal', 'Seal', 'Seal', 'Seal', 'Seal', 'Seal', 'Seal', 'Seal', 'Seal', 'Ampoule', 'Seal', 'Ampoule', 'Ampoule', 'Ampoule', 'Ampoule', 'Seal', 'Rubber Stopper', 'Vial', 'Ampoule', 'Seal', 'Seal', 'Ampoule', 'Vial', 'Vial', 'Vial', 'Vial', 'Vial', 'Seal', 'Seal', 'Seal', 'Ampoule', 'Ampoule', 'Ampoule', 'Seal', 'Seal', 'Plastic Bottles', 'Cap', 'Plastic Nozzle', 'Ampoule', 'Rubber Stopper', 'Vial', 'Seal', 'Rubber Stopper', 'Rubber Stopper', 'Cap', 'Collar', 'Collar', 'Vial', 'Vial', 'Vial', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Seal', 'Vial', 'Vial', 'Ampoule', 'Plastic Bottles', 'Cap', 'Vial', 'Vial', 'Rubber Stopper', 'Seal', 'Cap', 'Seal', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Vial', 'Ampoule', 'Seal', 'Vial', 'Vial', 'Vial', 'Vial', 'Vial', 'Vial', 'Vial', 'Vial', 'Vial', 'Vial', 'Seal', 'Seal', 'Seal', 'Seal', 'Seal', 'Seal', 'Ampoule', 'Ampoule', 'Ampoule', 'Plunger Stopper', 'PFS Syringe', 'PFS Syringe', 'PFS Syringe', 'Ampoule', 'Seal', 'Seal', 'Seal', 'Ampoule', 'Seal', 'Vial', 'Ampoule', 'Vial', 'Seal', 'Plastic Bottles', 'Cap', 'Plastic Nozzle', 'Plastic Nozzle', 'Vial', 'Ampoule', 'Seal', 'Seal', 'Seal', 'Seal', 'Vial', 'Ampoule', 'Vial', 'Vial', 'Vial', 'Seal', 'Vial', 'Ampoule', 'Ampoule', 'Vial', 'Seal', 'Seal', 'Ampoule', 'Ampoule', 'Ampoule', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Seal', 'Seal', 'Seal', 'Seal', 'Seal', 'Plastic Bottles', 'Plastic Bottles', 'Plastic Bottles', 'Plastic Bottles', 'Plastic Bottles', 'Plastic Bottles', 'Plastic Bottles', 'Cap', 'Cap', 'Cap', 'Cap', 'Cap', 'Cap', 'Cap', 'Vial', 'Vial', 'Seal', 'Seal', 'Seal', 'Seal', 'Vial', 'Vial', 'Vial', 'Vial', 'Ampoule', 'Seal', 'Vial', 'Vial', 'Seal', 'Seal', 'Seal', 'Seal', 'Ampoule', 'Seal', 'Seal', 'Seal', 'Vial', 'Vial', 'Vial', 'Vial', 'Vial', 'Vial', 'Vial', 'Ampoule', 'Ampoule', 'Seal', 'Seal', 'Seal', 'Seal', 'Seal', 'Vial', 'Plunger Stopper', 'Plunger Stopper', 'PFS Syringe', 'Vial', 'Vial', 'Ampoule', 'Ampoule', 'Seal', 'Seal', 'Seal', 'Seal', 'Seal', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Seal', 'Seal', 'Seal', 'Seal', 'Seal', 'Seal', 'Seal', 'Seal', 'Ampoule', 'Ampoule', 'Seal', 'Ampoule', 'Seal', 'Seal', 'Seal', 'Seal', 'Seal', 'Vial', 'Vial', 'Vial', 'Ampoule', 'Ampoule', 'Plastic Bottles', 'Cap', 'Plastic Nozzle', 'Seal', 'Vial', 'Vial', 'Seal', 'Seal', 'Seal', 'Seal', 'Seal', 'Vial', 'Seal', 'Seal', 'Seal', 'Seal', 'Vial', 'Seal', 'Seal', 'Vial', 'Vial', 'Vial', 'Vial', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Seal', 'Seal', 'Seal', 'Collar', 'Ampoule', 'Vial', 'Ampoule', 'Ampoule', 'Vial', 'Ampoule', 'Seal', 'Plastic Bottles', 'Cap', 'Plastic Nozzle', 'Seal', 'Vial', 'Vial', 'Vial', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Seal', 'Seal', 'Seal', 'Vial', 'Seal', 'Vial', 'Ampoule', 'Ampoule', 'Ampoule', 'Ampoule', 'Vial', 'Vial', 'Vial', 'Vial', 'Ampoule', 'Ampoule', 'Ampoule', 'Seal', 'Ampoule', 'Ampoule', 'Ampoule', 'Ampoule', 'Cap', 'Plunger Stopper', 'PFS Syringe', 'U plug', 'Vial', 'Ampoule', 'Vial', 'Vial', 'Vial', 'Ampoule', 'Ampoule', 'Ampoule', 'Seal', 'Seal', 'Vial', 'Vial', 'Seal', 'Seal', 'Seal', 'Cap', 'Seal', 'Vial', 'Vial', 'Vial', 'Vial', 'Vial', 'Vial', 'Vial', 'Vial', 'Seal', 'Seal', 'Seal', 'Seal', 'Seal', 'Plastic Bottles', 'Cap', 'Plastic Nozzle', 'Ampoule', 'Ampoule', 'Vial', 'Vial', 'Seal', 'Seal', 'Ampoule', 'Vial', 'Vial', 'Vial', 'Vial', 'Ampoule', 'Ampoule', 'Rubber Stopper', 'Seal', 'Vial', 'Vial', 'Vial', 'Vial', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Seal', 'Seal', 'Seal', 'Seal', 'Seal', 'Seal', 'Seal', 'Seal', 'Seal', 'Seal', 'Seal', 'Seal', 'Seal', 'Seal', 'Seal', 'Ampoule', 'Ampoule', 'Vial', 'Vial', 'Vial', 'Vial', 'Vial', 'Vial', 'Vial', 'Vial', 'Vial', 'Ampoule', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Seal', 'Seal', 'Seal', 'Ampoule', 'Ampoule', 'Plastic Bottles', 'Plastic Bottles', 'Plastic Bottles', 'Cap', 'Cap', 'Cap', 'Plastic Nozzle', 'Plastic Nozzle', 'Seal', 'Ampoule', 'Ampoule', 'Ampoule', 'Ampoule', 'Ampoule', 'Ampoule', 'Vial', 'Seal', 'Ampoule', 'Rubber Stopper', 'Seal', 'Seal', 'Vial', 'Ampoule', 'Cap', 'Ampoule', 'Rubber Stopper', 'Seal', 'Seal', 'Vial', 'Vial', 'Vial', 'Ampoule', 'Ampoule', 'Ampoule', 'Vial', 'Vial', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Seal', 'Ampoule', 'Ampoule', 'Ampoule', 'Vial', 'Plunger Rod', 'Vial', 'Vial', 'Ampoule', 'Seal', 'Ampoule', 'Seal', 'Seal', 'Vial', 'Ampoule', 'Seal', 'Plastic Bottles', 'Cap', 'Plastic Nozzle', 'Vial', 'Vial', 'Seal', 'Seal', 'Seal', 'Seal', 'Seal', 'Vial', 'Ampoule', 'Seal', 'Seal', 'Seal', 'Ampoule', 'Vial', 'Seal', 'Seal', 'Seal', 'Cap', 'Collar', 'Plastic Bottles', 'Cap', 'Plastic Nozzle', 'Cap', 'U plug', 'Collar', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Vial', 'Vial', 'Ampoule', 'Seal', 'Ampoule', 'Vial', 'Ampoule', 'Seal', 'Seal', 'Vial', 'Vial', 'Vial', 'Seal', 'Vial', 'Vial', 'Vial', 'Seal', 'Seal', 'Vial', 'Vial', 'Ampoule', 'Plunger Stopper', 'Plunger Stopper', 'PFS Syringe', 'PFS Syringe', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Plunger Rod', 'Rubber Stopper', 'Vial', 'Seal', 'Plastic Bottles', 'Cap', 'Plastic Nozzle', 'Vial', 'Vial', 'Ampoule', 'Seal', 'Vial', 'Seal', 'Vial', 'Seal', 'Seal', 'Seal', 'Vial', 'Plunger Rod', 'Rubber Stopper']})



DEFAULT_RULES = {"RM": 30, "SPM": 15, "TPM": 15}
DEFAULT_PPM_LT = 30
PPM_CATEGORY_MAP = {
    7:  ['Vial', 'Rubber Stopper', 'Rubber', 'Stopper', 'Seal', 'Cap', 'Collar', 'Inner Cap', 'Outer Cap'],
    12: ['Ampoule', 'Amp'],
    90: ['Pfs Syringe', 'Plunger Stopper', 'Plunger', 'U plug', 'U-plug'],
    15: ['Al Tube', 'Plastic Bottle', 'Plastic Nozzle', 'Nozzle'],
}

DEFAULT_UNKNOWN_LEAD_TIME = 30
custom_lead_times = {}  # optional override dict (leave empty or set before use)

# ------------------------- HELPERS -------------------------
def try_read_excel(uploaded_file):
    """
    Try read using openpyxl then xlrd. Returns DataFrame or raises informative error.
    uploaded_file: Streamlit UploadedFile (has .name and is file-like)
    """
    # rewind
    try:
        uploaded_file.seek(0)
    except Exception:
        pass

    errors = []
    for engine in ("openpyxl", "xlrd", None):
        try:
            if engine is None:
                df = pd.read_excel(uploaded_file)  # let pandas choose
            else:
                df = pd.read_excel(uploaded_file, engine=engine)
            return df
        except Exception as e:
            errors.append((engine, str(e)))
            try:
                uploaded_file.seek(0)
            except Exception:
                pass
    msg = "Failed to read Excel file. Tried engines:\n"
    for eng, err in errors:
        msg += f" - engine={eng}: {err}\n"
    raise ValueError(msg)

def standardize_column_names(df: pd.DataFrame) -> pd.DataFrame:
    """
    Map many column name variants to canonical dotted names used in logic.
    Preserves other columns.
    """
    def canon(col):
        s = str(col).strip()
        s = re.sub(r'\u00A0', ' ', s)
        s = re.sub(r'\s+', ' ', s).strip()
        k = re.sub(r'[^A-Za-z0-9]', '', s).lower()
        # map variants
        if k in ('mattype','materialtype'):
            return COL_MAT_TYPE
        if k in ('materialcode','matcode','material_code'):
            return COL_MATERIAL_CODE
        if k in ('materialname','itemname','material'):
            return COL_MATERIAL_NAME
        if k == 'uom':
            return COL_UOM
        if k in ('podt','podat','podate','podate'):
            return COL_PO_DT
        if k in ('pono','ponumber','po'):
            return COL_PO_NO
        if k in ('supplier','suppliername'):
            return COL_SUPPLIER
        if k in ('poqty','purchaseorderqty','quantityordered','quantity'):
            return COL_PO_QTY
        if k in ('gnrdt','grndt','grndate','grn','grndate'):
            return COL_GNR_DT
        if k in ('inwardqty','inwardquantity','receivedqty','receivedquantity'):
            return COL_INWARD_QTY
        if k in ('itemcategory','itemcat','category'):
            return COL_ITEM_CAT
        return col  # keep original
    df2 = df.copy()
    df2.columns = [canon(c) for c in df.columns]
    return df2

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

@st.cache_data(show_spinner=True)
def load_and_clean(file):
    # file is Streamlit UploadedFile (or file-like)
    df = try_read_excel(file)
    df = standardize_column_names(df)
    missing = [c for c in REQUIRED_COLS if c not in df.columns]
    if missing:
        raise ValueError(f"Input file missing required columns: {missing}. Columns found: {list(df.columns)}")
    # select only required columns and optional Item Category if present
    extras = [c for c in [COL_ITEM_CAT] if c in df.columns]
    df = df[REQUIRED_COLS + extras].copy()
    return df

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

def generate_failed_orders_pdf(breaches_df: pd.DataFrame, vendor_stats: pd.DataFrame, year: int) -> bytes:
    """
    Create a PDF of failed OTIF orders grouped by Supplier. Returns PDF bytes.
    """
    if not reportlab_available:
        raise RuntimeError("reportlab not available")

    buf = io.BytesIO()
    width, height = A4
    c = canvas.Canvas(buf, pagesize=A4)
    margin_x = 20 * mm
    y = height - 20 * mm
    line_height = 8 * mm

    # Title
    c.setFont("Helvetica-Bold", 16)
    c.drawString(margin_x, y, f"ALL OTIF FAILED ORDERS — {year}")
    y -= 12 * mm

    vendor_order = vendor_stats[vendor_stats["OTIF_Failures"]>0].sort_values("OTIF_Failures", ascending=False)

    vendor_idx = 1
    for _, vr in vendor_order.iterrows():
        vendor = vr[COL_SUPPLIER]
        failures = int(vr["OTIF_Failures"])
        total_orders = int(vr["Total_Orders"])
        otif_pct = float(vr["Vendor_OTIF_pct"])
        contrib_pct = float(vr["Total_Contribution_pct"])

        vendor_group = breaches_df[breaches_df[COL_SUPPLIER] == vendor].sort_values(COL_GNR_DT, ascending=False)

        # New page header if near bottom
        if y < 40 * mm:
            c.showPage()
            y = height - 20 * mm
            c.setFont("Helvetica-Bold", 16)
            c.drawString(margin_x, y, f"ALL OTIF FAILED ORDERS — {year} (cont.)")
            y -= 12 * mm

        c.setFont("Helvetica-Bold", 12)
        header = f"{vendor_idx}. {vendor}  (Failures: {failures})  OTIF: {otif_pct:.1f}%  Contribution: {contrib_pct:.1f}%  Total Orders: {total_orders}"
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
    return buf.getvalue()

# ------------------------- MAIN -------------------------
# File uploader (persistent across pages)
if not st.session_state.get("data_file_loaded"):
    uploaded = st.file_uploader("📤 Upload OTIF Excel (single sheet expected)", type=["xlsx", "xls"])
else:
    uploaded = None

# allow reusing main app-uploaded file if present
if "uploaded_file" in st.session_state and st.session_state["uploaded_file"] is not None:
    uploaded = st.session_state["uploaded_file"]

# store in session state if new upload
if uploaded is not None:
    st.session_state["uploaded_file"] = uploaded
    st.session_state["data_file_loaded"] = True

if not st.session_state.get("data_file_loaded"):
    st.info("Upload your Excel to begin. Expected columns (variants accepted): Mat Type, Material Code, Material Name, UOM, P.O. Dt., P. O. No., Supplier, PO Qty., GNR Dt., Inward Qty.")
    st.stop()

# Load and clean
try:
    df_raw = load_and_clean(st.session_state["uploaded_file"])
except Exception as e:
    st.error(f"❌ Processing error: {e}")
    st.stop()

if df_raw.empty:
    st.warning("No rows found after basic load.")
    st.stop()

# Sidebar: Mat Type filters
st.sidebar.header("🎛 Mat Type Filters & Lead Times")
all_types = sorted(df_raw[COL_MAT_TYPE].dropna().astype(str).unique().tolist())
select_all = st.sidebar.checkbox("Select ALL Mat Types", value=True)
selected_types = []

if select_all:
    selected_types = all_types
else:
    st.sidebar.caption("Tick the Mat Types you want to include:")
    for t in all_types:
        if st.sidebar.checkbox(f"{t}", value=False, key=f"mt_{t}"):
            selected_types.append(t)

if not selected_types:
    st.warning("Please select at least one Mat Type from the sidebar.")
    st.stop()

# Filter to chosen Mat Types
df = df_raw[df_raw[COL_MAT_TYPE].astype(str).isin(selected_types)].copy()
if df.empty:
    st.warning("No rows remain after Mat Type filtering.")
    st.stop()

# Convert dtypes & drop critical nulls
df = ensure_types_and_drop_nulls(df)
if df.empty:
    st.warning("No rows remain after dropping records with missing dates/quantities.")
    st.stop()

# Merge Item Category mapping
df = merge_item_category(df)

# Lead time rules: defaults + UI inputs for unknown selected types (except PPM)
lead_time_rules = DEFAULT_RULES.copy()
unknown_selected = [t for t in selected_types if t.upper() not in lead_time_rules and t.upper() != "PPM"]

if unknown_selected:
    st.sidebar.subheader("⏱ Lead Time (days) for other Mat Types")
for t in unknown_selected:
    lead_time_rules[t.upper()] = st.sidebar.number_input(
        f"Lead Time for {t}", min_value=1, max_value=365, value=30, step=1, key=f"lt_{t}"
    )

# compute Lead Time
df["Lead Time"] = df.apply(lambda row: compute_lead_time_for_row(row, lead_time_rules), axis=1)
if df["Lead Time"].isna().any():
    bad = df.loc[df["Lead Time"].isna(), COL_MAT_TYPE].unique().tolist()
    st.error(f"Lead Time missing for Mat Types: {bad}. Please provide values in the sidebar.")
    st.stop()

# PO-level metrics
df_line, df_po = compute_po_level_metrics(df)
if df_po.empty:
    st.warning("No P.O. rows available after processing.")
    st.stop()

# Year selection (based on last GRN date per PO)
years = sorted(df_po["Year"].dropna().unique().astype(int).tolist())
if not years:
    st.error("No valid 'GNR Dt.' years found after processing.")
    st.stop()

selected_year = st.selectbox("📅 Select Year", years, index=len(years)-1)
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

overall_yearly = po_year["OTIF"].mean()

# KPIs
k1, k2, k3, k4 = st.columns(4)
k1.metric("OTIF (Yearly PO-level mean)", f"{(overall_yearly*100):.1f}%")
k2.metric("On-Time (Yearly mean)", f"{(po_year['OnTime'].mean()*100):.1f}%")
k3.metric("In-Full (Yearly mean)", f"{(po_year['PO_Fulfilled'].mean()*100):.1f}%")
k4.metric("Total Orders (Year)", int(po_year.shape[0]))

# Monthly chart
st.subheader("📊 Monthly OTIF (Selected Year)")
if monthly.empty:
    st.info("No orders for the selected year.")
else:
    month_order = ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"]
    present = [m for m in month_order if m in monthly["Month"].tolist()]
    fig = px.bar(
        monthly,
        x="Month",
        y="Avg_OTIF",
        category_orders={"Month": present},
        text=monthly["Avg_OTIF"].map(lambda v: f"{v*100:.1f}%"),
        labels={"Month": "Month", "Avg_OTIF": "Average OTIF"},
        height=420
    )
    fig.update_traces(textposition="outside")
    fig.update_yaxes(range=[0, 1], tickformat=".0%")
    fig.update_layout(margin=dict(l=20, r=20, t=40, b=20))
    st.plotly_chart(fig, use_container_width=True)

# Monthly table
with st.expander("📄 Monthly Summary Table"):
    tbl = monthly.copy()
    for c in ["Avg_OTIF", "Avg_OnTime", "Avg_InFull"]:
        tbl[c] = (tbl[c] * 100).round(1)
    tbl = tbl.rename(columns={
        "MonthNum": "Month #",
        "Avg_OTIF": "Avg OTIF (%)",
        "Avg_OnTime": "Avg On-Time (%)",
        "Avg_InFull": "Avg In-Full (%)",
    })
    st.dataframe(tbl, use_container_width=True)

# Top 10 Vendors with breaches
st.subheader("🚨 Top 10 Vendors with OTIF Breaches (Selected Year)")
breaches = po_year[po_year["OTIF"] == 0].copy()

# Vendor stats for the year
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

if breaches.empty:
    st.success("No OTIF breaches in the selected year. 🎉")
    top10 = pd.DataFrame(columns=[COL_SUPPLIER, "OTIF_Failures", "Vendor_OTIF_pct", "Total_Contribution_pct", "Total_Orders"])
    st.dataframe(top10, use_container_width=True)
else:
    top10 = vendor_stats[vendor_stats["OTIF_Failures"]>0].sort_values("OTIF_Failures", ascending=False).head(10)
    display_top10 = top10[[COL_SUPPLIER, "OTIF_Failures", "Vendor_OTIF_pct", "Total_Contribution_pct", "Total_Orders"]].copy()
    display_top10["Vendor_OTIF_pct"] = display_top10["Vendor_OTIF_pct"].map(lambda v: f"{v:.1f}%")
    display_top10["Total_Contribution_pct"] = display_top10["Total_Contribution_pct"].map(lambda v: f"{v:.1f}%")
    st.dataframe(display_top10.reset_index(drop=True), use_container_width=True)

    # PDF download of failed orders grouped by supplier (include percentages in headings)
    if reportlab_available:
        try:
            pdf_bytes = generate_failed_orders_pdf(breaches[[COL_SUPPLIER, COL_GNR_DT, COL_PO_NO]], vendor_stats, selected_year)
            st.download_button(
                "⬇️ Download ALL Failed Orders (PDF)",
                data=pdf_bytes,
                file_name=f"OTIF_failed_orders_{selected_year}.pdf",
                mime="application/pdf"
            )
        except Exception as e:
            st.error(f"Error generating PDF: {e}")
            csv_bytes = breaches[[COL_SUPPLIER, COL_GNR_DT, COL_PO_NO]].sort_values([COL_SUPPLIER, COL_GNR_DT], ascending=[False, False]).to_csv(index=False).encode("utf-8")
            st.download_button(
                "⬇️ Download ALL Failed Orders (CSV fallback)",
                data=csv_bytes,
                file_name=f"OTIF_failed_orders_{selected_year}.csv",
                mime="text/csv"
            )
    else:
        st.warning("PDF export requires the `reportlab` package. Install it in your environment (requirements.txt).")
        csv_bytes = breaches[[COL_SUPPLIER, COL_GNR_DT, COL_PO_NO]].sort_values([COL_SUPPLIER, COL_GNR_DT], ascending=[False, False]).to_csv(index=False).encode("utf-8")
        st.download_button(
            "⬇️ Download ALL Failed Orders (CSV)",
            data=csv_bytes,
            file_name=f"OTIF_failed_orders_{selected_year}.csv",
            mime="text/csv"
        )

# Download PO-level & monthly data
col1, col2 = st.columns([1.5, 1])
with col1:
    st.download_button(
        "⬇️ Download PO-level Data (CSV)",
        data=po_year.to_csv(index=False).encode("utf-8"),
        file_name=f"po_level_{selected_year}.csv",
        mime="text/csv",
    )
with col2:
    st.download_button(
        "⬇️ Download Monthly Summary (CSV)",
        data=monthly.to_csv(index=False).encode("utf-8"),
        file_name=f"monthly_otif_{selected_year}.csv",
        mime="text/csv",
    )

with st.expander("🔎 Debug / Sanity"):
    st.write("Selected Mat Types:", selected_types)
    st.write("Lead Time Rules (non-PPM):", lead_time_rules)
    st.write("Number of POs in year:", int(po_year.shape[0]))
    st.write("Overall yearly (PO-level OTIF):", overall_yearly)

