# =============================================================
# Modality Lewisham – Private Services Invoice System (Non-NHS)
# STABLE VERSION:
# - Fixes pyarrow dataframe crash (arrow_safe)
# - Restores clinician dropdown from GP Excel
# - Full Manage Invoices functionality
# =============================================================

from pathlib import Path
from datetime import datetime, date
import pandas as pd
import streamlit as st

st.set_page_config(page_title="Modality Lewisham – Private Services", layout="wide")

APP_PASSWORD = "Modality2026!"

BASE_DIR = Path(__file__).parent
DATA_DIR = BASE_DIR / "data"
TRACKER_DIR = DATA_DIR / "Trackers"
INVOICE_DIR = DATA_DIR / "Invoices"

GP_PATH = DATA_DIR / "ModalityLewGP_List.xlsx"
PRICE_PATH = DATA_DIR / "Private Services Price List.xlsx"
TRACKER_PATH = TRACKER_DIR / "Private Services Tracker.xlsx"

for d in (DATA_DIR, TRACKER_DIR, INVOICE_DIR):
    d.mkdir(parents=True, exist_ok=True)

def password_gate():
    if st.session_state.get("auth_ok"):
        return
    st.title("🔐 Modality Lewisham – Private Services (Non-NHS)")
    pw = st.text_input("Password", type="password")
    if st.button("Login"):
        if pw == APP_PASSWORD:
            st.session_state["auth_ok"] = True
            st.rerun()
        else:
            st.error("Incorrect password")
            st.stop()
    else:
        st.stop()

password_gate()

def arrow_safe(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    for col in df.columns:
        if df[col].dtype == "object":
            df[col] = df[col].astype(str)
        if "date" in col.lower():
            df[col] = pd.to_datetime(df[col], errors="coerce").astype(str)
    return df.fillna("")

def norm_cols(df):
    df = df.copy()
    df.columns = df.columns.astype(str).str.strip().str.lower().str.replace(" ", "_")
    return df

def load_gp_list():
    if not GP_PATH.exists():
        st.warning("GP list missing: data/ModalityLewGP_List.xlsx")
        return []
    df = norm_cols(pd.read_excel(GP_PATH))
    if "role" in df.columns:
        df = df[df["role"].str.upper() == "GP"]
    if "cliniciandisplay" in df.columns:
        names = df["cliniciandisplay"]
    elif {"first_name","surname"}.issubset(df.columns):
        names = df["first_name"].astype(str) + " " + df["surname"].astype(str)
    else:
        names = df.iloc[:,0].astype(str)
    return sorted(names.dropna().astype(str).unique().tolist())

def load_price_list():
    if not PRICE_PATH.exists():
        return pd.DataFrame(columns=["Service","Price"])
    df = norm_cols(pd.read_excel(PRICE_PATH))
    svc = next((c for c in df.columns if "service" in c), None)
    prc = next((c for c in df.columns if "price" in c), None)
    if not svc or not prc:
        return pd.DataFrame(columns=["Service","Price"])
    out_df = df[[svc, prc]].copy()
    out_df.columns = ["Service","Price"]
    out_df["Service"] = out_df["Service"].astype(str).str.strip()
    out_df["Price"] = pd.to_numeric(out_df["Price"], errors="coerce").fillna(0.0)
    return out_df[out_df["Service"] != ""]

TRACKER_COLS = [
    "Invoice No","Clinician","Patient","Service","Amount","Status",
    "Invoice File","Date","Paid Date","Payment Method",
    "Payment Reversed Date","Cancelled Date","Cancellation Reason"
]

def ensure_cols(df):
    for c in TRACKER_COLS:
        if c not in df.columns:
            df[c] = None
    return df[TRACKER_COLS]

def load_tracker():
    if TRACKER_PATH.exists():
        sheets = pd.read_excel(TRACKER_PATH, sheet_name=None)
    else:
        sheets = {}
    return {
        "Sent": ensure_cols(sheets.get("Sent", pd.DataFrame(columns=TRACKER_COLS))),
        "Paid": ensure_cols(sheets.get("Paid", pd.DataFrame(columns=TRACKER_COLS))),
        "Cancelled": ensure_cols(sheets.get("Cancelled", pd.DataFrame(columns=TRACKER_COLS))),
    }

def save_tracker(t):
    with pd.ExcelWriter(TRACKER_PATH, engine="openpyxl") as w:
        for k,v in t.items():
            v.to_excel(w, sheet_name=k, index=False)

def build_invoice_html(inv, clinician, patient, service, amount):
    return f"""
<html><body style='font-family:Arial'>
<h2>Modality Lewisham – Private Services</h2>
<p><b>Invoice:</b> {inv}</p>
<p><b>Date:</b> {date.today().strftime('%d/%m/%Y')}</p>
<p><b>Clinician:</b> {clinician}</p>
<p><b>Patient:</b> {patient}</p>
<p><b>Service:</b> {service}</p>
<p><b>Total:</b> £{amount:.2f}</p>
</body></html>
"""

gps = load_gp_list()
prices = load_price_list()
tracker = load_tracker()

tab1, tab2, tab3 = st.tabs(["Create Invoice","Manage Invoices","Tracker"])

with tab1:
    clinician = st.selectbox("Clinician", [""] + gps)
    patient = st.text_input("Patient name")
    service = st.selectbox("Service", [""] + prices["Service"].tolist())
    default_price = float(prices.loc[prices["Service"]==service,"Price"].iloc[0]) if service in prices["Service"].values else 0.0
    amount = st.number_input("Amount (£)", value=default_price, format="%.2f")

    if st.button("Generate Invoice"):
        inv = f"INV-{datetime.now().strftime('%Y%m%d-%H%M%S')}"
        html = build_invoice_html(inv, clinician, patient, service, amount)
        year_dir = INVOICE_DIR / str(datetime.now().year)
        year_dir.mkdir(exist_ok=True)
        path = year_dir / f"{inv}.html"
        path.write_text(html, encoding="utf-8")

        tracker["Sent"] = pd.concat([tracker["Sent"], pd.DataFrame([{
            "Invoice No": inv,
            "Clinician": clinician,
            "Patient": patient,
            "Service": service,
            "Amount": amount,
            "Status": "Sent",
            "Invoice File": str(path),
            "Date": date.today()
        }])], ignore_index=True)
        save_tracker(tracker)
        st.download_button("Download invoice", html.encode("utf-8"), file_name=f"{inv}.html")

with tab2:
    st.subheader("Sent invoices")
    st.dataframe(arrow_safe(tracker["Sent"]))

    st.subheader("Mark as Paid")
    if not tracker["Sent"].empty:
        inv = st.selectbox("Invoice", tracker["Sent"]["Invoice No"])
        method = st.selectbox("Payment method", ["Card","Bank transfer","Cash","Cheque","Other"])
        if st.button("Confirm Paid"):
            row = tracker["Sent"][tracker["Sent"]["Invoice No"]==inv].iloc[0].copy()
            row["Status"] = "Paid"
            row["Paid Date"] = date.today()
            row["Payment Method"] = method
            tracker["Paid"] = pd.concat([tracker["Paid"], pd.DataFrame([row])], ignore_index=True)
            tracker["Sent"] = tracker["Sent"][tracker["Sent"]["Invoice No"]!=inv]
            save_tracker(tracker)
            st.success("Marked as paid")
            st.rerun()

    st.subheader("Cancel invoice")
    active = pd.concat([tracker["Sent"], tracker["Paid"]])
    if not active.empty:
        cancel_inv = st.selectbox("Invoice", active["Invoice No"], key="cancel_inv")
        reason = st.text_input("Reason")
        if st.button("Cancel"):
            row = active[active["Invoice No"]==cancel_inv].iloc[0].copy()
            row["Status"] = "Cancelled"
            row["Cancelled Date"] = date.today()
            row["Cancellation Reason"] = reason
            tracker["Sent"] = tracker["Sent"][tracker["Sent"]["Invoice No"]!=cancel_inv]
            tracker["Paid"] = tracker["Paid"][tracker["Paid"]["Invoice No"]!=cancel_inv]
            tracker["Cancelled"] = pd.concat([tracker["Cancelled"], pd.DataFrame([row])], ignore_index=True)
            save_tracker(tracker)
            st.warning("Invoice cancelled")
            st.rerun()

    st.subheader("Re-download invoice")
    all_inv = pd.concat(tracker.values())
    if not all_inv.empty:
        inv_sel = st.selectbox("Invoice", all_inv["Invoice No"], key="dl_inv")
        row = all_inv[all_inv["Invoice No"]==inv_sel].iloc[0]
        path = Path(row["Invoice File"])
        if path.exists():
            html = path.read_text(encoding="utf-8")
            st.download_button("Download existing invoice", html.encode("utf-8"), file_name=path.name)

with tab3:
    st.dataframe(arrow_safe(pd.concat(tracker.values(), ignore_index=True)))
