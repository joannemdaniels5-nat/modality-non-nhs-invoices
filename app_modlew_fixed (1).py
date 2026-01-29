# =============================================================
# Modality Lewisham – Private Services Invoice System (Non-NHS)
# FIXED VERSION:
# - Prices reliably pulled from Excel
# - Re-download existing invoices from Manage tab
# =============================================================

from pathlib import Path
from datetime import datetime, date
import pandas as pd
import streamlit as st
import altair as alt

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

def norm_cols(df):
    df = df.copy()
    df.columns = df.columns.astype(str).str.strip().str.lower().str.replace(" ", "_")
    return df

def load_gp_list():
    if not GP_PATH.exists():
        return []
    df = norm_cols(pd.read_excel(GP_PATH))
    if "role" in df.columns:
        df = df[df["role"].str.upper() == "GP"]
    if "cliniciandisplay" in df.columns:
        return sorted(df["cliniciandisplay"].dropna().astype(str).tolist())
    if {"first_name","surname"}.issubset(df.columns):
        return sorted((df["first_name"].astype(str)+" "+df["surname"].astype(str)).tolist())
    return sorted(df.iloc[:,0].astype(str).tolist())

def load_price_list():
    if not PRICE_PATH.exists():
        return pd.DataFrame(columns=["Service","Price"])
    df = norm_cols(pd.read_excel(PRICE_PATH))
    service_col = next((c for c in df.columns if "service" in c), None)
    price_col = next((c for c in df.columns if "price" in c), None)
    if not service_col or not price_col:
        return pd.DataFrame(columns=["Service","Price"])
    out = df[[service_col, price_col]].copy()
    out.columns = ["Service","Price"]
    out["Service"] = out["Service"].astype(str).str.strip()
    out["Price"] = pd.to_numeric(out["Price"], errors="coerce").fillna(0.0)
    return out[out["Service"] != ""]

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
    return f"<html><body><h2>Modality Lewisham – Private Services</h2><p>Invoice: {inv}</p><p>Clinician: {clinician}</p><p>Patient: {patient}</p><p>Service: {service}</p><p>Total: &pound;{amount:.2f}</p></body></html>"

gps = load_gp_list()
prices = load_price_list()
tracker = load_tracker()

tab1, tab2, tab3 = st.tabs(["Create Invoice","Manage Invoices","Tracker & Reports"])

with tab1:
    clinician = st.selectbox("Clinician", [""] + gps)
    patient = st.text_input("Patient name")
    service = st.selectbox("Service", [""] + prices["Service"].tolist())
    default_price = float(prices.loc[prices["Service"]==service,"Price"].iloc[0]) if service in prices["Service"].values else 0.0
    amount = st.number_input("Amount (&pound;)", value=default_price, format="%.2f")

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
    st.subheader("Re-download invoice")
    all_inv = pd.concat(tracker.values())
    if not all_inv.empty:
        inv_sel = st.selectbox("Invoice", all_inv["Invoice No"])
        row = all_inv[all_inv["Invoice No"]==inv_sel].iloc[0]
        file_path = Path(row["Invoice File"])
        if file_path.exists():
            html = file_path.read_text(encoding="utf-8")
            st.download_button("Download existing invoice", html.encode("utf-8"), file_name=file_path.name)

with tab3:
    st.dataframe(pd.concat(tracker.values(), ignore_index=True))
