# Modality Lewisham – Private Services Invoice System
# FINAL BUNDLED VERSION

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

for d in [DATA_DIR, TRACKER_DIR, INVOICE_DIR]:
    d.mkdir(exist_ok=True)

def password_gate():
    if st.session_state.get("auth_ok"):
        return
    st.title("🔐 Modality Lewisham – Private Services")
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

def load_gps():
    if not GP_PATH.exists():
        return []
    df = pd.read_excel(GP_PATH)
    df.columns = df.columns.str.lower().str.replace(" ","_")
    if "role" in df.columns:
        df = df[df["role"].str.upper()=="GP"]
    if "cliniciandisplay" in df.columns:
        return sorted(df["cliniciandisplay"].dropna().unique())
    return sorted(df.iloc[:,0].dropna().unique())

def load_prices():
    return pd.read_excel(PRICE_PATH)

def load_tracker():
    cols = ["Invoice No","Clinician","Patient","Service","Amount","Status",
            "Invoice File","Date","Paid Date","Payment Method",
            "Payment Reversed Date","Cancelled Date","Cancellation Reason"]
    if TRACKER_PATH.exists():
        sheets = pd.read_excel(TRACKER_PATH, sheet_name=None)
    else:
        sheets = {}
    for s in ["Sent","Paid"]:
        if s not in sheets:
            sheets[s] = pd.DataFrame(columns=cols)
        for c in cols:
            if c not in sheets[s]:
                sheets[s][c] = None
    return sheets

def save_tracker(t):
    with pd.ExcelWriter(TRACKER_PATH, engine="openpyxl") as w:
        for k,v in t.items():
            v.to_excel(w, sheet_name=k, index=False)

def invoice_html(inv, clinician, patient, service, amount):
    return f"""<html><body><h2>Modality Lewisham</h2>
    <p>Invoice {inv}</p>
    <p>Clinician: {clinician}</p>
    <p>Patient: {patient}</p>
    <p>Service: {service}</p>
    <p>Amount: £{amount:.2f}</p>
    </body></html>"""

tracker = load_tracker()
gps = load_gps()
prices = load_prices()

tab1, tab2 = st.tabs(["Create / Manage","Reports"])

with tab1:
    clinician = st.selectbox("Clinician", [""]+list(gps))
    patient = st.text_input("Patient")
    service = st.selectbox("Service", [""]+prices["Service"].tolist())
    price = float(prices.loc[prices["Service"]==service,"Price"].iloc[0]) if service else 0
    amount = st.number_input("Amount", value=price)
    if st.button("Generate Invoice"):
        inv = f"INV-{datetime.now().strftime('%Y%m%d%H%M%S')}"
        html = invoice_html(inv,clinician,patient,service,amount)
        year = str(datetime.now().year)
        out = INVOICE_DIR/year
        out.mkdir(exist_ok=True)
        path = out/f"{inv}.html"
        path.write_text(html)
        tracker["Sent"] = pd.concat([tracker["Sent"],pd.DataFrame([{
            "Invoice No":inv,"Clinician":clinician,"Patient":patient,
            "Service":service,"Amount":amount,"Status":"Sent",
            "Invoice File":str(path),"Date":date.today()
        }])])
        save_tracker(tracker)
        st.download_button("Download Invoice (HTML)", html, file_name=f"{inv}.html")

with tab2:
    if not tracker["Paid"].empty:
        df = tracker["Paid"].copy()
        df["Amount"]=pd.to_numeric(df["Amount"],errors="coerce").fillna(0)
        st.dataframe(df)
