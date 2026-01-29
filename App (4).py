# =============================================================
# Modality Lewisham – Private Services Invoice System (Non‑NHS)
# Streamlit app.py – 3 tabs: Create | Manage | Tracker & Reports
# =============================================================

from __future__ import annotations

from pathlib import Path
from datetime import datetime, date
from io import BytesIO
import base64

import pandas as pd
import streamlit as st
import altair as alt

# ----------------------- Page config ------------------------

st.set_page_config(
    page_title="Modality Lewisham – Private Services",
    layout="wide",
    page_icon="💷",
)

APP_PASSWORD = "Modality2026!"

BASE_DIR = Path(__file__).parent
DATA_DIR = BASE_DIR / "data"
TRACKER_DIR = DATA_DIR / "Trackers"
INVOICE_DIR = DATA_DIR / "Invoices"

GP_PATH = DATA_DIR / "ModalityLewGP_List.xlsx"
PRICE_PATH = DATA_DIR / "Private Services Price List.xlsx"
TRACKER_PATH = TRACKER_DIR / "Private Services Tracker.xlsx"

LOGO_CANDIDATES = [
    BASE_DIR / "logo.png",
    DATA_DIR / "logo.png",
]

for d in (DATA_DIR, TRACKER_DIR, INVOICE_DIR):
    d.mkdir(parents=True, exist_ok=True)

# ----------------------- Password gate ----------------------

def password_gate() -> None:
    if st.session_state.get("auth_ok", False):
        return

    st.title("🔐 Modality Lewisham – Private Services (Non‑NHS)")
    st.write("Internal system. Please enter the password to continue.")

    pw = st.text_input("Password", type="password")
    login = st.button("Login")

    if not login:
        st.stop()

    if pw == APP_PASSWORD:
        st.session_state["auth_ok"] = True
        st.rerun()
    else:
        st.error("Incorrect password.")
        st.stop()

password_gate()

# ----------------------- Helpers ----------------------------

def _norm_cols(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df.columns = (
        df.columns.astype(str)
        .str.strip()
        .str.lower()
        .str.replace(" ", "_")
        .str.replace("-", "_")
    )
    return df

def load_gp_list() -> list[str]:
    if not GP_PATH.exists():
        st.warning("GP list not found. Expected: data/ModalityLewGP_List.xlsx")
        return []

    try:
        df = pd.read_excel(GP_PATH)
    except Exception as e:
        st.warning(f"Could not read GP list: {e}")
        return []

    df = _norm_cols(df)

    if "role" in df.columns:
        df["role"] = df["role"].astype(str)
        df = df[df["role"].str.upper() == "GP"]

    if "cliniciandisplay" in df.columns:
        names = df["cliniciandisplay"].astype(str).str.strip()
    elif {"first_name", "surname"}.issubset(df.columns):
        names = df["first_name"].astype(str).str.strip() + " " + df["surname"].astype(str).str.strip()
    elif "name" in df.columns:
        names = df["name"].astype(str).str.strip()
    else:
        names = df.iloc[:, 0].astype(str).str.strip()

    names = names.replace("nan", "").dropna()
    names = names[names != ""].drop_duplicates().sort_values()
    return names.tolist()

def load_price_list() -> pd.DataFrame:
    if not PRICE_PATH.exists():
        df = pd.DataFrame({"Service": ["Private Letter"], "Price": [30.0]})
        df.to_excel(PRICE_PATH, index=False)

    df = pd.read_excel(PRICE_PATH)
    df = df.copy()

    if "Service" not in df.columns or "Price" not in df.columns:
        df.columns = ["Service", "Price"] + list(df.columns[2:])
        df = df[["Service", "Price"]]

    df["Service"] = df["Service"].astype(str).str.strip()
    df["Price"] = pd.to_numeric(df["Price"], errors="coerce").fillna(0.0)
    df = df[df["Service"] != ""]
    return df[["Service", "Price"]].drop_duplicates()

def get_logo_img_tag() -> str:
    logo_path = next((p for p in LOGO_CANDIDATES if p.exists()), None)
    if not logo_path:
        return "<div style='font-size:18px;font-weight:700;'>Modality Lewisham</div>"
    try:
        b64 = base64.b64encode(logo_path.read_bytes()).decode("ascii")
        return "<img src='data:image/png;base64,%s' style='height:60px; display:block; margin:0 auto;'>" % b64
    except Exception:
        return "<div style='font-size:18px;font-weight:700;'>Modality Lewisham</div>"

def build_invoice_html(
    invoice_no: str,
    invoice_date: date,
    clinician: str,
    patient: str,
    patient_dob: date | None,
    patient_address: str,
    service: str,
    amount: float,
    include_vat_note: bool = True,
) -> str:
    logo_html = get_logo_img_tag()
    dob_str = patient_dob.strftime("%d/%m/%Y") if patient_dob else ""
    address_html = patient_address.strip().replace("\n", "<br>") if patient_address else ""

    vat_line = "<p style='margin:0.25rem 0;'><i>VAT included where applicable.</i></p>" if include_vat_note else ""

    # Use &pound; to avoid any encoding surprises
    return f"""
<html>
<head>
  <meta charset='utf-8'>
  <title>{invoice_no}</title>
</head>
<body style='font-family: Arial, sans-serif; margin: 40px;'>
  <div style='text-align:center; margin-bottom: 12px;'>
    {logo_html}
    <div style='font-size:18px; font-weight:700; margin-top:8px;'>Private Services Invoice</div>
  </div>
  <hr>
  <table style='width:100%; border-collapse:collapse;'>
    <tr>
      <td style='width:50%; vertical-align:top;'>
        <p style='margin:0.25rem 0;'><b>Invoice No:</b> {invoice_no}</p>
        <p style='margin:0.25rem 0;'><b>Invoice Date:</b> {invoice_date.strftime('%d/%m/%Y')}</p>
        <p style='margin:0.25rem 0;'><b>Clinician:</b> {clinician}</p>
      </td>
      <td style='width:50%; vertical-align:top;'>
        <p style='margin:0.25rem 0;'><b>Patient:</b> {patient}</p>
        {("<p style='margin:0.25rem 0;'><b>DOB:</b> " + dob_str + "</p>") if dob_str else ""}
        {("<p style='margin:0.25rem 0;'><b>Address:</b><br>" + address_html + "</p>") if address_html else ""}
      </td>
    </tr>
  </table>
  <hr>
  <p style='margin:0.25rem 0;'><b>Service:</b> {service}</p>
  <p style='margin:0.25rem 0;'><b>Total:</b> &pound;{amount:.2f}</p>
  {vat_line}
  <hr>
  <p style='margin-top:18px;'>Please arrange payment to Modality Lewisham in line with local private services process.</p>
</body>
</html>
""".strip()

TRACKER_COLS = [
    "Invoice No","Clinician","Patient","Service","Amount","Status",
    "Invoice File","Date","Paid Date","Payment Method",
    "Payment Reversed Date","Cancelled Date","Cancellation Reason"
]

def ensure_tracker_columns(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    for c in TRACKER_COLS:
        if c not in df.columns:
            df[c] = None
    return df[TRACKER_COLS]

def load_tracker() -> dict[str, pd.DataFrame]:
    if TRACKER_PATH.exists():
        try:
            sheets = pd.read_excel(TRACKER_PATH, sheet_name=None)
        except Exception:
            sheets = {}
    else:
        sheets = {}

    out = {}
    for name in ["Sent","Paid","Cancelled"]:
        out[name] = ensure_tracker_columns(sheets.get(name, pd.DataFrame(columns=TRACKER_COLS)))
    return out

def save_tracker(t: dict[str, pd.DataFrame]) -> None:
    with pd.ExcelWriter(TRACKER_PATH, engine="openpyxl") as w:
        for sheet in ["Sent","Paid","Cancelled"]:
            t[sheet].to_excel(w, sheet_name=sheet, index=False)

def df_to_excel_bytes(sheets: dict[str, pd.DataFrame]) -> bytes:
    bio = BytesIO()
    with pd.ExcelWriter(bio, engine="openpyxl") as w:
        for name, df in sheets.items():
            df.to_excel(w, sheet_name=name[:31], index=False)
    bio.seek(0)
    return bio.read()

# ----------------------- Load data --------------------------

tracker = load_tracker()
gp_names = load_gp_list()
price_list = load_price_list()

st.title("💷 Modality Lewisham – Private Services (Non‑NHS)")
st.caption("Cancelled invoices remain visible for audit, but are excluded from revenue and outstanding.")

tab_create, tab_manage, tab_reports = st.tabs(["➕ Create Invoice","🧾 Manage Invoices","📊 Tracker & Reports"])

# ----------------------- TAB 1: Create ----------------------

def reset_invoice_form():
    for k in ["clinician","patient","include_dob","dob","address","service","amount","last_invoice"]:
        st.session_state.pop(k, None)

with tab_create:
    st.subheader("Create a new invoice")

    c1, c2 = st.columns(2)
    with c1:
        clinician = st.selectbox("Clinician", [""] + gp_names, key="clinician")
        patient = st.text_input("Patient full name", key="patient")
        include_dob = st.checkbox("Include patient DOB?", value=False, key="include_dob")
        patient_dob = st.date_input("Patient DOB", key="dob") if include_dob else None
    with c2:
        patient_address = st.text_area("Patient address (optional)", height=110, key="address")

    service_options = price_list["Service"].tolist()
    service = st.selectbox("Service", [""] + service_options, key="service")

    default_price = 0.0
    if service:
        try:
            default_price = float(price_list.loc[price_list["Service"] == service, "Price"].iloc[0])
        except Exception:
            default_price = 0.0

    amount = st.number_input("Total amount (&pound;)", min_value=0.0, step=1.0, value=float(st.session_state.get("amount", default_price)), format="%.2f", key="amount")

    include_vat = st.checkbox("Include VAT note", value=True)

    if st.button("Generate invoice", type="primary"):
        if not clinician or not patient or not service:
            st.error("Clinician, Patient and Service are required.")
        elif amount <= 0:
            st.error("Amount must be greater than zero.")
        else:
            now = datetime.now()
            invoice_no = f"INV-{now.strftime('%Y%m%d-%H%M%S')}"
            html = build_invoice_html(
                invoice_no=invoice_no,
                invoice_date=now.date(),
                clinician=clinician,
                patient=patient,
                patient_dob=patient_dob,
                patient_address=patient_address or "",
                service=service,
                amount=float(amount),
                include_vat_note=include_vat,
            )

            year_dir = INVOICE_DIR / str(now.year) / now.strftime("%B")
            year_dir.mkdir(parents=True, exist_ok=True)
            inv_path = year_dir / f"{invoice_no}.html"
            inv_path.write_text(html, encoding="utf-8")

            new_row = {
                "Invoice No": invoice_no,
                "Clinician": clinician,
                "Patient": patient,
                "Service": service,
                "Amount": float(amount),
                "Status": "Sent",
                "Invoice File": str(inv_path),
                "Date": now.date(),
                "Paid Date": None,
                "Payment Method": None,
                "Payment Reversed Date": None,
                "Cancelled Date": None,
                "Cancellation Reason": None,
            }

            tracker["Sent"] = pd.concat([tracker["Sent"], pd.DataFrame([new_row])], ignore_index=True)
            save_tracker(tracker)

            st.success(f"Invoice created: {invoice_no}")
            st.download_button("⬇️ Download invoice (HTML → Print to PDF)", data=html.encode("utf-8"), file_name=f"{invoice_no}.html", mime="text/html")
            st.info("Open the HTML in a browser → Ctrl+P → Save as PDF → email/send.")
            st.session_state["last_invoice"] = invoice_no

    if st.session_state.get("last_invoice"):
        st.divider()
        if st.button("➕ New invoice"):
            reset_invoice_form()
            st.rerun()

# ----------------------- TAB 2: Manage ----------------------

with tab_manage:
    st.subheader("Manage invoices")

    l, r = st.columns(2)
    with l:
        st.markdown("### Sent")
        st.dataframe(tracker["Sent"], use_container_width=True, height=260)
    with r:
        st.markdown("### Paid")
        st.dataframe(tracker["Paid"], use_container_width=True, height=260)

    st.divider()

    st.markdown("## Mark as Paid")
    if tracker["Sent"].empty:
        st.info("No sent invoices to mark as paid.")
    else:
        c1, c2, c3 = st.columns(3)
        with c1:
            inv_to_pay = st.selectbox("Invoice (Sent)", tracker["Sent"]["Invoice No"].tolist(), key="pay_inv")
        with c2:
            paid_date = st.date_input("Paid date", value=date.today(), key="pay_date")
        with c3:
            payment_method = st.selectbox("Payment method", ["Card","Bank transfer","Cash","Cheque","Other"], key="pay_method")

        if st.button("Confirm paid"):
            row = tracker["Sent"].loc[tracker["Sent"]["Invoice No"] == inv_to_pay].iloc[0].copy()
            row["Status"] = "Paid"
            row["Paid Date"] = paid_date
            row["Payment Method"] = payment_method
            row["Payment Reversed Date"] = None

            tracker["Paid"] = pd.concat([tracker["Paid"], pd.DataFrame([row])], ignore_index=True)
            tracker["Sent"] = tracker["Sent"].loc[tracker["Sent"]["Invoice No"] != inv_to_pay].reset_index(drop=True)
            save_tracker(tracker)
            st.success(f"{inv_to_pay} marked Paid ({payment_method}).")
            st.rerun()

    st.divider()

    st.markdown("## Reverse a payment (Paid → Sent)")
    if tracker["Paid"].empty:
        st.info("No paid invoices to reverse.")
    else:
        c1, c2 = st.columns(2)
        with c1:
            inv_to_reverse = st.selectbox("Invoice (Paid)", tracker["Paid"]["Invoice No"].tolist(), key="rev_inv")
        with c2:
            reversed_date = st.date_input("Reversal date", value=date.today(), key="rev_date")

        if st.button("Reverse payment"):
            row = tracker["Paid"].loc[tracker["Paid"]["Invoice No"] == inv_to_reverse].iloc[0].copy()
            row["Status"] = "Sent"
            row["Payment Reversed Date"] = reversed_date
            row["Paid Date"] = None
            row["Payment Method"] = None

            tracker["Sent"] = pd.concat([tracker["Sent"], pd.DataFrame([row])], ignore_index=True)
            tracker["Paid"] = tracker["Paid"].loc[tracker["Paid"]["Invoice No"] != inv_to_reverse].reset_index(drop=True)
            save_tracker(tracker)
            st.warning(f"Payment reversed for {inv_to_reverse}. Returned to Sent.")
            st.rerun()

    st.divider()

    st.markdown("## Cancel an invoice (Sent or Paid)")
    st.caption("Cancelled invoices remain visible for audit but are excluded from revenue/outstanding. A reason is required.")

    active = pd.concat([tracker["Sent"], tracker["Paid"]], ignore_index=True)
    if active.empty:
        st.info("No active invoices to cancel.")
    else:
        c1, c2, c3 = st.columns([2,2,4])
        with c1:
            inv_to_cancel = st.selectbox("Invoice", active["Invoice No"].tolist(), key="can_inv")
        with c2:
            cancelled_date = st.date_input("Cancelled date", value=date.today(), key="can_date")
        with c3:
            cancel_reason = st.text_input("Cancellation reason (required)", key="can_reason")

        if st.button("Cancel invoice", type="secondary"):
            if not cancel_reason.strip():
                st.error("Cancellation reason is required.")
            else:
                if (tracker["Sent"]["Invoice No"] == inv_to_cancel).any():
                    row = tracker["Sent"].loc[tracker["Sent"]["Invoice No"] == inv_to_cancel].iloc[0].copy()
                    tracker["Sent"] = tracker["Sent"].loc[tracker["Sent"]["Invoice No"] != inv_to_cancel].reset_index(drop=True)
                else:
                    row = tracker["Paid"].loc[tracker["Paid"]["Invoice No"] == inv_to_cancel].iloc[0].copy()
                    tracker["Paid"] = tracker["Paid"].loc[tracker["Paid"]["Invoice No"] != inv_to_cancel].reset_index(drop=True)

                row["Status"] = "Cancelled"
                row["Cancelled Date"] = cancelled_date
                row["Cancellation Reason"] = cancel_reason
                row["Paid Date"] = None
                row["Payment Method"] = None

                tracker["Cancelled"] = pd.concat([tracker["Cancelled"], pd.DataFrame([row])], ignore_index=True)
                save_tracker(tracker)
                st.warning(f"{inv_to_cancel} cancelled (audit retained).")
                st.rerun()

# ----------------------- TAB 3: Reports ---------------------

with tab_reports:
    st.subheader("Tracker & Reports")

    st.markdown("### Full tracker (read-only)")
    full = pd.concat(
        [
            tracker["Sent"].assign(_Sheet="Sent"),
            tracker["Paid"].assign(_Sheet="Paid"),
            tracker["Cancelled"].assign(_Sheet="Cancelled"),
        ],
        ignore_index=True,
    )
    st.dataframe(full, use_container_width=True, height=300)

    st.download_button(
        "⬇️ Download full tracker (Excel)",
        data=df_to_excel_bytes({"Sent": tracker["Sent"], "Paid": tracker["Paid"], "Cancelled": tracker["Cancelled"]}),
        file_name=f"Private_Services_Tracker_{date.today().isoformat()}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

    st.divider()
    st.markdown("### Reporting date range (Paid invoices)")

    paid = tracker["Paid"].copy()
    paid["Paid Date"] = pd.to_datetime(paid["Paid Date"], errors="coerce")
    paid["Amount"] = pd.to_numeric(paid["Amount"], errors="coerce").fillna(0.0)

    if paid["Paid Date"].notna().any():
        min_d = paid["Paid Date"].min().date()
        max_d = paid["Paid Date"].max().date()
    else:
        min_d = max_d = date.today()

    c1, c2 = st.columns(2)
    with c1:
        start_date = st.date_input("From", value=min_d, key="rep_from")
    with c2:
        end_date = st.date_input("To", value=max_d, key="rep_to")

    paid_f = paid[paid["Paid Date"].notna()].copy()
    paid_f["PaidDateOnly"] = paid_f["Paid Date"].dt.date
    paid_f = paid_f[(paid_f["PaidDateOnly"] >= start_date) & (paid_f["PaidDateOnly"] <= end_date)]

    sent = tracker["Sent"].copy()
    sent["Date"] = pd.to_datetime(sent["Date"], errors="coerce").dt.date
    if not sent.empty:
        today = date.today()
        sent["Days Outstanding"] = sent["Date"].apply(lambda d: (today - d).days if isinstance(d, date) else None)

    st.markdown("## Outstanding invoices (Sent)")
    if sent.empty:
        st.info("No outstanding invoices.")
    else:
        st.dataframe(sent, use_container_width=True)
        st.download_button("Download Outstanding (CSV)", data=sent.to_csv(index=False).encode("utf-8"), file_name="outstanding_invoices.csv", mime="text/csv")

    st.divider()
    st.markdown("## Clinician revenue totals (Paid)")
    if paid_f.empty:
        st.info("No paid invoices in selected date range.")
        totals = pd.DataFrame()
        pm_totals = pd.DataFrame()
        monthly = pd.DataFrame()
    else:
        totals = (
            paid_f.groupby("Clinician", dropna=False)["Amount"]
            .sum()
            .reset_index()
            .rename(columns={"Amount":"Total Revenue"})
            .sort_values("Total Revenue", ascending=False)
        )
        st.dataframe(totals, use_container_width=True)
        st.download_button("Download Clinician Totals (CSV)", data=totals.to_csv(index=False).encode("utf-8"), file_name="clinician_totals.csv", mime="text/csv")

        st.markdown("## Paid breakdown by payment method")
        pm = paid_f.copy()
        pm["Payment Method"] = pm["Payment Method"].fillna("Unknown").astype(str)
        pm_totals = (
            pm.groupby("Payment Method")["Amount"]
            .sum()
            .reset_index()
            .rename(columns={"Amount":"Total Revenue"})
            .sort_values("Total Revenue", ascending=False)
        )
        st.dataframe(pm_totals, use_container_width=True)

        st.markdown("## Monthly revenue chart (Paid)")
        tmp = paid_f.copy()
        tmp["YearMonth"] = pd.to_datetime(tmp["Paid Date"]).dt.to_period("M").astype(str)
        monthly = tmp.groupby("YearMonth")["Amount"].sum().reset_index().rename(columns={"Amount":"Revenue"}).sort_values("YearMonth")

        st.altair_chart(
            alt.Chart(monthly).mark_bar().encode(x="YearMonth:O", y="Revenue:Q", tooltip=["YearMonth","Revenue"]).properties(height=280),
            use_container_width=True
        )

    st.divider()
    st.markdown("## Status summary")
    status_counts = pd.DataFrame({"Status":["Sent","Paid","Cancelled"], "Count":[len(tracker["Sent"]), len(tracker["Paid"]), len(tracker["Cancelled"])]})
    st.altair_chart(alt.Chart(status_counts).mark_arc().encode(theta="Count:Q", color="Status:N", tooltip=["Status","Count"]), use_container_width=True)

    st.divider()
    st.markdown("## Export report (Excel)")
    if st.button("Generate report Excel"):
        report_bytes = df_to_excel_bytes({
            "Paid (filtered)": paid_f,
            "Sent (outstanding)": sent,
            "Clinician totals": totals if not totals.empty else pd.DataFrame(),
            "Payment method totals": pm_totals if not pm_totals.empty else pd.DataFrame(),
            "Monthly": monthly if not monthly.empty else pd.DataFrame(),
            "Status": status_counts,
            "Cancelled (audit)": tracker["Cancelled"],
        })
        st.download_button(
            "Download report Excel",
            data=report_bytes,
            file_name=f"PrivateServices_Report_{start_date}_{end_date}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
