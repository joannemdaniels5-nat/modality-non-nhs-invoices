import streamlit as st
import pandas as pd
from pathlib import Path
from datetime import datetime, date
from io import BytesIO
import altair as alt

# ===================== CONFIG =====================

# Base directory for data inside the app (works locally and on Streamlit Cloud)
BASE_DIR = Path(__file__).parent / "data"

TRACKER_PATH = BASE_DIR / "Trackers" / "Private Services Tracker.xlsx"
INVOICE_ROOT_DIR = BASE_DIR / "Invoices"
PRICE_LIST_PATH = BASE_DIR / "Private Services Price List.xlsx"
GP_LIST_PATH = BASE_DIR / "ModalityLewGP_List.xlsx"

# Ensure folders exist
(INVOICE_ROOT_DIR).mkdir(parents=True, exist_ok=True)
(TRACKER_PATH.parent).mkdir(parents=True, exist_ok=True)
(PRICE_LIST_PATH.parent).mkdir(parents=True, exist_ok=True)
(GP_LIST_PATH.parent).mkdir(parents=True, exist_ok=True)

# ===================== TRACKER LOAD / SAVE =====================

def load_tracker():
    """
    Load tracker as dict of DataFrames: {"Sent": df, "Paid": df}
    If file missing or corrupt, create fresh structure.
    """
    if TRACKER_PATH.exists():
        try:
            df_dict = pd.read_excel(TRACKER_PATH, sheet_name=None)
        except Exception:
            df_dict = {}
    else:
        df_dict = {}

    if "Sent" not in df_dict:
        df_dict["Sent"] = pd.DataFrame(columns=[
            "Invoice No", "Clinician", "Patient", "Service",
            "Amount", "Status", "Invoice File", "Date", "Paid Date"
        ])
    if "Paid" not in df_dict:
        df_dict["Paid"] = pd.DataFrame(columns=[
            "Invoice No", "Clinician", "Patient", "Service",
            "Amount", "Status", "Invoice File", "Date", "Paid Date"
        ])

    # Ensure all expected columns exist
    for sheet in ["Sent", "Paid"]:
        for col in ["Invoice No", "Clinician", "Patient", "Service",
                    "Amount", "Status", "Invoice File", "Date", "Paid Date"]:
            if col not in df_dict[sheet].columns:
                df_dict[sheet][col] = None

    return df_dict


def save_tracker(data_dict):
    """
    Save dict of DataFrames back to Excel with sheets Sent / Paid.
    """
    with pd.ExcelWriter(TRACKER_PATH, engine="openpyxl") as writer:
        for sheet_name, df in data_dict.items():
            df.to_excel(writer, sheet_name=sheet_name, index=False)

# ===================== PRICE LIST =====================

def load_price_list():
    """
    Load the price list from Excel.
    If it doesn't exist or is invalid, create a template.
    Expected columns: Service, Price
    """
    if PRICE_LIST_PATH.exists():
        try:
            df = pd.read_excel(PRICE_LIST_PATH)
        except Exception:
            df = pd.DataFrame()
    else:
        df = pd.DataFrame()

    if df.empty or "Service" not in df.columns or "Price" not in df.columns:
        df = pd.DataFrame({
            "Service": [
                "Private Sick Note",
                "To Whom It May Concern Letter",
                "Private Prescription",
                "Private Blood Test",
                "Private GP Consultation (15 mins)",
                "Medical & Report (30 mins)"
            ],
            "Price": [
                40.0,
                30.0,
                30.0,
                36.0,
                90.0,
                150.0,
            ]
        })
        PRICE_LIST_PATH.parent.mkdir(parents=True, exist_ok=True)
        df.to_excel(PRICE_LIST_PATH, index=False)

    df = df[["Service", "Price"]].copy()
    df["Service"] = df["Service"].astype(str)
    df["Price"] = pd.to_numeric(df["Price"], errors="coerce").fillna(0.0)
    return df

# ===================== GP LIST =====================

def load_gp_list():
    """
    Load GP list from Excel.
    Template is created if missing.
    Expected columns: First name, Surname, Role
    """
    if GP_LIST_PATH.exists():
        try:
            df = pd.read_excel(GP_LIST_PATH)
        except Exception:
            df = pd.DataFrame()
    else:
        df = pd.DataFrame()

    if df.empty or "First name" not in df.columns or "Surname" not in df.columns:
        # Create a simple template
        df = pd.DataFrame({
            "First name": ["Eoghan"],
            "Surname": ["MacSweeney"],
            "Role": ["GP"],
        })
        df.to_excel(GP_LIST_PATH, index=False)

    # Filter to GPs if Role exists; otherwise use all
    if "Role" in df.columns:
        df = df[df["Role"].astype(str).str.upper() == "GP"]

    df["ClinicianDisplay"] = df["First name"].astype(str).str.strip() + " " + df["Surname"].astype(str).str.strip()
    names = sorted(df["ClinicianDisplay"].unique().tolist())
    return names if names else None

# ===================== INVOICE GENERATION =====================

def generate_invoice_file(invoice_no, clinician, patient, service, amount):
    """
    Create HTML invoice file in the structure:
    data/Invoices/YYYY/MonthName/Invoice_<invoice_no>.html
    """
    now = datetime.now()
    year_str = str(now.year)
    month_name = now.strftime("%B")

    inv_dir = INVOICE_ROOT_DIR / year_str / month_name
    inv_dir.mkdir(parents=True, exist_ok=True)

    filename = f"Invoice_{invoice_no}.html"
    filepath = inv_dir / filename

    date_str = now.strftime("%Y-%m-%d")

    html = f"""
    <html>
    <body>
        <h2 style='text-align:center;'>Modality Lewisham - Private Services Invoice</h2>
        <hr>
        <p><b>Invoice No:</b> {invoice_no}</p>
        <p><b>Date:</b> {date_str}</p>
        <p><b>Clinician:</b> {clinician}</p>
        <p><b>Patient:</b> {patient}</p>
        <p><b>Service:</b> {service}</p>
        <p><b>Total Amount:</b> £{amount:.2f}</p>
        <hr>
        <p>Please arrange payment to Modality Lewisham.</p>
    </body>
    </html>
    """

    with open(filepath, "w", encoding="utf-8") as f:
        f.write(html)

    return str(filepath)

# ===================== REPORT HELPERS =====================

def parse_dates_for_reporting(df):
    """
    Ensure Date and Paid Date are parsed as datetime.date
    """
    if "Date" in df.columns:
        df["Date"] = pd.to_datetime(df["Date"], errors="coerce").dt.date
    if "Paid Date" in df.columns:
        df["Paid Date"] = pd.to_datetime(df["Paid Date"], errors="coerce").dt.date
    return df

def build_reports(data_dict):
    """
    Build derived reporting DataFrames from tracker.
    Returns a dict of tables: outstanding, clinician_totals, monthly_rev, status_counts
    """
    sent = data_dict["Sent"].copy()
    paid = data_dict["Paid"].copy()

    sent = parse_dates_for_reporting(sent)
    paid = parse_dates_for_reporting(paid)

    today = date.today()

    # Outstanding = Sent rows (not yet Paid)
    outstanding = sent.copy()
    if not outstanding.empty:
        outstanding["Days Outstanding"] = outstanding["Date"].apply(
            lambda d: (today - d).days if isinstance(d, date) else None
        )

    # Clinician revenue totals (Paid only)
    paid["Amount"] = pd.to_numeric(paid["Amount"], errors="coerce").fillna(0.0)
    clinician_totals = pd.DataFrame()
    if not paid.empty:
        clinician_totals = (
            paid.groupby("Clinician", dropna=False)["Amount"]
            .sum()
            .reset_index()
            .rename(columns={"Amount": "Total Revenue"})
            .sort_values("Total Revenue", ascending=False)
        )

    # Monthly revenue (Paid Date if available, else Date)
    monthly_rev = pd.DataFrame()
    if not paid.empty:
        paid["BasisDate"] = paid["Paid Date"].fillna(paid["Date"])
        paid["BasisDate"] = pd.to_datetime(paid["BasisDate"], errors="coerce")
        paid["YearMonth"] = paid["BasisDate"].dt.to_period("M").astype(str)
        monthly_rev = (
            paid.groupby("YearMonth")["Amount"]
            .sum()
            .reset_index()
            .rename(columns={"Amount": "Total Revenue"})
            .sort_values("YearMonth")
        )

    # Payment status (Sent vs Paid) from both
    df_all = pd.concat([
        sent.assign(_StatusGroup="Sent"),
        paid.assign(_StatusGroup="Paid")
    ], ignore_index=True)
    status_counts = pd.DataFrame()
    if not df_all.empty:
        status_counts = (
            df_all.groupby("_StatusGroup")["Invoice No"]
            .count()
            .reset_index()
            .rename(columns={"Invoice No": "Count", "_StatusGroup": "Status"})
        )

    return {
        "outstanding": outstanding,
        "clinician_totals": clinician_totals,
        "monthly_rev": monthly_rev,
        "status_counts": status_counts,
        "paid": paid,
    }

# ===================== STREAMLIT APP =====================

st.set_page_config("Modality Private Services", layout="wide")
st.title("💷 Modality Lewisham - Private Services")

data = load_tracker()
price_list = load_price_list()
gp_names = load_gp_list()
service_options = price_list["Service"].tolist()

tabs = st.tabs(["Create Invoice", "Reports"])

# ---------- TAB 1: CREATE INVOICE ----------
with tabs[0]:
    st.header("Create New Invoice")

    col1, col2 = st.columns(2)

    with col1:
        # Clinician from GP list, or free text fallback
        if gp_names:
            clinician = st.selectbox("Clinician", options=[""] + gp_names)
        else:
            clinician = st.text_input("Clinician")

        patient = st.text_input("Patient")

    with col2:
        # Service dropdown from price list
        service = st.selectbox("Service", options=[""] + service_options)

        # Default price
        if service and service in price_list["Service"].values:
            default_price = float(
                price_list.loc[price_list["Service"] == service, "Price"].iloc[0]
            )
        else:
            default_price = 0.0

        amount = st.number_input(
            "Amount (£)",
            step=1.0,
            format="%.2f",
            value=default_price
        )

    st.caption("Prices loaded from: " + str(PRICE_LIST_PATH))

    if st.button("Generate Invoice"):
        if not clinician or not patient or not service:
            st.error("Clinician, Patient and Service must be completed.")
        else:
            now = datetime.now()
            invoice_no = f"INV-{now.strftime('%Y%m%d-%H%M%S')}"
            invoice_file = generate_invoice_file(
                invoice_no, clinician, patient, service, amount
            )

            new_row = {
                "Invoice No": invoice_no,
                "Clinician": clinician,
                "Patient": patient,
                "Service": service,
                "Amount": amount,
                "Status": "Sent",
                "Invoice File": invoice_file,
                "Date": now.date(),
                "Paid Date": None
            }

            data["Sent"] = pd.concat(
                [data["Sent"], pd.DataFrame([new_row])],
                ignore_index=True
            )
            save_tracker(data)

            st.success(f"Invoice generated and saved: {invoice_file}")

    st.subheader("Sent Invoices")
    if data["Sent"].empty:
        st.info("No sent invoices yet.")
    else:
        st.dataframe(data["Sent"], use_container_width=True)

    st.subheader("Mark Invoice as Paid")
    if data["Sent"].empty:
        st.info("No invoices available to mark as paid.")
    else:
        inv_options = data["Sent"]["Invoice No"].tolist()
        selected = st.selectbox("Select Invoice to mark as paid", inv_options)
        paid_date = st.date_input("Paid Date", value=date.today())

        if st.button("Confirm Paid"):
            row = data["Sent"][data["Sent"]["Invoice No"] == selected].iloc[0].copy()
            row["Status"] = "Paid"
            row["Paid Date"] = paid_date

            data["Paid"] = pd.concat(
                [data["Paid"], pd.DataFrame([row])],
                ignore_index=True
            )
            data["Sent"] = data["Sent"][data["Sent"]["Invoice No"] != selected]
            save_tracker(data)
            st.success(f"Invoice {selected} marked as Paid.")

    st.subheader("Paid Invoices")
    if data["Paid"].empty:
        st.info("No paid invoices yet.")
    else:
        st.dataframe(data["Paid"], use_container_width=True)

# ---------- TAB 2: REPORTS ----------
with tabs[1]:
    st.header("Reports")

    reports = build_reports(data)
    outstanding = reports["outstanding"]
    clinician_totals = reports["clinician_totals"]
    monthly_rev = reports["monthly_rev"]
    status_counts = reports["status_counts"]
    paid = reports["paid"]

    # --- Outstanding Invoices Summary ---
    st.subheader("Outstanding Invoices (Sent, not yet Paid)")
    if outstanding.empty:
        st.info("No outstanding invoices.")
    else:
        st.dataframe(outstanding, use_container_width=True)

    # --- Clinician Revenue (Totals + Date Range) ---
    st.subheader("Clinician Revenue")

    if paid.empty:
        st.info("No paid invoices yet for revenue reporting.")
    else:
        paid["Amount"] = pd.to_numeric(paid["Amount"], errors="coerce").fillna(0.0)

        # Overall totals
        st.markdown("**Total Revenue by Clinician (all time)**")
        st.dataframe(clinician_totals, use_container_width=True)

        # Date range filter
        valid_dates = [d for d in paid["Paid Date"] if isinstance(d, date)]
        if valid_dates:
            min_date = min(valid_dates)
            max_date = max(valid_dates)
        else:
            min_date = max_date = date.today()

        c1, c2 = st.columns(2)
        with c1:
            start_date = st.date_input("From (Paid Date)", value=min_date)
        with c2:
            end_date = st.date_input("To (Paid Date)", value=max_date)

        mask = paid["Paid Date"].apply(
            lambda d: isinstance(d, date) and start_date <= d <= end_date
        )
        paid_filtered = paid[mask]

        if paid_filtered.empty:
            st.info("No paid invoices in the selected date range.")
        else:
            clin_range = (
                paid_filtered
                .groupby("Clinician", dropna=False)["Amount"]
                .sum()
                .reset_index()
                .rename(columns={"Amount": "Total Revenue"})
                .sort_values("Total Revenue", ascending=False)
            )
            st.markdown("**Clinician Revenue in Selected Date Range**")
            st.dataframe(clin_range, use_container_width=True)

    # --- Monthly Revenue Chart ---
    st.subheader("Monthly Revenue (Paid)")
    if monthly_rev.empty:
        st.info("No paid invoices yet for monthly revenue chart.")
    else:
        chart = (
            alt.Chart(monthly_rev)
            .mark_bar()
            .encode(
                x="YearMonth:O",
                y="Total Revenue:Q",
                tooltip=["YearMonth", "Total Revenue"]
            )
            .properties(height=300)
        )
        st.altair_chart(chart, use_container_width=True)

    # --- Payment Status Pie Chart ---
    st.subheader("Payment Status (Count of Invoices)")
    if status_counts.empty:
        st.info("No invoices yet for status chart.")
    else:
        pie_chart = (
            alt.Chart(status_counts)
            .mark_arc()
            .encode(
                theta="Count:Q",
                color="Status:N",
                tooltip=["Status", "Count"]
            )
        )
        st.altair_chart(pie_chart, use_container_width=True)

    # --- Export Reports to Excel ---
    st.subheader("Export Reports to Excel")
    if st.button("Generate Excel Report"):
        output = BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            data["Sent"].to_excel(writer, sheet_name="Sent", index=False)
            data["Paid"].to_excel(writer, sheet_name="Paid", index=False)
            outstanding.to_excel(writer, sheet_name="Outstanding", index=False)
            clinician_totals.to_excel(writer, sheet_name="ClinicianTotals", index=False)
            monthly_rev.to_excel(writer, sheet_name="MonthlyRevenue", index=False)
            status_counts.to_excel(writer, sheet_name="StatusCounts", index=False)

        output.seek(0)
        st.download_button(
            "Download Excel Report",
            data=output,
            file_name=f"PrivateServices_Report_{date.today()}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
