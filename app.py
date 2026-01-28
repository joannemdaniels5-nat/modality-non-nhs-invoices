from pathlib import Path
from datetime import datetime, date
from io import BytesIO
import base64

import altair as alt
import pandas as pd
import streamlit as st

# ===================== BASIC CONFIG =====================

st.set_page_config(
    page_title="Modality Lewisham - Private Services",
    layout="wide",
    page_icon="💷",
)

BASE_DIR = Path(__file__).parent
DATA_DIR = BASE_DIR / "data"
TRACKER_DIR = DATA_DIR / "Trackers"
TRACKER_PATH = TRACKER_DIR / "Private Services Tracker.xlsx"
INVOICE_ROOT_DIR = DATA_DIR / "Invoices"

PRICE_LIST_PATH = DATA_DIR / "Private Services Price List.xlsx"
GP_PATH = DATA_DIR / "ModalityLewGP_List.xlsx"

LOGO_CANDIDATES = [
    BASE_DIR / "logo.png",
    DATA_DIR / "logo.png",
]

# Ensure base folders exist
DATA_DIR.mkdir(parents=True, exist_ok=True)
TRACKER_DIR.mkdir(parents=True, exist_ok=True)
INVOICE_ROOT_DIR.mkdir(parents=True, exist_ok=True)

# ===================== PASSWORD GATE =====================

APP_PASSWORD = "Modality2026!"  # internal password for the app


def password_gate():
    """Simple password gate using Streamlit session_state."""
    if st.session_state.get("auth_ok", False):
        return

    st.title("🔐 Modality Lewisham - Private Services")
    st.write("This app is for internal Modality Lewisham use only.")

    pw = st.text_input("Enter password", type="password")
    if st.button("Login"):
        if pw == APP_PASSWORD:
            st.session_state["auth_ok"] = True
            st.experimental_rerun()
        else:
            st.error("Incorrect password.")
            st.stop()
    else:
        st.stop()


password_gate()

# ===================== HELPER FUNCTIONS =====================


def load_tracker():
    """
    Load tracker as dict of DataFrames: {'Sent': df, 'Paid': df}.
    If file missing or invalid, create fresh structure.
    """
    expected_cols = [
        "Invoice No",
        "Clinician",
        "Patient",
        "Service",
        "Amount",
        "Status",
        "Invoice File",
        "Date",
        "Paid Date",
    ]

    if TRACKER_PATH.exists():
        try:
            df_dict = pd.read_excel(TRACKER_PATH, sheet_name=None)
        except Exception:
            df_dict = {}
    else:
        df_dict = {}

    for sheet in ["Sent", "Paid"]:
        if sheet not in df_dict:
            df_dict[sheet] = pd.DataFrame(columns=expected_cols)
        else:
            # Ensure all expected columns exist
            for col in expected_cols:
                if col not in df_dict[sheet].columns:
                    df_dict[sheet][col] = None

    return df_dict


def save_tracker(df_dict: dict):
    """Save tracker dict back to Excel (Sent / Paid sheets)."""
    TRACKER_DIR.mkdir(parents=True, exist_ok=True)
    with pd.ExcelWriter(TRACKER_PATH, engine="openpyxl") as writer:
        for sheet_name, df in df_dict.items():
            df.to_excel(writer, sheet_name=sheet_name, index=False)


def load_price_list():
    """
    Load price list, or create a simple template if missing.
    Expects columns: Service, Price
    """
    if PRICE_LIST_PATH.exists():
        try:
            df = pd.read_excel(PRICE_LIST_PATH)
        except Exception:
            df = pd.DataFrame()
    else:
        df = pd.DataFrame()

    if df.empty or "Service" not in df.columns or "Price" not in df.columns:
        df = pd.DataFrame(
            {
                "Service": [
                    "Private Sick Note",
                    "To Whom It May Concern Letter",
                    "Private Prescription",
                    "Private Blood Test",
                    "Private GP Consultation (15 mins)",
                    "Medical & Report (30 mins)",
                ],
                "Price": [40.0, 30.0, 30.0, 36.0, 90.0, 150.0],
            }
        )
        df.to_excel(PRICE_LIST_PATH, index=False)

    df = df[["Service", "Price"]].copy()
    df["Service"] = df["Service"].astype(str)
    df["Price"] = pd.to_numeric(df["Price"], errors="coerce").fillna(0.0)
    return df


def load_gp_list():
    """
    Load GP list from data/ModalityLewGP_List.xlsx

    Your file format:
    first name | surname | role | ClinicianDisplay
    """
    if not GP_PATH.exists():
        st.warning("GP list not found at 'data/ModalityLewGP_List.xlsx'.")
        return []

    try:
        df = pd.read_excel(GP_PATH)
    except Exception as e:
        st.warning(f"Could not read GP list: {e}")
        return []

    # Normalise headers
    df.columns = (
        df.columns.str.strip()
        .str.lower()
        .str.replace(" ", "_")
        .str.replace("-", "_")
    )

    # Filter to GP if role present
    if "role" in df.columns:
        df["role"] = df["role"].astype(str)
        df = df[df["role"].str.upper() == "GP"]

    # Prefer ClinicianDisplay
    if "cliniciandisplay" in df.columns:
        names = df["cliniciandisplay"].astype(str).str.strip()
    elif {"first_name", "surname"}.issubset(df.columns):
        names = (
            df["first_name"].astype(str).str.strip()
            + " "
            + df["surname"].astype(str).str.strip()
        )
    elif "name" in df.columns:
        names = df["name"].astype(str).str.strip()
    else:
        first_col = df.columns[0]
        names = df[first_col].astype(str).str.strip()

    names = (
        names[names != ""]
        .dropna()
        .drop_duplicates()
        .sort_values()
    )

    return names.tolist()


def get_logo_html_tag() -> str:
    """
    Embed logo as base64 <img> tag if available; otherwise plain text heading.
    """
    logo_path = None
    for cand in LOGO_CANDIDATES:
        if cand.exists():
            logo_path = cand
            break

    if not logo_path:
        return "<h2>Modality Lewisham</h2>"

    try:
        with open(logo_path, "rb") as f:
            img_bytes = f.read()
        b64 = base64.b64encode(img_bytes).decode("ascii")
        return (
            f"<img src='data:image/png;base64,{b64}' "
            f"style='height:60px;'>"
        )
    except Exception:
        return "<h2>Modality Lewisham</h2>"


def build_invoice_html(
    invoice_no: str,
    invoice_date: date,
    clinician: str,
    patient: str,
    patient_dob: date | None,
    patient_address: str,
    service: str,
    amount: float,
) -> str:
    logo_html = get_logo_html_tag()
    dob_str = patient_dob.strftime("%d/%m/%Y") if patient_dob else ""
    date_str = invoice_date.strftime("%d/%m/%Y")

    patient_block = f"<p><b>Patient:</b> {patient}</p>"
    if dob_str:
        patient_block += f"<p><b>Date of Birth:</b> {dob_str}</p>"
    if patient_address.strip():
        patient_block += (
            f"<p><b>Address:</b><br>"
            f"{patient_address.strip().replace(chr(10), '<br>')}</p>"
        )

    html = f"""
    <html>
    <head>
        <meta charset="utf-8">
        <title>Invoice {invoice_no}</title>
    </head>
    <body style="font-family: Arial, sans-serif; margin: 40px;">
        <div style="display:flex; justify-content:space-between; align-items:center;">
            <div>{logo_html}</div>
            <div style="text-align:right;">
                <h3>Private Services Invoice</h3>
                <p>Modality Lewisham</p>
            </div>
        </div>
        <hr>
        <p><b>Invoice No:</b> {invoice_no}</p>
        <p><b>Invoice Date:</b> {date_str}</p>
        <p><b>Clinician:</b> {clinician}</p>
        {patient_block}
        <hr>
        <p><b>Service:</b> {service}</p>
        <p><b>Total Amount (incl. VAT where applicable):</b> £{amount:.2f}</p>
        <hr>
        <p>
            Please arrange payment to Modality Lewisham in line with our
            Private Services Policy. This invoice includes VAT where applicable.
        </p>
    </body>
    </html>
    """
    return html


def create_invoice_file(
    invoice_no: str,
    clinician: str,
    patient: str,
    patient_dob: date | None,
    patient_address: str,
    service: str,
    amount: float,
) -> tuple[str, str]:
    """
    Create invoice HTML file in data/Invoices/YYYY/Month/Invoice_<no>.html

    Returns (filepath_str, html_string)
    """
    now = datetime.now()
    year_str = str(now.year)
    month_str = now.strftime("%B")

    inv_dir = INVOICE_ROOT_DIR / year_str / month_str
    inv_dir.mkdir(parents=True, exist_ok=True)

    filename = f"Invoice_{invoice_no}.html"
    filepath = inv_dir / filename

    html_str = build_invoice_html(
        invoice_no,
        now.date(),
        clinician,
        patient,
        patient_dob,
        patient_address,
        service,
        amount,
    )

    with open(filepath, "w", encoding="utf-8") as f:
        f.write(html_str)

    return str(filepath), html_str


def parse_dates_for_reporting(df: pd.DataFrame) -> pd.DataFrame:
    """Ensure Date and Paid Date are parsed as date objects."""
    if "Date" in df.columns:
        df["Date"] = pd.to_datetime(df["Date"], errors="coerce").dt.date
    if "Paid Date" in df.columns:
        df["Paid Date"] = pd.to_datetime(df["Paid Date"], errors="coerce").dt.date
    return df


def build_reports(data_dict: dict, start_date: date | None, end_date: date | None):
    """
    Build reporting DataFrames based on filtered date range.
    BasisDate = Paid Date (if present) else Invoice Date.
    """
    sent = data_dict["Sent"].copy()
    paid = data_dict["Paid"].copy()

    sent = parse_dates_for_reporting(sent)
    paid = parse_dates_for_reporting(paid)

    for df in (sent, paid):
        df["Amount"] = pd.to_numeric(df["Amount"], errors="coerce").fillna(0.0)
        df["BasisDate"] = df["Paid Date"].fillna(df["Date"])
        df["BasisDate"] = pd.to_datetime(df["BasisDate"], errors="coerce").dt.date

    def filter_by_range(df: pd.DataFrame) -> pd.DataFrame:
        if start_date and end_date:
            return df[
                df["BasisDate"].apply(
                    lambda d: isinstance(d, date) and start_date <= d <= end_date
                )
            ]
        return df

    sent_f = filter_by_range(sent)
    paid_f = filter_by_range(paid)

    # outstanding = Sent not Paid (in filtered window)
    outstanding = sent_f.copy()
    today = date.today()
    if not outstanding.empty:
        outstanding["Days Outstanding"] = outstanding["Date"].apply(
            lambda d: (today - d).days if isinstance(d, date) else None
        )

    clinician_totals = pd.DataFrame()
    if not paid_f.empty:
        clinician_totals = (
            paid_f.groupby("Clinician", dropna=False)["Amount"]
            .sum()
            .reset_index()
            .rename(columns={"Amount": "Total Revenue"})
            .sort_values("Total Revenue", ascending=False)
        )

    monthly_rev = pd.DataFrame()
    if not paid_f.empty:
        tmp = paid_f.copy()
        tmp["BasisDT"] = pd.to_datetime(tmp["BasisDate"], errors="coerce")
        tmp["YearMonth"] = tmp["BasisDT"].dt.to_period("M").astype(str)
        monthly_rev = (
            tmp.groupby("YearMonth")["Amount"]
            .sum()
            .reset_index()
            .rename(columns={"Amount": "Total Revenue"})
            .sort_values("YearMonth")
        )

    df_all = pd.concat(
        [
            sent_f.assign(_StatusGroup="Sent"),
            paid_f.assign(_StatusGroup="Paid"),
        ],
        ignore_index=True,
    )

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
        "paid_filtered": paid_f,
        "sent_filtered": sent_f,
    }


# ===================== MAIN UI =====================

st.title("💷 Modality Lewisham - Private Services (Non-NHS)")

# Load data
tracker = load_tracker()
price_list = load_price_list()
gp_names = load_gp_list()

# Tabs
tab_create, tab_reports = st.tabs(["Create Invoice", "Reports"])

# ---------- TAB: CREATE INVOICE ----------
with tab_create:
    st.header("Create New Invoice")

    if not gp_names:
        st.warning(
            "No clinicians found in 'ModalityLewGP_List.xlsx'. "
            "Please check the file in the data/ folder."
        )

    col1, col2 = st.columns(2)

    with col1:
        clinician = st.selectbox(
            "Clinician",
            options=[""] + gp_names,
        )
        patient = st.text_input("Patient full name")
        include_dob = st.checkbox("Include date of birth?")
        patient_dob = None
        if include_dob:
            patient_dob = st.date_input("Patient date of birth")

    with col2:
        patient_address = st.text_area(
            "Patient address (optional)", height=100
        )

    col3, col4 = st.columns(2)
    with col3:
        service_options = price_list["Service"].tolist()
        service = st.selectbox("Service", options=[""] + service_options)
    with col4:
        default_price = 0.0
        if service and service in price_list["Service"].values:
            default_price = float(
                price_list.loc[price_list["Service"] == service, "Price"].iloc[0]
            )
        amount = st.number_input(
            "Total Amount (£, incl. VAT where applicable)",
            step=1.0,
            format="%.2f",
            value=default_price,
        )

    st.caption(f"Prices loaded from: `{PRICE_LIST_PATH.name}`")

    if st.button("Generate Invoice"):
        if not clinician or not patient or not service:
            st.error("Clinician, Patient and Service must be completed.")
        elif amount <= 0:
            st.error("Amount must be greater than zero.")
        else:
            now = datetime.now()
            invoice_no = f"INV-{now.strftime('%Y%m%d-%H%M%S')}"

            filepath, html_str = create_invoice_file(
                invoice_no,
                clinician,
                patient,
                patient_dob,
                patient_address,
                service,
                amount,
            )

            new_row = {
                "Invoice No": invoice_no,
                "Clinician": clinician,
                "Patient": patient,
                "Service": service,
                "Amount": amount,
                "Status": "Sent",
                "Invoice File": filepath,
                "Date": now.date(),
                "Paid Date": None,
            }
            tracker["Sent"] = pd.concat(
                [tracker["Sent"], pd.DataFrame([new_row])],
                ignore_index=True,
            )
            save_tracker(tracker)

            st.success(f"Invoice generated: `{invoice_no}`")

            st.download_button(
                "⬇️ Download Invoice (HTML)",
                data=html_str.encode("utf-8"),
                file_name=f"{invoice_no}.html",
                mime="text/html",
            )

            st.info(
                "Open the downloaded HTML file in your browser and use "
                "**Print → Save as PDF** to create a PDF for emailing or printing."
            )

    st.subheader("Sent Invoices")
    if tracker["Sent"].empty:
        st.info("No sent invoices yet.")
    else:
        st.dataframe(tracker["Sent"], use_container_width=True)

    st.subheader("Mark Invoice as Paid")
    if tracker["Sent"].empty:
        st.info("No invoices available to mark as paid.")
    else:
        inv_options = tracker["Sent"]["Invoice No"].tolist()
        selected = st.selectbox("Select invoice", inv_options)
        paid_date = st.date_input("Paid date", value=date.today())

        if st.button("Confirm Paid"):
            row = tracker["Sent"][tracker["Sent"]["Invoice No"] == selected].iloc[0].copy()
            row["Status"] = "Paid"
            row["Paid Date"] = paid_date

            tracker["Paid"] = pd.concat(
                [tracker["Paid"], pd.DataFrame([row])],
                ignore_index=True,
            )
            tracker["Sent"] = tracker["Sent"][tracker["Sent"]["Invoice No"] != selected]
            save_tracker(tracker)
            st.success(f"Invoice {selected} marked as Paid.")

    st.subheader("Paid Invoices")
    if tracker["Paid"].empty:
        st.info("No paid invoices yet.")
    else:
        st.dataframe(tracker["Paid"], use_container_width=True)

# ---------- TAB: REPORTS ----------
with tab_reports:
    st.header("Reports")

    combined = pd.concat(
        [
            tracker["Sent"].assign(_Source="Sent"),
            tracker["Paid"].assign(_Source="Paid"),
        ],
        ignore_index=True,
    )
    combined = parse_dates_for_reporting(combined)
    combined["BasisDate"] = combined["Paid Date"].fillna(combined["Date"])
    combined["BasisDate"] = pd.to_datetime(
        combined["BasisDate"], errors="coerce"
    ).dt.date

    valid_dates = [d for d in combined["BasisDate"] if isinstance(d, date)]
    if valid_dates:
        global_min = min(valid_dates)
        global_max = max(valid_dates)
    else:
        global_min = global_max = date.today()

    c1, c2 = st.columns(2)
    with c1:
        start_date = st.date_input("From (Basis date)", value=global_min)
    with c2:
        end_date = st.date_input("To (Basis date)", value=global_max)

    reports = build_reports(tracker, start_date, end_date)

    outstanding = reports["outstanding"]
    clinician_totals = reports["clinician_totals"]
    monthly_rev = reports["monthly_rev"]
    status_counts = reports["status_counts"]
    paid_filtered = reports["paid_filtered"]
    sent_filtered = reports["sent_filtered"]

    st.subheader("Outstanding Invoices")
    if outstanding.empty:
        st.info("No outstanding invoices in the selected date range.")
    else:
        st.dataframe(outstanding, use_container_width=True)
        csv = outstanding.to_csv(index=False).encode("utf-8")
        st.download_button(
            "Download Outstanding (CSV)",
            data=csv,
            file_name="outstanding_invoices.csv",
            mime="text/csv",
        )

    st.subheader("Clinician Revenue (Paid)")
    if paid_filtered.empty:
        st.info("No paid invoices in the selected date range.")
    else:
        st.dataframe(clinician_totals, use_container_width=True)
        csv = clinician_totals.to_csv(index=False).encode("utf-8")
        st.download_button(
            "Download Clinician Revenue (CSV)",
            data=csv,
            file_name="clinician_revenue.csv",
            mime="text/csv",
        )

    st.subheader("Monthly Revenue (Paid)")
    if monthly_rev.empty:
        st.info("No paid invoices for monthly revenue chart.")
    else:
        chart = (
            alt.Chart(monthly_rev)
            .mark_bar()
            .encode(
                x="YearMonth:O",
                y="Total Revenue:Q",
                tooltip=["YearMonth", "Total Revenue"],
            )
            .properties(height=300)
        )
        st.altair_chart(chart, use_container_width=True)

        csv = monthly_rev.to_csv(index=False).encode("utf-8")
        st.download_button(
            "Download Monthly Revenue (CSV)",
            data=csv,
            file_name="monthly_revenue.csv",
            mime="text/csv",
        )

    st.subheader("Payment Status (Sent vs Paid)")
    if status_counts.empty:
        st.info("No invoices in this date range for status chart.")
    else:
        pie = (
            alt.Chart(status_counts)
            .mark_arc()
            .encode(
                theta="Count:Q",
                color="Status:N",
                tooltip=["Status", "Count"],
            )
        )
        st.altair_chart(pie, use_container_width=True)

        csv = status_counts.to_csv(index=False).encode("utf-8")
        st.download_button(
            "Download Status Counts (CSV)",
            data=csv,
            file_name="status_counts.csv",
            mime="text/csv",
        )

    st.subheader("Full Excel Export (Filtered)")
    if st.button("Generate Excel Report"):
        output = BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            sent_filtered.to_excel(writer, sheet_name="Sent", index=False)
            paid_filtered.to_excel(writer, sheet_name="Paid", index=False)
            outstanding.to_excel(writer, sheet_name="Outstanding", index=False)
            clinician_totals.to_excel(writer, sheet_name="ClinicianTotals", index=False)
            monthly_rev.to_excel(writer, sheet_name="MonthlyRevenue", index=False)
            status_counts.to_excel(writer, sheet_name="StatusCounts", index=False)

        output.seek(0)
        st.download_button(
            "Download Excel Report",
            data=output,
            file_name=f"PrivateServices_Report_{date.today()}.xlsx",
            mime=(
                "application/vnd.openxmlformats-"
                "officedocument.spreadsheetml.sheet"
            ),
        )
