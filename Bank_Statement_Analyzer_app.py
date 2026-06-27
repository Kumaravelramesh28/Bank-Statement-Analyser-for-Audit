
import streamlit as st
import pandas as pd
import re
from io import BytesIO
from pathlib import Path

st.set_page_config(page_title="Bank Statement Analyzer", page_icon="🏦", layout="wide")

st.markdown("""
<style>
.main {padding-top:1rem;}
.stButton button, .stDownloadButton button {
    width:100%;
}
</style>
""", unsafe_allow_html=True)
def is_bank_code(token):

    token = str(token).strip().upper()

    if re.fullmatch(r'\d+', token):
        return True

    if re.fullmatch(r'[A-Z]{4}\d{4,}', token):
        return True

    if (
        re.search(r'[A-Z]', token)
        and re.search(r'\d', token)
        and len(token) >= 10
    ):
        return True

    return False

# ---------------------------
# EXISTING FUNCTION
# ---------------------------

def extract_narration(description):

    if pd.isna(description):
        return "UNKNOWN"

    desc = str(description).strip()

    try:

        # -------------------------------
        # UPI
        # -------------------------------

        if desc.upper().startswith("UPI/"):

            parts = desc.split("/")

            if len(parts) > 3:

                candidate = parts[3].strip()

                if candidate:
                    return candidate.upper()

            upi_match = re.search(
                r'([A-Za-z0-9._-]+@[A-Za-z0-9._-]+)',
                desc
            )

            if upi_match:
                return upi_match.group(1).upper()

        # -------------------------------
        # Generic Extraction
        # -------------------------------

        cleaned = desc.upper()

        # Remove OTH-VM, OTH-PO etc.
        cleaned = re.sub(
            r'\bOTH[- ]?[A-Z]*\b',
            ' ',
            cleaned
        )

        # Standardize separators
        cleaned = re.sub(
            r'[-_/]+',
            ' ',
            cleaned
        )

        cleaned = re.sub(
            r'\s+',
            ' ',
            cleaned
        )

        tokens = cleaned.split()

        meaningful = []

        stop_words = {
    "NEFT",
    "RTGS",
    "IMPS",
    "UPI",
    "CR",
    "DR",
    "ATTN",
    "INB",
    "TRANSFER",
    "DEBIT",
    "CREDIT",
    "FUNDS",
    "IFT",
    "BY",
    "TO",
    "FROM"
}

        IGNORE_WORDS = {

            "PAYMENT",
            "PAYMENTS",
            "RECEIPT",
            "RECEIPTS",
            "RECEIVED",
            "PAID",

            "JAN",
            "JANUARY",

            "FEB",
            "FEBRUARY",

            "MAR",
            "MARCH",

            "APR",
            "APRIL",

            "MAY",

            "JUN",
            "JUNE",

            "JUL",
            "JULY",

            "AUG",
            "AUGUST",

            "SEP",
            "SEPT",
            "SEPTEMBER",

            "OCT",
            "OCTOBER",

            "NOV",
            "NOVEMBER",

            "DEC",
            "DECEMBER"
        }

        for token in tokens:

            if token in stop_words:
                continue

            if token in IGNORE_WORDS:
                continue

            if is_bank_code(token):
                continue

            meaningful.append(token)

        narration = " ".join(meaningful)

        narration = re.sub(
            r'\s+',
            ' ',
            narration
        ).strip()

        if narration:
            return narration[:100]

        return cleaned[:100]
    
    except Exception:

        return str(description).upper()[:100]
    
def find_column(df, names):
    mapping = {str(c).strip().lower(): c for c in df.columns}
    for n in names:
        if n.lower() in mapping:
            return mapping[n.lower()]
    return None

def analyze(df, bank_name, preferred_value):
    desc_col = find_column(df, ["DESCRIPTION", "Narration", "Particulars", "Remarks"])
    credit_col = find_column(df, ["Credits", "Credit", "CR Amount"])
    debit_col = find_column(df, ["Debits", "Debit", "DR Amount"])

    if not desc_col:
        raise Exception("Description column not found")
    if not credit_col:
        raise Exception("Credits column not found")
    if not debit_col:
        raise Exception("Debits column not found")

    work = df.copy()

    work[credit_col] = pd.to_numeric(work[credit_col], errors="coerce").fillna(0)
    work[debit_col] = pd.to_numeric(work[debit_col], errors="coerce").fillna(0)

    work["Narration"] = work[desc_col].apply(extract_narration)

    result = (
        work.groupby("Narration", dropna=False)
        .agg(
            Credits=(credit_col, "sum"),
            Debits=(debit_col, "sum"),
            Count=("Narration", "size")
        )
        .reset_index()
    )

    qualified = result[
        (result["Credits"] >= preferred_value) |
        (result["Debits"] >= preferred_value)
    ].copy()

    qualified.insert(0, "Bank", bank_name)
    qualified["Explanation from Client"] = ""

    return qualified[[
        "Bank",
        "Narration",
        "Credits",
        "Debits",
        "Count",
        "Explanation from Client"
    ]]

def make_excel(df):
    output = BytesIO()

    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        df.to_excel(writer, sheet_name="Client Query Sheet", index=False)

        wb = writer.book
        ws = writer.sheets["Client Query Sheet"]

        header = wb.add_format({"bold": True})
        numfmt = wb.add_format({"num_format": '#,##0'})

        for col_num, value in enumerate(df.columns):
            ws.write(0, col_num, value, header)

        ws.freeze_panes(1, 0)
        ws.autofilter(0, 0, len(df), len(df.columns)-1)

        ws.set_column("A:A", 25)
        ws.set_column("B:B", 50)
        ws.set_column("C:D", 18, numfmt)
        ws.set_column("E:E", 10)
        ws.set_column("F:F", 35)

    output.seek(0)
    return output

st.title("🏦 Bank Statement Analyzer")
st.caption("Generate Client Query Sheet from Bank Statements")

st.subheader("Analysis Settings")

preferred_value = st.number_input(
    "Preferred Value (Minimum Individual/Cumulative Amount)",
    min_value=0,
    value=10000,
    step=1000,
    help="Narrations having cumulative Credits or Debits greater than or equal to this value will be included."
)

uploaded_file = st.file_uploader("Upload Excel File", type=["xlsx", "xls"])

if uploaded_file:
    try:
        bank_name = Path(uploaded_file.name).stem

        excel_file = pd.ExcelFile(uploaded_file)
        first_sheet = excel_file.sheet_names[0]

        df = pd.read_excel(uploaded_file, sheet_name=first_sheet)

        st.success(f"Bank Detected: {bank_name}")
        st.info(f"Reading First Worksheet: {first_sheet}")

        output_df = analyze(df, bank_name, preferred_value)

        st.info(f"Analysis Threshold: ₹{preferred_value:,.0f}")

        c1, c2, c3 = st.columns(3)
        c1.metric("Qualified Narrations", len(output_df))
        c2.metric("Total Credits", f"{output_df['Credits'].sum():,.0f}")
        c3.metric("Total Debits", f"{output_df['Debits'].sum():,.0f}")

        st.subheader("Output Preview")
        st.dataframe(output_df, use_container_width=True)

        excel_bytes = make_excel(output_df)

        st.download_button(
            label="📥 Download Client Query Sheet",
            data=excel_bytes,
            file_name=f"{bank_name}_Client_Query_Sheet.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        st.error(str(e))
