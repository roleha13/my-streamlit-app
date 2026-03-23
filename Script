import pandas as pd
from datetime import datetime
from io import BytesIO
import streamlit as st

# ================= CORE FUNCTIONS =================

def corrected_month_to_period(dt):
    if pd.isna(dt):
        return ""
    month = dt.month
    year = dt.year + 1 if month >= 7 else dt.year
    period_num = ((month - 7) % 12) + 2
    return f"{year}/{period_num:03d}"

def format_month_label(dt):
    if pd.isna(dt):
        return ""
    return dt.strftime("%B'%y")

def process_file(uploaded_file):
    xls = pd.ExcelFile(uploaded_file)
    rows = []

    for sheet_name in xls.sheet_names:
        df = pd.read_excel(xls, sheet_name=sheet_name, dtype=str)

        if df.shape[1] < 16:
            continue

        mask = df.iloc[:, 6].str.contains("POS Goldenkey Ltd", case=False, na=False)
        filtered = df[mask].copy()

        if filtered.empty:
            continue

        selected = filtered.iloc[:, [1, 9, 15]].copy()
        selected.columns = ["ColB", "ColJ", "ColP"]

        for _, r in selected.iterrows():
            dt = pd.to_datetime(r["ColB"], errors="coerce")
            if pd.isna(dt):
                continue

            rows.append({
                "1": "1;3;6",
                "TRANSACTION REFERENCE": format_month_label(dt),
                "DESCRIPTION": r["ColP"] if pd.notna(r["ColP"]) else "",
                "ACCOUNT CODE": "440100",
                "TRANSACTION DATE": dt,
                "PERIOD": corrected_month_to_period(dt),
                "BASE AMOUNT": float(str(r["ColJ"]).replace(",", "").strip() or 0),
                "DEBIT/CREDIT": "D",
                "TRANSACTION CURRENCY": "KSH"
            })

    return pd.DataFrame(rows)

def generate_master(files):
    all_data = []

    for file in files:
        df = process_file(file)
        if not df.empty:
            all_data.append(df)

    if not all_data:
        return None

    df = pd.concat(all_data, ignore_index=True)
    df["TRANSACTION DATE"] = pd.to_datetime(df["TRANSACTION DATE"])

    grouped = []

    for date, group in df.groupby("TRANSACTION DATE"):
        total_amt = group["BASE AMOUNT"].sum()
        month_label = format_month_label(date)
        period = corrected_month_to_period(date)

        grouped.append(group)

        grouped.append(pd.DataFrame([{
            "1": "1;3;6",
            "TRANSACTION REFERENCE": month_label,
            "DESCRIPTION": f"FOOD INV {date.strftime('%d/%m/%Y')}",
            "ACCOUNT CODE": "CT00311",
            "TRANSACTION DATE": date,
            "PERIOD": period,
            "BASE AMOUNT": total_amt,
            "DEBIT/CREDIT": "C",
            "TRANSACTION CURRENCY": "KSH"
        }]))

    df_final = pd.concat(grouped, ignore_index=True)

    output = BytesIO()
    df_final.to_excel(output, index=False)
    output.seek(0)

    return output

# ================= STREAMLIT UI =================

st.title("📊 GKC Food Invoice Generator (Web Version)")

uploaded_files = st.file_uploader(
    "Upload Excel Files",
    type=["xlsx", "xls"],
    accept_multiple_files=True
)

if st.button("Generate Master Excel"):
    if uploaded_files:
        result = generate_master(uploaded_files)

        if result:
            st.success("File generated successfully!")
            st.download_button(
                label="Download Excel",
                data=result,
                file_name="CRJ_ALL_TRANSFORMED.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        else:
            st.warning("No matching data found.")
    else:
        st.error("Please upload at least one file.")
