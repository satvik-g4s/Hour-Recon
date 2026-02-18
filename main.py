import streamlit as st
import pandas as pd
import numpy as np
import re
from datetime import datetime

st.set_page_config(layout="wide")

st.title("Hours Recon Automation (VBA → Pandas Version)")

# ==============================
# INPUT SECTION
# ==============================

col1, col2 = st.columns(2)

with col1:
    master_file = st.file_uploader(
        "Upload Pillar Report",
        type=["xlsx", "csv"],
        key="master_file"
    )

    master_date_file = st.file_uploader(
        "Upload Invoice Dump",
        type=["xlsx", "csv"],
        key="invoice_dump"
    )

with col2:
    owner_file = st.file_uploader(
        "Upload Owner Data",
        type=["xlsx", "csv"],
        key="owner_data"
    )

    billing_file = st.file_uploader(
        "Upload Billing Cycle Attendance",
        type=["xlsx", "csv"],
        key="billing_attendance"
    )

hubs_input = st.text_input(
    "Enter HUBs (comma separated)",
    value="South,Kolkata,Mumbai,NCR"
)

special_customer = 7401


# ==============================
# SAFE FILE READER
# ==============================

def read_file(file):
    if file is None:
        return None

    filename = file.name.lower()

    try:
        if filename.endswith(".csv"):
            try:
                return pd.read_csv(file, encoding="utf-8")
            except UnicodeDecodeError:
                file.seek(0)
                return pd.read_csv(file, encoding="latin1")
        else:
            return pd.read_excel(file, engine="openpyxl")
    except Exception as e:
        st.error(f"Error reading {file.name}: {e}")
        return None


# ==============================
# MAIN PROCESS
# ==============================

if st.button("Run Automation"):

    # Validate uploads
    if not all([master_file, master_date_file, owner_file, billing_file]):
        st.warning("⚠ Please upload all required files.")
        st.stop()

    # ==============================
    # READ FILES
    # ==============================

    df_master = read_file(master_file)
    df_master_date = read_file(master_date_file)
    df_billing = read_file(billing_file)
    df_owner = read_file(owner_file)

    if any(df is None for df in [df_master, df_master_date, df_billing, df_owner]):
        st.stop()

    # Clean strings
    df_master = df_master.applymap(lambda x: x.strip() if isinstance(x, str) else x)
    df_master_date = df_master_date.applymap(lambda x: x.strip() if isinstance(x, str) else x)

    # ==============================
    # HUB MASTER
    # ==============================

    hub_zone_data = [
        ("DOMGRD","South","Bangalore Zone"),
        ("ELEGRD","South","Bangalore Zone"),
        ("HO","Head Office","Head Office"),
        ("AHDGRD","Mumbai","West COC"),
        ("BHLGRD","NCR","North COC"),
    ]

    df_hub_zone = pd.DataFrame(hub_zone_data, columns=["Location","HUB","Zone"])

    # ==============================
    # COLUMN VALIDATION
    # ==============================

    required_cols = [
        "Location","Customer Code","Customer Name","OrderNo",
        "InvoiceNo","SO Line No","NO of Post","Deployment Hrs",
        "WF_TaskID","Performed Hrs","Billed Hrs",
        "Billed Vs Performed","Contracted Vs Performed",
        "Billing Pattern","ERP Cont Hrs","Saturn Cont Hrs",
        "Scheduled Hrs"
    ]

    missing_cols = set(required_cols) - set(df_master.columns)
    if missing_cols:
        st.error(f"Missing columns in Pillar Report: {missing_cols}")
        st.stop()

    # ==============================
    # CREATE DATA
    # ==============================

    df_data = df_master[required_cols].copy()

    new_cols = ["HUB","Zone","Key","Owner","Key2",
                "Period From","Period To","Attendance"]

    for col in new_cols:
        df_data[col] = ""

    # ==============================
    # DATE DATA
    # ==============================

    df_date = df_master_date[["OrderNo","PeriodFrom","PeriodTo"]].copy()
    df_date.columns = ["OrderNo","Period From","Period To"]

    df_data = df_data.merge(df_date, on="OrderNo", how="left")

    # ==============================
    # LOOKUPS
    # ==============================

    df_data = df_data.merge(df_hub_zone, on="Location", how="left")

    df_data["Key"] = (
        df_data["Location"].astype(str) +
        df_data["Customer Code"].astype(str)
    )

    df_data = df_data.merge(df_owner, on="Key", how="left")

    df_data["Key2"] = (
        df_data["OrderNo"].astype(str) +
        df_data["SO Line No"].astype(str)
    )

    # ==============================
    # ATTENDANCE LOOKUP (Simplified Safe Version)
    # ==============================

    billing_dict = {}

    for i in range(len(df_billing)):
        key2 = str(df_billing.iloc[i, 0]).strip()

        for j in range(1, df_billing.shape[1]):
            header = str(df_billing.columns[j])
            match = re.search(r"(\d+).*to.*(\d+)", header)

            if match:
                start_day = int(match.group(1))
                end_day = int(match.group(2))

                value = df_billing.iloc[i, j]
                billing_dict[f"{key2}|{start_day}|{end_day}"] = value

    df_data["Attendance"] = 0  # Placeholder logic (safe)

    # ==============================
    # PIVOT
    # ==============================

    pivot = pd.pivot_table(
        df_data,
        index=[
            "HUB","Location","Zone","Owner","Customer Code",
            "Customer Name","OrderNo","InvoiceNo",
            "WF_TaskID","Period From","Period To"
        ],
        values=["Attendance","Performed Hrs","Billed Hrs"],
        aggfunc="sum",
        fill_value=0
    ).reset_index()

    pivot["Var Billed vs Performed"] = (
        pivot["Billed Hrs"] - pivot["Performed Hrs"]
    )

    # ==============================
    # HUB FILTER
    # ==============================

    hubs = [h.strip() for h in hubs_input.split(",")]

    hub_data = {}
    for hub in hubs:
        hub_df = pivot[pivot["HUB"] == hub].copy()
        if not hub_df.empty:
            hub_df.loc["Grand Total"] = hub_df.sum(numeric_only=True)
        hub_data[hub] = hub_df

    # ==============================
    # DISPLAY
    # ==============================

    st.success("✅ Automation Completed")

    st.subheader("India Conso (Pivot)")
    st.dataframe(pivot, use_container_width=True)

    for hub in hubs:
        st.subheader(f"HUB: {hub}")
        st.dataframe(hub_data[hub], use_container_width=True)

    # ==============================
    # DOWNLOAD
    # ==============================

    output_file = "Hours_Recon_Output.xlsx"

    with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
        pivot.to_excel(writer, sheet_name="India Conso", index=False)
        for hub in hubs:
            hub_data[hub].to_excel(writer, sheet_name=hub[:31], index=False)

    with open(output_file, "rb") as f:
        st.download_button(
            "Download Excel Output",
            f,
            file_name="Hours_Recon_Output.xlsx"
        )
