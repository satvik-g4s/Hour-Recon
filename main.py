import streamlit as st
import pandas as pd
import re

st.set_page_config(layout="wide")

col1, col2, col3, col4 = st.columns(4)

with col1:
    uploaded_file_dump = st.file_uploader("Upload Dump CSV (columns: Order No, Period From, Period To, Invoice dt)", type=["csv"])
with col2:
    uploaded_file_pillar = st.file_uploader("Upload Pillar CSV (columns: Location, Customer Code, Customer Name, Order No, Invoice No, SO Line No, No of Post, Deployment Hrs, WF_TaskID, Performed Hrs, Billed Hrs, Billed Vs Performed, Contracted Vs Performed, Billing Pattern, ERP Cont Hrs, Saturn Cont Hrs, Scheduled Hrs)", type=["csv"])
with col3:
    uploaded_file_owner = st.file_uploader("Upload Owner Map CSV (columns:'id', 'company_no', 'hub', 'so_locn', 'billing_location',
       'hub_finance_head', 'branch_finance_lead', 'sscUser', 'sscUser1',
       'Cust_No', 'Cust_Name', 'isRefresh')", type=["csv"])
with col4:
    uploaded_file_attendance = st.file_uploader("Upload Attendance Excel("Row Labels" should be the name of the key column)", type=["xlsx"])

run = st.button("Run")

if run:
    if uploaded_file_dump and uploaded_file_pillar and uploaded_file_owner and uploaded_file_attendance:

        st.write("Reading files...")

        dump = pd.read_csv(uploaded_file_dump, header=2, encoding="latin1", index_col=False,
                           usecols=["Order No", "Period From", "Period To", "Invoice dt"])

        pillar = pd.read_csv(uploaded_file_pillar, header=2, encoding="latin1", index_col=False,
                             usecols=["Location", "Customer Code", "Customer Name", "Order No", "Invoice No",
                                      "SO Line No", "No of Post", "Deployment Hrs", "WF_TaskID",
                                      "Performed Hrs", "Billed Hrs", "Billed Vs Performed",
                                      "Contracted Vs Performed", "Billing Pattern",
                                      "ERP Cont Hrs", "Saturn Cont Hrs", "Scheduled Hrs"])

        owner_map = pd.read_csv(uploaded_file_owner, index_col=False)
        attendance = pd.read_excel(uploaded_file_attendance, header=2)

        st.write("Cleaning data...")

        str_cols = dump.select_dtypes(include="object").columns
        dump[str_cols] = dump[str_cols].apply(lambda col: col.str.strip())

        pillar = pillar[pillar["Performed Hrs"] + pillar["Billed Hrs"] > 0]

        str_cols = pillar.select_dtypes(include="object").columns
        pillar[str_cols] = pillar[str_cols].apply(lambda col: col.str.strip())

        def normalize_order(s):
            return s.astype(str).str.strip().str.replace(" ", "", regex=False).str.upper()

        dump["Order No"] = normalize_order(dump["Order No"])
        pillar["Order No"] = normalize_order(pillar["Order No"])

        dump["Period From"] = pd.to_datetime(dump["Period From"], errors="coerce")
        dump["Period To"] = pd.to_datetime(dump["Period To"], errors="coerce")

        dump = dump.sort_values(by="Invoice dt", ascending=False)
        dump_first = dump.drop_duplicates(subset=["Order No"], keep="first")

        dump_first["Date_Range"] = (
            dump_first["Period From"].dt.day.astype("Int64").astype(str)
            + "-"
            + dump_first["Period To"].dt.day.astype("Int64").astype(str)
        )

        pillar = pillar.merge(
            dump_first[["Order No", "Period From", "Period To", "Date_Range"]],
            on="Order No",
            how="left"
        )

        st.write("Processing attendance...")

        def normalize_attendance_col(col):
            if not isinstance(col, str):
                return col
            nums = re.findall(r'\d+', col)
            if len(nums) == 2:
                return f"{int(nums[0])}-{int(nums[1])}"
            return col

        attendance.columns = [normalize_attendance_col(c) for c in attendance.columns]

        def normalize_attendance_row_label(s):
            return s.astype(str).str.upper().str.strip().str.replace(" ", "", regex=False).str.replace("-", "", regex=False)

        pillar["row_key"] = pillar["Order No"].astype(str).str.strip() + pillar["SO Line No"].astype(str).str.strip()

        attendance["row_key"] = normalize_attendance_row_label(attendance["Row Labels"])

        attendance_long = attendance.melt(
            id_vars=["row_key"],
            var_name="Date_Range",
            value_name="Total Attendance"
        )

        pillar = pillar.merge(attendance_long, on=["row_key", "Date_Range"], how="left")
        pillar = pillar.drop(columns=["row_key"])

        st.write("Creating pivot...")

        pivot = (
            pillar.groupby([
                "Location",
                "Customer Code",
                "Customer Name",
                "Order No",
                "Invoice No",
                "WF_TaskID",
                "Period From",
                "Period To"
            ], dropna=False)[
                ["Total Attendance", "Performed Hrs", "Billed Hrs"]
            ].sum().reset_index()
        )

        pivot = pivot.rename(columns={
            "Performed Hrs": "Total Performed",
            "Billed Hrs": "Total Billed"
        })

        pivot["Var. Performed Vs. Billed"] = pivot["Total Billed"] - pivot["Total Performed"]

        st.dataframe(pivot, use_container_width=True)

        csv = pivot.to_csv(index=False).encode("utf-8")
        st.download_button("Download Output CSV", csv, "output.csv", "text/csv")

    else:
        st.write("Please upload all files.")    
