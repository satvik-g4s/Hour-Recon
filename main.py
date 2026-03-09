import streamlit as st
import pandas as pd
import re
from io import BytesIO
import numpy as np

st.set_page_config(layout="wide")

st.title("Hours Recon")

col1, col2, col3, col4 = st.columns(4)

with col1:
    uploaded_file_dump = st.file_uploader("Upload Dump CSV", type=["csv"])
    st.caption("Columns required:")
    st.caption("Order No, Period From, Period To, Invoice dt")

with col2:
    uploaded_file_pillar = st.file_uploader("Upload Pillar CSV", type=["csv"])
    st.caption("Columns required:")
    st.caption("Location, Customer Code, Customer Name, Order No, Invoice No, SO Line No, No of Post, Deployment Hrs, WF_TaskID, Performed Hrs, Billed Hrs, Billed Vs Performed, Contracted Vs Performed, Billing Pattern, ERP Cont Hrs, Saturn Cont Hrs, Scheduled Hrs")

with col3:
    uploaded_file_owner = st.file_uploader("Upload Owner Map CSV", type=["csv"])
    st.caption("Columns required:")
    st.caption("id, company_no, hub, so_locn, billing_location, hub_finance_head, branch_finance_lead, sscUser, sscUser1, Cust_No, Cust_Name, isRefresh")

with col4:
    uploaded_file_attendance = st.file_uploader("Upload Attendance Excel", type=["xlsx"])
    st.caption("Required key column:")
    st.caption("Row Labels")

run = st.button("Run")

if run:

    if uploaded_file_dump and uploaded_file_pillar and uploaded_file_owner and uploaded_file_attendance:

        st.write("Reading files...")

        dump = pd.read_csv(
            uploaded_file_dump,
            header=2,
            encoding="latin1",
            index_col=False,
            usecols=["Order No", "Period From", "Period To", "Invoice dt"]
        )

        pillar = pd.read_csv(
            uploaded_file_pillar,
            header=2,
            encoding="latin1",
            index_col=False,
            usecols=[
                "Location","Customer Code","Customer Name","Order No","Invoice No",
                "SO Line No","No of Post","Deployment Hrs","WF_TaskID",
                "Performed Hrs","Billed Hrs","Billed Vs Performed",
                "Contracted Vs Performed","Billing Pattern",
                "ERP Cont Hrs","Saturn Cont Hrs","Scheduled Hrs"
            ]
        )

        owner_map = pd.read_csv(uploaded_file_owner, index_col=False)

        attendance = pd.read_excel(uploaded_file_attendance, header=2)

        st.write("Cleaning data...")

        str_cols = dump.select_dtypes(include="object").columns
        dump[str_cols] = dump[str_cols].apply(lambda col: col.str.strip())

        pillar = pillar[pillar["Performed Hrs"] + pillar["Billed Hrs"] > 0]

        str_cols = pillar.select_dtypes(include="object").columns
        pillar[str_cols] = pillar[str_cols].apply(lambda col: col.str.strip())

        def normalize_order(s):
            return (
                s.astype(str)
                .str.strip()
                .str.replace(" ", "", regex=False)
                .str.upper()
            )

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
            dump_first[["Order No","Period From","Period To","Date_Range"]],
            on="Order No",
            how="left"
        )

        st.write("Processing attendance...")

        def normalize_attendance_col(col):
            if not isinstance(col, str):
                return col
            nums = re.findall(r"\d+", col)
            if len(nums) == 2:
                return f"{int(nums[0])}-{int(nums[1])}"
            return col

        attendance.columns = [normalize_attendance_col(c) for c in attendance.columns]

        def normalize_attendance_row_label(s):
            return (
                s.astype(str)
                .str.upper()
                .str.strip()
                .str.replace(" ", "", regex=False)
                .str.replace("-", "", regex=False)
            )

        pillar["row_key"] = (
            pillar["Order No"].astype(str).str.strip()
            + pillar["SO Line No"].astype(str).str.strip()
        )

        attendance["row_key"] = normalize_attendance_row_label(attendance["Row Labels"])

        attendance_long = attendance.melt(
            id_vars=["row_key"],
            var_name="Date_Range",
            value_name="Total Attendance"
        )

        pillar = pillar.merge(
            attendance_long,
            on=["row_key","Date_Range"],
            how="left"
        )

        pillar = pillar.drop(columns=["row_key"])

        st.write("Creating HUB Zone mapping...")

        hub_zone_data = [
            ("DOMGRD","South","Bangalore Zone"),
            ("ELEGRD","South","Bangalore Zone"),
            ("HOOGRD","South","Bangalore Zone"),
            ("HUBGRD","South","South COC"),
            ("MNGGRD","South","South COC"),
            ("MYOGRD","South","South COC"),
            ("SARGRD","South","Bangalore Zone"),
            ("YELGRD","South","Bangalore Zone"),
            ("YESGRD","South","Bangalore Zone"),
            ("ADYGRD","South","Chennai Zone"),
            ("ANNGRD","South","Chennai Zone"),
            ("CBTGRD","South","South COC"),
            ("CHNGRD","South","Chennai Zone"),
            ("COCGRD","South","South COC"),
            ("COMGRD","South","South COC"),
            ("GUIGRD","South","Chennai Zone"),
            ("HSRGRD","South","South COC"),
            ("MMNGRD","South","Chennai Zone"),
            ("PONGRD","South","South COC"),
            ("SRIGRD","South","Chennai Zone"),
            ("HO","Head Office","Head Office"),
            ("ANPGRD","South","Hyderabad Zone"),
            ("HYDGRD","South","Hyderabad Zone"),
            ("HYRGRD","South","Hyderabad Zone"),
            ("HYTGRD","South","Hyderabad Zone"),
            ("JBHGRD","South","Hyderabad Zone"),
            ("VIGGRD","South","South COC"),
            ("VJWGRD","South","South COC"),
            ("ALIGRD","Kolkata","Kolkata Zone"),
            ("ASLGRD","Kolkata","East COC"),
            ("BARGRD","Kolkata","Odisha Zone"),
            ("BBLGRD","Kolkata","Odisha Zone"),
            ("BBRGRD","Kolkata","Odisha Zone"),
            ("BHRGRD","Kolkata","East COC"),
            ("GHTGRD","Kolkata","East COC"),
            ("JAJGRD","Kolkata","Odisha Zone"),
            ("JAMGRD","Kolkata","East COC"),
            ("JASGRD","Kolkata","East COC"),
            ("JHAGRD","Kolkata","Odisha Zone"),
            ("PATGRD","Kolkata","East COC"),
            ("PTNGRD","Kolkata","East COC"),
            ("RAIGRD","Kolkata","East COC"),
            ("RJHGRD","Kolkata","Kolkata Zone"),
            ("ROUGRD","Kolkata","Odisha Zone"),
            ("SALGRD","Kolkata","Kolkata Zone"),
            ("SILGRD","Kolkata","East COC"),
            ("USCGRD","Kolkata","East COC"),
            ("AHDGRD","Mumbai","West COC"),
            ("AHMGRD","Mumbai","West COC"),
            ("AINGRD","Mumbai","West COC"),
            ("ANKGRD","Mumbai","West COC"),
            ("BODGRD","Mumbai","West COC"),
            ("CORMSP","Mumbai","Mumbai Zone"),
            ("DEUGRD","Mumbai","West COC"),
            ("GONGRD","Mumbai","West COC"),
            ("JNAGRD","Mumbai","West COC"),
            ("MNMGRD","Mumbai","Mumbai Zone"),
            ("MNVGRD","Mumbai","Mumbai Zone"),
            ("MONGRD","Mumbai","Mumbai Zone"),
            ("MSOGRD","Mumbai","Mumbai Zone"),
            ("MUCGRD","Mumbai","Mumbai Zone"),
            ("MUSGRD","Mumbai","West COC"),
            ("MUSMSP","Mumbai","Mumbai Zone"),
            ("NAGGRD","Mumbai","West COC"),
            ("PNEGRD","Mumbai","Pune Zone"),
            ("PNRGRD","Mumbai","Pune Zone"),
            ("PUNGRD","Mumbai","Pune Zone"),
            ("PUWGRD","Mumbai","Pune Zone"),
            ("BHLGRD","NCR","North COC"),
            ("CHDGRD","NCR","North COC"),
            ("CP1GRD","NCR","Delhi Zone"),
            ("DDNGRD","NCR","North COC"),
            ("DUNGRD","NCR","North COC"),
            ("EMBGRD","NCR","Delhi Zone"),
            ("FBDGRD","NCR","Gurgaon Zone"),
            ("FRMGRD","NCR","Delhi Zone"),
            ("GGNGRD","NCR","Gurgaon Zone"),
            ("GHAGRD","NCR","Noida Zone"),
            ("GNBGRD","NCR","Gurgaon Zone"),
            ("GNSGRD","NCR","Gurgaon Zone"),
            ("IDRGRD","NCR","North COC"),
            ("JALGRD","NCR","North COC"),
            ("JARGRD","NCR","North COC"),
            ("JAUGRD","NCR","North COC"),
            ("JMUGRD","NCR","North COC"),
            ("JNKGRD","NCR","North COC"),
            ("JPRGRD","NCR","North COC"),
            ("LKWGRD","NCR","North COC"),
            ("LUDGRD","NCR","North COC"),
            ("MNSGRD","NCR","Gurgaon Zone"),
            ("MRTGRD","NCR","North COC"),
            ("NDAGRD","NCR","Noida Zone"),
            ("NDGGRD","NCR","Noida Zone"),
            ("PSPGRD","NCR","Delhi Zone"),
            ("PWNGRD","NCR","North COC"),
            ("RPRGRD","NCR","North COC"),
            ("SPTGRD","NCR","North COC"),
            ("UDRGRD","NCR","North COC"),
            ("USEGRD","NCR","North COC"),
            ("UTKGRD","NCR","North COC"),
        ]

        hub_zone = pd.DataFrame(
            hub_zone_data,
            columns=["Location","HUB","Zone"]
        )

        hub_zone["Location"] = normalize_order(hub_zone["Location"])
        pillar["Location"] = normalize_order(pillar["Location"])

        pillar = pillar.merge(
            hub_zone,
            on="Location",
            how="left"
        )

        # normalize columns before building key
        st.write("Owner mapping...")
        
        
        # Create the keys with more careful handling
        owner_map["billing_location"] = (
            owner_map["billing_location"]
            .astype(str)
            .str.strip()
            .str.replace(" ", "", regex=False)
            .str.upper()
        )
        
        pillar["Location"] = (
            pillar["Location"]
            .astype(str)
            .str.strip()
            .str.replace(" ", "", regex=False)
            .str.upper()
        )
        
        # Handle customer codes - ensure they're strings and strip any special characters
        owner_map["Cust_No"] = (
            owner_map["Cust_No"]
            .astype(str)
            .str.strip()
            .str.replace(".0", "", regex=False)  # Remove decimal if any
            .str.upper()
        )
        
        pillar["Customer Code"] = (
            pillar["Customer Code"]
            .astype(str)
            .str.strip()
            .str.replace(".0", "", regex=False)  # Remove decimal if any
            .str.upper()
        )
        
        # Create keys
        pillar["Key"] = pillar["Location"] + "_" + pillar["Customer Code"]  # Add separator for clarity
        owner_map["Key"] = owner_map["billing_location"] + "_" + owner_map["Cust_No"]

        
        # Perform the merge
        pillar = pillar.merge(
            owner_map[["Key", "branch_finance_lead"]],
            on="Key",
            how="left"
        )
        
        
        # Show sample of unmatched records
       
        pillar = pillar.rename(columns={"branch_finance_lead": "Owner"})

        st.write("Creating pivot...")

        pivot = (
            pillar.groupby([
                "HUB","Location","Zone","Owner",
                "Customer Code","Customer Name",
                "Order No","Invoice No","WF_TaskID",
                "Period From","Period To"
            ], dropna=False)[
                ["Total Attendance","Performed Hrs","Billed Hrs"]
            ]
            .sum()
            .reset_index()
        )

        pivot = pivot.rename(columns={
            "Performed Hrs":"Total Performed",
            "Billed Hrs":"Total Billed"
        })

        pivot["Var. Performed Vs. Billed"] = (
            pivot["Total Billed"] - pivot["Total Performed"]
        )

        pivot["Office Duty/Office Patrolling"] = np.where(
            pivot["Customer Code"].astype(str) == "7401",
            pivot["Total Performed"],
            ""
        )

        extra_cols = [ "Excess Paid", "Reliever duty", "Excess billing", "Short billing", "Disciplinary Deduction", "Short / Missing Roster", "Inter assignment adjustment", "Indirect Hours Not Captured in Saturn", "Training & OJT", "Complimentary Hrs.", "Billing Cycle/ hours calculation other than calendar month", "Bill Hrs should being Cycle", "Diff with bill cycle should be", "Total ( B )", "Check (A - B)", "BFL Remarks", "SSC Query (If Any)" ]



        
        for col in extra_cols:
            pivot[col] = pd.NA
        pivot = pivot[[ "HUB", "Location", "Zone", "Owner", "Customer Code", "Customer Name", "Order No", "Invoice No", "WF_TaskID", "Period From", "Period To", "Total Attendance", "Total Performed", "Total Billed", "Var. Performed Vs. Billed", "Office Duty/Office Patrolling", "Excess Paid", "Reliever duty", "Excess billing", "Short billing", "Disciplinary Deduction", "Short / Missing Roster", "Inter assignment adjustment", "Indirect Hours Not Captured in Saturn", "Training & OJT", "Complimentary Hrs.", "Billing Cycle/ hours calculation other than calendar month", "Bill Hrs should being Cycle", "Diff with bill cycle should be", "Total ( B )", "Check (A - B)", "BFL Remarks", "SSC Query (If Any)" ]]


        pivot["Inter assignment adjustment"] = ""

        seen = {}
        
        for idx, row in pivot.iterrows():
        
            order = str(row["Order No"]).strip()
            val = row["Var. Performed Vs. Billed"]
        
            if pd.isna(val) or val == 0:
                continue
        
            key = (order, round(val, 6))
            reverse_key = (order, round(-val, 6))
        
            if reverse_key in seen:
        
                prev_idx = seen[reverse_key]
        
                pivot.loc[idx, "Inter assignment adjustment"] = -val
                pivot.loc[prev_idx, "Inter assignment adjustment"] = -pivot.loc[prev_idx, "Var. Performed Vs. Billed"]
        
                del seen[reverse_key]
        
            else:
                seen[key] = idx

        india_conso = pivot.copy()

        st.write("Creating HUB sheets...")

        hub_sheets = {}

        for hub in india_conso["HUB"].dropna().unique():
            hub_sheets[hub] = india_conso[india_conso["HUB"] == hub]

        st.write("Preparing Excel output...")

        output = BytesIO()

        with pd.ExcelWriter(output, engine="xlsxwriter") as writer:

            india_conso.to_excel(
                writer,
                sheet_name="India Conso",
                index=False
            )

            for hub, df in hub_sheets.items():

                df.to_excel(
                    writer,
                    sheet_name=str(hub)[:31],
                    index=False
                )

        st.dataframe(india_conso, use_container_width=True)

        st.download_button(
            "Download Output Excel",
            data=output.getvalue(),
            file_name="Hours_Recon_Output.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    else:
        st.warning("Please upload all files.")
