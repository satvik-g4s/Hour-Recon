import streamlit as st
import pandas as pd
import re
from io import BytesIO
import numpy as np

st.set_page_config(layout="wide")

def read_file(file, header=0, usecols=None):
    
    if file.name.endswith(".csv"):
        return pd.read_csv(
            file,
            header=header,
            encoding="latin1",
            index_col=False,
            usecols=usecols
        )

    elif file.name.endswith(".xlsx"):
        return pd.read_excel(
            file,
            header=header,
            usecols=usecols
        )

st.title("Hours Recon")

col1, col2, col3, col4 = st.columns(4)

with col1:
    uploaded_file_dump = st.file_uploader("Upload Dump CSV", type=["csv","xlsx"])
    st.caption("Columns required:")
    st.caption("Order No, Period From, Period To, Invoice dt")

with col2:
    uploaded_file_pillar = st.file_uploader("Upload Pillar CSV", type=["csv","xlsx"])
    st.caption("Columns required:")
    st.caption("Location, Customer Code, Customer Name, Order No, Invoice No, SO Line No, No of Post, Deployment Hrs, WF_TaskID, Performed Hrs, Billed Hrs, Billed Vs Performed, Contracted Vs Performed, Billing Pattern, ERP Cont Hrs, Saturn Cont Hrs, Scheduled Hrs")

with col3:
    uploaded_file_owner = st.file_uploader("Upload Owner Map CSV", type=["csv","xlsx"])
    st.caption("Columns required:")
    st.caption("id, company_no, hub, so_locn, billing_location, hub_finance_head, branch_finance_lead, sscUser, sscUser1, Cust_No, Cust_Name, isRefresh")

with col4:
    uploaded_file_attendance = st.file_uploader("Upload Attendance Excel", type=["csv","xlsx"])
    st.caption("Required key column:")
    st.caption("Row Labels")

run = st.button("Run")

if run:

    if uploaded_file_dump and uploaded_file_pillar and uploaded_file_owner and uploaded_file_attendance:

        st.write("Reading files...")

        dump = read_file( uploaded_file_dump, header=2, usecols=["Order No","Period From","Period To","Invoice dt"] )

        pillar = read_file( uploaded_file_pillar, header=2, usecols=[ "Location","Customer Code","Customer Name","Order No","Invoice No", "SO Line No","No of Post","Deployment Hrs","WF_TaskID", "Performed Hrs","Billed Hrs","Billed Vs Performed", "Contracted Vs Performed","Billing Pattern", "ERP Cont Hrs","Saturn Cont Hrs","Scheduled Hrs" ] )

        owner_map = read_file(uploaded_file_owner)

        attendance = read_file(uploaded_file_attendance, header=2)

        st.write("Cleaning data...")
        owner_map.columns = owner_map.columns .str.strip() .str.lower() .str.replace(" ", "_") 


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
        dump["Invoice dt"] = pd.to_datetime(dump["Invoice dt"], errors="coerce")


        dump = dump.sort_values(
            ["Invoice dt", "Period To", "Period From"],
            ascending=[False, False, False]
        )
        dump_first = dump.drop_duplicates(subset=["Order No"], keep="first").copy()

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
        pillar["SO Line No"] = (
            pillar["SO Line No"]
            .astype(str)
            .str.replace(".0","", regex=False)
            .str.strip()
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
            ("Kolkata","ALIGRD","Kolkata Zone"),
            ("Kolkata","ASLGRD","East COC"),
            ("Kolkata","BBRGRD","Odisha Zone"),
            ("Kolkata","BBLGRD","Odisha Zone"),
            ("Kolkata","JAJGRD","Odisha Zone"),
            ("Kolkata","JHAGRD","Odisha Zone"),
            ("Kolkata","JASGRD","East COC"),
            ("Kolkata","PATGRD","East COC"),
            ("Kolkata","PTNGRD","East COC"),
            ("Kolkata","BHRGRD","East COC"),
            ("Kolkata","DALGRD","Kolkata Zone"),
            ("Kolkata","GUWGRD","East COC"),
            ("Kolkata","GHTGRD","East COC"),
            ("Kolkata","HOWGRD","Kolkata Zone"),
            ("Kolkata","RJHGRD","Kolkata Zone"),
            ("Kolkata","KOLGRD","Kolkata Zone"),
            ("Kolkata","SALGRD","Kolkata Zone"),
            ("Kolkata","SILGRD","East COC"),
            ("Kolkata","USCGRD","East COC"),
            ("Kolkata","RAIGRD","East COC"),
            ("Kolkata","BARGRD","Odisha Zone"),
            ("Kolkata","ROUGRD","Odisha Zone"),
            ("Kolkata","JAMGRD","East COC"),
            ("Kolkata","KONGRD","Kolkata Zone"),            
            ("NCR","BHLGRD","North COC"),
            ("NCR","IDRGRD","North COC"),
            ("NCR","CP1GRD","Delhi Zone"),
            ("NCR","CP2GRD","Delhi Zone"),
            ("NCR","DROGRD","Delhi Zone"),
            ("NCR","EMBGRD","Delhi Zone"),
            ("NCR","FRMGRD","Delhi Zone"),
            ("NCR","PSPGRD","Delhi Zone"),
            ("NCR","GOLGRD","Delhi Zone"),
            ("NCR","VVRGRD","Delhi Zone"),
            ("NCR","USEGRD","North COC"),
            ("NCR","GHAGRD","Noida Zone"),
            ("NCR","LKWGRD","North COC"),
            ("NCR","MRTGRD","North COC"),
            ("NCR","NDAGRD","Noida Zone"),
            ("NCR","NDGGRD","Noida Zone"),
            ("NCR","CHDGRD","North COC"),
            ("NCR","CROGRD","North COC"),
            ("NCR","DDNGRD","North COC"),
            ("NCR","UTKGRD","North COC"),
            ("NCR","JMUGRD","North COC"),
            ("NCR","JAUGRD","North COC"),
            ("NCR","JNKGRD","North COC"),
            ("NCR","PRWGRD","North COC"),
            ("NCR","PWNGRD","North COC"),
            ("NCR","RUDGRD","North COC"),
            ("NCR","RPRGRD","North COC"),
            ("NCR","FBDGRD","Gurgaon Zone"),
            ("NCR","GGNGRD","Gurgaon Zone"),
            ("NCR","GNBGRD","Gurgaon Zone"),
            ("NCR","GNSGRD","Gurgaon Zone"),
            ("NCR","MNSGRD","Gurgaon Zone"),
            ("NCR","SPTGRD","North COC"),
            ("NCR","JALGRD","North COC"),
            ("NCR","LUDGRD","North COC"),
            ("NCR","JPRGRD","North COC"),
            ("NCR","DHRGRD","North COC"),
            ("NCR","JARGRD","North COC"),
            ("NCR","UDRGRD","North COC"),
            ("NCR","DUNGRD","North COC"),
            ("NCR","SNPGRD","North COC"),
            ("NCR","TYMGRD","Delhi Zone"),
            ("NCR","TEPGRD","Noida Zone"),
            ("NCR","OKLGRD","Delhi Zone"),            
            ("South","HUBGRD","South COC"),
            ("South","BELGRD","South COC"),
            ("South","BANGRD","Bangalore Zone"),
            ("South","BLRGRD","Bangalore Zone"),
            ("South","DOMGRD","Bangalore Zone"),
            ("South","ELEGRD","Bangalore Zone"),
            ("South","HOOGRD","Bangalore Zone"),
            ("South","ORRGRD","Bangalore Zone"),
            ("South","SARGRD","Bangalore Zone"),
            ("South","VASGRD","Bangalore Zone"),
            ("South","WHTGRD","Bangalore Zone"),
            ("South","YELGRD","Bangalore Zone"),
            ("South","YESGRD","Bangalore Zone"),
            ("South","MNGGRD","South COC"),
            ("South","MYOGRD","South COC"),
            ("South","MYSGRD","South COC"),
            ("South","HOPGRD","Bangalore Zone"),
            ("South","COMGRD","South COC"),
            ("South","CBTGRD","South COC"),
            ("South","ADYGRD","Chennai Zone"),
            ("South","ANNGRD","Chennai Zone"),
            ("South","CHNGRD","Chennai Zone"),
            ("South","GUIGRD","Chennai Zone"),
            ("South","MMNGRD","Chennai Zone"),
            ("South","NUGGRD","Chennai Zone"),
            ("South","SRIGRD","Chennai Zone"),
            ("South","COCGRD","South COC"),
            ("South","PONGRD","South COC"),
            ("South","MADGRD","South COC"),
            ("South","TRVGRD","South COC"),
            ("South","SIRGRD","Chennai Zone"),
            ("South","SLMGRD","South COC"),
            ("South","HYDGRD","Hyderabad Zone"),
            ("South","HYRGRD","Hyderabad Zone"),
            ("South","HYTGRD","Hyderabad Zone"),
            ("South","JBHGRD","Hyderabad Zone"),
            ("South","MHPGRD","Hyderabad Zone"),
            ("South","VIGGRD","South COC"),
            ("South","VIZGRD","South COC"),
            ("South","VJWGRD","South COC"),
            ("South","VWDGRD","South COC"),
            ("South","ANPGRD","Hyderabad Zone"),
            ("South","HSRGRD","South COC"),            
            ("Mumbai","AHDGRD","West COC"),
            ("Mumbai","AHMGRD","West COC"),
            ("Mumbai","AINGRD","West COC"),
            ("Mumbai","ANKGRD","West COC"),
            ("Mumbai","BODGRD","West COC"),
            ("Mumbai","JNAGRD","West COC"),
            ("Mumbai","MLDGRD","Mumbai Zone"),
            ("Mumbai","MNMGRD","Mumbai Zone"),
            ("Mumbai","MNVGRD","Mumbai Zone"),
            ("Mumbai","MSOGRD","Mumbai Zone"),
            ("Mumbai","MUCGRD","Mumbai Zone"),
            ("Mumbai","MUMGRD","Mumbai Zone"),
            ("Mumbai","MUSGRD","West COC"),
            ("Mumbai","GONGRD","West COC"),
            ("Mumbai","GOAGRD","West COC"),
            ("Mumbai","NAGGRD","West COC"),
            ("Mumbai","PROGRD","West COC"),
            ("Mumbai","PNEGRD","Pune Zone"),
            ("Mumbai","PNHGRD","Pune Zone"),
            ("Mumbai","RJGGRD","Pune Zone"),
            ("Mumbai","PNRGRD","Pune Zone"),
            ("Mumbai","PUWGRD","Pune Zone"),
            ("Mumbai","PUNGRD","Pune Zone"),
            ("Mumbai","DEUGRD","West COC"),
            ("Mumbai","PNIGRD","Pune Zone"),
            ("Mumbai","MONGRD","Mumbai Zone"),
            ("Mumbai","MUSMSP","Mumbai Zone"),
            ("Mumbai","CORMSP","Mumbai Zone"),            
            ("HeadOffice","INVGRD","Head Office"),
            ("HeadOffice","OTHGRD","Head Office"),
            ("HeadOffice","PSOGRD","Head Office"),
            ("HeadOffice","TRGGRD","Head Office"),
            ("HeadOffice","CORGRD","Head Office"),
            ("HeadOffice","HO","Head Office"),            
            ("NCR","HIMGRD","North COC")            ]

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
        owner_map["cust_no"] = (
            owner_map["cust_no"]
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
        owner_map["Key"] = owner_map["billing_location"] + "_" + owner_map["cust_no"]

        
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


        st.download_button(
            "Download Output Excel",
            data=output.getvalue(),
            file_name="Hours_Recon_Output.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    else:
        st.warning("Please upload all files.")
