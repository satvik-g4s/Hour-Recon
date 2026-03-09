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

        # ----- Dump -----
        dump = read_file(
            uploaded_file_dump,
            header=2,
            usecols=["Order No","Period From","Period To","Invoice dt"]
        )

        # ----- Pillar -----
        pillar = read_file(
            uploaded_file_pillar,
            header=2,
            usecols=[
                "Location","Customer Code","Customer Name","Order No","Invoice No",
                "SO Line No","No of Post","Deployment Hrs","WF_TaskID",
                "Performed Hrs","Billed Hrs","Billed Vs Performed",
                "Contracted Vs Performed","Billing Pattern",
                "ERP Cont Hrs","Saturn Cont Hrs","Scheduled Hrs"
            ]
        )

        # ----- Owner map -----
        owner_map = read_file(uploaded_file_owner)
        st.write("Owner map columns (before cleaning):", owner_map.columns.tolist())
        owner_map.columns = (
            owner_map.columns
            .str.strip()
            .str.lower()
            .str.replace(" ", "_")
        )
        st.write("Owner map columns (after cleaning):", owner_map.columns.tolist())

        # ----- Basic cleaning of dump and pillar -----
        str_cols = dump.select_dtypes(include="object").columns
        dump[str_cols] = dump[str_cols].apply(lambda col: col.str.strip())

        # Keep only rows with positive performed or billed hours (optional)
        pillar = pillar[pillar["Performed Hrs"] + pillar["Billed Hrs"] > 0]

        str_cols = pillar.select_dtypes(include="object").columns
        pillar[str_cols] = pillar[str_cols].apply(lambda col: col.str.strip())

        # Normalize order numbers
        def normalize_order(s):
            return (
                s.astype(str)
                .str.strip()
                .str.replace(" ", "", regex=False)
                .str.upper()
            )

        dump["Order No"] = normalize_order(dump["Order No"])
        pillar["Order No"] = normalize_order(pillar["Order No"])

        # Convert dates in dump
        dump["Period From"] = pd.to_datetime(dump["Period From"], errors="coerce")
        dump["Period To"]   = pd.to_datetime(dump["Period To"], errors="coerce")

        # Take the first occurrence per Order No (latest Invoice dt)
        dump = dump.sort_values(by="Invoice dt", ascending=False)
        dump_first = dump.drop_duplicates(subset=["Order No"], keep="first").copy()

        # Create Date_Range using day numbers (original logic – we'll keep it but not use for attendance)
        dump_first["Date_Range"] = (
            dump_first["Period From"].dt.day.astype("Int64").astype(str)
            + "-"
            + dump_first["Period To"].dt.day.astype("Int64").astype(str)
        )

        # Merge dump dates into pillar
        pillar = pillar.merge(
            dump_first[["Order No","Period From","Period To","Date_Range"]],
            on="Order No",
            how="left"
        )

        # ------------------------------------------------------------
        # ATTENDANCE – VBA‑style dictionary lookup
        # ------------------------------------------------------------
        st.write("Processing attendance (VBA‑style)...")

        # Read the two header rows (period from / to)
        date_headers = pd.read_excel(
            uploaded_file_attendance,
            header=None,
            nrows=2
        )

        # Read the data rows (starting from row 3, i.e. skiprows=2)
        att_data = pd.read_excel(
            uploaded_file_attendance,
            header=None,
            skiprows=2
        )

        # Build lookup dictionary
        attendance_dict = {}
        # Column 0 = Row Labels (Key2), columns 1..n = date ranges
        for col_idx in range(1, len(date_headers.columns)):
            from_date = date_headers.iloc[0, col_idx]
            to_date   = date_headers.iloc[1, col_idx]

            if pd.isna(from_date) or pd.isna(to_date):
                continue

            # Format exactly as VBA: "dd-mmm-yy"
            try:
                from_str = pd.to_datetime(from_date).strftime("%d-%b-%y")
                to_str   = pd.to_datetime(to_date).strftime("%d-%b-%y")
            except:
                continue

            for row_idx in range(len(att_data)):
                key2 = att_data.iloc[row_idx, 0]
                if pd.isna(key2):
                    continue

                # Normalize Key2: uppercase, no spaces, no hyphens, no ".0"
                key2_norm = str(key2).upper().strip()
                key2_norm = key2_norm.replace(" ", "").replace("-", "").replace(".0", "")

                lookup_key = f"{key2_norm}|{from_str}|{to_str}"
                att_value = att_data.iloc[row_idx, col_idx]

                if pd.notna(att_value):
                    attendance_dict[lookup_key] = att_value

        st.write(f"Built dictionary with {len(attendance_dict)} entries")

        # Prepare pillar for lookup
        def clean_key(x):
            if pd.isna(x):
                return ""
            return str(x).upper().strip().replace(" ", "").replace("-", "").replace(".0", "")

        pillar["Key2_VBA"] = (
            pillar["Order No"].apply(clean_key) +
            pillar["SO Line No"].apply(clean_key)
        )

        # Format dates exactly as VBA does
        pillar["From_VBA"] = pd.to_datetime(pillar["Period From"]).dt.strftime("%d-%b-%y")
        pillar["To_VBA"]   = pd.to_datetime(pillar["Period To"]).dt.strftime("%d-%b-%y")

        pillar["Lookup_VBA"] = (
            pillar["Key2_VBA"] + "|" +
            pillar["From_VBA"] + "|" +
            pillar["To_VBA"]
        )

        pillar["Total Attendance"] = pillar["Lookup_VBA"].map(attendance_dict)

        # Drop temporary columns
        pillar.drop(columns=["Key2_VBA", "From_VBA", "To_VBA", "Lookup_VBA"], inplace=True)

        matched = pillar["Total Attendance"].notna().sum()
        total = len(pillar)
        st.write(f"Matched: {matched}/{total} rows ({matched/total*100:.1f}%)")
        # ------------------------------------------------------------

        # ----- HUB Zone mapping -----
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

        hub_zone = pd.DataFrame(hub_zone_data, columns=["Location","HUB","Zone"])
        hub_zone["Location"] = normalize_order(hub_zone["Location"])
        pillar["Location"]   = normalize_order(pillar["Location"])

        pillar = pillar.merge(hub_zone, on="Location", how="left")

        # ----- Owner mapping -----
        st.write("Owner mapping...")

        # Prepare owner map keys
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

        owner_map["cust_no"] = (
            owner_map["cust_no"]
            .astype(str)
            .str.strip()
            .str.replace(".0", "", regex=False)
            .str.upper()
        )
        pillar["Customer Code"] = (
            pillar["Customer Code"]
            .astype(str)
            .str.strip()
            .str.replace(".0", "", regex=False)
            .str.upper()
        )

        # Create composite keys
        pillar["Key"] = pillar["Location"] + "_" + pillar["Customer Code"]
        owner_map["Key"] = owner_map["billing_location"] + "_" + owner_map["cust_no"]

        # Merge
        pillar = pillar.merge(
            owner_map[["Key", "branch_finance_lead"]],
            on="Key",
            how="left"
        )
        pillar = pillar.rename(columns={"branch_finance_lead": "Owner"})

        # ----- Pivot (aggregation) -----
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
            "Performed Hrs": "Total Performed",
            "Billed Hrs": "Total Billed"
        })

        pivot["Var. Performed Vs. Billed"] = (
            pivot["Total Billed"] - pivot["Total Performed"]
        )

        pivot["Office Duty/Office Patrolling"] = np.where(
            pivot["Customer Code"].astype(str) == "7401",
            pivot["Total Performed"],
            ""
        )

        # Add extra columns (empty for now)
        extra_cols = [
            "Excess Paid", "Reliever duty", "Excess billing",
            "Short billing", "Disciplinary Deduction",
            "Short / Missing Roster", "Inter assignment adjustment",
            "Indirect Hours Not Captured in Saturn",
            "Training & OJT", "Complimentary Hrs.",
            "Billing Cycle/ hours calculation other than calendar month",
            "Bill Hrs should being Cycle",
            "Diff with bill cycle should be",
            "Total ( B )", "Check (A - B)",
            "BFL Remarks", "SSC Query (If Any)"
        ]
        for col in extra_cols:
            pivot[col] = pd.NA

        # Reorder columns to match desired output
        pivot = pivot[[
            "HUB", "Location", "Zone", "Owner",
            "Customer Code", "Customer Name",
            "Order No", "Invoice No", "WF_TaskID",
            "Period From", "Period To",
            "Total Attendance", "Total Performed", "Total Billed",
            "Var. Performed Vs. Billed",
            "Office Duty/Office Patrolling",
            "Excess Paid", "Reliever duty", "Excess billing",
            "Short billing", "Disciplinary Deduction",
            "Short / Missing Roster", "Inter assignment adjustment",
            "Indirect Hours Not Captured in Saturn",
            "Training & OJT", "Complimentary Hrs.",
            "Billing Cycle/ hours calculation other than calendar month",
            "Bill Hrs should being Cycle",
            "Diff with bill cycle should be",
            "Total ( B )", "Check (A - B)",
            "BFL Remarks", "SSC Query (If Any)"
        ]]

        # ----- Inter assignment adjustment (match cancelling pairs) -----
        pivot["Inter assignment adjustment"] = ""  # will be filled with formulas later
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

        # ----- India Conso and Hub sheets -----
        india_conso = pivot.copy()
        st.write("Creating HUB sheets...")
        hub_sheets = {}
        for hub in india_conso["HUB"].dropna().unique():
            hub_sheets[hub] = india_conso[india_conso["HUB"] == hub]

        # ----- Write to Excel -----
        st.write("Preparing Excel output...")
        output = BytesIO()
        with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
            india_conso.to_excel(writer, sheet_name="India Conso", index=False)
            for hub, df in hub_sheets.items():
                df.to_excel(writer, sheet_name=str(hub)[:31], index=False)

        st.download_button(
            "Download Output Excel",
            data=output.getvalue(),
            file_name="Hours_Recon_Output.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    else:
        st.warning("Please upload all files.")
