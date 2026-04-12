import streamlit as st
import pandas as pd
import re
from io import BytesIO
import numpy as np

st.set_page_config(layout="wide")

# =========================
# FILE READER (UNCHANGED)
# =========================
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

# =========================
# HEADER
# =========================
st.title("Hours Recon Tool")
st.markdown("### 📂 Upload Files")

st.divider()

# =========================
# FILE UPLOADERS (TOP SECTION)
# =========================
col1, col2, col3, col4 = st.columns(4)

with col1:
    uploaded_file_dump = st.file_uploader("Upload Dump File", type=["csv","xlsx"])
    st.caption("Must contain columns: Order No, Period From, Period To, Invoice dt")

with col2:
    uploaded_file_pillar = st.file_uploader("Upload Pillar File", type=["csv","xlsx"])
    st.caption("Contains operational data such as performed hours, billed hours, and deployment details")

with col3:
    uploaded_file_owner = st.file_uploader("Upload Owner Mapping File", type=["csv","xlsx"])
    st.caption("Maps billing location and customer to finance owner (branch_finance_lead)")

with col4:
    uploaded_file_attendance = st.file_uploader("Upload Attendance File", type=["csv","xlsx"])
    st.caption("Pivot-style file with 'Row Labels' and date range columns showing attendance totals")

st.divider()

# =========================
# RUN BUTTON
# =========================
run = st.button("▶️ Run Processing")
log_container = st.container()

st.divider()

# =========================
# DOCUMENTATION SECTION
# =========================

with st.expander("📘 What This Tool Does"):
    st.write("""
    This tool reconciles **performed hours vs billed hours** across all orders.

    It combines multiple datasets:
    - Billing dump
    - Operational (pillar) data
    - Attendance records
    - Owner mapping

    Final output provides:
    - Variance between billed and performed hours
    - Attendance comparison
    - Finance ownership tagging
    - HUB and Zone classification
    """)

with st.expander("🛠️ How to Use"):
    st.write("""
    1. Upload all 4 required files at the top
    2. Ensure column names match expected structure
    3. Click **Run Processing**
    4. Download the generated Excel output

    ⚠️ Important:
    - Files must not be modified structurally
    - Date formats should be valid
    - Order numbers should be consistent across files
    """)

with st.expander("📊 Output Details"):
    st.write("""
    The output Excel contains:

    • **India Conso Sheet**
      - Consolidated data across all HUBs

    • **HUB-wise Sheets**
      - Separate sheets per HUB

    Key columns include:
    - Total Attendance
    - Total Performed Hours
    - Total Billed Hours
    - Variance (Performed vs Billed)
    - Owner (Finance Mapping)

    Additional columns are provided for manual finance adjustments.
    """)

with st.expander("💰 Financial Logic"):
    st.write("""
    **Core Formula:**

    Variance = Billed Hours − Performed Hours

    - Positive → Overbilling
    - Negative → Underbilling

    **Inter Assignment Adjustment Logic:**
    - If two entries within same Order have equal and opposite variance:
      They are auto-adjusted

    **Special Handling:**
    - Customer Code 7401 → Entire performed hours marked as Office Duty
    - Missing attendance → remains blank (no forced assumptions)
    - Owner mapping based on Location + Customer Code key

    **Classification:**
    - Data grouped by:
      HUB → Location → Zone → Customer → Order → Invoice

    This ensures finance-level reconciliation visibility.
    """)

# =========================
# MAIN PROCESSING (UNCHANGED)
# =========================
if run:

    if uploaded_file_dump and uploaded_file_pillar and uploaded_file_owner and uploaded_file_attendance:

        log_container.write("Reading files...")

        dump = read_file( uploaded_file_dump, header=2, usecols=["Order No","Period From","Period To","Invoice dt"] )

        pillar = read_file( uploaded_file_pillar, header=2, usecols=[ "Location","Customer Code","Customer Name","Order No","Invoice No", "SO Line No","No of Post","Deployment Hrs","WF_TaskID", "Performed Hrs","Billed Hrs","Billed Vs Performed", "Contracted Vs Performed","Billing Pattern", "ERP Cont Hrs","Saturn Cont Hrs","Scheduled Hrs" ] )

        owner_map = read_file(uploaded_file_owner)

        attendance = read_file(uploaded_file_attendance, header=2)

        log_container.write("Cleaning data...")
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

        log_container.write("Processing attendance...")

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

        log_container.write("Creating HUB Zone mapping...")

        # (HUB ZONE DATA UNCHANGED)
        hub_zone_data = [ ("ALIGRD","Kolkata","Kolkata Zone"), ("ASLGRD","Kolkata","East COC"), ("BBRGRD","Kolkata","Odisha Zone"), ("BBLGRD","Kolkata","Odisha Zone"), ("JAJGRD","Kolkata","Odisha Zone"), ("JHAGRD","Kolkata","Odisha Zone"), ("JASGRD","Kolkata","East COC"), ("PATGRD","Kolkata","East COC"), ("PTNGRD","Kolkata","East COC"), ("BHRGRD","Kolkata","East COC"), ("DALGRD","Kolkata","Kolkata Zone"), ("GUWGRD","Kolkata","East COC"), ("GHTGRD","Kolkata","East COC"), ("HOWGRD","Kolkata","Kolkata Zone"), ("RJHGRD","Kolkata","Kolkata Zone"), ("KOLGRD","Kolkata","Kolkata Zone"), ("SALGRD","Kolkata","Kolkata Zone"), ("SILGRD","Kolkata","East COC"), ("USCGRD","Kolkata","East COC"), ("RAIGRD","Kolkata","East COC"), ("BARGRD","Kolkata","Odisha Zone"), ("ROUGRD","Kolkata","Odisha Zone"), ("JAMGRD","Kolkata","East COC"), ("KONGRD","Kolkata","Kolkata Zone"), ("BHLGRD","NCR","North COC"), ("IDRGRD","NCR","North COC"), ("CP1GRD","NCR","Delhi Zone"), ("CP2GRD","NCR","Delhi Zone"), ("DROGRD","NCR","Delhi Zone"), ("EMBGRD","NCR","Delhi Zone"), ("FRMGRD","NCR","Delhi Zone"), ("PSPGRD","NCR","Delhi Zone"), ("GOLGRD","NCR","Delhi Zone"), ("VVRGRD","NCR","Delhi Zone"), ("USEGRD","NCR","North COC"), ("GHAGRD","NCR","Noida Zone"), ("LKWGRD","NCR","North COC"), ("MRTGRD","NCR","North COC"), ("NDAGRD","NCR","Noida Zone"), ("NDGGRD","NCR","Noida Zone"), ("CHDGRD","NCR","North COC"), ("CROGRD","NCR","North COC"), ("DDNGRD","NCR","North COC"), ("UTKGRD","NCR","North COC"), ("JMUGRD","NCR","North COC"), ("JAUGRD","NCR","North COC"), ("JNKGRD","NCR","North COC"), ("PRWGRD","NCR","North COC"), ("PWNGRD","NCR","North COC"), ("RUDGRD","NCR","North COC"), ("RPRGRD","NCR","North COC"), ("FBDGRD","NCR","Gurgaon Zone"), ("GGNGRD","NCR","Gurgaon Zone"), ("GNBGRD","NCR","Gurgaon Zone"), ("GNSGRD","NCR","Gurgaon Zone"), ("MNSGRD","NCR","Gurgaon Zone"), ("SPTGRD","NCR","North COC"), ("JALGRD","NCR","North COC"), ("LUDGRD","NCR","North COC"), ("JPRGRD","NCR","North COC"), ("DHRGRD","NCR","North COC"), ("JARGRD","NCR","North COC"), ("UDRGRD","NCR","North COC"), ("DUNGRD","NCR","North COC"), ("SNPGRD","NCR","North COC"), ("TYMGRD","NCR","Delhi Zone"), ("TEPGRD","NCR","Noida Zone"), ("OKLGRD","NCR","Delhi Zone"), ("HUBGRD","South","South COC"), ("BELGRD","South","South COC"), ("BANGRD","South","Bangalore Zone"), ("BLRGRD","South","Bangalore Zone"), ("DOMGRD","South","Bangalore Zone"), ("ELEGRD","South","Bangalore Zone"), ("HOOGRD","South","Bangalore Zone"), ("ORRGRD","South","Bangalore Zone"), ("SARGRD","South","Bangalore Zone"), ("VASGRD","South","Bangalore Zone"), ("WHTGRD","South","Bangalore Zone"), ("YELGRD","South","Bangalore Zone"), ("YESGRD","South","Bangalore Zone"), ("MNGGRD","South","South COC"), ("MYOGRD","South","South COC"), ("MYSGRD","South","South COC"), ("HOPGRD","South","Bangalore Zone"), ("COMGRD","South","South COC"), ("CBTGRD","South","South COC"), ("ADYGRD","South","Chennai Zone"), ("ANNGRD","South","Chennai Zone"), ("CHNGRD","South","Chennai Zone"), ("GUIGRD","South","Chennai Zone"), ("MMNGRD","South","Chennai Zone"), ("NUGGRD","South","Chennai Zone"), ("SRIGRD","South","Chennai Zone"), ("COCGRD","South","South COC"), ("PONGRD","South","South COC"), ("MADGRD","South","South COC"), ("TRVGRD","South","South COC"), ("SIRGRD","South","Chennai Zone"), ("SLMGRD","South","South COC"), ("HYDGRD","South","Hyderabad Zone"), ("HYRGRD","South","Hyderabad Zone"), ("HYTGRD","South","Hyderabad Zone"), ("JBHGRD","South","Hyderabad Zone"), ("MHPGRD","South","Hyderabad Zone"), ("VIGGRD","South","South COC"), ("VIZGRD","South","South COC"), ("VJWGRD","South","South COC"), ("VWDGRD","South","South COC"), ("ANPGRD","South","Hyderabad Zone"), ("HSRGRD","South","South COC"), ("AHDGRD","Mumbai","West COC"), ("AHMGRD","Mumbai","West COC"), ("AINGRD","Mumbai","West COC"), ("ANKGRD","Mumbai","West COC"), ("BODGRD","Mumbai","West COC"), ("JNAGRD","Mumbai","West COC"), ("MLDGRD","Mumbai","Mumbai Zone"), ("MNMGRD","Mumbai","Mumbai Zone"), ("MNVGRD","Mumbai","Mumbai Zone"), ("MSOGRD","Mumbai","Mumbai Zone"), ("MUCGRD","Mumbai","Mumbai Zone"), ("MUMGRD","Mumbai","Mumbai Zone"), ("MUSGRD","Mumbai","West COC"), ("GONGRD","Mumbai","West COC"), ("GOAGRD","Mumbai","West COC"), ("NAGGRD","Mumbai","West COC"), ("PROGRD","Mumbai","West COC"), ("PNEGRD","Mumbai","Pune Zone"), ("PNHGRD","Mumbai","Pune Zone"), ("RJGGRD","Mumbai","Pune Zone"), ("PNRGRD","Mumbai","Pune Zone"), ("PUWGRD","Mumbai","Pune Zone"), ("PUNGRD","Mumbai","Pune Zone"), ("DEUGRD","Mumbai","West COC"), ("PNIGRD","Mumbai","Pune Zone"), ("MONGRD","Mumbai","Mumbai Zone"), ("MUSMSP","Mumbai","Mumbai Zone"), ("CORMSP","Mumbai","Mumbai Zone"), ("INVGRD","HeadOffice","Head Office"), ("OTHGRD","HeadOffice","Head Office"), ("PSOGRD","HeadOffice","Head Office"), ("TRGGRD","HeadOffice","Head Office"), ("CORGRD","HeadOffice","Head Office"), ("HO","HeadOffice","Head Office"), ("HIMGRD","NCR","North COC") ]

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

        log_container.write("Owner mapping...")

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

        pillar["Key"] = pillar["Location"] + "_" + pillar["Customer Code"]
        owner_map["Key"] = owner_map["billing_location"] + "_" + owner_map["cust_no"]
        owner_map = owner_map.drop_duplicates(subset="Key", keep="last")

        pillar = pillar.merge(
            owner_map[["Key", "branch_finance_lead"]],
            on="Key",
            how="left"
        )

        pillar = pillar.rename(columns={"branch_finance_lead": "Owner"})

        log_container.write("Creating pivot...")

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

        pivot["Inter assignment adjustment"] = pd.NA
        pivot["Var. Performed Vs. Billed"] = pd.to_numeric( pivot["Var. Performed Vs. Billed"], errors="coerce" )

        seen = {}
        for idx, row in pivot.iterrows():
            order = str(row["Order No"]).strip()
            val = pd.to_numeric(row["Var. Performed Vs. Billed"], errors="coerce")

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

        log_container.write("Preparing Excel output...")

        output = BytesIO()

        with pd.ExcelWriter(output, engine="xlsxwriter") as writer:

            india_conso.to_excel(writer, sheet_name="India Conso", index=False)

            for hub in india_conso["HUB"].dropna().unique():
                india_conso[india_conso["HUB"] == hub].to_excel(
                    writer,
                    sheet_name=str(hub)[:31],
                    index=False
                )

        log_container.success("Processing complete ✅")


        log_container.download_button(
            "📥 Download Reconciliation Report",
            data=output.getvalue(),
            file_name="Hours_Recon_Output.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    else:
        st.warning("Please upload all required files.")
