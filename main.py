import streamlit as st
import pandas as pd
import re
from io import BytesIO
import numpy as np
import time

st.set_page_config(layout="wide")

# =========================
# FILE READER
# =========================
def read_file(file, header=0, usecols=None):

    try:

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

        else:
            raise ValueError("Unsupported file format")

    except Exception as e:
        raise Exception(f"Unable to read file '{file.name}': {e}")


# =========================
# HEADER
# =========================
st.title("Hours Recon Tool")

st.markdown("### Upload Files")

st.divider()

# =========================
# FILE UPLOADERS
# =========================
col1, col2, col3, col4 = st.columns(4)

with col1:
    uploaded_file_dump = st.file_uploader(
        "Upload Dump File",
        type=["csv", "xlsx"]
    )

    st.caption(
        "Required columns: Order No, Period From, Period To, Invoice dt"
    )

with col2:
    uploaded_file_pillar = st.file_uploader(
        "Upload Pillar File",
        type=["csv", "xlsx"]
    )

    st.caption(
        "Contains operational data such as performed hours, billed hours, deployment details, and invoice mapping"
    )

with col3:
    uploaded_file_owner = st.file_uploader(
        "Upload Owner Mapping File",
        type=["csv", "xlsx"]
    )

    st.caption(
        "Maps billing location and customer to finance owner (branch_finance_lead)"
    )

with col4:
    uploaded_file_attendance = st.file_uploader(
        "Upload Attendance File",
        type=["csv", "xlsx"]
    )

    st.caption(
        "Pivot-style attendance file with Row Labels and attendance date-range columns"
    )

st.divider()

# =========================
# RUN BUTTON
# =========================
run = st.button("Run")

# =========================
# LOG CONTAINER
# =========================
log_container = st.container()

st.divider()

# =========================
# DOCUMENTATION
# =========================
with st.expander("What This Tool Does"):

    st.write("""
    This tool reconciles performed hours vs billed hours across all orders.

    It combines:
    - Billing dump
    - Pillar operational data
    - Attendance records
    - Owner mapping

    Final output provides:
    - Variance between performed and billed hours
    - Attendance comparison
    - HUB and Zone classification
    - Finance ownership tagging
    """)

with st.expander("How to Use"):

    st.write("""
    1. Upload all required files
    2. Click Run
    3. Wait for processing completion
    4. Download the reconciliation report
    """)

with st.expander("Output Details"):

    st.write("""
    Output contains:

    - India Consolidated Sheet
    - HUB-wise sheets

    Key metrics:
    - Total Attendance
    - Total Performed Hours
    - Total Billed Hours
    - Variance
    - Owner Mapping
    - Adjustment columns

    Grouping hierarchy:
    HUB → Location → Zone → Customer → Order → Invoice
    """)

with st.expander("Financial Logic"):

    st.write("""
    Variance = Billed Hours − Performed Hours

    Positive variance:
    Overbilling

    Negative variance:
    Underbilling

    Inter Assignment Adjustment:
    Equal and opposite variances within same order are auto-adjusted.

    Special Handling:
    - Customer Code 7401 → Entire performed hours treated as Office Duty
    - Missing attendance remains blank
    - Owner mapping based on Location + Customer Code
    """)

# =========================
# MAIN PROCESSING
# =========================
if run:

    if not (
        uploaded_file_dump
        and uploaded_file_pillar
        and uploaded_file_owner
        and uploaded_file_attendance
    ):

        st.warning("Please upload all required files.")
        st.stop()

    with log_container:

        status_text = st.empty()
        progress_bar = st.progress(0)

        try:

            # =========================
            # READING FILES
            # =========================
            status_text.info("Reading input files...")
            progress_bar.progress(5)

            try:

                dump = read_file(
                    uploaded_file_dump,
                    header=0,
                    usecols=[
                        "Order No",
                        "Period From",
                        "Period To",
                        "Invoice dt"
                    ]
                )

            except Exception as e:
                st.error(f"Error reading Dump File: {e}")
                st.stop()

            try:

                pillar = read_file(
                    uploaded_file_pillar,
                    header=2,
                    usecols=[
                        "Location",
                        "Customer Code",
                        "Customer Name",
                        "Order No",
                        "Invoice No",
                        "SO Line No",
                        "No of Post",
                        "Deployment Hrs",
                        "WF_TaskID",
                        "Performed Hrs",
                        "Billed Hrs",
                        "Billed Vs Performed",
                        "Contracted Vs Performed",
                        "Billing Pattern",
                        "ERP Cont Hrs",
                        "Saturn Cont Hrs",
                        "Scheduled Hrs"
                    ]
                )

            except Exception as e:
                st.error(f"Error reading Pillar File: {e}")
                st.stop()

            try:

                owner_map = read_file(uploaded_file_owner)

            except Exception as e:
                st.error(f"Error reading Owner Mapping File: {e}")
                st.stop()

            try:

                attendance = read_file(
                    uploaded_file_attendance,
                    header=2
                )

            except Exception as e:
                st.error(f"Error reading Attendance File: {e}")
                st.stop()

            time.sleep(0.2)

            # =========================
            # VALIDATION
            # =========================
            status_text.info("Validating required columns...")
            progress_bar.progress(10)

            required_dump_cols = [
                "Order No",
                "Period From",
                "Period To",
                "Invoice dt"
            ]

            required_pillar_cols = [
                "Location",
                "Customer Code",
                "Customer Name",
                "Order No",
                "Invoice No",
                "SO Line No",
                "WF_TaskID",
                "Performed Hrs",
                "Billed Hrs"
            ]

            required_attendance_cols = [
                "Row Labels"
            ]

            owner_map.columns = (
                owner_map.columns
                .str.strip()
                .str.lower()
                .str.replace(" ", "_")
            )

            required_owner_cols = [
                "billing_location",
                "cust_no",
                "branch_finance_lead"
            ]

            missing_dump = [
                col for col in required_dump_cols
                if col not in dump.columns
            ]

            missing_pillar = [
                col for col in required_pillar_cols
                if col not in pillar.columns
            ]

            missing_attendance = [
                col for col in required_attendance_cols
                if col not in attendance.columns
            ]

            missing_owner = [
                col for col in required_owner_cols
                if col not in owner_map.columns
            ]

            if missing_dump:
                st.error(f"Missing columns in Dump File: {missing_dump}")
                st.stop()

            if missing_pillar:
                st.error(f"Missing columns in Pillar File: {missing_pillar}")
                st.stop()

            if missing_attendance:
                st.error(f"Missing columns in Attendance File: {missing_attendance}")
                st.stop()

            if missing_owner:
                st.error(f"Missing columns in Owner Mapping File: {missing_owner}")
                st.stop()

            time.sleep(0.2)

            # =========================
            # CLEANING
            # =========================
            status_text.info("Cleaning and standardizing datasets...")
            progress_bar.progress(18)

            try:

                str_cols = dump.select_dtypes(include="object").columns

                dump[str_cols] = dump[str_cols].apply(
                    lambda col: col.str.strip()
                )

            except Exception as e:
                st.error(f"Error cleaning Dump File: {e}")
                st.stop()

            try:

                pillar = pillar[
                    pillar["Performed Hrs"] + pillar["Billed Hrs"] > 0
                ]

                str_cols = pillar.select_dtypes(include="object").columns

                pillar[str_cols] = pillar[str_cols].apply(
                    lambda col: col.str.strip()
                )

            except Exception as e:
                st.error(f"Error cleaning Pillar File: {e}")
                st.stop()

            time.sleep(0.2)

            # =========================
            # NORMALIZATION
            # =========================
            status_text.info("Normalizing order numbers and dates...")
            progress_bar.progress(25)

            try:

                def normalize_order(s):

                    return (
                        s.astype(str)
                        .str.strip()
                        .str.replace(" ", "", regex=False)
                        .str.upper()
                    )

                dump["Order No"] = normalize_order(
                    dump["Order No"]
                )

                pillar["Order No"] = normalize_order(
                    pillar["Order No"]
                )

                dump["Period From"] = pd.to_datetime(
                    dump["Period From"],
                    errors="coerce"
                )

                dump["Period To"] = pd.to_datetime(
                    dump["Period To"],
                    errors="coerce"
                )

                dump["Invoice dt"] = pd.to_datetime(
                    dump["Invoice dt"],
                    errors="coerce"
                )

            except Exception as e:
                st.error(f"Error during normalization: {e}")
                st.stop()

            time.sleep(0.2)

            # =========================
            # PREPARE DUMP
            # =========================
            status_text.info("Preparing billing period mapping...")
            progress_bar.progress(32)

            try:

                dump = dump.sort_values(
                    ["Invoice dt", "Period To", "Period From"],
                    ascending=[False, False, False]
                )

                dump_first = dump.drop_duplicates(
                    subset=["Order No"],
                    keep="first"
                ).copy()

                dump_first["Date_Range"] = (
                    dump_first["Period From"]
                    .dt.day.astype("Int64")
                    .astype(str)
                    + "-"
                    + dump_first["Period To"]
                    .dt.day.astype("Int64")
                    .astype(str)
                )

            except Exception as e:
                st.error(f"Error preparing billing periods: {e}")
                st.stop()

            time.sleep(0.2)

            # =========================
            # MERGE DUMP
            # =========================
            status_text.info("Merging billing period data...")
            progress_bar.progress(38)

            try:

                pillar = pillar.merge(
                    dump_first[
                        [
                            "Order No",
                            "Period From",
                            "Period To",
                            "Date_Range"
                        ]
                    ],
                    on="Order No",
                    how="left"
                )

            except Exception as e:
                st.error(f"Error merging dump data: {e}")
                st.stop()

            time.sleep(0.2)

            # =========================
            # ATTENDANCE PROCESSING
            # =========================
            status_text.info("Processing attendance structure...")
            progress_bar.progress(45)

            try:

                def normalize_attendance_col(col):

                    if not isinstance(col, str):
                        return col

                    nums = re.findall(r"\d+", col)

                    if len(nums) == 2:
                        return f"{int(nums[0])}-{int(nums[1])}"

                    return col

                attendance.columns = [
                    normalize_attendance_col(c)
                    for c in attendance.columns
                ]

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
                    .str.replace(".0", "", regex=False)
                    .str.strip()
                )

                pillar["row_key"] = (
                    pillar["Order No"].astype(str).str.strip()
                    + pillar["SO Line No"].astype(str).str.strip()
                )

                attendance["row_key"] = normalize_attendance_row_label(
                    attendance["Row Labels"]
                )

                attendance_long = attendance.melt(
                    id_vars=["row_key"],
                    var_name="Date_Range",
                    value_name="Total Attendance"
                )

            except Exception as e:
                st.error(f"Error processing attendance data: {e}")
                st.stop()

            time.sleep(0.2)

            # =========================
            # ATTENDANCE MERGE
            # =========================
            status_text.info("Merging attendance values...")
            progress_bar.progress(52)

            try:

                pillar = pillar.merge(
                    attendance_long,
                    on=["row_key", "Date_Range"],
                    how="left"
                )

                pillar = pillar.drop(columns=["row_key"])

            except Exception as e:
                st.error(f"Error merging attendance data: {e}")
                st.stop()

            time.sleep(0.2)

            # =========================
            # HUB ZONE MAPPING
            # =========================
            status_text.info("Creating HUB and Zone mapping...")
            progress_bar.progress(60)

            try:

                hub_zone_data = [ ("ALIGRD","Kolkata","Kolkata Zone"), ("ASLGRD","Kolkata","East COC"), ("BBRGRD","Kolkata","Odisha Zone"), ("BBLGRD","Kolkata","Odisha Zone"), ("JAJGRD","Kolkata","Odisha Zone"), ("JHAGRD","Kolkata","Odisha Zone"), ("JASGRD","Kolkata","East COC"), ("PATGRD","Kolkata","East COC"), ("PTNGRD","Kolkata","East COC"), ("BHRGRD","Kolkata","East COC"), ("DALGRD","Kolkata","Kolkata Zone"), ("GUWGRD","Kolkata","East COC"), ("GHTGRD","Kolkata","East COC"), ("HOWGRD","Kolkata","Kolkata Zone"), ("RJHGRD","Kolkata","Kolkata Zone"), ("KOLGRD","Kolkata","Kolkata Zone"), ("SALGRD","Kolkata","Kolkata Zone"), ("SILGRD","Kolkata","East COC"), ("USCGRD","Kolkata","East COC"), ("RAIGRD","Kolkata","East COC"), ("BARGRD","Kolkata","Odisha Zone"), ("ROUGRD","Kolkata","Odisha Zone"), ("JAMGRD","Kolkata","East COC"), ("KONGRD","Kolkata","Kolkata Zone"), ("BHLGRD","NCR","North COC"), ("IDRGRD","NCR","North COC"), ("CP1GRD","NCR","Delhi Zone"), ("CP2GRD","NCR","Delhi Zone"), ("DROGRD","NCR","Delhi Zone"), ("EMBGRD","NCR","Delhi Zone"), ("FRMGRD","NCR","Delhi Zone"), ("PSPGRD","NCR","Delhi Zone"), ("GOLGRD","NCR","Delhi Zone"), ("VVRGRD","NCR","Delhi Zone"), ("USEGRD","NCR","North COC"), ("GHAGRD","NCR","Noida Zone"), ("LKWGRD","NCR","North COC"), ("MRTGRD","NCR","North COC"), ("NDAGRD","NCR","Noida Zone"), ("NDGGRD","NCR","Noida Zone"), ("CHDGRD","NCR","North COC"), ("CROGRD","NCR","North COC"), ("DDNGRD","NCR","North COC"), ("UTKGRD","NCR","North COC"), ("JMUGRD","NCR","North COC"), ("JAUGRD","NCR","North COC"), ("JNKGRD","NCR","North COC"), ("PRWGRD","NCR","North COC"), ("PWNGRD","NCR","North COC"), ("RUDGRD","NCR","North COC"), ("RPRGRD","NCR","North COC"), ("FBDGRD","NCR","Gurgaon Zone"), ("GGNGRD","NCR","Gurgaon Zone"), ("GNBGRD","NCR","Gurgaon Zone"), ("GNSGRD","NCR","Gurgaon Zone"), ("MNSGRD","NCR","Gurgaon Zone"), ("SPTGRD","NCR","North COC"), ("JALGRD","NCR","North COC"), ("LUDGRD","NCR","North COC"), ("JPRGRD","NCR","North COC"), ("DHRGRD","NCR","North COC"), ("JARGRD","NCR","North COC"), ("UDRGRD","NCR","North COC"), ("DUNGRD","NCR","North COC"), ("SNPGRD","NCR","North COC"), ("TYMGRD","NCR","Delhi Zone"), ("TEPGRD","NCR","Noida Zone"), ("OKLGRD","NCR","Delhi Zone"), ("HUBGRD","South","South COC"), ("BELGRD","South","South COC"), ("BANGRD","South","Bangalore Zone"), ("BLRGRD","South","Bangalore Zone"), ("DOMGRD","South","Bangalore Zone"), ("ELEGRD","South","Bangalore Zone"), ("HOOGRD","South","Bangalore Zone"), ("ORRGRD","South","Bangalore Zone"), ("SARGRD","South","Bangalore Zone"), ("VASGRD","South","Bangalore Zone"), ("WHTGRD","South","Bangalore Zone"), ("YELGRD","South","Bangalore Zone"), ("YESGRD","South","Bangalore Zone"), ("MNGGRD","South","South COC"), ("MYOGRD","South","South COC"), ("MYSGRD","South","South COC"), ("HOPGRD","South","Bangalore Zone"), ("COMGRD","South","South COC"), ("CBTGRD","South","South COC"), ("ADYGRD","South","Chennai Zone"), ("ANNGRD","South","Chennai Zone"), ("CHNGRD","South","Chennai Zone"), ("GUIGRD","South","Chennai Zone"), ("MMNGRD","South","Chennai Zone"), ("NUGGRD","South","Chennai Zone"), ("SRIGRD","South","Chennai Zone"), ("COCGRD","South","South COC"), ("PONGRD","South","South COC"), ("MADGRD","South","South COC"), ("TRVGRD","South","South COC"), ("SIRGRD","South","Chennai Zone"), ("SLMGRD","South","South COC"), ("HYDGRD","South","Hyderabad Zone"), ("HYRGRD","South","Hyderabad Zone"), ("HYTGRD","South","Hyderabad Zone"), ("JBHGRD","South","Hyderabad Zone"), ("MHPGRD","South","Hyderabad Zone"), ("VIGGRD","South","South COC"), ("VIZGRD","South","South COC"), ("VJWGRD","South","South COC"), ("VWDGRD","South","South COC"), ("ANPGRD","South","Hyderabad Zone"), ("HSRGRD","South","South COC"), ("AHDGRD","Mumbai","West COC"), ("AHMGRD","Mumbai","West COC"), ("AINGRD","Mumbai","West COC"), ("ANKGRD","Mumbai","West COC"), ("BODGRD","Mumbai","West COC"), ("JNAGRD","Mumbai","West COC"), ("MLDGRD","Mumbai","Mumbai Zone"), ("MNMGRD","Mumbai","Mumbai Zone"), ("MNVGRD","Mumbai","Mumbai Zone"), ("MSOGRD","Mumbai","Mumbai Zone"), ("MUCGRD","Mumbai","Mumbai Zone"), ("MUMGRD","Mumbai","Mumbai Zone"), ("MUSGRD","Mumbai","West COC"), ("GONGRD","Mumbai","West COC"), ("GOAGRD","Mumbai","West COC"), ("NAGGRD","Mumbai","West COC"), ("PROGRD","Mumbai","West COC"), ("PNEGRD","Mumbai","Pune Zone"), ("PNHGRD","Mumbai","Pune Zone"), ("RJGGRD","Mumbai","Pune Zone"), ("PNRGRD","Mumbai","Pune Zone"), ("PUWGRD","Mumbai","Pune Zone"), ("PUNGRD","Mumbai","Pune Zone"), ("DEUGRD","Mumbai","West COC"), ("PNIGRD","Mumbai","Pune Zone"), ("MONGRD","Mumbai","Mumbai Zone"), ("MUSMSP","Mumbai","Mumbai Zone"), ("CORMSP","Mumbai","Mumbai Zone"), ("INVGRD","HeadOffice","Head Office"), ("OTHGRD","HeadOffice","Head Office"), ("PSOGRD","HeadOffice","Head Office"), ("TRGGRD","HeadOffice","Head Office"), ("CORGRD","HeadOffice","Head Office"), ("HO","HeadOffice","Head Office"), ("HIMGRD","NCR","North COC") ]


                hub_zone = pd.DataFrame(
                    hub_zone_data,
                    columns=["Location","HUB","Zone"]
                )

                hub_zone["Location"] = normalize_order(
                    hub_zone["Location"]
                )

                pillar["Location"] = normalize_order(
                    pillar["Location"]
                )

                pillar = pillar.merge(
                    hub_zone,
                    on="Location",
                    how="left"
                )

            except Exception as e:
                st.error(f"Error during HUB/Zone mapping: {e}")
                st.stop()

            time.sleep(0.2)

            # =========================
            # OWNER MAPPING
            # =========================
            status_text.info("Mapping finance owners...")
            progress_bar.progress(68)

            try:

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

                pillar["Key"] = (
                    pillar["Location"]
                    + "_"
                    + pillar["Customer Code"]
                )

                owner_map["Key"] = (
                    owner_map["billing_location"]
                    + "_"
                    + owner_map["cust_no"]
                )

                owner_map = owner_map.drop_duplicates(
                    subset="Key",
                    keep="last"
                )

                pillar = pillar.merge(
                    owner_map[
                        ["Key", "branch_finance_lead"]
                    ],
                    on="Key",
                    how="left"
                )

                pillar = pillar.rename(
                    columns={
                        "branch_finance_lead": "Owner"
                    }
                )

            except Exception as e:
                st.error(f"Error during owner mapping: {e}")
                st.stop()

            time.sleep(0.2)

            # =========================
            # PIVOT
            # =========================
            status_text.info("Creating reconciliation pivot...")
            progress_bar.progress(78)

            try:

                pivot = (
                    pillar.groupby([
                        "HUB",
                        "Location",
                        "Zone",
                        "Owner",
                        "Customer Code",
                        "Customer Name",
                        "Order No",
                        "Invoice No",
                        "WF_TaskID",
                        "Period From",
                        "Period To"
                    ], dropna=False)[
                        [
                            "Total Attendance",
                            "Performed Hrs",
                            "Billed Hrs"
                        ]
                    ]
                    .sum()
                    .reset_index()
                )

            except Exception as e:
                st.error(f"Error creating pivot: {e}")
                st.stop()

            # =========================
            # FINAL CALCULATIONS
            # =========================
            status_text.info("Calculating variances and adjustments...")
            progress_bar.progress(85)

            try:

                pivot = pivot.rename(columns={
                    "Performed Hrs": "Total Performed",
                    "Billed Hrs": "Total Billed"
                })

                pivot["Var. Performed Vs. Billed"] = (
                    pivot["Total Billed"]
                    - pivot["Total Performed"]
                )

                pivot["Office Duty/Office Patrolling"] = np.where(
                    pivot["Customer Code"].astype(str) == "7401",
                    pivot["Total Performed"],
                    ""
                )

                extra_cols = [
                    "Excess Paid",
                    "Reliever duty",
                    "Excess billing",
                    "Short billing",
                    "Disciplinary Deduction",
                    "Short / Missing Roster",
                    "Inter assignment adjustment",
                    "Indirect Hours Not Captured in Saturn",
                    "Training & OJT",
                    "Complimentary Hrs.",
                    "Billing Cycle/ hours calculation other than calendar month",
                    "Bill Hrs should being Cycle",
                    "Diff with bill cycle should be",
                    "Total ( B )",
                    "Check (A - B)",
                    "BFL Remarks",
                    "SSC Query (If Any)"
                ]

                for col in extra_cols:
                    pivot[col] = pd.NA

            except Exception as e:
                st.error(f"Error during variance calculations: {e}")
                st.stop()

            # =========================
            # COLUMN ORDER
            # =========================
            status_text.info("Preparing final output structure...")
            progress_bar.progress(90)

            try:

                pivot = pivot[[
                    "HUB",
                    "Location",
                    "Zone",
                    "Owner",
                    "Customer Code",
                    "Customer Name",
                    "Order No",
                    "Invoice No",
                    "WF_TaskID",
                    "Period From",
                    "Period To",
                    "Total Attendance",
                    "Total Performed",
                    "Total Billed",
                    "Var. Performed Vs. Billed",
                    "Office Duty/Office Patrolling",
                    "Excess Paid",
                    "Reliever duty",
                    "Excess billing",
                    "Short billing",
                    "Disciplinary Deduction",
                    "Short / Missing Roster",
                    "Inter assignment adjustment",
                    "Indirect Hours Not Captured in Saturn",
                    "Training & OJT",
                    "Complimentary Hrs.",
                    "Billing Cycle/ hours calculation other than calendar month",
                    "Bill Hrs should being Cycle",
                    "Diff with bill cycle should be",
                    "Total ( B )",
                    "Check (A - B)",
                    "BFL Remarks",
                    "SSC Query (If Any)"
                ]]

                pivot["Inter assignment adjustment"] = pd.NA

                pivot["Var. Performed Vs. Billed"] = pd.to_numeric(
                    pivot["Var. Performed Vs. Billed"],
                    errors="coerce"
                )

            except Exception as e:
                st.error(f"Error preparing final columns: {e}")
                st.stop()

            # =========================
            # INTER ASSIGNMENT
            # =========================
            status_text.info("Applying inter-assignment adjustments...")
            progress_bar.progress(94)

            try:

                seen = {}

                for idx, row in pivot.iterrows():

                    order = str(row["Order No"]).strip()

                    val = pd.to_numeric(
                        row["Var. Performed Vs. Billed"],
                        errors="coerce"
                    )

                    if pd.isna(val) or val == 0:
                        continue

                    key = (order, round(val, 6))

                    reverse_key = (
                        order,
                        round(-val, 6)
                    )

                    if reverse_key in seen:

                        prev_idx = seen[reverse_key]

                        pivot.loc[
                            idx,
                            "Inter assignment adjustment"
                        ] = -val

                        pivot.loc[
                            prev_idx,
                            "Inter assignment adjustment"
                        ] = -pivot.loc[
                            prev_idx,
                            "Var. Performed Vs. Billed"
                        ]

                        del seen[reverse_key]

                    else:

                        seen[key] = idx

            except Exception as e:
                st.error(f"Error during inter-assignment adjustment logic: {e}")
                st.stop()

            # =========================
            # EXCEL OUTPUT
            # =========================
            status_text.info("Generating Excel output...")
            progress_bar.progress(98)

            try:

                india_conso = pivot.copy()

                output = BytesIO()

                with pd.ExcelWriter(
                    output,
                    engine="xlsxwriter"
                ) as writer:

                    india_conso.to_excel(
                        writer,
                        sheet_name="India Conso",
                        index=False
                    )

                    for hub in india_conso["HUB"].dropna().unique():

                        india_conso[
                            india_conso["HUB"] == hub
                        ].to_excel(
                            writer,
                            sheet_name=str(hub)[:31],
                            index=False
                        )

            except Exception as e:
                st.error(f"Error generating Excel output: {e}")
                st.stop()

            # =========================
            # SUCCESS
            # =========================
            progress_bar.progress(100)

            status_text.success(
                "Processing completed successfully."
            )

            st.download_button(
                label="Download Reconciliation Report",
                data=output.getvalue(),
                file_name="Hours_Recon_Output.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

        except Exception as e:

            st.error(f"Unexpected processing error: {e}")
            st.stop()
