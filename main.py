import streamlit as st
import pandas as pd
import numpy as np

st.set_page_config(layout="wide")

st.title("Hours Recon Automation (VBA → Pandas Version)")

# ==============================
# 7 INPUTS
# ==============================

col1, col2 = st.columns(2)

with col1:
    master_file = st.file_uploader(" Upload Pillar Report", type=["xlsx", "csv"])
    master_date_file = st.file_uploader(" Upload Invoice Dump", type=["xlsx", "csv"])

with col2:
    owner_file = st.file_uploader(" Upload Owner Data", type=["xlsx", "csv"])
    billing_file = st.file_uploader(" Upload Billing Cycle Attendance", type=["xlsx", "csv"])
special_customer = 7401    

# ==============================
# HELPER FUNCTION
# ==============================

def read_file(file):
    if file.name.endswith("csv"):
        return pd.read_csv(file)
    return pd.read_excel(file)

# ==============================
# MAIN PROCESS
# ==============================

if st.button("🚀 Run Full Automation"):

    # ==============================
    # READ FILES
    # ==============================
    
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
    
    df_hub_zone = pd.DataFrame(hub_zone_data, columns=["Location","HUB","Zone"])

    
    df_master = read_file(master_file)
    df_master_date = read_file(master_date_file)
    df_billing = read_file(billing_file)
    df_owner = read_file(owner_file)

    # Trim all strings
    df_master = df_master.applymap(lambda x: x.strip() if isinstance(x, str) else x)
    df_master_date = df_master_date.applymap(lambda x: x.strip() if isinstance(x, str) else x)

    # ==============================
    # MACRO 1 → CREATE DATA
    # ==============================

    required_cols = [
        "Location", "Customer Code", "Customer Name", "OrderNo",
        "InvoiceNo", "SO Line No", "NO of Post", "Deployment Hrs",
        "WF_TaskID", "Performed Hrs", "Billed Hrs",
        "Billed Vs Performed", "Contracted Vs Performed",
        "Billing Pattern", "ERP Cont Hrs", "Saturn Cont Hrs",
        "Scheduled Hrs"
    ]

    df_data = df_master[required_cols].copy()

    # Add new columns
    new_cols = ["HUB","Zone","Key","Owner","Key2",
                "Period From","Period To","Attendance"]

    for col in new_cols:
        df_data[col] = ""

    # ==============================
    # MACRO 2 → DATE DATA
    # ==============================

    df_date = df_master_date[["OrderNo","PeriodFrom","PeriodTo"]].copy()
    df_date.columns = ["OrderNo","Period From","Period To"]

    # ==============================
    # MACRO 3 → LOOKUPS (VLOOKUP → MERGE)
    # ==============================

    # HUB & Zone
    df_data = df_data.merge(df_hub_zone, on="Location", how="left")

    # Key
    df_data["Key"] = df_data["Location"].astype(str) + df_data["Customer Code"].astype(str)

    # Owner
    df_data = df_data.merge(df_owner, left_on="Key", right_on="Key", how="left")

    # Key2
    df_data["Key2"] = df_data["OrderNo"].astype(str) + df_data["SO Line No"].astype(str)

    # Period From/To
    df_data = df_data.merge(df_date, on="OrderNo", how="left")

    # ==============================
    # MACRO 4 → ATTENDANCE LOOKUP
    # ==============================

    import re
    from datetime import datetime
    
    # Build billing dictionary dynamically from column headers
    billing_dict = {}
    
    for i in range(len(df_billing)):
    
        key2 = str(df_billing.iloc[i, 0]).strip()
    
        for j in range(1, df_billing.shape[1]):
    
            header = str(df_billing.columns[j])
    
            # Extract pattern like "Sum of 15th to 14th"
            match = re.search(r"(\d+).*to.*(\d+)", header)
    
            if match:
    
                start_day = int(match.group(1))
                end_day = int(match.group(2))
    
                # Use Period From month/year from df_data dynamically
                for _, row in df_data.iterrows():
    
                    if pd.notna(row["Period From"]):
    
                        base_date = pd.to_datetime(row["Period From"])
    
                        month = base_date.month
                        year = base_date.year
    
                        # If cycle crosses month
                        if start_day > end_day:
                            period_from = datetime(year, month, start_day)
                            next_month = base_date + pd.DateOffset(months=1)
                            period_to = datetime(next_month.year, next_month.month, end_day)
                        else:
                            period_from = datetime(year, month, start_day)
                            period_to = datetime(year, month, end_day)
    
                        k = f"{key2}|{period_from.strftime('%d-%b-%y')}|{period_to.strftime('%d-%b-%y')}"
                        billing_dict[k] = df_billing.iloc[i, j]
    
    # Map attendance
    df_data["lookup_key"] = (
        df_data["Key2"].astype(str) + "|" +
        pd.to_datetime(df_data["Period From"]).dt.strftime("%d-%b-%y") + "|" +
        pd.to_datetime(df_data["Period To"]).dt.strftime("%d-%b-%y")
    )
    
    df_data["Attendance"] = df_data["lookup_key"].map(billing_dict)

    # ==============================
    # MACRO 5 → PIVOT TABLE
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

    pivot["Var Billed vs Performed"] = pivot["Billed Hrs"] - pivot["Performed Hrs"]

    # Special customer rule
    pivot.loc[pivot["Customer Code"] == special_customer,
              "Performed Hrs"] = pivot["Performed Hrs"]

    # ==============================
    # MACRO 6 → HUB WISE SHEETS
    # ==============================

    hubs = [h.strip() for h in hubs_input.split(",")]

    hub_data = {}
    for hub in hubs:
        hub_df = pivot[pivot["HUB"] == hub].copy()
        hub_df.loc["Grand Total"] = hub_df.sum(numeric_only=True)
        hub_data[hub] = hub_df

    # ==============================
    # DISPLAY OUTPUT
    # ==============================

    st.success("✅ Automation Completed")

    st.subheader("📊 India Conso (Pivot)")
    st.dataframe(pivot, use_container_width=True)

    for hub in hubs:
        st.subheader(f"📌 HUB: {hub}")
        st.dataframe(hub_data[hub], use_container_width=True)

    # ==============================
    # DOWNLOAD OPTION
    # ==============================

    with pd.ExcelWriter("Hours_Recon_Output.xlsx") as writer:
        pivot.to_excel(writer, sheet_name="India Conso", index=False)
        for hub in hubs:
            hub_data[hub].to_excel(writer, sheet_name=hub, index=False)

    with open("Hours_Recon_Output.xlsx", "rb") as f:
        st.download_button(
            "⬇ Download Excel Output",
            f,
            file_name="Hours_Recon_Output.xlsx"
        )
