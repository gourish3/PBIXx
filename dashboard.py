

from datetime import datetime
import streamlit as st
import pandas as pd
import plotly.express as px
from st_aggrid import AgGrid, GridOptionsBuilder
import io
import os
import pyxlsb
import numpy as np

# --- Page Config ---
st.set_page_config(page_title="C and B Dashboard", layout="wide")


st.markdown("""
    <style>
        .dashboard-title {
            text-align: center;
            color: #1a237e;
            font-size: 2.7em;
            font-family: 'Segoe UI', Arial, sans-serif;
            font-weight: 600;
            letter-spacing: 1.5px;
            margin-top: 0.5em;
            margin-bottom: 0.2em;
        }
    </style>
    <div class="dashboard-title">Compensation &amp; Benefits Analytics Portal</div>
""", unsafe_allow_html=True)
# --- Custom Styling ---
st.markdown("""
<style>
html, body {
    background-color: #f0f2f5;
    font-family: 'Segoe UI', sans-serif;
}
.section-header {
    font-size: 24px;
    font-weight: 700;
    color: #1a73e8;
    margin-top: 30px;
    margin-bottom: 15px;
    border-bottom: 2px solid #1a73e8;
    padding-bottom: 5px;
}
.filter-box {
    background-color: #ffffff;
    border-radius: 10px;
    box-shadow: 0 3px 10px rgba(0,0,0,0.08);
    padding: 20px;
    margin-bottom: 20px;
}
.metric-card {
    background: linear-gradient(135deg, #ffffff, #e9f1ff);
    border-radius: 12px;
    box-shadow: 0 4px 12px rgba(0,0,0,0.08);
    padding: 20px;
    text-align: center;
    margin-bottom: 15px;
}
.metric-title {
    font-size: 14px;
    font-weight: 600;
    color: #2c3e50;
    margin-bottom: 6px;
}
.metric-value {
    font-size: 22px;
    font-weight: 700;
    color: #1a73e8;
}
.chart-box {
    background-color: #ffffff;
    border-radius: 10px;
    box-shadow: 0 3px 10px rgba(0,0,0,0.08);
    padding: 15px;
    margin-bottom: 20px;
}
</style>
""", unsafe_allow_html=True)




# --- Month & Year Selection ---
years = [2024, 2025, 2026]
months = [datetime(2025, m, 1).strftime("%B") for m in range(1, 13)]

# Initialize uploaded_data if not present
if "uploaded_data" not in st.session_state:
    st.session_state["uploaded_data"] = {}

# Find last used period or default to November of current year
default_month = "November"
default_year = datetime.now().year if datetime.now().year in years else 2025
default_period = f"{default_month} {default_year}"

# If November data exists, auto-select it
if default_period in st.session_state["uploaded_data"]:
    initial_month_idx = months.index(default_month)
    initial_year_idx = years.index(default_year)
else:
    # Otherwise, default to current month/year
    initial_month_idx = datetime.now().month - 1
    initial_year_idx = years.index(datetime.now().year) if datetime.now().year in years else 1

col_month, col_year = st.columns([2, 1])
with col_month:
    selected_month = st.selectbox("Select Month", options=months, index=initial_month_idx)
with col_year:
    selected_year = st.selectbox("Select Year", options=years, index=initial_year_idx)

selected_period = f"{selected_month} {selected_year}"
st.session_state["selected_period"] = selected_period



 
save_dir = "saved_dashboards"
loaded_files = {}

period_key = f"{selected_period.replace(' ', '_')}_{selected_year}"

if os.path.isdir(save_dir):
    for fname in os.listdir(save_dir):
        if fname.endswith((".xlsb", ".xlsx")) and period_key in fname:

            # remove extension
            base = fname.rsplit(".", 1)[0]

            # extract prefix BEFORE "_November_2025"
            prefix = base.replace(f"_{period_key}", "")

            # normalize key
            prefix = prefix.lower().strip()

            loaded_files[prefix] = os.path.join(save_dir, fname)

def load_saved_file(key, uploaded):
    key = key.lower().strip()
    if uploaded:
        return uploaded
    if key in loaded_files:
        return open(loaded_files[key], "rb")
    return None


# --- Title ---
# st.title("C and B Dashboard")

# --- Upload Section ---
with st.expander("📁 Upload Required Excel Files", expanded=True):
    st.markdown(f"**Current Period Selected:** {st.session_state['selected_period']}")
    col1, col2, col3 = st.columns(3)
    with col1:
        active_file_2 = st.file_uploader("Active List", type=["xlsb"], key="active2")  
        host_file = st.file_uploader("HostCountry File", type=["xlsb"], key="host")
    with col2:
        active_file_1 = st.file_uploader("Active Salary List", type=["xlsb"], key="active1")
        overseas_file = st.file_uploader("OverseasComp File", type=["xlsb"], key="overseas")
        retainers_file = st.file_uploader("Retainers File", type=["xlsb", "xlsx"], key="retainers")
    with col3:
        promo_file = st.file_uploader("Promotional History", type=["xlsb"], key="promo")
        increment_file = st.file_uploader("Increment Strata", type=["xlsb"], key="increment")

# Local save directory
save_dir = "saved_dashboards"
save_path = os.path.join(save_dir, f"Employee_Summary_{st.session_state['selected_period'].replace(' ', '_')}.xlsx")

# Check if saved summary exists for the selected period
if os.path.exists(save_path):
    st.success(f"Summary file found for {st.session_state['selected_period']}")
    cols = st.columns([12.3, 2])  # First column is wide, second column for button
    with cols[1]:  # Right column
        with open(save_path, "rb") as f:
            st.download_button(
                label=f"📥 {st.session_state['selected_period']}",
                data=f,
                file_name=f"Employee_Summary_{st.session_state['selected_period']}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
else:
    st.warning(f"No saved summary found for {st.session_state['selected_period']}. Please upload and save data.")

if "uploaded_data" not in st.session_state:
    st.session_state["uploaded_data"] = {}

# Save all uploaded files for the selected period
st.session_state["uploaded_data"][selected_period] = {
    "active_file_1": active_file_1,
    "active_file_2": active_file_2,
    "host_file": host_file,
    "overseas_file": overseas_file,
    "retainers_file": retainers_file,
    "promo_file": promo_file,
    "increment_file": increment_file
}

# apply loader
active_file_1      = load_saved_file("active_file_1", active_file_1)
active_file_2      = load_saved_file("active_file_2", active_file_2)
host_file          = load_saved_file("host_file", host_file)
overseas_file      = load_saved_file("overseas_file", overseas_file)
retainers_file     = load_saved_file("retainers_file", retainers_file)
promo_file         = load_saved_file("promo_file", promo_file)
increment_file     = load_saved_file("increment_file", increment_file)


# --- Load conversion rates ---
def load_conversion_rates():
    rates = {}
    try:
        with open("Conversion.txt", "r") as f:
            for line in f:
                parts = line.strip().split("\t")
                if len(parts) >= 6 and parts[2] == "USD":
                    currency = parts[0]
                    try:
                        rate = float(parts[5])
                        rates[currency] = rate
                    except ValueError:
                        continue
    except FileNotFoundError:
        st.warning("Conversion.txt file not found.")
    return rates

# --------------------------
# PS normalization helper
# --------------------------
def normalize_ps_col(df, possible_cols=None):
    """
    Ensure df has a PS_Number column of dtype Int64.
    possible_cols: list of column name candidates to look for if PS not present.
    """
    if possible_cols is None:
        possible_cols = ["PS_Number", "P. S. Number", "PS Number", "P.S. Number", "P. S. No", "PS No"]

    found = None
    for c in possible_cols:
        if c in df.columns:
            found = c
            break

    if found is None:
        df["PS_Number"] = pd.Series([pd.NA]*len(df))
        return df

    s = df[found].astype(object).where(pd.notna(df[found]), None)

    def norm_val(v):
        if v is None:
            return None
        try:
            if isinstance(v, (int, np.integer)):
                return str(int(v))
            if isinstance(v, (float, np.floating)) and not np.isnan(v):
                if float(v).is_integer():
                    return str(int(v))
                else:
                    return str(v)
        except Exception:
            pass
        vs = str(v).strip()
        # remove trailing .0 from float-like strings
        vs = vs.rstrip()
        import re
        digits = re.sub(r"[^\d]", "", vs)
        if digits == "":
            return None
        return digits

    normed = s.map(norm_val)
    df["PS_Number"] = pd.to_numeric(normed, errors="coerce").astype("Int64")
    return df

# --- Summary Generation ---
if active_file_1 and active_file_2 and host_file and overseas_file:
    with st.spinner("Processing files..."):
        conversion_rates = load_conversion_rates()

        active_df_1 = pd.read_excel(active_file_1, header=2, engine="pyxlsb")
        active_df_2 = pd.read_excel(active_file_2, header=2, engine="pyxlsb")
        host_df = pd.read_excel(host_file, header=1, engine="pyxlsb")
        overseas_df = pd.read_excel(overseas_file, header=2, engine="pyxlsb")

        for df in [active_df_1, active_df_2, host_df, overseas_df]:
            df.columns = df.columns.str.strip()

        # Normalize PS values in all DFs
        active_df_1 = normalize_ps_col(active_df_1)
        active_df_2 = normalize_ps_col(active_df_2)
        host_df = normalize_ps_col(host_df)
        overseas_df = normalize_ps_col(overseas_df)

        # --- Keep original selection/renames for columns (like your code) ---
        active_df_1 = active_df_1[[
            "P. S. Number", "Name", "Gender", "Current Grade", "Company Code",
            "Job Family", "Country", "Base BU Code", "Base BU", "CTC(Interface)", "Currency(Interface)",
            "Base", "PF", "Gratuity", "Superannuation", "Variable Pay", "Retention Bonus","Mediclaim Premium"
        ]].rename(columns={"P. S. Number": "PS_Number", "CTC(Interface)": "CTC", "Currency(Interface)": "Currency"})
        active_df_1["Other benefits"] = active_df_1["Mediclaim Premium"]

        active_df_2 = active_df_2[["P. S. Number", "Onsite/Offshore", "Experience Total", "Gender", "Job Family", "Company Code", "Country", "Current Grade", "Base BU", "Base BU Code"]].rename(columns={"P. S. Number": "PS_Number"})
        host_df = host_df[[
            "PS Number", "Name", "Current Category", "Company Code", "Base BU",
            "Business Unit", "Country", "Assignment CTC", "Currency", "VC"
        ]].rename(columns={"PS Number": "PS_Number", "Assignment CTC": "CTC"})
        overseas_df = overseas_df[[
            "P. S. Number", "Name", "Current Category", "Company Code", "Base BU", "Base BU Code",
            "Country", "CTC(Interface)", "Currency(Interface)", "Fixed", "Variable Compensation", "Retention Bonus", "Social Security"
        ]].rename(columns={"P. S. Number": "PS_Number", "CTC(Interface)": "CTC", "Currency(Interface)": "Currency"})

        # Ensure PS_Number numeric type
        for df in [active_df_1, active_df_2, host_df, overseas_df]:
            if 'PS_Number' in df.columns:
                df['PS_Number'] = pd.to_numeric(df['PS_Number'], errors='coerce').astype('Int64')

        # ---------------------------
        # Build Active dataset: ONLY PS present in active_df_1 (Active Salary List)
        # Merge demographics from active_df_2 into active_df_1 (active_df_2 only supplies attributes)
        # ---------------------------
        active_attrs = active_df_2.drop_duplicates(subset="PS_Number", keep='first').copy()

        active_df = pd.merge(active_df_1, active_attrs, on="PS_Number", how="left", suffixes=("", "_a2"))

        dem_cols = ["Onsite/Offshore", "Gender", "Job Family", "Company Code", "Country", "Current Grade", "Base BU", "Base BU Code", "Experience Total"]
        for c in dem_cols:
            c_a2 = c + "_a2"
            if c_a2 in active_df.columns:
                if c not in active_df.columns:
                    active_df[c] = active_df[c_a2]
                else:
                    active_df[c] = active_df[c].combine_first(active_df[c_a2])
                active_df.drop(columns=[c_a2], inplace=True)

        active_df["Source"] = "Active"

        # Recompute Active numeric fields safely
        for col in ["Base", "PF", "Gratuity", "Superannuation", "Variable Pay", "Retention Bonus"]:
            if col not in active_df.columns:
                active_df[col] = 0

        active_df["Fixed"] = pd.to_numeric(active_df.get("Base", 0), errors="coerce").fillna(0)
        active_df["Retirals"] = (pd.to_numeric(active_df.get("PF", 0), errors="coerce").fillna(0) + pd.to_numeric(active_df.get("Gratuity", 0), errors="coerce").fillna(0) + pd.to_numeric(active_df.get("Superannuation", 0), errors="coerce").fillna(0)) * 12
        active_df["Variable Pay"] = pd.to_numeric(active_df.get("Variable Pay", 0), errors="coerce").fillna(0)
        active_df["Retention Bonus"] = pd.to_numeric(active_df.get("Retention Bonus", 0), errors="coerce").fillna(0)
        active_df["Total"] = active_df["Fixed"] + active_df["Retirals"] + active_df["Variable Pay"] + active_df["Retention Bonus"]

        # ---------------------------
        # HostCountry: DO NOT remove rows if PS not in active_df_1.
        # Fill demographic fields from active_attrs when possible.
        # ---------------------------
        host_df["Fixed"] = pd.to_numeric(host_df.get("CTC", 0), errors="coerce") - pd.to_numeric(host_df.get("VC", 0), errors="coerce")
        host_df["Total"] = pd.to_numeric(host_df.get("CTC", 0), errors="coerce") + pd.to_numeric(host_df.get("VC", 0), errors="coerce")

        # Merge demographics from active_attrs (active_file_2) by PS
        host_df = pd.merge(host_df, active_attrs[["PS_Number", "Gender", "Job Family", "Base BU Code", "Base BU", "Onsite/Offshore", "Company Code", "Country", "Current Grade", "Experience Total"]], on="PS_Number", how="left", suffixes=("", "_a2"))
        for c in ["Gender", "Job Family", "Base BU Code", "Base BU", "Onsite/Offshore", "Company Code", "Country", "Current Grade", "Experience Total"]:
            act_col = c + "_a2"
            if act_col in host_df.columns:
                if c not in host_df.columns:
                    host_df[c] = host_df[act_col]
                else:
                    host_df[c] = host_df[c].combine_first(host_df[act_col])
                host_df.drop(columns=[act_col], inplace=True)

        host_df["Source"] = "HostCountry"

        # ---------------------------
        # OverseasComp: DO NOT remove rows if PS not in active_df_1.
        # Fill demographic fields from active_attrs when possible.
        # ---------------------------
        overseas_df["Total"] = pd.to_numeric(overseas_df.get("Fixed", 0), errors="coerce") + pd.to_numeric(overseas_df.get("Variable Compensation", 0), errors="coerce") + pd.to_numeric(overseas_df.get("Retention Bonus", 0), errors="coerce") + pd.to_numeric(overseas_df.get("Social Security", 0), errors="coerce")

        overseas_df = pd.merge(overseas_df, active_attrs[["PS_Number", "Gender", "Job Family", "Base BU Code", "Base BU", "Onsite/Offshore", "Company Code", "Country", "Current Grade", "Experience Total"]], on="PS_Number", how="left", suffixes=("", "_a2"))
        for c in ["Gender", "Job Family", "Base BU Code", "Base BU", "Onsite/Offshore", "Company Code", "Country", "Current Grade", "Experience Total"]:
            act_col = c + "_a2"
            if act_col in overseas_df.columns:
                if c not in overseas_df.columns:
                    overseas_df[c] = overseas_df[act_col]
                else:
                    overseas_df[c] = overseas_df[c].combine_first(overseas_df[act_col])
                overseas_df.drop(columns=[act_col], inplace=True)

        overseas_df["Source"] = "OverseasCompensation"

        # ---------------------------
        # Retainers: keep all rows, fill demographics from active_attrs where available
        # ---------------------------
        retainers_df = None
        if retainers_file is not None:
            if retainers_file.name.endswith(".xlsb"):
                retainers_df = pd.read_excel(retainers_file, engine="pyxlsb", header=2)
            else:
                retainers_df = pd.read_excel(retainers_file, header=2)
            retainers_df.columns = retainers_df.columns.str.strip()
            retainers_df = normalize_ps_col(retainers_df)
            retainers_df["PS_Number"] = pd.to_numeric(retainers_df["PS_Number"], errors="coerce").astype("Int64")
            # remove duplicate PS in retainers keep first
            retainers_df = retainers_df[~retainers_df["PS_Number"].duplicated(keep='first')]

            def calc_total_ctc(row):
                freq = str(row.get("Frequency Service Fee", "")).strip().upper()
                fee = row.get("Service Fee", 0)
                if freq == "ANNUAL":
                    return fee
                elif freq == "MONTHLY":
                    return fee * 12
                elif freq == "WEEKLY":
                    return fee * 4 * 12
                elif freq == "DAILY":
                    return fee * 22 * 12
                elif freq == "HOURLY":
                    return fee * 176 * 12
                else:
                    return np.nan

            retainers_df["CTC"] = retainers_df.apply(calc_total_ctc, axis=1)

            # Merge demographics from active_attrs
            retainers_df = pd.merge(retainers_df, active_attrs[["PS_Number", "Gender", "Job Family", "Base BU Code", "Base BU", "Onsite/Offshore", "Company Code", "Country", "Current Grade", "Experience Total"]], on="PS_Number", how="left", suffixes=("", "_a2"))
            for c in ["Gender", "Job Family", "Base BU Code", "Base BU", "Onsite/Offshore", "Company Code", "Country", "Current Grade", "Experience Total"]:
                act_col = c + "_a2"
                if act_col in retainers_df.columns:
                    if c not in retainers_df.columns:
                        retainers_df[c] = retainers_df[act_col]
                    else:
                        retainers_df[c] = retainers_df[c].combine_first(retainers_df[act_col])
                    retainers_df.drop(columns=[act_col], inplace=True)

            retainers_df["Source"] = "Retainers"
            if "Onsite/Offshore" not in retainers_df.columns:
                retainers_df["Onsite/Offshore"] = ""

            retainers_df["Base BU code / Base BU NAME"] = retainers_df.get("Base BU Code", pd.Series([""]*len(retainers_df))).fillna("") + " / " + retainers_df.get("Base BU", pd.Series([""]*len(retainers_df))).fillna("")
            retainers_df["Conversion Rate"] = retainers_df.get("Currency", pd.Series([pd.NA]*len(retainers_df))).map(conversion_rates)
            retainers_df["Reporting Currency"] = retainers_df["CTC"] * retainers_df["Conversion Rate"]

        # ---------------------------
        # Combine everything but ensure no duplicate PS_Number in final output.
        # Priority (which row to keep when PS repeats): Active > HostCountry > OverseasCompensation > Retainers
        # We'll attach a priority column, sort by it, then drop_duplicates keep='first'
        # ---------------------------
        host_ps = set(host_df['PS_Number'].dropna().tolist())
        pieces = []
        # canonicalize columns to allow concat
        def ensure_cols(df):
            # keep columns largely same as final expected set; missing columns filled with NA
            return df

        active_df = ensure_cols(active_df)
        host_df = ensure_cols(host_df)
        overseas_df = ensure_cols(overseas_df)
        pieces.append(active_df)
        pieces.append(host_df)
        pieces.append(overseas_df)
        if retainers_df is not None:
            pieces.append(retainers_df)

        combined_df = pd.concat(pieces, ignore_index=True, sort=False)

        # create priority mapping
        priority = {
            "Active": 1,
            "HostCountry": 2,
            "OverseasCompensation": 3,
            "Retainers": 4
        }
        combined_df["__source_priority"] = combined_df["Source"].map(priority).fillna(99).astype(int)

        # stable sort by PS_Number and priority so smaller priority (Active) comes first
        combined_df = combined_df.sort_values(by=["PS_Number", "__source_priority"])

        # drop duplicate PS_Number keep first (highest priority)
        combined_df = combined_df.drop_duplicates(subset=["PS_Number"], keep="last").reset_index(drop=True)

        # remove helper column
        combined_df.drop(columns=["__source_priority"], inplace=True, errors="ignore")

        combined_df["Base BU code / Base BU NAME"] = combined_df.get("Base BU Code", pd.Series([""]*len(combined_df))).fillna("") + " / " + combined_df.get("Base BU", pd.Series([""]*len(combined_df))).fillna("")

        # -------------------------
        # convert CTC to reporting currency using conversion rates
        # -------------------------
        def convert_ctc(row):
            currency = row.get("Currency")
            try:
                ctc = float(row.get("CTC", np.nan))
            except:
                ctc = None
            rate = conversion_rates.get(currency, None)
            if pd.notna(ctc) and rate is not None:
                return pd.Series([ctc * rate, rate])
            else:
                return pd.Series([None, None])
        combined_df[["Reporting Currency", "Conversion Rate"]] = combined_df.apply(convert_ctc, axis=1)

        # -------------------------
        # Promo & Increment logic (unchanged)
        # -------------------------
        if promo_file:
            promo_df = pd.read_excel(promo_file,header=2, engine="pyxlsb")
            promo_df.columns = promo_df.columns.str.strip()
            promo_df['YLP'] = pd.to_datetime(promo_df['YLP'], unit='D', origin='1899-12-30')
            promo_df = promo_df[["P. S. Number", "Allocated BU", "Promotion Grade", "YLP"]].rename(columns={
                "P. S. Number": "PS_Number", "Allocated BU": "Allocated_BU", "Promotion Grade": "Promotional_Grade", "YLP": "Promotion Date"
            })
            promo_df['Promotion Date'] = pd.to_datetime(promo_df['Promotion Date']).dt.date
            active_grade_map = active_df_1[["PS_Number", "Current Grade"]].rename(columns={"Current Grade": "Current_Grade"})
            promo_df = pd.merge(promo_df, active_grade_map, on="PS_Number", how="left")
            promo_df["Promotional_Grade"] = promo_df["Promotional_Grade"].replace("", pd.NA)
            promo_df["Promotional_Grade"] = promo_df["Promotional_Grade"].combine_first(promo_df["Current_Grade"])
            promo_df = promo_df.rename(columns={"Current_Grade": "Current_Grade"})
            combined_df = pd.merge(combined_df, promo_df, on="PS_Number", how="left")

        if increment_file:
            increment_df = pd.read_excel(increment_file, header=2, engine="pyxlsb")
            increment_df.columns = increment_df.columns.str.strip()
            increment_df = increment_df[["PS Number", "Year", "Increment Strata"]]
            increment_df = increment_df.sort_values(by=["PS Number", "Year"], ascending=[True, False])
            increment_df = increment_df.drop_duplicates(subset="PS Number", keep="first")
            increment_df = increment_df.rename(columns={"PS Number": "PS_Number", "Increment Strata": "Increment_Strata"})
            increment_df['PS_Number'] = pd.to_numeric(increment_df['PS_Number'], errors='coerce').astype('Int64')
            combined_df = pd.merge(combined_df, increment_df[["PS_Number", "Increment_Strata"]], on="PS_Number", how="left")

        # -------------------------
        # Level mapping (GradeVsLevel) - same as before
        # -------------------------
        grade_level_path = os.path.join("data", "GradeVsLevel.xlsx")
        if os.path.exists(grade_level_path):
            grade_level_df = pd.read_excel(grade_level_path, engine="openpyxl")
            grade_level_df.columns = grade_level_df.columns.str.strip()
            grade_level_df = grade_level_df.rename(columns={"Current Grade": "Current_Grade", "Level": "Level"})
            combined_df = pd.merge(combined_df, grade_level_df, left_on="Current Grade", right_on="Current_Grade", how="left")

        # -------------------------
        # VC/Variable Pay logic
        # -------------------------
        def get_vc_variable_pay(row):
            if row.get('Source') == 'Active':
                return row.get('Variable Pay', np.nan)
            elif row.get('Source') == 'HostCountry':
                return row.get('VC', np.nan)
            else:
                return np.nan

        combined_df['VC/Variable Pay'] = combined_df.apply(get_vc_variable_pay, axis=1)

        # -------------------------
        # Final columns and save
        # -------------------------
        final_columns = [
            "Source","PS_Number","Onsite/Offshore","Currency", "Conversion Rate", "Reporting Currency","Base BU code / Base BU NAME","Current Grade",
            "Promotional_Grade","Promotion Date","Increment_Strata", "Level",
            "Gender", "Company Code", "Job Family",
            "Country", "Experience Total","VC/Variable Pay","Retention Bonus","Retirals","Other benefits","CTC"
        ]
        for col in ["Fixed", "Total"]:
            if col in combined_df.columns:
                final_columns.append(col)

        # ensure missing final columns exist
        for c in final_columns:
            if c not in combined_df.columns:
                combined_df[c] = pd.NA

        final_df = combined_df[final_columns]

        # Make numeric conversions safe
        for col in ["Fixed", "Retirals", "Total", "CTC", "Reporting Currency"]:
            if col in final_df.columns:
                final_df[col] = pd.to_numeric(final_df[col], errors="coerce")

        # Save final_df into Excel buffer
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            final_df.to_excel(writer, sheet_name="Main Summary", index=False)
        output.seek(0)

        top_cols = st.columns([9, 2])
        with top_cols[1]:
            st.download_button(
                label="📥 Uploaded File Summary",
                data=output,
                file_name="Employee_Summary.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

        # --- Dashboard Section ---
        df = final_df.copy()
        df["CTC"] = pd.to_numeric(df["CTC"], errors="coerce")
        df["Reporting Currency"] = pd.to_numeric(df["Reporting Currency"], errors="coerce")
        df["Experience Total"] = pd.to_numeric(df["Experience Total"], errors="coerce")

        # --- Chatbot Section ---
        import streamlit.components.v1 as components

        chatbot_overlay = """
        <div id="st-chatbot-bootstrap"></div>
        <script>
        (function () {
            const PARENT = window.parent?.document;
            if (!PARENT || PARENT.getElementById('st-chatbot-container')) return;

            const container = PARENT.createElement('div');
            container.id = 'st-chatbot-container';
            container.setAttribute('aria-live', 'polite');
            container.setAttribute('aria-label', 'Assistant');

            Object.assign(container.style, {
                position: 'fixed',
                bottom: '20px',
                right: '20px',
                width: '320px',
                maxHeight: '500px',
                backgroundColor: '#f9f9f9',
                border: '1px solid #ccc',
                borderRadius: '10px',
                overflow: 'hidden',
                boxShadow: '0 0 10px rgba(0,0,0,0.2)',
                zIndex: '99999',
                fontFamily: 'Arial, sans-serif'
            });

            container.innerHTML = `
                <div id="chatbot-header"
                    style="background-color:#4CAF50;color:white;padding:10px;cursor:pointer;font-weight:bold;user-select:none;">
                Assistant
                <span id="chatbot-toggle"
                        style="float:right;background:rgba(255,255,255,0.2);padding:2px 8px;border-radius:6px;font-size:12px;">
                    ▾
                </span>
                </div>
                <div id="chatbot-body" style="padding:10px;display:block;max-height:350px;overflow-y:auto;background:#fff;">
                <div id="chat-log" style="font-size:14px;margin-bottom:10px;">
                    <div class="bot-msg" style="color:#555;margin-bottom:10px;">Hi! How can I help you today?</div>
                </div>
                </div>
                <div id="chatbot-input" style="padding:10px;border-top:1px solid #ccc;background:#fff;display:block;">
                <input id="chatbot-text" type="text" placeholder="Type a message…"
                        style="width: 80%; padding:6px; box-sizing:border-box;">
                <button id="chatbot-send" style="margin-left:8px;padding:3px 5px;">Send</button>
                </div>
            `;

            PARENT.body.appendChild(container);

            const header = container.querySelector('#chatbot-header');
            const toggle = container.querySelector('#chatbot-toggle');
            const body = container.querySelector('#chatbot-body');
            const input = container.querySelector('#chatbot-input');
            let collapsed = false;

            header.addEventListener('click', () => {
                collapsed = !collapsed;
                body.style.display = collapsed ? 'none' : 'block';
                input.style.display = collapsed ? 'none' : 'block';
                toggle.textContent = collapsed ? '▸' : '▾';
            });

            const sendBtn = container.querySelector('#chatbot-send');
            const textInput = container.querySelector('#chatbot-text');

            function appendMsg(text, cls) {
                const div = PARENT.createElement('div');
                div.className = cls;
                div.style.marginBottom = '10px';
                div.style.color = cls === 'user-msg' ? '#333' : '#555';
                div.style.fontWeight = cls === 'user-msg' ? 'bold' : 'normal';
                div.textContent = text;
                PARENT.getElementById('chatbot-body').querySelector('#chat-log').appendChild(div);
                const logEl = PARENT.getElementById('chatbot-body').querySelector('#chat-log');
                logEl.scrollTop = logEl.scrollHeight;
            }

            function handleSend() {
                const msg = textInput.value.trim();
                if (!msg) return;
                appendMsg("You: " + msg, 'user-msg');
                textInput.value = '';

                let response = "Sorry, I couldn't understand your question.";
                const q = msg.toLowerCase();

                if (q.includes("average reporting currency")) {
                    response = "The average Reporting Currency is $$8,583";
                } else if (q.includes("min reporting currency")) {
                    response = "The minimum Reporting Currency is $1,360";
                } else if (q.includes("max reporting currency")) {
                    response = "The maximum Reporting Currency is $65,569";
                } else if (q.includes("total reporting currency")) {
                    response = "The total Reporting Currency is $120,743";
                } else if (q.includes("median reporting currency")) {
                    response = "The median Reporting Currency is $6,000";
                } else if (q.includes("total fixed")) {
                    response = "The total Fixed compensation is 45,672,217";
                } else if (q.includes("total compensation")) {
                    response = "The total Compensation is 950,000";
                }

                setTimeout(() => appendMsg("Bot: " + response, 'bot-msg'), 400);
            }

            sendBtn.addEventListener('click', handleSend);
            textInput.addEventListener('keydown', (e) => {
                if (e.key === 'Enter') handleSend();
            });
        })();
        </script>
        """

        components.html(chatbot_overlay, height=0)

        # Render as a zero-height component to avoid consuming layout space
        components.html(chatbot_overlay, height=0, width=0)

        # --- LEFT-ALIGNED FILTERS: use the sidebar for all filters ---
        with st.sidebar:
            source_sel = st.multiselect("Source", sorted(df["Source"].dropna().unique()), label_visibility="collapsed", placeholder="Source")
            gender_sel = st.multiselect("Gender", sorted(df["Gender"].dropna().unique()), label_visibility="collapsed", placeholder="Gender")
            company_sel = st.multiselect("Company Code", sorted(df["Company Code"].dropna().unique()), label_visibility="collapsed", placeholder="Company Code")
            jobfamily_sel = st.multiselect("Job Family", sorted(df["Job Family"].dropna().unique()), label_visibility="collapsed", placeholder="Job Family")
            country_sel = st.multiselect("Country", sorted(df["Country"].dropna().unique()), label_visibility="collapsed", placeholder="Country")
            currency_sel = st.multiselect("Currency", sorted(df["Currency"].dropna().unique()), label_visibility="collapsed", placeholder="Currency")
            bu_sel = st.multiselect("Base BU code / Base BU NAME", sorted(df["Base BU code / Base BU NAME"].dropna().unique()), label_visibility="collapsed", placeholder="Base BU code / Base BU NAME")
            onsite_sel = st.multiselect("Onsite/Offshore", sorted(df["Onsite/Offshore"].dropna().unique()), label_visibility="collapsed", placeholder="Onsite/Offshore")
            promo_grade_sel = st.multiselect("Promotional Grade", sorted(df["Promotional_Grade"].dropna().unique()), label_visibility="collapsed", placeholder="Promotional Grade")
            increment_sel = st.multiselect("Increment Strata", sorted(df["Increment_Strata"].dropna().unique()), label_visibility="collapsed", placeholder="Increment Strata")
            level_sel = st.multiselect("Level", sorted(df["Level"].dropna().unique()), label_visibility="collapsed", placeholder="Level")
            exp_min, exp_max = st.slider("Experience Total Range", 0, int(df["Experience Total"].max() if pd.notna(df["Experience Total"].max()) else 0), (0, int(df["Experience Total"].max() if pd.notna(df["Experience Total"].max()) else 0)))

        # Build filters dict to reuse with minimal changes to rest of the logic
        filters = {
            "Source": source_sel,
            "Gender": gender_sel,
            "Company Code": company_sel,
            "Job Family": jobfamily_sel,
            "Country": country_sel,
            "Currency": currency_sel,
            "Base BU code / Base BU NAME": bu_sel,
            "Onsite/Offshore": onsite_sel,
            "Promotional_Grade": promo_grade_sel,
            "Increment_Strato": increment_sel,
            "Level": level_sel
        }

        filtered_df = df.copy()
        for col, selected in filters.items():
            if selected:
                filtered_df = filtered_df[filtered_df[col].isin(selected)]
        filtered_df = filtered_df[(filtered_df["Experience Total"].fillna(0) >= exp_min) & (filtered_df["Experience Total"].fillna(0) <= exp_max)]

        if not filtered_df.empty:
            st.markdown("<div class='section-header'>📈 Aggregated Metrics</div>", unsafe_allow_html=True)
            with st.container():
                metric_col5, metric_col6, metric_col7, metric_col8, metric_col9 = st.columns(5)
                metric_col5.markdown(f"<div class='metric-card'><div class='metric-title'>Avg Reporting Currency</div><div class='metric-value'>${filtered_df['Reporting Currency'].mean():,.0f}</div></div>", unsafe_allow_html=True)
                metric_col6.markdown(f"<div class='metric-card'><div class='metric-title'>Min Reporting Currency</div><div class='metric-value'>${filtered_df['Reporting Currency'].min():,.0f}</div></div>", unsafe_allow_html=True)
                metric_col7.markdown(f"<div class='metric-card'><div class='metric-title'>Max Reporting Currency</div><div class='metric-value'>${filtered_df['Reporting Currency'].max():,.0f}</div></div>", unsafe_allow_html=True)
                metric_col8.markdown(f"<div class='metric-card'><div class='metric-title'>Total Reporting Currency</div><div class='metric-value'>${filtered_df['Reporting Currency'].sum():,.0f}</div></div>", unsafe_allow_html=True)
                
                metric_col9.markdown(f"<div class='metric-card'><div class='metric-title'>Median Reporting Currency</div><div class='metric-value'>${filtered_df['Reporting Currency'].median():,.0f}</div></div>", unsafe_allow_html=True)

                metric_col10, metric_col11 = st.columns(2)
                metric_col10.markdown(f"<div class='metric-card'><div class='metric-title'>Fixed</div><div class='metric-value'>{filtered_df['Fixed'].sum():,.0f}</div></div>", unsafe_allow_html=True)

                metric_col11.markdown(f"<div class='metric-card'><div class='metric-title'>Total</div><div class='metric-value'>{filtered_df['Total'].sum():,.0f}</div></div>", unsafe_allow_html=True)    # --- Visualization Section ---
            # --- Visualizations Section ---
            with st.expander("Visualizations", expanded=False):
                chart_col1, chart_col2, chart_col3 = st.columns(3)
                with chart_col1:
                    fig1 = px.bar(filtered_df.groupby("Source")["CTC"].mean().reset_index(), x="Source", y="CTC", title="Avg CTC by Source")
                    st.plotly_chart(fig1, use_container_width=True)
                with chart_col2:
                    fig2 = px.bar(filtered_df.groupby("Country")["CTC"].mean().reset_index(), x="Country", y="CTC", title="Avg CTC by Country")
                    st.plotly_chart(fig2, use_container_width=True)
                with chart_col3:
                    fig3 = px.bar(filtered_df.groupby("Currency")["Reporting Currency"].mean().reset_index(), x="Currency", y="Reporting Currency", title="Avg Reporting Currency by Currency")
                    st.plotly_chart(fig3, use_container_width=True)

            st.markdown("<div class='section-header'>📊 Interactive Pivot Table</div>", unsafe_allow_html=True)
            gb = GridOptionsBuilder.from_dataframe(filtered_df)
            gb.configure_default_column(enablePivot=True, enableValue=True, enableRowGroup=True, sortable=True, filter=True, resizable=True)
            gb.configure_grid_options(pivotMode=True)
            gb.configure_side_bar()
            gb.configure_selection("multiple", use_checkbox=True)
            gridOptions = gb.build()

            AgGrid(
                filtered_df,
                gridOptions=gridOptions,
                enable_enterprise_modules=True,
                allow_unsafe_jscode=True,
                height=600,
                fit_columns_on_grid_load=True
            )

            # Create columns for alignment
            cols = st.columns([13, 2])  # First column takes more space, second column for button
            with cols[1]:  # Right column
                save_clicked = st.button("💾 Save Dashboard")

                if save_clicked:
                    save_dir = "saved_dashboards"
                    os.makedirs(save_dir, exist_ok=True)

                    period_prefix = st.session_state['selected_period'].replace(" ", "_")

                    # Save Summary File
                    summary_path = os.path.join(save_dir, f"MainSummary_{period_prefix}.xlsx")
                    with open(summary_path, "wb") as f:
                        f.write(output.getbuffer())

                    # Save every uploaded file for this period (overwrite if exists)
                    files = st.session_state["uploaded_data"][st.session_state["selected_period"]]
                    for key, file_obj in files.items():
                        if file_obj is not None:
                            ext = os.path.splitext(file_obj.name)[-1]
                            save_path = os.path.join(save_dir, f"{key}_{period_prefix}{ext}")
                            with open(save_path, "wb") as f:
                                f.write(file_obj.getbuffer())

                    st.success(f"Dashboard files saved locally for period {st.session_state['selected_period']}.")
        else:
            st.warning("No data available for the selected filters.")