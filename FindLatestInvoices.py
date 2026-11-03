import io
import re
import pandas as pd
import streamlit as st

# --- VECTORIZED INVOICE STRING CLEANING ---
def clean_invoice_series(series): 
    """Standardizes invoice numbers by stripping whitespace, trailing '.0', 

    COR/REV suffixes, and leading zeros.
    """
    s = series.astype(str).str.strip() 
    s = s.str.replace(r'\.0$', '', regex=True) 
    base = s.str.replace(r'[\s\-]*(COR|REV)\d*$', '', regex=True, flags=re.IGNORECASE) 
    cleaned = base.str.lstrip('0') 
    return cleaned.where(cleaned != '', base).replace('nan', pd.NA)

def get_invoice_rank(inv): 
    """Ranks invoice versions to identify the latest transaction record.

    Returns a tuple of (num_val, type_val, dash_count).
    """
    if pd.isna(inv): 
        return (0, 0, 0) 
    inv_str = str(inv).upper().strip() 
    match = re.search(r'([\s\-]*)(COR|REV)(\d*)$', inv_str) 
    if not match: 
        return (0, 0, 0) 
    
    separator = match.group(1)
    type_str = match.group(2)
    num_str = match.group(3)
    
    type_val = 2 if type_str == 'COR' else 1
    num_val = int(num_str) if num_str else 1
    dash_count = separator.count('-') 
    
    return (num_val, type_val, dash_count)

def load_file_buffer(uploaded_file):
    """Loads Excel (.xlsx, .xls, .xlsb) or CSV/TXT files from Streamlit buffers."""
    filename = uploaded_file.name.lower()
    if filename.endswith(".xlsb"):
        return pd.read_excel(uploaded_file, sheet_name=None, engine="pyxlsb", dtype=str)
    elif filename.endswith((".xlsx", ".xls")):
        xls = pd.ExcelFile(uploaded_file)
        if len(xls.sheet_names) > 1:
            return pd.read_excel(xls, sheet_name=None, dtype=str)
        else:
            return pd.read_excel(xls, sheet_name=0, dtype=str)
    else:
        try:
            return pd.read_csv(uploaded_file, sep=None, engine="python", dtype=str)
        except Exception:
            uploaded_file.seek(0)
            return pd.read_excel(uploaded_file, sheet_name=0, dtype=str)

# --- STREAMLIT UI CONFIGURATION ---
def main():
    st.set_page_config(page_title="Find Latest Invoices Tool", page_icon="🔍", layout="wide")

    st.title("🔍 Find Latest Invoices from ATF")
    st.markdown(
        "Upload your **ATF File** (and optionally a **Requested Correction File**) "
        "to extract only the latest invoice records filtered by `Process Code == CCREC`."
    )

    st.sidebar.header("📁 File Uploads")
    atf_file = st.sidebar.file_uploader(
        "ATF File (Required)", type=["xlsx", "xls", "xlsb", "csv", "txt"]
    )
    correction_file = st.sidebar.file_uploader(
        "Requested Correction File (Optional - filter specific invoices)", 
        type=["xlsx", "xls", "xlsb", "csv", "txt"]
    )

    # --- SESSION STATE FOR PERSISTENT DOWNLOAD BUTTONS ---
    if "processed_results" not in st.session_state:
        st.session_state.processed_results = None

    if st.button("🚀 Find Latest Invoices", type="primary"):
        # Reset session state on re-click
        st.session_state.processed_results = None

        if not atf_file:
            st.error("⚠️ Please upload the ATF File to proceed.")
        else:
            with st.spinner("Processing ATF data to extract latest invoices..."):
                try:
                    # 1. Load ATF File
                    atf_data = load_file_buffer(atf_file)
                    if isinstance(atf_data, dict):
                        df_atf = pd.concat(atf_data.values(), ignore_index=True)
                    else:
                        df_atf = atf_data.copy()

                    df_atf.columns = df_atf.columns.str.strip()

                    # 2. Filter Process Code = 'CCREC'
                    if "Process Code" in df_atf.columns:
                        df_atf = df_atf[
                            df_atf["Process Code"].astype(str).str.strip().str.upper() == "CCREC"
                        ].copy()
                        if df_atf.empty:
                            st.warning("⚠️ No records found in ATF file with Process Code = 'CCREC'.")
                            return
                    else:
                        st.info("ℹ️ 'Process Code' column not found in ATF file. Processing all records.")

                    if "Invoice Number" not in df_atf.columns:
                        st.error("❌ Error: 'Invoice Number' column not found in ATF File.")
                        return

                    # Clean Invoice Numbers in ATF
                    df_atf["Original Invoice"] = clean_invoice_series(df_atf["Invoice Number"])

                    # 3. Optional Matching with Requested Correction File
                    if correction_file is not None:
                        corr_data = load_file_buffer(correction_file)
                        corr_original_invs = set()
                        ignore_sheets = ["Participant Earning Summary MAR"]

                        if isinstance(corr_data, dict):
                            for sheet_name, df_sheet in corr_data.items():
                                if str(sheet_name).strip().upper() in [x.upper() for x in ignore_sheets]:
                                    continue
                                df_sheet.columns = df_sheet.columns.str.strip()
                                if "Invoice Number" in df_sheet.columns:
                                    cleaned_invs = clean_invoice_series(df_sheet["Invoice Number"]).dropna()
                                    corr_original_invs.update(cleaned_invs)
                        else:
                            corr_data.columns = corr_data.columns.str.strip()
                            if "Invoice Number" in corr_data.columns:
                                cleaned_invs = clean_invoice_series(corr_data["Invoice Number"]).dropna()
                                corr_original_invs.update(cleaned_invs)

                        if corr_original_invs:
                            df_atf = df_atf[df_atf["Original Invoice"].isin(corr_original_invs)].copy()
                            if df_atf.empty:
                                st.warning("⚠️ No matching invoices found between ATF File and Requested Correction File.")
                                return

                    # 4. Priority Ranking & Extract Latest Invoices
                    df_atf["Priority_Rank"] = df_atf["Invoice Number"].apply(get_invoice_rank)

                    if "Transaction Amount" in df_atf.columns:
                        df_atf["Temp_Amount"] = (
                            pd.to_numeric(df_atf["Transaction Amount"], errors="coerce")
                            .round(2)
                            .abs()
                        )
                        max_ranks = df_atf.groupby(["Original Invoice", "Temp_Amount"], dropna=False)["Priority_Rank"].transform("max")
                        df_atf.drop(columns=["Temp_Amount"], inplace=True)
                    else:
                        max_ranks = df_atf.groupby("Original Invoice")["Priority_Rank"].transform("max")

                    latest_invoices_df = df_atf[df_atf["Priority_Rank"] == max_ranks].copy()
                    latest_invoices_df.drop(columns=["Priority_Rank", "Original Invoice"], errors="ignore", inplace=True)

                    # 5. Prepare Excel and CSV in-memory buffers
                    excel_buffer = io.BytesIO()
                    with pd.ExcelWriter(excel_buffer, engine="openpyxl") as writer:
                        latest_invoices_df.to_excel(writer, sheet_name="Latest_Invoices", index=False)
                    excel_data = excel_buffer.getvalue()

                    csv_data = latest_invoices_df.to_csv(index=False, encoding="utf-8-sig")

                    # Store in Session State so download buttons persist across reruns
                    st.session_state.processed_results = {
                        "df_latest": latest_invoices_df,
                        "excel_data": excel_data,
                        "csv_data": csv_data,
                        "msg": f"🎉 Successfully extracted {len(latest_invoices_df)} latest invoice records!"
                    }

                except Exception as e:
                    st.error(f"❌ Error during processing: {e}")

    # --- RENDER PERSISTENT DOWNLOAD BUTTONS & PREVIEW ---
    if st.session_state.processed_results is not None:
        results = st.session_state.processed_results
        st.success(results["msg"])

        st.subheader("📋 Preview Latest Invoices")
        st.dataframe(results["df_latest"], use_container_width=True)

        col1, col2 = st.columns(2)
        with col1:
            st.download_button(
                label="📥 Download Latest Invoices (.xlsx)",
                data=results["excel_data"],
                file_name="Latest_Invoices_ATF.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
                key="btn_dl_excel"
            )
        with col2:
            st.download_button(
                label="📥 Download Latest Invoices (.csv)",
                data=results["csv_data"],
                file_name="Latest_Invoices_ATF.csv",
                mime="text/csv",
                use_container_width=True,
                key="btn_dl_csv"
            )

if __name__ == "__main__":
    main()
