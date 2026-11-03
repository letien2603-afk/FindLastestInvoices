import io
import re
import pandas as pd
import streamlit as st


def format_invoice_number(val):
  """Ensures invoice numbers are text and carry the '00' prefix if numeric."""
  if pd.isnull(val):
    return ""
  s = str(val).strip()
  if s.endswith(".0"):
    s = s[:-2]
  if s.isdigit() and not s.startswith("00"):
    return "00" + s
  return s


def clean_original_invoice(inv):
  """Removes COR and REV suffixes to get the base invoice number."""
  if pd.isnull(inv):
    return ""
  inv = str(inv).strip()
  pattern = re.compile(r"[\s-]*(COR|REV)\d*\$", re.IGNORECASE)
  cleaned = re.sub(pattern, "", inv).strip().rstrip("- ")
  return cleaned


def parse_suffix_for_ranking(inv):
  """Ranks invoice versions to find the latest transaction record."""
  inv = str(inv).upper().strip()
  match = re.search(r"([\s-]*)(COR|REV)(\d*)\$", inv)
  if not match:
    return (0, 0, 0)

  separator = match.group(1)
  type_str = match.group(2)
  num_str = match.group(3)

  type_val = 2 if type_str == "COR" else 1
  num_val = int(num_str) if num_str else 1
  dash_count = separator.count("-")

  return (num_val, type_val, dash_count)


def increment_or_append_suffix(val, suffix_type):
  """Increments existing COR/REV version number or appends new suffix."""
  if pd.isnull(val):
    return val
  val = str(val).strip()
  match = re.search(r"(?i)(.*?)(?:[\s-]*)(COR|REV)(\d*)\$", val)
  if match:
    prefix = match.group(1).rstrip("- ")
    num_str = match.group(3)
    current_num = int(num_str) if num_str else 1
    next_num = current_num + 1
    return f"{prefix}-{suffix_type}{next_num}"
  else:
    return f"{val.rstrip('- ')}-{suffix_type}"


def replace_cor_with_rev(val):
  """Replaces COR with REV in suffix strings."""
  if pd.isnull(val):
    return val
  return re.sub(r"COR", "REV", str(val), flags=re.IGNORECASE)


def load_file_data(uploaded_file):
  """Loads Excel or delimited text files directly from Streamlit uploaded file buffers."""
  filename = uploaded_file.name.lower()
  if filename.endswith(".xlsb"):
    return pd.read_excel(
        uploaded_file,
        sheet_name=0,
        engine="pyxlsb",
        dtype={"Invoice Number": str},
    )
  elif filename.endswith((".xlsx", ".xls")):
    return pd.read_excel(
        uploaded_file, sheet_name=0, dtype={"Invoice Number": str}
    )
  else:
    try:
      return pd.read_csv(
          uploaded_file, sep=None, engine="python", dtype=str
      )
    except Exception:
      uploaded_file.seek(0)
      return pd.read_excel(
          uploaded_file, sheet_name=0, dtype={"Invoice Number": str}
      )


def main():
  st.set_page_config(
      page_title="Vertical Bulk Correction Tool", page_icon="📊", layout="wide"
  )

  st.title("📊 Vertical Bulk Correction Automation")
  st.markdown(
      "Upload your **Requested Correction File** and **ATF File** below to"
      " generate consolidated `COR` and `REV` upload entries."
  )

  st.sidebar.header("📁 File Uploads")
  req_file = st.sidebar.file_uploader(
      "Requested Correction File", type=["xlsx", "xls", "xlsb", "csv", "txt"]
  )
  atf_file = st.sidebar.file_uploader(
      "ATF File", type=["xlsx", "xls", "xlsb", "csv", "txt"]
  )

  user_comment = st.text_input(
      "Comment (Optional)",
      value="",
      help="Applied to the Comments field in output records.",
  )

  # Initialize session state for persistent results
  if "processed_results" not in st.session_state:
    st.session_state.processed_results = None

  if st.button("🚀 Process Bulk Correction", type="primary"):
    # Reset results when button is re-clicked
    st.session_state.processed_results = None

    if not req_file or not atf_file:
      st.error("Please upload both the Requested Correction File and ATF File.")
    else:
      with st.spinner("Processing files..."):
        try:
          # 1. Load Files
          df_req = load_file_data(req_file)
          df_atf = load_file_data(atf_file)

          df_req.columns = df_req.columns.str.strip()
          df_atf.columns = df_atf.columns.str.strip()

          # 2. Filter Process Code = 'CCREC' First
          if "Process Code" in df_atf.columns:
            df_atf = df_atf[
                df_atf["Process Code"].astype(str).str.strip().str.upper()
                == "CCREC"
            ].copy()
            if df_atf.empty:
              st.warning(
                  "No records found in ATF file with Process Code = 'CCREC'."
              )
              return
          else:
            st.info(
                "'Process Code' column not found in ATF file. Proceeding without"
                " filtering."
            )

          # 3. Standardize Invoice Numbers ('00' Prefix)
          df_req["Invoice Number"] = df_req["Invoice Number"].apply(
              format_invoice_number
          )
          df_atf["Invoice Number"] = df_atf["Invoice Number"].apply(
              format_invoice_number
          )

          df_req["Original Invoice"] = df_req["Invoice Number"].apply(
              clean_original_invoice
          )
          original_invoices_memory = (
              df_req["Original Invoice"].drop_duplicates().tolist()
          )

          if "Corrected Vertical" in df_req.columns:
            vertical_mapping = dict(
                zip(df_req["Original Invoice"], df_req["Corrected Vertical"])
            )
          else:
            vertical_mapping = {}

          # 4. Match ATF Records
          df_atf["Original Invoice"] = df_atf["Invoice Number"].apply(
              clean_original_invoice
          )
          matched_atf = df_atf[
              df_atf["Original Invoice"].isin(original_invoices_memory)
          ].copy()

          if matched_atf.empty:
            st.warning(
                "No matching invoice numbers found between ATF and Requested"
                " file."
            )
            return

          # 5. Rank & Select Latest Version
          matched_atf["SortKey"] = matched_atf["Invoice Number"].apply(
              parse_suffix_for_ranking
          )
          matched_atf["Temp_Amount"] = (
              pd.to_numeric(matched_atf["Transaction Amount"], errors="coerce")
              .round(2)
              .abs()
          )

          max_sort_keys = matched_atf.groupby(
              ["Original Invoice", "Temp_Amount"], dropna=False
          )["SortKey"].transform("max")
          matched_atf.drop(columns=["Temp_Amount"], inplace=True)
          latest_atf = matched_atf[
              matched_atf["SortKey"] == max_sort_keys
          ].copy()

          # 6. Skip If Vertical Already Matches
          col_vertical_atf = "Vertical"
          if col_vertical_atf in latest_atf.columns:
            latest_atf["Req_Vertical"] = latest_atf["Original Invoice"].map(
                vertical_mapping
            )
            val_atf = (
                latest_atf[col_vertical_atf]
                .astype(str)
                .str.strip()
                .str.lower()
            )
            val_req = (
                latest_atf["Req_Vertical"].astype(str).str.strip().str.lower()
            )
            latest_atf = latest_atf[val_atf != val_req].copy()
            latest_atf.drop(columns=["Req_Vertical"], inplace=True)

          if latest_atf.empty:
            st.warning(
                "Verticals in all requested invoices have been updated to match"
                " with the requested file or no matching invoices to process."
            )
            return

          # 7. Generate COR & REV Entries
          df_cor = latest_atf.copy()
          df_rev = latest_atf.copy()

          # COR
          df_cor["Transaction Number"] = df_cor["Transaction Number"].apply(
              lambda x: increment_or_append_suffix(x, "COR")
          )
          df_cor["Invoice Number"] = df_cor["Invoice Number"].apply(
              lambda x: increment_or_append_suffix(x, "COR")
          )
          df_cor["Transaction Type"] = "MANUAL_ADJ"
          df_cor["Vertical"] = df_cor["Original Invoice"].map(vertical_mapping)

          # REV
          df_rev["Transaction Number"] = df_cor["Transaction Number"].apply(
              replace_cor_with_rev
          )
          df_rev["Invoice Number"] = df_cor["Invoice Number"].apply(
              replace_cor_with_rev
          )
          df_rev["Transaction Type"] = "MANUAL_ADJ"

          cols_to_invert = [
              "Transaction Amount",
              "EUR Value",
              "CAD Value",
              "GBP Value",
              "Native Currency",
              "AUD Value",
          ]
          for col in cols_to_invert:
            if col in df_rev.columns:
              df_rev[col] = pd.to_numeric(df_rev[col], errors="coerce") * -1

          # Comments & Cleanup
          for df in [df_cor, df_rev]:
            if "Period" in df.columns:
              df.drop(columns=["Period"], inplace=True)

            if user_comment:
              if "Comments" in df.columns:
                df["Comments"] = user_comment
              elif "Comment" in df.columns:
                df["Comment"] = user_comment
              else:
                df["Comments"] = user_comment

            df.drop(
                columns=["SortKey", "Original Invoice"],
                errors="ignore",
                inplace=True,
            )

          # Consolidate
          df_upload = pd.concat([df_cor, df_rev], ignore_index=True)

          format_cols = ["Source Business Unit ID", "Business Unit ID"]
          for col in format_cols:
            if col in df_upload.columns:
              df_upload[col] = pd.to_numeric(
                  df_upload[col], errors="coerce"
              ).astype("Int64")

          # Prepare Downloads
          excel_buffer = io.BytesIO()
          with pd.ExcelWriter(excel_buffer, engine="openpyxl") as writer:
            df_upload.to_excel(writer, sheet_name="Upload", index=False)
          excel_data = excel_buffer.getvalue()

          csv_data = df_upload.to_csv(index=False, encoding="utf-8-sig")

          # Store results in session_state to persist across download reruns
          st.session_state.processed_results = {
              "df_upload": df_upload,
              "excel_data": excel_data,
              "csv_data": csv_data,
              "msg": (
                  f"🎉 Success! Generated {len(df_upload)} records"
                  f" ({len(df_cor)} COR + {len(df_rev)} REV)."
              ),
          }

        except Exception as e:
          st.error(f"An error occurred during processing: {e}")

  # Render results if present in session_state
  if st.session_state.processed_results is not None:
    results = st.session_state.processed_results
    st.success(results["msg"])

    # Preview Data
    st.subheader("Preview Output Data")
    st.dataframe(results["df_upload"], use_container_width=True)

    # Download Buttons
    col1, col2 = st.columns(2)
    with col1:
      st.download_button(
          label="📥 Download Excel (.xlsx)",
          data=results["excel_data"],
          file_name="Vertical Bulk Corrections.xlsx",
          mime=(
              "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
          ),
          use_container_width=True,
          key="btn_download_excel",
      )
    with col2:
      st.download_button(
          label="📥 Download CSV (.csv)",
          data=results["csv_data"],
          file_name="Vertical Bulk Corrections.csv",
          mime="text/csv",
          use_container_width=True,
          key="btn_download_csv",
      )


if __name__ == "__main__":
  main()
