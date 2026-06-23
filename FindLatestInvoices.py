import streamlit as st
import pandas as pd
import re
import io

def clean_original_invoice(inv):
    """Xoá bỏ các phần COR, REV và các số theo sau để lấy Original Invoice"""
    if pd.isnull(inv):
        return ""
    inv = str(inv).strip()
    pattern = re.compile(r'[\s\-]*(COR|REV)\d*$', re.IGNORECASE)
    cleaned = re.sub(pattern, '', inv).strip()
    return cleaned.rstrip('- ')

def parse_suffix_for_ranking(inv):
    """Phân loại để tìm ra Invoice mới nhất theo 3 cấp độ: Đuôi số -> Loại -> Số lượng gạch ngang"""
    inv = str(inv).upper().strip()
    match = re.search(r'([\s\-]*)(COR|REV)(\d*)$', inv)
    if not match:
        return (0, 0, 0)
    
    separator = match.group(1)
    type_str = match.group(2)
    num_str = match.group(3)
    
    type_val = 2 if type_str == 'COR' else 1
    num_val = int(num_str) if num_str else 1
    dash_count = separator.count('-')
    
    return (num_val, type_val, dash_count)

def increment_or_append_suffix(val, suffix_type):
    """Gắn hậu tố mới chuẩn 1 gạch ngang"""
    if pd.isnull(val): return val
    val = str(val).strip()
    match = re.search(r'(?i)(.*?)(?:[\s\-]*)(COR|REV)(\d*)$', val)
    if match:
        prefix = match.group(1).rstrip('- ') 
        num_str = match.group(3)
        current_num = int(num_str) if num_str else 1
        next_num = current_num + 1
        return f"{prefix}-{suffix_type}{next_num}"
    else:
        return f"{val.rstrip('- ')}-{suffix_type}"

def replace_cor_with_rev(val):
    """Thay thế COR thành REV để đồng bộ"""
    if pd.isnull(val): return val
    val = str(val)
    return re.sub(r'COR(\d*)$', r'REV\1', val)

def main():
    st.set_page_config(page_title="Invoice Correction Tool", layout="wide")
    st.title("Find the Latest Invoices from ATF")
    
    # [YÊU CẦU 4]: Sử dụng session_state để các nút download không bị biến mất
    if 'processed' not in st.session_state:
        st.session_state.processed = False
        st.session_state.excel_data = None
        st.session_state.csv_data = None

    col1, col2 = st.columns(2)
    with col1:
        req_file = st.file_uploader("1. Upload Requested Correction file", type=['xlsx', 'xls', 'xlsb'])
    with col2:
        atf_file = st.file_uploader("2. Upload ATF file", type=['xlsx', 'xls', 'xlsb'])

    if st.button("Start Data Processing"):
        if not req_file or not atf_file:
            st.error("Error: You need to upload 2 requested files before start!")
            return
            
        progress_bar = st.progress(0)
        status_text = st.empty()

        try:
            # Hàm phụ trợ để xử lý cho cả định dạng xlsx, xls và xlsb
            def load_excel(file_upload):
                if file_upload.name.endswith('.xlsb'):
                    return pd.read_excel(file_upload, sheet_name=0, engine='pyxlsb')
                else:
                    return pd.read_excel(file_upload, sheet_name=0)

            status_text.text("Scanning ATF file...")
            df_req = load_excel(req_file)
            df_atf = load_excel(atf_file)
            progress_bar.progress(20)

            # --- XỬ LÝ REQUESTED CORRECTION FILE ---
            status_text.text("Processing Requested Correction file...")
            df_req['Text'] = df_req['Invoice Number'].apply(lambda x: isinstance(x, str))

            def format_invoice(row):
                val = row['Invoice Number']
                is_text = row['Text']
                if not is_text and pd.notnull(val):
                    return '00' + str(int(val))
                return str(val) if pd.notnull(val) else val

            df_req['Invoice Number'] = df_req.apply(format_invoice, axis=1)
            df_req['Original Invoice'] = df_req['Invoice Number'].apply(clean_original_invoice)
            original_invoices_memory = df_req['Original Invoice'].drop_duplicates().tolist()

            progress_bar.progress(40)

            # --- XỬ LÝ ATF FILE ---
            status_text.text("Matching data against ATF...")
            df_atf['Original Invoice'] = df_atf['Invoice Number'].apply(clean_original_invoice)
            matched_atf = df_atf[df_atf['Original Invoice'].isin(original_invoices_memory)].copy()

            matched_atf['SortKey'] = matched_atf['Invoice Number'].apply(parse_suffix_for_ranking)
            
            # Làm sạch dữ liệu tiền để nhóm đúng số lượng dòng (giải quyết triệt để 13 dòng của 0043905960)
            if 'Transaction Amount' in matched_atf.columns:
                matched_atf['Temp_Amount'] = pd.to_numeric(matched_atf['Transaction Amount'], errors='coerce').abs()
                max_sort_keys = matched_atf.groupby(['Original Invoice', 'Temp_Amount'], dropna=False)['SortKey'].transform('max')
                matched_atf.drop(columns=['Temp_Amount'], inplace=True)
            else:
                max_sort_keys = matched_atf.groupby('Original Invoice')['SortKey'].transform('max')

            latest_atf = matched_atf[matched_atf['SortKey'] == max_sort_keys].copy()

            if latest_atf.empty:
                st.warning("Can't find any matching data!")
                progress_bar.progress(100)
                st.session_state.processed = False
                return

            progress_bar.progress(60)

            # --- TẠO DỮ LIỆU COR & REV ---
            status_text.text("Creating COR and REV sheets...")
            df_cor = latest_atf.copy()
            df_rev = latest_atf.copy()

            # --- XỬ LÝ DỮ LIỆU COR ---
            df_cor['Transaction Number'] = df_cor['Transaction Number'].apply(lambda x: increment_or_append_suffix(x, 'COR'))
            df_cor['Invoice Number'] = df_cor['Invoice Number'].apply(lambda x: increment_or_append_suffix(x, 'COR'))
            df_cor['Transaction Type'] = 'MANUAL_ADJ'
            # Cột Vertical được giữ nguyên bản từ ATF

            currency_cols = ['Transaction Amount', 'EUR Value', 'CAD Value', 'GBP Value', 'Native Currency', 'AUD Value']
            for col in currency_cols:
                if col in df_cor.columns:
                    df_cor[col] = pd.to_numeric(df_cor[col], errors='coerce').abs()

            # --- XỬ LÝ DỮ LIỆU REV ---
            df_rev['Transaction Number'] = df_cor['Transaction Number'].apply(replace_cor_with_rev)
            df_rev['Invoice Number'] = df_cor['Invoice Number'].apply(replace_cor_with_rev)
            df_rev['Transaction Type'] = 'MANUAL_ADJ'

            for col in currency_cols:
                if col in df_rev.columns:
                    df_rev[col] = df_cor[col] * -1

            # --- DỌN DẸP CÁC CỘT THỪA (Bao gồm YÊU CẦU 1: Xóa cột Period) ---
            cols_to_drop = ['SortKey', 'Original Invoice', 'Period']
            for df in [df_cor, df_rev]:
                df.drop(columns=cols_to_drop, errors='ignore', inplace=True)

            progress_bar.progress(80)
            status_text.text("Merging COR and REV sheets...")

            # --- [YÊU CẦU 2 & 3]: GỘP COR & REV THÀNH SHEET "Upload" ---
            df_upload = pd.concat([df_cor, df_rev], ignore_index=True)

            # Format lại ID columns (loại bỏ thập phân)
            format_cols = ['Source Business Unit ID', 'Business Unit ID']
            for col in format_cols:
                if col in df_upload.columns:
                    df_upload[col] = pd.to_numeric(df_upload[col], errors='coerce').astype('Int64')

            # --- XUẤT RA EXCEL & CSV BẰNG BUFFER ---
            # 1. Output Excel (Chỉ có duy nhất 1 sheet "Upload")
            excel_buffer = io.BytesIO()
            with pd.ExcelWriter(excel_buffer, engine='openpyxl') as writer:
                df_upload.to_excel(writer, sheet_name='Upload', index=False)
            
            # 2. Output CSV
            csv_buffer = df_upload.to_csv(index=False).encode('utf-8-sig')

            # Lưu vào Session State
            st.session_state.excel_data = excel_buffer.getvalue()
            st.session_state.csv_data = csv_buffer
            st.session_state.processed = True

            progress_bar.progress(100)
            status_text.success("Completed. Results are ready to download.")

        except Exception as e:
            st.error(f"Process is completed: {e}")
            progress_bar.empty()
            st.session_state.processed = False

    # --- HIỂN THỊ CÁC NÚT DOWNLOAD ---
    if st.session_state.processed:
        col_btn1, col_btn2 = st.columns(2)
        with col_btn1:
            st.download_button(
                label="📥 Download Excel (.xlsx)",
                data=st.session_state.excel_data,
                file_name="Upload file.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        with col_btn2:
            st.download_button(
                label="📥 Download CSV (.csv)",
                data=st.session_state.csv_data,
                file_name="Upload file.csv",
                mime="text/csv"
            )

if __name__ == "__main__":
    main()
