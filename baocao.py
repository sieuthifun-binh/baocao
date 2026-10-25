import streamlit as st
import pandas as pd
import io

st.set_page_config(
    page_title="Hệ thống So sánh & Đối soát Dữ liệu Excel",
    layout="wide"
)

st.title("📊 Hệ thống So sánh & Đối soát Dữ liệu Excel")
st.markdown("Tải lên **File dữ liệu nguồn** và **File dữ liệu gốc** để kiểm tra và đối soát thông tin.")

# --- BƯỚC 1: UPLOAD FILE ---
st.header("1. Tải lên dữ liệu")
col1, col2 = st.columns(2)

with col1:
    source_file = st.file_uploader("Chọn File Dữ liệu Nguồn (.xlsx, .xls)", type=["xlsx", "xls"], key="source")

with col2:
    master_file = st.file_uploader("Chọn File Dữ liệu Gốc (.xlsx, .xls)", type=["xlsx", "xls"], key="master")

# --- HÀM CHUẨN HÓA DỮ LIỆU ---
def normalize_string(val):
    if pd.isna(val) or val is None:
        return ""
    # Chuyển về chuỗi, loại bỏ khoảng trắng thừa, viết hoa để so sánh
    val_str = str(val).strip()
    # Nếu là dạng float như "123456789.0" thì xóa ".0"
    if val_str.endswith(".0"):
        val_str = val_str[:-2]
    return " ".join(val_str.split()).upper()

def normalize_date(val):
    if pd.isna(val) or val is None:
        return ""
    try:
        # Chuyển đổi về định dạng ngày YYYY-MM-DD
        dt = pd.to_datetime(val, dayfirst=True, errors='coerce')
        if pd.notna(dt):
            return dt.strftime('%Y-%m-%d')
    except Exception:
        pass
    return str(val).strip()

def normalize_id(val):
    if pd.isna(val) or val is None:
        return ""
    val_str = str(val).strip()
    if val_str.endswith(".0"):
        val_str = val_str[:-2]
    # Bổ sung số 0 ở đầu nếu CCCD bị mất số 0 (đủ 12 số hoặc 9 số)
    if len(val_str) == 11 and val_str.isdigit():
        val_str = "0" + val_str
    elif len(val_str) == 8 and val_str.isdigit():
        val_str = "0" + val_str
    return val_str

if source_file and master_file:
    # Đọc danh sách sheet
    xls_source = pd.ExcelFile(source_file)
    xls_master = pd.ExcelFile(master_file)
    
    st.header("2. Cấu hình Cột & Ánh xạ Dữ liệu")
    col_config1, col_config2 = st.columns(2)
    
    with col_config1:
        st.subheader("📌 File Nguồn")
        sheet_source = st.selectbox("Chọn Sheet File Nguồn:", xls_source.sheet_names, key="s_sheet")
        df_source_raw = pd.read_excel(source_file, sheet_name=sheet_source)
        
        cols_source = list(df_source_raw.columns)
        src_name = st.selectbox("Cột Họ và Tên (Nguồn):", cols_source, index=0 if len(cols_source) > 0 else 0)
        src_dob = st.selectbox("Cột Ngày Tháng Năm Sinh (Nguồn):", cols_source, index=1 if len(cols_source) > 1 else 0)
        src_id = st.selectbox("Cột Số CCCD/CMND (Nguồn):", cols_source, index=2 if len(cols_source) > 2 else 0)

    with col_config2:
        st.subheader("📌 File Dữ liệu Gốc")
        sheet_master = st.selectbox("Chọn Sheet File Gốc:", xls_master.sheet_names, key="m_sheet")
        df_master_raw = pd.read_excel(master_file, sheet_name=sheet_master)
        
        cols_master = list(df_master_raw.columns)
        mst_name = st.selectbox("Cột Họ và Tên (Gốc):", cols_master, index=0 if len(cols_master) > 0 else 0)
        mst_dob = st.selectbox("Cột Ngày Tháng Năm Sinh (Gốc):", cols_master, index=1 if len(cols_master) > 1 else 0)
        mst_id = st.selectbox("Cột Số CCCD/CMND (Gốc):", cols_master, index=2 if len(cols_master) > 2 else 0)

    if st.button("🚀 Thực hiện So sánh"):
        with st.spinner("Đang xử lý và so sánh dữ liệu..."):
            # Copy dataframe để xử lý
            df_src = df_source_raw.copy()
            df_mst = df_master_raw.copy()

            # Tạo cột chuẩn hóa
            df_src['_norm_name'] = df_src[src_name].apply(normalize_string)
            df_src['_norm_dob'] = df_src[src_dob].apply(normalize_date)
            df_src['_norm_id'] = df_src[src_id].apply(normalize_id)

            df_mst['_norm_name'] = df_mst[mst_name].apply(normalize_string)
            df_mst['_norm_dob'] = df_mst[mst_dob].apply(normalize_date)
            df_mst['_norm_id'] = df_mst[mst_id].apply(normalize_id)

            # Tạo khóa tổng hợp
            df_src['_key_exact'] = df_src['_norm_name'] + "|" + df_src['_norm_dob'] + "|" + df_src['_norm_id']
            df_mst['_key_exact'] = df_mst['_norm_name'] + "|" + df_mst['_norm_dob'] + "|" + df_mst['_norm_id']

            # Set dùng để tra cứu nhanh từ dữ liệu gốc
            master_exact_keys = set(df_mst['_key_exact'])
            master_id_map = df_mst.set_index('_norm_id').to_dict(orient='index')

            # Phân loại kết quả
            exact_matches = []
            mismatches = []
            not_found = []

            for idx, row in df_src.iterrows():
                key = row['_key_exact']
                cccd = row['_norm_id']
                
                # 1. Trùng khớp hoàn toàn cả 3 trường
                if key in master_exact_keys:
                    row_dict = row.to_dict()
                    row_dict['Trạng thái'] = "Trùng khớp hoàn toàn"
                    exact_matches.append(row_dict)
                
                # 2. Tìm thấy theo CCCD nhưng sai Tên hoặc Ngày sinh
                elif cccd in master_id_map and cccd != "":
                    matched_mst = master_id_map[cccd]
                    row_dict = row.to_dict()
                    row_dict['Trạng thái'] = "Sai lệch thông tin"
                    row_dict['Tên ở File Gốc'] = matched_mst.get(mst_name, "")
                    row_dict['Ngày sinh ở File Gốc'] = matched_mst.get(mst_dob, "")
                    row_dict['CCCD ở File Gốc'] = matched_mst.get(mst_id, "")
                    mismatches.append(row_dict)
                
                # 3. Không tìm thấy thông tin trong file gốc
                else:
                    row_dict = row.to_dict()
                    row_dict['Trạng thái'] = "Không có trong file gốc"
                    not_found.append(row_dict)

            # Chuyển đổi sang DataFrame
            df_exact = pd.DataFrame(exact_matches)
            df_mismatch = pd.DataFrame(mismatches)
            df_notfound = pd.DataFrame(not_found)

            # Xóa các cột hỗ trợ kỹ thuật trước khi hiển thị
            helper_cols = ['_norm_name', '_norm_dob', '_norm_id', '_key_exact']
            for df_res in [df_exact, df_mismatch, df_notfound]:
                if not df_res.empty:
                    df_res.drop(columns=[c for c in helper_cols if c in df_res.columns], inplace=True)

            # --- HIỂN THỊ KẾT QUẢ ---
            st.header("3. Kết quả So sánh")
            
            c1, c2, c3 = st.columns(3)
            c1.metric("Tổng số dòng nguồn", len(df_src))
            c2.metric("Trùng khớp hoàn toàn", len(df_exact))
            c3.metric("Lệch thông tin / Không thấy", len(df_mismatch) + len(df_notfound))

            tab1, tab2, tab3 = st.tabs([
                f"✅ Trùng khớp ({len(df_exact)})", 
                f"⚠️ Sai lệch thông tin ({len(df_mismatch)})", 
                f"❌ Không tìm thấy ({len(df_notfound)})"
            ])

            with tab1:
                st.dataframe(df_exact, use_container_width=True)

            with tab2:
                st.write("Các bản ghi khớp CCCD nhưng lệch Họ tên hoặc Ngày sinh:")
                st.dataframe(df_mismatch, use_container_width=True)

            with tab3:
                st.dataframe(df_notfound, use_container_width=True)

            # --- TẠO FILE EXCEL TẢI VỀ ---
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                if not df_exact.empty:
                    df_exact.to_excel(writer, sheet_name='Trung_Khop', index=False)
                if not df_mismatch.empty:
                    df_mismatch.to_excel(writer, sheet_name='Sai_Lech_Thong_Tin', index=False)
                if not df_notfound.empty:
                    df_notfound.to_excel(writer, sheet_name='Khong_Tim_Thay', index=False)

            output.seek(0)
            
            st.download_button(
                label="📥 Tải xuống Báo cáo Kết quả Đối soát (.xlsx)",
                data=output,
                file_name="Bao_Cao_Doi_Soat_Du_Lieu.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
