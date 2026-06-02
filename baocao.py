import streamlit as st
import pandas as pd
import plotly.express as px

# 1. Cấu hình giao diện Web
st.set_page_config(page_title="Dashboard Phân Tích BHXH", layout="wide")
st.title("📊 HỆ THỐNG PHÂN TÍCH DỮ LIỆU C12 THÔNG MINH")
st.markdown("---")

# 2. Tạo VÙNG UP FILE THÔNG MINH (Hỗ trợ từ Excel 2003 .xls, .xlsx đến .csv)
st.subheader("📂 Nạp Dữ Liệu Báo Cáo")
uploaded_file = st.file_uploader(
    label="Kéo và thả file báo cáo của bạn vào đây (Chấp nhận file .xlsx, .xls, .csv)",
    type=["xlsx", "xls", "csv"],
    help="Hỗ trợ tất cả các định dạng Excel từ phiên bản 2003 đến nay."
)

# Hàm đọc dữ liệu đa định dạng
def load_data(file):
    file_name = file.name.lower()
    
    # Trường hợp 1: File Excel hiện đại (.xlsx) hoặc file Excel cổ điển (.xls)
    if file_name.endswith('.xlsx') or file_name.endswith('.xls'):
        # Thư viện engine='xlrd' sẽ xử lý file .xls (Excel 2003)
        # Thư viện mặc định openpyxl sẽ xử lý file .xlsx
        engine_type = 'xlrd' if file_name.endswith('.xls') else None
        df = pd.read_excel(file, engine=engine_type)
        
    # Trường hợp 2: File CSV (như file hiện tại của bạn)
    elif file_name.endswith('.csv'):
        df = pd.read_csv(file)
    else:
        st.error("Định dạng file không được hỗ trợ!")
        return None
        
    # Chuẩn hóa dữ liệu số an toàn để tránh lỗi tính toán
    if 'sld' in df.columns:
        df['sld'] = pd.to_numeric(df['sld'], errors='coerce').fillna(0)
    if 'du_ck' in df.columns:
        df['du_ck'] = pd.to_numeric(df['du_ck'], errors='coerce').fillna(0)
        
    return df

# 3. Kiểm tra xem người dùng đã up file lên chưa
if uploaded_file is not None:
    try:
        df = load_data(uploaded_file)
        
        # Kiểm tra cấu trúc file có chứa các cột cốt lõi không
        required_cols = ['makhoi', 'sld', 'du_ck', 'madvi']
        missing_cols = [col for col in required_cols if col not in df.columns]
        
        if missing_cols:
            st.error(f"❌ File up lên không đúng mẫu chuẩn! Thiếu các cột bắt buộc: {', '.join(missing_cols)}")
        else:
            st.success(f"🎉 Tải file '{uploaded_file.name}' thành công! Đang xử lý dữ liệu...")
            
            # --- PHẦN XỬ LÝ BÁO CÁO TỰ ĐỘNG ---
            st.sidebar.header("Bộ Lọc Dữ Liệu")
            all_blocks = ["Tất cả"] + list(df['makhoi'].dropna().unique())
            selected_block = st.sidebar.selectbox("Chọn Mã Khối cần xem:", all_blocks)

            if selected_block != "Tất cả":
                filtered_df = df[df['makhoi'] == selected_block]
            else:
                filtered_df = df

            # Tính toán tổng hợp theo mã khối
            summary_df = df.groupby('makhoi').agg(
                Tong_Lao_Dong=('sld', 'sum'),
                Tong_No_Cuoi_Ky=('du_ck', 'sum'),
                So_Don_Vi=('madvi', 'count')
            ).reset_index()

            # Hiển thị KPI Số lớn
            st.markdown("### 📌 Số liệu tổng hợp kỳ báo cáo")
            col1, col2, col3 = st.columns(3)
            with col1:
                st.metric(label="Tổng Số Lao Động", value=f"{int(filtered_df['sld'].sum()):,}")
            with col2:
                st.metric(label="Tổng Tiền Thiếu Cuối Kỳ", value=f"{filtered_df['du_ck'].sum():,.0f} VNĐ")
            with col3:
                st.metric(label="Tổng Số Đơn Vị", value=f"{len(filtered_df):,}")

            # Vẽ biểu đồ trực quan
            st.markdown("---")
            chart_col1, chart_col2 = st.columns(2)

            with chart_col1:
                fig_sld = px.bar(summary_df.sort_values(by='Tong_Lao_Dong', ascending=False), 
                                 x='makhoi', y='Tong_Lao_Dong', 
                                 title='Tỷ Trọng Số Lao Động Theo Từng Khối',
                                 labels={'makhoi': 'Mã Khối', 'Tong_Lao_Dong': 'Số Lao Động'},
                                 color='Tong_Lao_Dong', color_continuous_scale='Blues')
                st.plotly_chart(fig_sld, use_container_width=True)

            with chart_col2:
                fig_no = px.bar(summary_df.sort_values(by='Tong_No_Cuoi_Ky', ascending=False), 
                                x='makhoi', y='Tong_No_Cuoi_Ky', 
                                title='Tình Hình Nợ/Thiếu Cuối Kỳ Theo Khối',
                                labels={'makhoi': 'Mã Khối', 'Tong_No_Cuoi_Ky': 'Số Tiền Thiếu'},
                                color='Tong_No_Cuoi_Ky', color_continuous_scale='Reds')
                st.plotly_chart(fig_no, use_container_width=True)

            # Hiện bảng dữ liệu chi tiết
            st.subheader("📋 Bảng Tổng Hợp Chi Tiết Theo Mã Khối")
            st.dataframe(summary_df.style.format({
                'Tong_Lao_Dong': '{:,.0f}',
                'Tong_No_Cuoi_Ky': '{:,.0f} VNĐ',
                'So_Don_Vi': '{:,.0f}'
            }), use_container_width=True)

    except Exception as e:
        st.error(f"Có lỗi hệ thống khi đọc dữ liệu: {e}")
else:
    # Trạng thái khi chưa có file nào được up lên Web
    st.info("☝️ Vui lòng chọn hoặc kéo thả file Excel dữ liệu tháng cần báo cáo vào vùng trống phía trên.")
