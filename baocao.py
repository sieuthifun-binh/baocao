
import streamlit as st
import pandas as pd
import plotly.express as px

# 1. Cấu hình giao diện Web
st.set_page_config(page_title="Dashboard Báo Cáo BHXH", layout="wide")
st.title("📊 HỆ THỐNG BÁO CÁO THU & LAO ĐỘNG THEO KHỐI")
st.markdown("---")

# 2. Đọc dữ liệu từ file Excel (Đảm bảo file nằm cùng thư mục với app.py)
@st.cache_data
def load_data():
    # File của bạn có tên mặc định là C12_CHI_TIEU (6).xlsx
    df = pd.read_excel("C12_CHI_TIEU (6).xlsx")
    # Đảm bảo ép kiểu dữ liệu số để không bị lỗi tính toán
    df['sld'] = pd.to_numeric(df['sld'], errors='coerce').fillna(0)
    df['du_ck'] = pd.to_numeric(df['du_ck'], errors='coerce').fillna(0)
    return df

try:
    df = load_data()

    # 3. Thanh bên trái (Sidebar) để người dùng lọc nhanh
    st.sidebar.header("Bộ Lọc Dữ Liệu")
    all_blocks = ["Tất cả"] + list(df['makhoi'].dropna().unique())
    selected_block = st.sidebar.selectbox("Chọn Mã Khối cần xem:", all_blocks)

    # Lọc dữ liệu dựa trên lựa chọn
    if selected_block != "Tất cả":
        filtered_df = df[df['makhoi'] == selected_block]
    else:
        filtered_df = df

    # 4. Tính toán tổng hợp theo từng mã khối (Tính năng thông minh)
    summary_df = df.groupby('makhoi').agg(
        Tong_Lao_Dong=('sld', 'sum'),
        Tong_No_Cuoi_Ky=('du_ck', 'sum'),
        So_Don_Vi=('madvi', 'count')
    ).reset_index()

    # 5. Hiển thị các chỉ số KPI Tổng quan (Thẻ số lớn)
    st.subheader("📌 Chỉ số Tổng Toàn Hệ Thống")
    col1, col2, col3 = st.columns(3)
    with col1:
        st.metric(label="Tổng Số Lao Động", value=f"{int(filtered_df['sld'].sum()):,}")
    with col2:
        st.metric(label="Tổng Số Tiền Thiếu Cuối Kỳ", value=f"{filtered_df['du_ck'].sum():,.0f} VNĐ")
    with col3:
        st.metric(label="Tổng Số Đơn Vị", value=f"{len(filtered_df):,}")

    st.markdown("---")

    # 6. Hiển thị Biểu đồ Trực quan hóa
    st.subheader("📈 Phân Tích Kéo Thả Trực Quan")
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

    # 7. Hiển thị Bảng Dữ Liệu Chi Tiết
    st.subheader("📋 Bảng Tổng Hợp Chi Tiết Theo Mã Khối")
    st.dataframe(summary_df.style.format({
        'Tong_Lao_Dong': '{:,.0f}',
        'Tong_No_Cuoi_Ky': '{:,.0f} VNĐ',
        'So_Don_Vi': '{:,.0f}'
    }), use_container_width=True)

except Exception as e:
    st.error(f"Đã xảy ra lỗi khi đọc dữ liệu. Vui lòng kiểm tra lại cấu trúc file: {e}")
