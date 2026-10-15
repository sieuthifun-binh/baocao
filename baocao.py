import io
import pandas as pd
import plotly.express as px
import streamlit as st

st.set_page_config(
    page_title="Hệ thống Quản lý & Phân tích Thu nộp Đơn vị",
    page_icon="📊",
    layout="wide",
)

st.title("📊 Hệ thống Phân tích & Báo cáo Tình hình Thu nộp Đơn vị")
st.markdown(
    "Hệ thống tự động đọc file nguồn, cho phép cập nhật dữ liệu và phân tích"
    " chuẩn xác theo các mẫu báo cáo: **Cập nhật**, **Chi tiết**, **Nợ**, và"
    " **KH Thu**."
)

tab1, tab2, tab3, tab4, tab5 = st.tabs([
    "📥 Nhập liệu & Cập nhật",
    "📋 Báo cáo Chi tiết",
    "📉 Phân tích Đơn vị Nợ",
    "📈 Tổng hợp KH Thu",
    "📊 Dashboard Trực quan",
])

uploaded_file = st.sidebar.file_uploader(
    "Tải lên file dữ liệu Excel mẫu (.xls / .xlsx)", type=["xls", "xlsx"]
)

if uploaded_file is not None:
  try:
    xls = pd.ExcelFile(uploaded_file)
    sheet_names = xls.sheet_names

    # Đọc sheet Cập nhật an toàn
    df_capnhat = pd.read_excel(uploaded_file, sheet_name="Cập nhật")

    # Chuẩn hóa tên cột để tránh lỗi khoảng trắng thừa
    df_capnhat.columns = [str(c).strip() for c in df_capnhat.columns]

    with tab1:
      st.subheader(
          "1. Khu vực Nhập liệu & Cập nhật Dữ liệu Nguồn (Sheet Cập nhật)"
      )
      st.markdown(
          "Cấu trúc dữ liệu nguồn chuẩn 100% như file gốc. Bạn có thể chỉnh"
          " sửa, thêm bớt dòng dữ liệu trực tiếp:"
      )

      edited_df = st.data_editor(
          df_capnhat, num_rows="dynamic", use_container_width=True
      )

      # Xuất file Excel bảo toàn các sheet khác
      buffer = io.BytesIO()
      with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        edited_df.to_excel(writer, sheet_name="Cập nhật", index=False)
        for s in sheet_names:
          if s != "Cập nhật":
            df_other = pd.read_excel(uploaded_file, sheet_name=s)
            df_other.to_excel(writer, sheet_name=s, index=False)
      buffer.seek(0)

      st.download_button(
          label="📥 Tải xuống File Excel đã cập nhật đầy đủ cấu trúc",
          data=buffer,
          file_name="Du_lieu_Cap_nhat_Moi.xlsx",
          mime=(
              "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
          ),
      )

    with tab2:
      st.subheader(
          "2. Sheet Chi tiết (Lọc theo tên User / Chuyên quản cột D)"
      )
      user_col = "dsnv" if "dsnv" in edited_df.columns else None
      if user_col:
        users = edited_df[user_col].dropna().unique().tolist()
        if users:
          selected_user = st.selectbox(
              "Chọn tên User / Chuyên quản để lọc báo cáo chi tiết:", users
          )
          df_filtered_user = edited_df[edited_df[user_col] == selected_user]
          st.info(
              f"Hiển thị danh sách đơn vị do chuyên quản **{selected_user}** phụ"
              f" trách ({len(df_filtered_user)} đơn vị):"
          )

          display_cols = [
              c
              for c in [
                  "madvi",
                  "tendvi",
                  "sld",
                  "tyleno",
                  "tylenoyt",
                  "_tien_ck",
              ]
              if c in edited_df.columns
          ]
          st.dataframe(df_filtered_user[display_cols], use_container_width=True)
        else:
          st.warning("Không có dữ liệu trong cột chuyên quản.")
      else:
        st.warning("Không tìm thấy cột chuyên quản (dsnv) trong file.")

    with tab3:
      st.subheader("3. Sheet Nợ (Phân tích số liệu đơn vị nợ)")
      debt_col = "_tien_ck" if "_tien_ck" in edited_df.columns else None
      if debt_col:
        edited_df[debt_col] = pd.to_numeric(
            edited_df[debt_col], errors="coerce"
        ).fillna(0)
        df_no = edited_df[edited_df[debt_col] > 0].copy()
        st.warning(
            f"Tổng số đơn vị còn nợ tiền cuối kỳ: **{len(df_no)} đơn vị**"
        )

        no_cols = [
            c
            for c in [
                "madvi",
                "tendvi",
                "sld",
                "tyleno",
                "_tien_dk",
                "_tien_unc",
                "_tien_ck",
            ]
            if c in edited_df.columns
        ]
        st.dataframe(df_no[no_cols], use_container_width=True)
      else:
        st.warning("Không tìm thấy cột số tiền cuối kỳ (_tien_ck).")

    with tab4:
      st.subheader("4. Sheet KH Thu (Tổng hợp kế hoạch thu)")
      pst_col = "pst" if "pst" in edited_df.columns else None
      unc_col = "_tien_unc" if "_tien_unc" in edited_df.columns else None

      if pst_col and debt_col:
        edited_df[pst_col] = pd.to_numeric(
            edited_df[pst_col], errors="coerce"
        ).fillna(0)
        if unc_col:
          edited_df[unc_col] = pd.to_numeric(
              edited_df[unc_col], errors="coerce"
          ).fillna(0)

        total_kh = edited_df[pst_col].sum()
        total_thu = edited_df[unc_col].sum() if unc_col else 0
        total_no = edited_df[debt_col].sum()

        col1, col2, col3 = st.columns(3)
        col1.metric("Tổng Kế hoạch Thu", f"{total_kh:,.0f} VNĐ")
        col2.metric("Tổng Đã Thu", f"{total_thu:,.0f} VNĐ")
        col3.metric("Tổng Nợ Cuối Kỳ", f"{total_no:,.0f} VNĐ")

        st.markdown("---")
        st.markdown("### Bảng Tổng Hợp Theo Khối Đơn Vị (`makhoi`)")
        if "makhoi" in edited_df.columns:
          summary_khoi = (
              edited_df.groupby("makhoi")
              .agg({
                  "madvi": "count",
                  "sld": "sum" if "sld" in edited_df.columns else "count",
                  pst_col: "sum",
                  debt_col: "sum",
              })
              .reset_index()
          )
          summary_khoi.columns = [
              "Khối đơn vị",
              "Số lượng đơn vị",
              "Tổng lao động",
              "Tổng kế hoạch thu",
              "Tổng nợ",
          ]
          st.dataframe(summary_khoi, use_container_width=True)
      else:
        st.warning("Thiếu cột dữ liệu tính toán kế hoạch thu.")

    with tab5:
      st.subheader("5. Biểu đồ Trực quan hóa Dữ liệu")
      if "makhoi" in edited_df.columns and debt_col:
        chart_data = edited_df.groupby("makhoi")[debt_col].sum().reset_index()
        fig = px.bar(
            chart_data,
            x="makhoi",
            y=debt_col,
            title="Tổng số tiền nợ theo Khối đơn vị",
            labels={"makhoi": "Khối đơn vị", debt_col: "Số tiền nợ (VNĐ)"},
            color="makhoi",
        )
        st.plotly_chart(fig, use_container_width=True)

        if "sld" in edited_df.columns and pst_col:
          fig2 = px.scatter(
              edited_df,
              x="sld",
              y=debt_col,
              size=pst_col,
              color="makhoi",
              hover_name="tendvi" if "tendvi" in edited_df.columns else None,
              title="Tương quan giữa Số lao động và Số tiền nợ",
              labels={"sld": "Số lao động", debt_col: "Số tiền nợ"},
          )
          st.plotly_chart(fig2, use_container_width=True)

  except Exception as e:
    st.error(f"Đã xảy ra lỗi khi xử lý file: {e}")
else:
  st.info(
      "👈 Vui lòng tải file Excel nguồn lên ở thanh bên trái (Sidebar) để khởi"
      " chạy hệ thống phân tích."
  )
