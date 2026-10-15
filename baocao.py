import google.genai as genai
import pandas as pd
import plotly.express as px
import streamlit as st

st.set_page_config(
    page_title="Hệ thống Quản lý & Phân tích Thu Nộp BHXH",
    page_icon="📊",
    layout="wide",
)

st.title("📊 Hệ thống Phân tích & Báo cáo Tổng hợp Thu Nộp BHXH/BHYT")
st.markdown(
    "Hệ thống tự động đồng bộ từ file nguồn **Cập nhật**, phân tách theo"
    " **Chi tiết chuyên quản**, phân tích **Đơn vị nợ** và tổng hợp **Kế hoạch"
    " thu** không giới hạn số dòng."
)

# Sidebar cấu hình
with st.sidebar:
  st.header("🔑 Cấu hình AI & Hệ thống")
  api_key = st.text_input("Nhập Google Gemini API Key:", type="password")
  st.markdown("---")
  st.markdown("### 📂 Tải file dữ liệu nguồn")
  uploaded_file = st.file_uploader(
      "Chọn file Excel nguồn (.xls, .xlsx)", type=["xls", "xlsx", "csv"]
  )

if uploaded_file is not None:
  try:
    # Đọc các sheet từ file Excel
    xls = pd.ExcelFile(uploaded_file)
    sheet_names = xls.sheet_names
    st.success(
        f"Đã tải thành công file! Các sheet có trong file: {sheet_names}"
    )

    # Đọc sheet Cập nhật làm gốc
    df_capnhat = pd.read_excel(
        uploaded_file, sheet_name="Cập nhật" if "Cập nhật" in sheet_names else 0
    )

    # Hiển thị tổng quan
    col1, col2, col3, col4 = st.columns(4)
    with col1:
      st.metric("Tổng số đơn vị", len(df_capnhat))
    with col2:
      total_ lao_dong = (
          df_capnhat["sld"].sum() if "sld" in df_capnhat.columns else 0
      )
      st.metric("Tổng số lao động", f"{total_ lao_dong:,.0f}")
    with col3:
      total_no = (
          df_capnhat["_tien_ck"].sum() if "_tien_ck" in df_capnhat.columns else 0
      )
      st.metric("Tổng tiền nợ cuối kỳ", f"{total_no:,.0f} đ")
    with col4:
      total_da_thu = (
          df_capnhat["_tien_unc"].sum()
          if "_tien_unc" in df_capnhat.columns
          else 0
      )
      st.metric("Đã thu trong kỳ", f"{total_da_thu:,.0f} đ")

    st.markdown("---")

    # Tabs chức năng theo đúng cấu trúc file Excel yêu cầu
    tab1, tab2, tab3, tab4, tab5 = st.tabs([
        "📋 Cập nhật (Nguồn)",
        "👤 Chi tiết (Theo Chuyên Quản)",
        "⚠️ Phân tích Nợ",
        "📈 KH Thu & Tổng hợp",
        "🤖 AI Phân tích & Báo cáo",
    ])

    with tab1:
      st.subheader("Dữ liệu gốc từ sheet Cập nhật")
      st.dataframe(df_capnhat, use_container_width=True)

    with tab2:
      st.subheader("Báo cáo Chi tiết theo Cán bộ Chuyên quản")
      if "nguoilh" in df_capnhat.columns:
        users = df_capnhat["nguoilh"].dropna().unique().tolist()
        selected_user = st.selectbox(
            "Chọn cán bộ chuyên quản / User:", users
        )
        df_user = df_capnhat[df_capnhat["nguoilh"] == selected_user]
        st.write(
            f"Hiển thị {len(df_user)} đơn vị do cán bộ **{selected_user}** phụ"
            " trách:"
        )
        st.dataframe(df_user, use_container_width=True)
      else:
        st.warning("Không tìm thấy cột quản lý (`nguoilh`) trong dữ liệu.")

    with tab3:
      st.subheader("Phân tích số liệu Đơn vị Nợ")
      if "_tien_ck" in df_capnhat.columns:
        df_no = df_capnhat[df_capnhat["_tien_ck"] > 0].sort_values(
            by="_tien_ck", ascending=False
        )
        st.write(f"Tìm thấy **{len(df_no)}** đơn vị còn nợ cuối kỳ.")
        st.dataframe(
            df_no[[
                "madvi",
                "tendvi",
                "sld",
                "tyleno",
                "tylenoyt",
                "_tien_ck",
                "nguoilh",
            ]],
            use_container_width=True,
        )

        # Biểu đồ trực quan top đơn vị nợ nhiều nhất
        if not df_no.empty:
          st.markdown("### 📊 Top Đơn vị nợ cao nhất")
          top_n = st.slider("Chọn số lượng đơn vị hiển thị:", 5, 20, 10)
          fig_no = px.bar(
              df_no.head(top_n),
              x="tendvi",
              y="_tien_ck",
              title=f"Top {top_n} đơn vị có số tiền nợ cuối kỳ lớn nhất",
              labels={"tendvi": "Tên đơn vị", "_tien_ck": "Số tiền nợ (VNĐ)"},
              color="_tien_ck",
              color_continuous_scale="Reds",
          )
          st.plotly_chart(fig_no, use_container_width=True)
      else:
        st.warning("Không tìm thấy cột số tiền nợ (`_tien_ck`).")

    with tab4:
      st.subheader("Tổng hợp Kế hoạch Thu (KH Thu)")
      if "makhoi" in df_capnhat.columns and "_tien_unc" in df_capnhat.columns:
        df_khoi = (
            df_capnhat.groupby("makhoi")
            .agg(
                So_Don_Vi=("madvi", "count"),
                Tong_Lao_Dong=("sld", "sum"),
                Da_Thu=("_tien_unc", "sum"),
                No_Cuoi_Ky=("_tien_ck", "sum"),
            )
            .reset_index()
        )
        st.dataframe(df_khoi, use_container_width=True)

        # Biểu đồ trực quan theo Khối
        fig_khoi = px.bar(
            df_khoi,
            x="makhoi",
            y=["Da_Thu", "No_Cuoi_Ky"],
            title="So sánh Số đã thu và Nợ cuối kỳ theo Khối đơn vị",
            barmode="group",
            labels={
                "value": "Số tiền (VNĐ)",
                "makhoi": "Khối",
                "variable": "Chỉ tiêu",
            },
        )
        st.plotly_chart(fig_khoi, use_container_width=True)
      else:
        st.info("Dữ liệu không đủ cột phân nhóm khối hoặc số thu.")

    with tab5:
      st.subheader("🤖 Trợ lý AI Phân tích Tự động & Lập Báo cáo")
      user_query = st.text_area(
          "Nhập yêu cầu phân tích của bạn (Ví dụ: Đánh giá tình hình nợ của"
          " các đơn vị khối TH, đề xuất giải pháp đôn đốc thu...):",
          placeholder="Nhập yêu cầu tại đây...",
      )

      if st.button("🚀 Gửi yêu cầu cho AI phân tích", type="primary"):
        if not api_key:
          st.error("Vui lòng nhập Google Gemini API Key ở Sidebar.")
        elif not user_query:
          st.warning("Vui lòng nhập câu hỏi/yêu cầu phân tích.")
        else:
          with st.spinner("AI đang tổng hợp số liệu và lập báo cáo..."):
            try:
              client = genai.Client(api_key=api_key)
              csv_summary = df_capnhat.head(100).to_csv(index=False)
              prompt = (
                  f"Dưới đây là dữ liệu tổng hợp từ file quản lý thu nộp BHXH"
                  f" (mẫu 100 dòng đầu):\n{csv_summary}\n\nTổng số đơn vị:"
                  f" {len(df_capnhat)}.\nDựa trên dữ liệu trên, hãy thực hiện"
                  f" yêu cầu sau bằng tiếng Việt thật chuyên nghiệp:\n{user_query}"
              )
              response = client.models.generate_content(
                  model="gemini-2.5-flash", contents=prompt
              )
              st.markdown("### 📋 Báo cáo kết quả từ AI")
              st.markdown(response.text)
            except Exception as e:
              st.error(f"Lỗi khi gọi AI: {e}")

  except Exception as e:
    st.error(f"Lỗi đọc file: {e}")
else:
  st.info("👈 Vui lòng tải file Excel nguồn lên ở thanh bên trái để bắt đầu.")
