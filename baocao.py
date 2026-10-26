import io
import re
import pandas as pd
import streamlit as st

st.set_page_config(
    page_title="Đối soát BHYT và Dân cư UBND", page_icon="📊", layout="wide"
)

st.title(
    "📊 Ứng dụng Tự động Đối soát Dữ liệu BHYT & Cơ sở Dữ liệu Dân cư UBND Phường"
    " An Dương"
)
st.markdown(
    "Hệ thống tự động đọc dữ liệu từ Cloud, chuẩn hóa định dạng **CCCD/Định"
    " danh**, đối soát chính xác và tìm kiếm theo **Họ tên & Năm sinh gần"
    " đúng**."
)

# Tên 2 file dữ liệu lưu trên GitHub
BHYT_PATH = "BHYT_03108  bình đồng bộ cccd.xlsx"
UBND_PATH = "21.8.2026 nhân khẩu thường trú gửi UBND p AD.xlsx"


@st.cache_data
def load_and_process_data():
  # Đọc 2 file Excel
  df_bhyt = pd.read_excel(BHYT_PATH, sheet_name=0)
  df_ubnd = pd.read_excel(UBND_PATH, sheet_name="Sheet1")

  # Hàm chuẩn hóa số CCCD / Định danh (đảm bảo đủ 12 chữ số)
  def clean_cccd(val):
    if pd.isna(val):
      return ""
    s = str(val).split(".")[0].strip().upper()
    if s.isdigit() and len(s) < 12:
      s = s.zfill(12)
    return s

  def clean_text(text):
    if pd.isna(text):
      return ""
    return re.sub(r"\s+", " ", str(text)).strip().upper()

  def extract_year(val):
    if pd.isna(val):
      return None
    match = re.search(r"\b(19\d{2}|20\d{2})\b", str(val))
    return int(match.group(1)) if match else None

  # Tạo các cột phụ trợ chuẩn hóa
  df_bhyt["CCCD_CLN"] = df_bhyt["SO_CMND"].apply(clean_cccd)
  df_ubnd["CCCD_CLN"] = df_ubnd["Đinh danh"].apply(clean_cccd)

  df_bhyt["HO_TEN_CLN"] = df_bhyt["HO_TEN"].apply(clean_text)
  df_ubnd["HO_TEN_CLN"] = df_ubnd["Họ và tên"].apply(clean_text)

  df_bhyt["NAM_SINH"] = df_bhyt["NGAY_SINH"].apply(extract_year)
  df_ubnd["NAM_SINH"] = df_ubnd["NTNS"].apply(extract_year)

  return df_bhyt, df_ubnd


try:
  with st.spinner(
      "Đang tải và đồng bộ dữ liệu từ Cloud (có thể mất vài giây)..."
  ):
    df_bhyt, df_ubnd = load_and_process_data()

  col1, col2 = st.columns(2)
  col1.metric("Tổng số bản ghi BHYT", f"{len(df_bhyt):,}")
  col2.metric("Tổng số nhân khẩu UBND", f"{len(df_ubnd):,}")

  # 1. Khớp chính xác theo CCCD
  matched_by_cccd = pd.merge(
      df_bhyt,
      df_ubnd,
      on="CCCD_CLN",
      how="inner",
      suffixes=("_bhyt", "_ubnd"),
  )

  st.subheader(
      f"✅ 1. Danh sách khớp chính xác theo Số CCCD / Định danh ("
      f"{len(matched_by_cccd)} người)"
  )
  if not matched_by_cccd.empty:
    st.dataframe(matched_by_cccd.head(100), use_container_width=True)
  else:
    st.info("Không có bản ghi nào trùng khớp tuyệt đối theo CCCD.")

  # 2. Xử lý các trường hợp chưa khớp CCCD -> Khớp theo Họ Tên và Năm sinh gần đúng
  unmatched_bhyt = df_bhyt[~df_bhyt["CCCD_CLN"].isin(matched_by_cccd["CCCD_CLN"])]

  matched_by_name_list = []
  for idx, row in unmatched_bhyt.iterrows():
    name = row["HO_TEN_CLN"]
    yob = row["NAM_SINH"]

    candidates = df_ubnd[df_ubnd["HO_TEN_CLN"] == name]
    if not candidates.empty:
      if yob is not None:
        candidates = candidates.copy()
        candidates["diff"] = abs(candidates["NAM_SINH"] - yob)
        candidates = candidates.sort_values(by="diff")
      best_match = candidates.iloc[0]

      # Gộp dữ liệu thành dictionary với hậu tố rõ ràng để tránh trùng lặp tên cột
      combined_dict = {}
      for col in df_bhyt.columns:
        combined_dict[f"{col}_bhyt"] = row[col]
      for col in df_ubnd.columns:
        combined_dict[f"{col}_ubnd"] = best_match[col]
      if "diff" in candidates.columns:
        combined_dict["chenh_lech_nam_ sinh"] = best_match["diff"]

      matched_by_name_list.append(combined_dict)

  df_matched_by_name = (
      pd.DataFrame(matched_by_name_list)
      if matched_by_name_list
      else pd.DataFrame()
  )

  st.subheader(
      "🔍 2. Danh sách trùng Họ tên & khớp/gần đúng Năm sinh (Chưa khớp CCCD):"
      f" {len(df_matched_by_name)} người"
  )
  if not df_matched_by_name.empty:
    st.dataframe(df_matched_by_name.head(100), use_container_width=True)
  else:
    st.info("Không tìm thấy trường hợp trùng tên nào khác.")

  # Tạo file Excel kết quả để tải xuống
  output = io.BytesIO()
  with pd.ExcelWriter(output, engine="openpyxl") as writer:
    matched_by_cccd.to_excel(writer, sheet_name="Khop_CCCD", index=False)
    if not df_matched_by_name.empty:
      df_matched_by_name.to_excel(
          writer, sheet_name="Khop_Theo_Ten_Va_Nam_Sinh", index=False
      )
  processed_data = output.getvalue()

  st.sidebar.markdown("---")
  st.sidebar.success("Trạng thái: Sẵn sàng xuất báo cáo!")
  st.sidebar.download_button(
      label="📥 Tải xuống File Kết Quả Tổng Hợp (.xlsx)",
      data=processed_data,
      file_name="Ket_qua_doi_soat_BHYT_UBND.xlsx",
      mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  )

except Exception as e:
  st.error(f"Đã xảy ra lỗi trong quá trình xử lý: {e}")
  st.info(
      "Vui lòng kiểm tra lại xem 2 file Excel đã được tải lên đầy đủ vào thư"
      " mục gốc của GitHub hay chưa."
  )
