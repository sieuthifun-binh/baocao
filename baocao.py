import io
import re
import pandas as pd
import streamlit as st

st.set_page_config(
    page_title="Hệ thống Đối soát Dữ liệu Thông minh", page_icon="📊", layout="wide"
)

st.title("📊 Ứng dụng Đối soát Dữ liệu Nâng cao (BHYT & Dân cư UBND)")
st.markdown(
    "Tải lên tệp cần so sánh (**Ô A**) và tệp dân cư gốc (**Ô B**), sau đó tùy"
    " chỉnh các tiêu chí khớp dữ liệu và khoảng chênh lệch năm sinh trực tiếp"
    " trên giao diện."
)

# --- THANH BÊN (SIDEBAR): UPLOAD FILE & CÀI ĐẶT ---
st.sidebar.header("📁 1. Tải lên tệp dữ liệu")
uploaded_file_a = st.sidebar.file_uploader(
    "Ô A: Tệp cần so sánh (Ví dụ: BHYT)", type=["xlsx", "xls"]
)
uploaded_file_b = st.sidebar.file_uploader(
    "Ô B: Tệp dân cư gốc (Ví dụ: Dân cư UBND)", type=["xlsx", "xls"]
)

st.sidebar.markdown("---")
st.sidebar.header("⚙️ 2. Tiêu chí tìm kiếm & Lọc")
tolerance_years = st.sidebar.slider(
    "Khoảng chênh lệch năm sinh cho phép (Năm)",
    min_value=0,
    max_value=5,
    value=2,
    step=1,
)


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


if uploaded_file_a and uploaded_file_b:
  try:
    # Đọc dữ liệu từ file người dùng upload
    df_a = pd.read_excel(uploaded_file_a, sheet_name=0)
    df_b = pd.read_excel(uploaded_file_b, sheet_name=0)

    st.success("Tải dữ liệu thành công!")

    # --- CẤU HÌNH CỘT ---
    st.sidebar.markdown("---")
    st.sidebar.header("Mapping Cột Dữ Liệu")

    cols_a = df_a.columns.tolist()
    cols_b = df_b.columns.tolist()

    # Tự động gợi ý tên cột nếu có sẵn
    default_cccd_a = (
        cols_a.index("SO_CMND") if "SO_CMND" in cols_a else 0
    )
    default_name_a = cols_a.index("HO_TEN") if "HO_TEN" in cols_a else 0
    default_dob_a = cols_a.index("NGAY_SINH") if "NGAY_SINH" in cols_a else 0

    default_cccd_b = cols_b.index("Đinh danh") if "Đinh danh" in cols_b else 0
    default_name_b = cols_b.index("Họ và tên") if "Họ và tên" in cols_b else 0
    default_dob_b = cols_b.index("NTNS") if "NTNS" in cols_b else 0

    col_opt1, col_opt2 = st.columns(2)
    with col_opt1:
      st.subheader("Cấu hình Tệp A")
      col_cccd_a = st.selectbox(
          "Cột CCCD / Định danh (Tệp A)", cols_a, index=default_cccd_a
      )
      col_name_a = st.selectbox("Cột Họ và tên (Tệp A)", cols_a, index=default_name_a)
      col_dob_a = st.selectbox(
          "Cột Ngày sinh / Năm sinh (Tệp A)", cols_a, index=default_dob_a
      )

    with col_opt2:
      st.subheader("Cấu hình Tệp B")
      col_cccd_b = st.selectbox(
          "Cột CCCD / Định danh (Tệp B)", cols_b, index=default_cccd_b
      )
      col_name_b = st.selectbox("Cột Họ và tên (Tệp B)", cols_b, index=default_name_b)
      col_dob_b = st.selectbox(
          "Cột Ngày sinh / Năm sinh (Tệp B)", cols_b, index=default_dob_b
      )

    # --- XỬ LÝ DỮ LIỆU ---
    df_a["CCCD_CLN"] = df_a[col_cccd_a].apply(clean_cccd)
    df_b["CCCD_CLN"] = df_b[col_cccd_b].apply(clean_cccd)

    df_a["HO_TEN_CLN"] = df_a[col_name_a].apply(clean_text)
    df_b["HO_TEN_CLN"] = df_b[col_name_b].apply(clean_text)

    df_a["NAM_SINH"] = df_a[col_dob_a].apply(extract_year)
    df_b["NAM_SINH"] = df_b[col_dob_b].apply(extract_year)

    st.markdown("---")
    m1, m2 = st.columns(2)
    m1.metric("Tổng bản ghi Tệp A", f"{len(df_a):,}")
    m2.metric("Tổng bản ghi Tệp B (Dân cư)", f"{len(df_b):,}")

    # 1. Khớp chính xác qua CCCD
    matched_by_cccd = pd.merge(
        df_a, df_b, on="CCCD_CLN", how="inner", suffixes=("_a", "_b")
    )

    st.subheader(
        f"✅ 1. Khớp chính xác theo Số CCCD / Định danh ({len(matched_by_cccd)}"
        " người)"
    )
    if not matched_by_cccd.empty:
      st.dataframe(matched_by_cccd.head(100), use_container_width=True)
    else:
      st.info("Không tìm thấy bản ghi khớp tuyệt đối theo CCCD.")

    # 2. Xử lý các trường hợp chưa khớp CCCD -> Khớp Họ tên & Năm sinh trong khoảng chênh lệch
    unmatched_a = df_a[~df_a["CCCD_CLN"].isin(matched_by_cccd["CCCD_CLN"])]

    matched_by_name_list = []
    for idx, row in unmatched_a.iterrows():
      name = row["HO_TEN_CLN"]
      yob = row["NAM_SINH"]

      candidates = df_b[df_b["HO_TEN_CLN"] == name]
      if not candidates.empty:
        if yob is not None:
          candidates = candidates.copy()
          candidates["diff"] = abs(candidates["NAM_SINH"] - yob)
          # Lọc theo dung sai năm sinh được chọn (tolerance_years)
          candidates = candidates[candidates["diff"] <= tolerance_years]
          candidates = candidates.sort_values(by="diff")

        if not candidates.empty:
          best_match = candidates.iloc[0]
          combined_dict = {}
          for col in df_a.columns:
            combined_dict[f"{col}_a"] = row[col]
          for col in df_b.columns:
            combined_dict[f"{col}_b"] = best_match[col]
          combined_dict["chenh_lech_nam_sinh"] = best_match["diff"]
          matched_by_name_list.append(combined_dict)

    df_matched_by_name = (
        pd.DataFrame(matched_by_name_list)
        if matched_by_name_list
        else pd.DataFrame()
    )

    st.subheader(
        f"🔍 2. Trùng Họ tên & Lệch năm sinh trong khoảng [±{tolerance_years}"
        f" tuổi]: {len(df_matched_by_name)} người"
    )
    if not df_matched_by_name.empty:
      st.dataframe(df_matched_by_name.head(100), use_container_width=True)
    else:
      st.info(
          "Không tìm thấy trường hợp trùng tên nào thỏa mãn khoảng chênh lệch"
          " năm sinh này."
      )

    # --- TẠO FILE EXCEL XUẤT BÁO CÁO ---
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
      matched_by_cccd.to_excel(writer, sheet_name="Khop_Chinh_Xac_CCCD", index=False)
      if not df_matched_by_name.empty:
        df_matched_by_name.to_excel(
            writer, sheet_name="Khop_Theo_Ten_Va_Nam_Sinh", index=False
        )
    processed_data = output.getvalue()

    st.sidebar.markdown("---")
    st.sidebar.success("Trạng thái: Sẵn sàng xuất báo cáo!")
    st.sidebar.download_button(
        label="📥 Tải xuống File Kết Quả (.xlsx)",
        data=processed_data,
        file_name="Ket_qua_doi_soat_thong_minh.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

  except Exception as e:
    st.error(f"Đã xảy ra lỗi khi xử lý tệp: {e}")
else:
  st.info(
      "👈 Vui lòng tải lên cả **Tệp A** và **Tệp B** ở thanh công cụ bên trái để"
      " bắt đầu."
  )
