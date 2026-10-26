import re
import io
import pandas as pd
import streamlit as st

st.set_page_config(
    page_title="Đối soát BHYT và Dân cư UBND", page_icon="📊", layout="wide"
)

st.title(
    "📊 Ứng dụng Đối soát Dữ liệu BHYT với Cơ sở Dữ liệu Dân cư UBND Phường An"
    " Dương"
)
st.markdown(
    "Tải lên tệp BHYT và tệp Dân cư UBND để hệ thống tự động lọc, so sánh theo"
    " **CCCD/Định danh** và **Họ tên / Năm sinh**."
)

# Khu vực tải file lên giao diện web
st.sidebar.header("📁 Tải tệp dữ liệu lên")
uploaded_bhyt = st.sidebar.file_uploader(
    "Chọn tệp BHYT (Excel)", type=["xlsx", "xls"]
)
uploaded_ubnd = st.sidebar.file_uploader(
    "Chọn tệp Dân cư UBND (Excel)", type=["xlsx", "xls"]
)


def clean_text(text):
  if pd.isna(text):
    return ""
  return re.sub(r"\s+", " ", str(text)).strip().upper()


def extract_year(val):
  if pd.isna(val):
    return None
  match = re.search(r"\b(19\d{2}|20\d{2})\b", str(val))
  return int(match.group(1)) if match else None


if uploaded_bhyt and uploaded_ubnd:
  with st.spinner("Đang đọc và xử lý dữ liệu, vui lòng chờ..."):
    # Đọc dữ liệu
    df_bhyt = pd.read_excel(uploaded_bhyt, sheet_name=0)
    df_ubnd = pd.read_excel(uploaded_ubnd, sheet_name=0)

    # Chuẩn hóa cột
    df_bhyt["CCCD_CLN"] = (
        df_bhyt["SO_CMND"]
        .astype(str)
        .str.replace(r"\.0$", "", regex=True)
        .str.strip()
        .str.upper()
    )
    df_ubnd["CCCD_CLN"] = (
        df_ubnd["Đinh danh"]
        .astype(str)
        .str.replace(r"\.0$", "", regex=True)
        .str.strip()
        .str.upper()
    )

    df_bhyt["HO_TEN_CLN"] = df_bhyt["HO_TEN"].apply(clean_text)
    df_ubnd["HO_TEN_CLN"] = df_ubnd["Họ và tên"].apply(clean_text)

    df_bhyt["NAM_SINH"] = df_bhyt["NGAY_SINH"].apply(extract_year)
    df_ubnd["NAM_SINH"] = df_ubnd["NTNS"].apply(extract_year)

    # Thống kê tổng quan
    col1, col2 = st.columns(2)
    col1.metric("Tổng số bản ghi BHYT", len(df_bhyt))
    col2.metric("Tổng số nhân khẩu UBND", len(df_ubnd))

    # 1. Khớp chính xác qua CCCD
    matched_by_cccd = pd.merge(
        df_bhyt,
        df_ubnd,
        on="CCCD_CLN",
        how="inner",
        suffixes=("_bhyt", "_ubnd"),
    )

    st.subheader(
        f"✅ 1. Danh sách khớp chính xác theo CCCD ({len(matched_by_cccd)} người)"
    )
    st.dataframe(matched_by_cccd.head(100))

    # 2. Xử lý phần chưa khớp CCCD -> Khớp theo Họ Tên & Năm sinh
    unmatched_bhyt = df_bhyt[
        ~df_bhyt["CCCD_CLN"].isin(matched_by_cccd["CCCD_CLN"])
    ]

    matched_by_name_list = []
    for idx, row in unmatched_bhyt.iterrows():
      name = row["HO_TEN_CLN"]
      yob = row["NAM_SINH"]

      candidates = df_ubnd[df_ubnd["HO_TEN_CLN"] == name]
      if not candidates.empty:
        if yob:
          candidates = candidates.copy()
          candidates["diff"] = abs(candidates["NAM_SINH"] - yob)
          candidates = candidates.sort_values(by="diff")
        best_match = candidates.iloc[0].copy()
        combined = pd.concat([row, best_match], axis=0)
        matched_by_name_list.append(combined)

    df_matched_by_name = (
        pd.DataFrame(matched_by_name_list)
        if matched_by_name_list
        else pd.DataFrame()
    )

    st.subheader(
        "🔍 2. Danh sách trùng Họ tên & khớp/gần đúng Năm sinh (Chưa khớp"
        f" CCCD): {len(df_matched_by_name)} người"
    )
    if not df_matched_by_name.empty:
      st.dataframe(df_matched_by_name.head(100))
    else:
      st.info("Không tìm thấy trường hợp trùng tên nào khác.")

    # Nút tải file kết quả tổng hợp
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
      matched_by_cccd.to_excel(writer, sheet_name="Khop_CCCD", index=False)
      if not df_matched_by_name.empty:
        df_matched_by_name.to_excel(
            writer, sheet_name="Khop_Theo_Ten_Va_Nam_Sinh", index=False
        )
    processed_data = output.getvalue()

    st.sidebar.markdown("---")
    st.sidebar.success("Xử lý thành công!")
    st.sidebar.download_button(
        label="📥 Tải xuống File Kết Quả (.xlsx)",
        data=processed_data,
        file_name="Ket_qua_doi_soat_BHYT_UBND.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
else:
  st.info(
      "👈 Vui lòng tải lên cả 2 tệp Excel (BHYT và Dân cư UBND) ở thanh bên trái"
      " để bắt đầu."
  )
