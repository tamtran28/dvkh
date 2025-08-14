import streamlit as st
import pandas as pd
import numpy as np
import zipfile
import re
from io import BytesIO
from datetime import datetime

st.set_page_config(page_title="DVKH | CKH/KKH + Ủy quyền", layout="wide")
st.title("📦 Xử lý CKH/KKH từ ZIP + Ủy quyền (Mục 30) + SMS/SCM010")

st.markdown("""
**Hướng dẫn nhanh**
1) Nén tất cả file **CKH/KKH** vào 1 file: `ckh_kkh.zip` (mẫu tên: `HDV_CHITIET_CKH_*.xls*`, `HDV_CHITIET_KKH_*.xls*`).  
2) Nén **các file tham chiếu** (Mục 30, SMS, SCM010) vào 1 file: `others.zip`.  
   - Ví dụ chứa:  
     - `MUC 30 *.xlsx` (bắt buộc)  
     - `Muc14_DK_SMS.txt` (tab-separated) (khuyến nghị)  
     - `Muc14_SCM010.xlsx` (khuyến nghị)
3) Tải 2 file ZIP lên, bấm **Xử lý** để nhận 1 file Excel tổng hợp tải về.
""")

# =========================
# Helpers
# =========================
def read_excel_safely(file_like, dtype=None):
    """
    Đọc Excel .xls/.xlsx an toàn. Cố gắng dùng engine phù hợp.
    Yêu cầu: openpyxl cho .xlsx, xlrd==1.2.0 cho .xls.
    """
    try:
        # Thử mặc định
        return pd.read_excel(file_like, dtype=dtype)
    except Exception:
        # Thử ép engine theo phần mở rộng
        try:
            return pd.read_excel(file_like, engine="openpyxl", dtype=dtype)
        except Exception:
            return pd.read_excel(file_like, engine="xlrd", dtype=dtype)

def extract_first_excel_or_txt_from_zip(zip_bytes, wanted_substrings, accept_txt=False):
    """
    Tìm file đầu tiên trong zip có tên chứa bất kỳ chuỗi con trong wanted_substrings.
    Trả về (df, filename). Với txt (tab-separated) nếu accept_txt=True.
    """
    with zipfile.ZipFile(zip_bytes, "r") as z:
        for name in z.namelist():
            low = name.lower()
            if any(s in low for s in wanted_substrings):
                with z.open(name) as f:
                    if accept_txt and (low.endswith(".txt") or low.endswith(".tsv")):
                        # cố gắng đọc TSV (tab)
                        try:
                            df = pd.read_csv(f, sep="\t", on_bad_lines="skip", dtype=str)
                            return df, name
                        except Exception:
                            # fallback: thử csv
                            f.seek(0)
                            df = pd.read_csv(f, dtype=str)
                            return df, name
                    else:
                        df = read_excel_safely(f, dtype=str)
                        return df, name
    return None, None

def read_all_ckh_kkh_from_zip(zip_bytes):
    """
    Đọc tất cả file CKH/KKH từ ZIP.
    - CKH: tên chứa 'HDV_CHITIET_CKH_'
    - KKH: tên chứa 'HDV_CHITIET_KKH_'
    Trả về: df_b_CKH, df_b_KKH, df_b_all
    """
    df_b_CKH_list, df_b_KKH_list = [], []
    with zipfile.ZipFile(zip_bytes, "r") as z:
        names = z.namelist()
        for name in names:
            low = name.lower()
            if low.endswith((".xls", ".xlsx", ".xlsm", ".xlsb")):
                with z.open(name) as f:
                    try:
                        df_tmp = read_excel_safely(f, dtype=str)
                        if "hdv_chitiet_ckh_" in low:
                            df_b_CKH_list.append(df_tmp)
                        elif "hdv_chitiet_kkh_" in low:
                            df_b_KKH_list.append(df_tmp)
                    except Exception as e:
                        st.warning(f"⚠️ Không đọc được file: {name}. Lỗi: {e}")

    df_b_CKH = pd.concat(df_b_CKH_list, ignore_index=True) if df_b_CKH_list else pd.DataFrame()
    df_b_KKH = pd.concat(df_b_KKH_list, ignore_index=True) if df_b_KKH_list else pd.DataFrame()
    df_b = pd.concat([df_b_CKH, df_b_KKH], ignore_index=True) if not df_b_CKH.empty or not df_b_KKH.empty else pd.DataFrame()

    return df_b_CKH, df_b_KKH, df_b

def extract_name_upper(value):
    """
    Tách tên (viết hoa) từ chuỗi có thể có '-' hoặc ','.
    Giữ phần có pattern A-Z và space, >= 3 ký tự.
    """
    parts = re.split(r'[-,]', str(value))
    for part in parts:
        name = part.strip()
        if re.fullmatch(r'[A-Z ]{3,}', name):
            return name
    return str(value).strip()

# =========================
# Uploaders
# =========================
zip_ckh_kkh = st.file_uploader("📁 Tải ZIP chứa CKH/KKH (tên file chứa 'HDV_CHITIET_CKH_' hoặc 'HDV_CHITIET_KKH_')", type="zip")
zip_others = st.file_uploader("📁 Tải ZIP chứa các file tham chiếu (Mục 30, SMS, SCM010)", type="zip")

run = st.button("▶️ Xử lý")

if run:
    if zip_ckh_kkh is None:
        st.error("Vui lòng tải ZIP chứa CKH/KKH trước.")
        st.stop()

    # 1) Đọc CKH/KKH
    df_b_CKH, df_b_KKH, df_b = read_all_ckh_kkh_from_zip(zip_ckh_kkh)
    if df_b.empty:
        st.error("Không tìm thấy file CKH/KKH hợp lệ trong ZIP.")
        st.stop()

    st.success(f"Đã đọc CKH: {len(df_b_CKH)} dòng, KKH: {len(df_b_KKH)} dòng, Tổng: {len(df_b)} dòng.")

    # Chuẩn hóa một số cột có thể dùng ở dưới
    for col in ["IDXACNO", "CUSTSEQ"]:
        if col in df_b.columns:
            df_b[col] = df_b[col].astype(str)

    # 2) Đọc các file tham chiếu từ others.zip
    df_a = pd.DataFrame()        # Mục 30
    df_sms = pd.DataFrame()      # DK_SMS (txt/tsv)
    df_scm10 = pd.DataFrame()    # SCM010 (xls/xlsx)

    if zip_others is not None:
        # Mục 30 - file tên chứa 'muc 30'
        df_a, name_a = extract_first_excel_or_txt_from_zip(zip_others, wanted_substrings=["muc 30"], accept_txt=False)
        if df_a is None:
            st.warning("Không tìm thấy file 'MUC 30 *.xlsx' trong others.zip. Một số logic ủy quyền sẽ bị bỏ qua.")

        # DK_SMS - file txt hoặc xlsx chứa 'muc14_dk_sms'
        df_sms, name_sms = extract_first_excel_or_txt_from_zip(zip_others, wanted_substrings=["muc14_dk_sms"], accept_txt=True)
        if df_sms is None:
            st.info("Không tìm thấy file DK_SMS (txt/xlsx). Bỏ qua gắn cờ 'TK có đăng ký SMS'.")

        # SCM010 - xlsx chứa 'muc14_scm010'
        df_scm10, name_scm = extract_first_excel_or_txt_from_zip(zip_others, wanted_substrings=["muc14_scm010"], accept_txt=False)
        if df_scm10 is None:
            st.info("Không tìm thấy file SCM010 (xlsx). Bỏ qua gắn cờ 'CIF có đăng ký SCM010'.")
    else:
        st.warning("Bạn chưa tải others.zip, sẽ chỉ xử lý phần CKH/KKH cơ bản.")

    # =========================
    # LOGIC ỦY QUYỀN (Mục 30) + ghép CKH/KKH
    # =========================
    merged = pd.DataFrame()
    df_uyquyen = pd.DataFrame()
    df_tc3 = pd.DataFrame()

    if not df_a.empty:
        # Giữ dạng chuỗi
        df_a = df_a.copy()
        for c in df_a.columns:
            df_a[c] = df_a[c].astype(str)

        req_cols = ["DESCRIPTION", "NGUOI_UY_QUYEN", "NGUOI_DUOC_UY_QUYEN", "TK_DUOC_UY_QUYEN",
                    "PRIMARY_SOL_ID", "EFFECTIVEDATE", "EXPIRYDATE"]
        # Thêm cột nếu thiếu
        for c in req_cols:
            if c not in df_a.columns:
                df_a[c] = ""

        # Lọc 'chữ ký'
        mask_sig = df_a["DESCRIPTION"].str.contains(r"chu\s*ky|chuky|cky", case=False, na=False)
        df_a = df_a[mask_sig].copy()

        # Chuẩn ngày
        def to_mmddyyyy(s):
            # cố gắng parse YYYYMMDD trước, sau đó ISO, nếu fail -> NaT
            s = str(s)
            out = pd.NaT
            for fmt in ("%Y%m%d", "%Y-%m-%d", "%d/%m/%Y", "%m/%d/%Y"):
                try:
                    out = pd.to_datetime(s, format=fmt)
                    break
                except Exception:
                    continue
            if pd.isna(out):
                try:
                    out = pd.to_datetime(s, errors="coerce")
                except Exception:
                    out = pd.NaT
            return out

        df_a["EFFECTIVEDATE"] = df_a["EFFECTIVEDATE"].apply(to_mmddyyyy)
        df_a["EXPIRYDATE"]    = df_a["EXPIRYDATE"].apply(to_mmddyyyy)
        df_a["EFFECTIVEDATE"] = df_a["EFFECTIVEDATE"].dt.strftime("%m/%d/%Y")
        df_a["EXPIRYDATE"]    = df_a["EXPIRYDATE"].dt.strftime("%m/%d/%Y")

        # Loại doanh nghiệp
        dn_keywords = ["CONG TY", "CTY", "CONGTY", "CÔNG TY", "CÔNGTY"]
        df_a = df_a[~df_a["NGUOI_UY_QUYEN"].str.upper().str.contains("|".join(dn_keywords), na=False)].copy()

        # Chuẩn tên người được ủy quyền
        df_a["NGUOI_DUOC_UY_QUYEN"] = df_a["NGUOI_DUOC_UY_QUYEN"].apply(extract_name_upper)

        # Ghép CIF từ CKH/KKH
        # Đồng nhất kiểu
        df_a["TK_DUOC_UY_QUYEN"] = df_a["TK_DUOC_UY_QUYEN"].astype(str)
        if "IDXACNO" in df_b.columns and "CUSTSEQ" in df_b.columns:
            df_b["IDXACNO"] = df_b["IDXACNO"].astype(str)
            df_b["CUSTSEQ"] = df_b["CUSTSEQ"].astype(str)

            merged = df_a.merge(
                df_b[["IDXACNO", "CUSTSEQ"]],
                left_on="TK_DUOC_UY_QUYEN",
                right_on="IDXACNO",
                how="left"
            )
            # Tạo CIF_NGUOI_UY_QUYEN
            def make_cif(x):
                return str(int(float(x))) if pd.notna(x) and str(x).strip() != "" else "NA"
            merged["CIF_NGUOI_UY_QUYEN"] = merged["CUSTSEQ"].apply(make_cif)

            # Điền CIF cho các bản ghi cùng NGUOI_DUOC_UY_QUYEN
            cif_updated = merged["CIF_NGUOI_UY_QUYEN"].copy()
            for nguoi_uq, group in merged.groupby("NGUOI_DUOC_UY_QUYEN"):
                vals = group["CIF_NGUOI_UY_QUYEN"].unique().tolist()
                has_na = "NA" in vals
                actual = [v for v in vals if v != "NA"]
                if has_na and actual:
                    fill_val = actual[0]
                    cif_updated.loc[group[merged["CIF_NGUOI_UY_QUYEN"] == "NA"].index] = fill_val

            merged["CIF_NGUOI_UY_QUYEN"] = cif_updated
            merged.drop(columns=[c for c in ["IDXACNO", "CUSTSEQ"] if c in merged.columns], inplace=True)

            # Phân loại TK thuộc CKH/KKH
            set_ckh = set(df_b_CKH.get("CUSTSEQ", pd.Series(dtype=str)).astype(str)) if not df_b_CKH.empty else set()
            set_kkh = set(df_b_KKH.get("IDXACNO", pd.Series(dtype=str)).astype(str)) if not df_b_KKH.empty else set()

            def phan_loai_tk(tk):
                tks = str(tk)
                if tks in set_ckh:
                    return "CKH"
                if tks in set_kkh:
                    return "KKH"
                return "NA"

            merged["LOAI_TK"] = merged["TK_DUOC_UY_QUYEN"].astype(str).apply(phan_loai_tk)

            # Cờ thời hạn ủy quyền
            # chuyển lại sang datetime để tính
            m = merged.copy()
            m["EFFECTIVEDATE_dt"] = pd.to_datetime(m["EFFECTIVEDATE"], errors="coerce")
            m["EXPIRYDATE_dt"]    = pd.to_datetime(m["EXPIRYDATE"], errors="coerce")
            year_diff = (m["EXPIRYDATE_dt"].dt.year - m["EFFECTIVEDATE_dt"].dt.year).fillna(0)
            merged["KHONG_NHAP_TGIAN_UQ"] = np.where(year_diff == 99, "X", "")
            merged["UQ_TREN_50_NAM"]      = np.where(year_diff >= 50, "X", "")

            # Chuẩn bị df_uyquyen để gắn thêm cờ SMS/SCM010
            df_uyquyen = merged.copy()
        else:
            st.warning("Không thấy cột 'IDXACNO' và 'CUSTSEQ' trong CKH/KKH để ghép ủy quyền. Bỏ qua phần ủy quyền.")

    # =========================
    # SMS & SCM010 flags
    # =========================
    if not df_uyquyen.empty:
        # SMS
        if not df_sms.empty:
            # Chuẩn cột
            for c in ["FORACID", "ORGKEY", "C_MOBILE_NO", "CUSTTPCD"]:
                if c in df_sms.columns:
                    df_sms[c] = df_sms[c].astype(str)

            # Loại bỏ foracid có chữ cái
            if "FORACID" in df_sms.columns:
                df_sms = df_sms[df_sms["FORACID"].str.match(r"^\d+$", na=False)]

            # Chỉ KH cá nhân
            if "CUSTTPCD" in df_sms.columns:
                df_sms = df_sms[df_sms["CUSTTPCD"].str.upper() != "KHDN"]

            tk_sms_set = set(df_sms.get("FORACID", pd.Series(dtype=str)))
            df_uyquyen["TK có đăng ký SMS"] = df_uyquyen["TK_DUOC_UY_QUYEN"].astype(str).apply(
                lambda x: "X" if x in tk_sms_set else ""
            )
        else:
            df_uyquyen["TK có đăng ký SMS"] = ""

        # SCM010
        if not df_scm10.empty:
            df_scm10 = df_scm10.rename(columns=lambda x: str(x).strip())
            if "CIF_ID" in df_scm10.columns:
                df_scm10["CIF_ID"] = df_scm10["CIF_ID"].astype(str)
                cif_scm10_set = set(df_scm10["CIF_ID"])
                df_uyquyen["CIF có đăng ký SCM010"] = df_uyquyen["CIF_NGUOI_UY_QUYEN"].astype(str).apply(
                    lambda x: "X" if x in cif_scm10_set else ""
                )
            else:
                df_uyquyen["CIF có đăng ký SCM010"] = ""
        else:
            df_uyquyen["CIF có đăng ký SCM010"] = ""

        # Tiêu chí 3: 1 người nhận UQ của nhiều người
        df_tc3 = df_uyquyen.copy()
        if "NGUOI_DUOC_UY_QUYEN" in df_tc3.columns and "NGUOI_UY_QUYEN" in df_tc3.columns:
            grouped = df_tc3.groupby("NGUOI_DUOC_UY_QUYEN")["NGUOI_UY_QUYEN"].nunique().reset_index()
            nguoi_nhan_nhieu = set(grouped[grouped["NGUOI_UY_QUYEN"] >= 2]["NGUOI_DUOC_UY_QUYEN"])
            df_tc3["1 người nhận UQ của nhiều người"] = df_tc3["NGUOI_DUOC_UY_QUYEN"].apply(
                lambda x: "X" if x in nguoi_nhan_nhieu else ""
            )

    # =========================
    # Xuất Excel
    # =========================
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        if not df_b_CKH.empty:
            df_b_CKH.to_excel(writer, sheet_name="CKH_raw", index=False)
        if not df_b_KKH.empty:
            df_b_KKH.to_excel(writer, sheet_name="KKH_raw", index=False)
        if not merged.empty:
            merged.to_excel(writer, sheet_name="tieu chi 1 (UYQ)", index=False)
        if not df_uyquyen.empty:
            df_uyquyen.to_excel(writer, sheet_name="tieu chi 2 (SMS/SCM010)", index=False)
        if not df_tc3.empty:
            df_tc3.to_excel(writer, sheet_name="tieu chi 3 (UQ nhiều)", index=False)

    st.success("✅ Hoàn tất. Bạn có thể tải file kết quả bên dưới.")
    st.download_button(
        label="⬇️ Tải Excel kết quả",
        data=output.getvalue(),
        file_name="DVKH_2241_KetQua.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

    # Hiển thị preview nhỏ
    if not df_uyquyen.empty:
        st.subheader("Preview — Tiêu chí 2 (SMS/SCM010)")
        st.dataframe(df_uyquyen.head(50), use_container_width=True)
