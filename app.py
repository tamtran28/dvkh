import io
import os
import re
import glob
import tempfile
from typing import List, Dict, Tuple, Set

import streamlit as st
import pandas as pd
import numpy as np

# ============ CẤU HÌNH TRANG ============
st.set_page_config(page_title="DVKH/HDV Toolkit (Drive)", layout="wide")

st.title("DVKH / HDV — Xử lý từ Google Drive")
st.caption("Đọc nhiều file từ 2 thư mục Drive (CKH riêng, còn lại chung), xử lý & xuất Excel.")

# ============ GDRIVE (Service Account) ============
# Bạn cần điền JSON Service Account vào st.secrets["gcp_service_account"]
# và share 2 folder ID cho email của Service Account (Viewer).
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload


@st.cache_resource
def get_drive_service():
    creds = service_account.Credentials.from_service_account_info(
        st.secrets["gcp_service_account"],
        scopes=["https://www.googleapis.com/auth/drive.readonly"],
    )
    return build("drive", "v3", credentials=creds)


def list_files_in_folder(service, folder_id: str, name_contains: str | None = None) -> List[Dict]:
    """Liệt kê file trong folder. Có thể lọc theo tên chứa chuỗi."""
    q = f"'{folder_id}' in parents and trashed=false"
    fields = "nextPageToken, files(id, name, mimeType)"
    files, token = [], None
    while True:
        resp = service.files().list(q=q, fields=fields, pageToken=token).execute()
        found = resp.get("files", [])
        if name_contains:
            found = [f for f in found if name_contains.lower() in f["name"].lower()]
        files.extend(found)
        token = resp.get("nextPageToken")
        if not token:
            break
    return files


def download_drive_file(service, file_id: str, out_path: str):
    req = service.files().get_media(fileId=file_id)
    with open(out_path, "wb") as fh:
        downloader = MediaIoBaseDownload(fh, req)
        done = False
        while not done:
            _, done = downloader.next_chunk()


# ============ HELPERS ============

def read_excel_any(path: str, dtype=None) -> pd.DataFrame:
    """
    Đọc excel .xlsx bằng openpyxl. Thử .xls bằng engine mặc định.
    Khuyến nghị: chuyển .xls -> .xlsx để ổn định trên Streamlit Cloud.
    """
    try:
        if path.lower().endswith(".xlsx"):
            return pd.read_excel(path, dtype=dtype, engine="openpyxl")
        return pd.read_excel(path, dtype=dtype)
    except Exception as e:
        raise RuntimeError(f"Lỗi đọc '{os.path.basename(path)}': {e}\n"
                           f"👉 Hãy 'Save As' sang .xlsx nếu đây là .xls.")

def extract_name(value: str) -> str:
    parts = re.split(r"[-,]", str(value))
    for part in parts:
        name = part.strip()
        if re.fullmatch(r"[A-Z ]{3,}", name):
            return name
    return str(value).strip()

def phan_loai_tk_factory(set_ckh: Set[str], set_kkh: Set[str]):
    def _f(tk: str) -> str:
        if tk in set_ckh:
            return "CKH"
        if tk in set_kkh:
            return "KKH"
        return "NA"
    return _f


# ============ GIAO DIỆN NHẬP ============
with st.sidebar:
    st.header("Thiết lập Drive")
    ckh_folder_id = st.text_input("Folder ID chứa CKH (chỉ CKH)", help="Share folder này cho service account.")
    common_folder_id = st.text_input("Folder ID chứa các file còn lại (KKH, MUC30, DK_SMS, SCM010)", help="Share folder này cho service account.")
    st.markdown("---")
    st.subheader("Tên file/mẫu cần tìm trong thư mục chung")
    kkh_pattern = st.text_input("Mẫu file KKH", value="HDV_CHITIET_KKH_", help="Khớp tên bắt đầu hoặc chứa chuỗi này")
    muc30_filename = st.text_input("Tên file MUC30", value="MUC 30 2241.xlsx")
    dk_sms_filename = st.text_input("Tên file DK_SMS", value="Muc14_DK_SMS.txt")
    scm010_filename = st.text_input("Tên file SCM010", value="Muc14_SCM010.xlsx")
    st.markdown("---")
    run = st.button("Chạy xử lý")

# ============ XỬ LÝ ============
if run:
    if not ckh_folder_id or not common_folder_id:
        st.error("Vui lòng điền đủ 2 Folder ID.")
        st.stop()

    service = get_drive_service()
    tmpdir = tempfile.mkdtemp()
    st.info(f"Thư mục tạm: {tmpdir}")

    # 1) CKH (folder riêng)
    st.subheader("1) Tải CKH")
    files_ckh = list_files_in_folder(service, ckh_folder_id, name_contains="HDV_CHITIET_CKH_")
    if not files_ckh:
        st.error("Không tìm thấy file CKH (HDV_CHITIET_CKH_*) trong folder CKH.")
        st.stop()

    local_ckh = []
    for f in files_ckh:
        outp = os.path.join(tmpdir, f["name"])
        download_drive_file(service, f["id"], outp)
        local_ckh.append(outp)
    st.success(f"Đã tải {len(local_ckh)} file CKH")

    # 2) KKH (trong folder chung)
    st.subheader("2) Tải KKH")
    files_kkh = list_files_in_folder(service, common_folder_id, name_contains=kkh_pattern or "HDV_CHITIET_KKH_")
    if not files_kkh:
        st.error("Không tìm thấy file KKH trong folder chung.")
        st.stop()
    local_kkh = []
    for f in files_kkh:
        outp = os.path.join(tmpdir, f["name"])
        download_drive_file(service, f["id"], outp)
        local_kkh.append(outp)
    st.success(f"Đã tải {len(local_kkh)} file KKH")

    # 3) 3 file còn lại trong folder chung: MUC30, DK_SMS, SCM010
    st.subheader("3) Tải file MUC30, DK_SMS, SCM010")
    def get_single_file(folder_id: str, exact_name: str) -> str:
        found = list_files_in_folder(service, folder_id)
        for f in found:
            if f["name"].strip().lower() == exact_name.strip().lower():
                outp = os.path.join(tmpdir, f["name"])
                download_drive_file(service, f["id"], outp)
                return outp
        raise FileNotFoundError(f"Không thấy '{exact_name}' trong folder chung.")

    try:
        path_muc30 = get_single_file(common_folder_id, muc30_filename)
        path_dksms = get_single_file(common_folder_id, dk_sms_filename)
        path_scm10 = get_single_file(common_folder_id, scm010_filename)
        st.success("Đã tải đủ MUC30, DK_SMS, SCM010")
    except Exception as e:
        st.error(str(e))
        st.stop()

    # 4) Đọc dữ liệu
    st.subheader("4) Đọc dữ liệu & chuẩn hóa")
    try:
        df_b_CKH = pd.concat([read_excel_any(p, dtype=str) for p in local_ckh], ignore_index=True)
        df_b_KKH = pd.concat([read_excel_any(p, dtype=str) for p in local_kkh], ignore_index=True)
        df_b = pd.concat([df_b_CKH, df_b_KKH], ignore_index=True)

        # MUC 30
        df_a = read_excel_any(path_muc30, dtype=str)
        df_a = df_a[df_a["DESCRIPTION"].str.contains(r"chu\s*ky|chuky|cky", case=False, na=False)]
        df_a["EXPIRYDATE"] = pd.to_datetime(df_a["EXPIRYDATE"], format="%Y%m%d", errors="coerce").dt.strftime("%m/%d/%Y")
        df_a["EFFECTIVEDATE"] = pd.to_datetime(df_a["EFFECTIVEDATE"], format="%Y%m%d", errors="coerce").dt.strftime("%m/%d/%Y")
        # loại DN
        keywords = ["CONG TY", "CTY", "CONGTY", "CÔNG TY", "CÔNGTY"]
        df_a = df_a[~df_a["NGUOI_UY_QUYEN"].str.upper().str.contains("|".join(keywords), na=False)]
        # tách tên người được UQ
        df_a["NGUOI_DUOC_UY_QUYEN"] = df_a["NGUOI_DUOC_UY_QUYEN"].apply(extract_name)
        df_a = df_a.drop_duplicates(subset=["PRIMARY_SOL_ID", "TK_DUOC_UY_QUYEN", "NGUOI_DUOC_UY_QUYEN"])

        # MERGE lấy CIF người ủy quyền
        df_a["TK_DUOC_UY_QUYEN"] = df_a["TK_DUOC_UY_QUYEN"].astype(str)
        df_b["IDXACNO"] = df_b["IDXACNO"].astype(str)

        merged = df_a.merge(
            df_b[["IDXACNO", "CUSTSEQ"]],
            left_on="TK_DUOC_UY_QUYEN",
            right_on="IDXACNO",
            how="left"
        )
        merged["CIF_NGUOI_UY_QUYEN"] = merged["CUSTSEQ"].apply(
            lambda x: str(int(float(x))) if pd.notna(x) and str(x).strip() != "" else "NA"
        )
        # bù CIF theo NGUOI_UY_QUYEN
        cif_updated = merged["CIF_NGUOI_UY_QUYEN"].copy()
        for _, grp in merged.groupby("NGUOI_UY_QUYEN"):
            cifs = [c for c in grp["CIF_NGUOI_UY_QUYEN"].unique() if c != "NA"]
            if cifs:
                cif_to_fill = cifs[0]
                idx_to_fill = grp[grp["CIF_NGUOI_UY_QUYEN"] == "NA"].index
                cif_updated.loc[idx_to_fill] = cif_to_fill
        merged["CIF_NGUOI_UY_QUYEN"] = cif_updated

        for col_drop in ["MODIFIEDDATE_NEW", "IDXACNO", "CUSTSEQ"]:
            if col_drop in merged.columns:
                merged.drop(columns=[col_drop], inplace=True)

        # phân loại LOAI_TK nhờ tập CKH/KKH
        set_ckh = set(df_b_CKH["CUSTSEQ"].astype(str)) if "CUSTSEQ" in df_b_CKH.columns else set()
        set_kkh = set(df_b_KKH["IDXACNO"].astype(str)) if "IDXACNO" in df_b_KKH.columns else set()
        merged["LOAI_TK"] = merged["TK_DUOC_UY_QUYEN"].astype(str).apply(phan_loai_tk_factory(set_ckh, set_kkh))

        # cờ thời hạn UQ
        merged["EXPIRYDATE_dt"] = pd.to_datetime(merged["EXPIRYDATE"], errors="coerce")
        merged["EFFECTIVEDATE_dt"] = pd.to_datetime(merged["EFFECTIVEDATE"], errors="coerce")
        merged["YEAR_DIFF"] = merged["EXPIRYDATE_dt"].dt.year - merged["EFFECTIVEDATE_dt"].dt.year
        merged["KHONG_NHAP_TGIAN_UQ"] = np.where(merged["YEAR_DIFF"] == 99, "X", "")
        merged["UQ_TREN_50_NAM"] = np.where(merged["YEAR_DIFF"] >= 50, "X", "")
        merged.drop(columns=["YEAR_DIFF"], inplace=True, errors="ignore")

        # DK_SMS
        df_sms = pd.read_csv(path_dksms, sep="\t", on_bad_lines="skip", dtype=str)
        df_sms["FORACID"] = df_sms["FORACID"].astype(str)
        df_sms["ORGKEY"] = df_sms["ORGKEY"].astype(str)
        df_sms["C_MOBILE_NO"] = df_sms["C_MOBILE_NO"].astype(str)
        df_sms["CRE DATE"] = pd.to_datetime(df_sms["CRE_DATE"], errors="coerce").dt.strftime("%m/%d/%Y")
        df_sms = df_sms[df_sms["FORACID"].str.match(r"^\d+$", na=False)]
        df_sms = df_sms[df_sms["CUSTTPCD"].str.upper() != "KHDN"]

        # SCM010
        df_scm10 = read_excel_any(path_scm10, dtype=str)
        df_scm10 = df_scm10.rename(columns=lambda x: x.strip())
        df_scm10["CIF_ID"] = df_scm10["CIF_ID"].astype(str)

        # Gắn cờ TK có SMS & CIF có SCM010
        df_uyquyen = merged.copy()
        tk_sms_set = set(df_sms["FORACID"])
        df_uyquyen["TK có đăng ký SMS"] = df_uyquyen["TK_DUOC_UY_QUYEN"].astype(str).apply(
            lambda x: "X" if x in tk_sms_set else ""
        )
        cif_scm10_set = set(df_scm10["CIF_ID"])
        df_uyquyen["CIF có đăng ký SCM010"] = df_uyquyen["CIF_NGUOI_UY_QUYEN"].astype(str).apply(
            lambda x: "X" if x in cif_scm10_set else ""
        )

        # Tiêu chí 3: 1 người nhận UQ của nhiều người
        df_tc3 = df_uyquyen.copy()
        g = df_tc3.groupby("NGUOI_DUOC_UY_QUYEN")["NGUOI_UY_QUYEN"].nunique().reset_index()
        nguoi_nhan_nhieu = set(g.loc[g["NGUOI_UY_QUYEN"] >= 2, "NGUOI_DUOC_UY_QUYEN"])
        df_tc3["1 người nhận UQ của nhiều người"] = df_tc3["NGUOI_DUOC_UY_QUYEN"].apply(
            lambda x: "X" if x in nguoi_nhan_nhieu else ""
        )

        st.success("Đọc & xử lý xong.")
    except Exception as e:
        st.exception(e)
        st.stop()

    # 5) Hiển thị nhanh
    st.subheader("Xem nhanh dữ liệu")
    tabs = st.tabs(["Tiêu chí 1", "Tiêu chí 2", "Tiêu chí 3"])
    with tabs[0]:
        st.dataframe(merged.head(100), use_container_width=True)
    with tabs[1]:
        st.dataframe(df_uyquyen.head(100), use_container_width=True)
    with tabs[2]:
        st.dataframe(df_tc3.head(100), use_container_width=True)

    # 6) Xuất Excel
    st.subheader("Tải kết quả")
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as wr:
        merged.to_excel(wr, sheet_name="tieu chi 1", index=False)
        df_uyquyen.to_excel(wr, sheet_name="tieu chi 2", index=False)
        df_tc3.to_excel(wr, sheet_name="tieu chi 3", index=False)
    st.download_button(
        "⬇️ Tải DVKH_2241.xlsx",
        data=buffer.getvalue(),
        file_name="DVKH_2241.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

st.caption("Gợi ý: nếu gặp lỗi đọc .xls, hãy mở bằng Excel rồi 'Save As' sang .xlsx.")
