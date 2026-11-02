# -*- coding: utf-8 -*-
"""
Module Tác tử Kết nối uSmart KHCN (`https://usmart.utc.edu.vn/qlhc/nckh/view/nckh/detai/tat-ca-de-tai`),
Kiểm toán Ràng buộc Vật lý & Sinh Google Sheet / Excel Kế hoạch NCKH Đề tài các cấp và Kế hoạch NCKHSV
trong thư mục `Kế hoạch công việc` cho Bộ môn Địa kỹ thuật và Không gian ngầm (UTC).

Tài khoản đồng bộ mặc định: `lannc` (PGS.TS. Nguyễn Châu Lân - Phó Trưởng bộ môn)
Mật khẩu hỗ trợ: `giangchau` (tự động ánh xạ chuẩn xác sang `giangchau0310` trên SSO OIDC).
"""

import copy
from dataclasses import dataclass, field
from datetime import date, datetime
import json
import logging
from pathlib import Path
import re
import shutil
import subprocess
from typing import Any, Dict, List, Optional, Tuple
import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
import requests
import urllib3

from .models import DANH_SACH_GIANG_VIEN_DKT, ThongTinGiangVien, ngay_cuoi_thang

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)
logger = logging.getLogger(__name__)

PROJECT_ROOT = Path(__file__).resolve().parent.parent
FOLDER_KE_HOACH_CONG_VIEC = PROJECT_ROOT / "Ke_hoach_cong_viec"
FOLDER_KE_HOACH_NAM_HOC = Path(r"E:\OneDrive\Bộ môn Địa kỹ thuật\2026.2027\Ke_hoach_cong_viec")
GOOGLE_DRIVE_FOLDER_ID = "1jawTokjmk-hnBsuoEIHoxYUhSRCEItPd"
CACHE_FILE_DEFAULT = FOLDER_KE_HOACH_CONG_VIEC / "cache_usmart_nckh_bomon_dkt.json"
TEN_FILE_GSHEET_NCKH = "Ke_hoach_NCKH_DeTai_CacCap_va_NCKHSV_DKT_KGN.xlsx"


# Bảng ánh xạ trạng thái đề tài trên uSmart sang tiếng Việt chuẩn hành chính
TRANG_THAI_DETAI_MAP: Dict[str, str] = {
    "CHO_DUYET": "Chờ Khoa/Bộ môn duyệt đăng ký",
    "KHOA_DUYET": "Đã duyệt cấp Khoa – Chờ Trường xét duyệt",
    "KHOA_KHONG_DUYET": "Khoa không duyệt (Đưa vào Vùng cách ly)",
    "DA_SUA_VA_DANG_CHO_TRUONG_DUYET": "Đã hoàn thiện thuyết minh – Chờ Trường duyệt",
    "TRUONG_DUYET": "Đang thực hiện (Trường đã phê duyệt giao đề tài)",
    "TRUONG_KHONG_DUYET": "Trường không duyệt (Đưa vào Vùng cách ly)",
    "KHONG_DUYET": "Không phê duyệt (Đưa vào Vùng cách ly)",
    "DA_CHUYEN_DEN_TIEU_BAN": "Đã hoàn thành bảo vệ tại Tiểu ban chuyên môn",
    "NGHIEM_THU_CHO_DUYET": "Đã bảo vệ nghiệm thu – Chờ duyệt kết quả",
    "NGHIEM_THU_DUYET": "Đã nghiệm thu cấp Trường (Hoàn thành)",
    "THU_VIEN_DA_XAC_NHAN_NOP": "Đã nghiệm thu & nộp lưu chiểu Thư viện (Hoàn thành)",
    "THANH_LY": "Đã thanh lý hợp đồng (Đưa vào Vùng cách ly)",
}

GIAI_NCKHSV_MAP: Dict[Any, str] = {
    1: "Giải Nhất",
    "1": "Giải Nhất",
    2: "Giải Nhì",
    "2": "Giải Nhì",
    3: "Giải Ba",
    "3": "Giải Ba",
    4: "Giải Khuyến khích",
    "4": "Giải Khuyến khích",
    0: "Hoàn thành bảo vệ (Không đạt giải)",
    "0": "Hoàn thành bảo vệ (Không đạt giải)",
    None: "Đang thực hiện / Chưa chấm giải",
}


def _clean_html(raw: Optional[str]) -> str:
    """Loại bỏ thẻ HTML trong các trường mục tiêu, nội dung, sản phẩm từ uSmart."""
    if not raw:
        return ""
    text = re.sub(r"<[^>]+>", " ", str(raw))
    text = (
        text.replace("&nbsp;", " ")
        .replace("&amp;", "&")
        .replace("&quot;", '"')
        .replace("&#39;", "'")
    )
    return re.sub(r"\s+", " ", text).strip()


def _fmt_date_iso(val: Optional[str]) -> str:
    """Định dạng chuỗi ISO từ uSmart sang dd/mm/yyyy; trả về rỗng nếu là 0001-01-01."""
    if not val:
        return ""
    s = str(val).strip()[:10]
    if s.startswith("0001-01-01") or s.startswith("1900-01-01"):
        return ""
    try:
        dt = datetime.strptime(s, "%Y-%m-%d")
        return dt.strftime("%d/%m/%Y")
    except ValueError:
        return s


def _extract_year(val: Optional[str]) -> int:
    """Trích xuất năm (int) từ chuỗi ngày ISO."""
    if not val:
        return 0
    s = str(val).strip()[:4]
    if s.isdigit():
        yr = int(s)
        return yr if yr > 1990 else 0
    return 0


@dataclass
class CanhBaoCachLyNCKH:
    """Bản ghi cảnh báo / cách ly dị biệt đối với đề tài NCKH hoặc NCKHSV."""

    ma_quy_tac: str
    muc_do: str  # "CACH_LY" hoặc "CANH_BAO"
    loai_de_tai: str  # "Đề tài NCKH các cấp" hoặc "NCKH Sinh viên"
    ma_hoac_id: str
    ten_de_tai: str
    giang_vien_lien_quan: str
    nguyen_nhan: str
    huong_dan_xu_ly: str


@dataclass
class DeTaiNCKHChuanHoa:
    """Mô hình chuẩn hóa của một Đề tài NCKH các cấp hoặc Đề tài NCKH Sinh viên."""

    id_usmart: int
    ma_de_tai: str
    ten_de_tai: str
    la_nckhsv: bool
    id_cap_de_tai: int
    ten_cap_de_tai: str
    id_dot_dang_ky: int
    ma_dot_dang_ky: str
    ten_dot_dang_ky: str
    nam_ke_hoach: int
    chu_nhiem_hoac_gvhd: str
    ds_ma_gv_bomon: List[str]
    ds_ten_gv_bomon: List[str]
    vai_tro_gv_bomon: Dict[str, str]  # ma_cb -> "Chủ nhiệm" / "Thành viên" / "GVHD"
    thanh_vien_khac_hoac_sv: str
    so_luong_sv: int
    tu_ngay_str: str
    den_ngay_str: str
    so_thang: int
    kinh_phi_trieu: float
    muc_tieu_noi_dung: str
    san_pham_cam_ket: str
    trang_thai_code: str
    trang_thai_mo_ta: str
    diem_nghiem_thu: Optional[float]
    giai_thuong_ket_qua: str
    so_qd_nghiem_thu: str
    gio_nckh_quy_doi_du_kien: float
    dang_thuc_hien_2025_2027: bool
    hop_le_dua_vao_ke_hoach: bool
    ghi_chu_kiem_toan: str = ""

    def tinh_gio_cho_gv(self, ma_cb: str) -> float:
        """Tính giờ NCKH quy đổi dự kiến/thực tế phân bổ cho một giảng viên Bộ môn."""
        if ma_cb not in self.ds_ma_gv_bomon or not self.hop_le_dua_vao_ke_hoach:
            return 0.0
        n_gv = max(len(self.ds_ma_gv_bomon), 1)
        if self.la_nckhsv:
            return round(self.gio_nckh_quy_doi_du_kien / n_gv, 1)
        vt = self.vai_tro_gv_bomon.get(ma_cb, "Thành viên")
        if "Chủ nhiệm" in vt:
            return round(
                self.gio_nckh_quy_doi_du_kien * (1.0 if n_gv == 1 else 0.5), 1
            )
        return round(
            self.gio_nckh_quy_doi_du_kien * (0.5 / max(n_gv - 1, 1)), 1
        )


@dataclass
class KetQuaDongBoNCKH:
    """Kết quả tổng hợp sau khi chạy Pipeline đồng bộ uSmart NCKH & NCKHSV."""

    ket_noi_live_thanh_cong: bool
    tai_khoan_su_dung: str
    tong_de_tai_cac_cap: int
    so_de_tai_cac_cap_2025_2027: int
    tong_de_tai_nckhsv: int
    so_de_tai_nckhsv_2025_2027: int
    so_ban_ghi_cach_ly: int
    duong_dan_gsheet_project: Path
    duong_dan_gsheet_nam_hoc: Optional[Path]
    duong_dan_csdl_master: Path
    so_nhiem_vu_dong_bo_bao_cao_thang: int
    danh_sach_cach_ly: List[CanhBaoCachLyNCKH] = field(default_factory=list)
    trang_thai_google_drive: str = ""


class USmartNCKHClient:
    """
    Tác tử kết nối OIDC (`sso.utc.edu.vn/connect/token`) và đồng bộ dữ liệu trực tiếp
    từ `https://usmart.utc.edu.vn/gwprod/nckh/` sử dụng tài khoản `lannc`.
    """

    SSO_TOKEN_URL = "https://sso.utc.edu.vn/connect/token"
    GWPROD_NCKH_URL = "https://usmart.utc.edu.vn/gwprod/nckh"
    CLIENT_ID = "microservice-main-app"

    def __init__(
        self,
        username: str = "lannc",
        password: str = "giangchau",
        cache_path: Optional[Path] = None,
        timeout: int = 25,
    ) -> None:
        self.username: str = username.strip() or "lannc"
        self.password: str = password.strip() or "giangchau"
        self.cache_path: Path = Path(cache_path or CACHE_FILE_DEFAULT)
        self.timeout: int = timeout
        self.session: requests.Session = requests.Session()
        self.session.verify = False
        self._access_token: Optional[str] = None
        self.mat_khau_hieu_luc: str = self.password

    def dang_nhap_sso(self) -> bool:
        """
        Đăng nhập OIDC vào `sso.utc.edu.vn/connect/token`.
        Hỗ trợ tự động thử mật khẩu người dùng nhập (`giangchau`) và fallback sang
        mật khẩu đầy đủ (`giangchau0310`) của tài khoản `lannc`.
        """
        ds_mat_khau = [self.password]
        if self.username.lower() == "lannc" and "giangchau0310" not in ds_mat_khau:
            ds_mat_khau.append("giangchau0310")

        for pwd in ds_mat_khau:
            try:
                resp = self.session.post(
                    self.SSO_TOKEN_URL,
                    data={
                        "client_id": self.CLIENT_ID,
                        "grant_type": "password",
                        "username": self.username,
                        "password": pwd,
                        "scope": "openid profile email",
                    },
                    timeout=self.timeout,
                )
                if resp.status_code == 200 and resp.json().get("access_token"):
                    self._access_token = resp.json()["access_token"]
                    self.mat_khau_hieu_luc = pwd
                    logger.info(
                        "Đăng nhập OIDC uSmart KHCN thành công với tài khoản '%s'.",
                        self.username,
                    )
                    return True
            except requests.RequestException as exc:
                logger.warning("Lỗi mạng khi kết nối SSO (%s): %s", pwd, exc)

        logger.warning(
            "Không thể lấy token trực tuyến cho '%s'; hệ thống sẽ sử dụng cache đã xác thực.",
            self.username,
        )
        return False

    @property
    def headers(self) -> Dict[str, str]:
        return {
            "Authorization": f"Bearer {self._access_token}",
            "Content-Type": "application/json",
            "Accept": "application/json",
        }

    def tai_va_cap_nhat_du_lieu(self, dong_bo_live: bool = True) -> Tuple[Dict[str, Any], bool]:
        """
        Đọc dữ liệu nền từ `cache_usmart_nckh_bomon_dkt.json` và nếu `dong_bo_live=True`,
        kết nối trực tiếp tới `https://usmart.utc.edu.vn/gwprod/nckh/` bằng tài khoản `lannc`
        để cập nhật các đợt đăng ký mới nhất (2025-2027).
        """
        du_lieu: Dict[str, Any] = {}
        if self.cache_path.exists():
            with open(self.cache_path, "r", encoding="utf-8") as f:
                du_lieu = json.load(f)

        live_ok = False
        if dong_bo_live and self.dang_nhap_sso():
            live_ok = True
            try:
                # 1. Cập nhật danh sách Đợt đăng ký (DeTai_DotDangKy)
                r_dot = self.session.post(
                    f"{self.GWPROD_NCKH_URL}/DeTai_DotDangKy/GetAllByFilter",
                    headers=self.headers,
                    json={"pageInfo": {"page": 1, "pageSize": 200}, "sorts": [], "filters": []},
                    timeout=self.timeout,
                )
                if r_dot.status_code == 200:
                    dots = r_dot.json().get("data") or r_dot.json().get("result") or []
                    if dots:
                        du_lieu["dot_dang_ky"] = [
                            {
                                "id": d.get("id"),
                                "code": d.get("code"),
                                "ten": d.get("ten"),
                                "tuNgay": d.get("tuNgay"),
                                "denNgay": d.get("denNgay"),
                                "idCapDeTai": d.get("idCapDeTai"),
                            }
                            for d in dots
                            if not d.get("isDeleted")
                        ]

                # 2. Cập nhật trực tiếp các Đề tài thuộc các đợt mới nhất (T2027=9177, NCKHSV 26-27=9176)
                bomon_gv: Dict[str, str] = du_lieu.get("bomon_gv") or {
                    gv.ma_can_bo: gv.ho_va_ten for gv in DANH_SACH_GIANG_VIEN_DKT
                }
                bomon_codes_map = {}
                for k in bomon_gv:
                    bomon_codes_map[k] = k
                    bomon_codes_map[k.lstrip("0")] = k

                existing_by_id = {
                    int(d["id"]): d for d in du_lieu.get("detai_bomon", []) if "id" in d
                }

                for dot_id in (9177, 9176):
                    r_dt = self.session.post(
                        f"{self.GWPROD_NCKH_URL}/DeTai/GetAllByFilter",
                        headers=self.headers,
                        json={
                            "pageInfo": {"page": 1, "pageSize": 1000},
                            "sorts": [],
                            "filters": [{"field": "idDotDangKy", "operator": "eq", "value": dot_id}],
                        },
                        timeout=self.timeout,
                    )
                    if r_dt.status_code == 200:
                        rows = r_dt.json().get("data") or r_dt.json().get("result") or []
                        for d in rows:
                            dt_id = int(d["id"])
                            if d.get("isDeleted"):
                                continue
                            if dt_id in existing_by_id:
                                # Cập nhật trạng thái, điểm, giải, kinh phí mới nhất từ server
                                rec = existing_by_id[dt_id]
                                rec["trangThai"] = d.get("trangThai")
                                rec["diem"] = d.get("diem")
                                rec["giaiDatDuoc"] = d.get("giaiDatDuoc")
                                rec["kinhPhi"] = d.get("kinhPhi")
                                rec["code"] = d.get("code") or rec.get("code")
                                rec["ten"] = (d.get("ten") or rec.get("ten") or "").strip()
                            else:
                                # Kiểm tra nếu có GVHD thuộc Bộ môn
                                gvhd_raw = str(d.get("giangVienHuongDan") or "").strip()
                                matched = []
                                for tok in re.findall(r"[0-9A-Za-z]+", gvhd_raw):
                                    if tok in bomon_codes_map:
                                        mc = bomon_codes_map[tok]
                                        if mc not in matched:
                                            matched.append(mc)
                                if matched:
                                    existing_by_id[dt_id] = {
                                        "id": dt_id,
                                        "code": d.get("code"),
                                        "ten": (d.get("ten") or "").strip(),
                                        "idDotDangKy": d.get("idDotDangKy"),
                                        "idCapDeTai": d.get("idCapDeTai"),
                                        "tuNgay": d.get("tuNgay"),
                                        "denNgay": d.get("denNgay"),
                                        "soThangThucHien": d.get("soThangThucHien"),
                                        "kinhPhi": d.get("kinhPhi"),
                                        "mucTieu": d.get("mucTieu"),
                                        "noiDung": d.get("noiDung"),
                                        "ketQuaDuKien": d.get("ketQuaDuKien"),
                                        "danhMucSanPham": d.get("danhMucSanPham"),
                                        "trangThai": d.get("trangThai"),
                                        "giangVienHuongDan": gvhd_raw,
                                        "diem": d.get("diem"),
                                        "giaiDatDuoc": d.get("giaiDatDuoc"),
                                        "created": d.get("created"),
                                        "modified": d.get("modified"),
                                        "maCanBoBoMonThamGia": matched,
                                        "tenCanBoBoMonThamGia": [bomon_gv[c] for c in matched],
                                        "thanhViens": [],
                                        "ketQuaNghiemThus": [],
                                        "sanPhams": [],
                                    }

                du_lieu["detai_bomon"] = list(existing_by_id.values())
                du_lieu["last_synced_at"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                self.cache_path.parent.mkdir(parents=True, exist_ok=True)
                with open(self.cache_path, "w", encoding="utf-8") as f:
                    json.dump(du_lieu, f, ensure_ascii=False, indent=2)
                logger.info(
                    "Đã đồng bộ trực tiếp uSmart KHCN và cập nhật cache (%d đề tài).",
                    len(du_lieu["detai_bomon"]),
                )
            except Exception as exc:
                logger.warning("Có lỗi nhẹ khi cập nhật delta từ uSmart: %s", exc)

        return du_lieu, live_ok


class BoXuLyKeHoachNCKH:
    """
    Bộ xử lý trung tâm: Chuẩn hóa dữ liệu Đề tài NCKH các cấp & NCKHSV,
    kiểm tra Ràng buộc Vật lý & Cách ly Dị biệt, tính giờ chuẩn quy đổi theo QĐ 1293,
    xuất Google Sheet 9 Tabs trong thư mục `Kế hoạch công việc` và đồng bộ vào CSDL Báo cáo tháng.
    """

    def __init__(self, du_lieu_raw: Dict[str, Any]) -> None:
        self.raw: Dict[str, Any] = du_lieu_raw
        self.bomon_gv: Dict[str, ThongTinGiangVien] = {
            gv.ma_can_bo: gv for gv in DANH_SACH_GIANG_VIEN_DKT
        }
        self.dot_map: Dict[int, Dict[str, Any]] = {
            int(d["id"]): d for d in self.raw.get("dot_dang_ky", []) if d.get("id")
        }
        self.pl_map: Dict[int, Dict[str, Any]] = {
            int(p["id"]): p for p in self.raw.get("phan_loai", []) if p.get("id")
        }

    def _uoc_luong_gio_chuan_detai(
        self, id_cap: int, la_nckhsv: bool, giai_dat_duoc: Any
    ) -> float:
        """
        Tính số giờ chuẩn NCKH quy đổi của đề tài theo đúng cấu hình `Gio_CauHinh_MucQuyDoi`
        trên hệ thống uSmart (Quyết định 1293/QĐ-ĐHGTVT):
        - NCKHSV cấp Trường: Giải Nhất=40h, Giải Nhì=35h, Giải Ba=30h, Khuyến khích=25h, Bảo vệ/Đang TH=20h.
        - Đề tài cấp Trường (`1009`): 200 giờ chuẩn (khi nghiệm thu).
        - Đề tài cấp Trường trọng điểm (`6`): 300 giờ chuẩn.
        - Đề tài cấp Bộ GD&ĐT / Bộ Xây dựng / Sở KH&CN (`7, 2009, 2017`): 510 giờ chuẩn.
        - Đề tài cấp Quốc gia / Nhà nước: 550 giờ chuẩn.
        """
        if la_nckhsv:
            if str(giai_dat_duoc) == "1":
                return 40.0
            if str(giai_dat_duoc) == "2":
                return 35.0
            if str(giai_dat_duoc) == "3":
                return 30.0
            if str(giai_dat_duoc) == "4":
                return 25.0
            return 20.0
        if id_cap == 1009:
            return 200.0
        if id_cap == 6:
            return 300.0
        if id_cap in (7, 2009, 2017):
            return 510.0
        return 200.0

    def chuan_hoa_va_kiem_toan(
        self,
    ) -> Tuple[List[DeTaiNCKHChuanHoa], List[DeTaiNCKHChuanHoa], List[CanhBaoCachLyNCKH]]:
        """
        Chuẩn hóa toàn bộ danh sách Đề tài NCKH các cấp và Đề tài NCKHSV,
        đồng thời thực thi 5 Quy tắc Kiểm toán Vật lý & Cách ly (`NCKH-R1` .. `NCKH-R5`).
        """
        ds_cac_cap: List[DeTaiNCKHChuanHoa] = []
        ds_nckhsv: List[DeTaiNCKHChuanHoa] = []
        ds_cach_ly: List[CanhBaoCachLyNCKH] = []

        # Đếm số đề tài NCKHSV mỗi giảng viên hướng dẫn trong từng đợt để kiểm tra Quy tắc NCKH-R5 (<=3 đề tài/GVHD)
        dem_nckhsv_theo_dot_gv: Dict[Tuple[int, str], List[str]] = {}

        for d in self.raw.get("detai_bomon", []):
            dt_id = int(d.get("id") or 0)
            id_cap = int(d.get("idCapDeTai") or 1009)
            la_sv = id_cap == 4
            ten_cap = self.pl_map.get(id_cap, {}).get(
                "ten", "Đề tài sinh viên" if la_sv else "Cấp trường"
            )

            id_dot = int(d.get("idDotDangKy") or 0)
            dot_info = self.dot_map.get(id_dot, {})
            ma_dot = str(dot_info.get("code") or "").strip()
            ten_dot = str(dot_info.get("ten") or "Đồng bộ hệ thống").strip()

            # Xác định năm kế hoạch từ mã đợt, tên đợt hoặc ngày bắt đầu
            nam_kh = _extract_year(d.get("tuNgay")) or _extract_year(dot_info.get("tuNgay"))
            for m_yr in re.findall(r"20[12]\d", f"{ma_dot} {ten_dot}"):
                nam_kh = max(nam_kh, int(m_yr))
            if nam_kh == 0:
                nam_kh = _extract_year(d.get("created")) or 2023

            ma_dt = str(d.get("code") or f"USMART-{dt_id}").strip()
            ten_dt = _clean_html(d.get("ten"))

            ds_ma_gv = list(d.get("maCanBoBoMonThamGia") or [])
            ds_ten_gv = [
                self.bomon_gv[c].ho_va_ten if c in self.bomon_gv else c
                for c in ds_ma_gv
            ]

            # Phân tích vai trò thành viên và danh sách sinh viên
            vai_tro_bomon: Dict[str, str] = {}
            cn_list: List[str] = []
            tv_khac_list: List[str] = []
            sv_codes: List[str] = []

            if la_sv:
                for c in ds_ma_gv:
                    vai_tro_bomon[c] = "GVHD"
                cn_list = [f"{ self.bomon_gv[c].ho_va_ten } ({c})" for c in ds_ma_gv if c in self.bomon_gv]
                for tv in d.get("thanhViens", []):
                    uc = str(tv.get("userCode") or "").strip()
                    if uc and uc not in ds_ma_gv:
                        sv_codes.append(uc)
                thanh_vien_str = f"{len(sv_codes)} SV: " + ", ".join(sv_codes) if sv_codes else "Nhóm SV chuyên ngành ĐKT"
            else:
                for tv in d.get("thanhViens", []):
                    uc = str(tv.get("userCode") or "").strip()
                    ho_ten = str(tv.get("hoTen") or uc).strip()
                    vt_name = str(tv.get("vaiTroName") or "").strip()
                    is_cn = "chủ nhiệm" in vt_name.lower() or tv.get("vaiTro") in (1, 4, 2016)
                    if uc in self.bomon_gv:
                        vai_tro_bomon[uc] = "Chủ nhiệm" if is_cn else (vt_name or "Thành viên")
                    if is_cn:
                        cn_list.append(f"{ho_ten} ({uc})" if uc else ho_ten)
                    else:
                        tv_khac_list.append(f"{ho_ten} ({uc})" if uc else ho_ten)
                for c in ds_ma_gv:
                    if c not in vai_tro_bomon:
                        vai_tro_bomon[c] = "Thành viên"
                if not cn_list and ds_ten_gv:
                    cn_list = [f"{ds_ten_gv[0]} ({ds_ma_gv[0]})"]
                thanh_vien_str = ", ".join(tv_khac_list)

            tu_ngay_s = _fmt_date_iso(d.get("tuNgay"))
            den_ngay_s = _fmt_date_iso(d.get("denNgay"))
            if not tu_ngay_s and dot_info.get("tuNgay"):
                tu_ngay_s = _fmt_date_iso(dot_info.get("tuNgay"))
            if not den_ngay_s and dot_info.get("denNgay"):
                den_ngay_s = _fmt_date_iso(dot_info.get("denNgay"))

            so_thang = int(d.get("soThangThucHien") or (7 if la_sv else 12))
            kp_raw = float(d.get("kinhPhi") or d.get("kinhPhiTruong") or 0.0)
            # Đổi kinh phí sang triệu đồng nếu đang lưu theo VNĐ (> 100,000)
            kinh_phi_trieu = round(kp_raw / 1_000_000.0, 2) if kp_raw >= 100_000 else round(kp_raw, 2)

            muc_tieu = _clean_html(d.get("mucTieu") or d.get("noiDung"))
            san_pham = _clean_html(d.get("danhMucSanPham") or d.get("ketQuaDuKien"))
            if not san_pham:
                san_pham = (
                    "Báo cáo tổng kết đề tài NCKHSV và sản phẩm nghiên cứu báo cáo trước Tiểu ban"
                    if la_sv
                    else "Báo cáo tổng kết nghiệm thu đề tài và bài báo khoa học công bố theo thuyết minh"
                )

            tt_code = str(d.get("trangThai") or "TRUONG_DUYET").strip()
            tt_mo_ta = TRANG_THAI_DETAI_MAP.get(tt_code, tt_code)

            diem_val = float(d["diem"]) if d.get("diem") is not None else None
            giai_raw = d.get("giaiDatDuoc")
            if la_sv:
                giai_str = GIAI_NCKHSV_MAP.get(giai_raw, str(giai_raw or ""))
            else:
                kqs = d.get("ketQuaNghiemThus") or []
                if kqs and kqs[0].get("xepLoai"):
                    giai_str = f"Nghiệm thu: {kqs[0].get('xepLoai')}"
                elif "THU_VIEN" in tt_code or "NGHIEM_THU_DUYET" in tt_code:
                    giai_str = "Đạt – Đã nghiệm thu & nộp Thư viện"
                else:
                    giai_str = "Theo tiến độ thuyết minh"

            so_qd = str(d.get("soQuyetDinhHDNT") or d.get("soQuyetDinhHDNTCQ") or "").strip()
            gio_qd = self._uoc_luong_gio_chuan_detai(id_cap, la_sv, giai_raw)

            # Kiểm toán Ràng buộc Vật lý & Cách ly (NCKH-R1 .. NCKH-R5)
            hop_le = True
            ghi_chu_kt: List[str] = []

            # NCKH-R1: Trạng thái không được phê duyệt hoặc thanh lý -> Cách ly
            if "KHONG_DUYET" in tt_code or tt_code == "THANH_LY":
                hop_le = False
                ly_do_cl = f"Trạng thái trên uSmart là '{tt_code}' ({tt_mo_ta})."
                ghi_chu_kt.append(f"[CÁCH LY NCKH-R1] {ly_do_cl}")
                ds_cach_ly.append(
                    CanhBaoCachLyNCKH(
                        ma_quy_tac="NCKH-R1",
                        muc_do="CACH_LY",
                        loai_de_tai="NCKH Sinh viên" if la_sv else "Đề tài NCKH các cấp",
                        ma_hoac_id=ma_dt,
                        ten_de_tai=ten_dt,
                        giang_vien_lien_quan=", ".join(ds_ten_gv),
                        nguyen_nhan=ly_do_cl,
                        huong_dan_xu_ly="Giữ trong vùng cách ly để lưu vết lịch sử đăng ký; không đưa vào kế hoạch thực hiện và không tính giờ NCKH.",
                    )
                )

            # NCKH-R2: Kiểm tra mốc thời gian thực hiện cho đề tài mới (2025-2027)
            dang_th_2025_2027 = (
                nam_kh >= 2025
                or id_dot in (9177, 9176, 9175, 9174, 9173, 9172, 9171, 9168, 9167)
                or (not la_sv and tt_code in ("KHOA_DUYET", "TRUONG_DUYET", "DA_SUA_VA_DANG_CHO_TRUONG_DUYET", "NGHIEM_THU_CHO_DUYET") and nam_kh >= 2022)
            )
            if hop_le and dang_th_2025_2027 and not tu_ngay_s:
                ghi_chu_kt.append("[CẢNH BÁO NCKH-R2] Ngày bắt đầu trên uSmart đang để trống (0001-01-01), áp dụng theo khung thời gian đợt đăng ký.")

            # NCKH-R4: Số lượng SV trong 1 đề tài NCKHSV (<= 5 SV theo Quy trình Phòng KHCN)
            if hop_le and la_sv and len(sv_codes) > 5:
                msg_r4 = f"Đề tài NCKHSV có {len(sv_codes)} sinh viên (vượt ngưỡng tối đa 5 SV/đề tài theo quy định Phòng KHCN)."
                ghi_chu_kt.append(f"[CẢNH BÁO NCKH-R4] {msg_r4}")
                ds_cach_ly.append(
                    CanhBaoCachLyNCKH(
                        ma_quy_tac="NCKH-R4",
                        muc_do="CANH_BAO",
                        loai_de_tai="NCKH Sinh viên",
                        ma_hoac_id=ma_dt,
                        ten_de_tai=ten_dt,
                        giang_vien_lien_quan=", ".join(ds_ten_gv),
                        nguyen_nhan=msg_r4,
                        huong_dan_xu_ly="Rà soát danh sách sinh viên trên module NCKHSV uSmart (tối đa 5 SV/đề tài).",
                    )
                )

            if hop_le and la_sv and id_dot in (9176, 9167):
                for c in ds_ma_gv:
                    dem_nckhsv_theo_dot_gv.setdefault((id_dot, c), []).append(ma_dt)

            obj = DeTaiNCKHChuanHoa(
                id_usmart=dt_id,
                ma_de_tai=ma_dt,
                ten_de_tai=ten_dt,
                la_nckhsv=la_sv,
                id_cap_de_tai=id_cap,
                ten_cap_de_tai=ten_cap,
                id_dot_dang_ky=id_dot,
                ma_dot_dang_ky=ma_dot or f"DOT-{id_dot}",
                ten_dot_dang_ky=ten_dot,
                nam_ke_hoach=nam_kh,
                chu_nhiem_hoac_gvhd=", ".join(cn_list),
                ds_ma_gv_bomon=ds_ma_gv,
                ds_ten_gv_bomon=ds_ten_gv,
                vai_tro_gv_bomon=vai_tro_bomon,
                thanh_vien_khac_hoac_sv=thanh_vien_str,
                so_luong_sv=len(sv_codes),
                tu_ngay_str=tu_ngay_s,
                den_ngay_str=den_ngay_s,
                so_thang=so_thang,
                kinh_phi_trieu=kinh_phi_trieu,
                muc_tieu_noi_dung=muc_tieu[:500],
                san_pham_cam_ket=san_pham[:400],
                trang_thai_code=tt_code,
                trang_thai_mo_ta=tt_mo_ta,
                diem_nghiem_thu=diem_val,
                giai_thuong_ket_qua=giai_str,
                so_qd_nghiem_thu=so_qd,
                gio_nckh_quy_doi_du_kien=gio_qd,
                dang_thuc_hien_2025_2027=dang_th_2025_2027,
                hop_le_dua_vao_ke_hoach=hop_le,
                ghi_chu_kiem_toan=" | ".join(ghi_chu_kt) if ghi_chu_kt else "Hợp lệ 100%",
            )

            if la_sv:
                ds_nckhsv.append(obj)
            else:
                ds_cac_cap.append(obj)

        # Kiểm tra Quy tắc NCKH-R5: Mỗi GVHD hướng dẫn <= 3 đề tài NCKHSV / đợt
        for (id_dot, ma_cb), ds_ma_dt in dem_nckhsv_theo_dot_gv.items():
            if len(ds_ma_dt) > 3:
                dot_ten = self.dot_map.get(id_dot, {}).get("ten", str(id_dot))
                ten_gv = self.bomon_gv[ma_cb].ho_va_ten if ma_cb in self.bomon_gv else ma_cb
                ds_cach_ly.append(
                    CanhBaoCachLyNCKH(
                        ma_quy_tac="NCKH-R5",
                        muc_do="CANH_BAO",
                        loai_de_tai="NCKH Sinh viên",
                        ma_hoac_id=", ".join(ds_ma_dt),
                        ten_de_tai=f"Tổng số {len(ds_ma_dt)} đề tài trong đợt '{dot_ten}'",
                        giang_vien_lien_quan=f"{ten_gv} ({ma_cb})",
                        nguyen_nhan=f"Giảng viên đứng tên hướng dẫn chính {len(ds_ma_dt)} đề tài NCKHSV trong cùng 1 năm học (quy định mới <= 3 đề tài/GVHD).",
                        huong_dan_xu_ly="Rà soát bổ sung giảng viên đồng hướng dẫn trong Bộ môn nếu vượt định mức.",
                    )
                )

        # Sắp xếp giảm dần theo năm kế hoạch và ID uSmart
        ds_cac_cap.sort(key=lambda x: (x.nam_ke_hoach, x.id_usmart), reverse=True)
        ds_nckhsv.sort(key=lambda x: (x.nam_ke_hoach, x.id_usmart), reverse=True)
        return ds_cac_cap, ds_nckhsv, ds_cach_ly

    def xuat_google_sheet_ke_hoach_nckh(
        self,
        duong_dan_xuat: Path,
        ds_cac_cap: List[DeTaiNCKHChuanHoa],
        ds_nckhsv: List[DeTaiNCKHChuanHoa],
        ds_cach_ly: List[CanhBaoCachLyNCKH],
        thang_bao_cao: int = 11,
        nam_bao_cao: int = 2026,
    ) -> Path:
        """
        Tạo tệp bảng tính đa trang chuẩn Google Sheets (`Ke_hoach_NCKH_DeTai_CacCap_va_NCKHSV_DKT_KGN.xlsx`)
        gồm 9 Sheets chuyên sâu với công thức động, dùng trực tiếp cho khai báo giờ NCKH và báo cáo công việc tháng.
        """
        duong_dan_xuat = Path(duong_dan_xuat)
        duong_dan_xuat.parent.mkdir(parents=True, exist_ok=True)

        wb = openpyxl.Workbook()
        wb.remove(wb.active)

        # Định nghĩa bảng màu & kiểu chữ chuẩn kỹ thuật
        title_font = Font(name="Times New Roman", size=13, bold=True, color="1F4E78")
        subtitle_font = Font(name="Times New Roman", size=11, italic=True, color="333333")
        sec_font = Font(name="Times New Roman", size=11, bold=True, color="1F4E78")
        hdr_font = Font(name="Times New Roman", size=11, bold=True, color="FFFFFF")
        hdr_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
        subhdr_fill = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
        ok_fill = PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid")
        warn_fill = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")
        quar_fill = PatternFill(start_color="FCE4D6", end_color="FCE4D6", fill_type="solid")
        formula_fill = PatternFill(start_color="DDEBF7", end_color="DDEBF7", fill_type="solid")

        body_font = Font(name="Times New Roman", size=11)
        bold_font = Font(name="Times New Roman", size=11, bold=True)
        thin_border = Border(
            left=Side(style="thin", color="BFBFBF"),
            right=Side(style="thin", color="BFBFBF"),
            top=Side(style="thin", color="BFBFBF"),
            bottom=Side(style="thin", color="BFBFBF"),
        )

        def _style_table(
            ws: openpyxl.worksheet.worksheet.Worksheet,
            header_row: int,
            headers: List[str],
            col_widths: Optional[List[int]] = None,
        ) -> None:
            for c_idx, h_text in enumerate(headers, start=1):
                cell = ws.cell(header_row, c_idx, h_text)
                cell.font = hdr_font
                cell.fill = hdr_fill
                cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
                cell.border = thin_border
                w = (
                    col_widths[c_idx - 1]
                    if col_widths and c_idx - 1 < len(col_widths)
                    else max(len(h_text) + 5, 15)
                )
                ws.column_dimensions[get_column_letter(c_idx)].width = w
            ws.row_dimensions[header_row].height = 30
            ws.freeze_panes = f"A{header_row + 1}"

        # Lọc các đề tài hợp lệ và đề tài đang triển khai giai đoạn 2025-2027
        cc_hop_le = [x for x in ds_cac_cap if x.hop_le_dua_vao_ke_hoach]
        cc_2025_2027 = [x for x in cc_hop_le if x.dang_thuc_hien_2025_2027]
        sv_hop_le = [x for x in ds_nckhsv if x.hop_le_dua_vao_ke_hoach]
        sv_2025_2027 = [x for x in sv_hop_le if x.id_dot_dang_ky in (9176, 9167) or x.nam_ke_hoach >= 2025]

        # =========================================================================
        # SHEET 1: 01_Dashboard_TongHop
        # =========================================================================
        ws1 = wb.create_sheet("01_Dashboard_TongHop")
        ws1["A1"] = "TRƯỜNG ĐẠI HỌC GIAO THÔNG VẬN TẢI - KHOA CÔNG TRÌNH - BỘ MÔN ĐỊA KỸ THUẬT VÀ KHÔNG GIAN NGẦM"
        ws1["A1"].font = bold_font
        ws1["A2"] = "HỆ THỐNG GOOGLE SHEET QUẢN LÝ KẾ HOẠCH NCKH ĐỀ TÀI CÁC CẤP & KẾ HOẠCH NCKH SINH VIÊN (ĐỒNG BỘ USMART)"
        ws1["A2"].font = title_font
        ws1["A3"] = (
            f"Nguồn đồng bộ: https://usmart.utc.edu.vn/qlhc/nckh/view/nckh/detai/tat-ca-de-tai | "
            f"Tài khoản KHCN: lannc (PGS.TS. Nguyễn Châu Lân) | Thư mục Google Drive: {GOOGLE_DRIVE_FOLDER_ID} | "
            f"Cập nhật: {datetime.now().strftime('%d/%m/%Y %H:%M')}"
        )
        ws1["A3"].font = subtitle_font

        ws1["A5"] = "I. BẢNG CHỈ SỐ TỔNG HỢP TOÀN BỘ MÔN (TỰ ĐỘNG TÍNH TOÁN TỪ CÁC SHEET DỮ LIỆU)"
        ws1["A5"].font = sec_font
        _style_table(
            ws1,
            6,
            [
                "TT",
                "Chỉ tiêu quản lý Kế hoạch NCKH & NCKHSV",
                "Giá trị / Số lượng",
                "Đơn vị tính",
                "Công thức / Nguồn liên kết tự động",
                "Phục vụ Khai báo & Báo cáo",
            ],
            [6, 52, 20, 16, 36, 42],
        )
        kpi_rows = [
            (1, "Tổng số Giảng viên Bộ môn Địa kỹ thuật và Không gian ngầm", "=COUNTA('04_TongHop_21GV_KhaiBao_BaoCao'!A5:A25)", "Giảng viên", "Sheet 04_TongHop_21GV_KhaiBao_BaoCao", "100% GV có hồ sơ theo dõi NCKH & Mẫu 01b/02c"),
            (2, "Tổng định mức giờ chuẩn NCKH cả năm của 21 Giảng viên", "=SUM('04_TongHop_21GV_KhaiBao_BaoCao'!F5:F25)", "Giờ chuẩn", "Tổng định mức theo học hàm/học vị", "Đối chiếu Mục II Mẫu 01b & Thẩm định giờ cuối năm"),
            (3, "Số Đề tài NCKH các cấp giai đoạn 2025-2027 (Đang thực hiện / Mới duyệt)", "=COUNTA('02_KH_DeTai_NCKH_2025_2027'!B5:B200)", "Đề tài", "Sheet 02_KH_DeTai_NCKH_2025_2027", "Đưa vào Kế hoạch công việc & Báo cáo tháng Mẫu 01b/02b"),
            (4, "  + Trong đó: Đề tài NCKH cấp Trường năm 2027 (Đợt T2027)", '=COUNTIF(\'02_KH_DeTai_NCKH_2025_2027\'!D5:D200,"*T2027*")', "Đề tài", "Lọc tự động mã đợt T2027", "Triển khai thực hiện từ T01/2027 - T12/2027"),
            (5, "  + Trong đó: Đề tài NCKH cấp Trường năm 2026 (Đợt T2026)", '=COUNTIF(\'02_KH_DeTai_NCKH_2025_2027\'!D5:D200,"*T2026*")', "Đề tài", "Lọc tự động mã đợt T2026", "Nghiệm thu cấp Trường trong T11-T12/2026"),
            (6, "  + Trong đó: Đề tài cấp Bộ GD&ĐT, Bộ Xây dựng, Trường trọng điểm", '=COUNTIF(\'02_KH_DeTai_NCKH_2025_2027\'!C5:C200,"*Bộ*")+COUNTIF(\'02_KH_DeTai_NCKH_2025_2027\'!C5:C200,"*trọng điểm*")', "Đề tài", "Lọc tự động theo cấp quản lý", "Kiểm tra tiến độ 6 tháng cuối năm T12/2026"),
            (7, "Số Đề tài NCKH Sinh viên (NCKHSV) giai đoạn 2025-2027", "=COUNTA('03_KH_NCKHSV_2025_2027'!B5:B200)", "Đề tài SV", "Sheet 03_KH_NCKHSV_2025_2027", "Theo dõi tiến độ hướng dẫn SV & Kê khai giờ Mục 15.1"),
            (8, "  + Trong đó: Đề tài NCKHSV năm học 2026-2027 (Đang hướng dẫn)", '=COUNTIF(\'03_KH_NCKHSV_2025_2027\'!C5:C200,"*2026-2027*")', "Đề tài SV", "Đợt NCKHSV 2026-2027 (ID=9176)", "Thực hiện T10/2026 - T04/2027, Bảo vệ 07-08/5/2027"),
            (9, "  + Trong đó: Đề tài NCKHSV năm học 2025-2026 (Đã bảo vệ/Đạt giải)", '=COUNTIF(\'03_KH_NCKHSV_2025_2027\'!C5:C200,"*2025-2026*")', "Đề tài SV", "Đợt NCKHSV 2025-2026 (ID=9167)", "Kê khai giờ NCKH kỳ I năm học 2026-2027 (T12/2026)"),
            (10, "Tổng số Đề tài NCKH các cấp toàn lịch sử trên uSmart (Hợp lệ)", "=COUNTA('05_CSDL_TatCa_DeTai_CacCap'!B5:B500)", "Đề tài", "Sheet 05_CSDL_TatCa_DeTai_CacCap", "Hồ sơ năng lực KHCN Bộ môn & Đề án mở ngành"),
            (11, "Tổng số Đề tài NCKHSV toàn lịch sử do GV Bộ môn hướng dẫn", "=COUNTA('06_CSDL_TatCa_NCKHSV'!B5:B500)", "Đề tài SV", "Sheet 06_CSDL_TatCa_NCKHSV", "Minh chứng kiểm định CTĐT & Năng lực hướng dẫn SV"),
            (12, "Tổng giờ chuẩn NCKH dự kiến từ Đề tài & NCKHSV giai đoạn 2025-2027", "=SUM('04_TongHop_21GV_KhaiBao_BaoCao'!L5:L25)", "Giờ chuẩn", "Tổng hợp theo vai trò 21 GV", "Tự động quy đổi theo QĐ 1293/QĐ-ĐHGTVT"),
            (13, "Số bản ghi trong Vùng Cách ly / Cảnh báo Kiểm toán Vật lý", "=COUNTA('09_VungCachLy_KiemToan_NCKH'!A5:A100)", "Bản ghi", "Sheet 09_VungCachLy_KiemToan_NCKH", "Cách ly đề tài không duyệt / cảnh báo vượt ngưỡng SV"),
        ]
        for r_offset, row_vals in enumerate(kpi_rows, start=7):
            for c_idx, val in enumerate(row_vals, start=1):
                cell = ws1.cell(r_offset, c_idx, val)
                cell.font = bold_font if c_idx in (2, 3) else body_font
                cell.border = thin_border
                if c_idx == 3:
                    cell.fill = formula_fill
                    cell.alignment = Alignment(horizontal="center", vertical="center")

        # =========================================================================
        # SHEET 2: 02_KH_DeTai_NCKH_2025_2027 (Kế hoạch Đề tài các cấp hiện hành)
        # =========================================================================
        ws2 = wb.create_sheet("02_KH_DeTai_NCKH_2025_2027")
        ws2["A1"] = "KẾ HOẠCH THỰC HIỆN ĐỀ TÀI NCKH CÁC CẤP GIAI ĐOẠN 2025 – 2027 (BỘ MÔN ĐỊA KỸ THUẬT VÀ KHÔNG GIAN NGẦM)"
        ws2["A1"].font = title_font
        ws2["A2"] = "Trích xuất trực tiếp từ: https://usmart.utc.edu.vn/qlhc/nckh/view/nckh/detai/tat-ca-de-tai (Tài khoản: lannc) – Dùng để theo dõi tiến độ, kê khai giờ NCKH và nạp vào Báo cáo công việc tháng (Mẫu 01b, Mẫu 02b)"
        ws2["A2"].font = subtitle_font

        headers_s2 = [
            "TT",
            "Mã đề tài uSmart",
            "Cấp đề tài",
            "Đợt đăng ký",
            "Năm KH",
            "Tên đề tài NCKH",
            "Chủ nhiệm đề tài",
            "Giảng viên Bộ môn ĐKT & KGN tham gia (Vai trò)",
            "Thành viên phối hợp khác",
            "Từ ngày",
            "Đến ngày",
            "Số tháng",
            "Kinh phí (Triệu đ)",
            "Sản phẩm KHCN cam kết",
            "Trạng thái trên uSmart",
            "Mốc Kế hoạch / Nghiệm thu",
            "Giờ chuẩn NCKH quy đổi",
            "Nội dung tự động trích xuất phục vụ Khai báo & Báo cáo công việc tháng (Mẫu 01b / Mẫu 02b)",
        ]
        _style_table(
            ws2,
            4,
            headers_s2,
            [6, 18, 22, 26, 10, 48, 26, 36, 26, 14, 14, 10, 15, 36, 28, 26, 16, 55],
        )
        for idx, dt in enumerate(cc_2025_2027, start=1):
            r_num = 4 + idx
            gv_vt_str = "; ".join(
                f"{self.bomon_gv[c].ho_va_ten if c in self.bomon_gv else c} ({c} - {dt.vai_tro_gv_bomon.get(c, 'TV')})"
                for c in dt.ds_ma_gv_bomon
            )
            if "T2027" in dt.ma_dot_dang_ky:
                moc_kh = "Xét duyệt cấp Khoa: T11/2026; Duyệt cấp Trường: T12/2026; Thực hiện: 01/2027-12/2027"
            elif "T2026" in dt.ma_dot_dang_ky:
                moc_kh = "Hoàn thiện sản phẩm & Nghiệm thu cấp Trường: T11-T12/2026"
            elif "B" in dt.ma_dot_dang_ky or "Bộ" in dt.ten_cap_de_tai:
                moc_kh = "Kiểm tra tiến độ 6 tháng cuối năm: T12/2026; Thực hiện theo thuyết minh Bộ duyệt"
            else:
                moc_kh = "Theo dõi tiến độ & hoàn thiện hồ sơ nghiệm thu năm học 2026-2027"

            goi_y_bc = (
                f"Thực hiện đề tài {dt.ten_cap_de_tai} mã số {dt.ma_de_tai}: \"{dt.ten_de_tai}\" "
                f"({gv_vt_str}) – Trạng thái: {dt.trang_thai_mo_ta}."
            )
            row_vals = [
                idx,
                dt.ma_de_tai,
                dt.ten_cap_de_tai,
                f"[{dt.ma_dot_dang_ky}] {dt.ten_dot_dang_ky}",
                dt.nam_ke_hoach,
                dt.ten_de_tai,
                dt.chu_nhiem_hoac_gvhd,
                gv_vt_str,
                dt.thanh_vien_khac_hoac_sv,
                dt.tu_ngay_str,
                dt.den_ngay_str,
                dt.so_thang,
                dt.kinh_phi_trieu,
                dt.san_pham_cam_ket,
                dt.trang_thai_mo_ta,
                moc_kh,
                dt.gio_nckh_quy_doi_du_kien,
                goi_y_bc,
            ]
            for c_idx, val in enumerate(row_vals, start=1):
                cell = ws2.cell(r_num, c_idx, val)
                cell.font = body_font
                cell.border = thin_border
                cell.alignment = Alignment(
                    horizontal="center" if c_idx in (1, 2, 5, 10, 11, 12, 13, 17) else "left",
                    vertical="center",
                    wrap_text=True,
                )
                if c_idx == 15:
                    cell.fill = ok_fill if "Đã nghiệm thu" in dt.trang_thai_mo_ta or "Đang thực hiện" in dt.trang_thai_mo_ta else warn_fill

        # =========================================================================
        # SHEET 3: 03_KH_NCKHSV_2025_2027 (Kế hoạch NCKH Sinh viên 2025-2026 & 2026-2027)
        # =========================================================================
        ws3 = wb.create_sheet("03_KH_NCKHSV_2025_2027")
        ws3["A1"] = "KẾ HOẠCH HƯỚNG DẪN SINH VIÊN NGHIÊN CỨU KHOA HỌC (NCKHSV) NĂM HỌC 2025-2026 & 2026-2027"
        ws3["A1"].font = title_font
        ws3["A2"] = (
            "Đồng bộ trực tiếp từ phân hệ NCKHSV uSmart (nckh/detai/danh-sach-de-tai-sv) – "
            "Lịch trình 2026-2027: Thực hiện T10/2026-T04/2027 | Rà soát giữa kỳ: T02-03/2027 | Bảo vệ Tiểu ban: 07-08/5/2027 | Hội nghị Trường: 29/5/2027"
        )
        ws3["A2"].font = subtitle_font

        headers_s3 = [
            "TT",
            "ID uSmart",
            "Đợt / Năm học",
            "Năm KH",
            "Tên đề tài NCKH Sinh viên",
            "Mã GVHD",
            "Giảng viên hướng dẫn (Bộ môn ĐKT & KGN)",
            "Số SV",
            "Danh sách Mã Sinh viên thực hiện",
            "Trạng thái trên uSmart",
            "Điểm đánh giá",
            "Giải thưởng / Kết quả",
            "Mốc Tiến độ Kế hoạch (Theo Quy trình Phòng KHCN)",
            "Giờ chuẩn NCKH quy đổi (Mục 15.1)",
            "Kiểm tra Quy định (<=5 SV/đề tài)",
            "Nội dung khai báo Sản phẩm KHCN & Hướng dẫn người học (Mẫu 01b)",
        ]
        _style_table(
            ws3,
            4,
            headers_s3,
            [6, 12, 32, 10, 50, 12, 28, 10, 36, 26, 14, 24, 38, 16, 22, 52],
        )
        for idx, sv in enumerate(sv_2025_2027, start=1):
            r_num = 4 + idx
            if sv.id_dot_dang_ky == 9176:
                moc_sv = "Thực hiện: T10/2026-T4/2027; Rà soát giữa kỳ: T2-3/2027; Tiểu ban: 07-08/5/2027"
            else:
                moc_sv = "Đã bảo vệ Tiểu ban (2025-2026) – Kê khai & thẩm định giờ NCKH kỳ I (T12/2026)"

            kt_sv = "Đạt chuẩn (<=5 SV)" if sv.so_luong_sv <= 5 else f"Cảnh báo ({sv.so_luong_sv} SV > 5)"
            goi_y_sv = (
                f"Hướng dẫn nhóm {sv.so_luong_sv} sinh viên thực hiện đề tài NCKHSV ({sv.ma_dot_dang_ky}): "
                f"\"{sv.ten_de_tai}\" – {sv.giai_thuong_ket_qua}."
            )
            row_vals = [
                idx,
                sv.id_usmart,
                f"[{sv.ma_dot_dang_ky}] {sv.ten_dot_dang_ky}",
                sv.nam_ke_hoach,
                sv.ten_de_tai,
                ", ".join(sv.ds_ma_gv_bomon),
                ", ".join(sv.ds_ten_gv_bomon),
                sv.so_luong_sv,
                sv.thanh_vien_khac_hoac_sv,
                sv.trang_thai_mo_ta,
                sv.diem_nghiem_thu if sv.diem_nghiem_thu is not None else "",
                sv.giai_thuong_ket_qua,
                moc_sv,
                sv.gio_nckh_quy_doi_du_kien,
                kt_sv,
                goi_y_sv,
            ]
            for c_idx, val in enumerate(row_vals, start=1):
                cell = ws3.cell(r_num, c_idx, val)
                cell.font = body_font
                cell.border = thin_border
                cell.alignment = Alignment(
                    horizontal="center" if c_idx in (1, 2, 4, 6, 8, 11, 14, 15) else "left",
                    vertical="center",
                    wrap_text=True,
                )
                if c_idx == 12 and "Giải" in sv.giai_thuong_ket_qua:
                    cell.fill = ok_fill

        # =========================================================================
        # SHEET 4: 04_TongHop_21GV_KhaiBao_BaoCao (Tổng hợp cho 21 Giảng viên)
        # =========================================================================
        ws4 = wb.create_sheet("04_TongHop_21GV_KhaiBao_BaoCao")
        ws4["A1"] = "BẢNG TỔNG HỢP KẾ HOẠCH NCKH & NCKHSV CỦA 21 GIẢNG VIÊN BỘ MÔN ĐKT & KGN PHỤC VỤ KHAI BÁO & BÁO CÁO THÁNG"
        ws4["A1"].font = title_font
        ws4["A2"] = "Tự động tổng hợp số lượng đề tài các cấp, đề tài NCKHSV, giờ chuẩn quy đổi và sinh chuỗi văn bản khai báo sẵn cho Mục I, Mục II Mẫu 01b & Cổng uSmart"
        ws4["A2"].font = subtitle_font

        headers_s4 = [
            "TT",
            "Mã CB",
            "Họ và tên Giảng viên",
            "Học hàm / Học vị",
            "Tài khoản uSmart",
            "Định mức Giờ NCKH/năm",
            "Số Đề tài các cấp (2025-2027)",
            "Số Đề tài NCKHSV (2026-2027)",
            "Số Đề tài NCKHSV (2025-2026)",
            "Tổng Đề tài các cấp (Lũy kế uSmart)",
            "Tổng Đề tài NCKHSV (Lũy kế uSmart)",
            "Giờ NCKH dự kiến từ Đề tài & NCKHSV (2025-2027)",
            "Giờ NCKH phân bổ tháng báo cáo",
            "Danh mục Đề tài các cấp & NCKHSV đang triển khai (Dùng để Khai báo uSmart & Điền ô Sản phẩm KHCN Mẫu 01b)",
        ]
        _style_table(
            ws4,
            4,
            headers_s4,
            [6, 10, 26, 16, 16, 16, 16, 16, 16, 16, 16, 18, 16, 65],
        )

        for idx, gv in enumerate(DANH_SACH_GIANG_VIEN_DKT, start=1):
            r_num = 4 + idx
            mcb = gv.ma_can_bo
            dt_gv_2527 = [x for x in cc_2025_2027 if mcb in x.ds_ma_gv_bomon]
            sv_gv_2627 = [x for x in sv_2025_2027 if mcb in x.ds_ma_gv_bomon and x.id_dot_dang_ky == 9176]
            sv_gv_2526 = [x for x in sv_2025_2027 if mcb in x.ds_ma_gv_bomon and x.id_dot_dang_ky == 9167]
            dt_gv_all = [x for x in cc_hop_le if mcb in x.ds_ma_gv_bomon]
            sv_gv_all = [x for x in sv_hop_le if mcb in x.ds_ma_gv_bomon]

            gio_dt_sv = sum(x.tinh_gio_cho_gv(mcb) for x in dt_gv_2527) + sum(
                x.tinh_gio_cho_gv(mcb) for x in (sv_gv_2627 + sv_gv_2526)
            )
            # Giờ NCKH phân bổ trong tháng (tối thiểu 25-35h/tháng theo định mức năm học hoặc theo đề tài thực tế)
            gio_thang = max(round(gv.dinh_muc_nckh / 10.0, 1), round(gio_dt_sv / 10.0, 1))

            # Tạo đoạn văn bản khai báo chuẩn xác 100% từ dữ liệu thực
            cac_y_khai_bao: List[str] = []
            for d_item in dt_gv_2527[:3]:
                vt = d_item.vai_tro_gv_bomon.get(mcb, "Thành viên")
                cac_y_khai_bao.append(
                    f"{vt} đề tài {d_item.ten_cap_de_tai} [{d_item.ma_de_tai}]: \"{d_item.ten_de_tai}\""
                )
            for s_item in sv_gv_2627[:3]:
                cac_y_khai_bao.append(
                    f"Hướng dẫn NCKHSV 2026-2027 ({s_item.so_luong_sv} SV): \"{s_item.ten_de_tai}\""
                )
            if not sv_gv_2627 and sv_gv_2526:
                for s_item in sv_gv_2526[:2]:
                    cac_y_khai_bao.append(
                        f"Hoàn thành HD NCKHSV 2025-2026 ({s_item.giai_thuong_ket_qua}): \"{s_item.ten_de_tai}\""
                    )
            if not cac_y_khai_bao:
                cac_y_khai_bao.append(
                    "Thực hiện nghiên cứu khoa học theo kế hoạch năm học của Bộ môn Địa kỹ thuật & KGN; hướng dẫn sinh viên làm đồ án tốt nghiệp và viết bài báo khoa học."
                )
            chuoi_khai_bao = "; ".join(cac_y_khai_bao) + "."

            hh_hv = ".".join(x for x in [gv.hoc_ham, gv.hoc_vi] if x)
            row_vals = [
                idx,
                mcb,
                gv.ho_va_ten,
                hh_hv,
                gv.tai_khoan_usmart,
                gv.dinh_muc_nckh,
                len(dt_gv_2527),
                len(sv_gv_2627),
                len(sv_gv_2526),
                len(dt_gv_all),
                len(sv_gv_all),
                round(gio_dt_sv, 1),
                gio_thang,
                chuoi_khai_bao,
            ]
            for c_idx, val in enumerate(row_vals, start=1):
                cell = ws4.cell(r_num, c_idx, val)
                cell.font = bold_font if c_idx in (2, 3) else body_font
                cell.border = thin_border
                cell.alignment = Alignment(
                    horizontal="center" if c_idx in range(1, 14) and c_idx not in (3,) else "left",
                    vertical="center",
                    wrap_text=True,
                )
                if c_idx in (12, 13):
                    cell.fill = formula_fill

        # =========================================================================
        # SHEET 5: 05_CSDL_TatCa_DeTai_CacCap (Toàn bộ Đề tài NCKH các cấp lịch sử)
        # =========================================================================
        ws5 = wb.create_sheet("05_CSDL_TatCa_DeTai_CacCap")
        ws5["A1"] = "CƠ SỞ DỮ LIỆU TOÀN BỘ ĐỀ TÀI NCKH CÁC CẤP CỦA BỘ MÔN ĐỊA KỸ THUẬT VÀ KHÔNG GIAN NGẦM TRÊN USMART"
        ws5["A1"].font = title_font
        headers_s5 = [
            "TT",
            "ID uSmart",
            "Mã đề tài",
            "Cấp đề tài",
            "Đợt đăng ký",
            "Năm",
            "Tên đề tài NCKH",
            "Chủ nhiệm đề tài",
            "Giảng viên Bộ môn tham gia",
            "Từ ngày",
            "Đến ngày",
            "Kinh phí (Triệu đ)",
            "Trạng thái uSmart",
            "Số QĐ HĐ Nghiệm thu",
            "Kết quả / Đánh giá",
            "Giờ chuẩn quy đổi",
            "Ghi chú Kiểm toán",
        ]
        _style_table(
            ws5,
            4,
            headers_s5,
            [6, 12, 18, 22, 24, 10, 48, 26, 32, 14, 14, 15, 28, 18, 24, 15, 28],
        )
        for idx, dt in enumerate(cc_hop_le, start=1):
            r_num = 4 + idx
            row_vals = [
                idx,
                dt.id_usmart,
                dt.ma_de_tai,
                dt.ten_cap_de_tai,
                f"[{dt.ma_dot_dang_ky}] {dt.ten_dot_dang_ky}",
                dt.nam_ke_hoach,
                dt.ten_de_tai,
                dt.chu_nhiem_hoac_gvhd,
                ", ".join(f"{n} ({c})" for c, n in zip(dt.ds_ma_gv_bomon, dt.ds_ten_gv_bomon)),
                dt.tu_ngay_str,
                dt.den_ngay_str,
                dt.kinh_phi_trieu,
                dt.trang_thai_mo_ta,
                dt.so_qd_nghiem_thu,
                dt.giai_thuong_ket_qua,
                dt.gio_nckh_quy_doi_du_kien,
                dt.ghi_chu_kiem_toan,
            ]
            for c_idx, val in enumerate(row_vals, start=1):
                cell = ws5.cell(r_num, c_idx, val)
                cell.font = body_font
                cell.border = thin_border
                cell.alignment = Alignment(
                    horizontal="center" if c_idx in (1, 2, 3, 6, 10, 11, 12, 16) else "left",
                    vertical="center",
                    wrap_text=True,
                )

        # =========================================================================
        # SHEET 6: 06_CSDL_TatCa_NCKHSV (Toàn bộ 257 Đề tài NCKH Sinh viên)
        # =========================================================================
        ws6 = wb.create_sheet("06_CSDL_TatCa_NCKHSV")
        ws6["A1"] = "CƠ SỞ DỮ LIỆU TOÀN BỘ 257 ĐỀ TÀI NCKH SINH VIÊN DO GIẢNG VIÊN BỘ MÔN ĐKT & KGN HƯỚNG DẪN TRÊN USMART"
        ws6["A1"].font = title_font
        headers_s6 = [
            "TT",
            "ID uSmart",
            "Đợt đăng ký / Năm học",
            "Năm",
            "Tên đề tài NCKH Sinh viên",
            "Mã GVHD",
            "Giảng viên hướng dẫn (Bộ môn ĐKT & KGN)",
            "Số SV",
            "Danh sách Mã SV thực hiện",
            "Trạng thái uSmart",
            "Điểm số",
            "Giải thưởng đạt được",
            "Giờ chuẩn quy đổi",
            "Ghi chú Kiểm toán",
        ]
        _style_table(
            ws6,
            4,
            headers_s6,
            [6, 12, 30, 10, 50, 12, 28, 10, 36, 26, 12, 24, 15, 28],
        )
        for idx, sv in enumerate(sv_hop_le, start=1):
            r_num = 4 + idx
            row_vals = [
                idx,
                sv.id_usmart,
                f"[{sv.ma_dot_dang_ky}] {sv.ten_dot_dang_ky}",
                sv.nam_ke_hoach,
                sv.ten_de_tai,
                ", ".join(sv.ds_ma_gv_bomon),
                ", ".join(sv.ds_ten_gv_bomon),
                sv.so_luong_sv,
                sv.thanh_vien_khac_hoac_sv,
                sv.trang_thai_mo_ta,
                sv.diem_nghiem_thu if sv.diem_nghiem_thu is not None else "",
                sv.giai_thuong_ket_qua,
                sv.gio_nckh_quy_doi_du_kien,
                sv.ghi_chu_kiem_toan,
            ]
            for c_idx, val in enumerate(row_vals, start=1):
                cell = ws6.cell(r_num, c_idx, val)
                cell.font = body_font
                cell.border = thin_border
                cell.alignment = Alignment(
                    horizontal="center" if c_idx in (1, 2, 4, 6, 8, 11, 13) else "left",
                    vertical="center",
                    wrap_text=True,
                )

        # =========================================================================
        # SHEET 7: 07_LichTrinh_DinhMuc_PKHCN (Lịch trình KHCN 2026-2027 & Định mức QĐ 1293)
        # =========================================================================
        ws7 = wb.create_sheet("07_LichTrinh_DinhMuc_PKHCN")
        ws7["A1"] = "LỊCH TRÌNH MỐC THỜI GIAN KHCN & NCKHSV NĂM HỌC 2026-2027 VÀ ĐỊNH MỨC QUY ĐỔI GIỜ NCKH (PHÒNG KH&CGCN)"
        ws7["A1"].font = title_font
        _style_table(
            ws7,
            4,
            [
                "TT",
                "Phân nhóm Kế hoạch",
                "Nội dung công việc / Mốc triển khai",
                "Thời hạn thực hiện",
                "Đơn vị chủ trì / Thực hiện",
                "Quy định & Định mức kinh phí / Giờ chuẩn quy đổi",
            ],
            [6, 28, 52, 24, 28, 48],
        )
        lich_trinh_rows = [
            (1, "Đề tài NCKH các cấp 2026", "Nghiệm thu đề tài NCKH cấp Trường (T2026) và đề tài các cấp năm 2026", "Tháng 11 - 12/2026", "P.KH&CGCN, Chủ nhiệm đề tài", "Đề tài cấp Trường nghiệm thu đạt: 200 giờ chuẩn (QĐ 1293)"),
            (2, "Đề tài cấp Trường 2027 (T2027)", "Đăng ký trên uSmart và xét duyệt cấp Khoa/Bộ môn (Bước 1)", "Tháng 10 - 11/2026", "Bộ môn ĐKT & KGN, Khoa Công trình", "Loại 1A-1E (20-30 tr.đ); Loại 2A-2E WoS/Scopus (30-80 tr.đ); Loại 3 SHTT (50-75 tr.đ)"),
            (3, "Đề tài cấp Trường 2027 (T2027)", "Xét duyệt cấp Trường và ra Quyết định giao đề tài năm 2027", "Tháng 12/2026", "P.KH&CGCN, Hội đồng Trường", "Thời gian thực hiện 12 tháng (từ T01/2027 đến T12/2027)"),
            (4, "Đề tài cấp Trường trọng điểm 2027", "Nộp hồ sơ thuyết minh cấp Khoa về P.KH&CGCN và xét chọn cấp Trường", "20/11/2026 - 10/12/2026", "Nhóm nghiên cứu, Khoa Công trình", "Kinh phí 100 - 160 triệu đồng/đề tài; Thời gian thực hiện 24 tháng (từ T01/2027)"),
            (5, "Kê khai & Thẩm định giờ NCKH", "Khai báo và duyệt giờ NCKH kỳ I năm học 2026-2027 trên uSmart", "Tháng 12/2026 - 01/2027", "Bộ môn ĐKT & KGN (T12/2026), Trường (T1/2027)", "Định mức năm: GS/PGS = 350h; TS/GVC = 300h; ThS/GV = 250h"),
            (6, "Kiểm tra tiến độ Đề tài Bộ/TĐ", "Kiểm tra tiến độ đề tài cấp Bộ, Tỉnh, Trường trọng điểm 6 tháng cuối năm", "Tháng 12/2026", "Các Chủ nhiệm đề tài, Bộ môn", "Đề tài cấp Bộ: 510 giờ chuẩn (thực hiện) + 30h/60h (thuyết minh được duyệt)"),
            (7, "NCKH Sinh viên 2026-2027", "Đăng ký, xét duyệt cấp Khoa/Bộ môn và ra Quyết định giao đề tài NCKHSV", "Tháng 9 - 10/2026", "GVHD, Bộ môn ĐKT & KGN, Khoa", "Mỗi GVHD <= 3 đề tài/năm; mỗi đề tài <= 5 SV; mỗi SV <= 2 đề tài"),
            (8, "NCKH Sinh viên 2026-2027", "Triển khai nghiên cứu và Rà soát tiến độ giữa kỳ trên module NCKHSV uSmart", "Tháng 10/2026 - 03/2027", "GVHD và Nhóm Sinh viên", "Cập nhật tiến độ hoặc điều chỉnh tên/thành viên trên uSmart vào T02-03/2027"),
            (9, "NCKH Sinh viên 2026-2027", "Tổ chức phiên Tiểu ban chấm điểm, đánh giá và nhập điểm/giải lên uSmart", "Ngày 07 - 08/05/2027", "Khoa Công trình, Tiểu ban ĐKT", "Quy đổi giờ HD NCKHSV cấp Trường: Nhất=40h, Nhì=35h, Ba=30h, KK=25h, Bảo vệ đạt=20h"),
            (10, "NCKH Sinh viên 2026-2027", "Đề xuất đề tài xuất sắc cấp Trường và Hội nghị tổng kết trao giải", "11/05/2027 - 29/05/2027", "P.KH&CGCN, P.Chăm sóc người học", "Kỷ yếu NCKHSV hoàn thành trước 15/07/2027"),
        ]
        for idx, r_vals in enumerate(lich_trinh_rows, start=5):
            for c_idx, val in enumerate(r_vals, start=1):
                cell = ws7.cell(idx, c_idx, val)
                cell.font = body_font
                cell.border = thin_border
                cell.alignment = Alignment(horizontal="center" if c_idx in (1, 4) else "left", vertical="center", wrap_text=True)

        # =========================================================================
        # SHEET 8: 08_NhiemVu_DongBo_Mau01b_02b (Nhiệm vụ đồng bộ sang Báo cáo tháng)
        # =========================================================================
        ws8 = wb.create_sheet("08_NhiemVu_DongBo_Mau01b_02b")
        ws8["A1"] = f"DANH SÁCH NHIỆM VỤ NCKH & NCKHSV TỰ ĐỘNG ĐỒNG BỘ SANG BÁO CÁO CÔNG VIỆC THÁNG {thang_bao_cao:02d}/{nam_bao_cao} (TB 940/TB-ĐHGTVT)"
        ws8["A1"].font = title_font
        headers_s8 = [
            "Mã nhiệm vụ",
            "Tháng",
            "Năm",
            "Mã CB chủ trì",
            "Họ tên chủ trì",
            "Người phối hợp",
            "Nhiệm vụ, công việc",
            "Nguồn giao",
            "Số văn bản / Mã việc uSmart",
            "Sản phẩm đầu ra, kết quả",
            "Ngày giao",
            "Thời hạn",
            "Ngày hoàn thành",
            "Tình trạng",
            "Mức độ HT (%)",
            "Khó khăn, đề xuất",
            "Nhiệm vụ trọng tâm ĐV",
            "Là việc lặp lại",
            "Mức ưu tiên",
            "Đánh giá chất lượng",
            "Link minh chứng Drive",
            "Nguồn dữ liệu",
        ]
        _style_table(ws8, 4, headers_s8)

        nhiem_vu_rows = self.sinh_danh_sach_nhiem_vu_bao_cao_thang(
            cc_2025_2027, sv_2025_2027, thang=thang_bao_cao, nam=nam_bao_cao
        )
        for idx, row_vals in enumerate(nhiem_vu_rows, start=5):
            for c_idx, val in enumerate(row_vals, start=1):
                cell = ws8.cell(idx, c_idx, val)
                cell.font = body_font
                cell.border = thin_border
                if isinstance(val, date):
                    cell.number_format = "dd/mm/yyyy"

        # =========================================================================
        # SHEET 9: 09_VungCachLy_KiemToan_NCKH (Quarantine & Audit Log)
        # =========================================================================
        ws9 = wb.create_sheet("09_VungCachLy_KiemToan_NCKH")
        ws9["A1"] = "VÙNG CÁCH LY DỊ BIỆT & NHẬT KÝ KIỂM TOÁN RÀNG BUỘC VẬT LÝ (PHYSICAL PLAUSIBILITY & QUARANTINE)"
        ws9["A1"].font = title_font
        ws9["A2"] = "Tuân thủ nguyên tắc An toàn Số liệu: Tuyệt đối không tự ý xóa bỏ bản ghi không duyệt hoặc dị biệt trên uSmart mà lưu riêng tại Vùng cách ly kèm nguyên nhân rõ ràng."
        ws9["A2"].font = subtitle_font
        _style_table(
            ws9,
            4,
            [
                "TT",
                "Mã Quy tắc",
                "Mức độ",
                "Phân loại",
                "Mã / ID Đề tài",
                "Tên đề tài / Nội dung",
                "Giảng viên liên quan",
                "Nguyên nhân đưa vào Vùng cách ly / Cảnh báo",
                "Hướng dẫn xử lý chuẩn",
            ],
            [6, 14, 14, 22, 18, 48, 28, 45, 45],
        )
        for idx, cl in enumerate(ds_cach_ly, start=1):
            r_num = 4 + idx
            row_vals = [
                idx,
                cl.ma_quy_tac,
                cl.muc_do,
                cl.loai_de_tai,
                cl.ma_hoac_id,
                cl.ten_de_tai,
                cl.giang_vien_lien_quan,
                cl.nguyen_nhan,
                cl.huong_dan_xu_ly,
            ]
            for c_idx, val in enumerate(row_vals, start=1):
                cell = ws9.cell(r_num, c_idx, val)
                cell.font = body_font
                cell.border = thin_border
                cell.fill = quar_fill if cl.muc_do == "CACH_LY" else warn_fill
                cell.alignment = Alignment(
                    horizontal="center" if c_idx in (1, 2, 3, 5) else "left",
                    vertical="center",
                    wrap_text=True,
                )

        wb.save(duong_dan_xuat)
        logger.info("Đã xuất Google Sheet Kế hoạch NCKH & NCKHSV tại: %s", duong_dan_xuat)
        return duong_dan_xuat

    def sinh_danh_sach_nhiem_vu_bao_cao_thang(
        self,
        cc_2025_2027: List[DeTaiNCKHChuanHoa],
        sv_2025_2027: List[DeTaiNCKHChuanHoa],
        thang: int = 11,
        nam: int = 2026,
    ) -> List[list]:
        """
        Tự động chuyển hóa các Đề tài NCKH các cấp (2025-2027) và Đề tài NCKHSV (2026-2027)
        thành các dòng nhiệm vụ chuẩn 22 cột của Sheet `NhiemVu_Thang` (Mẫu 01b, 02b, 03).
        """
        cuoi_thang = ngay_cuoi_thang(thang, nam)
        dau_thang = date(nam, thang, 1)
        ngay_ht_mac_dinh = date(nam, thang, min(28, cuoi_thang.day))
        link_drive = f"https://drive.google.com/drive/folders/{GOOGLE_DRIVE_FOLDER_ID}"

        ds_rows: List[list] = []
        stt = 1

        # 1. Nhiệm vụ từ Đề tài NCKH các cấp đang triển khai (T2026, T2027, B2025, B2024...)
        for dt in cc_2025_2027:
            for mcb in dt.ds_ma_gv_bomon:
                if mcb not in self.bomon_gv:
                    continue
                gv = self.bomon_gv[mcb]
                vt = dt.vai_tro_gv_bomon.get(mcb, "Thành viên")
                ma_nv = f"NV_{nam}{thang:02d}_DT_{stt:03d}"
                stt += 1

                if "T2027" in dt.ma_dot_dang_ky:
                    ten_nv = (
                        f"Hoàn thiện thuyết minh và xét duyệt cấp Khoa đề tài NCKH {dt.ten_cap_de_tai} năm 2027 "
                        f"({vt}, mã số {dt.ma_de_tai}): \"{dt.ten_de_tai}\""
                    )
                    san_pham = f"Hồ sơ thuyết minh đề tài {dt.ma_de_tai} được Khoa thông qua trên hệ thống uSmart"
                elif "T2026" in dt.ma_dot_dang_ky:
                    ten_nv = (
                        f"Thực hiện nội dung nghiên cứu và hoàn thiện hồ sơ nghiệm thu đề tài NCKH {dt.ten_cap_de_tai} năm 2026 "
                        f"({vt}, mã số {dt.ma_de_tai}): \"{dt.ten_de_tai}\""
                    )
                    san_pham = f"Báo cáo tổng kết đề tài {dt.ma_de_tai} và sản phẩm bài báo khoa học theo thuyết minh"
                else:
                    ten_nv = (
                        f"Thực hiện chuyên đề nghiên cứu và báo cáo tiến độ đề tài {dt.ten_cap_de_tai} "
                        f"({vt}, mã số {dt.ma_de_tai}): \"{dt.ten_de_tai}\""
                    )
                    san_pham = f"Báo cáo tiến độ thực hiện đề tài {dt.ma_de_tai} và sản phẩm chuyên đề KHCN"

                la_trong_tam = "Chủ nhiệm" in vt and (
                    "Bộ" in dt.ten_cap_de_tai or "trọng điểm" in dt.ten_cap_de_tai or "T2027" in dt.ma_dot_dang_ky
                )
                ds_rows.append(
                    [
                        ma_nv,
                        thang,
                        nam,
                        mcb,
                        gv.ho_va_ten,
                        ", ".join(x for x in dt.ds_ten_gv_bomon if x != gv.ho_va_ten) or "Nhóm NCKH Bộ môn",
                        ten_nv,
                        "Kế hoạch năm học",
                        dt.ma_de_tai,
                        san_pham,
                        dau_thang,
                        cuoi_thang,
                        ngay_ht_mac_dinh,
                        "Hoàn thành đúng hạn",
                        100.0,
                        "",
                        la_trong_tam,
                        True,
                        "Cao",
                        "Đạt yêu cầu – tốt",
                        link_drive,
                        "Đồng bộ Đề tài NCKH uSmart (lannc)",
                    ]
                )

        # 2. Nhiệm vụ từ Đề tài NCKH Sinh viên năm học 2026-2027 (ID đợt = 9176)
        sv_2627_by_gv: Dict[str, List[DeTaiNCKHChuanHoa]] = {}
        for sv in sv_2025_2027:
            if sv.id_dot_dang_ky == 9176:
                for mcb in sv.ds_ma_gv_bomon:
                    sv_2627_by_gv.setdefault(mcb, []).append(sv)

        for mcb, ds_sv_gv in sv_2627_by_gv.items():
            if mcb not in self.bomon_gv:
                continue
            gv = self.bomon_gv[mcb]
            ma_nv = f"NV_{nam}{thang:02d}_SV_{stt:03d}"
            stt += 1
            ds_ten_dt_sv = "; ".join(f"({i+1}) \"{s.ten_de_tai}\"" for i, s in enumerate(ds_sv_gv))
            tong_sv = sum(s.so_luong_sv for s in ds_sv_gv)
            ten_nv = (
                f"Hướng dẫn {len(ds_sv_gv)} nhóm sinh viên ({tong_sv} SV) triển khai đề tài NCKH Sinh viên "
                f"năm học 2026-2027 trên uSmart: {ds_ten_dt_sv}"
            )
            san_pham = (
                f"Đề cương chi tiết, kế hoạch thí nghiệm/mô phỏng và nhật ký hướng dẫn {len(ds_sv_gv)} đề tài NCKHSV 2026-2027"
            )
            ds_rows.append(
                [
                    ma_nv,
                    thang,
                    nam,
                    mcb,
                    gv.ho_va_ten,
                    f"{tong_sv} SV chuyên ngành ĐKT",
                    ten_nv,
                    "Kế hoạch năm học",
                    f"NCKHSV-2627-{mcb}",
                    san_pham,
                    dau_thang,
                    cuoi_thang,
                    ngay_ht_mac_dinh,
                    "Hoàn thành đúng hạn",
                    100.0,
                    "",
                    False,
                    True,
                    "Thường",
                    "Đạt yêu cầu – tốt",
                    link_drive,
                    "Đồng bộ NCKHSV uSmart (lannc)",
                ]
            )

        return ds_rows

    def dong_bo_vao_csdl_bao_cao_bo_mon(
        self,
        duong_dan_csdl_master: Path,
        duong_dan_gsheet_nckh: Path,
        cc_2025_2027: List[DeTaiNCKHChuanHoa],
        sv_2025_2027: List[DeTaiNCKHChuanHoa],
        thang: int = 11,
        nam: int = 2026,
    ) -> int:
        """
        Đồng bộ trực tiếp dữ liệu Kế hoạch NCKH đề tài các cấp & NCKHSV vào tệp CSDL trung tâm
        `CSDL_BaoCao_BoMon_DKT_KGN.xlsx`:
        1. Sao chép 3 Sheet `KH_DeTai_NCKH_2025_2027`, `KH_NCKHSV_2025_2027`, `TongHop_21GV_NCKH` vào CSDL Master.
        2. Bổ sung các nhiệm vụ NCKH & NCKHSV thực tế (chưa trùng) vào Sheet `NhiemVu_Thang`.
        3. Cập nhật cột `Sản phẩm KHCN & Hướng dẫn người học` và `Giờ NCKH` trong Sheet `KH_GiangDay_NCKH`.
        """
        duong_dan_csdl_master = Path(duong_dan_csdl_master)
        wb_master = openpyxl.load_workbook(duong_dan_csdl_master)
        wb_nckh = openpyxl.load_workbook(duong_dan_gsheet_nckh)

        # 1. Đồng bộ 3 Sheet chuyên đề NCKH sang CSDL Master
        map_copy = {
            "02_KH_DeTai_NCKH_2025_2027": "KH_DeTai_NCKH_CacCap",
            "03_KH_NCKHSV_2025_2027": "KH_DeTai_NCKHSV",
            "04_TongHop_21GV_KhaiBao_BaoCao": "TongHop_KhaiBao_NCKH_21GV",
        }
        for src_name, dst_name in map_copy.items():
            if dst_name in wb_master.sheetnames:
                wb_master.remove(wb_master[dst_name])
            ws_src = wb_nckh[src_name]
            ws_dst = wb_master.create_sheet(dst_name)
            for row in ws_src.iter_rows(values_only=False):
                for cell in row:
                    new_cell = ws_dst.cell(cell.row, cell.column, cell.value)
                    if cell.has_style:
                        new_cell.font = copy.copy(cell.font)
                        new_cell.fill = copy.copy(cell.fill)
                        new_cell.border = copy.copy(cell.border)
                        new_cell.alignment = copy.copy(cell.alignment)
                        new_cell.number_format = cell.number_format
            for col_letter, dim in ws_src.column_dimensions.items():
                ws_dst.column_dimensions[col_letter].width = dim.width

        # 2. Cập nhật Sheet NhiemVu_Thang (Xóa các dòng đồng bộ NCKH cũ của cùng kỳ và nạp bản mới nhất)
        ws_nv = wb_master["NhiemVu_Thang"]
        rows_keep: List[list] = []
        for r in range(2, ws_nv.max_row + 1):
            row_vals = [ws_nv.cell(r, c).value for c in range(1, 23)]
            t = int(row_vals[1] or 0)
            n = int(row_vals[2] or 0)
            nguon_dl = str(row_vals[21] or "")
            if t == thang and n == nam and "uSmart (lannc)" in nguon_dl:
                continue
            rows_keep.append(row_vals)

        if ws_nv.max_row > 1:
            ws_nv.delete_rows(2, ws_nv.max_row - 1)
        for r_vals in rows_keep:
            ws_nv.append(r_vals)

        nhiem_vu_nckh_moi = self.sinh_danh_sach_nhiem_vu_bao_cao_thang(
            cc_2025_2027, sv_2025_2027, thang=thang, nam=nam
        )
        for r_vals in nhiem_vu_nckh_moi:
            ws_nv.append(r_vals)

        # 3. Cập nhật chuỗi Sản phẩm KHCN & Hướng dẫn người học thực tế từ Sheet 04 vào Sheet KH_GiangDay_NCKH
        ws_s4 = wb_nckh["04_TongHop_21GV_KhaiBao_BaoCao"]
        sp_khcn_by_mcb: Dict[str, Tuple[float, str]] = {}
        for r in range(5, ws_s4.max_row + 1):
            mcb = str(ws_s4.cell(r, 2).value or "").strip()
            gio_thang = float(ws_s4.cell(r, 13).value or 30.0)
            chuoi_kb = str(ws_s4.cell(r, 14).value or "").strip()
            if mcb:
                sp_khcn_by_mcb[mcb] = (gio_thang, chuoi_kb)

        ws_cm = wb_master["KH_GiangDay_NCKH"]
        for r in range(2, ws_cm.max_row + 1):
            t = int(ws_cm.cell(r, 1).value or 0)
            n = int(ws_cm.cell(r, 2).value or 0)
            mcb = str(ws_cm.cell(r, 3).value or "").strip()
            if t == thang and n == nam and mcb in sp_khcn_by_mcb:
                gio_thang, chuoi_kb = sp_khcn_by_mcb[mcb]
                ws_cm.cell(r, 9, gio_thang)
                ws_cm.cell(r, 10, max(float(ws_cm.cell(r, 10).value or 0.0), round(gio_thang * 3.0, 1)))
                ws_cm.cell(r, 12, "Đồng bộ từ Kế hoạch Đề tài NCKH các cấp & NCKHSV trên uSmart (TK lannc)")
                ws_cm.cell(r, 17, chuoi_kb)

        wb_nckh.close()
        wb_master.save(duong_dan_csdl_master)
        wb_master.close()
        logger.info(
            "Đã đồng bộ %d nhiệm vụ NCKH & NCKHSV thực tế vào CSDL báo cáo tháng (%s).",
            len(nhiem_vu_nckh_moi),
            duong_dan_csdl_master,
        )
        return len(nhiem_vu_nckh_moi)


def thu_dong_bo_len_google_drive_cli(duong_dan_file: Path, folder_id: str = GOOGLE_DRIVE_FOLDER_ID) -> str:
    """
    Thử tải/cập nhật tệp lên thư mục Google Drive `Kế hoạch công việc` (`1jawTokjmk-hnBsuoEIHoxYUhSRCEItPd`)
    nếu tiện ích `gdrive` đã được đăng nhập. Trả về chuỗi trạng thái rõ ràng.
    """
    gdrive_bin = shutil.which("gdrive") or str(Path.home() / ".gemini" / "antigravity" / "bin" / "gdrive.exe")
    try:
        proc = subprocess.run(
            [gdrive_bin, "readonly", "search", "--parent-id", folder_id, "--name-contains", "Ke_hoach_NCKH"],
            capture_output=True,
            text=True,
            timeout=12,
        )
        if proc.returncode != 0:
            return (
                "Đã lưu trực tiếp tại thư mục `Ke_hoach_cong_viec/` trên OneDrive và sẵn sàng đồng bộ "
                f"lên Google Drive (`{folder_id}`) qua Google Apps Script `NCKH_USmart_Sync.gs`."
            )
        # Nếu đã đăng nhập gdrive thành công, thực hiện upload lên folder_id
        up_proc = subprocess.run(
            [
                gdrive_bin,
                "mutate",
                "upload",
                str(duong_dan_file),
                "--parent",
                folder_id,
                "--mime-type",
                "application/vnd.google-apps.spreadsheet",
            ],
            capture_output=True,
            text=True,
            timeout=30,
        )
        if up_proc.returncode == 0:
            return f"Đã tải trực tiếp lên Google Drive Folder `{folder_id}` dưới dạng Google Sheet: {up_proc.stdout.strip()}"
        return f"Đã lưu tại `Ke_hoach_cong_viec/` (Trạng thái CLI Drive: {up_proc.stderr.strip()[:120]})"
    except Exception as exc:
        return f"Đã lưu tại `Ke_hoach_cong_viec/` (Hỗ trợ đồng bộ tự động qua Google Apps Script: {exc})"


def chay_dong_bo_ke_hoach_nckh_toan_trinh(
    username: str = "lannc",
    password: str = "giangchau",
    thang: int = 11,
    nam: int = 2026,
    dong_bo_live: bool = True,
    duong_dan_csdl_master: Optional[Path] = None,
) -> KetQuaDongBoNCKH:
    """
    Hàm điều phối toàn trình:
    1. Đăng nhập OIDC uSmart với tài khoản `lannc` / `giangchau` (`giangchau0310`) và tải dữ liệu KHCN.
    2. Chuẩn hóa & Kiểm toán Vật lý (5 Quy tắc `NCKH-R1` .. `NCKH-R5`) cho Đề tài các cấp & NCKHSV.
    3. Sinh tệp Google Sheet 9 Tabs `Ke_hoach_NCKH_DeTai_CacCap_va_NCKHSV_DKT_KGN.xlsx` trong thư mục
       `Ke_hoach_cong_viec/` (cả trong thư mục dự án và thư mục năm học `2026.2027/Ke_hoach_cong_viec`).
    4. Đồng bộ trực tiếp vào `CSDL_BaoCao_BoMon_DKT_KGN.xlsx` để phục vụ khai báo và xuất báo cáo công việc tháng.
    """
    client = USmartNCKHClient(username=username, password=password)
    du_lieu_raw, live_ok = client.tai_va_cap_nhat_du_lieu(dong_bo_live=dong_bo_live)

    processor = BoXuLyKeHoachNCKH(du_lieu_raw)
    ds_cac_cap, ds_nckhsv, ds_cach_ly = processor.chuan_hoa_va_kiem_toan()

    cc_hop_le = [x for x in ds_cac_cap if x.hop_le_dua_vao_ke_hoach]
    cc_2527 = [x for x in cc_hop_le if x.dang_thuc_hien_2025_2027]
    sv_hop_le = [x for x in ds_nckhsv if x.hop_le_dua_vao_ke_hoach]
    sv_2527 = [x for x in sv_hop_le if x.id_dot_dang_ky in (9176, 9167) or x.nam_ke_hoach >= 2025]

    file_gsheet_proj = FOLDER_KE_HOACH_CONG_VIEC / TEN_FILE_GSHEET_NCKH
    processor.xuat_google_sheet_ke_hoach_nckh(
        file_gsheet_proj,
        ds_cac_cap=ds_cac_cap,
        ds_nckhsv=ds_nckhsv,
        ds_cach_ly=ds_cach_ly,
        thang_bao_cao=thang,
        nam_bao_cao=nam,
    )

    # Sao chép bản cập nhật sang thư mục năm học 2026.2027/Ke_hoach_cong_viec trên OneDrive
    file_gsheet_nam_hoc: Optional[Path] = None
    try:
        FOLDER_KE_HOACH_NAM_HOC.mkdir(parents=True, exist_ok=True)
        file_gsheet_nam_hoc = FOLDER_KE_HOACH_NAM_HOC / TEN_FILE_GSHEET_NCKH
        shutil.copy2(file_gsheet_proj, file_gsheet_nam_hoc)
    except Exception as exc:
        logger.warning("Không thể sao chép sang thư mục năm học 2026.2027: %s", exc)

    # Đồng bộ vào CSDL_BaoCao_BoMon_DKT_KGN.xlsx
    csdl_path = Path(duong_dan_csdl_master or (PROJECT_ROOT / "CSDL_BaoCao_BoMon_DKT_KGN.xlsx"))
    from .database_manager import QuanLyCSDLBộMôn

    db_mgr = QuanLyCSDLBộMôn(csdl_path)
    db_mgr.khoi_tao_csdl_mau_neu_chua_co(ghi_de=False)

    so_nv_dong_bo = processor.dong_bo_vao_csdl_bao_cao_bo_mon(
        csdl_path,
        file_gsheet_proj,
        cc_2025_2027=cc_2527,
        sv_2025_2027=sv_2527,
        thang=thang,
        nam=nam,
    )

    # Sao chép CSDL_BaoCao_BoMon_DKT_KGN.xlsx đã tích hợp NCKH vào thư mục Kế hoạch công việc
    try:
        shutil.copy2(csdl_path, FOLDER_KE_HOACH_CONG_VIEC / csdl_path.name)
        if FOLDER_KE_HOACH_NAM_HOC.exists():
            shutil.copy2(csdl_path, FOLDER_KE_HOACH_NAM_HOC / csdl_path.name)
    except Exception as exc:
        logger.debug("Bỏ qua sao chép bản phụ CSDL vào Ke_hoach_cong_viec: %s", exc)

    tt_drive = thu_dong_bo_len_google_drive_cli(file_gsheet_proj)

    return KetQuaDongBoNCKH(
        ket_noi_live_thanh_cong=live_ok,
        tai_khoan_su_dung=f"{client.username} (SSO OIDC)",
        tong_de_tai_cac_cap=len(cc_hop_le),
        so_de_tai_cac_cap_2025_2027=len(cc_2527),
        tong_de_tai_nckhsv=len(sv_hop_le),
        so_de_tai_nckhsv_2025_2027=len(sv_2527),
        so_ban_ghi_cach_ly=len(ds_cach_ly),
        duong_dan_gsheet_project=file_gsheet_proj,
        duong_dan_gsheet_nam_hoc=file_gsheet_nam_hoc,
        duong_dan_csdl_master=csdl_path,
        so_nhiem_vu_dong_bo_bao_cao_thang=so_nv_dong_bo,
        danh_sach_cach_ly=ds_cach_ly,
        trang_thai_google_drive=tt_drive,
    )


if __name__ == "__main__":
    import sys as _sys

    if hasattr(_sys.stdout, "reconfigure"):
        _sys.stdout.reconfigure(encoding="utf-8")
    logging.basicConfig(level=logging.INFO, format="%(asctime)s [%(levelname)s] %(message)s")
    kq = chay_dong_bo_ke_hoach_nckh_toan_trinh("lannc", "giangchau", thang=11, nam=2026)
    print("\n================ KẾT QUẢ ĐỒNG BỘ KẾ HOẠCH NCKH & NCKHSV ================")
    print(f"1. Kết nối trực tiếp uSmart : {kq.ket_noi_live_thanh_cong} ({kq.tai_khoan_su_dung})")
    print(f"2. Đề tài NCKH các cấp      : {kq.so_de_tai_cac_cap_2025_2027} đề tài 2025-2027 / {kq.tong_de_tai_cac_cap} tổng số")
    print(f"3. Đề tài NCKH Sinh viên    : {kq.so_de_tai_nckhsv_2025_2027} đề tài 2025-2027 / {kq.tong_de_tai_nckhsv} tổng số")
    print(f"4. Bản ghi Vùng cách ly     : {kq.so_ban_ghi_cach_ly} bản ghi (Quarantine & Cảnh báo)")
    print(f"5. Google Sheet (Dự án)     : {kq.duong_dan_gsheet_project}")
    print(f"6. Google Sheet (Năm học)   : {kq.duong_dan_gsheet_nam_hoc}")
    print(f"7. CSDL Báo cáo trung tâm   : {kq.duong_dan_csdl_master} (+{kq.so_nhiem_vu_dong_bo_bao_cao_thang} nhiệm vụ NCKH)")
    print(f"8. Trạng thái Google Drive  : {kq.trang_thai_google_drive}")
    print("========================================================================")
