# -*- coding: utf-8 -*-
"""
Module định nghĩa các mô hình dữ liệu (Data Models), danh mục chuẩn và danh sách
21 giảng viên của Bộ môn Địa kỹ thuật và Không gian ngầm (Mã đơn vị uSmart: 412).

Tuân thủ tiêu chuẩn PEP 8, Type Hints đầy đủ và Nguyên tắc An toàn Số liệu.
"""

from dataclasses import dataclass, field
from datetime import date, datetime
import calendar
import unicodedata
import re
from typing import Dict, List, Optional, Tuple


# Danh mục chuẩn theo đúng Sheet "Danh mục" của Mẫu 01, Mẫu 02, Mẫu 03 (TB 940/TB-ĐHGTVT)
DANH_MUC_NGUON_GIAO: Tuple[str, ...] = (
    "Kế hoạch năm học",
    "Kết luận giao ban",
    "Giao việc trên Usmart",
    "Chức năng, nhiệm vụ thường xuyên",
    "Nhiệm vụ đột xuất",
)

DANH_MUC_TINH_TRANG: Tuple[str, ...] = (
    "Hoàn thành đúng hạn",
    "Hoàn thành chậm",
    "Đang thực hiện – đúng tiến độ",
    "Đang thực hiện – chậm tiến độ",
    "Chưa thực hiện",
)

DANH_MUC_XEP_LOAI: Tuple[str, ...] = (
    "Hoàn thành xuất sắc nhiệm vụ",
    "Hoàn thành tốt nhiệm vụ",
    "Hoàn thành nhiệm vụ",
    "Không hoàn thành nhiệm vụ",
)

DANH_MUC_UU_TIEN: Tuple[str, ...] = (
    "Khẩn",
    "Cao",
    "Thường",
)

DANH_MUC_CHAT_LUONG: Tuple[str, ...] = (
    "Đạt yêu cầu – tốt",
    "Đạt yêu cầu",
    "Chưa đạt – làm lại",
)

TEN_DON_VI_CHUAN: str = "Bộ môn Địa kỹ thuật và Không gian ngầm"
MA_VIET_TAT_DON_VI: str = "DKT"
NGUOI_NHAN_BAO_CAO_MAC_DINH: str = "Trần Thế Truyền"
LANH_DAO_PHU_TRACH_DON_VI_MAC_DINH: str = "Trưởng khoa Công trình"


def bo_dau_tieng_viet(chuoi: str) -> str:
    """Chuyển chuỗi tiếng Việt có dấu thành không dấu, ghép liền kiểu PascalCase cho tên tệp."""
    if not chuoi:
        return ""
    s = chuoi.replace("đ", "d").replace("Đ", "D")
    s = unicodedata.normalize("NFD", s)
    s = "".join(c for c in s if unicodedata.category(c) != "Mn")
    cac_tu = re.findall(r"[A-Za-z0-9]+", s)
    return "".join(tu.capitalize() for tu in cac_tu)


def ngay_cuoi_thang(thang: int, nam: int) -> date:
    """Trả về ngày cuối cùng của tháng/năm tương ứng công thức =DATE(E5,C5+1,0)."""
    _, so_ngay = calendar.monthrange(nam, thang)
    return date(nam, thang, so_ngay)


@dataclass
class ThongTinGiangVien:
    """Thông tin hồ sơ và định mức chuyên môn của giảng viên Bộ môn ĐKT & KGN."""

    ma_can_bo: str
    ho_va_ten: str
    ngay_sinh: str
    hoc_ham: str
    hoc_vi: str
    chuc_danh_vi_tri: str = "Giảng viên"
    tai_khoan_usmart: str = ""
    email: str = ""
    nguoi_nhan_bao_cao: str = NGUOI_NHAN_BAO_CAO_MAC_DINH
    dinh_muc_giang_day: float = 270.0
    dinh_muc_nckh: float = 300.0
    dinh_muc_nhiem_vu_khac: float = 180.0

    @property
    def ten_viet_tat_khong_dau(self) -> str:
        """Trả về họ tên viết liền không dấu dùng để đặt tên tệp theo TB 940."""
        return bo_dau_tieng_viet(self.ho_va_ten)

    @property
    def chuc_danh_day_du(self) -> str:
        """Trả về chuỗi vị trí việc làm kèm học hàm, học vị."""
        tien_to = ".".join(x for x in [self.hoc_ham, self.hoc_vi] if x)
        if tien_to:
            return f"{self.chuc_danh_vi_tri} ({tien_to})"
        return self.chuc_danh_vi_tri


# Danh sách chính thức 22 cán bộ/giảng viên Bộ môn Địa kỹ thuật và Không gian ngầm & PTN Địa kỹ thuật (uSmart DonVi=412 & 494)
DANH_SACH_GIANG_VIEN_DKT: List[ThongTinGiangVien] = [
    ThongTinGiangVien("0694", "Trần Thế Truyền", "1978-07-27", "GS", "TS", "Trưởng bộ môn, Giảng viên cao cấp", "truyentt", "truyentt@utc.edu.vn", "Trưởng khoa Công trình", 270.0, 350.0, 180.0),
    ThongTinGiangVien("0882", "Nguyễn Châu Lân", "1981-06-26", "PGS", "TS", "Phó Trưởng bộ môn, Giảng viên cao cấp", "lannc", "lannc@utc.edu.vn", NGUOI_NHAN_BAO_CAO_MAC_DINH, 270.0, 350.0, 180.0),
    ThongTinGiangVien("0069", "Nguyễn Đức Mạnh", "1974-03-08", "PGS", "TS", "Giảng viên cao cấp", "manhnd", "manhnd@utc.edu.vn", NGUOI_NHAN_BAO_CAO_MAC_DINH, 270.0, 350.0, 180.0),
    ThongTinGiangVien("0935", "Phí Hồng Thịnh", "1974-01-28", "PGS", "TS", "Giảng viên cao cấp", "thinhph", "thinhph@utc.edu.vn", NGUOI_NHAN_BAO_CAO_MAC_DINH, 270.0, 350.0, 180.0),
    ThongTinGiangVien("0805", "Trần Thị Thu Hằng", "1980-09-02", "PGS", "TS", "Giảng viên cao cấp", "hangttt", "hangttt@utc.edu.vn", NGUOI_NHAN_BAO_CAO_MAC_DINH, 270.0, 350.0, 180.0),
    ThongTinGiangVien("1056", "Đặng Hồng Lam", "1981-05-23", "PGS", "TS", "Giảng viên cao cấp", "lamdh", "lamdh@utc.edu.vn", NGUOI_NHAN_BAO_CAO_MAC_DINH, 270.0, 350.0, 180.0),
    ThongTinGiangVien("1158", "Nguyễn Công Định", "1983-10-26", "", "TS", "Giảng viên chính", "dinhnc", "dinhnc@utc.edu.vn", NGUOI_NHAN_BAO_CAO_MAC_DINH, 270.0, 300.0, 180.0),
    ThongTinGiangVien("1472", "Nguyễn Thái Linh", "1990-03-30", "", "TS", "Giảng viên", "linhnt2", "thailinhdkt@utc.edu.vn", NGUOI_NHAN_BAO_CAO_MAC_DINH, 270.0, 300.0, 180.0),
    ThongTinGiangVien("1215", "Nguyễn Đức Vinh", "1984-07-27", "", "TS", "Giảng viên", "vinhnd", "vinhnd@utc.edu.vn", NGUOI_NHAN_BAO_CAO_MAC_DINH, 270.0, 300.0, 180.0),
    ThongTinGiangVien("1154", "Nguyễn Anh Tuấn", "1983-11-02", "", "TS", "Giảng viên", "tuanna", "tuanna@utc.edu.vn", NGUOI_NHAN_BAO_CAO_MAC_DINH, 270.0, 300.0, 180.0),
    ThongTinGiangVien("0068", "Nguyễn Đức Hạnh", "1974-01-21", "", "TS", "Giảng viên chính", "hanhnd", "hanhnd@utc.edu.vn", NGUOI_NHAN_BAO_CAO_MAC_DINH, 270.0, 300.0, 180.0),
    ThongTinGiangVien("0844", "Nguyễn Thạch Bích", "1978-06-20", "", "TS", "Giảng viên chính", "bichnt", "bichnt@utc.edu.vn", NGUOI_NHAN_BAO_CAO_MAC_DINH, 270.0, 300.0, 180.0),
    ThongTinGiangVien("1051", "Nguyễn Hải Hà", "1983-04-28", "", "ThS", "Giảng viên", "hanh1", "haihadkt@utc.edu.vn", NGUOI_NHAN_BAO_CAO_MAC_DINH, 270.0, 250.0, 180.0),
    ThongTinGiangVien("1216", "Trần Thị Thu Hiền", "1984-10-20", "", "ThS", "Giảng viên", "hienttt", "hienttt@utc.edu.vn", NGUOI_NHAN_BAO_CAO_MAC_DINH, 270.0, 250.0, 180.0),
    ThongTinGiangVien("0745", "Lê Thị Hồng Vân", "1979-01-30", "", "ThS", "Giảng viên", "vanlth", "vanlth@utc.edu.vn", NGUOI_NHAN_BAO_CAO_MAC_DINH, 270.0, 250.0, 180.0),
    ThongTinGiangVien("0074", "Võ Thị Hồng Minh", "1969-10-11", "", "ThS", "Giảng viên chính", "minhvth", "minhvth@utc.edu.vn", NGUOI_NHAN_BAO_CAO_MAC_DINH, 270.0, 250.0, 180.0),
    ThongTinGiangVien("0064", "Lê Anh Đức", "1974-12-06", "", "ThS", "Giảng viên", "ducla", "laduc@utc.edu.vn", NGUOI_NHAN_BAO_CAO_MAC_DINH, 270.0, 250.0, 180.0),
    ThongTinGiangVien("0837", "Nguyễn Bá Đồng", "1979-04-30", "", "ThS", "Giảng viên", "dongnb", "dongnb@utc.edu.vn", NGUOI_NHAN_BAO_CAO_MAC_DINH, 270.0, 250.0, 180.0),
    ThongTinGiangVien("0072", "Phạm Thanh Hà", "1972-02-10", "", "ThS", "Giảng viên", "hapt", "hapt@utc.edu.vn", NGUOI_NHAN_BAO_CAO_MAC_DINH, 270.0, 250.0, 180.0),
    ThongTinGiangVien("0067", "Nguyễn Đình Dũng", "1973-12-19", "", "ThS", "Giảng viên", "dungnd", "dungnd@utc.edu.vn", NGUOI_NHAN_BAO_CAO_MAC_DINH, 270.0, 250.0, 180.0),
    ThongTinGiangVien("1209", "Phạm Thu Trang", "1982-10-09", "", "ThS", "Giảng viên", "trangpt", "trangpt@utc.edu.vn", NGUOI_NHAN_BAO_CAO_MAC_DINH, 270.0, 250.0, 180.0),
    ThongTinGiangVien("0537", "Nguyễn Thị Thanh Giang", "1980-01-01", "", "ThS", "Giảng viên thực hành, thí nghiệm", "giangntt", "nttgiang@utc.edu.vn", NGUOI_NHAN_BAO_CAO_MAC_DINH, 210.0, 75.0, 180.0),
]

# Từ điển mật khẩu uSmart đã cấu hình sẵn của các thành viên Bộ môn
MAT_KHAU_USMART_MAC_DINH: dict[str, str] = {
    "giangntt": "123456",
    "ducla": "123456",
    "linhnt2": "30032013",
    "linhnt": "30032013",
    "hanh1": "182128",
    "lannc": "giangchau0310",
}


def lay_mat_khau_usmart_mac_dinh(tai_khoan: str) -> str:
    """Trả về mật khẩu uSmart mặc định đã cấu hình cho tài khoản (nếu có)."""
    tk = (tai_khoan or "").strip().lower()
    return MAT_KHAU_USMART_MAC_DINH.get(tk, "")


def tim_ma_can_bo_theo_tai_khoan_usmart(tai_khoan: str) -> Optional[str]:
    """Ánh xạ tên đăng nhập uSmart (`giangntt`, `ducla`, `linhnt2`, `hanh1`, ...) sang mã cán bộ."""
    tk = (tai_khoan or "").strip().lower()
    if not tk:
        return None
    if tk == "linhnt":
        return "1472"
    for gv in DANH_SACH_GIANG_VIEN_DKT:
        if gv.tai_khoan_usmart.strip().lower() == tk or gv.ma_can_bo.strip().lower() == tk:
            return gv.ma_can_bo
    return None


@dataclass
class NhiemVuCongViec:
    """Mô hình một dòng nhiệm vụ trong Mẫu 01a/01b, Mẫu 02a/02b và Sổ giao việc Mẫu 03."""

    ma_nhiem_vu: str
    nhiem_vu: str
    nguon_giao: str
    so_van_ban_ma_viec: str = ""
    san_pham_ket_qua: str = ""
    thoi_han: Optional[date] = None
    ngay_hoan_thanh: Optional[date] = None
    tinh_trang: str = "Hoàn thành đúng hạn"
    muc_do_hoan_thanh: float = 100.0  # Lưu theo thang 0..100 (khi ghi vào Excel sẽ chia 100 thành 0..1)
    kho_khan_de_xuat: str = ""
    ma_can_bo_chu_tri: str = "1051"
    ten_can_bo_chu_tri: str = "Nguyễn Hải Hà"
    nguoi_phoi_hop: str = ""
    don_vi_phoi_hop: str = ""
    lanh_dao_phu_trach: str = LANH_DAO_PHU_TRACH_DON_VI_MAC_DINH
    ngay_giao: Optional[date] = None
    muc_uu_tien: str = "Thường"
    danh_gia_chat_luong: str = "Đạt yêu cầu – tốt"
    ghi_chu_so_giao_viec: str = ""
    la_nhiem_vu_trong_tam: bool = False
    la_cong_viec_lap_lai: bool = False
    link_minh_chung_drive: str = ""
    nguon_du_lieu: str = "Kế hoạch Bộ môn"

    def da_den_han_trong_ky(self, ngay_cuoi_ky: date) -> int:
        """
        Mô phỏng chính xác công thức cột K (Mẫu 01b):
        =IF(B13="","",IF(OR(AND(F13<>"",$G$5<>"",F13<=$G$5),G13<>""),1,0))
        """
        if not self.nhiem_vu.strip():
            return 0
        if (self.thoi_han is not None and self.thoi_han <= ngay_cuoi_ky) or (
            self.ngay_hoan_thanh is not None
        ):
            return 1
        return 0

    def hoan_thanh_dung_han_excel(self) -> int:
        """
        Mô phỏng chính xác công thức cột L (Mẫu 01b Excel):
        =IF(B13="","",IF(AND(G13<>"",OR(F13="",G13<=F13)),1,0))
        """
        if not self.nhiem_vu.strip():
            return 0
        if self.ngay_hoan_thanh is not None and (
            self.thoi_han is None or self.ngay_hoan_thanh <= self.thoi_han
        ):
            return 1
        return 0

    def hoan_thanh_dung_han_usmart(self) -> int:
        """
        Mô phỏng chính xác quy tắc tính của máy chủ uSmart (app.js dòng 176):
        Nhiệm vụ phải có thời hạn (thoi_han != None) và ngay_ht <= thoi_han mới được tính đúng hạn.
        """
        if not self.nhiem_vu.strip():
            return 0
        if (
            self.ngay_hoan_thanh is not None
            and self.thoi_han is not None
            and self.ngay_hoan_thanh <= self.thoi_han
        ):
            return 1
        return 0

    def canh_bao_excel(self) -> str:
        """
        Mô phỏng chính xác công thức cột M (Mẫu 01b):
        =IF(B13="","",IF(AND(LEFT(H13,10)="Hoàn thành",G13=""),"Thiếu ngày hoàn thành",
          IF(AND(G13<>"",LEFT(H13,10)<>"Hoàn thành"),"Đã có ngày hoàn thành – sửa tình trạng",
          IF(AND(G13<>"",F13<>"",G13>F13,H13="Hoàn thành đúng hạn"),"Hoàn thành sau thời hạn – sửa tình trạng",
          IF(AND(G13<>"",F13<>"",G13<=F13,H13="Hoàn thành chậm"),"Hoàn thành trong hạn – sửa tình trạng","")))))
        """
        if not self.nhiem_vu.strip():
            return ""
        tt = self.tinh_trang.strip()
        if tt.startswith("Hoàn thành") and self.ngay_hoan_thanh is None:
            return "Thiếu ngày hoàn thành"
        if self.ngay_hoan_thanh is not None and not tt.startswith("Hoàn thành"):
            return "Đã có ngày hoàn thành – sửa tình trạng"
        if (
            self.ngay_hoan_thanh is not None
            and self.thoi_han is not None
            and self.ngay_hoan_thanh > self.thoi_han
            and tt == "Hoàn thành đúng hạn"
        ):
            return "Hoàn thành sau thời hạn – sửa tình trạng"
        if (
            self.ngay_hoan_thanh is not None
            and self.thoi_han is not None
            and self.ngay_hoan_thanh <= self.thoi_han
            and tt == "Hoàn thành chậm"
        ):
            return "Hoàn thành trong hạn – sửa tình trạng"
        return ""


@dataclass
class KeHoachKyToi:
    """Mô hình một dòng kế hoạch công việc tuần tới / tháng tới (Mục IV Mẫu 01b / Mẫu 02b)."""

    nhiem_vu: str
    san_pham_du_kien: str
    thoi_han: Optional[date] = None
    ghi_chu: str = ""
    ca_nhan_chu_tri: str = ""
    don_vi_phoi_hop: str = ""


@dataclass
class HoatDongChuyenMon:
    """Mô hình Mục II – Hoạt động chuyên môn của giảng viên trong Mẫu 01b."""

    giang_day_thang: float = 0.0
    giang_day_luy_ke: float = 0.0
    giang_day_dinh_muc: float = 270.0
    giang_day_ghi_chu: str = ""

    nckh_thang: float = 0.0
    nckh_luy_ke: float = 0.0
    nckh_dinh_muc: float = 250.0
    nckh_ghi_chu: str = ""

    nhiem_vu_khac_thang: float = 0.0
    nhiem_vu_khac_luy_ke: float = 0.0
    nhiem_vu_khac_dinh_muc: float = 180.0
    nhiem_vu_khac_ghi_chu: str = ""

    san_pham_khcn_huong_dan: str = ""

    @property
    def ty_le_giang_day(self) -> Optional[float]:
        return (
            self.giang_day_luy_ke / self.giang_day_dinh_muc
            if self.giang_day_dinh_muc > 0
            else None
        )

    @property
    def ty_le_nckh(self) -> Optional[float]:
        return (
            self.nckh_luy_ke / self.nckh_dinh_muc
            if self.nckh_dinh_muc > 0
            else None
        )

    @property
    def ty_le_nhiem_vu_khac(self) -> Optional[float]:
        return (
            self.nhiem_vu_khac_luy_ke / self.nhiem_vu_khac_dinh_muc
            if self.nhiem_vu_khac_dinh_muc > 0
            else None
        )


@dataclass
class BaoCaoCaNhanThang:
    """Mô hình đầy đủ của một Báo cáo công việc tháng cá nhân (Mẫu 01b)."""

    giang_vien: ThongTinGiangVien
    thang: int = 11
    nam: int = 2026
    danh_sach_nhiem_vu: List[NhiemVuCongViec] = field(default_factory=list)
    hoat_dong_chuyen_mon: HoatDongChuyenMon = field(
        default_factory=HoatDongChuyenMon
    )
    diem_tu_cham_tieu_chi_chung: float = 28.0  # Tối đa 30
    diem_tu_cham_ket_qua_nhiem_vu: Optional[float] = None  # Tối đa 70
    diem_truong_don_vi_tieu_chi_chung: Optional[float] = None
    diem_truong_don_vi_ket_qua_nhiem_vu: Optional[float] = None
    ke_hoach_thang_toi: List[KeHoachKyToi] = field(default_factory=list)
    de_xuat_kien_nghi: str = ""
    ngay_ky_bao_cao: Optional[date] = None
    so_bc_tuan_phai_nop: int = 0  # Giảng viên không bắt buộc nộp BC tuần
    so_bc_tuan_nop_dung_han: int = 0

    @property
    def ngay_cuoi_ky(self) -> date:
        return ngay_cuoi_thang(self.thang, self.nam)

    @property
    def tong_so_nhiem_vu(self) -> int:
        return sum(1 for nv in self.danh_sach_nhiem_vu if nv.nhiem_vu.strip())

    @property
    def so_nhiem_vu_da_den_han(self) -> int:
        ck = self.ngay_cuoi_ky
        return sum(nv.da_den_han_trong_ky(ck) for nv in self.danh_sach_nhiem_vu)

    @property
    def so_nhiem_vu_hoan_thanh_dung_han(self) -> int:
        return sum(
            nv.hoan_thanh_dung_han_usmart() for nv in self.danh_sach_nhiem_vu
        )

    @property
    def ty_le_hoan_thanh_dung_han(self) -> Optional[float]:
        den_han = self.so_nhiem_vu_da_den_han
        if den_han == 0:
            return None
        return self.so_nhiem_vu_hoan_thanh_dung_han / den_han

    @property
    def diem_tham_chieu_nhiem_vu(self) -> Optional[int]:
        """Mô phỏng công thức ô J50: =IFERROR(ROUND(70*E38,0),"")"""
        ty_le = self.ty_le_hoan_thanh_dung_han
        if ty_le is None:
            return None
        return int(round(70.0 * ty_le))

    @property
    def diem_tu_cham_nv_hieu_luc(self) -> float:
        if self.diem_tu_cham_ket_qua_nhiem_vu is not None:
            return float(self.diem_tu_cham_ket_qua_nhiem_vu)
        tc = self.diem_tham_chieu_nhiem_vu
        return float(tc if tc is not None else 65.0)

    @property
    def tong_diem_tu_cham(self) -> float:
        return round(
            self.diem_tu_cham_tieu_chi_chung + self.diem_tu_cham_nv_hieu_luc, 1
        )

    @property
    def tong_diem_truong_don_vi(self) -> Optional[float]:
        if (
            self.diem_truong_don_vi_tieu_chi_chung is None
            and self.diem_truong_don_vi_ket_qua_nhiem_vu is None
        ):
            return None
        return round(
            (self.diem_truong_don_vi_tieu_chi_chung or 0.0)
            + (self.diem_truong_don_vi_ket_qua_nhiem_vu or 0.0),
            1,
        )

    @property
    def muc_xep_loai_tham_chieu(self) -> str:
        """
        Mô phỏng chính xác công thức ô E52 Mẫu 01b và cột L Mẫu 02c:
        Ngưỡng 90 - 70 - 50 điểm theo Nghị định 233/2026/NĐ-CP.
        """
        diem = (
            self.tong_diem_truong_don_vi
            if self.tong_diem_truong_don_vi is not None
            else self.tong_diem_tu_cham
        )
        if diem >= 90:
            return "Hoàn thành xuất sắc nhiệm vụ"
        if diem >= 70:
            return "Hoàn thành tốt nhiệm vụ"
        if diem >= 50:
            return "Hoàn thành nhiệm vụ"
        return "Không hoàn thành nhiệm vụ"

    @property
    def ten_file_xlsx_chuan(self) -> str:
        """Trả về tên tệp chuẩn: BCThang_DKT_<HoTen>_<tháng>-<năm>.xlsx"""
        return f"BCThang_{MA_VIET_TAT_DON_VI}_{self.giang_vien.ten_viet_tat_khong_dau}_{self.thang:02d}-{self.nam}.xlsx"

    @property
    def ten_file_docx_chuan(self) -> str:
        """Trả về tên tệp bản cứng Word: BCThang_DKT_<HoTen>_<tháng>-<năm>.docx"""
        return f"BCThang_{MA_VIET_TAT_DON_VI}_{self.giang_vien.ten_viet_tat_khong_dau}_{self.thang:02d}-{self.nam}.docx"


@dataclass
class ThongKeUsmartDonVi:
    """Số liệu thống kê xử lý văn bản, công việc trên hệ thống uSmart (Mục II Mẫu 02a/02b)."""

    so_van_ban_den_han: int = 0
    hoan_thanh_dung_han: int = 0
    hoan_thanh_cham: int = 0
    qua_han_chua_hoan_thanh: int = 0

    @property
    def kiem_tra_sai_lech(self) -> int:
        """Phải bằng 0 theo công thức =IF(E40="","",E40-E41-E42-E43)."""
        return (
            self.so_van_ban_den_han
            - self.hoan_thanh_dung_han
            - self.hoan_thanh_cham
            - self.qua_han_chua_hoan_thanh
        )
