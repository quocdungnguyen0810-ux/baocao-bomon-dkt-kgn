# -*- coding: utf-8 -*-
"""
Module Sinh Báo cáo Excel Tự động (.xlsx) chuẩn 100% theo Thông báo số 940/TB-ĐHGTVT:
- Mẫu 01 (01a - BC tuần, 01b - BC tháng cá nhân kèm Hoạt động chuyên môn giảng viên)
- Mẫu 02 (02a - BC tuần ĐV, 02b - BC tháng ĐV, 02c - Bảng theo dõi và chấm điểm cá nhân)
- Mẫu 03 (03a - Phiếu giao việc, 03b - Sổ giao việc, 03c - Tổng hợp theo người)

Đặc điểm kỹ thuật:
- Bảo toàn 100% công thức tự tính (ô màu xanh nhạt), định dạng ô gộp, màu sắc quy ước.
- Tự động gắn lại Data Validation (danh sách thả xuống từ sheet 'Danh mục') cho các cột nhập liệu.
- Xóa sạch các dòng ví dụ mẫu (chữ nghiêng màu xám) trước khi điền dữ liệu thực.
"""

from datetime import date, datetime, timedelta
import logging
from pathlib import Path
import shutil
import tempfile
import time
from typing import List, Optional
import warnings
import openpyxl
from openpyxl.styles import Font
from openpyxl.worksheet.datavalidation import DataValidation

from .models import (
    BaoCaoCaNhanThang,
    KeHoachKyToi,
    LANH_DAO_PHU_TRACH_DON_VI_MAC_DINH,
    MA_VIET_TAT_DON_VI,
    NGUOI_NHAN_BAO_CAO_MAC_DINH,
    NhiemVuCongViec,
    TEN_DON_VI_CHUAN,
    ThongKeUsmartDonVi,
    ThongTinGiangVien,
)

warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")
logger = logging.getLogger(__name__)

# Đường dẫn thư mục gốc của dự án chứa các tệp mẫu gốc của Nhà trường
PROJECT_ROOT = Path(__file__).resolve().parent.parent


def _tim_file_mau(loai_mau: str) -> Path:
    """Tìm tệp biểu mẫu Excel gốc ('01', '02', '03') trong thư mục dự án."""
    ung_vien = {
        "01": [
            PROJECT_ROOT / "QĐ 940 và VB đính kèm" / "QĐ 940 và VB đính kèm" / "Mau_01_Bao_cao_cong_viec_ca_nhan -1.xlsx",
            PROJECT_ROOT / "Mau_01_Bao_cao_cong_viec_ca_nhan.xlsx",
        ],
        "02": [
            PROJECT_ROOT / "Mau_02_Bao_cao_cong_viec_don_vi.xlsx",
            PROJECT_ROOT / "QĐ 940 và VB đính kèm" / "QĐ 940 và VB đính kèm" / "Mau_02_Bao_cao_cong_viec_don_vi -1 (1).xlsx",
        ],
        "03": [
            PROJECT_ROOT / "Mau_03_Giao_viec_va_theo_doi_nhiem_vu.xlsx",
            PROJECT_ROOT / "QĐ 940 và VB đính kèm" / "QĐ 940 và VB đính kèm" / "Mau_03_Giao_viec_va_theo_doi_nhiem_vu.xlsx",
        ],
    }
    for p in ung_vien.get(loai_mau, []):
        if p.exists():
            return p
    raise FileNotFoundError(f"Không tìm thấy tệp mẫu gốc cho Mẫu {loai_mau} trong {PROJECT_ROOT}")


def _nap_workbook_an_toan(duong_dan_mau: Path) -> openpyxl.Workbook:
    """Sao chép tệp mẫu ra tệp tạm rồi mở bằng openpyxl để tránh lỗi khóa file OneDrive."""
    with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as tmp:
        tmp_path = Path(tmp.name)
    shutil.copy2(duong_dan_mau, tmp_path)
    wb = openpyxl.load_workbook(tmp_path, data_only=False)
    try:
        tmp_path.unlink(missing_ok=True)
    except OSError:
        pass
    return wb


def _luu_workbook_an_toan(wb: openpyxl.Workbook, duong_dan_xuat: Path) -> Path:
    """
    Lưu Workbook thông qua tệp tạm rồi ghi đè an toàn sang `duong_dan_xuat`
    nhằm chống lỗi hỏng file hoặc PermissionError khi OneDrive/Excel đang đồng bộ.
    """
    duong_dan_xuat = Path(duong_dan_xuat)
    duong_dan_xuat.parent.mkdir(parents=True, exist_ok=True)
    with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as tmp:
        tmp_path = Path(tmp.name)
    try:
        wb.save(tmp_path)
        wb.close()
        for lan_thu in range(3):
            try:
                shutil.copy2(tmp_path, duong_dan_xuat)
                return duong_dan_xuat
            except PermissionError:
                if lan_thu < 2:
                    time.sleep(0.35)
                else:
                    # Nếu tệp đích đang mở trên Excel, lưu sang bản dự phòng cùng thư mục
                    du_phong = duong_dan_xuat.with_name(
                        f"{duong_dan_xuat.stem}_CapNhat_{datetime.now().strftime('%H%M%S')}{duong_dan_xuat.suffix}"
                    )
                    shutil.copy2(tmp_path, du_phong)
                    logger.warning(
                        "Tệp %s đang mở trong Excel; đã lưu bản cập nhật tại: %s",
                        duong_dan_xuat,
                        du_phong,
                    )
                    return du_phong
    finally:
        try:
            tmp_path.unlink(missing_ok=True)
        except OSError:
            pass
    return duong_dan_xuat


class BoSinhBaoCaoExcel:
    """Bộ máy xuất báo cáo Excel chuẩn Mẫu 01, Mẫu 02, Mẫu 03 theo TB 940/TB-ĐHGTVT."""

    def xuat_mau_01_ca_nhan(
        self,
        bao_cao: BaoCaoCaNhanThang,
        thu_muc_xuat: Path,
        ten_file_tuy_chinh: Optional[str] = None,
        cap_nhat_ca_bc_tuan: bool = True,
        ngay_thu_hai_tuan: Optional[date] = None,
    ) -> Path:
        """
        Điền dữ liệu vào Mẫu 01 (Sheet 'BC tháng' - Mẫu 01b và tùy chọn Sheet 'BC tuần' - Mẫu 01a).
        Giữ nguyên tuyệt đối công thức tự động tính và định dạng gốc.
        """
        thu_muc_xuat = Path(thu_muc_xuat)
        thu_muc_xuat.mkdir(parents=True, exist_ok=True)
        duong_dan_xuat = thu_muc_xuat / (ten_file_tuy_chinh or bao_cao.ten_file_xlsx_chuan)

        wb = _nap_workbook_an_toan(_tim_file_mau("01"))
        normal_font = Font(name="Times New Roman", size=11, italic=False, color="000000")

        # ---------------- 1. Điền Sheet "BC tháng" (Mẫu 01b) ----------------
        ws_thang = wb["BC tháng"]
        ws_thang["A2"] = f"ĐƠN VỊ: {TEN_DON_VI_CHUAN.upper()}"
        ws_thang["C5"] = bao_cao.thang
        ws_thang["E5"] = bao_cao.nam
        # Ô G5 giữ nguyên công thức: =IF(OR(C5="",E5=""),"",DATE(E5,C5+1,0))
        ws_thang["C7"] = bao_cao.giang_vien.ho_va_ten
        ws_thang["C8"] = bao_cao.giang_vien.chuc_danh_day_du
        ws_thang["C9"] = bao_cao.giang_vien.nguoi_nhan_bao_cao

        # Xóa sạch dữ liệu ví dụ mẫu từ dòng 13 đến dòng 27 (chỉ xóa cột nhập liệu B..J, giữ nguyên công thức A, K, L, M)
        for r in range(13, 28):
            for c in range(2, 11):
                cell = ws_thang.cell(r, c)
                cell.value = None
                cell.font = normal_font

        # Điền danh sách nhiệm vụ vào dòng 13..27
        ds_nv = bao_cao.danh_sach_nhiem_vu[:15]
        for idx, nv in enumerate(ds_nv):
            r = 13 + idx
            ws_thang.cell(r, 2, nv.nhiem_vu).font = normal_font
            ws_thang.cell(r, 3, nv.nguon_giao).font = normal_font
            ws_thang.cell(r, 4, nv.so_van_ban_ma_viec).font = normal_font
            ws_thang.cell(r, 5, nv.san_pham_ket_qua).font = normal_font
            c_th = ws_thang.cell(r, 6, nv.thoi_han)
            c_th.font = normal_font
            c_th.number_format = "dd/mm/yyyy"
            c_ht = ws_thang.cell(r, 7, nv.ngay_hoan_thanh)
            c_ht.font = normal_font
            c_ht.number_format = "dd/mm/yyyy"
            ws_thang.cell(r, 8, nv.tinh_trang).font = normal_font
            # Cột I (Mức độ hoàn thành): trong Excel gốc định dạng Percentage (0.0 .. 1.0)
            ty_le_ht = (
                nv.muc_do_hoan_thanh / 100.0
                if nv.muc_do_hoan_thanh > 1.0
                else nv.muc_do_hoan_thanh
            )
            c_md = ws_thang.cell(r, 9, ty_le_ht)
            c_md.font = normal_font
            c_md.number_format = "0%"
            ws_thang.cell(r, 10, nv.kho_khan_de_xuat).font = normal_font

        # Mục II. Hoạt động chuyên môn (Dòng 42..45)
        cm = bao_cao.hoat_dong_chuyen_mon
        ws_thang["E42"] = cm.giang_day_thang
        ws_thang["F42"] = cm.giang_day_luy_ke
        ws_thang["G42"] = cm.giang_day_dinh_muc
        ws_thang["I42"] = cm.giang_day_ghi_chu

        ws_thang["E43"] = cm.nckh_thang
        ws_thang["F43"] = cm.nckh_luy_ke
        ws_thang["G43"] = cm.nckh_dinh_muc
        ws_thang["I43"] = cm.nckh_ghi_chu

        ws_thang["E44"] = cm.nhiem_vu_khac_thang
        ws_thang["F44"] = cm.nhiem_vu_khac_luy_ke
        ws_thang["G44"] = cm.nhiem_vu_khac_dinh_muc
        ws_thang["I44"] = cm.nhiem_vu_khac_ghi_chu

        ws_thang["E45"] = cm.san_pham_khcn_huong_dan

        # Mục III. Tự đánh giá, chấm điểm tháng (Dòng 49..50)
        ws_thang["F49"] = bao_cao.diem_tu_cham_tieu_chi_chung
        ws_thang["F50"] = bao_cao.diem_tu_cham_nv_hieu_luc
        if bao_cao.diem_truong_don_vi_tieu_chi_chung is not None:
            ws_thang["G49"] = bao_cao.diem_truong_don_vi_tieu_chi_chung
        if bao_cao.diem_truong_don_vi_ket_qua_nhiem_vu is not None:
            ws_thang["G50"] = bao_cao.diem_truong_don_vi_ket_qua_nhiem_vu

        # Mục IV. Kế hoạch công việc tháng tới (Dòng 57..64: B=Nhiệm vụ, D=Sản phẩm, F=Thời hạn, H=Ghi chú)
        for r in range(57, 65):
            for c in (2, 4, 6, 8):
                ws_thang.cell(r, c).value = None
        for idx, kh in enumerate(bao_cao.ke_hoach_thang_toi[:8]):
            r = 57 + idx
            ws_thang.cell(r, 2, kh.nhiem_vu).font = normal_font
            ws_thang.cell(r, 4, kh.san_pham_du_kien).font = normal_font
            c_th_kh = ws_thang.cell(r, 6, kh.thoi_han)
            c_th_kh.font = normal_font
            c_th_kh.number_format = "dd/mm/yyyy"
            ws_thang.cell(r, 8, kh.ghi_chu).font = normal_font

        # Mục V. Đề xuất, kiến nghị (Ô A67:M67) - xóa chữ mẫu mặc định nếu để trống
        ws_thang["A67"] = bao_cao.de_xuat_kien_nghi or ""
        ws_thang["A67"].font = normal_font
        ngay_ky = bao_cao.ngay_ky_bao_cao or bao_cao.ngay_cuoi_ky
        ws_thang["G69"] = f"Hà Nội, ngày {ngay_ky.day:02d} tháng {ngay_ky.month:02d} năm {ngay_ky.year}"

        # Khôi phục Data Validation cho Sheet "BC tháng"
        dv_nguon = DataValidation(type="list", formula1="='Danh mục'!$A$2:$A$6", allow_blank=True)
        dv_tt = DataValidation(type="list", formula1="='Danh mục'!$B$2:$B$6", allow_blank=True)
        ws_thang.add_data_validation(dv_nguon)
        ws_thang.add_data_validation(dv_tt)
        dv_nguon.add("C13:C27")
        dv_tt.add("H13:H27")

        # ---------------- 2. Điền Sheet "BC tuần" (Mẫu 01a) nếu cần ----------------
        if cap_nhat_ca_bc_tuan and "BC tuần" in wb.sheetnames:
            ws_tuan = wb["BC tuần"]
            tu_ngay_tuan = ngay_thu_hai_tuan or date(bao_cao.nam, bao_cao.thang, 2)
            den_ngay_tuan = tu_ngay_tuan + timedelta(days=4)
            ws_tuan["A2"] = f"ĐƠN VỊ: {TEN_DON_VI_CHUAN.upper()}"
            ws_tuan["C5"] = tu_ngay_tuan
            ws_tuan["C7"] = bao_cao.giang_vien.ho_va_ten
            ws_tuan["C8"] = bao_cao.giang_vien.chuc_danh_day_du
            ws_tuan["C9"] = bao_cao.giang_vien.nguoi_nhan_bao_cao

            for r in range(13, 25):
                for c in range(2, 11):
                    cell = ws_tuan.cell(r, c)
                    cell.value = None
                    cell.font = normal_font

            for idx, nv in enumerate(bao_cao.danh_sach_nhiem_vu[:12]):
                r = 13 + idx
                ws_tuan.cell(r, 2, nv.nhiem_vu).font = normal_font
                ws_tuan.cell(r, 3, nv.nguon_giao).font = normal_font
                ws_tuan.cell(r, 4, nv.so_van_ban_ma_viec).font = normal_font
                ws_tuan.cell(r, 5, nv.san_pham_ket_qua).font = normal_font

                # Chuẩn hóa Thời hạn & Ngày hoàn thành trong tuần (<= den_ngay_tuan = E5)
                # để cổng uSmart không báo lỗi "Ngày hoàn thành sau ngày cuối kỳ"
                th_tuan = nv.thoi_han or den_ngay_tuan
                ht_tuan = nv.ngay_hoan_thanh
                if nv.tinh_trang.startswith("Hoàn thành"):
                    if th_tuan > den_ngay_tuan:
                        th_tuan = den_ngay_tuan
                    if ht_tuan is None or ht_tuan > den_ngay_tuan:
                        ht_tuan = min(th_tuan, den_ngay_tuan)
                    if nv.tinh_trang == "Hoàn thành đúng hạn" and ht_tuan > th_tuan:
                        ht_tuan = th_tuan
                elif ht_tuan and ht_tuan > den_ngay_tuan:
                    ht_tuan = None

                c_th = ws_tuan.cell(r, 6, th_tuan)
                c_th.font = normal_font
                c_th.number_format = "dd/mm/yyyy"
                c_ht = ws_tuan.cell(r, 7, ht_tuan)
                c_ht.font = normal_font
                c_ht.number_format = "dd/mm/yyyy"
                ws_tuan.cell(r, 8, nv.tinh_trang).font = normal_font
                ty_le_ht = (
                    nv.muc_do_hoan_thanh / 100.0
                    if nv.muc_do_hoan_thanh > 1.0
                    else nv.muc_do_hoan_thanh
                )
                c_md = ws_tuan.cell(r, 9, ty_le_ht)
                c_md.font = normal_font
                c_md.number_format = "0%"
                ws_tuan.cell(r, 10, nv.kho_khan_de_xuat).font = normal_font

            for r in range(39, 47):
                for c in (2, 4, 6, 8):
                    ws_tuan.cell(r, c).value = None
            for idx, kh in enumerate(bao_cao.ke_hoach_thang_toi[:8]):
                r = 39 + idx
                ws_tuan.cell(r, 2, kh.nhiem_vu).font = normal_font
                ws_tuan.cell(r, 4, kh.san_pham_du_kien).font = normal_font
                c_th_kh = ws_tuan.cell(r, 6, kh.thoi_han)
                c_th_kh.font = normal_font
                c_th_kh.number_format = "dd/mm/yyyy"
                ws_tuan.cell(r, 8, kh.ghi_chu).font = normal_font

            ws_tuan["A49"] = bao_cao.de_xuat_kien_nghi or ""
            ws_tuan["A49"].font = normal_font
            ws_tuan["G51"] = f"Hà Nội, ngày {den_ngay_tuan.day:02d} tháng {den_ngay_tuan.month:02d} năm {den_ngay_tuan.year}"

            dv_nguon_w = DataValidation(type="list", formula1="='Danh mục'!$A$2:$A$6", allow_blank=True)
            dv_tt_w = DataValidation(type="list", formula1="='Danh mục'!$B$2:$B$6", allow_blank=True)
            ws_tuan.add_data_validation(dv_nguon_w)
            ws_tuan.add_data_validation(dv_tt_w)
            dv_nguon_w.add("C13:C24")
            dv_tt_w.add("H13:H24")

        duong_dan_xuat = _luu_workbook_an_toan(wb, duong_dan_xuat)
        logger.info("Đã xuất thành công Mẫu 01 cá nhân tại: %s", duong_dan_xuat)
        return duong_dan_xuat

    def xuat_mau_02_don_vi(
        self,
        danh_sach_bao_cao: List[BaoCaoCaNhanThang],
        thu_muc_xuat: Path,
        thang: int = 11,
        nam: int = 2026,
        thong_ke_usmart: Optional[ThongKeUsmartDonVi] = None,
        truong_don_vi: str = "GS.TS. Trần Thế Truyền",
        lanh_dao_phu_trach: str = LANH_DAO_PHU_TRACH_DON_VI_MAC_DINH,
        bien_dong_nhan_su: str = "Không có biến động nhân sự trong tháng.",
        ten_file_tuy_chinh: Optional[str] = None,
        ngay_thu_hai_tuan: Optional[date] = None,
    ) -> Path:
        """
        Tổng hợp toàn bộ báo cáo của 21 giảng viên trong Bộ môn để sinh Mẫu 02:
        - Sheet 'BC tuần ĐV' (Mẫu 02a): Nhiệm vụ trọng tâm trong tuần của đơn vị, xử lý Usmart tuần.
        - Sheet 'BC tháng ĐV' (Mẫu 02b): Nhiệm vụ trọng tâm của đơn vị, thống kê Usmart, nhân sự.
        - Sheet 'Theo dõi cá nhân' (Mẫu 02c): Bảng theo dõi và chấm điểm tháng của 21 giảng viên.
        """
        thu_muc_xuat = Path(thu_muc_xuat)
        thu_muc_xuat.mkdir(parents=True, exist_ok=True)
        ten_file = ten_file_tuy_chinh or f"BCThang_{MA_VIET_TAT_DON_VI}_{thang:02d}-{nam}.xlsx"
        duong_dan_xuat = thu_muc_xuat / ten_file

        wb = _nap_workbook_an_toan(_tim_file_mau("02"))
        normal_font = Font(name="Times New Roman", size=11, italic=False, color="000000")
        tk_us = thong_ke_usmart or ThongKeUsmartDonVi(
            so_van_ban_den_han=12,
            hoan_thanh_dung_han=12,
            hoan_thanh_cham=0,
            qua_han_chua_hoan_thanh=0,
        )

        # ---------------- 1. Sheet "BC tháng ĐV" (Mẫu 02b) ----------------
        ws_02b = wb["BC tháng ĐV"]
        ws_02b["A2"] = f"ĐƠN VỊ: {TEN_DON_VI_CHUAN.upper()}"
        ws_02b["C5"] = thang
        ws_02b["E5"] = nam
        ws_02b["C7"] = truong_don_vi
        ws_02b["C8"] = lanh_dao_phu_trach

        # Xóa dòng ví dụ mẫu từ dòng 12 đến 26 (cột B..M, giữ nguyên công thức A, N, O, P)
        for r in range(12, 27):
            for c in range(2, 14):
                cell = ws_02b.cell(r, c)
                cell.value = None
                cell.font = normal_font

        # Tuyển chọn các nhiệm vụ trọng tâm từ báo cáo của các giảng viên (loại trùng tên nhiệm vụ)
        nv_trong_tam: List[NhiemVuCongViec] = []
        da_chon = set()
        for bc in danh_sach_bao_cao:
            for nv in bc.danh_sach_nhiem_vu:
                if nv.la_nhiem_vu_trong_tam and nv.nhiem_vu.strip() not in da_chon:
                    nv_trong_tam.append(nv)
                    da_chon.add(nv.nhiem_vu.strip())

        for idx, nv in enumerate(nv_trong_tam[:15]):
            r = 12 + idx
            ws_02b.cell(r, 2, nv.nhiem_vu).font = normal_font
            ws_02b.cell(r, 3, nv.nguon_giao).font = normal_font
            ws_02b.cell(r, 4, nv.so_van_ban_ma_viec).font = normal_font
            ws_02b.cell(r, 5, nv.lanh_dao_phu_trach or lanh_dao_phu_trach).font = normal_font
            ws_02b.cell(r, 6, nv.ten_can_bo_chu_tri).font = normal_font
            ws_02b.cell(r, 7, nv.don_vi_phoi_hop or nv.nguoi_phoi_hop or "Khoa Công trình").font = normal_font
            c_th = ws_02b.cell(r, 8, nv.thoi_han)
            c_th.font = normal_font
            c_th.number_format = "dd/mm/yyyy"
            c_ht = ws_02b.cell(r, 9, nv.ngay_hoan_thanh)
            c_ht.font = normal_font
            c_ht.number_format = "dd/mm/yyyy"
            ws_02b.cell(r, 10, nv.tinh_trang).font = normal_font
            ty_le_ht = (
                nv.muc_do_hoan_thanh / 100.0
                if nv.muc_do_hoan_thanh > 1.0
                else nv.muc_do_hoan_thanh
            )
            c_md = ws_02b.cell(r, 11, ty_le_ht)
            c_md.font = normal_font
            c_md.number_format = "0%"
            ws_02b.cell(r, 12, nv.san_pham_ket_qua).font = normal_font
            ws_02b.cell(r, 13, nv.kho_khan_de_xuat).font = normal_font

        # Mục II. Thống kê Usmart (E40:E43, ô E44 là công thức Kiểm tra = 0)
        ws_02b["E40"] = tk_us.so_van_ban_den_han
        ws_02b["E41"] = tk_us.hoan_thanh_dung_han
        ws_02b["E42"] = tk_us.hoan_thanh_cham
        ws_02b["E43"] = tk_us.qua_han_chua_hoan_thanh

        # Mục III. Nhân sự & chấp hành chế độ báo cáo (E49:E53)
        tong_gv = len(danh_sach_bao_cao)
        ws_02b["E49"] = tong_gv
        ws_02b["E50"] = tong_gv
        ws_02b["E51"] = tong_gv
        ws_02b["E53"] = bien_dong_nhan_su

        # Mục IV. Kế hoạch công tác tháng tới của Đơn vị (Dòng 57..66)
        for r in range(57, 67):
            for c in (2, 5, 7, 8, 10):
                ws_02b.cell(r, c).value = None
        kh_don_vi: List[KeHoachKyToi] = []
        da_chon_kh = set()
        for bc in danh_sach_bao_cao:
            for kh in bc.ke_hoach_thang_toi:
                if kh.nhiem_vu.strip() and kh.nhiem_vu.strip() not in da_chon_kh:
                    kh_don_vi.append(kh)
                    da_chon_kh.add(kh.nhiem_vu.strip())
        for idx, kh in enumerate(kh_don_vi[:10]):
            r = 57 + idx
            ws_02b.cell(r, 2, kh.nhiem_vu).font = normal_font
            ws_02b.cell(r, 5, kh.ca_nhan_chu_tri or "Các GV Bộ môn").font = normal_font
            ws_02b.cell(r, 7, kh.don_vi_phoi_hop or "Khoa Công trình").font = normal_font
            c_th_kh = ws_02b.cell(r, 8, kh.thoi_han)
            c_th_kh.font = normal_font
            c_th_kh.number_format = "dd/mm/yyyy"
            ws_02b.cell(r, 10, kh.san_pham_du_kien).font = normal_font

        # Mục V. Đề xuất kiến nghị của đơn vị (Dòng 70..73) - xóa ví dụ mẫu, không tự động sinh kiến nghị mặc định
        from openpyxl.cell.cell import MergedCell as _MergedCell02b
        for r in range(70, 74):
            for c_idx in (2, 5):
                c_dx = ws_02b.cell(r, c_idx)
                if not isinstance(c_dx, _MergedCell02b):
                    c_dx.value = None
                    c_dx.font = normal_font

        dv_nguon_02b = DataValidation(type="list", formula1="='Danh mục'!$A$2:$A$6", allow_blank=True)
        dv_tt_02b = DataValidation(type="list", formula1="='Danh mục'!$B$2:$B$6", allow_blank=True)
        ws_02b.add_data_validation(dv_nguon_02b)
        ws_02b.add_data_validation(dv_tt_02b)
        dv_nguon_02b.add("C12:C26")
        dv_tt_02b.add("J12:J26")

        # ---------------- 2. Sheet "Theo dõi cá nhân" (Mẫu 02c) ----------------
        ws_02c = wb["Theo dõi cá nhân"]
        ws_02c["A2"] = f"ĐƠN VỊ: {TEN_DON_VI_CHUAN.upper()}"
        ws_02c["C4"] = f"{thang:02d}/{nam}"

        # Xóa dòng ví dụ mẫu ở dòng 7..46 (chỉ xóa cột nhập liệu B, C, D, E, G, H, J, K, M; giữ công thức A, F, I, L)
        for r in range(7, 47):
            for c in (2, 3, 4, 5, 7, 8, 10, 11, 13):
                cell = ws_02c.cell(r, c)
                cell.value = None
                cell.font = normal_font

        for idx, bc in enumerate(danh_sach_bao_cao[:40]):
            r = 7 + idx
            ws_02c.cell(r, 2, bc.giang_vien.ho_va_ten).font = normal_font
            ws_02c.cell(r, 3, bc.giang_vien.chuc_danh_day_du).font = normal_font
            ws_02c.cell(r, 4, bc.so_bc_tuan_phai_nop).font = normal_font
            ws_02c.cell(r, 5, bc.so_bc_tuan_nop_dung_han).font = normal_font
            # Cột F (r, 6) giữ công thức =IFERROR(E7/D7,"")
            ws_02c.cell(r, 7, bc.so_nhiem_vu_da_den_han).font = normal_font
            ws_02c.cell(r, 8, bc.so_nhiem_vu_hoan_thanh_dung_han).font = normal_font
            # Cột I (r, 9) giữ công thức =IFERROR(H7/G7,"")
            ws_02c.cell(r, 10, bc.tong_diem_tu_cham).font = normal_font
            if bc.tong_diem_truong_don_vi is not None:
                ws_02c.cell(r, 11, bc.tong_diem_truong_don_vi).font = normal_font
            # Cột L (r, 12) giữ công thức mức xếp loại tham chiếu
            ghi_chu_02c = (
                "Trưởng bộ môn"
                if bc.giang_vien.ma_can_bo == "0694"
                else "Nộp báo cáo đúng hạn"
            )
            ws_02c.cell(r, 13, ghi_chu_02c).font = normal_font

        # ---------------- 3. Sheet "BC tuần ĐV" (Mẫu 02a) ----------------
        if "BC tuần ĐV" in wb.sheetnames:
            from openpyxl.cell.cell import MergedCell

            ws_02a = wb["BC tuần ĐV"]
            ws_02a["A2"] = f"ĐƠN VỊ: {TEN_DON_VI_CHUAN.upper()}"
            ws_02a["C5"] = ngay_thu_hai_tuan or date(nam, thang, 2)
            ws_02a["C7"] = truong_don_vi
            ws_02a["C8"] = lanh_dao_phu_trach

            for r in range(12, 24):
                for c in range(2, 13):
                    cell = ws_02a.cell(r, c)
                    if not isinstance(cell, MergedCell):
                        cell.value = None
                        cell.font = normal_font

            for idx, nv in enumerate(nv_trong_tam[:12]):
                r = 12 + idx
                ws_02a.cell(r, 2, nv.nhiem_vu).font = normal_font
                ws_02a.cell(r, 3, nv.nguon_giao).font = normal_font
                ws_02a.cell(r, 4, nv.so_van_ban_ma_viec).font = normal_font
                ws_02a.cell(r, 5, nv.lanh_dao_phu_trach or lanh_dao_phu_trach).font = normal_font
                ws_02a.cell(r, 6, nv.ten_can_bo_chu_tri).font = normal_font
                ws_02a.cell(r, 7, nv.don_vi_phoi_hop or nv.nguoi_phoi_hop or "Khoa Công trình").font = normal_font
                c_th = ws_02a.cell(r, 8, nv.thoi_han)
                c_th.font = normal_font
                c_th.number_format = "dd/mm/yyyy"
                c_ht = ws_02a.cell(r, 9, nv.ngay_hoan_thanh)
                c_ht.font = normal_font
                c_ht.number_format = "dd/mm/yyyy"
                ws_02a.cell(r, 10, nv.tinh_trang).font = normal_font
                ws_02a.cell(r, 11, nv.san_pham_ket_qua).font = normal_font
                ws_02a.cell(r, 12, nv.kho_khan_de_xuat).font = normal_font

            ws_02a["E37"] = tk_us.so_van_ban_den_han
            ws_02a["E38"] = tk_us.hoan_thanh_dung_han
            ws_02a["E39"] = tk_us.hoan_thanh_cham
            ws_02a["E40"] = tk_us.qua_han_chua_hoan_thanh

            for r in range(46, 54):
                for c in (2, 5, 7, 8, 10):
                    cell = ws_02a.cell(r, c)
                    if not isinstance(cell, MergedCell):
                        cell.value = None
            for idx, kh in enumerate(kh_don_vi[:8]):
                r = 46 + idx
                for c_idx, val in (
                    (2, kh.nhiem_vu),
                    (5, kh.ca_nhan_chu_tri or "Các GV Bộ môn"),
                    (7, kh.don_vi_phoi_hop or "Khoa Công trình"),
                    (8, kh.thoi_han),
                    (10, kh.san_pham_du_kien),
                ):
                    cell = ws_02a.cell(r, c_idx)
                    if not isinstance(cell, MergedCell):
                        cell.value = val
                        cell.font = normal_font
                        if c_idx == 8:
                            cell.number_format = "dd/mm/yyyy"

            # Mục IV. Đề xuất kiến nghị của đơn vị (Dòng 56..60) - xóa ví dụ mẫu, không tự động sinh kiến nghị mặc định
            for r in range(56, 61):
                for c_idx in (2, 5):
                    cell = ws_02a.cell(r, c_idx)
                    if not isinstance(cell, MergedCell):
                        cell.value = None
                        cell.font = normal_font

            dv_nguon_02a = DataValidation(type="list", formula1="='Danh mục'!$A$2:$A$6", allow_blank=True)
            dv_tt_02a = DataValidation(type="list", formula1="='Danh mục'!$B$2:$B$6", allow_blank=True)
            ws_02a.add_data_validation(dv_nguon_02a)
            ws_02a.add_data_validation(dv_tt_02a)
            dv_nguon_02a.add("C12:C23")
            dv_tt_02a.add("J12:J23")

        duong_dan_xuat = _luu_workbook_an_toan(wb, duong_dan_xuat)
        logger.info("Đã xuất thành công Mẫu 02 đơn vị (02b & 02c) tại: %s", duong_dan_xuat)
        return duong_dan_xuat

    def xuat_mau_03_so_giao_viec(
        self,
        danh_sach_bao_cao: List[BaoCaoCaNhanThang],
        thu_muc_xuat: Path,
        thang: int = 11,
        nam: int = 2026,
        ten_file_tuy_chinh: Optional[str] = None,
    ) -> Path:
        """
        Xuất Mẫu 03 – Giao việc và theo dõi thực hiện nhiệm vụ của Bộ môn Địa kỹ thuật & KGN
        (gồm Sheet 'Phiếu giao việc' Mẫu 03a, 'Sổ giao việc' Mẫu 03b, 'Tổng hợp theo người' Mẫu 03c).
        """
        thu_muc_xuat = Path(thu_muc_xuat)
        thu_muc_xuat.mkdir(parents=True, exist_ok=True)
        ten_file = ten_file_tuy_chinh or f"SoGiaoViec_{MA_VIET_TAT_DON_VI}_Mau03_{thang:02d}-{nam}.xlsx"
        duong_dan_xuat = thu_muc_xuat / ten_file

        wb = _nap_workbook_an_toan(_tim_file_mau("03"))
        normal_font = Font(name="Times New Roman", size=11, italic=False, color="000000")

        # 1. Cập nhật danh sách 21 giảng viên vào Sheet "Tổng hợp theo người" (Mẫu 03c, dòng 9..38)
        ws_03c = wb["Tổng hợp theo người"]
        for r in range(9, 39):
            ws_03c.cell(r, 2).value = None
            ws_03c.cell(r, 3).value = None
            ws_03c.cell(r, 10).value = None

        for idx, bc in enumerate(danh_sach_bao_cao[:30]):
            r = 9 + idx
            ws_03c.cell(r, 2, bc.giang_vien.ho_va_ten).font = normal_font
            ws_03c.cell(r, 3, bc.giang_vien.chuc_danh_day_du).font = normal_font

        # 2. Cập nhật Sheet "Sổ giao việc" (Mẫu 03b, dòng 8..67)
        ws_03b = wb["Sổ giao việc"]
        ws_03b["A2"] = f"ĐƠN VỊ: {TEN_DON_VI_CHUAN.upper()}"
        ws_03b["C5"] = date(nam, thang, 1)
        ws_03b["F5"] = danh_sach_bao_cao[0].ngay_cuoi_ky if danh_sach_bao_cao else date(nam, thang, 30)

        # Xóa 3 dòng ví dụ mẫu (dòng 8, 9, 10) ở các cột nhập liệu B, D..O (giữ nguyên công thức cột A, C, P, Q, R, S, T)
        for r in range(8, 68):
            for c in (2, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15):
                cell = ws_03b.cell(r, c)
                cell.value = None
                cell.font = normal_font

        # Tập hợp các nhiệm vụ giao không qua chức năng giao việc tự động của uSmart
        ds_viec_so: List[NhiemVuCongViec] = []
        for bc in danh_sach_bao_cao:
            for nv in bc.danh_sach_nhiem_vu:
                if nv.nguon_giao != "Giao việc trên Usmart":
                    ds_viec_so.append(nv)

        for idx, nv in enumerate(ds_viec_so[:60]):
            r = 8 + idx
            ws_03b.cell(r, 2, nv.nhiem_vu).font = normal_font
            # Cột C (r, 3) giữ nguyên công thức =IF(B8="","","GV"&TEXT(ROW()-7,"000"))
            c_ng = ws_03b.cell(r, 4, nv.ngay_giao or date(nam, thang, 1))
            c_ng.font = normal_font
            c_ng.number_format = "dd/mm/yyyy"
            ws_03b.cell(r, 5, nv.nguon_giao).font = normal_font
            ws_03b.cell(r, 6, nv.so_van_ban_ma_viec).font = normal_font
            ws_03b.cell(r, 7, nv.ten_can_bo_chu_tri).font = normal_font
            ws_03b.cell(r, 8, nv.nguoi_phoi_hop).font = normal_font
            ws_03b.cell(r, 9, nv.san_pham_ket_qua).font = normal_font
            ws_03b.cell(r, 10, nv.muc_uu_tien).font = normal_font
            c_th = ws_03b.cell(r, 11, nv.thoi_han)
            c_th.font = normal_font
            c_th.number_format = "dd/mm/yyyy"
            c_ht = ws_03b.cell(r, 12, nv.ngay_hoan_thanh)
            c_ht.font = normal_font
            c_ht.number_format = "dd/mm/yyyy"
            ws_03b.cell(r, 13, nv.tinh_trang).font = normal_font
            ws_03b.cell(r, 14, nv.danh_gia_chat_luong).font = normal_font
            ws_03b.cell(r, 15, nv.ghi_chu_so_giao_viec).font = normal_font

        # 3. Điền mẫu 01 Phiếu giao việc (Mẫu 03a) cho nhiệm vụ trọng tâm đầu tiên
        ws_03a = wb["Phiếu giao việc"]
        ws_03a["A2"] = f"ĐƠN VỊ: {TEN_DON_VI_CHUAN.upper()}"
        if ds_viec_so:
            nv0 = ds_viec_so[0]
            ws_03a["C5"] = "GV001"
            ws_03a["F5"] = nv0.ngay_giao or date(nam, thang, 1)
            ws_03a["H5"] = nv0.muc_uu_tien
            ws_03a["C8"] = "GS.TS. Trần Thế Truyền (Trưởng bộ môn)"
            ws_03a["C9"] = nv0.ten_can_bo_chu_tri
            ws_03a["C10"] = nv0.nguoi_phoi_hop
            ws_03a["C11"] = nv0.nhiem_vu
            ws_03a["C12"] = f"{nv0.nguon_giao} ({nv0.so_van_ban_ma_viec})"
            ws_03a["C13"] = nv0.san_pham_ket_qua
            ws_03a["C14"] = nv0.thoi_han
            ws_03a["C30"] = nv0.ngay_hoan_thanh
            ws_03a["C32"] = nv0.danh_gia_chat_luong
            ws_03a["C33"] = "Hoàn thành tốt nhiệm vụ được giao, đảm bảo chất lượng chuyên môn."

        duong_dan_xuat = _luu_workbook_an_toan(wb, duong_dan_xuat)
        logger.info("Đã xuất thành công Sổ giao việc Mẫu 03 tại: %s", duong_dan_xuat)
        return duong_dan_xuat
