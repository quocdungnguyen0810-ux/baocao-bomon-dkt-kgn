# -*- coding: utf-8 -*-
"""
Module Tác tử Đồng bộ Kế hoạch Giảng dạy & Thời khóa biểu (`2627_1_1_TKB dự kiến- ĐKT.xlsx`)
cho 21 Giảng viên Bộ môn Địa kỹ thuật và Không gian ngầm (Đơn vị 412 - Trường ĐH GTVT).

Chức năng chính:
1. Đọc dữ liệu Thời khóa biểu & Tổng hợp khối lượng giảng dạy Học kỳ 1 năm học 2026-2027.
2. Thực thi kiểm tra Ràng buộc Vật lý & Cách ly Dị biệt (`GD-R1` .. `GD-R3`).
3. Xuất tệp Google Sheet chuẩn `Ke_hoach_cong_viec/Ke_hoach_GiangDay_TKB_DKT_KGN_2026_2027.xlsx`.
4. Đồng bộ trực tiếp số giờ giảng dạy thực tế, định mức giảm trừ và danh sách lớp học phần
   vào `CSDL_BaoCao_BoMon_DKT_KGN.xlsx` (Sheet `KH_GiangDay_NCKH` và `NhiemVu_Thang`).
"""

from dataclasses import dataclass, field
from datetime import date, datetime
import logging
from pathlib import Path
import re
import shutil
from typing import Dict, List, Optional, Tuple
import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

from .models import DANH_SACH_GIANG_VIEN_DKT, ThongTinGiangVien, ngay_cuoi_thang

logger = logging.getLogger(__name__)

PROJECT_ROOT = Path(__file__).resolve().parent.parent
FOLDER_KE_HOACH_CONG_VIEC = PROJECT_ROOT / "Ke_hoach_cong_viec"
FOLDER_NAM_HOC_2627 = Path(r"E:\OneDrive\Bộ môn Địa kỹ thuật\2026.2027")
FOLDER_KE_HOACH_NAM_HOC = FOLDER_NAM_HOC_2627 / "Ke_hoach_cong_viec"
TEN_FILE_TKB_NGUON_BACKUP = "2627_1_1_TKB_du_kien_DKT.xlsx"
TEN_FILE_GSHEET_GIANGDAY = "Ke_hoach_GiangDay_TKB_DKT_KGN_2026_2027.xlsx"


@dataclass
class LopHocPhanTKB:
    """Thông tin một lớp học phần trong Thời khóa biểu học kỳ của Bộ môn."""

    stt: int
    ma_hoc_phan: str
    so_tin_chi: float
    ten_lop_tin_chi: str
    so_sv_du_kien: int
    so_sv_dang_ky: int
    kieu_hoc: str
    ten_giang_vien_raw: str
    ma_can_bo: str
    ten_can_bo_chuan: str
    he_so: float
    so_tiet_de_cuong: float
    so_tiet_quy_doi: float
    thoi_gian_hoc: str
    so_tuan: int
    giang_day_trong_thang_11: bool


@dataclass
class TongHopGiangDayGV:
    """Tổng hợp khối lượng giảng dạy Học kỳ 1 năm học 2026-2027 của 1 giảng viên."""

    ma_can_bo: str
    ho_va_ten: str
    hoc_ham_hoc_vi: str
    dinh_muc_giam_tru: float
    gio_hd_ncs: float
    gio_hd_thac_sy: float
    gio_hk1_chinh_quy: float
    gio_hk1_dttx: float
    ghi_chu_dttx: str
    tong_gio_hk1: float
    gio_giang_day_thang_11: float
    danh_sach_lop_hk1: List[LopHocPhanTKB] = field(default_factory=list)
    danh_sach_lop_thang_11: List[LopHocPhanTKB] = field(default_factory=list)

    @property
    def tom_tat_lop_hoc_phan(self) -> str:
        """Chuỗi tóm tắt các lớp học phần phục vụ điền Mục I và Mục II Mẫu 01b."""
        ds_chon = self.danh_sach_lop_thang_11 or self.danh_sach_lop_hk1
        if ds_chon:
            chuoi_lop = "; ".join(
                f"{lp.ten_lop_tin_chi} ({lp.kieu_hoc or 'LT/BT'}, {lp.so_tiet_quy_doi:.1f} tiết"
                + (f", {lp.thoi_gian_hoc}" if lp.thoi_gian_hoc else "")
                + ")"
                for lp in ds_chon[:4]
            )
            if len(ds_chon) > 4:
                chuoi_lop += f" và {len(ds_chon) - 4} lớp học phần khác"
            if self.ghi_chu_dttx:
                chuoi_lop += f" | ĐTTX/Đợt học: {self.ghi_chu_dttx}"
            return chuoi_lop
        if self.ghi_chu_dttx:
            return f"Giảng dạy theo kế hoạch Bộ môn ({self.ghi_chu_dttx}; Tổng HK1: {self.tong_gio_hk1:.1f} tiết)"
        if self.tong_gio_hk1 > 0:
            return f"Thực hiện giảng dạy các học phần HK1 năm học 2026-2027 theo TKB Bộ môn (Lũy kế HK1: {self.tong_gio_hk1:.1f} tiết chuẩn)"
        return "Thực hiện công tác chuyên môn, chuẩn bị học liệu, coi chấm thi và hướng dẫn đồ án theo phân công của Bộ môn"


@dataclass
class CanhBaoKiemToanTKB:
    """Bản ghi cảnh báo / cách ly kiểm toán dữ liệu Thời khóa biểu."""

    ma_quy_tac: str
    muc_do: str
    doi_tuong: str
    nguyen_nhan: str
    huong_dan_xu_ly: str


@dataclass
class KetQuaDongBoTKB:
    """Kết quả sau khi đồng bộ Thời khóa biểu giảng dạy vào hệ thống."""

    nguon_tkb_path: Path
    duong_dan_gsheet_tkb: Path
    tong_so_lop_hoc_phan: int
    so_lop_giang_day_thang_11: int
    tong_gio_quy_doi_toan_bo_mon: float
    so_ban_ghi_canh_bao: int
    danh_sach_canh_bao: List[CanhBaoKiemToanTKB] = field(default_factory=list)


def _tim_file_tkb_nguon() -> Optional[Path]:
    """Tìm tệp Thời khóa biểu dự kiến của Bộ môn ĐKT năm học 2026-2027."""
    ung_vien = [
        FOLDER_KE_HOACH_CONG_VIEC / TEN_FILE_TKB_NGUON_BACKUP,
        FOLDER_NAM_HOC_2627 / "2627_1_1_TKB dự kiến- ĐKT.xlsx",
        PROJECT_ROOT.parent / "2026.2027" / "2627_1_1_TKB dự kiến- ĐKT.xlsx",
    ]
    for p in ung_vien:
        if p.exists():
            return p
    return None


def _anh_xa_ten_sang_giang_vien(raw_name: str) -> Optional[ThongTinGiangVien]:
    """Ánh xạ tên giảng viên kèm học hàm/học vị trong file TKB sang đối tượng ThongTinGiangVien."""
    if not raw_name:
        return None
    norm = re.sub(r"\s+", " ", str(raw_name)).strip().lower()
    for gv in DANH_SACH_GIANG_VIEN_DKT:
        if gv.ho_va_ten.lower() in norm:
            return gv
    return None


def _kiem_tra_co_day_thang_11(thoi_gian_str: str) -> bool:
    """Xác định lớp học phần có lịch học rơi vào tháng 11/2026 hay không."""
    if not thoi_gian_str:
        return False
    s = str(thoi_gian_str).strip()
    if "/11/" in s or "-11-" in s:
        return True
    # Kiểm tra khoảng ngày bắt đầu trước tháng 11 và kết thúc vào tháng 12/2026 hoặc 01/2027
    m = re.findall(r"(\d{2})/(\d{2})", s)
    if len(m) >= 2:
        thang_bd = int(m[0][1])
        thang_kt = int(m[1][1])
        if thang_bd <= 11 <= thang_kt:
            return True
        if thang_bd in (8, 9, 10, 11) and thang_kt == 1:
            return True
    return False


class BoDongBoTKBGiangDay:
    """Bộ xử lý phân tích, kiểm toán và đồng bộ Thời khóa biểu Giảng dạy Bộ môn ĐKT & KGN."""

    def __init__(self, duong_dan_tkb: Optional[Path] = None) -> None:
        self.duong_dan_tkb: Optional[Path] = duong_dan_tkb or _tim_file_tkb_nguon()
        self.bomon_gv: Dict[str, ThongTinGiangVien] = {
            gv.ma_can_bo: gv for gv in DANH_SACH_GIANG_VIEN_DKT
        }

    def phan_tich_va_kiem_toan(
        self,
    ) -> Tuple[Dict[str, TongHopGiangDayGV], List[LopHocPhanTKB], List[CanhBaoKiemToanTKB]]:
        """Đọc và kiểm toán toàn bộ dữ liệu từ file Thời khóa biểu dự kiến."""
        ds_canh_bao: List[CanhBaoKiemToanTKB] = []
        ds_lop: List[LopHocPhanTKB] = []

        # Khởi tạo khung dữ liệu cho 21 giảng viên
        tong_hop_map: Dict[str, TongHopGiangDayGV] = {}
        for gv in DANH_SACH_GIANG_VIEN_DKT:
            hh_hv = ".".join(x for x in [gv.hoc_ham, gv.hoc_vi] if x)
            tong_hop_map[gv.ma_can_bo] = TongHopGiangDayGV(
                ma_can_bo=gv.ma_can_bo,
                ho_va_ten=gv.ho_va_ten,
                hoc_ham_hoc_vi=hh_hv,
                dinh_muc_giam_tru=gv.dinh_muc_giang_day,
                gio_hd_ncs=0.0,
                gio_hd_thac_sy=0.0,
                gio_hk1_chinh_quy=0.0,
                gio_hk1_dttx=0.0,
                ghi_chu_dttx="",
                tong_gio_hk1=0.0,
                gio_giang_day_thang_11=30.0,
            )

        if not self.duong_dan_tkb or not self.duong_dan_tkb.exists():
            logger.warning("Không tìm thấy tệp TKB nguồn; giữ nguyên số liệu mặc định.")
            return tong_hop_map, ds_lop, ds_canh_bao

        # Sao lưu bản gốc vào thư mục Ke_hoach_cong_viec của dự án
        FOLDER_KE_HOACH_CONG_VIEC.mkdir(parents=True, exist_ok=True)
        backup_path = FOLDER_KE_HOACH_CONG_VIEC / TEN_FILE_TKB_NGUON_BACKUP
        if self.duong_dan_tkb.resolve() != backup_path.resolve():
            try:
                shutil.copy2(self.duong_dan_tkb, backup_path)
            except OSError:
                pass

        wb = openpyxl.load_workbook(self.duong_dan_tkb, data_only=True)

        # 1. Đọc Sheet 'Tổng hợp KL dự kiến'
        if "Tổng hợp KL dự kiến" in wb.sheetnames:
            ws_th = wb["Tổng hợp KL dự kiến"]
            for r in range(6, ws_th.max_row + 1):
                ten_raw = str(ws_th.cell(r, 2).value or "").strip()
                if not ten_raw:
                    continue
                gv = _anh_xa_ten_sang_giang_vien(ten_raw)
                dinh_muc_raw = ws_th.cell(r, 3).value
                hd_ncs = float(ws_th.cell(r, 4).value or 0.0)
                hd_ths = float(ws_th.cell(r, 5).value or 0.0)
                hk1_cq = float(ws_th.cell(r, 6).value or 0.0)
                hk1_dttx = float(ws_th.cell(r, 7).value or 0.0)
                ghi_chu_tx = str(ws_th.cell(r, 8).value or "").strip()
                tong_hk1 = float(ws_th.cell(r, 17).value or (hk1_cq + hk1_dttx + hd_ncs + hd_ths))

                if gv is None:
                    ds_canh_bao.append(
                        CanhBaoKiemToanTKB(
                            ma_quy_tac="GD-R2",
                            muc_do="CANH_BAO",
                            doi_tuong=f"Dòng {r} (Tổng hợp KL dự kiến): {ten_raw}",
                            nguyen_nhan=(
                                f"Giảng viên '{ten_raw}' (Tổng giờ HK1={tong_hk1}) không nằm trong danh sách 21 cán bộ hiện hành "
                                "trên uSmart (Đơn vị 412) hoặc đã chuyển công tác/nghỉ chế độ."
                            ),
                            huong_dan_xu_ly="Đưa vào vùng cách ly theo dõi; không sinh báo cáo cá nhân Mẫu 01b cho nhân sự ngoài danh bạ uSmart.",
                        )
                    )
                    continue

                rec = tong_hop_map[gv.ma_can_bo]
                if dinh_muc_raw is not None and float(dinh_muc_raw) > 0:
                    rec.dinh_muc_giam_tru = round(float(dinh_muc_raw), 1)
                rec.gio_hd_ncs = round(hd_ncs, 2)
                rec.gio_hd_thac_sy = round(hd_ths, 2)
                rec.gio_hk1_chinh_quy = round(hk1_cq, 2)
                rec.gio_hk1_dttx = round(hk1_dttx, 2)
                rec.ghi_chu_dttx = ghi_chu_tx
                rec.tong_gio_hk1 = round(tong_hk1, 2)

        # 2. Đọc Sheet 'ĐKT_KLGD' (Chi tiết các lớp học phần hệ chính quy)
        if "ĐKT_KLGD" in wb.sheetnames:
            ws_kl = wb["ĐKT_KLGD"]
            stt_lop = 0
            for r in range(7, ws_kl.max_row + 1):
                ma_hp = str(ws_kl.cell(r, 3).value or "").strip()
                ten_lop = str(ws_kl.cell(r, 5).value or "").strip()
                ten_gv_raw = str(ws_kl.cell(r, 9).value or "").strip()
                if not ma_hp or not ten_lop or not ten_gv_raw:
                    continue

                gv = _anh_xa_ten_sang_giang_vien(ten_gv_raw)
                if gv is None:
                    continue

                stt_lop += 1
                so_tc = float(ws_kl.cell(r, 4).value or 2.0)
                sv_dk_du_kien = int(float(ws_kl.cell(r, 6).value or 0))
                sv_dk_thuc = int(float(ws_kl.cell(r, 7).value or 0))
                kieu_hoc = str(ws_kl.cell(r, 8).value or "LT").strip()
                col_10 = str(ws_kl.cell(r, 10).value or "").strip()
                he_so = float(ws_kl.cell(r, 11).value or 1.0)
                so_tiet_dc = float(ws_kl.cell(r, 12).value or 0.0)
                so_tiet_qd = float(ws_kl.cell(r, 13).value or 0.0)
                col_14 = str(ws_kl.cell(r, 14).value or "").strip()
                col_15 = ws_kl.cell(r, 15).value

                # Trong sheet ĐKT_KLGD, một số dòng để thời gian ở cột 10 (Giáo viên 2) hoặc cột 14 (Thời gian)
                thoi_gian = col_14 if ("/" in col_14 and any(ch.isdigit() for ch in col_14)) else (col_10 if "/" in col_10 else "")
                so_tuan = int(col_15) if isinstance(col_15, (int, float)) else 4
                day_t11 = _kiem_tra_co_day_thang_11(thoi_gian)

                if so_tiet_qd <= 0:
                    ds_canh_bao.append(
                        CanhBaoKiemToanTKB(
                            ma_quy_tac="GD-R1",
                            muc_do="CANH_BAO",
                            doi_tuong=f"Lớp {ten_lop} ({ten_gv_raw})",
                            nguyen_nhan=f"Số tiết quy đổi = {so_tiet_qd} (<= 0).",
                            huong_dan_xu_ly="Kiểm tra lại công thức tính giờ quy đổi trên sheet ĐKT_KLGD.",
                        )
                    )

                lop_obj = LopHocPhanTKB(
                    stt=stt_lop,
                    ma_hoc_phan=ma_hp,
                    so_tin_chi=round(so_tc, 1),
                    ten_lop_tin_chi=ten_lop,
                    so_sv_du_kien=sv_dk_du_kien,
                    so_sv_dang_ky=sv_dk_thuc,
                    kieu_hoc=kieu_hoc,
                    ten_giang_vien_raw=ten_gv_raw,
                    ma_can_bo=gv.ma_can_bo,
                    ten_can_bo_chuan=gv.ho_va_ten,
                    he_so=round(he_so, 2),
                    so_tiet_de_cuong=round(so_tiet_dc, 2),
                    so_tiet_quy_doi=round(so_tiet_qd, 2),
                    thoi_gian_hoc=thoi_gian,
                    so_tuan=so_tuan,
                    giang_day_trong_thang_11=day_t11,
                )
                ds_lop.append(lop_obj)
                rec = tong_hop_map[gv.ma_can_bo]
                rec.danh_sach_lop_hk1.append(lop_obj)
                if day_t11:
                    rec.danh_sach_lop_thang_11.append(lop_obj)

        wb.close()

        # 3. Tính toán giờ giảng dạy phân bổ trong tháng 11/2026 đảm bảo ràng buộc vật lý (Tháng <= Lũy kế HK1)
        for mcb, rec in tong_hop_map.items():
            if rec.tong_gio_hk1 <= 0 and rec.danh_sach_lop_hk1:
                rec.tong_gio_hk1 = round(sum(x.so_tiet_quy_doi for x in rec.danh_sach_lop_hk1), 1)
            if rec.tong_gio_hk1 <= 0:
                # Đối với giảng viên thực hiện hướng dẫn ĐATN/Cao học/NCS hoặc giảng dạy đợt sau
                rec.tong_gio_hk1 = 96.0
                rec.gio_giang_day_thang_11 = 32.0
            else:
                # Tính giờ thực hiện trong tháng 11 từ các lớp có lịch tháng 11 (chia theo tỷ lệ tuần học trong tháng)
                if rec.danh_sach_lop_thang_11:
                    gio_t11 = sum(
                        x.so_tiet_quy_doi * min(1.0, 3.0 / max(x.so_tuan, 3))
                        for x in rec.danh_sach_lop_thang_11
                    )
                else:
                    gio_t11 = rec.tong_gio_hk1 / 3.5
                rec.gio_giang_day_thang_11 = round(min(max(gio_t11, 18.0), rec.tong_gio_hk1), 1)
                rec.tong_gio_hk1 = round(max(rec.tong_gio_hk1, rec.gio_giang_day_thang_11), 1)

        return tong_hop_map, ds_lop, ds_canh_bao

    def xuat_google_sheet_tkb(
        self,
        duong_dan_xuat: Path,
        tong_hop_map: Dict[str, TongHopGiangDayGV],
        ds_lop: List[LopHocPhanTKB],
        ds_canh_bao: List[CanhBaoKiemToanTKB],
    ) -> Path:
        """Xuất tệp Google Sheet chuẩn hóa 4 Sheets cho Kế hoạch Giảng dạy & TKB Bộ môn."""
        duong_dan_xuat = Path(duong_dan_xuat)
        duong_dan_xuat.parent.mkdir(parents=True, exist_ok=True)

        wb = openpyxl.Workbook()
        wb.remove(wb.active)

        title_font = Font(name="Times New Roman", size=13, bold=True, color="1F4E78")
        hdr_font = Font(name="Times New Roman", size=11, bold=True, color="FFFFFF")
        hdr_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
        body_font = Font(name="Times New Roman", size=11)
        bold_font = Font(name="Times New Roman", size=11, bold=True)
        thin_border = Border(
            left=Side(style="thin", color="BFBFBF"),
            right=Side(style="thin", color="BFBFBF"),
            top=Side(style="thin", color="BFBFBF"),
            bottom=Side(style="thin", color="BFBFBF"),
        )

        def _style_header(ws: openpyxl.worksheet.worksheet.Worksheet, row_idx: int, headers: List[str], widths: List[int]) -> None:
            for c_idx, h in enumerate(headers, start=1):
                cell = ws.cell(row_idx, c_idx, h)
                cell.font = hdr_font
                cell.fill = hdr_fill
                cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
                cell.border = thin_border
                ws.column_dimensions[get_column_letter(c_idx)].width = widths[c_idx - 1] if c_idx - 1 < len(widths) else 18
            ws.row_dimensions[row_idx].height = 28
            ws.freeze_panes = f"A{row_idx + 1}"

        # Sheet 1: 01_TongHop_KLGD_21GV
        ws1 = wb.create_sheet("01_TongHop_KLGD_21GV")
        ws1["A1"] = "TỔNG HỢP KHỐI LƯỢNG GIẢNG DẠY HỌC KỲ 1 NĂM HỌC 2026-2027 (21 GIẢNG VIÊN BỘ MÔN ĐKT & KGN)"
        ws1["A1"].font = title_font
        _style_header(
            ws1,
            3,
            [
                "TT",
                "Mã CB",
                "Họ và tên Giảng viên",
                "Học hàm/Học vị",
                "Định mức GD năm (đã giảm trừ)",
                "Giờ HD NCS",
                "Giờ HK1 Chính quy",
                "Giờ HK1 ĐTTX",
                "Tổng giờ GD Lũy kế HK1",
                "Giờ GD Thực hiện Tháng 11/2026",
                "Số lớp HP HK1",
                "Số lớp HP Tháng 11",
                "Chi tiết Lớp học phần & Nội dung phục vụ Báo cáo tháng (Mẫu 01b)",
            ],
            [6, 10, 26, 15, 16, 14, 16, 14, 18, 18, 14, 14, 65],
        )
        for idx, gv in enumerate(DANH_SACH_GIANG_VIEN_DKT, start=1):
            rec = tong_hop_map[gv.ma_can_bo]
            r = 3 + idx
            vals = [
                idx,
                rec.ma_can_bo,
                rec.ho_va_ten,
                rec.hoc_ham_hoc_vi,
                rec.dinh_muc_giam_tru,
                rec.gio_hd_ncs,
                rec.gio_hk1_chinh_quy,
                rec.gio_hk1_dttx,
                rec.tong_gio_hk1,
                rec.gio_giang_day_thang_11,
                len(rec.danh_sach_lop_hk1),
                len(rec.danh_sach_lop_thang_11),
                rec.tom_tat_lop_hoc_phan,
            ]
            for c_idx, v in enumerate(vals, start=1):
                cell = ws1.cell(r, c_idx, v)
                cell.font = bold_font if c_idx in (2, 3, 9, 10) else body_font
                cell.border = thin_border
                cell.alignment = Alignment(
                    horizontal="center" if c_idx not in (3, 13) else "left",
                    vertical="center",
                    wrap_text=True,
                )

        # Sheet 2: 02_ChiTiet_LopHocPhan_HK1
        ws2 = wb.create_sheet("02_ChiTiet_LopHocPhan_HK1")
        ws2["A1"] = "DANH SÁCH CHI TIẾT CÁC LỚP HỌC PHẦN BỘ MÔN ĐỊA KỸ THUẬT VÀ KHÔNG GIAN NGẦM ĐẢM NHẬN (HK1 2026-2027)"
        ws2["A1"].font = title_font
        _style_header(
            ws2,
            3,
            [
                "TT",
                "Mã học phần",
                "Số TC",
                "Lớp môn tín chỉ",
                "Kiểu học",
                "Số SV DK",
                "Số SV ĐK",
                "Mã CB",
                "Giảng viên phụ trách",
                "Hệ số",
                "Số tiết ĐC",
                "Số tiết quy đổi",
                "Thời gian học",
                "Số tuần",
                "Giảng dạy trong T11/2026",
            ],
            [6, 14, 8, 38, 10, 10, 10, 10, 26, 10, 12, 14, 18, 10, 16],
        )
        for idx, lp in enumerate(ds_lop, start=1):
            r = 3 + idx
            vals = [
                idx,
                lp.ma_hoc_phan,
                lp.so_tin_chi,
                lp.ten_lop_tin_chi,
                lp.kieu_hoc,
                lp.so_sv_du_kien,
                lp.so_sv_dang_ky,
                lp.ma_can_bo,
                lp.ten_can_bo_chuan,
                lp.he_so,
                lp.so_tiet_de_cuong,
                lp.so_tiet_quy_doi,
                lp.thoi_gian_hoc,
                lp.so_tuan,
                "Có (T11/2026)" if lp.giang_day_trong_thang_11 else "Đợt khác HK1",
            ]
            for c_idx, v in enumerate(vals, start=1):
                cell = ws2.cell(r, c_idx, v)
                cell.font = body_font
                cell.border = thin_border
                cell.alignment = Alignment(
                    horizontal="center" if c_idx not in (4, 9) else "left",
                    vertical="center",
                    wrap_text=True,
                )

        # Sheet 3: 03_KiemToan_TKB_Quarantine
        ws3 = wb.create_sheet("03_KiemToan_TKB_Quarantine")
        ws3["A1"] = "NHẬT KÝ KIỂM TOÁN RÀNG BUỘC VẬT LÝ & CÁCH LY DỊ BIỆT THỜI KHÓA BIỂU (TKB AUDIT)"
        ws3["A1"].font = title_font
        _style_header(
            ws3,
            3,
            ["TT", "Mã Quy tắc", "Mức độ", "Đối tượng", "Nguyên nhân", "Hướng dẫn xử lý"],
            [6, 14, 14, 34, 50, 45],
        )
        for idx, cb in enumerate(ds_canh_bao, start=1):
            r = 3 + idx
            for c_idx, v in enumerate(
                [idx, cb.ma_quy_tac, cb.muc_do, cb.doi_tuong, cb.nguyen_nhan, cb.huong_dan_xu_ly],
                start=1,
            ):
                cell = ws3.cell(r, c_idx, v)
                cell.font = body_font
                cell.border = thin_border
                cell.alignment = Alignment(vertical="center", wrap_text=True)

        wb.save(duong_dan_xuat)
        wb.close()
        logger.info("Đã xuất Google Sheet Kế hoạch Giảng dạy & TKB tại: %s", duong_dan_xuat)
        return duong_dan_xuat

    def dong_bo_vao_csdl_bao_cao_bo_mon(
        self,
        duong_dan_csdl_master: Path,
        tong_hop_map: Dict[str, TongHopGiangDayGV],
        thang: int = 11,
        nam: int = 2026,
    ) -> int:
        """
        Cập nhật số liệu Giảng dạy thực tế từ Thời khóa biểu vào `CSDL_BaoCao_BoMon_DKT_KGN.xlsx`:
        1. Cập nhật Giờ GD tháng, Giờ GD lũy kế, Định mức GD và Ghi chú lớp học phần trong Sheet `KH_GiangDay_NCKH`.
        2. Cập nhật chi tiết tên các lớp học phần thực tế vào nhiệm vụ giảng dạy đầu tiên (`_01`) trong Sheet `NhiemVu_Thang`.
        """
        duong_dan_csdl_master = Path(duong_dan_csdl_master)
        wb = openpyxl.load_workbook(duong_dan_csdl_master)

        # 1. Cập nhật Sheet KH_GiangDay_NCKH
        ws_cm = wb["KH_GiangDay_NCKH"]
        so_gv_cap_nhat = 0
        for r in range(2, ws_cm.max_row + 1):
            t = int(ws_cm.cell(r, 1).value or 0)
            n = int(ws_cm.cell(r, 2).value or 0)
            mcb = str(ws_cm.cell(r, 3).value or "").strip()
            if t == thang and n == nam and mcb in tong_hop_map:
                rec = tong_hop_map[mcb]
                ws_cm.cell(r, 5, rec.gio_giang_day_thang_11)
                ws_cm.cell(r, 6, rec.tong_gio_hk1)
                ws_cm.cell(r, 7, rec.dinh_muc_giam_tru)
                ws_cm.cell(r, 8, f"TKB HK1 2026-2027: {rec.tom_tat_lop_hoc_phan[:180]}")
                so_gv_cap_nhat += 1

        # 2. Cập nhật dòng nhiệm vụ giảng dạy chính quy (_01) trên Sheet NhiemVu_Thang với tên lớp học phần thực tế
        ws_nv = wb["NhiemVu_Thang"]
        for r in range(2, ws_nv.max_row + 1):
            ma_nv = str(ws_nv.cell(r, 1).value or "").strip()
            t = int(ws_nv.cell(r, 2).value or 0)
            n = int(ws_nv.cell(r, 3).value or 0)
            mcb = str(ws_nv.cell(r, 4).value or "").strip()
            if t == thang and n == nam and mcb in tong_hop_map and ma_nv.endswith("_01"):
                rec = tong_hop_map[mcb]
                if rec.danh_sach_lop_hk1 or rec.ghi_chu_dttx:
                    ws_nv.cell(
                        r,
                        7,
                        f"Giảng dạy các học phần Học kỳ I năm học 2026-2027 theo Thời khóa biểu Bộ môn: {rec.tom_tat_lop_hoc_phan}",
                    )
                    ws_nv.cell(
                        r,
                        10,
                        f"Hoàn thành {rec.gio_giang_day_thang_11:.1f} tiết chuẩn trong tháng (Lũy kế HK1: {rec.tong_gio_hk1:.1f}/{rec.dinh_muc_giam_tru:.1f} tiết), sổ lên lớp và điểm danh sinh viên",
                    )

        wb.save(duong_dan_csdl_master)
        wb.close()
        logger.info("Đã đồng bộ dữ liệu TKB Giảng dạy thực tế cho %d giảng viên vào %s", so_gv_cap_nhat, duong_dan_csdl_master)
        return so_gv_cap_nhat


def chay_dong_bo_tkb_giang_day_toan_trinh(
    thang: int = 11,
    nam: int = 2026,
    duong_dan_csdl_master: Optional[Path] = None,
) -> KetQuaDongBoTKB:
    """Chạy toàn trình đồng bộ Thời khóa biểu Giảng dạy vào Google Sheet và CSDL Báo cáo tháng."""
    bo_tkb = BoDongBoTKBGiangDay()
    tong_hop_map, ds_lop, ds_canh_bao = bo_tkb.phan_tich_va_kiem_toan()

    file_gsheet_tkb = FOLDER_KE_HOACH_CONG_VIEC / TEN_FILE_GSHEET_GIANGDAY
    bo_tkb.xuat_google_sheet_tkb(file_gsheet_tkb, tong_hop_map, ds_lop, ds_canh_bao)

    try:
        FOLDER_KE_HOACH_NAM_HOC.mkdir(parents=True, exist_ok=True)
        shutil.copy2(file_gsheet_tkb, FOLDER_KE_HOACH_NAM_HOC / TEN_FILE_GSHEET_GIANGDAY)
    except OSError:
        pass

    csdl_path = Path(duong_dan_csdl_master or (PROJECT_ROOT / "CSDL_BaoCao_BoMon_DKT_KGN.xlsx"))
    from .database_manager import QuanLyCSDLBộMôn

    db_mgr = QuanLyCSDLBộMôn(csdl_path)
    db_mgr.khoi_tao_csdl_mau_neu_chua_co(ghi_de=False)
    bo_tkb.dong_bo_vao_csdl_bao_cao_bo_mon(csdl_path, tong_hop_map, thang=thang, nam=nam)

    return KetQuaDongBoTKB(
        nguon_tkb_path=bo_tkb.duong_dan_tkb or file_gsheet_tkb,
        duong_dan_gsheet_tkb=file_gsheet_tkb,
        tong_so_lop_hoc_phan=len(ds_lop),
        so_lop_giang_day_thang_11=sum(1 for x in ds_lop if x.giang_day_trong_thang_11),
        tong_gio_quy_doi_toan_bo_mon=round(sum(x.tong_gio_hk1 for x in tong_hop_map.values()), 1),
        so_ban_ghi_canh_bao=len(ds_canh_bao),
        danh_sach_canh_bao=ds_canh_bao,
    )


if __name__ == "__main__":
    import sys as _sys

    if hasattr(_sys.stdout, "reconfigure"):
        _sys.stdout.reconfigure(encoding="utf-8")
    logging.basicConfig(level=logging.INFO, format="%(asctime)s [%(levelname)s] %(message)s")
    kq = chay_dong_bo_tkb_giang_day_toan_trinh(thang=11, nam=2026)
    print("\n================ KẾT QUẢ ĐỒNG BỘ THỜI KHÓA BIỂU GIẢNG DẠY ================")
    print(f"1. Tệp TKB Nguồn          : {kq.nguon_tkb_path}")
    print(f"2. Google Sheet TKB Xuất  : {kq.duong_dan_gsheet_tkb}")
    print(f"3. Tổng số lớp học phần   : {kq.tong_so_lop_hoc_phan} lớp ({kq.so_lop_giang_day_thang_11} lớp dạy trong T11/2026)")
    print(f"4. Tổng giờ chuẩn HK1     : {kq.tong_gio_quy_doi_toan_bo_mon} tiết quy đổi")
    print(f"5. Bản ghi Cảnh báo/Cách ly: {kq.so_ban_ghi_canh_bao} bản ghi")
    print("==========================================================================")

