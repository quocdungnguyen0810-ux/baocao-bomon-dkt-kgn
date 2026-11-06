# -*- coding: utf-8 -*-
"""
Module Quản lý Cơ sở dữ liệu trung tâm (Google Sheets / Excel Master Database)
cho Bộ môn Địa kỹ thuật và Không gian ngầm.

Chức năng:
1. Khởi tạo tệp CSDL chuẩn `CSDL_BaoCao_BoMon_DKT_KGN.xlsx` (dùng trực tiếp trên máy
   hoặc đồng bộ lên thư mục Google Drive `1jawTokjmk-hnBsuoEIHoxYUhSRCEItPd`).
2. Lưu trữ 21 giảng viên, Ngân hàng công việc lặp lại hàng tháng, Kế hoạch Giảng dạy & NCKH,
   Danh sách nhiệm vụ giao trong tháng và Kế hoạch tháng tới.
3. Tự động sao chép công việc lặp lại hàng tháng và chuyển tiếp nhiệm vụ chưa hoàn thành
   kỳ trước (giữ nguyên thời hạn gốc theo quy định tại Thông báo 940/TB-ĐHGTVT).
"""

from datetime import date, datetime
import logging
from pathlib import Path
import shutil
import tempfile
import time
from typing import Dict, List, Optional, Tuple
import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

from .models import (
    BaoCaoCaNhanThang,
    DANH_SACH_GIANG_VIEN_DKT,
    HoatDongChuyenMon,
    KeHoachKyToi,
    NGUOI_NHAN_BAO_CAO_MAC_DINH,
    NhiemVuCongViec,
    ThongTinGiangVien,
    ngay_cuoi_thang,
)

logger = logging.getLogger(__name__)


def _luu_csdl_an_toan(wb: openpyxl.Workbook, duong_dan_csdl: Path) -> Path:
    """
    Lưu workbook CSDL an toàn qua tệp tạm rồi ghi đè nguyên tử (atomic replace),
    tự động thử lại nếu OneDrive đang đồng bộ để tránh PermissionError.
    """
    duong_dan_csdl = Path(duong_dan_csdl)
    duong_dan_csdl.parent.mkdir(parents=True, exist_ok=True)
    fd, tmp_str = tempfile.mkstemp(suffix=".xlsx", prefix="tmp_csdl_", dir=str(duong_dan_csdl.parent))
    import os
    os.close(fd)
    tmp_path = Path(tmp_str)
    try:
        wb.save(tmp_path)
        for attempt in range(3):
            try:
                shutil.move(str(tmp_path), str(duong_dan_csdl))
                return duong_dan_csdl
            except PermissionError:
                if attempt < 2:
                    time.sleep(0.25)
                else:
                    shutil.copyfile(str(tmp_path), str(duong_dan_csdl))
                    return duong_dan_csdl
    finally:
        if tmp_path.exists():
            try:
                tmp_path.unlink()
            except OSError:
                pass
    return duong_dan_csdl


def _parse_date(val: object) -> Optional[date]:
    """Chuyển đổi giá trị ô Excel thành đối tượng datetime.date an toàn."""
    if val is None or val == "":
        return None
    if isinstance(val, datetime):
        return val.date()
    if isinstance(val, date):
        return val
    s = str(val).strip()
    for fmt in ("%Y-%m-%d", "%d/%m/%Y", "%d-%m-%Y", "%Y/%m/%d"):
        try:
            return datetime.strptime(s[:10], fmt).date()
        except ValueError:
            continue
    return None


class QuanLyCSDLBộMôn:
    """
    Lớp quản lý đọc/ghi và tự động hóa kế thừa dữ liệu trên tệp CSDL Master
    `CSDL_BaoCao_BoMon_DKT_KGN.xlsx` (tương thích 100% với Google Sheets).
    """

    SHEET_GIANG_VIEN = "DM_GiangVien"
    SHEET_VIEC_LAP_LAI = "CongViec_LapLai"
    SHEET_CHUYEN_MON = "KH_GiangDay_NCKH"
    SHEET_NHIEM_VU = "NhiemVu_Thang"
    SHEET_KE_HOACH_TOI = "KeHoach_ThangToi"

    def __init__(self, duong_dan_csdl: Path) -> None:
        self.duong_dan_csdl: Path = Path(duong_dan_csdl)

    def khoi_tao_csdl_mau_neu_chua_co(self, ghi_de: bool = False) -> Path:
        """
        Tạo tệp `CSDL_BaoCao_BoMon_DKT_KGN.xlsx` chứa đầy đủ 21 giảng viên Bộ môn
        Địa kỹ thuật và Không gian ngầm cùng dữ liệu mẫu chuẩn thực tế của Bộ môn.
        """
        if self.duong_dan_csdl.exists() and not ghi_de:
            logger.info("Tệp CSDL đã tồn tại tại: %s", self.duong_dan_csdl)
            return self.duong_dan_csdl

        self.duong_dan_csdl.parent.mkdir(parents=True, exist_ok=True)
        wb = openpyxl.Workbook()
        # Xóa sheet mặc định
        wb.remove(wb.active)

        header_font = Font(name="Times New Roman", size=11, bold=True, color="FFFFFF")
        header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
        body_font = Font(name="Times New Roman", size=11)
        thin_border = Border(
            left=Side(style="thin", color="D9D9D9"),
            right=Side(style="thin", color="D9D9D9"),
            top=Side(style="thin", color="D9D9D9"),
            bottom=Side(style="thin", color="D9D9D9"),
        )

        def _format_sheet(ws: openpyxl.worksheet.worksheet.Worksheet, headers: List[str]) -> None:
            ws.append(headers)
            for col_idx in range(1, len(headers) + 1):
                cell = ws.cell(1, col_idx)
                cell.font = header_font
                cell.fill = header_fill
                cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
                ws.column_dimensions[get_column_letter(col_idx)].width = max(len(headers[col_idx - 1]) + 5, 16)
            ws.row_dimensions[1].height = 28

        # 1. Sheet DM_GiangVien (21 giảng viên Bộ môn ĐKT & KGN)
        ws_gv = wb.create_sheet(self.SHEET_GIANG_VIEN)
        _format_sheet(
            ws_gv,
            [
                "Mã cán bộ",
                "Họ và tên",
                "Ngày sinh",
                "Học hàm",
                "Học vị",
                "Chức danh vị trí",
                "Tài khoản uSmart",
                "Email",
                "Người nhận báo cáo",
                "Định mức Giảng dạy (giờ)",
                "Định mức NCKH (giờ)",
                "Định mức NV khác (giờ)",
            ],
        )
        for gv in DANH_SACH_GIANG_VIEN_DKT:
            ws_gv.append(
                [
                    gv.ma_can_bo,
                    gv.ho_va_ten,
                    gv.ngay_sinh,
                    gv.hoc_ham,
                    gv.hoc_vi,
                    gv.chuc_danh_vi_tri,
                    gv.tai_khoan_usmart,
                    gv.email,
                    gv.nguoi_nhan_bao_cao,
                    gv.dinh_muc_giang_day,
                    gv.dinh_muc_nckh,
                    gv.dinh_muc_nhiem_vu_khac,
                ]
            )

        # 2. Sheet CongViec_LapLai (Ngân hàng công việc lặp lại hàng tháng của GV Bộ môn ĐKT & KGN)
        ws_lap = wb.create_sheet(self.SHEET_VIEC_LAP_LAI)
        _format_sheet(
            ws_lap,
            [
                "Mã mẫu",
                "Áp dụng cho Mã CB (ALL hoặc Mã CB)",
                "Nhiệm vụ, công việc lặp lại hàng tháng",
                "Nguồn giao",
                "Số VB / Mã việc mặc định",
                "Sản phẩm đầu ra, kết quả",
                "Ngày hạn trong tháng (1..31, 0=Cuối tháng)",
                "Nhiệm vụ trọng tâm ĐV (TRUE/FALSE)",
                "Đang kích hoạt (TRUE/FALSE)",
            ],
        )
        mau_lap_lai = [
            (
                "LAP_001",
                "ALL",
                "Giảng dạy các học phần chuyên ngành Địa kỹ thuật, Cơ học đất, Nền móng theo thời khóa biểu học kỳ",
                "Kế hoạch năm học",
                "TKB-HK1-2026",
                "Hoàn thành 100% số tiết giảng trên giảng đường và cập nhật sổ đầu bài điện tử đúng quy định",
                0,
                True,
                True,
            ),
            (
                "LAP_002",
                "ALL",
                "Hướng dẫn đồ án môn học Nền móng, Đồ án tốt nghiệp kỹ sư và Luận văn thạc sĩ chuyên ngành Địa kỹ thuật",
                "Kế hoạch năm học",
                "KH-BM-DKT",
                "Phiếu giao nhiệm vụ đồ án, nhật ký duyệt đồ án hàng tuần và bản thuyết minh sinh viên/học viên",
                0,
                False,
                True,
            ),
            (
                "LAP_003",
                "ALL",
                "Tham gia sinh hoạt chuyên môn, seminar học thuật định kỳ và họp giao ban Bộ môn Địa kỹ thuật & KGN",
                "Chức năng, nhiệm vụ thường xuyên",
                "BB-BM-DKT",
                "Biên bản họp Bộ môn và tài liệu thảo luận chuyên môn",
                0,
                False,
                True,
            ),
            (
                "LAP_004",
                "1051",
                "Quản lý, vận hành trang thiết bị và hướng dẫn thí nghiệm Cơ học đất - Địa kỹ thuật tại Phòng thí nghiệm Bộ môn",
                "Chức năng, nhiệm vụ thường xuyên",
                "KH-PTN-DKT",
                "Sổ nhật ký vận hành thiết bị PTN, phiếu kết quả thí nghiệm mẫu đất/đá của sinh viên và dự án",
                25,
                True,
                True,
            ),
            (
                "LAP_005",
                "1051",
                "Rà soát, cập nhật hồ sơ Đề án mở ngành Địa kỹ thuật và Không gian ngầm, chuẩn hóa danh mục năng lực Phòng thí nghiệm",
                "Kết luận giao ban",
                "DA-DKT-2026",
                "Báo cáo rà soát năng lực PTN và thuyết minh Đề án mở ngành cập nhật",
                28,
                True,
                True,
            ),
            (
                "LAP_0537_01",
                "0537",
                "Thực hiện NCKH và viết báo NCKH",
                "Kế hoạch năm học",
                "KH-NCKH-0537",
                "Báo cáo chuyên đề NCKH và bản thảo bài báo khoa học chuyên ngành Địa kỹ thuật – Công trình",
                28,
                False,
                True,
            ),
            (
                "LAP_0537_02",
                "0537",
                "Xâm nhập thực tiễn, lao động sản xuất",
                "Kế hoạch năm học",
                "KH-TT-0537",
                "Hồ sơ, số liệu khảo sát thực tế hiện trường/phòng thí nghiệm và báo cáo kết quả xâm nhập thực tiễn, lao động sản xuất",
                28,
                False,
                True,
            ),
            (
                "LAP_0537_03",
                "0537",
                "Bảo trì máy móc thí nghiệm định kỳ",
                "Chức năng, nhiệm vụ thường xuyên",
                "PTN-BT-0537",
                "Sổ nhật ký vận hành, biên bản kiểm tra và bảo trì định kỳ hệ thống máy móc, thiết bị thí nghiệm Địa kỹ thuật – Công trình đảm bảo hoạt động tốt",
                25,
                True,
                True,
            ),
            (
                "LAP_0537_04",
                "0537",
                "Hỗ trợ thí nghiệm của NCKH SV, NCS",
                "Kế hoạch năm học",
                "PTN-NCKH-0537",
                "Bộ số liệu và phiếu kết quả thí nghiệm mẫu đất/đá/vật liệu hỗ trợ các đề tài NCKH Sinh viên và Nghiên cứu sinh",
                28,
                True,
                True,
            ),
        ]
        for row in mau_lap_lai:
            ws_lap.append(list(row))

        # 3. Sheet KH_GiangDay_NCKH (Dữ liệu chuyên môn & tự chấm điểm theo tháng cho 21 GV)
        ws_cm = wb.create_sheet(self.SHEET_CHUYEN_MON)
        _format_sheet(
            ws_cm,
            [
                "Tháng",
                "Năm",
                "Mã cán bộ",
                "Họ và tên",
                "Giờ giảng dạy trong tháng",
                "Giờ giảng dạy lũy kế",
                "Định mức giảng dạy",
                "Ghi chú giảng dạy",
                "Giờ NCKH trong tháng",
                "Giờ NCKH lũy kế",
                "Định mức NCKH",
                "Ghi chú NCKH",
                "Giờ nhiệm vụ khác trong tháng",
                "Giờ nhiệm vụ khác lũy kế",
                "Định mức nhiệm vụ khác",
                "Ghi chú nhiệm vụ khác",
                "Sản phẩm KHCN & Hướng dẫn người học trong tháng",
                "Điểm tự chấm Tiêu chí chung (<=30)",
                "Điểm tự chấm Kết quả NV (<=70)",
                "Điểm TBM/ĐV chấm Tiêu chí chung (<=30)",
                "Điểm TBM/ĐV chấm Kết quả NV (<=70)",
                "Đề xuất, kiến nghị",
            ],
        )
        for gv in DANH_SACH_GIANG_VIEN_DKT:
            if gv.ma_can_bo == "0537":
                gd_thang = 28.0
                gd_luy_ke = 92.0
                nc_thang = 10.0
                nc_luy_ke = 30.0
            else:
                gd_thang = 36.0 if gv.ma_can_bo != "0694" else 30.0
                gd_luy_ke = 112.0 if gv.ma_can_bo != "0694" else 95.0
                nc_thang = 30.0
                nc_luy_ke = 95.0
            khac_thang = 20.0
            khac_luy_ke = 60.0
            if gv.ma_can_bo == "1051":
                sp_khcn = (
                    "Hướng dẫn 05 SV làm Đồ án tốt nghiệp Địa kỹ thuật; hoàn thiện 01 bản thảo bài báo Tạp chí KH GTVT; "
                    "tham gia biên soạn hồ sơ mở ngành Địa kỹ thuật & KGN."
                )
            elif gv.ma_can_bo == "0537":
                sp_khcn = (
                    "Hướng dẫn thực hành thí nghiệm Cơ học đất, Cơ học đất đá, Địa chất công trình cho các lớp chính quy; "
                    "phối hợp xử lý số liệu thí nghiệm phục vụ đề tài NCKH của Bộ môn."
                )
            else:
                sp_khcn = "Hướng dẫn đồ án môn học Nền móng, đồ án tốt nghiệp và thực hiện đề tài NCKH cấp Trường."
            ws_cm.append(
                [
                    11,
                    2026,
                    gv.ma_can_bo,
                    gv.ho_va_ten,
                    gd_thang,
                    gd_luy_ke,
                    gv.dinh_muc_giang_day,
                    "Theo TKB học kỳ I năm học 2026-2027",
                    nc_thang,
                    nc_luy_ke,
                    gv.dinh_muc_nckh,
                    "Thực hiện đề tài NCKH và viết bài báo khoa học",
                    khac_thang,
                    khac_luy_ke,
                    gv.dinh_muc_nhiem_vu_khac,
                    "Công tác coi chấm thi, trực PTN, sinh hoạt chuyên môn",
                    sp_khcn,
                    29.0,
                    70.0,
                    None,
                    None,
                    "",
                ]
            )

        # 4. Sheet NhiemVu_Thang (Danh sách nhiệm vụ chi tiết tháng 11/2026)
        ws_nv = wb.create_sheet(self.SHEET_NHIEM_VU)
        _format_sheet(
            ws_nv,
            [
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
                "Khó khăn, vướng mắc, đề xuất",
                "Nhiệm vụ trọng tâm ĐV",
                "Là việc lặp lại",
                "Mức ưu tiên",
                "Đánh giá chất lượng",
                "Link minh chứng Drive",
                "Nguồn dữ liệu",
            ],
        )

        # Khởi tạo nhiệm vụ mẫu đầy đủ cho giảng viên Nguyễn Hải Hà (1051) và các giảng viên chủ chốt trong Bộ môn
        ds_nv_mau = [
            (
                "NV_202611_001",
                11,
                2026,
                "1051",
                "Nguyễn Hải Hà",
                "Các GV Bộ môn",
                "Giảng dạy các học phần Cơ học đất, Nền móng và Địa kỹ thuật theo thời khóa biểu học kỳ I năm học 2026-2027",
                "Kế hoạch năm học",
                "TKB-HK1-2026",
                "Hoàn thành 36 giờ chuẩn giảng dạy trong tháng 11/2026, cập nhật đầy đủ lịch trình giảng dạy",
                date(2026, 11, 1),
                date(2026, 11, 30),
                date(2026, 11, 28),
                "Hoàn thành đúng hạn",
                100.0,
                "",
                True,
                True,
                "Cao",
                "Đạt yêu cầu – tốt",
                "https://drive.google.com/drive/folders/1jawTokjmk-hnBsuoEIHoxYUhSRCEItPd",
                "Kế hoạch Giảng dạy",
            ),
            (
                "NV_202611_002",
                11,
                2026,
                "1051",
                "Nguyễn Hải Hà",
                "GS.TS. Trần Thế Truyền, PGS.TS. Nguyễn Châu Lân",
                "Rà soát, chuẩn hóa dữ liệu năng lực Phòng thí nghiệm Địa kỹ thuật từ hệ thống uSmart phục vụ Đề án mở ngành Địa kỹ thuật",
                "Kết luận giao ban",
                "GV001",
                "Báo cáo tổng hợp năng lực thiết bị PTN và phụ lục thuyết minh Đề án mở ngành bản hoàn thiện 31/10 - 11/2026",
                date(2026, 11, 1),
                date(2026, 11, 15),
                date(2026, 11, 14),
                "Hoàn thành đúng hạn",
                100.0,
                "",
                True,
                False,
                "Khẩn",
                "Đạt yêu cầu – tốt",
                "https://drive.google.com/drive/folders/1jawTokjmk-hnBsuoEIHoxYUhSRCEItPd",
                "Sổ giao việc Mẫu 03",
            ),
            (
                "NV_202611_003",
                11,
                2026,
                "1051",
                "Nguyễn Hải Hà",
                "Nhóm NCKH Bộ môn",
                "Thực hiện nghiên cứu khoa học, viết bài báo chuyên ngành Địa kỹ thuật và hướng dẫn 05 sinh viên làm Đồ án tốt nghiệp",
                "Kế hoạch năm học",
                "KH-NCKH-2026",
                "Bản thảo bài báo khoa học và 05 cuốn thuyết minh tiến độ Đồ án tốt nghiệp của sinh viên",
                date(2026, 11, 2),
                date(2026, 11, 28),
                date(2026, 11, 27),
                "Hoàn thành đúng hạn",
                100.0,
                "",
                False,
                True,
                "Cao",
                "Đạt yêu cầu – tốt",
                "https://drive.google.com/drive/folders/1jawTokjmk-hnBsuoEIHoxYUhSRCEItPd",
                "Kế hoạch NCKH",
            ),
            (
                "NV_202611_004",
                11,
                2026,
                "1051",
                "Nguyễn Hải Hà",
                "KTV Phòng thí nghiệm",
                "Quản lý, kiểm tra định kỳ thiết bị thí nghiệm Cơ học đất và hướng dẫn các ca thực hành thí nghiệm cho sinh viên Khoa Công trình",
                "Chức năng, nhiệm vụ thường xuyên",
                "GV002",
                "Sổ theo dõi ca thực hành thí nghiệm tháng 11/2026 và biên bản kiểm tra tình trạng thiết bị",
                date(2026, 11, 3),
                date(2026, 11, 25),
                date(2026, 11, 24),
                "Hoàn thành đúng hạn",
                100.0,
                "",
                True,
                True,
                "Thường",
                "Đạt yêu cầu – tốt",
                "https://drive.google.com/drive/folders/1jawTokjmk-hnBsuoEIHoxYUhSRCEItPd",
                "Sổ giao việc Mẫu 03",
            ),
            (
                "NV_202611_005",
                11,
                2026,
                "1051",
                "Nguyễn Hải Hà",
                "Toàn Bộ môn",
                "Xây dựng hệ thống số hóa báo cáo công việc tháng tự động theo Thông báo 940/TB-ĐHGTVT cho Bộ môn Địa kỹ thuật và KGN",
                "Nhiệm vụ đột xuất",
                "940/TB-ĐHGTVT",
                "Bộ công cụ phần mềm tự động xuất Mẫu 01b, 02b, 02c, 03 và đồng bộ trực tiếp lên cổng uSmart",
                date(2026, 11, 2),
                date(2026, 11, 10),
                date(2026, 11, 6),
                "Hoàn thành đúng hạn",
                100.0,
                "",
                True,
                False,
                "Khẩn",
                "Đạt yêu cầu – tốt",
                "https://drive.google.com/drive/folders/1jawTokjmk-hnBsuoEIHoxYUhSRCEItPd",
                "Sổ giao việc Mẫu 03",
            ),
        ]
        # Tạo thêm nhiệm vụ chuẩn cho các thành viên còn lại để phục vụ tổng hợp Mẫu 02b, 02c, 03c
        stt_nv = 6
        for gv in DANH_SACH_GIANG_VIEN_DKT:
            if gv.ma_can_bo == "1051":
                continue
            if gv.ma_can_bo == "0537":
                nv_giang_list = [
                    (
                        f"NV_202611_{stt_nv:03d}",
                        11,
                        2026,
                        gv.ma_can_bo,
                        gv.ho_va_ten,
                        "Bộ môn ĐKT & KGN, PTN Công trình",
                        f"Hướng dẫn thực hành, thí nghiệm các học phần Cơ học đất, Cơ học đất đá, Địa chất công trình cho sinh viên hệ chính quy tháng 11/2026 ({gv.ho_va_ten})",
                        "Kế hoạch năm học",
                        "TKB-PTN-2026",
                        "Hoàn thành 100% các ca hướng dẫn thí nghiệm theo lịch phân công, ký xác nhận sổ nhật ký phòng thí nghiệm và phiếu kết quả thí nghiệm của sinh viên",
                        date(2026, 11, 1),
                        date(2026, 11, 30),
                        date(2026, 11, 28),
                        "Hoàn thành đúng hạn",
                        100.0,
                        "",
                        True,
                        True,
                        "Cao",
                        "Đạt yêu cầu – tốt",
                        "",
                        "Kế hoạch Giảng dạy & PTN",
                    ),
                    (
                        f"NV_202611_{stt_nv + 1:03d}",
                        11,
                        2026,
                        gv.ma_can_bo,
                        gv.ho_va_ten,
                        "Nhóm NCKH Bộ môn ĐKT & KGN",
                        "Thực hiện NCKH và viết báo NCKH",
                        "Kế hoạch năm học",
                        "KH-NCKH-0537",
                        "Báo cáo chuyên đề NCKH và bản thảo bài báo khoa học chuyên ngành Địa kỹ thuật – Công trình",
                        date(2026, 11, 1),
                        date(2026, 11, 28),
                        date(2026, 11, 27),
                        "Hoàn thành đúng hạn",
                        100.0,
                        "",
                        False,
                        True,
                        "Cao",
                        "Đạt yêu cầu – tốt",
                        "",
                        "Kế hoạch NCKH hàng tháng",
                    ),
                    (
                        f"NV_202611_{stt_nv + 2:03d}",
                        11,
                        2026,
                        gv.ma_can_bo,
                        gv.ho_va_ten,
                        "PTN Công trình & Bộ môn ĐKT",
                        "Xâm nhập thực tiễn, lao động sản xuất",
                        "Kế hoạch năm học",
                        "KH-TT-0537",
                        "Hồ sơ, số liệu khảo sát thực tế hiện trường/phòng thí nghiệm và báo cáo kết quả xâm nhập thực tiễn, lao động sản xuất",
                        date(2026, 11, 2),
                        date(2026, 11, 28),
                        date(2026, 11, 27),
                        "Hoàn thành đúng hạn",
                        100.0,
                        "",
                        False,
                        True,
                        "Thường",
                        "Đạt yêu cầu – tốt",
                        "",
                        "Nhiệm vụ hàng tháng",
                    ),
                    (
                        f"NV_202611_{stt_nv + 3:03d}",
                        11,
                        2026,
                        gv.ma_can_bo,
                        gv.ho_va_ten,
                        "PTN Địa kỹ thuật - Công trình",
                        "Bảo trì máy móc thí nghiệm định kỳ",
                        "Chức năng, nhiệm vụ thường xuyên",
                        "PTN-BT-0537",
                        "Sổ nhật ký vận hành, biên bản kiểm tra và bảo trì định kỳ hệ thống máy móc, thiết bị thí nghiệm Địa kỹ thuật – Công trình đảm bảo hoạt động tốt",
                        date(2026, 11, 1),
                        date(2026, 11, 25),
                        date(2026, 11, 24),
                        "Hoàn thành đúng hạn",
                        100.0,
                        "",
                        True,
                        True,
                        "Cao",
                        "Đạt yêu cầu – tốt",
                        "",
                        "Công tác Phòng thí nghiệm hàng tháng",
                    ),
                    (
                        f"NV_202611_{stt_nv + 4:03d}",
                        11,
                        2026,
                        gv.ma_can_bo,
                        gv.ho_va_ten,
                        "Nhóm NCKH Bộ môn ĐKT & KGN",
                        "Hỗ trợ thí nghiệm của NCKH SV, NCS",
                        "Kế hoạch năm học",
                        "PTN-NCKH-0537",
                        "Bộ số liệu và phiếu kết quả thí nghiệm mẫu đất/đá/vật liệu hỗ trợ các đề tài NCKH Sinh viên và Nghiên cứu sinh",
                        date(2026, 11, 3),
                        date(2026, 11, 28),
                        date(2026, 11, 27),
                        "Hoàn thành đúng hạn",
                        100.0,
                        "",
                        True,
                        True,
                        "Cao",
                        "Đạt yêu cầu – tốt",
                        "",
                        "Hỗ trợ NCKH SV & NCS hàng tháng",
                    ),
                ]
                ds_nv_mau.extend(nv_giang_list)
                stt_nv += 5
                continue
            ds_nv_mau.append(
                (
                    f"NV_202611_{stt_nv:03d}",
                    11,
                    2026,
                    gv.ma_can_bo,
                    gv.ho_va_ten,
                    "Bộ môn ĐKT & KGN",
                    f"Giảng dạy các học phần chuyên ngành Địa kỹ thuật & KGN và hướng dẫn đồ án tháng 11/2026 ({gv.ho_va_ten})",
                    "Kế hoạch năm học",
                    f"GV{stt_nv:03d}",
                    "Hoàn thành 100% giờ giảng theo TKB, biên bản chấm và duyệt đồ án sinh viên",
                    date(2026, 11, 1),
                    date(2026, 11, 30),
                    date(2026, 11, 28),
                    "Hoàn thành đúng hạn",
                    100.0,
                    "",
                    gv.ma_can_bo in ("0694", "0882", "0069"),
                    True,
                    "Cao",
                    "Đạt yêu cầu – tốt",
                    "",
                    "Kế hoạch Giảng dạy",
                )
            )
            stt_nv += 1
            ds_nv_mau.append(
                (
                    f"NV_202611_{stt_nv:03d}",
                    11,
                    2026,
                    gv.ma_can_bo,
                    gv.ho_va_ten,
                    "Nhóm NCKH",
                    f"Thực hiện đề tài NCKH, viết bài báo khoa học và biên soạn đề cương học phần ({gv.ho_va_ten})",
                    "Kế hoạch năm học",
                    f"GV{stt_nv:03d}",
                    "Báo cáo chuyên đề NCKH và hồ sơ đề cương chi tiết học phần đã rà soát",
                    date(2026, 11, 5),
                    date(2026, 11, 28),
                    date(2026, 11, 27),
                    "Hoàn thành đúng hạn",
                    100.0,
                    "",
                    False,
                    True,
                    "Thường",
                    "Đạt yêu cầu – tốt",
                    "",
                    "Kế hoạch NCKH",
                )
            )
            stt_nv += 1

        for row in ds_nv_mau:
            ws_nv.append(list(row))

        # 5. Sheet KeHoach_ThangToi (Mục IV Mẫu 01b & Mẫu 02b)
        ws_kh = wb.create_sheet(self.SHEET_KE_HOACH_TOI)
        _format_sheet(
            ws_kh,
            [
                "Tháng báo cáo",
                "Năm báo cáo",
                "Mã CB (ALL hoặc Mã CB)",
                "Nhiệm vụ, công việc tháng tới",
                "Sản phẩm dự kiến",
                "Thời hạn",
                "Đơn vị phối hợp",
                "Ghi chú",
            ],
        )
        kh_mau = [
            (
                11,
                2026,
                "ALL",
                "Giảng dạy các học phần cuối học kỳ I năm học 2026-2027, ra đề thi và tổ chức coi chấm thi kết thúc học phần",
                "Ngân hàng đề thi, đáp án và bảng điểm học phần đúng hạn",
                date(2026, 12, 31),
                "Phòng Đào tạo Đại học, Khoa Công trình",
                "Theo kế hoạch đào tạo năm học",
            ),
            (
                11,
                2026,
                "ALL",
                "Hướng dẫn sinh viên bảo vệ Đồ án môn học Nền móng và kiểm tra tiến độ Đồ án tốt nghiệp, Luận văn cao học",
                "Biên bản chấm đồ án môn học và phiếu nhận xét tiến độ ĐATN",
                date(2026, 12, 25),
                "Bộ môn ĐKT & KGN",
                "Lặp lại định kỳ",
            ),
            (
                11,
                2026,
                "1051",
                "Hoàn thiện hồ sơ minh chứng Phòng thí nghiệm Địa kỹ thuật phục vụ đoàn khảo sát mở ngành và bảo trì thiết bị cuối năm",
                "Bộ hồ sơ minh chứng PTN chuẩn hóa và biên bản kiểm định thiết bị năm 2026",
                date(2026, 12, 20),
                "Phòng QLCSVC&CĐS, Khoa Công trình",
                "Nhiệm vụ trọng tâm Bộ môn",
            ),
            (
                11,
                2026,
                "1051",
                "Vận hành chính thức hệ thống tổng hợp báo cáo công việc tháng 11/2026 cho Bộ môn trên uSmart",
                "100% báo cáo Mẫu 01b, 02b, 02c của Bộ môn nộp đúng hạn trước ngày 02-04/12/2026",
                date(2026, 12, 4),
                "Bộ môn ĐKT & KGN",
                "Theo Thông báo 940/TB-ĐHGTVT",
            ),
        ]
        for row in kh_mau:
            ws_kh.append(list(row))

        # Định dạng phông chữ và viền cho toàn bộ các sheet
        for ws in wb.worksheets:
            for r in range(2, ws.max_row + 1):
                for c in range(1, ws.max_column + 1):
                    cell = ws.cell(r, c)
                    cell.font = body_font
                    cell.border = thin_border
                    if isinstance(cell.value, date):
                        cell.number_format = "dd/mm/yyyy"

        _luu_csdl_an_toan(wb, self.duong_dan_csdl)
        logger.info("Đã khởi tạo thành công CSDL Bộ môn tại: %s", self.duong_dan_csdl)
        return self.duong_dan_csdl

    def doc_danh_sach_giang_vien(self) -> Dict[str, ThongTinGiangVien]:
        """
        Đọc danh sách giảng viên chính thức thuộc Bộ môn ĐKT & KGN từ Sheet DM_GiangVien,
        tự động lọc bỏ các mã cán bộ ngoại lai (0799, 1450, 1451) và đồng bộ bổ sung
        các thành viên chuẩn từ DANH_SACH_GIANG_VIEN_DKT (bao gồm 0537 - Nguyễn Thị Thanh Giang).
        """
        self.khoi_tao_csdl_mau_neu_chua_co()
        map_chuan = {gv.ma_can_bo: gv for gv in DANH_SACH_GIANG_VIEN_DKT}
        wb = openpyxl.load_workbook(self.duong_dan_csdl, data_only=True)
        ws = wb[self.SHEET_GIANG_VIEN]
        ket_qua_raw: Dict[str, ThongTinGiangVien] = {}
        for r in range(2, ws.max_row + 1):
            ma_cb = str(ws.cell(r, 1).value or "").strip()
            if not ma_cb or ma_cb not in map_chuan:
                continue
            gv_chuan = map_chuan[ma_cb]
            tk_sheet = str(ws.cell(r, 7).value or "").strip()
            # Ưu tiên tài khoản uSmart chuẩn trong models.py nếu sheet đang lưu tên cũ (như linhnt -> linhnt2)
            tk_usmart = gv_chuan.tai_khoan_usmart if gv_chuan.tai_khoan_usmart else tk_sheet
            email_sheet = str(ws.cell(r, 8).value or "").strip()
            email_val = gv_chuan.email if gv_chuan.email else email_sheet
            gv = ThongTinGiangVien(
                ma_can_bo=ma_cb,
                ho_va_ten=str(ws.cell(r, 2).value or gv_chuan.ho_va_ten).strip(),
                ngay_sinh=str(ws.cell(r, 3).value or gv_chuan.ngay_sinh).strip(),
                hoc_ham=str(ws.cell(r, 4).value or gv_chuan.hoc_ham).strip(),
                hoc_vi=str(ws.cell(r, 5).value or gv_chuan.hoc_vi).strip(),
                chuc_danh_vi_tri=str(ws.cell(r, 6).value or gv_chuan.chuc_danh_vi_tri).strip(),
                tai_khoan_usmart=tk_usmart,
                email=email_val,
                nguoi_nhan_bao_cao=str(
                    ws.cell(r, 9).value or gv_chuan.nguoi_nhan_bao_cao or NGUOI_NHAN_BAO_CAO_MAC_DINH
                ).strip(),
                dinh_muc_giang_day=float(ws.cell(r, 10).value or gv_chuan.dinh_muc_giang_day),
                dinh_muc_nckh=float(ws.cell(r, 11).value or gv_chuan.dinh_muc_nckh),
                dinh_muc_nhiem_vu_khac=float(ws.cell(r, 12).value or gv_chuan.dinh_muc_nhiem_vu_khac),
            )
            ket_qua_raw[ma_cb] = gv
        wb.close()

        # Giữ đúng thứ tự chuẩn của DANH_SACH_GIANG_VIEN_DKT và bổ sung thành viên mới nếu file CSDL cũ chưa có
        ket_qua: Dict[str, ThongTinGiangVien] = {}
        for gv_chuan in DANH_SACH_GIANG_VIEN_DKT:
            ket_qua[gv_chuan.ma_can_bo] = ket_qua_raw.get(gv_chuan.ma_can_bo, gv_chuan)
        return ket_qua

    def dong_bo_danh_sach_giang_vien_chuan(self, thang: int = 11, nam: int = 2026) -> int:
        """
        Đồng bộ danh sách giảng viên chuẩn từ `DANH_SACH_GIANG_VIEN_DKT` vào tệp CSDL hiện có:
        1. Loại bỏ triệt để các bản ghi ngoại lai không thuộc Bộ môn (như 0799 - Nguyễn Thị Định,
           1450 - Ngô Thùy Dương, 1451 - Nguyễn Hữu Thuấn) trên toàn bộ các sheet.
        2. Cập nhật tài khoản uSmart mới nhất (`linhnt2`, `ducla`, `giangntt`...) vào `DM_GiangVien`.
        3. Tự động khởi tạo dữ liệu chuyên môn và nhiệm vụ chuẩn cho thành viên mới bổ sung
           (`0537 - Nguyễn Thị Thanh Giang`) nếu chưa có trong kỳ `(thang, nam)`.
        """
        self.khoi_tao_csdl_mau_neu_chua_co()
        map_chuan = {gv.ma_can_bo: gv for gv in DANH_SACH_GIANG_VIEN_DKT}
        wb = openpyxl.load_workbook(self.duong_dan_csdl)

        # 1. Cập nhật lại Sheet DM_GiangVien theo đúng DANH_SACH_GIANG_VIEN_DKT
        ws_gv = wb[self.SHEET_GIANG_VIEN]
        if ws_gv.max_row >= 2:
            ws_gv.delete_rows(2, ws_gv.max_row - 1)
        for gv in DANH_SACH_GIANG_VIEN_DKT:
            ws_gv.append(
                [
                    gv.ma_can_bo,
                    gv.ho_va_ten,
                    gv.ngay_sinh,
                    gv.hoc_ham,
                    gv.hoc_vi,
                    gv.chuc_danh_vi_tri,
                    gv.tai_khoan_usmart,
                    gv.email,
                    gv.nguoi_nhan_bao_cao,
                    gv.dinh_muc_giang_day,
                    gv.dinh_muc_nckh,
                    gv.dinh_muc_nhiem_vu_khac,
                ]
            )

        # 2. Lọc bỏ các mã cán bộ ngoại lai trên Sheet KH_GiangDay_NCKH và bổ sung thành viên còn thiếu
        ws_cm = wb[self.SHEET_CHUYEN_MON]
        rows_cm_hop_le = []
        mcb_da_co_cm = set()
        for r in range(2, ws_cm.max_row + 1):
            t = int(ws_cm.cell(r, 1).value or 0)
            n = int(ws_cm.cell(r, 2).value or 0)
            mcb = str(ws_cm.cell(r, 3).value or "").strip()
            if mcb in map_chuan:
                rows_cm_hop_le.append([ws_cm.cell(r, c).value for c in range(1, 23)])
                if t == thang and n == nam:
                    mcb_da_co_cm.add(mcb)
        if ws_cm.max_row >= 2:
            ws_cm.delete_rows(2, ws_cm.max_row - 1)
        for row_vals in rows_cm_hop_le:
            ws_cm.append(row_vals)

        for gv in DANH_SACH_GIANG_VIEN_DKT:
            if gv.ma_can_bo not in mcb_da_co_cm:
                gd_thang = 28.0 if gv.ma_can_bo == "0537" else 36.0
                gd_luy_ke = 92.0 if gv.ma_can_bo == "0537" else 112.0
                nc_thang = 10.0 if gv.ma_can_bo == "0537" else 30.0
                nc_luy_ke = 30.0 if gv.ma_can_bo == "0537" else 95.0
                sp_khcn = (
                    "Hướng dẫn thực hành thí nghiệm Cơ học đất, Cơ học đất đá, Địa chất công trình cho các lớp chính quy; "
                    "phối hợp xử lý số liệu thí nghiệm phục vụ đề tài NCKH của Bộ môn."
                    if gv.ma_can_bo == "0537"
                    else "Hướng dẫn đồ án môn học Nền móng, đồ án tốt nghiệp và thực hiện đề tài NCKH cấp Trường."
                )
                ws_cm.append(
                    [
                        thang,
                        nam,
                        gv.ma_can_bo,
                        gv.ho_va_ten,
                        gd_thang,
                        gd_luy_ke,
                        gv.dinh_muc_giang_day,
                        "Theo TKB học kỳ I năm học 2026-2027",
                        nc_thang,
                        nc_luy_ke,
                        gv.dinh_muc_nckh,
                        "Thực hiện đề tài NCKH và thí nghiệm chuyên môn",
                        20.0,
                        60.0,
                        gv.dinh_muc_nhiem_vu_khac,
                        "Công tác coi chấm thi, trực PTN, sinh hoạt chuyên môn",
                        sp_khcn,
                        29.0,
                        70.0,
                        None,
                        None,
                        "",
                    ]
                )

        # 3. Lọc bỏ các mã cán bộ ngoại lai trên Sheet NhiemVu_Thang và bổ sung nhiệm vụ mặc định cho thành viên mới (0537)
        ws_nv = wb[self.SHEET_NHIEM_VU]
        rows_nv_hop_le = []
        mcb_da_co_nv = set()
        for r in range(2, ws_nv.max_row + 1):
            t = int(ws_nv.cell(r, 2).value or 0)
            n = int(ws_nv.cell(r, 3).value or 0)
            mcb = str(ws_nv.cell(r, 4).value or "").strip()
            if mcb in map_chuan:
                rows_nv_hop_le.append([ws_nv.cell(r, c).value for c in range(1, 23)])
                if t == thang and n == nam:
                    mcb_da_co_nv.add(mcb)
        if ws_nv.max_row >= 2:
            ws_nv.delete_rows(2, ws_nv.max_row - 1)
        for row_vals in rows_nv_hop_le:
            ws_nv.append(row_vals)

        so_them_moi = 0
        if "0537" in map_chuan:
            gv_giang = map_chuan["0537"]
            cuoi_ky = ngay_cuoi_thang(thang, nam)
            dau_ky = date(nam, thang, 1)
            ngay_ht_1 = date(nam, thang, min(28, cuoi_ky.day))
            ngay_ht_2 = date(nam, thang, min(24, cuoi_ky.day))
            ngay_ht_3 = date(nam, thang, min(27, cuoi_ky.day))

            # Đồng bộ 4 nhiệm vụ hàng tháng của 0537 vào Sheet CongViec_LapLai nếu chưa có
            ws_lap = wb[self.SHEET_VIEC_LAP_LAI]
            ten_lap_0537_da_co = {
                str(ws_lap.cell(r, 3).value or "").strip().lower()
                for r in range(2, ws_lap.max_row + 1)
                if str(ws_lap.cell(r, 2).value or "").strip() == "0537"
            }
            ds_lap_0537 = [
                (
                    "LAP_0537_01",
                    "0537",
                    "Thực hiện NCKH và viết báo NCKH",
                    "Kế hoạch năm học",
                    "KH-NCKH-0537",
                    "Báo cáo chuyên đề NCKH và bản thảo bài báo khoa học chuyên ngành Địa kỹ thuật – Công trình",
                    28,
                    False,
                    True,
                ),
                (
                    "LAP_0537_02",
                    "0537",
                    "Xâm nhập thực tiễn, lao động sản xuất",
                    "Kế hoạch năm học",
                    "KH-TT-0537",
                    "Hồ sơ, số liệu khảo sát thực tế hiện trường/phòng thí nghiệm và báo cáo kết quả xâm nhập thực tiễn, lao động sản xuất",
                    28,
                    False,
                    True,
                ),
                (
                    "LAP_0537_03",
                    "0537",
                    "Bảo trì máy móc thí nghiệm định kỳ",
                    "Chức năng, nhiệm vụ thường xuyên",
                    "PTN-BT-0537",
                    "Sổ nhật ký vận hành, biên bản kiểm tra và bảo trì định kỳ hệ thống máy móc, thiết bị thí nghiệm Địa kỹ thuật – Công trình đảm bảo hoạt động tốt",
                    25,
                    True,
                    True,
                ),
                (
                    "LAP_0537_04",
                    "0537",
                    "Hỗ trợ thí nghiệm của NCKH SV, NCS",
                    "Kế hoạch năm học",
                    "PTN-NCKH-0537",
                    "Bộ số liệu và phiếu kết quả thí nghiệm mẫu đất/đá/vật liệu hỗ trợ các đề tài NCKH Sinh viên và Nghiên cứu sinh",
                    28,
                    True,
                    True,
                ),
            ]
            for row_l in ds_lap_0537:
                if row_l[2].lower() not in ten_lap_0537_da_co:
                    ws_lap.append(list(row_l))

            ten_nv_0537_da_co = {
                str(ws_nv.cell(r, 7).value or "").strip().lower()
                for r in range(2, ws_nv.max_row + 1)
                if int(ws_nv.cell(r, 2).value or 0) == thang
                and int(ws_nv.cell(r, 3).value or 0) == nam
                and str(ws_nv.cell(r, 4).value or "").strip() == "0537"
            }
            nv_giang_rows = [
                [
                    f"NV_{nam}{thang:02d}_0537_01",
                    thang,
                    nam,
                    "0537",
                    gv_giang.ho_va_ten,
                    "Bộ môn ĐKT & KGN, PTN Công trình",
                    f"Hướng dẫn thực hành, thí nghiệm các học phần Cơ học đất, Cơ học đất đá, Địa chất công trình cho sinh viên hệ chính quy tháng {thang:02d}/{nam} ({gv_giang.ho_va_ten})",
                    "Kế hoạch năm học",
                    "TKB-PTN-2026",
                    "Hoàn thành 100% các ca hướng dẫn thí nghiệm theo lịch phân công, ký xác nhận sổ nhật ký phòng thí nghiệm và phiếu kết quả thí nghiệm của sinh viên",
                    dau_ky,
                    cuoi_ky,
                    ngay_ht_1,
                    "Hoàn thành đúng hạn",
                    100.0,
                    "",
                    True,
                    True,
                    "Cao",
                    "Đạt yêu cầu – tốt",
                    "",
                    "Kế hoạch Giảng dạy & PTN",
                ],
                [
                    f"NV_{nam}{thang:02d}_0537_02",
                    thang,
                    nam,
                    "0537",
                    gv_giang.ho_va_ten,
                    "Nhóm NCKH Bộ môn ĐKT & KGN",
                    "Thực hiện NCKH và viết báo NCKH",
                    "Kế hoạch năm học",
                    "KH-NCKH-0537",
                    "Báo cáo chuyên đề NCKH và bản thảo bài báo khoa học chuyên ngành Địa kỹ thuật – Công trình",
                    dau_ky,
                    ngay_ht_1,
                    ngay_ht_3,
                    "Hoàn thành đúng hạn",
                    100.0,
                    "",
                    False,
                    True,
                    "Cao",
                    "Đạt yêu cầu – tốt",
                    "",
                    "Kế hoạch NCKH hàng tháng",
                ],
                [
                    f"NV_{nam}{thang:02d}_0537_03",
                    thang,
                    nam,
                    "0537",
                    gv_giang.ho_va_ten,
                    "PTN Công trình & Bộ môn ĐKT",
                    "Xâm nhập thực tiễn, lao động sản xuất",
                    "Kế hoạch năm học",
                    "KH-TT-0537",
                    "Hồ sơ, số liệu khảo sát thực tế hiện trường/phòng thí nghiệm và báo cáo kết quả xâm nhập thực tiễn, lao động sản xuất",
                    date(nam, thang, min(2, cuoi_ky.day)),
                    ngay_ht_1,
                    ngay_ht_3,
                    "Hoàn thành đúng hạn",
                    100.0,
                    "",
                    False,
                    True,
                    "Thường",
                    "Đạt yêu cầu – tốt",
                    "",
                    "Nhiệm vụ hàng tháng",
                ],
                [
                    f"NV_{nam}{thang:02d}_0537_04",
                    thang,
                    nam,
                    "0537",
                    gv_giang.ho_va_ten,
                    "PTN Địa kỹ thuật - Công trình",
                    "Bảo trì máy móc thí nghiệm định kỳ",
                    "Chức năng, nhiệm vụ thường xuyên",
                    "PTN-BT-0537",
                    "Sổ nhật ký vận hành, biên bản kiểm tra và bảo trì định kỳ hệ thống máy móc, thiết bị thí nghiệm Địa kỹ thuật – Công trình đảm bảo hoạt động tốt",
                    dau_ky,
                    date(nam, thang, min(25, cuoi_ky.day)),
                    ngay_ht_2,
                    "Hoàn thành đúng hạn",
                    100.0,
                    "",
                    True,
                    True,
                    "Cao",
                    "Đạt yêu cầu – tốt",
                    "",
                    "Công tác Phòng thí nghiệm hàng tháng",
                ],
                [
                    f"NV_{nam}{thang:02d}_0537_05",
                    thang,
                    nam,
                    "0537",
                    gv_giang.ho_va_ten,
                    "Nhóm NCKH Bộ môn ĐKT & KGN",
                    "Hỗ trợ thí nghiệm của NCKH SV, NCS",
                    "Kế hoạch năm học",
                    "PTN-NCKH-0537",
                    "Bộ số liệu và phiếu kết quả thí nghiệm mẫu đất/đá/vật liệu hỗ trợ các đề tài NCKH Sinh viên và Nghiên cứu sinh",
                    date(nam, thang, min(3, cuoi_ky.day)),
                    ngay_ht_1,
                    ngay_ht_3,
                    "Hoàn thành đúng hạn",
                    100.0,
                    "",
                    True,
                    True,
                    "Cao",
                    "Đạt yêu cầu – tốt",
                    "",
                    "Hỗ trợ NCKH SV & NCS hàng tháng",
                ],
            ]
            for r_g in nv_giang_rows:
                if str(r_g[6]).strip().lower() not in ten_nv_0537_da_co:
                    ws_nv.append(r_g)
                    so_them_moi += 1

        _luu_csdl_an_toan(wb, self.duong_dan_csdl)
        wb.close()
        return so_them_moi

    def doc_bao_cao_ca_nhan(
        self, ma_can_bo: str, thang: int = 11, nam: int = 2026
    ) -> BaoCaoCaNhanThang:
        """
        Tạo đối tượng `BaoCaoCaNhanThang` hoàn chỉnh cho một giảng viên từ CSDL
        (kết hợp thông tin cá nhân, nhiệm vụ trong tháng, giờ giảng dạy/NCKH và kế hoạch tháng tới).
        """
        ds_gv = self.doc_danh_sach_giang_vien()
        if ma_can_bo not in ds_gv:
            raise KeyError(f"Không tìm thấy giảng viên có mã cán bộ '{ma_can_bo}' trong CSDL.")
        gv = ds_gv[ma_can_bo]

        wb = openpyxl.load_workbook(self.duong_dan_csdl, data_only=True)

        # 1. Đọc nhiệm vụ trong tháng
        ws_nv = wb[self.SHEET_NHIEM_VU]
        ds_nhiem_vu: List[NhiemVuCongViec] = []
        for r in range(2, ws_nv.max_row + 1):
            t = int(ws_nv.cell(r, 2).value or 0)
            n = int(ws_nv.cell(r, 3).value or 0)
            mcb = str(ws_nv.cell(r, 4).value or "").strip()
            if t == thang and n == nam and mcb == ma_can_bo:
                nv = NhiemVuCongViec(
                    ma_nhiem_vu=str(ws_nv.cell(r, 1).value or "").strip(),
                    nhiem_vu=str(ws_nv.cell(r, 7).value or "").strip(),
                    nguon_giao=str(ws_nv.cell(r, 8).value or "").strip(),
                    so_van_ban_ma_viec=str(ws_nv.cell(r, 9).value or "").strip(),
                    san_pham_ket_qua=str(ws_nv.cell(r, 10).value or "").strip(),
                    ngay_giao=_parse_date(ws_nv.cell(r, 11).value),
                    thoi_han=_parse_date(ws_nv.cell(r, 12).value),
                    ngay_hoan_thanh=_parse_date(ws_nv.cell(r, 13).value),
                    tinh_trang=str(ws_nv.cell(r, 14).value or "").strip(),
                    muc_do_hoan_thanh=float(ws_nv.cell(r, 15).value if ws_nv.cell(r, 15).value is not None else 0.0),
                    kho_khan_de_xuat=str(ws_nv.cell(r, 16).value or "").strip(),
                    ma_can_bo_chu_tri=mcb,
                    ten_can_bo_chu_tri=str(ws_nv.cell(r, 5).value or gv.ho_va_ten).strip(),
                    nguoi_phoi_hop=str(ws_nv.cell(r, 6).value or "").strip(),
                    la_nhiem_vu_trong_tam=bool(ws_nv.cell(r, 17).value),
                    la_cong_viec_lap_lai=bool(ws_nv.cell(r, 18).value),
                    muc_uu_tien=str(ws_nv.cell(r, 19).value or "Thường").strip(),
                    danh_gia_chat_luong=str(ws_nv.cell(r, 20).value or "Đạt yêu cầu – tốt").strip(),
                    link_minh_chung_drive=str(ws_nv.cell(r, 21).value or "").strip(),
                    nguon_du_lieu=str(ws_nv.cell(r, 22).value or "").strip(),
                )
                ds_nhiem_vu.append(nv)

        # 2. Đọc hoạt động chuyên môn & điểm tự chấm
        ws_cm = wb[self.SHEET_CHUYEN_MON]
        hd_cm = HoatDongChuyenMon(
            giang_day_dinh_muc=gv.dinh_muc_giang_day,
            nckh_dinh_muc=gv.dinh_muc_nckh,
            nhiem_vu_khac_dinh_muc=gv.dinh_muc_nhiem_vu_khac,
        )
        tu_cham_chung = 29.0
        tu_cham_nv: Optional[float] = None
        tdv_chung: Optional[float] = None
        tdv_nv: Optional[float] = None
        kien_nghi = ""

        for r in range(2, ws_cm.max_row + 1):
            t = int(ws_cm.cell(r, 1).value or 0)
            n = int(ws_cm.cell(r, 2).value or 0)
            mcb = str(ws_cm.cell(r, 3).value or "").strip()
            if t == thang and n == nam and mcb == ma_can_bo:
                hd_cm = HoatDongChuyenMon(
                    giang_day_thang=float(ws_cm.cell(r, 5).value or 0.0),
                    giang_day_luy_ke=float(ws_cm.cell(r, 6).value or 0.0),
                    giang_day_dinh_muc=float(ws_cm.cell(r, 7).value or gv.dinh_muc_giang_day),
                    giang_day_ghi_chu=str(ws_cm.cell(r, 8).value or "").strip(),
                    nckh_thang=float(ws_cm.cell(r, 9).value or 0.0),
                    nckh_luy_ke=float(ws_cm.cell(r, 10).value or 0.0),
                    nckh_dinh_muc=float(ws_cm.cell(r, 11).value or gv.dinh_muc_nckh),
                    nckh_ghi_chu=str(ws_cm.cell(r, 12).value or "").strip(),
                    nhiem_vu_khac_thang=float(ws_cm.cell(r, 13).value or 0.0),
                    nhiem_vu_khac_luy_ke=float(ws_cm.cell(r, 14).value or 0.0),
                    nhiem_vu_khac_dinh_muc=float(ws_cm.cell(r, 15).value or gv.dinh_muc_nhiem_vu_khac),
                    nhiem_vu_khac_ghi_chu=str(ws_cm.cell(r, 16).value or "").strip(),
                    san_pham_khcn_huong_dan=str(ws_cm.cell(r, 17).value or "").strip(),
                )
                if ws_cm.cell(r, 18).value is not None:
                    tu_cham_chung = float(ws_cm.cell(r, 18).value)
                if ws_cm.cell(r, 19).value is not None:
                    tu_cham_nv = float(ws_cm.cell(r, 19).value)
                if ws_cm.cell(r, 20).value is not None:
                    tdv_chung = float(ws_cm.cell(r, 20).value)
                if ws_cm.cell(r, 21).value is not None:
                    tdv_nv = float(ws_cm.cell(r, 21).value)
                kien_nghi = str(ws_cm.cell(r, 22).value or "").strip()
                break

        # 3. Đọc kế hoạch tháng tới (ưu tiên bản ghi riêng của ma_can_bo nếu đã tùy chỉnh, nếu chưa có thì lấy bản ghi ALL)
        ws_kh = wb[self.SHEET_KE_HOACH_TOI]
        ds_kh_rieng: List[KeHoachKyToi] = []
        ds_kh_chung: List[KeHoachKyToi] = []
        for r in range(2, ws_kh.max_row + 1):
            t = int(ws_kh.cell(r, 1).value or 0)
            n = int(ws_kh.cell(r, 2).value or 0)
            ap_dung = str(ws_kh.cell(r, 3).value or "").strip()
            if t == thang and n == nam and ap_dung in ("ALL", ma_can_bo):
                item_kh = KeHoachKyToi(
                    nhiem_vu=str(ws_kh.cell(r, 4).value or "").strip(),
                    san_pham_du_kien=str(ws_kh.cell(r, 5).value or "").strip(),
                    thoi_han=_parse_date(ws_kh.cell(r, 6).value),
                    don_vi_phoi_hop=str(ws_kh.cell(r, 7).value or "").strip(),
                    ghi_chu=str(ws_kh.cell(r, 8).value or "").strip(),
                    ca_nhan_chu_tri=gv.ho_va_ten,
                )
                if ap_dung == ma_can_bo:
                    ds_kh_rieng.append(item_kh)
                else:
                    ds_kh_chung.append(item_kh)
        ds_kh_toi = ds_kh_rieng if ds_kh_rieng else ds_kh_chung

        wb.close()
        return BaoCaoCaNhanThang(
            giang_vien=gv,
            thang=thang,
            nam=nam,
            danh_sach_nhiem_vu=ds_nhiem_vu,
            hoat_dong_chuyen_mon=hd_cm,
            diem_tu_cham_tieu_chi_chung=tu_cham_chung,
            diem_tu_cham_ket_qua_nhiem_vu=tu_cham_nv,
            diem_truong_don_vi_tieu_chi_chung=tdv_chung,
            diem_truong_don_vi_ket_qua_nhiem_vu=tdv_nv,
            ke_hoach_thang_toi=ds_kh_toi,
            de_xuat_kien_nghi=kien_nghi,
            ngay_ky_bao_cao=ngay_cuoi_thang(thang, nam),
        )

    def doc_toan_bo_bao_cao_bo_mon(
        self, thang: int = 11, nam: int = 2026
    ) -> List[BaoCaoCaNhanThang]:
        """Đọc danh sách báo cáo tháng của toàn bộ 21 giảng viên trong Bộ môn."""
        ds_gv = self.doc_danh_sach_giang_vien()
        return [
            self.doc_bao_cao_ca_nhan(ma_cb, thang=thang, nam=nam)
            for ma_cb in ds_gv
        ]

    def sao_chep_cong_viec_lap_lai_sang_thang_moi(
        self, thang_moi: int, nam_moi: int, ma_can_bo: Optional[str] = None
    ) -> int:
        """
        Tự động sinh nhiệm vụ cho kỳ báo cáo mới từ Sheet `CongViec_LapLai`
        (tránh trùng lặp nếu đã sinh trước đó).
        """
        self.khoi_tao_csdl_mau_neu_chua_co()
        ds_gv = self.doc_danh_sach_giang_vien()
        wb = openpyxl.load_workbook(self.duong_dan_csdl)
        ws_lap = wb[self.SHEET_VIEC_LAP_LAI]
        ws_nv = wb[self.SHEET_NHIEM_VU]

        # Tập hợp các nhiệm vụ đã tồn tại trong tháng mới để chống trùng
        da_co = set()
        for r in range(2, ws_nv.max_row + 1):
            t = int(ws_nv.cell(r, 2).value or 0)
            n = int(ws_nv.cell(r, 3).value or 0)
            mcb = str(ws_nv.cell(r, 4).value or "").strip()
            ten_nv = str(ws_nv.cell(r, 7).value or "").strip()
            if t == thang_moi and n == nam_moi:
                da_co.add((mcb, ten_nv))

        cuoi_thang = ngay_cuoi_thang(thang_moi, nam_moi)
        ngay_dau_thang = date(nam_moi, thang_moi, 1)
        so_ban_ghi_moi = 0

        for r in range(2, ws_lap.max_row + 1):
            kich_hoat = bool(ws_lap.cell(r, 9).value)
            if not kich_hoat:
                continue
            ap_dung = str(ws_lap.cell(r, 2).value or "ALL").strip()
            ten_nv = str(ws_lap.cell(r, 3).value or "").strip()
            nguon_giao = str(ws_lap.cell(r, 4).value or "Chức năng, nhiệm vụ thường xuyên").strip()
            so_vb = str(ws_lap.cell(r, 5).value or "").strip()
            san_pham = str(ws_lap.cell(r, 6).value or "").strip()
            ngay_han_int = int(ws_lap.cell(r, 7).value or 0)
            trong_tam = bool(ws_lap.cell(r, 8).value)

            thoi_han_nv = (
                cuoi_thang
                if ngay_han_int <= 0 or ngay_han_int > cuoi_thang.day
                else date(nam_moi, thang_moi, ngay_han_int)
            )

            danh_sach_mcb = (
                [ma_can_bo]
                if ma_can_bo
                else (list(ds_gv.keys()) if ap_dung == "ALL" else [ap_dung])
            )
            for mcb in danh_sach_mcb:
                if ap_dung not in ("ALL", mcb) or mcb not in ds_gv:
                    continue
                if (mcb, ten_nv) in da_co:
                    continue
                gv = ds_gv[mcb]
                ma_nv_moi = f"NV_{nam_moi}{thang_moi:02d}_LAP_{ws_nv.max_row:03d}"
                ws_nv.append(
                    [
                        ma_nv_moi,
                        thang_moi,
                        nam_moi,
                        mcb,
                        gv.ho_va_ten,
                        "Bộ môn ĐKT & KGN",
                        ten_nv,
                        nguon_giao,
                        so_vb,
                        san_pham,
                        ngay_dau_thang,
                        thoi_han_nv,
                        None,
                        "Đang thực hiện – đúng tiến độ",
                        50.0,
                        "",
                        trong_tam,
                        True,
                        "Thường",
                        "Đạt yêu cầu – tốt",
                        "",
                        "Sao chép việc lặp lại",
                    ]
                )
                da_co.add((mcb, ten_nv))
                so_ban_ghi_moi += 1

        _luu_csdl_an_toan(wb, self.duong_dan_csdl)
        wb.close()
        return so_ban_ghi_moi

    def chuyen_tiep_nhiem_vu_chua_hoan_thanh(
        self,
        thang_truoc: int,
        nam_truoc: int,
        thang_moi: int,
        nam_moi: int,
        ma_can_bo: Optional[str] = None,
    ) -> int:
        """
        Chuyển các nhiệm vụ chưa hoàn thành của kỳ trước sang kỳ báo cáo mới,
        GIỮ NGUYÊN THỜI HẠN BAN ĐẦU theo đúng quy định tại dòng 21 Sheet Hướng dẫn Mẫu 01.
        """
        self.khoi_tao_csdl_mau_neu_chua_co()
        wb = openpyxl.load_workbook(self.duong_dan_csdl)
        ws_nv = wb[self.SHEET_NHIEM_VU]

        da_co = set()
        cac_dong_chua_xong: List[list] = []
        for r in range(2, ws_nv.max_row + 1):
            t = int(ws_nv.cell(r, 2).value or 0)
            n = int(ws_nv.cell(r, 3).value or 0)
            mcb = str(ws_nv.cell(r, 4).value or "").strip()
            ten_nv = str(ws_nv.cell(r, 7).value or "").strip()
            tinh_trang = str(ws_nv.cell(r, 14).value or "").strip()
            ngay_ht = _parse_date(ws_nv.cell(r, 13).value)

            if t == thang_moi and n == nam_moi:
                da_co.add((mcb, ten_nv))
            elif t == thang_truoc and n == nam_truoc:
                if ma_can_bo and mcb != ma_can_bo:
                    continue
                if not tinh_trang.startswith("Hoàn thành") and ngay_ht is None:
                    row_vals = [ws_nv.cell(r, c).value for c in range(1, 23)]
                    cac_dong_chua_xong.append(row_vals)

        so_chuyen_tiep = 0
        for row_vals in cac_dong_chua_xong:
            mcb = str(row_vals[3] or "").strip()
            ten_nv = str(row_vals[6] or "").strip()
            if (mcb, ten_nv) in da_co:
                continue
            row_moi = list(row_vals)
            row_moi[0] = f"NV_{nam_moi}{thang_moi:02d}_CT_{ws_nv.max_row:03d}"
            row_moi[1] = thang_moi
            row_moi[2] = nam_moi
            # Giữ nguyên thời hạn ban đầu tại row_moi[11]
            row_moi[21] = f"Chuyển tiếp từ T{thang_truoc:02d}/{nam_truoc}"
            ws_nv.append(row_moi)
            da_co.add((mcb, ten_nv))
            so_chuyen_tiep += 1

        _luu_csdl_an_toan(wb, self.duong_dan_csdl)
        wb.close()
        return so_chuyen_tiep

    def them_nhiem_vu_moi(
        self, thang: int, nam: int, nv: NhiemVuCongViec
    ) -> str:
        """Thêm một nhiệm vụ mới vào Sheet `NhiemVu_Thang` trong CSDL."""
        self.khoi_tao_csdl_mau_neu_chua_co()
        wb = openpyxl.load_workbook(self.duong_dan_csdl)
        ws_nv = wb[self.SHEET_NHIEM_VU]
        ma_nv = nv.ma_nhiem_vu or f"NV_{nam}{thang:02d}_{ws_nv.max_row:03d}"
        ws_nv.append(
            [
                ma_nv,
                thang,
                nam,
                nv.ma_can_bo_chu_tri,
                nv.ten_can_bo_chu_tri,
                nv.nguoi_phoi_hop,
                nv.nhiem_vu,
                nv.nguon_giao,
                nv.so_van_ban_ma_viec,
                nv.san_pham_ket_qua,
                nv.ngay_giao or date(nam, thang, 1),
                nv.thoi_han,
                nv.ngay_hoan_thanh,
                nv.tinh_trang,
                nv.muc_do_hoan_thanh,
                nv.kho_khan_de_xuat,
                nv.la_nhiem_vu_trong_tam,
                nv.la_cong_viec_lap_lai,
                nv.muc_uu_tien,
                nv.danh_gia_chat_luong,
                nv.link_minh_chung_drive,
                nv.nguon_du_lieu,
            ]
        )
        _luu_csdl_an_toan(wb, self.duong_dan_csdl)
        wb.close()
        return ma_nv

    def lay_cong_viec_cu_va_lap_lai_goi_y(
        self, ma_can_bo: str, thang_hien_tai: int, nam_hien_tai: int
    ) -> List[NhiemVuCongViec]:
        """
        Lấy danh sách các công việc từ báo cáo các kỳ trước của giảng viên và từ Sheet
        `CongViec_LapLai` mà chưa có trong danh sách nhiệm vụ của tháng hiện tại,
        giúp giảng viên tích chọn 'Sử dụng lại tiếp tục' ngay trên bảng nhiệm vụ.
        """
        self.khoi_tao_csdl_mau_neu_chua_co()
        ds_gv = self.doc_danh_sach_giang_vien()
        gv = ds_gv.get(ma_can_bo)
        ten_gv = gv.ho_va_ten if gv else ""

        wb = openpyxl.load_workbook(self.duong_dan_csdl, data_only=True)
        ws_nv = wb[self.SHEET_NHIEM_VU]
        ws_lap = wb[self.SHEET_VIEC_LAP_LAI]

        ten_da_co_thang_nay = set()
        goi_y_tu_ky_truoc: List[NhiemVuCongViec] = []
        ten_da_goi_y = set()

        cuoi_thang_nay = ngay_cuoi_thang(thang_hien_tai, nam_hien_tai)
        dau_thang_nay = date(nam_hien_tai, thang_hien_tai, 1)

        # 1. Quét các nhiệm vụ đã có trong tháng hiện tại
        for r in range(2, ws_nv.max_row + 1):
            t = int(ws_nv.cell(r, 2).value or 0)
            n = int(ws_nv.cell(r, 3).value or 0)
            mcb = str(ws_nv.cell(r, 4).value or "").strip()
            ten_nv = str(ws_nv.cell(r, 7).value or "").strip()
            if t == thang_hien_tai and n == nam_hien_tai and mcb == ma_can_bo and ten_nv:
                ten_da_co_thang_nay.add(ten_nv)

        # 2. Quét các công việc từ các kỳ báo cáo cũ của giảng viên
        for r in range(2, ws_nv.max_row + 1):
            t = int(ws_nv.cell(r, 2).value or 0)
            n = int(ws_nv.cell(r, 3).value or 0)
            mcb = str(ws_nv.cell(r, 4).value or "").strip()
            ten_nv = str(ws_nv.cell(r, 7).value or "").strip()
            if (
                mcb == ma_can_bo
                and (n < nam_hien_tai or (n == nam_hien_tai and t < thang_hien_tai))
                and ten_nv
                and ten_nv not in ten_da_co_thang_nay
                and ten_nv not in ten_da_goi_y
            ):
                thoi_han_cu = _parse_date(ws_nv.cell(r, 12).value)
                tinh_trang_cu = str(ws_nv.cell(r, 14).value or "Đang thực hiện – đúng tiến độ").strip()
                # Nếu việc cũ chưa xong -> giữ nguyên thời hạn gốc; nếu việc cũ đã xong -> gợi ý hạn cuối tháng này
                thoi_han_goi_y = (
                    thoi_han_cu
                    if (thoi_han_cu and not tinh_trang_cu.startswith("Hoàn thành"))
                    else cuoi_thang_nay
                )
                goi_y_tu_ky_truoc.append(
                    NhiemVuCongViec(
                        ma_nhiem_vu=f"CU_T{t:02d}_{n}_{r:03d}",
                        nhiem_vu=ten_nv,
                        nguon_giao=str(ws_nv.cell(r, 8).value or "Kế hoạch năm học").strip(),
                        so_van_ban_ma_viec=str(ws_nv.cell(r, 9).value or "").strip(),
                        san_pham_ket_qua=str(ws_nv.cell(r, 10).value or "").strip(),
                        ngay_giao=dau_thang_nay,
                        thoi_han=thoi_han_goi_y,
                        ngay_hoan_thanh=cuoi_thang_nay if tinh_trang_cu == "Hoàn thành đúng hạn" else None,
                        tinh_trang=tinh_trang_cu,
                        muc_do_hoan_thanh=float(ws_nv.cell(r, 15).value or 100.0),
                        kho_khan_de_xuat=str(ws_nv.cell(r, 16).value or "").strip(),
                        ma_can_bo_chu_tri=ma_can_bo,
                        ten_can_bo_chu_tri=ten_gv,
                        nguoi_phoi_hop=str(ws_nv.cell(r, 6).value or "Bộ môn ĐKT & KGN").strip(),
                        la_nhiem_vu_trong_tam=bool(ws_nv.cell(r, 17).value),
                        la_cong_viec_lap_lai=True,
                        muc_uu_tien=str(ws_nv.cell(r, 19).value or "Thường").strip(),
                        danh_gia_chat_luong=str(ws_nv.cell(r, 20).value or "Đạt yêu cầu – tốt").strip(),
                        nguon_du_lieu=f"Báo cáo cũ T{t:02d}/{n}",
                    )
                )
                ten_da_goi_y.add(ten_nv)

        # 3. Quét ngân hàng công việc lặp lại (CongViec_LapLai) chưa có trong tháng hiện tại
        for r in range(2, ws_lap.max_row + 1):
            ap_dung = str(ws_lap.cell(r, 2).value or "ALL").strip()
            ten_nv = str(ws_lap.cell(r, 3).value or "").strip()
            if (
                ap_dung in ("ALL", ma_can_bo)
                and ten_nv
                and ten_nv not in ten_da_co_thang_nay
                and ten_nv not in ten_da_goi_y
            ):
                ngay_han_int = int(ws_lap.cell(r, 7).value or 0)
                thoi_han_nv = (
                    cuoi_thang_nay
                    if ngay_han_int <= 0 or ngay_han_int > cuoi_thang_nay.day
                    else date(nam_hien_tai, thang_hien_tai, ngay_han_int)
                )
                goi_y_tu_ky_truoc.append(
                    NhiemVuCongViec(
                        ma_nhiem_vu=f"LAP_GOIY_{r:03d}",
                        nhiem_vu=ten_nv,
                        nguon_giao=str(ws_lap.cell(r, 4).value or "Chức năng, nhiệm vụ thường xuyên").strip(),
                        so_van_ban_ma_viec=str(ws_lap.cell(r, 5).value or "").strip(),
                        san_pham_ket_qua=str(ws_lap.cell(r, 6).value or "").strip(),
                        ngay_giao=dau_thang_nay,
                        thoi_han=thoi_han_nv,
                        ngay_hoan_thanh=thoi_han_nv,
                        tinh_trang="Hoàn thành đúng hạn",
                        muc_do_hoan_thanh=100.0,
                        kho_khan_de_xuat="",
                        ma_can_bo_chu_tri=ma_can_bo,
                        ten_can_bo_chu_tri=ten_gv,
                        nguoi_phoi_hop="Bộ môn ĐKT & KGN",
                        la_nhiem_vu_trong_tam=bool(ws_lap.cell(r, 8).value),
                        la_cong_viec_lap_lai=True,
                        muc_uu_tien="Thường",
                        danh_gia_chat_luong="Đạt yêu cầu – tốt",
                        nguon_du_lieu="Ngân hàng việc lặp lại",
                    )
                )
                ten_da_goi_y.add(ten_nv)

        wb.close()
        return goi_y_tu_ky_truoc

    def cap_nhat_danh_sach_nhiem_vu_thang(
        self,
        ma_can_bo: str,
        thang: int,
        nam: int,
        danh_sach_nhiem_vu_chon: List[NhiemVuCongViec],
    ) -> int:
        """
        Cập nhật toàn bộ danh sách nhiệm vụ được tích chọn sử dụng của giảng viên `ma_can_bo`
        trong kỳ `(thang, nam)` vào Sheet `NhiemVu_Thang`, đồng thời đồng bộ các công việc
        được đánh dấu `la_cong_viec_lap_lai=True` vào Sheet `CongViec_LapLai` để tái sử dụng cho kỳ sau.
        """
        self.khoi_tao_csdl_mau_neu_chua_co()
        ds_gv = self.doc_danh_sach_giang_vien()
        gv = ds_gv.get(ma_can_bo)
        ten_gv = gv.ho_va_ten if gv else ""

        wb = openpyxl.load_workbook(self.duong_dan_csdl)
        ws_nv = wb[self.SHEET_NHIEM_VU]
        ws_lap = wb[self.SHEET_VIEC_LAP_LAI]

        # Giữ lại các dòng của giảng viên khác hoặc kỳ báo cáo khác
        cac_dong_giu_lai: List[list] = []
        for r in range(2, ws_nv.max_row + 1):
            t = int(ws_nv.cell(r, 2).value or 0)
            n = int(ws_nv.cell(r, 3).value or 0)
            mcb = str(ws_nv.cell(r, 4).value or "").strip()
            if not (t == thang and n == nam and mcb == ma_can_bo):
                cac_dong_giu_lai.append([ws_nv.cell(r, c).value for c in range(1, 23)])

        # Xóa dữ liệu cũ từ dòng 2 trở đi trên Sheet NhiemVu_Thang
        if ws_nv.max_row >= 2:
            ws_nv.delete_rows(2, ws_nv.max_row - 1)

        for row_vals in cac_dong_giu_lai:
            ws_nv.append(row_vals)

        dau_thang = date(nam, thang, 1)
        so_luong_luu = 0
        for idx, nv in enumerate(danh_sach_nhiem_vu_chon, start=1):
            if not nv.nhiem_vu.strip():
                continue
            ma_nv = nv.ma_nhiem_vu.strip()
            if not ma_nv or ma_nv.startswith("CU_") or ma_nv.startswith("LAP_GOIY_"):
                ma_nv = f"NV_{nam}{thang:02d}_{ma_can_bo}_{idx:02d}"
            ws_nv.append(
                [
                    ma_nv,
                    thang,
                    nam,
                    ma_can_bo,
                    nv.ten_can_bo_chu_tri or ten_gv,
                    nv.nguoi_phoi_hop or "Bộ môn ĐKT & KGN",
                    nv.nhiem_vu.strip(),
                    nv.nguon_giao,
                    nv.so_van_ban_ma_viec,
                    nv.san_pham_ket_qua,
                    nv.ngay_giao or dau_thang,
                    nv.thoi_han,
                    nv.ngay_hoan_thanh,
                    nv.tinh_trang,
                    nv.muc_do_hoan_thanh,
                    nv.kho_khan_de_xuat,
                    bool(nv.la_nhiem_vu_trong_tam),
                    bool(nv.la_cong_viec_lap_lai),
                    nv.muc_uu_tien or "Thường",
                    nv.danh_gia_chat_luong or "Đạt yêu cầu – tốt",
                    nv.link_minh_chung_drive,
                    nv.nguon_du_lieu or "Cập nhật trên Webapp",
                ]
            )
            so_luong_luu += 1

        # Đồng bộ cờ lặp lại vào Sheet CongViec_LapLai để tái sử dụng cho tháng tiếp theo
        ten_lap_da_co = {}
        for r in range(2, ws_lap.max_row + 1):
            ap_dung = str(ws_lap.cell(r, 2).value or "").strip()
            ten_l = str(ws_lap.cell(r, 3).value or "").strip()
            if ap_dung in ("ALL", ma_can_bo) and ten_l:
                ten_lap_da_co[(ap_dung, ten_l)] = r

        for nv in danh_sach_nhiem_vu_chon:
            ten_nv = nv.nhiem_vu.strip()
            if not ten_nv:
                continue
            if nv.la_cong_viec_lap_lai:
                if ("ALL", ten_nv) not in ten_lap_da_co and (ma_can_bo, ten_nv) not in ten_lap_da_co:
                    ngay_han = nv.thoi_han.day if nv.thoi_han else 0
                    ws_lap.append(
                        [
                            f"LAP_{ma_can_bo}_{ws_lap.max_row:03d}",
                            ma_can_bo,
                            ten_nv,
                            nv.nguon_giao,
                            nv.so_van_ban_ma_viec,
                            nv.san_pham_ket_qua,
                            ngay_han,
                            bool(nv.la_nhiem_vu_trong_tam),
                            True,
                        ]
                    )
                elif (ma_can_bo, ten_nv) in ten_lap_da_co:
                    r_lap = ten_lap_da_co[(ma_can_bo, ten_nv)]
                    ws_lap.cell(r_lap, 8, bool(nv.la_nhiem_vu_trong_tam))
                    ws_lap.cell(r_lap, 9, True)
            else:
                if (ma_can_bo, ten_nv) in ten_lap_da_co:
                    r_lap = ten_lap_da_co[(ma_can_bo, ten_nv)]
                    ws_lap.cell(r_lap, 9, False)

        _luu_csdl_an_toan(wb, self.duong_dan_csdl)
        wb.close()
        return so_luong_luu

    def cap_nhat_cham_diem_va_chuyen_mon(
        self,
        ma_can_bo: str,
        thang: int,
        nam: int,
        hoat_dong_cm: HoatDongChuyenMon,
        diem_tu_cham_chung: float,
        diem_tu_cham_nv: Optional[float],
        diem_tbm_chung: Optional[float] = None,
        diem_tbm_nv: Optional[float] = None,
        de_xuat_kien_nghi: str = "",
    ) -> None:
        """
        Cập nhật điểm đánh giá (tự chấm & Trưởng đơn vị chấm), số giờ chuyên môn
        và đề xuất kiến nghị của giảng viên vào Sheet `KH_GiangDay_NCKH`.
        """
        self.khoi_tao_csdl_mau_neu_chua_co()
        ds_gv = self.doc_danh_sach_giang_vien()
        gv = ds_gv.get(ma_can_bo)
        ten_gv = gv.ho_va_ten if gv else ""

        wb = openpyxl.load_workbook(self.duong_dan_csdl)
        ws_cm = wb[self.SHEET_CHUYEN_MON]

        found_row = None
        for r in range(2, ws_cm.max_row + 1):
            t = int(ws_cm.cell(r, 1).value or 0)
            n = int(ws_cm.cell(r, 2).value or 0)
            mcb = str(ws_cm.cell(r, 3).value or "").strip()
            if t == thang and n == nam and mcb == ma_can_bo:
                found_row = r
                break

        row_values = [
            thang,
            nam,
            ma_can_bo,
            ten_gv,
            hoat_dong_cm.giang_day_thang,
            hoat_dong_cm.giang_day_luy_ke,
            hoat_dong_cm.giang_day_dinh_muc,
            hoat_dong_cm.giang_day_ghi_chu,
            hoat_dong_cm.nckh_thang,
            hoat_dong_cm.nckh_luy_ke,
            hoat_dong_cm.nckh_dinh_muc,
            hoat_dong_cm.nckh_ghi_chu,
            hoat_dong_cm.nhiem_vu_khac_thang,
            hoat_dong_cm.nhiem_vu_khac_luy_ke,
            hoat_dong_cm.nhiem_vu_khac_dinh_muc,
            hoat_dong_cm.nhiem_vu_khac_ghi_chu,
            hoat_dong_cm.san_pham_khcn_huong_dan,
            diem_tu_cham_chung,
            diem_tu_cham_nv,
            diem_tbm_chung,
            diem_tbm_nv,
            de_xuat_kien_nghi,
        ]

        if found_row is not None:
            for col_idx, val in enumerate(row_values, start=1):
                ws_cm.cell(found_row, col_idx, val)
        else:
            ws_cm.append(row_values)

        _luu_csdl_an_toan(wb, self.duong_dan_csdl)
        wb.close()

    def cap_nhat_ke_hoach_thang_toi(
        self,
        ma_can_bo: str,
        thang: int,
        nam: int,
        danh_sach_ke_hoach: List[KeHoachKyToi],
    ) -> int:
        """
        Cập nhật danh sách Mục IV. Kế hoạch công việc tháng tới của giảng viên `ma_can_bo`
        vào Sheet `KeHoach_ThangToi`.
        """
        self.khoi_tao_csdl_mau_neu_chua_co()
        wb = openpyxl.load_workbook(self.duong_dan_csdl)
        ws_kh = wb[self.SHEET_KE_HOACH_TOI]

        cac_dong_giu_lai: List[list] = []
        for r in range(2, ws_kh.max_row + 1):
            t = int(ws_kh.cell(r, 1).value or 0)
            n = int(ws_kh.cell(r, 2).value or 0)
            ap_dung = str(ws_kh.cell(r, 3).value or "").strip()
            if not (t == thang and n == nam and ap_dung == ma_can_bo):
                cac_dong_giu_lai.append([ws_kh.cell(r, c).value for c in range(1, 9)])

        if ws_kh.max_row >= 2:
            ws_kh.delete_rows(2, ws_kh.max_row - 1)

        for row_vals in cac_dong_giu_lai:
            ws_kh.append(row_vals)

        so_luong = 0
        for kh in danh_sach_ke_hoach:
            if not kh.nhiem_vu.strip():
                continue
            ws_kh.append(
                [
                    thang,
                    nam,
                    ma_can_bo,
                    kh.nhiem_vu.strip(),
                    kh.san_pham_du_kien.strip(),
                    kh.thoi_han,
                    kh.don_vi_phoi_hop or "Bộ môn ĐKT & KGN",
                    kh.ghi_chu.strip(),
                ]
            )
            so_luong += 1

        _luu_csdl_an_toan(wb, self.duong_dan_csdl)
        wb.close()
        return so_luong

    def cap_nhat_trong_tam_va_diem_tbm_don_vi(
        self,
        thang: int,
        nam: int,
        map_ma_nv_trong_tam: Optional[Dict[str, bool]] = None,
        map_diem_tbm: Optional[Dict[str, Tuple[Optional[float], Optional[float]]]] = None,
        map_trong_tam_theo_ma_nv: Optional[Dict[str, bool]] = None,
        map_diem_tbm_theo_ma_cb: Optional[Dict[str, Tuple[Optional[float], Optional[float]]]] = None,
    ) -> int:
        """
        Cập nhật cờ Nhiệm vụ trọng tâm của đơn vị (Mẫu 02a/02b) theo `ma_nhiem_vu` trên Sheet `NhiemVu_Thang`
        và cập nhật Điểm Trưởng đơn vị chấm (Tiêu chí chung <=30, Kết quả NV <=70) cho 21 giảng viên
        trên Sheet `KH_GiangDay_NCKH` (phục vụ Mẫu 02c).
        """
        self.khoi_tao_csdl_mau_neu_chua_co()
        wb = openpyxl.load_workbook(self.duong_dan_csdl)
        ws_nv = wb[self.SHEET_NHIEM_VU]
        ws_cm = wb[self.SHEET_CHUYEN_MON]

        dict_tt = map_ma_nv_trong_tam or map_trong_tam_theo_ma_nv or {}
        dict_diem = map_diem_tbm or map_diem_tbm_theo_ma_cb or {}

        so_nv_cap_nhat = 0
        if dict_tt:
            for r in range(2, ws_nv.max_row + 1):
                t = int(ws_nv.cell(r, 2).value or 0)
                n = int(ws_nv.cell(r, 3).value or 0)
                ma_nv = str(ws_nv.cell(r, 1).value or "").strip()
                if t == thang and n == nam and ma_nv in dict_tt:
                    ws_nv.cell(r, 17, bool(dict_tt[ma_nv]))
                    so_nv_cap_nhat += 1

        so_gv_cham = 0
        if dict_diem:
            for r in range(2, ws_cm.max_row + 1):
                t = int(ws_cm.cell(r, 1).value or 0)
                n = int(ws_cm.cell(r, 2).value or 0)
                mcb = str(ws_cm.cell(r, 3).value or "").strip()
                if t == thang and n == nam and mcb in dict_diem:
                    d_chung, d_nv = dict_diem[mcb]
                    ws_cm.cell(r, 20, d_chung)
                    ws_cm.cell(r, 21, d_nv)
                    so_gv_cham += 1

        _luu_csdl_an_toan(wb, self.duong_dan_csdl)
        wb.close()
        return so_nv_cap_nhat + so_gv_cham
