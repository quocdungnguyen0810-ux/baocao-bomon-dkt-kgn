# -*- coding: utf-8 -*-
"""
Module Sinh Báo cáo Bản cứng Word (.docx) Tự động cho Giảng viên (Mẫu 01b)
và Bản tổng hợp Bộ môn (Mẫu 02b/02c) theo Thông báo số 940/TB-ĐHGTVT.

Sử dụng thư viện `python-docx` với định dạng khổ giấy A4 ngang (Landscape),
phông chữ Times New Roman chuẩn văn bản hành chính.
"""

from datetime import date
import logging
from pathlib import Path
from typing import List, Optional
import docx
from docx.enum.section import WD_ORIENT
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Cm, Pt, RGBColor

from .models import (
    BaoCaoCaNhanThang,
    MA_VIET_TAT_DON_VI,
    TEN_DON_VI_CHUAN,
)

logger = logging.getLogger(__name__)


def _thiet_lap_font_mac_dinh(doc: docx.Document) -> None:
    """Thiết lập phông chữ Times New Roman cỡ 11pt cho toàn bộ tài liệu Word."""
    style = doc.styles["Normal"]
    style.font.name = "Times New Roman"
    style.font.size = Pt(11)
    style.font.color.rgb = RGBColor(0, 0, 0)


def _chuyen_kho_a4_ngang(doc: docx.Document) -> None:
    """Chuyển trang sang khổ A4 ngang (Landscape) để hiển thị bảng 10-13 cột không bị vỡ dòng."""
    section = doc.sections[0]
    section.orientation = WD_ORIENT.LANDSCAPE
    section.page_width = Cm(29.7)
    section.page_height = Cm(21.0)
    section.top_margin = Cm(1.8)
    section.bottom_margin = Cm(1.8)
    section.left_margin = Cm(2.0)
    section.right_margin = Cm(1.8)


def _them_quoc_hieu_tieu_ngu(doc: docx.Document, ma_mau: str = "Mẫu 01b") -> None:
    """Thêm bảng tiêu đề Quốc hiệu - Tiêu ngữ chuẩn văn bản hành chính Trường ĐH GTVT."""
    tbl = doc.add_table(rows=1, cols=2)
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    cell_trai = tbl.cell(0, 0)
    cell_phai = tbl.cell(0, 1)

    p1 = cell_trai.paragraphs[0]
    p1.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r1 = p1.add_run("TRƯỜNG ĐẠI HỌC GIAO THÔNG VẬN TẢI\n")
    r1.bold = False
    r1.font.size = Pt(11)
    r2 = p1.add_run(f"{TEN_DON_VI_CHUAN.upper()}")
    r2.bold = True
    r2.font.size = Pt(11)

    p2 = cell_phai.paragraphs[0]
    p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r3 = p2.add_run("CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM\n")
    r3.bold = True
    r3.font.size = Pt(11)
    r4 = p2.add_run(f"Độc lập - Tự do - Hạnh phúc\n({ma_mau} - TB số 940/TB-ĐHGTVT)")
    r4.bold = True
    r4.font.size = Pt(10.5)


class BoSinhBaoCaoDocx:
    """Bộ máy xuất báo cáo bản cứng Word (.docx) phục vụ in ấn và ký duyệt."""

    def xuat_docx_ca_nhan_01b(
        self,
        bao_cao: BaoCaoCaNhanThang,
        thu_muc_xuat: Path,
        ten_file_tuy_chinh: Optional[str] = None,
    ) -> Path:
        """Xuất báo cáo cá nhân tháng (Mẫu 01b) ra tệp .docx chuẩn in ấn."""
        thu_muc_xuat = Path(thu_muc_xuat)
        thu_muc_xuat.mkdir(parents=True, exist_ok=True)
        duong_dan_xuat = thu_muc_xuat / (ten_file_tuy_chinh or bao_cao.ten_file_docx_chuan)

        doc = docx.Document()
        _thiet_lap_font_mac_dinh(doc)
        _chuyen_kho_a4_ngang(doc)
        _them_quoc_hieu_tieu_ngu(doc, ma_mau="Mẫu 01b")

        p_title = doc.add_paragraph()
        p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run_t = p_title.add_run(
            f"\nBÁO CÁO KẾT QUẢ CÔNG VIỆC THÁNG {bao_cao.thang:02d}/{bao_cao.nam} VÀ TỰ CHẤM ĐIỂM\n"
        )
        run_t.bold = True
        run_t.font.size = Pt(13)
        run_sub = p_title.add_run(
            f"(Ngày cuối kỳ báo cáo: {bao_cao.ngay_cuoi_ky.strftime('%d/%m/%Y')})"
        )
        run_sub.italic = True
        run_sub.font.size = Pt(11)

        # Thông tin giảng viên
        p_info = doc.add_paragraph()
        p_info.add_run("Họ và tên: ").bold = True
        p_info.add_run(f"{bao_cao.giang_vien.ho_va_ten}\t\t")
        p_info.add_run("Mã cán bộ: ").bold = True
        p_info.add_run(f"{bao_cao.giang_vien.ma_can_bo}\t\t")
        p_info.add_run("Vị trí việc làm, chức danh: ").bold = True
        p_info.add_run(f"{bao_cao.giang_vien.chuc_danh_day_du}\n")
        p_info.add_run("Người nhận báo cáo: ").bold = True
        p_info.add_run(f"{bao_cao.giang_vien.nguoi_nhan_bao_cao}")

        # Mục I. Kết quả thực hiện nhiệm vụ trong tháng
        p_m1 = doc.add_paragraph()
        p_m1.add_run("I. KẾT QUẢ THỰC HIỆN NHIỆM VỤ TRONG THÁNG").bold = True

        headers_m1 = [
            "TT",
            "Nhiệm vụ, công việc",
            "Nguồn giao",
            "Số VB / Mã việc",
            "Sản phẩm đầu ra, kết quả",
            "Thời hạn",
            "Ngày HT",
            "Tình trạng",
            "Mức độ (%)",
            "Đã đến hạn",
            "Đúng hạn",
        ]
        tbl_m1 = doc.add_table(rows=1, cols=len(headers_m1))
        tbl_m1.style = "Table Grid"
        for i, h in enumerate(headers_m1):
            c = tbl_m1.cell(0, i)
            c.text = h
            c.paragraphs[0].runs[0].bold = True
            c.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER

        ck = bao_cao.ngay_cuoi_ky
        for idx, nv in enumerate(bao_cao.danh_sach_nhiem_vu, start=1):
            row_cells = tbl_m1.add_row().cells
            row_cells[0].text = str(idx)
            row_cells[1].text = nv.nhiem_vu
            row_cells[2].text = nv.nguon_giao
            row_cells[3].text = nv.so_van_ban_ma_viec
            row_cells[4].text = nv.san_pham_ket_qua
            row_cells[5].text = nv.thoi_han.strftime("%d/%m/%Y") if nv.thoi_han else ""
            row_cells[6].text = (
                nv.ngay_hoan_thanh.strftime("%d/%m/%Y") if nv.ngay_hoan_thanh else ""
            )
            row_cells[7].text = nv.tinh_trang
            row_cells[8].text = f"{nv.muc_do_hoan_thanh:.0f}%"
            row_cells[9].text = str(nv.da_den_han_trong_ky(ck))
            row_cells[10].text = str(nv.hoan_thanh_dung_han_usmart())

        ty_le_str = (
            f"{bao_cao.ty_le_hoan_thanh_dung_han * 100:.1f}%"
            if bao_cao.ty_le_hoan_thanh_dung_han is not None
            else "N/A"
        )
        p_sum = doc.add_paragraph()
        p_sum.add_run(
            f"Tổng hợp Mục I: Tổng số nhiệm vụ: {bao_cao.tong_so_nhiem_vu} | "
            f"Số nhiệm vụ đã đến hạn: {bao_cao.so_nhiem_vu_da_den_han} | "
            f"Số nhiệm vụ hoàn thành đúng hạn: {bao_cao.so_nhiem_vu_hoan_thanh_dung_han} | "
            f"Tỷ lệ hoàn thành đúng hạn: {ty_le_str}."
        ).italic = True

        # Mục II. Hoạt động chuyên môn
        p_m2 = doc.add_paragraph()
        p_m2.add_run("II. HOẠT ĐỘNG CHUYÊN MÔN (GIẢNG VIÊN KÊ KHAI)").bold = True
        cm = bao_cao.hoat_dong_chuyen_mon
        tbl_m2 = doc.add_table(rows=4, cols=6)
        tbl_m2.style = "Table Grid"
        headers_m2 = [
            "Nội dung",
            "Thực hiện trong tháng",
            "Lũy kế năm học",
            "Định mức năm học",
            "Tỷ lệ lũy kế / định mức",
            "Ghi chú",
        ]
        for i, h in enumerate(headers_m2):
            c = tbl_m2.cell(0, i)
            c.text = h
            c.paragraphs[0].runs[0].bold = True

        dong_cm = [
            (
                "1. Giảng dạy (giờ chuẩn)",
                cm.giang_day_thang,
                cm.giang_day_luy_ke,
                cm.giang_day_dinh_muc,
                cm.ty_le_giang_day,
                cm.giang_day_ghi_chu,
            ),
            (
                "2. Nghiên cứu khoa học (giờ chuẩn)",
                cm.nckh_thang,
                cm.nckh_luy_ke,
                cm.nckh_dinh_muc,
                cm.ty_le_nckh,
                cm.nckh_ghi_chu,
            ),
            (
                "3. Nhiệm vụ khác theo quy định (giờ)",
                cm.nhiem_vu_khac_thang,
                cm.nhiem_vu_khac_luy_ke,
                cm.nhiem_vu_khac_dinh_muc,
                cm.ty_le_nhiem_vu_khac,
                cm.nhiem_vu_khac_ghi_chu,
            ),
        ]
        for r_idx, (ten, th, lk, dm, tl, gc) in enumerate(dong_cm, start=1):
            cells = tbl_m2.rows[r_idx].cells
            cells[0].text = ten
            cells[1].text = f"{th:.1f}"
            cells[2].text = f"{lk:.1f}"
            cells[3].text = f"{dm:.1f}"
            cells[4].text = f"{tl * 100:.1f}%" if tl is not None else ""
            cells[5].text = gc

        p_sp = doc.add_paragraph()
        p_sp.add_run(
            "Sản phẩm KHCN, hướng dẫn người học trong tháng: "
        ).bold = True
        p_sp.add_run(cm.san_pham_khcn_huong_dan)

        # Mục III. Tự đánh giá, chấm điểm tháng
        p_m3 = doc.add_paragraph()
        p_m3.add_run("III. TỰ ĐÁNH GIÁ, CHẤM ĐIỂM THÁNG").bold = True
        p_m3.add_run(
            f"\n- 1. Tiêu chí chung (tối đa 30 điểm): {bao_cao.diem_tu_cham_tieu_chi_chung:.1f} / 30"
            f"\n- 2. Kết quả thực hiện nhiệm vụ (tối đa 70 điểm): {bao_cao.diem_tu_cham_nv_hieu_luc:.1f} / 70 "
            f"(Điểm tham chiếu theo tỷ lệ đúng hạn: {bao_cao.diem_tham_chieu_nhiem_vu})"
            f"\n- Tổng điểm tự chấm: {bao_cao.tong_diem_tu_cham:.1f} / 100  —  "
            f"Mức xếp loại tham chiếu: {bao_cao.muc_xep_loai_tham_chieu}."
        )

        # Mục IV. Kế hoạch công việc tháng tới
        p_m4 = doc.add_paragraph()
        p_m4.add_run("IV. KẾ HOẠCH CÔNG VIỆC THÁNG TỚI").bold = True
        tbl_m4 = doc.add_table(rows=1, cols=5)
        tbl_m4.style = "Table Grid"
        for i, h in enumerate(["TT", "Nhiệm vụ, công việc", "Sản phẩm dự kiến", "Thời hạn", "Ghi chú"]):
            c = tbl_m4.cell(0, i)
            c.text = h
            c.paragraphs[0].runs[0].bold = True
        for idx, kh in enumerate(bao_cao.ke_hoach_thang_toi, start=1):
            rc = tbl_m4.add_row().cells
            rc[0].text = str(idx)
            rc[1].text = kh.nhiem_vu
            rc[2].text = kh.san_pham_du_kien
            rc[3].text = kh.thoi_han.strftime("%d/%m/%Y") if kh.thoi_han else ""
            rc[4].text = kh.ghi_chu

        # Mục V. Đề xuất, kiến nghị & Chữ ký
        p_m5 = doc.add_paragraph()
        p_m5.add_run("V. ĐỀ XUẤT, KIẾN NGHỊ: ").bold = True
        p_m5.add_run(bao_cao.de_xuat_kien_nghi or "Không.")

        ngay_ky = bao_cao.ngay_ky_bao_cao or bao_cao.ngay_cuoi_ky
        tbl_ky = doc.add_table(rows=1, cols=2)
        tbl_ky.alignment = WD_TABLE_ALIGNMENT.CENTER
        c_left = tbl_ky.cell(0, 0).paragraphs[0]
        c_left.alignment = WD_ALIGN_PARAGRAPH.CENTER
        c_left.add_run("NHẬN XÉT CỦA TRƯỞNG BỘ MÔN / TRƯỞNG ĐƠN VỊ\n").bold = True
        c_left.add_run("(Ghi nhận xét, ký và ghi rõ họ tên)\n\n\n\n")
        c_left.add_run(bao_cao.giang_vien.nguoi_nhan_bao_cao).bold = True

        c_right = tbl_ky.cell(0, 1).paragraphs[0]
        c_right.alignment = WD_ALIGN_PARAGRAPH.CENTER
        c_right.add_run(
            f"Hà Nội, ngày {ngay_ky.day:02d} tháng {ngay_ky.month:02d} năm {ngay_ky.year}\n"
        ).italic = True
        c_right.add_run("NGƯỜI BÁO CÁO\n").bold = True
        c_right.add_run("(Ký và ghi rõ họ tên)\n\n\n\n")
        c_right.add_run(bao_cao.giang_vien.ho_va_ten).bold = True

        doc.save(duong_dan_xuat)
        logger.info("Đã xuất thành công bản cứng Word (.docx) tại: %s", duong_dan_xuat)
        return duong_dan_xuat

    def xuat_docx_tong_hop_don_vi(
        self,
        danh_sach_bao_cao: List[BaoCaoCaNhanThang],
        thu_muc_xuat: Path,
        thang: int = 11,
        nam: int = 2026,
        ten_file_tuy_chinh: Optional[str] = None,
    ) -> Path:
        """Xuất báo cáo tổng hợp đơn vị (Mẫu 02b & 02c) ra bản cứng Word (.docx)."""
        thu_muc_xuat = Path(thu_muc_xuat)
        thu_muc_xuat.mkdir(parents=True, exist_ok=True)
        ten_file = ten_file_tuy_chinh or f"BCThang_{MA_VIET_TAT_DON_VI}_DonVi_{thang:02d}-{nam}.docx"
        duong_dan_xuat = thu_muc_xuat / ten_file

        doc = docx.Document()
        _thiet_lap_font_mac_dinh(doc)
        _chuyen_kho_a4_ngang(doc)
        _them_quoc_hieu_tieu_ngu(doc, ma_mau="Mẫu 02b & 02c")

        p_title = doc.add_paragraph()
        p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r_t = p_title.add_run(
            f"\nBÁO CÁO TỔNG HỢP KẾT QUẢ CÔNG VIỆC VÀ CHẤM ĐIỂM THÁNG {thang:02d}/{nam} CỦA ĐƠN VỊ\n"
        )
        r_t.bold = True
        r_t.font.size = Pt(13)

        tbl = doc.add_table(rows=1, cols=8)
        tbl.style = "Table Grid"
        headers = [
            "TT",
            "Họ và tên",
            "Vị trí việc làm",
            "NV đã đến hạn",
            "NV đúng hạn",
            "Tỷ lệ đúng hạn",
            "Điểm tự chấm",
            "Mức xếp loại tham chiếu",
        ]
        for i, h in enumerate(headers):
            c = tbl.cell(0, i)
            c.text = h
            c.paragraphs[0].runs[0].bold = True

        for idx, bc in enumerate(danh_sach_bao_cao, start=1):
            rc = tbl.add_row().cells
            rc[0].text = str(idx)
            rc[1].text = bc.giang_vien.ho_va_ten
            rc[2].text = bc.giang_vien.chuc_danh_day_du
            rc[3].text = str(bc.so_nhiem_vu_da_den_han)
            rc[4].text = str(bc.so_nhiem_vu_hoan_thanh_dung_han)
            tl = bc.ty_le_hoan_thanh_dung_han
            rc[5].text = f"{tl * 100:.0f}%" if tl is not None else "100%"
            rc[6].text = f"{bc.tong_diem_tu_cham:.1f}"
            rc[7].text = bc.muc_xep_loai_tham_chieu

        doc.save(duong_dan_xuat)
        return duong_dan_xuat
