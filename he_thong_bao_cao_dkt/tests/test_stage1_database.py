# -*- coding: utf-8 -*-
"""
Unit Test & Kiểm toán Giai đoạn 1 (Audit Stage 1):
Kiểm tra Mô hình dữ liệu (models.py) và Quản lý CSDL Bộ môn (database_manager.py).
"""

from datetime import date
from pathlib import Path
import tempfile
import unittest

from he_thong_bao_cao_dkt.database_manager import QuanLyCSDLBộMôn
from he_thong_bao_cao_dkt.models import (
    DANH_SACH_GIANG_VIEN_DKT,
    NhiemVuCongViec,
    bo_dau_tieng_viet,
    ngay_cuoi_thang,
)


class TestStage1Database(unittest.TestCase):
    """Kiểm toán tính đúng đắn của CSDL 21 giảng viên Bộ môn ĐKT & KGN và logic kế thừa."""

    def setUp(self) -> None:
        self.tmp_dir = tempfile.TemporaryDirectory()
        self.db_path = Path(self.tmp_dir.name) / "CSDL_Test_DKT.xlsx"
        self.db = QuanLyCSDLBộMôn(self.db_path)
        self.db.khoi_tao_csdl_mau_neu_chua_co(ghi_de=True)

    def tearDown(self) -> None:
        self.tmp_dir.cleanup()

    def test_01_danh_sach_21_giang_vien(self) -> None:
        """Đảm bảo có đúng 22 thành viên Bộ môn Địa kỹ thuật và Không gian ngầm (bao gồm Nguyễn Thị Thanh Giang - 0537)."""
        self.assertEqual(len(DANH_SACH_GIANG_VIEN_DKT), 22)
        ds_gv = self.db.doc_danh_sach_giang_vien()
        self.assertEqual(len(ds_gv), 22)
        self.assertIn("1051", ds_gv)
        self.assertIn("0694", ds_gv)
        self.assertIn("0537", ds_gv)
        self.assertNotIn("0799", ds_gv)
        self.assertEqual(ds_gv["1051"].ho_va_ten, "Nguyễn Hải Hà")
        self.assertEqual(ds_gv["1051"].tai_khoan_usmart, "hanh1")
        self.assertEqual(ds_gv["0694"].ho_va_ten, "Trần Thế Truyền")
        self.assertEqual(ds_gv["0537"].ho_va_ten, "Nguyễn Thị Thanh Giang")
        self.assertEqual(ds_gv["0537"].tai_khoan_usmart, "giangntt")
        self.assertEqual(ds_gv["1472"].tai_khoan_usmart, "linhnt2")
        self.assertEqual(ds_gv["0064"].tai_khoan_usmart, "ducla")

    def test_02_doc_bao_cao_ca_nhan_va_ten_file_chuan(self) -> None:
        """Kiểm tra đọc báo cáo tháng 11/2026 của giảng viên Nguyễn Hải Hà (1051)."""
        bc = self.db.doc_bao_cao_ca_nhan("1051", thang=11, nam=2026)
        self.assertEqual(bc.giang_vien.ho_va_ten, "Nguyễn Hải Hà")
        self.assertEqual(bc.ngay_cuoi_ky, date(2026, 11, 30))
        self.assertEqual(
            bc.ten_file_xlsx_chuan, "BCThang_DKT_NguyenHaiHa_11-2026.xlsx"
        )
        self.assertEqual(
            bc.ten_file_docx_chuan, "BCThang_DKT_NguyenHaiHa_11-2026.docx"
        )
        self.assertEqual(bc.tong_so_nhiem_vu, 5)
        self.assertEqual(bc.so_nhiem_vu_da_den_han, 5)
        self.assertEqual(bc.so_nhiem_vu_hoan_thanh_dung_han, 5)
        self.assertAlmostEqual(bc.ty_le_hoan_thanh_dung_han or 0.0, 1.0)
        self.assertEqual(bc.diem_tham_chieu_nhiem_vu, 70)
        self.assertEqual(bc.muc_xep_loai_tham_chieu, "Hoàn thành xuất sắc nhiệm vụ")

    def test_03_sao_chep_viec_lap_lai_va_chuyen_tiep_viec_chua_xong(self) -> None:
        """Kiểm tra sao chép công việc lặp lại sang tháng 12/2026 và chuyển việc chưa xong giữ nguyên thời hạn gốc."""
        # Thêm 1 việc chưa xong ở tháng 11/2026 với thời hạn gốc 25/11/2026
        nv_chua_xong = NhiemVuCongViec(
            ma_nhiem_vu="NV_202611_999",
            nhiem_vu="Hoàn thiện chuyên đề thí nghiệm nén ba trục chu kỳ",
            nguon_giao="Kế hoạch năm học",
            so_van_ban_ma_viec="DT-2026",
            san_pham_ket_qua="Báo cáo chuyên đề",
            thoi_han=date(2026, 11, 25),
            ngay_hoan_thanh=None,
            tinh_trang="Đang thực hiện – chậm tiến độ",
            muc_do_hoan_thanh=70.0,
            ma_can_bo_chu_tri="1051",
            ten_can_bo_chu_tri="Nguyễn Hải Hà",
        )
        self.db.them_nhiem_vu_moi(11, 2026, nv_chua_xong)

        # Sao chép việc lặp lại sang tháng 12/2026 cho GV 1051 (có 5 mẫu: 3 ALL + 2 riêng 1051)
        so_lap = self.db.sao_chep_cong_viec_lap_lai_sang_thang_moi(
            12, 2026, ma_can_bo="1051"
        )
        self.assertEqual(so_lap, 5)

        # Chuyển tiếp việc chưa xong từ tháng 11/2026 sang tháng 12/2026
        so_ct = self.db.chuyen_tiep_nhiem_vu_chua_hoan_thanh(
            11, 2026, 12, 2026, ma_can_bo="1051"
        )
        self.assertEqual(so_ct, 1)

        # Đọc báo cáo tháng 12/2026 để kiểm chứng thời hạn gốc được giữ nguyên
        bc_t12 = self.db.doc_bao_cao_ca_nhan("1051", thang=12, nam=2026)
        self.assertEqual(bc_t12.tong_so_nhiem_vu, 6)
        nv_ct = [
            x
            for x in bc_t12.danh_sach_nhiem_vu
            if "nén ba trục chu kỳ" in x.nhiem_vu
        ][0]
        self.assertEqual(nv_ct.thoi_han, date(2026, 11, 25))


if __name__ == "__main__":
    unittest.main()
