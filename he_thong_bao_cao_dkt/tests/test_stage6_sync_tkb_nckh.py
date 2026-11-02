# -*- coding: utf-8 -*-
"""
Unit Test & Kiểm toán Giai đoạn 6 (Audit Stage 6):
Kiểm tra đồng bộ Thời khóa biểu Giảng dạy (`2627_1_1_TKB dự kiến- ĐKT.xlsx`),
Kế hoạch NCKH Đề tài các cấp & NCKHSV từ uSmart (`lannc`), và đối chiếu từng ô công thức
Excel so với biểu mẫu gốc của Trường Đại học Giao thông Vận tải.
"""

from pathlib import Path
import tempfile
import unittest
import openpyxl

from he_thong_bao_cao_dkt.database_manager import QuanLyCSDLBộMôn
from he_thong_bao_cao_dkt.excel_generator import BoSinhBaoCaoExcel, _tim_file_mau
from he_thong_bao_cao_dkt.tkb_giang_day_sync import chay_dong_bo_tkb_giang_day_toan_trinh
from he_thong_bao_cao_dkt.usmart_nckh_sync import chay_dong_bo_ke_hoach_nckh_toan_trinh
from he_thong_bao_cao_dkt.validator import BoKiemTraRangBuocVatLy


class TestStage6SyncTKBNckhAndFormulaAudit(unittest.TestCase):
    """Kiểm toán chuyên sâu Giai đoạn 6: Đồng bộ TKB, NCKH/NCKHSV và Đối chiếu công thức gốc."""

    def test_01_dong_bo_tkb_giang_day_va_kiem_toan(self) -> None:
        """Kiểm tra trích xuất 110 lớp học phần thực tế từ TKB dự kiến Học kỳ 1 năm học 2026-2027."""
        with tempfile.TemporaryDirectory() as tmp_dir:
            csdl_tmp = Path(tmp_dir) / "CSDL_Test_TKB.xlsx"
            kq = chay_dong_bo_tkb_giang_day_toan_trinh(
                thang=11, nam=2026, duong_dan_csdl_master=csdl_tmp
            )
            self.assertTrue(kq.duong_dan_gsheet_tkb.exists())
            self.assertGreaterEqual(kq.tong_so_lop_hoc_phan, 100)
            self.assertGreaterEqual(kq.so_lop_giang_day_thang_11, 40)

            db = QuanLyCSDLBộMôn(csdl_tmp)
            bc_ha = db.doc_bao_cao_ca_nhan("1051", thang=11, nam=2026)
            self.assertAlmostEqual(bc_ha.hoat_dong_chuyen_mon.giang_day_luy_ke, 124.13, places=1)
            self.assertAlmostEqual(bc_ha.hoat_dong_chuyen_mon.giang_day_dinh_muc, 187.5, places=1)
            self.assertIn("Cơ học đất", bc_ha.hoat_dong_chuyen_mon.giang_day_ghi_chu)

    def test_02_dong_bo_usmart_nckh_va_nckhsv(self) -> None:
        """Kiểm tra đồng bộ đề tài NCKH các cấp & NCKHSV từ uSmart (tài khoản lannc) vào CSDL."""
        with tempfile.TemporaryDirectory() as tmp_dir:
            csdl_tmp = Path(tmp_dir) / "CSDL_Test_NCKH.xlsx"
            kq = chay_dong_bo_ke_hoach_nckh_toan_trinh(
                username="lannc",
                password="giangchau",
                thang=11,
                nam=2026,
                dong_bo_live=False,
                duong_dan_csdl_master=csdl_tmp,
            )
            self.assertTrue(kq.duong_dan_gsheet_project.exists())
            self.assertEqual(kq.so_de_tai_cac_cap_2025_2027, 25)
            self.assertEqual(kq.so_de_tai_nckhsv_2025_2027, 77)
            self.assertGreaterEqual(kq.so_nhiem_vu_dong_bo_bao_cao_thang, 60)

            # Kiểm tra toàn bộ 21 giảng viên đều đạt chuẩn Ràng buộc Vật lý sau khi đồng bộ
            db = QuanLyCSDLBộMôn(csdl_tmp)
            val = BoKiemTraRangBuocVatLy()
            for bc in db.doc_toan_bo_bao_cao_bo_mon(11, 2026):
                res = val.kiem_tra_bao_cao_ca_nhan(bc)
                self.assertTrue(
                    res.hop_le_de_chot,
                    f"Giảng viên {bc.giang_vien.ho_va_ten} có lỗi chặn: {res.tom_tat_kiem_toan()}",
                )

    def test_03_doi_chieu_cong_thuc_va_dinh_dang_excel_goc(self) -> None:
        """Kiểm toán từng ô công thức (Cell-by-Cell Formula Audit) giữa tệp xuất và tệp biểu mẫu gốc."""
        with tempfile.TemporaryDirectory() as tmp_dir:
            tmp_path = Path(tmp_dir)
            csdl_tmp = tmp_path / "CSDL_Formula_Audit.xlsx"
            db = QuanLyCSDLBộMôn(csdl_tmp)
            db.khoi_tao_csdl_mau_neu_chua_co(ghi_de=True)

            gen = BoSinhBaoCaoExcel()
            bc_ha = db.doc_bao_cao_ca_nhan("1051", thang=11, nam=2026)
            out_01 = gen.xuat_mau_01_ca_nhan(bc_ha, tmp_path)

            wb_tpl = openpyxl.load_workbook(_tim_file_mau("01"), data_only=False)
            wb_out = openpyxl.load_workbook(out_01, data_only=False)

            ws_tpl = wb_tpl["BC tháng"]
            ws_out = wb_out["BC tháng"]

            # Kiểm tra tất cả các ô chứa công thức (bắt đầu bằng '=') trên biểu mẫu gốc phải được giữ nguyên 100%
            so_cong_thuc = 0
            for r in range(1, ws_tpl.max_row + 1):
                for c in range(1, ws_tpl.max_column + 1):
                    v_tpl = ws_tpl.cell(r, c).value
                    if isinstance(v_tpl, str) and v_tpl.startswith("="):
                        so_cong_thuc += 1
                        v_out = ws_out.cell(r, c).value
                        self.assertEqual(
                            v_out,
                            v_tpl,
                            f"Sai lệch công thức tại ô ({r},{c}) trên Sheet 'BC tháng': {v_out} != {v_tpl}",
                        )
            self.assertGreaterEqual(so_cong_thuc, 40)
            wb_tpl.close()
            wb_out.close()


if __name__ == "__main__":
    unittest.main()

