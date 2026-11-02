# -*- coding: utf-8 -*-
"""
Unit Test & Kiểm toán Giai đoạn 4 (Audit Stage 4 - Live uSmart Integration):
Kiểm thử kết nối thực tế tới `https://usmart.utc.edu.vn/baocaocongviec/`
với tài khoản `hanh1` (mật khẩu `182128`), gửi báo cáo tháng 11/2026 (Mẫu 01b)
ở chế độ Lưu nháp (`luu_nhap`) và đọc ngược lại từ máy chủ uSmart để đối chiếu từng trường.
"""

from pathlib import Path
import tempfile
import unittest

from he_thong_bao_cao_dkt.database_manager import QuanLyCSDLBộMôn
from he_thong_bao_cao_dkt.usmart_baocao_client import USmartBaoCaoAgent


class TestStage4USmartLive(unittest.TestCase):
    """Kiểm toán trực tiếp luồng OIDC và lưu nháp báo cáo Mẫu 01b trên máy chủ uSmart."""

    def setUp(self) -> None:
        self.tmp_dir = tempfile.TemporaryDirectory()
        self.db = QuanLyCSDLBộMôn(Path(self.tmp_dir.name) / "CSDL_Test.xlsx")
        self.db.khoi_tao_csdl_mau_neu_chua_co(ghi_de=True)
        self.agent = USmartBaoCaoAgent("hanh1", "182128")

    def tearDown(self) -> None:
        self.tmp_dir.cleanup()

    def test_01_dang_nhap_va_lay_ky_bao_cao_usmart(self) -> None:
        """Kiểm tra đăng nhập OIDC và lấy thông tin giảng viên Nguyễn Hải Hà trên uSmart."""
        ok = self.agent.dang_nhap()
        self.assertTrue(ok)
        self.assertIn("Nguyễn Hải Hà", self.agent.ten_nguoi_dung_usmart)
        self.assertIn("Địa kỹ thuật", self.agent.don_vi_usmart)
        self.assertEqual(self.agent.nguoi_nhan_usmart, "Trần Thế Truyền")

        ds_ky = self.agent.lay_danh_sach_ky_bao_cao()
        self.assertGreaterEqual(len(ds_ky), 2)
        ma_mau_list = [k.ma_mau for k in ds_ky]
        self.assertIn("01a", ma_mau_list)
        self.assertIn("01b", ma_mau_list)

    def test_02_luu_nhap_mau_01b_thang_11_2026_va_doi_chieu_server(self) -> None:
        """Gửi bản nháp Mẫu 01b tháng 11/2026 lên uSmart và đọc ngược lại để kiểm toán."""
        bc_ha = self.db.doc_bao_cao_ca_nhan("1051", thang=11, nam=2026)
        kq = self.agent.gui_bao_cao_webform(
            bc_ha, ky_id="2", ma_mau="01b", che_do="luu_nhap"
        )
        self.assertTrue(kq.thanh_cong)
        self.assertEqual(kq.http_status, 200)

        # Kiểm toán dữ liệu đọc ngược từ server uSmart
        server_data = kq.du_lieu_xac_nhan_tren_server
        self.assertEqual(server_data["so_nhiem_vu"], 5)
        self.assertEqual(server_data["cells"].get("gd_thang"), "36")
        self.assertEqual(server_data["cells"].get("nc_thang"), "30")
        self.assertEqual(server_data["cells"].get("tu_cham_chung"), "29")
        self.assertEqual(server_data["cells"].get("tu_cham_nv"), "70")


if __name__ == "__main__":
    unittest.main()
