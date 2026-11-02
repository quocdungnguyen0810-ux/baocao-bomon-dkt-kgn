# -*- coding: utf-8 -*-
"""
Unit Test & Kiểm toán Giai đoạn 5 (Audit Stage 5 - End-to-End Pipeline):
Kiểm toán toàn trình từ CSDL Bộ môn -> Kiểm toán Vật lý -> Xuất Mẫu 01, 02, 03 (.xlsx & .docx)
cho toàn bộ 21 giảng viên Bộ môn Địa kỹ thuật và Không gian ngầm -> Đồng bộ trực tiếp uSmart.
"""

from pathlib import Path
import tempfile
import unittest

from he_thong_bao_cao_dkt.cli_baocao import TacTuDieuPhoiBaoCao


class TestStage5EndToEnd(unittest.TestCase):
    """Kiểm toán tích hợp toàn bộ hệ thống tác tử báo cáo công việc Bộ môn ĐKT & KGN."""

    def setUp(self) -> None:
        self.tmp_dir = tempfile.TemporaryDirectory()
        self.base_dir = Path(self.tmp_dir.name)
        self.orchestrator = TacTuDieuPhoiBaoCao(
            duong_dan_csdl=self.base_dir / "CSDL_BaoCao_BoMon_DKT_KGN.xlsx",
            thu_muc_xuat=self.base_dir / "output_bao_cao_dkt",
        )

    def tearDown(self) -> None:
        self.tmp_dir.cleanup()

    def test_01_chay_pipeline_toan_trinh_21_gv_va_usmart(self) -> None:
        """Chạy toàn bộ pipeline End-to-End kèm xuất 21 báo cáo giảng viên và lưu nháp uSmart."""
        kq = self.orchestrator.chay_pipeline_hoan_chinh(
            ma_can_bo="1051",
            thang=11,
            nam=2026,
            xuat_toan_bo_21_gv=True,
            usmart_username="hanh1",
            usmart_password="182128",
            che_do_usmart="luu_nhap",
        )
        self.assertTrue(kq.duong_dan_csdl.exists())
        self.assertTrue(kq.ket_qua_kiem_toan_ca_nhan.hop_le_de_chot)
        self.assertTrue(kq.file_xlsx_ca_nhan.exists())
        self.assertTrue(kq.file_docx_ca_nhan.exists())
        self.assertTrue(kq.file_xlsx_mau_02_don_vi.exists())
        self.assertTrue(kq.file_docx_mau_02_don_vi.exists())
        self.assertTrue(kq.file_xlsx_mau_03_so_giao_viec.exists())
        # 2 file chính + 42 file của 21 giảng viên
        self.assertEqual(len(kq.danh_sach_file_toan_bo_gv), 44)
        self.assertIsNotNone(kq.ket_qua_usmart)
        self.assertTrue(kq.ket_qua_usmart.thanh_cong)


if __name__ == "__main__":
    unittest.main()
