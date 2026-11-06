# -*- coding: utf-8 -*-
"""
Bộ kiểm thử tự động Giai đoạn 6 (NCKH Sync):
Kiểm tra Tác tử Đồng bộ Kế hoạch NCKH Đề tài các cấp & Kế hoạch NCKH Sinh viên (`NCKHSV`)
từ cổng `https://usmart.utc.edu.vn/qlhc/nckh/view/nckh/detai/tat-ca-de-tai` (tài khoản `lannc` / `giangchau`),
Kiểm toán Ràng buộc Vật lý (`NCKH-R1`..`NCKH-R5`) và Xuất Google Sheet 9 trang trong thư mục `Kế hoạch công việc`.
"""

from pathlib import Path
import tempfile
import unittest
import openpyxl

from he_thong_bao_cao_dkt.database_manager import QuanLyCSDLBộMôn
from he_thong_bao_cao_dkt.usmart_nckh_sync import (
    BoXuLyKeHoachNCKH,
    CACHE_FILE_DEFAULT,
    FOLDER_KE_HOACH_CONG_VIEC,
    TEN_FILE_GSHEET_NCKH,
    USmartNCKHClient,
    chay_dong_bo_ke_hoach_nckh_toan_trinh,
)
from he_thong_bao_cao_dkt.validator import BoKiemTraRangBuocVatLy


class TestStage6NCKHSync(unittest.TestCase):
    """Kiểm thử độc lập và tích hợp cho Tác tử Đồng bộ uSmart KHCN (TK lannc)."""

    def test_01_usmart_nckh_client_sso_login_lannc(self) -> None:
        """Kiểm tra đăng nhập OIDC SSO với tài khoản `lannc` và mật khẩu `giangchau`."""
        client = USmartNCKHClient("lannc", "giangchau")
        ok = client.dang_nhap_sso()
        self.assertTrue(ok)
        self.assertIsNotNone(client._access_token)
        self.assertGreater(len(client._access_token or ""), 50)

    def test_02_bo_xu_ly_ke_hoach_nckh_va_vung_cach_ly(self) -> None:
        """
        Kiểm tra chuẩn hóa dữ liệu 307 đề tài từ cache uSmart của Bộ môn ĐKT & KGN
        và xác minh nguyên tắc an toàn số liệu (đưa vào Vùng cách ly thay vì xóa bỏ).
        """
        self.assertTrue(CACHE_FILE_DEFAULT.exists(), "Tệp cache uSmart NCKH phải tồn tại trong Ke_hoach_cong_viec")
        client = USmartNCKHClient("lannc", "giangchau")
        du_lieu_raw, _ = client.tai_va_cap_nhat_du_lieu(dong_bo_live=False)

        processor = BoXuLyKeHoachNCKH(du_lieu_raw)
        ds_cac_cap, ds_nckhsv, ds_cach_ly = processor.chuan_hoa_va_kiem_toan()

        # Tổng số đề tài = 307 (50 đề tài các cấp + 257 đề tài NCKHSV)
        self.assertEqual(len(ds_cac_cap) + len(ds_nckhsv), 307)
        self.assertGreaterEqual(len(ds_cac_cap), 45)
        self.assertGreaterEqual(len(ds_nckhsv), 250)

        # Phải phát hiện và cách ly các đề tài KHOA_KHONG_DUYET (NCKH-R1) mà không làm mất bản ghi gốc
        ma_quy_tac_cl = {item.ma_quy_tac for item in ds_cach_ly}
        self.assertIn("NCKH-R1", ma_quy_tac_cl)

        # Kiểm tra các đề tài cấp Trường T2027 (idDotDangKy=9177) và NCKHSV 2026-2027 (idDotDangKy=9176)
        dt_t2027 = [x for x in ds_cac_cap if x.id_dot_dang_ky == 9177 and x.hop_le_dua_vao_ke_hoach]
        self.assertEqual(len(dt_t2027), 7)
        dt_sv_2627 = [x for x in ds_nckhsv if x.id_dot_dang_ky == 9176 and x.hop_le_dua_vao_ke_hoach]
        self.assertEqual(len(dt_sv_2627), 11)

    def test_03_xuat_google_sheet_9_sheets_va_dong_bo_csdl(self) -> None:
        """
        Kiểm tra xuất sổ tính Google Sheet 9 trang và đồng bộ vào CSDL Báo cáo công việc tháng,
        đảm bảo 100% nhiệm vụ đồng bộ vượt qua bộ Kiểm toán Ràng buộc Vật lý (TB 940).
        """
        with tempfile.TemporaryDirectory() as tmp_dir_str:
            csdl_tmp = Path(tmp_dir_str) / "CSDL_BaoCao_Test_NCKH.xlsx"
            kq = chay_dong_bo_ke_hoach_nckh_toan_trinh(
                username="lannc",
                password="giangchau",
                thang=11,
                nam=2026,
                dong_bo_live=False,
                duong_dan_csdl_master=csdl_tmp,
            )

            self.assertTrue(kq.duong_dan_gsheet_project.exists())
            self.assertEqual(kq.duong_dan_gsheet_project, FOLDER_KE_HOACH_CONG_VIEC / TEN_FILE_GSHEET_NCKH)
            self.assertEqual(kq.so_de_tai_cac_cap_2025_2027, 25)
            self.assertEqual(kq.so_de_tai_nckhsv_2025_2027, 77)
            self.assertGreaterEqual(kq.so_nhiem_vu_dong_bo_bao_cao_thang, 50)

            # Kiểm tra cấu trúc 9 trang tính của Google Sheet Kế hoạch NCKH & NCKHSV
            wb = openpyxl.load_workbook(kq.duong_dan_gsheet_project, data_only=False)
            expected_sheets = [
                "01_Dashboard_TongHop",
                "02_KH_DeTai_NCKH_2025_2027",
                "03_KH_NCKHSV_2025_2027",
                "04_TongHop_21GV_KhaiBao_BaoCao",
                "05_CSDL_TatCa_DeTai_CacCap",
                "06_CSDL_TatCa_NCKHSV",
                "07_LichTrinh_DinhMuc_PKHCN",
                "08_NhiemVu_DongBo_Mau01b_02b",
                "09_VungCachLy_KiemToan_NCKH",
            ]
            self.assertEqual(wb.sheetnames, expected_sheets)
            wb.close()

            # Kiểm tra CSDL Báo cáo tháng sau khi đồng bộ NCKH phải đạt 100% kiểm toán vật lý cho cả 22 thành viên
            db = QuanLyCSDLBộMôn(csdl_tmp)
            validator = BoKiemTraRangBuocVatLy()
            ds_bc = db.doc_toan_bo_bao_cao_bo_mon(thang=11, nam=2026)
            self.assertEqual(len(ds_bc), 22)
            for bc in ds_bc:
                audit = validator.kiem_tra_bao_cao_ca_nhan(bc)
                self.assertTrue(audit.hop_le_de_chot, f"GV {bc.giang_vien.ho_va_ten} có lỗi kiểm toán: {audit.vung_cach_ly}")


if __name__ == "__main__":
    unittest.main()
