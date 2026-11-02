# -*- coding: utf-8 -*-
"""
Unit Test & Kiểm toán Giai đoạn 7 (Turnkey Deployment & Cycle Transition Audit):
Kiểm tra tính toàn vẹn của gói triển khai bàn giao (42 tệp báo cáo 21 GV, Mẫu 02, Mẫu 03,
Google Sheets NCKH & TKB, Google Apps Script, tập lệnh 1-Click) và chu kỳ chuyển tiếp sang Tháng 12/2026.
"""

from pathlib import Path
import tempfile
import unittest

from he_thong_bao_cao_dkt.database_manager import QuanLyCSDLBộMôn
from he_thong_bao_cao_dkt.excel_generator import PROJECT_ROOT
from he_thong_bao_cao_dkt.models import DANH_SACH_GIANG_VIEN_DKT


class TestStage7DeploymentAudit(unittest.TestCase):
    """Kiểm toán gói triển khai bàn giao hoàn thiện (Giai đoạn 7)."""

    def test_01_kiem_tra_tron_bo_ho_so_xuat_thuc_te_21_gv(self) -> None:
        """Đảm bảo đủ 42 tệp (.xlsx & .docx) của 21 giảng viên + Mẫu 02 + Mẫu 03 tồn tại trên ổ đĩa."""
        ky_dir = PROJECT_ROOT / "output_bao_cao_dkt" / "Ky_11_2026"
        self.assertTrue((ky_dir / "BCThang_DKT_NguyenHaiHa_11-2026.xlsx").exists())
        self.assertTrue((ky_dir / "BCThang_DKT_NguyenHaiHa_11-2026.docx").exists())
        self.assertTrue((ky_dir / "BCThang_DKT_11-2026.xlsx").exists())
        self.assertTrue((ky_dir / "BCThang_DKT_DonVi_11-2026.docx").exists())
        self.assertTrue((ky_dir / "SoGiaoViec_DKT_Mau03_11-2026.xlsx").exists())

        gv_dir = ky_dir / "BaoCao_CaNhan_21_GV"
        ds_xlsx = list(gv_dir.glob("BCThang_DKT_*_11-2026.xlsx"))
        ds_docx = list(gv_dir.glob("BCThang_DKT_*_11-2026.docx"))
        self.assertEqual(len(ds_xlsx), 21)
        self.assertEqual(len(ds_docx), 21)

    def test_02_kiem_tra_goi_google_sheets_va_apps_script(self) -> None:
        """Đảm bảo các tệp Google Sheets Kế hoạch NCKH, TKB và Google Apps Script đồng bộ 21 GV."""
        kh_dir = PROJECT_ROOT / "Ke_hoach_cong_viec"
        self.assertTrue((kh_dir / "Ke_hoach_NCKH_DeTai_CacCap_va_NCKHSV_DKT_KGN.xlsx").exists())
        self.assertTrue((kh_dir / "Ke_hoach_GiangDay_TKB_DKT_KGN_2026_2027.xlsx").exists())
        self.assertTrue((PROJECT_ROOT / "CSDL_BaoCao_BoMon_DKT_KGN.xlsx").exists())
        self.assertTrue((PROJECT_ROOT / "CHAY_HE_THONG_BAO_CAO_DKT.bat").exists())
        self.assertTrue((PROJECT_ROOT / "README.md").exists())

        gs_sync = (PROJECT_ROOT / "he_thong_bao_cao_dkt" / "google_apps_script" / "NCKH_USmart_Sync.gs").read_text(
            encoding="utf-8"
        )
        for gv in DANH_SACH_GIANG_VIEN_DKT:
            self.assertIn(f'"{gv.ma_can_bo}": "{gv.ho_va_ten}"', gs_sync)

    def test_03_kiem_tra_chuyen_tiep_chu_ky_thang_12_2026(self) -> None:
        """Kiểm tra tự động khởi tạo công việc lặp lại sang kỳ Tháng 12/2026 cho toàn bộ 21 giảng viên."""
        with tempfile.TemporaryDirectory() as tmp_dir:
            csdl_tmp = Path(tmp_dir) / "CSDL_Cycle_12_2026.xlsx"
            db = QuanLyCSDLBộMôn(csdl_tmp)
            db.khoi_tao_csdl_mau_neu_chua_co(ghi_de=True)
            so_moi = db.sao_chep_cong_viec_lap_lai_sang_thang_moi(12, 2026)
            self.assertGreaterEqual(so_moi, 42)
            bc_t12 = db.doc_bao_cao_ca_nhan("1051", thang=12, nam=2026)
            self.assertGreaterEqual(len(bc_t12.danh_sach_nhiem_vu), 2)
            self.assertEqual(bc_t12.ten_file_xlsx_chuan, "BCThang_DKT_NguyenHaiHa_12-2026.xlsx")

    def test_04_kiem_tra_phan_quyen_admin_an_dong_bo_du_lieu(self) -> None:
        """Đảm bảo nội dung hiển thị & đồng bộ dữ liệu nguồn chỉ mở khóa khi đăng nhập tài khoản Admin."""
        from he_thong_bao_cao_dkt.app_webapp import xac_thuc_tai_khoan_admin

        # Tài khoản trống hoặc tài khoản giảng viên thường không có quyền Admin
        self.assertFalse(xac_thuc_tai_khoan_admin("", ""))
        self.assertFalse(xac_thuc_tai_khoan_admin("hanh1", "182128"))
        self.assertFalse(xac_thuc_tai_khoan_admin("admin", "sai_mat_khau"))

        # Chỉ tài khoản Admin mới được xác thực hợp lệ
        self.assertTrue(xac_thuc_tai_khoan_admin("admin", "dkt412@admin"))
        self.assertTrue(xac_thuc_tai_khoan_admin("lannc", "giangchau0310"))
        self.assertTrue(xac_thuc_tai_khoan_admin("lannc", "giangchau"))
        self.assertTrue(xac_thuc_tai_khoan_admin("truyentt", "dkt412@admin"))

    def test_05_tuy_chon_su_dung_lai_viec_cu_doi_ngay_cham_diem_va_trong_tam(self) -> None:
        """
        Kiểm tra chức năng hiển thị công việc báo cáo cũ để tái sử dụng, thay đổi hạng mục,
        ngày hoàn thành, tình trạng, chuyển thành nhiệm vụ trọng tâm đơn vị và đổi điểm đánh giá.
        """
        from datetime import date
        from he_thong_bao_cao_dkt.models import HoatDongChuyenMon

        with tempfile.TemporaryDirectory() as tmp_dir:
            csdl_tmp = Path(tmp_dir) / "CSDL_Interactive_Edit.xlsx"
            db = QuanLyCSDLBộMôn(csdl_tmp)
            db.khoi_tao_csdl_mau_neu_chua_co(ghi_de=True)

            # 1. Ở kỳ Tháng 12/2026 (chưa có việc), lấy danh sách gợi ý từ báo cáo cũ T11/2026 & ngân hàng lặp lại
            goi_y = db.lay_cong_viec_cu_va_lap_lai_goi_y("1051", thang_hien_tai=12, nam_hien_tai=2026)
            self.assertGreaterEqual(len(goi_y), 3)

            # 2. Người dùng tích chọn sử dụng 2 việc cũ, đổi ngày hoàn thành, đổi hạng mục và bật cờ Trọng tâm ĐV
            nv_chon_1 = goi_y[0]
            nv_chon_1.nhiem_vu = "Giảng dạy Cơ học đất & Nền móng (đã điều chỉnh hạng mục T12/2026)"
            nv_chon_1.ngay_hoan_thanh = date(2026, 12, 25)
            nv_chon_1.thoi_han = date(2026, 12, 28)
            nv_chon_1.tinh_trang = "Hoàn thành đúng hạn"
            nv_chon_1.danh_gia_chat_luong = "Đạt yêu cầu – tốt"
            nv_chon_1.la_nhiem_vu_trong_tam = True
            nv_chon_1.la_cong_viec_lap_lai = True

            nv_chon_2 = goi_y[1]
            nv_chon_2.ngay_hoan_thanh = date(2026, 12, 30)
            nv_chon_2.thoi_han = date(2026, 12, 29)
            nv_chon_2.tinh_trang = "Hoàn thành chậm"
            nv_chon_2.la_nhiem_vu_trong_tam = False
            nv_chon_2.la_cong_viec_lap_lai = False

            so_luu = db.cap_nhat_danh_sach_nhiem_vu_thang("1051", 12, 2026, [nv_chon_1, nv_chon_2])
            self.assertEqual(so_luu, 2)

            # 3. Cập nhật điểm đánh giá và số giờ chuyên môn
            hd_cm = HoatDongChuyenMon(
                giang_day_thang=42.5,
                giang_day_luy_ke=166.6,
                giang_day_dinh_muc=187.5,
                nckh_thang=35.0,
                nckh_luy_ke=130.0,
                nckh_dinh_muc=250.0,
                nhiem_vu_khac_thang=15.0,
                nhiem_vu_khac_luy_ke=75.0,
                nhiem_vu_khac_dinh_muc=180.0,
                san_pham_khcn_huong_dan="Hoàn thành hướng dẫn NCKHSV và nghiệm thu đề tài cấp Trường.",
            )
            db.cap_nhat_cham_diem_va_chuyen_mon(
                ma_can_bo="1051",
                thang=12,
                nam=2026,
                hoat_dong_cm=hd_cm,
                diem_tu_cham_chung=28.5,
                diem_tu_cham_nv=65.0,
                de_xuat_kien_nghi="Đề nghị bổ sung thiết bị PTN Địa kỹ thuật.",
            )

            # 3b. Cập nhật Mục IV. Kế hoạch công việc tháng tới
            from he_thong_bao_cao_dkt.models import KeHoachKyToi

            so_kh = db.cap_nhat_ke_hoach_thang_toi(
                ma_can_bo="1051",
                thang=12,
                nam=2026,
                danh_sach_ke_hoach=[
                    KeHoachKyToi(
                        nhiem_vu="Tổ chức coi chấm thi kết thúc học phần Học kỳ I năm học 2026-2027",
                        san_pham_du_kien="Bảng điểm học phần và túi bài thi lưu trữ",
                        thoi_han=date(2027, 1, 25),
                        ghi_chu="Theo lịch thi Phòng ĐTĐH",
                    )
                ],
            )
            self.assertEqual(so_kh, 1)

            # 4. Đọc lại báo cáo và kiểm tra toàn bộ thay đổi đã được lưu chính xác
            bc_sau = db.doc_bao_cao_ca_nhan("1051", thang=12, nam=2026)
            self.assertEqual(len(bc_sau.danh_sach_nhiem_vu), 2)
            self.assertEqual(
                bc_sau.danh_sach_nhiem_vu[0].nhiem_vu,
                "Giảng dạy Cơ học đất & Nền móng (đã điều chỉnh hạng mục T12/2026)",
            )
            self.assertEqual(bc_sau.danh_sach_nhiem_vu[0].ngay_hoan_thanh, date(2026, 12, 25))
            self.assertTrue(bc_sau.danh_sach_nhiem_vu[0].la_nhiem_vu_trong_tam)
            self.assertEqual(bc_sau.diem_tu_cham_tieu_chi_chung, 28.5)
            self.assertEqual(bc_sau.diem_tu_cham_ket_qua_nhiem_vu, 65.0)
            self.assertAlmostEqual(bc_sau.tong_diem_tu_cham, 93.5)
            self.assertEqual(bc_sau.hoat_dong_chuyen_mon.giang_day_thang, 42.5)
            self.assertEqual(len(bc_sau.ke_hoach_thang_toi), 1)
            self.assertIn("Tổ chức coi chấm thi", bc_sau.ke_hoach_thang_toi[0].nhiem_vu)

    def test_06_kiem_tra_kieu_bao_cao_tuan_thang_don_vi_va_so_giao_viec(self) -> None:
        """
        Kiểm tra đầy đủ các kiểu báo cáo:
        - Báo cáo công việc cá nhân: Mẫu 01a (BC tuần) & Mẫu 01b (BC tháng)
        - Báo cáo công việc đơn vị: Mẫu 02a (BC tuần ĐV), Mẫu 02b (BC tháng ĐV), Mẫu 02c (Theo dõi cá nhân)
        - Giao việc và theo dõi nhiệm vụ: Mẫu 03a (Phiếu giao việc), Mẫu 03b (Sổ giao việc), Mẫu 03c (Tổng hợp theo người)
        """
        import openpyxl
        from he_thong_bao_cao_dkt.excel_generator import BoSinhBaoCaoExcel

        with tempfile.TemporaryDirectory() as tmp_dir:
            tmp_path = Path(tmp_dir)
            csdl_tmp = tmp_path / "CSDL_Full_Report_Types.xlsx"
            db = QuanLyCSDLBộMôn(csdl_tmp)
            db.khoi_tao_csdl_mau_neu_chua_co(ghi_de=True)
            excel_gen = BoSinhBaoCaoExcel()

            # 1. Cập nhật nhiệm vụ trọng tâm và điểm Trưởng đơn vị chấm (Mẫu 02c)
            so_cn = db.cap_nhat_trong_tam_va_diem_tbm_don_vi(
                thang=11,
                nam=2026,
                map_diem_tbm_theo_ma_cb={"1051": (29.0, 70.0), "0694": (30.0, 70.0)},
            )
            self.assertGreaterEqual(so_cn, 2)

            # 2. Xuất Mẫu 01 (có cả Sheet 'BC tuần' Mẫu 01a và 'BC tháng' Mẫu 01b)
            bc_1051 = db.doc_bao_cao_ca_nhan("1051", thang=11, nam=2026)
            self.assertEqual(bc_1051.tong_diem_truong_don_vi, 99.0)
            f01 = excel_gen.xuat_mau_01_ca_nhan(bc_1051, tmp_path, cap_nhat_ca_bc_tuan=True)
            wb01 = openpyxl.load_workbook(f01, data_only=False)
            self.assertIn("BC tuần", wb01.sheetnames)
            self.assertIn("BC tháng", wb01.sheetnames)
            self.assertEqual(wb01["BC tuần"]["C7"].value, "Nguyễn Hải Hà")
            self.assertEqual(wb01["BC tháng"]["C7"].value, "Nguyễn Hải Hà")
            wb01.close()

            # 3. Xuất Mẫu 02 (có đủ 'BC tuần ĐV' Mẫu 02a, 'BC tháng ĐV' Mẫu 02b, 'Theo dõi cá nhân' Mẫu 02c)
            ds_all = db.doc_toan_bo_bao_cao_bo_mon(thang=11, nam=2026)
            f02 = excel_gen.xuat_mau_02_don_vi(ds_all, tmp_path, thang=11, nam=2026)
            wb02 = openpyxl.load_workbook(f02, data_only=False)
            self.assertIn("BC tuần ĐV", wb02.sheetnames)
            self.assertIn("BC tháng ĐV", wb02.sheetnames)
            self.assertIn("Theo dõi cá nhân", wb02.sheetnames)
            self.assertIsNotNone(wb02["BC tuần ĐV"]["B12"].value)
            self.assertIsNotNone(wb02["BC tháng ĐV"]["B12"].value)
            self.assertEqual(wb02["Theo dõi cá nhân"]["B7"].value, "Trần Thế Truyền")
            wb02.close()

            # 4. Xuất Mẫu 03 (có đủ 'Phiếu giao việc' Mẫu 03a, 'Sổ giao việc' Mẫu 03b, 'Tổng hợp theo người' Mẫu 03c)
            f03 = excel_gen.xuat_mau_03_so_giao_viec(ds_all, tmp_path, thang=11, nam=2026)
            wb03 = openpyxl.load_workbook(f03, data_only=False)
            self.assertIn("Phiếu giao việc", wb03.sheetnames)
            self.assertIn("Sổ giao việc", wb03.sheetnames)
            self.assertIn("Tổng hợp theo người", wb03.sheetnames)
            wb03.close()

    def test_07_an_bao_cao_don_vi_voi_tai_khoan_thuong_va_chon_vi_tri_luu_file(self) -> None:
        """
        Kiểm tra:
        1. Module Báo cáo công việc đơn vị (Mẫu 02) chỉ hiển thị với tài khoản Admin.
        2. Chức năng chọn vị trí lưu file (`xac_dinh_thu_muc_luu_file`) hỗ trợ thư mục tùy chỉnh & tạo thư mục tự động.
        3. Cấu hình Webapp chi phí tối ưu (`streamlit_app.py`, `requirements.txt`, `.streamlit/config.toml`).
        """
        from he_thong_bao_cao_dkt.app_webapp import (
            lay_danh_sach_kieu_bao_cao_theo_quyen,
            xac_dinh_thu_muc_luu_file,
        )

        # 1. Tài khoản thường KHÔNG thấy Báo cáo công việc đơn vị (Mẫu 02) và KHÔNG thấy Đồng bộ dữ liệu nguồn
        ds_thuong = lay_danh_sach_kieu_bao_cao_theo_quyen(is_admin=False)
        self.assertEqual(len(ds_thuong), 2)
        self.assertFalse(any("Mẫu 02" in x or "đơn vị" in x for x in ds_thuong))
        self.assertFalse(any("Đồng bộ" in x for x in ds_thuong))

        # 2. Tài khoản Admin thấy đầy đủ Báo cáo công việc đơn vị (Mẫu 02) và Đồng bộ dữ liệu nguồn
        ds_admin = lay_danh_sach_kieu_bao_cao_theo_quyen(is_admin=True)
        self.assertEqual(len(ds_admin), 4)
        self.assertTrue(any("Mẫu 02" in x for x in ds_admin))
        self.assertTrue(any("Đồng bộ" in x for x in ds_admin))

        # 3. Kiểm tra chọn vị trí lưu file tùy chỉnh
        with tempfile.TemporaryDirectory() as tmp_dir:
            custom_dir = Path(tmp_dir) / "Thu_muc_luu_tuy_chinh_DKT"
            resolved = xac_dinh_thu_muc_luu_file("✍️ Tự nhập đường dẫn thư mục tùy chỉnh...", str(custom_dir))
            self.assertEqual(resolved, custom_dir)
            self.assertTrue(resolved.exists())

        # 4. Kiểm tra các tệp cấu hình Webapp Chi phí tối ưu (0 đồng)
        self.assertTrue((PROJECT_ROOT / "streamlit_app.py").exists())
        self.assertTrue((PROJECT_ROOT / "requirements.txt").exists())
        self.assertTrue((PROJECT_ROOT / ".streamlit" / "config.toml").exists())


if __name__ == "__main__":
    unittest.main()
