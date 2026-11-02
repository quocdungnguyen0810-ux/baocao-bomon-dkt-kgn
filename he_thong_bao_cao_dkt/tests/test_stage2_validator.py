# -*- coding: utf-8 -*-
"""
Unit Test & Kiểm toán Giai đoạn 2 (Audit Stage 2):
Kiểm tra Bộ Kiểm tra Ràng buộc Vật lý & Vùng Cách ly Dị biệt (validator.py).
"""

from datetime import date
import unittest

from he_thong_bao_cao_dkt.models import (
    BaoCaoCaNhanThang,
    DANH_SACH_GIANG_VIEN_DKT,
    HoatDongChuyenMon,
    NhiemVuCongViec,
    ThongKeUsmartDonVi,
)
from he_thong_bao_cao_dkt.validator import BoKiemTraRangBuocVatLy


class TestStage2Validator(unittest.TestCase):
    """Kiểm toán 9 quy tắc Physical Plausibility và cơ chế Vùng Cách ly (Quarantine)."""

    def setUp(self) -> None:
        self.validator = BoKiemTraRangBuocVatLy()
        self.gv_ha = [gv for gv in DANH_SACH_GIANG_VIEN_DKT if gv.ma_can_bo == "1051"][0]

    def test_01_bao_cao_hop_le_hoan_toan(self) -> None:
        """Báo cáo có số liệu hợp lý phải đạt hop_le_de_chot = True và 0 lỗi."""
        bc = BaoCaoCaNhanThang(
            giang_vien=self.gv_ha,
            thang=11,
            nam=2026,
            danh_sach_nhiem_vu=[
                NhiemVuCongViec(
                    ma_nhiem_vu="NV01",
                    nhiem_vu="Giảng dạy học phần Cơ học đất",
                    nguon_giao="Kế hoạch năm học",
                    so_van_ban_ma_viec="TKB-11",
                    san_pham_ket_qua="Hoàn thành 36 tiết giảng dạy",
                    thoi_han=date(2026, 11, 30),
                    ngay_hoan_thanh=date(2026, 11, 28),
                    tinh_trang="Hoàn thành đúng hạn",
                    muc_do_hoan_thanh=100.0,
                )
            ],
            hoat_dong_chuyen_mon=HoatDongChuyenMon(
                giang_day_thang=36.0,
                giang_day_luy_ke=110.0,
                giang_day_dinh_muc=270.0,
                nckh_thang=25.0,
                nckh_luy_ke=80.0,
                nckh_dinh_muc=250.0,
                nhiem_vu_khac_thang=15.0,
                nhiem_vu_khac_luy_ke=45.0,
                nhiem_vu_khac_dinh_muc=180.0,
                san_pham_khcn_huong_dan="Hướng dẫn 05 SV làm ĐATN",
            ),
            diem_tu_cham_tieu_chi_chung=29.0,
            diem_tu_cham_ket_qua_nhiem_vu=70.0,
        )
        kq = self.validator.kiem_tra_bao_cao_ca_nhan(bc)
        self.assertTrue(kq.hop_le_de_chot)
        self.assertEqual(kq.tong_so_loi_nghiem_trong, 0)
        self.assertEqual(len(kq.nhiem_vu_bi_cach_ly), 0)

    def test_02_phat_hien_sai_lech_ngay_va_cach_ly_nhiem_vu(self) -> None:
        """Phát hiện lỗi chọn 'Hoàn thành đúng hạn' nhưng Ngày HT > Thời hạn, thiếu thời hạn, thiếu mã việc uSmart."""
        nv_sai_han = NhiemVuCongViec(
            ma_nhiem_vu="NV_SAI_1",
            nhiem_vu="Nhiệm vụ hoàn thành trễ nhưng khai đúng hạn",
            nguon_giao="Giao việc trên Usmart",
            so_van_ban_ma_viec="",  # Lỗi thiếu mã việc uSmart
            san_pham_ket_qua="",  # Lỗi thiếu sản phẩm
            thoi_han=date(2026, 11, 10),
            ngay_hoan_thanh=date(2026, 11, 15),
            tinh_trang="Hoàn thành đúng hạn",  # Lỗi mâu thuẫn ngày hoàn thành
            muc_do_hoan_thanh=120.0,  # Lỗi vượt 100%
        )
        bc = BaoCaoCaNhanThang(
            giang_vien=self.gv_ha,
            thang=11,
            nam=2026,
            danh_sach_nhiem_vu=[nv_sai_han],
            hoat_dong_chuyen_mon=HoatDongChuyenMon(
                giang_day_thang=50.0,
                giang_day_luy_ke=20.0,  # Lỗi lũy kế < tháng
                san_pham_khcn_huong_dan="Có sản phẩm",
            ),
            diem_tu_cham_tieu_chi_chung=35.0,  # Lỗi vượt 30 điểm
            diem_tu_cham_ket_qua_nhiem_vu=75.0,  # Lỗi vượt 70 điểm
        )
        kq = self.validator.kiem_tra_bao_cao_ca_nhan(bc)
        self.assertFalse(kq.hop_le_de_chot)
        self.assertEqual(len(kq.nhiem_vu_bi_cach_ly), 1)
        ma_loi = {item.ma_quy_tac for item in kq.vung_cach_ly}
        self.assertIn("R1_MISSING_OUTPUT", ma_loi)
        self.assertIn("R2_DATE_STATUS_MISMATCH", ma_loi)
        self.assertIn("R3_INVALID_PERCENTAGE", ma_loi)
        self.assertIn("R6_USMART_MISSING_CODE", ma_loi)
        self.assertIn("R4_CUMULATIVE_LESS_THAN_MONTHLY", ma_loi)
        self.assertIn("R5_SCORE_GENERAL_OUT_OF_RANGE", ma_loi)
        self.assertIn("R5_SCORE_TASK_OUT_OF_RANGE", ma_loi)

    def test_03_kiem_tra_thong_ke_usmart_don_vi(self) -> None:
        """Kiểm tra công thức dòng Kiểm tra = 0 ở Mục II Mẫu 02."""
        tk_sai = ThongKeUsmartDonVi(
            so_van_ban_den_han=10,
            hoan_thanh_dung_han=7,
            hoan_thanh_cham=1,
            qua_han_chua_hoan_thanh=1,
        )
        loi = self.validator.kiem_tra_thong_ke_usmart_don_vi(tk_sai)
        self.assertEqual(len(loi), 1)
        self.assertEqual(loi[0].ma_quy_tac, "R7_USMART_SUM_MISMATCH")


if __name__ == "__main__":
    unittest.main()
