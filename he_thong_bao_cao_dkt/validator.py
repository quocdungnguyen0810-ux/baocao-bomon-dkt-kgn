# -*- coding: utf-8 -*-
"""
Module Kiểm tra Ràng buộc Vật lý & Vùng Cách ly Dữ liệu Dị biệt (Physical Plausibility & Quarantine)
cho Hệ thống Báo cáo Công việc Bộ môn Địa kỹ thuật và Không gian ngầm.

Tuân thủ Nguyên tắc An toàn Số liệu:
- Tuyệt đối KHÔNG tự ý suy diễn hoặc bịa số liệu khi thiếu thông tin.
- Khi gặp lỗi hoặc điểm dị biệt (outlier), thông báo rõ nguyên nhân và đưa vào vùng cách ly
  (Quarantine) thay vì tự ý xóa bỏ.
"""

from dataclasses import dataclass, field
from typing import List, Optional, Set, Tuple
import logging

from .models import (
    BaoCaoCaNhanThang,
    DANH_MUC_NGUON_GIAO,
    DANH_MUC_TINH_TRANG,
    NhiemVuCongViec,
    ThongKeUsmartDonVi,
)

logger = logging.getLogger(__name__)

# Ngưỡng vật lý tối đa hợp lý trong 1 tháng cho 1 giảng viên (cảnh báo outlier)
NGUONG_GIO_GIANG_DAY_THANG_MAX: float = 200.0
NGUONG_GIO_NCKH_THANG_MAX: float = 250.0
NGUONG_GIO_KHAC_THANG_MAX: float = 200.0


@dataclass
class BanGhiCachLy:
    """Thông tin một lỗi hoặc điểm dị biệt (Outlier) được đưa vào Vùng Cách ly."""

    ma_quy_tac: str
    muc_do: str  # "ERROR" (Chặn chốt báo cáo) hoặc "WARNING" (Cảnh báo điểm dị biệt cần xác nhận)
    doi_tuong: str  # Tên nhiệm vụ, mục chuyên môn hoặc tên giảng viên
    nguyen_nhan: str
    huong_dan_xu_ly: str
    du_lieu_goc: Optional[object] = None


@dataclass
class KetQuaKiemToanBaoCao:
    """Kết quả kiểm tra ràng buộc vật lý và nghiệp vụ của một báo cáo cá nhân hoặc đơn vị."""

    hop_le_de_chot: bool
    tong_so_loi_nghiem_trong: int
    tong_so_canh_bao_di_biet: int
    vung_cach_ly: List[BanGhiCachLy] = field(default_factory=list)
    nhiem_vu_hop_le: List[NhiemVuCongViec] = field(default_factory=list)
    nhiem_vu_bi_cach_ly: List[NhiemVuCongViec] = field(default_factory=list)

    def tom_tat_kiem_toan(self) -> str:
        trang_thai = "ĐẠT (SẴN SÀNG CHỐT & NỘP)" if self.hop_le_de_chot else "CHƯA ĐẠT (CÓ BẢN GHI TRONG VÙNG CÁCH LY)"
        return (
            f"[{trang_thai}] Lỗi chặn: {self.tong_so_loi_nghiem_trong} | "
            f"Cảnh báo dị biệt (Outlier): {self.tong_so_canh_bao_di_biet} | "
            f"Nhiệm vụ hợp lệ: {len(self.nhiem_vu_hop_le)} | "
            f"Nhiệm vụ cách ly: {len(self.nhiem_vu_bi_cach_ly)}"
        )


class BoKiemTraRangBuocVatLy:
    """
    Bộ máy kiểm toán 9 quy tắc ràng buộc vật lý (Physical Plausibility) và quy định
    Thông báo số 940/TB-ĐHGTVT trước khi xuất báo cáo hoặc gửi lên uSmart.
    """

    def kiem_tra_bao_cao_ca_nhan(
        self, bao_cao: BaoCaoCaNhanThang
    ) -> KetQuaKiemToanBaoCao:
        vung_cach_ly: List[BanGhiCachLy] = []
        nv_hop_le: List[NhiemVuCongViec] = []
        nv_cach_ly: List[NhiemVuCongViec] = []

        da_gap_ten_nv: Set[str] = set()
        da_gap_so_vb: Set[str] = set()

        if not bao_cao.danh_sach_nhiem_vu:
            vung_cach_ly.append(
                BanGhiCachLy(
                    ma_quy_tac="R0_EMPTY_TASKS",
                    muc_do="ERROR",
                    doi_tuong=bao_cao.giang_vien.ho_va_ten,
                    nguyen_nhan="Báo cáo tháng chưa có bất kỳ nhiệm vụ nào được kê khai.",
                    huong_dan_xu_ly="Bấm 'Đồng bộ KH Giảng dạy & NCKH' hoặc 'Sao chép công việc lặp lại' để nạp nhiệm vụ.",
                )
            )

        for idx, nv in enumerate(bao_cao.danh_sach_nhiem_vu, start=1):
            co_loi_nv = False
            ten_ht = f"Dòng {idx} ({nv.ma_nhiem_vu or 'Không mã'}): {nv.nhiem_vu[:50]}"

            # Quy tắc R1: Kiểm tra thiếu Thời hạn hoặc Sản phẩm đầu ra
            if not nv.nhiem_vu.strip():
                vung_cach_ly.append(
                    BanGhiCachLy(
                        ma_quy_tac="R1_MISSING_NAME",
                        muc_do="ERROR",
                        doi_tuong=ten_ht,
                        nguyen_nhan="Nội dung nhiệm vụ, công việc bị để trống.",
                        huong_dan_xu_ly="Nhập tên nhiệm vụ cụ thể hoặc xóa dòng trống.",
                        du_lieu_goc=nv,
                    )
                )
                co_loi_nv = True

            if nv.thoi_han is None:
                vung_cach_ly.append(
                    BanGhiCachLy(
                        ma_quy_tac="R1_MISSING_DEADLINE",
                        muc_do="ERROR",
                        doi_tuong=ten_ht,
                        nguyen_nhan="Thiếu 'Thời hạn' hoàn thành (Bắt buộc theo TB 940; uSmart không tính đúng hạn nếu thiếu thời hạn).",
                        huong_dan_xu_ly="Bổ sung ngày Thời hạn (dd/mm/yyyy) dựa trên kế hoạch hoặc sổ giao việc.",
                        du_lieu_goc=nv,
                    )
                )
                co_loi_nv = True

            if not nv.san_pham_ket_qua.strip():
                vung_cach_ly.append(
                    BanGhiCachLy(
                        ma_quy_tac="R1_MISSING_OUTPUT",
                        muc_do="ERROR",
                        doi_tuong=ten_ht,
                        nguyen_nhan="Thiếu 'Sản phẩm đầu ra, kết quả' cụ thể (TB 940 mục 1c yêu cầu rõ sản phẩm).",
                        huong_dan_xu_ly="Ghi rõ sản phẩm đầu ra (văn bản, báo cáo, hồ sơ, bài giảng, số tiết...).",
                        du_lieu_goc=nv,
                    )
                )
                co_loi_nv = True

            # Quy tắc R2: Kiểm tra danh mục Nguồn giao, Tình trạng & Cảnh báo sai lệch Ngày hoàn thành
            if nv.nguon_giao not in DANH_MUC_NGUON_GIAO:
                vung_cach_ly.append(
                    BanGhiCachLy(
                        ma_quy_tac="R2_INVALID_SOURCE",
                        muc_do="ERROR",
                        doi_tuong=ten_ht,
                        nguyen_nhan=f"Nguồn giao '{nv.nguon_giao}' không nằm trong Sheet Danh mục chuẩn.",
                        huong_dan_xu_ly=f"Chọn một trong 5 nguồn giao chuẩn: {', '.join(DANH_MUC_NGUON_GIAO)}.",
                        du_lieu_goc=nv,
                    )
                )
                co_loi_nv = True

            if nv.tinh_trang not in DANH_MUC_TINH_TRANG:
                vung_cach_ly.append(
                    BanGhiCachLy(
                        ma_quy_tac="R2_INVALID_STATUS",
                        muc_do="ERROR",
                        doi_tuong=ten_ht,
                        nguyen_nhan=f"Tình trạng '{nv.tinh_trang}' không nằm trong Sheet Danh mục chuẩn.",
                        huong_dan_xu_ly=f"Chọn một trong 5 tình trạng chuẩn: {', '.join(DANH_MUC_TINH_TRANG)}.",
                        du_lieu_goc=nv,
                    )
                )
                co_loi_nv = True

            canh_bao_m = nv.canh_bao_excel()
            if canh_bao_m:
                vung_cach_ly.append(
                    BanGhiCachLy(
                        ma_quy_tac="R2_DATE_STATUS_MISMATCH",
                        muc_do="ERROR",
                        doi_tuong=ten_ht,
                        nguyen_nhan=f"Cảnh báo công thức cột M: '{canh_bao_m}' (Thời hạn={nv.thoi_han}, Ngày HT={nv.ngay_hoan_thanh}, Tình trạng='{nv.tinh_trang}').",
                        huong_dan_xu_ly="Chỉnh lại Tình trạng hoặc Ngày hoàn thành cho khớp logic thời gian thực tế.",
                        du_lieu_goc=nv,
                    )
                )
                co_loi_nv = True

            if (
                nv.ngay_giao is not None
                and nv.ngay_hoan_thanh is not None
                and nv.ngay_hoan_thanh < nv.ngay_giao
            ):
                vung_cach_ly.append(
                    BanGhiCachLy(
                        ma_quy_tac="R2_PHYSICAL_TIME_PARADOX",
                        muc_do="ERROR",
                        doi_tuong=ten_ht,
                        nguyen_nhan=f"Phi lý vật lý thời gian: Ngày hoàn thành ({nv.ngay_hoan_thanh}) sớm hơn Ngày giao ({nv.ngay_giao}).",
                        huong_dan_xu_ly="Kiểm tra lại Ngày giao và Ngày hoàn thành của nhiệm vụ.",
                        du_lieu_goc=nv,
                    )
                )
                co_loi_nv = True

            # Quy tắc R3: Kiểm tra ràng buộc vật lý Mức độ hoàn thành (%)
            if not (0.0 <= nv.muc_do_hoan_thanh <= 100.0):
                vung_cach_ly.append(
                    BanGhiCachLy(
                        ma_quy_tac="R3_INVALID_PERCENTAGE",
                        muc_do="ERROR",
                        doi_tuong=ten_ht,
                        nguyen_nhan=f"Mức độ hoàn thành ({nv.muc_do_hoan_thanh}%) nằm ngoài miền vật lý [0%, 100%].",
                        huong_dan_xu_ly="Nhập giá trị phần trăm từ 0 đến 100.",
                        du_lieu_goc=nv,
                    )
                )
                co_loi_nv = True
            elif nv.tinh_trang.startswith("Hoàn thành") and nv.muc_do_hoan_thanh < 100.0:
                vung_cach_ly.append(
                    BanGhiCachLy(
                        ma_quy_tac="R3_INCOMPLETE_PERCENT_ON_DONE",
                        muc_do="WARNING",
                        doi_tuong=ten_ht,
                        nguyen_nhan=f"Tình trạng là '{nv.tinh_trang}' nhưng mức độ hoàn thành chỉ ghi {nv.muc_do_hoan_thanh}% (< 100%).",
                        huong_dan_xu_ly="Cập nhật mức độ hoàn thành lên 100% nếu đã nghiệm thu xong sản phẩm.",
                        du_lieu_goc=nv,
                    )
                )

            # Quy tắc R6: Kiểm tra trùng lặp và quy định Usmart mục 1c
            if (
                nv.nguon_giao == "Giao việc trên Usmart"
                and not nv.so_van_ban_ma_viec.strip()
            ):
                vung_cach_ly.append(
                    BanGhiCachLy(
                        ma_quy_tac="R6_USMART_MISSING_CODE",
                        muc_do="ERROR",
                        doi_tuong=ten_ht,
                        nguyen_nhan="Nhiệm vụ chọn nguồn 'Giao việc trên Usmart' nhưng thiếu Số văn bản / Mã việc (TB 940 mục 1c).",
                        huong_dan_xu_ly="Tra cứu cột 'Số, ký hiệu' văn bản trên uSmart và điền vào cột Số VB / mã việc.",
                        du_lieu_goc=nv,
                    )
                )
                co_loi_nv = True

            khoa_ten = nv.nhiem_vu.strip().lower()
            if khoa_ten and khoa_ten in da_gap_ten_nv:
                vung_cach_ly.append(
                    BanGhiCachLy(
                        ma_quy_tac="R6_DUPLICATE_TASK",
                        muc_do="WARNING",
                        doi_tuong=ten_ht,
                        nguyen_nhan="Phát hiện nhiệm vụ trùng tên trong cùng kỳ báo cáo của cá nhân.",
                        huong_dan_xu_ly="Kiểm tra và gộp các nhiệm vụ trùng lặp giữa Kế hoạch, Mẫu 03 và uSmart.",
                        du_lieu_goc=nv,
                    )
                )
            if khoa_ten:
                da_gap_ten_nv.add(khoa_ten)

            if co_loi_nv:
                nv_cach_ly.append(nv)
            else:
                nv_hop_le.append(nv)

        # Quy tắc R4: Kiểm tra ràng buộc vật lý Hoạt động chuyên môn của giảng viên (Mục II)
        cm = bao_cao.hoat_dong_chuyen_mon
        for ten_muc, thang_val, luy_ke_val, dinh_muc_val, nguong_max in [
            (
                "Giảng dạy (giờ chuẩn)",
                cm.giang_day_thang,
                cm.giang_day_luy_ke,
                cm.giang_day_dinh_muc,
                NGUONG_GIO_GIANG_DAY_THANG_MAX,
            ),
            (
                "Nghiên cứu khoa học (giờ chuẩn)",
                cm.nckh_thang,
                cm.nckh_luy_ke,
                cm.nckh_dinh_muc,
                NGUONG_GIO_NCKH_THANG_MAX,
            ),
            (
                "Nhiệm vụ khác (giờ)",
                cm.nhiem_vu_khac_thang,
                cm.nhiem_vu_khac_luy_ke,
                cm.nhiem_vu_khac_dinh_muc,
                NGUONG_GIO_KHAC_THANG_MAX,
            ),
        ]:
            if thang_val < 0 or luy_ke_val < 0 or dinh_muc_val <= 0:
                vung_cach_ly.append(
                    BanGhiCachLy(
                        ma_quy_tac="R4_NEGATIVE_OR_ZERO_QUOTA",
                        muc_do="ERROR",
                        doi_tuong=f"Mục II - {ten_muc}",
                        nguyen_nhan=f"Giá trị giờ âm hoặc định mức <= 0 (Tháng={thang_val}, Lũy kế={luy_ke_val}, Định mức={dinh_muc_val}).",
                        huong_dan_xu_ly="Nhập số giờ >= 0 và định mức năm học > 0.",
                    )
                )
            if luy_ke_val < thang_val:
                vung_cach_ly.append(
                    BanGhiCachLy(
                        ma_quy_tac="R4_CUMULATIVE_LESS_THAN_MONTHLY",
                        muc_do="ERROR",
                        doi_tuong=f"Mục II - {ten_muc}",
                        nguyen_nhan=f"Phi lý vật lý: Giờ lũy kế năm học ({luy_ke_val}) nhỏ hơn giờ thực hiện trong tháng ({thang_val}).",
                        huong_dan_xu_ly="Cập nhật lại Giờ lũy kế năm học >= Giờ thực hiện trong tháng.",
                    )
                )
            if thang_val > nguong_max:
                vung_cach_ly.append(
                    BanGhiCachLy(
                        ma_quy_tac="R4_OUTLIER_HOURS",
                        muc_do="WARNING",
                        doi_tuong=f"Mục II - {ten_muc}",
                        nguyen_nhan=f"Điểm dị biệt (Outlier): Giờ thực hiện trong 1 tháng ({thang_val} giờ) vượt ngưỡng cảnh báo ({nguong_max} giờ).",
                        huong_dan_xu_ly="Kiểm tra xem có nhầm lẫn giữa giờ chuẩn trong tháng và giờ lũy kế cả năm không.",
                    )
                )

        if not cm.san_pham_khcn_huong_dan.strip():
            vung_cach_ly.append(
                BanGhiCachLy(
                    ma_quy_tac="R4_MISSING_SCIENTIFIC_OUTPUT",
                    muc_do="WARNING",
                    doi_tuong="Mục II - Sản phẩm KHCN & hướng dẫn người học",
                    nguyen_nhan="Chưa kê khai chi tiết sản phẩm KHCN, bài báo, đề tài hoặc hướng dẫn đồ án/luận văn trong tháng.",
                    huong_dan_xu_ly="Bổ sung thông tin hướng dẫn sinh viên/học viên hoặc bài báo/đề tài đang thực hiện.",
                )
            )

        # Quy tắc R5: Kiểm tra điểm tự chấm (Mục III)
        if not (0.0 <= bao_cao.diem_tu_cham_tieu_chi_chung <= 30.0):
            vung_cach_ly.append(
                BanGhiCachLy(
                    ma_quy_tac="R5_SCORE_GENERAL_OUT_OF_RANGE",
                    muc_do="ERROR",
                    doi_tuong="Mục III.1 - Điểm tự chấm Tiêu chí chung",
                    nguyen_nhan=f"Điểm tự chấm tiêu chí chung ({bao_cao.diem_tu_cham_tieu_chi_chung}) vượt khung [0, 30].",
                    huong_dan_xu_ly="Chỉnh lại điểm tự chấm tiêu chí chung trong khoảng từ 0 đến tối đa 30 điểm.",
                )
            )

        diem_nv = bao_cao.diem_tu_cham_nv_hieu_luc
        if not (0.0 <= diem_nv <= 70.0):
            vung_cach_ly.append(
                BanGhiCachLy(
                    ma_quy_tac="R5_SCORE_TASK_OUT_OF_RANGE",
                    muc_do="ERROR",
                    doi_tuong="Mục III.2 - Điểm tự chấm Kết quả thực hiện nhiệm vụ",
                    nguyen_nhan=f"Điểm tự chấm kết quả nhiệm vụ ({diem_nv}) vượt khung [0, 70].",
                    huong_dan_xu_ly="Chỉnh lại điểm tự chấm kết quả nhiệm vụ trong khoảng từ 0 đến tối đa 70 điểm.",
                )
            )

        tham_chieu = bao_cao.diem_tham_chieu_nhiem_vu
        if tham_chieu is not None and abs(diem_nv - tham_chieu) > 10.0:
            vung_cach_ly.append(
                BanGhiCachLy(
                    ma_quy_tac="R5_SCORE_DEVIATION_FROM_REF",
                    muc_do="WARNING",
                    doi_tuong="Mục III.2 - So sánh Điểm tự chấm và Điểm tham chiếu",
                    nguyen_nhan=f"Điểm tự chấm nhiệm vụ ({diem_nv}) chênh lệch > 10 điểm so với Điểm tham chiếu theo tỷ lệ đúng hạn ({tham_chieu}).",
                    huong_dan_xu_ly="Rà soát lại điểm tự chấm hoặc bổ sung giải trình về độ phức tạp/chất lượng công việc.",
                )
            )

        so_loi = sum(1 for item in vung_cach_ly if item.muc_do == "ERROR")
        so_cb = sum(1 for item in vung_cach_ly if item.muc_do == "WARNING")

        return KetQuaKiemToanBaoCao(
            hop_le_de_chot=(so_loi == 0),
            tong_so_loi_nghiem_trong=so_loi,
            tong_so_canh_bao_di_biet=so_cb,
            vung_cach_ly=vung_cach_ly,
            nhiem_vu_hop_le=nv_hop_le,
            nhiem_vu_bi_cach_ly=nv_cach_ly,
        )

    def kiem_tra_thong_ke_usmart_don_vi(
        self, tk: ThongKeUsmartDonVi
    ) -> List[BanGhiCachLy]:
        """Kiểm tra ràng buộc Mục II Mẫu 02a/02b: dòng 'Kiểm tra' phải bằng 0."""
        loi: List[BanGhiCachLy] = []
        if tk.kiem_tra_sai_lech != 0:
            loi.append(
                BanGhiCachLy(
                    ma_quy_tac="R7_USMART_SUM_MISMATCH",
                    muc_do="ERROR",
                    doi_tuong="Mẫu 02 - Mục II (Xử lý công việc trên uSmart)",
                    nguyen_nhan=(
                        f"Dòng Kiểm tra khác 0 ({tk.kiem_tra_sai_lech}): Tổng đến hạn ({tk.so_van_ban_den_han}) "
                        f"!= Đúng hạn ({tk.hoan_thanh_dung_han}) + Chậm ({tk.hoan_thanh_cham}) + Quá hạn ({tk.qua_han_chua_hoan_thanh})."
                    ),
                    huong_dan_xu_ly="Điều chỉnh số liệu thống kê uSmart sao cho tổng 3 thành phần bằng đúng số văn bản đến hạn.",
                )
            )
        return loi
