# -*- coding: utf-8 -*-
"""
Module Điều phối Tác tử Trung tâm (Orchestrator Agent & CLI)
cho Hệ thống Báo cáo Công việc Bộ môn Địa kỹ thuật và Không gian ngầm (TB 940/TB-ĐHGTVT).

Cách sử dụng:
    # Chạy toàn bộ quy trình (Kiểm toán + Xuất Mẫu 01, 02, 03, DOCX + Lưu nháp uSmart cho TK hanh1):
    python -m he_thong_bao_cao_dkt.cli_baocao --thang 11 --nam 2026 --ma-cb 1051 --usmart-user hanh1 --usmart-pass 182128 --che-do-usmart luu_nhap

    # Xuất đồng loạt 21 báo cáo cá nhân (.xlsx & .docx) cho cả Bộ môn:
    python -m he_thong_bao_cao_dkt.cli_baocao --thang 11 --nam 2026 --xuat-toan-bo-bo-mon
"""

import argparse
from dataclasses import dataclass, field
import logging
from pathlib import Path
import sys
from typing import List, Optional

from .database_manager import QuanLyCSDLBộMôn
from .docx_generator import BoSinhBaoCaoDocx
from .excel_generator import BoSinhBaoCaoExcel, PROJECT_ROOT
from .usmart_baocao_client import KetQuaGuiUSmart, USmartBaoCaoAgent
from .usmart_nckh_sync import KetQuaDongBoNCKH, chay_dong_bo_ke_hoach_nckh_toan_trinh
from .validator import BoKiemTraRangBuocVatLy, KetQuaKiemToanBaoCao

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(name)s: %(message)s",
)
logger = logging.getLogger(__name__)


@dataclass
class KetQuaChayPipeline:
    """Kết quả tổng hợp sau khi chạy Pipeline Tác tử Báo cáo Bộ môn ĐKT & KGN."""

    duong_dan_csdl: Path
    ket_qua_kiem_toan_ca_nhan: KetQuaKiemToanBaoCao
    file_xlsx_ca_nhan: Path
    file_docx_ca_nhan: Path
    file_xlsx_mau_02_don_vi: Path
    file_docx_mau_02_don_vi: Path
    file_xlsx_mau_03_so_giao_viec: Path
    danh_sach_file_toan_bo_gv: List[Path] = field(default_factory=list)
    ket_qua_usmart: Optional[KetQuaGuiUSmart] = None
    ket_qua_dong_bo_nckh: Optional[KetQuaDongBoNCKH] = None


class TacTuDieuPhoiBaoCao:
    """
    Tác tử Điều phối (Orchestrator Agent) kết nối toàn bộ 4 thành phần:
    CSDL Google Sheets/Excel -> Kiểm toán Vật lý -> Bộ sinh XLSX/DOCX -> Tác tử OIDC uSmart.
    """

    def __init__(
        self,
        duong_dan_csdl: Optional[Path] = None,
        thu_muc_xuat: Optional[Path] = None,
    ) -> None:
        self.duong_dan_csdl: Path = duong_dan_csdl or (
            PROJECT_ROOT / "CSDL_BaoCao_BoMon_DKT_KGN.xlsx"
        )
        self.thu_muc_xuat: Path = thu_muc_xuat or (
            PROJECT_ROOT / "output_bao_cao_dkt"
        )
        self.db = QuanLyCSDLBộMôn(self.duong_dan_csdl)
        self.validator = BoKiemTraRangBuocVatLy()
        self.excel_gen = BoSinhBaoCaoExcel()
        self.docx_gen = BoSinhBaoCaoDocx()

    def chay_pipeline_hoan_chinh(
        self,
        ma_can_bo: str = "1051",
        thang: int = 11,
        nam: int = 2026,
        xuat_toan_bo_21_gv: bool = False,
        usmart_username: Optional[str] = None,
        usmart_password: Optional[str] = None,
        che_do_usmart: str = "luu_nhap",  # "khong", "luu_nhap", "nop", "upload_xlsx"
        dong_bo_nckh: bool = False,
        nckh_username: str = "lannc",
        nckh_password: str = "giangchau",
        dong_bo_nckh_live: bool = True,
    ) -> KetQuaChayPipeline:
        """Thực thi toàn trình từ khởi tạo/đọc CSDL đến kiểm toán, xuất file và gửi uSmart."""
        # 1. Đảm bảo CSDL trung tâm đã sẵn sàng
        self.db.khoi_tao_csdl_mau_neu_chua_co(ghi_de=False)

        # 1b. Đồng bộ Kế hoạch NCKH đề tài các cấp & NCKHSV từ uSmart (TK lannc) nếu được yêu cầu
        kq_nckh: Optional[KetQuaDongBoNCKH] = None
        if dong_bo_nckh:
            kq_nckh = chay_dong_bo_ke_hoach_nckh_toan_trinh(
                username=nckh_username,
                password=nckh_password,
                thang=thang,
                nam=nam,
                dong_bo_live=dong_bo_nckh_live,
                duong_dan_csdl_master=self.duong_dan_csdl,
            )

        # 2. Đọc và kiểm toán báo cáo cá nhân của giảng viên mục tiêu
        bc_ca_nhan = self.db.doc_bao_cao_ca_nhan(ma_can_bo, thang=thang, nam=nam)
        kq_kt = self.validator.kiem_tra_bao_cao_ca_nhan(bc_ca_nhan)
        logger.info(
            "Kết quả kiểm toán cho %s (%s): %s",
            bc_ca_nhan.giang_vien.ho_va_ten,
            ma_can_bo,
            kq_kt.tom_tat_kiem_toan(),
        )
        if not kq_kt.hop_le_de_chot:
            for item in kq_kt.vung_cach_ly:
                logger.warning(
                    "  [CÁCH LY - %s] %s: %s -> %s",
                    item.muc_do,
                    item.ma_quy_tac,
                    item.nguyen_nhan,
                    item.huong_dan_xu_ly,
                )

        # 3. Xuất báo cáo cá nhân (.xlsx chuẩn Mẫu 01 và bản cứng .docx)
        thu_muc_ky = self.thu_muc_xuat / f"Ky_{thang:02d}_{nam}"
        file_01_xlsx = self.excel_gen.xuat_mau_01_ca_nhan(bc_ca_nhan, thu_muc_ky)
        file_01_docx = self.docx_gen.xuat_docx_ca_nhan_01b(bc_ca_nhan, thu_muc_ky)

        # 4. Đọc toàn bộ 21 giảng viên và tổng hợp Mẫu 02 (02b, 02c) + Mẫu 03 (03a, 03b, 03c)
        ds_toan_bo = self.db.doc_toan_bo_bao_cao_bo_mon(thang=thang, nam=nam)
        file_02_xlsx = self.excel_gen.xuat_mau_02_don_vi(
            ds_toan_bo, thu_muc_ky, thang=thang, nam=nam
        )
        file_02_docx = self.docx_gen.xuat_docx_tong_hop_don_vi(
            ds_toan_bo, thu_muc_ky, thang=thang, nam=nam
        )
        file_03_xlsx = self.excel_gen.xuat_mau_03_so_giao_viec(
            ds_toan_bo, thu_muc_ky, thang=thang, nam=nam
        )

        # 5. Nếu yêu cầu xuất toàn bộ 21 giảng viên
        ds_file_21_gv: List[Path] = [file_01_xlsx, file_01_docx]
        if xuat_toan_bo_21_gv:
            thu_muc_21 = thu_muc_ky / "BaoCao_CaNhan_21_GV"
            for bc in ds_toan_bo:
                f_x = self.excel_gen.xuat_mau_01_ca_nhan(bc, thu_muc_21)
                f_d = self.docx_gen.xuat_docx_ca_nhan_01b(bc, thu_muc_21)
                ds_file_21_gv.extend([f_x, f_d])

        # 6. Kết nối và gửi trực tiếp lên uSmart nếu cung cấp tài khoản/mật khẩu
        kq_usmart: Optional[KetQuaGuiUSmart] = None
        if usmart_username and usmart_password and che_do_usmart != "khong":
            agent_us = USmartBaoCaoAgent(usmart_username, usmart_password)
            if che_do_usmart in ("luu_nhap", "nop"):
                # Gửi cả Mẫu 01a (Tuần 45, ky=1) và Mẫu 01b (Tháng 11, ky=2) ở chế độ lưu nháp nếu là luu_nhap
                if che_do_usmart == "luu_nhap":
                    agent_us.gui_bao_cao_webform(
                        bc_ca_nhan, ky_id="1", ma_mau="01a", che_do="luu_nhap"
                    )
                kq_usmart = agent_us.gui_bao_cao_webform(
                    bc_ca_nhan, ky_id="2", ma_mau="01b", che_do=che_do_usmart
                )
            elif che_do_usmart == "upload_xlsx":
                kq_usmart = agent_us.tai_tep_excel_len_usmart(
                    file_01_xlsx, ky_id="2", ma_mau="01b"
                )

        return KetQuaChayPipeline(
            duong_dan_csdl=self.duong_dan_csdl,
            ket_qua_kiem_toan_ca_nhan=kq_kt,
            file_xlsx_ca_nhan=file_01_xlsx,
            file_docx_ca_nhan=file_01_docx,
            file_xlsx_mau_02_don_vi=file_02_xlsx,
            file_docx_mau_02_don_vi=file_02_docx,
            file_xlsx_mau_03_so_giao_viec=file_03_xlsx,
            danh_sach_file_toan_bo_gv=ds_file_21_gv,
            ket_qua_usmart=kq_usmart,
            ket_qua_dong_bo_nckh=kq_nckh,
        )


def main() -> None:
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8")

    parser = argparse.ArgumentParser(
        description="Hệ thống Tác tử Tự động Sinh Báo cáo Công việc & Nộp uSmart (Bộ môn ĐKT & KGN)"
    )
    parser.add_argument("--ma-cb", default="1051", help="Mã cán bộ giảng viên (mặc định: 1051 - Nguyễn Hải Hà)")
    parser.add_argument("--thang", type=int, default=11, help="Tháng báo cáo (1..12)")
    parser.add_argument("--nam", type=int, default=2026, help="Năm báo cáo")
    parser.add_argument("--xuat-toan-bo-bo-mon", action="store_true", help="Xuất báo cáo cho toàn bộ 21 giảng viên Bộ môn")
    parser.add_argument("--dong-bo-nckh", action="store_true", help="Đồng bộ Kế hoạch NCKH & NCKHSV từ uSmart (TK lannc)")
    parser.add_argument("--nckh-user", default="lannc", help="Tài khoản uSmart KHCN (mặc định: lannc)")
    parser.add_argument("--nckh-pass", default="giangchau", help="Mật khẩu uSmart KHCN (mặc định: giangchau -> giangchau0310)")
    parser.add_argument("--usmart-user", default="", help="Tài khoản đăng nhập uSmart Báo cáo (VD: hanh1)")
    parser.add_argument("--usmart-pass", default="", help="Mật khẩu đăng nhập uSmart Báo cáo")
    parser.add_argument(
        "--che-do-usmart",
        choices=["khong", "luu_nhap", "nop", "upload_xlsx"],
        default="luu_nhap",
        help="Chế độ gửi lên uSmart (mặc định: luu_nhap an toàn)",
    )
    args = parser.parse_args()

    orchestrator = TacTuDieuPhoiBaoCao()
    kq = orchestrator.chay_pipeline_hoan_chinh(
        ma_can_bo=args.ma_cb,
        thang=args.thang,
        nam=args.nam,
        xuat_toan_bo_21_gv=args.xuat_toan_bo_bo_mon,
        usmart_username=args.usmart_user or None,
        usmart_password=args.usmart_pass or None,
        che_do_usmart=args.che_do_usmart,
        dong_bo_nckh=args.dong_bo_nckh,
        nckh_username=args.nckh_user,
        nckh_password=args.nckh_pass,
    )

    print("\n================ KẾT QUẢ THỰC THI PIPELINE TÁC TỬ ================")
    print(f"1. CSDL Trung tâm Bộ môn : {kq.duong_dan_csdl}")
    if kq.ket_qua_dong_bo_nckh:
        nckh = kq.ket_qua_dong_bo_nckh
        print(
            f"1b. Kế hoạch NCKH/NCKHSV : {nckh.duong_dan_gsheet_project} "
            f"({nckh.so_de_tai_cac_cap_2025_2027} ĐT các cấp, {nckh.so_de_tai_nckhsv_2025_2027} NCKHSV)"
        )
    print(f"2. Kiểm toán Số liệu     : {kq.ket_qua_kiem_toan_ca_nhan.tom_tat_kiem_toan()}")
    print(f"3. Mẫu 01b Cá nhân (.xlsx): {kq.file_xlsx_ca_nhan}")
    print(f"4. Bản cứng Cá nhân (.docx): {kq.file_docx_ca_nhan}")
    print(f"5. Mẫu 02 Đơn vị (.xlsx)  : {kq.file_xlsx_mau_02_don_vi}")
    print(f"6. Bản cứng Đơn vị (.docx): {kq.file_docx_mau_02_don_vi}")
    print(f"7. Sổ giao việc Mẫu 03    : {kq.file_xlsx_mau_03_so_giao_viec}")
    if kq.ket_qua_usmart:
        print(f"8. Trạng thái gửi uSmart  : {kq.ket_qua_usmart.thong_bao}")
    print("==================================================================")


if __name__ == "__main__":
    main()
