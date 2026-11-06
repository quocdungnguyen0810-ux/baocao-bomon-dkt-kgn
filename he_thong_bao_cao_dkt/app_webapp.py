# -*- coding: utf-8 -*-
"""
Giao diện Webapp Tương tác Chuyên nghiệp (Streamlit) cho Giảng viên và Trưởng Bộ môn
Địa kỹ thuật và Không gian ngầm - Trường Đại học Giao thông vận tải (UTC).

Hỗ trợ đầy đủ các kiểu lựa chọn báo cáo theo Thông báo số 940/TB-ĐHGTVT:
1. Báo cáo công việc cá nhân (Mẫu 01):
   - Mẫu 01a: Báo cáo kết quả công việc tuần (BC tuần)
   - Mẫu 01b: Báo cáo kết quả công việc tháng và tự chấm điểm (BC tháng)
2. Báo cáo công việc đơn vị (Mẫu 02):
   - Mẫu 02a: Báo cáo kết quả công việc tuần của đơn vị (BC tuần ĐV)
   - Mẫu 02b: Báo cáo kết quả công việc tháng của đơn vị (BC tháng ĐV)
   - Mẫu 02c: Bảng theo dõi kết quả công việc và chấm điểm tháng 21 giảng viên
3. Giao việc và theo dõi nhiệm vụ (Mẫu 03):
   - Mẫu 03a: Phiếu giao việc & Nghiệm thu kết quả
   - Mẫu 03b: Sổ giao việc và theo dõi thực hiện nhiệm vụ (đủ 20 cột)
   - Mẫu 03c: Tổng hợp kết quả thực hiện nhiệm vụ được giao theo người
4. Quản trị Đồng bộ & Hiển thị Dữ liệu Nguồn (Chỉ hiển thị khi đăng nhập tài khoản Admin).

Chạy ứng dụng:
    streamlit run he_thong_bao_cao_dkt/app_webapp.py
"""

from datetime import date, timedelta
from pathlib import Path
import streamlit as st

from he_thong_bao_cao_dkt.database_manager import QuanLyCSDLBộMôn, _parse_date
from he_thong_bao_cao_dkt.docx_generator import BoSinhBaoCaoDocx
from he_thong_bao_cao_dkt.excel_generator import BoSinhBaoCaoExcel, PROJECT_ROOT
from he_thong_bao_cao_dkt.models import (
    DANH_MUC_CHAT_LUONG,
    DANH_MUC_NGUON_GIAO,
    DANH_MUC_TINH_TRANG,
    DANH_MUC_UU_TIEN,
    DANH_MUC_XEP_LOAI,
    HoatDongChuyenMon,
    KeHoachKyToi,
    NhiemVuCongViec,
    TEN_DON_VI_CHUAN,
    lay_mat_khau_usmart_mac_dinh,
    tim_ma_can_bo_theo_tai_khoan_usmart,
)
from he_thong_bao_cao_dkt.tkb_giang_day_sync import (
    chay_dong_bo_tkb_giang_day_toan_trinh,
)
from he_thong_bao_cao_dkt.ui_donvi_giaoviec import (
    _tinh_thu_hai_mac_dinh,
    hien_thi_bao_cao_don_vi_mau_02,
    hien_thi_giao_viec_va_theo_doi_mau_03,
    hien_thi_nut_tai_file_truc_tiep,
)
from he_thong_bao_cao_dkt.usmart_baocao_client import USmartBaoCaoAgent
from he_thong_bao_cao_dkt.usmart_nckh_sync import (
    BoXuLyKeHoachNCKH,
    FOLDER_KE_HOACH_CONG_VIEC,
    TEN_FILE_GSHEET_NCKH,
    USmartNCKHClient,
    chay_dong_bo_ke_hoach_nckh_toan_trinh,
)
from he_thong_bao_cao_dkt.validator import BoKiemTraRangBuocVatLy


# Danh sách tài khoản Quản trị viên (Admin) được phép xem & đồng bộ dữ liệu nguồn
TAI_KHOAN_ADMIN_HOP_LE = {
    "admin": "dkt412@admin",
    "lannc": "giangchau0310",
    "truyentt": "dkt412@admin",
}


def xac_thuc_tai_khoan_admin(username: str, password: str) -> bool:
    """Kiểm tra thông tin đăng nhập có thuộc quyền Quản trị viên (Admin) hay không."""
    u = (username or "").strip().lower()
    p = (password or "").strip()
    if not u or not p:
        return False
    if u == "lannc" and p in ("giangchau0310", "giangchau"):
        return True
    return TAI_KHOAN_ADMIN_HOP_LE.get(u) == p


def lay_danh_sach_kieu_bao_cao_theo_quyen(is_admin: bool) -> list[str]:
    """
    Trả về danh sách các kiểu lựa chọn báo cáo theo quyền người dùng:
    - Tài khoản thường (`is_admin=False`): Chỉ hiển thị Báo cáo công việc cá nhân (Mẫu 01) và Giao việc & theo dõi nhiệm vụ (Mẫu 03).
    - Tài khoản Admin (`is_admin=True`): Hiển thị thêm Báo cáo công việc đơn vị (Mẫu 02) và Đồng bộ & Hiển thị dữ liệu nguồn.
    """
    ds = [
        "📄 1. Báo cáo công việc cá nhân (Mẫu 01 — Theo tuần & Theo tháng)",
        "📋 2. Giao việc và theo dõi nhiệm vụ (Mẫu 03 — Phiếu giao việc, Sổ giao việc & Tổng hợp)",
    ]
    if is_admin:
        ds.extend(
            [
                "🏛️ 3. Báo cáo công việc đơn vị (Mẫu 02 — Theo tuần, Theo tháng & Chấm điểm 21 GV — Chỉ Admin)",
                "🔄 4. Đồng bộ & Hiển thị Dữ liệu Nguồn (TKB & NCKH uSmart — Chỉ Admin)",
            ]
        )
    return ds


def xac_dinh_thu_muc_luu_file(kieu_vi_tri: str, duong_dan_tuy_chinh: str = "") -> Path:
    """Giải mã lựa chọn vị trí lưu file của người dùng thành đường dẫn `Path` hợp lệ và tự tạo thư mục nếu chưa có."""
    if kieu_vi_tri.startswith("☁️"):
        target = FOLDER_KE_HOACH_CONG_VIEC / "Bao_cao_xuat"
    elif kieu_vi_tri.startswith("💻"):
        dl_dir = Path.home() / "Downloads" / "Bao_cao_Bo_mon_DKT_KGN"
        target = dl_dir
    elif kieu_vi_tri.startswith("🖥️"):
        dt_dir = Path.home() / "Desktop" / "Bao_cao_Bo_mon_DKT_KGN"
        target = dt_dir
    elif kieu_vi_tri.startswith("✍️") and duong_dan_tuy_chinh.strip():
        target = Path(duong_dan_tuy_chinh.strip())
    else:
        target = PROJECT_ROOT / "output_bao_cao_dkt"
    try:
        target.mkdir(parents=True, exist_ok=True)
    except OSError:
        target = PROJECT_ROOT / "output_bao_cao_dkt"
        target.mkdir(parents=True, exist_ok=True)
    return target


def thiet_lap_css_chuyen_nghiep() -> None:
    """Thiết lập giao diện CSS chuyên nghiệp chuẩn hệ thống hành chính - học thuật Trường ĐH GTVT."""
    st.markdown(
        """
        <style>
        .utc-top-banner {
            background: linear-gradient(135deg, #0f2d59 0%, #1d4ed8 100%);
            color: #ffffff;
            padding: 16px 22px;
            border-radius: 10px;
            margin-bottom: 14px;
            box-shadow: 0 4px 12px rgba(15, 45, 89, 0.16);
        }
        .utc-top-banner h2 {
            color: #ffffff !important;
            margin: 0 0 4px 0 !important;
            font-size: 1.35rem !important;
            font-weight: 700 !important;
        }
        .utc-top-banner p {
            color: #e2e8f0 !important;
            margin: 0 !important;
            font-size: 0.92rem !important;
        }
        .section-banner {
            background: #f0f6ff;
            border-left: 5px solid #1d4ed8;
            padding: 12px 16px;
            border-radius: 6px;
            margin-bottom: 12px;
        }
        .section-banner h3 {
            color: #0f2d59 !important;
            margin: 0 0 4px 0 !important;
            font-size: 1.1rem !important;
        }
        .section-banner p {
            color: #334155 !important;
            margin: 0 !important;
            font-size: 0.88rem !important;
        }
        .form-badge {
            display: inline-block;
            background: #1d4ed8;
            color: #ffffff;
            font-weight: 700;
            padding: 4px 12px;
            border-radius: 999px;
            font-size: 0.85rem;
        }
        .utc-footer {
            margin-top: 36px;
            padding: 14px 20px;
            background: #f8fafc;
            border-top: 2px solid #cbd5e1;
            border-radius: 8px;
            text-align: center;
            color: #1e293b;
            font-size: 0.92rem;
            font-weight: 500;
        }
        .utc-footer a {
            color: #1d4ed8;
            text-decoration: none;
            font-weight: 600;
        }
        .utc-footer a:hover {
            text-decoration: underline;
        }
        .roadmap-grid {
            display: grid;
            grid-template-columns: repeat(5, 1fr);
            gap: 10px;
            margin: 10px 0 14px 0;
        }
        @media (max-width: 992px) {
            .roadmap-grid {
                grid-template-columns: 1fr;
            }
        }
        .roadmap-card {
            background: #ffffff;
            border: 1.5px solid #cbd5e1;
            border-top: 4px solid #1d4ed8;
            border-radius: 8px;
            padding: 12px 12px;
            box-shadow: 0 2px 6px rgba(15, 45, 89, 0.06);
        }
        .roadmap-step-badge {
            display: inline-block;
            background: #eff6ff;
            color: #1d4ed8;
            border: 1px solid #bfdbfe;
            font-weight: 700;
            font-size: 0.78rem;
            padding: 2px 8px;
            border-radius: 999px;
            margin-bottom: 6px;
        }
        .roadmap-card h4 {
            color: #0f2d59 !important;
            font-size: 0.95rem !important;
            margin: 0 0 6px 0 !important;
            font-weight: 700 !important;
        }
        .roadmap-card p {
            color: #334155 !important;
            font-size: 0.83rem !important;
            line-height: 1.42 !important;
            margin: 0 !important;
        }
        </style>
        """,
        unsafe_allow_html=True,
    )


def hien_thi_so_do_roadmap_huong_dan(is_admin: bool = False) -> None:
    """Hiển thị Sơ đồ Lộ trình (Roadmap) hướng dẫn sử dụng trực quan từng bước cho giảng viên Bộ môn ĐKT & KGN."""
    with st.expander(
        "🗺️ SƠ ĐỒ LỘ TRÌNH HƯỚNG DẪN SỬ DỤNG NHANH CHO GIẢNG VIÊN BỘ MÔN ĐỊA KỸ THUẬT & KGN (5 BƯỚC — 2 PHÚT)",
        expanded=False,
    ):
        st.markdown(
            """
            <div class='roadmap-grid'>
                <div class='roadmap-card'>
                    <span class='roadmap-step-badge'>BƯỚC 1 • THANH BÊN TRÁI</span>
                    <h4>👤 Chọn Giảng viên & Kỳ</h4>
                    <p>
                        1. Chọn <b>Tháng / Năm</b> báo cáo.<br>
                        2. Chọn đúng <b>Mã CB - Họ và tên</b> trong danh sách 21 GV.<br>
                        3. Chọn <b>Vị trí lưu file</b> mong muốn.<br>
                        4. <b>Nhập tài khoản uSmart (tùy chọn)</b> để link trực tiếp đến module Báo cáo công việc cá nhân trên <a href="https://usmart.utc.edu.vn/baocaocongviec/" target="_blank"><b>uSmart</b></a>.
                    </p>
                </div>
                <div class='roadmap-card'>
                    <span class='roadmap-step-badge'>BƯỚC 2 • LOẠI BÁO CÁO</span>
                    <h4>🗂️ Chọn Mẫu Tuần / Tháng</h4>
                    <p>
                        • Chọn <b>📄 1. Báo cáo công việc cá nhân</b>:<br>
                        &nbsp;&nbsp;↳ <b>📅 Mẫu 01a</b>: Báo cáo Tuần.<br>
                        &nbsp;&nbsp;↳ <b>📆 Mẫu 01b</b>: Báo cáo Tháng.<br>
                        • Hoặc <b>📋 2. Giao việc (Mẫu 03)</b>.
                    </p>
                </div>
                <div class='roadmap-card'>
                    <span class='roadmap-step-badge'>BƯỚC 3 • MỤC I.2 NHIỆM VỤ</span>
                    <h4>☑️ Tích chọn & Sửa việc</h4>
                    <p>
                        • Tích <b>☑️ Sử dụng</b> việc trong kỳ hoặc bật lấy việc cũ để dùng lại.<br>
                        • Sửa trực tiếp <b>Nguồn giao, Thời hạn, Ngày HT, % HT</b>.<br>
                        • Tích <b>⭐ Trọng tâm ĐV</b> nếu là việc chính của Bộ môn.
                    </p>
                </div>
                <div class='roadmap-card'>
                    <span class='roadmap-step-badge'>BƯỚC 4 • MỤC II, III, IV & LƯU</span>
                    <h4>💾 Tự chấm điểm & Lưu</h4>
                    <p>
                        • Kiểm tra giờ <b>Giảng dạy, NCKH</b> (Mục II).<br>
                        • Nhập <b>Điểm tự chấm</b> (Mục III: TC chung ≤30, KQ NV ≤70).<br>
                        • Bấm nút <b>💾 Lưu toàn bộ thay đổi vào CSDL</b>.
                    </p>
                </div>
                <div class='roadmap-card'>
                    <span class='roadmap-step-badge'>BƯỚC 5 • XUẤT FILE & NỘP USMART</span>
                    <h4>📥 Tải File / Nộp uSmart</h4>
                    <p>
                        • <b>Xuất file:</b> Bấm <b>📊 Xuất Excel</b> / <b>📝 Xuất Word</b> ➔ <b>⬇️ Tải về máy</b>.<br>
                        • <b>Nộp lên uSmart (Nhập dữ liệu tự động sang uSmart):</b><br>
                        &nbsp;&nbsp;↳ Mở <a href="https://usmart.utc.edu.vn/baocaocongviec/" target="_blank"><b>🔗 Cổng Báo cáo uSmart</b></a> ➔ <b>Giao diện Báo cáo cá nhân</b> ➔ Chọn <b>Nhập Web</b> ➔ <b>Kiểm tra và Nộp báo cáo</b>.
                    </p>
                </div>
            </div>
            """,
            unsafe_allow_html=True,
        )

        st.markdown("##### 🔄 Sơ đồ luồng thực hiện & kế thừa dữ liệu tự động trong Bộ môn")
        st.markdown(
            """
            <div style="background:#f8fafc; border:1px solid #cbd5e1; border-radius:10px; padding:14px 16px; margin-bottom:12px;">
                <div style="display:flex; flex-wrap:wrap; align-items:center; justify-content:space-between; gap:8px; text-align:center;">
                    <div style="flex:1; min-width:155px; background:#dbeafe; border:1.5px solid #2563eb; border-radius:8px; padding:10px;">
                        <b style="color:#1e3a8a;">1️⃣ KHỞI ĐẦU</b><br>
                        <span style="font-size:0.81rem; color:#1e293b;">Chọn GV, Kỳ, Vị trí lưu<br>+ Nhập TK uSmart (tùy chọn)</span>
                    </div>
                    <div style="font-size:1.3rem; font-weight:bold; color:#2563eb;">➔</div>
                    <div style="flex:1; min-width:145px; background:#e0f2fe; border:1.5px solid #0284c7; border-radius:8px; padding:10px;">
                        <b style="color:#0c4a6e;">2️⃣ CHỌN BIỂU MẪU</b><br>
                        <span style="font-size:0.81rem; color:#1e293b;">Mẫu 01a (BC Tuần)<br>Mẫu 01b (BC Tháng)</span>
                    </div>
                    <div style="font-size:1.3rem; font-weight:bold; color:#0284c7;">➔</div>
                    <div style="flex:1; min-width:165px; background:#fef3c7; border:1.5px solid #d97706; border-radius:8px; padding:10px;">
                        <b style="color:#78350f;">3️⃣ RÀ SOÁT & CHẤM ĐIỂM</b><br>
                        <span style="font-size:0.81rem; color:#1e293b;">Tích ☑️ Sử dụng / 🔁 Tiếp tục<br>Sửa Ngày HT, Điểm, ⭐ Trọng tâm</span>
                    </div>
                    <div style="font-size:1.3rem; font-weight:bold; color:#d97706;">➔</div>
                    <div style="flex:1; min-width:145px; background:#dcfce7; border:1.5px solid #16a34a; border-radius:8px; padding:10px;">
                        <b style="color:#14532d;">4️⃣ LƯU DỮ LIỆU</b><br>
                        <span style="font-size:0.81rem; color:#1e293b;">Bấm 💾 Lưu vào CSDL<br>(Tự động cộng vào Mẫu 02/03)</span>
                    </div>
                    <div style="font-size:1.3rem; font-weight:bold; color:#16a34a;">➔</div>
                    <div style="flex:1.25; min-width:190px; background:#ede9fe; border:1.5px solid #7c3aed; border-radius:8px; padding:10px;">
                        <b style="color:#4c1d95;">5️⃣ XUẤT FILE & NỘP USMART</b><br>
                        <span style="font-size:0.80rem; color:#1e293b;">
                            ⬇️ Tải .xlsx/.docx & Đẩy tự động sang uSmart:<br>
                            <a href="https://usmart.utc.edu.vn/baocaocongviec/" target="_blank" style="color:#5b21b6; font-weight:700;">🔗 Mở Cổng Báo cáo uSmart</a> ➔ Giao diện BC cá nhân ➔ Chọn <b>Nhập Web</b> ➔ Kiểm tra & Nộp
                        </span>
                    </div>
                </div>
            </div>
            """,
            unsafe_allow_html=True,
        )

        st.info(
            "⏰ **Mốc thời gian quy định (Thông báo số 940/TB-ĐHGTVT):**  \n"
            "- **Báo cáo Tuần (Mẫu 01a):** Hoàn thành trước **16h00 chiều Thứ Sáu** hàng tuần.  \n"
            "- **Báo cáo Tháng Cá nhân (Mẫu 01b):** Giảng viên hoàn thành từ **ngày 25 đến hết ngày 02 tháng kế tiếp**.  \n"
            "- **Báo cáo Tháng Đơn vị (Mẫu 02b, 02c - Dành cho Trưởng Bộ môn / Admin):** Tổng hợp và duyệt từ **ngày 02 đến ngày 05 tháng kế tiếp**."
        )


def _chuan_bi_dong_nhiem_vu_editor(bc, ds_viec_goi_y, ngay_moc_han: date) -> list[dict]:
    """Tạo danh sách dòng dữ liệu tương tác cho bảng Nhiệm vụ (dùng chung cho Mẫu 01a và Mẫu 01b)."""
    rows_editor = []
    for idx_nv, nv in enumerate(bc.danh_sach_nhiem_vu, start=1):
        rows_editor.append(
            {
                "tt": idx_nv,
                "su_dung": True,
                "tiep_tuc_ky_sau": bool(nv.la_cong_viec_lap_lai),
                "trong_tam_dv": bool(nv.la_nhiem_vu_trong_tam),
                "ma_nv": nv.ma_nhiem_vu,
                "nhiem_vu": nv.nhiem_vu,
                "nguon_giao": nv.nguon_giao if nv.nguon_giao in DANH_MUC_NGUON_GIAO else DANH_MUC_NGUON_GIAO[0],
                "so_vb": nv.so_van_ban_ma_viec,
                "san_pham": nv.san_pham_ket_qua,
                "thoi_han": nv.thoi_han,
                "ngay_hoan_thanh": nv.ngay_hoan_thanh,
                "tinh_trang": nv.tinh_trang if nv.tinh_trang in DANH_MUC_TINH_TRANG else DANH_MUC_TINH_TRANG[0],
                "muc_do": float(nv.muc_do_hoan_thanh),
                "kho_khan": nv.kho_khan_de_xuat,
                "da_den_han": nv.da_den_han_trong_ky(ngay_moc_han),
                "dung_han": nv.hoan_thanh_dung_han_usmart(),
                "canh_bao": nv.canh_bao_excel(),
                "chat_luong": nv.danh_gia_chat_luong if nv.danh_gia_chat_luong in DANH_MUC_CHAT_LUONG else DANH_MUC_CHAT_LUONG[0],
                "uu_tien": nv.muc_uu_tien if nv.muc_uu_tien in DANH_MUC_UU_TIEN else DANH_MUC_UU_TIEN[2],
                "nguon_goc": nv.nguon_du_lieu or "Trong kỳ",
            }
        )

    for idx_cu, nv_cu in enumerate(ds_viec_goi_y, start=len(rows_editor) + 1):
        rows_editor.append(
            {
                "tt": idx_cu,
                "su_dung": False,
                "tiep_tuc_ky_sau": True,
                "trong_tam_dv": bool(nv_cu.la_nhiem_vu_trong_tam),
                "ma_nv": nv_cu.ma_nhiem_vu,
                "nhiem_vu": nv_cu.nhiem_vu,
                "nguon_giao": nv_cu.nguon_giao if nv_cu.nguon_giao in DANH_MUC_NGUON_GIAO else DANH_MUC_NGUON_GIAO[0],
                "so_vb": nv_cu.so_van_ban_ma_viec,
                "san_pham": nv_cu.san_pham_ket_qua,
                "thoi_han": nv_cu.thoi_han,
                "ngay_hoan_thanh": nv_cu.ngay_hoan_thanh,
                "tinh_trang": nv_cu.tinh_trang if nv_cu.tinh_trang in DANH_MUC_TINH_TRANG else DANH_MUC_TINH_TRANG[0],
                "muc_do": float(nv_cu.muc_do_hoan_thanh),
                "kho_khan": nv_cu.kho_khan_de_xuat,
                "da_den_han": nv_cu.da_den_han_trong_ky(ngay_moc_han),
                "dung_han": nv_cu.hoan_thanh_dung_han_usmart(),
                "canh_bao": "Tích '☑️ Sử dụng' để lấy vào kỳ này",
                "chat_luong": nv_cu.danh_gia_chat_luong if nv_cu.danh_gia_chat_luong in DANH_MUC_CHAT_LUONG else DANH_MUC_CHAT_LUONG[0],
                "uu_tien": nv_cu.muc_uu_tien if nv_cu.muc_uu_tien in DANH_MUC_UU_TIEN else DANH_MUC_UU_TIEN[2],
                "nguon_goc": f"🔄 Gợi ý ({nv_cu.nguon_du_lieu})",
            }
        )
    return rows_editor


def _chuyen_edited_rows_thanh_nhiem_vu(edited_rows: list[dict], ma_cb_chon: str, ho_ten_gv: str) -> list[NhiemVuCongViec]:
    """Chuyển các dòng đang tích chọn `☑️ Sử dụng` trên `st.data_editor` thành danh sách `NhiemVuCongViec`."""
    ds_nv: list[NhiemVuCongViec] = []
    for r_item in edited_rows:
        if not r_item.get("su_dung", True):
            continue
        ten_nv_str = str(r_item.get("nhiem_vu") or "").strip()
        if not ten_nv_str:
            continue
        ds_nv.append(
            NhiemVuCongViec(
                ma_nhiem_vu=str(r_item.get("ma_nv") or ""),
                nhiem_vu=ten_nv_str,
                nguon_giao=str(r_item.get("nguon_giao") or DANH_MUC_NGUON_GIAO[0]),
                so_van_ban_ma_viec=str(r_item.get("so_vb") or ""),
                san_pham_ket_qua=str(r_item.get("san_pham") or ""),
                thoi_han=_parse_date(r_item.get("thoi_han")),
                ngay_hoan_thanh=_parse_date(r_item.get("ngay_hoan_thanh")),
                tinh_trang=str(r_item.get("tinh_trang") or DANH_MUC_TINH_TRANG[0]),
                muc_do_hoan_thanh=float(r_item.get("muc_do") if r_item.get("muc_do") is not None else 100.0),
                kho_khan_de_xuat=str(r_item.get("kho_khan") or ""),
                ma_can_bo_chu_tri=ma_cb_chon,
                ten_can_bo_chu_tri=ho_ten_gv,
                la_nhiem_vu_trong_tam=bool(r_item.get("trong_tam_dv", False)),
                la_cong_viec_lap_lai=bool(r_item.get("tiep_tuc_ky_sau", True)),
                muc_uu_tien=str(r_item.get("uu_tien") or "Thường"),
                danh_gia_chat_luong=str(r_item.get("chat_luong") or "Đạt yêu cầu – tốt"),
                nguon_du_lieu=str(r_item.get("nguon_goc") or "Cập nhật trên Webapp").replace("🔄 Gợi ý (", "").rstrip(")"),
            )
        )
    return ds_nv


def hien_thi_phan_he_giang_vien(
    db: QuanLyCSDLBộMôn,
    validator: BoKiemTraRangBuocVatLy,
    excel_gen: BoSinhBaoCaoExcel,
    docx_gen: BoSinhBaoCaoDocx,
    ma_cb_chon: str,
    gv_hien_tai,
    thang: int,
    nam: int,
    out_dir: Path,
    usmart_user: str,
    usmart_pass: str,
) -> None:
    """Hiển thị giao diện Báo cáo Công việc Cá nhân (đầy đủ cả Mẫu 01a - Theo Tuần và Mẫu 01b - Theo Tháng)."""
    bc = db.doc_bao_cao_ca_nhan(ma_cb_chon, thang=int(thang), nam=int(nam))
    kq_kt = validator.kiem_tra_bao_cao_ca_nhan(bc)

    st.markdown(
        "<div class='section-banner'>"
        f"<h3>📄 BÁO CÁO CÔNG VIỆC CÁ NHÂN — {gv_hien_tai.ho_va_ten.upper()} (MẪU 01A THEO TUẦN & MẪU 01B THEO THÁNG)</h3>"
        "<p>Hỗ trợ lập Báo cáo Tuần (Mẫu 01a), Báo cáo Tháng & Tự chấm điểm (Mẫu 01b), tích chọn tái sử dụng công việc cũ, "
        "chuyển thành nhiệm vụ trọng tâm đơn vị và nộp trực tiếp lên uSmart.</p>"
        "</div>",
        unsafe_allow_html=True,
    )

    # Hàng nút công cụ nhanh (Action Buttons)
    b1, b2, b3, b4, b5 = st.columns(5)
    if b1.button("🔄 Sao chép công việc lặp lại", use_container_width=True):
        so_moi = db.sao_chep_cong_viec_lap_lai_sang_thang_moi(
            int(thang), int(nam), ma_can_bo=ma_cb_chon
        )
        st.success(f"Đã tự động thêm {so_moi} công việc lặp lại vào kỳ {thang:02d}/{nam}!")
        st.rerun()

    if b2.button("⏪ Chuyển việc chưa xong kỳ trước", use_container_width=True):
        thang_tr = 12 if int(thang) == 1 else int(thang) - 1
        nam_tr = int(nam) - 1 if int(thang) == 1 else int(nam)
        so_ct = db.chuyen_tiep_nhiem_vu_chua_hoan_thanh(
            thang_tr, nam_tr, int(thang), int(nam), ma_can_bo=ma_cb_chon
        )
        st.info(f"Đã chuyển tiếp {so_ct} nhiệm vụ chưa hoàn thành (giữ nguyên thời hạn gốc).")
        st.rerun()

    if b3.button("🔎 Tự rà soát uSmart & Đề xuất việc", use_container_width=True):
        agent_scan = USmartBaoCaoAgent(
            usmart_user or gv_hien_tai.tai_khoan_usmart,
            usmart_pass or lay_mat_khau_usmart_mac_dinh(usmart_user or gv_hien_tai.tai_khoan_usmart),
        )
        ds_de_xuat, nhat_ky_scan = agent_scan.ra_soat_va_de_xuat_nhiem_vu_tu_usmart(
            gv_hien_tai, thang=int(thang), nam=int(nam)
        )
        ten_da_co = {nv.nhiem_vu.strip().lower() for nv in bc.danh_sach_nhiem_vu}
        ds_gop = list(bc.danh_sach_nhiem_vu)
        so_moi_usmart = 0
        for nv_dx in ds_de_xuat:
            if nv_dx.nhiem_vu.strip().lower() not in ten_da_co:
                ds_gop.append(nv_dx)
                ten_da_co.add(nv_dx.nhiem_vu.strip().lower())
                so_moi_usmart += 1
        if so_moi_usmart > 0:
            db.cap_nhat_danh_sach_nhiem_vu_thang(ma_cb_chon, int(thang), int(nam), ds_gop)
            bc = db.doc_bao_cao_ca_nhan(ma_cb_chon, thang=int(thang), nam=int(nam))
        st.session_state[f"scan_log_{ma_cb_chon}_{thang}_{nam}"] = (
            ds_de_xuat,
            nhat_ky_scan,
            so_moi_usmart,
        )
        st.rerun()

    if b4.button("🔍 Kiểm tra Ràng buộc Vật lý", use_container_width=True):
        if kq_kt.hop_le_de_chot:
            st.success(kq_kt.tom_tat_kiem_toan())
        else:
            st.error(kq_kt.tom_tat_kiem_toan())

    if b5.button("📥 Xuất Mẫu 01 (01a Tuần & 01b Tháng .xlsx + .docx)", type="primary", use_container_width=True):
        thu_muc_ky = out_dir / f"Ky_{int(thang):02d}_{int(nam)}"
        f_xlsx = excel_gen.xuat_mau_01_ca_nhan(bc, thu_muc_ky, cap_nhat_ca_bc_tuan=True)
        f_docx = docx_gen.xuat_docx_ca_nhan_01b(bc, thu_muc_ky)
        st.session_state[f"last_mau01_{ma_cb_chon}_{thang}_{nam}"] = [str(f_xlsx), str(f_docx)]
        st.success(f"Đã lưu báo cáo vào vị trí đã chọn:\n- Excel (01a & 01b): `{f_xlsx}`\n- Word bản cứng: `{f_docx}`")

    if st.session_state.get(f"scan_log_{ma_cb_chon}_{thang}_{nam}"):
        ds_dx_log, nhat_ky_log, so_moi_log = st.session_state[f"scan_log_{ma_cb_chon}_{thang}_{nam}"]
        with st.expander(
            f"🔎 Kết quả tự động rà soát nhiệm vụ trên uSmart cho {gv_hien_tai.ho_va_ten} "
            f"(Phát hiện {len(ds_dx_log)} nhiệm vụ, bổ sung mới {so_moi_log} việc vào CSDL)",
            expanded=True,
        ):
            for dong_nk in nhat_ky_log:
                st.markdown(f"- {dong_nk}")

    if st.session_state.get(f"last_mau01_{ma_cb_chon}_{thang}_{nam}"):
        hien_thi_nut_tai_file_truc_tiep(
            [Path(p) for p in st.session_state[f"last_mau01_{ma_cb_chon}_{thang}_{nam}"]],
            key_prefix=f"dl_mau01_{ma_cb_chon}_{thang}_{nam}",
        )

    # Hiển thị cảnh báo Vùng Cách ly nếu có
    if kq_kt.vung_cach_ly:
        with st.expander("⚠️ Danh sách Bản ghi trong Vùng Cách ly / Cảnh báo Dị biệt", expanded=not kq_kt.hop_le_de_chot):
            for item in kq_kt.vung_cach_ly:
                st.warning(
                    f"**[{item.muc_do} - {item.ma_quy_tac}]** `{item.doi_tuong}`: {item.nguyen_nhan}  \n"
                    f"👉 *Hướng dẫn xử lý:* {item.huong_dan_xu_ly}"
                )

    # Bộ chọn kiểu báo cáo cá nhân: Theo Tuần (Mẫu 01a) hoặc Theo Tháng (Mẫu 01b)
    chon_mau_01 = st.radio(
        "📌 Chọn loại Báo cáo Công việc Cá nhân cần hiển thị & chỉnh sửa:",
        options=[
            "📆 Mẫu 01b — Báo cáo kết quả công việc tháng và tự chấm điểm (BC tháng)",
            "📅 Mẫu 01a — Báo cáo kết quả công việc tuần (BC tuần)",
        ],
        horizontal=True,
        key=f"radio_mau01_{ma_cb_chon}_{thang}_{nam}",
    )

    hien_thi_viec_cu = st.checkbox(
        "📂 Hiển thị thêm công việc từ báo cáo cũ & ngân hàng việc lặp lại để tích chọn sử dụng tiếp (☑️ Sử dụng)",
        value=True,
        key=f"chk_viec_cu_{ma_cb_chon}_{thang}_{nam}",
    )
    ds_viec_goi_y = (
        db.lay_cong_viec_cu_va_lap_lai_goi_y(ma_cb_chon, int(thang), int(nam))
        if hien_thi_viec_cu
        else []
    )

    # Cấu hình cột dùng chung cho bảng nhiệm vụ Mẫu 01a và Mẫu 01b
    cau_hinh_cot_nv = {
        "tt": st.column_config.NumberColumn("TT", disabled=True, width="small"),
        "su_dung": st.column_config.CheckboxColumn(
            "☑️ Sử dụng",
            help="Tích chọn để đưa công việc này vào báo cáo hiện tại (bỏ tích để loại khỏi kỳ này)",
            default=True,
        ),
        "tiep_tuc_ky_sau": st.column_config.CheckboxColumn(
            "🔁 Tiếp tục kỳ sau",
            help="Tích chọn nếu công việc này tiếp tục sử dụng lặp lại ở các kỳ tiếp theo",
            default=True,
        ),
        "trong_tam_dv": st.column_config.CheckboxColumn(
            "⭐ Trọng tâm ĐV",
            help="Tích chọn để chuyển thành Nhiệm vụ trọng tâm của Đơn vị (tự động tổng hợp sang Mẫu 02a/02b của Bộ môn)",
            default=False,
        ),
        "nhiem_vu": st.column_config.TextColumn("Nhiệm vụ, công việc *", width="large", required=True),
        "nguon_giao": st.column_config.SelectboxColumn("Nguồn giao *", options=list(DANH_MUC_NGUON_GIAO), required=True),
        "so_vb": st.column_config.TextColumn("Số văn bản / mã việc trên Usmart", width="small"),
        "san_pham": st.column_config.TextColumn("Sản phẩm đầu ra, kết quả *", width="medium"),
        "thoi_han": st.column_config.DateColumn("Thời hạn *", format="DD/MM/YYYY"),
        "ngay_hoan_thanh": st.column_config.DateColumn("Ngày hoàn thành", format="DD/MM/YYYY"),
        "tinh_trang": st.column_config.SelectboxColumn("Tình trạng *", options=list(DANH_MUC_TINH_TRANG), required=True),
        "muc_do": st.column_config.NumberColumn("Mức độ hoàn thành (%)", min_value=0, max_value=100, step=5, format="%d%%"),
        "kho_khan": st.column_config.TextColumn("Khó khăn, vướng mắc, đề xuất", width="medium"),
        "da_den_han": st.column_config.NumberColumn("Đã đến hạn (tự tính)", disabled=True, width="small"),
        "dung_han": st.column_config.NumberColumn("Đúng hạn (tự tính)", disabled=True, width="small"),
        "canh_bao": st.column_config.TextColumn("Cảnh báo (tự tính)", disabled=True, width="medium"),
        "chat_luong": st.column_config.SelectboxColumn("Đánh giá chất lượng", options=list(DANH_MUC_CHAT_LUONG)),
        "uu_tien": st.column_config.SelectboxColumn("Mức ưu tiên", options=list(DANH_MUC_UU_TIEN)),
        "nguon_goc": st.column_config.TextColumn("Nguồn gốc", disabled=True, width="small"),
    }
    thu_tu_cot_nv = [
        "tt",
        "su_dung",
        "tiep_tuc_ky_sau",
        "trong_tam_dv",
        "nhiem_vu",
        "nguon_giao",
        "so_vb",
        "san_pham",
        "thoi_han",
        "ngay_hoan_thanh",
        "tinh_trang",
        "muc_do",
        "kho_khan",
        "da_den_han",
        "dung_han",
        "canh_bao",
        "chat_luong",
        "uu_tien",
        "nguon_goc",
    ]

    # =========================================================================
    # TRƯỜNG HỢP 1: MẪU 01A — BÁO CÁO KẾT QUẢ CÔNG VIỆC TUẦN (SHEET 'BC tuần')
    # =========================================================================
    if chon_mau_01.startswith("📅 Mẫu 01a"):
        cw1, cw2, cw3 = st.columns(3)
        tu_ngay_tuan = cw1.date_input(
            "Từ ngày (thứ Hai đầu tuần) *",
            value=_tinh_thu_hai_mac_dinh(int(thang), int(nam)),
            key=f"d_01a_{ma_cb_chon}_{thang}_{nam}",
        )
        den_ngay_tuan = tu_ngay_tuan + timedelta(days=4)
        tuan_so = tu_ngay_tuan.isocalendar()[1]
        cw2.metric("Đến ngày (cuối kỳ = C5 + 4)", den_ngay_tuan.strftime("%d/%m/%Y"))
        cw3.metric("Tuần số (=WEEKNUM(C5,2))", f"Tuần {tuan_so}")

        with st.container(border=True):
            col_hdr_l, col_hdr_r = st.columns([4, 1])
            col_hdr_l.markdown(f"**TRƯỜNG ĐẠI HỌC GIAO THÔNG VẬN TẢI**  \n**ĐƠN VỊ: {TEN_DON_VI_CHUAN.upper()}**")
            col_hdr_r.markdown("<div style='text-align:right;'><span class='form-badge'>Mẫu 01a</span></div>", unsafe_allow_html=True)
            st.markdown(
                "<h4 style='text-align:center; color:#0f2d59; margin-top:4px; margin-bottom:6px;'>"
                "BÁO CÁO KẾT QUẢ CÔNG VIỆC TUẦN"
                "</h4>",
                unsafe_allow_html=True,
            )
            info1, info2, info3 = st.columns(3)
            info1.markdown(
                f"- **Từ ngày (thứ Hai):** `{tu_ngay_tuan.strftime('%d/%m/%Y')}`  \n"
                f"- **Đến ngày (cuối kỳ):** `{den_ngay_tuan.strftime('%d/%m/%Y')}` (Tuần số: `{tuan_so}`)"
            )
            info2.markdown(
                f"- **Họ và tên:** **{bc.giang_vien.ho_va_ten}** (`{bc.giang_vien.ma_can_bo}`)  \n"
                f"- **Vị trí việc làm, chức danh:** {bc.giang_vien.chuc_danh_day_du}"
            )
            info3.markdown(
                f"- **Người nhận báo cáo:** **{bc.giang_vien.nguoi_nhan_bao_cao}**  \n"
                f"- **Biểu mẫu gốc:** Sheet `BC tuần` (`{bc.ten_file_xlsx_chuan}`)"
            )

        st.subheader("I. KẾT QUẢ THỰC HIỆN NHIỆM VỤ TRONG TUẦN")
        rows_01a = _chuan_bi_dong_nhiem_vu_editor(bc, ds_viec_goi_y, den_ngay_tuan)
        edited_rows_01a = st.data_editor(
            rows_01a,
            key=f"editor_nv_01a_{ma_cb_chon}_{thang}_{nam}",
            use_container_width=True,
            num_rows="dynamic",
            column_order=thu_tu_cot_nv,
            column_config=cau_hinh_cot_nv,
        )

        ds_nv_tuan = _chuyen_edited_rows_thanh_nhiem_vu(edited_rows_01a, ma_cb_chon, gv_hien_tai.ho_va_ten)
        tong_w = len(ds_nv_tuan)
        ht_dh_w = sum(1 for x in ds_nv_tuan if x.tinh_trang == "Hoàn thành đúng hạn")
        ht_cham_w = sum(1 for x in ds_nv_tuan if x.tinh_trang == "Hoàn thành chậm")
        dth_dung_w = sum(1 for x in ds_nv_tuan if x.tinh_trang == "Đang thực hiện – đúng tiến độ")
        dth_cham_w = sum(1 for x in ds_nv_tuan if x.tinh_trang == "Đang thực hiện – chậm tiến độ")
        chua_th_w = sum(1 for x in ds_nv_tuan if x.tinh_trang == "Chưa thực hiện")
        da_den_han_w = sum(x.da_den_han_trong_ky(den_ngay_tuan) for x in ds_nv_tuan)
        dung_han_w = sum(x.hoan_thanh_dung_han_usmart() for x in ds_nv_tuan)
        ty_le_w = (dung_han_w / da_den_han_w) if da_den_han_w > 0 else None

        with st.expander("📊 TỔNG HỢP MỤC I (Tự động tính theo đúng Dòng 26–35 của Mẫu 01a)", expanded=True):
            c_th_w1, c_th_w2 = st.columns([3, 2])
            c_th_w1.dataframe(
                [
                    {"Chỉ tiêu tổng hợp (Dòng 27–35 Mẫu 01a)": "Tổng số nhiệm vụ trong báo cáo", "Giá trị tự tính": str(tong_w)},
                    {"Chỉ tiêu tổng hợp (Dòng 27–35 Mẫu 01a)": "  • Tình trạng: Hoàn thành đúng hạn", "Giá trị tự tính": str(ht_dh_w)},
                    {"Chỉ tiêu tổng hợp (Dòng 27–35 Mẫu 01a)": "  • Tình trạng: Hoàn thành chậm", "Giá trị tự tính": str(ht_cham_w)},
                    {"Chỉ tiêu tổng hợp (Dòng 27–35 Mẫu 01a)": "  • Tình trạng: Đang thực hiện – đúng tiến độ", "Giá trị tự tính": str(dth_dung_w)},
                    {"Chỉ tiêu tổng hợp (Dòng 27–35 Mẫu 01a)": "  • Tình trạng: Đang thực hiện – chậm tiến độ", "Giá trị tự tính": str(dth_cham_w)},
                    {"Chỉ tiêu tổng hợp (Dòng 27–35 Mẫu 01a)": "  • Tình trạng: Chưa thực hiện", "Giá trị tự tính": str(chua_th_w)},
                    {"Chỉ tiêu tổng hợp (Dòng 27–35 Mẫu 01a)": "Số nhiệm vụ đã đến hạn trong kỳ (=SUM(K13:K24))", "Giá trị tự tính": str(da_den_han_w)},
                    {"Chỉ tiêu tổng hợp (Dòng 27–35 Mẫu 01a)": "Số nhiệm vụ hoàn thành đúng hạn (=SUM(L13:L24))", "Giá trị tự tính": str(dung_han_w)},
                    {
                        "Chỉ tiêu tổng hợp (Dòng 27–35 Mẫu 01a)": "Tỷ lệ hoàn thành đúng hạn (=E34/E33)",
                        "Giá trị tự tính": f"{ty_le_w * 100:.1f}%" if ty_le_w is not None else "N/A",
                    },
                ],
                use_container_width=True,
                hide_index=True,
            )
            c_th_w2.info(
                "📌 **Quy tắc tự tính của Mẫu 01a (BC tuần):**\n"
                "- **Thời điểm nộp:** Trước 15h00 thứ Sáu hàng tuần.\n"
                "- **Số BC tuần nộp đúng hạn** sẽ được chuyển tự động sang bảng theo dõi **Mẫu 02c** của Đơn vị."
            )

        st.subheader("II. KẾ HOẠCH CÔNG VIỆC TUẦN TỚI")
        rows_kh_w = [
            {
                "tt": idx_kh,
                "su_dung": True,
                "nhiem_vu": kh.nhiem_vu,
                "san_pham_du_kien": kh.san_pham_du_kien,
                "thoi_han": kh.thoi_han,
                "ghi_chu": kh.ghi_chu,
            }
            for idx_kh, kh in enumerate(bc.ke_hoach_thang_toi, start=1)
        ]
        edited_kh_w = st.data_editor(
            rows_kh_w,
            key=f"editor_kh_01a_{ma_cb_chon}_{thang}_{nam}",
            use_container_width=True,
            num_rows="dynamic",
            column_config={
                "tt": st.column_config.NumberColumn("TT", disabled=True, width="small"),
                "su_dung": st.column_config.CheckboxColumn("☑️ Sử dụng", default=True),
                "nhiem_vu": st.column_config.TextColumn("Nhiệm vụ, công việc *", width="large", required=True),
                "san_pham_du_kien": st.column_config.TextColumn("Sản phẩm dự kiến *", width="medium"),
                "thoi_han": st.column_config.DateColumn("Thời hạn", format="DD/MM/YYYY"),
                "ghi_chu": st.column_config.TextColumn("Ghi chú", width="medium"),
            },
        )

        st.subheader("III. ĐỀ XUẤT, KIẾN NGHỊ")
        kien_nghi_w = st.text_area(
            "Nội dung đề xuất, kiến nghị trong tuần (Ô A49:M49 Mẫu 01a):",
            value=bc.de_xuat_kien_nghi,
            height=68,
            key=f"kn_01a_{ma_cb_chon}_{thang}_{nam}",
        )
        ngay_ky_w_str = den_ngay_tuan.strftime("Hà Nội, ngày %d tháng %m năm %Y")
        sw_l, sw_r = st.columns(2)
        sw_l.markdown(
            "**NHẬN XÉT CỦA NGƯỜI NHẬN BÁO CÁO**  \n"
            "*(Trưởng bộ môn/trưởng đơn vị ghi nhận xét, ký và ghi rõ họ tên)*  \n"
            f"**{bc.giang_vien.nguoi_nhan_bao_cao}**"
        )
        sw_r.markdown(
            f"<div style='text-align:center;'><i>{ngay_ky_w_str}</i><br>"
            f"<b>NGƯỜI BÁO CÁO</b><br><i>(Ký và ghi rõ họ tên)</i><br><br>"
            f"<b>{bc.giang_vien.ho_va_ten}</b></div>",
            unsafe_allow_html=True,
        )

        def _luu_nhanh_01a() -> None:
            db.cap_nhat_danh_sach_nhiem_vu_thang(ma_cb_chon, int(thang), int(nam), ds_nv_tuan)
            ds_kh_w_moi = [
                KeHoachKyToi(
                    nhiem_vu=str(r.get("nhiem_vu") or "").strip(),
                    san_pham_du_kien=str(r.get("san_pham_du_kien") or "").strip(),
                    thoi_han=_parse_date(r.get("thoi_han")),
                    ghi_chu=str(r.get("ghi_chu") or "").strip(),
                    ca_nhan_chu_tri=gv_hien_tai.ho_va_ten,
                )
                for r in edited_kh_w
                if r.get("su_dung", True) and str(r.get("nhiem_vu") or "").strip()
            ]
            db.cap_nhat_ke_hoach_thang_toi(ma_cb_chon, int(thang), int(nam), ds_kh_w_moi)
            if kien_nghi_w.strip() != bc.de_xuat_kien_nghi.strip():
                db.cap_nhat_cham_diem_va_chuyen_mon(
                    ma_can_bo=ma_cb_chon,
                    thang=int(thang),
                    nam=int(nam),
                    hoat_dong_cm=bc.hoat_dong_chuyen_mon,
                    diem_tu_cham_chung=bc.diem_tu_cham_tieu_chi_chung,
                    diem_tu_cham_nv=bc.diem_tu_cham_ket_qua_nhiem_vu,
                    diem_tbm_chung=bc.diem_truong_don_vi_tieu_chi_chung,
                    diem_tbm_nv=bc.diem_truong_don_vi_ket_qua_nhiem_vu,
                    de_xuat_kien_nghi=kien_nghi_w.strip(),
                )

        bw_save, bw_xlsx, bw_usmart = st.columns(3)
        if bw_save.button("💾 Lưu Báo cáo Tuần (Mẫu 01a) vào CSDL", type="primary", use_container_width=True):
            _luu_nhanh_01a()
            st.success(f"Đã lưu Báo cáo Tuần Mẫu 01a ({len(ds_nv_tuan)} nhiệm vụ) vào CSDL!")
            st.rerun()

        if bw_xlsx.button("📥 Xuất Sheet 'BC tuần' (Mẫu 01a .xlsx)", use_container_width=True):
            _luu_nhanh_01a()
            bc_w = db.doc_bao_cao_ca_nhan(ma_cb_chon, thang=int(thang), nam=int(nam))
            thu_muc_ky = out_dir / f"Ky_{int(thang):02d}_{int(nam)}"
            f_xlsx = excel_gen.xuat_mau_01_ca_nhan(
                bc_w, thu_muc_ky, cap_nhat_ca_bc_tuan=True, ngay_thu_hai_tuan=tu_ngay_tuan
            )
            st.session_state[f"last_mau01_{ma_cb_chon}_{thang}_{nam}"] = [str(f_xlsx)]
            st.success(f"Đã xuất Mẫu 01a (Sheet `BC tuần` Tuần {tuan_so}): `{f_xlsx}`")

        if bw_usmart.button("☁️ Nộp Báo cáo Tuần (Mẫu 01a) lên uSmart", use_container_width=True):
            _luu_nhanh_01a()
            bc_w = db.doc_bao_cao_ca_nhan(ma_cb_chon, thang=int(thang), nam=int(nam))
            agent = USmartBaoCaoAgent(
                usmart_user or gv_hien_tai.tai_khoan_usmart,
                usmart_pass or lay_mat_khau_usmart_mac_dinh(usmart_user or gv_hien_tai.tai_khoan_usmart),
            )
            kq_u = agent.gui_bao_cao_webform(bc_w, ky_id="1", ma_mau="01a", che_do="nop")
            if kq_u.thanh_cong:
                st.success(kq_u.thong_bao)
            else:
                st.error(kq_u.thong_bao)
        st.caption(
            "👉 **Quy trình nộp lên uSmart (Nhập dữ liệu tự động sang uSmart):** "
            "Mở [🔗 **Cổng Báo cáo uSmart**](https://usmart.utc.edu.vn/baocaocongviec/) "
            "➔ **Giao diện Báo cáo cá nhân** ➔ Chọn **Nhập Web** ➔ **Kiểm tra và Nộp báo cáo**."
        )
        return

    # =========================================================================
    # TRƯỜNG HỢP 2: MẪU 01B — BÁO CÁO KẾT QUẢ CÔNG VIỆC THÁNG VÀ TỰ CHẤM ĐIỂM
    # =========================================================================
    with st.container(border=True):
        col_hdr_l, col_hdr_r = st.columns([4, 1])
        col_hdr_l.markdown(
            f"**TRƯỜNG ĐẠI HỌC GIAO THÔNG VẬN TẢI**  \n"
            f"**ĐƠN VỊ: {TEN_DON_VI_CHUAN.upper()}**"
        )
        col_hdr_r.markdown("<div style='text-align:right;'><span class='form-badge'>Mẫu 01b</span></div>", unsafe_allow_html=True)
        st.markdown(
            "<h4 style='text-align:center; color:#0f2d59; margin-top:4px; margin-bottom:6px;'>"
            "BÁO CÁO KẾT QUẢ CÔNG VIỆC THÁNG VÀ TỰ CHẤM ĐIỂM"
            "</h4>",
            unsafe_allow_html=True,
        )
        info1, info2, info3 = st.columns(3)
        info1.markdown(
            f"- **Tháng:** `{bc.thang:02d}` / **Năm:** `{bc.nam}`  \n"
            f"- **Ngày cuối kỳ:** `{bc.ngay_cuoi_ky.strftime('%d/%m/%Y')}`"
        )
        info2.markdown(
            f"- **Họ và tên:** **{bc.giang_vien.ho_va_ten}** (`{bc.giang_vien.ma_can_bo}`)  \n"
            f"- **Vị trí việc làm, chức danh:** {bc.giang_vien.chuc_danh_day_du}"
        )
        info3.markdown(
            f"- **Người nhận báo cáo:** **{bc.giang_vien.nguoi_nhan_bao_cao}**  \n"
            f"- **Tệp xuất chuẩn:** `{bc.ten_file_xlsx_chuan}`"
        )

    st.subheader("I. KẾT QUẢ THỰC HIỆN NHIỆM VỤ TRONG THÁNG")
    st.caption(
        "💡 **Hướng dẫn:** Tích chọn `☑️ Sử dụng` để đưa công việc (kể cả việc từ báo cáo cũ) vào báo cáo tháng này; "
        "tích `🔁 Tiếp tục kỳ sau` để tái sử dụng cho tháng sau; tích `⭐ Trọng tâm ĐV` để chuyển thành nhiệm vụ trọng tâm của đơn vị (Mẫu 02b). "
        "Mọi cột hạng mục, nguồn giao, thời hạn, ngày hoàn thành, tình trạng, mức độ (%) đều có thể chọn và sửa trực tiếp."
    )

    rows_editor = _chuan_bi_dong_nhiem_vu_editor(bc, ds_viec_goi_y, bc.ngay_cuoi_ky)
    edited_rows = st.data_editor(
        rows_editor,
        key=f"editor_nv_{ma_cb_chon}_{thang}_{nam}",
        use_container_width=True,
        num_rows="dynamic",
        column_order=thu_tu_cot_nv,
        column_config=cau_hinh_cot_nv,
    )

    ds_nv_dang_chon = _chuyen_edited_rows_thanh_nhiem_vu(edited_rows, ma_cb_chon, gv_hien_tai.ho_va_ten)
    tong_nv_live = len(ds_nv_dang_chon)
    so_ht_dung_han_tt = sum(1 for x in ds_nv_dang_chon if x.tinh_trang == "Hoàn thành đúng hạn")
    so_ht_cham_tt = sum(1 for x in ds_nv_dang_chon if x.tinh_trang == "Hoàn thành chậm")
    so_dang_th_dung = sum(1 for x in ds_nv_dang_chon if x.tinh_trang == "Đang thực hiện – đúng tiến độ")
    so_dang_th_cham = sum(1 for x in ds_nv_dang_chon if x.tinh_trang == "Đang thực hiện – chậm tiến độ")
    so_chua_th = sum(1 for x in ds_nv_dang_chon if x.tinh_trang == "Chưa thực hiện")
    so_da_den_han_live = sum(x.da_den_han_trong_ky(bc.ngay_cuoi_ky) for x in ds_nv_dang_chon)
    so_dung_han_live = sum(x.hoan_thanh_dung_han_usmart() for x in ds_nv_dang_chon)
    ty_le_dung_han_live = (so_dung_han_live / so_da_den_han_live) if so_da_den_han_live > 0 else None
    diem_tham_chieu_live = round(70.0 * ty_le_dung_han_live) if ty_le_dung_han_live is not None else 70.0

    with st.expander("📊 TỔNG HỢP MỤC I (Tự động tính theo đúng Dòng 29–38 của Mẫu 01b)", expanded=True):
        col_th_table, col_th_note = st.columns([3, 2])
        bang_tong_hop_muc_1 = [
            {"Chỉ tiêu tổng hợp (Dòng 30–38 Mẫu 01b)": "Tổng số nhiệm vụ trong báo cáo", "Giá trị tự tính": str(tong_nv_live)},
            {"Chỉ tiêu tổng hợp (Dòng 30–38 Mẫu 01b)": "  • Tình trạng: Hoàn thành đúng hạn", "Giá trị tự tính": str(so_ht_dung_han_tt)},
            {"Chỉ tiêu tổng hợp (Dòng 30–38 Mẫu 01b)": "  • Tình trạng: Hoàn thành chậm", "Giá trị tự tính": str(so_ht_cham_tt)},
            {"Chỉ tiêu tổng hợp (Dòng 30–38 Mẫu 01b)": "  • Tình trạng: Đang thực hiện – đúng tiến độ", "Giá trị tự tính": str(so_dang_th_dung)},
            {"Chỉ tiêu tổng hợp (Dòng 30–38 Mẫu 01b)": "  • Tình trạng: Đang thực hiện – chậm tiến độ", "Giá trị tự tính": str(so_dang_th_cham)},
            {"Chỉ tiêu tổng hợp (Dòng 30–38 Mẫu 01b)": "  • Tình trạng: Chưa thực hiện", "Giá trị tự tính": str(so_chua_th)},
            {"Chỉ tiêu tổng hợp (Dòng 30–38 Mẫu 01b)": "Số nhiệm vụ đã đến hạn trong kỳ (=SUM(K13:K27))", "Giá trị tự tính": str(so_da_den_han_live)},
            {"Chỉ tiêu tổng hợp (Dòng 30–38 Mẫu 01b)": "Số nhiệm vụ hoàn thành đúng hạn (=SUM(L13:L27))", "Giá trị tự tính": str(so_dung_han_live)},
            {
                "Chỉ tiêu tổng hợp (Dòng 30–38 Mẫu 01b)": "Tỷ lệ hoàn thành đúng hạn (=E37/E36)",
                "Giá trị tự tính": f"{ty_le_dung_han_live * 100:.1f}%" if ty_le_dung_han_live is not None else "N/A",
            },
        ]
        col_th_table.dataframe(bang_tong_hop_muc_1, use_container_width=True, hide_index=True)
        col_th_note.info(
            "📌 **Quy tắc tự tính của Mẫu 01b:**\n"
            "- **Tỷ lệ hoàn thành đúng hạn** = Số nhiệm vụ hoàn thành đúng hạn / Số nhiệm vụ đã đến hạn trong kỳ.\n"
            "- **Nhiệm vụ đã đến hạn:** có thời hạn đến ngày cuối kỳ hoặc đã hoàn thành (kể cả hoàn thành trước hạn).\n"
            "- **Điểm tham chiếu Mục III.2:** `70 × Tỷ lệ hoàn thành đúng hạn` = "
            f"**{diem_tham_chieu_live}/70 điểm**."
        )

    # ==================== MỤC II. HOẠT ĐỘNG CHUYÊN MÔN (DÒNG 40–45 MẪU 01B) ====================
    st.subheader("II. HOẠT ĐỘNG CHUYÊN MÔN (chỉ giảng viên kê khai)")
    cm_hien_tai = bc.hoat_dong_chuyen_mon
    rows_chuyen_mon = [
        {
            "noi_dung": "Giảng dạy (giờ chuẩn)",
            "thuc_hien_thang": float(cm_hien_tai.giang_day_thang),
            "luy_ke_nam": float(cm_hien_tai.giang_day_luy_ke),
            "dinh_muc_nam": float(cm_hien_tai.giang_day_dinh_muc),
            "ty_le_luy_ke": round((cm_hien_tai.giang_day_luy_ke / cm_hien_tai.giang_day_dinh_muc) * 100.0, 1)
            if cm_hien_tai.giang_day_dinh_muc > 0
            else 0.0,
            "ghi_chu": cm_hien_tai.giang_day_ghi_chu,
        },
        {
            "noi_dung": "Nghiên cứu khoa học (giờ chuẩn)",
            "thuc_hien_thang": float(cm_hien_tai.nckh_thang),
            "luy_ke_nam": float(cm_hien_tai.nckh_luy_ke),
            "dinh_muc_nam": float(cm_hien_tai.nckh_dinh_muc),
            "ty_le_luy_ke": round((cm_hien_tai.nckh_luy_ke / cm_hien_tai.nckh_dinh_muc) * 100.0, 1)
            if cm_hien_tai.nckh_dinh_muc > 0
            else 0.0,
            "ghi_chu": cm_hien_tai.nckh_ghi_chu,
        },
        {
            "noi_dung": "Nhiệm vụ khác theo quy định (giờ)",
            "thuc_hien_thang": float(cm_hien_tai.nhiem_vu_khac_thang),
            "luy_ke_nam": float(cm_hien_tai.nhiem_vu_khac_luy_ke),
            "dinh_muc_nam": float(cm_hien_tai.nhiem_vu_khac_dinh_muc),
            "ty_le_luy_ke": round((cm_hien_tai.nhiem_vu_khac_luy_ke / cm_hien_tai.nhiem_vu_khac_dinh_muc) * 100.0, 1)
            if cm_hien_tai.nhiem_vu_khac_dinh_muc > 0
            else 0.0,
            "ghi_chu": cm_hien_tai.nhiem_vu_khac_ghi_chu,
        },
    ]
    edited_cm = st.data_editor(
        rows_chuyen_mon,
        key=f"editor_cm_{ma_cb_chon}_{thang}_{nam}",
        use_container_width=True,
        hide_index=True,
        num_rows="fixed",
        column_config={
            "noi_dung": st.column_config.TextColumn("Nội dung", disabled=True, width="medium"),
            "thuc_hien_thang": st.column_config.NumberColumn(
                "Thực hiện trong tháng", min_value=0.0, max_value=1000.0, step=0.5, format="%.1f"
            ),
            "luy_ke_nam": st.column_config.NumberColumn(
                "Lũy kế năm học", min_value=0.0, max_value=3000.0, step=0.5, format="%.1f"
            ),
            "dinh_muc_nam": st.column_config.NumberColumn(
                "Định mức năm học", min_value=0.0, max_value=3000.0, step=0.5, format="%.1f"
            ),
            "ty_le_luy_ke": st.column_config.NumberColumn(
                "Tỷ lệ lũy kế / định mức (%)", disabled=True, format="%.1f%%"
            ),
            "ghi_chu": st.column_config.TextColumn("Ghi chú", width="large"),
        },
    )
    sp_khcn_moi = st.text_area(
        "Sản phẩm khoa học công nghệ, hướng dẫn người học trong tháng (bài báo, đề tài, sách, hướng dẫn đồ án/luận văn/luận án…):",
        value=cm_hien_tai.san_pham_khcn_huong_dan,
        height=68,
    )

    # ==================== MỤC III. TỰ ĐÁNH GIÁ, CHẤM ĐIỂM THÁNG (DÒNG 47–53 MẪU 01B) ====================
    st.subheader("III. TỰ ĐÁNH GIÁ, CHẤM ĐIỂM THÁNG")
    cd1, cd2, cd3, cd4 = st.columns(4)
    diem_tc_chung_moi = cd1.number_input(
        "1. Cá nhân tự chấm Tiêu chí chung (Tối đa 30)",
        min_value=0.0,
        max_value=30.0,
        value=float(bc.diem_tu_cham_tieu_chi_chung),
        step=0.5,
    )
    che_do_diem_nv = cd2.selectbox(
        "Cách tính Điểm mục 2 (Kết quả thực hiện NV - Tối đa 70)",
        options=[
            "Tự động tính = 70 × tỷ lệ hoàn thành đúng hạn",
            "Tự chọn / Chỉnh sửa điểm thủ công",
        ],
        index=0 if bc.diem_tu_cham_ket_qua_nhiem_vu is None else 1,
    )
    diem_tc_nv_goi_y = float(
        bc.diem_tu_cham_ket_qua_nhiem_vu
        if bc.diem_tu_cham_ket_qua_nhiem_vu is not None
        else diem_tham_chieu_live
    )
    diem_tc_nv_input = cd3.number_input(
        "2. Cá nhân tự chấm Kết quả NV (Tối đa 70)",
        min_value=0.0,
        max_value=70.0,
        value=diem_tc_nv_goi_y,
        step=1.0,
        disabled=(che_do_diem_nv.startswith("Tự động")),
    )
    diem_tc_nv_luu = None if che_do_diem_nv.startswith("Tự động") else float(diem_tc_nv_input)
    diem_tc_nv_hien_thi = float(diem_tham_chieu_live) if diem_tc_nv_luu is None else diem_tc_nv_luu

    tbm1, tbm2 = st.columns(2)
    co_diem_tbm = tbm1.checkbox(
        "Nhập điểm đánh giá của Trưởng đơn vị (Cột G - Mẫu 01b)",
        value=(bc.diem_truong_don_vi_tieu_chi_chung is not None or bc.diem_truong_don_vi_ket_qua_nhiem_vu is not None),
    )
    if co_diem_tbm:
        c_tbm1, c_tbm2 = tbm2.columns(2)
        diem_tbm_chung_moi: float | None = float(
            c_tbm1.number_input(
                "Trưởng ĐV chấm Tiêu chí chung (<=30)",
                min_value=0.0,
                max_value=30.0,
                value=float(bc.diem_truong_don_vi_tieu_chi_chung if bc.diem_truong_don_vi_tieu_chi_chung is not None else diem_tc_chung_moi),
                step=0.5,
            )
        )
        diem_tbm_nv_moi: float | None = float(
            c_tbm2.number_input(
                "Trưởng ĐV chấm Kết quả NV (<=70)",
                min_value=0.0,
                max_value=70.0,
                value=float(bc.diem_truong_don_vi_ket_qua_nhiem_vu if bc.diem_truong_don_vi_ket_qua_nhiem_vu is not None else diem_tc_nv_hien_thi),
                step=1.0,
            )
        )
    else:
        diem_tbm_chung_moi = None
        diem_tbm_nv_moi = None

    tong_diem_ca_nhan = diem_tc_chung_moi + diem_tc_nv_hien_thi
    tong_diem_tbm = (
        (diem_tbm_chung_moi or 0.0) + (diem_tbm_nv_moi or 0.0)
        if (diem_tbm_chung_moi is not None or diem_tbm_nv_moi is not None)
        else None
    )
    diem_xet_xep_loai = tong_diem_tbm if tong_diem_tbm is not None else tong_diem_ca_nhan
    xep_loai_tham_chieu = (
        DANH_MUC_XEP_LOAI[0]
        if diem_xet_xep_loai >= 90
        else (
            DANH_MUC_XEP_LOAI[1]
            if diem_xet_xep_loai >= 70
            else (DANH_MUC_XEP_LOAI[2] if diem_xet_xep_loai >= 50 else DANH_MUC_XEP_LOAI[3])
        )
    )
    cd4.metric("Tổng điểm & Mức xếp loại tham chiếu", f"{diem_xet_xep_loai:.1f} / 100", xep_loai_tham_chieu)

    bang_muc_3 = [
        {
            "Nhóm tiêu chí (Dòng 48–52 Mẫu 01b)": (
                "1. Tiêu chí chung (chính trị tư tưởng; đạo đức, lối sống; tác phong, lề lối làm việc; "
                "ý thức tổ chức kỷ luật, kể cả chấp hành chế độ báo cáo; tinh thần phối hợp)"
            ),
            "Điểm tối đa": "30",
            "Cá nhân tự chấm": f"{diem_tc_chung_moi:.1f}",
            "Trưởng đơn vị chấm": f"{diem_tbm_chung_moi:.1f}" if diem_tbm_chung_moi is not None else "",
            "Ghi chú": "Theo tiêu chí chi tiết tại Quy chế đánh giá, xếp loại của Trường",
        },
        {
            "Nhóm tiêu chí (Dòng 48–52 Mẫu 01b)": "2. Kết quả thực hiện nhiệm vụ (số lượng, chất lượng, tiến độ)",
            "Điểm tối đa": "70",
            "Cá nhân tự chấm": f"{diem_tc_nv_hien_thi:.1f}",
            "Trưởng đơn vị chấm": f"{diem_tbm_nv_moi:.1f}" if diem_tbm_nv_moi is not None else "",
            "Ghi chú": f"Điểm tham chiếu = 70 × tỷ lệ hoàn thành đúng hạn (mục I): {diem_tham_chieu_live}",
        },
        {
            "Nhóm tiêu chí (Dòng 48–52 Mẫu 01b)": "Tổng điểm",
            "Điểm tối đa": "100",
            "Cá nhân tự chấm": f"{tong_diem_ca_nhan:.1f}",
            "Trưởng đơn vị chấm": f"{tong_diem_tbm:.1f}" if tong_diem_tbm is not None else "",
            "Ghi chú": "Thang 100 điểm và ngưỡng 90 – 70 – 50 điểm theo Nghị định số 233/2026/NĐ-CP",
        },
        {
            "Nhóm tiêu chí (Dòng 48–52 Mẫu 01b)": "Mức xếp loại tham chiếu (lấy điểm của trưởng đơn vị nếu đã chấm)",
            "Điểm tối đa": "",
            "Cá nhân tự chấm": xep_loai_tham_chieu,
            "Trưởng đơn vị chấm": xep_loai_tham_chieu if tong_diem_tbm is not None else "",
            "Ghi chú": "Số liệu 'Số nhiệm vụ đã đến hạn' và 'Số nhiệm vụ hoàn thành đúng hạn' ở mục I được chuyển sang Mẫu 02c",
        },
    ]
    st.dataframe(bang_muc_3, use_container_width=True, hide_index=True)

    # ==================== MỤC IV. KẾ HOẠCH CÔNG VIỆC THÁNG TỚI (DÒNG 55–64 MẪU 01B) ====================
    st.subheader("IV. KẾ HOẠCH CÔNG VIỆC THÁNG TỚI")
    rows_kh_toi = [
        {
            "tt": idx_kh,
            "su_dung": True,
            "nhiem_vu": kh.nhiem_vu,
            "san_pham_du_kien": kh.san_pham_du_kien,
            "thoi_han": kh.thoi_han,
            "ghi_chu": kh.ghi_chu,
        }
        for idx_kh, kh in enumerate(bc.ke_hoach_thang_toi, start=1)
    ]
    edited_kh_toi = st.data_editor(
        rows_kh_toi,
        key=f"editor_kh_toi_{ma_cb_chon}_{thang}_{nam}",
        use_container_width=True,
        num_rows="dynamic",
        column_config={
            "tt": st.column_config.NumberColumn("TT", disabled=True, width="small"),
            "su_dung": st.column_config.CheckboxColumn("☑️ Sử dụng", default=True),
            "nhiem_vu": st.column_config.TextColumn("Nhiệm vụ, công việc *", width="large", required=True),
            "san_pham_du_kien": st.column_config.TextColumn("Sản phẩm dự kiến *", width="medium"),
            "thoi_han": st.column_config.DateColumn("Thời hạn", format="DD/MM/YYYY"),
            "ghi_chu": st.column_config.TextColumn("Ghi chú", width="medium"),
        },
    )

    # ==================== MỤC V. ĐỀ XUẤT, KIẾN NGHỊ & CHỮ KÝ (DÒNG 66–71 MẪU 01B) ====================
    st.subheader("V. ĐỀ XUẤT, KIẾN NGHỊ")
    kien_nghi_moi = st.text_area(
        "Nội dung đề xuất, kiến nghị với Bộ môn / Khoa / Trường (Ô A67:M67 Mẫu 01b):",
        value=bc.de_xuat_kien_nghi,
        height=72,
    )
    ngay_ky_str = (bc.ngay_ky_bao_cao or bc.ngay_cuoi_ky).strftime("Hà Nội, ngày %d tháng %m năm %Y")
    sig_l, sig_r = st.columns(2)
    sig_l.markdown(
        "**NHẬN XÉT CỦA NGƯỜI NHẬN BÁO CÁO**  \n"
        "*(Trưởng bộ môn/trưởng đơn vị ghi nhận xét, ký và ghi rõ họ tên)*  \n"
        f"**{bc.giang_vien.nguoi_nhan_bao_cao}**"
    )
    sig_r.markdown(
        f"<div style='text-align:center;'><i>{ngay_ky_str}</i><br>"
        f"<b>NGƯỜI BÁO CÁO</b><br><i>(Ký và ghi rõ họ tên)</i><br><br>"
        f"<b>{bc.giang_vien.ho_va_ten}</b></div>",
        unsafe_allow_html=True,
    )

    def _luu_toan_bo_mau_01b_hien_tai() -> tuple[int, int]:
        ds_nv_cap_nhat = _chuyen_edited_rows_thanh_nhiem_vu(edited_rows, ma_cb_chon, gv_hien_tai.ho_va_ten)
        so_nv_luu = db.cap_nhat_danh_sach_nhiem_vu_thang(
            ma_can_bo=ma_cb_chon,
            thang=int(thang),
            nam=int(nam),
            danh_sach_nhiem_vu_chon=ds_nv_cap_nhat,
        )

        r_gd = edited_cm[0] if len(edited_cm) > 0 else {}
        r_nc = edited_cm[1] if len(edited_cm) > 1 else {}
        r_kh = edited_cm[2] if len(edited_cm) > 2 else {}
        hd_moi = HoatDongChuyenMon(
            giang_day_thang=float(r_gd.get("thuc_hien_thang", cm_hien_tai.giang_day_thang) or 0.0),
            giang_day_luy_ke=float(r_gd.get("luy_ke_nam", cm_hien_tai.giang_day_luy_ke) or 0.0),
            giang_day_dinh_muc=float(r_gd.get("dinh_muc_nam", cm_hien_tai.giang_day_dinh_muc) or 0.0),
            giang_day_ghi_chu=str(r_gd.get("ghi_chu") or ""),
            nckh_thang=float(r_nc.get("thuc_hien_thang", cm_hien_tai.nckh_thang) or 0.0),
            nckh_luy_ke=float(r_nc.get("luy_ke_nam", cm_hien_tai.nckh_luy_ke) or 0.0),
            nckh_dinh_muc=float(r_nc.get("dinh_muc_nam", cm_hien_tai.nckh_dinh_muc) or 0.0),
            nckh_ghi_chu=str(r_nc.get("ghi_chu") or ""),
            nhiem_vu_khac_thang=float(r_kh.get("thuc_hien_thang", cm_hien_tai.nhiem_vu_khac_thang) or 0.0),
            nhiem_vu_khac_luy_ke=float(r_kh.get("luy_ke_nam", cm_hien_tai.nhiem_vu_khac_luy_ke) or 0.0),
            nhiem_vu_khac_dinh_muc=float(r_kh.get("dinh_muc_nam", cm_hien_tai.nhiem_vu_khac_dinh_muc) or 0.0),
            nhiem_vu_khac_ghi_chu=str(r_kh.get("ghi_chu") or ""),
            san_pham_khcn_huong_dan=sp_khcn_moi.strip(),
        )
        db.cap_nhat_cham_diem_va_chuyen_mon(
            ma_can_bo=ma_cb_chon,
            thang=int(thang),
            nam=int(nam),
            hoat_dong_cm=hd_moi,
            diem_tu_cham_chung=float(diem_tc_chung_moi),
            diem_tu_cham_nv=diem_tc_nv_luu,
            diem_tbm_chung=diem_tbm_chung_moi,
            diem_tbm_nv=diem_tbm_nv_moi,
            de_xuat_kien_nghi=kien_nghi_moi.strip(),
        )

        ds_kh_cap_nhat: list[KeHoachKyToi] = []
        for r_kh_item in edited_kh_toi:
            if not r_kh_item.get("su_dung", True):
                continue
            ten_kh_str = str(r_kh_item.get("nhiem_vu") or "").strip()
            if not ten_kh_str:
                continue
            ds_kh_cap_nhat.append(
                KeHoachKyToi(
                    nhiem_vu=ten_kh_str,
                    san_pham_du_kien=str(r_kh_item.get("san_pham_du_kien") or "").strip(),
                    thoi_han=_parse_date(r_kh_item.get("thoi_han")),
                    ghi_chu=str(r_kh_item.get("ghi_chu") or "").strip(),
                    ca_nhan_chu_tri=gv_hien_tai.ho_va_ten,
                )
            )
        so_kh_luu = db.cap_nhat_ke_hoach_thang_toi(
            ma_can_bo=ma_cb_chon,
            thang=int(thang),
            nam=int(nam),
            danh_sach_ke_hoach=ds_kh_cap_nhat,
        )
        return so_nv_luu, so_kh_luu

    if st.button(
        "💾 Lưu toàn bộ thay đổi (Mục I, II, III, IV, V) vào CSDL",
        type="primary",
        use_container_width=True,
    ):
        so_nv_luu, so_kh_luu = _luu_toan_bo_mau_01b_hien_tai()
        st.success(
            f"Đã lưu thành công toàn bộ Mẫu 01b ({so_nv_luu} nhiệm vụ tháng này, hoạt động chuyên môn Mục II, "
            f"điểm đánh giá Mục III, {so_kh_luu} kế hoạch tháng tới Mục IV và đề xuất kiến nghị Mục V) vào CSDL!"
        )
        st.rerun()

    # Biểu mẫu bổ sung nhanh nhiệm vụ mới
    with st.expander("➕ Bổ sung nhanh nhiệm vụ mới bằng biểu mẫu"):
        with st.form("form_them_nv"):
            ten_nv_moi = st.text_input("Nội dung nhiệm vụ, công việc *")
            c_ng, c_vb, c_tt = st.columns(3)
            nguon_moi = c_ng.selectbox("Nguồn giao *", options=list(DANH_MUC_NGUON_GIAO))
            so_vb_moi = c_vb.text_input("Số văn bản / Mã việc uSmart")
            tinh_trang_moi = c_tt.selectbox("Tình trạng *", options=list(DANH_MUC_TINH_TRANG))
            san_pham_moi = st.text_input("Sản phẩm đầu ra, kết quả cụ thể *")
            c_d1, c_d2, c_md = st.columns(3)
            thoi_han_moi = c_d1.date_input("Thời hạn hoàn thành *", value=bc.ngay_cuoi_ky)
            da_xong = c_d2.checkbox("Đã hoàn thành", value=True)
            ngay_ht_moi = c_d2.date_input("Ngày hoàn thành", value=bc.ngay_cuoi_ky) if da_xong else None
            muc_do_moi = c_md.number_input("Mức độ hoàn thành (%)", min_value=0, max_value=100, value=100)
            c_chk1, c_chk2 = st.columns(2)
            trong_tam_moi = c_chk1.checkbox("⭐ Là nhiệm vụ trọng tâm của Bộ môn (đưa vào Mẫu 02a/02b)")
            lap_lai_moi = c_chk2.checkbox("🔁 Sử dụng lặp lại tiếp tục ở các tháng sau", value=True)

            if st.form_submit_button("Lưu bổ sung vào CSDL"):
                nv_obj = NhiemVuCongViec(
                    ma_nhiem_vu="",
                    nhiem_vu=ten_nv_moi,
                    nguon_giao=nguon_moi,
                    so_van_ban_ma_viec=so_vb_moi,
                    san_pham_ket_qua=san_pham_moi,
                    thoi_han=thoi_han_moi,
                    ngay_hoan_thanh=ngay_ht_moi,
                    tinh_trang=tinh_trang_moi,
                    muc_do_hoan_thanh=float(muc_do_moi),
                    ma_can_bo_chu_tri=ma_cb_chon,
                    ten_can_bo_chu_tri=gv_hien_tai.ho_va_ten,
                    la_nhiem_vu_trong_tam=trong_tam_moi,
                    la_cong_viec_lap_lai=lap_lai_moi,
                    nguon_du_lieu="Giảng viên nhập bổ sung",
                )
                ma_sinh = db.them_nhiem_vu_moi(int(thang), int(nam), nv_obj)
                st.success(f"Đã thêm nhiệm vụ `{ma_sinh}` thành công!")
                st.rerun()

    # Khu vực gửi trực tiếp lên uSmart
    st.divider()
    st.subheader("🚀 Gửi trực tiếp Báo cáo Công việc Tháng (Mẫu 01b) & Báo cáo Tuần (Mẫu 01a) lên Hệ thống uSmart (`usmart.utc.edu.vn/baocaocongviec/`)")
    st.markdown(
        "📌 **Lộ trình nộp lên uSmart (Nhập dữ liệu tự động sang uSmart):**  \n"
        "1. Bấm một trong các nút đồng bộ tự động bên dưới (Hệ thống tự động lưu dữ liệu đang sửa vào CSDL trước khi gửi) ➔ "
        "2. Mở [**🔗 Cổng Báo cáo uSmart (`https://usmart.utc.edu.vn/baocaocongviec/`)**](https://usmart.utc.edu.vn/baocaocongviec/) ➔ "
        "3. Vào **Giao diện Báo cáo cá nhân** ➔ "
        "4. Chọn **Nhập Web** ➔ "
        "5. **Kiểm tra và Nộp báo cáo**."
    )
    tk_gui = usmart_user or gv_hien_tai.tai_khoan_usmart
    mk_gui = usmart_pass or lay_mat_khau_usmart_mac_dinh(tk_gui)
    u1, u2, u3 = st.columns(3)
    if u1.button("☁️ Lưu nháp trực tiếp lên uSmart (An toàn)", use_container_width=True):
        _luu_toan_bo_mau_01b_hien_tai()
        bc_moi = db.doc_bao_cao_ca_nhan(ma_cb_chon, thang=int(thang), nam=int(nam))
        agent = USmartBaoCaoAgent(tk_gui, mk_gui)
        kq_u = agent.gui_bao_cao_webform(bc_moi, ky_id="2", ma_mau="01b", che_do="luu_nhap")
        if kq_u.thanh_cong:
            st.success(kq_u.thong_bao)
        else:
            st.error(kq_u.thong_bao)

    if u2.button("🚀 Chốt & Nộp chính thức Webform lên uSmart", use_container_width=True):
        _luu_toan_bo_mau_01b_hien_tai()
        bc_moi = db.doc_bao_cao_ca_nhan(ma_cb_chon, thang=int(thang), nam=int(nam))
        agent = USmartBaoCaoAgent(tk_gui, mk_gui)
        kq_u = agent.gui_bao_cao_webform(bc_moi, ky_id="2", ma_mau="01b", che_do="nop")
        if kq_u.thanh_cong:
            st.success(kq_u.thong_bao)
        else:
            st.warning(kq_u.thong_bao)

    if u3.button("📤 Xuất & Upload tệp .xlsx chuẩn lên uSmart", use_container_width=True):
        _luu_toan_bo_mau_01b_hien_tai()
        bc_moi = db.doc_bao_cao_ca_nhan(ma_cb_chon, thang=int(thang), nam=int(nam))
        thu_muc_ky = out_dir / f"Ky_{int(thang):02d}_{int(nam)}"
        f_xlsx = excel_gen.xuat_mau_01_ca_nhan(bc_moi, thu_muc_ky, cap_nhat_ca_bc_tuan=True)
        st.session_state[f"last_mau01_{ma_cb_chon}_{thang}_{nam}"] = [str(f_xlsx)]
        agent = USmartBaoCaoAgent(tk_gui, mk_gui)
        kq_up = agent.tai_tep_excel_len_usmart(f_xlsx, ky_id="2", ma_mau="01b", bao_cao=bc_moi)
        if kq_up.thanh_cong:
            st.success(f"Phản hồi từ máy chủ uSmart: {kq_up.thong_bao}")
        else:
            st.warning(f"Phản hồi từ máy chủ uSmart: {kq_up.thong_bao}")


def hien_thi_phan_he_dong_bo_admin(
    db_path: Path,
    thang: int,
    nam: int,
    nckh_user: str,
    nckh_pass: str,
) -> None:
    """Hiển thị phân hệ Đồng bộ & Hiển thị Dữ liệu Nguồn (chỉ xuất hiện khi đăng nhập Admin)."""
    st.markdown(
        "<div class='section-banner'>"
        "<h3>🔄 QUẢN TRỊ ĐỒNG BỘ & HIỂN THỊ DỮ LIỆU NGUỒN (CHỈ DÀNH CHO TÀI KHOẢN ADMIN)</h3>"
        "<p>Nguồn dữ liệu: TKB Học kỳ (<code>2627_1_1_TKB dự kiến- ĐKT.xlsx</code>) & Cổng uSmart KHCN (<code>lannc</code>) | "
        "Thư mục lưu trữ: <code>Ke_hoach_cong_viec/</code> & Google Drive <code>1jawTokjmk-hnBsuoEIHoxYUhSRCEItPd</code></p>"
        "</div>",
        unsafe_allow_html=True,
    )

    nc0, nc1, nc2 = st.columns([2, 2, 2])
    gsheet_nckh_path = FOLDER_KE_HOACH_CONG_VIEC / TEN_FILE_GSHEET_NCKH
    if nc0.button("📅 Đồng bộ Thời khóa biểu Giảng dạy vào CSDL", use_container_width=True):
        kq_tkb = chay_dong_bo_tkb_giang_day_toan_trinh(
            thang=int(thang),
            nam=int(nam),
            duong_dan_csdl_master=db_path,
        )
        st.success(
            f"Đã đồng bộ TKB Giảng dạy: {kq_tkb.tong_so_lop_hoc_phan} lớp học phần "
            f"({kq_tkb.so_lop_giang_day_thang_11} lớp T11/2026), tổng {kq_tkb.tong_gio_quy_doi_toan_bo_mon:.1f} tiết chuẩn!"
        )

    if nc1.button("🔄 Đồng bộ NCKH & NCKHSV từ uSmart KHCN", type="primary", use_container_width=True):
        kq_sync = chay_dong_bo_ke_hoach_nckh_toan_trinh(
            username=nckh_user,
            password=nckh_pass,
            thang=int(thang),
            nam=int(nam),
            dong_bo_live=True,
            duong_dan_csdl_master=db_path,
        )
        st.success(
            f"Đã đồng bộ thành công từ uSmart ({kq_sync.tai_khoan_su_dung}): "
            f"{kq_sync.so_de_tai_cac_cap_2025_2027} đề tài các cấp (2025-2027), "
            f"{kq_sync.so_de_tai_nckhsv_2025_2027} đề tài NCKHSV (2025-2027), "
            f"tích hợp {kq_sync.so_nhiem_vu_dong_bo_bao_cao_thang} nhiệm vụ vào CSDL báo cáo tháng!"
        )

    if gsheet_nckh_path.exists():
        with open(gsheet_nckh_path, "rb") as f_bin:
            nc2.download_button(
                label="📥 Tải Google Sheet Kế hoạch NCKH (.xlsx)",
                data=f_bin.read(),
                file_name=TEN_FILE_GSHEET_NCKH,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )

    client_nckh = USmartNCKHClient(username=nckh_user, password=nckh_pass)
    raw_nckh, _ = client_nckh.tai_va_cap_nhat_du_lieu(dong_bo_live=False)
    proc_nckh = BoXuLyKeHoachNCKH(raw_nckh)
    ds_cc, ds_sv, ds_cl = proc_nckh.chuan_hoa_va_kiem_toan()

    cc_active = [x for x in ds_cc if x.hop_le_dua_vao_ke_hoach and x.dang_thuc_hien_2025_2027]
    sv_active = [x for x in ds_sv if x.hop_le_dua_vao_ke_hoach and (x.id_dot_dang_ky in (9176, 9167) or x.nam_ke_hoach >= 2025)]

    k1, k2, k3, k4 = st.columns(4)
    k1.metric("Đề tài các cấp (2025-2027)", f"{len(cc_active)} / {len([x for x in ds_cc if x.hop_le_dua_vao_ke_hoach])}")
    k2.metric("Đề tài NCKHSV (2025-2027)", f"{len(sv_active)} / {len([x for x in ds_sv if x.hop_le_dua_vao_ke_hoach])}")
    k3.metric("Tổng giờ chuẩn ĐT các cấp", f"{sum(x.gio_chuan_quy_doi for x in cc_active):.0f} giờ")
    k4.metric("Bản ghi Vùng cách ly", f"{len(ds_cl)} bản ghi")

    st.markdown("#### 1. Kế hoạch thực hiện Đề tài NCKH các cấp (Trường, Trường trọng điểm, Bộ, Tỉnh/Sở) giai đoạn 2025–2027")
    st.dataframe(
        [
            {
                "ID": x.id_de_tai,
                "Mã số": x.ma_de_tai,
                "Cấp đề tài": x.cap_de_tai,
                "Năm KH": x.nam_ke_hoach,
                "Tên đề tài NCKH": x.ten_de_tai,
                "Chủ nhiệm": x.chu_nhiem_ten,
                "Thành viên Bộ môn": x.thanh_vien_bo_mon_str,
                "Kinh phí (Tr.Đ)": x.kinh_phi_trieu_dong,
                "Giờ chuẩn": x.gio_chuan_quy_doi,
                "Trạng thái uSmart": x.trang_thai_hien_thi,
            }
            for x in cc_active
        ],
        use_container_width=True,
    )

    st.markdown("#### 2. Kế hoạch Hướng dẫn Sinh viên NCKH (`NCKHSV`) năm học 2025–2026 & 2026–2027")
    st.dataframe(
        [
            {
                "ID": x.id_de_tai,
                "Mã đề tài": x.ma_de_tai,
                "Đợt NCKHSV": x.dot_dang_ky,
                "Tên đề tài NCKH Sinh viên": x.ten_de_tai,
                "GV Hướng dẫn": x.gvhd_ten,
                "Số SV": x.so_sinh_vien,
                "Điểm nghiệm thu": x.diem_nghiem_thu or "",
                "Xếp giải": x.xep_giai,
                "Giờ quy đổi GVHD": x.gio_chuan_gvhd,
                "Trạng thái": x.trang_thai_hien_thi,
            }
            for x in sv_active
        ],
        use_container_width=True,
    )

    if ds_cl:
        with st.expander(f"⚠️ Vùng Cách ly & Kiểm toán Vật lý NCKH ({len(ds_cl)} bản ghi)", expanded=False):
            for item in ds_cl:
                st.warning(
                    f"**[{item.muc_do} - {item.ma_quy_tac}]** ID `{item.id_de_tai}` ({item.phan_loai}) — "
                    f"**{item.ten_de_tai}**: {item.nguyen_nhan}  \n"
                    f"👉 *Hướng dẫn xử lý:* {item.huong_dan_xu_ly}"
                )


@st.cache_resource
def _khoi_tao_bo_xu_ly_dung_chung():
    """Khởi tạo bộ kiểm tra ràng buộc vật lý và bộ sinh Excel/Word dùng chung (tối ưu RAM/CPU cho Webapp Cloud)."""
    return BoKiemTraRangBuocVatLy(), BoSinhBaoCaoExcel(), BoSinhBaoCaoDocx()


def main() -> None:
    st.set_page_config(
        page_title="Hệ thống Báo cáo Công việc Bộ môn ĐKT & KGN (TB 940)",
        page_icon="📊",
        layout="wide",
    )
    thiet_lap_css_chuyen_nghiep()

    st.markdown(
        "<div class='utc-top-banner'>"
        "<h2>📊 HỆ THỐNG BÁO CÁO CÔNG VIỆC & GIAO VIỆC BỘ MÔN ĐỊA KỸ THUẬT VÀ KHÔNG GIAN NGẦM</h2>"
        "<p>Trường Đại học Giao thông vận tải (UTC) — Thực hiện theo <b>Thông báo số 940/TB-ĐHGTVT</b> & "
        "<b>Nghị định số 233/2026/NĐ-CP</b> | Trưởng bộ môn: <b>GS.TS. Trần Thế Truyền</b></p>"
        "</div>",
        unsafe_allow_html=True,
    )

    db_path = PROJECT_ROOT / "CSDL_BaoCao_BoMon_DKT_KGN.xlsx"
    db = QuanLyCSDLBộMôn(db_path)
    db.khoi_tao_csdl_mau_neu_chua_co(ghi_de=False)
    if not st.session_state.get("da_dong_bo_gv_chuan_22"):
        db.dong_bo_danh_sach_giang_vien_chuan(thang=11, nam=2026)
        st.session_state["da_dong_bo_gv_chuan_22"] = True
    validator, excel_gen, docx_gen = _khoi_tao_bo_xu_ly_dung_chung()

    # Thanh bên (Sidebar): Chọn kỳ báo cáo, Giảng viên, Vị trí lưu file và Tài khoản
    with st.sidebar:
        st.header("⚙️ Cấu hình Kỳ & Thành viên")
        col_t, col_n = st.columns(2)
        thang = col_t.number_input("Tháng", min_value=1, max_value=12, value=11)
        nam = col_n.number_input("Năm", min_value=2025, max_value=2030, value=2026)

        ds_gv = db.doc_danh_sach_giang_vien()
        lua_chon_gv = {
            f"{gv.ma_can_bo} - {gv.ho_va_ten} ({gv.hoc_ham + '.' if gv.hoc_ham else ''}{gv.hoc_vi})": ma
            for ma, gv in ds_gv.items()
        }
        map_ma_sang_nhan = {v: k for k, v in lua_chon_gv.items()}
        nhan_mac_dinh = map_ma_sang_nhan.get("1051", list(lua_chon_gv.keys())[0])

        if "sb_chon_gv_label" not in st.session_state or st.session_state["sb_chon_gv_label"] not in lua_chon_gv:
            st.session_state["sb_chon_gv_label"] = nhan_mac_dinh
            gv_init = ds_gv[lua_chon_gv[nhan_mac_dinh]]
            st.session_state["sb_usmart_user"] = gv_init.tai_khoan_usmart
            st.session_state["sb_usmart_pass"] = lay_mat_khau_usmart_mac_dinh(gv_init.tai_khoan_usmart)

        def _khi_doi_giang_vien() -> None:
            lbl = st.session_state.get("sb_chon_gv_label", nhan_mac_dinh)
            mcb = lua_chon_gv.get(lbl, "1051")
            gv_obj = ds_gv[mcb]
            st.session_state["sb_usmart_user"] = gv_obj.tai_khoan_usmart
            st.session_state["sb_usmart_pass"] = lay_mat_khau_usmart_mac_dinh(gv_obj.tai_khoan_usmart)

        def _khi_nhap_tk_usmart() -> None:
            tk_nhap = st.session_state.get("sb_usmart_user", "").strip()
            mk_goi_y = lay_mat_khau_usmart_mac_dinh(tk_nhap)
            if mk_goi_y:
                st.session_state["sb_usmart_pass"] = mk_goi_y
            mcb_tim = tim_ma_can_bo_theo_tai_khoan_usmart(tk_nhap)
            if mcb_tim and mcb_tim in map_ma_sang_nhan:
                st.session_state["sb_chon_gv_label"] = map_ma_sang_nhan[mcb_tim]

        chon_key = st.selectbox(
            "👤 Chọn Thành viên Bộ môn (22 CB/GV)",
            options=list(lua_chon_gv.keys()),
            key="sb_chon_gv_label",
            on_change=_khi_doi_giang_vien,
        )
        ma_cb_chon = lua_chon_gv[chon_key]
        gv_hien_tai = ds_gv[ma_cb_chon]

        st.divider()
        st.subheader("🔐 Tài khoản Nộp Báo cáo uSmart")
        usmart_user = st.text_input(
            "Tài khoản uSmart Cá nhân",
            key="sb_usmart_user",
            on_change=_khi_nhap_tk_usmart,
            help="Nhập tài khoản uSmart (ví dụ: giangntt, ducla, linhnt2, hanh1, lannc...) để tự động chuyển sang đúng cán bộ và điền mật khẩu đã cấu hình.",
        )
        usmart_pass = st.text_input(
            "Mật khẩu uSmart Cá nhân",
            type="password",
            key="sb_usmart_pass",
            placeholder="Nhập mật khẩu uSmart cá nhân...",
        )
        st.caption(
            "💡 **Tài khoản đã cấu hình sẵn:** `giangntt` (`123456`), `ducla` (`123456`), `linhnt2` (`30032013`), `hanh1`, `lannc`."
        )
        st.markdown(
            "🔗 [Mở Cổng Báo cáo uSmart](https://usmart.utc.edu.vn/baocaocongviec/)  \n"
            "📂 [Mở Thư mục Google Drive Bộ môn](https://drive.google.com/drive/folders/1jawTokjmk-hnBsuoEIHoxYUhSRCEItPd)"
        )

        st.divider()
        st.subheader("📂 Lựa chọn Vị trí Lưu File")
        kieu_vi_tri = st.selectbox(
            "Nơi lưu tệp báo cáo (.xlsx & .docx)",
            options=[
                "📂 Thư mục mặc định của Dự án (output_bao_cao_dkt)",
                "☁️ Thư mục Kế hoạch công việc (Ke_hoach_cong_viec/Bao_cao_xuat)",
                "💻 Thư mục Downloads cá nhân (~/Downloads/Bao_cao_Bo_mon_DKT_KGN)",
                "🖥️ Thư mục Desktop cá nhân (~/Desktop/Bao_cao_Bo_mon_DKT_KGN)",
                "✍️ Tự nhập đường dẫn thư mục tùy chỉnh...",
            ],
            index=0,
        )
        duong_dan_tuy_chinh = ""
        if kieu_vi_tri.startswith("✍️"):
            duong_dan_tuy_chinh = st.text_input(
                "Nhập đường dẫn thư mục lưu file",
                value=str(PROJECT_ROOT / "output_bao_cao_dkt"),
                placeholder="Ví dụ: D:\\BaoCao_BoMon_DKT",
            )
        out_dir = xac_dinh_thu_muc_luu_file(kieu_vi_tri, duong_dan_tuy_chinh)
        st.caption(f"📍 **Đang lưu tại:** `{out_dir}`")

        st.divider()
        with st.expander("🛡️ Đăng nhập Quản trị viên (Admin)", expanded=False):
            st.caption("Chỉ tài khoản Admin mới hiển thị Module Báo cáo công việc đơn vị (Mẫu 02) & Đồng bộ dữ liệu nguồn.")
            admin_user = st.text_input("Tài khoản Admin", value="", key="adm_user", placeholder="admin / lannc / truyentt")
            admin_pass = st.text_input("Mật khẩu Admin", value="", type="password", key="adm_pass")
            is_admin = xac_thuc_tai_khoan_admin(admin_user, admin_pass)
            if admin_user or admin_pass:
                if is_admin:
                    st.success(f"Đã mở khóa quyền Admin ({admin_user}): Báo cáo Đơn vị (Mẫu 02) & Đồng bộ dữ liệu!")
                else:
                    st.error("Tài khoản hoặc mật khẩu Admin không đúng.")

            if is_admin:
                st.markdown("**Cấu hình tài khoản đồng bộ uSmart KHCN:**")
                nckh_user = st.text_input("TK uSmart KHCN (Bộ môn)", value="lannc")
                nckh_pass = st.text_input("Mật khẩu uSmart KHCN", value="giangchau0310", type="password")
                st.markdown("**☁️ Sao lưu & Khôi phục CSDL Tổng (Cloud <-> OneDrive):**")
                if db_path.exists():
                    with open(db_path, "rb") as f_db:
                        st.download_button(
                            label="⬇️ Tải CSDL Tổng (.xlsx) về máy/OneDrive",
                            data=f_db.read(),
                            file_name="CSDL_BaoCao_BoMon_DKT_KGN.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            use_container_width=True,
                        )
                up_db = st.file_uploader("📤 Nạp CSDL Tổng (.xlsx) lên Webapp Cloud", type=["xlsx"], key="up_csdl_master")
                if up_db is not None and st.button("🔄 Xác nhận ghi đè CSDL Tổng từ file tải lên", use_container_width=True):
                    db_path.write_bytes(up_db.getvalue())
                    st.success("Đã cập nhật CSDL Tổng thành công!")
                    st.rerun()
            else:
                nckh_user = "lannc"
                nckh_pass = "giangchau0310"

    # Danh sách Kiểu lựa chọn báo cáo:
    # Module Báo cáo đơn vị (Mẫu 02) và Module Đồng bộ dữ liệu nguồn CHỈ hiển thị với tài khoản Admin
    hien_thi_so_do_roadmap_huong_dan(is_admin)
    cac_kieu_bao_cao = lay_danh_sach_kieu_bao_cao_theo_quyen(is_admin)

    kieu_bao_cao_chon = st.radio(
        "🗂️ KIỂU LỰA CHỌN BÁO CÁO & THEO DÕI NHIỆM VỤ:",
        options=cac_kieu_bao_cao,
        horizontal=True,
        key="nav_kieu_bao_cao_chinh",
    )

    mk_hien_tai = usmart_pass or lay_mat_khau_usmart_mac_dinh(usmart_user or gv_hien_tai.tai_khoan_usmart) or "182128"

    if kieu_bao_cao_chon.startswith("📄 1."):
        hien_thi_phan_he_giang_vien(
            db=db,
            validator=validator,
            excel_gen=excel_gen,
            docx_gen=docx_gen,
            ma_cb_chon=ma_cb_chon,
            gv_hien_tai=gv_hien_tai,
            thang=int(thang),
            nam=int(nam),
            out_dir=out_dir,
            usmart_user=usmart_user,
            usmart_pass=mk_hien_tai,
        )
    elif kieu_bao_cao_chon.startswith("📋 2."):
        hien_thi_giao_viec_va_theo_doi_mau_03(
            db=db,
            excel_gen=excel_gen,
            thang=int(thang),
            nam=int(nam),
            out_dir=out_dir,
        )
    elif is_admin and kieu_bao_cao_chon.startswith("🏛️ 3."):
        hien_thi_bao_cao_don_vi_mau_02(
            db=db,
            excel_gen=excel_gen,
            docx_gen=docx_gen,
            thang=int(thang),
            nam=int(nam),
            out_dir=out_dir,
            db_path=db_path,
            ma_cb_chon=ma_cb_chon,
            usmart_user=usmart_user,
            usmart_pass=mk_hien_tai,
        )
    elif is_admin and kieu_bao_cao_chon.startswith("🔄 4."):
        hien_thi_phan_he_dong_bo_admin(
            db_path=db_path,
            thang=int(thang),
            nam=int(nam),
            nckh_user=nckh_user,
            nckh_pass=nckh_pass,
        )

    # Chân trang (Footer) cố định ở cuối giao diện Webapp
    st.markdown(
        "<div class='utc-footer'>"
        "📧 Email: <b>Nguyễn Hải Hà_Bộ môn Địa kỹ thuật và không gian ngầm</b> "
        "(<a href='mailto:haihadkt@utc.edu.vn'>haihadkt@utc.edu.vn</a>)"
        "</div>",
        unsafe_allow_html=True,
    )


if __name__ == "__main__":
    main()
