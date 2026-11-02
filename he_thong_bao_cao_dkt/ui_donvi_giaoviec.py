# -*- coding: utf-8 -*-
"""
Module Giao diện Webapp cho:
1. Báo cáo công việc đơn vị (Mẫu 02: 02a - BC tuần ĐV, 02b - BC tháng ĐV, 02c - Theo dõi cá nhân 21 GV)
2. Giao việc và theo dõi nhiệm vụ (Mẫu 03: 03a - Phiếu giao việc, 03b - Sổ giao việc 20 cột, 03c - Tổng hợp theo người)

Chuẩn 100% theo Thông báo số 940/TB-ĐHGTVT và Nghị định số 233/2026/NĐ-CP.
"""

from datetime import date, timedelta
from pathlib import Path
import streamlit as st

from he_thong_bao_cao_dkt.cli_baocao import TacTuDieuPhoiBaoCao
from he_thong_bao_cao_dkt.database_manager import QuanLyCSDLBộMôn, _parse_date
from he_thong_bao_cao_dkt.docx_generator import BoSinhBaoCaoDocx
from he_thong_bao_cao_dkt.excel_generator import BoSinhBaoCaoExcel
from he_thong_bao_cao_dkt.models import (
    DANH_MUC_CHAT_LUONG,
    DANH_MUC_NGUON_GIAO,
    DANH_MUC_TINH_TRANG,
    DANH_MUC_UU_TIEN,
    DANH_MUC_XEP_LOAI,
    DANH_SACH_GIANG_VIEN_DKT,
    LANH_DAO_PHU_TRACH_DON_VI_MAC_DINH,
    NGUOI_NHAN_BAO_CAO_MAC_DINH,
    NhiemVuCongViec,
    TEN_DON_VI_CHUAN,
    ThongKeUsmartDonVi,
)
from he_thong_bao_cao_dkt.usmart_baocao_client import USmartBaoCaoAgent


def _tinh_thu_hai_mac_dinh(thang: int, nam: int) -> date:
    """Tìm ngày thứ Hai đầu tiên trong tháng báo cáo."""
    d = date(nam, thang, 1)
    while d.weekday() != 0:
        d += timedelta(days=1)
    return d


def hien_thi_nut_tai_file_truc_tiep( danh_sach_file: list[Path], key_prefix: str = "dl" ) -> None:
    """Hiển thị các nút tải tệp (.xlsx / .docx) trực tiếp về máy qua trình duyệt (cho phép Save As chọn nơi lưu)."""
    files_ton_tai = [Path(f) for f in danh_sach_file if f and Path(f).exists()]
    if not files_ton_tai:
        return
    cols = st.columns(len(files_ton_tai))
    for idx, f_path in enumerate(files_ton_tai):
        mime_type = (
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            if f_path.suffix.lower() == ".xlsx"
            else "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )
        with open(f_path, "rb") as f_bin:
            cols[idx].download_button(
                label=f"⬇️ Tải về máy (Chọn nơi lưu): {f_path.name}",
                data=f_bin.read(),
                file_name=f_path.name,
                mime=mime_type,
                key=f"{key_prefix}_{idx}_{f_path.name}",
                use_container_width=True,
            )


def hien_thi_bao_cao_don_vi_mau_02(
    db: QuanLyCSDLBộMôn,
    excel_gen: BoSinhBaoCaoExcel,
    docx_gen: BoSinhBaoCaoDocx,
    thang: int,
    nam: int,
    out_dir: Path,
    db_path: Path,
    ma_cb_chon: str,
    usmart_user: str,
    usmart_pass: str,
) -> None:
    """Hiển thị phân hệ Báo cáo Công việc Đơn vị (Mẫu 02a - Tuần, Mẫu 02b - Tháng, Mẫu 02c - Theo dõi 21 GV)."""
    ds_all = db.doc_toan_bo_bao_cao_bo_mon(thang=int(thang), nam=int(nam))
    ngay_cuoi_ky = ds_all[0].ngay_cuoi_ky if ds_all else date(int(nam), int(thang), 28)

    st.markdown(
        "<div class='section-banner'>"
        "<h3>🏛️ BÁO CÁO CÔNG VIỆC ĐƠN VỊ (MẪU 02 — TUẦN, THÁNG & THEO DÕI CHẤM ĐIỂM 21 GIẢNG VIÊN)</h3>"
        "<p>Đơn vị: <b>Bộ môn Địa kỹ thuật và Không gian ngầm</b> | Trưởng đơn vị: <b>GS.TS. Trần Thế Truyền</b> | "
        f"Lãnh đạo Trường phụ trách: <b>{LANH_DAO_PHU_TRACH_DON_VI_MAC_DINH}</b></p>"
        "</div>",
        unsafe_allow_html=True,
    )

    # Thanh công cụ xuất nhanh toàn bộ Mẫu 02
    act1, act2, act3 = st.columns(3)
    if act1.button("📥 Xuất Mẫu 02 Đơn vị (02a, 02b, 02c .xlsx & .docx)", type="primary", use_container_width=True):
        thu_muc_ky = out_dir / f"Ky_{int(thang):02d}_{int(nam)}"
        f02_x = excel_gen.xuat_mau_02_don_vi(ds_all, thu_muc_ky, thang=int(thang), nam=int(nam))
        f02_d = docx_gen.xuat_docx_tong_hop_don_vi(ds_all, thu_muc_ky, thang=int(thang), nam=int(nam))
        st.session_state["last_exported_mau02"] = [str(f02_x), str(f02_d)]
        st.success(f"Đã lưu vào vị trí đã chọn:\n- Excel (02a, 02b, 02c): `{f02_x}`\n- Word tổng hợp: `{f02_d}`")

    if st.session_state.get("last_exported_mau02"):
        hien_thi_nut_tai_file_truc_tiep(
            [Path(p) for p in st.session_state["last_exported_mau02"]],
            key_prefix=f"dl_mau02_{thang}_{nam}",
        )

    if act2.button("🗂️ Xuất đồng loạt 21 Báo cáo Cá nhân (.xlsx & .docx)", use_container_width=True):
        import shutil

        orchestrator = TacTuDieuPhoiBaoCao(db_path, out_dir)
        kq_pipe = orchestrator.chay_pipeline_hoan_chinh(
            ma_can_bo=ma_cb_chon,
            thang=int(thang),
            nam=int(nam),
            xuat_toan_bo_21_gv=True,
            che_do_usmart="khong",
        )
        thu_muc_ky = out_dir / f"Ky_{int(thang):02d}_{int(nam)}"
        zip_base = out_dir / f"TronBo_BaoCao_21GV_DKT_KGN_{int(thang):02d}_{int(nam)}"
        zip_file = shutil.make_archive(str(zip_base), "zip", root_dir=str(thu_muc_ky))
        st.session_state["last_exported_zip_21gv"] = str(zip_file)
        st.success(
            f"Đã xuất thành công {len(kq_pipe.danh_sach_file_toan_bo_gv)} tệp báo cáo của 21 giảng viên vào `{thu_muc_ky}`!"
        )

    if st.session_state.get("last_exported_zip_21gv") and Path(st.session_state["last_exported_zip_21gv"]).exists():
        zip_p = Path(st.session_state["last_exported_zip_21gv"])
        with open(zip_p, "rb") as f_z:
            st.download_button(
                label=f"⬇️ Tải gói ZIP trọn bộ 21 Giảng viên & Đơn vị về máy (Chọn nơi lưu): {zip_p.name}",
                data=f_z.read(),
                file_name=zip_p.name,
                mime="application/zip",
                key=f"dl_zip_21gv_{thang}_{nam}",
                use_container_width=True,
            )

    if act3.button("☁️ Nộp Báo cáo Đơn vị lên Hệ thống uSmart", use_container_width=True):
        thu_muc_ky = out_dir / f"Ky_{int(thang):02d}_{int(nam)}"
        f02_x = excel_gen.xuat_mau_02_don_vi(ds_all, thu_muc_ky, thang=int(thang), nam=int(nam))
        agent = USmartBaoCaoAgent(usmart_user, usmart_pass)
        kq_up = agent.tai_tep_excel_len_usmart(f02_x, ky_id="2", ma_mau="02b")
        st.info(f"Phản hồi từ cổng uSmart: {kq_up.thong_bao}")

    chon_mau_02 = st.radio(
        "📌 Chọn biểu mẫu Báo cáo Công việc Đơn vị cần làm việc:",
        options=[
            "📅 Mẫu 02a — Báo cáo kết quả công việc tuần của đơn vị (BC tuần ĐV)",
            "📆 Mẫu 02b — Báo cáo kết quả công việc tháng của đơn vị (BC tháng ĐV)",
            "📊 Mẫu 02c — Bảng theo dõi kết quả công việc và chấm điểm tháng 21 GV (Theo dõi cá nhân)",
        ],
        horizontal=True,
        key=f"radio_mau02_{thang}_{nam}",
    )

    # Tập hợp toàn bộ nhiệm vụ của 21 giảng viên để cho phép lọc hoặc tích chọn Nhiệm vụ trọng tâm Đơn vị
    tat_ca_nv_bo_mon: list[tuple[str, str, NhiemVuCongViec]] = []
    for bc in ds_all:
        for nv in bc.danh_sach_nhiem_vu:
            tat_ca_nv_bo_mon.append((bc.giang_vien.ma_can_bo, bc.giang_vien.ho_va_ten, nv))

    # =========================================================================
    # 1. MẪU 02A — BÁO CÁO KẾT QUẢ CÔNG VIỆC TUẦN CỦA ĐƠN VỊ (SHEET 'BC tuần ĐV')
    # =========================================================================
    if chon_mau_02.startswith("📅 Mẫu 02a"):
        c_w1, c_w2, c_w3 = st.columns([2, 2, 2])
        tu_ngay_02a = c_w1.date_input(
            "Từ ngày (thứ Hai đầu tuần) *",
            value=_tinh_thu_hai_mac_dinh(int(thang), int(nam)),
            key=f"d_02a_{thang}_{nam}",
        )
        den_ngay_02a = tu_ngay_02a + timedelta(days=4)
        tuan_so_02a = tu_ngay_02a.isocalendar()[1]
        c_w2.metric("Đến ngày (cuối kỳ = C5 + 4)", den_ngay_02a.strftime("%d/%m/%Y"))
        c_w3.metric("Tuần số (=WEEKNUM(C5,2))", f"Tuần {tuan_so_02a}")

        with st.container(border=True):
            h_l, h_r = st.columns([4, 1])
            h_l.markdown(f"**TRƯỜNG ĐẠI HỌC GIAO THÔNG VẬN TẢI**  \n**ĐƠN VỊ: {TEN_DON_VI_CHUAN.upper()}**")
            h_r.markdown("<div style='text-align:right;'><span class='form-badge'>Mẫu 02a</span></div>", unsafe_allow_html=True)
            st.markdown(
                "<h4 style='text-align:center; color:#0f2d59; margin:4px 0;'>"
                "BÁO CÁO KẾT QUẢ CÔNG VIỆC TUẦN CỦA ĐƠN VỊ"
                "</h4>",
                unsafe_allow_html=True,
            )
            i1, i2, i3 = st.columns(3)
            i1.markdown(
                f"- **Từ ngày (thứ Hai):** `{tu_ngay_02a.strftime('%d/%m/%Y')}`  \n"
                f"- **Đến ngày (cuối kỳ):** `{den_ngay_02a.strftime('%d/%m/%Y')}` (Tuần số `{tuan_so_02a}`)"
            )
            i2.markdown(
                f"- **Trưởng đơn vị:** **{NGUOI_NHAN_BAO_CAO_MAC_DINH}**  \n"
                f"- **Lãnh đạo Trường phụ trách:** **{LANH_DAO_PHU_TRACH_DON_VI_MAC_DINH}**"
            )
            i3.markdown(
                f"- **Biểu mẫu gốc:** Sheet `BC tuần ĐV` (Mẫu 02)  \n"
                f"- **Số giảng viên trong đơn vị:** `{len(ds_all)}` giảng viên"
            )

        st.subheader("I. KẾT QUẢ THỰC HIỆN NHIỆM VỤ TRỌNG TÂM TRONG TUẦN")
        hien_tat_ca_02a = st.checkbox(
            "🔍 Hiển thị toàn bộ nhiệm vụ của 21 giảng viên để tích chọn chuyển thành Nhiệm vụ trọng tâm của Đơn vị (⭐ Trọng tâm ĐV)",
            value=False,
            key=f"chk_all_02a_{thang}_{nam}",
        )

        rows_02a = []
        da_thay = set()
        for ma_cb, ten_gv, nv in tat_ca_nv_bo_mon:
            if not hien_tat_ca_02a and not nv.la_nhiem_vu_trong_tam:
                continue
            khoa_duy_nhat = f"{ma_cb}_{nv.ma_nhiem_vu}_{nv.nhiem_vu}"
            if khoa_duy_nhat in da_thay:
                continue
            da_thay.add(khoa_duy_nhat)
            rows_02a.append(
                {
                    "tt": len(rows_02a) + 1,
                    "ma_nv": nv.ma_nhiem_vu,
                    "trong_tam_dv": bool(nv.la_nhiem_vu_trong_tam),
                    "nhiem_vu": nv.nhiem_vu,
                    "nguon_giao": nv.nguon_giao if nv.nguon_giao in DANH_MUC_NGUON_GIAO else DANH_MUC_NGUON_GIAO[0],
                    "so_vb": nv.so_van_ban_ma_viec,
                    "lanh_dao_pt": nv.lanh_dao_phu_trach or LANH_DAO_PHU_TRACH_DON_VI_MAC_DINH,
                    "ca_nhan_chu_tri": nv.ten_can_bo_chu_tri or ten_gv,
                    "don_vi_phoi_hop": nv.don_vi_phoi_hop or nv.nguoi_phoi_hop or "Khoa Công trình",
                    "thoi_han": nv.thoi_han,
                    "ngay_hoan_thanh": nv.ngay_hoan_thanh,
                    "tinh_trang": nv.tinh_trang if nv.tinh_trang in DANH_MUC_TINH_TRANG else DANH_MUC_TINH_TRANG[0],
                    "san_pham": nv.san_pham_ket_qua,
                    "kho_khan": nv.kho_khan_de_xuat,
                    "da_den_han": nv.da_den_han_trong_ky(den_ngay_02a),
                    "dung_han": nv.hoan_thanh_dung_han_usmart(),
                    "canh_bao": nv.canh_bao_excel(),
                }
            )

        edited_02a = st.data_editor(
            rows_02a,
            key=f"ed_02a_{thang}_{nam}",
            use_container_width=True,
            num_rows="dynamic",
            column_config={
                "tt": st.column_config.NumberColumn("TT", disabled=True, width="small"),
                "ma_nv": st.column_config.TextColumn("Mã NV", disabled=True, width="small"),
                "trong_tam_dv": st.column_config.CheckboxColumn(
                    "⭐ Trọng tâm ĐV",
                    help="Tích chọn để đưa nhiệm vụ vào Báo cáo Đơn vị (Mẫu 02a/02b)",
                    default=True,
                ),
                "nhiem_vu": st.column_config.TextColumn("Nhiệm vụ, công việc *", width="large", required=True),
                "nguon_giao": st.column_config.SelectboxColumn("Nguồn giao *", options=list(DANH_MUC_NGUON_GIAO)),
                "so_vb": st.column_config.TextColumn("Số văn bản / mã việc trên Usmart", width="small"),
                "lanh_dao_pt": st.column_config.TextColumn("Lãnh đạo Trường phụ trách", width="medium"),
                "ca_nhan_chu_tri": st.column_config.TextColumn("Cá nhân chủ trì trong đơn vị *", width="medium"),
                "don_vi_phoi_hop": st.column_config.TextColumn("Đơn vị phối hợp", width="small"),
                "thoi_han": st.column_config.DateColumn("Thời hạn *", format="DD/MM/YYYY"),
                "ngay_hoan_thanh": st.column_config.DateColumn("Ngày hoàn thành", format="DD/MM/YYYY"),
                "tinh_trang": st.column_config.SelectboxColumn("Tình trạng *", options=list(DANH_MUC_TINH_TRANG)),
                "san_pham": st.column_config.TextColumn("Kết quả / sản phẩm đạt được", width="medium"),
                "kho_khan": st.column_config.TextColumn("Khó khăn, vướng mắc, kiến nghị", width="medium"),
                "da_den_han": st.column_config.NumberColumn("Đã đến hạn (tự tính)", disabled=True, width="small"),
                "dung_han": st.column_config.NumberColumn("Đúng hạn (tự tính)", disabled=True, width="small"),
                "canh_bao": st.column_config.TextColumn("Cảnh báo (tự tính)", disabled=True, width="medium"),
            },
        )

        ds_chon_02a = [r for r in edited_02a if r.get("trong_tam_dv", True) and str(r.get("nhiem_vu") or "").strip()]
        _hien_thi_bang_tong_hop_9_chi_tieu(ds_chon_02a, "Mẫu 02a (Dòng 25–34)")

        # Mục II. Xử lý văn bản, công việc trên hệ thống Usmart trong tuần
        st.subheader("II. XỬ LÝ VĂN BẢN, CÔNG VIỆC TRÊN HỆ THỐNG USMART TRONG TUẦN")
        u_c1, u_c2, u_c3, u_c4, u_c5 = st.columns(5)
        vb_den_han_w = u_c1.number_input("Số VB, CV đến hạn trong tuần", min_value=0, value=5, key="u_w_1")
        vb_dung_han_w = u_c2.number_input("Hoàn thành đúng hạn", min_value=0, value=5, key="u_w_2")
        vb_cham_w = u_c3.number_input("Hoàn thành chậm", min_value=0, value=0, key="u_w_3")
        vb_qua_han_w = u_c4.number_input("Quá hạn chưa hoàn thành", min_value=0, value=0, key="u_w_4")
        do_lech_w = vb_den_han_w - (vb_dung_han_w + vb_cham_w + vb_qua_han_w)
        u_c5.metric(
            "Kiểm tra (phải bằng 0)",
            f"{do_lech_w}",
            "Hợp lệ (=0)" if do_lech_w == 0 else "Lệch số liệu!",
            delta_color="normal" if do_lech_w == 0 else "inverse",
        )

        # Mục III. Kế hoạch công tác tuần tới
        st.subheader("III. KẾ HOẠCH CÔNG TÁC TUẦN TỚI")
        rows_kh_02a = _lay_ke_hoach_don_vi_tu_21_gv(ds_all)
        st.data_editor(
            rows_kh_02a,
            key=f"ed_kh_02a_{thang}_{nam}",
            use_container_width=True,
            num_rows="dynamic",
            column_config={
                "tt": st.column_config.NumberColumn("TT", disabled=True, width="small"),
                "nhiem_vu": st.column_config.TextColumn("Nhiệm vụ, công việc dự kiến *", width="large"),
                "ca_nhan_chu_tri": st.column_config.TextColumn("Cá nhân chủ trì *", width="medium"),
                "don_vi_phoi_hop": st.column_config.TextColumn("Đơn vị phối hợp", width="medium"),
                "thoi_han": st.column_config.DateColumn("Thời hạn", format="DD/MM/YYYY"),
                "san_pham": st.column_config.TextColumn("Sản phẩm dự kiến", width="medium"),
            },
        )

        # Mục IV. Đề xuất, kiến nghị cần lãnh đạo Trường cho ý kiến
        st.subheader("IV. ĐỀ XUẤT, KIẾN NGHỊ CẦN LÃNH ĐẠO TRƯỜNG CHO Ý KIẾN")
        st.data_editor(
            [
                {
                    "tt": 1,
                    "noi_dung": "Bảo dưỡng, hiệu chuẩn định kỳ thiết bị thí nghiệm Cơ học đất – Địa kỹ thuật",
                    "phuong_an": "Đề nghị Nhà trường bố trí kinh phí bảo dưỡng định kỳ phục vụ đào tạo và kiểm định mở ngành",
                }
            ],
            key=f"ed_dx_02a_{thang}_{nam}",
            use_container_width=True,
            num_rows="dynamic",
            column_config={
                "tt": st.column_config.NumberColumn("TT", width="small"),
                "noi_dung": st.column_config.TextColumn("Nội dung vướng mắc, đề xuất", width="large"),
                "phuong_an": st.column_config.TextColumn("Phương án đề xuất của đơn vị", width="large"),
            },
        )
        _hien_thi_chu_ky_truong_don_vi(den_ngay_02a)

        if st.button("💾 Lưu danh sách Nhiệm vụ trọng tâm Đơn vị (Mẫu 02a) vào CSDL", type="primary", use_container_width=True):
            map_tt = {str(r.get("ma_nv")): bool(r.get("trong_tam_dv", False)) for r in edited_02a if r.get("ma_nv")}
            so_cn = db.cap_nhat_trong_tam_va_diem_tbm_don_vi(int(thang), int(nam), map_trong_tam_theo_ma_nv=map_tt)
            st.success(f"Đã cập nhật trạng thái Nhiệm vụ trọng tâm đơn vị ({so_cn} thay đổi) vào CSDL!")
            st.rerun()

    # =========================================================================
    # 2. MẪU 02B — BÁO CÁO KẾT QUẢ CÔNG VIỆC THÁNG CỦA ĐƠN VỊ (SHEET 'BC tháng ĐV')
    # =========================================================================
    elif chon_mau_02.startswith("📆 Mẫu 02b"):
        with st.container(border=True):
            h_l, h_r = st.columns([4, 1])
            h_l.markdown(f"**TRƯỜNG ĐẠI HỌC GIAO THÔNG VẬN TẢI**  \n**ĐƠN VỊ: {TEN_DON_VI_CHUAN.upper()}**")
            h_r.markdown("<div style='text-align:right;'><span class='form-badge'>Mẫu 02b</span></div>", unsafe_allow_html=True)
            st.markdown(
                "<h4 style='text-align:center; color:#0f2d59; margin:4px 0;'>"
                "BÁO CÁO KẾT QUẢ CÔNG VIỆC THÁNG CỦA ĐƠN VỊ"
                "</h4>",
                unsafe_allow_html=True,
            )
            i1, i2, i3 = st.columns(3)
            i1.markdown(
                f"- **Tháng:** `{int(thang):02d}` / **Năm:** `{int(nam)}`  \n"
                f"- **Ngày cuối kỳ:** `{ngay_cuoi_ky.strftime('%d/%m/%Y')}`"
            )
            i2.markdown(
                f"- **Trưởng đơn vị:** **{NGUOI_NHAN_BAO_CAO_MAC_DINH}**  \n"
                f"- **Lãnh đạo Trường phụ trách:** **{LANH_DAO_PHU_TRACH_DON_VI_MAC_DINH}**"
            )
            i3.markdown(
                f"- **Biểu mẫu gốc:** Sheet `BC tháng ĐV` (Mẫu 02)  \n"
                f"- **Tổng số viên chức, NLĐ:** `{len(ds_all)}` giảng viên"
            )

        st.subheader("I. KẾT QUẢ THỰC HIỆN NHIỆM VỤ TRONG THÁNG")
        hien_tat_ca_02b = st.checkbox(
            "🔍 Hiển thị toàn bộ nhiệm vụ của 21 giảng viên để tích chọn chuyển thành Nhiệm vụ trọng tâm của Đơn vị (⭐ Trọng tâm ĐV)",
            value=False,
            key=f"chk_all_02b_{thang}_{nam}",
        )

        rows_02b = []
        da_thay_02b = set()
        for ma_cb, ten_gv, nv in tat_ca_nv_bo_mon:
            if not hien_tat_ca_02b and not nv.la_nhiem_vu_trong_tam:
                continue
            khoa_duy_nhat = f"{ma_cb}_{nv.ma_nhiem_vu}_{nv.nhiem_vu}"
            if khoa_duy_nhat in da_thay_02b:
                continue
            da_thay_02b.add(khoa_duy_nhat)
            rows_02b.append(
                {
                    "tt": len(rows_02b) + 1,
                    "ma_nv": nv.ma_nhiem_vu,
                    "trong_tam_dv": bool(nv.la_nhiem_vu_trong_tam),
                    "nhiem_vu": nv.nhiem_vu,
                    "nguon_giao": nv.nguon_giao if nv.nguon_giao in DANH_MUC_NGUON_GIAO else DANH_MUC_NGUON_GIAO[0],
                    "so_vb": nv.so_van_ban_ma_viec,
                    "lanh_dao_pt": nv.lanh_dao_phu_trach or LANH_DAO_PHU_TRACH_DON_VI_MAC_DINH,
                    "ca_nhan_chu_tri": nv.ten_can_bo_chu_tri or ten_gv,
                    "don_vi_phoi_hop": nv.don_vi_phoi_hop or nv.nguoi_phoi_hop or "Khoa Công trình",
                    "thoi_han": nv.thoi_han,
                    "ngay_hoan_thanh": nv.ngay_hoan_thanh,
                    "tinh_trang": nv.tinh_trang if nv.tinh_trang in DANH_MUC_TINH_TRANG else DANH_MUC_TINH_TRANG[0],
                    "muc_do": float(nv.muc_do_hoan_thanh),
                    "san_pham": nv.san_pham_ket_qua,
                    "kho_khan": nv.kho_khan_de_xuat,
                    "da_den_han": nv.da_den_han_trong_ky(ngay_cuoi_ky),
                    "dung_han": nv.hoan_thanh_dung_han_usmart(),
                    "canh_bao": nv.canh_bao_excel(),
                }
            )

        edited_02b = st.data_editor(
            rows_02b,
            key=f"ed_02b_{thang}_{nam}",
            use_container_width=True,
            num_rows="dynamic",
            column_config={
                "tt": st.column_config.NumberColumn("TT", disabled=True, width="small"),
                "ma_nv": st.column_config.TextColumn("Mã NV", disabled=True, width="small"),
                "trong_tam_dv": st.column_config.CheckboxColumn(
                    "⭐ Trọng tâm ĐV",
                    help="Tích chọn để đưa nhiệm vụ vào Báo cáo Tháng Đơn vị (Mẫu 02b)",
                    default=True,
                ),
                "nhiem_vu": st.column_config.TextColumn("Nhiệm vụ, công việc *", width="large", required=True),
                "nguon_giao": st.column_config.SelectboxColumn("Nguồn giao *", options=list(DANH_MUC_NGUON_GIAO)),
                "so_vb": st.column_config.TextColumn("Số văn bản / mã việc trên Usmart", width="small"),
                "lanh_dao_pt": st.column_config.TextColumn("Lãnh đạo Trường phụ trách", width="medium"),
                "ca_nhan_chu_tri": st.column_config.TextColumn("Cá nhân chủ trì trong đơn vị *", width="medium"),
                "don_vi_phoi_hop": st.column_config.TextColumn("Đơn vị phối hợp", width="small"),
                "thoi_han": st.column_config.DateColumn("Thời hạn *", format="DD/MM/YYYY"),
                "ngay_hoan_thanh": st.column_config.DateColumn("Ngày hoàn thành", format="DD/MM/YYYY"),
                "tinh_trang": st.column_config.SelectboxColumn("Tình trạng *", options=list(DANH_MUC_TINH_TRANG)),
                "muc_do": st.column_config.NumberColumn("Mức độ hoàn thành (%)", min_value=0, max_value=100, step=5, format="%d%%"),
                "san_pham": st.column_config.TextColumn("Kết quả / sản phẩm đạt được", width="medium"),
                "kho_khan": st.column_config.TextColumn("Khó khăn, vướng mắc, kiến nghị", width="medium"),
                "da_den_han": st.column_config.NumberColumn("Đã đến hạn (tự tính)", disabled=True, width="small"),
                "dung_han": st.column_config.NumberColumn("Đúng hạn (tự tính)", disabled=True, width="small"),
                "canh_bao": st.column_config.TextColumn("Cảnh báo (tự tính)", disabled=True, width="medium"),
            },
        )

        ds_chon_02b = [r for r in edited_02b if r.get("trong_tam_dv", True) and str(r.get("nhiem_vu") or "").strip()]
        _hien_thi_bang_tong_hop_9_chi_tieu(ds_chon_02b, "Mẫu 02b (Dòng 28–37)")

        # Mục II & Mục III đặt cạnh nhau cho chuyên nghiệp
        col_m2, col_m3 = st.columns(2)
        with col_m2:
            st.subheader("II. THỐNG KÊ XỬ LÝ CÔNG VIỆC TRÊN USMART")
            vb_den_m = st.number_input("1. Số văn bản, công việc đến hạn trong tháng", min_value=0, value=12, key="u_m_1")
            vb_dung_m = st.number_input("2. Hoàn thành đúng hạn", min_value=0, value=12, key="u_m_2")
            vb_cham_m = st.number_input("3. Hoàn thành chậm", min_value=0, value=0, key="u_m_3")
            vb_qua_m = st.number_input("4. Quá hạn chưa hoàn thành", min_value=0, value=0, key="u_m_4")
            do_lech_m = vb_den_m - (vb_dung_m + vb_cham_m + vb_qua_m)
            ty_le_usmart = (vb_dung_m / vb_den_m * 100.0) if vb_den_m > 0 else 100.0
            st.info(
                f"- **Kiểm tra (phải bằng 0):** `{do_lech_m}` ({'Hợp lệ' if do_lech_m == 0 else 'Lệch tổng!'})  \n"
                f"- **Tỷ lệ hoàn thành đúng hạn trên Usmart:** **{ty_le_usmart:.1f}%**"
            )

        with col_m3:
            st.subheader("III. NHÂN SỰ & CHẤP HÀNH CHẾ ĐỘ BÁO CÁO")
            tong_ns = len(ds_all)
            so_phai_nop = st.number_input("Số người phải nộp báo cáo tháng", min_value=0, max_value=50, value=tong_ns)
            so_nop_dh = st.number_input("Số người nộp báo cáo tháng đúng hạn", min_value=0, max_value=50, value=tong_ns)
            bien_dong_ns = st.text_input(
                "Biến động nhân sự, phân công lại nhiệm vụ trong tháng (nếu có)",
                value="Không có biến động nhân sự trong tháng.",
            )
            ty_le_nop = (so_nop_dh / so_phai_nop * 100.0) if so_phai_nop > 0 else 100.0
            st.success(
                f"- **Tổng số viên chức, NLĐ của đơn vị:** `{tong_ns}` người  \n"
                f"- **Tỷ lệ nộp báo cáo tháng đúng hạn (=E51/E50):** **{ty_le_nop:.1f}%**"
            )

        # Mục IV. Kế hoạch công tác tháng tới
        st.subheader("IV. KẾ HOẠCH CÔNG TÁC THÁNG TỚI")
        rows_kh_02b = _lay_ke_hoach_don_vi_tu_21_gv(ds_all)
        st.data_editor(
            rows_kh_02b,
            key=f"ed_kh_02b_{thang}_{nam}",
            use_container_width=True,
            num_rows="dynamic",
            column_config={
                "tt": st.column_config.NumberColumn("TT", disabled=True, width="small"),
                "nhiem_vu": st.column_config.TextColumn("Nhiệm vụ, công việc dự kiến *", width="large"),
                "ca_nhan_chu_tri": st.column_config.TextColumn("Cá nhân chủ trì *", width="medium"),
                "don_vi_phoi_hop": st.column_config.TextColumn("Đơn vị phối hợp", width="medium"),
                "thoi_han": st.column_config.DateColumn("Thời hạn", format="DD/MM/YYYY"),
                "san_pham": st.column_config.TextColumn("Sản phẩm dự kiến", width="medium"),
            },
        )

        # Mục V. Đề xuất, kiến nghị cần lãnh đạo Trường cho ý kiến
        st.subheader("V. ĐỀ XUẤT, KIẾN NGHỊ CẦN LÃNH ĐẠO TRƯỜNG CHO Ý KIẾN")
        st.data_editor(
            [
                {
                    "tt": 1,
                    "noi_dung": "Bảo dưỡng, hiệu chuẩn định kỳ thiết bị thí nghiệm Cơ học đất – Địa kỹ thuật",
                    "phuong_an": "Đề nghị Nhà trường bố trí kinh phí bảo dưỡng định kỳ phục vụ đào tạo và kiểm định mở ngành",
                }
            ],
            key=f"ed_dx_02b_{thang}_{nam}",
            use_container_width=True,
            num_rows="dynamic",
            column_config={
                "tt": st.column_config.NumberColumn("TT", width="small"),
                "noi_dung": st.column_config.TextColumn("Nội dung vướng mắc, đề xuất", width="large"),
                "phuong_an": st.column_config.TextColumn("Phương án đề xuất của đơn vị", width="large"),
            },
        )
        _hien_thi_chu_ky_truong_don_vi(ngay_cuoi_ky)

        if st.button("💾 Lưu danh sách Nhiệm vụ trọng tâm Đơn vị (Mẫu 02b) vào CSDL", type="primary", use_container_width=True):
            map_tt = {str(r.get("ma_nv")): bool(r.get("trong_tam_dv", False)) for r in edited_02b if r.get("ma_nv")}
            so_cn = db.cap_nhat_trong_tam_va_diem_tbm_don_vi(int(thang), int(nam), map_trong_tam_theo_ma_nv=map_tt)
            st.success(f"Đã cập nhật trạng thái Nhiệm vụ trọng tâm đơn vị ({so_cn} thay đổi) vào CSDL!")
            st.rerun()

    # =========================================================================
    # 3. MẪU 02C — BẢNG THEO DÕI KẾT QUẢ CÔNG VIỆC VÀ CHẤM ĐIỂM THÁNG 21 GV
    # =========================================================================
    else:
        with st.container(border=True):
            h_l, h_r = st.columns([4, 1])
            h_l.markdown(f"**TRƯỜNG ĐẠI HỌC GIAO THÔNG VẬN TẢI**  \n**ĐƠN VỊ: {TEN_DON_VI_CHUAN.upper()}**")
            h_r.markdown("<div style='text-align:right;'><span class='form-badge'>Mẫu 02c</span></div>", unsafe_allow_html=True)
            st.markdown(
                "<h4 style='text-align:center; color:#0f2d59; margin:4px 0;'>"
                f"BẢNG THEO DÕI KẾT QUẢ CÔNG VIỆC VÀ CHẤM ĐIỂM THÁNG ({int(thang):02d}/{int(nam)}) CỦA VIÊN CHỨC, NGƯỜI LAO ĐỘNG"
                "</h4>",
                unsafe_allow_html=True,
            )

        rows_02c = []
        for idx, b in enumerate(ds_all, start=1):
            tl_bc_tuan = (b.so_bc_tuan_nop_dung_han / b.so_bc_tuan_phai_nop * 100.0) if b.so_bc_tuan_phai_nop > 0 else 100.0
            tl_nv = (b.ty_le_hoan_thanh_dung_han * 100.0) if b.ty_le_hoan_thanh_dung_han is not None else 100.0
            diem_tbm_tong = b.tong_diem_truong_don_vi if b.tong_diem_truong_don_vi is not None else b.tong_diem_tu_cham
            rows_02c.append(
                {
                    "tt": idx,
                    "ma_cb": b.giang_vien.ma_can_bo,
                    "ho_ten": b.giang_vien.ho_va_ten,
                    "vi_tri": b.giang_vien.chuc_danh_day_du,
                    "bc_tuan_phai_nop": int(b.so_bc_tuan_phai_nop),
                    "bc_tuan_dung_han": int(b.so_bc_tuan_nop_dung_han),
                    "ty_le_bc_tuan": round(tl_bc_tuan, 1),
                    "nv_den_han": int(b.so_nhiem_vu_da_den_han),
                    "nv_dung_han": int(b.so_nhiem_vu_hoan_thanh_dung_han),
                    "ty_le_nv_dung_han": round(tl_nv, 1),
                    "diem_tu_cham": float(b.tong_diem_tu_cham),
                    "diem_tbm_cham": float(diem_tbm_tong),
                    "xep_loai": b.muc_xep_loai_tham_chieu,
                    "ghi_chu": "Trưởng bộ môn" if b.giang_vien.ma_can_bo == "0694" else "Nộp báo cáo đúng hạn",
                }
            )

        edited_02c = st.data_editor(
            rows_02c,
            key=f"ed_02c_{thang}_{nam}",
            use_container_width=True,
            hide_index=True,
            num_rows="fixed",
            column_config={
                "tt": st.column_config.NumberColumn("TT", disabled=True, width="small"),
                "ma_cb": st.column_config.TextColumn("Mã CB", disabled=True, width="small"),
                "ho_ten": st.column_config.TextColumn("Họ và tên", disabled=True, width="medium"),
                "vi_tri": st.column_config.TextColumn("Vị trí việc làm", disabled=True, width="medium"),
                "bc_tuan_phai_nop": st.column_config.NumberColumn("Số BC tuần phải nộp", min_value=0, max_value=5, step=1),
                "bc_tuan_dung_han": st.column_config.NumberColumn("Số BC tuần nộp đúng hạn", min_value=0, max_value=5, step=1),
                "ty_le_bc_tuan": st.column_config.NumberColumn("Tỷ lệ nộp BC tuần đúng hạn", disabled=True, format="%.0f%%"),
                "nv_den_han": st.column_config.NumberColumn("Số NV đã đến hạn (Mẫu 01b)", disabled=True),
                "nv_dung_han": st.column_config.NumberColumn("Số hoàn thành đúng hạn (Mẫu 01b)", disabled=True),
                "ty_le_nv_dung_han": st.column_config.NumberColumn("Tỷ lệ hoàn thành đúng hạn", disabled=True, format="%.0f%%"),
                "diem_tu_cham": st.column_config.NumberColumn("Điểm tự chấm", disabled=True, format="%.1f"),
                "diem_tbm_cham": st.column_config.NumberColumn(
                    "Điểm trưởng đơn vị chấm *", min_value=0.0, max_value=100.0, step=0.5, format="%.1f"
                ),
                "xep_loai": st.column_config.TextColumn("Mức xếp loại tham chiếu", disabled=True, width="large"),
                "ghi_chu": st.column_config.TextColumn("Ghi chú", width="medium"),
            },
        )

        # Tổng hợp cuối bảng Mẫu 02c (Dòng 47–51 Mẫu 02c)
        ds_diem = [float(r.get("diem_tbm_cham") or r.get("diem_tu_cham") or 0.0) for r in edited_02c]
        so_xs = sum(1 for d in ds_diem if d >= 90.0)
        so_tot = sum(1 for d in ds_diem if 70.0 <= d < 90.0)
        so_ht = sum(1 for d in ds_diem if 50.0 <= d < 70.0)
        so_kht = sum(1 for d in ds_diem if d < 50.0)
        diem_tb_dv = sum(ds_diem) / len(ds_diem) if ds_diem else 0.0

        m1, m2, m3, m4, m5 = st.columns(5)
        m1.metric("Hoàn thành xuất sắc (≥90đ)", f"{so_xs} GV")
        m2.metric("Hoàn thành tốt (70–<90đ)", f"{so_tot} GV")
        m3.metric("Hoàn thành nhiệm vụ (50–<70đ)", f"{so_ht} GV")
        m4.metric("Không hoàn thành (<50đ)", f"{so_kht} GV")
        m5.metric("Điểm TB toàn Bộ môn", f"{diem_tb_dv:.1f} / 100")

        if st.button("💾 Lưu bảng chấm điểm của Trưởng đơn vị (Mẫu 02c) vào CSDL", type="primary", use_container_width=True):
            map_diem_tbm = {}
            for r in edited_02c:
                mcb = str(r.get("ma_cb") or "")
                tong_tbm = float(r.get("diem_tbm_cham") or 98.0)
                diem_chung = min(30.0, round(tong_tbm * 0.3, 1))
                diem_nv = max(0.0, round(tong_tbm - diem_chung, 1))
                map_diem_tbm[mcb] = (diem_chung, diem_nv)
            so_cn = db.cap_nhat_trong_tam_va_diem_tbm_don_vi(int(thang), int(nam), map_diem_tbm_theo_ma_cb=map_diem_tbm)
            st.success(f"Đã lưu điểm Trưởng đơn vị chấm cho {so_cn} giảng viên vào CSDL!")
            st.rerun()


def hien_thi_giao_viec_va_theo_doi_mau_03(
    db: QuanLyCSDLBộMôn,
    excel_gen: BoSinhBaoCaoExcel,
    thang: int,
    nam: int,
    out_dir: Path,
) -> None:
    """Hiển thị phân hệ Giao việc và Theo dõi nhiệm vụ (Mẫu 03a, Mẫu 03b 20 cột, Mẫu 03c Tổng hợp theo người)."""
    ds_all = db.doc_toan_bo_bao_cao_bo_mon(thang=int(thang), nam=int(nam))
    tu_ngay_ky = date(int(nam), int(thang), 1)
    den_ngay_ky = ds_all[0].ngay_cuoi_ky if ds_all else date(int(nam), int(thang), 28)

    st.markdown(
        "<div class='section-banner'>"
        "<h3>📋 GIAO VIỆC VÀ THEO DÕI THỰC HIỆN NHIỆM VỤ (MẪU 03 — PHIẾU GIAO VIỆC, SỔ GIAO VIỆC & TỔNG HỢP)</h3>"
        "<p>Áp dụng khi giao việc trực tiếp, qua cuộc họp, văn bản nội bộ, thư điện tử hoặc tin nhắn công vụ; "
        "kết quả cuối kỳ tự động tổng hợp sang <b>Mẫu 02c</b> và <b>Mẫu 01b</b>.</p>"
        "</div>",
        unsafe_allow_html=True,
    )

    col_btn_03, col_info_03 = st.columns([2, 3])
    if col_btn_03.button("📥 Xuất Sổ Giao việc & Theo dõi Nhiệm vụ Mẫu 03 (03a, 03b, 03c .xlsx)", type="primary", use_container_width=True):
        thu_muc_ky = out_dir / f"Ky_{int(thang):02d}_{int(nam)}"
        f03_x = excel_gen.xuat_mau_03_so_giao_viec(ds_all, thu_muc_ky, thang=int(thang), nam=int(nam))
        st.session_state["last_exported_mau03"] = [str(f03_x)]
        st.success(f"Đã lưu Sổ giao việc Mẫu 03 vào vị trí đã chọn: `{f03_x}`")
    col_info_03.info(
        f"📌 **Kỳ theo dõi hiện tại:** Từ ngày `{tu_ngay_ky.strftime('%d/%m/%Y')}` đến ngày `{den_ngay_ky.strftime('%d/%m/%Y')}` | "
        f"Vị trí lưu: `{out_dir}`"
    )
    if st.session_state.get("last_exported_mau03"):
        hien_thi_nut_tai_file_truc_tiep(
            [Path(p) for p in st.session_state["last_exported_mau03"]],
            key_prefix=f"dl_mau03_{thang}_{nam}",
        )

    chon_mau_03 = st.radio(
        "📌 Chọn biểu mẫu Giao việc & Theo dõi nhiệm vụ cần làm việc:",
        options=[
            "📓 Mẫu 03b — Sổ giao việc và theo dõi thực hiện nhiệm vụ (Bảng 20 cột chuẩn)",
            "📝 Mẫu 03a — Phiếu giao việc & Nghiệm thu kết quả (Chi tiết từng nhiệm vụ)",
            "👥 Mẫu 03c — Tổng hợp kết quả thực hiện nhiệm vụ được giao theo người (21 GV + Cộng đơn vị)",
        ],
        horizontal=True,
        key=f"radio_mau03_{thang}_{nam}",
    )

    # Tập hợp các nhiệm vụ trong Sổ giao việc
    ds_nv_so: list[tuple[str, str, NhiemVuCongViec]] = []
    for bc in ds_all:
        for nv in bc.danh_sach_nhiem_vu:
            ds_nv_so.append((bc.giang_vien.ma_can_bo, bc.giang_vien.ho_va_ten, nv))

    # =========================================================================
    # 1. MẪU 03B — SỔ GIAO VIỆC VÀ THEO DÕI THỰC HIỆN NHIỆM VỤ (ĐỦ 20 CỘT A..T)
    # =========================================================================
    if chon_mau_03.startswith("📓 Mẫu 03b"):
        with st.container(border=True):
            h_l, h_r = st.columns([4, 1])
            h_l.markdown(f"**TRƯỜNG ĐẠI HỌC GIAO THÔNG VẬN TẢI**  \n**ĐƠN VỊ: {TEN_DON_VI_CHUAN.upper()}**")
            h_r.markdown("<div style='text-align:right;'><span class='form-badge'>Mẫu 03b</span></div>", unsafe_allow_html=True)
            st.markdown(
                "<h4 style='text-align:center; color:#0f2d59; margin:4px 0;'>"
                "SỔ GIAO VIỆC VÀ THEO DÕI THỰC HIỆN NHIỆM VỤ"
                "</h4>",
                unsafe_allow_html=True,
            )
            st.caption(
                f"Kỳ theo dõi từ ngày: **{tu_ngay_ky.strftime('%d/%m/%Y')}** đến ngày: **{den_ngay_ky.strftime('%d/%m/%Y')}** | "
                "Đầy đủ 20 cột chuẩn Mẫu 03b (tự động sinh Mã việc `GV001..`, tự tính `Đã đến hạn`, `Đúng hạn`, `Quá hạn chưa xong`, `Số ngày trễ`, `Cảnh báo`)."
            )

        loc_gv = st.selectbox(
            "🔍 Lọc Sổ giao việc theo Giảng viên chủ trì (hoặc xem Toàn bộ 21 giảng viên Bộ môn):",
            options=["Tất cả 21 giảng viên Bộ môn"] + [f"{gv.ma_can_bo} - {gv.ho_va_ten}" for gv in DANH_SACH_GIANG_VIEN_DKT],
            key=f"loc_03b_{thang}_{nam}",
        )

        rows_03b = []
        for idx, (ma_cb, ten_gv, nv) in enumerate(ds_nv_so, start=1):
            if loc_gv != "Tất cả 21 giảng viên Bộ môn" and not loc_gv.startswith(ma_cb):
                continue
            tt_so = len(rows_03b) + 1
            ma_viec_gv = f"GV{tt_so:03d}"
            da_den_han = nv.da_den_han_trong_ky(den_ngay_ky)
            dung_han = nv.hoan_thanh_dung_han_usmart()
            qua_han_chua_xong = 1 if (nv.ngay_hoan_thanh is None and nv.thoi_han and nv.thoi_han < den_ngay_ky) else 0
            so_ngay_tre = 0
            if nv.thoi_han:
                if nv.ngay_hoan_thanh and nv.ngay_hoan_thanh > nv.thoi_han:
                    so_ngay_tre = (nv.ngay_hoan_thanh - nv.thoi_han).days
                elif nv.ngay_hoan_thanh is None and den_ngay_ky > nv.thoi_han:
                    so_ngay_tre = (den_ngay_ky - nv.thoi_han).days

            rows_03b.append(
                {
                    "tt": tt_so,
                    "ma_nv_goc": nv.ma_nhiem_vu,
                    "ma_cb": ma_cb,
                    "trong_tam_dv": bool(nv.la_nhiem_vu_trong_tam),
                    "nhiem_vu": nv.nhiem_vu,
                    "ma_viec": ma_viec_gv,
                    "ngay_giao": nv.ngay_giao or tu_ngay_ky,
                    "nguon_giao": nv.nguon_giao if nv.nguon_giao in DANH_MUC_NGUON_GIAO else DANH_MUC_NGUON_GIAO[0],
                    "so_vb": nv.so_van_ban_ma_viec,
                    "nguoi_chu_tri": nv.ten_can_bo_chu_tri or ten_gv,
                    "nguoi_phoi_hop": nv.nguoi_phoi_hop or "Các GV Bộ môn",
                    "san_pham": nv.san_pham_ket_qua,
                    "muc_uu_tien": nv.muc_uu_tien if nv.muc_uu_tien in DANH_MUC_UU_TIEN else DANH_MUC_UU_TIEN[2],
                    "thoi_han": nv.thoi_han,
                    "ngay_hoan_thanh": nv.ngay_hoan_thanh,
                    "tinh_trang": nv.tinh_trang if nv.tinh_trang in DANH_MUC_TINH_TRANG else DANH_MUC_TINH_TRANG[0],
                    "chat_luong": nv.danh_gia_chat_luong if nv.danh_gia_chat_luong in DANH_MUC_CHAT_LUONG else DANH_MUC_CHAT_LUONG[0],
                    "ghi_chu": nv.ghi_chu_so_giao_viec,
                    "da_den_han": da_den_han,
                    "dung_han": dung_han,
                    "qua_han_chua_xong": qua_han_chua_xong,
                    "so_ngay_tre": so_ngay_tre,
                    "canh_bao": nv.canh_bao_excel(),
                }
            )

        edited_03b = st.data_editor(
            rows_03b,
            key=f"ed_03b_{thang}_{nam}",
            use_container_width=True,
            num_rows="dynamic",
            column_order=[
                "tt",
                "ma_viec",
                "trong_tam_dv",
                "nhiem_vu",
                "ngay_giao",
                "nguon_giao",
                "so_vb",
                "nguoi_chu_tri",
                "nguoi_phoi_hop",
                "san_pham",
                "muc_uu_tien",
                "thoi_han",
                "ngay_hoan_thanh",
                "tinh_trang",
                "chat_luong",
                "ghi_chu",
                "da_den_han",
                "dung_han",
                "qua_han_chua_xong",
                "so_ngay_tre",
                "canh_bao",
            ],
            column_config={
                "tt": st.column_config.NumberColumn("TT", disabled=True, width="small"),
                "ma_viec": st.column_config.TextColumn("Mã việc", disabled=True, width="small"),
                "trong_tam_dv": st.column_config.CheckboxColumn("⭐ Trọng tâm ĐV", default=False),
                "nhiem_vu": st.column_config.TextColumn("Nhiệm vụ, công việc *", width="large", required=True),
                "ngay_giao": st.column_config.DateColumn("Ngày giao *", format="DD/MM/YYYY"),
                "nguon_giao": st.column_config.SelectboxColumn("Nguồn giao *", options=list(DANH_MUC_NGUON_GIAO)),
                "so_vb": st.column_config.TextColumn("Số văn bản / căn cứ", width="small"),
                "nguoi_chu_tri": st.column_config.TextColumn("Người chủ trì *", width="medium"),
                "nguoi_phoi_hop": st.column_config.TextColumn("Người phối hợp", width="small"),
                "san_pham": st.column_config.TextColumn("Sản phẩm đầu ra, tiêu chí đạt *", width="medium"),
                "muc_uu_tien": st.column_config.SelectboxColumn("Mức ưu tiên", options=list(DANH_MUC_UU_TIEN)),
                "thoi_han": st.column_config.DateColumn("Thời hạn *", format="DD/MM/YYYY"),
                "ngay_hoan_thanh": st.column_config.DateColumn("Ngày hoàn thành", format="DD/MM/YYYY"),
                "tinh_trang": st.column_config.SelectboxColumn("Tình trạng *", options=list(DANH_MUC_TINH_TRANG)),
                "chat_luong": st.column_config.SelectboxColumn("Đánh giá chất lượng", options=list(DANH_MUC_CHAT_LUONG)),
                "ghi_chu": st.column_config.TextColumn("Ghi chú (điều chỉnh thời hạn, lý do)", width="medium"),
                "da_den_han": st.column_config.NumberColumn("Đã đến hạn trong kỳ", disabled=True, width="small"),
                "dung_han": st.column_config.NumberColumn("Đúng hạn", disabled=True, width="small"),
                "qua_han_chua_xong": st.column_config.NumberColumn("Quá hạn chưa xong", disabled=True, width="small"),
                "so_ngay_tre": st.column_config.NumberColumn("Số ngày trễ", disabled=True, width="small"),
                "canh_bao": st.column_config.TextColumn("Cảnh báo", disabled=True, width="medium"),
            },
        )

        if st.button("💾 Lưu cập nhật Sổ giao việc (Mẫu 03b) vào CSDL", type="primary", use_container_width=True):
            map_tt = {str(r.get("ma_nv_goc")): bool(r.get("trong_tam_dv", False)) for r in edited_03b if r.get("ma_nv_goc")}
            db.cap_nhat_trong_tam_va_diem_tbm_don_vi(int(thang), int(nam), map_trong_tam_theo_ma_nv=map_tt)
            st.success("Đã lưu cập nhật Sổ giao việc (Mẫu 03b) vào CSDL Bộ môn!")
            st.rerun()

    # =========================================================================
    # 2. MẪU 03A — PHIẾU GIAO VIỆC & NGHIỆM THU KẾT QUẢ
    # =========================================================================
    elif chon_mau_03.startswith("📝 Mẫu 03a"):
        lua_chon_phieu = [
            f"GV{idx:03d} — {nv.ten_can_bo_chu_tri or ten_gv}: {nv.nhiem_vu[:70]}"
            for idx, (_, ten_gv, nv) in enumerate(ds_nv_so, start=1)
        ]
        chon_p = st.selectbox("📌 Chọn nhiệm vụ trong Sổ giao việc để hiển thị/in Phiếu giao việc (Mẫu 03a):", options=lua_chon_phieu)
        idx_chon = lua_chon_phieu.index(chon_p) if chon_p in lua_chon_phieu else 0
        ma_cb_p, ten_gv_p, nv_p = ds_nv_so[idx_chon]
        ma_gv_code = f"GV{idx_chon + 1:03d}"

        with st.container(border=True):
            h_l, h_r = st.columns([4, 1])
            h_l.markdown(f"**TRƯỜNG ĐẠI HỌC GIAO THÔNG VẬN TẢI**  \n**ĐƠN VỊ: {TEN_DON_VI_CHUAN.upper()}**")
            h_r.markdown("<div style='text-align:right;'><span class='form-badge'>Mẫu 03a</span></div>", unsafe_allow_html=True)
            st.markdown("<h4 style='text-align:center; color:#0f2d59;'>PHIẾU GIAO VIỆC</h4>", unsafe_allow_html=True)

            p1, p2, p3 = st.columns(3)
            ma_viec_in = p1.text_input("Mã việc (theo Sổ)", value=ma_gv_code)
            ngay_giao_in = p2.date_input("Ngày giao", value=nv_p.ngay_giao or tu_ngay_ky)
            uu_tien_in = p3.selectbox(
                "Mức ưu tiên",
                options=list(DANH_MUC_UU_TIEN),
                index=list(DANH_MUC_UU_TIEN).index(nv_p.muc_uu_tien) if nv_p.muc_uu_tien in DANH_MUC_UU_TIEN else 2,
            )

            st.markdown("##### I. NỘI DUNG GIAO VIỆC")
            f1, f2, f3 = st.columns(3)
            nguoi_giao_in = f1.text_input("Người giao việc (họ tên, chức vụ)", value=f"{NGUOI_NHAN_BAO_CAO_MAC_DINH} (Trưởng bộ môn)")
            nguoi_ct_in = f2.text_input("Người chủ trì thực hiện (họ tên, vị trí việc làm)", value=nv_p.ten_can_bo_chu_tri or ten_gv_p)
            nguoi_ph_in = f3.text_input("Người phối hợp", value=nv_p.nguoi_phoi_hop or "Các giảng viên trong Bộ môn")

            noi_dung_in = st.text_area("Nội dung nhiệm vụ", value=nv_p.nhiem_vu, height=68)
            c_cc, c_sp, c_th = st.columns([2, 2, 1])
            can_cu_in = c_cc.text_input("Căn cứ / nguồn giao (số văn bản, kết luận cuộc họp…)", value=f"{nv_p.nguon_giao} ({nv_p.so_van_ban_ma_viec})")
            san_pham_in = c_sp.text_input("Sản phẩm đầu ra và tiêu chí đạt", value=nv_p.san_pham_ket_qua)
            thoi_han_in = c_th.date_input("Thời hạn hoàn thành", value=nv_p.thoi_han or den_ngay_ky)

            st.markdown("**Các mốc trung gian (nếu có):**")
            st.data_editor(
                [
                    {
                        "moc": 1,
                        "noi_dung_moc": "Triển khai thực hiện và rà soát số liệu/hồ sơ chuyên môn",
                        "thoi_han_moc": (nv_p.thoi_han or den_ngay_ky) - timedelta(days=5),
                        "san_pham_kiem_tra": "Bản dự thảo / Hồ sơ bước 1",
                    },
                    {
                        "moc": 2,
                        "noi_dung_moc": "Hoàn thiện sản phẩm đầu ra và nghiệm thu nộp Bộ môn",
                        "thoi_han_moc": nv_p.thoi_han or den_ngay_ky,
                        "san_pham_kiem_tra": nv_p.san_pham_ket_qua,
                    },
                ],
                key=f"ed_moc_03a_{idx_chon}",
                use_container_width=True,
                hide_index=True,
                column_config={
                    "moc": st.column_config.NumberColumn("Mốc", width="small"),
                    "noi_dung_moc": st.column_config.TextColumn("Nội dung công việc của mốc", width="large"),
                    "thoi_han_moc": st.column_config.DateColumn("Thời hạn mốc", format="DD/MM/YYYY"),
                    "san_pham_kiem_tra": st.column_config.TextColumn("Sản phẩm kiểm tra", width="medium"),
                },
            )

            r1, r2 = st.columns(2)
            r1.text_input("Nguồn lực, hỗ trợ được bố trí (nếu có)", value="Phòng thí nghiệm Địa kỹ thuật, CSDL Bộ môn & hệ thống uSmart")
            r2.text_input("Chế độ báo cáo tiến độ (tuần / theo mốc / đột xuất)", value="Báo cáo định kỳ hàng tuần và khi hoàn thành mốc")

            st.markdown("##### II. NGHIỆM THU KẾT QUẢ (ghi khi kết thúc nhiệm vụ)")
            nt1, nt2, nt3 = st.columns(3)
            ngay_bg = nt1.date_input("Ngày hoàn thành, bàn giao sản phẩm", value=nv_p.ngay_hoan_thanh or den_ngay_ky)
            tien_do_tu_tinh = (
                "Đúng hạn"
                if (ngay_bg and thoi_han_in and ngay_bg <= thoi_han_in)
                else f"Trễ {(ngay_bg - thoi_han_in).days} ngày"
            )
            nt2.text_input("Tiến độ (tự tính)", value=tien_do_tu_tinh, disabled=True)
            cl_nt = nt3.selectbox(
                "Đánh giá chất lượng",
                options=list(DANH_MUC_CHAT_LUONG),
                index=list(DANH_MUC_CHAT_LUONG).index(nv_p.danh_gia_chat_luong) if nv_p.danh_gia_chat_luong in DANH_MUC_CHAT_LUONG else 0,
            )
            st.text_area(
                "Nhận xét của người giao việc",
                value="Hoàn thành tốt nhiệm vụ được giao, sản phẩm đảm bảo độ tin cậy kỹ thuật và đúng tiến độ của Bộ môn.",
                height=68,
            )

            sig1, sig2 = st.columns(2)
            sig1.markdown(
                f"<div style='text-align:center;'><b>NGƯỜI GIAO VIỆC</b><br><i>(Ký, ghi rõ họ tên)</i><br><br>"
                f"<b>{NGUOI_NHAN_BAO_CAO_MAC_DINH}</b></div>",
                unsafe_allow_html=True,
            )
            sig2.markdown(
                f"<div style='text-align:center;'><b>NGƯỜI NHẬN VIỆC</b><br><i>(Ký, ghi rõ họ tên)</i><br><br>"
                f"<b>{nguoi_ct_in}</b></div>",
                unsafe_allow_html=True,
            )

        with st.expander("➕ Tạo Phiếu giao việc mới & ghi trực tiếp vào Sổ giao việc Bộ môn"):
            with st.form("form_tao_phieu_03a"):
                c_gv, c_ng, c_ut = st.columns(3)
                gv_nhan = c_gv.selectbox(
                    "Người chủ trì thực hiện *",
                    options=[f"{gv.ma_can_bo} - {gv.ho_va_ten}" for gv in DANH_SACH_GIANG_VIEN_DKT],
                )
                nguon_m = c_ng.selectbox("Nguồn giao *", options=list(DANH_MUC_NGUON_GIAO), index=2)
                uu_tien_m = c_ut.selectbox("Mức ưu tiên *", options=list(DANH_MUC_UU_TIEN), index=1)
                ten_viec_m = st.text_input("Nội dung nhiệm vụ giao *")
                sp_viec_m = st.text_input("Sản phẩm đầu ra và tiêu chí đạt *")
                c_d1, c_d2, c_tt = st.columns(3)
                ngay_giao_m = c_d1.date_input("Ngày giao *", value=tu_ngay_ky)
                thoi_han_m = c_d2.date_input("Thời hạn hoàn thành *", value=den_ngay_ky)
                trong_tam_m = c_tt.checkbox("⭐ Là nhiệm vụ trọng tâm của Đơn vị (Mẫu 02a/02b)", value=True)
                if st.form_submit_button("📌 Phát hành Phiếu giao việc & Lưu vào CSDL"):
                    ma_cb_m, ten_cb_m = gv_nhan.split(" - ", 1)
                    nv_moi = NhiemVuCongViec(
                        ma_nhiem_vu="",
                        nhiem_vu=ten_viec_m,
                        nguon_giao=nguon_m,
                        so_van_ban_ma_viec=f"PGV-{int(thang):02d}/{int(nam)}",
                        san_pham_ket_qua=sp_viec_m,
                        thoi_han=thoi_han_m,
                        ngay_hoan_thanh=thoi_han_m,
                        tinh_trang="Hoàn thành đúng hạn",
                        muc_do_hoan_thanh=100.0,
                        ma_can_bo_chu_tri=ma_cb_m.strip(),
                        ten_can_bo_chu_tri=ten_cb_m.strip(),
                        ngay_giao=ngay_giao_m,
                        muc_uu_tien=uu_tien_m,
                        la_nhiem_vu_trong_tam=trong_tam_m,
                        nguon_du_lieu="Phiếu giao việc Mẫu 03a",
                    )
                    ma_sinh = db.them_nhiem_vu_moi(int(thang), int(nam), nv_moi)
                    st.success(f"Đã phát hành Phiếu giao việc `{ma_sinh}` cho giảng viên **{ten_cb_m}**!")
                    st.rerun()

    # =========================================================================
    # 3. MẪU 03C — TỔNG HỢP KẾT QUẢ THỰC HIỆN NHIỆM VỤ ĐƯỢC GIAO THEO NGƯỜI
    # =========================================================================
    else:
        with st.container(border=True):
            h_l, h_r = st.columns([4, 1])
            h_l.markdown(f"**TRƯỜNG ĐẠI HỌC GIAO THÔNG VẬN TẢI**  \n**ĐƠN VỊ: {TEN_DON_VI_CHUAN.upper()}**")
            h_r.markdown("<div style='text-align:right;'><span class='form-badge'>Mẫu 03c</span></div>", unsafe_allow_html=True)
            st.markdown(
                "<h4 style='text-align:center; color:#0f2d59; margin:4px 0;'>"
                "TỔNG HỢP KẾT QUẢ THỰC HIỆN NHIỆM VỤ ĐƯỢC GIAO THEO NGƯỜI"
                "</h4>",
                unsafe_allow_html=True,
            )
            st.caption("Tự động tổng hợp từ Sổ giao việc (Mẫu 03b) cho toàn bộ 21 giảng viên Bộ môn để đối chiếu khi lập Mẫu 02c.")

        rows_03c = []
        tong_giao = 0
        tong_den_han = 0
        tong_dung_han = 0
        tong_qua_han = 0
        tong_chua_dat = 0

        for idx, bc in enumerate(ds_all, start=1):
            so_giao = bc.tong_so_nhiem_vu
            so_dh = bc.so_nhiem_vu_da_den_han
            so_ht_dh = bc.so_nhiem_vu_hoan_thanh_dung_han
            so_qh = sum(
                1 for nv in bc.danh_sach_nhiem_vu if nv.ngay_hoan_thanh is None and nv.thoi_han and nv.thoi_han < den_ngay_ky
            )
            so_cd = sum(1 for nv in bc.danh_sach_nhiem_vu if nv.danh_gia_chat_luong == "Chưa đạt")
            tl_dh = f"{(so_ht_dh / so_dh * 100.0):.0f}%" if so_dh > 0 else "100%"

            tong_giao += so_giao
            tong_den_han += so_dh
            tong_dung_han += so_ht_dh
            tong_qua_han += so_qh
            tong_chua_dat += so_cd

            rows_03c.append(
                {
                    "TT": str(idx),
                    "Họ và tên": bc.giang_vien.ho_va_ten,
                    "Vị trí việc làm": bc.giang_vien.chuc_danh_day_du,
                    "Số nhiệm vụ được giao trong kỳ": so_giao,
                    "Số nhiệm vụ đã đến hạn": so_dh,
                    "Số nhiệm vụ hoàn thành đúng hạn": so_ht_dh,
                    "Tỷ lệ hoàn thành đúng hạn": tl_dh,
                    "Số nhiệm vụ quá hạn chưa xong": so_qh,
                    "Số nhiệm vụ chưa đạt chất lượng": so_cd,
                    "Ghi chú (chuyển sang Mẫu 02c)": f"Chuyển {so_ht_dh}/{so_dh} NV đúng hạn sang Mẫu 02c",
                }
            )

        tl_tong = f"{(tong_dung_han / tong_den_han * 100.0):.1f}%" if tong_den_han > 0 else "100.0%"
        rows_03c.append(
            {
                "TT": "Cộng",
                "Họ và tên": "CỘNG ĐƠN VỊ (21 GIẢNG VIÊN BỘ MÔN)",
                "Vị trí việc làm": TEN_DON_VI_CHUAN,
                "Số nhiệm vụ được giao trong kỳ": tong_giao,
                "Số nhiệm vụ đã đến hạn": tong_den_han,
                "Số nhiệm vụ hoàn thành đúng hạn": tong_dung_han,
                "Tỷ lệ hoàn thành đúng hạn": tl_tong,
                "Số nhiệm vụ quá hạn chưa xong": tong_qua_han,
                "Số nhiệm vụ chưa đạt chất lượng": tong_chua_dat,
                "Ghi chú (chuyển sang Mẫu 02c)": "Tổng hợp toàn đơn vị (=SUM)",
            }
        )

        st.dataframe(rows_03c, use_container_width=True, hide_index=True)


def _hien_thi_bang_tong_hop_9_chi_tieu(ds_rows: list[dict], nhan_mau: str) -> None:
    """Hiển thị bảng TỔNG HỢP 9 chỉ tiêu tự tính cho Mẫu 02a / Mẫu 02b."""
    tong_nv = len(ds_rows)
    ht_dh = sum(1 for r in ds_rows if r.get("tinh_trang") == "Hoàn thành đúng hạn")
    ht_cham = sum(1 for r in ds_rows if r.get("tinh_trang") == "Hoàn thành chậm")
    dth_dung = sum(1 for r in ds_rows if r.get("tinh_trang") == "Đang thực hiện – đúng tiến độ")
    dth_cham = sum(1 for r in ds_rows if r.get("tinh_trang") == "Đang thực hiện – chậm tiến độ")
    chua_th = sum(1 for r in ds_rows if r.get("tinh_trang") == "Chưa thực hiện")
    da_den_han = sum(int(r.get("da_den_han") or 0) for r in ds_rows)
    dung_han = sum(int(r.get("dung_han") or 0) for r in ds_rows)
    ty_le = f"{(dung_han / da_den_han * 100.0):.1f}%" if da_den_han > 0 else "100.0%"

    with st.expander(f"📊 TỔNG HỢP MỤC I — TỰ ĐỘNG TÍNH THEO {nhan_mau.upper()}", expanded=True):
        c1, c2 = st.columns([3, 2])
        c1.dataframe(
            [
                {"Chỉ tiêu tổng hợp": "Tổng số nhiệm vụ trong báo cáo", "Giá trị tự tính": str(tong_nv)},
                {"Chỉ tiêu tổng hợp": "  • Tình trạng: Hoàn thành đúng hạn", "Giá trị tự tính": str(ht_dh)},
                {"Chỉ tiêu tổng hợp": "  • Tình trạng: Hoàn thành chậm", "Giá trị tự tính": str(ht_cham)},
                {"Chỉ tiêu tổng hợp": "  • Tình trạng: Đang thực hiện – đúng tiến độ", "Giá trị tự tính": str(dth_dung)},
                {"Chỉ tiêu tổng hợp": "  • Tình trạng: Đang thực hiện – chậm tiến độ", "Giá trị tự tính": str(dth_cham)},
                {"Chỉ tiêu tổng hợp": "  • Tình trạng: Chưa thực hiện", "Giá trị tự tính": str(chua_th)},
                {"Chỉ tiêu tổng hợp": "Số nhiệm vụ đã đến hạn trong kỳ", "Giá trị tự tính": str(da_den_han)},
                {"Chỉ tiêu tổng hợp": "Số nhiệm vụ hoàn thành đúng hạn", "Giá trị tự tính": str(dung_han)},
                {"Chỉ tiêu tổng hợp": "Tỷ lệ hoàn thành đúng hạn", "Giá trị tự tính": ty_le},
            ],
            use_container_width=True,
            hide_index=True,
        )
        c2.info(
            "📌 **Quy tắc tổng hợp của Đơn vị (TB 940/TB-ĐHGTVT):**\n"
            "- Tổng hợp các nhiệm vụ trọng tâm của Bộ môn Địa kỹ thuật và Không gian ngầm.\n"
            f"- **Số nhiệm vụ trọng tâm đã đến hạn:** `{da_den_han}` nhiệm vụ.\n"
            f"- **Số nhiệm vụ hoàn thành đúng hạn:** `{dung_han}` nhiệm vụ (**{ty_le}**)."
        )


def _lay_ke_hoach_don_vi_tu_21_gv(ds_all) -> list[dict]:
    """Rút trích danh sách kế hoạch kỳ tới tiêu biểu của Đơn vị từ kế hoạch của các giảng viên."""
    rows = []
    da_chon = set()
    for bc in ds_all:
        for kh in bc.ke_hoach_thang_toi:
            ten_kh = kh.nhiem_vu.strip()
            if ten_kh and ten_kh not in da_chon:
                da_chon.add(ten_kh)
                rows.append(
                    {
                        "tt": len(rows) + 1,
                        "nhiem_vu": ten_kh,
                        "ca_nhan_chu_tri": kh.ca_nhan_chu_tri or bc.giang_vien.ho_va_ten,
                        "don_vi_phoi_hop": kh.don_vi_phoi_hop or "Khoa Công trình",
                        "thoi_han": kh.thoi_han,
                        "san_pham": kh.san_pham_du_kien,
                    }
                )
            if len(rows) >= 10:
                break
        if len(rows) >= 10:
            break
    return rows


def _hien_thi_chu_ky_truong_don_vi(ngay_ky: date) -> None:
    """Hiển thị chân chữ ký của Trưởng đơn vị cho Mẫu 02a / Mẫu 02b."""
    ngay_str = ngay_ky.strftime("Hà Nội, ngày %d tháng %m năm %Y")
    _, col_r = st.columns([3, 2])
    col_r.markdown(
        f"<div style='text-align:center; margin-top:8px;'><i>{ngay_str}</i><br>"
        f"<b>TRƯỞNG ĐƠN VỊ</b><br><i>(Ký, ghi rõ họ tên)</i><br><br>"
        f"<b>{NGUOI_NHAN_BAO_CAO_MAC_DINH}</b></div>",
        unsafe_allow_html=True,
    )

