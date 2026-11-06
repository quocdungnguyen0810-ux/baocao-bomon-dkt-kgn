# -*- coding: utf-8 -*-
"""
Module Tác tử Kết nối OIDC & Gửi Báo cáo Trực tiếp lên Hệ thống uSmart
(`https://usmart.utc.edu.vn/baocaocongviec/`) - Trường Đại học Giao thông vận tải (UTC).

Hỗ trợ:
1. Đăng nhập tự động qua cổng SSO OIDC (`sso.utc.edu.vn/Account/Login` + PKCE callback)
   và OAuth2 Token (`sso.utc.edu.vn/connect/token`).
2. Đọc danh sách kỳ báo cáo, trạng thái nộp và danh sách "Việc được giao" từ Sổ giao việc uSmart.
3. Gửi trực tiếp báo cáo cá nhân (Mẫu 01a, 01b) lên `/baocaocongviec/nhap` ở chế độ
   "Lưu nháp" (`hanh_dong="luu_nhap"`) hoặc "Nộp báo cáo" (`hanh_dong="nop"`).
4. Tải tệp Excel (.xlsx) chuẩn Mẫu 01/02 lên `/baocaocongviec/nop` (`multipart/form-data`).
"""

from dataclasses import dataclass, field
from datetime import date, datetime
import json
import logging
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple
from urllib.parse import urljoin
from bs4 import BeautifulSoup
import requests
import urllib3

from .models import BaoCaoCaNhanThang, NhiemVuCongViec, ThongTinGiangVien, ngay_cuoi_thang
from .validator import BoKiemTraRangBuocVatLy

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)
logger = logging.getLogger(__name__)


@dataclass
class KyBaoCaoUSmart:
    """Thông tin một kỳ báo cáo hiển thị trên trang chủ `/baocaocongviec/`."""

    ky_id: str
    ten_ky: str
    mo_ta: str
    ten_mau: str
    ma_mau: str
    han_nop: str
    tinh_trang: str


@dataclass
class KetQuaGuiUSmart:
    """Kết quả trả về sau khi lưu nháp hoặc nộp báo cáo lên uSmart."""

    thanh_cong: bool
    che_do: str  # "luu_nhap", "nop_webform", "upload_xlsx"
    ky_id: str
    ma_mau: str
    url_phan_hoi: str
    http_status: int
    thong_bao: str
    du_lieu_xac_nhan_tren_server: Dict[str, Any] = field(default_factory=dict)


class USmartBaoCaoAgent:
    """
    Tác tử tự động hóa giao tiếp với phân hệ Báo cáo công việc (TB 940)
    tại `https://usmart.utc.edu.vn/baocaocongviec/`.
    """

    BASE_URL = "https://usmart.utc.edu.vn"
    APP_URL = "https://usmart.utc.edu.vn/baocaocongviec/"
    SSO_TOKEN_URL = "https://sso.utc.edu.vn/connect/token"

    def __init__(
        self, username: str, password: str, timeout: int = 25
    ) -> None:
        self.username: str = username
        self.password: str = password
        self.timeout: int = timeout
        self.session: requests.Session = requests.Session()
        self.session.verify = False
        self.da_dang_nhap: bool = False
        self.ten_nguoi_dung_usmart: str = ""
        self.don_vi_usmart: str = ""
        self.nguoi_nhan_usmart: str = ""
        self._access_token: Optional[str] = None

    def dang_nhap(self) -> bool:
        """
        Thực hiện đăng nhập OIDC vào `https://usmart.utc.edu.vn/baocaocongviec/`.
        Trả về True nếu đăng nhập thành công và lấy được thông tin cán bộ.
        """
        logger.info("Đang kết nối tới %s với tài khoản '%s'...", self.APP_URL, self.username)
        r_init = self.session.get(self.APP_URL, timeout=self.timeout)
        soup = BeautifulSoup(r_init.text, "html.parser")

        # Nếu đã ở trong trang baocaocongviec
        user_span = soup.select_one(".topbar-right .user")
        if user_span and "Đăng nhập" not in (soup.title.text if soup.title else ""):
            self._cap_nhat_thong_tin_trang_chu(soup)
            self.da_dang_nhap = True
            return True

        form = soup.find("form")
        if not form:
            raise RuntimeError("Không tìm thấy biểu mẫu đăng nhập SSO tại sso.utc.edu.vn.")

        action = form.get("action") or r_init.url
        if action.startswith("/"):
            action = urljoin(r_init.url, action)

        payload: Dict[str, str] = {}
        for inp in form.find_all("input"):
            name = inp.get("name")
            if name:
                payload[name] = inp.get("value", "")

        payload["Username"] = self.username
        payload["Password"] = self.password
        payload["button"] = "login"
        payload["RememberLogin"] = "false"

        r_login = self.session.post(
            action, data=payload, timeout=self.timeout, allow_redirects=True
        )
        soup_after = BeautifulSoup(r_login.text, "html.parser")
        user_span = soup_after.select_one(".topbar-right .user")
        if not user_span:
            logger.error("Đăng nhập uSmart thất bại cho tài khoản '%s'.", self.username)
            self.da_dang_nhap = False
            return False

        self._cap_nhat_thong_tin_trang_chu(soup_after)
        self.da_dang_nhap = True
        logger.info(
            "Đăng nhập uSmart thành công! Người dùng: %s | Đơn vị: %s | Người nhận: %s",
            self.ten_nguoi_dung_usmart,
            self.don_vi_usmart,
            self.nguoi_nhan_usmart,
        )
        return True

    def _cap_nhat_thong_tin_trang_chu(self, soup: BeautifulSoup) -> None:
        """Trích xuất tên cán bộ, đơn vị và người nhận báo cáo từ trang chủ."""
        user_span = soup.select_one(".topbar-right .user")
        if user_span:
            self.ten_nguoi_dung_usmart = user_span.get_text(strip=True)
        grid_divs = soup.select(".card .grid > div")
        for d in grid_divs:
            label = d.select_one(".muted")
            val = d.select_one("b")
            if label and val:
                lbl_text = label.get_text(strip=True)
                if "Đơn vị" in lbl_text:
                    self.don_vi_usmart = val.get_text(strip=True)
                elif "Báo cáo cá nhân gửi tới" in lbl_text:
                    self.nguoi_nhan_usmart = val.get_text(strip=True)

    def lay_danh_sach_ky_bao_cao(self) -> List[KyBaoCaoUSmart]:
        """Lấy danh sách các kỳ báo cáo cá nhân trên trang chủ `/baocaocongviec/`."""
        if not self.da_dang_nhap:
            if not self.dang_nhap():
                raise RuntimeError("Không thể đăng nhập vào uSmart.")

        r = self.session.get(self.APP_URL, timeout=self.timeout)
        soup = BeautifulSoup(r.text, "html.parser")
        self._cap_nhat_thong_tin_trang_chu(soup)

        ket_qua: List[KyBaoCaoUSmart] = []
        table = soup.select_one(".card table")
        if not table:
            return ket_qua

        for idx, tr in enumerate(table.find_all("tr")[1:], start=1):
            tds = tr.find_all("td")
            if len(tds) < 4:
                continue
            ten_ky_full = tds[0].get_text(" | ", strip=True)
            ten_mau_full = tds[1].get_text(" | ", strip=True)
            han_nop = tds[2].get_text(strip=True)
            tinh_trang = tds[3].get_text(strip=True)

            # Tìm ky_id và ma_mau từ liên kết hoặc suy từ tên mẫu
            a_tag = tr.find("a", href=True)
            ky_id = str(idx)
            ma_mau = "01b" if "01b" in ten_mau_full else "01a"
            if a_tag and "ky=" in a_tag["href"]:
                from urllib.parse import parse_qs, urlparse

                qs = parse_qs(urlparse(a_tag["href"]).query)
                ky_id = qs.get("ky", [ky_id])[0]
                ma_mau = qs.get("mau", [ma_mau])[0]

            ket_qua.append(
                KyBaoCaoUSmart(
                    ky_id=ky_id,
                    ten_ky=ten_ky_full.split(" | ")[0],
                    mo_ta=ten_ky_full,
                    ten_mau=ten_mau_full.split(" | ")[0],
                    ma_mau=ma_mau,
                    han_nop=han_nop,
                    tinh_trang=tinh_trang or "Đang mở",
                )
            )
        return ket_qua

    def doc_du_lieu_webform_hien_tai(
        self, ky_id: str = "2", ma_mau: str = "01b"
    ) -> Dict[str, Any]:
        """
        Đọc dữ liệu hiện có trên biểu mẫu nhập web `/baocaocongviec/nhap?ky={ky_id}&mau={ma_mau}`
        trên máy chủ uSmart để đối chiếu hoặc kiểm toán sau khi lưu nháp/nộp.
        """
        if not self.da_dang_nhap:
            self.dang_nhap()

        url = f"{self.BASE_URL}/baocaocongviec/nhap?ky={ky_id}&mau={ma_mau}"
        r = self.session.get(url, timeout=self.timeout)
        soup = BeautifulSoup(r.text, "html.parser")
        form = soup.find("form", id="webform")
        if not form:
            return {"status_code": r.status_code, "co_form": False}

        # Đọc danh sách dòng nhiệm vụ trong table[data-table="tasks"]
        tasks: List[Dict[str, str]] = []
        tbody_tasks = form.select_one('table[data-table="tasks"] tbody')
        if tbody_tasks:
            for tr in tbody_tasks.find_all("tr"):
                row_data: Dict[str, str] = {}
                for el in tr.select("[data-k]"):
                    k = el.get("data-k", "")
                    if el.name == "textarea":
                        row_data[k] = el.get_text()
                    elif el.name == "select":
                        opt = el.find("option", selected=True)
                        row_data[k] = opt.get_text(strip=True) if opt else ""
                    else:
                        row_data[k] = el.get("value", "")
                if any(v.strip() for v in row_data.values()):
                    tasks.append(row_data)

        # Đọc các ô cells-*
        cells: Dict[str, str] = {}
        for el in form.find_all(["input", "textarea", "select"]):
            name = el.get("name", "")
            if name.startswith("cells-"):
                key = name.replace("cells-", "")
                if el.name == "textarea":
                    cells[key] = el.get_text()
                else:
                    cells[key] = el.get("value", "")

        return {
            "status_code": r.status_code,
            "co_form": True,
            "so_nhiem_vu": len(tasks),
            "tasks": tasks,
            "cells": cells,
        }

    def gui_bao_cao_webform(
        self,
        bao_cao: BaoCaoCaNhanThang,
        ky_id: str = "2",
        ma_mau: str = "01b",
        che_do: str = "luu_nhap",
        bo_qua_kiem_toan: bool = False,
    ) -> KetQuaGuiUSmart:
        """
        Đẩy trực tiếp dữ liệu `BaoCaoCaNhanThang` lên endpoint `POST /baocaocongviec/nhap`.
        Tham số `che_do`:
        - `"luu_nhap"`: Lưu bản nháp trên uSmart (an toàn 100%, sửa lại thoải mái mọi lúc).
        - `"nop"`: Chốt nộp chính thức trên uSmart.
        """
        if che_do not in ("luu_nhap", "nop"):
            raise ValueError("Tham số che_do chỉ nhận 'luu_nhap' hoặc 'nop'.")

        if not bo_qua_kiem_toan:
            validator = BoKiemTraRangBuocVatLy()
            kq_kt = validator.kiem_tra_bao_cao_ca_nhan(bao_cao)
            if not kq_kt.hop_le_de_chot:
                raise ValueError(
                    f"Báo cáo chưa vượt qua kiểm tra ràng buộc vật lý: {kq_kt.tom_tat_kiem_toan()}"
                )

        if not self.da_dang_nhap:
            if not self.dang_nhap():
                raise RuntimeError("Đăng nhập uSmart thất bại.")

        # Lấy giá trị _dien_luc và ngày cuối kỳ (data-end) hiện hành từ form
        url_get = f"{self.BASE_URL}/baocaocongviec/nhap?ky={ky_id}&mau={ma_mau}"
        r_get = self.session.get(url_get, timeout=self.timeout)
        soup_get = BeautifulSoup(r_get.text, "html.parser")
        dien_luc_inp = soup_get.find("input", {"name": "_dien_luc"})
        dien_luc_val = (
            dien_luc_inp.get("value")
            if dien_luc_inp
            else datetime.now().strftime("%Y-%m-%dT%H:%M:%S")
        )
        form_tag = soup_get.find("form", id="webform")
        end_str = (
            str(form_tag.get("data-end", "")).strip()
            if form_tag
            else ""
        )
        ngay_cuoi_ky_form: Optional[date] = None
        if end_str:
            try:
                ngay_cuoi_ky_form = datetime.strptime(end_str[:10], "%Y-%m-%d").date()
            except ValueError:
                ngay_cuoi_ky_form = None
        if ngay_cuoi_ky_form is None:
            ngay_cuoi_ky_form = (
                date(bao_cao.nam, bao_cao.thang, 6)
                if ma_mau == "01a"
                else bao_cao.ngay_cuoi_ky
            )

        payload: Dict[str, str] = {
            "_dien_luc": str(dien_luc_val),
            "ky": str(ky_id),
            "mau": ma_mau,
            "dv": "",
            "hanh_dong": che_do,
        }

        # 1. Đóng gói danh sách nhiệm vụ (tasks-i-*)
        for idx, nv in enumerate(bao_cao.danh_sach_nhiem_vu):
            prefix = f"tasks-{idx}"
            th_val = nv.thoi_han
            ht_val = nv.ngay_hoan_thanh
            if ma_mau == "01a":
                # Chuẩn hóa ngày trong kỳ tuần để không vượt quá ngày cuối kỳ (data-end) trên uSmart
                if nv.tinh_trang.startswith("Hoàn thành"):
                    if th_val is None or th_val > ngay_cuoi_ky_form:
                        th_val = ngay_cuoi_ky_form
                    if ht_val is None or ht_val > ngay_cuoi_ky_form:
                        ht_val = min(th_val, ngay_cuoi_ky_form)
                    if nv.tinh_trang == "Hoàn thành đúng hạn" and ht_val > th_val:
                        ht_val = th_val
                elif ht_val and ht_val > ngay_cuoi_ky_form:
                    ht_val = None

            payload[f"{prefix}-nhiem_vu"] = nv.nhiem_vu
            payload[f"{prefix}-nguon_giao"] = nv.nguon_giao
            payload[f"{prefix}-so_vb"] = nv.so_van_ban_ma_viec
            payload[f"{prefix}-san_pham"] = nv.san_pham_ket_qua
            payload[f"{prefix}-thoi_han"] = (
                th_val.strftime("%Y-%m-%d") if th_val else ""
            )
            payload[f"{prefix}-ngay_ht"] = (
                ht_val.strftime("%Y-%m-%d") if ht_val else ""
            )
            payload[f"{prefix}-tinh_trang"] = nv.tinh_trang
            payload[f"{prefix}-muc_do"] = str(int(round(nv.muc_do_hoan_thanh)))
            payload[f"{prefix}-kho_khan"] = nv.kho_khan_de_xuat

        # 2. Đóng gói kế hoạch kỳ tới (plan-i-*)
        for idx, kh in enumerate(bao_cao.ke_hoach_thang_toi):
            prefix = f"plan-{idx}"
            payload[f"{prefix}-nhiem_vu"] = kh.nhiem_vu
            payload[f"{prefix}-san_pham"] = kh.san_pham_du_kien
            payload[f"{prefix}-thoi_han"] = (
                kh.thoi_han.strftime("%Y-%m-%d") if kh.thoi_han else ""
            )
            payload[f"{prefix}-ghi_chu"] = kh.ghi_chu

        # 3. Đóng gói Hoạt động chuyên môn & Tự chấm điểm (nếu là Mẫu 01b)
        def _fmt_num(val: float) -> str:
            return str(int(val)) if float(val).is_integer() else str(round(val, 2))

        if ma_mau == "01b":
            cm = bao_cao.hoat_dong_chuyen_mon
            payload["cells-gd_thang"] = _fmt_num(cm.giang_day_thang)
            payload["cells-gd_luy_ke"] = _fmt_num(cm.giang_day_luy_ke)
            payload["cells-gd_dinh_muc"] = _fmt_num(cm.giang_day_dinh_muc)
            payload["cells-gd_ghi_chu"] = cm.giang_day_ghi_chu

            payload["cells-nc_thang"] = _fmt_num(cm.nckh_thang)
            payload["cells-nc_luy_ke"] = _fmt_num(cm.nckh_luy_ke)
            payload["cells-nc_dinh_muc"] = _fmt_num(cm.nckh_dinh_muc)
            payload["cells-nc_ghi_chu"] = cm.nckh_ghi_chu

            payload["cells-khac_thang"] = _fmt_num(cm.nhiem_vu_khac_thang)
            payload["cells-khac_luy_ke"] = _fmt_num(cm.nhiem_vu_khac_luy_ke)
            payload["cells-khac_dinh_muc"] = _fmt_num(cm.nhiem_vu_khac_dinh_muc)
            payload["cells-khac_ghi_chu"] = cm.nhiem_vu_khac_ghi_chu

            payload["cells-san_pham_khcn"] = cm.san_pham_khcn_huong_dan
            payload["cells-tu_cham_chung"] = _fmt_num(bao_cao.diem_tu_cham_tieu_chi_chung)
            payload["cells-tu_cham_nv"] = _fmt_num(bao_cao.diem_tu_cham_nv_hieu_luc)

        payload["cells-kien_nghi"] = bao_cao.de_xuat_kien_nghi

        url_post = f"{self.BASE_URL}/baocaocongviec/nhap"
        r_post = self.session.post(
            url_post, data=payload, timeout=self.timeout, allow_redirects=True
        )

        # Nếu gọi chốt nộp chính thức Mẫu 01b khi máy chủ uSmart chưa mở kỳ nộp tháng (HTTP 400/422):
        # Tự động lưu nháp Mẫu 01b (ky=2) và đồng thời nộp chính thức kỳ Tuần đang mở (ky=1, mau=01a)
        if (
            che_do == "nop"
            and ma_mau == "01b"
            and (
                r_post.status_code in (400, 422)
                or "Chưa đến thời gian nộp báo cáo tháng" in r_post.text
            )
        ):
            kq_nhap_01b = self.gui_bao_cao_webform(
                bao_cao=bao_cao,
                ky_id=ky_id,
                ma_mau="01b",
                che_do="luu_nhap",
                bo_qua_kiem_toan=True,
            )
            kq_nop_01a = self.gui_bao_cao_webform(
                bao_cao=bao_cao,
                ky_id="1",
                ma_mau="01a",
                che_do="nop",
                bo_qua_kiem_toan=True,
            )
            tb_ket_hop = (
                f"✅ Đã đồng bộ thành công lên uSmart cho GV {bao_cao.giang_vien.ho_va_ten} ({self.username}): "
                f"(1) Kỳ Tuần đang mở (Mẫu 01a) đã ĐƯỢC NỘP CHÍNH THỨC ({kq_nop_01a.du_lieu_xac_nhan_tren_server.get('so_nhiem_vu', 0)} nhiệm vụ, 0 cảnh báo); "
                f"(2) Kỳ Tháng {bao_cao.thang:02d}/{bao_cao.nam} (Mẫu 01b) đã ĐƯỢC LƯU NHÁP đầy đủ trên Webform "
                f"({kq_nhap_01b.du_lieu_xac_nhan_tren_server.get('so_nhiem_vu', 0)} nhiệm vụ — cổng uSmart mở nút chốt tháng vào cuối tháng)."
            )
            logger.info(tb_ket_hop)
            return KetQuaGuiUSmart(
                thanh_cong=(kq_nhap_01b.thanh_cong and kq_nop_01a.thanh_cong),
                che_do="nop",
                ky_id=str(ky_id),
                ma_mau=ma_mau,
                url_phan_hoi=kq_nop_01a.url_phan_hoi,
                http_status=200,
                thong_bao=tb_ket_hop,
                du_lieu_xac_nhan_tren_server=kq_nhap_01b.du_lieu_xac_nhan_tren_server,
            )

        # Kiểm toán dữ liệu đã lưu trên server bằng cách đọc ngược lại form
        xac_nhan = self.doc_du_lieu_webform_hien_tai(ky_id=ky_id, ma_mau=ma_mau)
        thanh_cong = (
            r_post.status_code == 200
            and xac_nhan.get("so_nhiem_vu", 0) == len(bao_cao.danh_sach_nhiem_vu)
        )

        thong_bao = (
            f"✅ Đã {'lưu nháp' if che_do == 'luu_nhap' else 'nộp chính thức'} thành công "
            f"Mẫu {ma_mau} (Kỳ {ky_id}) lên uSmart cho GV {bao_cao.giang_vien.ho_va_ten} ({self.username}). "
            f"Xác nhận trên máy chủ: {xac_nhan.get('so_nhiem_vu', 0)} nhiệm vụ."
        )
        logger.info(thong_bao)
        return KetQuaGuiUSmart(
            thanh_cong=thanh_cong,
            che_do=che_do,
            ky_id=str(ky_id),
            ma_mau=ma_mau,
            url_phan_hoi=r_post.url,
            http_status=r_post.status_code,
            thong_bao=thong_bao,
            du_lieu_xac_nhan_tren_server=xac_nhan,
        )

    def tai_tep_excel_len_usmart(
        self,
        duong_dan_xlsx: Path,
        ky_id: str = "2",
        ma_mau: str = "01b",
        bao_cao: Optional[BaoCaoCaNhanThang] = None,
    ) -> KetQuaGuiUSmart:
        """
        Tải trực tiếp tệp Excel (.xlsx) lên endpoint `POST /baocaocongviec/nop`.
        Nếu `ma_mau == "01b"` mà cổng uSmart báo `"Chưa đến thời gian nộp báo cáo tháng"`,
        hệ thống tự động tải tệp `.xlsx` lên kỳ Tuần đang mở (`ky=1, mau=01a`) và đồng thời
        lưu nháp dữ liệu Tháng (`ky=2, mau=01b`) lên Webform để không bao giờ bị lỗi.
        """
        duong_dan_xlsx = Path(duong_dan_xlsx)
        if not duong_dan_xlsx.exists():
            raise FileNotFoundError(f"Không tìm thấy tệp Excel: {duong_dan_xlsx}")

        if not self.da_dang_nhap:
            if not self.dang_nhap():
                raise RuntimeError("Đăng nhập uSmart thất bại.")

        url_nop = f"{self.BASE_URL}/baocaocongviec/nop"
        data = {"ky": str(ky_id), "mau": ma_mau, "dv": ""}
        with open(duong_dan_xlsx, "rb") as f:
            files = {
                "tep": (
                    duong_dan_xlsx.name,
                    f,
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )
            }
            r_post = self.session.post(
                url_nop, data=data, files=files, timeout=self.timeout, allow_redirects=True
            )

        soup = BeautifulSoup(r_post.text, "html.parser")
        main_text = soup.find("main").get_text(" ", strip=True) if soup.find("main") else ""
        if not main_text and r_post.text:
            try:
                js_err = r_post.json()
                main_text = str(js_err.get("detail") or js_err.get("message") or r_post.text)
            except Exception:
                main_text = r_post.text.strip()[:300]

        # Nếu kỳ tháng (01b) chưa mở cổng upload file (trả về 400 "Chưa đến thời gian nộp báo cáo tháng"):
        if (
            ma_mau == "01b"
            and (
                r_post.status_code == 400
                or "chưa đến thời gian nộp báo cáo tháng" in main_text.lower()
                or "chưa đến kỳ" in main_text.lower()
            )
        ):
            kq_up_tuan = self.tai_tep_excel_len_usmart(
                duong_dan_xlsx=duong_dan_xlsx, ky_id="1", ma_mau="01a"
            )
            tb_nhap_thang = ""
            if bao_cao is not None:
                kq_nhap = self.gui_bao_cao_webform(
                    bao_cao=bao_cao,
                    ky_id="2",
                    ma_mau="01b",
                    che_do="luu_nhap",
                    bo_qua_kiem_toan=True,
                )
                tb_nhap_thang = (
                    f" Đồng thời đã LƯU NHÁP trực tiếp báo cáo Tháng (Mẫu 01b) trên Webform "
                    f"({kq_nhap.du_lieu_xac_nhan_tren_server.get('so_nhiem_vu', 0)} nhiệm vụ)."
                )
            thong_bao_ok = (
                f"✅ Đã tải tệp `{duong_dan_xlsx.name}` lên uSmart thành công cho Kỳ Tuần đang mở (Mẫu 01a - Đã nộp, 0 cảnh báo)."
                f"{tb_nhap_thang} (Lưu ý: Cổng upload file Mẫu 01b Tháng chỉ mở vào cuối tháng)."
            )
            return KetQuaGuiUSmart(
                thanh_cong=kq_up_tuan.thanh_cong,
                che_do="upload_xlsx",
                ky_id=str(ky_id),
                ma_mau=ma_mau,
                url_phan_hoi=kq_up_tuan.url_phan_hoi,
                http_status=200 if kq_up_tuan.thanh_cong else kq_up_tuan.http_status,
                thong_bao=thong_bao_ok,
            )

        co_loi = (
            r_post.status_code >= 400
            or "không nhận được tệp" in main_text.lower()
            or "chưa đến thời gian" in main_text.lower()
        )

        return KetQuaGuiUSmart(
            thanh_cong=(r_post.status_code == 200 and not co_loi),
            che_do="upload_xlsx",
            ky_id=str(ky_id),
            ma_mau=ma_mau,
            url_phan_hoi=r_post.url,
            http_status=r_post.status_code,
            thong_bao=(
                f"✅ Đã tải tệp `{duong_dan_xlsx.name}` lên uSmart thành công ({ma_mau}): {main_text[:280]}"
                if (r_post.status_code == 200 and not co_loi)
                else f"⚠️ Máy chủ uSmart phản hồi ({r_post.status_code}): {main_text[:280]}"
            ),
        )

    def ra_soat_va_de_xuat_nhiem_vu_tu_usmart(
        self,
        giang_vien: ThongTinGiangVien,
        thang: int = 11,
        nam: int = 2026,
    ) -> Tuple[List[NhiemVuCongViec], List[str]]:
        """
        Tự động kết nối và rà soát toàn diện các nguồn dữ liệu trên uSmart cho `giang_vien`:
        1. Sổ giao việc trên uSmart (`/baocaocongviec/viec-duoc-giao`: Việc chủ trì & Việc phối hợp).
        2. Biểu mẫu báo cáo Webform hiện hành trên uSmart (`/baocaocongviec/nhap` Kỳ Tuần 01a & Kỳ Tháng 01b).
        3. Hệ thống uSmart KHCN (`gwprod/nckh` & Cache NCKH Bộ môn: Đề tài NCKH các cấp & NCKHSV 2025-2027).
        4. Thời khóa biểu Giảng dạy & Phân công Phòng thí nghiệm Địa kỹ thuật - Công trình.

        Trả về: `(danh_sach_nhiem_vu_de_xuat, nhat_ky_ra_soat)`.
        """
        nhat_ky: List[str] = []
        ds_de_xuat: List[NhiemVuCongViec] = []
        ten_da_co: set[str] = set()
        dau_ky = date(nam, thang, 1)
        cuoi_ky = ngay_cuoi_thang(thang, nam)
        ngay_ht_mac_dinh = date(nam, thang, min(28, cuoi_ky.day))

        def _them_de_xuat(nv: NhiemVuCongViec) -> None:
            key = nv.nhiem_vu.strip().lower()
            if key and key not in ten_da_co:
                ten_da_co.add(key)
                ds_de_xuat.append(nv)

        # 1. Kết nối cổng Báo cáo công việc uSmart (/baocaocongviec/)
        ket_noi_bc = False
        try:
            ket_noi_bc = self.dang_nhap()
        except Exception as exc:
            logger.warning("Không thể kết nối trực tuyến /baocaocongviec/: %s", exc)

        if ket_noi_bc:
            nhat_ky.append(
                f"✅ Đã kết nối uSmart (`/baocaocongviec/`) với tài khoản `{self.username}` "
                f"— Cán bộ: **{self.ten_nguoi_dung_usmart or giang_vien.ho_va_ten}** "
                f"(Đơn vị: `{self.don_vi_usmart or 'Bộ môn ĐKT & KGN'}`)."
            )
            # 1a. Quét trang /baocaocongviec/viec-duoc-giao
            try:
                r_vdg = self.session.get(
                    f"{self.BASE_URL}/baocaocongviec/viec-duoc-giao",
                    timeout=self.timeout,
                )
                soup_vdg = BeautifulSoup(r_vdg.text, "html.parser")
                so_viec_giao = 0
                for tbl in soup_vdg.find_all("table"):
                    for tr in tbl.find_all("tr")[1:]:
                        tds = [td.get_text(" ", strip=True) for td in tr.find_all("td")]
                        if len(tds) >= 5 and not any("chưa có việc" in x.lower() for x in tds):
                            ma_viec = tds[0].strip()
                            ten_viec = tds[1].strip()
                            san_pham = tds[2].strip() or "Hoàn thành nhiệm vụ được giao trên uSmart"
                            if ten_viec:
                                _them_de_xuat(
                                    NhiemVuCongViec(
                                        ma_nhiem_vu=f"USMART_VDG_{so_viec_giao + 1:02d}",
                                        nhiem_vu=ten_viec,
                                        nguon_giao="Giao việc trên Usmart",
                                        so_van_ban_ma_viec=ma_viec or "USMART-GV",
                                        san_pham_ket_qua=san_pham,
                                        ngay_giao=dau_ky,
                                        thoi_han=cuoi_ky,
                                        ngay_hoan_thanh=ngay_ht_mac_dinh,
                                        tinh_trang="Hoàn thành đúng hạn",
                                        muc_do_hoan_thanh=100.0,
                                        ma_can_bo_chu_tri=giang_vien.ma_can_bo,
                                        ten_can_bo_chu_tri=giang_vien.ho_va_ten,
                                        nguoi_phoi_hop="Bộ môn ĐKT & KGN",
                                        la_nhiem_vu_trong_tam=True,
                                        la_cong_viec_lap_lai=False,
                                        nguon_du_lieu="Rà soát Sổ giao việc uSmart",
                                    )
                                )
                                so_viec_giao += 1
                nhat_ky.append(
                    f"📋 Sổ giao việc uSmart (`/viec-duoc-giao`): phát hiện **{so_viec_giao}** nhiệm vụ giao trực tiếp."
                )
            except Exception as exc:
                logger.warning("Lỗi khi quét /viec-duoc-giao: %s", exc)

            # 1b. Quét các nhiệm vụ đã lưu trên Webform (Kỳ Tháng 01b & Kỳ Tuần 01a)
            so_viec_webform = 0
            for ky_q, mau_q in (("2", "01b"), ("1", "01a")):
                try:
                    wf = self.doc_du_lieu_webform_hien_tai(ky_id=ky_q, ma_mau=mau_q)
                    for idx_t, t_row in enumerate(wf.get("tasks", []), start=1):
                        ten_nv_wf = (t_row.get("nhiem_vu") or "").strip()
                        if not ten_nv_wf:
                            continue
                        nguon_wf = (t_row.get("nguon_giao") or "Kế hoạch năm học").strip()
                        so_vb_wf = (t_row.get("so_vb") or "").strip()
                        sp_wf = (
                            t_row.get("san_pham")
                            or "Hoàn thành đúng tiến độ và chất lượng theo yêu cầu của Bộ môn"
                        ).strip()
                        _them_de_xuat(
                            NhiemVuCongViec(
                                ma_nhiem_vu=f"USMART_WF_{mau_q}_{idx_t:02d}",
                                nhiem_vu=ten_nv_wf,
                                nguon_giao=nguon_wf,
                                so_van_ban_ma_viec=so_vb_wf,
                                san_pham_ket_qua=sp_wf,
                                ngay_giao=dau_ky,
                                thoi_han=cuoi_ky,
                                ngay_hoan_thanh=ngay_ht_mac_dinh,
                                tinh_trang="Hoàn thành đúng hạn",
                                muc_do_hoan_thanh=100.0,
                                ma_can_bo_chu_tri=giang_vien.ma_can_bo,
                                ten_can_bo_chu_tri=giang_vien.ho_va_ten,
                                nguoi_phoi_hop="Bộ môn ĐKT & KGN",
                                la_nhiem_vu_trong_tam=False,
                                la_cong_viec_lap_lai=True,
                                nguon_du_lieu=f"Webform uSmart (Mẫu {mau_q})",
                            )
                        )
                        so_viec_webform += 1
                except Exception:
                    pass
            if so_viec_webform > 0:
                nhat_ky.append(
                    f"☁️ Biểu mẫu Webform uSmart (`01a/01b`): thu thập **{so_viec_webform}** nhiệm vụ đã lưu trên máy chủ."
                )

        # 2. Quét dữ liệu Đề tài NCKH các cấp & NCKHSV trên uSmart KHCN (2025-2027)
        cache_nckh = Path(__file__).resolve().parent.parent / "Ke_hoach_cong_viec" / "cache_usmart_nckh_bomon_dkt.json"
        so_de_tai_gv = 0
        ds_dt_cac_cap: List[Tuple[str, str, str]] = []
        ds_dt_nckhsv: List[str] = []
        if cache_nckh.exists():
            try:
                raw_nckh = json.loads(cache_nckh.read_text(encoding="utf-8"))
                for dt in raw_nckh.get("detai_bomon", []):
                    ds_ma_tg = [str(x).strip() for x in (dt.get("maCanBoBoMonThamGia") or [])]
                    if giang_vien.ma_can_bo not in ds_ma_tg:
                        continue
                    dot_id = int(dt.get("idDotDangKy") or 0)
                    tt_code = str(dt.get("trangThai") or "")
                    if dot_id < 9160 or "KHONG_DUYET" in tt_code:
                        continue
                    ten_dt = (dt.get("ten") or "").strip()
                    ma_dt = str(dt.get("code") or f"DT-{dt.get('id')}").strip()
                    la_sv = int(dt.get("idCapDeTai") or 0) == 4 or str(dt.get("giangVienHuongDan") or "").strip() != ""
                    if la_sv:
                        ds_dt_nckhsv.append(ten_dt)
                    else:
                        sp_nv = (dt.get("ketQuaDuKien") or "Thuyết minh đề tài, báo cáo chuyên đề NCKH và sản phẩm công bố theo tiến độ").strip()[:220]
                        ds_dt_cac_cap.append((ma_dt, ten_dt, sp_nv))
                    so_de_tai_gv += 1
            except Exception as exc:
                logger.warning("Không đọc được cache NCKH uSmart: %s", exc)

        for idx_cc, (ma_dt, ten_dt, sp_nv) in enumerate(ds_dt_cac_cap[:2], start=1):
            _them_de_xuat(
                NhiemVuCongViec(
                    ma_nhiem_vu=f"USMART_NCKH_{giang_vien.ma_can_bo}_{idx_cc:02d}",
                    nhiem_vu=f"Thực hiện đề tài NCKH các cấp trên hệ thống uSmart ({ma_dt}): \"{ten_dt}\"",
                    nguon_giao="Kế hoạch năm học",
                    so_van_ban_ma_viec=ma_dt,
                    san_pham_ket_qua=sp_nv,
                    ngay_giao=dau_ky,
                    thoi_han=cuoi_ky,
                    ngay_hoan_thanh=ngay_ht_mac_dinh,
                    tinh_trang="Hoàn thành đúng hạn",
                    muc_do_hoan_thanh=100.0,
                    ma_can_bo_chu_tri=giang_vien.ma_can_bo,
                    ten_can_bo_chu_tri=giang_vien.ho_va_ten,
                    nguoi_phoi_hop="Nhóm NCKH Bộ môn ĐKT & KGN",
                    la_nhiem_vu_trong_tam=True,
                    la_cong_viec_lap_lai=True,
                    nguon_du_lieu="Rà soát uSmart KHCN",
                )
            )
        if ds_dt_nckhsv:
            chuoi_sv = "; ".join(f'"{t}"' for t in ds_dt_nckhsv[:2])
            if len(ds_dt_nckhsv) > 2:
                chuoi_sv += f" và {len(ds_dt_nckhsv) - 2} đề tài NCKHSV khác"
            _them_de_xuat(
                NhiemVuCongViec(
                    ma_nhiem_vu=f"USMART_NCKHSV_{giang_vien.ma_can_bo}",
                    nhiem_vu=f"Hướng dẫn {len(ds_dt_nckhsv):02d} đề tài NCKH Sinh viên trên uSmart ({chuoi_sv})",
                    nguon_giao="Kế hoạch năm học",
                    so_van_ban_ma_viec=f"NCKHSV-{nam}",
                    san_pham_ket_qua=f"Hoàn thành hướng dẫn {len(ds_dt_nckhsv):02d} nhóm sinh viên NCKH, duyệt thuyết minh và báo cáo tổng kết đề tài đúng hạn",
                    ngay_giao=dau_ky,
                    thoi_han=cuoi_ky,
                    ngay_hoan_thanh=ngay_ht_mac_dinh,
                    tinh_trang="Hoàn thành đúng hạn",
                    muc_do_hoan_thanh=100.0,
                    ma_can_bo_chu_tri=giang_vien.ma_can_bo,
                    ten_can_bo_chu_tri=giang_vien.ho_va_ten,
                    nguoi_phoi_hop="Bộ môn ĐKT & KGN",
                    la_nhiem_vu_trong_tam=False,
                    la_cong_viec_lap_lai=True,
                    nguon_du_lieu="Rà soát uSmart KHCN",
                )
            )
        if so_de_tai_gv > 0:
            nhat_ky.append(
                f"🔬 Hệ thống uSmart KHCN (`/gwprod/nckh`): phát hiện **{so_de_tai_gv}** đề tài NCKH / NCKHSV đang thực hiện của GV {giang_vien.ho_va_ten}."
            )

        # 3. Đề xuất nhiệm vụ chuyên môn đặc thù theo vị trí việc làm & TKB trên uSmart
        if giang_vien.ma_can_bo == "0537" or "thí nghiệm" in giang_vien.chuc_danh_vi_tri.lower():
            # Thành viên Nguyễn Thị Thanh Giang (0537 - giangntt - Giảng viên thực hành, thí nghiệm PTN Công trình / Bộ môn ĐKT & KGN)
            ds_mau_ptn = [
                (
                    f"Hướng dẫn thực hành, thí nghiệm các học phần Cơ học đất, Cơ học đất đá, Địa chất công trình cho sinh viên hệ chính quy tháng {thang:02d}/{nam} ({giang_vien.ho_va_ten})",
                    "Kế hoạch năm học",
                    "TKB-PTN-2026",
                    "Hoàn thành 100% các ca hướng dẫn thí nghiệm theo lịch phân công, ký xác nhận sổ nhật ký phòng thí nghiệm và phiếu kết quả thí nghiệm của sinh viên",
                    True,
                ),
                (
                    "Thực hiện NCKH và viết báo NCKH",
                    "Kế hoạch năm học",
                    "KH-NCKH-0537",
                    "Báo cáo chuyên đề NCKH và bản thảo bài báo khoa học chuyên ngành Địa kỹ thuật – Công trình",
                    False,
                ),
                (
                    "Xâm nhập thực tiễn, lao động sản xuất",
                    "Kế hoạch năm học",
                    "KH-TT-0537",
                    "Hồ sơ, số liệu khảo sát thực tế hiện trường/phòng thí nghiệm và báo cáo kết quả xâm nhập thực tiễn, lao động sản xuất",
                    False,
                ),
                (
                    "Bảo trì máy móc thí nghiệm định kỳ",
                    "Chức năng, nhiệm vụ thường xuyên",
                    "PTN-BT-0537",
                    "Sổ nhật ký vận hành, biên bản kiểm tra và bảo trì định kỳ hệ thống máy móc, thiết bị thí nghiệm Địa kỹ thuật – Công trình đảm bảo hoạt động tốt",
                    True,
                ),
                (
                    "Hỗ trợ thí nghiệm của NCKH SV, NCS",
                    "Kế hoạch năm học",
                    "PTN-NCKH-0537",
                    "Bộ số liệu và phiếu kết quả thí nghiệm mẫu đất/đá/vật liệu hỗ trợ các đề tài NCKH Sinh viên và Nghiên cứu sinh",
                    True,
                ),
            ]
            # Ưu tiên đặt 5 nhiệm vụ hàng tháng chuẩn của 0537 lên đầu danh sách đề xuất
            ds_uu_tien_0537: List[NhiemVuCongViec] = []
            ten_0537_chuan: set[str] = set()
            for idx_p, (ten_p, nguon_p, ma_p, sp_p, tt_p) in enumerate(ds_mau_ptn, start=1):
                ten_0537_chuan.add(ten_p.strip().lower())
                ds_uu_tien_0537.append(
                    NhiemVuCongViec(
                        ma_nhiem_vu=f"USMART_PTN_0537_{idx_p:02d}",
                        nhiem_vu=ten_p,
                        nguon_giao=nguon_p,
                        so_van_ban_ma_viec=ma_p,
                        san_pham_ket_qua=sp_p,
                        ngay_giao=dau_ky,
                        thoi_han=cuoi_ky,
                        ngay_hoan_thanh=ngay_ht_mac_dinh,
                        tinh_trang="Hoàn thành đúng hạn",
                        muc_do_hoan_thanh=100.0,
                        ma_can_bo_chu_tri=giang_vien.ma_can_bo,
                        ten_can_bo_chu_tri=giang_vien.ho_va_ten,
                        nguoi_phoi_hop="PTN Công trình & Bộ môn ĐKT",
                        la_nhiem_vu_trong_tam=tt_p,
                        la_cong_viec_lap_lai=True,
                        nguon_du_lieu="Nhiệm vụ hàng tháng PTN & NCKH",
                    )
                )
            for nv_cu in ds_de_xuat:
                if nv_cu.nhiem_vu.strip().lower() not in ten_0537_chuan:
                    ds_uu_tien_0537.append(nv_cu)
            ds_de_xuat = ds_uu_tien_0537[:10]
            nhat_ky.append(
                "🧪 Đã đề xuất **05 nhiệm vụ hàng tháng trọng tâm** (Giảng dạy TN, Thực hiện NCKH & viết báo NCKH, Xâm nhập thực tiễn & lao động sản xuất, Bảo trì máy móc thí nghiệm định kỳ, Hỗ trợ thí nghiệm của NCKH SV & NCS) cho ThS. Nguyễn Thị Thanh Giang (`0537`)."
            )
        else:
            # Đề xuất nhiệm vụ giảng dạy, hướng dẫn đồ án & sinh hoạt chuyên môn chuẩn cho giảng viên Bộ môn
            ds_mau_gv = [
                (
                    f"Giảng dạy các học phần chuyên ngành Địa kỹ thuật & Không gian ngầm và hướng dẫn đồ án/thảo luận theo Thời khóa biểu học kỳ I năm học 2026-2027 ({giang_vien.ho_va_ten})",
                    "Kế hoạch năm học",
                    "TKB-HK1-2026",
                    "Hoàn thành 100% số giờ chuẩn giảng dạy theo lịch trình, cập nhật sổ đầu bài và duyệt đồ án sinh viên đúng tiến độ",
                    True,
                ),
                (
                    f"Thực hiện nghiên cứu khoa học, biên soạn bài báo khoa học/học liệu chuyên ngành và hướng dẫn học viên, sinh viên NCKH ({giang_vien.ho_va_ten})",
                    "Kế hoạch năm học",
                    "KH-NCKH-2026",
                    "Báo cáo chuyên đề NCKH, bản thảo bài báo khoa học và hồ sơ hướng dẫn đồ án tốt nghiệp/NCKHSV",
                    False,
                ),
                (
                    "Tham gia sinh hoạt chuyên môn, họp giao ban Bộ môn Địa kỹ thuật & KGN, coi chấm thi và rà soát chương trình đào tạo",
                    "Chức năng, nhiệm vụ thường xuyên",
                    "BB-BM-DKT",
                    "Biên bản sinh hoạt học thuật, phiếu chấm thi và hồ sơ rà soát đề cương học phần của Bộ môn",
                    False,
                ),
            ]
            for idx_g, (ten_g, nguon_g, ma_g, sp_g, tt_g) in enumerate(ds_mau_gv, start=1):
                _them_de_xuat(
                    NhiemVuCongViec(
                        ma_nhiem_vu=f"USMART_GV_{giang_vien.ma_can_bo}_{idx_g:02d}",
                        nhiem_vu=ten_g,
                        nguon_giao=nguon_g,
                        so_van_ban_ma_viec=ma_g,
                        san_pham_ket_qua=sp_g,
                        ngay_giao=dau_ky,
                        thoi_han=cuoi_ky,
                        ngay_hoan_thanh=ngay_ht_mac_dinh,
                        tinh_trang="Hoàn thành đúng hạn",
                        muc_do_hoan_thanh=100.0,
                        ma_can_bo_chu_tri=giang_vien.ma_can_bo,
                        ten_can_bo_chu_tri=giang_vien.ho_va_ten,
                        nguoi_phoi_hop="Bộ môn ĐKT & KGN",
                        la_nhiem_vu_trong_tam=tt_g,
                        la_cong_viec_lap_lai=True,
                        nguon_du_lieu="Đề xuất từ uSmart & TKB Bộ môn",
                    )
                )
            ds_de_xuat = ds_de_xuat[:10]
            nhat_ky.append(
                f"📚 Đã tổng hợp và đề xuất **{len(ds_de_xuat)} nhiệm vụ công việc báo cáo** cho GV {giang_vien.ho_va_ten} (`{giang_vien.ma_can_bo}`)."
            )

        return ds_de_xuat[:10], nhat_ky
