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
from datetime import datetime
import logging
from pathlib import Path
from typing import Any, Dict, List, Optional
from urllib.parse import urljoin
from bs4 import BeautifulSoup
import requests
import urllib3

from .models import BaoCaoCaNhanThang, NhiemVuCongViec
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

        # Lấy giá trị _dien_luc hiện hành từ form
        url_get = f"{self.BASE_URL}/baocaocongviec/nhap?ky={ky_id}&mau={ma_mau}"
        r_get = self.session.get(url_get, timeout=self.timeout)
        soup_get = BeautifulSoup(r_get.text, "html.parser")
        dien_luc_inp = soup_get.find("input", {"name": "_dien_luc"})
        dien_luc_val = (
            dien_luc_inp.get("value")
            if dien_luc_inp
            else datetime.now().strftime("%Y-%m-%dT%H:%M:%S")
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
            payload[f"{prefix}-nhiem_vu"] = nv.nhiem_vu
            payload[f"{prefix}-nguon_giao"] = nv.nguon_giao
            payload[f"{prefix}-so_vb"] = nv.so_van_ban_ma_viec
            payload[f"{prefix}-san_pham"] = nv.san_pham_ket_qua
            payload[f"{prefix}-thoi_han"] = (
                nv.thoi_han.strftime("%Y-%m-%d") if nv.thoi_han else ""
            )
            payload[f"{prefix}-ngay_ht"] = (
                nv.ngay_hoan_thanh.strftime("%Y-%m-%d") if nv.ngay_hoan_thanh else ""
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

        # Kiểm toán dữ liệu đã lưu trên server bằng cách đọc ngược lại form
        xac_nhan = self.doc_du_lieu_webform_hien_tai(ky_id=ky_id, ma_mau=ma_mau)
        thanh_cong = (
            r_post.status_code == 200
            and xac_nhan.get("so_nhiem_vu", 0) == len(bao_cao.danh_sach_nhiem_vu)
        )

        thong_bao = (
            f"Đã {'lưu nháp' if che_do == 'luu_nhap' else 'nộp chính thức'} thành công "
            f"Mẫu {ma_mau} (Kỳ {ky_id}) lên uSmart cho GV {bao_cao.giang_vien.ho_va_ten}. "
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
    ) -> KetQuaGuiUSmart:
        """
        Tải trực tiếp tệp Excel (.xlsx) lên endpoint `POST /baocaocongviec/nop`.
        Lưu ý: Hành động này sẽ nộp chính thức tệp Excel lên hệ thống uSmart.
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
        co_loi = "lỗi" in main_text.lower() or "chưa đến kỳ" in main_text.lower()

        return KetQuaGuiUSmart(
            thanh_cong=(r_post.status_code == 200 and not co_loi),
            che_do="upload_xlsx",
            ky_id=str(ky_id),
            ma_mau=ma_mau,
            url_phan_hoi=r_post.url,
            http_status=r_post.status_code,
            thong_bao=main_text[:400],
        )
