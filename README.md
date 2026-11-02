# Hệ Thống Tác Tử Tự Động Hóa Báo Cáo Công Việc & Đồng Bộ uSmart (TB 940/TB-ĐHGTVT)

- **Đơn vị áp dụng**: Bộ môn Địa kỹ thuật và Không gian ngầm (Mã đơn vị: `412`) – Khoa Công trình – Trường Đại học Giao thông Vận tải (UTC)
- **Trưởng Bộ môn (Người nhận báo cáo Mẫu 01)**: GS.TS. Trần Thế Truyền (Mã CB: `0694`)
- **Phó Trưởng Bộ môn (Phụ trách đồng bộ KHCN)**: PGS.TS. Nguyễn Châu Lân (Mã CB: `0882`, TK uSmart: `lannc`)
- **Văn bản pháp lý**: Thông báo số `940/TB-ĐHGTVT` ngày 21/10/2026 và Nghị định số `233/2026/NĐ-CP`
- **Thư mục Google Drive Bộ môn**: `https://drive.google.com/drive/folders/1jawTokjmk-hnBsuoEIHoxYUhSRCEItPd`
- **Cổng nộp báo cáo uSmart**: `https://usmart.utc.edu.vn/baocaocongviec/`

---

## 1. Khởi Chạy Nhanh (3 Cách)

### Cách 1: Nhấp đúp tệp `CHAY_HE_THONG_BAO_CAO_DKT.bat`
Nhấp đúp chuột vào tệp [`CHAY_HE_THONG_BAO_CAO_DKT.bat`](./CHAY_HE_THONG_BAO_CAO_DKT.bat) tại thư mục gốc để mở bảng điều khiển tiếng Việt với 5 chức năng tự động.

### Cách 2: Mở Giao diện Webapp Tương tác (Streamlit)
```powershell
python -m streamlit run he_thong_bao_cao_dkt/app_webapp.py
```

### Cách 3: Chạy Pipeline Dòng lệnh Toàn trình (CLI)
```powershell
python -m he_thong_bao_cao_dkt.cli_baocao --thang 11 --nam 2026 --ma-cb 1051 --dong-bo-nckh --xuat-toan-bo-bo-mon --usmart-user hanh1 --usmart-pass 182128 --che-do-usmart luu_nhap
```

---

## 2. Cấu Trúc Thư Mục & Hồ Sơ Đầu Ra

1. **`CSDL_BaoCao_BoMon_DKT_KGN.xlsx`**: Cơ sở dữ liệu trung tâm của 21 giảng viên Bộ môn Địa kỹ thuật và Không gian ngầm (bao gồm danh bạ, nhiệm vụ tháng, kế hoạch giảng dạy/NCKH, công việc lặp lại, kế hoạch tháng tới, vùng cách ly lỗi và 3 trang tính chuyên đề NCKH/NCKHSV).
2. **`Ke_hoach_cong_viec/`**:
   - `Ke_hoach_NCKH_DeTai_CacCap_va_NCKHSV_DKT_KGN.xlsx`: Google Sheet 9 trang tổng hợp **25 đề tài NCKH các cấp** và **77 đề tài NCKH Sinh viên** giai đoạn 2025–2027 (trích xuất từ tổng số 307 đề tài trên uSmart).
   - `Ke_hoach_GiangDay_TKB_DKT_KGN_2026_2027.xlsx`: Google Sheet 3 trang tổng hợp **110 lớp học phần** Học kỳ 1 năm học 2026–2027 (**45 lớp** giảng dạy trong tháng 11/2026, tổng cộng **2.922,2 tiết quy đổi**).
   - `HUONG_DAN_SU_DUNG_GOOGLE_SHEET_NCKH.md`: Hướng dẫn tích hợp Google Sheets & Google Apps Script trên Google Drive.
3. **`output_bao_cao_dkt/Ky_11_2026/`**:
   - `BCThang_DKT_NguyenHaiHa_11-2026.xlsx` & `.docx`: Báo cáo cá nhân Mẫu 01a/01b của ThS. Nguyễn Hải Hà.
   - `BCThang_DKT_11-2026.xlsx` & `BCThang_DKT_DonVi_11-2026.docx`: Báo cáo tổng hợp đơn vị Mẫu 02a/02b/02c.
   - `SoGiaoViec_DKT_Mau03_11-2026.xlsx`: Sổ giao việc và theo dõi nhiệm vụ Mẫu 03a/03b/03c.
   - `BaoCao_CaNhan_21_GV/`: Trọn bộ **42 tệp báo cáo (`.xlsx` & `.docx`)** của toàn bộ 21 giảng viên Bộ môn.
4. **`he_thong_bao_cao_dkt/`**:
   - Mã nguồn Python chuẩn PEP 8, Type Hints và bộ kiểm thử tự động `he_thong_bao_cao_dkt/tests/`.
   - Bộ mã Google Apps Script tại `he_thong_bao_cao_dkt/google_apps_script/` (`Code.gs`, `NCKH_USmart_Sync.gs`, `Index.html`).
