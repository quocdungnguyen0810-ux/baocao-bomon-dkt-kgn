# HƯỚNG DẪN SỬ DỤNG & CẬP NHẬT TỰ ĐỘNG GOOGLE SHEET KẾ HOẠCH NCKH ĐỀ TÀI CÁC CẤP & NCKH SINH VIÊN

- **Đơn vị**: Bộ môn Địa kỹ thuật và Không gian ngầm – Khoa Công trình – Trường Đại học Giao thông vận tải (UTC)
- **Cổng dữ liệu nguồn**: `https://usmart.utc.edu.vn/qlhc/nckh/view/nckh/detai/tat-ca-de-tai`
- **Tài khoản đồng bộ uSmart KHCN**: `lannc` (**PGS.TS. Nguyễn Châu Lân** – Phó Trưởng Bộ môn, Mã CB: `0882`)
- **Thư mục lưu trữ**: `Ke_hoach_cong_viec/` (OneDrive & Google Drive Folder ID: `1jawTokjmk-hnBsuoEIHoxYUhSRCEItPd`)

---

## 1. Danh mục tệp dữ liệu trong thư mục `Kế hoạch công việc` (`Ke_hoach_cong_viec/`)

1. **`Ke_hoach_NCKH_DeTai_CacCap_va_NCKHSV_DKT_KGN.xlsx`** *(Sổ tính Google Sheet / Excel 9 trang chuyên đề NCKH & NCKHSV)*:
   - `01_Dashboard_TongHop`: Bảng điều khiển (Dashboard) sử dụng công thức động `COUNTIF`/`SUM` tổng hợp KPI NCKH toàn Bộ môn.
   - `02_KH_DeTai_NCKH_2025_2027`: Kế hoạch thực hiện **25 đề tài NCKH các cấp** đang triển khai giai đoạn 2025–2027 (Cấp Bộ GD&ĐT, Cấp Bộ Xây dựng, Cấp Tỉnh/Sở KH&CN, Cấp Trường trọng điểm, Cấp Trường `T2026` & `T2027`).
   - `03_KH_NCKHSV_2025_2027`: Kế hoạch thực hiện **77 đề tài NCKH Sinh viên (`NCKHSV`)** năm học 2025–2026 (36 đề tài) và đăng ký mới năm học 2026–2027 (11 đề tài) kèm GVHD, danh sách sinh viên, điểm nghiệm thu, giải thưởng và giờ chuẩn quy đổi.
   - `04_TongHop_21GV_KhaiBao_BaoCao`: Bảng tổng hợp định mức, giờ quy đổi thực tế và **Chuỗi chuẩn copy-paste khai báo uSmart / Báo cáo tháng (Mục II.1.2 Mẫu 01b)** cho từng giảng viên trong 21 giảng viên Bộ môn.
   - `05_CSDL_TatCa_DeTai_CacCap`: Toàn bộ **45 đề tài NCKH các cấp** hợp lệ của Bộ môn qua các năm.
   - `06_CSDL_TatCa_NCKHSV`: Toàn bộ **257 đề tài NCKH Sinh viên** của Bộ môn trên uSmart.
   - `07_LichTrinh_DinhMuc_PKHCN`: Lịch trình mốc thời gian KHCN & NCKHSV năm học 2026–2027 theo văn bản Phòng KH&CGCN và định mức quy đổi giờ chuẩn QĐ 1293.
   - `08_NhiemVu_DongBo_Mau01b_02b`: Bảng trích xuất **63 nhiệm vụ NCKH & NCKHSV thực tế** tương thích 100% cấu trúc 22 cột của Sheet `NhiemVu_Thang` (Mẫu 01a/01b, Mẫu 02b, Mẫu 03 theo Thông báo 940/TB-ĐHGTVT).
   - `09_VungCachLy_KiemToan_NCKH`: **Vùng cách ly (Quarantine Zone)** lưu trữ **7 bản ghi** vi phạm quy tắc vật lý/quy chế (`NCKH-R1`..`NCKH-R5`) kèm nguyên nhân và hướng dẫn xử lý, tuyệt đối không tự ý xóa bỏ dữ liệu gốc.

2. **`Ke_hoach_GiangDay_TKB_DKT_KGN_2026_2027.xlsx`** *(Sổ tính Google Sheet / Excel 3 trang Thời khóa biểu & Khối lượng giảng dạy HK1 2026–2027)*:
   - `01_TongHop_KLGD_21GV`: Bảng tổng hợp định mức giảng dạy đã giảm trừ, giờ hướng dẫn NCS, giờ chính quy, giờ ĐTTX, tổng giờ lũy kế Học kỳ 1 (`2.922,2` tiết chuẩn) và giờ giảng dạy thực hiện trong tháng 11/2026 của 21 giảng viên Bộ môn.
   - `02_ChiTiet_LopHocPhan_HK1`: Danh sách chi tiết **110 lớp học phần** (**45 lớp** giảng dạy trực tiếp trong tháng 11/2026) trích xuất từ `2627_1_1_TKB dự kiến- ĐKT.xlsx`.
   - `03_KiemToan_TKB_Quarantine`: Vùng cách ly & Nhật ký kiểm toán dữ liệu Thời khóa biểu (`GD-R1`..`GD-R3`).

3. **`CSDL_BaoCao_BoMon_DKT_KGN.xlsx`**:
   - Tệp CSDL Báo cáo công việc tháng của Bộ môn đã được tự động đồng bộ 3 trang tính NCKH (`KH_DeTai_NCKH_CacCap`, `KH_DeTai_NCKHSV`, `TongHop_KhaiBao_NCKH_21GV`), nạp **63 nhiệm vụ NCKH & NCKHSV** vào `NhiemVu_Thang` và cập nhật khối lượng giảng dạy thực tế của **110 lớp học phần** vào `KH_GiangDay_NCKH`.

---

## 2. Cách cập nhật dữ liệu thường xuyên (3 phương thức linh hoạt)

### Cách 1: Chạy lệnh CLI đồng bộ trực tiếp (1 giây)
```powershell
# Chỉ đồng bộ và cập nhật Google Sheet Kế hoạch NCKH & NCKHSV + CSDL Báo cáo:
python -m he_thong_bao_cao_dkt.usmart_nckh_sync

# Chạy toàn trình: Đồng bộ NCKH (TK lannc) + Xuất Báo cáo 21 giảng viên & Mẫu 02, 03 + Lưu nháp uSmart:
python -m he_thong_bao_cao_dkt.cli_baocao --thang 11 --nam 2026 --dong-bo-nckh --xuat-toan-bo-bo-mon
```

### Cách 2: Cập nhật qua Giao diện Webapp Streamlit
```powershell
streamlit run he_thong_bao_cao_dkt/app_webapp.py
```
- Chuyển sang Tab thứ 3: **"🔬 Kế hoạch NCKH Đề tài các cấp & NCKHSV (uSmart lannc)"**.
- Nhấn nút **"🔄 Đồng bộ trực tiếp từ uSmart KHCN (TK lannc) & Cập nhật Google Sheet"**.

### Cách 3: Tự động cập nhật hàng ngày trên Google Sheets (Google Drive Folder `1jawTokjmk-hnBsuoEIHoxYUhSRCEItPd`)
1. Mở tệp `Ke_hoach_NCKH_DeTai_CacCap_va_NCKHSV_DKT_KGN.xlsx` trên Google Drive bằng **Google Trang tính (Google Sheets)**.
2. Vào **Tiện ích mở rộng (Extensions) -> Apps Script**, dán nội dung 2 tệp trong `he_thong_bao_cao_dkt/google_apps_script/`:
   - `Code.gs`
   - `NCKH_USmart_Sync.gs`
3. Trên thanh công cụ của Google Sheets, chọn menu **"📊 Báo cáo ĐKT-KGN (TB 940)"**:
   - Nhấn **"🔬 Đồng bộ Kế hoạch NCKH & NCKHSV từ uSmart (TK lannc)"** để cập nhật tức thời.
   - Nhấn **"⏰ Bật tự động cập nhật NCKH hàng ngày (07:00)"** để Google Cloud tự động gọi API `usmart.utc.edu.vn/gwprod/nckh/` mỗi sáng.
