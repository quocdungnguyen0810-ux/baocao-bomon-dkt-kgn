# 🗺️ ROADMAP HƯỚNG DẪN SỬ DỤNG HỆ THỐNG BÁO CÁO CÔNG VIỆC
**Đơn vị:** Bộ môn Địa kỹ thuật và Không gian ngầm — Khoa Công trình — Trường Đại học Giao thông vận tải (UTC)  
**Căn cứ áp dụng:** Thông báo số **940/TB-ĐHGTVT** & Nghị định số **233/2026/NĐ-CP**  
**Trưởng bộ môn:** GS.TS. Trần Thế Truyền  

---

## 1. SƠ ĐỒ TỔNG QUAN 5 BƯỚC DÀNH CHO GIẢNG VIÊN TRONG BỘ MÔN (THAO TÁC TRONG 2 PHÚT)

```mermaid
flowchart LR
    B1["1️⃣ BƯỚC 1: KHỞI ĐẦU<br/>(Thanh bên trái)<br/>• Chọn Tháng / Năm<br/>• Chọn Tên Giảng viên<br/>• Chọn Vị trí lưu file"]
    B2["2️⃣ BƯỚC 2: CHỌN MẪU<br/>(Thanh điều hướng trên)<br/>• Mẫu 01a: Báo cáo Tuần<br/>• Mẫu 01b: Báo cáo Tháng<br/>• Mẫu 03: Sổ giao việc"]
    B3["3️⃣ BƯỚC 3: RÀ SOÁT VIỆC<br/>(Bảng tương tác Mục I.2)<br/>• Tích ☑️ Sử dụng việc<br/>• Tích 🔁 Tiếp tục kỳ sau<br/>• Đổi Ngày HT, Điểm, ⭐ Trọng tâm"]
    B4["4️⃣ BƯỚC 4: CHẤM ĐIỂM & LƯU<br/>(Mục II, III, IV, V)<br/>• Kiểm tra Giờ GD, NCKH<br/>• Tự chấm điểm (≤100)<br/>• Bấm 💾 Lưu vào CSDL"]
    B5["5️⃣ BƯỚC 5: XUẤT & NỘP<br/>(Cuối trang báo cáo)<br/>• Bấm 📊 Xuất Excel / 📝 Word<br/>• Bấm ⬇️ Tải về máy<br/>• Bấm 🚀 Nộp lên uSmart"]

    B1 --> B2 --> B3 --> B4 --> B5
```

---

## 2. SƠ ĐỒ CHI TIẾT QUY TRÌNH PHỐI HỢP GIỮA 21 GIẢNG VIÊN & TRƯỞNG BỘ MÔN (ADMIN)

```mermaid
flowchart TD
    subgraph GV["👨‍🏫 DÀNH CHO 21 GIẢNG VIÊN BỘ MÔN ĐỊA KỸ THUẬT & KGN"]
        A1["Mở Webapp Báo cáo Bộ môn<br/>Chọn đúng Họ tên & Kỳ báo cáo"] --> A2{"Chọn loại báo cáo<br/>cần thực hiện"}
        A2 -->|"Hàng tuần (Trước 16h Thứ Sáu)"| A3["📅 Mẫu 01a: Báo cáo Tuần cá nhân<br/>Kiểm tra nhiệm vụ tuần & Kế hoạch tuần tới"]
        A2 -->|"Cuối tháng (Ngày 25 -> 02 tháng sau)"| A4["📆 Mẫu 01b: Báo cáo Tháng cá nhân<br/>Mục I (Nhiệm vụ) + Mục II (Giờ GD/NCKH) + Mục III (Tự chấm điểm)"]
        A2 -->|"Khi nhận / tra cứu việc"| A5["📋 Mẫu 03a/03b/03c: Giao việc & Theo dõi<br/>Cập nhật tiến độ, sản phẩm đầu ra, link minh chứng"]

        A3 --> A6["Tại Mục I.2: Tích ☑️ Sử dụng / 🔁 Tiếp tục kỳ sau<br/>Chỉnh sửa Hạng mục, Thời hạn, Ngày hoàn thành<br/>Tích ⭐ Trọng tâm ĐV (nếu là việc chính của Bộ môn)"]
        A4 --> A6
        A6 --> A7["Bấm nút: 💾 Lưu toàn bộ thay đổi vào CSDL"]
        A7 --> A8["Bấm 📊 Xuất Excel (.xlsx) / 📝 Xuất Word (.docx)<br/>Bấm ⬇️ Tải trực tiếp về máy hoặc 🚀 Nộp lên uSmart"]
    end

    subgraph ADMIN["🛡️ DÀNH CHO TRƯỞNG BỘ MÔN / QUẢN TRỊ VIÊN (ADMIN)"]
        B1["Đăng nhập Quyền Admin ở Thanh bên trái<br/>(Mở khóa Module Báo cáo Đơn vị & Đồng bộ)"] --> B2["🏛️ Chọn Mục 3: Báo cáo công việc Đơn vị (Mẫu 02)"]
        B2 --> B3["📅 Mẫu 02a (BC Tuần ĐV) & 📆 Mẫu 02b (BC Tháng ĐV)<br/>Tự động gom các nhiệm vụ có dấu ⭐ Trọng tâm ĐV từ 21 GV"]
        B2 --> B4["📊 Mẫu 02c: Bảng theo dõi & Chấm điểm 21 Giảng viên<br/>Nhập Điểm Trưởng bộ môn chấm (Thang 100) & Xếp loại A/B/C/D"]
        B3 --> B5["Xuất trọn bộ Hồ sơ Đơn vị (.xlsx / .docx)<br/>Lưu vào ổ đĩa / Google Drive Bộ môn"]
        B4 --> B5
    end

    A7 -.->|"Tự động đồng bộ vào CSDL Bộ môn"| B3
    A7 -.->|"Tự động tổng hợp điểm & tiến độ"| B4
```

---

## 3. HƯỚNG DẪN THAO TÁC TỪNG BƯỚC CHO THÀNH VIÊN BỘ MÔN

| Bước | Vị trí trên màn hình | Thao tác cụ thể của Giảng viên | Kết quả đạt được |
| :--- | :--- | :--- | :--- |
| **Bước 1** | **Thanh bên trái (Sidebar)** | • Chọn **Tháng / Năm** báo cáo.<br/>• Chọn đúng **Mã cán bộ – Họ và tên** của mình.<br/>• Chọn **Vị trí lưu file** (Thư mục dự án, Downloads, Desktop hoặc tự nhập đường dẫn). | Hệ thống tự động nạp toàn bộ dữ liệu giảng dạy (TKB), đề tài NCKH và công việc của đúng giảng viên đó. |
| **Bước 2** | **Thanh chọn Kiểu báo cáo** | • Chọn **📄 1. Báo cáo công việc cá nhân**.<br/>• Chọn tiếp **📅 Mẫu 01a — Báo cáo Tuần** hoặc **📆 Mẫu 01b — Báo cáo Tháng**. | Hiển thị đúng mẫu biểu hành chính theo Thông báo 940/TB-ĐHGTVT ngay trên màn hình. |
| **Bước 3** | **Mục I.2 — Danh sách nhiệm vụ** | • **Tái sử dụng việc cũ:** Tích vào ô *"Hiển thị thêm công việc từ báo cáo kỳ trước..."* rồi tích `☑️ Sử dụng` ở dòng muốn lấy lại.<br/>• **Chỉnh sửa trực tiếp:** Nhấp đúp vào ô để sửa Nội dung, Nguồn giao, Thời hạn, Ngày hoàn thành, Tình trạng, `% Hoàn thành`.<br/>• **Chuyển lên Đơn vị:** Tích `⭐ Trọng tâm ĐV` nếu muốn đưa công việc đó vào Báo cáo Đơn vị (Mẫu 02a/02b). | Không phải gõ lại các công việc lặp lại hàng tháng (giảng dạy, hướng dẫn ĐATN, trực PTN, NCKH). |
| **Bước 4** | **Mục II, III, IV, V & Nút Lưu** | • **Mục II (Mẫu 01b):** Kiểm tra số giờ Giảng dạy, NCKH, Nhiệm vụ khác.<br/>• **Mục III (Mẫu 01b):** Nhập **Điểm tự chấm Tiêu chí chung** ($\le 30$) và **Điểm tự chấm Kết quả nhiệm vụ** ($\le 70$).<br/>• **Mục IV & V:** Rà soát kế hoạch kỳ tới và đề xuất kiến nghị (nếu có, mặc định để trống).<br/>• Bấm nút màu xanh **💾 Lưu toàn bộ thay đổi vào CSDL**. | Dữ liệu được kiểm tra logic vật lý và lưu an toàn vào CSDL Bộ môn. |
| **Bước 5** | **Khu vực Xuất file & uSmart** | • Bấm **📊 Xuất Mẫu 01 (.xlsx)** hoặc **📝 Xuất Bản cứng (.docx)**.<br/>• Bấm nút **⬇️ Tải trực tiếp tệp về máy (Save As...)** để lưu vào máy cá nhân.<br/>• Hoặc nhập mật khẩu uSmart ở thanh bên trái và bấm **🚀 Gửi lên uSmart**. | Có ngay file Excel/Word đúng chuẩn phông chữ, công thức của Nhà trường để ký nộp và lưu trữ. |

---

## 4. LỊCH TRÌNH ĐỊNH KỲ TRONG THÁNG (TIMELINE CHUẨN TB 940/TB-ĐHGTVT)

```mermaid
flowchart LR
    T1["📅 Thứ Sáu hàng tuần<br/>(Trước 16h00)<br/>Giảng viên xuất & nộp<br/>Mẫu 01a (BC Tuần)"]
    T2["📆 Ngày 25 -> Ngày 02<br/>tháng kế tiếp<br/>21 Giảng viên hoàn thành<br/>Mẫu 01b (BC Tháng)"]
    T3["🏛️ Ngày 02 -> Ngày 04<br/>tháng kế tiếp<br/>Trưởng Bộ môn chấm điểm<br/>& xuất Mẫu 02b, 02c"]
    T4["☁️ Trước Ngày 05<br/>tháng kế tiếp<br/>Đồng bộ lên cổng uSmart<br/>& Lưu Drive Bộ môn"]

    T1 --> T2 --> T3 --> T4
```

---
📧 **Hỗ trợ kỹ thuật & Quản trị hệ thống:** **Nguyễn Hải Hà_Bộ môn Địa kỹ thuật và không gian ngầm** ([haihadkt@utc.edu.vn](mailto:haihadkt@utc.edu.vn))
