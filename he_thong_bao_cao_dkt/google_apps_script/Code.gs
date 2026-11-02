/**
 * BỘ MÃ GOOGLE APPS SCRIPT TỰ ĐỘNG HÓA BÁO CÁO CÔNG VIỆC THÁNG (TB 940/TB-ĐHGTVT)
 * Đơn vị: Bộ môn Địa kỹ thuật và Không gian ngầm - Trường Đại học Giao thông vận tải (UTC)
 * Thư mục Google Drive gốc: https://drive.google.com/drive/folders/1jawTokjmk-hnBsuoEIHoxYUhSRCEItPd
 *
 * Hướng dẫn gắn vào Google Sheets:
 * 1. Tải tệp `CSDL_BaoCao_BoMon_DKT_KGN.xlsx` lên thư mục Google Drive của Bộ môn, mở bằng Google Sheets
 *    và chọn "Tệp -> Lưu dưới dạng Google Trang tính".
 * 2. Vào "Tiện ích mở rộng -> Apps Script", dán nội dung `Code.gs` và tạo tệp HTML `Index.html`.
 * 3. Nhấn "Triển khai -> Triển khai mới -> Ứng dụng web (Web App)" để tạo giao diện Web cho 21 giảng viên.
 */

const DRIVE_FOLDER_ID = "1jawTokjmk-hnBsuoEIHoxYUhSRCEItPd";
const USMART_BAOCAO_URL = "https://usmart.utc.edu.vn/baocaocongviec/";

/**
 * Tạo Menu tác vụ nhanh ngay trên thanh công cụ của Google Sheets
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  const menuAdmin = ui.createMenu("🛡️ Quản trị & Đồng bộ Dữ liệu (Chỉ Admin)")
    .addItem("🔬 Đồng bộ Kế hoạch NCKH & NCKHSV từ uSmart", "dongBoKeHoachNCKHTuUSmartUI")
    .addItem("⏰ Bật tự động cập nhật NCKH hàng ngày (07:00)", "caiDatTriggerCapNhatNCKHDinhKy")
    .addItem("🎓 Mở Cổng Quản lý Đề tài NCKH uSmart", "moTrangTatCaDeTaiUSmart");

  ui.createMenu("📊 Báo cáo ĐKT-KGN (TB 940)")
    .addItem("🖥️ Mở Giao diện Webapp Báo cáo", "moSideBarBaoCao")
    .addSeparator()
    .addItem("🔄 1. Sao chép công việc lặp lại sang tháng mới", "saoChepCongViecLapLaiUI")
    .addItem("⏪ 2. Chuyển việc chưa hoàn thành kỳ trước", "chuyenViecChuaXongUI")
    .addItem("🔍 3. Kiểm tra Ràng buộc Vật lý & Cảnh báo", "kiemTraDuLieuToanBoSheet")
    .addSeparator()
    .addItem("🌐 Mở Cổng Báo cáo công việc uSmart", "moTrangUSmart")
    .addSeparator()
    .addSubMenu(menuAdmin)
    .addToUi();
}

/**
 * Phục vụ giao diện Webapp khi truy cập qua URL triển khai
 */
function doGet(e) {
  return HtmlService.createTemplateFromFile("Index")
    .evaluate()
    .setTitle("Hệ thống Báo cáo Công việc Bộ môn Địa kỹ thuật & KGN (TB 940)")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function moSideBarBaoCao() {
  const html = HtmlService.createTemplateFromFile("Index")
    .evaluate()
    .setTitle("Báo cáo ĐKT & KGN (TB 940)")
    .setWidth(450);
  SpreadsheetApp.getUi().showSidebar(html);
}

function moTrangUSmart() {
  const html = HtmlService.createHtmlOutput(
    '<script>window.open("' + USMART_BAOCAO_URL + '", "_blank");google.script.host.close();</script>'
  ).setHeight(60);
  SpreadsheetApp.getUi().showModalDialog(html, "Đang mở Cổng Báo cáo uSmart...");
}

/**
 * Lấy danh sách 21 giảng viên từ Sheet DM_GiangVien
 */
function layDanhSachGiangVien() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ws = ss.getSheetByName("DM_GiangVien");
  if (!ws) return [];
  const data = ws.getDataRange().getValues();
  const result = [];
  for (let i = 1; i < data.length; i++) {
    if (!data[i][0]) continue;
    result.push({
      ma_cb: String(data[i][0]),
      ho_ten: String(data[i][1]),
      hoc_ham: String(data[i][3] || ""),
      hoc_vi: String(data[i][4] || ""),
      chuc_danh: String(data[i][5] || "Giảng viên"),
      tk_usmart: String(data[i][6] || ""),
      nguoi_nhan: String(data[i][8] || "Trần Thế Truyền")
    });
  }
  return result;
}

/**
 * Sao chép công việc lặp lại từ Sheet CongViec_LapLai sang tháng mới
 */
function saoChepCongViecLapLai(thang, nam, maCanBoFilter) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const wsLap = ss.getSheetByName("CongViec_LapLai");
  const wsNv = ss.getSheetByName("NhiemVu_Thang");
  const dsGv = layDanhSachGiangVien();
  const mapGv = {};
  dsGv.forEach(g => { mapGv[g.ma_cb] = g; });

  const dataLap = wsLap.getDataRange().getValues();
  const dataNv = wsNv.getDataRange().getValues();

  const daCo = new Set();
  for (let i = 1; i < dataNv.length; i++) {
    if (Number(dataNv[i][1]) === Number(thang) && Number(dataNv[i][2]) === Number(nam)) {
      daCo.add(String(dataNv[i][3]).trim() + "||" + String(dataNv[i][6]).trim());
    }
  }

  const ngayDauThang = new Date(nam, thang - 1, 1);
  const ngayCuoiThang = new Date(nam, thang, 0);
  let dem = 0;

  for (let i = 1; i < dataLap.length; i++) {
    const kichHoat = Boolean(dataLap[i][8]);
    if (!kichHoat) continue;
    const apDung = String(dataLap[i][1]).trim();
    const tenNv = String(dataLap[i][2]).trim();
    const nguonGiao = String(dataLap[i][3]).trim();
    const soVb = String(dataLap[i][4]).trim();
    const sanPham = String(dataLap[i][5]).trim();
    const ngayHanInt = Number(dataLap[i][6] || 0);
    const trongTam = Boolean(dataLap[i][7]);

    const thoiHan = (ngayHanInt <= 0 || ngayHanInt > ngayCuoiThang.getDate())
      ? ngayCuoiThang
      : new Date(nam, thang - 1, ngayHanInt);

    const dsMcb = maCanBoFilter ? [maCanBoFilter] : (apDung === "ALL" ? Object.keys(mapGv) : [apDung]);
    dsMcb.forEach(mcb => {
      if ((apDung !== "ALL" && apDung !== mcb) || !mapGv[mcb]) return;
      const key = mcb + "||" + tenNv;
      if (daCo.has(key)) return;
      const maMoi = "NV_" + nam + String(thang).padStart(2, "0") + "_LAP_" + (wsNv.getLastRow() + 1);
      wsNv.appendRow([
        maMoi, thang, nam, mcb, mapGv[mcb].ho_ten, "Bộ môn ĐKT & KGN",
        tenNv, nguonGiao, soVb, sanPham, ngayDauThang, thoiHan, "",
        "Đang thực hiện – đúng tiến độ", 50, "", trongTam, true, "Thường",
        "Đạt yêu cầu – tốt", "", "Sao chép việc lặp lại"
      ]);
      daCo.add(key);
      dem++;
    });
  }
  return { thanh_cong: true, so_ban_ghi: dem };
}

function saoChepCongViecLapLaiUI() {
  const now = new Date();
  const kq = saoChepCongViecLapLai(now.getMonth() + 1, now.getFullYear(), null);
  SpreadsheetApp.getUi().alert("Đã sao chép " + kq.so_ban_ghi + " công việc lặp lại vào kỳ hiện tại!");
}

/**
 * Chuyển tiếp nhiệm vụ chưa hoàn thành kỳ trước sang kỳ mới (Giữ nguyên thời hạn gốc)
 */
function chuyenViecChuaXong(thangTruoc, namTruoc, thangMoi, namMoi, maCanBoFilter) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const wsNv = ss.getSheetByName("NhiemVu_Thang");
  const dataNv = wsNv.getDataRange().getValues();

  const daCo = new Set();
  const chuaXong = [];

  for (let i = 1; i < dataNv.length; i++) {
    const t = Number(dataNv[i][1]);
    const n = Number(dataNv[i][2]);
    const mcb = String(dataNv[i][3]).trim();
    const tenNv = String(dataNv[i][6]).trim();
    const ngayHt = dataNv[i][12];
    const tinhTrang = String(dataNv[i][13]).trim();

    if (t === Number(thangMoi) && n === Number(namMoi)) {
      daCo.add(mcb + "||" + tenNv);
    } else if (t === Number(thangTruoc) && n === Number(namTruoc)) {
      if (maCanBoFilter && mcb !== maCanBoFilter) continue;
      if (!tinhTrang.startsWith("Hoàn thành") && !ngayHt) {
        chuaXong.push(dataNv[i].slice());
      }
    }
  }

  let dem = 0;
  chuaXong.forEach(row => {
    const key = String(row[3]).trim() + "||" + String(row[6]).trim();
    if (daCo.has(key)) return;
    row[0] = "NV_" + namMoi + String(thangMoi).padStart(2, "0") + "_CT_" + (wsNv.getLastRow() + 1);
    row[1] = Number(thangMoi);
    row[2] = Number(namMoi);
    // Giữ nguyên Thời hạn gốc ở row[11] đúng hướng dẫn Mẫu 01
    row[21] = "Chuyển tiếp từ T" + String(thangTruoc).padStart(2, "0") + "/" + namTruoc;
    wsNv.appendRow(row);
    daCo.add(key);
    dem++;
  });
  return { thanh_cong: true, so_chuyen_tiep: dem };
}

function chuyenViecChuaXongUI() {
  const now = new Date();
  const tMoi = now.getMonth() + 1;
  const nMoi = now.getFullYear();
  const tTruoc = tMoi === 1 ? 12 : tMoi - 1;
  const nTruoc = tMoi === 1 ? nMoi - 1 : nMoi;
  const kq = chuyenViecChuaXong(tTruoc, nTruoc, tMoi, nMoi, null);
  SpreadsheetApp.getUi().alert("Đã chuyển tiếp " + kq.so_chuyen_tiep + " công việc chưa hoàn thành (giữ nguyên thời hạn gốc).");
}

/**
 * Kiểm tra ràng buộc vật lý và cảnh báo sai lệch trên toàn bộ Sheet NhiemVu_Thang
 */
function kiemTraDuLieuToanBoSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const wsNv = ss.getSheetByName("NhiemVu_Thang");
  const data = wsNv.getDataRange().getValues();
  const dsLoi = [];

  for (let i = 1; i < data.length; i++) {
    const dong = i + 1;
    const tenNv = String(data[i][6] || "").trim();
    if (!tenNv) continue;
    const nguon = String(data[i][7] || "").trim();
    const soVb = String(data[i][8] || "").trim();
    const sanPham = String(data[i][9] || "").trim();
    const thoiHan = data[i][11];
    const ngayHt = data[i][12];
    const tinhTrang = String(data[i][13] || "").trim();

    if (!thoiHan || !sanPham) {
      dsLoi.push("Dòng " + dong + ": Thiếu Thời hạn hoặc Sản phẩm đầu ra.");
    }
    if (tinhTrang.startsWith("Hoàn thành") && !ngayHt) {
      dsLoi.push("Dòng " + dong + ": Chọn Hoàn thành nhưng thiếu Ngày hoàn thành.");
    }
    if (ngayHt && !tinhTrang.startsWith("Hoàn thành")) {
      dsLoi.push("Dòng " + dong + ": Đã có Ngày hoàn thành nhưng tình trạng vẫn là '" + tinhTrang + "'.");
    }
    if (ngayHt && thoiHan && new Date(ngayHt) > new Date(thoiHan) && tinhTrang === "Hoàn thành đúng hạn") {
      dsLoi.push("Dòng " + dong + ": Ngày hoàn thành sau Thời hạn nhưng chọn 'Hoàn thành đúng hạn'.");
    }
    if (nguon === "Giao việc trên Usmart" && !soVb) {
      dsLoi.push("Dòng " + dong + ": Giao việc trên Usmart nhưng thiếu Số văn bản / Mã việc.");
    }
  }

  const ui = SpreadsheetApp.getUi();
  if (dsLoi.length === 0) {
    ui.alert("✅ KIỂM TOÁN ĐẠT 100%: Toàn bộ nhiệm vụ đều hợp lệ, sẵn sàng xuất báo cáo và nộp uSmart!");
  } else {
    ui.alert("⚠️ PHÁT HIỆN " + dsLoi.length + " LỖI TRONG VÙNG CÁCH LY:\n\n" + dsLoi.slice(0, 15).join("\n"));
  }
  return dsLoi;
}
