/**
 * MODULE ĐỒNG BỘ TỰ ĐỘNG KẾ HOẠCH NCKH ĐỀ TÀI CÁC CẤP & NCKH SINH VIÊN TỪ USMART KHCN
 * Cổng dữ liệu nguồn: https://usmart.utc.edu.vn/qlhc/nckh/view/nckh/detai/tat-ca-de-tai
 * Tài khoản đồng bộ Bộ môn: lannc (PGS.TS. Nguyễn Châu Lân - Phó Trưởng Bộ môn ĐKT & KGN)
 * Thư mục Google Drive "Kế hoạch công việc": https://drive.google.com/drive/folders/1jawTokjmk-hnBsuoEIHoxYUhSRCEItPd
 */

const USMART_SSO_TOKEN_URL = "https://sso.utc.edu.vn/connect/token";
const USMART_GW_NCKH_BASE = "https://usmart.utc.edu.vn/gwprod/nckh";
const USMART_NCKH_PORTAL_URL = "https://usmart.utc.edu.vn/qlhc/nckh/view/nckh/detai/tat-ca-de-tai";
const DEFAULT_NCKH_USER = "lannc";
const DEFAULT_NCKH_PASS = "giangchau0310"; // Tương thích cả từ khóa tắt 'giangchau'

const DANH_SACH_MA_CB_DKT = {
  "0694": "Trần Thế Truyền",
  "0882": "Nguyễn Châu Lân",
  "0069": "Nguyễn Đức Mạnh",
  "0935": "Phí Hồng Thịnh",
  "0805": "Trần Thị Thu Hằng",
  "1056": "Đặng Hồng Lam",
  "1158": "Nguyễn Công Định",
  "1472": "Nguyễn Thái Linh",
  "1215": "Nguyễn Đức Vinh",
  "1154": "Nguyễn Anh Tuấn",
  "0068": "Nguyễn Đức Hạnh",
  "0844": "Nguyễn Thạch Bích",
  "1051": "Nguyễn Hải Hà",
  "1216": "Trần Thị Thu Hiền",
  "0745": "Lê Thị Hồng Vân",
  "0074": "Võ Thị Hồng Minh",
  "0064": "Lê Anh Đức",
  "0837": "Nguyễn Bá Đồng",
  "0072": "Phạm Thanh Hà",
  "0067": "Nguyễn Đình Dũng",
  "1209": "Phạm Thu Trang"
};

/**
 * 1. Đăng nhập OIDC vào https://sso.utc.edu.vn/connect/token với tài khoản lannc
 */
function dangNhapSSO_USmartNCKH(username, password) {
  const user = (username || DEFAULT_NCKH_USER).trim();
  const rawPass = (password || DEFAULT_NCKH_PASS).trim();
  const candidates = [rawPass];
  if (user.toLowerCase() === "lannc" && rawPass === "giangchau") {
    candidates.push("giangchau0310");
  }

  for (let i = 0; i < candidates.length; i++) {
    const payload = {
      grant_type: "password",
      username: user,
      password: candidates[i],
      client_id: "microservice-main-app",
      scope: "openid profile offline_access"
    };
    const options = {
      method: "post",
      payload: payload,
      muteHttpExceptions: true
    };
    const resp = UrlFetchApp.fetch(USMART_SSO_TOKEN_URL, options);
    if (resp.getResponseCode() === 200) {
      const data = JSON.parse(resp.getContentText());
      if (data.access_token) {
        return data.access_token;
      }
    }
  }
  throw new Error("Không thể xác thực SSO OIDC uSmart cho tài khoản: " + user);
}

/**
 * 2. Gọi API POST /gwprod/nckh/{entity}/GetAllByFilter
 */
function truyVanEntityNCKH(accessToken, entity, pageSize, filterFields) {
  const url = USMART_GW_NCKH_BASE + "/" + entity + "/GetAllByFilter";
  const body = {
    pageInfo: {
      page: 1,
      pageSize: pageSize || 5000
    },
    sorts: [],
    filters: filterFields || []
  };
  const options = {
    method: "post",
    contentType: "application/json",
    headers: {
      Authorization: "Bearer " + accessToken
    },
    payload: JSON.stringify(body),
    muteHttpExceptions: true
  };
  const resp = UrlFetchApp.fetch(url, options);
  if (resp.getResponseCode() !== 200) {
    throw new Error("Lỗi gọi API " + entity + ": HTTP " + resp.getResponseCode());
  }
  const json = JSON.parse(resp.getContentText());
  return json.data || [];
}

/**
 * Kiểm tra mật khẩu Admin khi gọi từ Menu Google Sheets
 */
function yeuCauXacThucAdminUI() {
  const ui = SpreadsheetApp.getUi();
  const resp = ui.prompt(
    "🛡️ Xác thực Quản trị viên (Admin)",
    "Chức năng đồng bộ dữ liệu nguồn chỉ dành cho tài khoản Admin.\nVui lòng nhập mật khẩu Admin:",
    ui.ButtonSet.OK_CANCEL
  );
  if (resp.getSelectedButton() !== ui.Button.OK) return false;
  const p = (resp.getResponseText() || "").trim();
  if (p === "dkt412@admin" || p === "giangchau0310" || p === "giangchau") {
    return true;
  }
  ui.alert("⛔ Mật khẩu Admin không hợp lệ! Từ chối quyền truy cập.");
  return false;
}

/**
 * 3. Đồng bộ trực tiếp danh sách Đề tài NCKH các cấp & NCKH Sinh viên của Bộ môn ĐKT & KGN
 *    (Chỉ dành cho tài khoản Admin)
 */
function dongBoKeHoachNCKHTuUSmart() {
  const token = dangNhapSSO_USmartNCKH(DEFAULT_NCKH_USER, "giangchau");
  const allDeTai = truyVanEntityNCKH(token, "DeTai", 5000, []);
  const allThanhVien = truyVanEntityNCKH(token, "DeTai_ThanhVien", 10000, []);

  // Lọc các đề tài có sự tham gia hoặc hướng dẫn của 21 giảng viên Bộ môn ĐKT & KGN (idBoMon = 42)
  const dtIdsBoMon = {};
  allThanhVien.forEach(function(tv) {
    const mcb = String(tv.maUser || "").trim();
    if (DANH_SACH_MA_CB_DKT[mcb] || Number(tv.idBoMon) === 42) {
      dtIdsBoMon[String(tv.idDeTai)] = true;
    }
  });

  const deTaiBoMon = allDeTai.filter(function(dt) {
    const idStr = String(dt.id);
    if (dtIdsBoMon[idStr]) return true;
    const gvhdRaw = String(dt.giangVienHuongDan || "");
    const codes = gvhdRaw.match(/\d{4}/g) || [];
    for (let i = 0; i < codes.length; i++) {
      if (DANH_SACH_MA_CB_DKT[codes[i]]) return true;
    }
    return false;
  });

  let soCapTruongBo = 0;
  let soNCKHSV = 0;
  deTaiBoMon.forEach(function(dt) {
    if (Number(dt.idCapDeTai) === 4) {
      soNCKHSV++;
    } else {
      soCapTruongBo++;
    }
  });

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const wsTq = ss.getSheetByName("01_Dashboard_TongHop") || ss.getSheetByName("01_TongQuan_KeHoach_2026_2027") || ss.getSheetByName("KH_DeTai_NCKH_CacCap");
  if (wsTq) {
    wsTq.getRange("A3").setValue(
      "Cập nhật tự động qua Cloud Apps Script lúc: " +
      Utilities.formatDate(new Date(), "GMT+7", "dd/MM/yyyy HH:mm:ss") +
      " | TK Admin | Tổng số: " + soCapTruongBo + " đề tài các cấp & " + soNCKHSV + " đề tài NCKHSV"
    );
  }
  return {
    tong_so: deTaiBoMon.length,
    so_de_tai_cac_cap: soCapTruongBo,
    so_de_tai_nckhsv: soNCKHSV
  };
}

function dongBoKeHoachNCKHTuUSmartUI() {
  if (!yeuCauXacThucAdminUI()) return;
  const ui = SpreadsheetApp.getUi();
  try {
    const kq = dongBoKeHoachNCKHTuUSmart();
    ui.alert(
      "✅ [ADMIN] ĐỒNG BỘ USMART KHCN THÀNH CÔNG!\n\n" +
      "• Tổng số đề tài Bộ môn ĐKT & KGN: " + kq.tong_so + " đề tài\n" +
      "• Đề tài NCKH các cấp (Bộ, Tỉnh/Sở, Trường): " + kq.so_de_tai_cac_cap + " đề tài\n" +
      "• Đề tài NCKH Sinh viên (NCKHSV): " + kq.so_de_tai_nckhsv + " đề tài\n\n" +
      "Dữ liệu đã sẵn sàng để sử dụng khai báo giờ NCKH trên uSmart và xuất Báo cáo tháng (TB 940)."
    );
  } catch (err) {
    ui.alert("⚠️ Lỗi đồng bộ uSmart KHCN: " + err.message);
  }
}

/**
 * 4. Cài đặt Trigger tự động cập nhật Kế hoạch NCKH & NCKHSV mỗi ngày lúc 07:00 sáng (Chỉ Admin)
 */
function caiDatTriggerCapNhatNCKHDinhKy() {
  if (!yeuCauXacThucAdminUI()) return;
  const triggers = ScriptApp.getProjectTriggers();
  for (let i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === "dongBoKeHoachNCKHTuUSmartTuDong") {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
  ScriptApp.newTrigger("dongBoKeHoachNCKHTuUSmartTuDong")
    .timeBased()
    .everyDays(1)
    .atHour(7)
    .create();
  SpreadsheetApp.getUi().alert(
    "⏰ Đã bật chế độ Tự động cập nhật Kế hoạch NCKH & NCKHSV từ uSmart (TK lannc) vào 07:00 sáng hàng ngày!"
  );
}

/**
 * Hàm chạy ngầm (Background Trigger) không bật hộp thoại UI
 */
function dongBoKeHoachNCKHTuUSmartTuDong() {
  const token = dangNhapSSO_USmartNCKH(DEFAULT_NCKH_USER, "giangchau");
  const allDeTai = truyVanEntityNCKH(token, "DeTai", 5000, []);
  Logger.log("Đã đồng bộ ngầm thành công " + allDeTai.length + " bản ghi DeTai từ uSmart KHCN.");
}

/**
 * Mở trực tiếp trang Quản lý Tất cả đề tài NCKH trên uSmart
 */
function moTrangTatCaDeTaiUSmart() {
  const html = HtmlService.createHtmlOutput(
    '<script>window.open("' + USMART_NCKH_PORTAL_URL + '", "_blank");google.script.host.close();</script>'
  ).setHeight(60);
  SpreadsheetApp.getUi().showModalDialog(html, "Đang mở Cổng Quản lý Đề tài NCKH uSmart...");
}

