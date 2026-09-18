function doGet(e) {
  // [BẠN CÓ THỂ XÓA DÒNG NÀY SAU KHI CHẠY 1 LẦN ĐỂ DỌN SẠCH TOÀN BỘ CACHE HỆ THỐNG]:
  resetAllSystemCache();

  var page = e.parameter.page || "";
  var requestId = e.parameter.requestId || "";
  var template = HtmlService.createTemplateFromFile("index");

  template.initialPage = page;

  var currentEmail = Session.getActiveUser().getEmail();

  // Đọc thông tin user trực tiếp từ Google Sheet trong thời gian thực (Không dùng Cache)
  var userInfo = getUser(currentEmail);

  template.userFromServer = userInfo || {
    msnv: "",
    dept: "", // Để trống hoặc điền mặc định
    name: "User (Chưa đăng ký)",
    mail: currentEmail,
    position: "",
    avatarUrl: "",
    initials: "U"
  };
  template.requestId = requestId;

  return template.evaluate()
    .setTitle("Quản lý bản vẽ đóng gói G2G")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// Hàm dọn sạch TOÀN BỘ cache của hệ thống (User, Profile, Email, Thống kê...) trên mọi tầng cache
function resetAllSystemCache() {
  try {
    var scriptCache = CacheService.getScriptCache();
    var userCache = CacheService.getUserCache();
    var docCache = CacheService.getDocumentCache();
    var keys = [];

    // 1. Quét dọn toàn bộ Cache User & Profile từ Sheet User
    var ss = SpreadsheetApp.openById("1DRteBSFT1cj4R_OUPMoDxeLMzAIJexWF3HPT-rpMOoM");
    var sheet = ss.getSheetByName("User");
    if (sheet) {
      var lastrow = sheet.getLastRow();
      if (lastrow >= 2) {
        var data = sheet.getRange(2, 4, lastrow - 1, 1).getValues(); // Cột D: Email
        for (var i = 0; i < data.length; i++) {
          var m = String(data[i][0] || '').trim().toLowerCase();
          if (m) {
            keys.push('user:' + Utilities.base64EncodeWebSafe(m));
            keys.push('profile:' + Utilities.base64EncodeWebSafe(m));
          }
        }
      }
    }

    // 2. Dọn Cache của user hiện tại đang chạy phiên
    var curMail = Session.getActiveUser().getEmail();
    if (curMail) {
      var normCur = String(curMail).trim().toLowerCase();
      keys.push('user:' + Utilities.base64EncodeWebSafe(normCur));
      keys.push('profile:' + Utilities.base64EncodeWebSafe(normCur));
    }

    // 3. Dọn tất cả các Cache nghiệp vụ khác trong toàn bộ dự án (Mail, Stats, v.v.)
    keys.push("MAIL_STATS");
    keys.push("MAIL_LIST_0_");
    for (var page = 0; page < 30; page++) {
      keys.push("MAIL_LIST_" + page + "_");
      keys.push("MAIL_LIST_" + page);
    }

    // 4. Xóa đồng loạt trên cả 3 tầng Cache của Google Apps Script
    if (keys.length > 0) {
      try { if (scriptCache) scriptCache.removeAll(keys); } catch (e1) {}
      try { if (userCache) userCache.removeAll(keys); } catch (e2) {}
      try { if (docCache) docCache.removeAll(keys); } catch (e3) {}
    }
  } catch (err) {
    Logger.log("Lỗi xóa cache toàn hệ thống: " + err);
  }
}

function getUser(mail) {
  try {
    var normalizedMail = String(mail || '').trim().toLowerCase();
    if (!normalizedMail) return null;

    // Đọc trực tiếp từ Sheet User, KHÔNG dùng Cache
    const ss = SpreadsheetApp.openById("1DRteBSFT1cj4R_OUPMoDxeLMzAIJexWF3HPT-rpMOoM");
    const sheet = ss.getSheetByName("User");
    if (!sheet) return null;

    const lastrow = sheet.getLastRow();
    const lastcol = Math.min(sheet.getLastColumn(), 6);
    let data = [];

    if (lastrow > 1) {
      data = sheet.getRange(2, 1, lastrow - 1, lastcol).getValues();
    }

    const userRow = data.find(row => row[3] && row[3].toString().trim().toLowerCase() === normalizedMail);

    if (!userRow) {
      return null;
    }

    var userName = userRow[2] ? String(userRow[2]).trim() : "";
    var names = userName.split(" ");
    var initials = names.length > 0 && names[names.length - 1] ? names[names.length - 1].charAt(0).toUpperCase() : "U";
    var avatarUrl = (userRow[5] && String(userRow[5]).trim()) ? String(userRow[5]).trim() : "";
    var userPosition = userRow[4] ? String(userRow[4]).trim() : "";

    return {
      msnv: userRow[0] || "",
      dept: userRow[1] || "",
      name: userName,
      mail: userRow[3] || normalizedMail,
      position: userPosition,
      avatarUrl: avatarUrl,
      initials: initials
    };
  } catch (e) {
    return null;
  }
}

// =========================================================================
// HỆ THỐNG USER PROFILE & NOTIFICATION
// =========================================================================

function getUserProfile() {
  var email = Session.getActiveUser().getEmail();
  if (!email) email = "test@example.com";

  var profile = {
    name: email.split('@')[0],
    position: "Thành viên",
    email: email,
    initials: "U",
    avatarUrl: ""
  };

  try {
    var ss = SpreadsheetApp.openById("1DRteBSFT1cj4R_OUPMoDxeLMzAIJexWF3HPT-rpMOoM");
    var sheet = ss.getSheetByName("User");

    if (sheet) {
      var profileLastRow = sheet.getLastRow();
      var profileLastColumn = Math.min(sheet.getLastColumn(), 6);
      var data = profileLastRow > 0 && profileLastColumn > 0
        ? sheet.getRange(1, 1, profileLastRow, profileLastColumn).getValues()
        : [];
      for (var i = 1; i < data.length; i++) {
        if (data[i][3] && data[i][3].toString().trim().toLowerCase() === email.toLowerCase()) {
          if (data[i][2]) profile.name = data[i][2].toString().trim();
          if (data[i][4]) profile.position = data[i][4].toString().trim();
          if (data[i][5] && data[i][5].toString().trim() !== "") {
            profile.avatarUrl = data[i][5].toString().trim();
          }
          break;
        }
      }
    }
  } catch (err) {
  }

  var names = profile.name.trim().split(" ");
  profile.initials = names.length > 0 && names[names.length - 1] ? names[names.length - 1].charAt(0).toUpperCase() : "U";

  if (!profile.avatarUrl) {
    try {
      var res = People.People.get('people/me', { personFields: 'photos' });
      if (res && res.photos && res.photos.length > 0) {
        var photoUrl = res.photos[0].url;
        profile.avatarUrl = photoUrl.replace(/=s\d+(-c)?$/, '=s128-c');
      }
    } catch (e) {
    }
  }

  if (!profile.avatarUrl) {
    var svg = '<svg xmlns="http://www.w3.org/2000/svg" width="128" height="128" viewBox="0 0 128 128"><rect width="100%" height="100%" rx="32" fill="#0284c7"/><text x="50%" y="55%" dominant-baseline="middle" text-anchor="middle" fill="#ffffff" font-family="sans-serif" font-size="52" font-weight="bold">' + profile.initials + '</text></svg>';
    profile.avatarUrl = "data:image/svg+xml;base64," + Utilities.base64Encode(svg);
  }

  return profile;
}

function addNotification(notiText) {
  if (!notiText || !String(notiText).trim()) {
    return { success: false, error: "Nội dung thông báo trống" };
  }

  try {
    var ss = SpreadsheetApp.openById("1DRteBSFT1cj4R_OUPMoDxeLMzAIJexWF3HPT-rpMOoM");
    var sheet = ss.getSheetByName("Note");
    if (!sheet) return { success: false, error: "Không tìm thấy sheet Noti" };

    var profile = getUserProfile();
    var now = new Date();
    var dateStr = Utilities.formatDate(now, Session.getScriptTimeZone(), "yyyy/MM/dd");
    var timeStr = Utilities.formatDate(now, Session.getScriptTimeZone(), "HH:mm:ss");

    sheet.appendRow([profile.email, profile.name, String(notiText).trim(), dateStr, timeStr]);
    return { success: true };
  } catch (err) {
    return { success: false, error: err && err.message ? err.message : "Lỗi lưu thông báo" };
  }
}

function getLatestNotifications() {
  var ss = SpreadsheetApp.openById("1DRteBSFT1cj4R_OUPMoDxeLMzAIJexWF3HPT-rpMOoM");
  var sheet = ss.getSheetByName("Note");
  if (!sheet) return [];

  var lastRow = sheet.getLastRow();
  if (lastRow <= 1) return [];

  // Chỉ lấy phần dữ liệu cần hiển thị; tránh đọc toàn bộ lịch sử Note.
  var startRow = Math.max(2, lastRow - 19);
  var data = sheet.getRange(startRow, 1, lastRow - startRow + 1, 5).getDisplayValues();

  var notis = [];

  // Chạy ngược từ dưới lên để lấy thông báo mới nhất
  for (var i = data.length - 1; i >= 0; i--) {
    notis.push({
      email: data[i][0],
      user: data[i][1],
      noti: data[i][2],
      date: data[i][3],
      time: data[i][4]
    });
  }

  return notis;
}

function include(filename) {
  try {
    return HtmlService.createHtmlOutputFromFile(filename).getContent();
  } catch (e) {
    return "<p>Lỗi: Không tìm thấy file " + filename + "</p>";
  }
}

