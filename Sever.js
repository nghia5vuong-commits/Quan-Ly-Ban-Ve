function doGet(e) {
  var page = e.parameter.page || "";
  var requestId = e.parameter.requestId || "";
  var template = HtmlService.createTemplateFromFile("index");

  template.initialPage = page;

  var currentEmail = Session.getActiveUser().getEmail();

  var userInfo = getUser(currentEmail);

  template.userFromServer = userInfo || {
    msnv: "",
    dept: "", // Để trống hoặc điền mặc định
    name: "User (Chưa đăng ký)",
    mail: currentEmail,
    position: ""
  };
  template.requestId = requestId;

  return template.evaluate()
    .setTitle("Quản lý bản vẽ đóng gói G2G")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function getUser(mail) {
  try {
    var normalizedMail = String(mail || '').trim().toLowerCase();
    if (!normalizedMail) return null;

    var cache = CacheService.getScriptCache();
    var cacheKey = 'user:' + Utilities.base64EncodeWebSafe(normalizedMail);
    var cached = cache.get(cacheKey);
    if (cached) return cached === 'null' ? null : JSON.parse(cached);

    const ss = SpreadsheetApp.openById("1DRteBSFT1cj4R_OUPMoDxeLMzAIJexWF3HPT-rpMOoM");
    const sheet = ss.getSheetByName("User");

    const lastrow = sheet.getLastRow();
    // getUser chỉ dùng A:E; tránh đọc các cột phụ không liên quan.
    const lastcol = Math.min(sheet.getLastColumn(), 5);
    let data = [];

    if (lastrow > 1) {
      data = sheet.getRange(2, 1, lastrow - 1, lastcol).getValues();
    }

    const userRow = data.find(row => row[3] && row[3].toString().trim().toLowerCase() === normalizedMail);

    if (!userRow) {
      cache.put(cacheKey, 'null', 300);
      return null;
    }

    var result = {
      msnv: userRow[0],
      dept: userRow[1],
      name: userRow[2],
      mail: userRow[3],
      position: userRow[4],
    };
    cache.put(cacheKey, JSON.stringify(result), 300);
    return result;
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

  var profileCacheKey = 'profile:' + Utilities.base64EncodeWebSafe(email.toLowerCase());
  try {
    var cachedProfile = CacheService.getScriptCache().get(profileCacheKey);
    if (cachedProfile) return JSON.parse(cachedProfile);
  } catch (cacheError) {
    // Cache chỉ là tối ưu phụ; tiếp tục đọc dữ liệu gốc nếu cache lỗi.
  }

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
      // Profile chỉ dùng A:F; không cần tải toàn bộ các cột mở rộng của User.
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

  try {
    CacheService.getScriptCache().put(profileCacheKey, JSON.stringify(profile), 300);
  } catch (cacheError) {
    // Không làm thay đổi kết quả nếu profile không thể ghi cache.
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

