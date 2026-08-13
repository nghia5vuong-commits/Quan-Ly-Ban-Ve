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
    const ss = SpreadsheetApp.openById("1DRteBSFT1cj4R_OUPMoDxeLMzAIJexWF3HPT-rpMOoM");
    const sheet = ss.getSheetByName("User");

    const lastrow = sheet.getLastRow();
    const lastcol = sheet.getLastColumn();
    let data = [];

    if (lastrow > 1) {
      data = sheet.getRange(2, 1, lastrow - 1, lastcol).getValues();
    }

    const userRow = data.find(row => row[3].toString().trim().toLowerCase() === mail.toString().trim().toLowerCase());

    if (!userRow) return null;

    return {
      msnv: userRow[0],
      dept: userRow[1],
      name: userRow[2],
      mail: userRow[3],
      position: userRow[4],
    };
  } catch (e) {
    Logger.log("Lỗi lấy User: " + e.toString());
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
      var data = sheet.getDataRange().getValues();
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
    Logger.log(err.toString());
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
      Logger.log(e.toString());
    }
  }

  if (!profile.avatarUrl) {
    var svg = '<svg xmlns="http://www.w3.org/2000/svg" width="128" height="128" viewBox="0 0 128 128"><rect width="100%" height="100%" rx="32" fill="#0284c7"/><text x="50%" y="55%" dominant-baseline="middle" text-anchor="middle" fill="#ffffff" font-family="sans-serif" font-size="52" font-weight="bold">' + profile.initials + '</text></svg>';
    profile.avatarUrl = "data:image/svg+xml;base64," + Utilities.base64Encode(svg);
  }

  return profile;
}

function addNotification(notiText) {
  var ss = SpreadsheetApp.openById("1DRteBSFT1cj4R_OUPMoDxeLMzAIJexWF3HPT-rpMOoM");
  var sheet = ss.getSheetByName("Note");
  if (!sheet) return { success: false, error: "Không tìm thấy sheet Noti" };

  var profile = getUserProfile();
  var now = new Date();
  var dateStr = Utilities.formatDate(now, Session.getScriptTimeZone(), "yyyy/MM/dd");
  var timeStr = Utilities.formatDate(now, Session.getScriptTimeZone(), "HH:mm:ss");

  sheet.appendRow([profile.email, profile.name, notiText, dateStr, timeStr]);
  return { success: true };
}

function getLatestNotifications() {
  var ss = SpreadsheetApp.openById("1DRteBSFT1cj4R_OUPMoDxeLMzAIJexWF3HPT-rpMOoM");
  var sheet = ss.getSheetByName("Note");
  if (!sheet) return [];

  var data = sheet.getDataRange().getDisplayValues();
  if (data.length <= 1) return [];

  var notis = [];
  var start = Math.max(1, data.length - 20); // Chỉ lấy 20 dòng cuối

  // Chạy ngược từ dưới lên để lấy thông báo mới nhất
  for (var i = data.length - 1; i >= start; i--) {
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

