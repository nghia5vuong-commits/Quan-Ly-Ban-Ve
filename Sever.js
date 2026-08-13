function doGet(e) {
  var page = e.parameter.page || ""; 
  var requestId = e.parameter.requestId|| ""; 
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
      signVal: userRow[5],
      signFormula: userRow[5] ? String(userRow[5]).trim() : "",

    };
  } catch (e) {
    Logger.log("Lỗi lấy User: " + e.toString());
    return null;
  }
}

function include(filename) {
  try {
    return HtmlService.createHtmlOutputFromFile(filename).getContent();
  } catch (e) {
    return "<p>Lỗi: Không tìm thấy file " + filename + "</p>";
  }
}


