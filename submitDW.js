var SUBMIT_SS_ID = "1DRteBSFT1cj4R_OUPMoDxeLMzAIJexWF3HPT-rpMOoM";
var SUBMIT_SHEET = "Data";
var SUBMIT_NOTE = "Note";
var SUBMIT_USER = "User";

var COL_STATUS = 2;
var COL_DW_NO = 11;
var COL_NOTE = 22;
var COL_CHECKER_BY = 17;
var COL_CHECKER_DATE = 18;
var COL_APPROVAL_BY = 19;
var COL_APPROVAL_DATE = 20;

function _cleanStr(s) {
  return String(s || '')
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .toLowerCase()
    .trim();
}

function _getNextStep(currentStatus) {
  var st = _cleanStr(currentStatus);

  if (st.indexOf('checker 1') !== -1 || st.indexOf('checker1') !== -1 ||
    st.indexOf('cho charger') !== -1 || st.indexOf('cho checker 1') !== -1 ||
    st.indexOf('cho duyet') !== -1) {
    return { nextStatus: 'Chờ Checker 2', byCol: COL_CHECKER_BY, dateCol: COL_CHECKER_DATE, level: 'Checker 1', appendMode: true };
  }

  if (st.indexOf('checker 2') !== -1 || st.indexOf('checker2') !== -1) {
    return { nextStatus: 'Chờ Approval', byCol: COL_CHECKER_BY, dateCol: COL_CHECKER_DATE, level: 'Checker 2', appendMode: true };
  }

  if (st.indexOf('approval') !== -1 || st.indexOf('duyet') !== -1 ||
    st.indexOf('trinh ky') !== -1 || st.indexOf('pending') !== -1) {
    return { nextStatus: 'Hoàn thành', byCol: COL_APPROVAL_BY, dateCol: COL_APPROVAL_DATE, level: 'Approval' };
  }

  return null;
}

function _getUserNameByEmail(email) {
  if (!email || email === 'System') return 'System';
  
  try {
    var ss = SpreadsheetApp.openById(SUBMIT_SS_ID);
    
    var userSheetName = (typeof USER_SHEET_NAME !== 'undefined') ? USER_SHEET_NAME : 'User';
    var userSheet = ss.getSheetByName(userSheetName);
    
    if (!userSheet) {
      Logger.log('Không tìm thấy sheet tên là: ' + userSheetName);
      return email; 
    }

    var data = userSheet.getDataRange().getValues();
    
    for (var i = 1; i < data.length; i++) {
      var rowMail = String(data[i][3]).trim().toLowerCase(); // Cột D (Mail) là index 3
      if (rowMail === email.toLowerCase()) {
        var rowName = String(data[i][2]).trim(); // Cột C (Name) là index 2
        return rowName || email;
      }
    }
  } catch (e) {
    Logger.log('Lỗi dò tên user: ' + e.message);
  }
  
  return email; 
}

function _findRowByIdInColA(sheet, drawingId) {
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return -1;
  var ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
  for (var i = 0; i < ids.length; i++) {
    if (String(ids[i][0]).trim() === String(drawingId).trim()) return i + 2;
  }
  return -1;
}

function _writeNote(ss, notiText, userEmail) {
  var noteSheet = ss.getSheetByName(SUBMIT_NOTE);
  if (!noteSheet) return;
  var profile = getUserProfile();
  var now = new Date();
  var tz = Session.getScriptTimeZone();
  noteSheet.appendRow([
    userEmail || profile.email,
    profile.name,
    notiText,
    Utilities.formatDate(now, tz, 'yyyy/MM/dd'),
    Utilities.formatDate(now, tz, 'HH:mm:ss')
  ]);
}

function approveDrawingOnServer(drawingId, approvedLevel) {
  if (!drawingId) throw new Error('Thieu ID ban ve.');

  var ss = SpreadsheetApp.openById(SUBMIT_SS_ID);
  var sheet = ss.getSheetByName(SUBMIT_SHEET);
  if (!sheet) throw new Error('Khong tim thay sheet "' + SUBMIT_SHEET + '".');

  var sheetRow = _findRowByIdInColA(sheet, drawingId);
  if (sheetRow === -1) {
    throw new Error('Khong tim thay ban ve co ID: ' + drawingId + '. Du lieu co the da bi thay doi, vui long tai lai trang.');
  }

  var currentStatus = String(sheet.getRange(sheetRow, COL_STATUS + 1).getValue() || '');
  var step = null;

  if (approvedLevel === 'Checker 1') {
    step = { nextStatus: 'Chờ Checker 2', byCol: COL_CHECKER_BY, dateCol: COL_CHECKER_DATE, level: 'Checker 1', appendMode: true };
  } else if (approvedLevel === 'Checker 2') {
    step = { nextStatus: 'Chờ Approval', byCol: COL_CHECKER_BY, dateCol: COL_CHECKER_DATE, level: 'Checker 2', appendMode: true };
  } else if (approvedLevel === 'Approval' || approvedLevel === 'Approver') {
    step = { nextStatus: 'Chờ ban hành', byCol: COL_APPROVAL_BY, dateCol: COL_APPROVAL_DATE, level: 'Approval', appendMode: false };
  } else {
    step = _getNextStep(currentStatus);
  }

  if (!step) {
    throw new Error('Ban ve dang o trang thai "' + currentStatus + '" - khong xac dinh duoc buoc trinh ky tiep theo.');
  }

  var approverEmail = Session.getActiveUser().getEmail() || 'System'
  var approverName = _getUserNameByEmail(approverEmail);
  var timezone = Session.getScriptTimeZone();
  var nowStr = Utilities.formatDate(new Date(), timezone, 'dd/MM/yyyy');

  sheet.getRange(sheetRow, COL_STATUS + 1).setValue(step.nextStatus);

  if (step.appendMode) {
    var oldBy = String(sheet.getRange(sheetRow, step.byCol + 1).getValue() || '');
    sheet.getRange(sheetRow, step.byCol + 1).setValue(oldBy ? oldBy + '\n' + approverName : approverName);
    var oldDate = String(sheet.getRange(sheetRow, step.dateCol + 1).getValue() || '');
    sheet.getRange(sheetRow, step.dateCol + 1).setValue(oldDate ? oldDate + '\n' + nowStr : nowStr);
  } else {
    sheet.getRange(sheetRow, step.byCol + 1).setValue(approverName);
    sheet.getRange(sheetRow, step.dateCol + 1).setValue(nowStr);
  }

  var dwNo = String(sheet.getRange(sheetRow, COL_DW_NO + 1).getValue() || drawingId);
  var notiMsg = '[TRÌNH KÝ] Bản vẽ: ' + dwNo +
    ' | Cấp: ' + step.level +
    ' | Trạng thái mới: ' + step.nextStatus +
    ' | Người ký: ' + approverName +
    ' | Lúc: ' + nowStr;
  _writeNote(ss, notiMsg, approverName);

  Logger.log('[approveDrawingOnServer] ID=' + drawingId + ' | ' + currentStatus + ' -> ' + step.nextStatus + ' | By=' + approverName);

  return {
    success: true,
    drawingId: drawingId,
    prevStatus: currentStatus,
    nextStatus: step.nextStatus,  
    level: step.level,
    approvedBy: approverName,
    approvedAt: nowStr
  };
}

function rejectDrawingOnServer(drawingId, reason) {
  if (!drawingId) throw new Error('Thieu ID ban ve.');
  if (!reason || !reason.trim()) throw new Error('Vui long nhap ly do tu choi.');

  var ss = SpreadsheetApp.openById(SUBMIT_SS_ID);
  var sheet = ss.getSheetByName(SUBMIT_SHEET);
  if (!sheet) throw new Error('Khong tim thay sheet "' + SUBMIT_SHEET + '".');

  var sheetRow = _findRowByIdInColA(sheet, drawingId);
  if (sheetRow === -1) {
    throw new Error('Khong tim thay ban ve co ID: ' + drawingId);
  }

  var rejectorEmail = Session.getActiveUser().getEmail() || 'System';
  var timezone = Session.getScriptTimeZone();
  var nowStr = Utilities.formatDate(new Date(), timezone, 'dd/MM/yyyy HH:mm');

  var prevStatus = String(sheet.getRange(sheetRow, COL_STATUS + 1).getValue() || '');
  sheet.getRange(sheetRow, COL_STATUS + 1).setValue('Tra lai - ' + prevStatus);

  var oldNote = String(sheet.getRange(sheetRow, COL_NOTE + 1).getValue() || '');
  sheet.getRange(sheetRow, COL_NOTE + 1).setValue(
    '[' + nowStr + ' - ' + rejectorEmail + '] TU CHOI: ' + reason.trim() +
    (oldNote ? '\n' + oldNote : '')
  );

  var dwNo = String(sheet.getRange(sheetRow, COL_DW_NO + 1).getValue() || drawingId);
  var notiMsg = '[REJECT] Bản vẽ: ' + dwNo +
    ' | Trạng thái cũ: ' + prevStatus +
    ' | Người từ chối: ' + rejectorEmail +
    ' | Lúc: ' + nowStr +
    ' | Lý do: ' + reason.trim();
  _writeNote(ss, notiMsg, rejectorEmail);

  Logger.log('[rejectDrawingOnServer] ID=' + drawingId + ' | By=' + rejectorEmail + ' | Reason=' + reason);

  return { success: true };
}

