var SUBMIT_SS_ID = "1DRteBSFT1cj4R_OUPMoDxeLMzAIJexWF3HPT-rpMOoM";
var SUBMIT_SS_BH = "1t5PWyoJHrxElWP3QgmB16BEMIvEC015NHq0tsxu_TpE";
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

function generateRandomString(length) {
  const chars = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789';
  let result = '';
  for (let i = 0; i < length; i++) {
    result += chars.charAt(Math.floor(Math.random() * chars.length));
  }
  return result;
}

function _toDateKey(dateValue) {
  var d = dateValue instanceof Date ? dateValue : new Date(dateValue);
  if (isNaN(d.getTime())) return '';
  var y = d.getFullYear();
  var m = String(d.getMonth() + 1).padStart(2, '0');
  var day = String(d.getDate()).padStart(2, '0');
  return y + '-' + m + '-' + day;
}

function _loadCalendarRows() {
  try {
    var ss = SpreadsheetApp.openById(SUBMIT_SS_ID);
    var calSheet = ss.getSheetByName('Cal');
    if (!calSheet) return [];
    var values = calSheet.getDataRange().getValues();
    return values.slice(1);
  } catch (e) {
    Logger.log('[Cal] Lỗi đọc sheet Cal: ' + e.message);
    return [];
  }
}

function _isCalendarOffDate(dateKey) {
  if (!dateKey) return false;

  var calendarRows = _loadCalendarRows();
  for (var i = 0; i < calendarRows.length; i++) {
    var row = calendarRows[i];
    var rowDate = String(row[0] || '').trim();
    if (!rowDate) continue;

    if (String(rowDate).slice(0, 10) === dateKey) {
      var status = String(row[2] || '').toLowerCase();
      var desc = String(row[3] || '').toLowerCase();
      var text = status + ' ' + desc;
      if (text.indexOf('weekend') !== -1 || text.indexOf('holiday') !== -1 || text.indexOf('nghi') !== -1 || text.indexOf('off') !== -1 || text.indexOf('rest') !== -1 || text.indexOf('sunday') !== -1 || text.indexOf('saturday') !== -1) {
        return true;
      }
      return false;
    }
  }

  var d = new Date(dateKey + 'T00:00:00');
  if (isNaN(d.getTime())) return false;
  var dayNum = d.getDay();
  return dayNum === 0 || dayNum === 6;
}

function getReleaseDueDate(requestDate, requestType) {
  var type = String(requestType || 'Normal').toLowerCase();
  var businessDays = (type === 'urgent') ? 3 : 7;
  var startDate = requestDate ? new Date(requestDate + 'T00:00:00') : new Date();
  if (isNaN(startDate.getTime())) {
    startDate = new Date();
  }

  var workingDays = 0;
  var cursor = new Date(startDate.getTime());
  var maxLoops = 366;

  while (workingDays < businessDays && maxLoops > 0) {
    cursor.setDate(cursor.getDate() + 1);
    var key = _toDateKey(cursor);
    if (_isCalendarOffDate(key)) {
      maxLoops--;
      continue;
    }
    workingDays++;
    maxLoops--;
  }

  return _toDateKey(cursor);
}

function sendMailFromDeptToCharge(dataToNotify, deptName) {
  const sheetUser = SpreadsheetApp.openById("1t5PWyoJHrxElWP3QgmB16BEMIvEC015NHq0tsxu_TpE").getSheetByName("User");
  const userLastRow = sheetUser.getLastRow();
  if (userLastRow < 2) return;

  const dataUser = sheetUser.getRange(2, 1, userLastRow - 1, 5).getValues();

  const targetUsers = dataUser.filter(u =>
    String(u[1]).trim().toUpperCase() === "QA" &&
    String(u[4]).trim().toUpperCase() === "CHARGER"
  );

  if (targetUsers.length === 0) {
    console.log("Không tìm thấy Charger nào thuộc bộ phận QA để gửi mail.");
    return;
  }

  const emailList = targetUsers.map(u => String(u[3]).trim()).join(",");
  const requests = {};

  dataToNotify.forEach(item => {
    if (!requests[item.requestId]) {
      requests[item.requestId] = {
        dept: item.dept || "Không rõ",
        name: item.name || "Không rõ",
        dwList: []
      };
    }
    if (!requests[item.requestId].dwList.includes(item.dw)) {
      requests[item.requestId].dwList.push(item.dw);
    }
  });

  let body = "Xin chào,\n\n";
  body += "Có yêu cầu phát hành bản vẽ mới từ bộ phận " + deptName + ".\n\n";

  Object.keys(requests).forEach(reqId => {
    const reqInfo = requests[reqId];
    const dwString = reqInfo.dwList.join(", ");
    body += "Mã yêu cầu: " + reqId + "\n";
    body += "Người gửi: " + reqInfo.name + " (Bộ phận: " + reqInfo.dept + ")\n";
    body += "Danh sách bản vẽ: " + dwString + "\n";
    body += "Trạng thái: Chờ ban hành\n\n";
  });

  body += "Vui lòng truy cập hệ thống để xử lý yêu cầu.\n\n";
  body += "Trân trọng,\n";
  body += "Hệ thống quản lý bản vẽ";

  const rowsHtml = Object.keys(requests).map(reqId => {
    const reqInfo = requests[reqId];
    const dwString = reqInfo.dwList.join(", ");
    return `
      <tr>
        <td style="padding: 12px 14px; border: 1px solid #e5e7eb; font-size: 13px; color: #374151;">${reqId}</td>
        <td style="padding: 12px 14px; border: 1px solid #e5e7eb; font-size: 13px; color: #374151;">${reqInfo.name}</td>
        <td style="padding: 12px 14px; border: 1px solid #e5e7eb; font-size: 13px; color: #374151;">${reqInfo.dept}</td>
        <td style="padding: 12px 14px; border: 1px solid #e5e7eb; font-size: 13px; color: #374151;">${dwString}</td>
      </tr>`;
  }).join("");

  const htmlBody = `
    <div style="font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; background: #f4f7fb; padding: 24px;">
      <div style="max-width: 760px; margin: 0 auto; background: #ffffff; border-radius: 14px; overflow: hidden; box-shadow: 0 6px 18px rgba(15, 23, 42, 0.08); border: 1px solid #e5e7eb;">
        <div style="background: linear-gradient(135deg, #0f172a 0%, #1e3a8a 100%); padding: 28px 24px; text-align: center; color: #ffffff;">
          <div style="font-size: 11px; letter-spacing: 1.4px; text-transform: uppercase; opacity: 0.8; margin-bottom: 8px;">Hệ thống quản lý bản vẽ</div>
          <h2 style="margin: 0; font-size: 24px; font-weight: 700;">Yêu cầu phát hành bản vẽ mới</h2>
          <div style="margin-top: 10px; font-size: 13px; color: #dbeafe;">Bộ phận: <strong style="color: #facc15;">${deptName}</strong></div>
        </div>

        <div style="padding: 24px;">
          <p style="margin: 0 0 18px; font-size: 15px; line-height: 1.7; color: #374151;">
            Xin chào,<br><br>
            Có <strong>yêu cầu phát hành bản vẽ mới</strong> cần bạn xử lý trong hệ thống QA.
          </p>

          <div style="background: #f8fafc; border-left: 4px solid #2563eb; padding: 14px 16px; border-radius: 8px; margin-bottom: 18px;">
            <strong style="color: #1d4ed8; font-size: 14px;">Thông tin chi tiết</strong>
          </div>

          <table style="width: 100%; border-collapse: collapse; background: #ffffff; border: 1px solid #e5e7eb; border-radius: 10px; overflow: hidden;">
            <thead>
              <tr style="background: #eff6ff;">
                <th style="padding: 12px 14px; text-align: left; border: 1px solid #e5e7eb; font-size: 12px; text-transform: uppercase; color: #1e3a8a;">Mã yêu cầu</th>
                <th style="padding: 12px 14px; text-align: left; border: 1px solid #e5e7eb; font-size: 12px; text-transform: uppercase; color: #1e3a8a;">Người gửi</th>
                <th style="padding: 12px 14px; text-align: left; border: 1px solid #e5e7eb; font-size: 12px; text-transform: uppercase; color: #1e3a8a;">Bộ phận</th>
                <th style="padding: 12px 14px; text-align: left; border: 1px solid #e5e7eb; font-size: 12px; text-transform: uppercase; color: #1e3a8a;">Danh sách bản vẽ</th>
              </tr>
            </thead>
            <tbody>
              ${rowsHtml}
            </tbody>
          </table>

          <div style="margin-top: 22px; text-align: center;">
            <a href="https://script.google.com/a/macros/listing.com/s/AKfycbw2MPLHbLNrNn3PhvU0Zr7V8D1-ouzWVTQDxW/usercache" style="display: inline-block; background: linear-gradient(135deg, #2563eb 0%, #1d4ed8 100%); color: #ffffff; padding: 12px 28px; border-radius: 8px; text-decoration: none; font-weight: 600; font-size: 14px;">
              → Vào hệ thống xử lý
            </a>
          </div>

          <div style="margin-top: 22px; border-top: 1px solid #e5e7eb; padding-top: 18px; font-size: 13px; color: #6b7280; line-height: 1.8;">
            Trạng thái: <strong style="color: #0f172a;">Chờ ban hành</strong><br>
            Vui lòng truy cập hệ thống để xác nhận và xử lý yêu cầu.<br><br>
            Trân trọng,<br>
            <strong>Hệ thống quản lý bản vẽ</strong>
          </div>
        </div>
      </div>
    </div>
  `;

  GmailApp.sendEmail(
    emailList,
    "[THÔNG BÁO] Có Yêu Cầu Phát Hành Bản Vẽ Lên QA System",
    body,
    { htmlBody: htmlBody }
  );
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

  if (st.indexOf('cho ban hanh') !== -1 || st.indexOf('ban hanh') !== -1 ||
    st.indexOf('dang ban hanh') !== -1 || st.indexOf('released') !== -1 ||
    st.indexOf('hoan thanh') !== -1) {
    return { nextStatus: 'Hoàn thành', byCol: COL_APPROVAL_BY, dateCol: COL_APPROVAL_DATE, level: 'Approval' };
  }

  if (st.indexOf('approval') !== -1 || st.indexOf('duyet') !== -1 ||
    st.indexOf('trinh ky') !== -1 || st.indexOf('pending') !== -1) {
    return { nextStatus: 'Chờ ban hành', byCol: COL_APPROVAL_BY, dateCol: COL_APPROVAL_DATE, level: 'Approval' };
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

function _normalizeRoleName(value) {
  return String(value || '')
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .replace(/[^a-zA-Z0-9]/g, '')
    .toLowerCase();
}

function _getUserEmailByRole(roleNames) {
  if (!roleNames) return '';

  var roles = Array.isArray(roleNames) ? roleNames : [roleNames];
  var normalizedRoles = [];
  for (var i = 0; i < roles.length; i++) {
    normalizedRoles.push(_normalizeRoleName(roles[i]));
  }

  try {
    var ss = SpreadsheetApp.openById(SUBMIT_SS_ID);
    var userSheet = ss.getSheetByName(SUBMIT_USER || 'User');
    if (!userSheet) return '';

    var data = userSheet.getDataRange().getValues();
    for (var r = 1; r < data.length; r++) {
      var mail = String(data[r][3] || '').trim();
      if (!mail) continue;

      var positionValue = String(data[r][4] || '');
      var normalizedPosition = _normalizeRoleName(positionValue);
      if (normalizedRoles.indexOf(normalizedPosition) !== -1) {
        return mail;
      }
    }
  } catch (e) {
    Logger.log('Lỗi dò email theo role: ' + e.message);
  }

  return '';
}

function _sendApprovalWorkflowEmail(targetEmail, dwNo, currentLevel, nextStatus, actorName) {
  if (!targetEmail) {
    Logger.log('[ApprovalEmail] Không có email đích cho level: ' + currentLevel);
    return false;
  }

  var subject = '[TRÌNH KÝ] Yêu cầu xác nhận bản vẽ ' + dwNo + ' - ' + nextStatus;
  var bodyText = '';
  var titleText = '';

  if (currentLevel === 'Checker 1') {
    bodyText = 'Bản vẽ ' + dwNo + ' vừa được ' + actorName + ' ký duyệt ở cấp Checker 1.\n'
      + 'Hiện tại cần bạn xác nhận ở cấp Checker 2 để tiếp tục quy trình.';
    titleText = 'Xác Nhận Bản Vẽ - Cấp Checker 2';
  } else if (currentLevel === 'Checker 2') {
    bodyText = 'Bản vẽ ' + dwNo + ' đã được ' + actorName + ' xác nhận ở cấp Checker 2.\n'
      + 'Hiện tại cần bạn ký duyệt ở cấp Approval để tiếp tục quy trình.';
    titleText = 'Phê Duyệt Bản Vẽ - Cấp Approval';
  } else if (currentLevel === 'Approval') {
    bodyText = 'Bản vẽ ' + dwNo + ' đã được ' + actorName + ' ký duyệt ở cấp Approval.\n'
      + 'Đã chuyển sang trạng thái Chờ ban hành và đang chờ Charge xử lý ban hành.';
    titleText = 'Bản Vẽ Đã Phê Duyệt - Chờ Ban Hành';
  } else {
    bodyText = 'Bản vẽ ' + dwNo + ' đã cập nhật trạng thái: ' + nextStatus;
    titleText = 'Cập Nhật Trạng Thái Bản Vẽ';
  }

  var htmlBody = `
    <div style="font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; background: #f5f7fa; padding: 20px; max-width: 700px; margin: 0 auto; color: #333;">
      <!-- HEADER -->
      <div style="background: linear-gradient(135deg, #1a3a52 0%, #2d5a7b 100%); border-radius: 12px 12px 0 0; padding: 30px 25px; text-align: center; color: white; box-shadow: 0 4px 6px rgba(0,0,0,0.1);">
        <div style="font-size: 12px; font-weight: 600; letter-spacing: 1px; margin-bottom: 12px; opacity: 0.9; text-transform: uppercase;">Hệ thống quản lý bản vẽ</div>
        <h1 style="margin: 0 0 8px 0; font-size: 22px; font-weight: 700; letter-spacing: 0.5px;">${titleText}</h1>
        <div style="font-size: 13px; font-weight: 500; color: #b0d0f0; margin-top: 8px;">Bản vẽ: <span style="color: #ffc107; font-weight: 700;">${dwNo}</span></div>
      </div>

      <!-- MAIN CONTENT -->
      <div style="background: white; padding: 32px 25px; border-radius: 0 0 12px 12px; box-shadow: 0 4px 15px rgba(0,0,0,0.08);">
        <p style="margin: 0 0 20px 0; font-size: 15px; line-height: 1.6; color: #555;">
          <strong>Xin chào,</strong>
        </p>

        <!-- STATUS INFO BOX -->
        <div style="background: #f0f7ff; border-left: 4px solid #0d6efd; padding: 18px; border-radius: 6px; margin-bottom: 24px;">
          <p style="margin: 0; font-size: 14px; line-height: 1.7; color: #0d6efd;">
            <strong>ℹ Thông tin cần xử lý:</strong><br>
            ${bodyText.split('\n').join('<br>')}
          </p>
        </div>

        <!-- INFO TABLE -->
        <table style="width: 100%; border-collapse: collapse; background: #f8f9fa; margin-bottom: 24px;">
          <tbody>
            <tr>
              <td style="padding: 12px 15px; font-weight: 600; background: #e8eef5; width: 35%; border-bottom: 1px solid #ddd; border-right: 1px solid #ddd;">Mã bản vẽ</td>
              <td style="padding: 12px 15px; border-bottom: 1px solid #ddd;"><strong style="color: #0d6efd; font-size: 15px;">${dwNo}</strong></td>
            </tr>
            <tr>
              <td style="padding: 12px 15px; font-weight: 600; background: #e8eef5; border-right: 1px solid #ddd;">Cấp xét duyệt</td>
              <td style="padding: 12px 15px; border-bottom: 1px solid #ddd;"><strong>${currentLevel}</strong></td>
            </tr>
            <tr>
              <td style="padding: 12px 15px; font-weight: 600; background: #e8eef5; border-right: 1px solid #ddd;">Trạng thái tiếp theo</td>
              <td style="padding: 12px 15px; border-bottom: 1px solid #ddd;">
                <span style="display: inline-block; background: #fff3cd; color: #856404; padding: 6px 12px; border-radius: 20px; font-weight: 600; font-size: 13px;">${nextStatus}</span>
              </td>
            </tr>
            <tr>
              <td style="padding: 12px 15px; font-weight: 600; background: #e8eef5; border-right: 1px solid #ddd;">Người xử lý</td>
              <td style="padding: 12px 15px;"><strong>${actorName}</strong></td>
            </tr>
          </tbody>
        </table>

        <!-- CALL TO ACTION -->
        <div style="text-align: center; margin: 28px 0;">
          <a href="https://script.google.com/a/macros/listing.com/s/AKfycbw2MPLHbLNrNn3PhvU0Zr7V8D1-ouzWVTQDxW/usercache" style="display: inline-block; background: linear-gradient(135deg, #0d6efd 0%, #0056b3 100%); color: white; padding: 12px 32px; text-decoration: none; border-radius: 6px; font-weight: 600; font-size: 14px; box-shadow: 0 4px 8px rgba(13,110,253,0.3); transition: all 0.3s;">
            → Vào hệ thống xử lý
          </a>
        </div>

        <!-- FOOTER -->
        <div style="border-top: 1px solid #e0e0e0; margin-top: 32px; padding-top: 20px; text-align: center;">
          <p style="margin: 0 0 8px 0; font-size: 13px; color: #999; line-height: 1.5;">
            <strong>Trân trọng,</strong><br>
            <span style="color: #0d6efd; font-weight: 600;">Hệ thống Web App Quản Lý Bản Vẽ</span>
          </p>
          <p style="margin: 12px 0 0 0; font-size: 11px; color: #bbb;">
            Email này được gửi tự động từ hệ thống. Vui lòng không trả lời trực tiếp.
          </p>
        </div>
      </div>

      <!-- SPACING -->
      <div style="height: 20px;"></div>
    </div>
  `;

  try {
    GmailApp.sendEmail(targetEmail, subject, "", { htmlBody: htmlBody });
    Logger.log('[ApprovalEmail] Gửi email thành công tới ' + targetEmail + ' cho bản vẽ ' + dwNo);
    return true;
  } catch (e) {
    Logger.log('[ApprovalEmail] Lỗi gửi email tới ' + targetEmail + ': ' + e.message);
    return false;
  }
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

  // Gửi email thông báo theo luồng ký duyệt
  var notifyEmail = '';
  if (step.level === 'Checker 1') {
    notifyEmail = _getUserEmailByRole(['Checker 2', 'checker2']);
  } else if (step.level === 'Checker 2') {
    notifyEmail = _getUserEmailByRole(['Approval', 'Approver', 'GMQA', 'QA', 'Approval QA']);
  } else if (step.level === 'Approval') {
    notifyEmail = _getUserEmailByRole(['Charger', 'Charge', 'QA Charge', 'Charge QA']);
  }

  if (notifyEmail) {
    _sendApprovalWorkflowEmail(notifyEmail, dwNo, step.level, step.nextStatus, approverName);
  } else {
    Logger.log('[approveDrawingOnServer] Không tìm thấy email đích để thông báo: level=' + step.level + ', dw=' + dwNo);
  }

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








function  requestReleaseDW(arr) {
  const ss = SpreadsheetApp.openById("1t5PWyoJHrxElWP3QgmB16BEMIvEC015NHq0tsxu_TpE");
  const sheet = ss.getSheetByName("data");

  const sheetSoure = SpreadsheetApp.openById("1DRteBSFT1cj4R_OUPMoDxeLMzAIJexWF3HPT-rpMOoM").getSheetByName("Data");

  const lastrow = sheet.getLastRow();
  const lastcol = sheet.getLastColumn();
  let data = [];

  if (lastrow > 1) {
    data = sheet.getRange(2, 1, lastrow - 1, lastcol).getValues();
  }

  const existingIds = new Set();
  data.forEach(row => {
    if(row[0]) existingIds.add(String(row[0]));
  });

  const requestType = String((arr && arr[2]) || 'Normal');
  const requestDateRaw = String((arr && arr[4]) || new Date().toISOString().slice(0, 10));
  const computedDueDate = getReleaseDueDate(requestDateRaw, requestType);
  const finalDueDate = String((arr && arr[5]) || computedDueDate);
  const releaseReason = String((arr && arr[3]) || 'New Issue');

  const inputDate = new Date(requestDateRaw);
  const currentMonth = inputDate.getMonth() + 1;
  const currentFullYear = inputDate.getFullYear();

  const strMonth = currentMonth < 10 ? "0" + currentMonth : currentMonth;
  const strYear = currentFullYear.toString().slice(-2);
  const orderCheck = "WI-DW-" + strMonth + strYear + "-";

  let check = [];
  data.forEach(row => {
    if (row[8]) {
      const idStr = String(row[8]);
      if (idStr.startsWith(orderCheck)) {
        let numberPart = parseInt(idStr.slice(-3), 10);
        if (!isNaN(numberPart)) {
             check.push(numberPart);
        }
      }
    }
  });

  let maxNum = 0;
  if (check.length > 0) {
    maxNum = Math.max(...check);
  }

  const newCount = maxNum + 1;
  const strCount = newCount.toString().padStart(3, '0');
  const orderNo = "WI-DW-" + strMonth + strYear + "-" + strCount;
  let mailValue = [];

  const list = arr[9];
  let output = [];
  let ids = [];
  list.forEach(row => {
    let newID;
    do {
      newID = generateRandomString(6);
    } while (existingIds.has(newID));
    ids.push(row[0]);

    existingIds.add(newID);
     mailValue.push({
              requestId: orderNo,
              dw: row[1],
              status:"Yêu Cầu Ban Hàng Bản Vẽ Mới",
              dept: arr[0],
              name: arr[1],
            });

    output.push([
      newID,
      "BUNDLING & PACKING",
      row[4],
      "",
      row[1],
      row[2],
      row[3],
      "",
      orderNo,
      "",
      "QA-G2G",
      "",
      arr[1],
      requestDateRaw,
      requestType === 'Urgent' ? 'Urgent' : 'Normal',
      finalDueDate,
      "",
      row[5],
      "",
      requestType === 'Urgent' ? '' : releaseReason,
      "",
      "",
      "",
      "Đã tạo",
    ]);
  });

  if (output.length > 0) {
    sheet.getRange(lastrow + 1, 1, output.length, output[0].length).setValues(output);
    sendMailFromDeptToCharge(mailValue, arr[0]);
  }
  if(ids.length>0){
    updateBulkStatusByIds(ids,"Đang ban hành")
  }

  return "Đã tạo thành công: " + orderNo + " | Hạn hoàn thành: " + finalDueDate;
}




function updateBulkStatusByIds(targetIds, stt) {
  // Đảm bảo targetIds là một mảng. Ví dụ: ["ID001", "ID002", "ID003"]
  if (!Array.isArray(targetIds) || targetIds.length === 0) return;

  const sheet = SpreadsheetApp.openById("1DRteBSFT1cj4R_OUPMoDxeLMzAIJexWF3HPT-rpMOoM").getSheetByName("Data");
  const lastRow = sheet.getLastRow();
  if (lastRow < 1) return; // Thoát nếu sheet trống

  // 1. ĐỌC 1 LẦN: Chỉ lấy dữ liệu Cột A (ID) và Cột C (Status)
  // Lấy riêng biệt để tránh ghi đè làm mất công thức ở cột B (nếu có)
  const idValues = sheet.getRange(1, 1, lastRow, 1).getValues(); // Đọc toàn bộ cột A
  const statusRange = sheet.getRange(1, 3, lastRow, 1); 
  const statusValues = statusRange.getValues(); // Đọc toàn bộ cột C
  
  // Chuyển mảng targetIds thành Set để tìm kiếm với tốc độ siêu tốc (O(1))
  const idSet = new Set(targetIds); 
  let hasChanges = false;

  // 2. XỬ LÝ TRONG RAM: Vòng lặp này chạy bằng tốc độ của CPU/RAM, gần như tức thời
  for (let i = 0; i < idValues.length; i++) {
    const currentId = idValues[i][0];
    
    // Nếu ID ở dòng hiện tại nằm trong mảng cần thay đổi
    if (idSet.has(currentId)) {
      statusValues[i][0] = stt; // Cập nhật trạng thái mới vào mảng
      hasChanges = true;
    }
  }

  // 3. GHI 1 LẦN: Đổ ngược mảng đã cập nhật vào lại cột C
  if (hasChanges) {
    statusRange.setValues(statusValues);
  }
}
