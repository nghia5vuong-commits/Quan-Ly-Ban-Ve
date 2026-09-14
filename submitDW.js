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

const _cleanStr = (s) => String(s || '')
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .toLowerCase()
    .trim();

const generateRandomString = (length) => {
  const chars = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789';
  let result = '';
  for (let i = 0; i < length; i++) {
    result += chars.charAt(Math.floor(Math.random() * chars.length));
  }
  return result;
};

const _toDateKey = (dateValue) => {
  var d = dateValue instanceof Date ? dateValue : new Date(dateValue);
  if (isNaN(d.getTime())) return '';
  var y = d.getFullYear();
  var m = String(d.getMonth() + 1).padStart(2, '0');
  var day = String(d.getDate()).padStart(2, '0');
  return y + '-' + m + '-' + day;
};

const _loadCalendarRows = () => {
  try {
    var ss = SpreadsheetApp.openById(SUBMIT_SS_ID);
    var calSheet = ss.getSheetByName('Cal');
    if (!calSheet) return [];
    var values = calSheet.getDataRange().getValues();
    if (!values || values.length < 2) return [];
    return values.slice(1);
  } catch (e) {
    return [];
  }
};

const _normalizeCalendarText = (value) => {
  return String(value || '')
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .toLowerCase()
    .trim();
}

function _isCalendarOffDate(dateKey) {
  if (!dateKey) return false;

  var calendarRows = _loadCalendarRows();
  for (var i = 0; i < calendarRows.length; i++) {
    var row = calendarRows[i];
    var rowDate = String(row[0] || '').trim();
    if (!rowDate) continue;

    if (String(rowDate).slice(0, 10) === dateKey) {
      var status = _normalizeCalendarText(row[2]);
      var description = _normalizeCalendarText(row[3]);
      var text = (status + ' ' + description).trim();

      var offPatterns = [
        'weekend', 'holiday', 'nghi', 'day off', 'off', 'rest', 'sunday', 'saturday',
        'khong lam', 'khong hoat dong', 'ngay nghi', 'leave', 'non working', 'non-working',
        'close', 'closed'
      ];

      for (var j = 0; j < offPatterns.length; j++) {
        if (text.indexOf(offPatterns[j]) !== -1) {
          return true;
        }
      }

      return false;
    }
  }

  var d = new Date(dateKey + 'T00:00:00');
  if (isNaN(d.getTime())) return false;
  var dayNum = d.getDay();
  return dayNum === 0 || dayNum === 6;
}

function _getBusinessDateAfterDays(dateValue, businessDays) {
  var requestedDays = Number(businessDays) || 5;
  var startDate = dateValue ? new Date(dateValue + 'T00:00:00') : new Date();
  if (isNaN(startDate.getTime())) {
    startDate = new Date();
  }

  var workingDays = 0;
  var cursor = new Date(startDate.getTime());
  var maxLoops = 366;

  while (workingDays < requestedDays && maxLoops > 0) {
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

const getBusinessDateFromToday = (workingDays) => _getBusinessDateAfterDays(new Date(), Number(workingDays) || 5);

const getReleaseDueDate = (requestDate, requestType) => {
  var type = String(requestType || 'Normal');
  var businessDays = /^normal$/i.test(type) ? 3 : 1;
  return _getBusinessDateAfterDays(requestDate || new Date(), businessDays);
};

const sendMailFromDeptToCharge = (dataToNotify, deptName) => {
  deptName = "QA G2G";
  const sheetUser = SpreadsheetApp.openById("1t5PWyoJHrxElWP3QgmB16BEMIvEC015NHq0tsxu_TpE").getSheetByName("User");
  const userLastRow = sheetUser.getLastRow();
  if (userLastRow < 2) return;

  const dataUser = sheetUser.getRange(2, 1, userLastRow - 1, 5).getValues();

  const targetUsers = dataUser.filter(u =>
    String(u[1]).trim().toUpperCase() === "QA" &&
    String(u[4]).trim().toUpperCase() === "CHARGER"
  );

  if (targetUsers.length === 0) {
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

  const reqId = Object.keys(requests)[0] || '';
  const reqInfo = Object.keys(requests).length > 0 ? requests[reqId] : { name: 'Không rõ', dept: deptName, dwList: [] };
  const dwString = reqInfo.dwList.join(", ");
  const webAppBaseUrl = 'https://script.google.com/a/macros/lixil.com/s/AKfycbyEBaqWQ_UHTu1eahFFx4xdcpteLV93DkLZM50nYGHzCcYVdeNfVuCxXxdRqPpiVHhXzA/exec';
  const requestWebUrl = reqId ? `${webAppBaseUrl}?requestId=${encodeURIComponent(reqId)}` : webAppBaseUrl;

  const htmlBody = `
    <div style="font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; background: #eef4f8; padding: 20px;">
      <div style="max-width: 780px; margin: 0 auto; background: #ffffff; border: 1px solid #d1d5db; border-radius: 4px; box-shadow: 0 4px 14px rgba(0,0,0,0.1); overflow: hidden;">
        <div style="background: linear-gradient(135deg, #13b8ae 0%, #0f8d93 100%); color: #ffffff; padding: 14px 18px; font-size: 24px; font-weight: 700; line-height: 1.4;">
          Yêu Cầu Phát Hành Bản Vẽ Lên QA System
        </div>

        <div style="padding: 22px 26px 26px; background: #ffffff;">
          <p style="margin: 0 0 12px; font-size: 16px; color: #1f2937; line-height: 1.6;">
            Xin chào Ms. Nhung,<br>
            Hệ thống ghi nhận có một yêu cầu bạn phát hành bản vẽ mới vừa được tạo từ bộ phận ${deptName}. Thông tin chi tiết được thể hiện dưới đây:
          </p>

          <div style="background: #eef4f6; border: 1px solid #cbd5e1; border-radius: 8px; padding: 10px 12px; margin: 12px 0 14px;">
            <table style="width: 100%; border-collapse: collapse;">
              <tr>
                <td style="font-size: 15px; color: #374151; font-weight: 600; text-align: left; vertical-align: middle;">
                  Mã yêu cầu (Request ID): <span style="font-weight: 700; color: #0f766e;">${reqId}</span>
                </td>
                <td style="text-align: right; vertical-align: middle;">
                  <a href="${requestWebUrl}" style="background: #2563eb; color: #ffffff; text-decoration: none; padding: 8px 14px; border-radius: 4px; font-size: 13px; font-weight: 700; white-space: nowrap; display: inline-block;">
                    Tới Web App →
                  </a>
                </td>
              </tr>
            </table>
          </div>

          <table style="width: 100%; border-collapse: collapse; border: 1px solid #d1d5db; background: #ffffff; margin-top: 10px;">
            <tr>
              <td style="width: 50%; padding: 10px 12px; border: 1px solid #d1d5db; font-size: 14px; color: #334155; background: #f8fafc; font-weight: 700;">
                Người gửi yêu cầu
              </td>
              <td style="width: 50%; padding: 10px 12px; border: 1px solid #d1d5db; font-size: 14px; color: #334155; background: #f8fafc; font-weight: 700;">
                Nghĩa (Bộ phận: ${deptName})
              </td>
            </tr>
            <tr>
              <td style="width: 50%; padding: 10px 12px; border: 1px solid #d1d5db; font-size: 14px; color: #334155; background: #ffffff; font-weight: 700;">
                Danh sách bản vẽ
              </td>
              <td style="width: 50%; padding: 10px 12px; border: 1px solid #d1d5db; font-size: 14px; color: #334155; background: #ffffff;">
                ${dwString}
              </td>
            </tr>
          </table>

          <p style="margin: 18px 0 8px; font-size: 15px; color: #374151; line-height: 1.7;">
            Vui lòng truy cập hệ thống để tiến hành kiểm tra và xử lý yêu cầu.
          </p>

          <div style="margin-top: 12px; font-size: 15px; color: #374151; line-height: 1.7;">
            <div>Trân trọng,</div>
            <div><strong>Hệ thống Web App Ban Hành Bản Vẽ</strong></div>
          </div>
        </div>
      </div>
    </div>`;

  GmailApp.sendEmail(
    emailList,
    "[THÔNG BÁO] Có Yêu Cầu Phát Hành Bản Vẽ Lên QA System",
    body,
    { htmlBody: htmlBody }
  );
};

const _getNextStep = (currentStatus) => {
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
  }
  
  return email; 
};

const _normalizeRoleName = (value) => {
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
  }

  return '';
}

function _escapeApprovalEmailHtml(value) {
  return String(value || '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

function _formatApprovalEmailDate(value) {
  if (!value) return '';
  var date = value instanceof Date ? value : new Date(value);
  if (isNaN(date.getTime())) {
    var raw = String(value).trim();
    var match = raw.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})/);
    if (match) date = new Date(Number(match[3]), Number(match[2]) - 1, Number(match[1]));
  }
  if (isNaN(date.getTime())) return String(value);
  return Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyy/MM/dd');
}

function _getApprovalWebUrl(drawingId, action) {
  var baseUrl = 'https://script.google.com/a/macros/lixil.com/s/AKfycbyEBaqWQ_UHTu1eahFFx4xdcpteLV93DkLZM50nYGHzCcYVdeNfVuCxXxdRqPpiVHhXzA/exec';
  var query = '?drawingId=' + encodeURIComponent(drawingId || '');
  if (action) query += '&action=' + encodeURIComponent(action);
  return baseUrl + query;
}

function _approvalEmailButton(url, label, background) {
  return '<a href="' + _escapeApprovalEmailHtml(url) + '" style="display:inline-block;margin:4px 4px;padding:11px 16px;background:' + background + ';color:#ffffff;text-decoration:none;border-radius:7px;font-weight:700;font-size:13px;">' + label + '</a>';
}

function _getApprovalDrawingDetails(sheet, sheetRow, drawingId) {
  var details = {
    id: drawingId,
    group: '',
    type: '',
    version: '',
    to: '',
    project: '',
    customer: '',
    dwNo: '',
    typeDw: '',
    assignee: '',
    receivedDate: '',
    dueDate: '',
    actualDoneDate: '',
    fye: '',
    width: '',
    height: '',
    so: '',
    excelUrl: '',
    pdfUrl: '',
    cutDrawingBlob: null
  };

  if (!sheet || !sheetRow) return details;

  var row = sheet.getRange(sheetRow, 1, 1, Math.max(34, sheet.getLastColumn())).getValues()[0];
  details.group = row[3] || '';
  details.type = row[4] || '';
  details.version = row[5] || '';
  details.receivedDate = row[6] || '';
  details.dueDate = row[7] || '';
  details.to = row[8] || '';
  details.project = row[9] || '';
  details.customer = row[10] || '';
  details.dwNo = row[11] || drawingId || '';
  details.typeDw = row[12] || row[24] || '';
  details.assignee = row[15] || '';
  details.actualDoneDate = row[18] || '';
  details.fye = row[23] || '';
  details.width = row[28] || '';
  details.height = row[29] || '';
  details.excelUrl = row[30] || '';
  details.pdfUrl = row[31] || row[14] || '';
  details.so = row[33] || '';

  try {
    var imageValue = row[32];
    if (imageValue && typeof imageValue.getBlob === 'function') {
      details.cutDrawingBlob = imageValue.getBlob();
    }
  } catch (error) {
    Logger.log('Không thể lấy hình mặt cắt cho email ký duyệt: ' + error.toString());
  }

  return details;
}

function _sendApprovalWorkflowEmail(targetEmail, dwNo, currentLevel, nextStatus, actorName, details) {
  if (!targetEmail) {
    return false;
  }

  details = details || {};
  var displayDwNo = details.dwNo || dwNo || '';
  var receivedDate = _formatApprovalEmailDate(details.receivedDate);
  var dueDate = _formatApprovalEmailDate(details.dueDate);
  var systemUrl = _getApprovalWebUrl(details.id || displayDwNo, 'view');
  var approvalButtons = _approvalEmailButton(_getApprovalWebUrl(details.id || displayDwNo, 'approve'), 'Approval', '#16a34a')
    + _approvalEmailButton(_getApprovalWebUrl(details.id || displayDwNo, 'reject'), 'Reject', '#dc2626');
  var drawingUrl = details.pdfUrl || systemUrl;
  var actionButtons = approvalButtons
    + _approvalEmailButton(drawingUrl, 'Xem bản vẽ', '#2563eb')
    + _approvalEmailButton(systemUrl, 'Vào hệ thống', '#475569');

  var subject = '[TRÌNH KÝ] Yêu cầu xác nhận bản vẽ ' + displayDwNo + ' - ' + nextStatus;
  var bodyText = '';
  var titleText = '';

  if (currentLevel === 'Checker 1') {
    bodyText = 'Bản vẽ ' + displayDwNo + ' vừa được ' + actorName + ' ký duyệt ở cấp Checker 1.\n'
      + 'Hiện tại cần bạn xác nhận ở cấp Checker 2 để tiếp tục quy trình.';
    titleText = 'Xác Nhận Bản Vẽ - Cấp Checker 2';
  } else if (currentLevel === 'Checker 2') {
    bodyText = 'Bản vẽ ' + displayDwNo + ' đã được ' + actorName + ' xác nhận ở cấp Checker 2.\n'
      + 'Hiện tại cần bạn ký duyệt ở cấp Approval để tiếp tục quy trình.';
    titleText = 'Phê Duyệt Bản Vẽ - Cấp Approval';
  } else if (currentLevel === 'Approval') {
    bodyText = 'Bản vẽ ' + displayDwNo + ' đã được ' + actorName + ' ký duyệt ở cấp Approval.\n'
      + 'Đã chuyển sang trạng thái Chờ ban hành và đang chờ Charge xử lý ban hành.';
    titleText = 'Bản Vẽ Đã Phê Duyệt - Chờ Ban Hành';
  } else {
    bodyText = 'Bản vẽ ' + displayDwNo + ' đã cập nhật trạng thái: ' + nextStatus;
    titleText = 'Cập Nhật Trạng Thái Bản Vẽ';
  }

  var imageHtml = details.cutDrawingBlob
    ? '<p style="margin:24px 0 8px;font-weight:700;color:#334155;">Hình mặt cắt</p><div style="padding:12px;text-align:center;background:#f8fafc;border:1px solid #dbeafe;border-radius:8px;"><img src="cid:approvalCutDrawing" style="max-width:100%;max-height:280px;object-fit:contain;border-radius:6px;" alt="Hình mặt cắt" /></div>'
    : '';
  var linkHtml = details.pdfUrl
    ? '<a href="' + _escapeApprovalEmailHtml(details.pdfUrl) + '" style="display:inline-block;margin:4px 6px 4px 0;color:#2563eb;">Mở file PDF</a>'
    : '';

  var htmlBody = `
    <div style="font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; background: #f5f7fa; padding: 20px; max-width: 700px; margin: 0 auto; color: #333;">
      <!-- HEADER -->
      <div style="background: linear-gradient(135deg, #1a3a52 0%, #2d5a7b 100%); border-radius: 12px 12px 0 0; padding: 30px 25px; text-align: center; color: white; box-shadow: 0 4px 6px rgba(0,0,0,0.1);">
        <div style="font-size: 12px; font-weight: 600; letter-spacing: 1px; margin-bottom: 12px; opacity: 0.9; text-transform: uppercase;">Hệ thống quản lý bản vẽ</div>
        <h1 style="margin: 0 0 8px 0; font-size: 22px; font-weight: 700; letter-spacing: 0.5px;">${titleText}</h1>
        <div style="font-size: 13px; font-weight: 500; color: #b0d0f0; margin-top: 8px;">Bản vẽ: <span style="color: #ffc107; font-weight: 700;">${_escapeApprovalEmailHtml(displayDwNo)}</span></div>
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
            ${_escapeApprovalEmailHtml(bodyText).split('\n').join('<br>')}
          </p>
        </div>

        <!-- INFO TABLE -->
        <table style="width: 100%; border-collapse: collapse; background: #f8f9fa; margin-bottom: 24px;">
          <tbody>
            <tr>
              <td style="padding: 12px 15px; font-weight: 600; background: #e8eef5; width: 35%; border-bottom: 1px solid #ddd; border-right: 1px solid #ddd;">Mã bản vẽ</td>
              <td style="padding: 12px 15px; border-bottom: 1px solid #ddd;"><strong style="color: #0d6efd; font-size: 15px;">${_escapeApprovalEmailHtml(displayDwNo)}</strong></td>
            </tr>
            <tr>
              <td style="padding: 12px 15px; font-weight: 600; background: #e8eef5; border-right: 1px solid #ddd;">Cấp xét duyệt</td>
              <td style="padding: 12px 15px; border-bottom: 1px solid #ddd;"><strong>${_escapeApprovalEmailHtml(currentLevel)}</strong></td>
            </tr>
            <tr>
              <td style="padding: 12px 15px; font-weight: 600; background: #e8eef5; border-right: 1px solid #ddd;">Trạng thái tiếp theo</td>
              <td style="padding: 12px 15px; border-bottom: 1px solid #ddd;">
                <span style="display: inline-block; background: #fff3cd; color: #856404; padding: 6px 12px; border-radius: 20px; font-weight: 600; font-size: 13px;">${_escapeApprovalEmailHtml(nextStatus)}</span>
              </td>
            </tr>
            <tr>
              <td style="padding: 12px 15px; font-weight: 600; background: #e8eef5; border-right: 1px solid #ddd;">Người xử lý</td>
              <td style="padding: 12px 15px;"><strong>${_escapeApprovalEmailHtml(actorName)}</strong></td>
            </tr>
          </tbody>
        </table>

        <table style="width:100%;border-collapse:collapse;background:#ffffff;margin-top:10px;">
          <tr><td style="padding:8px 10px;border:1px solid #d1d5db;background:#f8fafc;font-weight:700;width:35%;">TO</td><td style="padding:8px 10px;border:1px solid #d1d5db;">${_escapeApprovalEmailHtml(details.to)}</td></tr>
          <tr><td style="padding:8px 10px;border:1px solid #d1d5db;background:#f8fafc;font-weight:700;">Khách hàng</td><td style="padding:8px 10px;border:1px solid #d1d5db;">${_escapeApprovalEmailHtml(details.customer)}</td></tr>
          <tr><td style="padding:8px 10px;border:1px solid #d1d5db;background:#f8fafc;font-weight:700;">Dự án</td><td style="padding:8px 10px;border:1px solid #d1d5db;">${_escapeApprovalEmailHtml(details.project)}</td></tr>
          <tr><td style="padding:8px 10px;border:1px solid #d1d5db;background:#f8fafc;font-weight:700;">Type / Version</td><td style="padding:8px 10px;border:1px solid #d1d5db;">${_escapeApprovalEmailHtml(details.type)} / ${_escapeApprovalEmailHtml(details.version)}</td></tr>
          <tr><td style="padding:8px 10px;border:1px solid #d1d5db;background:#f8fafc;font-weight:700;">Type DW</td><td style="padding:8px 10px;border:1px solid #d1d5db;">${_escapeApprovalEmailHtml(details.typeDw)}</td></tr>
          <tr><td style="padding:8px 10px;border:1px solid #d1d5db;background:#f8fafc;font-weight:700;">Người đảm trách</td><td style="padding:8px 10px;border:1px solid #d1d5db;">${_escapeApprovalEmailHtml(details.assignee)}</td></tr>
          <tr><td style="padding:8px 10px;border:1px solid #d1d5db;background:#f8fafc;font-weight:700;">Ngày tiếp nhận / Deadline</td><td style="padding:8px 10px;border:1px solid #d1d5db;">${_escapeApprovalEmailHtml(receivedDate)} / ${_escapeApprovalEmailHtml(dueDate)}</td></tr>
          <tr><td style="padding:8px 10px;border:1px solid #d1d5db;background:#f8fafc;font-weight:700;">Width / Height</td><td style="padding:8px 10px;border:1px solid #d1d5db;">${_escapeApprovalEmailHtml(details.width)} / ${_escapeApprovalEmailHtml(details.height)} mm</td></tr>
          <tr><td style="padding:8px 10px;border:1px solid #d1d5db;background:#f8fafc;font-weight:700;">Số SO / FYE</td><td style="padding:8px 10px;border:1px solid #d1d5db;">${_escapeApprovalEmailHtml(details.so)} / ${_escapeApprovalEmailHtml(details.fye)}</td></tr>
        </table>
        <div style="margin-top:10px;">${linkHtml}</div>
        ${imageHtml}

        <!-- CALL TO ACTION -->
        <div style="text-align: center; margin: 28px 0;">
          ${actionButtons}
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
    var mailOptions = { htmlBody: htmlBody };
    if (details.cutDrawingBlob) mailOptions.inlineImages = { approvalCutDrawing: details.cutDrawingBlob };
    GmailApp.sendEmail(targetEmail, subject, bodyText, mailOptions);
    return true;
  } catch (e) {
    return false;
  }
};

function _sendRejectWorkflowEmail(targetEmail, dwNo, rejectorName, previousStatus, reason, details) {
  if (!targetEmail) return false;

  details = details || {};
  var displayDwNo = details.dwNo || dwNo || '';
  var rejectDate = _formatApprovalEmailDate(new Date());
  var subject = '[REJECT] Bản vẽ ' + displayDwNo + ' đã bị trả về';
  var bodyText = 'Bản vẽ ' + displayDwNo + ' đã bị ' + rejectorName + ' từ chối.\n'
    + 'Lý do: ' + reason + '\n'
    + 'Vui lòng truy cập hệ thống để kiểm tra và xử lý lại.';
  var systemUrl = _getApprovalWebUrl(details.id || displayDwNo, 'view');
  var actionButtons = _approvalEmailButton(details.pdfUrl || systemUrl, 'Xem bản vẽ', '#2563eb');
  actionButtons += _approvalEmailButton(systemUrl, 'Vào hệ thống', '#475569');
  var imageHtml = details.cutDrawingBlob
    ? '<p style="margin:24px 0 8px;font-weight:700;color:#334155;">Hình mặt cắt</p><div style="padding:12px;text-align:center;background:#f8fafc;border:1px solid #fecaca;border-radius:8px;"><img src="cid:rejectCutDrawing" style="max-width:100%;max-height:280px;object-fit:contain;border-radius:6px;" alt="Hình mặt cắt" /></div>'
    : '';

  var htmlBody = '<div style="font-family:Segoe UI,Tahoma,sans-serif;background:#f5f7fa;padding:20px;color:#1f2937;">'
    + '<div style="max-width:700px;margin:0 auto;background:#fff;border-radius:12px;overflow:hidden;box-shadow:0 4px 15px rgba(0,0,0,.08);">'
    + '<div style="background:linear-gradient(135deg,#991b1b,#dc2626);padding:28px 25px;color:#fff;text-align:center;">'
    + '<div style="font-size:12px;font-weight:700;letter-spacing:1px;text-transform:uppercase;opacity:.9;">Hệ thống quản lý bản vẽ</div>'
    + '<h1 style="margin:10px 0 6px;font-size:22px;">Bản vẽ bị từ chối</h1>'
    + '<div style="font-size:14px;">Mã bản vẽ: <strong>' + _escapeApprovalEmailHtml(displayDwNo) + '</strong></div></div>'
    + '<div style="padding:28px 25px;">'
    + '<p style="font-size:15px;line-height:1.7;margin-top:0;">Xin chào,<br>Bản vẽ cần được kiểm tra và xử lý lại theo thông tin dưới đây.</p>'
    + '<div style="background:#fff1f2;border-left:4px solid #dc2626;padding:16px;border-radius:6px;margin:18px 0;line-height:1.7;">'
    + '<strong style="color:#991b1b;">Lý do từ chối</strong><br>' + _escapeApprovalEmailHtml(reason) + '</div>'
    + '<table style="width:100%;border-collapse:collapse;margin-bottom:18px;">'
    + '<tr><td style="padding:9px 10px;border:1px solid #d1d5db;background:#f8fafc;font-weight:700;width:35%;">Trạng thái trước</td><td style="padding:9px 10px;border:1px solid #d1d5db;">' + _escapeApprovalEmailHtml(previousStatus) + '</td></tr>'
    + '<tr><td style="padding:9px 10px;border:1px solid #d1d5db;background:#f8fafc;font-weight:700;">Người từ chối</td><td style="padding:9px 10px;border:1px solid #d1d5db;">' + _escapeApprovalEmailHtml(rejectorName) + '</td></tr>'
    + '<tr><td style="padding:9px 10px;border:1px solid #d1d5db;background:#f8fafc;font-weight:700;">Ngày xử lý</td><td style="padding:9px 10px;border:1px solid #d1d5db;">' + rejectDate + '</td></tr>'
    + '<tr><td style="padding:9px 10px;border:1px solid #d1d5db;background:#f8fafc;font-weight:700;">TO / Khách hàng</td><td style="padding:9px 10px;border:1px solid #d1d5db;">' + _escapeApprovalEmailHtml(details.to) + ' / ' + _escapeApprovalEmailHtml(details.customer) + '</td></tr>'
    + '</table>' + imageHtml
    + '<div style="text-align:center;margin:26px 0;">' + actionButtons + '</div>'
    + '<div style="border-top:1px solid #e5e7eb;padding-top:18px;text-align:center;color:#64748b;font-size:12px;">Email được gửi tự động từ Hệ thống Web App Quản Lý Bản Vẽ.</div>'
    + '</div></div></div>';

  try {
    var mailOptions = { htmlBody: htmlBody };
    if (details.cutDrawingBlob) mailOptions.inlineImages = { rejectCutDrawing: details.cutDrawingBlob };
    GmailApp.sendEmail(targetEmail, subject, bodyText, mailOptions);
    return true;
  } catch (error) {
    Logger.log('Lỗi gửi email Reject: ' + error.toString());
    return false;
  }
}

function _getRejectNotificationEmail(previousStatus) {
  var st = _cleanStr(previousStatus);
  if (st.indexOf('checker 2') !== -1 || st.indexOf('checker2') !== -1) {
    return _getUserEmailByRole(['Checker 1', 'checker1']);
  }
  if (st.indexOf('approval') !== -1 || st.indexOf('trinh ky') !== -1 || st.indexOf('duyet') !== -1) {
    return _getUserEmailByRole(['Checker 2', 'checker2']);
  }
  return _getUserEmailByRole(['Charger', 'Charge', 'QA Charge', 'Charge QA']);
}

const _findRowByIdInColA = (sheet, drawingId) => {
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return -1;
  var dId = String(drawingId || '').trim().toLowerCase();
  if (!dId) return -1;

  var lastCol = Math.min(sheet.getLastColumn(), 34);
  var vals = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
  for (var i = 0; i < vals.length; i++) {
    var rId = String(vals[i][0] || '').trim().toLowerCase();
    var rDw = String(vals[i][11] || '').trim().toLowerCase();
    var rTo = String(vals[i][8] || '').trim().toLowerCase();
    if (rId === dId || (rDw && rDw === dId) || (rTo && rTo === dId)) {
      return i + 2;
    }
  }
  return -1;
};

const _writeNote = (ss, notiText, userEmail) => {
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
};

const approveDrawingOnServer = (drawingId, approvedLevel) => {
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
  var nowStr = Utilities.formatDate(new Date(), timezone, 'yyyy/MM/dd');

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
  var drawingDetails = _getApprovalDrawingDetails(sheet, sheetRow, drawingId);
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
    _sendApprovalWorkflowEmail(notifyEmail, dwNo, step.level, step.nextStatus, approverName, drawingDetails);
  }

  return {
    success: true,
    drawingId: drawingId,
    prevStatus: currentStatus,
    nextStatus: step.nextStatus,  
    level: step.level,
    approvedBy: approverName,
    approvedAt: nowStr
  };
};

const rejectDrawingOnServer = (drawingId, reason) => {
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
  var rejectorName = _getUserNameByEmail(rejectorEmail);
  var timezone = Session.getScriptTimeZone();
  var nowStr = Utilities.formatDate(new Date(), timezone, 'yyyy/MM/dd HH:mm');

  var prevStatus = String(sheet.getRange(sheetRow, COL_STATUS + 1).getValue() || '');
  sheet.getRange(sheetRow, COL_STATUS + 1).setValue('Tra lai - ' + prevStatus);

  var oldNote = String(sheet.getRange(sheetRow, COL_NOTE + 1).getValue() || '');
  sheet.getRange(sheetRow, COL_NOTE + 1).setValue(
    '[' + nowStr + ' - ' + rejectorEmail + '] TU CHOI: ' + reason.trim() +
    (oldNote ? '\n' + oldNote : '')
  );

  var dwNo = String(sheet.getRange(sheetRow, COL_DW_NO + 1).getValue() || drawingId);
  var drawingDetails = _getApprovalDrawingDetails(sheet, sheetRow, drawingId);
  var notiMsg = '[REJECT] Bản vẽ: ' + dwNo +
    ' | Trạng thái cũ: ' + prevStatus +
    ' | Người từ chối: ' + rejectorEmail +
    ' | Lúc: ' + nowStr +
    ' | Lý do: ' + reason.trim();
  _writeNote(ss, notiMsg, rejectorEmail);

  var rejectNotifyEmail = _getRejectNotificationEmail(prevStatus);
  if (rejectNotifyEmail) {
    _sendRejectWorkflowEmail(rejectNotifyEmail, dwNo, rejectorName, prevStatus, reason.trim(), drawingDetails);
  }

  return { success: true };
};



function requestReleaseDW(arr) {
  if (!Array.isArray(arr)) {
    throw new Error('Du lieu ban hanh khong hop le.');
  }

  const ss = SpreadsheetApp.openById("1t5PWyoJHrxElWP3QgmB16BEMIvEC015NHq0tsxu_TpE");
  const sheet = ss.getSheetByName("data") || ss.getSheetByName("Data");
  if (!sheet) {
    throw new Error('Khong tim thay sheet Data de luu du lieu ban hanh.');
  }

  const lastcol = sheet.getLastColumn();
  let data = [];

  const currentLastRow = sheet.getLastRow();
  if (currentLastRow > 1) {
    data = sheet.getRange(2, 1, currentLastRow - 1, lastcol).getValues();
  }

  const existingIds = new Set();
  data.forEach(row => {
    if(row[0]) existingIds.add(String(row[0]));
  });

  const requestTypeRaw = String((arr && arr[2]) || 'Normal');
  const requestType = /^urgent$/i.test(requestTypeRaw) ? 'Urgent' : 'Normal';
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

  const list = Array.isArray(arr[9])
    ? arr[9].filter(row => Array.isArray(row) && String(row[0] || '').trim())
    : [];
  if (list.length === 0) {
    throw new Error('Khong co ban ve nao duoc chon de ban hanh.');
  }

  let output = [];
  let ids = [];
  list.forEach(row => {
    let newID;
    do {
      newID = generateRandomString(6);
    } while (existingIds.has(newID));
    ids.push(String(row[0]).trim());

    const drawingReviseRaw = String(row[8] || '').trim();
    const drawingRevise = drawingReviseRaw || '0';
    const depName = String(arr[0] || '').trim();

    existingIds.add(newID);
     mailValue.push({
              requestId: orderNo,
              dw: row[1],
              status:"Yêu Cầu Ban Hàng Bản Vẽ Mới",
              dept: depName,
              name: arr[1],
            });

    output.push([
      newID,
      "BUNDLING & PACKING",
      row[4],
      "",
      row[1],
      drawingRevise,
      row[3] === 'New' ? 'New' : 'Change',
      "",
      orderNo,
      "",
      "QA-G2G",
      String(row[7] || ''), 
      arr[1],
      requestDateRaw,
      requestType,
      finalDueDate,
      "",
      "",
      "",
      requestType === 'Urgent' ? '' : releaseReason,
      "",
      "OK",
      "",
      "Hoàn thành",
    ]);
  });

  if (output.length === 0) {
    throw new Error('Khong tao duoc du lieu ban hanh tu danh sach da chon.');
  }

  // Khóa thao tác để hai yêu cầu ban hành đồng thời không dùng chung lastRow.
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    let writeRow = sheet.getLastRow() + 1;
    const batchSize = 100;
    for (let start = 0; start < output.length; start += batchSize) {
      const batch = output.slice(start, start + batchSize);
      sheet.getRange(writeRow, 1, batch.length, batch[0].length).setValues(batch);
      writeRow += batch.length;
    }
  } finally {
    lock.releaseLock();
  }

  sendMailFromDeptToCharge(mailValue, arr[0]);
  if (ids.length > 0) {
    updateBulkStatusByIds(ids, "Đang ban hành");
  }

  return "Đã tạo thành công " + output.length + " bản vẽ: " + orderNo + " | Hạn hoàn thành: " + finalDueDate;
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
  const idSet = new Set(targetIds.map(id => String(id || '').trim()));
  let hasChanges = false;

  // 2. XỬ LÝ TRONG RAM: Vòng lặp này chạy bằng tốc độ của CPU/RAM, gần như tức thời
  for (let i = 0; i < idValues.length; i++) {
    const currentId = String(idValues[i][0] || '').trim();
    
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
