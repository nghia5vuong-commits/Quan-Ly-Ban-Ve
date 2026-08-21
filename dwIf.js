var MAIL_CONFIG = {
  SUBJECT_FILTERS: ['MANUFACTURING ORDER', 'NEW PROJECT', 'THÔNG BÁO TỰ ĐỘNG'],
  MAX_THREADS: 100,
  TIMEZONE: 'Asia/Ho_Chi_Minh'
};

function toSafeString(value) {
  if (value === null || value === undefined) return '';
  if (typeof value === 'string') return value.trim();
  if (typeof value === 'number' || typeof value === 'boolean') return String(value).trim();
  if (value instanceof Date) return Utilities.formatDate(value, MAIL_CONFIG.TIMEZONE, 'dd/MM/yyyy');

  try {
    if (typeof value.toString === 'function') {
      var text = String(value.toString());
      if (text && text !== '[object Object]' && text !== '[object CellImage]') return text.trim();
    }
  } catch (err) {
  }

  return '';
}

function normalizeImageSource(value) {
  if (value === null || value === undefined || value === '') return '';

  if (typeof value === 'string') return value.trim();
  if (typeof value === 'number' || typeof value === 'boolean') return String(value).trim();

  if (typeof value === 'object') {
    try {
      if (typeof value.getContentUrl === 'function') {
        var contentUrl = value.getContentUrl();
        if (contentUrl) return String(contentUrl).trim();
      }
      if (typeof value.getUrl === 'function') {
        var imgUrl = value.getUrl();
        if (imgUrl) return String(imgUrl).trim();
      }
      if (typeof value.toString === 'function') {
        var strVal = String(value.toString());
        if (strVal && strVal !== '[object Object]' && strVal !== '[object CellImage]') {
          if (/^(https?:\/\/|data:|\/|drive\.google|googleusercontent)/i.test(strVal)) {
            return strVal.trim();
          }
        }
      }
    } catch (err) {
    }
  }

  return '';
}

function buildQuery(keyword) {
  var query = 'subject:("MANUFACTURING ORDER" OR "NEW PROJECT" OR "THÔNG BÁO TỰ ĐỘNG")';
  if (keyword && keyword.trim() !== '') {
    query = '(' + query + ') ' + keyword.trim();
  }
  return query;
}

function getManufacturingEmails(page, keyword) {
  try {
    page = page || 0;
    keyword = keyword || '';
    var query = buildQuery(keyword);
    var limit = MAIL_CONFIG.MAX_THREADS;
    var start = page * limit;

    // Chỉ load MAX_THREADS + 1 để biết có trang tiếp theo không
    var threads = GmailApp.search(query, start, limit + 1);
    var hasMore = threads.length > limit;
    if (hasMore) threads.pop(); // Loại bỏ thread dư thừa lấy làm mốc

    if (threads.length === 0) return { success: true, emails: [], total: 0, page: page, hasMore: false };

    var sheetId = '1DRteBSFT1cj4R_OUPMoDxeLMzAIJexWF3HPT-rpMOoM';
    var ss = SpreadsheetApp.openById(sheetId);
    var logSheet = ss.getSheetByName('Log');
    var receivedThreadIds = new Set();

    if (logSheet) {
      var data = logSheet.getDataRange().getValues();
      for (var r = 0; r < data.length; r++) {
        for (var c = 0; c < data[r].length; c++) {
          var cellValue = String(data[r][c]).trim();
          // Quét và nạp các giá trị giống định dạng ID (chuỗi dài) vào Set
          if (cellValue.length > 10) {
            receivedThreadIds.add(cellValue);
          }
        }
      }
    }

    // TỐI ƯU CỐT LÕI: Lấy messages của TOÀN BỘ threads bằng 1 lệnh duy nhất
    var threadsMessages = GmailApp.getMessagesForThreads(threads);

    function detectTag(subject) {
      if (!subject) return '';
      var sUpper = subject.toUpperCase();
      for (var j = 0; j < MAIL_CONFIG.SUBJECT_FILTERS.length; j++) {
        if (sUpper.indexOf(MAIL_CONFIG.SUBJECT_FILTERS[j]) !== -1) return MAIL_CONFIG.SUBJECT_FILTERS[j];
      }
      return '';
    }

    var emails = [];
    for (var i = 0; i < threads.length; i++) {
      try {
        var thread = threads[i];
        var threadId = thread.getId();
        var messages = threadsMessages[i];
        if (!messages || messages.length === 0) continue;

        var lastMsg = messages[messages.length - 1];
        var from = lastMsg.getFrom() || '';

        emails.push({
          id: threadId,
          subject: lastMsg.getSubject() || '(Không có tiêu đề)',
          sender: from,
          senderName: extractName(from),
          senderEmail: extractEmail(from),
          date: formatDate(lastMsg.getDate()),
          dateRaw: lastMsg.getDate().getTime(),
          isUnread: thread.isUnread(),
          messageCount: messages.length,
          attachments: [],
          tag: detectTag(lastMsg.getSubject()),

          // BỔ SUNG TRẠNG THÁI TIẾP NHẬN: Đối chiếu với Set vừa tạo ở trên
          isReceived: receivedThreadIds.has(threadId)
        });
      } catch (err) { }
    }

    return {
      success: true,
      emails: emails,
      total: hasMore ? (start + limit + "+") : (start + emails.length),
      page: page,
      hasMore: hasMore
    };
  } catch (err) {
    return { success: false, error: err.toString(), emails: [], total: 0 };
  }
}

// 🚀 TỐI ƯU 2: HÀM LẤY CHI TIẾT EMAIL (Giữ nguyên phần đính kèm ở đây vì người dùng thực sự đang xem nó)
function getEmailDetail(threadId) {
  try {
    var thread = GmailApp.getThreadById(threadId);
    var messages = thread.getMessages();

    var messageList = messages.map(function (msg, idx) {
      return {
        index: idx,
        messageId: msg.getId(),
        subject: msg.getSubject(),
        from: msg.getFrom(),
        senderName: extractName(msg.getFrom()),
        senderEmail: extractEmail(msg.getFrom()),
        to: msg.getTo(),
        cc: msg.getCc(),
        date: formatDate(msg.getDate()),
        dateRaw: msg.getDate().getTime(),
        body: sanitizeHtml(msg.getBody()),
        plainBody: msg.getPlainBody(),
        isUnread: msg.isUnread(),
        attachments: getAttachmentInfo(msg) // Chỉ parse file khi click mở email
      };
    });

    thread.markRead();
    var lastMsg = messages[messages.length - 1];

    return {
      success: true,
      threadId: threadId,
      subject: lastMsg.getSubject(),
      messageCount: messages.length,
      messages: messageList
    };
  } catch (err) {
    return { success: false, error: err.toString() };
  }
}

function getAttachmentParsedData(threadId, msgIndex, attIndex) {
  try {
    var thread = GmailApp.getThreadById(threadId);
    if (!thread) {
      throw new Error("Không tìm thấy chuỗi email.");
    }

    var messages = thread.getMessages();
    if (!messages || msgIndex >= messages.length) {
      throw new Error("Không tìm thấy tin nhắn.");
    }

    var attachments = messages[msgIndex].getAttachments();
    if (!attachments || attIndex >= attachments.length) {
      throw new Error("Không tìm thấy file đính kèm.");
    }

    var attachment = attachments[attIndex];
    var base64Data = Utilities.base64Encode(attachment.getBytes());

    return {
      success: true,
      dataBase64: base64Data
    };
  } catch (err) {
    return {
      success: false,
      error: err.toString()
    };
  }
}

// 🚀 TỐI ƯU 3: HÀM THỐNG KÊ (Sử dụng Cache lâu hơn, giới hạn quét)
function getStatistics() {
  try {
    var cache = CacheService.getScriptCache();
    var cached = cache.get("MAIL_STATS");
    if (cached) return JSON.parse(cached);

    var query = buildQuery('');
    var allThreads = GmailApp.search(query, 0, 50); // Cắt xuống 50 để thống kê nhẹ nhàng hơn

    var totalUnread = 0, todayCount = 0, weekCount = 0;
    var today = new Date(); today.setHours(0, 0, 0, 0);
    var weekAgo = new Date(today.getTime() - 7 * 24 * 60 * 60 * 1000);

    allThreads.forEach(function (t) {
      if (t.isUnread()) totalUnread++;
      var d = t.getLastMessageDate();
      if (d >= today) todayCount++;
      if (d >= weekAgo) weekCount++;
    });

    var result = {
      success: true,
      total: allThreads.length === 50 ? '50+' : allThreads.length,
      unread: totalUnread,
      read: allThreads.length - totalUnread,
      today: todayCount,
      thisWeek: weekCount
    };

    // Tăng thời gian lưu cache lên 5 phút (300 giây) để chống request dồn dập
    cache.put("MAIL_STATS", JSON.stringify(result), 300);
    return result;
  } catch (err) {
    return { success: false, error: err.toString(), total: 0, unread: 0, read: 0 };
  }
}

function getEmailReceptionStats() {
  try {
    var query = buildQuery('');
    var threads = GmailApp.search(query, 0, 100);

    var sheetId = '1DRteBSFT1cj4R_OUPMoDxeLMzAIJexWF3HPT-rpMOoM';
    var ss = SpreadsheetApp.openById(sheetId);
    var logSheet = ss.getSheetByName('Log');
    var receivedThreadIds = new Set();

    if (logSheet) {
      var data = logSheet.getDataRange().getValues();
      for (var r = 0; r < data.length; r++) {
        for (var c = 0; c < data[r].length; c++) {
          var cellValue = String(data[r][c]).trim();
          if (cellValue.length > 10) {
            receivedThreadIds.add(cellValue);
          }
        }
      }
    }

    var receivedCount = 0;
    var unreceivedCount = 0;

    for (var i = 0; i < threads.length; i++) {
      var threadId = threads[i].getId();
      if (receivedThreadIds.has(threadId)) {
        receivedCount++;
      } else {
        unreceivedCount++;
      }
    }

    return {
      success: true,
      received: receivedCount,
      unreceived: unreceivedCount,
      total: threads.length
    };
  } catch (err) {
    return { success: false, error: err.toString(), received: 0, unreceived: 0, total: 0 };
  }
}


function toggleDrawingForm(rowIdx, subject, customer, to, so) {
  var container = document.getElementById('form-container-' + rowIdx);
  var form = document.getElementById('drawingForm');
  var btn = document.getElementById('btn-toggle-' + rowIdx);

  // Chốt chặn an toàn: Tránh lỗi "Cannot read properties of null"
  if (!form) {
    showToast('Hệ thống đang tải form, vui lòng thử lại!', 'error');
    return;
  }

  if (pendingState.selectedLogIdx === rowIdx && !container.classList.contains('d-none')) {
    closeDrawingForm(); return;
  }

  if (pendingState.selectedLogIdx) {
    var oldBtn = document.getElementById('btn-toggle-' + pendingState.selectedLogIdx);
    if (oldBtn) {
      oldBtn.innerHTML = '<i class="fa fa-pen me-2"></i>Nhập dữ liệu';
      oldBtn.classList.replace('btn-primary', 'btn-outline-primary');
      oldBtn.classList.remove('text-secondary');
    }
    var oldContainer = document.getElementById('form-container-' + pendingState.selectedLogIdx);
    if (oldContainer) oldContainer.classList.add('d-none');
  }

  pendingState.selectedLogIdx = rowIdx;
  pendingState.subject = subject;

  form.reset();
  var rDate = document.getElementById('formReceivedDate');
  if (rDate) rDate.value = new Date().toISOString().split('T')[0];
  if (document.getElementById('formCustomer')) document.getElementById('formCustomer').value = customer || '';
  if (document.getElementById('formTO')) document.getElementById('formTO').value = to || '';
  if (document.getElementById('formSO')) document.getElementById('formSO').value = so || '';

  generateDwCode();

  container.appendChild(form);
  container.classList.remove('d-none');

  btn.innerHTML = '<i class="fa fa-chevron-up me-2"></i>Đóng form';
  btn.classList.replace('btn-outline-primary', 'btn-primary');
  btn.classList.add('text-secondary');
}

function markEmail(threadId, markAsRead) {
  try {
    var thread = GmailApp.getThreadById(threadId);
    if (markAsRead) thread.markRead(); else thread.markUnread();
    return { success: true };
  } catch (err) {
    return { success: false, error: err.toString() };
  }
}

function extractName(s) { var m = s.match(/^([^<]+)</); return m ? m[1].trim() : s; }
function extractEmail(s) { var m = s.match(/<([^>]+)>/); return m ? m[1] : s; }
function formatDate(date) { return Utilities.formatDate(date, MAIL_CONFIG.TIMEZONE, 'dd/MM/yyyy HH:mm'); }

function getAttachmentInfo(message) {
  try {
    return message.getAttachments().map(function (att, idx) {
      return {
        index: idx, name: att.getName(), contentType: att.getContentType(),
        size: formatFileSize(att.getSize()), isSheet: isSpreadsheetType(att.getName(), att.getContentType())
      };
    });
  } catch (e) { return []; }
}

function isSpreadsheetType(name, contentType) {
  var n = (name || '').toLowerCase(), t = (contentType || '').toLowerCase();
  return !!(n.match(/\.(xlsx|xls|csv|ods|xlsm|xlsb)$/) || t.indexOf('spreadsheet') !== -1 || t.indexOf('excel') !== -1 || t.indexOf('csv') !== -1);
}

function formatFileSize(bytes) {
  if (bytes < 1024) return bytes + ' B';
  if (bytes < 1048576) return (bytes / 1024).toFixed(1) + ' KB';
  return (bytes / 1048576).toFixed(1) + ' MB';
}

function sanitizeHtml(html) { return html.replace(/<script[^>]*>[\s\S]*?<\/script>/gi, '').replace(/on\w+\s*=\s*["'][^"']*["']/gi, ''); }

function getAttachmentForDownload(threadId, messageIndex, attachmentIndex) {
  try {
    var thread = GmailApp.getThreadById(threadId);
    var messages = thread.getMessages();
    if (messageIndex >= messages.length) return { success: false, error: 'Không tìm thấy tin nhắn.' };
    var att = messages[messageIndex].getAttachments()[attachmentIndex];
    return { success: true, dataBase64: Utilities.base64Encode(att.getBytes()), fileName: att.getName() };
  } catch (err) {
    return { success: false, error: err.toString() };
  }
}

function saveToLogSheet(threadId, subject, parsedData, fullDataString) {
  try {
    var sheetId = '1DRteBSFT1cj4R_OUPMoDxeLMzAIJexWF3HPT-rpMOoM';
    var ss = SpreadsheetApp.openById(sheetId);
    var logSheet = ss.getSheetByName('Log');

    if (!logSheet) {
      logSheet = ss.insertSheet('Log');
    }

    var timestamp = Utilities.formatDate(new Date(), MAIL_CONFIG.TIMEZONE, 'dd/MM/yyyy HH:mm:ss');
    var rowData = [
      timestamp,                  // A
      threadId,                   // B
      subject,                    // C
      parsedData.customer || '',  // D
      parsedData.to || '',        // E
      parsedData.so || '',        // F
      'Đã tiếp nhận',             // G
      fullDataString || ''        // H
    ];

    logSheet.appendRow(rowData);
    CacheService.getScriptCache().remove("MAIL_STATS");

    return { success: true, message: 'Đã lưu tiếp nhận thiết kế thành công!' };
  } catch (err) {
    return { success: false, error: err.toString() };
  }
}

// 🚀 TỐI ƯU 4: KIỂM TRA TRẠNG THÁI TIẾP NHẬN (Chỉ quét cột ThreadID thay vì toàn bộ bảng)
function checkEmailReceived(threadId) {
  try {
    var sheetId = '1DRteBSFT1cj4R_OUPMoDxeLMzAIJexWF3HPT-rpMOoM';
    var ss = SpreadsheetApp.openById(sheetId);
    var logSheet = ss.getSheetByName('Log');
    if (!logSheet) return { success: true, isReceived: false };

    var lastRow = logSheet.getLastRow();
    if (lastRow < 2) return { success: true, isReceived: false };

    // TỐI ƯU CỐT LÕI: Chỉ tải dữ liệu của CỘT B (ThreadID) về máy chủ. Nhanh hơn tải hàng ngàn cột chéo x10 lần.
    var threadIds = logSheet.getRange(2, 2, lastRow - 1, 1).getValues();

    for (var i = 0; i < threadIds.length; i++) {
      if (String(threadIds[i][0]).trim() === String(threadId).trim()) {
        // Chỉ lấy ngày khi khớp (tránh dùng getDisplayValues cho toàn bộ bảng làm chậm máy)
        var receivedDate = logSheet.getRange(i + 2, 1).getDisplayValue();
        return { success: true, isReceived: true, receivedDate: receivedDate };
      }
    }
    return { success: true, isReceived: false };
  } catch (err) {
    return { success: false, error: err.toString(), isReceived: false };
  }
}

function getPendingLogData() {
  try {
    var sheetId = '1DRteBSFT1cj4R_OUPMoDxeLMzAIJexWF3HPT-rpMOoM';
    var ss = SpreadsheetApp.openById(sheetId);
    var logSheet = ss.getSheetByName('Log');
    if (!logSheet) return { success: true, data: [] };

    // Đọc danh sách các Mã SO đã có trong sheet "Data"
    var dataSheet = ss.getSheetByName('Data');
    var existingSoMap = {};
    if (dataSheet && dataSheet.getLastRow() > 1) {
      var dataVals = dataSheet.getRange(2, 1, dataSheet.getLastRow() - 1, dataSheet.getLastColumn()).getValues();
      for (var d = 0; d < dataVals.length; d++) {
        var r = dataVals[d];
        var soCode = "";
        if (r && r[33]) soCode = toSafeString(r[33]);
        else if (r && r[5]) soCode = toSafeString(r[5]);

        if (!soCode && r) {
          for (var c = 0; c < r.length; c++) {
            var cellStr = toSafeString(r[c]);
            if (/^SO\d+/i.test(cellStr)) {
              soCode = cellStr;
              break;
            }
          }
        }

        if (soCode) {
          var cleanSo = soCode.toUpperCase().replace(/\s+/g, "");
          var digitsSo = cleanSo.replace(/[^0-9]/g, "");

          var imgVal = r[32] || "";
          var imgUrl = normalizeImageSource(imgVal);

          existingSoMap[cleanSo] = imgUrl || true;
          if (digitsSo) existingSoMap[digitsSo] = imgUrl || true;
        }
      }
    }

    var data = logSheet.getDataRange().getDisplayValues();
    var logs = [];

    for (var i = 1; i < data.length; i++) {
      if (String(data[i][6]).trim() === 'Đã tiếp nhận') {
        var rawSo = (data[i][5] || "").toString().trim();
        var cleanRaw = rawSo.toUpperCase().replace(/\s+/g, "");
        var digitsRaw = cleanRaw.replace(/[^0-9]/g, "");

        var foundInfo = rawSo ? (existingSoMap[cleanRaw] || (digitsRaw ? existingSoMap[digitsRaw] : null)) : null;
        var existsInData = !!foundInfo;
        var imageUrl = (typeof foundInfo === 'string') ? foundInfo : "";

        logs.push({
          rowIdx: i + 1,
          date: String(data[i][0] || ''),
          threadId: String(data[i][1] || ''),
          subject: String(data[i][2] || ''),
          customer: String(data[i][3] || ''),
          to: String(data[i][4] || ''),
          so: rawSo,
          existsInData: existsInData,
          imageUrl: imageUrl
        });
      }
    }
    return { success: true, data: logs.reverse() };
  } catch (err) {
    return { success: false, error: err.toString() };
  }
}

function getLogDataFromSheet(threadId) {
  try {
    var sheetId = '1DRteBSFT1cj4R_OUPMoDxeLMzAIJexWF3HPT-rpMOoM';
    var ss = SpreadsheetApp.openById(sheetId);
    var logSheet = ss.getSheetByName('Log');
    if (!logSheet) return { success: false, error: 'Chưa có tab Log' };

    var data = logSheet.getDataRange().getDisplayValues();

    // Quét từ dưới lên trên tìm dòng log mới nhất
    for (var i = data.length - 1; i >= 0; i--) {
      if (String(data[i][1]).trim() === String(threadId).trim()) {
        var jsonStr = data[i][7]; // Cột H
        var customer = '', to = '', so = '';

        if (jsonStr && jsonStr.startsWith('[')) {
          try {
            var rows = JSON.parse(jsonStr);
            var headerRowIdx = -1, colSO = -1, colCustomer = -1, colMaterial = -1;

            for (var r = 0; r < rows.length; r++) {
              if (!rows[r]) continue;
              for (var c = 0; c < rows[r].length; c++) {
                var val = String(rows[r][c] || '').trim().toLowerCase();
                if (val === 'sales document') { headerRowIdx = r; colSO = c; }
                if (val === 'item descr.') { headerRowIdx = r; colCustomer = c; }
                if (val === 'material') { colMaterial = c; }
              }
              if (headerRowIdx !== -1) break;
            }

            if (headerRowIdx !== -1 && rows.length > headerRowIdx + 1) {
              var dataRow = rows[headerRowIdx + 1];
              if (colSO !== -1 && dataRow[colSO]) so = 'SO' + String(dataRow[colSO]).replace(/[^0-9]/g, '');
              if (colCustomer !== -1 && dataRow[colCustomer]) customer = String(dataRow[colCustomer]).trim().toUpperCase();
              if (colMaterial !== -1 && dataRow[colMaterial]) {
                var rawMat = String(dataRow[colMaterial]).trim();
                to = rawMat.length > 1 ? rawMat.substring(1) : rawMat;
              }
            }
          } catch (e) { }
        }

        return {
          success: true,
          customer: customer || data[i][3],
          to: to || data[i][4],
          so: so || data[i][5]
        };
      }
    }
    return { success: false, error: 'Chưa có thông tin tiếp nhận trong Log' };
  } catch (err) {
    return { success: false, error: err.toString() };
  }
}

function generateUniqueId() {
  var now = new Date();
  var dateStr = Utilities.formatDate(now, Session.getScriptTimeZone(), "yyyyMMdd");
  var randomNum = Math.floor(1000 + Math.random() * 9000); // 4 chữ số ngẫu nhiên
  return "REQ-" + dateStr + "-" + randomNum;
}

function getCurrentUserEmail() {
  try {
    var email = Session.getActiveUser().getEmail();
    return email || "System";
  } catch (e) {
    return "System";
  }
}

function saveDataToTestSheet(subject, matrixData, rowIdx) {
  try {
    var ss = SpreadsheetApp.openById("1DRteBSFT1cj4R_OUPMoDxeLMzAIJexWF3HPT-rpMOoM");
    var sectionImageFolderId = "19Faip8STiBLJ58aVLiqwVs4SaotKb1bm";
    var sheetData = ss.getSheetByName("Data") || ss.getSheetByName("data") || ss.insertSheet("Data");
    var sheetSO = ss.getSheetByName("SO");
    var sheetCustomer = ss.getSheetByName("Customer");

    if (!matrixData || matrixData.length === 0) {
      return { success: false, error: "Dữ liệu gửi lên rỗng!" };
    }

    // Kiểm tra xem có phải là chế độ cập nhật (upgrade) hay không
    var isUpgradeMode = rowIdx && typeof rowIdx === 'number' && rowIdx >= 2;

    var formRow = matrixData[0];
    var imageObj = formRow.pop();
    var cellImage = null;
    var imageSourceUrl = '';

    if (imageObj && imageObj.base64) {
      try {
        var imageBytes = Utilities.base64Decode(imageObj.base64);
        var imageMimeType = imageObj.mimeType || 'image/png';
        var imageBlob = Utilities.newBlob(imageBytes, imageMimeType, imageObj.name || 'section-image');
        var sectionImageFolder = DriveApp.getFolderById(sectionImageFolderId);
        var imageFile = sectionImageFolder.createFile(imageBlob);

        // Cho phép người dùng trong cùng miền mở ảnh từ liên kết Drive.
        try {
          imageFile.setSharing(DriveApp.Access.DOMAIN_WITH_LINK, DriveApp.Permission.VIEW);
        } catch (sharingErr) {
          Logger.log('Không thể cập nhật quyền chia sẻ ảnh mặt cắt: ' + sharingErr.toString());
        }

        imageSourceUrl = "data:" + imageObj.mimeType + ";base64," + imageObj.base64;
        cellImage = SpreadsheetApp.newCellImage()
          .setSourceUrl(imageSourceUrl)
          .build();
      } catch (imgErr) {
        throw new Error('Không thể lưu hình ảnh mặt cắt vào Drive: ' + imgErr.toString());
      }
    }

    // 1. Trích xuất 17 trường dữ liệu từ HTML Form
    var group = formRow[0] || "";
    var type = formRow[1] || "";
    var customerName = formRow[2] || "";
    var toCode = formRow[3] || "";
    var project = formRow[4] || "";
    var soNo = formRow[5] || "";
    var dwCode = formRow[6] || "";
    var version = formRow[7] || "";
    var typeDw = formRow[8] || "";
    var assignee = formRow[9] || "";
    var widthVal = formRow[10] || "";
    var heightVal = formRow[11] || "";
    var receivedDate = formRow[12] || "";
    var assigneeDoneDate = formRow[13] || "";
    var actualDoneDate = formRow[14] || "";
    var fyeVal = formRow[15] || "";
    var noteVal = formRow[16] || "";

    // ========================================================================
    // KIỂM TRA: TO + Customer đã tồn tại trong sheet "Data" chưa?
    // Chỉ bỏ qua ghi Data khi cùng TO + Customer đã tồn tại, không phân biệt "TO-123" / "123" hay chữ hoa/thường.
    // ========================================================================
    function normalizeToValue(value) {
      if (value === null || value === undefined) return "";
      return String(value).replace(/^TO-?/i, "").replace(/\s+/g, "").trim().toUpperCase();
    }

    function normalizeCustomerValue(value) {
      if (value === null || value === undefined) return "";
      return String(value).replace(/\s+/g, " ").trim().toUpperCase();
    }

    var toCodeNormalized = normalizeToValue(toCode);
    var customerNameNormalized = normalizeCustomerValue(customerName);
    var toCustomerExists = false;

    if (toCodeNormalized && customerNameNormalized && sheetData && sheetData.getLastRow() > 1) {
      try {
        var existingData = sheetData.getRange(2, 1, sheetData.getLastRow() - 1, sheetData.getLastColumn()).getValues();
        for (var checkIdx = 0; checkIdx < existingData.length; checkIdx++) {
          var checkRow = existingData[checkIdx];
          var existingTO = normalizeToValue(checkRow[8]);      // Column I (index 8): TO
          var existingCustomer = normalizeCustomerValue(checkRow[10]); // Column K (index 10): Customer

          if (existingTO === toCodeNormalized && existingCustomer === customerNameNormalized) {
            toCustomerExists = true;
            break;
          }
        }
      } catch (checkErr) {
      }
    }

    // Luôn luôn tạo một dòng MỚI ở cuối sheet "Data" khi lưu dữ liệu (NẾU chưa tồn tại)
    // TRƯỜNG HỢP UPGRADE: cập nhật dòng hiện có thay vì tạo dòng mới
    var targetRow = isUpgradeMode ? rowIdx : (sheetData.getLastRow() + 1);
    var autoId = isUpgradeMode ? sheetData.getRange(rowIdx, 1).getValue() : generateUniqueId();
    var userEmail = getCurrentUserEmail();

    // 2. Ghi dữ liệu vào sheet "Data" (33 cột từ A -> AG)
    var mappedRowData = new Array(33).fill("");
    mappedRowData[0] = autoId;             // A: ID
    mappedRowData[1] = userEmail;          // B: Mail ID
    mappedRowData[2] = "Đang thực hiện";           // C: Status (không thay đổi nếu upgrade)
    mappedRowData[3] = group;              // D: Group
    mappedRowData[4] = type;               // E: Type
    mappedRowData[5] = version;                   // F: Drawing Revise
    mappedRowData[6] = receivedDate;       // G: Ngày tiếp nhận
    mappedRowData[7] = assigneeDoneDate;   // H: Ngày hoàn thành dự kiến
    mappedRowData[8] = toCode;             // I: TO
    mappedRowData[9] = project;            // J: Dự án
    mappedRowData[10] = customerName;       // K: Customer
    mappedRowData[11] = dwCode;             // L: Drawing code
    mappedRowData[12] = typeDw;            // M: Version Drawing
    mappedRowData[13] = "";                 // N: Nội dung thay đổi
    mappedRowData[14] = "";                 // O: File Drawing (PDF)
    mappedRowData[15] = assignee;           // P: Người đảm trách
    mappedRowData[16] = "";   // Q: Ngày người đảm trách hoàn thành
    mappedRowData[17] = "";                 // R: Checker
    mappedRowData[18] = "";                 // S: Ngày hoàn thành thực tế
    mappedRowData[19] = "";                 // T: Approver
    mappedRowData[20] = "";                 // U: Ngày phát hành thực tế
    mappedRowData[21] = "RUNNING";          // V: Trạng thái bản vẽ
    mappedRowData[22] = noteVal;            // W: Note
    mappedRowData[23] = fyeVal;             // X: FYE
    mappedRowData[24] = typeDw;             // Y: Type Drawing
    mappedRowData[25] = "";                 // Z: Type Product
    mappedRowData[26] = "";                 // AA: Product name
    mappedRowData[27] = "";                 // AB: Product code + Name
    mappedRowData[28] = widthVal;           // AC: W
    mappedRowData[29] = heightVal;          // AD: H
    mappedRowData[30] = "";                 // AE: Sample File Excel
    mappedRowData[31] = "";                 // AF: Sample File PDF
    mappedRowData[32] = "";                  // AG: Hình ảnh mặt cắt (để trống, ảnh được chèn bởi CellImage riêng)
    mappedRowData[33] = soNo;              // AH: Mã SO

    // Chỉ kiểm tra TO+Customer trong chế độ INSERT (không upgrade)
    if (isUpgradeMode) {
      // Chế độ UPDATE: luôn ghi lại dữ liệu vào hàng hiện có
      sheetData.getRange(targetRow, 1, 1, mappedRowData.length).setValues([mappedRowData]);
      if (cellImage) {
        sheetData.getRange(targetRow, 33).setValue(cellImage);
      }
    } else {
      // Chế độ INSERT: kiểm tra TO+Customer trước khi ghi
      // ========================================================================
      // KIỂM TRA: TO + Customer đã tồn tại trong sheet "Data" chưa?
      // Chỉ bỏ qua ghi Data khi cùng TO + Customer đã tồn tại, không phân biệt "TO-123" / "123" hay chữ hoa/thường.
      // ========================================================================
      function normalizeToValue(value) {
        if (value === null || value === undefined) return "";
        return String(value).replace(/^TO-?/i, "").replace(/\s+/g, "").trim().toUpperCase();
      }

      function normalizeCustomerValue(value) {
        if (value === null || value === undefined) return "";
        return String(value).replace(/\s+/g, " ").trim().toUpperCase();
      }

      var toCodeNormalized = normalizeToValue(toCode);
      var customerNameNormalized = normalizeCustomerValue(customerName);
      var toCustomerExists = false;

      if (toCodeNormalized && customerNameNormalized && sheetData && sheetData.getLastRow() > 1) {
        try {
          var existingData = sheetData.getRange(2, 1, sheetData.getLastRow() - 1, sheetData.getLastColumn()).getValues();
          for (var checkIdx = 0; checkIdx < existingData.length; checkIdx++) {
            var checkRow = existingData[checkIdx];
            var existingTO = normalizeToValue(checkRow[8]);      // Column I (index 8): TO
            var existingCustomer = normalizeCustomerValue(checkRow[10]); // Column K (index 10): Customer

            if (existingTO === toCodeNormalized && existingCustomer === customerNameNormalized) {
              toCustomerExists = true;
              break;
            }
          }
        } catch (checkErr) {
        }
      }

      // Chỉ lưu vào sheet "Data" nếu TO + Customer chưa tồn tại
      if (!toCustomerExists) {
        sheetData.getRange(targetRow, 1, 1, mappedRowData.length).setValues([mappedRowData]);
        if (cellImage) {
          sheetData.getRange(targetRow, 33).setValue(cellImage);
        }
      } else {
      }
    }

    // Luôn luôn lưu vào sheet "SO" (dù TO+Customer đã tồn tại hay chưa)
    if (sheetSO) {
      var colorCode = "";
      var pureToCode = toCode;
      var cleanTo = toCode.replace(/^TO-?/i, "").trim();

      if (cleanTo.length >= 2) {
        colorCode = cleanTo.substring(0, 2).toUpperCase();
        pureToCode = cleanTo.substring(2).trim().split(" ")[0];
      }

      var toMau = "TO-" + colorCode + pureToCode;

      var customerCode = customerName;
      if (sheetCustomer) {
        var custData = sheetCustomer.getDataRange().getValues();
        for (var c = 1; c < custData.length; c++) {
          if (custData[c][2] && custData[c][2].toString().trim().toLowerCase() === customerName.trim().toLowerCase()) {
            customerCode = custData[c][1] || customerName;
            break;
          }
        }
      }

      var toCustomer = "TO-" + pureToCode + customerCode;
      var toMauNormalized = toMau.trim().toUpperCase();
      var toCustomerNormalized = toCustomer.trim().toUpperCase();

      // ========================================================================
      // KIỂM TRA: (TO MÀU + TO theo khách hàng) có tồn tại trong SO chưa?
      // ========================================================================
      var soRowExists = -1;
      var soData = sheetSO.getDataRange().getValues();
      
      for (var soIdx = 1; soIdx < soData.length; soIdx++) { // Bắt đầu từ row 2 (index 1)
        var soRow = soData[soIdx];
        var existingToMau = (soRow[5] || "").toString().trim().toUpperCase();      // Column F (index 5): TO MÀU
        var existingToCustomer = (soRow[8] || "").toString().trim().toUpperCase();  // Column I (index 8): TO theo khách hàng
        
        if (existingToMau === toMauNormalized || existingToCustomer === toCustomerNormalized) {
          soRowExists = soIdx + 1; // +1 vì hàng sheet bắt đầu từ 1, array index bắt đầu từ 0
          break;
        }
      }

      if (soRowExists !== -1) {
        // ========================================================================
        // CẬP NHẬT: Tần xuất TO, Tần xuất MÀU, Tần xuất theo khách
        // ========================================================================
        try {
          var updateRow = soData[soRowExists - 1]; // Convert row number to array index
          
          // Column G (index 6): Tần xuất TO - luôn +1
          var currentTanXuatTO = parseInt(updateRow[6] || 0) || 0;
          var newTanXuatTO = currentTanXuatTO + 1;
          sheetSO.getRange(soRowExists, 7).setValue(newTanXuatTO); // Column G
          
          // Column H (index 7): Tần xuất MÀU theo TO - nếu màu giống +1
          var existingColorCode = (updateRow[2] || "").toString().trim().toUpperCase();
          if (existingColorCode === colorCode) {
            var currentTanXuatMau = parseInt(updateRow[7] || 0) || 0;
            var newTanXuatMau = currentTanXuatMau + 1;
            sheetSO.getRange(soRowExists, 8).setValue(newTanXuatMau); // Column H
          }
          
          // Column J (index 9): Tần xuất theo khách hàng - nếu khách giống +1
          var existingToCustomer = (updateRow[8] || "").toString().trim().toUpperCase();
          if (existingToCustomer === toCustomerNormalized) {
            var currentTanXuatKhach = parseInt(updateRow[9] || 0) || 0;
            var newTanXuatKhach = currentTanXuatKhach + 1;
            sheetSO.getRange(soRowExists, 10).setValue(newTanXuatKhach); // Column J
          }
        } catch (updateErr) {
        }
      } else {
        // ========================================================================
        // TẠO HÀNG MỚI: (TO MÀU) chưa tồn tại trong SO
        // Nhưng cần đếm: Tần xuất TO và Tần xuất theo khách
        // ========================================================================
        
        // Tìm max Tần xuất TO (Column G) cho TO này (7Y090A)
        var maxTanXuatTO = 0;
        var toToFind = ("TO-" + pureToCode).toUpperCase();
        for (var countIdx = 1; countIdx < soData.length; countIdx++) {
          var countRow = soData[countIdx];
          var existingTO = (countRow[0] || "").toString().trim().toUpperCase(); // Column A: TO
          if (existingTO === toToFind) {
            var rowTanXuatTO = parseInt(countRow[6] || 0) || 0;
            if (rowTanXuatTO > maxTanXuatTO) {
              maxTanXuatTO = rowTanXuatTO;
            }
          }
        }
        var newTanXuatTO = maxTanXuatTO + 1;
        
        // Tìm max Tần xuất theo khách (Column J) cho TO+khách này
        var maxTanXuatKhach = 0;
        var toCustomerToFind = toCustomerNormalized;
        for (var countIdx2 = 1; countIdx2 < soData.length; countIdx2++) {
          var countRow2 = soData[countIdx2];
          var existingToCustomer = (countRow2[8] || "").toString().trim().toUpperCase(); // Column I: TO theo khách
          if (existingToCustomer === toCustomerToFind) {
            var rowTanXuatKhach = parseInt(countRow2[9] || 0) || 0;
            if (rowTanXuatKhach > maxTanXuatKhach) {
              maxTanXuatKhach = rowTanXuatKhach;
            }
          }
        }
        var newTanXuatKhach = maxTanXuatKhach + 1;
        
        // Tạo hàng mới với tần xuất được tính toán
        var mappedRowSO = new Array(12).fill("");
        mappedRowSO[0] = "TO-" + pureToCode; // A: TO
        mappedRowSO[1] = version;            // B: Version
        mappedRowSO[2] = colorCode;          // C: MÀU COLOR
        mappedRowSO[3] = customerCode;       // D: Khách hàng CUSTOMER
        mappedRowSO[4] = project;            // E: Dự án PROJECT
        mappedRowSO[5] = toMau;              // F: TO MÀU
        mappedRowSO[6] = newTanXuatTO;       // G: Tần xuất TO (đếm được)
        mappedRowSO[7] = 1;                  // H: Tần xuất MÀU theo TO (lần đầu màu này)
        mappedRowSO[8] = toCustomer;         // I: TO theo khách hàng
        mappedRowSO[9] = newTanXuatKhach;    // J: Tần xuất theo khách (đếm được)
        mappedRowSO[10] = soNo;               // K: Số SO
        mappedRowSO[11] = "";                 // L: Tình trạng phát hành

        sheetSO.appendRow(mappedRowSO);
      }
    }


    // Trả về kết quả chi tiết
    var resultMsg = toCustomerExists 
      ? "Lưu thành công! (Chỉ cập nhật SO, TO [" + toCodeNormalized + "] - Customer [" + customerNameNormalized + "] đã tồn tại)"
      : "Lưu thành công! (Lưu Data & cập nhật/tạo SO)";
    
    return { 
      success: true, 
      message: resultMsg,
      dataOnly: !toCustomerExists,
      soOnly: toCustomerExists
    };

  } catch (err) {
    return { success: false, error: err.toString() };
  }
}

// ========================================================================
// HỆ THỐNG QUẢN LÍ BẢN VẼ - MANAGED DRAWINGS
// ========================================================================

function getManagedDrawings() {
  try {
    var sheetId = '1DRteBSFT1cj4R_OUPMoDxeLMzAIJexWF3HPT-rpMOoM';
    var ss = SpreadsheetApp.openById(sheetId);
    var dataSheet = ss.getSheetByName('Data');

    if (!dataSheet || dataSheet.getLastRow() <= 1) {
      return { success: true, data: [] };
    }

    var data = dataSheet.getRange(2, 1, dataSheet.getLastRow() - 1, dataSheet.getLastColumn()).getValues();
    var managed = [];

    for (var i = 0; i < data.length; i++) {
      var row = data[i] || [];

      // Kiểm tra nếu có dữ liệu đủ (AC hoặc AG không rỗng)
      var acValue = row[28];  // AC: Width
      var agValue = row[32];  // AG: Hình mặt cắt

      // Chuyển sang string an toàn để tránh CellImage object
      var acStr = acValue ? String(acValue).trim() : '';
      var agStr = normalizeImageSource(agValue);

      if ((acStr !== '') || (agStr !== '')) {
        var typeVal = row[3] || 'New';        // D: Type
        var toVal = row[8] || 'N/A';          // I: TO
        var dwCodeVal = row[11] || '';        // L: DW Code
        var customerVal = row[10] || 'Unknown'; // K: Customer

        managed.push({
          type: String(typeVal),
          to: String(toVal),
          dwCode: String(dwCodeVal),
          customer: String(customerVal),
          cutDrawing: agStr,    // Already converted to string
          so: String(row[33] || ''),
          rowIdx: i + 2                 // Row index (1-based)
        });
      }
    }

    return { success: true, data: managed.reverse() };
  } catch (err) {
    return { success: false, error: err.toString() };
  }
}

function removeManagedDrawing(rowId) {
  try {
    var sheetId = '1DRteBSFT1cj4R_OUPMoDxeLMzAIJexWF3HPT-rpMOoM';
    var ss = SpreadsheetApp.openById(sheetId);
    var dataSheet = ss.getSheetByName('Data');

    if (!dataSheet || rowId < 2) {
      return { success: false, error: 'Invalid row ID' };
    }

    // Dữ liệu thực tế lưu ở AC, AD, AG => cột 29, 30, 33 (1-based)
    dataSheet.getRange(rowId, 29).clearContent();  // AC: Width
    dataSheet.getRange(rowId, 30).clearContent();  // AD: Height
    dataSheet.getRange(rowId, 33).clearContent();  // AG: Image

    return {
      success: true,
      message: 'Bản vẽ đã được chuyển về "Thêm bản vẽ"'
    };
  } catch (err) {
    return { success: false, error: err.toString() };
  }
}

// Fetch dữ liệu đầy đủ 1 row từ Data sheet
function getDrawingByRowIdx(rowIdx) {
  try {
    var sheetId = '1DRteBSFT1cj4R_OUPMoDxeLMzAIJexWF3HPT-rpMOoM';
    var ss = SpreadsheetApp.openById(sheetId);
    var dataSheet = ss.getSheetByName('Data');

    if (!dataSheet || rowIdx < 2) {
      return { success: false, error: 'Invalid row index' };
    }

    var rowData = dataSheet.getRange(rowIdx, 1, 1, dataSheet.getLastColumn()).getValues()[0];
    
    // Ánh xạ chính xác theo thứ tự 34 cột (A->AH) của sheet Data:
    // A(0): ID, B(1): Mail, C(2): Status, D(3): Group, E(4): Type, F(5): Version,
    // G(6): Ngày tiếp nhận, H(7): Ngày hoàn thành dự kiến, I(8): TO, J(9): Dự án,
    // K(10): Customer, L(11): Drawing code, M(12): Type DW, N(13): Nội dung thay đổi,
    // O(14): File PDF, P(15): Người đảm trách, Q(16): Ngày HT người đảm trách,
    // R(17): Checker, S(18): Ngày HT thực tế, T(19): Approver, U(20): Ngày PH thực tế,
    // V(21): Status running, W(22): Note, X(23): FYE, Y(24): Type Drawing,
    // Z(25): Type Product, AA(26): Product name, AB(27): Product code,
    // AC(28): W, AD(29): H, AE(30): Excel, AF(31): PDF, AG(32): Image, AH(33): Mã SO
    return {
      success: true,
      data: {
        group: String(rowData[3] || ''),             // D: Group
        type: String(rowData[4] || ''),              // E: Type
        version: String(rowData[5] || ''),           // F: Version / Revise
        receivedDate: String(rowData[6] || ''),     // G: Ngày tiếp nhận
        assigneeDoneDate: String(rowData[7] || ''), // H: Ngày HT dự kiến
        to: String(rowData[8] || ''),                // I: TO
        project: String(rowData[9] || ''),           // J: Dự án
        customer: String(rowData[10] || ''),         // K: Customer
        dwCode: String(rowData[11] || ''),           // L: Drawing Code
        typeDw: String(rowData[12] || rowData[24] || ''), // M / Y: Type DW
        assignee: String(rowData[15] || ''),         // P: Người đảm trách
        actualDoneDate: String(rowData[18] || ''),   // S: Ngày hoàn thành thực tế
        note: String(rowData[22] || ''),             // W: Note / Change log
        fye: String(rowData[23] || ''),              // X: FYE
        width: String(rowData[28] || ''),            // AC: Width
        height: String(rowData[29] || ''),           // AD: Height
        image: rowData[32],                          // AG: Image (CellImage)
        so: String(rowData[33] || '')                // AH: Mã SO
      }
    };
  } catch (err) {
    return { success: false, error: err.toString() };
  }
}

function getDrawingsWithoutImage() {
  try {
    var sheetId = '1DRteBSFT1cj4R_OUPMoDxeLMzAIJexWF3HPT-rpMOoM';
    var ss = SpreadsheetApp.openById(sheetId);
    var logSheet = ss.getSheetByName('Log');
    var dataSheet = ss.getSheetByName('Data');

    var existingSoMap = {};
    if (dataSheet && dataSheet.getLastRow() > 1) {
      var dataVals = dataSheet.getRange(2, 1, dataSheet.getLastRow() - 1, dataSheet.getLastColumn()).getValues();
      for (var d = 0; d < dataVals.length; d++) {
        var r = dataVals[d] || [];
        var soCode = '';
        if (r[33]) soCode = toSafeString(r[33]);
        else if (r[5]) soCode = toSafeString(r[5]);

        if (!soCode) {
          for (var c = 0; c < r.length; c++) {
            var cellStr = toSafeString(r[c]);
            if (/^SO\d+/i.test(cellStr)) {
              soCode = cellStr;
              break;
            }
          }
        }

        if (soCode) {
          var cleanSo = soCode.toUpperCase().replace(/\s+/g, "");
          var digitsSo = cleanSo.replace(/[^0-9]/g, "");
          existingSoMap[cleanSo] = true;
          if (digitsSo) existingSoMap[digitsSo] = true;
        }
      }
    }

    if (logSheet && logSheet.getLastRow() > 1) {
      var logValues = logSheet.getDataRange().getValues();
      var pending = [];

      for (var i = 1; i < logValues.length; i++) {
        if (String(logValues[i][6] || '').trim() !== 'Đã tiếp nhận') continue;

        var rawSo = String(logValues[i][5] || '').trim();
        var cleanSo = rawSo.toUpperCase().replace(/\s+/g, "");
        var digitsSo = cleanSo.replace(/[^0-9]/g, "");

        if ((rawSo && existingSoMap[cleanSo]) || (digitsSo && existingSoMap[digitsSo])) {
          continue;
        }

        pending.push({
          rowIdx: i + 1,
          type: 'New',
          to: String(logValues[i][4] || 'N/A'),
          dwCode: '',
          customer: String(logValues[i][3] || 'Unknown'),
          subject: String(logValues[i][2] || ''),
          so: rawSo,
          existsInData: false,
          imageUrl: ''
        });
      }

      return { success: true, data: pending.reverse() };
    }

    if (!dataSheet || dataSheet.getLastRow() <= 1) {
      return { success: true, data: [] };
    }

    var data = dataSheet.getRange(2, 1, dataSheet.getLastRow() - 1, dataSheet.getLastColumn()).getValues();
    var pending = [];

    for (var i = 0; i < data.length; i++) {
      var row = data[i] || [];
      var width = row[28];
      var height = row[29];
      var widthStr = width ? String(width).trim() : '';
      var heightStr = height ? String(height).trim() : '';

      if ((widthStr === '') && (heightStr === '')) {
        var type = row[3] || 'New';
        var to = row[8] || 'N/A';
        var dwCode = row[11] || '';
        var customer = row[10] || 'Unknown';

        pending.push({
          rowIdx: i + 2,
          type: String(type),
          to: String(to),
          dwCode: String(dwCode),
          customer: String(customer)
        });
      }
    }

    return { success: true, data: pending.reverse() };
  } catch (err) {
    return { success: false, error: err.toString() };
  }
}

function updateBulkDrawing(rowIds, bulkData) {
  try {
    var sheetId = '1DRteBSFT1cj4R_OUPMoDxeLMzAIJexWF3HPT-rpMOoM';
    var ss = SpreadsheetApp.openById(sheetId);
    var dataSheet = ss.getSheetByName('Data');
    
    if (!dataSheet || !rowIds || rowIds.length === 0) {
      return { success: false, error: 'Invalid row IDs or no data' };
    }

    // Column mapping
    // B=2: Type, D=4: TO, E=5: Project, F=6: SO, G=7: Dw Code, H=8: Version
    // I=9: Type DW, J=10: Assignee, N=14: Assignee Done Date, P=16: Status
    
    for (var idx = 0; idx < rowIds.length; idx++) {
      var rowId = rowIds[idx];
      if (rowId < 2) continue;

      // Cập nhật các field từ bulkData
      if (bulkData.type) dataSheet.getRange(rowId, 2).setValue(bulkData.type);
      if (bulkData.version) dataSheet.getRange(rowId, 8).setValue(bulkData.version);
      if (bulkData.typeDw) dataSheet.getRange(rowId, 9).setValue(bulkData.typeDw);
      if (bulkData.assignee) dataSheet.getRange(rowId, 10).setValue(bulkData.assignee);
      if (bulkData.assigneeDoneDate) dataSheet.getRange(rowId, 14).setValue(bulkData.assigneeDoneDate);
      if (bulkData.status) dataSheet.getRange(rowId, 16).setValue(bulkData.status);
    }

    return { 
      success: true, 
      message: 'Đã cập nhật ' + rowIds.length + ' bản vẽ' 
    };
  } catch (err) {
    return { success: false, error: err.toString() };
  }
}