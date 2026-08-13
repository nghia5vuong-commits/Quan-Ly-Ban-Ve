var MAIL_CONFIG = {
  SUBJECT_FILTERS: ['MANUFACTURING ORDER', 'NEW PROJECT', 'THÔNG BÁO TỰ ĐỘNG'],
  MAX_THREADS    : 50,
  TIMEZONE       : 'Asia/Ho_Chi_Minh'
};

function buildQuery(keyword) {
  var query = 'subject:("MANUFACTURING ORDER" OR "NEW PROJECT" OR "THÔNG BÁO TỰ ĐỘNG")';
  if (keyword && keyword.trim() !== '') {
    query = '(' + query + ') ' + keyword.trim();
  }
  return query;
}

// 🚀 TỐI ƯU 1: HÀM TẢI DANH SÁCH (Nhanh hơn x10 lần)
function getManufacturingEmails(page, keyword) {
  try {
    page = page || 0;
    keyword = keyword || '';
    var query = buildQuery(keyword);
    var limit = MAIL_CONFIG.MAX_THREADS;
    var start = page * limit;

    // Chỉ load MAX_THREADS + 1 để biết có trang tiếp theo không (Bỏ hẳn lệnh load 500)
    var threads = GmailApp.search(query, start, limit + 1);
    var hasMore = threads.length > limit;
    if (hasMore) threads.pop(); // Loại bỏ thread dư thừa lấy làm mốc

    if (threads.length === 0) return { success: true, emails: [], total: 0, page: page, hasMore: false };

    // TỐI ƯU CỐT LÕI: Lấy messages của TOÀN BỘ threads bằng 1 lệnh duy nhất (Xóa sổ lỗi N+1 query)
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
        var messages = threadsMessages[i]; // Lấy từ mảng đã truy xuất 1 lần
        if (!messages || messages.length === 0) continue;
        
        var lastMsg = messages[messages.length - 1];
        var from = lastMsg.getFrom() || '';
        
        emails.push({
          id          : thread.getId(),
          subject     : lastMsg.getSubject() || '(Không có tiêu đề)',
          sender      : from,
          senderName  : extractName(from),
          senderEmail : extractEmail(from),
          date        : formatDate(lastMsg.getDate()),
          dateRaw     : lastMsg.getDate().getTime(),
          isUnread    : thread.isUnread(),
          messageCount: messages.length,
          // TỐI ƯU KHỦNG: Tắt getAttachmentInfo ở màn hình danh sách. Nó gây giật lag vì tải dữ liệu blob file.
          attachments : [], 
          tag         : detectTag(lastMsg.getSubject())
        });
      } catch (err) {}
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

    var messageList = messages.map(function(msg, idx) {
      return {
        index       : idx,
        messageId   : msg.getId(),
        subject     : msg.getSubject(),
        from        : msg.getFrom(),
        senderName  : extractName(msg.getFrom()),
        senderEmail : extractEmail(msg.getFrom()),
        to          : msg.getTo(),
        cc          : msg.getCc(),
        date        : formatDate(msg.getDate()),
        dateRaw     : msg.getDate().getTime(),
        body        : sanitizeHtml(msg.getBody()),
        plainBody   : msg.getPlainBody(),
        isUnread    : msg.isUnread(),
        attachments : getAttachmentInfo(msg) // Chỉ parse file khi click mở email
      };
    });

    thread.markRead();
    var lastMsg = messages[messages.length - 1];

    return {
      success     : true,
      threadId    : threadId,
      subject     : lastMsg.getSubject(),
      messageCount: messages.length,
      messages    : messageList
    };
  } catch (err) {
    return { success: false, error: err.toString() };
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

    allThreads.forEach(function(t) {
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
    return message.getAttachments().map(function(att, idx) {
      return {
        index: idx, name: att.getName(), contentType: att.getContentType(),
        size: formatFileSize(att.getSize()), isSheet: isSpreadsheetType(att.getName(), att.getContentType())
      };
    });
  } catch(e) { return []; }
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

function getAttachmentParsedData(threadId, messageIndex, attachmentIndex) {
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
    var sheetId = '19U_o7QYsAqNRJ7NHwYXH2UPY8pLnndCcsHCyB_qRNeg';
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
    var sheetId = '19U_o7QYsAqNRJ7NHwYXH2UPY8pLnndCcsHCyB_qRNeg';
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
    var sheetId = '19U_o7QYsAqNRJ7NHwYXH2UPY8pLnndCcsHCyB_qRNeg';
    var ss = SpreadsheetApp.openById(sheetId);
    var logSheet = ss.getSheetByName('Log');
    if (!logSheet) return { success: true, data: [] };

    var data = logSheet.getDataRange().getDisplayValues();
    var logs = [];
    
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][6]).trim() === 'Đã tiếp nhận') {
        logs.push({
          rowIdx: i + 1, 
          date: data[i][0],
          threadId: data[i][1],
          subject: data[i][2],
          customer: data[i][3],
          to: data[i][4],
          so: data[i][5]
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
    var sheetId = '19U_o7QYsAqNRJ7NHwYXH2UPY8pLnndCcsHCyB_qRNeg';
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
          } catch (e) {}
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