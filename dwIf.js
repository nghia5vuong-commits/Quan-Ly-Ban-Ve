var MAIL_CONFIG = {
  SUBJECT_FILTERS: ['MANUFACTURING ORDER', 'NEW PROJECT', 'DIE REQUEST', 'NEW CUSTOMER', 'NEW CUSTOMER + NEW PROJECT'],
  MAX_THREADS: 50,
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
      var imageUrl = typeof value.getContentUrl === 'function' ? value.getContentUrl() : '';
      if (imageUrl) {
        var response = UrlFetchApp.fetch(imageUrl, {
          headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() },
          muteHttpExceptions: true
        });
        if (response.getResponseCode() === 200) {
          var blob = response.getBlob();
          return 'data:' + blob.getContentType() + ';base64,' + Utilities.base64Encode(blob.getBytes());
        }
      }
    } catch (err) {
      Logger.log('Lỗi đọc CellImage AG: ' + err.toString());
    }
    return '';
  }

  return '';
}

function buildQuery(keyword) {
  var query = 'subject:("MANUFACTURING ORDER" OR "NEW PROJECT" OR "DIE REQUEST" OR "NEW CUSTOMER" OR "NEW CUSTOMER + NEW PROJECT")';
  if (keyword && keyword.trim() !== '') {
    query = '(' + query + ') ' + keyword.trim();
  }
  return query;
}

function normalizeLogTo(value) {
  return String(value || '').replace(/^TO-?/i, '').replace(/\s+/g, '').trim().toUpperCase();
}

function normalizeLogCustomer(value) {
  return String(value || '').replace(/\s+/g, ' ').trim().toUpperCase();
}

function normalizeMaterialTo(value) {
  var material = String(value || '').trim();
  if (material.length <= 7) return '';
  // Ví dụ ATO-FV7C519 60000 -> TO-FV7C519.
  return material.substring(1, material.length - 6).trim().toUpperCase();
}

function getLogToValues(value, stripMaterialCharacter) {
  return splitDrawingToValues(value).map(function (toValue) {
    var normalized = normalizeLogTo(toValue);
    return stripMaterialCharacter && normalized.length > 1 ? normalized.substring(1) : normalized;
  }).filter(Boolean);
}

function getLogReceiptProgress(logToValues, customer, dataRows) {
  var toValues = Array.from(new Set((logToValues || []).map(normalizeLogTo).filter(Boolean)));
  var customerValue = normalizeLogCustomer(customer);
  var receivedSet = new Set();

  (dataRows || []).forEach(function (row) {
    if (normalizeLogCustomer(row[10]) !== customerValue) return;
    getLogToValues(row[8], false).forEach(function (toValue) { receivedSet.add(toValue); });
  });

  var receivedCount = toValues.filter(function (toValue) { return receivedSet.has(toValue); }).length;
  return {
    total: toValues.length,
    received: receivedCount,
    complete: toValues.length > 0 && receivedCount >= toValues.length,
    label: toValues.length > 0 && receivedCount >= toValues.length ? 'Đã tiếp nhận' : receivedCount + '/' + toValues.length
  };
}

function getLogToClassification(toValues, customer, dataRows, isMaterialSource) {
  var customerValue = normalizeLogCustomer(customer);
  var dataToSet = new Set();
  (dataRows || []).forEach(function (row) {
    if (normalizeLogCustomer(row[10]) !== customerValue) return;
    // TO OLD chỉ cần tồn tại trong Data cùng Customer, không phụ thuộc Width/Height/hình ảnh.
    getLogToValues(row[8], false).forEach(function (toValue) { dataToSet.add(toValue); });
  });

  var newValues = [];
  var oldValues = [];
  Array.from(new Set((toValues || []).map(function (rawValue) {
    var rawText = String(rawValue || '').trim();
    var materialValue = isMaterialSource ? normalizeMaterialTo(rawText) : '';
    var normalized = materialValue || normalizeLogTo(rawText);
    var displayValue = /^TO-?/i.test(normalized) ? normalized : 'TO-' + normalized;
    return {
      displayValue: displayValue,
      candidates: normalized ? [normalized, normalizeLogTo(normalized)] : []
    };
  }).filter(function (entry) { return entry.displayValue; }).map(function (entry) {
    return JSON.stringify(entry);
  }))).forEach(function (serializedValue) {
    var value = JSON.parse(serializedValue);
    var isOld = value.candidates.some(function (candidate) { return dataToSet.has(candidate); });
    (isOld ? oldValues : newValues).push(value.displayValue);
  });
  return { newValues: newValues, oldValues: oldValues };
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
    var threadSoMap = {};
    var drawingSoSet = new Set();
    var checkAdjustSoSet = typeof getCheckAdjustSoSet === 'function' ? getCheckAdjustSoSet() : new Set();
    var emailProgressMap = {};

    var dataSheet = ss.getSheetByName('Data');
    var dataRows = [];
    if (dataSheet && dataSheet.getLastRow() >= 2) {
      dataRows = dataSheet.getRange(2, 1, dataSheet.getLastRow() - 1, Math.max(34, dataSheet.getLastColumn())).getValues();
      var drawingRows = dataSheet.getRange(2, 34, dataSheet.getLastRow() - 1, 1).getDisplayValues();
      drawingRows.forEach(function (row) {
        var normalizedSo = typeof normalizeComparisonSo === 'function' ? normalizeComparisonSo(row[0]) : String(row[0] || '').trim().toUpperCase();
        if (normalizedSo) drawingSoSet.add(normalizedSo);
      });
    }

    if (logSheet) {
      var data = logSheet.getDataRange().getValues();
      for (var r = 1; r < data.length; r++) {
        var logRow = data[r];
        var logThreadId = String(logRow[1] || '').trim();
        if (!logThreadId) continue;
        var progress = getLogReceiptProgress(getLogToValues(logRow[4], true), logRow[3], dataRows);
        emailProgressMap[logThreadId] = progress;
        var logSo = typeof normalizeComparisonSo === 'function' ? normalizeComparisonSo(logRow[5]) : String(logRow[5] || '').trim().toUpperCase();
        if (logSo) threadSoMap[logThreadId] = logSo;
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

          receivedProgress: emailProgressMap[threadId] || { total: 0, received: 0, complete: false, label: 'Chưa tiếp nhận' },
          isInLog: !!emailProgressMap[threadId],
          isReceived: !!(emailProgressMap[threadId] && emailProgressMap[threadId].complete),
          isDrawingDone: !!(threadSoMap[threadId] && checkAdjustSoSet.has(threadSoMap[threadId])),
          drawingStatus: threadSoMap[threadId] && checkAdjustSoSet.has(threadSoMap[threadId]) ? 'Check SO OK' : 'Không có SO'
        });
      } catch (err) { }
    }

    var result = {
      success: true,
      emails: emails,
      total: hasMore ? (start + limit + "+") : (start + emails.length),
      page: page,
      hasMore: hasMore
    };
    return result;
  } catch (err) {
    return { success: false, error: err.toString(), emails: [], total: 0 };
  }
}

// 🚀 TỐI ƯU 2: HÀM LẤY CHI TIẾT EMAIL (Giữ nguyên phần đính kèm ở đây vì người dùng thực sự đang xem nó)
function getEmailDetail(threadId) {
  try {
    var cache = CacheService.getScriptCache();
    var cacheKey = 'MAIL_DETAIL_' + String(threadId);
    var cached = cache.get(cacheKey);
    if (cached) {
      return JSON.parse(cached);
    }

    var thread = GmailApp.getThreadById(threadId);
    var messages = thread.getMessages();

    var messageList = messages.map(function (msg, idx) {
      var plainBody = String(msg.getPlainBody() || '');
      var safePlainBody = sanitizePlainBody(plainBody);
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
        plainBody: safePlainBody,
        isUnread: msg.isUnread(),
        attachments: getAttachmentInfo(msg) // Chỉ parse file khi click mở email
      };
    });

    thread.markRead();
    var lastMsg = messages[messages.length - 1];

    var isInLog = false;
    try {
      var ss = SpreadsheetApp.openById('1DRteBSFT1cj4R_OUPMoDxeLMzAIJexWF3HPT-rpMOoM');
      var logSheet = ss.getSheetByName('Log');
      if (logSheet && logSheet.getLastRow() > 1) {
        var threadValues = logSheet.getRange(2, 2, logSheet.getLastRow() - 1, 1).getValues();
        for (var tIdx = 0; tIdx < threadValues.length; tIdx++) {
          if (String(threadValues[tIdx][0] || '').trim() === String(threadId || '').trim()) {
            isInLog = true;
            break;
          }
        }
      }
    } catch (eLog) { }

    var detail = {
      success: true,
      threadId: threadId,
      isInLog: isInLog,
      subject: lastMsg.getSubject(),
      messageCount: messages.length,
      messages: messageList
    };

    // CacheService giới hạn kích thước value khoảng 100 KB. Thread nhiều mail
    // có thể vượt giới hạn dù nội dung từng message đã được rút gọn.
    var detailJson = JSON.stringify(detail);
    if (detailJson.length <= 90000) {
      try {
        cache.put(cacheKey, detailJson, 300);
      } catch (cacheError) {
        Logger.log('Bỏ qua cache chi tiết email quá lớn: ' + cacheError.toString());
      }
    }
    return detail;
  } catch (err) {
    return { success: false, error: err.toString() };
  }
}

function sanitizePlainBody(text) {
  try {
    return String(text || '').replace(/\r\n/g, '\n').replace(/\n{3,}/g, '\n\n').slice(0, 12000);
  } catch (e) {
    return String(text || '');
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
    var name = String(attachment.getName() || 'attachment');
    var mimeType = attachment.getContentType() || guessMimeType(name);
    var size = attachment.getSize() || 0;
    if (size > 5 * 1024 * 1024) {
      return {
        success: false,
        error: 'File đính kèm quá lớn để parse trên trình duyệt. Vui lòng dùng file nhỏ hơn 5MB.'
      };
    }

    var base64Data = Utilities.base64Encode(attachment.getBytes());

    return {
      success: true,
      dataBase64: base64Data,
      fileName: name,
      mimeType: mimeType
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
      var lastRow = logSheet.getLastRow();
      if (lastRow >= 2) {
        var threadIdValues = logSheet.getRange(2, 2, lastRow - 1, 1).getValues();
        for (var r = 0; r < threadIdValues.length; r++) {
          var val = String(threadIdValues[r][0] || '').trim();
          if (val) {
            receivedThreadIds.add(val);
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
        index: idx,
        name: att.getName(),
        contentType: att.getContentType(),
        size: formatFileSize(att.getSize()),
        isSheet: isSpreadsheetType(att.getName(), att.getContentType()),
        isPdf: isPdfType(att.getName(), att.getContentType())
      };
    });
  } catch (e) { return []; }
}

function isSpreadsheetType(name, contentType) {
  var n = (name || '').toLowerCase(), t = (contentType || '').toLowerCase();
  return !!(n.match(/\.(xlsx|xls|csv|ods|xlsm|xlsb)$/) || t.indexOf('spreadsheet') !== -1 || t.indexOf('excel') !== -1 || t.indexOf('csv') !== -1);
}

function isPdfType(name, contentType) {
  var n = (name || '').toLowerCase(), t = (contentType || '').toLowerCase();
  return !!(n.match(/\.pdf$/) || t.indexOf('pdf') !== -1);
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
    var bytes = att.getBytes();
    if (bytes.length > 5 * 1024 * 1024) {
      return { success: false, error: 'File đính kèm quá lớn để xem/preview trên trình duyệt. Vui lòng dùng file nhỏ hơn 5MB.' };
    }
    return {
      success: true,
      dataBase64: Utilities.base64Encode(bytes),
      fileName: att.getName(),
      mimeType: att.getContentType() || guessMimeType(att.getName())
    };
  } catch (err) {
    return { success: false, error: err.toString() };
  }
}

function guessMimeType(fileName) {
  try {
    var name = String(fileName || '').toLowerCase();
    if (name.endsWith('.pdf')) return 'application/pdf';
    if (name.endsWith('.png')) return 'image/png';
    if (name.endsWith('.jpg') || name.endsWith('.jpeg')) return 'image/jpeg';
    if (name.endsWith('.gif')) return 'image/gif';
    if (name.endsWith('.webp')) return 'image/webp';
    if (name.endsWith('.xlsx')) return 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
    if (name.endsWith('.xls')) return 'application/vnd.ms-excel';
    if (name.endsWith('.csv')) return 'text/csv';
    if (name.endsWith('.zip')) return 'application/zip';
    return 'application/octet-stream';
  } catch (e) {
    return 'application/octet-stream';
  }
}

function getAttachmentForDownloadByUrl(fileUrl) {
  try {
    if (!fileUrl) return { success: false, error: 'Thiếu đường link tải file.' };

    var url = String(fileUrl).trim();
    var response = UrlFetchApp.fetch(url, {
      method: 'get',
      muteHttpExceptions: true,
      followRedirects: true
    });

    var status = response.getResponseCode();
    if (status < 200 || status >= 300) {
      return { success: false, error: 'Không tải được file từ link: HTTP ' + status };
    }

    var blob = response.getBlob();
    var fileName = blob.getName() || url.split('/').pop() || 'remote-file';
    var mimeType = blob.getContentType() || guessMimeType(fileName);
    var bytes = blob.getBytes();
    if (bytes.length > 5 * 1024 * 1024) {
      return { success: false, error: 'File từ link quá lớn để preview trên trình duyệt. Vui lòng dùng file nhỏ hơn 5MB.' };
    }

    return {
      success: true,
      dataBase64: Utilities.base64Encode(bytes),
      fileName: fileName,
      mimeType: mimeType
    };
  } catch (err) {
    return { success: false, error: err.toString() };
  }
}

function saveToLogSheet(threadId, subject, parsedData, fullDataString, clientUserEmail) {
  try {
    var sheetId = '1DRteBSFT1cj4R_OUPMoDxeLMzAIJexWF3HPT-rpMOoM';
    var ss = SpreadsheetApp.openById(sheetId);
    var logSheet = ss.getSheetByName('Log');

    if (!logSheet) {
      logSheet = ss.insertSheet('Log');
    }

    var headers = logSheet.getRange(1, 1, 1, Math.max(11, logSheet.getLastColumn())).getValues()[0];
    if (!headers[8]) logSheet.getRange(1, 9).setValue('TO NEW');
    if (!headers[9]) logSheet.getRange(1, 10).setValue('TO OLD');
    if (!headers[10]) logSheet.getRange(1, 11).setValue('NGƯỜI TIẾP NHẬN');

    var timestamp = Utilities.formatDate(new Date(), MAIL_CONFIG.TIMEZONE, 'dd/MM/yyyy HH:mm:ss');
    var isMaterialSource = Array.isArray(parsedData.materialTos);
    var sourceToValues = isMaterialSource ? parsedData.materialTos : (Array.isArray(parsedData.tos) ? parsedData.tos : [parsedData.to || '']);
    var toValues = sourceToValues.map(function (value) {
      return isMaterialSource ? normalizeMaterialTo(value) : normalizeLogTo(value);
    }).filter(Boolean);
    var toValue = toValues.join(', ');
    var dataSheet = ss.getSheetByName('Data');
    var dataRows = dataSheet && dataSheet.getLastRow() > 1
      ? dataSheet.getRange(2, 1, dataSheet.getLastRow() - 1, Math.max(34, dataSheet.getLastColumn())).getValues()
      : [];
    var classification = getLogToClassification(toValues, parsedData.customer || '', dataRows, false);
    var progress = getLogReceiptProgress(toValues, parsedData.customer || '', dataRows);

    var receiverEmail = String(clientUserEmail || getCurrentUserEmail() || '').trim();

    var existingRow = -1;
    if (logSheet.getLastRow() >= 2) {
      var threadValues = logSheet.getRange(2, 2, logSheet.getLastRow() - 1, 1).getValues();
      for (var threadIndex = 0; threadIndex < threadValues.length; threadIndex++) {
        if (String(threadValues[threadIndex][0] || '').trim() === String(threadId || '').trim()) {
          existingRow = threadIndex + 2;
          break;
        }
      }
    }

    // Nếu không có receiverEmail mới nhưng dòng cũ đã có thì giữ lại
    if (!receiverEmail && existingRow > 1) {
      var prevReceiver = logSheet.getRange(existingRow, 11).getValue();
      if (!prevReceiver) {
        var prevStatus = String(logSheet.getRange(existingRow, 7).getValue() || '');
        var match = prevStatus.match(/-\s*([a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,})/i);
        if (match) prevReceiver = match[1];
      }
      if (prevReceiver) receiverEmail = String(prevReceiver).trim();
    }

    var statusLabel = progress.label + (receiverEmail ? ' - ' + receiverEmail : '');

    var rowData = [
      timestamp,
      threadId,
      subject,
      parsedData.customer || '',
      classification.newValues.join(', '),
      parsedData.so || '',
      statusLabel,
      fullDataString || '',
      classification.newValues.join(', '),
      classification.oldValues.join(', '),
      receiverEmail
    ];

    if (existingRow > 1) logSheet.getRange(existingRow, 1, 1, rowData.length).setValues([rowData]);
    else logSheet.getRange(logSheet.getLastRow() + 1, 1, 1, rowData.length).setValues([rowData]);
    var cache = CacheService.getScriptCache();
    cache.remove("MAIL_STATS");
    cache.remove("MAIL_LIST_0_");
    try { cache.remove("MAIL_DETAIL_" + String(threadId)); } catch (e) { }

    return { success: true, message: 'Đã lưu tiếp nhận thiết kế thành công!', progress: progress, receiverEmail: receiverEmail };
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
    var logRows = logSheet.getRange(2, 1, lastRow - 1, Math.max(10, logSheet.getLastColumn())).getValues();
    var dataSheet = ss.getSheetByName('Data');
    var dataRows = dataSheet && dataSheet.getLastRow() > 1
      ? dataSheet.getRange(2, 1, dataSheet.getLastRow() - 1, Math.max(34, dataSheet.getLastColumn())).getValues()
      : [];

    for (var i = 0; i < logRows.length; i++) {
      if (String(logRows[i][1]).trim() === String(threadId).trim()) {
        var progress = getLogReceiptProgress(getLogToValues(logRows[i][4], true), logRows[i][3], dataRows);
        var receivedDate = logSheet.getRange(i + 2, 1).getDisplayValue();
        return { success: true, isReceived: progress.complete, progress: progress, receivedDate: receivedDate };
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
    var dataRows = [];
    var existingSoMap = {};
    if (dataSheet && dataSheet.getLastRow() > 1) {
      var dataVals = dataSheet.getRange(2, 1, dataSheet.getLastRow() - 1, dataSheet.getLastColumn()).getValues();
      dataRows = dataVals;
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
      var rawSo = (data[i][5] || "").toString().trim();
      var newToValues = getLogToClassification(splitDrawingToValues(data[i][4] || ''), data[i][3], dataRows).newValues;
      newToValues.forEach(function (toValue) {
        var cleanRaw = rawSo.toUpperCase().replace(/\s+/g, "");
        var digitsRaw = cleanRaw.replace(/[^0-9]/g, "");
        var foundInfo = rawSo ? (existingSoMap[cleanRaw] || (digitsRaw ? existingSoMap[digitsRaw] : null)) : null;
        logs.push({
          rowIdx: i + 1,
          date: String(data[i][0] || ''),
          threadId: String(data[i][1] || ''),
          subject: String(data[i][2] || ''),
          customer: String(data[i][3] || ''),
          to: String(toValue || ''),
          so: rawSo,
          existsInData: !!foundInfo,
          imageUrl: typeof foundInfo === 'string' ? foundInfo : ''
        });
      });
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

function getDrawingSyncHistory_() {
  try {
    var raw = PropertiesService.getUserProperties().getProperty('drawing_sync_history');
    return raw ? JSON.parse(raw) : {};
  } catch (err) {
    return {};
  }
}

function rememberDrawingSyncResult_(operationId, result) {
  if (!operationId) return;
  try {
    var history = getDrawingSyncHistory_();
    history[String(operationId)] = result;
    var keys = Object.keys(history);
    while (keys.length > 50) {
      delete history[keys.shift()];
    }
    PropertiesService.getUserProperties().setProperty('drawing_sync_history', JSON.stringify(history));
  } catch (err) {
    Logger.log('Không thể lưu lịch sử đồng bộ bản vẽ: ' + err.toString());
  }
}

function saveDataToTestSheet(subject, matrixData, rowIdx, operationId) {
  try {
    var normalizedOperationId = String(operationId || '').trim();
    var syncHistory = normalizedOperationId ? getDrawingSyncHistory_() : {};
    if (normalizedOperationId && syncHistory[normalizedOperationId]) {
      return syncHistory[normalizedOperationId];
    }

    var ss = SpreadsheetApp.openById("1DRteBSFT1cj4R_OUPMoDxeLMzAIJexWF3HPT-rpMOoM");
    var sectionImageFolderId = "19Faip8STiBLJ58aVLiqwVs4SaotKb1bm";
    var sheetData = ss.getSheetByName("Data") || ss.getSheetByName("data") || ss.insertSheet("Data");
    var sheetSO = ss.getSheetByName("SO");
    var sheetCustomer = ss.getSheetByName("Customer");

    if (!matrixData || matrixData.length === 0) {
      return { success: false, error: "Dữ liệu gửi lên rỗng!" };
    }

    // Nếu request đã ghi Data nhưng client mất phản hồi, retry phải trả về thành công
    // thay vì ghi thêm SO lần nữa.
    if (normalizedOperationId && sheetData.getLastRow() > 1) {
      var existingIds = sheetData.getRange(2, 1, sheetData.getLastRow() - 1, 1).getValues();
      var alreadyWritten = existingIds.some(function (row) {
        return String(row[0] || '').trim() === normalizedOperationId;
      });
      if (alreadyWritten) {
        var duplicateResult = { success: true, operationId: normalizedOperationId, duplicate: true };
        rememberDrawingSyncResult_(normalizedOperationId, duplicateResult);
        return duplicateResult;
      }
    }

    // Kiểm tra xem có phải là chế độ cập nhật (upgrade) hay không
    var isUpgradeMode = rowIdx && typeof rowIdx === 'number' && rowIdx >= 2;

    var formRow = matrixData[0];
    var imageObj = formRow.pop();
    var cellImageObj = null;

    if (imageObj && imageObj.base64) {
      try {
        var imageMimeType = imageObj.mimeType || 'image/png';
        var dataUri = 'data:' + imageMimeType + ';base64,' + imageObj.base64;
        cellImageObj = SpreadsheetApp.newCellImage()
          .setSourceUrl(dataUri)
          .setAltTextTitle(imageObj.name || 'Hinh mat cat')
          .build();
      } catch (imgErr) {
        Logger.log('Lỗi tạo CellImage: ' + imgErr.toString());
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
    var autoId = isUpgradeMode
      ? sheetData.getRange(rowIdx, 1).getValue()
      : (normalizedOperationId || generateUniqueId());
    var userEmail = getCurrentUserEmail();

    // 2. Ghi dữ liệu vào sheet "Data" (33 cột từ A -> AG)
    // Khi Update Version, chỉ thay Version; các trường còn lại phải giữ nguyên.
    var mappedRowData = isUpgradeMode
      ? sheetData.getRange(rowIdx, 1, 1, Math.max(sheetData.getLastColumn(), 34)).getValues()[0]
      : new Array(34).fill("");
    if (isUpgradeMode) {
      mappedRowData[5] = version; // F: Version
      sheetData.getRange(targetRow, 1, 1, mappedRowData.length).setValues([mappedRowData]);
      return { success: true };
    }

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
    mappedRowData[14] = "";
    mappedRowData[11] = dwCode;             // L: Drawing code
    mappedRowData[12] = typeDw;            // M: Version Drawing
    mappedRowData[13] = noteVal;           // N: Nội dung thay đổi
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
    if (cellImageObj) mappedRowData[32] = cellImageObj; // AG: Hình ảnh mặt cắt (chèn trực tiếp vào ô)
    mappedRowData[33] = soNo;              // AH: Mã SO

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
      if (cellImageObj) {
        try { sheetData.getRange(targetRow, 33).setValue(cellImageObj); } catch (e) { }
      }
    }

    // Luôn luôn lưu vào sheet "SO" (dù TO+Customer đã tồn tại hay chưa)
    if (sheetSO) {
      var colorCode = "";
      var cleanTo = String(toCode || "").replace(/^TO-?/i, "").replace(/\s+/g, "").trim().toUpperCase();

      // Tách phần màu trước chữ số đầu tiên và bỏ hậu tố chữ của mã TO.
      // Ví dụ: TO-FV7H748B -> màu FV, TO gốc 7H748.
      var firstDigitMatch = cleanTo.match(/\d/);
      var firstDigitIdx = firstDigitMatch ? firstDigitMatch.index : -1;
      var pureToCode = "";
      if (firstDigitIdx > 0) {
        colorCode = cleanTo.substring(0, firstDigitIdx);
      }
      if (firstDigitIdx >= 0) {
        pureToCode = cleanTo.substring(firstDigitIdx).replace(/[A-Z]+$/, "");
      } else {
        colorCode = cleanTo;
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
      // LUÔN TẠO HÀNG MỚI: Không cập nhật dòng cũ, mà tạo dòng mới mỗi lần
      // ========================================================================
      var soData = sheetSO.getDataRange().getValues();

      // Tìm tần suất lớn nhất nhưng luôn tạo dòng mới, không sửa dòng SO cũ.
      var maxTanXuatTO = 0;
      var toToFind = ("TO-" + pureToCode).toUpperCase();
      var maxTanXuatMau = 0;
      function normalizeSoBaseTo(value) {
        var normalized = String(value || "").replace(/^TO-?/i, "").replace(/\s+/g, "").trim().toUpperCase();
        var digitMatch = normalized.match(/\d/);
        if (!digitMatch) return normalized ? "TO-" + normalized : "";
        return "TO-" + normalized.substring(digitMatch.index).replace(/[A-Z]+$/, "");
      }

      for (var countIdx = 1; countIdx < soData.length; countIdx++) {
        var countRow = soData[countIdx];
        var existingTO = normalizeSoBaseTo(countRow[0]); // Column A: TO
        if (existingTO === toToFind) {
          var rowTanXuatTO = parseInt(countRow[6] || 0) || 0;
          if (rowTanXuatTO > maxTanXuatTO) {
            maxTanXuatTO = rowTanXuatTO;
          }

          if ((countRow[2] || "").toString().trim().toUpperCase() === colorCode) {
            var rowTanXuatMau = parseInt(countRow[7] || 0) || 0;
            if (rowTanXuatMau > maxTanXuatMau) {
              maxTanXuatMau = rowTanXuatMau;
            }
          }
        }
      }
      var newTanXuatTO = maxTanXuatTO + 1;
      var newTanXuatMau = maxTanXuatMau + 1;

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

      // Tạo hàng mới
      var mappedRowSO = new Array(12).fill("");
      mappedRowSO[0] = "TO-" + pureToCode; // A: TO
      mappedRowSO[1] = version;            // B: Version
      mappedRowSO[2] = colorCode;          // C: MÀU COLOR
      mappedRowSO[3] = customerCode;       // D: Khách hàng CUSTOMER
      mappedRowSO[4] = project;            // E: Dự án PROJECT
      mappedRowSO[5] = toMau;              // F: TO MÀU
      mappedRowSO[6] = newTanXuatTO;       // G: Tần xuất TO (đếm được)
      mappedRowSO[7] = newTanXuatMau;      // H: Tần xuất MÀU theo TO
      mappedRowSO[8] = toCustomer;         // I: TO theo khách hàng
      mappedRowSO[9] = newTanXuatKhach;    // J: Tần xuất theo khách (đếm được)
      mappedRowSO[10] = soNo;               // K: Số SO
      mappedRowSO[11] = "";                 // L: Tình trạng phát hành

      sheetSO.appendRow(mappedRowSO);
    }


    // Trả về kết quả chi tiết
    var resultMsg = toCustomerExists
      ? "Lưu thành công! (Chỉ cập nhật SO, TO [" + toCodeNormalized + "] - Customer [" + customerNameNormalized + "] đã tồn tại)"
      : "Lưu thành công! (Lưu Data & cập nhật/tạo SO)";

    var result = {
      success: true,
      message: resultMsg,
      dataOnly: !toCustomerExists,
      soOnly: toCustomerExists
    };
    if (normalizedOperationId) {
      result.operationId = normalizedOperationId;
      rememberDrawingSyncResult_(normalizedOperationId, result);
    }
    return result;

  } catch (err) {
    return { success: false, error: err.toString() };
  }
}

function syncDrawingOperations(operations) {
  if (!Array.isArray(operations)) {
    return { success: false, results: [], error: 'Danh sách thao tác đồng bộ không hợp lệ.' };
  }

  var lock = LockService.getUserLock();
  lock.waitLock(30000);
  try {
    var results = operations.slice(0, 20).map(function (operation) {
      var operationId = String(operation && operation.operationId || '').trim();
      if (!operationId || operation.type !== 'create_drawing' || !Array.isArray(operation.rowData)) {
        return { success: false, operationId: operationId, error: 'Thao tác tạo bản vẽ không hợp lệ.' };
      }

      try {
        var result = saveDataToTestSheet(
          operation.subject || '',
          [operation.rowData.slice()],
          null,
          operationId
        );
        return {
          success: !!(result && result.success),
          operationId: operationId,
          message: result && result.message,
          error: result && result.error
        };
      } catch (err) {
        return { success: false, operationId: operationId, error: err.toString() };
      }
    });
    return { success: true, results: results };
  } finally {
    lock.releaseLock();
  }
}

// ========================================================================
// HỆ THỐNG MÃ KHÁCH HÀNG & QUẢN LÍ BẢN VẼ - CUSTOMER MAP & MANAGED DRAWINGS
// ========================================================================

function getCustomerMap() {
  try {
    var sheetId = '1DRteBSFT1cj4R_OUPMoDxeLMzAIJexWF3HPT-rpMOoM';
    var ss = SpreadsheetApp.openById(sheetId);
    var sheet = ss.getSheetByName('Customer');
    if (!sheet) return { success: false, error: 'Sheet Customer không tồn tại' };

    var data = sheet.getDataRange().getValues();
    // Cột A (index 0): Mã khách hàng (ví dụ BACT)
    // Cột B (index 1): Tên khách hàng (ví dụ BACH TUNG)
    var map = {};
    for (var i = 1; i < data.length; i++) {
      var code = String(data[i][0] || '').trim().toUpperCase();
      var name = String(data[i][1] || '').trim().toUpperCase();
      if (name && code) {
        map[name] = code;
      }
    }
    return { success: true, data: map };
  } catch (err) {
    return { success: false, error: err.toString() };
  }
}

function getManagedDrawings() {
  try {
    var sheetId = '1DRteBSFT1cj4R_OUPMoDxeLMzAIJexWF3HPT-rpMOoM';
    var ss = SpreadsheetApp.openById(sheetId);
    var dataSheet = ss.getSheetByName('Data');

    if (!dataSheet || dataSheet.getLastRow() <= 1) {
      return { success: true, data: [] };
    }

    var data = dataSheet.getRange(2, 1, dataSheet.getLastRow() - 1, Math.max(34, dataSheet.getLastColumn())).getValues();
    var displayData = dataSheet.getRange(2, 1, dataSheet.getLastRow() - 1, Math.max(34, dataSheet.getLastColumn())).getDisplayValues();
    var currentYear = Number(Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy'));
    var managed = [];
    var currentYearRows = [];
    var recentRows = [];

    for (var i = 0; i < data.length; i++) {
      var row = data[i] || [];
      var displayRow = displayData[i] || [];
      var receivedDate = row[6];
      var receivedYear = receivedDate instanceof Date && !isNaN(receivedDate.getTime())
        ? Number(Utilities.formatDate(receivedDate, Session.getScriptTimeZone(), 'yyyy'))
        : Number(String(receivedDate || '').match(/20\d{2}/) || 0);

      var rowContext = { row: row, displayRow: displayRow, index: i };
      if (receivedYear === currentYear) {
        currentYearRows.push(rowContext);
      }
      recentRows.push(rowContext);
    }

    var rowsToProcess = currentYearRows.length > 0 ? currentYearRows : recentRows.slice(-250);

    for (var j = 0; j < rowsToProcess.length; j++) {
      var rowInfo = rowsToProcess[j];
      var row = rowInfo.row || [];
      var displayRow = rowInfo.displayRow || [];

      // Dữ liệu cơ bản quyết định việc hiển thị; lỗi đọc hình ảnh không được loại bỏ dòng.
      var hasDrawingData = [row[8], row[9], row[10], row[11], row[33]].some(function (value) {
        return value !== null && value !== undefined && String(value).trim() !== '';
      });

      var acValue = row[28];  // AC: Width
      var adValue = row[29];  // AD: Height
      var agValue = row[32];  // AG: Hình mặt cắt

      var acStr = acValue === null || acValue === undefined ? '' : String(acValue).trim();
      var adStr = adValue === null || adValue === undefined ? '' : String(adValue).trim();
      var agStr = '';
      try {
        agStr = normalizeImageSource(agValue);
      } catch (imageError) {
        Logger.log('Không thể đọc hình mặt cắt dòng ' + (rowInfo.index + 2) + ': ' + imageError.toString());
      }
      var hasImageValue = agValue !== null && agValue !== undefined && agValue !== '';
      var displayAcStr = String(displayRow[28] || '').trim();
      var displayAdStr = String(displayRow[29] || '').trim();

      if (hasDrawingData || (acStr !== '') || (adStr !== '') || (displayAcStr !== '') || (displayAdStr !== '') || (agStr !== '') || hasImageValue) {
        var typeVal = row[3] || 'New';
        var toVal = row[8] || 'N/A';
        var dwCodeVal = row[11] || '';
        var customerVal = row[10] || 'Unknown';
        var oldDwCodeVal = row[14] || '';

        managed.push({
          status: String(row[2] || ''),
          type: String(typeVal),
          to: String(toVal),
          dwCode: String(dwCodeVal),
          oldDwCode: String(oldDwCodeVal),
          customer: String(customerVal),
          group: String(row[3] || ''),
          project: String(row[9] || ''),
          version: row[5] === null || row[5] === undefined ? '' : String(row[5]),
          receivedDate: String(row[6] || ''),
          assigneeDoneDate: String(row[7] || ''),
          typeDw: String(row[12] || row[24] || ''),
          assignee: String(row[15] || ''),
          actualDoneDate: String(row[18] || ''),
          note: String(row[22] || ''),
          fye: String(row[23] || ''),
          width: String(row[28] || ''),
          height: String(row[29] || ''),
          cutDrawing: agStr,
          image: agStr,
          so: String(row[33] || ''),
          rowIdx: rowInfo.index + 2
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

    // Kiểm tra trạng thái: Chỉ cho phép xóa khi "Đang thực hiện" hoặc "Trả về charger"
    var statusVal = String(dataSheet.getRange(rowId, 3).getValue() || '');
    var cleanSt = statusVal.toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g, '').replace(/đ/g, 'd').trim();
    var isEditable = cleanSt.includes('dang thuc hien') || cleanSt.includes('tra ve') || cleanSt.includes('tra lai') || cleanSt.includes('tu choi');
    if (!isEditable && cleanSt !== '') {
      return { success: false, error: 'Chỉ có bản vẽ ở trạng thái "Đang thực hiện" hoặc "Trả về charger" mới được xóa!' };
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
        status: String(rowData[2] || ''),            // C: Status
        group: String(rowData[3] || ''),             // D: Group
        type: String(rowData[4] || ''),              // E: Type
        version: rowData[5] === null || rowData[5] === undefined ? '' : String(rowData[5]), // F: Version / Revise
        receivedDate: String(rowData[6] || ''),     // G: Ngày tiếp nhận
        assigneeDoneDate: String(rowData[7] || ''), // H: Ngày HT dự kiến
        to: String(rowData[8] || ''),                // I: TO
        project: String(rowData[9] || ''),           // J: Dự án
        customer: String(rowData[10] || ''),         // K: Customer
        dwCode: String(rowData[11] || ''),           // L: Drawing Code
        oldDwCode: String(rowData[14] || ''),        // O: DW Code cũ trước khi thay đổi
        typeDw: String(rowData[12] || rowData[24] || ''), // M / Y: Type DW
        assignee: String(rowData[15] || ''),         // P: Người đảm trách
        actualDoneDate: String(rowData[18] || ''),   // S: Ngày hoàn thành thực tế
        note: String(rowData[22] || ''),             // W: Note / Change log
        fye: String(rowData[23] || ''),              // X: FYE
        width: String(rowData[28] || ''),            // AC: Width
        height: String(rowData[29] || ''),           // AD: Height
        image: normalizeImageSource(rowData[32]),     // AG: ảnh mặt cắt
        so: String(rowData[33] || '')                // AH: Mã SO
      }
    };
  } catch (err) {
    return { success: false, error: err.toString() };
  }
}

function getDrawingsWithoutImage(userEmailFilter) {
  try {
    var sheetId = '1DRteBSFT1cj4R_OUPMoDxeLMzAIJexWF3HPT-rpMOoM';
    var ss = SpreadsheetApp.openById(sheetId);
    var logSheet = ss.getSheetByName('Log');
    var dataSheet = ss.getSheetByName('Data');

    var filterEmail = String(userEmailFilter || '').trim().toLowerCase();

    // Cột E của Log là danh sách TO gốc. Chỉ ẩn TO sau khi đã có
    // trong Data; cột I/J chỉ dùng để phân loại NEW/OLD.
    var receivedToSet = new Set();
    if (dataSheet && dataSheet.getLastRow() > 1) {
      var dataRows = dataSheet.getRange(2, 1, dataSheet.getLastRow() - 1, Math.max(34, dataSheet.getLastColumn())).getValues();
      dataRows.forEach(function (row) {
        getLogToValues(row[8], false).forEach(function (toValue) {
          receivedToSet.add(normalizeLogTo(toValue));
        });
      });
    }

    // Nếu không có filter từ frontend thì tự lấy từ Session
    if (!filterEmail) {
      try { filterEmail = String(getCurrentUserEmail() || '').trim().toLowerCase(); } catch(e) {}
    }

    if (logSheet && logSheet.getLastRow() > 1) {
      var logValues = logSheet.getDataRange().getDisplayValues();
      var pending = [];

      for (var i = 1; i < logValues.length; i++) {
        var rawSo = String(logValues[i][5] || '').trim();
        var customer = String(logValues[i][3] || 'Unknown');
        var statusStr = String(logValues[i][6] || '').trim();
        var receiverCol = String(logValues[i][10] || '').trim();

        // Xác định email người tiếp nhận
        var rowReceiver = receiverCol ? receiverCol.toLowerCase() : '';
        if (!rowReceiver) {
          var match = statusStr.match(/-\s*([a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,})/i);
          if (match) rowReceiver = match[1].toLowerCase();
        }

        // Bỏ lọc theo người tiếp nhận: hiển thị tất cả bản vẽ đã tiếp nhận trong Log chưa vào Data

        // Nếu cột F (index 5) trống, trích xuất SO từ cột H (JSON fullDataString) hoặc subject
        if (!rawSo && logValues[i][7]) {
          try {
            var jsonStr = String(logValues[i][7]);
            if (jsonStr.startsWith('[')) {
              var pRows = JSON.parse(jsonStr);
              for (var pr = 0; pr < pRows.length; pr++) {
                if (!pRows[pr]) continue;
                for (var pc = 0; pc < pRows[pr].length; pc++) {
                  if (String(pRows[pr][pc] || '').trim().toLowerCase() === 'sales document' && pRows[pr + 1]) {
                    rawSo = 'SO' + String(pRows[pr + 1][pc] || '').replace(/[^0-9]/g, '');
                    break;
                  }
                }
                if (rawSo) break;
              }
            }
          } catch(e) {}
        }
        if (!rawSo && logValues[i][2]) {
          var sMatch = String(logValues[i][2]).match(/\bSO\s*[-:_]?\s*(\d+)\b/i);
          if (sMatch) rawSo = 'SO' + sMatch[1];
        }

        var logTos = splitDrawingToValues(logValues[i][4]).map(function (toValue) {
          var normalizedTo = normalizeLogTo(toValue);
          return normalizedTo ? 'TO-' + normalizedTo : '';
        }).filter(Boolean);
        logTos.forEach(function (toValue, toIndex) {
          if (receivedToSet.has(normalizeLogTo(toValue))) return;

          pending.push({
            rowIdx: i + 1,
            displayRowIdx: i + 1 + '-' + toIndex,
            type: 'New',
            to: toValue || 'N/A',
            dwCode: '',
            customer: customer,
            subject: String(logValues[i][2] || ''),
            threadId: String(logValues[i][1] || '').trim(),
            so: rawSo,
            receiverEmail: rowReceiver,
            existsInData: false,
            imageUrl: ''
          });
        });
      }

      if (pending.length === 0 && logValues.length > 1) {
        var fallbackLogRows = logValues.slice(1).slice(-80);
        for (var j = 0; j < fallbackLogRows.length; j++) {
          var logRow = fallbackLogRows[j];
          var fallbackRawSo = String(logRow[5] || '').trim();
          var fallbackCustomer = String(logRow[3] || 'Unknown');
          var fallbackTo = String(logRow[4] || '').trim();
          var fallbackStatus = String(logRow[6] || '').trim();
          var fallbackReceiverCol = String(logRow[10] || '').trim();
          var fallbackReceiver = fallbackReceiverCol.toLowerCase();
          if (!fallbackReceiver) {
            var fbMatch = fallbackStatus.match(/-\s*([a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,})/i);
            if (fbMatch) fallbackReceiver = fbMatch[1].toLowerCase();
          }

          if (!fallbackRawSo && logRow[7]) {
            try {
              var fJson = String(logRow[7]);
              if (fJson.startsWith('[')) {
                var fpRows = JSON.parse(fJson);
                for (var fpr = 0; fpr < fpRows.length; fpr++) {
                  if (!fpRows[fpr]) continue;
                  for (var fpc = 0; fpc < fpRows[fpr].length; fpc++) {
                    if (String(fpRows[fpr][fpc] || '').trim().toLowerCase() === 'sales document' && fpRows[fpr + 1]) {
                      fallbackRawSo = 'SO' + String(fpRows[fpr + 1][fpc] || '').replace(/[^0-9]/g, '');
                      break;
                    }
                  }
                  if (fallbackRawSo) break;
                }
              }
            } catch(e) {}
          }
          if (!fallbackRawSo && logRow[2]) {
            var fbSMatch = String(logRow[2]).match(/\bSO\s*[-:_]?\s*(\d+)\b/i);
            if (fbSMatch) fallbackRawSo = 'SO' + fbSMatch[1];
          }

          if (!fallbackTo) continue;
          pending.push({
            rowIdx: j + 2,
            displayRowIdx: j + 2 + '-fallback',
            type: 'New',
            to: fallbackTo || 'N/A',
            dwCode: '',
            customer: fallbackCustomer,
            subject: String(logRow[2] || ''),
            so: fallbackRawSo,
            receiverEmail: fallbackReceiver,
            existsInData: false,
            imageUrl: ''
          });
        }
      }

      return { success: true, data: pending.reverse() };
    }

    // Log không có TO NEW thì danh sách rỗng; tuyệt đối không lấy fallback từ Data.
    return { success: true, data: [] };
  } catch (err) {
    return { success: false, error: err.toString() };
  }
}

function splitDrawingToValues(rawValue) {
  var text = String(rawValue || '').trim();
  if (!text) return [];

  return text.split(/[;,\n|]+/).map(function (value) {
    return String(value).trim();
  }).filter(Boolean);
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

function updateDrawingData(rowIdx, formRow) {
  try {
    var ss = SpreadsheetApp.openById('1DRteBSFT1cj4R_OUPMoDxeLMzAIJexWF3HPT-rpMOoM');
    var sheet = ss.getSheetByName('Data');
    if (!sheet || !rowIdx || rowIdx < 2 || !formRow) return { success: false, error: 'Dữ liệu cập nhật không hợp lệ' };

    var row = sheet.getRange(rowIdx, 1, 1, sheet.getLastColumn()).getValues()[0];

    // Kiểm tra trạng thái: Chỉ cho phép sửa khi "Đang thực hiện" hoặc "Trả về charger"
    var currentStatus = String(row[2] || '');
    var cleanSt = currentStatus.toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g, '').replace(/đ/g, 'd').trim();
    var isEditable = cleanSt.includes('dang thuc hien') || cleanSt.includes('tra ve') || cleanSt.includes('tra lai') || cleanSt.includes('tu choi');
    if (!isEditable && cleanSt !== '') {
      return { success: false, error: 'Chỉ có bản vẽ ở trạng thái "Đang thực hiện" hoặc "Trả về charger" mới được sửa dữ liệu!' };
    }

    var currentDwInSheet = String(row[11] || '');
    var newDwCode = formRow[6] ? String(formRow[6]).trim() : '';
    if (currentDwInSheet && newDwCode && currentDwInSheet !== newDwCode) {
      row[14] = currentDwInSheet; // Cột O: lưu lại DW Code trước khi thay đổi
    }

    var fields = [
      [3, formRow[0]], [4, formRow[1]], [10, formRow[2]], [8, formRow[3]],
      [9, formRow[4]], [33, formRow[5]], [11, formRow[6]], [5, formRow[7]],
      [12, formRow[8]], [15, formRow[9]], [28, formRow[10]], [29, formRow[11]],
      [6, formRow[12]], [7, formRow[13]], [18, formRow[14]], [23, formRow[15]], [22, formRow[16]]
    ];
    fields.forEach(function (field) { row[field[0]] = field[1] === undefined ? '' : field[1]; });

    var imageObj = formRow[17];
    var cellImageObj = null;
    if (imageObj && imageObj.base64) {
      try {
        var imageMimeType = imageObj.mimeType || 'image/png';
        var dataUri = 'data:' + imageMimeType + ';base64,' + imageObj.base64;
        cellImageObj = SpreadsheetApp.newCellImage()
          .setSourceUrl(dataUri)
          .setAltTextTitle(imageObj.name || 'Hinh mat cat')
          .build();
        row[32] = cellImageObj;
      } catch (e) {
        Logger.log('Lỗi tạo CellImage: ' + e.toString());
      }
    }

    sheet.getRange(rowIdx, 1, 1, row.length).setValues([row]);
    if (cellImageObj) {
      try { sheet.getRange(rowIdx, 33).setValue(cellImageObj); } catch (e) { }
    }
    return { success: true };
  } catch (err) {
    return { success: false, error: err.toString() };
  }
}

function saveBulkDrawingData(subject, formRow, selectedItems) {
  try {
    if (!formRow || !Array.isArray(selectedItems) || selectedItems.length === 0) {
      return { success: false, error: 'Chưa có bản vẽ được chọn để sửa hàng loạt.' };
    }

    selectedItems = selectedItems.filter(function (item) {
      return item && item.rowIdx !== undefined && item.rowIdx !== null;
    });
    if (selectedItems.length === 0) {
      return { success: false, error: 'Danh sách bản vẽ được chọn không hợp lệ.' };
    }

    var results = [];
    selectedItems.forEach(function (item) {
      var row = formRow.slice();
      row[2] = item.customer || row[2] || '';
      row[3] = item.to || row[3] || '';
      row[5] = item.so || row[5] || '';
      results.push(saveDataToTestSheet(item.subject || subject, [row], null));
    });

    var failed = results.filter(function (result) { return !result || !result.success; });
    return failed.length
      ? { success: false, error: 'Có ' + failed.length + '/' + results.length + ' bản vẽ lưu không thành công.', results: results }
      : { success: true, count: results.length, results: results };
  } catch (err) {
    return { success: false, error: err.toString() };
  }
}

// ========================================================================
// TẦN SUẤT TO - FREQUENCY STATISTICS
// ========================================================================

function getToFrequencyData() {
  try {
    var sheetId = '1DRteBSFT1cj4R_OUPMoDxeLMzAIJexWF3HPT-rpMOoM';
    var ss = SpreadsheetApp.openById(sheetId);
    var soSheet = ss.getSheetByName('SO');

    if (!soSheet || soSheet.getLastRow() <= 1) {
      return { success: true, data: [], summary: { total: 0, highestTO: '', highestFrequency: 0 } };
    }

    var soData = soSheet.getRange(2, 1, soSheet.getLastRow() - 1, soSheet.getLastColumn()).getValues();
    var frequencyMap = {}; // {TO: {TO, colorCount, customerCount, maxFrequency}}

    // Duyệt qua từng hàng trong sheet SO
    soData.forEach(function (row) {
      var to = String(row[0] || '').trim().toUpperCase();     // Column A: TO
      var colorCode = String(row[2] || '').trim().toUpperCase(); // Column C: MÀU
      var customer = String(row[3] || '').trim().toUpperCase(); // Column D: Khách hàng
      var tanXuatTO = parseInt(row[6] || 0) || 0;             // Column G: Tần xuất TO
      var tanXuatMau = parseInt(row[7] || 0) || 0;            // Column H: Tần xuất MÀU
      var tanXuatKhach = parseInt(row[9] || 0) || 0;          // Column J: Tần xuất khách

      if (!to) return;

      if (!frequencyMap[to]) {
        frequencyMap[to] = {
          to: to,
          colors: {},
          customers: {},
          maxFrequency: 0,
          totalFrequency: 0,
          colorCount: 0,
          customerCount: 0
        };
      }

      // Đếm số lần xuất hiện của TO
      frequencyMap[to].maxFrequency = Math.max(frequencyMap[to].maxFrequency, tanXuatTO);
      frequencyMap[to].totalFrequency = (frequencyMap[to].totalFrequency || 0) + 1;

      // Đếm số màu khác nhau
      if (colorCode && !frequencyMap[to].colors[colorCode]) {
        frequencyMap[to].colors[colorCode] = tanXuatMau;
        frequencyMap[to].colorCount++;
      } else if (colorCode) {
        frequencyMap[to].colors[colorCode] = Math.max(frequencyMap[to].colors[colorCode], tanXuatMau);
      }

      // Đếm số khách hàng khác nhau
      if (customer && !frequencyMap[to].customers[customer]) {
        frequencyMap[to].customers[customer] = tanXuatKhach;
        frequencyMap[to].customerCount++;
      } else if (customer) {
        frequencyMap[to].customers[customer] = Math.max(frequencyMap[to].customers[customer], tanXuatKhach);
      }
    });

    // Chuyển đổi map thành array và sắp xếp
    var data = Object.keys(frequencyMap).map(function (toKey) {
      var toObj = frequencyMap[toKey];
      return {
        to: toObj.to,
        frequency: toObj.maxFrequency,
        totalOccurrences: toObj.totalFrequency,
        colorCount: toObj.colorCount,
        customerCount: toObj.customerCount,
        colors: Object.keys(toObj.colors).join(', '),
        customers: Object.keys(toObj.customers).join(', ')
      };
    }).sort(function (a, b) {
      return b.frequency - a.frequency; // Sắp xếp giảm dần theo tần suất
    });

    var summary = {
      total: data.length,
      highestTO: data.length > 0 ? data[0].to : '',
      highestFrequency: data.length > 0 ? data[0].frequency : 0,
      totalFrequencyCount: data.reduce(function (sum, item) { return sum + item.frequency; }, 0),
      averageFrequency: data.length > 0 ? (data.reduce(function (sum, item) { return sum + item.frequency; }, 0) / data.length).toFixed(2) : 0
    };

    return { success: true, data: data, summary: summary };
  } catch (err) {
    return { success: false, error: err.toString() };
  }
}

// ========================================================================
// LỊCH SỬ CHUYỂN TAB - TAB TRANSFER HISTORY
// ========================================================================

function getTabTransferHistory() {
  try {
    var sheetId = '1DRteBSFT1cj4R_OUPMoDxeLMzAIJexWF3HPT-rpMOoM';
    var ss = SpreadsheetApp.openById(sheetId);
    var dataSheet = ss.getSheetByName('Data');

    if (!dataSheet || dataSheet.getLastRow() <= 1) {
      return { success: true, data: [], summary: { totalTransfers: 0, uniqueDrawings: 0 } };
    }

    var data = dataSheet.getRange(2, 1, dataSheet.getLastRow() - 1, dataSheet.getLastColumn()).getValues();
    var transferHistory = [];
    var transferMap = {}; // Để tránh duplicate

    data.forEach(function (row, idx) {
      var id = String(row[0] || '').trim();                  // Column A: ID
      var to = String(row[8] || '').trim();                  // Column I: TO
      var dwCode = String(row[11] || '').trim();             // Column L: Drawing Code
      var customer = String(row[10] || '').trim();            // Column K: Customer
      var status = String(row[2] || '').trim();              // Column C: Status
      var receivedDate = String(row[6] || '').trim();        // Column G: Ngày tiếp nhận
      var assigneeDoneDate = String(row[7] || '').trim();    // Column H: Ngày hoàn thành dự kiến
      var actualDoneDate = String(row[18] || '').trim();     // Column S: Ngày HT thực tế
      var assignee = String(row[15] || '').trim();           // Column P: Người đảm trách
      var checker = String(row[17] || '').trim();            // Column R: Checker
      var approver = String(row[19] || '').trim();           // Column T: Approver
      var releaseStatus = String(row[21] || '').trim();      // Column V: Release Status
      var oldDwCode = String(row[14] || '').trim();          // Column O: Old DW Code
      var so = String(row[33] || '').trim();                 // Column AH: SO

      if (!to || !dwCode) return;

      var key = to + '|' + dwCode + '|' + customer;
      if (transferMap[key]) return; // Bỏ qua duplicate

      transferMap[key] = true;
      transferHistory.push({
        id: id,
        to: to,
        customer: customer,
        dwCode: dwCode,
        oldDwCode: oldDwCode,
        status: status,
        receivedDate: receivedDate,
        assigneeDoneDate: assigneeDoneDate,
        actualDoneDate: actualDoneDate,
        assignee: assignee,
        checker: checker,
        approver: approver,
        releaseStatus: releaseStatus,
        so: so,
        rowIdx: idx + 2,
        isUpgraded: oldDwCode ? true : false
      });
    });

    // Sắp xếp theo thứ tự mới nhất trước
    transferHistory.sort(function (a, b) {
      return (String(b.id || '')).localeCompare(String(a.id || ''), undefined, { numeric: true });
    });

    var summary = {
      totalTransfers: transferHistory.length,
      uniqueDrawings: Object.keys(transferMap).length,
      upgradedDrawings: transferHistory.filter(function (t) { return t.isUpgraded; }).length,
      statusDistribution: {}
    };

    // Đếm phân bố trạng thái
    transferHistory.forEach(function (t) {
      summary.statusDistribution[t.status] = (summary.statusDistribution[t.status] || 0) + 1;
    });

    return { success: true, data: transferHistory, summary: summary };
  } catch (err) {
    return { success: false, error: err.toString() };
  }
}

