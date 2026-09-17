const sheetConfig = {
  mainSpreadsheetId: "1DRteBSFT1cj4R_OUPMoDxeLMzAIJexWF3HPT-rpMOoM",
  releaseSpreadsheetId: "1t5PWyoJHrxElWP3QgmB16BEMIvEC015NHq0tsxu_TpE",
  mainSheetName: "Data",
  releaseSheetName: "data",
  holidaySheetName: "Cal",
};

const padZero = (n) => (n < 10 ? '0' + n : '' + n);

const formatSheetCellValue = (cell) => {
  if (cell instanceof Date) {
    if (isNaN(cell.getTime())) return '';
    return `${padZero(cell.getDate())}/${padZero(cell.getMonth() + 1)}/${cell.getFullYear()}`;
  }

  if (cell && typeof cell === "object") {
    try {
      const imageUrl = typeof cell.getContentUrl === "function" ? cell.getContentUrl() : "";
      if (imageUrl) {
        const response = UrlFetchApp.fetch(imageUrl, {
          headers: { Authorization: `Bearer ${ScriptApp.getOAuthToken()}` },
          muteHttpExceptions: true,
        });
        if (response.getResponseCode() === 200) {
          const blob = response.getBlob();
          return `data:${blob.getContentType()};base64,${Utilities.base64Encode(blob.getBytes())}`;
        }
      }
    } catch (error) {
      Logger.log(`Lỗi đọc CellImage: ${error.toString()}`);
    }
    return "";
  }

  return cell;
};

// Chuyển đổi theo lô để tránh UrlFetchApp.fetch tuần tự cho từng ảnh trong sheet.
// Giá trị trả về vẫn giữ nguyên như normalizeSheetRows trước đây.
const normalizeSheetRows = (rows, options) => {
  const includeImages = !(options && options.includeImages === false);
  const imageCells = [];
  const normalized = rows.map((row, rowIndex) => row.map((cell, columnIndex) => {
    if (cell && typeof cell === 'object' && typeof cell.getContentUrl === 'function') {
      if (!includeImages) return '';
      try {
        const imageUrl = cell.getContentUrl();
        if (imageUrl) {
          imageCells.push({ rowIndex, columnIndex, imageUrl });
          return null;
        }
      } catch (error) {
        Logger.log(`Lỗi đọc URL CellImage: ${error.toString()}`);
      }
    }
    return formatSheetCellValue(cell);
  }));

  if (!imageCells.length) return normalized;

  let responses = [];
  try {
    responses = UrlFetchApp.fetchAll(imageCells.map(({ imageUrl }) => ({
      url: imageUrl,
      headers: { Authorization: `Bearer ${ScriptApp.getOAuthToken()}` },
      muteHttpExceptions: true,
    })));
  } catch (error) {
    Logger.log(`Lỗi tải CellImage theo lô: ${error.toString()}`);
  }

  imageCells.forEach((item, index) => {
    const response = responses[index];
    if (response && response.getResponseCode() === 200) {
      const blob = response.getBlob();
      normalized[item.rowIndex][item.columnIndex] =
        `data:${blob.getContentType()};base64,${Utilities.base64Encode(blob.getBytes())}`;
    } else {
      normalized[item.rowIndex][item.columnIndex] = '';
    }
  });

  return normalized;
};

const getSheetByIdAndName = (spreadsheetId, sheetName) => {
  const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
  return spreadsheet.getSheetByName(sheetName);
};

const getSheetDataByConfig = (spreadsheetId, sheetName) => {
  try {
    const sheet = getSheetByIdAndName(spreadsheetId, sheetName);
    if (!sheet) {
      return [];
    }

    const lastRow = sheet.getLastRow();
    const lastColumn = sheet.getLastColumn();

    if (lastRow <= 1) {
      return [];
    }

    const rows = sheet.getRange(2, 1, lastRow - 1, lastColumn).getValues();
    return normalizeSheetRows(rows);
  } catch (error) {
    Logger.log(`Lỗi đọc dữ liệu sheet ${sheetName}: ${error.toString()}`);
    return [];
  }
};

const getAllData = () => getMainData();

const getOptimizedSheetRows = (spreadsheetId, sheetName, maxColumns, options) => {
  try {
    const sheet = getSheetByIdAndName(spreadsheetId, sheetName);
    if (!sheet) return [];

    const lastRow = sheet.getLastRow();
    const lastColumn = Math.min(sheet.getLastColumn(), maxColumns || sheet.getLastColumn());
    if (lastRow <= 1 || lastColumn <= 0) return [];

    const rows = sheet.getRange(2, 1, lastRow - 1, lastColumn).getValues();
    return normalizeSheetRows(rows, options);
  } catch (error) {
    Logger.log(`Lỗi đọc dữ liệu tối ưu hóa cho sheet ${sheetName}: ${error.toString()}`);
    return [];
  }
};

// Cột mốc năm của sheet Data (zero-based). Cột G (index 6) là thời gian
// tiếp nhận bản vẽ, nên đây là mốc duy nhất để xác định dữ liệu năm hiện tại.
// Không thay đổi row shape mà các màn hình hiện tại đang sử dụng.
const CURRENT_YEAR_COLUMNS = [6]; // G: thời gian tiếp nhận bản vẽ
const RELEASE_CURRENT_YEAR_COLUMN = 13; // N: ngày yêu cầu hoàn thành trong sheet release
const ACTIVE_STATUS_PARTS = [
  'dang thuc hien', 'cho checker', 'cho approval', 'cho ban hanh',
  'dang ban hanh', 'cho thiet ke', 'trinh ky', 'pending', 'tra ve',
  'tra lai', 'tu choi', 'yeu cau ban hanh', 'cho duyet', 'cho kiem tra',
  'cho charger', 'dang xu ly', 'dang thiet ke'
];

const getCurrentYearInScriptTimezone = () => Number(
  Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy')
);

const normalizeFilterText = (value) => String(value || '')
  .toLowerCase()
  .normalize('NFD')
  .replace(/[\u0300-\u036f]/g, '')
  .replace(/đ/g, 'd')
  .trim();

const parseSheetYear = (value) => {
  if (value instanceof Date && !isNaN(value.getTime())) {
    return Number(Utilities.formatDate(value, Session.getScriptTimeZone(), 'yyyy'));
  }

  const text = String(value || '').trim();
  if (!text) return 0;
  // Ưu tiên năm bốn chữ số trong FYE hoặc ngày hiển thị dd/MM/yyyy.
  const yearMatch = text.match(/(?:^|[^0-9])(20\d{2})(?:[^0-9]|$)/);
  if (yearMatch) return Number(yearMatch[1]);
  const parsed = new Date(text);
  return isNaN(parsed.getTime())
    ? 0
    : Number(Utilities.formatDate(parsed, Session.getScriptTimeZone(), 'yyyy'));
};

const isActiveOperationalRow = (row) => {
  const status = normalizeFilterText(row && row[2]);
  if (!status) return false;
  return ACTIVE_STATUS_PARTS.some((part) => status.indexOf(part) !== -1);
};

const isCurrentOperationalRow = (row, currentYear) => {
  if (!Array.isArray(row)) return false;
  // Phân loại năm phải bám đúng thời gian tiếp nhận ở cột G của sheet Data.
  // Không dùng trạng thái vận hành để kéo các dòng của năm cũ vào payload đầu tiên.
  return parseSheetYear(row[CURRENT_YEAR_COLUMNS[0]]) === currentYear;
};

const isCurrentReleaseRow = (row, currentYear) => {
  if (!Array.isArray(row)) return false;
  return parseSheetYear(row[RELEASE_CURRENT_YEAR_COLUMN]) === currentYear;
};

// Đọc giá trị gốc để lọc trước khi tải/giải mã CellImage. Lịch sử vẫn dùng
// đường đọc cũ khi được yêu cầu, còn request đầu tiên chỉ normalize các dòng cần thiết.
const getCurrentYearSheetRows = (spreadsheetId, sheetName, maxColumns, options) => {
  try {
    const sheet = getSheetByIdAndName(spreadsheetId, sheetName);
    if (!sheet) return [];
    const lastRow = sheet.getLastRow();
    const lastColumn = Math.min(sheet.getLastColumn(), maxColumns || sheet.getLastColumn());
    if (lastRow <= 1 || lastColumn <= 0) return [];

    const rows = sheet.getRange(2, 1, lastRow - 1, lastColumn).getValues();
    const currentYear = getCurrentYearInScriptTimezone();
    return normalizeSheetRows(rows.filter((row) => isCurrentOperationalRow(row, currentYear)), options);
  } catch (error) {
    Logger.log(`Lỗi đọc dữ liệu năm hiện tại cho sheet ${sheetName}: ${error.toString()}`);
    return [];
  }
};

const normalizeComparisonSo = (value) => {
  const text = String(value || '').trim().toUpperCase();
  const match = text.match(/(?:SO\s*)?(\d{6,})/);
  return match ? `SO${match[1]}` : '';
};

const getSoDrawingComparison = () => {
  const checkAdjustSpreadsheetId = '1sXSge8vRINt2x8PNjmxSUKMbX4nvVfBl1SBHmaOrUNI';
  const resultBySo = {};

  const addSource = (rawSo, source, details, flags) => {
    const so = normalizeComparisonSo(rawSo);
    if (!so) return;
    if (!resultBySo[so]) {
      resultBySo[so] = { so, email: false, drawing: false, isNew: false, isOld: false, details: [] };
    }
    resultBySo[so][source] = true;
    if (flags) {
      resultBySo[so].isNew = resultBySo[so].isNew || flags.isNew;
      resultBySo[so].isOld = resultBySo[so].isOld || flags.isOld;
    }
    if (details) resultBySo[so].details.push(details);
  };

  const checkAdjustSheet = getSheetByIdAndName(checkAdjustSpreadsheetId, 'Check adjust');
  if (checkAdjustSheet && checkAdjustSheet.getLastRow() >= 3) {
    const values = checkAdjustSheet.getRange(3, 2, checkAdjustSheet.getLastRow() - 2, Math.max(1, checkAdjustSheet.getLastColumn() - 1)).getDisplayValues();
    values.forEach((row) => {
      const newText = row.some((cell) => String(cell).trim().toUpperCase() === 'NEW');
      const oldText = row.some((cell) => String(cell).trim().toUpperCase() === 'OLD');
      addSource(row[0], 'checkAdjust', '', { isNew: newText, isOld: oldText });
    });
  }

  const mainSpreadsheet = SpreadsheetApp.openById(sheetConfig.mainSpreadsheetId);
  const logSheet = mainSpreadsheet.getSheetByName('Log');
  if (logSheet && logSheet.getLastRow() >= 2) {
    const logValues = logSheet.getRange(2, 1, logSheet.getLastRow() - 1, Math.max(7, logSheet.getLastColumn())).getDisplayValues();
    logValues.forEach((row) => addSource(row[5], 'email', row[2] ? String(row[2]) : ''));
  }

  const dataSheet = mainSpreadsheet.getSheetByName(sheetConfig.mainSheetName);
  if (dataSheet && dataSheet.getLastRow() >= 2) {
    const dataValues = dataSheet.getRange(2, 1, dataSheet.getLastRow() - 1, Math.max(34, dataSheet.getLastColumn())).getDisplayValues();
    dataValues.forEach((row) => addSource(row[33] || row[5], 'drawing', row[11] ? String(row[11]) : ''));
  }

  const rows = Object.keys(resultBySo).map((so) => resultBySo[so]);
  rows.sort((a, b) => a.so.localeCompare(b.so, undefined, { numeric: true }));
  return { success: true, data: rows };
};

const getCheckAdjustSoSet = () => {
  const checkAdjustSpreadsheetId = '1sXSge8vRINt2x8PNjmxSUKMbX4nvVfBl1SBHmaOrUNI';
  const checkAdjustSheet = getSheetByIdAndName(checkAdjustSpreadsheetId, 'Check adjust');
  const soSet = new Set();
  if (!checkAdjustSheet || checkAdjustSheet.getLastRow() < 3) return soSet;

  const values = checkAdjustSheet.getRange(3, 2, checkAdjustSheet.getLastRow() - 2, Math.max(1, checkAdjustSheet.getLastColumn() - 1)).getDisplayValues();
  values.forEach((row) => {
    const normalizedSo = normalizeComparisonSo(row[0]);
    if (normalizedSo) soSet.add(normalizedSo);
  });
  return soSet;
};

const getDrawingStatusBySo = (rawSo) => {
  const normalizedSo = normalizeComparisonSo(rawSo);
  const isDrawingDone = !!(normalizedSo && getCheckAdjustSoSet().has(normalizedSo));
  return {
    success: true,
    so: normalizedSo,
    isDrawingDone,
    drawingStatus: isDrawingDone ? 'Check SO OK' : 'Không có SO'
  };
};

const getMainData = () => getOptimizedSheetRows(sheetConfig.mainSpreadsheetId, sheetConfig.mainSheetName, 35);

const getReleaseData = (providedMainRows, options) => {
  const releaseSpreadsheet = SpreadsheetApp.openById(sheetConfig.releaseSpreadsheetId);
  const expectedSheetName = String(sheetConfig.releaseSheetName).trim().toLowerCase();
  const releaseSheet = releaseSpreadsheet.getSheets().find((candidate) => String(candidate.getName()).trim().toLowerCase() === expectedSheetName);
  if (!releaseSheet) {
    throw new Error(`Không tìm thấy sheet "${sheetConfig.releaseSheetName}" trong file dữ liệu ban hành.`);
  }

  const rawReleaseRows = releaseSheet.getLastRow() > 1
    ? releaseSheet.getRange(2, 1, releaseSheet.getLastRow() - 1, Math.min(releaseSheet.getLastColumn(), 35)).getValues()
    : [];
  const currentOnly = !!(options && options.currentOnly);
  const currentYear = getCurrentYearInScriptTimezone();
  const isQaG2gRow = (row) => String(row && row[10] || '')
    .trim()
    .toUpperCase()
    .replace(/[^A-Z0-9]/g, '') === 'QAG2G';
  const releaseRows = normalizeSheetRows(currentOnly
    ? rawReleaseRows.filter((row) => isCurrentReleaseRow(row, currentYear))
    : rawReleaseRows, options);

  // Một số yêu cầu QA-G2G đang nằm ở Data trước khi được ghi sang sheet release.
  // Dashboard đã đọc mainData ở ngay trước đó; dùng lại để tránh đọc Data lần thứ hai.
  const mainRows = Array.isArray(providedMainRows)
    ? providedMainRows
    : getOptimizedSheetRows(sheetConfig.mainSpreadsheetId, sheetConfig.mainSheetName, 35, options);
  const qaReleaseRows = releaseRows.filter(isQaG2gRow);
  const qaRows = mainRows.filter((row) => isQaG2gRow(row) &&
    (!currentOnly || isCurrentReleaseRow(row, currentYear)));
  const mergedRows = [];
  const seenKeys = new Set();

  qaReleaseRows.concat(qaRows).forEach((row) => {
    const key = String(row[0] || '').trim() || `${row[8] || ''}|${row[4] || ''}|${row[5] || ''}`;
    if (seenKeys.has(key)) return;
    seenKeys.add(key);
    mergedRows.push(row);
  });

  return mergedRows;
};

const getSheetLastUpdatedAt = (spreadsheetId, sheetName) => {
  try {
    const sheet = getSheetByIdAndName(spreadsheetId, sheetName);
    if (!sheet) return 0;
    const updatedAt = sheet.getLastUpdated();
    return updatedAt ? updatedAt.getTime() : 0;
  } catch (error) {
    return 0;
  }
};

const getDashboardDataVersion = () => {
  const mainSheetUpdated = getSheetLastUpdatedAt(sheetConfig.mainSpreadsheetId, sheetConfig.mainSheetName);
  const releaseSheetUpdated = getSheetLastUpdatedAt(sheetConfig.releaseSpreadsheetId, sheetConfig.releaseSheetName);
  return {
    success: true,
    currentYear: getCurrentYearInScriptTimezone(),
    mainDataVersion: mainSheetUpdated,
    releaseDataVersion: releaseSheetUpdated,
    serverTime: Date.now()
  };
};

const getDashboardPayload = () => {
  // Trả về toàn bộ dữ liệu hệ thống (tất cả các năm) không lọc giới hạn năm hiện tại
  const mainData = getOptimizedSheetRows(sheetConfig.mainSpreadsheetId, sheetConfig.mainSheetName, 35, { includeImages: false });
  return {
    mainData,
    releaseData: getReleaseData(mainData, { includeImages: false }),
    isCurrentYearOnly: false,
    currentYear: getCurrentYearInScriptTimezone()
  };
};

const getCurrentYearMainData = (options) => getOptimizedSheetRows(
  sheetConfig.mainSpreadsheetId,
  sheetConfig.mainSheetName,
  35,
  options
);

const getHistoricalDashboardPayload = () => getDashboardPayload();

const getHolidaysFromCal = () => {
  try {
    const sheet = getSheetByIdAndName(sheetConfig.mainSpreadsheetId, sheetConfig.holidaySheetName);
    if (!sheet) {
      return [];
    }

    const lastRow = sheet.getLastRow();
    if (lastRow < 2) {
      return [];
    }

    const holidays = sheet.getRange(2, 1, lastRow - 1, 1).getValues().flat();

    return holidays
      .filter((holiday) => holiday && holiday !== "")
      .map((holiday) => {
        if (holiday instanceof Date) {
          return new Date(holiday.getFullYear(), holiday.getMonth(), holiday.getDate());
        }

        const parsedHoliday = new Date(holiday);
        if (!Number.isNaN(parsedHoliday.getTime())) {
          return new Date(parsedHoliday.getFullYear(), parsedHoliday.getMonth(), parsedHoliday.getDate());
        }

        return null;
      })
      .filter((holiday) => holiday !== null);
  } catch (error) {
    Logger.log(`Lỗi lấy danh sách ngày nghỉ từ Cal: ${error.toString()}`);
    return [];
  }
};

const isHoliday = (date, holidays) => {
  if (!holidays || holidays.length === 0) {
    return false;
  }

  const targetDateString = Utilities.formatDate(date, Session.getScriptTimeZone(), "yyyy-MM-dd");

  return holidays.some((holiday) => {
    const holidayDateString = Utilities.formatDate(holiday, Session.getScriptTimeZone(), "yyyy-MM-dd");
    return targetDateString === holidayDateString;
  });
};

const getBusinessDateFromTodayFromCalendar = (workingDays) => {
  const targetWorkingDays = workingDays || 5;

  try {
    const holidays = getHolidaysFromCal();
    let currentDate = new Date();
    let count = 0;

    while (count < targetWorkingDays) {
      currentDate.setDate(currentDate.getDate() + 1);

      const dayOfWeek = currentDate.getDay();
      if (dayOfWeek === 0 || dayOfWeek === 6) {
        continue;
      }

      if (isHoliday(currentDate, holidays)) {
        continue;
      }

      count += 1;
    }

    const year = currentDate.getFullYear();
    const month = String(currentDate.getMonth() + 1).padStart(2, "0");
    const day = String(currentDate.getDate()).padStart(2, "0");

    return `${year}-${month}-${day}`;
  } catch (error) {
    Logger.log(`Lỗi tính toán ngày hoàn thành: ${error.toString()}`);

    let fallbackDate = new Date();
    let count = 0;

    while (count < targetWorkingDays) {
      fallbackDate.setDate(fallbackDate.getDate() + 1);
      const dayOfWeek = fallbackDate.getDay();
      if (dayOfWeek !== 0 && dayOfWeek !== 6) {
        count += 1;
      }
    }

    const year = fallbackDate.getFullYear();
    const month = String(fallbackDate.getMonth() + 1).padStart(2, "0");
    const day = String(fallbackDate.getDate()).padStart(2, "0");

    return `${year}-${month}-${day}`;
  }
};
