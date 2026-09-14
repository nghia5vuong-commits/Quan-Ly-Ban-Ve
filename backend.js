const sheetConfig = {
  mainSpreadsheetId: "1DRteBSFT1cj4R_OUPMoDxeLMzAIJexWF3HPT-rpMOoM",
  releaseSpreadsheetId: "1t5PWyoJHrxElWP3QgmB16BEMIvEC015NHq0tsxu_TpE",
  mainSheetName: "Data",
  releaseSheetName: "data",
  holidaySheetName: "Cal",
};

const formatSheetCellValue = (cell) => {
  if (cell instanceof Date) {
    return Utilities.formatDate(cell, Session.getScriptTimeZone(), "dd/MM/yyyy");
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

const normalizeSheetRows = (rows) => rows.map((row) => row.map((cell) => formatSheetCellValue(cell)));

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

const getMainData = () => getSheetDataByConfig(sheetConfig.mainSpreadsheetId, sheetConfig.mainSheetName);

const getReleaseData = () => {
  const releaseSpreadsheet = SpreadsheetApp.openById(sheetConfig.releaseSpreadsheetId);
  const expectedSheetName = String(sheetConfig.releaseSheetName).trim().toLowerCase();
  const releaseSheet = releaseSpreadsheet.getSheets().find((candidate) => String(candidate.getName()).trim().toLowerCase() === expectedSheetName);
  if (!releaseSheet) {
    throw new Error(`Không tìm thấy sheet "${sheetConfig.releaseSheetName}" trong file dữ liệu ban hành.`);
  }

  const releaseRows = releaseSheet.getLastRow() > 1
    ? normalizeSheetRows(releaseSheet.getRange(2, 1, releaseSheet.getLastRow() - 1, releaseSheet.getLastColumn()).getValues())
    : [];

  // Một số yêu cầu QA-G2G đang nằm ở Data trước khi được ghi sang sheet release.
  // Hợp nhất chúng để dashboard không bị rỗng giữa hai bước của quy trình.
  const mainSheet = getSheetByIdAndName(sheetConfig.mainSpreadsheetId, sheetConfig.mainSheetName);
  const mainRows = mainSheet && mainSheet.getLastRow() > 1
    ? normalizeSheetRows(mainSheet.getRange(2, 1, mainSheet.getLastRow() - 1, mainSheet.getLastColumn()).getValues())
    : [];
  const isQaG2gRow = (row) => String(row[10] || '').trim().toUpperCase() === 'QA-G2G';
  const qaReleaseRows = releaseRows.filter(isQaG2gRow);
  const qaRows = mainRows.filter(isQaG2gRow);
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
