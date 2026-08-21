const sheetConfig = {
  mainSpreadsheetId: "1DRteBSFT1cj4R_OUPMoDxeLMzAIJexWF3HPT-rpMOoM",
  releaseSpreadsheetId: "1t5PWyoJHrxElWP3QgmB16BEMIvEC015NHq0tsxu_TpE",
  mainSheetName: "Data",
  releaseSheetName: "data",
  holidaySheetName: "Cal",
};

const formatSheetCellValue = (cell) => {
  if (cell && typeof cell.getContentUrl === "function") {
    return cell.getContentUrl() || "";
  }

  if (String(cell) === "[object CellImage]") {
    return "";
  }

  if (cell instanceof Date) {
    return Utilities.formatDate(cell, Session.getScriptTimeZone(), "dd/MM/yyyy");
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

const getMainData = () => getSheetDataByConfig(sheetConfig.mainSpreadsheetId, sheetConfig.mainSheetName);

const getReleaseData = () => {
  const spreadsheet = SpreadsheetApp.openById(sheetConfig.releaseSpreadsheetId);
  const expectedSheetName = String(sheetConfig.releaseSheetName).trim().toLowerCase();
  const sheet = spreadsheet.getSheets().find((candidate) => String(candidate.getName()).trim().toLowerCase() === expectedSheetName);
  if (!sheet) {
    throw new Error(`Không tìm thấy sheet "${sheetConfig.releaseSheetName}" trong file dữ liệu ban hành.`);
  }

  const lastRow = sheet.getLastRow();
  const lastColumn = sheet.getLastColumn();
  if (lastRow <= 1 || lastColumn === 0) {
    return [];
  }

  return normalizeSheetRows(sheet.getRange(2, 1, lastRow - 1, lastColumn).getValues());
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
