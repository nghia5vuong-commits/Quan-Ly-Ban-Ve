const drawingSheetConfig = {
  spreadsheetId: "1t5PWyoJHrxElWP3QgmB16BEMIvEC015NHq0tsxu_TpE",
  sheetName: "data",
  idColumnIndex: 1,
  requestNumberColumnIndex: 9,
  departmentStatusColumnIndex: 24,
  noteColumnIndex: 34,
};

const getDrawingSheet = () => SpreadsheetApp.openById(drawingSheetConfig.spreadsheetId).getSheetByName(drawingSheetConfig.sheetName);

const normalizeCellValue = (value) => String(value ?? "").trim();

const groupContiguousRowNumbers = (rowNumbers) => {
  if (!rowNumbers || rowNumbers.length === 0) {
    return [];
  }

  const sortedRowNumbers = [...new Set(rowNumbers)].sort((left, right) => left - right);
  const groupedRows = [];
  let startRow = sortedRowNumbers[0];
  let previousRow = sortedRowNumbers[0];

  for (let index = 1; index < sortedRowNumbers.length; index += 1) {
    const currentRow = sortedRowNumbers[index];

    if (currentRow === previousRow + 1) {
      previousRow = currentRow;
      continue;
    }

    groupedRows.push([startRow, previousRow - startRow + 1]);
    startRow = currentRow;
    previousRow = currentRow;
  }

  groupedRows.push([startRow, previousRow - startRow + 1]);
  return groupedRows;
};

const clearRowsByGroups = (sheet, rowNumbers, lastColumn) => {
  const groupedRows = groupContiguousRowNumbers(rowNumbers);

  groupedRows.forEach(([startRow, length]) => {
    sheet.getRange(startRow, 1, length, lastColumn).clearContent();
  });
};

const setStatusOnRowGroups = (sheet, rowNumbers, columnIndex, statusText) => {
  const groupedRows = groupContiguousRowNumbers(rowNumbers);

  groupedRows.forEach(([startRow, length]) => {
    const values = Array.from({ length }, () => [statusText]);
    sheet.getRange(startRow, columnIndex, length, 1).setValues(values);
  });
};

const getAllDataForStats = () => {
  try {
    const sheet = getDrawingSheet();
    if (!sheet) {
      return [];
    }

    const lastRow = sheet.getLastRow();
    const lastColumn = sheet.getLastColumn();

    if (lastRow <= 1) {
      return [];
    }

    const rows = sheet.getRange(2, 1, lastRow - 1, lastColumn).getValues();
    return rows.map((row) => row.map((cell) => (cell instanceof Date ? Utilities.formatDate(cell, Session.getScriptTimeZone(), "dd/MM/yyyy") : cell)));
  } catch (error) {
    Logger.log(`Lỗi đọc dữ liệu thống kê: ${error.toString()}`);
    return [];
  }
};

const delDwNo = (id) => {
  const sheet = getDrawingSheet();
  if (!sheet) {
    throw new Error("Không tìm thấy sheet dữ liệu release.");
  }

  const targetId = normalizeCellValue(id);
  const idColumnValues = sheet.getRange(1, 1, sheet.getLastRow(), 1).getValues();
  const rowIndex = idColumnValues.findIndex((row) => normalizeCellValue(row[0]) === targetId) + 1;

  if (rowIndex <= 0) {
    throw new Error("Không tìm thấy ID bản vẽ. Có thể dữ liệu đã được thay đổi bởi người khác.");
  }

  sheet.getRange(rowIndex, 1, 1, sheet.getLastColumn()).clearContent();
  return "Thành công";
};

const delOrderNo = (idList) => {
  const sheet = getDrawingSheet();
  if (!sheet) {
    return;
  }

  const lastRow = sheet.getLastRow();
  const lastColumn = sheet.getLastColumn();

  if (lastRow < 2) {
    return;
  }

  const targetIds = Array.isArray(idList) ? idList.map((item) => normalizeCellValue(item)) : [normalizeCellValue(idList)];
  const data = sheet.getRange(2, 1, lastRow - 1, lastColumn).getValues();
  const rowNumbers = [];

  data.forEach((row, index) => {
    const currentRequestNumber = normalizeCellValue(row[drawingSheetConfig.requestNumberColumnIndex - 1]);
    if (targetIds.includes(currentRequestNumber)) {
      rowNumbers.push(index + 2);
    }
  });

  clearRowsByGroups(sheet, rowNumbers, lastColumn);
};

const updateDrawingFull = (payload) => {
  try {
    if (!payload || !payload.id) {
      return "Dữ liệu không hợp lệ";
    }

    const sheet = getDrawingSheet();
    if (!sheet) {
      throw new Error("Không tìm thấy Sheet 'data'");
    }

    const lastRow = sheet.getLastRow();
    const lastColumn = sheet.getLastColumn();

    if (lastRow < 2) {
      return "Bảng dữ liệu trống";
    }

    const data = sheet.getRange(2, 1, lastRow - 1, lastColumn).getValues();
    const targetRowIndex = data.findIndex((row) => normalizeCellValue(row[0]) === normalizeCellValue(payload.id));

    if (targetRowIndex === -1) {
      throw new Error("Không tìm thấy ID bản vẽ để cập nhật!");
    }

    const updatedRow = [...data[targetRowIndex]];
    updatedRow[1] = payload.type;
    updatedRow[4] = payload.dwNo;
    updatedRow[5] = payload.ver;
    updatedRow[2] = payload.lineCus;
    updatedRow[3] = payload.lineCode;
    updatedRow[9] = payload.designNo;
    updatedRow[7] = payload.shape;
    updatedRow[11] = payload.deptLq;
    updatedRow[17] = payload.link;
    updatedRow[33] = payload.note;
    updatedRow[39] = payload.changeType;

    const rowNumber = targetRowIndex + 2;
    sheet.getRange(rowNumber, 1, 1, lastColumn).setValues([updatedRow]);
    return "Cập nhật thành công";
  } catch (error) {
    throw new Error(error.message || "Lỗi khi cập nhật dữ liệu bản vẽ");
  }
};

const sendToChargeQA = (idList) => {
  const sheet = getDrawingSheet();
  if (!sheet) {
    return [];
  }

  const lastRow = sheet.getLastRow();
  const lastColumn = sheet.getLastColumn();

  if (lastRow < 2) {
    return [];
  }

  const targetIds = Array.isArray(idList) ? idList.map((item) => normalizeCellValue(item)) : [normalizeCellValue(idList)];
  const data = sheet.getRange(2, 1, lastRow - 1, lastColumn).getValues();
  const rowsToUpdate = [];
  const approvedIds = [];

  data.forEach((row, index) => {
    const currentRequestNumber = normalizeCellValue(row[drawingSheetConfig.requestNumberColumnIndex - 1]);
    const hasNote = normalizeCellValue(row[drawingSheetConfig.noteColumnIndex - 1]) !== "";

    if (targetIds.includes(currentRequestNumber) && hasNote) {
      rowsToUpdate.push(index + 2);
      approvedIds.push(row[0]);
    }
  });

  setStatusOnRowGroups(sheet, rowsToUpdate, drawingSheetConfig.departmentStatusColumnIndex, "Dept đã chọn phương án");
  return approvedIds;
};
