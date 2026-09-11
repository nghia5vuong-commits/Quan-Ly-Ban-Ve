const designSheetConfig = {
  spreadsheetId: "1DRteBSFT1cj4R_OUPMoDxeLMzAIJexWF3HPT-rpMOoM",
  sheetName: "Data",
  designFolderId: "1PGwlJ484P8xki8i1k0GiyJvtwdMHu8A0",
  approvalFolderId: "1lJEJnMhRW8C-ykqR-EHMMQWkcRtO2Abz",
  fileLinkStartColumn: 31,
  fileLinkColumnCount: 2,
};

const getDesignSheet = () => SpreadsheetApp.openById(designSheetConfig.spreadsheetId).getSheetByName(designSheetConfig.sheetName);

const normalizeText = (value) => String(value ?? "").trim();

const findRowIndexById = (sheet, dataId, dwCode, toCode) => {
  if (!sheet) return -1;

  const idVal = normalizeText(dataId);
  const dwVal = normalizeText(dwCode);
  const toVal = normalizeText(toCode);

  // 1. Tìm theo dataId ở Cột A (ID)
  if (idVal) {
    const foundCell = sheet.getRange("A:A").createTextFinder(idVal).matchEntireCell(true).findNext();
    if (foundCell && foundCell.getRow() > 1) return foundCell.getRow();
  }

  // 2. Tìm theo dwCode hoặc dataId ở Cột L (Drawing Code - Cột 12)
  const searchDw = dwVal || idVal;
  if (searchDw) {
    const foundCell = sheet.getRange("L:L").createTextFinder(searchDw).matchEntireCell(true).findNext();
    if (foundCell && foundCell.getRow() > 1) return foundCell.getRow();
  }

  // 3. Tìm theo toCode hoặc dataId ở Cột I (TO - Cột 9)
  const searchTo = toVal || idVal;
  if (searchTo) {
    const foundCell = sheet.getRange("I:I").createTextFinder(searchTo).matchEntireCell(true).findNext();
    if (foundCell && foundCell.getRow() > 1) return foundCell.getRow();
  }

  // 4. Quét mảng dữ liệu (fallback đề phòng định dạng ký tự, khoảng trắng thừa hoặc chữ hoa/thường)
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    const dataRange = sheet.getRange(2, 1, lastRow - 1, Math.min(sheet.getLastColumn(), 34)).getValues();
    for (let r = 0; r < dataRange.length; r++) {
      const row = dataRange[r];
      const rId = String(row[0] || '').trim().toLowerCase();
      const rTo = String(row[8] || '').trim().toLowerCase();
      const rDw = String(row[11] || '').trim().toLowerCase();

      if (idVal && rId && rId === idVal.toLowerCase()) return r + 2;
      if (searchDw && rDw && rDw === searchDw.toLowerCase()) return r + 2;
      if (searchTo && rTo && rTo === searchTo.toLowerCase()) return r + 2;
    }
  }

  return -1;
};

const uploadPdfDesignToDrive = (pdfBase64, excelBase64, filePdfName, fileExcelName, dataId, dwCode, toCode) => {
  try {
    const decodedPdfData = Utilities.base64Decode(pdfBase64);
    const pdfBlob = Utilities.newBlob(decodedPdfData, MimeType.PDF, filePdfName);

    const decodedExcelData = Utilities.base64Decode(excelBase64);
    const excelBlob = Utilities.newBlob(decodedExcelData, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", fileExcelName);

    const sheet = getDesignSheet();
    const folder = DriveApp.getFolderById(designSheetConfig.designFolderId);

    const pdfFile = folder.createFile(pdfBlob);
    const excelFile = folder.createFile(excelBlob);

    pdfFile.setSharing(DriveApp.Access.DOMAIN_WITH_LINK, DriveApp.Permission.VIEW);
    excelFile.setSharing(DriveApp.Access.DOMAIN_WITH_LINK, DriveApp.Permission.VIEW);

    const pdfUrl = pdfFile.getUrl();
    const excelUrl = excelFile.getUrl();

    let resolvedId = dataId;
    if (sheet) {
      const targetRowIndex = findRowIndexById(sheet, dataId, dwCode, toCode);
      if (targetRowIndex > 1) {
        // Ghi link Excel vào cột AE (31) và PDF vào cột AF (32)
        sheet.getRange(targetRowIndex, designSheetConfig.fileLinkStartColumn, 1, designSheetConfig.fileLinkColumnCount).setValues([[excelUrl, pdfUrl]]);

        // Tự động chuyển Trạng thái sang "Chờ Checker 1" (Cột C - cột 3)
        sheet.getRange(targetRowIndex, 3).setValue("Chờ Checker 1");

        // Ghi ngày người đảm trách hoàn thành thiết kế (Cột Q - cột 17) nếu chưa có
        try {
          const currentDoneDate = String(sheet.getRange(targetRowIndex, 17).getValue() || '').trim();
          if (!currentDoneDate) {
            sheet.getRange(targetRowIndex, 17).setValue(Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy"));
          }
        } catch (eDate) {}

        // Nếu Cột A (ID) đang trống, tự động cấp ID duy nhất để các bước sau (Ký duyệt, tra cứu) hoạt động chính xác
        const currentId = String(sheet.getRange(targetRowIndex, 1).getValue() || '').trim();
        if (!currentId) {
          const dateStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyyMMdd");
          const randomNum = Math.floor(1000 + Math.random() * 9000);
          resolvedId = "REQ-" + dateStr + "-" + randomNum;
          sheet.getRange(targetRowIndex, 1).setValue(resolvedId);
        } else {
          resolvedId = currentId;
        }
      } else {
        throw new Error(`Đã tạo file Drive nhưng không tìm thấy dòng bản vẽ '${dataId || dwCode || toCode}' trong Sheet Data để ghi nhận link.`);
      }
    }

    return {
      pdfUrl,
      excelUrl,
      id: resolvedId,
      status: "Chờ Checker 1"
    };
  } catch (error) {
    throw new Error(`Lỗi khi lưu file lên Drive: ${error.message}`);
  }
};

const submitForApproval = (dataId, base64Data, stt) => {
  try {
    if (!dataId || !base64Data) {
      throw new Error("Thiếu mã dữ liệu hoặc file PDF.");
    }

    const sheet = getDesignSheet();
    if (!sheet) {
      throw new Error("Không tìm thấy sheet dữ liệu thiết kế.");
    }

    const folder = DriveApp.getFolderById(designSheetConfig.approvalFolderId);
    const decodedData = Utilities.base64Decode(base64Data);
    const pdfBlob = Utilities.newBlob(decodedData, "application/pdf", `${dataId}.pdf`);
    const newFile = folder.createFile(pdfBlob);

    try {
      newFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    } catch (sharingError) {
      newFile.setSharing(DriveApp.Access.DOMAIN_WITH_LINK, DriveApp.Permission.VIEW);
    }

    const fileUrl = newFile.getUrl();
    const targetRowIndex = findRowIndexById(sheet, dataId, dataId, dataId);
    if (targetRowIndex <= 1) {
      throw new Error(`Không tìm thấy ID '${dataId}' trong sheet để ghi link.`);
    }

    const currentStatus = String(sheet.getRange(targetRowIndex, 3).getValue() || '').trim();
    if (stt) {
      sheet.getRange(targetRowIndex, 3).setValue(stt);
    } else if (!currentStatus || currentStatus.toLowerCase().includes('dang thuc hien') || currentStatus.toLowerCase().includes('tra ve') || currentStatus.toLowerCase().includes('cho thiet ke')) {
      sheet.getRange(targetRowIndex, 3).setValue("Chờ Checker 1");
    }

    sheet.getRange(targetRowIndex, 32).setValue(fileUrl);
    return fileUrl;
  } catch (error) {
    throw new Error(`Lỗi khi lưu trữ data: ${error.message}`);
  }
};

const getFilesFromDrive = (pdfLink, userMail) => {
  try {
    if (!pdfLink || !String(pdfLink).includes('http')) {
      throw new Error("Link PDF trên bảng dữ liệu đang rỗng hoặc không phải URL Google Drive hợp lệ.");
    }

    const url = String(pdfLink);
    let fileId = '';

    // Hỗ trợ cả: https://drive.google.com/open?id=... , /file/d/.../view, /uc?id=...
    const idMatch = url.match(/[?&]id=([a-zA-Z0-9_-]{10,})/i);
    const fileDMatch = url.match(/\/d\/([a-zA-Z0-9_-]{10,})/i);
    const fileIdLongMatch = url.match(/[-\w]{25,}/);

    if (idMatch && idMatch[1]) fileId = idMatch[1];
    else if (fileDMatch && fileDMatch[1]) fileId = fileDMatch[1];
    else if (fileIdLongMatch) fileId = fileIdLongMatch[0];

    if (!fileId) {
      throw new Error("Link Drive không hợp lệ hoặc không chứa ID file PDF.");
    }

    const pdfBlob = DriveApp.getFileById(fileId).getBlob();
    const stampBlob = getSignBlob(userMail);

    if (!stampBlob) {
      throw new Error(`Không thể trích xuất chữ ký của tài khoản: ${userMail}. Vui lòng kiểm tra lại Cột F (Chữ ký) trong Sheet 'User'!`);
    }

    return {
      success: true,
      pdfBase64: Utilities.base64Encode(pdfBlob.getBytes()),
      stampBase64: Utilities.base64Encode(stampBlob.getBytes()),
    };
  } catch (error) {
    return {
      success: false,
      message: error && error.toString ? error.toString() : String(error),
    };
  }
};

const getSignBlob = (userMail) => {
  try {
    if (!userMail) return null;
    const cleanMail = String(userMail).trim().toLowerCase();

    // 1. Thử tìm sheet User từ 2 Spreadsheet cấu hình trong hệ thống
    const ssIds = [
      designSheetConfig.spreadsheetId,
      "1t5PWyoJHrxElWP3QgmB16BEMIvEC015NHq0tsxu_TpE"
    ];

    let targetSheet = null;
    let rowIndex = -1;

    for (const ssId of ssIds) {
      try {
        const ss = SpreadsheetApp.openById(ssId);
        const s = ss ? ss.getSheetByName("User") : null;
        if (!s || s.getLastRow() < 2) continue;

        // Quét cột D (Mail)
        const mails = s.getRange(2, 4, s.getLastRow() - 1, 1).getValues();
        for (let i = 0; i < mails.length; i++) {
          if (String(mails[i][0] || '').trim().toLowerCase() === cleanMail) {
            rowIndex = i + 2;
            targetSheet = s;
            break;
          }
        }
        if (rowIndex > 1) break;
      } catch (eSs) {}
    }

    if (!targetSheet || rowIndex <= 1) {
      Logger.log(`Không tìm thấy email ${userMail} trong sheet User.`);
      return null;
    }

    const signCell = targetSheet.getRange(rowIndex, 6); // Cột F: Chữ ký
    const rawValue = signCell.getValue();
    const rawFormula = signCell.getFormula();

    // TRƯỜNG HỢP 1: Chèn hình ảnh trực tiếp trong ô (CellImage)
    if (rawValue && typeof rawValue === "object") {
      try {
        if (typeof rawValue.getBlob === "function") {
          const b = rawValue.getBlob();
          if (b && b.getBytes().length > 0) return b;
        }
      } catch (eBlob) {}

      try {
        const getContentUrlFn = rawValue.getContentUrl;
        if (typeof getContentUrlFn === "function") {
          const contentUrl = rawValue.getContentUrl();
          if (contentUrl) {
            // Thử 1: Fetch kèm Bearer OAuth Token (chuẩn cho Google Sheets CellImage)
            try {
              const res = UrlFetchApp.fetch(contentUrl, {
                headers: { Authorization: "Bearer " + ScriptApp.getOAuthToken() },
                muteHttpExceptions: true,
                followRedirects: true
              });
              if (res.getResponseCode() === 200) {
                const b = res.getBlob();
                if (b && b.getBytes().length > 0) return b;
              }
            } catch (eAuth) {
              Logger.log("Fetch CellImage with Auth error: " + eAuth.toString());
            }

            // Thử 2: Fetch thông thường
            try {
              const res2 = UrlFetchApp.fetch(contentUrl, {
                muteHttpExceptions: true,
                followRedirects: true
              });
              if (res2.getResponseCode() === 200) {
                const b2 = res2.getBlob();
                if (b2 && b2.getBytes().length > 0) return b2;
              }
            } catch (eDirect) {
              Logger.log("Fetch CellImage direct error: " + eDirect.toString());
            }
          }
        }
      } catch (eCellImg) {
        Logger.log("Lỗi trích xuất CellImage: " + eCellImg.toString());
      }
    }

    // TRƯỜNG HỢP 2: Hình ảnh chèn đè lên ô (OverGridImage) tại ô F{rowIndex}
    try {
      const overGridImages = targetSheet.getImages();
      for (let i = 0; i < overGridImages.length; i++) {
        const img = overGridImages[i];
        const anchor = img.getAnchorCell();
        if (anchor && anchor.getRow() === rowIndex && anchor.getColumn() === 6) {
          const b = img.getBlob();
          if (b && b.getBytes().length > 0) return b;
        }
      }
    } catch (eOver) {}

    // TRƯỜNG HỢP 3: Dùng công thức =IMAGE("URL")
    if (rawFormula && rawFormula.toUpperCase().indexOf("=IMAGE") === 0) {
      const imageMatch = rawFormula.match(/=IMAGE\(\s*"([^"]+)"/i);
      if (imageMatch) {
        try {
          return UrlFetchApp.fetch(imageMatch[1]).getBlob();
        } catch (eImg) {}
      }
    }

    // TRƯỜNG HỢP 4: Nhập link ảnh trực tiếp (Drive URL hoặc HTTP link)
    if (typeof rawValue === "string" && rawValue.indexOf("http") === 0) {
      const rawUrl = rawValue.trim();
      let fileId = "";

      if (rawUrl.indexOf("drive.google.com/file/d/") !== -1) {
        const driveMatch = rawUrl.match(/\/d\/(.+?)(\/|$)/);
        if (driveMatch) fileId = driveMatch[1];
      } else {
        const paramMatch = rawUrl.match(/[?&]id=([^&]+)/);
        if (paramMatch) fileId = paramMatch[1];
      }

      if (fileId) {
        return DriveApp.getFileById(fileId).getBlob();
      }

      return UrlFetchApp.fetch(rawUrl).getBlob();
    }

    return null;
  } catch (error) {
    Logger.log(`Lỗi lấy chữ ký người dùng: ${error.toString()}`);
    return null;
  }
};
