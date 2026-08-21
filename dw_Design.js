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

const findRowIndexById = (sheet, dataId) => {
  if (!sheet || !dataId) {
    return -1;
  }

  const foundCell = sheet.getRange("A:A").createTextFinder(normalizeText(dataId)).matchEntireCell(true).findNext();
  return foundCell ? foundCell.getRow() : -1;
};

const uploadPdfDesignToDrive = (pdfBase64, excelBase64, filePdfName, fileExcelName, dataId) => {
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

    if (dataId && sheet) {
      const targetRowIndex = findRowIndexById(sheet, dataId);
      if (targetRowIndex > 0) {
        sheet.getRange(targetRowIndex, designSheetConfig.fileLinkStartColumn, 1, designSheetConfig.fileLinkColumnCount).setValues([[excelUrl, pdfUrl]]);
      }
    }

    return {
      pdfUrl,
      excelUrl,
    };
  } catch (error) {
    throw new Error(`Lỗi khi lưu file lên Drive: ${error.message}`);
  }
};

const submitForApproval = (dataId, base64Data, stt) => {
  try {
<<<<<<< HEAD
    const sheet = getDesignSheet();
    if (!dataId || !sheet) {
      return null;
=======
    var sheet = SpreadsheetApp.openById("1DRteBSFT1cj4R_OUPMoDxeLMzAIJexWF3HPT-rpMOoM").getSheetByName("Data");

    if (dataId && sheet) {
      // 1. Tạo file mới trên Drive từ dữ liệu Local Server trả về
      var decodedData = Utilities.base64Decode(base64Data);
      var fileName = dataId + ".pdf";
      var blob = Utilities.newBlob(decodedData, 'application/pdf', fileName);

      var folder = DriveApp.getFolderById(targetFolderId);
      var newFile = folder.createFile(blob);

      // Cấp quyền xem cho file 
      try {
        newFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      } catch (e) {
        newFile.setSharing(DriveApp.Access.DOMAIN_WITH_LINK, DriveApp.Permission.VIEW); 
      }

      var fileUrl = newFile.getUrl();

      // 2. Tìm ID bằng TextFinder (Nhanh gấp 10 lần vòng lặp For)
      var foundCell = sheet.getRange("A:A").createTextFinder(dataId).matchEntireCell(true).findNext();

      if (foundCell) {
        var targetRowIndex = foundCell.getRow();

        // Cập nhật Cột 3 (Trạng thái) và Cột 32 (AF - Link File PDF Ký)
        sheet.getRange(targetRowIndex, 3).setValue("Chờ Checker 1");
        sheet.getRange(targetRowIndex, 32).setValue(fileUrl);

        return fileUrl; // Trả link về cho giao diện Swal.fire hiện lên
      } else {
        throw new Error("Không tìm thấy ID '" + dataId + "' trong sheet để ghi link.");
      }
>>>>>>> 9bc7433d99a2711db9b15354f7e483b5890cda31
    }

    const folder = DriveApp.getFolderById(designSheetConfig.approvalFolderId);
    const decodedData = Utilities.base64Decode(base64Data);
    const fileName = `${dataId}.pdf`;
    const pdfBlob = Utilities.newBlob(decodedData, "application/pdf", fileName);
    const newFile = folder.createFile(pdfBlob);

    try {
      newFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    } catch (error) {
      newFile.setSharing(DriveApp.Access.DOMAIN_WITH_LINK, DriveApp.Permission.VIEW);
    }

    const fileUrl = newFile.getUrl();
    const targetRowIndex = findRowIndexById(sheet, dataId);

    if (targetRowIndex <= 0) {
      throw new Error(`Không tìm thấy ID '${dataId}' trong sheet để ghi link.`);
    }

    sheet.getRange(targetRowIndex, 3).setValue("Chờ Checker 1");
    sheet.getRange(targetRowIndex, 32).setValue(fileUrl);

    return fileUrl;
  } catch (error) {
    throw new Error(`Lỗi khi lưu trữ data: ${error.message}`);
  }
};

const getFilesFromDrive = (pdfLink, userMail) => {
  try {
    const fileIdMatch = pdfLink.match(/[-\w]{25,}/);
    if (!fileIdMatch) {
      throw new Error("Link Drive không hợp lệ hoặc không chứa ID.");
    }

    const fileId = fileIdMatch[0];
    const pdfBlob = DriveApp.getFileById(fileId).getBlob();
    const stampBlob = getSignBlob(userMail);

    if (!stampBlob) {
      throw new Error(`Không thể trích xuất chữ ký của tài khoản: ${userMail}`);
    }

    return {
      success: true,
      pdfBase64: Utilities.base64Encode(pdfBlob.getBytes()),
      stampBase64: Utilities.base64Encode(stampBlob.getBytes()),
    };
  } catch (error) {
    return {
      success: false,
      message: error.toString(),
    };
  }
};

const getSignBlob = (userMail) => {
  try {
    const sheet = SpreadsheetApp.openById(designSheetConfig.spreadsheetId).getSheetByName("User");
    if (!sheet) {
      return null;
    }

    const foundCell = sheet.getRange("D:D").createTextFinder(userMail).matchEntireCell(true).findNext();
    if (!foundCell) {
      return null;
    }

    const rowIndex = foundCell.getRow();
    const signCell = sheet.getRange(rowIndex, 6);
    const rawValue = signCell.getValue();
    const rawFormula = signCell.getFormula();

    if (rawValue && typeof rawValue === "object" && typeof rawValue.getContentUrl === "function") {
      const imageUrl = rawValue.getContentUrl();
      return UrlFetchApp.fetch(imageUrl).getBlob();
    }

    if (rawFormula && rawFormula.toUpperCase().indexOf("=IMAGE") === 0) {
      const imageMatch = rawFormula.match(/=IMAGE\(\s*"([^"]+)"/i);
      if (imageMatch) {
        return UrlFetchApp.fetch(imageMatch[1]).getBlob();
      }
    }

    if (typeof rawValue === "string" && rawValue.indexOf("http") === 0) {
      const rawUrl = rawValue.trim();
      let fileId = "";

      if (rawUrl.indexOf("drive.google.com/file/d/") !== -1) {
        const driveMatch = rawUrl.match(/\/d\/(.+?)(\/|$)/);
        if (driveMatch) {
          fileId = driveMatch[1];
        }
      } else {
        const paramMatch = rawUrl.match(/[?&]id=([^&]+)/);
        if (paramMatch) {
          fileId = paramMatch[1];
        }
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
