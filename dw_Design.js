function uploadPdfDesignToDrive(pdfBase64, excelBase64, filePdfName, fileExcelName, dataId) {
  try {
    // 1. Giải mã Base64 thành Blob
    var decodedData = Utilities.base64Decode(pdfBase64);
    var blobPDF = Utilities.newBlob(decodedData, MimeType.PDF, filePdfName);

    var decodedExcelData = Utilities.base64Decode(excelBase64);
    // Lưu ý: Dùng định dạng chuẩn cho file Excel .xlsx
    var blobExcel = Utilities.newBlob(decodedExcelData, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", fileExcelName);

    var sheet = SpreadsheetApp.openById("1DRteBSFT1cj4R_OUPMoDxeLMzAIJexWF3HPT-rpMOoM").getSheetByName("Data");

    // 2. Lấy Folder trên Drive
    var FOLDER_ID = "1PGwlJ484P8xki8i1k0GiyJvtwdMHu8A0"; 
    var folder = DriveApp.getFolderById(FOLDER_ID);

    // 3. Tạo file trên Drive
    var pdfFile = folder.createFile(blobPDF);
    var excelFile = folder.createFile(blobExcel);

// Chia sẻ cho bất kỳ ai TRONG CÔNG TY có link đều xem được
    pdfFile.setSharing(DriveApp.Access.DOMAIN_WITH_LINK, DriveApp.Permission.VIEW);
    excelFile.setSharing(DriveApp.Access.DOMAIN_WITH_LINK, DriveApp.Permission.VIEW);
    // Lấy URL của file vừa tạo
    var pdfUrl = pdfFile.getUrl();
    var excelUrl = excelFile.getUrl();

    // 4. TÌM ID VÀ GHI LINK VÀO GOOGLE SHEET
    if (dataId && sheet) {
      var lastRow = sheet.getLastRow();
      if (lastRow >= 1) {
        // Lấy toàn bộ dữ liệu cột A (Cột 1) để tìm ID
        var idColumn = sheet.getRange(1, 1, lastRow, 1).getValues();
        var targetRowIndex = -1;

        // Quét tìm ID
        for (var i = 0; i < idColumn.length; i++) {
          if (String(idColumn[i][0]).trim() === String(dataId).trim()) {
            targetRowIndex = i + 1; // Index hàng của Sheet bắt đầu từ 1
            break;
          }
        }

        if (targetRowIndex !== -1) {
          // Cột AE là 31, cột AF là 32
          // Cập nhật 1 hàng, 2 cột liên tiếp với [ [Excel, PDF] ]
          sheet.getRange(targetRowIndex, 31, 1, 2).setValues([[excelUrl, pdfUrl]]);
        } else {
          Logger.log("Cảnh báo: Tải file thành công nhưng không tìm thấy ID '" + dataId + "' trong sheet để ghi link.");
        }
      }
    }

    // 5. Trả về Link URL để frontend truyền tiếp cho Local Server (nếu cần)
    return {
      pdfUrl: pdfUrl,
      excelUrl: excelUrl
    };
    
  } catch (error) {
    Logger.log("Lỗi upload file: " + error.toString());
    throw new Error("Lỗi khi lưu file lên Drive: " + error.message);
  }
}



function submitForApproval(dataId, base64Data) {
  var targetFolderId = "1lJEJnMhRW8C-ykqR-EHMMQWkcRtO2Abz";
  try {
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
      } catch(e) {
        newFile.setSharing(DriveApp.Access.DOMAIN_WITH_LINK, DriveApp.Permission.VIEW);
      }

      var fileUrl = newFile.getUrl();

      // 2. Tìm ID bằng TextFinder (Nhanh gấp 10 lần vòng lặp For)
      var foundCell = sheet.getRange("A:A").createTextFinder(dataId).matchEntireCell(true).findNext();
      
      if (foundCell) {
        var targetRowIndex = foundCell.getRow();
        
        // Cập nhật Cột 3 (Trạng thái) và Cột 32 (AF - Link File PDF Ký)
        sheet.getRange(targetRowIndex, 3).setValue("Chờ Duyệt");
        sheet.getRange(targetRowIndex, 32).setValue(fileUrl);
        
        return fileUrl; // Trả link về cho giao diện Swal.fire hiện lên
      } else {
        throw new Error("Không tìm thấy ID '" + dataId + "' trong sheet để ghi link.");
      }
    }
  } catch (error) {
    Logger.log("Lỗi lưu trữ: " + error.toString());
    throw new Error("Lỗi khi lưu trữ data: " + error.message);
  }
}



function getFilesFromDrive(pdfLink, userMail) {  
  try {
    var fileIdMatch = pdfLink.match(/[-\w]{25,}/);
    
    if (!fileIdMatch) {
      throw new Error("Link Drive không hợp lệ hoặc không chứa ID.");
    }
    
    // 1. Lấy PDF Blob
    var fileId = fileIdMatch[0]; 
    var pdfBlob = DriveApp.getFileById(fileId).getBlob();
    
    // 2. Lấy Chữ ký Blob dựa vào Email
    var stampBlob = getSignBlob(userMail);
    
    if (!stampBlob) {
      throw new Error("Không thể trích xuất chữ ký của tài khoản: " + userMail);
    }
    
    return {
      success: true,
      pdfBase64: Utilities.base64Encode(pdfBlob.getBytes()),
      stampBase64: Utilities.base64Encode(stampBlob.getBytes())
    };
  } catch(e) {
    return { success: false, message: e.toString() };
  }
}


function getSignBlob(userMail) {
  try {
    var sheet = SpreadsheetApp.openById("1DRteBSFT1cj4R_OUPMoDxeLMzAIJexWF3HPT-rpMOoM").getSheetByName("User");
    
    // Dùng TextFinder quét Cột D (Cột Email) siêu tốc
    var foundCell = sheet.getRange("D:D").createTextFinder(userMail).matchEntireCell(true).findNext();
    
    if (!foundCell) return null;
    
    // Lấy vị trí ô chứa chữ ký (Cột F - Tương ứng số 6)
    var rowIndex = foundCell.getRow();
    var signCell = sheet.getRange(rowIndex, 6); 
    
    // Rút xuất trực tiếp Giá trị và Công thức của ô đó ngay trên Backend
    var rawVal = signCell.getValue();
    var rawFormula = signCell.getFormula();
    var url = "";

    // TRƯỜNG HỢP 1: Chèn ảnh trực tiếp vào ô (CellImage)
    if (rawVal && typeof rawVal === 'object' && typeof rawVal.getContentUrl === 'function') {
      url = rawVal.getContentUrl();
      return UrlFetchApp.fetch(url).getBlob();
    } 
    
    // TRƯỜNG HỢP 2: Chèn ảnh bằng hàm =IMAGE()
    else if (rawFormula && rawFormula.toUpperCase().indexOf('=IMAGE') === 0) {
      var match = rawFormula.match(/=IMAGE\(\s*"([^"]+)"/i);
      if (match) {
        url = match[1];
        return UrlFetchApp.fetch(url).getBlob();
      }
    } 
    
    // TRƯỜNG HỢP 3: Dán link Text (Google Drive hoặc Web URL)
    else if (typeof rawVal === 'string' && rawVal.indexOf('http') === 0) {
      var rawUrl = rawVal.trim();
      var fileId = "";
      
      if (rawUrl.indexOf("drive.google.com/file/d/") !== -1) {
        var matchDrive = rawUrl.match(/\/d\/(.+?)(\/|$)/);
        if (matchDrive) fileId = matchDrive[1];
      } else if (rawUrl.match(/[?&]id=([^&]+)/)) {
        var matchParam = rawUrl.match(/[?&]id=([^&]+)/);
        if (matchParam) fileId = matchParam[1];
      }

      if (fileId) {
        return DriveApp.getFileById(fileId).getBlob();
      } else {
        return UrlFetchApp.fetch(rawUrl).getBlob();
      }
    }
    
    return null; // Không rơi vào trường hợp nào
    
  } catch (e) {
    Logger.log("Lỗi khi lấy Blob Chữ ký: " + e.message);
    return null;
  }
}

