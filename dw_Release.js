function getAllDataForStats() {
  try {
    const ss = SpreadsheetApp.openById("1t5PWyoJHrxElWP3QgmB16BEMIvEC015NHq0tsxu_TpE");
    const sheet = ss.getSheetByName("data");
    
    if (!sheet) {
      Logger.log("Lỗi: Không tìm thấy sheet tên 'data'");
      return []; 
    }
    
    const lastrow = sheet.getLastRow();
    const lastcol = sheet.getLastColumn();
    
    if (lastrow <= 1) {
      return [];
    }
    
    const range = sheet.getRange(2, 1, lastrow - 1, lastcol);
    const data = range.getValues();
    
    const formattedData = data.map(row => {
      return row.map(cell => {
        if (cell instanceof Date) {
          return Utilities.formatDate(cell, Session.getScriptTimeZone(), "dd/MM/yyyy");
        }
        return cell;
      });
    });
    // console.log(formattedData[5][37]);
    
    return formattedData;
    
  } catch (e) {
    Logger.log("Lỗi trong getAllDataForStats: " + e.toString());
    return []; 
  }
}
function delDwNo(id) {
  const ss = SpreadsheetApp.openById("1t5PWyoJHrxElWP3QgmB16BEMIvEC015NHq0tsxu_TpE");
  const sheet = ss.getSheetByName("data");
  
  // Lấy toàn bộ ID ở cột A (Cột 1)
  const idColumnValues = sheet.getRange(1, 1, sheet.getLastRow(), 1).getValues();
  
  // Tìm vị trí của ID trong cột
  // Dùng String() để tránh lỗi so sánh số và chuỗi
  const rowIndex = idColumnValues.findIndex(row => String(row[0]) === String(id)) + 1;

  if (rowIndex > 0) {
    // Nếu tìm thấy, dọn dẹp nội dung hàng đó
    sheet.getRange(rowIndex, 1, 1, sheet.getLastColumn()).clearContent();
    return "Thành công";
  } else {
    throw new Error("Không tìm thấy ID bản vẽ. Có thể dữ liệu đã được thay đổi bởi người khác.");
  }
}



function delOrderNo(idList) {
  const ss = SpreadsheetApp.openById("1t5PWyoJHrxElWP3QgmB16BEMIvEC015NHq0tsxu_TpE");
  const sheet = ss.getSheetByName("data");
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();

  if (lastRow < 2) return;

  const range = sheet.getRange(2, 1, lastRow - 1, lastCol);
  const data = range.getValues();
  
  // Ép kiểu idList về mảng string để so sánh chính xác
  const targetIds = Array.isArray(idList) ? idList.map(String) : [String(idList)];

  // Duyệt dữ liệu thực tế ngay lúc này trên server
  data.forEach((row, index) => {
    const currentReqNo = String(row[8]); // Cột I
    if (targetIds.indexOf(currentReqNo) !== -1) {
      // index + 2 vì data bắt đầu từ hàng 2
      sheet.getRange(index + 2, 1, 1, lastCol).clearContent();
    }
  });
}

function updateDrawingFull(payload) {
  try {
    // 1. Sửa lỗi kiểm tra Object rỗng
    if (!payload || !payload.id) return "Dữ liệu không hợp lệ"; 
    
    const ss = SpreadsheetApp.openById("1t5PWyoJHrxElWP3QgmB16BEMIvEC015NHq0tsxu_TpE");
    const sheet = ss.getSheetByName("data");
    if (!sheet) throw new Error("Không tìm thấy Sheet 'data'");

    const lastrow = sheet.getLastRow();
    const lastcol = sheet.getLastColumn();
    
    if (lastrow < 2) return "Bảng dữ liệu trống";

    // Lấy toàn bộ mảng dữ liệu lên xử lý
    const dataRange = sheet.getRange(2, 1, lastrow - 1, lastcol);
    const data = dataRange.getValues();

    // 2. Dùng vòng lặp for thay vì forEach để có thể dừng sớm
    for (let i = 0; i < data.length; i++) {
      // Ép kiểu String để so sánh an toàn
      if (String(data[i][0]) === String(payload.id)) {
        
        // Cập nhật giá trị ngay trên Mảng (rất nhanh vì xử lý trong bộ nhớ)
        // Lưu ý: Index của mảng data[i] bắt đầu từ 0 (tương ứng cột A = 0, B = 1...)
        data[i][1] = payload.type;       // Cột B (2) -> Index 1
        data[i][4] = payload.dwNo;       // Cột E (5) -> Index 4
        data[i][5] = payload.ver;        // Cột F (6) -> Index 5
        data[i][2] = payload.lineCus;    // Cột C (3) -> Index 2
        data[i][3] = payload.lineCode;   // Cột D (4) -> Index 3
        data[i][9] = payload.designNo;   // Cột J (10)-> Index 9
        data[i][7] = payload.shape;      // Cột H (8) -> Index 7
        data[i][11] = payload.deptLq;    // Cột L (12)-> Index 11
        data[i][17] = payload.link;      // Cột R (18)-> Index 17
        data[i][33] = payload.note;      // Cột AH (34)-> Index 33
        data[i][39] = payload.changeType; // Cột Ah (40)-> Index 39

        // 3. Đổ lại ĐÚNG MỘT HÀNG đó xuống Sheet (Gọi API 1 lần duy nhất)
        const rowIndex = i + 2;
        sheet.getRange(rowIndex, 1, 1, lastcol).setValues([data[i]]);
        
        return "Cập nhật thành công"; // Thoát hàm ngay lập tức
      }
    }

    throw new Error("Không tìm thấy ID bản vẽ để cập nhật!");

  } catch (e) {
    Logger.log("Lỗi updateDrawingFull: " + e.message);
    // Quăng lỗi ra để client (withFailureHandler) có thể bắt được và hiển thị Swal.fire
    throw new Error(e.message); 
  }
}



function sendToChargeQA(idList) {
  const ss = SpreadsheetApp.openById("1t5PWyoJHrxElWP3QgmB16BEMIvEC015NHq0tsxu_TpE");
  const sheet = ss.getSheetByName("data");
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();

  if (lastRow < 2) return;
  let supID = [];
  const range = sheet.getRange(2, 1, lastRow - 1, lastCol);
  const data = range.getValues();
  
  // Ép kiểu idList về mảng string để so sánh chính xác
  const targetIds = Array.isArray(idList) ? idList.map(String) : [String(idList)];

  // Duyệt dữ liệu thực tế ngay lúc này trên server
  data.forEach((row, index) => {
    const currentReqNo = String(row[8]); // Cột I
    if (targetIds.indexOf(currentReqNo) !== -1 && row[33] !=="") {
      // index + 2 vì data bắt đầu từ hàng 2
      sheet.getRange(index + 2, 24).setValue("Dept đã chọn phương án");
      supID.push(row[0]);
    }
  });
  return supID;
}
