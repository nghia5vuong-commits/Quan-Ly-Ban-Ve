function getAllData() {
    const ss = SpreadsheetApp.openById("1DRteBSFT1cj4R_OUPMoDxeLMzAIJexWF3HPT-rpMOoM");
    const sheet = ss.getSheetByName("Data");
    const lastrow = sheet.getLastRow();
    const lastcol = sheet.getLastColumn();
    
    if (lastrow <= 1) return [];
    
    const range = sheet.getRange(2, 1, lastrow - 1, lastcol);
    const data = range.getValues();
    
    // Xử lý định dạng ngày tháng cho toàn bộ mảng dữ liệu
    const formattedData = data.map(row => {
        return row.map(cell => {
            // Kiểm tra nếu ô dữ liệu là kiểu Ngày (Date)
            if (cell instanceof Date) {
                // Định dạng lại thành Ngày/Tháng/Năm
                return Utilities.formatDate(cell, Session.getScriptTimeZone(), "dd/MM/yyyy");
            }
            return cell; // Nếu không phải ngày tháng thì giữ nguyên
        });
    });
    
    return formattedData;
}

 