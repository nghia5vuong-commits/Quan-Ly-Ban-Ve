/**
 * TOÀN CỤC: Khai báo các ID tĩnh (Nên đưa vào PropertiesService trong tương lai)
 */
const CONFIG = {
  FOLDER_CACHE: "1o_51QiYFRMGZnJHhMVZkISsHr6XkLDI6",
  FOLDER_TARGET: "1IiydBBRI1olOz5aPTsGXRYBq3kW0vCFR",
  SHEET_SOURCE: "1gbDjHE6i0_G6gxTcLQUD26HkMI6TzQ0J8Q3UzscknBM",
  SHEET_LOG: "1tDdXQDl1k7yc-9ZeXP4EQzrx-1TRQNKGYTdCNYSJ1jc",
  STD_SHEET: "1PEGA2lFc7y71mcox8dDvXZBgG2XCUJfASqp1hbTdrEw"
};

/**
 * 1. TẢI TRƯỚC DỮ LIỆU KHI MỞ TRANG
 * @returns {Object} Dữ liệu Map, danh sách khách hàng và DW
 */
function getInitialData() {
  try {
    const sheet = SpreadsheetApp.openById(CONFIG.SHEET_SOURCE).getSheetByName("TRN_checksheet");
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return { dataMap: {}, customers: [], dwList: [] };

    const rawData = sheet.getRange(2, 1, lastRow - 1, 31).getValues();
    const dataMap = {};
    const uniqueCust = new Set();
    const dwListMap = {}; 

    rawData.forEach(row => {
      // 1. Khách hàng
      const rawCust = String(row[10]).trim();
      let cleanCust = (rawCust && rawCust !== "undefined" && rawCust !== "null") ? rawCust : "Không có dữ liệu";
      const dashIndex = cleanCust.indexOf('-');
      if (dashIndex !== -1) cleanCust = cleanCust.substring(dashIndex + 1).trim();
      if (cleanCust !== "Không có dữ liệu") uniqueCust.add(cleanCust);

      // 2. Data Map
      const code = String(row[8]).trim();
      if (code) {
        dataMap[code] = {
          customer: cleanCust, 
          h: row[29] || 0, w: row[28] || 0, l: row[5] || 0, link: row[14] || "", link_xlsx: row[30] || ""
        };
      }

      // 3. DW No & Version
      const dwno = String(row[11]).trim();
      const version = String(row[12]).trim();
      const to = String(row[8]).trim();
      const xlsx = String(row[30]).trim() || "";
      
      if (dwno && !xlsx) {
        dwListMap[`${dwno}|${version}|${to}`] = { dwno: dwno, version: version, to: to };
      }
    });

    return { 
      dataMap: dataMap, 
      customers: Array.from(uniqueCust).sort(),
      dwList: Object.values(dwListMap) 
    };
  } catch (e) {
    console.error("Lỗi getInitialData: ", e);
    return { dataMap: {}, customers: [], dwList: [] };
  }
}

/**
 * Xử lý copy file Excel, lưu ảnh và cập nhật Sheet
 * @param {Object} payload Dữ liệu từ Client
 * @returns {Object} Trạng thái và link mới
 */
function processCopyAndSave(payload) {
  try {
    const targetFolder = DriveApp.getFolderById(CONFIG.FOLDER_TARGET);
    
    // --- 1. XỬ LÝ LƯU ẢNH ---
    if (payload.imageBytes) {
      const decodedBytes = Utilities.base64Decode(payload.imageBytes);
      const newImageName = `${payload.to}.${payload.extension}`; 
      const blob = Utilities.newBlob(decodedBytes, payload.mimeType, newImageName);
      targetFolder.createFile(blob);
    }

    // --- 2. COPY FILE EXCEL ---
    const fileIdMatch = payload.link_xlsx.match(/[-\w]{25,}/);
    if (!fileIdMatch) throw new Error("Link Excel gốc không hợp lệ để copy.");
    
    const sourceFile = DriveApp.getFileById(fileIdMatch[0]);
    const copiedFile = sourceFile.makeCopy(payload.dwno, targetFolder);
    const newFileUrl = copiedFile.getUrl(); 

    // --- 3. CẬP NHẬT GOOGLE SHEET BẰNG SETVALUES (O(1)) ---
    const sheet = SpreadsheetApp.openById(CONFIG.SHEET_SOURCE).getSheetByName("TRN_checksheet");
    const data = sheet.getDataRange().getValues();
    let targetRowIndex = -1;

    // Duyệt mảng trên RAM để tìm dòng
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][11]).trim() === payload.dwno && String(data[i][12]).trim() === payload.version) {
        targetRowIndex = i + 1; 
        break;
      }
    }

    if (targetRowIndex === -1) {
      throw new Error("Copy file thành công, nhưng không tìm thấy dòng trùng khớp với DW No & Version trong Sheet gốc!");
    }

    // TỐI ƯU: Ghi 1 lần duy nhất cho 3 cột liên tiếp: W (cột 29), H (cột 30), Link (cột 31)
    sheet.getRange(targetRowIndex, 29, 1, 3).setValues([[payload.w, payload.h, newFileUrl]]);

    return { status: "success", newLink: newFileUrl }; 

  } catch(e) {
    console.error("Lỗi processCopyAndSave: ", e);
    throw new Error(e.message);
  }
}

/**
 * 2. HÀM TÌM KIẾM CỰC NHANH TRÊN GOOGLE DRIVE
 * @param {Array<string>} fileNames Danh sách tên file cần tìm
 * @returns {Object} Map chứa Base64 của ảnh
 */
function fetchImagesFromDrive(fileNames) {
  try {
    const fileIdMap = lookupFileIds(fileNames);
    const token = ScriptApp.getOAuthToken();
    const allReqs = [];
    const imageOrder = [];

    fileNames.forEach(name => {
      const fileId = fileIdMap[name.trim().replace(/\.[^/.]+$/, "").toLowerCase()];
      if (fileId) {
        allReqs.push({
          url: `https://www.googleapis.com/drive/v3/files/${fileId}?alt=media`,
          method: "get", 
          headers: { Authorization: `Bearer ${token}` }, 
          muteHttpExceptions: true
        });
        imageOrder.push(name);
      }
    });

    if (allReqs.length === 0) return {};

    const allRes = UrlFetchApp.fetchAll(allReqs);
    const imageMap = {};
    
    for (let r = 0; r < allRes.length; r++) {
      const imgName = imageOrder[r];
      if (imgName && allRes[r].getResponseCode() === 200) {
        const blob = allRes[r].getBlob();
        imageMap[imgName] = `data:${blob.getContentType()};base64,${Utilities.base64Encode(blob.getBytes())}`;
      }
    }

    return imageMap;
  } catch (e) { 
    console.error("Lỗi fetchImagesFromDrive: ", e);
    throw new Error("Lỗi tải tệp tin từ Google Drive: " + e.message); 
  }
}

/**
 * 3. GHI LOG NGẦM
 */
function logSearchActionAsync() {
  try {
    const email = Session.getActiveUser().getEmail();
    const now = new Date();
    SpreadsheetApp.openById(CONFIG.SHEET_LOG).getSheetByName("Thống kê")
      .appendRow([now.getFullYear(), now.getMonth() + 1, now.getDate(), email, now.toLocaleTimeString()]);
  } catch (e) {
    console.warn("Không thể ghi log: ", e);
  }
}

/**
 * Hàm Cache ID file tối ưu hiệu năng
 * @param {Array<string>} fileNames 
 * @returns {Object} Map File IDs
 */
function lookupFileIds(fileNames) {
  const cache = CacheService.getScriptCache();
  const fileIdMap = {}; 
  const uncached = [];
  const baseNames = [];
  
  fileNames.forEach(fn => {
    if (!fn) return;
    baseNames.push(fn.trim().replace(/\.[^/.]+$/, "").toLowerCase());
  });
  
  const cachedKeys = baseNames.map(b => `fid_${b}`);
  const cached = cache.getAll(cachedKeys);
  
  baseNames.forEach(base => {
    if (cached[`fid_${base}`]) fileIdMap[base] = cached[`fid_${base}`];
    else uncached.push(base);
  });

  if (uncached.length === 0) return fileIdMap;

  const folder = DriveApp.getFolderById(CONFIG.FOLDER_CACHE);
  const queryParts = uncached.map(base => `title contains '${base.replace(/'/g, "\\'")}'`);
  const files = folder.searchFiles(`(${queryParts.join(" or ")}) and trashed = false`);

  const uncachedSet = {}; 
  uncached.forEach(b => { uncachedSet[b] = true; });
  const newCache = {};
  
  while (files.hasNext()) {
    const f = files.next();
    const fBase = f.getName().replace(/\.[^/.]+$/, "").toLowerCase();
    
    if (uncachedSet[fBase]) {
      const fId = f.getId();
      fileIdMap[fBase] = fId; 
      newCache[`fid_${fBase}`] = fId;
      delete uncachedSet[fBase];
      if (Object.keys(uncachedSet).length === 0) break;
    }
  }
  
  if (Object.keys(newCache).length > 0) cache.putAll(newCache, 21600);
  return fileIdMap;
}

/**
 * LẤY DỮ LIỆU TIÊU CHUẨN THEO KHÁCH HÀNG
 * @param {string} customerName 
 * @returns {Array<Array>} Data mảng 2D
 */
function getCustomerStandard(customerName) {
  try {
    const sheet = SpreadsheetApp.openById(CONFIG.STD_SHEET).getSheetByName("STD");
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return [];

    const data = sheet.getRange(4, 2, lastRow - 1, 27).getValues();
    const results = [];
    const searchName = customerName.trim().toUpperCase();

    data.forEach(row => {
      const cust = String(row[0]).trim().toUpperCase();
      if (cust === searchName) {
        results.push(row.map(cell => {
          if (cell instanceof Date) {
            return Utilities.formatDate(cell, Session.getScriptTimeZone(), "dd/MM/yyyy");
          }
          if (cell === null || cell === undefined || cell === "") return "";
          return String(cell);
        }));
      }
    });

    return results;
  } catch (e) {
    console.error("Lỗi getCustomerStandard: ", e);
    throw new Error("Lỗi đọc Sheet: " + e.message);
  }
}