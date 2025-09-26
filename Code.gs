// ===================================================
// === 1. การตั้งค่าและฟังก์ชันหลัก (Configuration & Core Functions) ===
// ===================================================
const CONFIG = {
  salesSheetName: "ข้อมูลการขาย",
  customerSheetName: "Customer",
  webAppInfo: { configSheet: "config", documentInfoSheet: "ข้อมูลเอกสาร", settingsSheet: "Settings" },
  stockSheetId: "19MvkCOZfUuQKjaeCYHKV5UTgSv-09PqpgIiTbX6qKWk",
  masterStockSheetName: "คลัง"
};
function doGet() {
  if (checkUserAccess_()) {
    return HtmlService.createTemplateFromFile('WebApp').evaluate().setTitle("ระบบขายและจัดการลูกค้า").setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  } else {
    return HtmlService.createHtmlOutputFromFile('AccessDenied').setTitle("Access Denied");
  }
}
function include(filename) { return HtmlService.createHtmlOutputFromFile(filename).getContent(); }
function _fetchAndProcessStockData() {
  try {
    const stockSs = SpreadsheetApp.openById(CONFIG.stockSheetId);
    const sheet = stockSs.getSheetByName(CONFIG.masterStockSheetName);
    if (!sheet || sheet.getLastRow() < 2) return { stockData: { headers: [], data: [] }, productList: [] };
    const range = sheet.getDataRange();
    const displayValues = range.getDisplayValues();
    const values = range.getValues();
    const headers = displayValues.shift();
    values.shift();
    const productList = values.map(row => {
      const [productId, productName, stock] = row;
      if (productId && productName) {
        return { displayValue: `${productId} - ${productName}`, productName, stockCentral: parseFloat(stock) || 0 };
      }
      return null;
    }).filter(Boolean);
    const stockData = { headers: headers, data: displayValues };
    return { stockData, productList };
  } catch (e) {
    console.error("_fetchAndProcessStockData Error:", e.message);
    return { stockData: { headers: [], data: [], error: e.message }, productList: [] };
  }
}
function getInitialData() {
  try {
    const sales = getSalesHistory();
    const customers = getCustomers();
    const employees = getEmployeeList();
    const { stockData, productList } = _fetchAndProcessStockData();
    return { salesData: sales, customerData: customers, productList: productList, employeeList: employees, stockData: stockData };
  } catch (e) {
    console.error("getInitialData Error: " + e.message);
    return { error: e.message };
  }
}
function checkUserAccess_() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const settingsSheet = ss.getSheetByName(CONFIG.webAppInfo.settingsSheet);
    if (!settingsSheet) return true;
    const lastRow = settingsSheet.getLastRow();
    if (lastRow < 2) return false;
    const allowedEmails = new Set(settingsSheet.getRange(`A2:A${lastRow}`).getValues().flat().map(e => e.toString().trim().toLowerCase()).filter(Boolean));
    const currentUser = Session.getActiveUser().getEmail().toLowerCase();
    return allowedEmails.has(currentUser);
  } catch (e) { console.error("checkUserAccess_ Error: " + e.toString()); return false; }
}
// ===================================================
// === 2. ฟังก์ชันจัดการข้อมูลลูกค้า (Customer CRUD) ===
// ===================================================
function getCustomers() {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.customerSheetName);
    if (!sheet || sheet.getLastRow() < 2) return [];
    const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 4).getValues();
    return data.map(row => ({ id: row[0], name: row[1], tel: row[2], address: row[3] }));
  } catch (e) { console.error("getCustomers Error: " + e.message); return []; }
}
function addCustomer(customerData) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(CONFIG.customerSheetName) || ss.insertSheet(CONFIG.customerSheetName);
    if (sheet.getLastRow() === 0) { sheet.appendRow(['CustomerID', 'ชื่อลูกค้า', 'เบอร์ติดต่อ', 'ที่อยู่']); }
    const newId = generateCustomerId_(sheet);
    sheet.appendRow([newId, customerData.name, "'" + customerData.tel, customerData.address]);
    return { success: true, message: "เพิ่มข้อมูลลูกค้าสำเร็จ", newCustomer: { id: newId, name: customerData.name, tel: customerData.tel, address: customerData.address }};
  } catch (e) { console.error("addCustomer Error: " + e.message); return { success: false, message: e.message }; }
}
/**
 * อัปเดตข้อมูลลูกค้า
 * @param {object} customerData ข้อมูลลูกค้าที่ต้องการแก้ไข
 * @returns {object} ผลลัพธ์การทำงานและข้อมูลที่อัปเดต
 */
function updateCustomer(customerData) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.customerSheetName);
    const result = sheet.getRange("A:A").createTextFinder(customerData.id).findNext();
    
    if (!result) throw new Error("ไม่พบข้อมูลลูกค้า");
    
    const targetRow = result.getRow();

    // [FIXED] เพิ่ม single quote (') ข้างหน้าเบอร์โทรศัพท์เพื่อบังคับให้เป็น Text
    sheet.getRange(targetRow, 2, 1, 3).setValues([[customerData.name, "'" + customerData.tel, customerData.address]]);
    
    // บรรทัดนี้ยังคงไว้เผื่อกรณีสร้างเซลล์ใหม่
    sheet.getRange(targetRow, 3).setNumberFormat("@");
    
    return { 
      success: true, 
      message: "แก้ไขข้อมูลสำเร็จ",
      updatedCustomer: customerData
    };
  } catch (e) { 
    console.error("updateCustomer Error: " + e.message); 
    return { success: false, message: e.message }; 
  }
}
function deleteCustomer(customerId) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.customerSheetName);
    const result = sheet.getRange("A:A").createTextFinder(customerId).findNext();
    if (!result) throw new Error("ไม่พบข้อมูลลูกค้า");
    sheet.deleteRow(result.getRow());
    return { success: true, message: "ลบข้อมูลลูกค้าสำเร็จ" };
  } catch (e) { console.error("deleteCustomer Error: " + e.message); return { success: false, message: e.message }; }
}
// ===================================================
// === 3. ฟังก์ชันจัดการข้อมูลการขาย (Sales Functions) ===
// ===================================================
function saveSalesData(formData) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const salesSheet = ss.getSheetByName(CONFIG.salesSheetName) || ss.insertSheet(CONFIG.salesSheetName);
    const currentUser = Session.getActiveUser().getEmail() || 'Unknown';
    const timestamp = new Date();
    const invoiceId = generateInvoiceId_(salesSheet);
    if (salesSheet.getLastRow() === 0) {
      salesSheet.appendRow(['เลขที่เอกสาร', 'วันที่ขาย', 'ชื่อลูกค้า', 'เบอร์ติดต่อ', 'พนักงานขาย', 'ยอดรวม', 'ส่วนลด', 'ยอดสุทธิ', 'ชื่อสินค้า', 'จำนวน', 'หน่วย', 'ราคาต่อหน่วยย่อย', 'ราคารวม', 'จำนวนหน่วยย่อย', 'ผู้บันทึก', 'เวลาที่บันทึก']);
    }
    const recordsToSave = formData.items.map((item, index) => {
      const commonData = index === 0 ? [ invoiceId, timestamp, formData.customerName, "'" + formData.customerTel, formData.salesperson, formData.subtotal, formData.discount, formData.grandTotal ] : Array(8).fill('');
      return [...commonData, item.name, item.quantity, item.unitName, item.price, item.total, item.baseQuantity, currentUser, timestamp];
    });
    if (recordsToSave.length > 0) {
      salesSheet.getRange(salesSheet.getLastRow() + 1, 1, recordsToSave.length, recordsToSave[0].length).setValues(recordsToSave);
      updateStock_(formData.items, 'DEDUCT');
      return { success: true, docId: invoiceId, message: "บันทึกข้อมูลการขายสำเร็จ" };
    } else {
      throw new Error("ไม่พบรายการสินค้าที่จะบันทึก");
    }
  } catch (e) { console.error("saveSalesData Error: " + e.message); return { success: false, message: `เกิดข้อผิดพลาด: ${e.message}` }; }
}
function updateSale(formData) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const salesSheet = ss.getSheetByName(CONFIG.salesSheetName);
    const docId = formData.docId;
    if (!docId) throw new Error("ไม่พบเลขที่เอกสารสำหรับการอัปเดต");
    const oldItems = getSaleRecordByDocId_(salesSheet, docId);
    if (oldItems.length > 0) {
      updateStock_(oldItems, 'RETURN');
    }
    updateStock_(formData.items, 'DEDUCT');
    deleteRowsByDocId_(salesSheet, docId);
    const currentUser = Session.getActiveUser().getEmail() || 'Unknown';
    const timestamp = new Date();
    const recordsToSave = formData.items.map((item, index) => {
      const commonData = index === 0 ? [ docId, timestamp, formData.customerName, "'" + formData.customerTel, formData.salesperson, formData.subtotal, formData.discount, formData.grandTotal ] : Array(8).fill('');
      return [...commonData, item.name, item.quantity, item.unitName, item.price, item.total, item.baseQuantity, currentUser, timestamp];
    });
    if (recordsToSave.length > 0) {
      salesSheet.getRange(salesSheet.getLastRow() + 1, 1, recordsToSave.length, recordsToSave[0].length).setValues(recordsToSave);
    }
    return { success: true, message: `อัปเดตเอกสาร ${docId} สำเร็จ` };
  } catch (e) { console.error("updateSale Error: " + e.message); return { success: false, message: e.message }; }
}
function deleteSaleById(docId) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const salesSheet = ss.getSheetByName(CONFIG.salesSheetName);
    const itemsToReturn = getSaleRecordByDocId_(salesSheet, docId);
    if (itemsToReturn.length > 0) {
      updateStock_(itemsToReturn, 'RETURN');
    }
    deleteRowsByDocId_(salesSheet, docId);
    return { success: true, message: `ลบเอกสาร ${docId} และคืนสต็อกสำเร็จ` };
  } catch (e) { console.error("deleteSaleById Error: " + e.message); return { success: false, message: e.message }; }
}
function getSalesHistory() {
  try{
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const salesSheet = ss.getSheetByName(CONFIG.salesSheetName);
    if (!salesSheet || salesSheet.getLastRow() < 2) return [];
    const data = salesSheet.getRange(2, 1, salesSheet.getLastRow() - 1, salesSheet.getLastColumn()).getValues();
    const groupedSales = {}, timezone = ss.getSpreadsheetTimeZone();
    let currentDocId = '';
    data.forEach(row => {
      const docId = row[0] || currentDocId;
      if (row[0]) {
        currentDocId = docId;
        groupedSales[docId] = { docId: docId, date: Utilities.formatDate(new Date(row[1]), timezone, 'dd/MM/yyyy'), customerName: row[2], customerTel: row[3], salesperson: row[4], subtotal: row[5], discount: row[6], grandTotal: row[7], items: [] };
      }
      if (groupedSales[docId]) {
        groupedSales[docId].items.push({ name: row[8], quantity: row[9], unitName: row[10], price: row[11], total: row[12] });
      }
    });
    return Object.values(groupedSales).reverse();
  } catch (e) { console.error("getSalesHistory Error: " + e.message); return []; }
}
// ===================================================
// === 4. ฟังก์ชันสำหรับหน้าคลังสินค้า (Stock Functions) ===
// ===================================================
function updateStock_(items, mode = 'DEDUCT') {
  if (!items || items.length === 0) return;
  try {
    const stockSs = SpreadsheetApp.openById(CONFIG.stockSheetId);
    const sheet = stockSs.getSheetByName(CONFIG.masterStockSheetName);
    if (!sheet) throw new Error(`ไม่พบชีต '${CONFIG.masterStockSheetName}'`);
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return;
    const stockDataRange = sheet.getRange(2, 2, lastRow - 1, 2);
    const stockValues = stockDataRange.getValues();
    const stockMap = new Map();
    stockValues.forEach((row, index) => {
      const productName = row[0].toString().trim();
      if (productName) {
        stockMap.set(productName, { stock: parseFloat(row[1]) || 0, arrayIndex: index });
      }
    });
    items.forEach(item => {
      const productName = item.name.trim();
      const quantity = item.baseQuantity;
      if (stockMap.has(productName)) {
        const productData = stockMap.get(productName);
        if (mode === 'DEDUCT') {
          productData.stock -= quantity;
        } else if (mode === 'RETURN') {
          productData.stock += quantity;
        }
        stockValues[productData.arrayIndex][1] = productData.stock;
      } else {
        console.warn(`Product "${productName}" not found in stock sheet.`);
      }
    });
    stockDataRange.setValues(stockValues);
  } catch (e) {
    console.error(`Failed to update stock in ${mode} mode. Error: ${e.toString()}`);
    throw new Error(`ไม่สามารถอัปเดตสต็อกได้: ${e.message}`);
  }
}
// ===================================================
// === 5. ฟังก์ชันดึงข้อมูลสำหรับ Dropdown (Data Fetchers) ===
// ===================================================

/**
 * [HELPER] ฟังก์ชันกลางสำหรับจัดการ Cache
 */
function getCachedData_(key, dataFetcher, expirationInSeconds = 3600) {
  const cache = CacheService.getScriptCache();
  const cached = cache.get(key);
  if (cached != null) {
    return JSON.parse(cached);
  }
  const data = dataFetcher();
  cache.put(key, JSON.stringify(data), expirationInSeconds);
  return data;
}

function getEmployeeList() {
    return getCachedData_('employeeList', () => {
        try {
            const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.webAppInfo.configSheet);
            if (!sheet || sheet.getLastRow() < 3) return [];
            const values = sheet.getRange(`N3:N${sheet.getLastRow()}`).getValues().flat();
            return [...new Set(values.filter(name => String(name).trim()))];
        } catch (e) {
            console.error("getEmployeeList Error: " + e.message); return [];
        }
    });
}
// ===================================================
// === 6. ฟังก์ชันจัดการเอกสารและ ID (Utility Functions) ===
// ===================================================
function getSaleRecordByDocId_(sheet, docId) {
  if (!sheet || !docId) return [];
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  const data = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();
  const items = [];
  let found = false;
  for (const row of data) {
    if (row[0].toString().trim() === docId) {
      found = true;
    }
    if (found) {
      const productName = row[8];
      const baseQuantity = parseFloat(row[13]);
      if (productName && !isNaN(baseQuantity)) {
        items.push({ name: productName, baseQuantity: baseQuantity });
      }
      const nextRowIndex = data.indexOf(row) + 1;
      if (nextRowIndex < data.length && data[nextRowIndex][0]) {
        break; 
      }
    }
  }
  return items;
}
function getDocumentInfo() {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.webAppInfo.documentInfoSheet);
    if (!sheet) throw new Error(`ไม่พบชีต '${CONFIG.webAppInfo.documentInfoSheet}'`);
    const data = sheet.getRange("A2:B5").getValues();
    const docInfo = {};
    data.forEach(row => {
      const key = row[0].toString().trim().replace('ที่อยู่ 1', 'address1').replace('ที่อยู่ 2', 'address2').replace('ชื่อบริษัท', 'companyName').replace('ข้อมูลติดต่อ', 'contactInfo');
      docInfo[key] = row[1];
    });
    return docInfo;
  } catch (e) { return { companyName: "ชื่อบริษัทของคุณ", address1: "ที่อยู่ 1", address2: "ที่อยู่ 2", contactInfo: "ข้อมูลติดต่อ" }; }
}
function saveDocumentInfo(settingsData) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.webAppInfo.documentInfoSheet);
    if (!sheet) throw new Error(`ไม่พบชีต '${CONFIG.webAppInfo.documentInfoSheet}'`);
    sheet.getRange("B2").setValue(settingsData.companyName);
    sheet.getRange("B3").setValue(settingsData.address1);
    sheet.getRange("B4").setValue(settingsData.address2);
    const contactCell = sheet.getRange("B5");
    contactCell.setNumberFormat("@");
    contactCell.setValue(settingsData.contactInfo);
    return { success: true, message: "บันทึกการตั้งค่าสำเร็จ!" };
  } catch (e) { return { success: false, message: e.message }; }
}
function generateCustomerId_(sheet) {
  if (!sheet || sheet.getLastRow() < 2) return "CUS-001";
  const lastId = sheet.getRange(sheet.getLastRow(), 1).getValue().toString();
  const match = lastId.match(/(\d+)$/);
  if (!match) return "CUS-001";
  return "CUS-" + String(parseInt(match[1], 10) + 1).padStart(3, '0');
}
function generateInvoiceId_(sheet) {
  const today = new Date();
  const datePrefix = `INV-${today.getFullYear()}${String(today.getMonth() + 1).padStart(2, '0')}${String(today.getDate()).padStart(2, '0')}-`;
  if (!sheet || sheet.getLastRow() < 2) return datePrefix + "1";
  const range = sheet.getRange("A:A");
  const allMatches = range.createTextFinder(datePrefix).matchEntireCell(false).matchCase(true).findAll();
  if (allMatches.length === 0) return datePrefix + "1";
  let maxNum = 0;
  allMatches.forEach(cell => {
    const num = parseInt(cell.getValue().split('-')[2], 10);
    if (num > maxNum) maxNum = num;
  });
  return datePrefix + (maxNum + 1);
}
function deleteRowsByDocId_(sheet, docId) {
  if (!sheet || !docId) return;
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;
  const docIdColumnValues = sheet.getRange(2, 1, lastRow - 1, 1).getValues().flat();
  const firstIndex = docIdColumnValues.indexOf(docId.toString().trim());
  if (firstIndex === -1) return;
  let rowCount = 1;
  for (let i = firstIndex + 1; i < docIdColumnValues.length; i++) {
    if (!docIdColumnValues[i]) { rowCount++; } else { break; }
  }
  sheet.deleteRows(firstIndex + 2, rowCount);
}
function clearServerCache() {
  try {
    CacheService.getScriptCache().removeAll(['productList', 'employeeList']);
    return { success: true, message: 'ล้างแคชฝั่งเซิร์ฟเวอร์สำเร็จ' };
  } catch (e) { console.error("clearServerCache Error: " + e.message); return { success: false, message: e.message }; }
}
