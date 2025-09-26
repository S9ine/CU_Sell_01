// ===================================================
// === 1. การตั้งค่าและฟังก์ชันหลัก (Configuration & Core Functions) ===
// ===================================================
const CONFIG = {
  salesSheetName: "ข้อมูลการขาย",
  customer: {
    sheetId: "19MvkCOZfUuQKjaeCYHKV5UTgSv-09PqpgIiTbX6qKWk",
    sheetName: "Contacts"
  },
  // [NEW] เพิ่มการตั้งค่าสำหรับชีตสต็อกแผงไข่
  trayStock: {
    sheetId: "19MvkCOZfUuQKjaeCYHKV5UTgSv-09PqpgIiTbX6qKWk",
    sheetName: "TrayStock"
  },
  webAppInfo: { configSheet: "config", documentInfoSheet: "ข้อมูลเอกสาร", settingsSheet: "Settings" },
  stockSheetId: "19MvkCOZfUuQKjaeCYHKV5UTgSv-09PqpgIiTbX6qKWk",
  masterStockSheetName: "คลัง"
};

function doGet(e) {
  if (checkUserAccess_()) {
    if (e.parameter.page) {
      const template = HtmlService.createTemplateFromFile('WebApp');
      template.initialPage = e.parameter.page;
      template.dashboardUrl = ScriptApp.getService().getUrl();
      return template.evaluate()
        .setTitle("ระบบขายและจัดการลูกค้า")
        .setSandboxMode(HtmlService.SandboxMode.IFRAME);
    } else {
      const template = HtmlService.createTemplateFromFile('Dashboard');
      template.metrics = getDashboardMetrics(); // ✨ ส่งข้อมูล metrics ไปที่หน้า Dashboard
      return template.evaluate()
        .setTitle("Dashboard | ระบบขาย")
        .setSandboxMode(HtmlService.SandboxMode.IFRAME);
    }
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

/**
 * [NEW] Helper function to get the customer sheet from the correct spreadsheet.
 * @returns {GoogleAppsScript.Spreadsheet.Sheet} The customer sheet object.
 */
function _getCustomerSheet() {
  try {
    const ss = SpreadsheetApp.openById(CONFIG.customer.sheetId);
    const sheet = ss.getSheetByName(CONFIG.customer.sheetName);
    if (!sheet) {
      throw new Error(`Sheet with name "${CONFIG.customer.sheetName}" not found in spreadsheet ID ${CONFIG.customer.sheetId}`);
    }
    return sheet;
  } catch(e) {
    console.error("Failed to open customer spreadsheet: " + e.message);
    throw new Error("ไม่สามารถเปิดไฟล์ข้อมูลลูกค้าได้ กรุณาตรวจสอบการตั้งค่า");
  }
}


function getCustomers() {
  try {
    const sheet = _getCustomerSheet(); 
    if (!sheet || sheet.getLastRow() < 2) return [];
    // [CHANGED] ดึงข้อมูล 4 คอลัมน์
    const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 4).getValues();
    // [CHANGED] ปรับการ map ข้อมูลให้ตรงกับคอลัมน์ใหม่ (เบอร์โทรศัพท์อยู่คอลัมน์ที่ 4) และตัดที่อยู่ออก
    return data.map(row => ({ id: row[0], name: row[1], tel: row[3] }));
  } catch (e) { console.error("getCustomers Error: " + e.message); return []; }
}

function addCustomer(customerData) {
  try {
    const sheet = _getCustomerSheet(); 
    // [CHANGED] ปรับ Header ให้ตรงกับโครงสร้างใหม่
    if (sheet.getLastRow() === 0) { 
      sheet.appendRow(['ContactID', 'ContactName', 'Type', 'Phone']); 
    }
    const newId = _generateNextId_(sheet, 'CUS');
    // [CHANGED] เพิ่มข้อมูลตามลำดับใหม่ โดยใส่ Type เป็น "Customer" และตัดที่อยู่ออก
    sheet.appendRow([newId, customerData.name, "Customer", "'" + customerData.tel]);
    
    _updateTrayBalance_(newId, customerData.name, 0, 0);

    // [CHANGED] ส่งข้อมูลกลับโดยไม่มีที่อยู่
    return { success: true, message: "เพิ่มข้อมูลลูกค้าสำเร็จ", newCustomer: { id: newId, name: customerData.name, tel: customerData.tel }};
  } catch (e) { console.error("addCustomer Error: " + e.message); return { success: false, message: e.message }; }
} 

function updateCustomer(customerData) {
  try {
    const sheet = _getCustomerSheet();
    const result = sheet.getRange("A:A").createTextFinder(customerData.id).findNext();
    if (!result) throw new new Error("ไม่พบข้อมูลลูกค้า");
    
    const targetRow = result.getRow();
    // [CHANGED] อัปเดตเฉพาะชื่อ (คอลัมน์ B) และเบอร์โทร (คอลัมน์ D)
    sheet.getRange(targetRow, 2).setValue(customerData.name);
    sheet.getRange(targetRow, 4).setValue("'" + customerData.tel);
    
    // อัปเดตชื่อใน TrayStock ด้วย
    const traySheet = SpreadsheetApp.openById(CONFIG.trayStock.sheetId).getSheetByName(CONFIG.trayStock.sheetName);
    const trayResult = traySheet.getRange("A:A").createTextFinder(customerData.id).findNext();
    if(trayResult) {
      traySheet.getRange(trayResult.getRow(), 2).setValue(customerData.name);
    }

    return { success: true, message: "แก้ไขข้อมูลสำเร็จ", updatedCustomer: customerData };
  } catch (e) { 
    console.error("updateCustomer Error: " + e.message); 
    return { success: false, message: e.message }; 
  }
}

function deleteCustomer(customerId) {
  // หมายเหตุ: การลบลูกค้าจาก Contacts จะไม่ลบออกจาก TrayStock เพื่อรักษประวัติ
  // หากต้องการลบด้วย ให้เขียนโค้ดเพิ่มเติมในส่วนนี้
  try {
    const sheet = _getCustomerSheet();
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
      // [CHANGED] เพิ่ม Header "รหัสลูกค้า"
      salesSheet.appendRow(['เลขที่เอกสาร', 'วันที่ขาย', 'รหัสลูกค้า', 'ชื่อลูกค้า', 'เบอร์ติดต่อ', 'พนักงานขาย', 'ยอดรวม', 'ส่วนลด', 'ยอดสุทธิ', 'แผงไข่ส่ง', 'แผงไข่รับ', 'ชื่อสินค้า', 'จำนวน', 'หน่วย', 'ราคาต่อหน่วยย่อย', 'ราคารวม', 'จำนวนหน่วยย่อย', 'ผู้บันทึก', 'เวลาที่บันทึก']);
    }
    const recordsToSave = formData.items.map((item, index) => {
      // [CHANGED] เพิ่ม formData.customerId และปรับขนาด Array
      const commonData = index === 0 ? [ invoiceId, timestamp, formData.customerId, formData.customerName, "'" + formData.customerTel, formData.salesperson, formData.subtotal, formData.discount, formData.grandTotal, formData.traysSent, formData.traysReceived ] : Array(11).fill('');
      return [...commonData, item.name, item.quantity, item.unitName, item.price, item.total, item.baseQuantity, currentUser, timestamp];
    });
    if (recordsToSave.length > 0) {
      salesSheet.getRange(salesSheet.getLastRow() + 1, 1, recordsToSave.length, recordsToSave[0].length).setValues(recordsToSave);
      updateStock_(formData.items, 'DEDUCT');

      // [IMPROVED] ใช้ Customer ID โดยตรง ไม่ต้องค้นหาจากชื่ออีกต่อไป
      if (formData.customerId) {
        _updateTrayBalance_(formData.customerId, formData.customerName, formData.traysSent, formData.traysReceived);
      } else {
        console.warn(`Customer ID was missing for "${formData.customerName}". Tray stock not updated.`);
      }

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

    const oldSaleData = getSaleRecordByDocId_(salesSheet, docId);
    
    if (oldSaleData.items.length > 0) {
      updateStock_(oldSaleData.items, 'RETURN');
      // [IMPROVED] ใช้ oldSaleData.customerId โดยตรง
      if (oldSaleData.customerId) {
        _updateTrayBalance_(oldSaleData.customerId, oldSaleData.customerName, -oldSaleData.traysSent, -oldSaleData.traysReceived);
      }
    }
    
    updateStock_(formData.items, 'DEDUCT');
    // [IMPROVED] ใช้ formData.customerId โดยตรง
    if(formData.customerId) {
      _updateTrayBalance_(formData.customerId, formData.customerName, formData.traysSent, formData.traysReceived);
    }

    deleteRowsByDocId_(salesSheet, docId);
    const currentUser = Session.getActiveUser().getEmail() || 'Unknown';
    const timestamp = new Date();
    const recordsToSave = formData.items.map((item, index) => {
      // [CHANGED] เพิ่ม formData.customerId และปรับขนาด Array
      const commonData = index === 0 ? [ docId, timestamp, formData.customerId, formData.customerName, "'" + formData.customerTel, formData.salesperson, formData.subtotal, formData.discount, formData.grandTotal, formData.traysSent, formData.traysReceived ] : Array(11).fill('');
      return [...commonData, item.name, item.quantity, item.unitName, item.price, item.total, item.baseQuantity, currentUser, timestamp];
    });
    if (recordsToSave.length > 0) {
      salesSheet.getRange(salesSheet.getLastRow() + 1, 1, recordsToSave.length, recordsToSave[0].length).setValues(recordsToSave);
    }
    return { success: true, message: `อัปเดตเอกสาร ${docId} สำเร็จ` };
  } catch (e) { console.error("updateSale Error: " + e.message); return { success: false, message: e.message }; }
}


function deleteSaleById(docId) {
  if (!docId) {
    return { success: false, message: "ไม่พบเลขที่เอกสาร" };
  }
  
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const salesSheet = ss.getSheetByName(CONFIG.salesSheetName);
    if (!salesSheet) throw new Error(`ไม่พบชีต '${CONFIG.salesSheetName}'`);
    
    const dataRange = salesSheet.getDataRange();
    const allData = dataRange.getValues();
    const header = allData.shift(); // Tách dòng tiêu đề ra

    let rowsToDeleteIndices = [];
    let saleDataToReturn = { items: [] };
    let isFirstRowOfBill = true;

    // 1. ค้นหาข้อมูลทั้งหมดที่เกี่ยวข้องกับบิล และรวบรวมรายการที่จะคืนสต็อก
    allData.forEach((row, index) => {
      const currentRowDocId = row[0].toString().trim();
      // ค้นหาแถวแรกและแถวลูกของบิลที่ต้องการลบ
      if (currentRowDocId === docId || (rowsToDeleteIndices.length > 0 && currentRowDocId === "")) {
        rowsToDeleteIndices.push(index + 2); // +2 เพราะ data ไม่มี header และ index เริ่มจาก 0

        // ดึงข้อมูลหลักจากแถวแรกของบิลเท่านั้น
        if (isFirstRowOfBill) {
            isFirstRowOfBill = false;
            saleDataToReturn.customerId = row[2];     // คอลัมน์ C
            saleDataToReturn.customerName = row[3];   // คอลัมน์ D
            saleDataToReturn.traysSent = parseInt(row[9]) || 0;     // คอลัมน์ J
            saleDataToReturn.traysReceived = parseInt(row[10]) || 0; // คอลัมน์ K
        }

        // ดึงข้อมูลสินค้า (Item) จากทุกแถวของบิล
        const productName = row[11]; // คอลัมน์ L
        const baseQuantity = parseFloat(row[16]); // คอลัมน์ Q
        if (productName && !isNaN(baseQuantity)) {
          saleDataToReturn.items.push({ name: productName, baseQuantity: baseQuantity });
        }
      }
    });

    // 2. ถ้าเจอข้อมูล ให้ทำการคืนสต็อกทั้งหมดก่อน
    if (rowsToDeleteIndices.length > 0) {
      // 2.1 คืนสต็อกสินค้า (ไข่) กลับเข้าชีต "คลัง"
      if (saleDataToReturn.items.length > 0) {
        updateStock_(saleDataToReturn.items, 'RETURN');
      }

      // 2.2 คืนสต็อกแผงไข่ กลับเข้าชีต "TrayStock"
      if (saleDataToReturn.customerId) {
        // ใช้ - เพื่อทำการย้อนรายการ (ส่งกลายเป็นลบ, รับกลายเป็นบวก)
        _updateTrayBalance_(saleDataToReturn.customerId, saleDataToReturn.customerName, -saleDataToReturn.traysSent, -saleDataToReturn.traysReceived);
      }

      // 3. ทำการลบแถวทั้งหมดทีเดียว (เริ่มลบจากล่างขึ้นบนเสมอ เพื่อป้องกัน index เพี้ยน)
      for (let i = rowsToDeleteIndices.length - 1; i >= 0; i--) {
        salesSheet.deleteRow(rowsToDeleteIndices[i]);
      }

      return { success: true, message: `ลบเอกสาร ${docId} และคืนสต็อกสำเร็จ` };
    } else {
      return { success: false, message: `ไม่พบเอกสาร ${docId} ในระบบ` };
    }

  } catch (e) {
    console.error("deleteSaleById Error for docId " + docId + ": " + e.message);
    return { success: false, message: "เกิดข้อผิดพลาดขณะลบ: " + e.message };
  }
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
        groupedSales[docId] = { 
            docId: docId, 
            date: Utilities.formatDate(new Date(row[1]), timezone, 'dd/MM/yyyy'), 
            customerId: row[2],
            customerName: row[3], 
            customerTel: row[4], 
            salesperson: row[5], 
            subtotal: row[6], 
            discount: row[7], 
            grandTotal: row[8],
            traysSent: row[9] || 0,
            traysReceived: row[10] || 0,
            items: [] 
        };
      }
      if (groupedSales[docId]) {
        // [IMPROVED] เพิ่ม baseQuantity (row[16]) เข้ามาใน object ของ item
        groupedSales[docId].items.push({ 
            name: row[11], 
            quantity: row[12], 
            unitName: row[13], 
            price: row[14], 
            total: row[15],
            baseQuantity: row[16] || 0
        });
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
/**
 * [UPDATED] ฟังก์ชันสำหรับอัปเดตยอดคงเหลือแผงไข่ของลูกค้า (เพิ่มความทนทานต่อ Error)
 * @param {string} customerId - ID ของลูกค้า (เช่น CUS-001)
 * @param {string} customerName - ชื่อของลูกค้า
 * @param {number} traysSent - จำนวนแผงที่ส่ง (เป็น +)
 * @param {number} traysReceived - จำนวนแผงที่รับคืน (เป็น -)
 */
function _updateTrayBalance_(customerId, customerName, traysSent, traysReceived) {
  try {
    const ss = SpreadsheetApp.openById(CONFIG.trayStock.sheetId);
    const sheet = ss.getSheetByName(CONFIG.trayStock.sheetName);
    if (!sheet) throw new Error(`ไม่พบชีต '${CONFIG.trayStock.sheetName}'`);

    if (sheet.getLastRow() === 0) {
      sheet.appendRow(['ContactID', 'ContactName', 'TrayBalance']);
    }

    // ใช้ TextFinder เพื่อการค้นหาที่แม่นยำและรวดเร็วกว่า
    const finder = sheet.getRange("A:A").createTextFinder(customerId).findNext();
    const netChange = (Number(traysSent) || 0) - (Number(traysReceived) || 0);

    if (finder) {
      // ลูกค้ามีอยู่แล้ว: อัปเดตยอดคงเหลือ
      const balanceCell = sheet.getRange(finder.getRow(), 3);
      // [IMPROVED] ทำให้การแปลงค่าเป็นตัวเลขมีความปลอดภัยสูงสุด
      const currentBalance = Number(balanceCell.getValue()) || 0;
      balanceCell.setValue(currentBalance + netChange);
    } else {
      // ลูกค้าใหม่: เพิ่มแถวใหม่
      sheet.appendRow([customerId, customerName, netChange]);
    }
  } catch (e) {
    console.error(`_updateTrayBalance_ Error for CUS_ID ${customerId}: ${e.message}`);
  }
}

/**
 * [NEW] Helper function to find a customer's ID by their name.
 * @param {string} name - The name of the customer to find.
 * @returns {string|null} The customer ID or null if not found.
 */
function _findCustomerIdByName_(name) {
  try {
    const sheet = _getCustomerSheet();
    if (sheet.getLastRow() < 2) return null;
    const names = sheet.getRange(2, 2, sheet.getLastRow() - 1, 1).getValues().flat();
    const ids = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues().flat();
    const customerIndex = names.findIndex(n => n.trim() === name.trim());
    return customerIndex !== -1 ? ids[customerIndex] : null;
  } catch (e) {
    console.error("_findCustomerIdByName_ Error: " + e.message);
    return null;
  }
}

/**
 * ดึงข้อมูลการขายทั้งหมดที่เกี่ยวข้องกับ docId หนึ่งๆ รวมถึงสินค้าทุกรายการ
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - ชีตข้อมูลการขาย
 * @param {string} docId - เลขที่เอกสารที่ต้องการค้นหา
 * @returns {object} ออบเจ็กต์ข้อมูลการขายที่สมบูรณ์
 */
function getSaleRecordByDocId_(sheet, docId) {
  if (!sheet || !docId) return { items: [] }; // คืนค่า object ที่มี items เสมอ
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return { items: [] };

  const data = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();
  let saleData = { items: [] };
  let isCapturing = false; // ใช้ตัวแปรเพื่อบอกสถานะว่า "กำลังอ่านข้อมูลของบิลนี้อยู่"

  for (const row of data) {
    const currentRowDocId = row[0].toString().trim();

    // หากเจอ ID ที่ตรงกัน ให้เริ่มกระบวนการอ่านข้อมูล
    if (currentRowDocId === docId) {
      isCapturing = true;
      // ดึงข้อมูลหลักจากแถวแรกของบิล
      saleData.customerId = row[2];
      saleData.customerName = row[3];
      saleData.traysSent = parseInt(row[9]) || 0;
      saleData.traysReceived = parseInt(row[10]) || 0;
    }

    // ตราบใดที่ยังอยู่ในโหมด "กำลังอ่าน" ให้เก็บข้อมูลสินค้าไปเรื่อยๆ
    if (isCapturing) {
      const productName = row[11]; // คอลัมน์ L: ชื่อสินค้า
      const baseQuantity = parseFloat(row[16]); // คอลัมน์ Q: จำนวนหน่วยย่อย

      if (productName && !isNaN(baseQuantity) && baseQuantity > 0) {
        saleData.items.push({ name: productName, baseQuantity: baseQuantity });
      }

      // ตรวจสอบแถวถัดไป: ถ้าแถวถัดไปไม่มีอยู่แล้ว หรือมี ID ใหม่, ให้หยุดการอ่าน
      const currentRowIndex = data.indexOf(row);
      const nextRow = data[currentRowIndex + 1];
      if (!nextRow || nextRow[0].toString().trim() !== "") {
        break; // หยุด Loop ทันที
      }
    }
  }
  return saleData;
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

/**
 * สร้าง ID ใหม่โดยค้นหาเลขล่าสุดจาก Prefix ที่ระบุ
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - ชีตที่ต้องการค้นหา
 * @param {string} prefix - คำนำหน้า ID ที่ต้องการ (เช่น "CUS", "SUP", "BR")
 * @returns {string} ID ใหม่ที่สร้างขึ้น
 */
function _generateNextId_(sheet, prefix) {
  const PADDING_LENGTH = 3; // กำหนดจำนวนหลักของตัวเลข เช่น 3 คือ 001, 002

  if (!sheet || sheet.getLastRow() < 2) {
    return `${prefix}-${'1'.padStart(PADDING_LENGTH, '0')}`;
  }
  const allIds = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues().flat();

  let maxNum = 0;
  allIds.forEach(id => {
    const idString = id.toString();
    if (idString.startsWith(`${prefix}-`)) {
      const match = idString.match(/(\d+)$/);
      if (match) {
        const num = parseInt(match[1], 10);
        if (num > maxNum) {
          maxNum = num;
        }
      }
    }
  });

  const nextNum = maxNum + 1;
  return `${prefix}-${String(nextNum).padStart(PADDING_LENGTH, '0')}`;
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

const ssCache = {};

function getSpreadsheet_(ssId) {
  if (!ssCache[ssId]) {
    ssCache[ssId] = SpreadsheetApp.openById(ssId);
  }
  return ssCache[ssId];
}

// เวลาใช้งาน
const stockSs = getSpreadsheet_(CONFIG.stockSheetId);

/**
 * [NEW] ดึงข้อมูลสรุปยอดบิลสำหรับหน้า Dashboard
 * @returns {object} ออบเจ็กต์ที่ประกอบด้วยยอดบิล วันนี้, เดือนนี้, และทั้งหมด
 */
function getDashboardMetrics() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const salesSheet = ss.getSheetByName(CONFIG.salesSheetName);
    if (!salesSheet || salesSheet.getLastRow() < 2) {
      return { today: 0, month: 0, total: 0 };
    }

    // ดึงข้อมูลแค่คอลัมน์ A (เลขที่เอกสาร) และ B (วันที่ขาย) เพื่อประสิทธิภาพ
    const data = salesSheet.getRange(2, 1, salesSheet.getLastRow() - 1, 2).getValues();
    const uniqueDocs = new Set();
    let todayCount = 0;
    let monthCount = 0;

    const now = new Date();
    const today = now.getDate();
    const thisMonth = now.getMonth();
    const thisYear = now.getFullYear();

    data.forEach(row => {
      const docId = row[0];
      if (docId) { // ประมวลผลเฉพาะแถวแรกของบิลที่มีเลขที่เอกสาร
        uniqueDocs.add(docId);
        const saleDate = new Date(row[1]);
        
        // ตรวจสอบว่าเป็นบิลของวันนี้หรือไม่
        if (saleDate.getDate() === today && saleDate.getMonth() === thisMonth && saleDate.getFullYear() === thisYear) {
          todayCount++;
        }
        // ตรวจสอบว่าเป็นบิลของเดือนนี้หรือไม่
        if (saleDate.getMonth() === thisMonth && saleDate.getFullYear() === thisYear) {
          monthCount++;
        }
      }
    });

    return {
      today: todayCount,
      month: monthCount,
      total: uniqueDocs.size
    };
  } catch (e) {
    console.error("getDashboardMetrics Error: " + e.message);
    return { today: 0, month: 0, total: 0, error: e.message };
  }
}
