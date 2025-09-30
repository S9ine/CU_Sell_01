// ===================================================
// === 1. การตั้งค่าและฟังก์ชันหลัก (Configuration & Core Functions) ===
// ===================================================
const CONFIG = {
  salesSheetName: "ข้อมูลการขาย",
  logSheetName: "Log",
  customer: {
    sheetId: "19MvkCOZfUuQKjaeCYHKV5UTgSv-09PqpgIiTbX6qKWk",
    sheetName: "Contacts"
  },
  trayStock: {
    sheetId: "19MvkCOZfUuQKjaeCYHKV5UTgSv-09PqpgIiTbX6qKWk",
    sheetName: "TrayStock"
  },
  webAppInfo: { configSheet: "config", documentInfoSheet: "ข้อมูลเอกสาร", settingsSheet: "Settings" },
  stockSheetId: "19MvkCOZfUuQKjaeCYHKV5UTgSv-09PqpgIiTbX6qKWk",
  masterStockSheetName: "คลัง"
};

/**
 * [IMPROVED] เพิ่มรายละเอียด Log ให้มากขึ้น
 * @param {string} emoji - Emoji for the log entry.
 * @param {string} details - The detailed log message.
 * @param {string} context - Optional context like DocID or Customer Name.
 */
function _logActivity_(emoji, details, context = '') {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const logSheet = ss.getSheetByName(CONFIG.logSheetName) || ss.insertSheet(CONFIG.logSheetName);

    if (logSheet.getLastRow() === 0) {
      logSheet.appendRow(['Timestamp', 'User', 'Activity', 'Context']);
      logSheet.setFrozenRows(1);
      logSheet.getRange("A:A").setNumberFormat("yyyy-mm-dd hh:mm:ss");
      logSheet.getRange("C:C").setWrap(true);
      logSheet.setColumnWidth(1, 150);
      logSheet.setColumnWidth(2, 200);
      logSheet.setColumnWidth(3, 500);
      logSheet.setColumnWidth(4, 200);
    }

    const user = Session.getActiveUser().getEmail() || 'Unknown';
    const timestamp = new Date();
    logSheet.appendRow([timestamp, user, `${emoji} ${details}`, context]);
  } catch (e) {
    console.error("Failed to log activity: " + e.message);
  }
}


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
      template.metrics = getDashboardMetrics();
      return template.evaluate()
        .setTitle("Dashboard | ระบบขาย")
        .setSandboxMode(HtmlService.SandboxMode.IFRAME);
    }
  } else {
    return HtmlService.createHtmlOutputFromFile('AccessDenied').setTitle("Access Denied");
  }
}

function include(filename) { return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

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
  } catch (e) { console.error("checkUserAccess_ Error: " + e.toString()); return false;
  }
}

// ===================================================
// === 2. ฟังก์ชันจัดการข้อมูลลูกค้า (Customer CRUD) ===
// ===================================================

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
    const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 4).getValues();
    return data.map(row => ({ id: row[0], name: row[1], tel: row[3] }));
  } catch (e) { console.error("getCustomers Error: " + e.message); return [];
  }
}

function addCustomer(customerData) {
  try {
    const sheet = _getCustomerSheet();
    if (sheet.getLastRow() === 0) { 
      sheet.appendRow(['ContactID', 'ContactName', 'Type', 'Phone']);
    }
    const newId = _generateNextId_(sheet, 'CUS');
    sheet.appendRow([newId, customerData.name, "Customer", "'" + customerData.tel]);
    _updateTrayBalance_(newId, customerData.name, 0, 0);
    _logActivity_('👤', `เพิ่มลูกค้าใหม่: "${customerData.name}" (ID: ${newId})`);
    return { success: true, message: "เพิ่มข้อมูลลูกค้าสำเร็จ", newCustomer: { id: newId, name: customerData.name, tel: customerData.tel }};
  } catch (e) { console.error("addCustomer Error: " + e.message); return { success: false, message: e.message };
  }
} 

function updateCustomer(customerData) {
  try {
    const sheet = _getCustomerSheet();
    const result = sheet.getRange("A:A").createTextFinder(customerData.id).findNext();
    if (!result) throw new new Error("ไม่พบข้อมูลลูกค้า");
    
    const targetRow = result.getRow();
    sheet.getRange(targetRow, 2).setValue(customerData.name);
    sheet.getRange(targetRow, 4).setValue("'" + customerData.tel);
    
    const traySheet = SpreadsheetApp.openById(CONFIG.trayStock.sheetId).getSheetByName(CONFIG.trayStock.sheetName);
    const trayResult = traySheet.getRange("A:A").createTextFinder(customerData.id).findNext();
    if(trayResult) {
      traySheet.getRange(trayResult.getRow(), 2).setValue(customerData.name);
    }
    
    _logActivity_('✏️', `แก้ไขข้อมูลลูกค้า "${customerData.name}" (ID: ${customerData.id})`);
    return { success: true, message: "แก้ไขข้อมูลสำเร็จ", updatedCustomer: customerData };
  } catch (e) { 
    console.error("updateCustomer Error: " + e.message); 
    return { success: false, message: e.message };
  }
}

function deleteCustomer(customerId) {
  try {
    const sheet = _getCustomerSheet();
    const result = sheet.getRange("A:A").createTextFinder(customerId).findNext();
    if (!result) throw new Error("ไม่พบข้อมูลลูกค้า");
    const customerName = sheet.getRange(result.getRow(), 2).getValue();
    sheet.deleteRow(result.getRow());
    _logActivity_('🗑️', `ลบข้อมูลลูกค้า "${customerName}" (ID: ${customerId})`);
    return { success: true, message: "ลบข้อมูลลูกค้าสำเร็จ" };
  } catch (e) { console.error("deleteCustomer Error: " + e.message); return { success: false, message: e.message };
  }
}

// ===================================================
// === 3. ฟังก์ชันจัดการข้อมูลการขาย (Sales Functions) ===
// ===================================================
function saveSalesData(formData) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000); // รอ Lock สูงสุด 20 วินาที

  try {
    // [NEW] ตรวจสอบสต็อกใน Backend อีกครั้งเพื่อความปลอดภัย
    const stockCheck = _validateStockAvailability(formData.items);
    if (!stockCheck.isValid) {
      throw new Error(`สต็อกไม่พอสำหรับ: ${stockCheck.insufficientItems.join(', ')}`);
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const salesSheet = ss.getSheetByName(CONFIG.salesSheetName) || ss.insertSheet(CONFIG.salesSheetName);
    const currentUser = Session.getActiveUser().getEmail() || 'Unknown';
    const timestamp = new Date();
    const invoiceId = generateInvoiceId_(salesSheet);

    if (salesSheet.getLastRow() === 0) {
      salesSheet.appendRow(['เลขที่เอกสาร', 'วันที่ขาย', 'รหัสลูกค้า', 'ชื่อลูกค้า', 'เบอร์ติดต่อ', 'พนักงานขาย', 'ยอดรวม', 'ส่วนลด', 'ยอดสุทธิ', 'แผงไข่ส่ง', 'แผงไข่รับ', 'ชื่อสินค้า', 'จำนวน', 'หน่วย', 'ราคาต่อหน่วยย่อย', 'ราคารวม', 'จำนวนหน่วยย่อย', 'ผู้บันทึก', 'เวลาที่บันทึก']);
    }

    const recordsToSave = formData.items.map((item, index) => {
      const commonData = index === 0 ? [ invoiceId, timestamp, formData.customerId, formData.customerName, "'" + formData.customerTel, formData.salesperson, formData.subtotal, formData.discount, formData.grandTotal, formData.traysSent, formData.traysReceived ] : Array(11).fill('');
      return [...commonData, item.name, item.quantity, item.unitName, item.price, item.total, item.baseQuantity, currentUser, timestamp];
    });

    if (recordsToSave.length > 0) {
      salesSheet.getRange(salesSheet.getLastRow() + 1, 1, recordsToSave.length, recordsToSave[0].length).setValues(recordsToSave);
      
      // ทำ Operation สำคัญหลังจากบันทึกชีตสำเร็จ
      updateStock_(formData.items, 'DEDUCT');
      if (formData.customerId) {
        _updateTrayBalance_(formData.customerId, formData.customerName, formData.traysSent, formData.traysReceived);
      }
      
      _logActivity_('🧾', `สร้างเอกสารขาย #${invoiceId} (${formData.items.length} รายการ)`, `ลูกค้า: ${formData.customerName}`);
      return { success: true, docId: invoiceId, message: "บันทึกข้อมูลการขายสำเร็จ" };
    } else {
      throw new Error("ไม่พบรายการสินค้าที่จะบันทึก");
    }
  } catch (e) {
    console.error("saveSalesData Error: " + e.message, e.stack);
    return { success: false, message: `เกิดข้อผิดพลาด: ${e.message}` };
  } finally {
    lock.releaseLock(); // ปลด Lock ทุกครั้ง
  }
}

/**
 * [NEW] ฟังก์ชันสำหรับตรวจสอบสต็อกใน Backend
 * @param {Array} items - รายการสินค้าที่ต้องการตรวจสอบ
 * @returns {{isValid: boolean, insufficientItems: Array<string>}}
 */
function _validateStockAvailability(items) {
  try {
    const { productList } = _fetchAndProcessStockData();
    const stockMap = new Map(productList.map(p => [p.productName, p.stockCentral]));
    let insufficientItems = [];

    items.forEach(item => {
      const availableStock = stockMap.get(item.name.trim()) || 0;
      if (item.baseQuantity > availableStock) {
        insufficientItems.push(`${item.name} (ต้องการ: ${item.baseQuantity}, มี: ${availableStock})`);
      }
    });

    return {
      isValid: insufficientItems.length === 0,
      insufficientItems: insufficientItems
    };
  } catch (e) {
    console.error("Stock Validation Error:", e.message);
    // กรณีฉุกเฉิน ให้ถือว่าไม่ผ่านเพื่อความปลอดภัย
    return { isValid: false, insufficientItems: ["เกิดข้อผิดพลาดในการตรวจสอบสต็อก"] };
  }
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
      if (oldSaleData.customerId) {
        _updateTrayBalance_(oldSaleData.customerId, oldSaleData.customerName, -oldSaleData.traysSent, -oldSaleData.traysReceived);
      }
    }
    
    updateStock_(formData.items, 'DEDUCT');
    if(formData.customerId) {
      _updateTrayBalance_(formData.customerId, formData.customerName, formData.traysSent, formData.traysReceived);
    }

    deleteRowsByDocId_(salesSheet, docId);
    const currentUser = Session.getActiveUser().getEmail() || 'Unknown';
    const timestamp = new Date();
    const recordsToSave = formData.items.map((item, index) => {
      const commonData = index === 0 ? [ docId, timestamp, formData.customerId, formData.customerName, "'" + formData.customerTel, formData.salesperson, formData.subtotal, formData.discount, formData.grandTotal, formData.traysSent, formData.traysReceived ] : Array(11).fill('');
      return [...commonData, item.name, item.quantity, item.unitName, item.price, item.total, item.baseQuantity, currentUser, timestamp];
    });
    if (recordsToSave.length > 0) {
      salesSheet.getRange(salesSheet.getLastRow() + 1, 1, recordsToSave.length, recordsToSave[0].length).setValues(recordsToSave);
    }
    _logActivity_('✏️', `แก้ไขเอกสารขาย #${docId} ของ "${formData.customerName}" (${formData.items.length} รายการ)`);
    return { success: true, message: `อัปเดตเอกสาร ${docId} สำเร็จ` };
  } catch (e) { console.error("updateSale Error: " + e.message); return { success: false, message: e.message };
  }
}


function deleteSaleById(docId) {
  if (!docId) {
    return { success: false, message: "ไม่พบเลขที่เอกสาร" };
  }
  
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const salesSheet = ss.getSheetByName(CONFIG.salesSheetName);
    if (!salesSheet) throw new Error(`ไม่พบชีต '${CONFIG.salesSheetName}'`);
    
    // ดึงข้อมูลเก่าเพื่อใช้คืนสต็อก (Refactored)
    const oldSaleData = getSaleRecordByDocId_(salesSheet, docId);
    if (oldSaleData.items.length === 0) {
      return { success: false, message: `ไม่พบข้อมูลเอกสาร ${docId}` };
    }
    
    // 1. คืนสต็อกและแผงไข่ก่อน
    if (oldSaleData.items.length > 0) {
      updateStock_(oldSaleData.items, 'RETURN');
    }
    if (oldSaleData.customerId) {
      _updateTrayBalance_(oldSaleData.customerId, oldSaleData.customerName, -oldSaleData.traysSent, -oldSaleData.traysReceived);
    }

    // 2. ลบแถวในชีต
    deleteRowsByDocId_(salesSheet, docId);
      
    const logDetails = `ลบเอกสารขาย #${docId} และคืน ${oldSaleData.items.length} รายการสู่สต็อก`;
    _logActivity_('🗑️', logDetails, `ลูกค้า: ${oldSaleData.customerName}`);
    return { success: true, message: `ลบเอกสาร ${docId} และคืนสต็อกสำเร็จ` };

  } catch (e) {
    console.error("deleteSaleById Error for docId " + docId + ": " + e.message, e.stack);
    return { success: false, message: "เกิดข้อผิดพลาดขณะลบ: " + e.message };
  } finally {
    lock.releaseLock();
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
  } catch (e) { console.error("getSalesHistory Error: " + e.message); return [];
  }
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
        } catch (e) 
        {
            console.error("getEmployeeList Error: " + e.message); return [];
        }
    });
}

// ===================================================
// === 6. ฟังก์ชันจัดการเอกสารและ ID (Utility Functions) ===
// ===================================================
function _updateTrayBalance_(customerId, customerName, traysSent, traysReceived) {
  try {
    const ss = SpreadsheetApp.openById(CONFIG.trayStock.sheetId);
    const sheet = ss.getSheetByName(CONFIG.trayStock.sheetName);
    if (!sheet) throw new Error(`ไม่พบชีต '${CONFIG.trayStock.sheetName}'`);
    if (sheet.getLastRow() === 0) {
      sheet.appendRow(['ContactID', 'ContactName', 'TrayBalance']);
    }

    const finder = sheet.getRange("A:A").createTextFinder(customerId).findNext();
    const netChange = (Number(traysSent) || 0) - (Number(traysReceived) || 0);
    
    if (netChange === 0 && !finder) return; // ไม่ต้องทำอะไรถ้าไม่มีการเปลี่ยนแปลงและเป็นลูกค้าใหม่

    let currentBalance = 0;
    let newBalance = 0;
    
    if (finder) {
      const balanceCell = sheet.getRange(finder.getRow(), 3);
      currentBalance = Number(balanceCell.getValue()) || 0;
      newBalance = currentBalance + netChange;
      balanceCell.setValue(newBalance);
    } else {
      newBalance = netChange;
      sheet.appendRow([customerId, customerName, newBalance]);
    }
    
    if (netChange !== 0) {
        _logActivity_('📦', `อัปเดตยอดแผงไข่ของ "${customerName}" จำนวน ${netChange > 0 ? '+' + netChange : netChange} แผง (ยอดใหม่: ${newBalance})`);
    }

  } catch (e) {
    console.error(`_updateTrayBalance_ Error for CUS_ID ${customerId}: ${e.message}`);
  }
}

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

function getSaleRecordByDocId_(sheet, docId) {
  if (!sheet || !docId) return { items: [] };
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return { items: [] };
  const data = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();
  let saleData = { items: [] };
  let isCapturing = false;

  for (const row of data) {
    const currentRowDocId = row[0].toString().trim();
    if (currentRowDocId === docId) {
      isCapturing = true;
      saleData.customerId = row[2];
      saleData.customerName = row[3];
      saleData.traysSent = parseInt(row[9]) || 0;
      saleData.traysReceived = parseInt(row[10]) || 0;
    }

    if (isCapturing) {
      const productName = row[11];
      const baseQuantity = parseFloat(row[16]);

      if (productName && !isNaN(baseQuantity) && baseQuantity > 0) {
        saleData.items.push({ name: productName, baseQuantity: baseQuantity });
      }

      const currentRowIndex = data.indexOf(row);
      const nextRow = data[currentRowIndex + 1];
      if (!nextRow || nextRow[0].toString().trim() !== "") {
        break;
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
  } catch (e) { return { companyName: "ชื่อบริษัทของคุณ", address1: "ที่อยู่ 1", address2: "ที่อยู่ 2", contactInfo: "ข้อมูลติดต่อ" };
  }
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
     _logActivity_('⚙️', `อัปเดตข้อมูลเอกสาร (ชื่อบริษัท: ${settingsData.companyName})`);
    return { success: true, message: "บันทึกการตั้งค่าสำเร็จ!" };
  } catch (e) { return { success: false, message: e.message };
  }
}

function _generateNextId_(sheet, prefix) {
  const PADDING_LENGTH = 3;

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
    if (!docIdColumnValues[i]) { rowCount++;
    } else { break; }
  }
  sheet.deleteRows(firstIndex + 2, rowCount);
}

function clearServerCache() {
  try {
    CacheService.getScriptCache().removeAll(['productList', 'employeeList']);
    _logActivity_('🧹', 'ทำการล้างแคชของเซิร์ฟเวอร์');
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

const stockSs = getSpreadsheet_(CONFIG.stockSheetId);

function getDashboardMetrics() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const salesSheet = ss.getSheetByName(CONFIG.salesSheetName);
    if (!salesSheet || salesSheet.getLastRow() < 2) {
      return { today: 0, month: 0, total: 0 };
    }

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
      if (docId) { 
        uniqueDocs.add(docId);
        const saleDate = new Date(row[1]);
        
        if (saleDate.getDate() === today && saleDate.getMonth() === thisMonth && saleDate.getFullYear() === thisYear) {
          todayCount++;
        }
    
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
