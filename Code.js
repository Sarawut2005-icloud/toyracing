// ==========================================
// ⚙️ CONFIGURATION & SETTINGS
// ==========================================
const SPREADSHEET_ID = "1ZwDgSlOZpU64WjI8f1fGqgsLBq8Ftd2DyJqNPuBfcQU";
const LINE_TOKEN = "NHetPh5dKYYydjK96riOww1TgUZ7b8wTSiB1stAw3FO3XS60O6u5gBinrioCm67NEywy4bX8GBIiHSDXEXpinQYRFxTAdslahCO6j6eYp883mql8noMfQkgQ5CZVbFvCZ2CStqQ+jQBDbGGiiyjgPQdB04t89/1O/w1cDnyilFU=";
const SAFETY_LIMIT = 2; // จำนวนสินค้าขั้นต่ำในการแจ้งเตือน

function doGet(e) {
  initSheetsIfNotExist();
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('Toy Racing Pro - Stock Management System (Light UI)')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ------------------------------------------
// 📊 API ดึงข้อมูลทั้งหมดเข้า WEB APP
// ------------------------------------------
function getDashboardData() {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const masterSheet = ss.getSheetByName('สินค้าทั้งหมด');
    const logSheet = ss.getSheetByName('Stock Log');

    const masterValues = masterSheet.getDataRange().getValues();
    if (masterValues.length > 0) masterValues.shift(); // ลบ Header

    const logValues = logSheet.getDataRange().getValues();
    if (logValues.length > 0) logValues.shift(); // ลบ Header

    const stockMap = {};
    let totalInCount = 0;
    let totalOutCount = 0;

    logValues.forEach(row => {
      const type = row[1];
      const sku = String(row[2]).trim();
      const qty = Number(row[4]) || 0;

      if (!sku) return;
      if (!stockMap[sku]) stockMap[sku] = { in: 0, out: 0 };

      if (type === 'รับเข้า') {
        stockMap[sku].in += qty;
        totalInCount += qty;
      } else if (type === 'จ่ายออก') {
        stockMap[sku].out += qty;
        totalOutCount += qty;
      }
    });

    let totalStockRemaining = 0;
    let totalStockValue = 0;
    let lowStockCount = 0;
    let lowStockList = [];

    const items = masterValues.map((row, index) => {
      const sku = String(row[2]).trim();
      const name = row[3] || '-';
      const price = Number(row[4]) || 0;
      const initialStock = Number(row[5]) || 0;

      const inQty = initialStock + (stockMap[sku] ? stockMap[sku].in : 0);
      const outQty = stockMap[sku] ? stockMap[sku].out : 0;
      const remaining = inQty - outQty;
      const itemValue = remaining * price;

      totalStockRemaining += remaining;
      totalStockValue += itemValue;

      let status = 'ปกติ';
      if (remaining <= 0) {
        status = 'หมดสต๊อก';
        lowStockCount++;
        lowStockList.push({ id: index + 1, sku, name, remaining, status });
      } else if (remaining <= SAFETY_LIMIT) {
        status = 'ใกล้หมด';
        lowStockCount++;
        lowStockList.push({ id: index + 1, sku, name, remaining, status });
      }

      return {
        id: index + 1,
        category: row[1] || 'ทั่วไป',
        sku: sku,
        name: name,
        price: price,
        inQty: inQty,
        outQty: outQty,
        remaining: remaining,
        status: status
      };
    });

    const allLogs = logValues.slice().reverse().map(row => ({
      date: row[0] ? Utilities.formatDate(new Date(row[0]), "GMT+7", "dd/MM/yyyy HH:mm") : "-",
      type: row[1],
      sku: row[2],
      note: row[3] || "-",
      qty: row[4],
      user: row[5] || "Admin"
    }));

    return {
      kpi: {
        totalIn: totalInCount,
        totalOut: totalOutCount,
        totalRemaining: totalStockRemaining,
        totalItems: items.length,
        totalValue: totalStockValue,
        lowStockCount: lowStockCount
      },
      items: items,
      allLogs: allLogs,
      lowStockList: lowStockList
    };
  } catch (err) {
    Logger.log("Error in getDashboardData: " + err.toString());
    throw new Error("ดึงข้อมูลไม่สำเร็จ: " + err.message);
  }
}

// ------------------------------------------
// 💾 API บันทึกข้อมูลเข้า/ออก & สินค้าใหม่
// ------------------------------------------
function saveTransaction(data) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const logSheet = ss.getSheetByName('Stock Log');
  const now = new Date();
  const qty = Number(data.qty);
  
  logSheet.appendRow([
    now,
    data.type,
    data.sku,
    data.note || (data.type === 'รับเข้า' ? 'นำสินค้าเข้าสต๊อก' : 'จ่ายสินค้าออก'),
    qty,
    data.user || 'Admin'
  ]);

  const icon = data.type === 'รับเข้า' ? '📥' : '📤';
  const lineMsg = `${icon} [รายการ${data.type}ใหม่]\n🏎️ Toy Racing\n• รหัสสินค้า: ${data.sku}\n• จำนวน: ${qty} ชิ้น\n• หมายเหตุ: ${data.note || '-'}\n• ผู้บันทึก: ${data.user || 'Admin'}\n⏱️ เวลา: ${Utilities.formatDate(now, "GMT+7", "HH:mm น.")}`;
  sendLineNotification(lineMsg);

  return { status: 'success' };
}

function addNewProduct(data) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const masterSheet = ss.getSheetByName('สินค้าทั้งหมด');
  const lastRow = masterSheet.getLastRow();
  
  masterSheet.appendRow([
    lastRow,
    data.category || 'อะไหล่ทั่วไป',
    data.sku,
    data.name,
    Number(data.price) || 0,
    Number(data.initialStock) || 0
  ]);
  
  return { status: 'success' };
}

// ------------------------------------------
// 🔔 LINE NOTIFICATION SYSTEM
// ------------------------------------------
function sendLineNotification(textMessage) {
  if (!LINE_TOKEN) return;
  try {
    UrlFetchApp.fetch("https://api.line.me/v2/bot/message/broadcast", {
      method: "post",
      headers: {
        "Authorization": "Bearer " + LINE_TOKEN,
        "Content-Type": "application/json"
      },
      payload: JSON.stringify({ messages: [{ type: "text", text: textMessage }] }),
      muteHttpExceptions: true
    });
  } catch (e) {
    Logger.log("LINE Send Error: " + e.toString());
  }
}

function initSheetsIfNotExist() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  if (!ss.getSheetByName('สินค้าทั้งหมด')) {
    const master = ss.insertSheet('สินค้าทั้งหมด');
    master.appendRow(["ลำดับ", "หมวดหมู่", "รหัสสินค้า", "ชื่อรุ่น / รายละเอียดสินค้า", "ราคาต้นทุน (บาท)", "สต๊อกตั้งต้น"]);
  }
  if (!ss.getSheetByName('Stock Log')) {
    const log = ss.insertSheet('Stock Log');
    log.appendRow(["วันที่-เวลา", "ประเภท", "รหัสสินค้า", "รายการ", "จำนวน", "ผู้บันทึก"]);
  }
}