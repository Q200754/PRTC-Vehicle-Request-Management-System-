// ==========================================
// PRTC Vehicle Request Management System (PRTC-VRMS)
// File: Code.gs
// ==========================================

const SPREADSHEET_ID = '1mfuOdQQG6ryphI7zme6mnlUGADMnSEMmIMqkMadtfJo'; 
const SHEET_NAME = 'คำร้องขอใช้รถ';

const LINE_CHANNEL_TOKEN = '3AEyfuGANlWSyi9KT46RJ5SsHyTm//eUK8MZKNEvpFw8NL+zpNxpGfLWW7OJ3LRjAs+1TC0T1nb1retUL9HqjAr5DKRn65BcJ6oKAvC7UGMIv+O7cr7QStFvV5aypAIXlaYSjxaRl8E7BvWlrq3dfgdB04t89/1O/w1cDnyilFU='; 
const LINE_TARGET_ID = 'U56af96f6772c80456d03df04fc84d3bf'; 

function getTargetSheet() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  if (!ss) throw new Error("ไม่สามารถเปิดไฟล์ Google Sheet ด้วย ID ที่ระบุได้");
  const sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) throw new Error(`ไม่พบ Tab ชีตที่ชื่อ '${SHEET_NAME}' ในไฟล์ Google Sheets`);
  return sheet;
}

// 1. ฟังก์ชัน doGet สำหรับแสดงหน้าเว็บและส่งข้อมูลผ่าน API
function doGet(e) {
  // รองรับการดึงข้อมูลตารางผ่าน fetch (สำหรับ Vercel / Client)
  if (e && e.parameter && e.parameter.action === 'getRequests') {
    const data = getRequests();
    return ContentService
      .createTextOutput(JSON.stringify(data))
      .setMimeType(ContentService.MimeType.JSON);
  }

  // รองรับการเปิดหน้าปกติใน Apps Script
  const page = e && e.parameter.p ? e.parameter.p.toLowerCase() : 'index';
  let htmlName = 'index';
  
  if (page === 'admin') htmlName = 'Admin';
  else if (page === 'driver') htmlName = 'Driver';
  
  return HtmlService.createTemplateFromFile(htmlName)
    .evaluate()
    .setTitle('PRTC-VRMS | ' + htmlName)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// 2. บันทึกคำร้องขอใช้รถใหม่
function submitRequest(data) {
  try {
    const sheet = getTargetSheet();
    const timestamp = new Date();
    const dateStr = Utilities.formatDate(timestamp, "GMT+7", "yyyyMMdd");
    const randomNum = Math.floor(1000 + Math.random() * 9000);
    const requestId = `REQ-${dateStr}-${randomNum}`;
    
    const rowData = [
      timestamp, requestId, data.fullname, data.department, data.objective,
      data.startDate, data.startTime, data.endDate, data.endTime,
      data.durationDays || '', data.tripType || '', data.passengerCount,
      data.vehicleType, data.driverName || '', data.activity || '',
      'รออนุมัติ', '', data.note || '', '', '', '', '', '', '', '',
      'รออนุมัติ', '', 'รออนุมัติ', ''
    ];
    
    sheet.appendRow(rowData);
    return { success: true, requestId: requestId };
  } catch (error) {
    return { success: false, message: error.toString() };
  }
}

// 3. ดึงรายการคำร้องทั้งหมด
function getRequests() {
  try {
    const sheet = getTargetSheet();
    const data = sheet.getDataRange().getValues();
    if (!data || data.length <= 1) return [];
    
    return data.slice(1).map((row, index) => {
      const formatTimeOnly = (val) => {
        if (!val) return '';
        if (val instanceof Date) return Utilities.formatDate(val, "GMT+7", "HH:mm") + ' น.';
        const str = String(val).trim();
        return str.length <= 5 && str !== '' ? str + ' น.' : str;
      };
      const formatDateOnly = (val) => {
        if (!val) return '';
        try { return Utilities.formatDate(new Date(val), "GMT+7", "dd/MM/yyyy"); }
        catch(e) { return String(val); }
      };

      return {
        rowIndex: index + 2,
        timestamp: row[0] ? Utilities.formatDate(new Date(row[0]), "GMT+7", "dd/MM/yyyy HH:mm") : '',
        requestId: row[1] || '',
        fullname: row[2] || '',
        department: row[3] || '',
        objective: row[4] || '',
        startDate: formatDateOnly(row[5]),
        startTime: formatTimeOnly(row[6]),
        endDate: formatDateOnly(row[7]),
        endTime: formatTimeOnly(row[8]),
        passengerCount: row[11] || 1,
        vehicleType: row[12] || '',
        driverName: row[13] || '',
        status: row[15] || 'รออนุมัติ',
        startMile: row[18] || '',
        endMile: row[19] || '',
        totalDistance: row[20] || '',
        fuelAmount: row[22] || ''
      };
    }).filter(item => item.requestId !== '');
  } catch (error) {
    return [];
  }
}

// 4. อัปเดตสถานะอนุมัติ/ไม่อนุมัติ
function updateAdminApproval(rowIndex, status, driverName, adminName) {
  try {
    const sheet = getTargetSheet();
    const today = Utilities.formatDate(new Date(), "GMT+7", "yyyy-MM-dd HH:mm");
    
    const requestId = sheet.getRange(rowIndex, 2).getValue();
    const fullname = sheet.getRange(rowIndex, 3).getValue();
    const department = sheet.getRange(rowIndex, 4).getValue();
    const objective = sheet.getRange(rowIndex, 5).getValue();
    const startDateVal = sheet.getRange(rowIndex, 6).getValue();
    const startDate = startDateVal ? Utilities.formatDate(new Date(startDateVal), "GMT+7", "dd/MM/yyyy") : '';
    const startTime = sheet.getRange(rowIndex, 7).getValue();
    const vehicleType = sheet.getRange(rowIndex, 13).getValue();

    sheet.getRange(rowIndex, 14).setValue(driverName);
    sheet.getRange(rowIndex, 16).setValue(status);
    sheet.getRange(rowIndex, 17).setValue(`${adminName} (${today})`);
    sheet.getRange(rowIndex, 26).setValue(status);
    sheet.getRange(rowIndex, 27).setValue(adminName);

    sendApprovalLineMessagingApi(requestId, fullname, department, objective, vehicleType, startDate, startTime, driverName, status, adminName);
    return { success: true };
  } catch (error) {
    return { success: false, message: error.toString() };
  }
}

// 5. แก้ไขข้อมูลคำร้อง
function updateRequestDetail(data) {
  try {
    const sheet = getTargetSheet();
    const rowIndex = Number(data.rowIndex);
    if (!rowIndex || rowIndex < 2) return { success: false, message: 'ไม่พบตำแหน่งข้อมูล' };

    sheet.getRange(rowIndex, 3).setValue(data.fullname);
    sheet.getRange(rowIndex, 4).setValue(data.department);
    sheet.getRange(rowIndex, 5).setValue(data.objective);
    sheet.getRange(rowIndex, 6).setValue(data.startDate);
    sheet.getRange(rowIndex, 7).setValue(data.startTime);
    sheet.getRange(rowIndex, 8).setValue(data.endDate);
    sheet.getRange(rowIndex, 9).setValue(data.endTime);
    sheet.getRange(rowIndex, 13).setValue(data.vehicleType);
    sheet.getRange(rowIndex, 14).setValue(data.driverName);
    sheet.getRange(rowIndex, 16).setValue(data.status);
    sheet.getRange(rowIndex, 26).setValue(data.status);

    return { success: true };
  } catch (error) {
    return { success: false, message: error.toString() };
  }
}

// 6. ส่งแจ้งเตือน LINE Flex Card
function sendApprovalLineMessagingApi(requestId, fullname, department, objective, vehicleType, startDate, startTime, driverName, status, adminName) {
  if (!LINE_CHANNEL_TOKEN || !LINE_TARGET_ID) return;

  const isApproved = status === 'อนุมัติ';
  const headerColor = isApproved ? '#0f172a' : '#991b1b';
  const statusColor = isApproved ? '#10b981' : '#ef4444';
  const sheetUrl = `https://docs.google.com/spreadsheets/d/${SPREADSHEET_ID}/edit`;

  const flexMessageJson = {
    "type": "flex",
    "altText": `ผลพิจารณาคำร้องขอใช้รถ (${requestId})`,
    "contents": {
      "type": "bubble",
      "header": {
        "type": "box", "layout": "horizontal", "backgroundColor": headerColor, "paddingAll": "15px",
        "contents": [{ "type": "text", "text": "🚘 ผลพิจารณาคำร้องขอใช้รถ", "weight": "bold", "color": "#38bdf8", "size": "md" }]
      },
      "body": {
        "type": "box", "layout": "vertical", "spacing": "md",
        "contents": [
          { "type": "text", "text": fullname || "ไม่ระบุชื่อ", "weight": "bold", "size": "xl", "color": "#0f172a" },
          { "type": "text", "text": `ฝ่าย/แผนก: ${department || '-'}`, "size": "sm", "color": "#64748b", "margin": "xs" },
          { "type": "separator", "margin": "md" },
          {
            "type": "box", "layout": "vertical", "margin": "md", "spacing": "sm",
            "contents": [
              { "type": "box", "layout": "baseline", "spacing": "sm", "contents": [{ "type": "text", "text": "รหัสคำร้อง:", "color": "#64748b", "size": "sm", "flex": 3 }, { "type": "text", "text": requestId, "weight": "bold", "color": "#1e293b", "size": "sm", "flex": 5 }] },
              { "type": "box", "layout": "baseline", "spacing": "sm", "contents": [{ "type": "text", "text": "ผลอนุมัติ:", "color": "#64748b", "size": "sm", "flex": 3 }, { "type": "text", "text": status, "weight": "bold", "color": statusColor, "size": "sm", "flex": 5 }] },
              { "type": "box", "layout": "baseline", "spacing": "sm", "contents": [{ "type": "text", "text": "รถที่ขอใช้:", "color": "#64748b", "size": "sm", "flex": 3 }, { "type": "text", "text": vehicleType || "-", "weight": "bold", "color": "#1e293b", "size": "sm", "flex": 5 }] },
              { "type": "box", "layout": "baseline", "spacing": "sm", "contents": [{ "type": "text", "text": "วัตถุประสงค์:", "color": "#64748b", "size": "sm", "flex": 3 }, { "type": "text", "text": objective || "-", "color": "#334155", "size": "sm", "flex": 5, "wrap": true }] },
              { "type": "box", "layout": "baseline", "spacing": "sm", "contents": [{ "type": "text", "text": "วัน-เวลาออก:", "color": "#64748b", "size": "sm", "flex": 3 }, { "type": "text", "text": `${startDate} ${startTime}`, "color": "#334155", "size": "sm", "flex": 5 }] },
              { "type": "box", "layout": "baseline", "spacing": "sm", "contents": [{ "type": "text", "text": "คนขับรถ:", "color": "#64748b", "size": "sm", "flex": 3 }, { "type": "text", "text": driverName || "-", "color": "#334155", "size": "sm", "flex": 5 }] },
              { "type": "box", "layout": "baseline", "spacing": "sm", "contents": [{ "type": "text", "text": "ผู้อนุมัติ:", "color": "#64748b", "size": "sm", "flex": 3 }, { "type": "text", "text": adminName || "-", "color": "#334155", "size": "sm", "flex": 5 }] }
            ]
          }
        ]
      },
      "footer": {
        "type": "box", "layout": "vertical",
        "contents": [{ "type": "button", "action": { "type": "uri", "label": "เปิดดูตารางใน Google Sheet", "uri": sheetUrl }, "style": "primary", "color": "#2563eb", "height": "sm" }]
      }
    }
  };

  UrlFetchApp.fetch('https://api.line.me/v2/bot/message/push', {
    'method': 'post',
    'headers': { 'Content-Type': 'application/json', 'Authorization': 'Bearer ' + LINE_CHANNEL_TOKEN },
    'payload': JSON.stringify(flexMessageJson),
    'muteHttpExceptions': true
  });
}

// 7. บันทึกงานคนขับ
function updateDriverLog(data) {
  try {
    const sheet = getTargetSheet();
    const startMile = Number(data.startMile) || 0;
    const endMile = Number(data.endMile) || 0;
    const totalDistance = endMile > startMile ? (endMile - startMile) : 0;
    
    sheet.getRange(data.rowIndex, 19).setValue(startMile);
    sheet.getRange(data.rowIndex, 20).setValue(endMile);
    sheet.getRange(data.rowIndex, 21).setValue(totalDistance);
    sheet.getRange(data.rowIndex, 22).setValue(data.fuelType || '');
    sheet.getRange(data.rowIndex, 23).setValue(data.fuelAmount || 0);
    sheet.getRange(data.rowIndex, 24).setValue(data.driverNote || '');
    sheet.getRange(data.rowIndex, 16).setValue('เสร็จสิ้น');
    
    return { success: true };
  } catch (error) {
    return { success: false, message: error.toString() };
  }
}

// 8. ลบคำร้อง
function deleteRequest(rowIndex) {
  try {
    const sheet = getTargetSheet();
    const index = Number(rowIndex);
    if (!index || index < 2) return { success: false, message: 'ไม่พบตำแหน่งข้อมูล' };
    sheet.deleteRow(index);
    return { success: true };
  } catch (error) {
    return { success: false, message: error.toString() };
  }
}

// 9. รองรับ HTTP POST จาก Vercel
// รองรับ HTTP POST จาก Vercel/Client
function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    let result = { success: false, message: 'Invalid action' };
    
    // 1. ยื่นคำร้องขอใช้รถ
    if (data.action === 'submitRequest') {
      result = submitRequest(data);
    } 
    // 2. อนุมัติ / ไม่อนุมัติ
    else if (data.action === 'updateAdminApproval') {
      result = updateAdminApproval(data.rowIndex, data.status, data.driverName, data.adminName);
    } 
    // 3. บันทึกแก้ไขข้อมูลคำร้อง
    else if (data.action === 'updateRequestDetail') {
      result = updateRequestDetail(data.data);
    } 
    // 4. ลบคำร้อง
    else if (data.action === 'deleteRequest') {
      result = deleteRequest(data.rowIndex);
    }
    // 5. คนขับบันทึกระยะทาง/ค่าน้ำมัน
    else if (data.action === 'updateDriverLog') {
      result = updateDriverLog(data.data);
    }
    
    return ContentService
      .createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (error) {
    return ContentService
      .createTextOutput(JSON.stringify({ success: false, message: error.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}
