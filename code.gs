// ============================================================
//  GOOGLE APPS SCRIPT — REST API Backend v2
//  รองรับ: เข้างาน / ออกงาน (type: in/out)
//  Deploy > New deployment > Web App
//  Execute as: Me | Who has access: Anyone
// ============================================================

function doGet(e) {
  const action = e.parameter.action;
  let result;

  if      (action === 'getConfig')     result = getConfig();
  else if (action === 'getKnownFaces') result = getKnownFaces();
  else                                  result = { error: 'Unknown action: ' + action };

  return ContentService
    .createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  let data;
  try {
    data = JSON.parse(e.postData.contents);
  } catch (err) {
    return jsonOut({ error: 'Invalid JSON body' });
  }

  const action = data.action;
  let result;

  if      (action === 'registerUser')  result = registerUser(data.name, data.faceDescriptor);
  else if (action === 'logAttendance') result = logAttendance(data.name, data.lat, data.lng, data.type);
  else if (action === 'saveConfig')    result = saveConfig(data.lat, data.lng, data.radius);
  else                                  result = { error: 'Unknown action: ' + action };

  return jsonOut(result);
}

function jsonOut(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// ─── ลงทะเบียนใบหน้า ────────────────────────────────────────
function registerUser(name, faceDescriptor) {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  let sheet   = ss.getSheetByName('Users');
  if (!sheet) {
    sheet = ss.insertSheet('Users');
    sheet.appendRow(['Name', 'Face Descriptor (128D JSON)', 'Registered At']);
    sheet.setFrozenRows(1);
  }

  sheet.appendRow([name, JSON.stringify(faceDescriptor), new Date()]);
  return { success: true, message: 'บันทึกข้อมูลหน้าเรียบร้อย' };
}

function getKnownFaces() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Users');
  if (!sheet) return [];

  const data  = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];

  const users = [];
  // รวม descriptor หลายแถวของคนเดียวกัน (multi-sample)
  const map   = {};
  for (let i = 1; i < data.length; i++) {
    const name = String(data[i][0]).trim();
    const json = data[i][1];
    if (!name || !json) continue;
    try {
      const desc = JSON.parse(json);
      if (!map[name]) map[name] = [];
      map[name].push(desc);
    } catch (e) {}
  }

  // ส่ง descriptor แรกของแต่ละคน (face-api จะ match ทีละ descriptor)
  // ถ้ามีหลายตัวอย่าง ส่งทุกตัวเป็น array ของ records แยก
  for (const name in map) {
    map[name].forEach(desc => {
      users.push({ label: name, descriptor: desc });
    });
  }

  return users;
}

// ─── บันทึกเวลาเข้า/ออก ────────────────────────────────────
function logAttendance(name, lat, lng, type) {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  let sheet   = ss.getSheetByName('Attendance');

  const isNew = !sheet;
  if (isNew) {
    sheet = ss.insertSheet('Attendance');
    sheet.setFrozenRows(1);
  }

  // สร้าง/ตรวจ header (8 คอลัมน์)
  const headers = ['ชื่อพนักงาน', 'ประเภท', 'สถานะ', 'เวลา', 'วันที่', 'Latitude', 'Longitude', 'Google Map Link'];
  if (isNew) {
    sheet.appendRow(headers);
    const hRange = sheet.getRange(1, 1, 1, headers.length);
    hRange.setBackground('#1a3a5c')
          .setFontColor('#ffffff')
          .setFontWeight('bold')
          .setHorizontalAlignment('center');
    sheet.setColumnWidth(1, 160);
    sheet.setColumnWidth(2,  90);
    sheet.setColumnWidth(3, 120);
    sheet.setColumnWidth(4,  90);
    sheet.setColumnWidth(5, 100);
    sheet.setColumnWidth(6,  90);
    sheet.setColumnWidth(7,  90);
    sheet.setColumnWidth(8, 260);
  }

  const now     = new Date();
  const tz      = Session.getScriptTimeZone();
  const dateStr = Utilities.formatDate(now, tz, 'd/M/yyyy');
  const timeStr = Utilities.formatDate(now, tz, 'HH:mm:ss');

  const isOut   = (type === 'out');
  const typeStr = isOut ? 'ออกงาน'  : 'เข้างาน';
  // สถานะ — ใช้ emoji + ข้อความให้เห็นชัดใน Sheets
  const status  = isOut ? '🔴 ออกงานแล้ว' : '🟢 เข้างานแล้ว';
  const mapLink = (lat && lng) ? `https://www.google.com/maps?q=${lat},${lng}` : '';

  sheet.appendRow([
    name,
    typeStr,
    status,
    timeStr,
    "'" + dateStr,
    lat  || '-',
    lng  || '-',
    mapLink
  ]);

  // ระบายสีแถวตามประเภท
  const rowIdx  = sheet.getLastRow();
  const numCols = headers.length;
  const bgColor = isOut ? '#fff3e0' : '#e8f5e9';
  const txColor = isOut ? '#7c3009' : '#14532d';
  const statusCell = sheet.getRange(rowIdx, 3); // คอลัมน์ สถานะ
  statusCell.setFontColor(isOut ? '#c2410c' : '#15803d').setFontWeight('bold');
  sheet.getRange(rowIdx, 1, 1, numCols).setBackground(bgColor).setFontColor(txColor);
  // คืนสีตัวอักษรปกติยกเว้นช่องสถานะ
  sheet.getRange(rowIdx, 1, 1, numCols).setFontColor('#1e293b');
  statusCell.setFontColor(isOut ? '#c2410c' : '#15803d').setFontWeight('bold');

  return {
    success: true,
    message: `บันทึก${typeStr}สำเร็จ — ${name} ${timeStr}`
  };
}

// ─── Config GPS ─────────────────────────────────────────────
function saveConfig(lat, lng, radius) {
  const ss  = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('Config');

  if (!sheet) {
    sheet = ss.insertSheet('Config');
    sheet.getRange('A1:B1').setValues([['Parameter', 'Value']]);
    sheet.getRange('A2:A4').setValues([
      ['Target Latitude'],
      ['Target Longitude'],
      ['Allowed Radius (KM)']
    ]);
    sheet.getRange(1, 1, 1, 2).setBackground('#1e3a5f').setFontColor('#ffffff').setFontWeight('bold');
    sheet.setColumnWidth(1, 180);
    sheet.setColumnWidth(2, 120);
  }

  sheet.getRange('B2').setValue(parseFloat(lat));
  sheet.getRange('B3').setValue(parseFloat(lng));
  sheet.getRange('B4').setValue(parseFloat(radius));

  return { success: true, message: 'บันทึกการตั้งค่าลง Google Sheets เรียบร้อย' };
}

function getConfig() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Config');
  let config  = { lat: 0, lng: 0, radius: 0 };

  if (sheet) {
    const lat    = sheet.getRange('B2').getValue();
    const lng    = sheet.getRange('B3').getValue();
    const radius = sheet.getRange('B4').getValue();
    if (lat    !== '') config.lat    = parseFloat(lat);
    if (lng    !== '') config.lng    = parseFloat(lng);
    if (radius !== '') config.radius = parseFloat(radius);
  }

  return config;
}
