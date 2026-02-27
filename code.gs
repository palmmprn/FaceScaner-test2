function doGet(e) {
  var page = e.parameter.page || 'menu';
  var template;

  // อ่านค่า Config ล่าสุดจาก Google Sheets ทุกครั้งที่โหลดหน้า
  var currentConfig = getConfig();

  if (page == 'register') template = HtmlService.createTemplateFromFile('register');
  else if (page == 'scan') template = HtmlService.createTemplateFromFile('scan');
  else if (page == 'config') template = HtmlService.createTemplateFromFile('config'); 
  else template = HtmlService.createTemplateFromFile('menu');
  
  // ส่งค่า Config ไปให้หน้าเว็บใช้ได้เลย
  template.config = currentConfig;

  return template.evaluate()
      .addMetaTag('viewport', 'width=device-width, initial-scale=1')
      .setTitle('Face Recognition System')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function getScriptUrl() {
  return ScriptApp.getService().getUrl();
}

// --- ส่วนจัดการใบหน้า (Users) ---
function registerUser(name, faceDescriptor) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Users');
  if (!sheet) sheet = ss.insertSheet('Users'); 
  
  sheet.appendRow([name, JSON.stringify(faceDescriptor), new Date()]); 
  return "บันทึกข้อมูลหน้าเรียบร้อย";
}

function getKnownFaces() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Users');
  if (!sheet) return [];
  
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];

  let users = [];
  for (let i = 1; i < data.length; i++) {
    const name = data[i][0];
    const jsonStr = data[i][1];
    if (name && jsonStr) {
      try {
        users.push({
          label: name, 
          descriptor: JSON.parse(jsonStr)
        });
      } catch (e) {}
    }
  }
  return users;
}

// --- ส่วนบันทึกเวลา (Attendance) ---
function logAttendance(name, lat, lng) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Attendance');
  if (!sheet) {
    sheet = ss.insertSheet('Attendance');
    sheet.appendRow(['Name', 'Time', 'Date', 'Latitude', 'Longitude', 'Google Map Link']);
  }

  const now = new Date();
  const mapLink = (lat && lng) ? `https://www.google.com/maps?q=${lat},${lng}` : "";
  
  // --- ปรับแก้ตรงนี้: เปลี่ยนรูปแบบวันที่ให้เป็น ค.ศ. ---
  // ใช้ Utilities.formatDate กำหนด pattern เป็น "d/M/yyyy" (ปี ค.ศ.)
  // Session.getScriptTimeZone() ใช้ Timezone ของ Script (ควรตั้งเป็น GMT+7 Bangkok)
  const dateStr = Utilities.formatDate(now, Session.getScriptTimeZone(), "d/M/yyyy");
  const timeStr = Utilities.formatDate(now, Session.getScriptTimeZone(), "HH:mm:ss");
  
  sheet.appendRow([
    name, 
    timeStr, // เวลา
    "'" + dateStr, // วันที่ (ใส่ ' นำหน้าเพื่อให้ Google Sheets มองเป็น Text และไม่แปลงกลับเป็น พ.ศ. อัตโนมัติ)
    lat || "-",
    lng || "-",
    mapLink
  ]);
  return "บันทึกเวลาสำเร็จ";
}

// --- ส่วนจัดการ Config (GPS) ---
function saveConfig(lat, lng, radius) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('Config');
  
  if (!sheet) {
    sheet = ss.insertSheet('Config');
    sheet.getRange("A1:B1").setValues([["Parameter", "Value"]]);
    sheet.getRange("A2").setValue("Target Latitude");
    sheet.getRange("A3").setValue("Target Longitude");
    sheet.getRange("A4").setValue("Allowed Radius (KM)");
    sheet.setColumnWidth(1, 150); 
  }
  
  sheet.getRange("B2").setValue(lat);
  sheet.getRange("B3").setValue(lng);
  sheet.getRange("B4").setValue(radius);
  
  return "บันทึกการตั้งค่าลง Google Sheets เรียบร้อย";
}

function getConfig() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Config');
  
  let config = {
    lat: 0,
    lng: 0,
    radius: 0.5 
  };

  if (sheet) {
    const latVal = sheet.getRange("B2").getValue();
    const lngVal = sheet.getRange("B3").getValue();
    const radiusVal = sheet.getRange("B4").getValue();

    if (latVal !== "") config.lat = parseFloat(latVal);
    if (lngVal !== "") config.lng = parseFloat(lngVal);
    if (radiusVal !== "") config.radius = parseFloat(radiusVal);
  }
  
  return config;
}
