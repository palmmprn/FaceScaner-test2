function doGet(e) {
  var page = e.parameter.page || 'menu';
  var template;

  var currentConfig = getConfig();

  if (page == 'register') template = HtmlService.createTemplateFromFile('register');
  else if (page == 'scan') template = HtmlService.createTemplateFromFile('scan');
  else if (page == 'config') template = HtmlService.createTemplateFromFile('config'); 
  else if (page == 'dashboard') template = HtmlService.createTemplateFromFile('dashboard');
  else if (page == 'emp_dashboard') template = HtmlService.createTemplateFromFile('emp_dashboard');
  else template = HtmlService.createTemplateFromFile('menu');
  
  template.config = currentConfig;

  return template.evaluate()
      .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1, user-scalable=0')
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
      try { users.push({ label: name, descriptor: JSON.parse(jsonStr) }); } catch (e) {}
    }
  }
  return users;
}

// --- ส่วนบันทึกเวลา (Attendance) รองรับ เข้า-ออก และคำนวณเวลาชดเชย ---
function logAttendance(name, lat, lng, scanType) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Attendance');
  if (!sheet) {
    sheet = ss.insertSheet('Attendance');
    sheet.appendRow(['Name', 'Date', 'Time In', 'Time Out', 'Earned Comp (Mins)', 'Used Comp (Mins)', 'Latitude', 'Longitude']);
  }

  const now = new Date();
  const dateStr = Utilities.formatDate(now, Session.getScriptTimeZone(), "d/M/yyyy");
  const timeStr = Utilities.formatDate(now, Session.getScriptTimeZone(), "HH:mm"); 
  
  // แปลงเวลาเป็นนาทีเพื่อคำนวณ (09:00 = 540 นาที, 18:00 = 1080 นาที)
  const currentMins = now.getHours() * 60 + now.getMinutes();
  const startWorkMins = 9 * 60; // 09:00
  const endWorkMins = 18 * 60;  // 18:00

  const data = sheet.getDataRange().getValues();
  let rowIndexToUpdate = -1;

  // หาว่าวันนี้คนนี้เคยสแกนหรือยัง
  for (let i = data.length - 1; i >= 1; i--) {
    let rowDate = String(data[i][1]).replace(/'/g, '');
    if (data[i][0] === name && rowDate === dateStr) {
      rowIndexToUpdate = i + 1; // +1 เพราะ array เริ่มที่ 0 แต่ sheet เริ่มที่ 1
      break;
    }
  }

  if (scanType === 'IN') {
    let earnedComp = 0;
    if (currentMins < startWorkMins) {
      earnedComp = startWorkMins - currentMins; // มาก่อนเวลา ได้นาทีสะสม
    }
    
    if (rowIndexToUpdate > -1) {
      // ถ้าเคยกดเข้าแล้วมากดเข้าอีก ให้อัปเดตเวลาเข้าล่าสุด
      sheet.getRange(rowIndexToUpdate, 3).setValue(timeStr);
      sheet.getRange(rowIndexToUpdate, 5).setValue(earnedComp);
    } else {
      // บันทึกแถวใหม่
      sheet.appendRow([name, "'" + dateStr, timeStr, "-", earnedComp, 0, lat || "-", lng || "-"]);
    }
    return "บันทึกเวลาเข้างานเรียบร้อย ได้เวลาสะสม: " + earnedComp + " นาที";

  } else if (scanType === 'OUT') {
    let usedComp = 0;
    if (currentMins < endWorkMins) {
      usedComp = endWorkMins - currentMins; // เลิกก่อนเวลา โดนหักนาทีสะสม
    }

    if (rowIndexToUpdate > -1) {
      // อัปเดตเวลาออก ในแถวของวันนี้
      sheet.getRange(rowIndexToUpdate, 4).setValue(timeStr);
      sheet.getRange(rowIndexToUpdate, 6).setValue(usedComp);
      return "บันทึกเวลาออกงานเรียบร้อย ใช้เวลาสะสมไป: " + usedComp + " นาที";
    } else {
      // ลืมสแกนเข้า แต่มากดสแกนออก
      sheet.appendRow([name, "'" + dateStr, "-", timeStr, 0, usedComp, lat || "-", lng || "-"]);
      return "บันทึกเวลาออกงาน (ไม่พบข้อมูลสแกนเข้าวันนี้)";
    }
  }
}

// --- ฟังก์ชันดึงข้อมูลเวลาชดเชยของพนักงานรายบุคคล ---
function getMyCompStats(empName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Attendance');
  if (!sheet) return { totalEarned: 0, totalUsed: 0, balance: 0, logs: [] };

  const data = sheet.getDataRange().getValues();
  let totalEarned = 0;
  let totalUsed = 0;
  let logs = [];

  for (let i = data.length - 1; i >= 1; i--) {
    if (data[i][0] === empName) {
      let earned = Number(data[i][4]) || 0;
      let used = Number(data[i][5]) || 0;
      totalEarned += earned;
      totalUsed += used;
      
      if(logs.length < 5) { // เก็บประวัติล่าสุด 5 วัน
        logs.push({
          date: String(data[i][1]).replace(/'/g, ''),
          timeIn: data[i][2],
          timeOut: data[i][3],
          earned: earned,
          used: used
        });
      }
    }
  }

  return {
    totalEarned: totalEarned,
    totalUsed: totalUsed,
    balance: totalEarned - totalUsed, // ยอดคงเหลือ
    logs: logs
  };
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
  return "บันทึกการตั้งค่าเรียบร้อย";
}

function getConfig() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Config');
  let config = { lat: 0, lng: 0, radius: 0.5 };
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

// --- ส่วน Dashboard (Admin) ---
function getDashboardData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Attendance');
  if (!sheet) return { todayCount: 0, leaderboard: [], recentLogs: [] };

  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return { todayCount: 0, leaderboard: [], recentLogs: [] };

  const records = data.slice(1).map(row => ({
    name: row[0],
    dateStr: String(row[1]).replace(/'/g, ''),
    timeIn: row[2],
    timeOut: row[3]
  })).reverse(); 

  const now = new Date();
  const todayStr = Utilities.formatDate(now, Session.getScriptTimeZone(), "d/M/yyyy");

  const todayRecords = records.filter(r => r.dateStr === todayStr);
  const uniqueTodayNames = [...new Set(todayRecords.map(r => r.name))];
  const todayCount = uniqueTodayNames.length;

  let leaderboardMap = {};
  let sevenDaysAgo = new Date(now.getTime() - (7 * 24 * 60 * 60 * 1000));
  
  records.forEach(r => {
    let parts = r.dateStr.split('/');
    if(parts.length === 3) {
      let rowDate = new Date(parts[2], parts[1]-1, parts[0]);
      if(rowDate >= sevenDaysAgo && r.timeIn && r.timeIn !== "-") {
        if(!leaderboardMap[r.name] || r.timeIn < leaderboardMap[r.name]) {
          leaderboardMap[r.name] = r.timeIn;
        }
      }
    }
  });

  let leaderboard = [];
  for (const [name, time] of Object.entries(leaderboardMap)) {
    leaderboard.push({name: name, time: time});
  }
  leaderboard.sort((a, b) => a.time.localeCompare(b.time));
  leaderboard = leaderboard.slice(0, 5);

  const recentLogs = records.slice(0, 10).map(r => ({
    name: r.name,
    dateStr: r.dateStr,
    timeStr: r.timeIn !== "-" ? r.timeIn : r.timeOut,
    mapLink: ""
  }));

  return { todayCount: todayCount, leaderboard: leaderboard, recentLogs: recentLogs };
}
