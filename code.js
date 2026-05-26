/**
 * BUTC Online Teaching Report System
 * (Copyright © 2026 Assawin Namsert. All Rights Reserved.)
 */

const CONFIG = {
  SYSTEM_NAME: "BUTC Online Teaching Report System",
  COLLEGE_NAME: "วิทยาลัยเทคนิคบูรพาปราจีน",
  DEPARTMENTS: ["ช่างยนต์", "ช่างไฟฟ้า", "ช่างซ่อมบำรุง", "เมคคาทรอนิกส์", "การบัญชี", "เทคโนโลยีธุรกิจดิจิทัล", "การจัดการสำนักงานดิจิทัล", "เทคนิคพื้นฐาน"],
  SHEET_USERS: "Users",
  SHEET_REPORTS: "Reports"
};

function doGet(e) {
  try {
    const template = HtmlService.createTemplateFromFile('Index');
    return template.evaluate()
      .setTitle(CONFIG.SYSTEM_NAME)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  } catch (err) {
    return HtmlService.createHtmlOutput("<h3>เกิดข้อผิดพลาดในการโหลดระบบ</h3><p>" + err.message + "</p>");
  }
}

function include(filename) { return HtmlService.createHtmlOutputFromFile(filename).getContent(); }

function getDb() {
  const props = PropertiesService.getScriptProperties();
  const ssId = props.getProperty('spreadsheet_id');
  if (ssId) { try { return SpreadsheetApp.openById(ssId); } catch (e) {} }
  return SpreadsheetApp.getActiveSpreadsheet();
}

function setupSystem() {
  const props = PropertiesService.getScriptProperties();
  let ssId = props.getProperty('spreadsheet_id');
  let ss;
  if (ssId) { try { ss = SpreadsheetApp.openById(ssId); } catch (e) { ss = null; } }
  if (!ss) {
    let folder;
    try {
      const fileId = ScriptApp.getScriptId();
      const parentFolders = DriveApp.getFileById(fileId).getParents();
      folder = parentFolders.hasNext() ? parentFolders.next() : DriveApp.getRootFolder();
    } catch (e) { folder = DriveApp.getRootFolder(); }
    ss = SpreadsheetApp.create(CONFIG.SYSTEM_NAME + "_DB");
    ssId = ss.getId();
    const file = DriveApp.getFileById(ssId);
    file.moveTo(folder);
    props.setProperty('spreadsheet_id', ssId);
  }

  // Users Sheet
  let userSheet = ss.getSheetByName(CONFIG.SHEET_USERS);
  if (!userSheet) {
    userSheet = ss.insertSheet(CONFIG.SHEET_USERS);
    userSheet.appendRow(["userId", "email", "password", "fullName", "department", "phone", "role", "status", "timestamp"]);
    userSheet.getRange(1, 1, 1, 9).setFontWeight("bold").setBackground("#f3f3f3");
    const defaultSheet = ss.getSheetByName("Sheet1");
    if (defaultSheet) ss.deleteSheet(defaultSheet);
  }
  
  // Reports Sheet
  let reportSheet = ss.getSheetByName(CONFIG.SHEET_REPORTS);
  const reportHeaders = [
    "reportId", "userEmail", "teachingDate", "teachingLevel", "studentDepts", "periods", "totalHours", 
    "subjectName", "teachingTopic", "totalStudents", "presentCount", "absentCount", "leaveCount", "problems", "solutions"
  ];
  // เพิ่มคอลัมน์สำหรับเก็บรูป (Split Base64 20 Cells)
  for(let i=1; i<=20; i++) reportHeaders.push("imageData_" + i);
  reportHeaders.push("status", "timestamp");

  if (!reportSheet) {
    reportSheet = ss.insertSheet(CONFIG.SHEET_REPORTS);
    reportSheet.appendRow(reportHeaders);
    reportSheet.getRange(1, 1, 1, reportHeaders.length).setFontWeight("bold").setBackground("#f3f3f3");
  } else {
    // อัปเดตหัวตารางกรณีมีการเปลี่ยนแปลงโครงสร้าง
    reportSheet.getRange(1, 1, 1, reportHeaders.length).setValues([reportHeaders]).setFontWeight("bold").setBackground("#f3f3f3");
  }
  return "ระบบถูกตั้งค่าเรียบร้อยแล้ว (ไม่ใช้ DriveApp แล้ว)\nSpreadsheet: " + ss.getUrl();
}

function loginUser(credentials) {
  const ss = getDb();
  const sheet = ss.getSheetByName(CONFIG.SHEET_USERS);
  const data = sheet.getDataRange().getValues();
  const email = credentials.email.trim().toLowerCase();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][1]).trim().toLowerCase() === email && String(data[i][2]).trim() === credentials.password.trim()) {
      if (data[i][7] === 'Pending') return { success: false, status: 'Pending', message: "รอการอนุมัติ" };
      return { success: true, user: { email: data[i][1], fullName: data[i][3], department: data[i][4], phone: data[i][5], role: data[i][6], status: data[i][7] } };
    }
  }
  return { success: false, message: "อีเมลหรือรหัสผ่านไม่ถูกต้อง" };
}

function registerUser(userData) {
  const ss = getDb();
  const sheet = ss.getSheetByName(CONFIG.SHEET_USERS);
  const data = sheet.getDataRange().getValues();
  const email = userData.email.trim().toLowerCase();
  for (let i = 1; i < data.length; i++) { if (String(data[i][1]).trim().toLowerCase() === email) return { success: false, message: "อีเมลนี้ถูกใช้งานไปแล้ว" }; }
  const role = data.length === 1 ? 'Admin' : 'Teacher';
  sheet.appendRow(["U-" + new Date().getTime(), email, userData.password.trim(), userData.fullName.trim(), userData.department, userData.phone.trim(), role, (role === 'Admin' ? 'Approved' : 'Pending'), new Date()]);
  return { success: true, message: "สมัครสำเร็จ" };
}

function updateUserProfile(userData) {
  try {
    const ss = getDb();
    const sheet = ss.getSheetByName(CONFIG.SHEET_USERS);
    const data = sheet.getDataRange().getValues();
    const email = userData.email.trim().toLowerCase();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][1]).trim().toLowerCase() === email) {
        sheet.getRange(i + 1, 3, 1, 4).setValues([[userData.password, userData.fullName, userData.department, userData.phone]]);
        return { success: true, user: { email: data[i][1], fullName: userData.fullName, department: userData.department, phone: userData.phone, role: data[i][6], status: data[i][7] } };
      }
    }
    return { success: false, message: "ไม่พบข้อมูลผู้ใช้" };
  } catch (e) { return { success: false, message: e.toString() }; }
}

function submitReport(reportData) {
  try {
    const ss = getDb();
    const sheet = ss.getSheetByName(CONFIG.SHEET_REPORTS);
    
    // แยก Base64 เป็นส่วนๆ ส่วนละ 45,000 ตัวอักษร (ไม่เกิน 50,000)
    const fullBase64 = reportData.images.map(img => {
      if (img.data && img.data.startsWith('data:')) return img.data;
      return `data:${img.mimeType};base64,${img.data}`;
    }).join('|');
    
    const chunks = [];
    const CHUNK_SIZE = 45000;
    for (let i = 0; i < 20; i++) {
      chunks.push(fullBase64.substring(i * CHUNK_SIZE, (i + 1) * CHUNK_SIZE));
    }

    const timestampStr = Utilities.formatDate(new Date(), "Asia/Bangkok", "dd/MM/yyyy HH:mm:ss");

    const rowData = [
      reportData.reportId || "R-" + new Date().getTime(),
      reportData.userEmail,
      reportData.teachingDate,
      reportData.teachingLevel,
      reportData.targetDepartments,
      reportData.periods,
      reportData.totalHours,
      reportData.subjectName,
      reportData.teachingTopic,
      reportData.totalStudents,
      reportData.presentCount,
      reportData.absentCount,
      reportData.leaveCount,
      reportData.problems,
      reportData.solutions,
      ...chunks, // คอลัมน์ที่ 16 - 35
      'Pending', // สถานะ
      timestampStr // Timestamp
    ];

    if (reportData.reportId) {
      const data = sheet.getDataRange().getValues();
      for (let i = 1; i < data.length; i++) {
        if (String(data[i][0]) === String(reportData.reportId)) {
          sheet.getRange(i + 1, 1, 1, rowData.length).setValues([rowData]);
          return { success: true };
        }
      }
    }
    sheet.appendRow(rowData);
    return { success: true };
  } catch (err) { return { success: false, message: err.toString() }; }
}

function getUserReports(email) {
  try {
    const ss = getDb();
    const sheet = ss.getSheetByName(CONFIG.SHEET_REPORTS);
    if (!sheet || sheet.getLastRow() < 2) return [];
    const data = sheet.getRange(1, 1, sheet.getLastRow(), sheet.getLastColumn()).getDisplayValues();
    const results = [];
    const target = String(email).trim().toLowerCase();
    for (let i = data.length - 1; i >= 1; i--) { 
      if (String(data[i][1]).trim().toLowerCase() === target) {
        results.push(reconstructReport(data[i]));
      }
    }
    return results;
  } catch (e) { return []; }
}

function getAllReports() {
  try {
    const ss = getDb();
    const sheet = ss.getSheetByName(CONFIG.SHEET_REPORTS);
    const userSheet = ss.getSheetByName(CONFIG.SHEET_USERS);
    if (!sheet || sheet.getLastRow() < 2) return [];
    const data = sheet.getRange(1, 1, sheet.getLastRow(), sheet.getLastColumn()).getDisplayValues();
    const users = userSheet.getDataRange().getValues();
    const userMap = {};
    users.forEach(u => userMap[String(u[1]).toLowerCase()] = u[3]);
    const results = [];
    for (let i = data.length - 1; i >= 1; i--) {
      const r = reconstructReport(data[i]);
      r.teacherName = userMap[String(r.email).toLowerCase()] || r.email;
      results.push(r);
    }
    return results;
  } catch (e) { return []; }
}

function reconstructReport(row) {
  let fullImageData = "";
  for (let j = 15; j < 35; j++) {
    fullImageData += row[j];
  }
  
  let isLate = false;
  try {
    const teachDateParts = String(row[2]).split('/');
    const teachDate = new Date(parseInt(teachDateParts[2]) - 543, parseInt(teachDateParts[1]) - 1, parseInt(teachDateParts[0]));
    
    const tsStr = String(row[36] || "");
    const tsDateParts = tsStr.split(' ')[0].split('/');
    let tsDate = new Date();
    if(tsDateParts.length === 3) {
      tsDate = new Date(parseInt(tsDateParts[2]), parseInt(tsDateParts[1]) - 1, parseInt(tsDateParts[0]));
    }
    
    teachDate.setHours(0,0,0,0);
    tsDate.setHours(0,0,0,0);
    
    if (tsDate > teachDate) {
      isLate = true;
    }
  } catch(e) {}

  return { 
    id: row[0], email: row[1], date: row[2], level: row[3], depts: row[4], periods: row[5], hours: row[6], 
    subject: row[7], topic: row[8], totalStud: row[9], present: row[10], absent: row[11], leave: row[12], 
    problems: row[13], solutions: row[14], images: fullImageData.split('|').filter(img => img.startsWith('data:')), 
    status: row[35] || 'Pending', timestamp: row[36] || '', isLate: isLate 
  };
}

function getPendingUsers() {
  try {
    const ss = getDb();
    const sheet = ss.getSheetByName(CONFIG.SHEET_USERS);
    const data = sheet.getDataRange().getValues();
    const results = [];
    for (let i = 1; i < data.length; i++) { if (String(data[i][7]).trim() === 'Pending') results.push({ email: data[i][1], fullName: data[i][3], department: data[i][4], phone: data[i][5] }); }
    return results;
  } catch (e) { return []; }
}

function getAllTeachers() {
  try {
    const ss = getDb();
    const sheet = ss.getSheetByName(CONFIG.SHEET_USERS);
    const data = sheet.getDataRange().getValues();
    const results = [];
    for (let i = 1; i < data.length; i++) { results.push({ fullName: data[i][3], email: data[i][1], department: data[i][4], phone: data[i][5], role: data[i][6], status: data[i][7] }); }
    return results;
  } catch (e) { return []; }
}

function approveUser(email, action) {
  const ss = getDb(), sheet = ss.getSheetByName(CONFIG.SHEET_USERS), data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) { if (String(data[i][1]).trim().toLowerCase() === String(email).trim().toLowerCase()) { sheet.getRange(i + 1, 8).setValue(action === 'approve' ? 'Approved' : 'Rejected'); return { success: true }; } }
  return { success: false };
}

function approveReport(id, action) {
  const ss = getDb(), sheet = ss.getSheetByName(CONFIG.SHEET_REPORTS), data = sheet.getDataRange().getValues();
  const targetStatus = action === 'approve' ? 'Approved' : (action === 'cancel' ? 'Pending' : 'Rejected');
  for (let i = 1; i < data.length; i++) { if (String(data[i][0]) === String(id)) { sheet.getRange(i + 1, 36).setValue(targetStatus); return { success: true }; } }
  return { success: false };
}

function bulkApproveReports(ids) {
  const ss = getDb(), sheet = ss.getSheetByName(CONFIG.SHEET_REPORTS), data = sheet.getDataRange().getValues();
  let count = 0;
  ids.forEach(id => {
    for (let i = 1; i < data.length; i++) { if (String(data[i][0]) === String(id)) { sheet.getRange(i + 1, 36).setValue('Approved'); count++; break; } }
  });
  return { success: true, message: `อนุมัติแล้ว ${count} รายการ` };
}

function generateDailyPDF(dateStr, endDateStr = null) {
  try {
    const ss = getDb(), reportSheet = ss.getSheetByName(CONFIG.SHEET_REPORTS), userSheet = ss.getSheetByName(CONFIG.SHEET_USERS);
    if (!reportSheet || reportSheet.getLastRow() < 2) throw new Error("ไม่พบข้อมูลรายงาน");
    
    let searchStart = dateStr;
    const parts = dateStr.split('-');
    if (parts.length === 3) searchStart = `${parts[2]}/${parts[1]}/${parseInt(parts[0]) + 543}`;
    
    let searchEnd = endDateStr;
    if (endDateStr) {
      const partsE = endDateStr.split('-');
      if (partsE.length === 3) searchEnd = `${partsE[2]}/${partsE[1]}/${parseInt(partsE[0]) + 543}`;
    } else {
      searchEnd = searchStart;
    }

    const parseBE = (d) => { const p = d.split('/'); return new Date(parseInt(p[2])-543, parseInt(p[1])-1, parseInt(p[0])); };
    const reports = reportSheet.getRange(2, 1, reportSheet.getLastRow() - 1, reportSheet.getLastColumn()).getDisplayValues();
    const users = userSheet.getRange(2, 1, userSheet.getLastRow() - 1, userSheet.getLastColumn()).getValues();
    const userMap = {};
    users.forEach(u => { userMap[String(u[1]).toLowerCase()] = { name: String(u[3] || ""), dept: String(u[4] || "") }; });

    const filtered = reports.filter(r => {
      const d = parseBE(r[2]);
      return d >= parseBE(searchStart) && d <= parseBE(searchEnd) && r[35] === 'Approved';
    }).map(r => {
      const u = userMap[String(r[1]).toLowerCase()] || { name: String(r[1]), dept: "ไม่ระบุ" };
      let imgs = ""; for(let j=15; j<35; j++) imgs += r[j];
      
      let isLate = false;
      try {
        const p = r[2].split('/');
        const tDate = new Date(parseInt(p[2])-543, parseInt(p[1])-1, parseInt(p[0]));
        const tsParts = String(r[36] || "").split(' ')[0].split('/');
        if(tsParts.length === 3) {
           const tsDate = new Date(parseInt(tsParts[2]), parseInt(tsParts[1]) - 1, parseInt(tsParts[0]));
           tDate.setHours(0,0,0,0); tsDate.setHours(0,0,0,0);
           if (tsDate > tDate) isLate = true;
        }
      } catch(e) {}

      return { 
        teacherDept: u.dept, name: u.name, date: r[2], studentDept: r[4], 
        subject: r[7], topic: r[8], level: r[3], periods: r[5], hours: r[6],
        totalStud: r[9], present: r[10], absent: r[11], leave: r[12],
        problems: r[13] || '-', solutions: r[14] || '-',
        images: imgs.split('|').filter(i => i.startsWith('data:')),
        timestamp: r[36] || '-', isLate: isLate
      };
    });

    if (filtered.length === 0) throw new Error("ไม่พบรายงานที่อนุมัติแล้วในช่วงวันที่เลือก");
    filtered.sort((a, b) => a.teacherDept.localeCompare(b.teacherDept) || a.name.localeCompare(b.name));

    let html = `<html><head><style>
      @import url('https://fonts.googleapis.com/css2?family=Sarabun&display=swap'); 
      body { font-family: 'Sarabun', sans-serif; font-size: 14px; padding: 20px; color: #333; } 
      h1 { text-align: center; color: #333; margin-bottom: 5px; font-size: 22px; }
      h2 { text-align: center; color: #1a73e8; border-bottom: 2px solid #1a73e8; padding-bottom: 10px; font-size: 18px; margin-top: 0; } 
      .report-card { border: 1px solid #dee2e6; padding: 12px; margin-bottom: 15px; border-radius: 6px; page-break-inside: avoid; background: #fff; } 
      .dept-header { background: #f1f3f4; padding: 6px 12px; font-weight: bold; margin: 15px 0 8px 0; border-left: 4px solid #1a73e8; font-size: 14px; } 
      .img-container { display: flex; flex-wrap: wrap; gap: 8px; margin-top: 10px; } 
      .report-img { width: 180px; height: 130px; object-fit: cover; border: 1px solid #ddd; } 
      table { width: 100%; border-collapse: collapse; font-size: 14px; } 
      td { padding: 4px; vertical-align: top; border-bottom: 1px solid #f8f9fa; } 
      .label { font-weight: bold; width: 110px; color: #555; }
      .text-content { padding: 6px; background: #f8f9fa; border-radius: 4px; margin-top: 4px; }
      .text-danger { color: red; font-weight: bold; }
    </style></head><body>
    <h1>${CONFIG.SYSTEM_NAME}</h1>
    <h2>สรุปรายงานการสอน (${endDateStr ? searchStart + ' ถึง ' + searchEnd : searchStart})</h2>`;
    let currentDept = "";
    filtered.forEach(r => {
      if (r.teacherDept !== currentDept) { currentDept = r.teacherDept; html += `<div class="dept-header">แผนกวิชา: ${currentDept}</div>`; }
      html += `<div class="report-card">
        <table>
          <tr><td class="label">ครูผู้สอน:</td><td>${r.name}</td><td class="label">วันที่สอน:</td><td>${r.date}</td></tr>
          <tr><td class="label">วิชา:</td><td>${r.subject}</td><td class="label">เวลาที่ส่ง:</td><td>${r.timestamp} ${r.isLate ? '<span class="text-danger">(ล่าช้า)</span>' : ''}</td></tr>
          <tr><td class="label">แผนกนักเรียน:</td><td>${r.studentDept}</td><td class="label">คาบเรียน:</td><td>${r.periods} (${r.hours || '-'} ชม.)</td></tr>
          <tr><td class="label">ระดับชั้น:</td><td colspan="3">${r.level}</td></tr>
          <tr><td class="label">เรื่องที่สอน:</td><td colspan="3">${r.topic}</td></tr>
          <tr><td class="label">นักเรียน:</td><td colspan="3">ทั้งหมด ${r.totalStud || '-'} | มา ${r.present || '-'} | ขาด ${r.absent || '-'} | ลา ${r.leave || '-'}</td></tr>
          <tr><td class="label">ปัญหา/อุปสรรค:</td><td colspan="3"><div class="text-content">${r.problems || '-'}</div></td></tr>
          <tr><td class="label">วิธีแก้ปัญหา:</td><td colspan="3"><div class="text-content">${r.solutions || '-'}</div></td></tr>
        </table>
        <div class="img-container">${r.images.map(src => `<img src="${src}" class="report-img">`).join('')}</div>
      </div>`;
    });
    const blob = HtmlService.createHtmlOutput(html + `</body></html>`).getAs('application/pdf').setName(`BUTC_Report.pdf`);
    return { success: true, data: Utilities.base64Encode(blob.getBytes()), filename: blob.getName() };
  } catch (e) { return { success: false, message: e.toString() }; }
}
