/**
 * ระบบเกาะกล้วยหอมจอมพลัง 🍌 - โค้ดสำหรับ Google Apps Script (GAS) (v2.1)
 * คุณครูนำโค้ดนี้ไปเปิดใน Extension > Apps Script ของ Google Sheets ของคุณครู
 * จากนั้นกด Deploy > New Deployment > เลือกเป็น Web App > ตั้งค่า Who has access เป็น Anyone
 */

function doGet(e) {
  var params = e.parameter;
  var action = params.action || "";
  var studentId = params.id || "";
  
  // 1. ตรวจสอบการดึงข้อมูลนักเรียนแบบเรียลไทม์ (GET action=getStudent&id=XXXX)
  if (action === "getStudent") {
    return ContentService.createTextOutput(JSON.stringify(getStudentData(studentId)))
      .setMimeType(ContentService.MimeType.JSON);
  }
  
  // 2. การบันทึกคะแนนปกติ (สแกน QR Code หรือส่งคะแนนท้ายเกม)
  var studentName = params.name || "ไม่ระบุชื่อ";
  var room = params.room || "ทั่วไป";
  var num = params.num || "";
  var score = params.score || "0";
  var taskName = params.task || "เกาะกล้วยหอมจอมพลัง";
  
  if (!studentId) {
    return HtmlService.createHtmlOutput(getHtmlResponse("เกิดข้อผิดพลาด", "ไม่พบรหัสนักเรียน กรุณาลองสแกนใหม่อีกครั้งนะจ๊ะ! ❌", false));
  }
  
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var cleanRoom = room.replace(/\//g, "-");
    var sheetName = cleanRoom + " - " + taskName.substring(0, 15);
    var sheet = ss.getSheetByName(sheetName);
    
    if (!sheet) {
      sheet = ss.insertSheet(sheetName);
      sheet.appendRow(["วัน-เวลาที่ส่ง", "รหัสประจำตัว", "ชื่อ-นามสกุล", "เลขที่", "ห้อง", "คะแนนเต็ม 6", "สถานะการผ่านด่าน"]);
      sheet.getRange("A1:G1").setFontWeight("bold").setBackground("#fef08a").setHorizontalAlignment("center");
      sheet.setColumnWidth(1, 180);
      sheet.setColumnWidth(2, 100);
      sheet.setColumnWidth(3, 200);
      sheet.setColumnWidth(4, 80);
      sheet.setColumnWidth(5, 80);
      sheet.setColumnWidth(6, 100);
      sheet.setColumnWidth(7, 150);
    }
    
    var data = sheet.getDataRange().getValues();
    var foundIndex = -1;
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][1]).trim() === String(studentId).trim()) {
        foundIndex = i + 1;
        break;
      }
    }
    
    var timestamp = new Date();
    var formattedDate = Utilities.formatDate(timestamp, Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss");
    var status = (parseInt(score) >= 6) ? "ผ่านด่านครบสมบูรณ์ 🏆" : "เรียนรู้ระหว่างด่าน 🍌";
    
    if (foundIndex !== -1) {
      sheet.getRange(foundIndex, 1).setValue(formattedDate);
      sheet.getRange(foundIndex, 3).setValue(studentName);
      sheet.getRange(foundIndex, 4).setValue(num);
      sheet.getRange(foundIndex, 5).setValue(room);
      sheet.getRange(foundIndex, 6).setValue(score);
      sheet.getRange(foundIndex, 7).setValue(status);
    } else {
      sheet.appendRow([formattedDate, studentId, studentName, num, room, score, status]);
    }
    
    var numRows = sheet.getLastRow();
    if (numRows > 1) {
      var rangeToSort = sheet.getRange(2, 1, numRows - 1, 7);
      rangeToSort.sort({column: 4, ascending: true});
    }
    
    var messageHtml = "ระบบได้บันทึกคะแนนของ <strong>" + studentName + "</strong> (เลขที่ " + num + " ห้อง " + room + ") เรียบร้อยแล้วจ้า!<br>" +
                      "คะแนนที่ได้: <strong style='font-size: 24px; color: #eab308;'>" + score + " / 6</strong> คะแนน ⭐<br>" +
                      "สเปรดชีตเป้าหมาย: <code>" + sheetName + "</code> (ระบบบันทึกคะแนนเดิมปลอดภัยไม่เสียหายแน่นอน 💖)";
                      
    return HtmlService.createHtmlOutput(getHtmlResponse("บันทึกคะแนนสำเร็จ! 🎉", messageHtml, true));
    
  } catch (error) {
    return HtmlService.createHtmlOutput(getHtmlResponse("เกิดข้อผิดพลาดในการบันทึก", "ข้อผิดพลาด: " + error.toString() + " ❌", false));
  }
}

// ฟังก์ชันดึงข้อมูลนักเรียนจาก Google Sheets เพื่อนำรูปและข้อมูลไปแสดงในเกมแบบเรียลไทม์
function getStudentData(studentId) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    // ค้นหาชีทที่มีคำว่า "รายชื่อ" หรือชีทแรกสุด
    var sheet = ss.getSheetByName("รายชื่อนักเรียน") || ss.getSheets()[0];
    if (!sheet) {
      return { success: false, message: "ไม่พบชีทรายชื่อนักเรียน" };
    }
    
    var range = sheet.getDataRange();
    var values = range.getValues();
    if (values.length <= 1) {
      return { success: false, message: "ชีทไม่มีข้อมูล" };
    }
    
    // หาคอลัมน์จากแถวหัวตาราง (Header)
    var headers = values[5] || values[0]; // สมมติหัวตารางอยู่แถว 6 (ป้อนจาก Excel) หรือแถวแรก
    // ค้นหาตำแหน่งดัชนีหัวตารางที่ยืดหยุ่น
    var colId = -1, colRoom = -1, colNum = -1, colPrefix = -1, colName = -1, colSurname = -1, colPhoto = -1;
    
    for (var i = 0; i < headers.length; i++) {
      var h = String(headers[i]).trim();
      if (h.indexOf("รหัสนักศึกษา") > -1 || h.indexOf("รหัสประจำตัว") > -1 || h.indexOf("รหัส") > -1) colId = i;
      else if (h.indexOf("ชั้น/ห้อง") > -1 || h.indexOf("ห้อง") > -1) colRoom = i;
      else if (h.indexOf("เลขที่") > -1) colNum = i;
      else if (h.indexOf("คำนำหน้า") > -1) colPrefix = i;
      else if (h.indexOf("ชื่อ") === 0) colName = i; // ชื่อต้น
      else if (h.indexOf("นามสกุล") > -1) colSurname = i;
      else if (h.indexOf("รูปภาพ") > -1 || h.indexOf("รูป") > -1 || h.indexOf("Photo") > -1) colPhoto = i;
    }
    
    // หากหาคอลัมน์ไม่เจอ ให้ใช้ค่าเริ่มต้นในการเดา (Fallback Index)
    if (colId === -1) colId = 5;       // คอลัมน์ F (รหัสนักศึกษา)
    if (colRoom === -1) colRoom = 1;   // คอลัมน์ B (ชั้น/ห้อง)
    if (colNum === -1) colNum = 2;     // คอลัมน์ C (เลขที่)
    if (colPrefix === -1) colPrefix = 9; // คอลัมน์ J (คำนำหน้า)
    if (colName === -1) colName = 10;   // คอลัมน์ K (ชื่อ)
    if (colSurname === -1) colSurname = 11; // คอลัมน์ L (นามสกุล)
    // คอลัมน์รูปภาพสามารถเพิ่มต่อท้ายสุดของคุณครูได้ (เช่น คอลัมน์ FG หรือท้ายสุด)
    if (colPhoto === -1) {
      // ค้นหาจากหัวตารางที่มีคำว่ารูป
      for (var colIdx = 0; colIdx < headers.length; colIdx++) {
        if (String(headers[colIdx]).includes("รูป")) {
          colPhoto = colIdx;
          break;
        }
      }
    }
    
    // วนลูปค้นหารายชื่อจากแถวถัดจากหัวตาราง
    var startRow = (values[5] && String(values[5][0]).includes("ลำดับ")) ? 6 : 1;
    for (var rowIdx = startRow; rowIdx < values.length; rowIdx++) {
      var row = values[rowIdx];
      var rowId = String(row[colId]).trim();
      
      if (rowId === String(studentId).trim()) {
        var prefix = colPrefix !== -1 ? String(row[colPrefix]) : "";
        var nameStr = colName !== -1 ? String(row[colName]) : "";
        var surname = colSurname !== -1 ? String(row[colSurname]) : "";
        var fullName = (prefix + nameStr + " " + surname).trim();
        
        var gender = "male";
        if (prefix.indexOf("หญิง") > -1 || prefix.indexOf("สาว") > -1 || prefix.indexOf("ด.ญ.") > -1) {
          gender = "female";
        }
        
        var photoUrl = (colPhoto !== -1 && colPhoto < row.length) ? String(row[colPhoto]).trim() : "";
        
        return {
          success: true,
          student: {
            id: rowId,
            name: fullName,
            room: colRoom !== -1 ? String(row[colRoom]).trim() : "",
            num: colNum !== -1 ? String(row[colNum]).trim() : "",
            gender: gender,
            photo: photoUrl || null // คืนลิงก์รูปภาพออนไลน์
          }
        };
      }
    }
    
    return { success: false, message: "ไม่พบข้อมูลนักเรียนรหัสนี้ในชีท" };
  } catch (err) {
    return { success: false, message: "เกิดข้อผิดพลาด: " + err.toString() };
  }
}

// ฟังก์ชันสร้างหน้าเว็บตอบรับสไตล์การ์ตูนสำหรับแสดงผลบนมือถือเด็ก
function getHtmlResponse(title, message, isSuccess) {
  var bgColor = isSuccess ? "radial-gradient(circle, #fef9c3 0%, #fef08a 100%)" : "radial-gradient(circle, #fee2e2 0%, #fecaca 100%)";
  var icon = isSuccess ? "🍌🏆✨" : "⚠️❌";
  
  return '<!DOCTYPE html>' +
    '<html>' +
    '<head>' +
    '  <meta charset="UTF-8">' +
    '  <meta name="viewport" content="width=device-width, initial-scale=1.0">' +
    '  <title>' + title + '</title>' +
    '  <script src="https://cdn.tailwindcss.com"></script>' +
    '  <link rel="preconnect" href="https://fonts.googleapis.com">' +
    '  <link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>' +
    '  <link href="https://fonts.googleapis.com/css2?family=Mali:wght@400;600;700&display=swap" rel="stylesheet">' +
    '  <style>' +
    '    body { font-family: "Mali", cursive; background: ' + bgColor + '; }' +
    '    .floating { animation: float 3s ease-in-out infinite; }' +
    '    @keyframes float {' +
    '      0%, 100% { transform: translateY(0); }' +
    '      50% { transform: translateY(-10px); }' +
    '    }' +
    '  </style>' +
    '</head>' +
    '<body class="min-h-screen flex items-center justify-center p-4">' +
    '  <div class="w-full max-w-md bg-white/95 rounded-3xl border-8 border-yellow-400 shadow-2xl p-6 text-center">' +
    '    <div class="floating text-6xl mb-4">' + icon + '</div>' +
    '    <h1 class="text-2xl font-bold text-yellow-800 mb-4">' + title + '</h1>' +
    '    <div class="text-gray-700 text-sm leading-relaxed mb-6 bg-yellow-50/50 p-4 rounded-2xl border border-yellow-100">' +
    '      ' + message + '' +
    '    </div>' +
    '    <button onclick="window.close()" class="bg-green-500 hover:bg-green-600 text-white font-bold py-3 px-8 rounded-full shadow-lg transition-all text-md border-b-4 border-green-700 active:border-b-0">' +
    '      ปิดหน้านี้ 🚪' +
    '    </button>' +
    '    <p class="text-[10px] text-gray-400 mt-4">เกาะกล้วยหอมจอมพลัง v2.1 - พัฒนาโดยครูกิ่งร่วมกับ AI</p>' +
    '  </div>' +
    '</body>' +
    '</html>';
}
