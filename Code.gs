/**
 * ระบบจัดการหลังบ้าน HR Portal (Leave & Shift Swap)
 * เชื่อมต่อ Google Sheets, Google Drive และ Lark Bot แบบ Interactive Cards
 */

// --- ตั้งค่า Lark Webhook URLs ---
var LARK_WEBHOOK_NEW_REQUEST = "https://open.larksuite.com/open-apis/bot/v2/hook/90fbc357-c20e-427e-8334-e5c72774156a";
var LARK_WEBHOOK_STATUS_UPDATE = "https://open.larksuite.com/open-apis/bot/v2/hook/caa1c74e-0004-47e3-93bf-326d653360c6";

function doGet() {
  return HtmlService.createTemplateFromFile('index')
      .evaluate()
      .setTitle('HR Portal - Leave & Shift Swap')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1.0')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * แปลงรูปแบบวันที่จาก YYYY-MM-DD เป็น วันที่ เดือน พ.ศ. เช่น "19 มิถุนายน 2569"
 */
function formatThaiDate(dateStr) {
  if (!dateStr) return "ไม่ระบุ";
  try {
    var dateObj;
    if (dateStr instanceof Date) {
      dateObj = dateStr;
    } else {
      var parts = dateStr.toString().split("-");
      if (parts.length === 3) {
        dateObj = new Date(parts[0], parts[1] - 1, parts[2]);
      } else {
        dateObj = new Date(dateStr);
      }
    }
    
    if (isNaN(dateObj.getTime())) return dateStr;
    
    var thaiMonths = [
      "มกราคม", "กุมภาพันธ์", "มีนาคม", "เมษายน", "พฤษภาคม", "มิถุนายน",
      "กรกฎาคม", "สิงหาคม", "กันยายน", "ตุลาคม", "พฤศจิกายน", "ธันวาคม"
    ];
    
    var day = dateObj.getDate();
    var month = thaiMonths[dateObj.getMonth()];
    var year = dateObj.getFullYear() + 543; // แปลงปี ค.ศ. เป็น พ.ศ.
    
    return day + " " + month + " " + year;
  } catch (e) {
    Logger.log("Error formatting date: " + e.toString());
    return dateStr;
  }
}

/**
 * ฟังก์ชันส่งข้อความรูปแบบ Interactive Card ไปยัง Lark Bot (ตรงตามต้นแบบภาพแนบ)
 */
function sendLarkInteractiveCard(webhookUrl, headerTitle, headerTemplate, contentMarkdown) {
  try {
    var payload = {
      "msg_type": "interactive",
      "card": {
        "config": {
          "wide_screen_mode": true
        },
        "header": {
          "title": {
            "tag": "plain_text",
            "content": headerTitle
          },
          "template": headerTemplate // ใช้โทนสี เช่น orange, green, red, purple
        },
        "elements": [
          {
            "tag": "div",
            "text": {
              "tag": "lark_md",
              "content": contentMarkdown
            }
          },
          {
            "tag": "note",
            "elements": [
              {
                "tag": "plain_text",
                "content": "ระบบลาอัตโนมัติ"
              }
            ]
          }
        ]
      }
    };
    
    var options = {
      "method": "post",
      "contentType": "application/json",
      "payload": JSON.stringify(payload)
    };
    
    UrlFetchApp.fetch(webhookUrl, options);
  } catch (e) {
    Logger.log("Lark Card Notification Error: " + e.toString());
  }
}

/**
 * ดึงข้อมูลรายชื่อพนักงานจากชีท "รายชื่อ"
 */
function getEmployeeData() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("รายชื่อ");
    if (!sheet) {
      return []; 
    }
    
    var data = sheet.getDataRange().getValues();
    var employees = [];
    
    for (var i = 1; i < data.length; i++) {
      if (data[i][0]) { 
        employees.push({
          id: data[i][0].toString().trim(),
          name: data[i][1] ? data[i][1].toString().trim() : "",
          position: data[i][2] ? data[i][2].toString().trim() : ""
        });
      }
    }
    return employees;
  } catch (e) {
    Logger.log("Error fetching employee data: " + e.toString());
    return [];
  }
}

/**
 * บันทึกการลาพักผ่อน อัปโหลดไฟล์ และส่งแจ้งเตือนการลาใหม่เป็นแบบการ์ดต้นแบบ
 */
function saveLeave(data) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("ข้อมูลการลา");
    
    if (!sheet) {
      sheet = ss.insertSheet("ข้อมูลการลา");
      sheet.appendRow([
        "วันที่ทำรายการ", "รหัสพนักงาน", "ชื่อ - สกุล", "ตำแหน่ง", 
        "วันที่ขอลา", "กะทำงาน", "ประเภทการลา", "เหตุผลการลา", 
        "ลิงก์เอกสารแนบ", "สถานะ"
      ]);
    }
    
    var fileUrl = "ไม่มีเอกสารแนบ";
    if (data.fileData && data.fileName) {
      fileUrl = uploadFileToDrive(data.fileData, data.fileName, data.leaveDate);
    }
    
    sheet.appendRow([
      new Date(),                     
      data.empId,                     
      data.name,                      
      data.position,                  
      data.leaveDate,                 
      data.shift,                     
      data.leaveType,                 
      data.reason,                    
      fileUrl,                        
      "รออนุมัติ"                      
    ]);

    // แปลงรูปแบบวันที่ให้เป็นแบบไทยแสนอบอุ่นตามภาพ "image_2d5159.png"
    var formattedDate = formatThaiDate(data.leaveDate);

    // ประกอบ Markdown โครงสร้างแบบเดียวกับแบบฟอร์มขอลาในรูป "image_2d5159.png"
    var cardMarkdown = "**👤 ข้อมูลพนักงาน**\n" +
                       "• รหัส : " + data.empId + "\n" +
                       "• ชื่อ : " + data.name + "\n" +
                       "• ตำแหน่ง : " + data.position + "\n\n" +
                       "**📅 รายละเอียดลา**\n" +
                       "• วันที่ : " + formattedDate + "\n" +
                       "• เวลา : " + data.shift + "\n" +
                       "• ประเภท : " + data.leaveType + "\n" +
                       "• เหตุผล : " + data.reason + "\n\n" +
                       "**📌 สถานะ :** 🟡 รออนุมัติ";
                       
    // ส่งข้อความการ์ดไปยัง Lark Group (Webhook 1) ด้วยธีมสีส้ม
    sendLarkInteractiveCard(LARK_WEBHOOK_NEW_REQUEST, "📝 มีคำขอลางานใหม่", "orange", cardMarkdown);
    
    return { success: true };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

/**
 * บันทึกการสลับกะการทำงาน และแจ้งเตือนด้วยดีไซน์การ์ดสวยงาม
 */
function saveSwap(data) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("ข้อมูลการสลับกะ");
    
    if (!sheet) {
      sheet = ss.insertSheet("ข้อมูลการสลับกะ");
      sheet.appendRow([
        "วันที่ทำรายการ", 
        "รหัสผู้ขอแลก", "ชื่อผู้ขอแลก", "ตำแหน่งผู้ขอแลก", "วันที่เรามาแทนเพื่อน", "กะของเรา",
        "รหัสเพื่อน", "ชื่อเพื่อน", "ตำแหน่งเพื่อน", "วันที่เพื่อนมาแทนเรา", "กะของเพื่อน",
        "เหตุผลความจำเป็น", "สถานะ"
      ]);
    }
    
    sheet.appendRow([
      new Date(),
      data.s1_empId, data.s1_name, data.s1_position, data.s1_date, data.s1_shift,
      data.s2_empId, data.s2_name, data.s2_position, data.s2_date, data.s2_shift,
      data.reason,
      "รออนุมัติ"
    ]);

    // แปลงวันที่สลับกะเป็นรูปแบบปฏิทินไทยที่สวยงาม
    var formattedS1Date = formatThaiDate(data.s1_date);
    var formattedS2Date = formatThaiDate(data.s2_date);

    // ประกอบร่างข้อความการ์ดสลับวันทำงาน
    var cardMarkdown = "**👤 ข้อมูลผู้ขอสลับวันทำงาน**\n" +
                       "• รหัส : " + data.s1_empId + "\n" +
                       "• ชื่อ : " + data.s1_name + "\n" +
                       "• ตำแหน่ง : " + data.s1_position + "\n" +
                       "• วันที่มาทำงานแทนเพื่อน : " + formattedS1Date + " (กะ: " + data.s1_shift + ")\n\n" +
                       "**👥 ข้อมูลเพื่อนที่มาแทน**\n" +
                       "• รหัส : " + data.s2_empId + "\n" +
                       "• ชื่อ : " + data.s2_name + "\n" +
                       "• ตำแหน่ง : " + data.s2_position + "\n" +
                       "• วันที่เพื่อนมาแทนเรา : " + formattedS2Date + " (กะ: " + data.s2_shift + ")\n\n" +
                       "**📝 เหตุผลความจำเป็น**\n" +
                       "• เหตุผล : " + data.reason + "\n\n" +
                       "**📌 สถานะ :** 🟡 รออนุมัติ";

    // ส่งข้อความการ์ดไปยัง Lark Group (Webhook 1) ด้วยธีมสีม่วง
    sendLarkInteractiveCard(LARK_WEBHOOK_NEW_REQUEST, "🔄 มีคำขอสลับวันทำงานใหม่", "purple", cardMarkdown);
    
    return { success: true };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

/**
 * ทริกเกอร์ On Edit (แบบติดตั้ง) ตรวจจับเมื่อผู้ดูแลระบบแก้สถานะ
 * หากเปลี่ยนสถานะใน Google Sheet เป็น "อนุมัติ" หรือ "ไม่อนุมัติ"
 * จะยิงการ์ดแจ้งผลการอนุมัติสไตล์เขียว/แดง ไปที่ Webhook ตัวที่สองทันที (ตรงตามภาพ "image_2d547b.png")
 */
function onStatusChange(e) {
  if (!e || !e.range) return;
  
  var sheet = e.range.getSheet();
  var sheetName = sheet.getName();
  var value = e.value;
  
  // ตรวจสอบว่ามีการเปลี่ยนค่าเป็น "อนุมัติ" หรือ "ไม่อนุมัติ"
  if (value === "อนุมัติ" || value === "ไม่อนุมัติ") {
    var row = e.range.getRow();
    var col = e.range.getColumn();
    
    // ดึงหัวตารางของคอลัมน์นั้นมาตรวจสอบว่าเป็นคอลัมน์ "สถานะ" หรือไม่
    var header = sheet.getRange(1, col).getValue();
    
    if (header && header.toString().trim() === "สถานะ") {
      var dataRow = sheet.getRange(row, 1, 1, sheet.getLastColumn()).getValues()[0];
      var cardMarkdown = "";
      var cardTitle = "";
      var cardTemplate = "";
      var statusEmoji = (value === "อนุมัติ") ? "🟢 อนุมัติ" : "🔴 ไม่อนุมัติ";
      
      // กำหนดสีหัวการ์ดตามผลพิจารณา (เขียวเมื่ออนุมัติ แดงเมื่อไม่อนุมัติ)
      cardTemplate = (value === "อนุมัติ") ? "green" : "red";
      
      if (sheetName === "ข้อมูลการลา") {
        // จัดรูปแบบวันที่ให้เป็นภาษาไทยพ.ศ. แสนสวยงาม
        var formattedLeaveDate = formatThaiDate(dataRow[4]);
        
        cardTitle = "📢 ผลการอนุมัติลา";
        
        // จัดวางข้อความ Markdown ให้ตรงตามรูปภาพต้นแบบ "image_2d547b.png"
        cardMarkdown = "**👤 ข้อมูลพนักงาน**\n" +
                       "• รหัส : " + dataRow[1] + "\n" +
                       "• ชื่อ : " + dataRow[2] + "\n" +
                       "• ตำแหน่ง : " + dataRow[3] + "\n\n" +
                       "**📅 รายละเอียดลา**\n" +
                       "• วันที่ : " + formattedLeaveDate + "\n" +
                       "• เวลา : " + dataRow[5] + "\n" +
                       "• ประเภท : " + dataRow[6] + "\n" +
                       "• เหตุผล : " + dataRow[7] + "\n\n" +
                       "**📢 ผลการพิจารณา :** " + statusEmoji;
                        
      } else if (sheetName === "ข้อมูลการสลับกะ") {
        var formattedS1Date = formatThaiDate(dataRow[4]);
        var formattedS2Date = formatThaiDate(dataRow[9]);
        
        cardTitle = "📢 ผลการพิจารณาสลับวันทำงาน";
        
        cardMarkdown = "**👤 ข้อมูลผู้ขอสลับวันทำงาน**\n" +
                       "• รหัส : " + dataRow[1] + "\n" +
                       "• ชื่อ : " + dataRow[2] + "\n" +
                       "• วันที่มาทำงานแทนเพื่อน : " + formattedS1Date + " (กะ: " + dataRow[5] + ")\n\n" +
                       "**👥 ข้อมูลเพื่อนที่มาแทน**\n" +
                       "• รหัส : " + dataRow[6] + "\n" +
                       "• ชื่อ : " + dataRow[7] + "\n" +
                       "• วันที่เพื่อนมาแทนเรา : " + formattedS2Date + " (กะ: " + dataRow[10] + ")\n\n" +
                       "**📢 ผลการพิจารณา :** " + statusEmoji;
      }
      
      // ส่งแจ้งเตือนแบบ Interactive Card ไปที่ Webhook ตัวที่สองทันที
      if (cardMarkdown !== "") {
        sendLarkInteractiveCard(LARK_WEBHOOK_STATUS_UPDATE, cardTitle, cardTemplate, cardMarkdown);
      }
    }
  }
}

/**
 * อัปโหลดไฟล์ขึ้น Google Drive โดยแบ่งโฟลเดอร์ย่อยตาม เดือน-ปี ที่ลา
 */
function uploadFileToDrive(base64Data, fileName, leaveDate) {
  var dateParts = leaveDate.split("-"); 
  var year = dateParts[0];
  var month = dateParts[1];
  var folderName = year + "-" + month;
  
  var parentFolderName = "HR_Leave_Documents";
  var parentFolder;
  var folders = DriveApp.getFoldersByName(parentFolderName);
  
  if (folders.hasNext()) {
    parentFolder = folders.next();
  } else {
    parentFolder = DriveApp.createFolder(parentFolderName);
  }
  
  var subFolder;
  var subFolders = parentFolder.getFoldersByName(folderName);
  if (subFolders.hasNext()) {
    subFolder = subFolders.next();
  } else {
    subFolder = parentFolder.createFolder(folderName);
  }
  
  var contentType = base64Data.substring(base64Data.indexOf(":") + 1, base64Data.indexOf(";"));
  var bytes = Utilities.base64Decode(base64Data.split(",")[1]);
  var blob = Utilities.newBlob(bytes, contentType, fileName);
  
  var file = subFolder.createFile(blob);
  
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  
  return file.getUrl();
}
