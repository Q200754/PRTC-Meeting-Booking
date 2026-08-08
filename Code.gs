// ======================================
// PRTC Meeting Room Booking System
// Google Apps Script - Code.gs
// ======================================

const SPREADSHEET_ID = "1gBKYqxArNntcKlUIv81klwJe16r4d6Dpc8EiVvCkS5E";
const SHEET_NAME = "Bookings";
const LINE_CHANNEL_ACCESS_TOKEN = "9/64Pbwp+aVE4piPq6ZAe2qPuGqPTZNBRNWJP5zdfzB5bm9DMzO4o6Hfj6RLDAUwYwTYAGLrfnUq138o1+W0FQUF7po/wuKKMbpUDyjTCVknFQZ/vaSzVYyHiJJEaVUPxta+Wxv6VeuogYc0L821aAdB04t89/1O/w1cDnyilFU=";

function doGet(e) {
  var action = e && e.parameter ? e.parameter.action : "";
  var result = {};

  try {
    if (action === "getBookings") {
      result = getBookingsSafe();
    } else if (action === "getRooms") {
      result = { ok: true, data: getMeetingRooms() };
    } else {
      result = { ok: false, errorMessage: "Invalid action parameter" };
    }
  } catch (err) {
    result = { ok: false, errorMessage: err.message };
  }

  return ContentService.createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  var result = {};
  try {
    var data = JSON.parse(e.postData.contents);
    if (data.action === "submitBooking") {
      result = submitBooking(data.payload);
    } else if (data.action === "updateStatus") {
      result = { ok: updateBookingStatus(data.payload.id, data.payload.status) };
    } else if (data.action === "deleteBooking") {
      result = { ok: deleteBooking(data.payload.id) };
    } else if (data.action === "editBooking") {
      result = { ok: editBookingData(data.payload) };
    } else if (data.action === "uploadImageAndNotify" || data.action === "uploadImageColumnK") {
      result = uploadMeetingImageToColumnK(data.payload);
    }
  } catch (err) {
    result = { ok: false, errorMessage: err.message };
  }

  return ContentService.createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

// 1. ฟังก์ชันแก้ไขข้อมูลการจอง (ไม่ส่ง LINE)
function editBookingData(payload) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEET_NAME);
  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (data[i][0] == payload.id) {
      sheet.getRange(i + 1, 2).setValue(payload.name);       // B: Name
      sheet.getRange(i + 1, 3).setValue(payload.position);   // C: Position
      sheet.getRange(i + 1, 4).setValue(payload.room);       // D: Room
      sheet.getRange(i + 1, 5).setValue(payload.building);   // E: Building
      sheet.getRange(i + 1, 6).setValue(payload.date);       // F: Date
      sheet.getRange(i + 1, 7).setValue(payload.startTime);  // G: StartTime
      sheet.getRange(i + 1, 8).setValue(payload.people);     // H: People
      sheet.getRange(i + 1, 10).setValue(payload.type);      // J: Type
      return true;
    }
  }
  return false;
}

// 2. ฟังก์ชันอัปโหลดรูปภาพลง Google Drive บันทึกลงคอลัมน์ K เป็น Plain Text (ไม่ส่ง LINE)
function uploadMeetingImageToColumnK(payload) {
  try {
    const id = payload.id;
    const base64Data = payload.base64Data;
    const fileName = payload.fileName || `Meeting_${id}.jpg`;

    let folder;
    const folders = DriveApp.getFoldersByName("PRTC_Meeting_Photos");
    if (folders.hasNext()) {
      folder = folders.next();
    } else {
      folder = DriveApp.createFolder("PRTC_Meeting_Photos");
    }

    const splitData = base64Data.split(",");
    const contentType = splitData[0].match(/:(.*?);/)[1];
    const bytes = Utilities.base64Decode(splitData[1]);
    const blob = Utilities.newBlob(bytes, contentType, fileName);

    const file = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    
    const viewUrl = file.getUrl();

    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(SHEET_NAME);
    const data = sheet.getDataRange().getValues();

    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim() === String(id).trim()) {
        const cell = sheet.getRange(i + 1, 11); // คอลัมน์ K (index 11)
        cell.setNumberFormat("@");             // กำหนดรูปแบบเป็น Plain Text
        cell.setValue(String(viewUrl));        // บันทึกลิงก์ข้อความ
        return { ok: true, status: "success", fileUrl: viewUrl };
      }
    }

    return { ok: false, errorMessage: "ไม่พบ Booking ID ใน Google Sheets" };
  } catch (err) {
    return { ok: false, errorMessage: err.message };
  }
}

// 3. ฟังก์ชันเปลี่ยนสถานะ (อนุมัติ/ไม่อนุมัติ) -> ส่งการแจ้งเตือนเข้า LINE
function updateBookingStatus(id, status) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEET_NAME);
  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (data[i][0] == id) {
      sheet.getRange(i + 1, 9).setValue(status);
      sheet.getRange(i + 1, 12).setValue(new Date());

      if (status === "อนุมัติ") sendApproveLine(data[i]);
      if (status === "ไม่อนุมัติ") sendRejectLine(data[i]);

      return true;
    }
  }
  return false;
}

// ======================================
// LINE Notification Functions
// ======================================
const THAI_MONTHS = ["มกราคม","กุมภาพันธ์","มีนาคม","เมษายน","พฤษภาคม","มิถุนายน","กรกฎาคม","สิงหาคม","กันยายน","ตุลาคม","พฤศจิกายน","ธันวาคม"];

function formatThaiDate(value) {
  if (!value) return "-";
  let day, month, year;

  if (value instanceof Date) {
    day = Number(Utilities.formatDate(value, "Asia/Bangkok", "d"));
    month = Number(Utilities.formatDate(value, "Asia/Bangkok", "M")) - 1;
    year = Number(Utilities.formatDate(value, "Asia/Bangkok", "yyyy"));
  } else {
    const datePart = String(value).split(" ")[0];
    const parts = datePart.split("-");
    if (parts.length !== 3) return String(value);
    year = Number(parts[0]);
    month = Number(parts[1]) - 1;
    day = Number(parts[2]);
  }

  if (isNaN(day) || isNaN(month) || isNaN(year)) return String(value);
  return `${day} ${THAI_MONTHS[month]} ${year + 543}`;
}

function formatThaiTime(value) {
  if (!value) return "-";
  let hour, minute;

  if (value instanceof Date) {
    hour = Number(Utilities.formatDate(value, "Asia/Bangkok", "H"));
    minute = Number(Utilities.formatDate(value, "Asia/Bangkok", "m"));
  } else {
    const str = String(value);
    const timePart = str.includes(" ") ? str.split(" ")[1] : str;
    const parts = timePart.split(":");
    hour = Number(parts[0]);
    minute = Number(parts[1]);
  }

  if (isNaN(hour) || isNaN(minute)) return String(value);
  return `${hour}:${String(minute).padStart(2, "0")} โมง`;
}

function buildBookingFlexBubble(row, isApproved) {
  const headerColor = isApproved ? "#17B37C" : "#E5484D";
  const headerIcon = isApproved ? "✅" : "❌";
  const headerTitle = isApproved ? "แจ้งอนุมัติการจองห้องประชุม" : "แจ้งไม่อนุมัติการจองห้องประชุม";
  const statusText = isApproved ? "การจองห้องประชุมได้รับการอนุมัติ" : "การจองห้องประชุมไม่ได้รับการอนุมัติ";
  const statusBg = isApproved ? "#E7F9F1" : "#FDECEC";
  const statusColor = isApproved ? "#0E7A54" : "#B3261E";

  return {
    type: "bubble",
    header: {
      type: "box",
      layout: "horizontal",
      backgroundColor: headerColor,
      paddingAll: "20px",
      spacing: "sm",
      contents: [
        { type: "text", text: headerIcon, size: "lg", color: "#FFFFFF", flex: 0 },
        { type: "text", text: headerTitle, size: "md", color: "#FFFFFF", weight: "bold", wrap: true, gravity: "center" }
      ]
    },
    body: {
      type: "box",
      layout: "vertical",
      paddingAll: "20px",
      spacing: "md",
      contents: [
        { type: "text", text: String(row[3] || "-") + (row[4] ? " (" + row[4] + ")" : ""), weight: "bold", size: "xl", color: "#1E2A44", wrap: true },
        { type: "text", text: "ผู้จอง: " + String(row[1] || "-") + (row[2] ? " (" + row[2] + ")" : ""), size: "sm", color: "#8B96B4", wrap: true },
        { type: "separator", margin: "md" },
        {
          type: "box",
          layout: "vertical",
          margin: "md",
          spacing: "sm",
          contents: [
            {
              type: "box",
              layout: "baseline",
              contents: [
                { type: "text", text: "รหัส", size: "sm", color: "#8B96B4", flex: 3 },
                { type: "text", text: String(row[0] || "-"), size: "sm", color: "#1E2A44", weight: "bold", flex: 5, wrap: true }
              ]
            },
            {
              type: "box",
              layout: "baseline",
              contents: [
                { type: "text", text: "วันที่", size: "sm", color: "#8B96B4", flex: 3 },
                { type: "text", text: formatThaiDate(row[5]), size: "sm", color: "#1E2A44", weight: "bold", flex: 5, wrap: true }
              ]
            },
            {
              type: "box",
              layout: "baseline",
              contents: [
                { type: "text", text: "เวลา", size: "sm", color: "#8B96B4", flex: 3 },
                { type: "text", text: formatThaiTime(row[6]), size: "sm", color: "#1E2A44", weight: "bold", flex: 5, wrap: true }
              ]
            }
          ]
        },
        {
          type: "box",
          layout: "vertical",
          margin: "lg",
          paddingAll: "12px",
          cornerRadius: "10px",
          backgroundColor: statusBg,
          contents: [
            { type: "text", text: headerIcon + " " + statusText, size: "sm", color: statusColor, weight: "bold", wrap: true, align: "center" }
          ]
        }
      ]
    }
  };
}

function sendApproveLine(row) {
  sendLineFlex(buildBookingFlexBubble(row, true), "✅ การจองห้องประชุม " + row[0] + " ได้รับการอนุมัติ");
}

function sendRejectLine(row) {
  sendLineFlex(buildBookingFlexBubble(row, false), "❌ การจองห้องประชุม " + row[0] + " ไม่ได้รับการอนุมัติ");
}

function sendLineFlex(bubble, altText) {
  const payload = {
    messages: [{ type: "flex", altText: altText || "แจ้งเตือนจากระบบจองห้องประชุม", contents: bubble }]
  };

  UrlFetchApp.fetch("https://api.line.me/v2/bot/message/broadcast", {
    method: "post",
    headers: {
      "Authorization": "Bearer " + LINE_CHANNEL_ACCESS_TOKEN,
      "Content-Type": "application/json"
    },
    payload: JSON.stringify(payload)
  });
}

function submitBooking(data) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(SHEET_NAME);
    if (!sheet) throw new Error("ไม่พบ Sheet: " + SHEET_NAME);

    const bookingID = generateBookingID(sheet);
    const name = String(data.name).trim();
    const position = String(data.position || "").trim();
    const room = String(data.room).trim();
    const building = String(data.building).trim();
    const date = data.date;
    const startTime = data.startTime;
    const people = Number(data.people);
    const type = String(data.type).trim();

    sheet.appendRow([
      bookingID, name, position, room, building, date, startTime, people, "รออนุมัติ", type, ""
    ]);

    return { status: "success", bookingID: bookingID };
  } catch (error) {
    return { status: "error", message: error.message };
  }
}

function generateBookingID(sheet) {
  const now = new Date();
  const day = Utilities.formatDate(now, "Asia/Bangkok", "dd");
  const month = Utilities.formatDate(now, "Asia/Bangkok", "MM");
  const prefix = "BK" + day + month;

  const data = sheet.getDataRange().getValues();
  let count = 0;

  data.forEach(row => {
    if (row[0] && row[0].toString().startsWith(prefix)) count++;
  });

  const number = String(count + 1).padStart(3, "0");
  return `${prefix}-${number}`;
}

function getBookings() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) throw new Error("ไม่พบ Sheet Bookings");

  const data = sheet.getDataRange().getValues();

  return data.map(function(row) {
    return row.map(function(cell) {
      if (cell instanceof Date) {
        const isTimeOnly = cell.getFullYear() === 1899;
        return Utilities.formatDate(
          cell,
          "Asia/Bangkok",
          isTimeOnly ? "HH:mm" : "yyyy-MM-dd HH:mm:ss"
        );
      }
      return cell;
    });
  });
}

function getBookingsSafe() {
  try {
    const data = getBookings();
    return { ok: true, rowCount: data.length, data: data };
  } catch (error) {
    return { ok: false, errorMessage: error.message, errorStack: error.stack };
  }
}

function deleteBooking(id) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEET_NAME);
  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (data[i][0] == id) {
      sheet.deleteRow(i + 1);
      return true;
    }
  }
  return false;
}

function getMeetingRooms() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName("MeetingRooms");
  return sheet.getDataRange().getValues();
}
