// ======================================
// PRTC Meeting Room Booking System
// Google Apps Script - Code.gs
// ======================================

// Google Sheet ID
const SPREADSHEET_ID = "1gBKYqxArNntcKlUIv81klwJe16r4d6Dpc8EiVvCkS5E";

// Sheet Name
const SHEET_NAME = "Bookings";

// LINE Messaging API Token
const LINE_CHANNEL_ACCESS_TOKEN = "9/64Pbwp+aVE4piPq6ZAe2qPuGqPTZNBRNWJP5zdfzB5bm9DMzO4o6Hfj6RLDAUwYwTYAGLrfnUq138o1+W0FQUF7po/wuKKMbpUDyjTCVknFQZ/vaSzVYyHiJJEaVUPxta+Wxv6VeuogYc0L821aAdB04t89/1O/w1cDnyilFU=";

// ======================================
// เปิดหน้า Web App (ปรับรองรับ Safari / iOS)
// ======================================

function doGet(e) {
  let templateName = "index";
  let title = "PRTC Meeting Room Booking System";

  if (e && e.parameter && e.parameter.page === "admin") {
    templateName = "admin";
    title = "PRTC Admin Dashboard";
  }

  return HtmlService
    .createHtmlOutputFromFile(templateName)
    .setTitle(title)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL); // ช่วยแก้ไขปัญหากรอบ iFrame โดนบล็อกบน Safari
}

// ======================================
// ระบบบันทึกข้อมูลการจอง
// ======================================

function submitBooking(data) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(SHEET_NAME);

    if (!sheet) {
      throw new Error("ไม่พบ Sheet: " + SHEET_NAME);
    }

    const bookingID = generateBookingID(sheet);

    const name = String(data.name).trim();
    const position = String(data.position || "").trim();
    const room = String(data.room).trim();
    const building = String(data.building).trim();
    const date = data.date;
    const startTime = data.startTime;
    const people = Number(data.people);
    const type = String(data.type).trim();

    // โครงสร้างแถวใหม่:
    // 0=รหัส, 1=ผู้จอง, 2=ตำแหน่ง/ฝ่าย, 3=ห้อง, 4=อาคาร, 5=วันที่,
    // 6=เวลาเริ่ม, 7=จำนวนคน, 8=สถานะ, 9=ประเภท, 10=ผู้อนุมัติ, 11=เวลาอนุมัติ
    sheet.appendRow([
      bookingID,
      name,
      position,
      room,
      building,
      date,
      startTime,
      people,
      "รออนุมัติ",
      type
    ]);

    return {
      status: "success",
      bookingID: bookingID
    };

  } catch (error) {
    return {
      status: "error",
      message: error.message
    };
  }
}

// สร้าง Booking ID (รูปแบบ BKDDMM-001)
function generateBookingID(sheet) {
  const now = new Date();
  const day = Utilities.formatDate(now, "Asia/Bangkok", "dd");
  const month = Utilities.formatDate(now, "Asia/Bangkok", "MM");
  const prefix = "BK" + day + month;

  const data = sheet.getDataRange().getValues();
  let count = 0;

  data.forEach(row => {
    if (row[0] && row[0].toString().startsWith(prefix)) {
      count++;
    }
  });

  const number = String(count + 1).padStart(3, "0");
  return `${prefix}-${number}`;
}

// ======================================
// Admin Dashboard & Data Handling
// ======================================

// ดึงข้อมูล Booking
function getBookings() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEET_NAME);

  if (!sheet) {
    throw new Error("ไม่พบ Sheet Bookings");
  }

  const data = sheet.getDataRange().getValues();

  // แปลงค่า Date ในทุกเซลล์ให้เป็น string ก่อนส่งกลับเพื่อป้องกันปัญหา null บน client side
  const safeData = data.map(function(row) {
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

  return safeData;
}

// เวอร์ชันดักจับ error ของ getBookings
function getBookingsSafe() {
  try {
    const data = getBookings();
    return {
      ok: true,
      rowCount: data.length,
      data: data
    };
  } catch (error) {
    return {
      ok: false,
      errorMessage: error.message,
      errorStack: error.stack
    };
  }
}

// ลบ Booking
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

// เปลี่ยนสถานะ Booking (อนุมัติ / ไม่อนุมัติ)
function updateBookingStatus(id, status) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEET_NAME);
  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (data[i][0] == id) {
      // เปลี่ยนสถานะ (Column I)
      sheet.getRange(i + 1, 9).setValue(status);

      // ผู้อนุมัติ (Column K)
      sheet.getRange(i + 1, 11).setValue("Admin");

      // เวลาอนุมัติ (Column L)
      sheet.getRange(i + 1, 12).setValue(new Date());

      // ส่ง LINE เฉพาะตอน Admin กด
      if (status === "อนุมัติ") {
        sendApproveLine(data[i]);
      }

      if (status === "ไม่อนุมัติ") {
        sendRejectLine(data[i]);
      }

      return true;
    }
  }
  return false;
}

// ======================================
// Helper Functions (การแปลงวันที่/เวลา)
// ======================================

const THAI_MONTHS = [
  "มกราคม", "กุมภาพันธ์", "มีนาคม", "เมษายน", "พฤษภาคม", "มิถุนายน",
  "กรกฎาคม", "สิงหาคม", "กันยายน", "ตุลาคม", "พฤศจิกายน", "ธันวาคม"
];

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
  const buddhistYear = year + 543;

  return `${day} ${THAI_MONTHS[month]} ${buddhistYear}`;
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
  const mm = String(minute).padStart(2, "0");

  return `${hour}:${mm} โมง`;
}

// ======================================
// LINE Flex Message Components
// ======================================

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
  const bubble = buildBookingFlexBubble(row, true);
  sendLineFlex(bubble, "✅ การจองห้องประชุม " + row[0] + " ได้รับการอนุมัติ");
}

function sendRejectLine(row) {
  const bubble = buildBookingFlexBubble(row, false);
  sendLineFlex(bubble, "❌ การจองห้องประชุม " + row[0] + " ไม่ได้รับการอนุมัติ");
}

function sendLineText(message) {
  const payload = {
    messages: [{ type: "text", text: message }]
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

// ======================================
// ดึงข้อมูลห้องประชุม & ฟังก์ชันทดสอบ
// ======================================

function getMeetingRooms() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName("MeetingRooms");
  return sheet.getDataRange().getValues();
}

function testMeetingRooms() {
  const data = getMeetingRooms();
  Logger.log(data);
}

function testGetBookings() {
  const data = getBookings();
  Logger.log("จำนวนแถวทั้งหมด (รวม header): " + data.length);
  Logger.log(data);
}

function testLINE() {
  const payload = {
    messages: [
      {
        type: "text",
        text: "✅ PRTC Meeting Room Booking System เชื่อม LINE สำเร็จ"
      }
    ]
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

function sendLineFlex(bubble, altText) {
  const payload = {
    messages: [
      {
        type: "flex",
        altText: altText || "แจ้งเตือนจากระบบจองห้องประชุม",
        contents: bubble
      }
    ]
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
