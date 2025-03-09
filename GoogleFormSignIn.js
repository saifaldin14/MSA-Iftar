// MODIFY THIS WHEN MODIFYING THIS AND REPLOYING
const DEPLOYMENT_ID = 'AKfycbxnAudIEHBRk92m3vIq-GQdXrLpB2bzThSFp5YzOvwDQQV_DA0Cp32PWCTcKqKigAirag';
const TIMESTAMP = "Timestamp", EMAIL = "UWaterloo Email",
      TICKET_CODE = "Ticket Code", STATUS = "Status",
      VALID = "Valid", USED = "Used";

function onFormSubmit(e) {
  const sheet = SpreadsheetApp.getActiveSheet();
  addExtraColumns();
  const row = e.range.getRow();
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0]; // Get the header row
  
  // Create a map of column names to their indices
  const columnMap = {};
  headers.forEach((header, index) => {
    columnMap[header.trim()] = index;
  });
  
  const timestamp = e.values[columnMap[TIMESTAMP]];
  const email = e.values[columnMap[EMAIL]] ? e.values[columnMap[EMAIL]].trim() : "";
  
  const submittedDate = new Date(timestamp);
  const submittedDay   = submittedDate.getDate();
  const submittedMonth = submittedDate.getMonth();
  const submittedYear  = submittedDate.getFullYear();
  const formattedDate = `${submittedYear}-${submittedMonth}-${submittedDay}`;

  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    const rowTimestamp = data[i][columnMap[TIMESTAMP]];
    const rowEmail = data[i][columnMap[EMAIL]];

    const rowDate = new Date(rowTimestamp);
    const rowDay   = rowDate.getDate();
    const rowMonth = rowDate.getMonth();
    const rowYear  = rowDate.getFullYear();

    if (
      rowDay === submittedDay &&
      rowMonth === submittedMonth &&
      rowYear === submittedYear &&
      rowEmail === email &&
      i !== (row - 1)
    ) {
      // Send an email letting them know they already submitted today
      MailApp.sendEmail({
        to: email,
        subject: "Duplicate Submission for Today 🚫",
        htmlBody: `
          <p>Hi there,</p>
          <p>⚠️ Our records show that you have already submitted today.</p>
          <p>Each day is limited to one submission. Please try again tomorrow.</p>
          <p>Thank you for your interest! 🙌</p>
        `
      });      
      return;
    }
  }
  
  const randomString = Math.random().toString(36).substring(2, 10);
  const uniqueCode = `${email}_${submittedDay}-${submittedMonth+1}-${submittedYear}_${randomString}`;

  const ticketPageUrl = `https://script.google.com/macros/s/${DEPLOYMENT_ID}/exec?code=${encodeURIComponent(uniqueCode)}&date=${encodeURIComponent(submittedDate.toISOString())}`;
  const qrUrl = `https://quickchart.io/qr?format=png&margin=1&size=300&text=${encodeURIComponent(ticketPageUrl)}`;

  // Save the ticket details in the Sheet
  sheet.getRange(row, columnMap[TICKET_CODE] + 1).setValue(uniqueCode);
  sheet.getRange(row, columnMap[STATUS] + 1).setValue(VALID); // Default status as "Valid"

  MailApp.sendEmail({
    to: email,
    subject: "Your One-Time Use QR Code",
    htmlBody: `
      <p>Hi there,</p>
      <p>✅ Your ticket for <strong>${formattedDate}</strong> has been successfully generated!</p>
      <p>Please present this QR code at the entrance:</p>
      <img src="${qrUrl}" alt="Your QR Code" />
      <hr>
      <p>🎟️ <strong>Ticket Details:</strong></p>
      <ul>
        <li>📧 <strong>Email:</strong> ${email}</li>
        <li>📅 <strong>Date:</strong> ${formattedDate}</li>
        <li>🎉 <strong>Event:</strong> MSA Iftar</li>
        <li>ℹ️ <strong>Instructions:</strong> Show this QR code at the entrance.</li>
      </ul>
      <p>We look forward to seeing you! 🎊</p>
    `
  });
}

// Web app function that displays the ticket and a check-in button
function doGet(e) {  
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const code = e.parameter.code;
  if (!code) {
    return HtmlService.createHtmlOutput("❌ Invalid or missing ticket code.");
  }
  
  // Parse the submitted date parameter
  const submittedDate = new Date(e.parameter.date);
  
  // Build column mapping from header row
  const data = sheet.getDataRange().getValues();
  let columnMap = {};
  const headers = data[0];
  headers.forEach((header, index) => {
    columnMap[header.trim()] = index;
  });
  
  // Locate the ticket row in the sheet
  let ticketRow = -1;
  let email = "";
  let status = "";
  for (let i = 1; i < data.length; i++) {
    if (data[i][columnMap[TICKET_CODE]].trim() === code.trim()) {
      ticketRow = i;
      email = data[i][columnMap[EMAIL]];
      status = data[i][columnMap[STATUS]].trim();
      break;
    }
  }
  
  if (ticketRow === -1) {
    return HtmlService.createHtmlOutput("❌ Ticket not found.");
  }
  
  // Determine if the ticket is valid (status is "Valid" and the date matches today)
  const currentDate = new Date();
  const isSameDay = currentDate.toISOString().split("T")[0] === submittedDate.toISOString().split("T")[0];
  const isValid = (status === VALID && isSameDay);
  
  // Build the ticket page HTML
  let ticketHtml = `
    <style>
      body {
        font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
        background-color: #f2f2f2;
        margin: 0;
        padding: 20px;
      }
      .ticket {
        max-width: 400px;
        margin: 40px auto;
        padding: 20px;
        background-color: #fff;
        border-radius: 8px;
        box-shadow: 0 2px 8px rgba(0, 0, 0, 0.15);
        text-align: center;
      }
      h1 {
        font-size: 24px;
        margin-bottom: 20px;
        color: #333;
      }
      p {
        font-size: 18px;
        margin: 8px 0;
      }
      .status {
        font-size: 20px;
        font-weight: bold;
        color: #2ecc71; /* Green for valid */
      }
      .status.invalid {
        color: #e74c3c; /* Red for invalid */
      }
      .checkin-button {
        display: block;
        width: 100%;
        padding: 15px;
        margin-top: 20px;
        font-size: 18px;
        color: #fff;
        background-color: #3498db;
        border: none;
        border-radius: 5px;
        cursor: pointer;
      }
      .checkin-button:disabled {
        background-color: #95a5a6;
        cursor: not-allowed;
      }
      #result {
        margin-top: 15px;
        font-size: 18px;
        font-weight: bold;
        color: #34495e;
      }
    </style>
    <div class="ticket">
      <h1>🎟️ IFTAR TICKET 🎟️</h1>
      <p><strong>Unique Code:</strong> ${code}</p>
      <p><strong>Email:</strong> ${email}</p>
      <p><strong>Date:</strong> ${submittedDate.toDateString()}</p>
      <p><strong>Status:</strong> <span class="status ${isValid ? '' : 'invalid'}">${status}</span></p>
  `;

  if (isValid) {
    ticketHtml += `
      <button class="checkin-button" onclick="checkIn()">Check In</button>
      <p id="result"></p>
      <script>
        function checkIn() {
          google.script.run.withSuccessHandler(function(msg) {
            document.getElementById("result").innerText = msg;
            document.querySelector(".checkin-button").disabled = true;
          }).markTicketUsed("${code}");
        }
      </script>
    `;
  } else {
    ticketHtml += `<p style="color:#e74c3c; font-size: 18px; font-weight: bold;">Ticket is already used or invalid.</p>`;
  }

  ticketHtml += `</div>`;

  return HtmlService.createHtmlOutput(ticketHtml).setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// Server-side function to mark the ticket as used
function markTicketUsed(ticketCode) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const data = sheet.getDataRange().getValues();
  let columnMap = {};
  const headers = data[0];
  headers.forEach((header, index) => {
    columnMap[header.trim()] = index;
  });
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][columnMap[TICKET_CODE]].trim() === ticketCode.trim()) {
      if (data[i][columnMap[STATUS]].trim() === VALID) {
        sheet.getRange(i + 1, columnMap[STATUS] + 1).setValue(USED);
        return "Ticket successfully checked in!";
      } else {
        return "Ticket is already used.";
      }
    }
  }
  return "Ticket not found.";
}

function addExtraColumns() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const requiredColumns = [TICKET_CODE, STATUS];
  let addedColumns = 0;
  
  requiredColumns.forEach(colName => {
    if (!headers.includes(colName)) {
      sheet.getRange(1, headers.length + 1 + addedColumns).setValue(colName);
      addedColumns++;
    }
  });
  
  if (addedColumns > 0) {
    SpreadsheetApp.flush();
  }
}

