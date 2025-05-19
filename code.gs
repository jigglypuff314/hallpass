const SHEET_NAME = 'Form';
const END_PASS_SHEET_NAME = 'EndPass';
const EMAIL_SUBJECT = 'Your Hall Pass Status';
const MAX_PASSES_PER_BLOCK = 2; // Define the maximum number of passes per time block
const END_PASS_FORM_URL = 'https://forms.gle/yxrgRf9rV48SNWmc9';
const MAX_PASS_DURATION_MINUTES = 20; // Maximum duration for automatic expiration
const VISUAL_PASS_DURATION_MINUTES = 5; // Default visual duration shown in the email
const EXEMPT_VISUAL_DURATION_MINUTES = 60; // Longer visual duration for exempt students
const FRIENDS_SHEET_NAME = 'Friends'; // Define the Friends sheet name

function extractNameFromEmail(email) {
  const parts = email.split('@')[0].split('.');
  return parts.map(p => p.charAt(0).toUpperCase() + p.slice(1)).join(' ');
}

// For ACTIVE Hall Passes
function generatePassHtml(bg, name, from, to, expiry) {
  const destinationDisplay = destinationTitleCase(to);
  // Format the expiry date and time for the vClock URL, ensuring local time
  const year = expiry.getFullYear();
  const month = String(expiry.getMonth() + 1).padStart(2, '0'); // Month is 0-indexed
  const day = String(expiry.getDate()).padStart(2, '0');
  const hours = String(expiry.getHours()).padStart(2, '0');
  const minutes = String(expiry.getMinutes()).padStart(2, '0');
  const seconds = String(expiry.getSeconds()).padStart(2, '0');

  const timerLink = `https://vclock.com/timer/#date=${year}-${month}-${day}T${hours}:${minutes}:${seconds}&showmessage=1&message=${encodeURIComponent('Your Hall Pass Has Expired')}&sound=twinkle&loop=0`;

  const expiryDateFormatted = expiry.toLocaleDateString();
  const expiryTimeFormatted = expiry.toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' });

  return `
  <div style="font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; color: #333; background-color: #f4f7f6; padding: 20px; margin: 0 auto; max-width: 600px;">
    <div style="background-color: #ffffff; border: 1px solid #e0e0e0; border-radius: 10px; box-shadow: 0 4px 12px rgba(0,0,0,0.08);">
      <div style="background-color: ${bg}; color: white; padding: 25px; text-align: center; border-top-left-radius: 9px; border-top-right-radius: 9px;">
        <h1 style="margin: 0; font-size: 30px; font-weight: bold;">HALL PASS</h1>
      </div>
      <div style="padding: 25px 30px; font-size: 17px; text-align: center;">
        <p style="margin: 12px 0; line-height: 1.6;"><strong>Name:</strong> ${name}</p>
        <hr style="border: 0; border-top: 1px solid #eeeeee; margin: 15px 0;">
        <p style="margin: 12px 0; line-height: 1.6;"><strong>From:</strong> Room ${from}</p>
        <hr style="border: 0; border-top: 1px solid #eeeeee; margin: 15px 0;">
        <p style="margin: 12px 0; line-height: 1.6;"><strong>To:</strong> ${destinationDisplay}</p>
        <hr style="border: 0; border-top: 1px solid #eeeeee; margin: 15px 0;">
        <p style="margin: 12px 0; line-height: 1.6;"><strong>Expected to Expire On:</strong> ${expiryDateFormatted} at ${expiryTimeFormatted}</p>
        <p style="margin-top: 20px; text-align: center;">
          <a href="${timerLink}" style="display: inline-block; background-color: ${bg}; color: white; padding: 10px 18px; text-decoration: none; border-radius: 5px; font-weight: bold; font-size: 16px; margin-right: 10px;" target="_blank">View Timer</a>
          <a href="${END_PASS_FORM_URL}" style="display: inline-block; background-color: #f44336; color: white; padding: 10px 18px; text-decoration: none; border-radius: 5px; font-weight: bold; font-size: 16px;" target="_blank">End Pass</a>
        </p>
      </div>
    </div>
    <div style="text-align: center; margin-top: 25px; font-size: 12px; color: #888;">
      This is an automated email. Do not reply to this email.
    </div>
  </div>`;
}

// For WAITLISTED Hall Passes
function generateWaitlistHtml(name, from, to, position, unlockTime) {
  const destinationDisplay = destinationTitleCase(to);
  const now = new Date();
  const diff = unlockTime.getTime() - now.getTime();
  let remainingTime = "Your pass may be active soon. Check for a new email.";
  let timerColor = "green";

  if (diff > 0) {
    const mins = Math.floor(diff / 60000);
    const secs = Math.floor((diff % 60000) / 1000);
    remainingTime = `${mins}m ${String(secs).padStart(2, '0')}s`;
    timerColor = "#333";
  }

  const unlockDateFormatted = unlockTime.toLocaleDateString();
  const unlockTimeFormatted = unlockTime.toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' });

  return `
  <div style="font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; color: #333; background-color: #f4f7f6; padding: 20px; margin: 0 auto; max-width: 600px;">
    <div style="background-color: #ffffff; border: 1px solid #e0e0e0; border-radius: 10px; box-shadow: 0 4px 12px rgba(0,0,0,0.08);">
      <div style="background-color: orange; color: white; padding: 25px; text-align: center; border-top-left-radius: 9px; border-top-right-radius: 9px;">
        <h1 style="margin: 0; font-size: 30px; font-weight: bold;">HALL PASS - WAITLISTED</h1>
      </div>
      <div style="padding: 25px 30px; font-size: 17px;">
        <p style="margin: 12px 0; line-height: 1.6;"><strong>Name:</strong> ${name}</p>
        <hr style="border: 0; border-top: 1px solid #eeeeee; margin: 15px 0;">
        <p style="margin: 12px 0; line-height: 1.6;"><strong>From:</strong> Room ${from}</p>
        <hr style="border: 0; border-top: 1px solid #eeeeee; margin: 15px 0;">
        <p style="margin: 12px 0; line-height: 1.6;"><strong>To:</strong> ${destinationDisplay}</p>
        <hr style="border: 0; border-top: 1px solid #eeeeee; margin: 15px 0;">
        <p style="margin: 12px 0; line-height: 1.6;"><strong>Position in line:</strong> <span style="font-weight: bold; color: #E67E22;">${position}</span></p>
        <hr style="border: 0; border-top: 1px solid #eeeeee; margin: 15px 0;">
        <p style="margin: 12px 0; line-height: 1.6;"><strong>Estimated Activation:</strong> ${unlockDateFormatted} at ${unlockTimeFormatted}</p>
        <p style="margin: 20px 0 10px 0; text-align: center;">
          <strong style="font-size: 18px; color: orange; display: block; margin-bottom: 5px;">Estimated Wait Time:</strong>
          <span id="timer" style="font-weight: bold; font-size: 28px; color: ${timerColor};">${remainingTime}</span>
        </p>
      </div>
    </div>
    <div style="text-align: center; margin-top: 25px; font-size: 12px; color: #888;">
      This is an automated email. Do not reply to this email.
    </div>
  </div>`;
}

function createHtml(bgColor, message) {
  return `<div style="font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; color: #333; background-color: #f4f7f6; padding: 20px; margin: 0 auto; max-width: 600px;">
    <div style="background-color: #ffffff; border: 1px solid #e0e0e0; border-radius: 10px; box-shadow: 0 4px 12px rgba(0,0,0,0.08);">
      <div style="background-color: ${bgColor}; color: white; padding: 25px; text-align: center; border-top-left-radius: 9px; border-top-right-radius: 9px;">
        <h1 style="margin: 0; font-size: 30px; font-weight: bold;">HALL PASS</h1>
      </div>
      <div style="padding: 25px 30px; font-size: 17px; text-align: center;">
        <p style="margin: 12px 0; line-height: 1.6;">${message}</p>
      </div>
    </div>
    <div style="text-align: center; margin-top: 25px; font-size: 12px; color: #888;">
      This is an automated email. Do not reply to this email.
    </div>
  </div>`;
}

function destinationTitleCase(str) {
  return str.toLowerCase().split(' ').map(function (word) {
    return (word.charAt(0).toUpperCase() + word.slice(1));
  }).join(' ');
}

function isExemptStudent(email) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const exemptSheet = ss.getSheetByName('Exempt');
  if (!exemptSheet) {
    Logger.log('Warning: "Exempt" sheet not found.');
    return false;
  }
  const exemptData = exemptSheet.getDataRange().getValues();
  // Assuming email addresses are in the second column (index 1)
  for (let i = 0; i < exemptData.length; i++) {
    if (exemptData[i][1] === email) {
      return true;
    }
  }
  return false;
}

function onFormSubmit(e) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);

  const row = e.range.getRow();
  const rowData = sheet.getRange(row, 1, 1, sheet.getLastColumn()).getValues()[0];

  const [timestamp, email, , roomNumber, destinationRaw] = rowData;
  const destination = destinationRaw.toLowerCase();
  const name = extractNameFromEmail(email);
  const timeSubmitted = new Date(timestamp);

  // --- Check if the student is exempt ---
  if (isExemptStudent(email)) {
    const actualExpiry = new Date(timeSubmitted.getTime() + MAX_PASS_DURATION_MINUTES * 60000);
    const visualExpiry = new Date(timeSubmitted.getTime() + EXEMPT_VISUAL_DURATION_MINUTES * 60000);
    const emailBody = generatePassHtml("green", name, roomNumber, destination, visualExpiry);
    sheet.getRange(row, 12).setValue("ACTIVE"); // Column L - Status
    sheet.getRange(row, 3).setValue(name);    // Column C - Name
    // Send email as BCC
    GmailApp.sendEmail('', EMAIL_SUBJECT, 'Your Hall Pass is Approved (Exempt)', {
      htmlBody: emailBody,
      bcc: email // Send to the student as BCC
    });
    return; // Stop the rest of the onFormSubmit function for exempt students
  }

  const currentTime = new Date();
  const hour = timeSubmitted.getHours();
  const timeBlock = hour >= 8 && hour < 12 ? "AM" : hour >= 12 && hour <= 23 ? "PM" : "OUT";

  // Visual duration for the email
  let visualDuration = VISUAL_PASS_DURATION_MINUTES;
  if (destination.includes("nurse")) visualDuration = 8;
  else if (destination.includes("locker")) visualDuration = 5;
  else if (destination.includes("bathroom")) visualDuration = 5;
  else if (destination.includes("guidance")) visualDuration = 12;

  const allData = sheet.getDataRange().getValues();
  const studentPasses = allData.filter(r => r[1] === email);

  // Check if the student already has an active pass
  const hasActivePass = studentPasses.some(r => r[11] === "ACTIVE");

  // Count the number of active, expired, or waitlisted passes for the current student in the current time block
  const usedPassesInBlock = studentPasses.filter(r => {
    const rHour = new Date(r[0]).getHours();
    const rTimeBlock = rHour >= 8 && rHour < 12 ? "AM" : rHour >= 12 && rHour <= 23 ? "PM" : "OUT";
    const rStatus = r[11];
    return rTimeBlock === timeBlock && (rStatus === "ACTIVE" || rStatus === "EXPIRED" || rStatus === "WAITLISTED");
  }).length;

  let status = "REJECTED";
  let emailBody = "";
  let waitlistPosition = 0;

  // Check if there is already an active pass for this classroom
  const activePassInClassroom = allData.some(r => r[3] === roomNumber && r[11] === "ACTIVE");

  if (hasActivePass) {
    emailBody = createHtml("red", `You already have an active hall pass. Please wait until it is marked as used.`);
  } else if (timeBlock === "OUT") {
    emailBody = createHtml("red", `Hall passes are unavailable at this time.`);
  } else if (usedPassesInBlock >= MAX_PASSES_PER_BLOCK) {
    emailBody = createHtml("red", `You have already reached your limit of ${MAX_PASSES_PER_BLOCK} hall passes this ${timeBlock === "AM" ? "morning" : "afternoon"}.`);
  } else if (activePassInClassroom) {
    status = "WAITLISTED";
    emailBody = generateWaitlistHtml(name, roomNumber, destination, 'N/A', new Date(Date.now() + 300000)); // Approximate unlock time
  } else {
    status = "ACTIVE";
    // Actual expiry for the sheet will be MAX_PASS_DURATION_MINUTES
    const actualExpiry = new Date(timeSubmitted.getTime() + MAX_PASS_DURATION_MINUTES * 60000);
    // Visual expiry for the email will be visualDuration
    const visualExpiry = new Date(timeSubmitted.getTime() + visualDuration * 60000);
    emailBody = generatePassHtml("green", name, roomNumber, destination, visualExpiry);
  }

  // Update Sheet
  sheet.getRange(row, 12).setValue(status); // Column L - Status
  sheet.getRange(row, 3).setValue(name);    // Column C - Name

  // Send email as BCC
  GmailApp.sendEmail('', EMAIL_SUBJECT, 'Your hall pass info', {
    htmlBody: emailBody,
    bcc: email // Send to the student as BCC
  });
}

function onEndPassFormSubmit(e) {
  Logger.log("onEndPassFormSubmit triggered with event:");
  Logger.log(JSON.stringify(e)); // Log the entire event object

  const endPassSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(END_PASS_SHEET_NAME);
  try {
    const row = e.range.getRow();
    const rowData = endPassSheet.getRange(row, 1, 1, endPassSheet.getLastColumn()).getValues()[0];
    const [endPassTimestamp, studentEmail] = rowData;
    const formSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
    const formData = formSheet.getDataRange().getValues();

    let passesEnded = 0;

    Logger.log(`Processing End Pass for email: ${studentEmail}`);

    for (let i = 1; i < formData.length; i++) { // Start from the second row to skip headers
      const rowEmail = formData[i][1];
      const status = formData[i][11];

      Logger.log(`Checking row ${i + 1}, Email: ${rowEmail}, Status: ${status}`);

      if (rowEmail === studentEmail && (status === "ACTIVE" || status === "WAITLISTED")) {
        // Get the original timestamp from column A
        const originalTimestamp = new Date(formData[i][0]);
        // Calculate the duration in milliseconds
        const durationMillis = endPassTimestamp.getTime() - originalTimestamp.getTime();

        // Convert milliseconds to hours, minutes, and seconds
        const hours = Math.floor(durationMillis / (1000 * 60 * 60));
        const minutes = Math.floor((durationMillis % (1000 * 60 * 60)) / (1000 * 60));
        const seconds = Math.floor((durationMillis % (1000 * 60)) / 1000);

        // Format the duration as HH:MM:SS
        const formattedDuration = `${String(hours).padStart(2, '0')}:${String(minutes).padStart(2, '0')}:${String(seconds).padStart(2, '0')}`;

        // Write the formatted duration to column M (index 12)
        formSheet.getRange(i + 1, 13).setValue(formattedDuration);
        Logger.log(`Setting duration in row ${i + 1} to: ${formattedDuration}`);
        // Update the status to EXPIRED in column L (index 11)
        formSheet.getRange(i + 1, 12).setValue("EXPIRED");
        Logger.log(`Setting status in row ${i + 1} to: EXPIRED`);
        passesEnded++;
      }
    }

    if (passesEnded > 0) {
      Logger.log(`Ended ${passesEnded} passes for ${studentEmail}`);
      // Optionally, you could send a confirmation email to the student here
    }

    // Call promoteWaitlist() after processing the end pass form
    promoteWaitlist();
    Logger.log("promoteWaitlist() called.");

  } catch (error) {
    Logger.log(`Error in onEndPassFormSubmit: ${error}`);
    if (e && !e.range) {
      Logger.log("The 'range' property is missing from the event object. Ensure the trigger is set to 'On form submit'.");
    }
  }
}

const MAX_ACTIVE_DURATION_MINUTES = 20; // Define the maximum active duration

function markExpiredPasses() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  const data = sheet.getDataRange().getValues();
  const now = new Date();

  for (let i = 1; i < data.length; i++) {
    const status = data[i][11];
    const timestamp = new Date(data[i][0]);

    if (status === "ACTIVE") {
      // Calculate the expiry time based on the original timestamp
      const expiryTime = new Date(timestamp.getTime() + MAX_ACTIVE_DURATION_MINUTES * 60000);

      // If the current time is past the calculated expiry time
      if (now > expiryTime) {
        sheet.getRange(i + 1, 12).setValue("EXPIRED");
        Logger.log(`Pass for ${data[i][1]} expired automatically.`);
      }
    }
  }
}

function promoteWaitlist() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const [timestamp, email, , roomNumber, destination, , , , , , , status] = row;

    if (status === "WAITLISTED") {
      // Check if there are ANY active passes for this student's classroom
      const activePassInClassroom = data.some(r => r[3] === roomNumber && r[11] === "ACTIVE");

      if (!activePassInClassroom) {
        sheet.getRange(i + 1, 12).setValue("ACTIVE");

        // Visual duration for the email
        let visualDuration = VISUAL_PASS_DURATION_MINUTES;
        if (destination.toLowerCase().includes("nurse")) visualDuration = 8;
        else if (destination.toLowerCase().includes("locker")) visualDuration = 5;
        else if (destination.toLowerCase().includes("bathroom")) visualDuration = 5;
        else if (destination.toLowerCase().includes("guidance")) visualDuration = 12;

        const visualExpiry = new Date(new Date().getTime() + visualDuration * 60000);
        const name = extractNameFromEmail(email);
        const html = generatePassHtml("green", name, roomNumber, destination, visualExpiry);
        GmailApp.sendEmail('', EMAIL_SUBJECT, "Your hall pass is now active", { htmlBody: html, bcc: email });
        // Once promoted, we can break the loop to ensure only the first waitlisted student for that room is activated
        break;
      }
    }
  }
}

function hidePreviousDaysPasses() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const formSheet = ss.getSheetByName(SHEET_NAME); // Assuming SHEET_NAME is defined as 'Form'
  const data = formSheet.getDataRange().getValues();
  const today = new Date();
  today.setHours(0, 0, 0, 0); // Get the start of today

  // Start from the second row to skip headers
  for (let i = 1; i < data.length; i++) {
    const timestampValue = data[i][0];
    if (timestampValue instanceof Date) {
      const passDate = new Date(timestampValue);
      passDate.setHours(0, 0, 0, 0); // Get the start of the pass date

      // If the pass date is before today, hide the row
      if (passDate.getTime() < today.getTime()) {
        formSheet.hideRows(i + 1, 1); // i + 1 because sheet rows are 1-indexed
      }
    }
  }
  Logger.log('Previous days\' passes have been hidden.');
}

function populateRoomDropdown() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const formSheet = ss.getSheetByName('Form');
  const dashboardSheet = ss.getSheetByName('Dashboard');
  const formData = formSheet.getDataRange().getValues();
  const headerRow = formData[0];
  const roomColIndex = headerRow.indexOf('Room #');
  const rooms = [...new Set(formData.slice(1).map(row => row[roomColIndex]))].filter(Boolean).sort(); // Get unique rooms, remove blanks, and sort

  // Clear any existing dropdown
  const roomDropdownCell = dashboardSheet.getRange('G2');
  roomDropdownCell.clearDataValidations();
  dashboardSheet.getRange('G1').setValue('Filter by Room:');

  // Create the dropdown
  const rule = SpreadsheetApp.newDataValidation()
      .requireValueInList(rooms)
      .setAllowInvalid(false)
      .build();
  roomDropdownCell.setDataValidation(rule);
  roomDropdownCell.setValue(''); // Set default value to blank
}

function onDashboardChange(e) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dashboardSheet = ss.getSheetByName('Dashboard');
  const formSheet = ss.getSheetByName('Form');
  const changedCell = dashboardSheet.getActiveRange();

  // Check if the changed cell is the room number dropdown (assuming it's at G2)
  if (changedCell.getRow() === 2 && changedCell.getColumn() === 7) {
    const selectedRoom = changedCell.getValue();
    const formData = formSheet.getDataRange().getValues();
    const headerRow = formData[0];
    const dataRows = formData.slice(1);

    const timestampCol = headerRow.indexOf('Timestamp');
    const emailCol = headerRow.indexOf('Email Address');
    const roomCol = headerRow.indexOf('Room #');
    const destinationCol = headerRow.indexOf('Destination');

    const today = new Date();
    today.setHours(0, 0, 0, 0);
    const tomorrow = new Date(today);
    tomorrow.setDate(today.getDate() + 1);

    const filteredPasses = dataRows.filter(row => {
      const timestamp = new Date(row[timestampCol]);
      const roomNumber = row[roomCol];
      return roomNumber === selectedRoom && timestamp >= today && timestamp < tomorrow;
    });

    // Clear previous filtered data
    const startRow = 8; // Adjust this based on where you want to display the filtered list
    const numRowsToClear = dashboardSheet.getLastRow() - startRow + 1;
    if (numRowsToClear > 0) {
      dashboardSheet.getRange(startRow, 1, numRowsToClear, dashboardSheet.getLastColumn()).clearContent();
    }

    // Write the filtered passes
    if (filteredPasses.length > 0) {
      dashboardSheet.getRange(startRow, 1, 1, 4).setValues([['Timestamp', 'Student Email', 'Room', 'Destination']]).setFontWeight('bold');
      filteredPasses.forEach((pass, index) => {
        dashboardSheet.getRange(startRow + 1 + index, 1, 1, 4).setValues([[new Date(pass[timestampCol]), pass[emailCol], pass[roomCol], pass[destinationCol]]]);
      });
    } else if (selectedRoom !== '') {
      dashboardSheet.getRange(startRow, 1).setValue('No passes found for this room today.');
    }
  }
}

function updateHallPassDashboard() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const formSheet = ss.getSheetByName('Form');
  const dashboardSheet = ss.getSheetByName('Dashboard');
  const formData = formSheet.getDataRange().getValues();
  const headerRow = formData[0];
  const dataRows = formData.slice(1);

  const today = new Date();
  today.setHours(0, 0, 0, 0);
  const startOfWeek = new Date(today);
  const dayOfWeek = startOfWeek.getDay();
  const diff = startOfWeek.getDate() - dayOfWeek + (dayOfWeek === 0 ? -6 : 1);
  startOfWeek.setDate(diff);
  startOfWeek.setHours(0, 0, 0, 0);

  const classroomPassData = {};
  const studentPassData = {};

  const timestampCol = headerRow.indexOf('Timestamp');
  const emailCol = headerRow.indexOf('Email Address');
  const roomCol = headerRow.indexOf('Room #');
  const statusCol = headerRow.indexOf('Status');

  dataRows.forEach(row => {
    const timestamp = new Date(row[timestampCol]);
    const email = row[emailCol];
    const roomNumber = row[roomCol];
    const status = row[statusCol];

    if (roomNumber) {
      if (!classroomPassData[roomNumber]) {
        classroomPassData[roomNumber] = { active: 0, today: 0, thisWeek: 0, overall: 0 };
      }
      classroomPassData[roomNumber].overall++;
      if (status === 'ACTIVE') classroomPassData[roomNumber].active++;
      if (timestamp >= today) classroomPassData[roomNumber].today++;
      if (timestamp >= startOfWeek) classroomPassData[roomNumber].thisWeek++;
    }

    if (email) {
      if (!studentPassData[email]) {
        studentPassData[email] = { today: 0, thisWeek: 0, overall: 0 };
      }
      studentPassData[email].overall++;
      if (timestamp >= today) studentPassData[email].today++;
      if (timestamp >= startOfWeek) studentPassData[email].thisWeek++;
    }
  });

  dashboardSheet.getRange('A1:E1').clearContent();
  dashboardSheet.getRange('A1:E1').setValues([['Classroom', 'Active', 'Today', 'This Week', 'Overall']]).setFontWeight('bold');
  let rowNum = 2;
  for (const room in classroomPassData) {
    dashboardSheet.getRange(rowNum, 1, 1, 5).setValues([[room, classroomPassData[room].active, classroomPassData[room].today, classroomPassData[room].thisWeek, classroomPassData[room].overall]]);
    rowNum++;
  }

  const studentDataStartRow = rowNum + 2;
  dashboardSheet.getRange(studentDataStartRow, 1, 1, 4).clearContent();
  dashboardSheet.getRange(studentDataStartRow, 1, 1, 4).setValues([['Student Email', 'Today', 'This Week', 'Overall']]).setFontWeight('bold');
  let studentRowNum = studentDataStartRow + 1;
  for (const email in studentPassData) {
    dashboardSheet.getRange(studentRowNum, 1, 1, 4).setValues([[email, studentPassData[email].today, studentPassData[email].thisWeek, studentPassData[email].overall]]);
    studentRowNum++;
  }

  Logger.log('Hall Pass Dashboard Updated');
}

function onOpen() {
  SpreadsheetApp.getUi()
      .createMenu('Hall Pass Dashboard')
      .addItem('Populate Room Dropdown', 'populateRoomDropdown')
      .addToUi();
}
