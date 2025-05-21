const SCRIPT_PROPERTIES = PropertiesService.getScriptProperties();

function updateConstants() {
  SCRIPT_PROPERTIES.setProperty('SHEET_NAME', 'Form');
  SCRIPT_PROPERTIES.setProperty('END_PASS_SHEET_NAME', 'EndPass');
  SCRIPT_PROPERTIES.setProperty('EMAIL_SUBJECT', 'Your Hall Pass Status');
  SCRIPT_PROPERTIES.setProperty('MAX_PASSES_PER_BLOCK', '2');
  SCRIPT_PROPERTIES.setProperty('END_PASS_FORM_URL', 'https://forms.gle/yxrgRf9rV48SNWmc9');
  SCRIPT_PROPERTIES.setProperty('MAX_PASS_DURATION_MINUTES', '20');
  SCRIPT_PROPERTIES.setProperty('VISUAL_PASS_DURATION_MINUTES', '5');
  SCRIPT_PROPERTIES.setProperty('EXEMPT_VISUAL_DURATION_MINUTES', '6');
  SCRIPT_PROPERTIES.setProperty('FRIENDS_SHEET_NAME', 'Friends');
  SCRIPT_PROPERTIES.setProperty('EXEMPT_SHEET_NAME', 'Exempt');
  SCRIPT_PROPERTIES.setProperty('MAX_ACTIVE_DURATION_MINUTES', '20');
  SCRIPT_PROPERTIES.setProperty('EMAILS_SHEET_NAME', 'Emails'); // New constant for the Emails sheet name
}

function getConstant(key) {
  return SCRIPT_PROPERTIES.getProperty(key);
}

const SHEET_NAME = getConstant('SHEET_NAME');
const END_PASS_SHEET_NAME = getConstant('END_PASS_SHEET_NAME');
const EMAIL_SUBJECT = getConstant('EMAIL_SUBJECT');
const MAX_PASSES_PER_BLOCK = parseInt(getConstant('MAX_PASSES_PER_BLOCK'));
const END_PASS_FORM_URL = getConstant('END_PASS_FORM_URL');
const MAX_PASS_DURATION_MINUTES = parseInt(getConstant('MAX_PASS_DURATION_MINUTES'));
const VISUAL_PASS_DURATION_MINUTES = parseInt(getConstant('VISUAL_PASS_DURATION_MINUTES'));
const EXEMPT_VISUAL_DURATION_MINUTES = parseInt(getConstant('EXEMPT_VISUAL_DURATION_MINUTES'));
const FRIENDS_SHEET_NAME = getConstant('FRIENDS_SHEET_NAME');
const EXEMPT_SHEET_NAME = getConstant('EXEMPT_SHEET_NAME');
const MAX_ACTIVE_DURATION_MINUTES = parseInt(getConstant('MAX_ACTIVE_DURATION_MINUTES'));
const EMAILS_SHEET_NAME = getConstant('EMAILS_SHEET_NAME'); // Get the Emails sheet name

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
        <p style="margin: 12px 0; line-height: 1.6;"><strong>Expires On:</strong> ${expiryDateFormatted} at ${expiryTimeFormatted}</p>
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
  const exemptSheet = ss.getSheetByName(EXEMPT_SHEET_NAME);
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
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME);
  const emailsSheet = ss.getSheetByName(EMAILS_SHEET_NAME); // Get the Emails sheet
  let nameLookup = {};

  if (emailsSheet) {
    const emailsData = emailsSheet.getDataRange().getValues();
    // Create a lookup object (dictionary) where key is email and value is name
    for (let i = 0; i < emailsData.length; i++) {
      const name = emailsData[i][0]; // Column A is Name
      const email = emailsData[i][1]; // Column B is Email
      if (email) {
        nameLookup[email] = name;
      }
    }
  } else {
    Logger.log(`Warning: Sheet "${EMAILS_SHEET_NAME}" not found. Using email to extract name.`);
  }

  const allData = sheet.getDataRange().getValues();
  const row = e.range.getRow();
  Logger.log('Row number of submission:', row);
  Logger.log('Number of rows in allData:', allData.length);

  // Adjust row number to be 0-based index
  const rowIndex = row - 1;

  if (rowIndex >= 0 && rowIndex < allData.length) {
    const rowData = allData[rowIndex];
    Logger.log('rowData:', rowData);

    if (rowData && rowData.length >= 5) { // Ensure rowData is not undefined and has at least 5 elements
      const [timestamp, email, , roomNumber, destinationRaw] = rowData;
      const destination = destinationRaw.toLowerCase();
      let name = nameLookup[email]; // Try to get the name from the Emails sheet

      if (!name) {
        name = extractNameFromEmail(email); // Fallback to extracting from email if not found
        Logger.log(`Name not found in "${EMAILS_SHEET_NAME}" for email: ${email}. Extracted "${name}" from email.`);
      }

      const timeSubmitted = new Date(timestamp);
      const today = timeSubmitted.toDateString(); // Get the date part only

      if (isExemptStudent(email)) {
        const actualExpiry = new Date(timeSubmitted.getTime() + MAX_PASS_DURATION_MINUTES * 60000);
        const visualExpiry = new Date(timeSubmitted.getTime() + EXEMPT_VISUAL_DURATION_MINUTES * 60000);
        const emailBody = generatePassHtml("green", name, roomNumber, destination, visualExpiry);
        sheet.getRange(row, 12).setValue("ACTIVE"); // Column L - Status
        sheet.getRange(row, 14).setValue(timestamp); // Column N - Activation Timestamp
        sheet.getRange(row,3).setValue(name);    // Column C - Name
        GmailApp.sendEmail('', EMAIL_SUBJECT, 'Your Hall Pass is Approved (Exempt)', {
          htmlBody: emailBody,
          bcc: email
        });
        sortFormSheetNewestToOldest(); // Sort after processing the submission
        return;
      }

      const currentTime = new Date();
      const hour = timeSubmitted.getHours();
      const timeBlock = hour >= 8 && hour < 12 ? "AM" : hour >= 12 && hour <= 23 ? "PM" : "OUT";

      let visualDuration = VISUAL_PASS_DURATION_MINUTES;
      if (destination.includes("nurse")) visualDuration = 8;
      else if (destination.includes("locker")) visualDuration = 5;
      else if (destination.includes("bathroom")) visualDuration = 5;
      else if (destination.includes("guidance")) visualDuration = 12;

      const studentPasses = allData.filter(r => r[1] === email);
      const hasActivePass = studentPasses.some(r => r[11] === "ACTIVE");

      let status = "REJECTED";
      let emailBody = "";

      if (hasActivePass) {
        emailBody = createHtml("red", `You already have an active hall pass. Please wait until it is marked as used.`);
      } else if (timeBlock === "OUT") {
        emailBody = createHtml("red", `Hall passes are unavailable at this time.`);
      } else {
        // Only count EXPIRED passes from the current day within the current time block
        const usedPassesInBlock = studentPasses.filter(r => {
          const rTimestamp = new Date(r[0]);
          const rDate = rTimestamp.toDateString();
          const rHour = rTimestamp.getHours();
          const rTimeBlock = rHour >= 8 && rHour < 12 ? "AM" : rHour >= 12 && rHour <= 23 ? "PM" : "OUT";
          const rStatus = r[11];
          return rDate === today && rTimeBlock === timeBlock && rStatus === "EXPIRED";
        }).length;

        const nonExemptActivePassInClassroom = allData.some(r => r[3] === roomNumber && r[11] === "ACTIVE" && !isExemptStudent(r[1]));

        if (usedPassesInBlock >= MAX_PASSES_PER_BLOCK) {
          emailBody = createHtml("red", `You have already reached your limit of ${MAX_PASSES_PER_BLOCK} hall passes this ${timeBlock === "AM" ? "morning" : "afternoon"}.`);
        } else if (nonExemptActivePassInClassroom) {
          status = "WAITLISTED";
          emailBody = generateWaitlistHtml(name, roomNumber, destination, 'N/A', new Date(Date.now() + 300000));
        } else {
          status = "ACTIVE";
          const actualExpiry = new Date(timeSubmitted.getTime() + MAX_PASS_DURATION_MINUTES * 60000);
          const visualExpiry = new Date(timeSubmitted.getTime() + visualDuration * 60000);
          emailBody = generatePassHtml("green", name, roomNumber, destination, visualExpiry);
          sheet.getRange(row, 14).setValue(timestamp); // Column N - Activation Timestamp
        }
      }

      sheet.getRange(row, 12).setValue(status);
      sheet.getRange(row, 3).setValue(name);

      GmailApp.sendEmail('', EMAIL_SUBJECT, 'Your hall pass info', {
        htmlBody: emailBody,
        bcc: email
      });
    } else {
      Logger.log('Error: rowData is undefined or has insufficient length.');
    }
  } else {
    Logger.log('Error: Row number from form submission is out of bounds.');
  }
  sortFormSheetNewestToOldest(); // Sort after processing the submission
}

function onEndPassFormSubmit(e) {
  Logger.log("onEndPassFormSubmit triggered with event:");
  Logger.log(JSON.stringify(e));

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const endPassSheet = ss.getSheetByName(END_PASS_SHEET_NAME);
  const endPassRow = e.range.getRow();
  const endPassRowData = endPassSheet.getRange(endPassRow, 1, 1, endPassSheet.getLastColumn()).getValues()[0];
  const [endPassTimestamp, studentEmail] = endPassRowData;

  const formSheet = ss.getSheetByName(SHEET_NAME);
  const formData = formSheet.getDataRange().getValues();

  const updates = [];
  let passesEnded = 0;

  for (let i = 1; i < formData.length; i++) {
    const rowData = formData[i];
    const rowEmail = rowData[1];
    const status = rowData[11];
    const activationTimestampValue = rowData[13]; // Column N

    if (rowEmail === studentEmail && (status === "ACTIVE" || status === "WAITLISTED")) {
      let startTime;
      if (activationTimestampValue instanceof Date) {
        startTime = activationTimestampValue;
      } else {
        startTime = new Date(rowData[0]); // Fallback
        Logger.log(`Warning: Activation timestamp missing for row ${i + 1}. Using submission time.`);
      }

      const durationMillis = endPassTimestamp.getTime() - startTime.getTime();
      const totalSeconds = Math.floor(durationMillis / 1000);
      const minutes = Math.floor(totalSeconds / 60);
      const seconds = totalSeconds % 60;
      const formattedDuration = `${String(minutes).padStart(1, '0')}:${String(seconds).padStart(2, '0')}`;

      updates.push([formattedDuration, "EXPIRED", i + 1]); // [duration, status, row number]
      passesEnded++;
    }
  }

  if (updates.length > 0) {
    updates.forEach(update => {
      const [duration, status, row] = update;
      formSheet.getRange(row, 13).setValue(duration); // Column M - Duration
      formSheet.getRange(row, 12).setValue(status);  // Column L - Status
    });
    Logger.log(`Updated ${updates.length} passes for ${studentEmail}.`);
  }

  promoteWaitlist();
  Logger.log("promoteWaitlist() called.");
}

function markExpiredPasses() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME);
  const data = sheet.getDataRange().getValues();
  const now = new Date();
  const updates = [];

  for (let i = 1; i < data.length; i++) {
    const rowData = data[i];
    const status = rowData[11]; // Column L - Status (0-indexed 11)
    const activationTimestampValue = rowData[13]; // Column N - Activation Timestamp (0-indexed 13)

    if (status === "ACTIVE") {
      let startTime;
      if (activationTimestampValue instanceof Date) {
        startTime = activationTimestampValue;
      } else {
        startTime = new Date(rowData[0]); // Fallback to submission time
      }
      const expiryTime = new Date(startTime.getTime() + MAX_ACTIVE_DURATION_MINUTES * 60000);
      if (now > expiryTime) {
        updates.push({ row: i + 1, col: 12, value: "EXPIRED" }); // Store row, col (1-indexed), and new status
        Logger.log(`Pass for ${rowData[1]} expired automatically.`);
      }
    }
  }

  // Perform updates
  if (updates.length > 0) {
    // Collect all unique rows to update to minimize calls
    const rowsToUpdate = {};
    updates.forEach(update => {
      if (!rowsToUpdate[update.row]) {
        rowsToUpdate[update.row] = [];
      }
      rowsToUpdate[update.row].push({ col: update.col, value: update.value });
    });

    for (const rowNum in rowsToUpdate) {
      rowsToUpdate[rowNum].forEach(updateItem => {
        sheet.getRange(parseInt(rowNum), updateItem.col).setValue(updateItem.value);
      });
    }

    Logger.log(`Updated ${updates.length} passes to EXPIRED.`);
  } else {
    Logger.log('No active passes to expire.');
  }
}

function promoteWaitlist() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME);
  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const [timestamp, email, , roomNumber, destination, , , , , , , status] = row;

    if (status === "WAITLISTED") {
      const activePassInClassroom = data.some(r => r[3] === roomNumber && r[11] === "ACTIVE");

      if (!activePassInClassroom) {
        sheet.getRange(i + 1, 12).setValue("ACTIVE");
        sheet.getRange(i + 1, 14).setValue(new Date()); // Set the activation timestamp

        let visualDuration = VISUAL_PASS_DURATION_MINUTES;
        if (destination.toLowerCase().includes("nurse")) visualDuration = 8;
        else if (destination.toLowerCase().includes("locker")) visualDuration = 5;
        else if (destination.toLowerCase().includes("bathroom")) visualDuration = 5;
        else if (destination.toLowerCase().includes("guidance")) visualDuration = 12;

        const visualExpiry = new Date(new Date().getTime() + visualDuration * 60000);
        const name = extractNameFromEmail(email);
        const html = generatePassHtml("green", name, roomNumber, destination, visualExpiry);
        GmailApp.sendEmail('', EMAIL_SUBJECT, "Your hall pass is now active", { htmlBody: html, bcc: email });
        break; // Promote only the first waitlisted student for that room
      }
    }
  }
}

function hidePreviousDaysPasses() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const formSheet = ss.getSheetByName(SHEET_NAME);
  const data = formSheet.getDataRange().getValues();
  const today = new Date();
  today.setHours(0, 0, 0, 0);

  const rowsToHide = [];
  for (let i = 1; i < data.length; i++) {
    const timestampValue = data[i][0];
    if (timestampValue instanceof Date) {
      const passDate = new Date(timestampValue);
      passDate.setHours(0, 0, 0, 0);
      if (passDate.getTime() < today.getTime()) {
        rowsToHide.push(i + 1);
      }
    }
  }

  if (rowsToHide.length > 0) {
    formSheet.hideRows(rowsToHide[0], rowsToHide.length);
    Logger.log(`Hidden ${rowsToHide.length} previous days' passes.`);
  } else {
    Logger.log('No previous days\' passes to hide.');
  }
}

function sortFormSheetNewestToOldest() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const formSheet = ss.getSheetByName(SHEET_NAME);
  if (!formSheet) {
    Logger.log(`Error: Sheet "${SHEET_NAME}" not found.`);
    return;
  }
  const lastRow = formSheet.getLastRow();
  if (lastRow < 2) {
    Logger.log('No data to sort (less than 2 rows).');
    return;
  }
  const rangeToSort = formSheet.getRange(2, 1, lastRow - 1, formSheet.getLastColumn()); // Start from row 2 to exclude headers
  rangeToSort.sort({ column: 1, ascending: false }); // Sort by the first column (Timestamp), descending (newest first)
  Logger.log('Sheet "${SHEET_NAME}" sorted from newest to oldest.');
}

// Initialize constants in script properties (run this once)
// updateConstants();
