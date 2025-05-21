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
  SCRIPT_PROPERTIES.setProperty('EMAILS_SHEET_NAME', 'Emails');
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
const EMAILS_SHEET_NAME = getConstant('EMAILS_SHEET_NAME');

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

// For ALL WAITLISTED Hall Passes (classroom capacity OR friend-based)
function generateWaitlistHtml(name, from, to, position, unlockTime) { // Removed reasonMessage parameter
  const destinationDisplay = destinationTitleCase(to);
  const now = new Date();
  const diff = unlockTime.getTime() - now.getTime();
  let remainingTime = "Estimating activation...";
  let timerColor = "#333";

  if (diff > 0) {
    const mins = Math.floor(diff / 60000);
    const secs = Math.floor((diff % 60000) / 1000);
    remainingTime = `${mins}m ${String(secs).padStart(2, '0')}s`;
  } else {
    remainingTime = "Awaiting activation.";
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
        <p style="margin: 12px 0; line-height: 1.6; color: #E67E22; font-weight: bold;">
          Your pass may be active soon. Check for a new email.
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

/**
 * Loads the friends data from the "Friends" sheet and returns a lookup object.
 * The lookup object maps each email to a Set of their friends' emails.
 * E.g., { "studentA@example.com": Set {"studentB@example.com", "studentC@example.com"} }
 */
function getFriendsLookup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const friendsSheet = ss.getSheetByName(FRIENDS_SHEET_NAME);
  const friendsLookup = new Map(); // Using Map for potentially better performance with string keys

  if (!friendsSheet) {
    Logger.log(`Warning: "Friends" sheet named "${FRIENDS_SHEET_NAME}" not found. Friend-based waitlisting will not be applied.`);
    return friendsLookup; // Return empty map if sheet is missing
  }

  const friendsData = friendsSheet.getDataRange().getValues();
  // Assuming friends are in columns A and B, starting from row 1 (no header assumed for simplicity, adjust if needed)
  for (let i = 0; i < friendsData.length; i++) {
    const email1 = String(friendsData[i][0]).toLowerCase().trim();
    const email2 = String(friendsData[i][1]).toLowerCase().trim();

    if (email1 && email2) { // Ensure both emails are present
      // Add email2 as a friend of email1
      if (!friendsLookup.has(email1)) {
        friendsLookup.set(email1, new Set());
      }
      friendsLookup.get(email1).add(email2);

      // Add email1 as a friend of email2 (friendship is mutual)
      if (!friendsLookup.has(email2)) {
        friendsLookup.set(email2, new Set());
      }
      friendsLookup.get(email2).add(email1);
    }
  }
  Logger.log(`Loaded ${friendsLookup.size} distinct students with friends.`);
  return friendsLookup;
}


function onFormSubmit(e) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME);
  const emailsSheet = ss.getSheetByName(EMAILS_SHEET_NAME);
  const friendsLookup = getFriendsLookup(); // Load friends data once per submission

  let nameLookup = {};
  if (emailsSheet) {
    const emailsData = emailsSheet.getDataRange().getValues();
    for (let i = 0; i < emailsData.length; i++) {
      const name = emailsData[i][0];
      const email = emailsData[i][1];
      if (email) {
        nameLookup[String(email).toLowerCase().trim()] = name; // Store emails as lowercase for consistent lookup
      }
    }
  } else {
    Logger.log(`Warning: Sheet "${EMAILS_SHEET_NAME}" not found. Using email to extract name.`);
  }

  const allData = sheet.getDataRange().getValues(); // Get all data to check for active passes
  const row = e.range.getRow();
  Logger.log('Row number of submission:', row);

  // Adjust row number to be 0-based index
  const rowIndex = row - 1;

  if (rowIndex >= 0 && rowIndex < allData.length) {
    const rowData = allData[rowIndex];
    Logger.log('rowData:', rowData);

    if (rowData && rowData.length >= 5) {
      const [timestamp, emailRaw, , roomNumber, destinationRaw] = rowData;
      const studentEmail = String(emailRaw).toLowerCase().trim(); // Ensure submitted email is clean
      const destination = destinationRaw.toLowerCase();
      let name = nameLookup[studentEmail]; // Try to get the name from the Emails sheet

      if (!name) {
        name = extractNameFromEmail(studentEmail);
        Logger.log(`Name not found in "${EMAILS_SHEET_NAME}" for email: ${studentEmail}. Extracted "${name}" from email.`);
      }

      const timeSubmitted = new Date(timestamp);
      const today = timeSubmitted.toDateString();

      // --- Start Friend Check Logic ---
      let friendHasActivePass = false;
      const studentFriends = friendsLookup.has(studentEmail) ? friendsLookup.get(studentEmail) : new Set();

      if (studentFriends.size > 0) {
        Logger.log(`Checking for active passes among friends of ${studentEmail}: ${Array.from(studentFriends).join(', ')}`);
        // Iterate through all existing passes to check for active friends
        for (let i = 1; i < allData.length; i++) { // Start from 1 to skip header row
          const existingPass = allData[i];
          const existingPassEmail = String(existingPass[1]).toLowerCase().trim();
          const existingPassStatus = existingPass[11]; // Column L

          if (existingPassStatus === "ACTIVE" && studentFriends.has(existingPassEmail)) {
            friendHasActivePass = true;
            Logger.log(`Friend ${existingPassEmail} has an active pass.`);
            break; // Found an active friend, no need to check further
          }
        }
      }
      // --- End Friend Check Logic ---

      if (isExemptStudent(studentEmail)) {
        const actualExpiry = new Date(timeSubmitted.getTime() + MAX_PASS_DURATION_MINUTES * 60000);
        const visualExpiry = new Date(timeSubmitted.getTime() + EXEMPT_VISUAL_DURATION_MINUTES * 60000);
        const emailBody = generatePassHtml("green", name, roomNumber, destination, visualExpiry);
        sheet.getRange(row, 12).setValue("ACTIVE"); // Column L - Status
        sheet.getRange(row, 14).setValue(timestamp); // Column N - Activation Timestamp
        sheet.getRange(row, 3).setValue(name);    // Column C - Name
        GmailApp.sendEmail('', EMAIL_SUBJECT, 'Your Hall Pass is Approved (Exempt)', {
          htmlBody: emailBody,
          bcc: studentEmail
        });
        sortFormSheetNewestToOldest();
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

      const studentPasses = allData.filter(r => String(r[1]).toLowerCase().trim() === studentEmail);
      const hasActivePass = studentPasses.some(r => r[11] === "ACTIVE");

      let status = "REJECTED";
      let emailBody = "";
      let waitlistReasonForLog = ""; // For internal logging only

      if (hasActivePass) {
        emailBody = createHtml("red", `You already have an active hall pass. Please wait until it is marked as used.`);
        waitlistReasonForLog = "Already active pass";
      } else if (timeBlock === "OUT") {
        emailBody = createHtml("red", `Hall passes are unavailable at this time.`);
        waitlistReasonForLog = "Outside operating hours";
      } else {
        const usedPassesInBlock = studentPasses.filter(r => {
          const rTimestamp = new Date(r[0]);
          const rDate = rTimestamp.toDateString();
          const rHour = rTimestamp.getHours();
          const rTimeBlock = rHour >= 8 && rHour < 12 ? "AM" : rHour >= 12 && rHour <= 23 ? "PM" : "OUT";
          const rStatus = r[11];
          return rDate === today && rTimeBlock === timeBlock && rStatus === "EXPIRED";
        }).length;

        const nonExemptActivePassInClassroom = allData.some(r => r[3] === roomNumber && r[11] === "ACTIVE" && !isExemptStudent(String(r[1]).toLowerCase().trim()));

        if (usedPassesInBlock >= MAX_PASSES_PER_BLOCK) {
          emailBody = createHtml("red", `You have already reached your limit of ${MAX_PASSES_PER_BLOCK} hall passes this ${timeBlock === "AM" ? "morning" : "afternoon"}.`);
          waitlistReasonForLog = "Pass limit reached";
        } else if (friendHasActivePass || nonExemptActivePassInClassroom) { // Both conditions now result in generic waitlist
          status = "WAITLISTED";
          // Call generateWaitlistHtml without a specific reason message
          emailBody = generateWaitlistHtml(name, roomNumber, destination, 'N/A', new Date(Date.now() + 300000));
          if (friendHasActivePass) {
            waitlistReasonForLog = "Friend has active pass";
          } else { // nonExemptActivePassInClassroom must be true here
            waitlistReasonForLog = "Active pass in classroom";
          }
          Logger.log(`Student ${studentEmail} waitlisted. Internal reason: ${waitlistReasonForLog}`);
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
        bcc: studentEmail
      });
      Logger.log(`Hall pass for ${studentEmail} set to status: ${status}. Internal reason: ${waitlistReasonForLog || 'Approved'}`);
    } else {
      Logger.log('Error: rowData is undefined or has insufficient length.');
    }
  } else {
    Logger.log('Error: Row number from form submission is out of bounds.');
  }
  sortFormSheetNewestToOldest();
}

function onEndPassFormSubmit(e) {
  Logger.log("onEndPassFormSubmit triggered with event:");
  Logger.log(JSON.stringify(e));

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const endPassSheet = ss.getSheetByName(END_PASS_SHEET_NAME);
  const endPassRow = e.range.getRow();
  const endPassRowData = endPassSheet.getRange(endPassRow, 1, 1, endPassSheet.getLastColumn()).getValues()[0];
  const [endPassTimestamp, studentEmailRaw] = endPassRowData;
  const studentEmail = String(studentEmailRaw).toLowerCase().trim();

  const formSheet = ss.getSheetByName(SHEET_NAME);
  const formData = formSheet.getDataRange().getValues();

  const updates = [];
  let passesEnded = 0;

  for (let i = 1; i < formData.length; i++) {
    const rowData = formData[i];
    const rowEmail = String(rowData[1]).toLowerCase().trim();
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

      updates.push({ row: i + 1, duration: formattedDuration, status: "EXPIRED" }); // Store row, duration, and new status
      passesEnded++;
    }
  }

  if (updates.length > 0) {
    updates.forEach(update => {
      formSheet.getRange(update.row, 13).setValue(update.duration); // Column M - Duration
      formSheet.getRange(update.row, 12).setValue(update.status);  // Column L - Status
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
    for (const updateItem of updates) {
      sheet.getRange(updateItem.row, updateItem.col).setValue(updateItem.value);
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
  const friendsLookup = getFriendsLookup(); // Load friends data for promotion as well

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const [timestamp, emailRaw, , roomNumber, destination, , , , , , , status] = row;
    const studentEmail = String(emailRaw).toLowerCase().trim();

    if (status === "WAITLISTED") {
      const activePassInClassroom = data.some(r => r[3] === roomNumber && r[11] === "ACTIVE" && !isExemptStudent(String(r[1]).toLowerCase().trim()));

      // Check if a friend currently has an active pass
      let friendHasActivePass = false;
      const studentFriends = friendsLookup.has(studentEmail) ? friendsLookup.get(studentEmail) : new Set();
      if (studentFriends.size > 0) {
        for (let j = 1; j < data.length; j++) { // Check all other passes
          if (i === j) continue; // Skip checking the current student's pass against itself
          const existingPass = data[j];
          const existingPassEmail = String(existingPass[1]).toLowerCase().trim();
          const existingPassStatus = existingPass[11];

          if (existingPassStatus === "ACTIVE" && studentFriends.has(existingPassEmail)) {
            friendHasActivePass = true;
            break;
          }
        }
      }

      // Promote ONLY if neither classroom nor friend condition prevents it
      if (!activePassInClassroom && !friendHasActivePass) {
        sheet.getRange(i + 1, 12).setValue("ACTIVE");
        sheet.getRange(i + 1, 14).setValue(new Date()); // Set the activation timestamp

        let visualDuration = VISUAL_PASS_DURATION_MINUTES;
        if (destination.toLowerCase().includes("nurse")) visualDuration = 8;
        else if (destination.toLowerCase().includes("locker")) visualDuration = 5;
        else if (destination.toLowerCase().includes("bathroom")) visualDuration = 5;
        else if (destination.toLowerCase().includes("guidance")) visualDuration = 12;

        const visualExpiry = new Date(new Date().getTime() + visualDuration * 60000);
        const name = extractNameFromEmail(studentEmail);
        const html = generatePassHtml("green", name, roomNumber, destination, visualExpiry);
        GmailApp.sendEmail('', EMAIL_SUBJECT, "Your hall pass is now active", { htmlBody: html, bcc: studentEmail });
        Logger.log(`Student ${studentEmail} promoted from waitlist to active.`);
        break;
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
