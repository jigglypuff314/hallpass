const SCRIPT_PROPERTIES = PropertiesService.getScriptProperties();

function updateConstants() {
  SCRIPT_PROPERTIES.setProperty('SHEET_NAME', 'Form');
  SCRIPT_PROPERTIES.setProperty('END_PASS_SHEET_NAME', 'EndPass');
  SCRIPT_PROPERTIES.setProperty('EMAIL_SUBJECT', 'Your Hall Pass Status');
  SCRIPT_PROPERTIES.setProperty('MAX_PASSES_PER_BLOCK', '2');
  SCRIPT_PROPERTIES.setProperty('END_PASS_FORM_URL', 'https://forms.gle/yxrgRf9rV48SNWmc9'); // REPLACE WITH YOUR ACTUAL END PASS FORM URL
  SCRIPT_PROPERTIES.setProperty('MAX_PASS_DURATION_MINUTES', '20');
  SCRIPT_PROPERTIES.setProperty('VISUAL_PASS_DURATION_MINUTES', '5');
  SCRIPT_PROPERTIES.setProperty('EXEMPT_VISUAL_DURATION_MINUTES', '6');
  SCRIPT_PROPERTIES.setProperty('FRIENDS_SHEET_NAME', 'Friends');
  SCRIPT_PROPERTIES.setProperty('EXEMPT_SHEET_NAME', 'Exempt');
  SCRIPT_PROPERTIES.setProperty('MAX_ACTIVE_DURATION_MINUTES', '20');
  SCRIPT_PROPERTIES.setProperty('EMAILS_SHEET_NAME', 'Emails');
  Logger.log("Script properties updated. Remember to run setup() once manually.");
}

// Retrieve constants,
// Cache them globally after first call to avoid re-reading properties service
let SHEET_NAME_CACHED;
let END_PASS_SHEET_NAME_CACHED;
let EMAIL_SUBJECT_CACHED;
let MAX_PASSES_PER_BLOCK_CACHED;
let END_PASS_FORM_URL_CACHED;
let MAX_PASS_DURATION_MINUTES_CACHED;
let VISUAL_PASS_DURATION_MINUTES_CACHED;
let EXEMPT_VISUAL_DURATION_MINUTES_CACHED;
let FRIENDS_SHEET_NAME_CACHED;
let EXEMPT_SHEET_NAME_CACHED;
let MAX_ACTIVE_DURATION_MINUTES_CACHED;
let EMAILS_SHEET_NAME_CACHED;

function getConstant(key) {
  switch (key) {
    case 'SHEET_NAME': return SHEET_NAME_CACHED || (SHEET_NAME_CACHED = SCRIPT_PROPERTIES.getProperty(key));
    case 'END_PASS_SHEET_NAME': return END_PASS_SHEET_NAME_CACHED || (END_PASS_SHEET_NAME_CACHED = SCRIPT_PROPERTIES.getProperty(key));
    case 'EMAIL_SUBJECT': return EMAIL_SUBJECT_CACHED || (EMAIL_SUBJECT_CACHED = SCRIPT_PROPERTIES.getProperty(key));
    case 'MAX_PASSES_PER_BLOCK': return MAX_PASSES_PER_BLOCK_CACHED || (MAX_PASSES_PER_BLOCK_CACHED = parseInt(SCRIPT_PROPERTIES.getProperty(key)));
    case 'END_PASS_FORM_URL': return END_PASS_FORM_URL_CACHED || (END_PASS_FORM_URL_CACHED = SCRIPT_PROPERTIES.getProperty(key));
    case 'MAX_PASS_DURATION_MINUTES': return MAX_PASS_DURATION_MINUTES_CACHED || (MAX_PASS_DURATION_MINUTES_CACHED = parseInt(SCRIPT_PROPERTIES.getProperty(key)));
    case 'VISUAL_PASS_DURATION_MINUTES': return VISUAL_PASS_DURATION_MINUTES_CACHED || (VISUAL_PASS_DURATION_MINUTES_CACHED = parseInt(SCRIPT_PROPERTIES.getProperty(key)));
    case 'EXEMPT_VISUAL_DURATION_MINUTES': return EXEMPT_VISUAL_DURATION_MINUTES_CACHED || (EXEMPT_VISUAL_DURATION_MINUTES_CACHED = parseInt(SCRIPT_PROPERTIES.getProperty(key)));
    case 'FRIENDS_SHEET_NAME': return FRIENDS_SHEET_NAME_CACHED || (FRIENDS_SHEET_NAME_CACHED = SCRIPT_PROPERTIES.getProperty(key));
    case 'EXEMPT_SHEET_NAME': return EXEMPT_SHEET_NAME_CACHED || (EXEMPT_SHEET_NAME_CACHED = SCRIPT_PROPERTIES.getProperty(key));
    case 'MAX_ACTIVE_DURATION_MINUTES': return MAX_ACTIVE_DURATION_MINUTES_CACHED || (MAX_ACTIVE_DURATION_MINUTES_CACHED = parseInt(SCRIPT_PROPERTIES.getProperty(key)));
    case 'EMAILS_SHEET_NAME': return EMAILS_SHEET_NAME_CACHED || (EMAILS_SHEET_NAME_CACHED = SCRIPT_PROPERTIES.getProperty(key));
    default: return null;
  }
}

const SHEET_NAME = getConstant('SHEET_NAME');
const END_PASS_SHEET_NAME = getConstant('END_PASS_SHEET_NAME');
const EMAIL_SUBJECT = getConstant('EMAIL_SUBJECT');
const MAX_PASSES_PER_BLOCK = getConstant('MAX_PASSES_PER_BLOCK');
const END_PASS_FORM_URL = getConstant('END_PASS_FORM_URL');
const MAX_PASS_DURATION_MINUTES = getConstant('MAX_PASS_DURATION_MINUTES');
const VISUAL_PASS_DURATION_MINUTES = getConstant('VISUAL_PASS_DURATION_MINUTES');
const EXEMPT_VISUAL_DURATION_MINUTES = getConstant('EXEMPT_VISUAL_DURATION_MINUTES');
const FRIENDS_SHEET_NAME = getConstant('FRIENDS_SHEET_NAME');
const EXEMPT_SHEET_NAME = getConstant('EXEMPT_SHEET_NAME');
const MAX_ACTIVE_DURATION_MINUTES = getConstant('MAX_ACTIVE_DURATION_MINUTES');
const EMAILS_SHEET_NAME = getConstant('EMAILS_SHEET_NAME');

// Global variable (cached) for exempt students - initialized on first call to getExemptStudentsSet
let exemptStudentsCache = null;

/**
 * Loads all exempt student emails into a Set for quick lookups.
 * @returns {Set<string>} A Set containing lowercase, trimmed exempt student emails.
 */
function getExemptStudentsSet() {
  if (exemptStudentsCache) {
    return exemptStudentsCache; // Return cached data if available
  }
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const exemptSheet = ss.getSheetByName(EXEMPT_SHEET_NAME);
  const exemptStudents = new Set();

  if (!exemptSheet) {
    Logger.log(`Warning: "Exempt" sheet named "${EXEMPT_SHEET_NAME}" not found. Exemptions will not be applied.`);
    exemptStudentsCache = exemptStudents; // Cache empty set to avoid repeated warnings
    return exemptStudents;
  }

  const exemptData = exemptSheet.getDataRange().getValues();
  // Assuming email addresses are in the second column (index 1)
  for (let i = 0; i < exemptData.length; i++) {
    const email = String(exemptData[i][1]).toLowerCase().trim();
    if (email) {
      exemptStudents.add(email);
    }
  }
  Logger.log(`Loaded ${exemptStudents.size} exempt students.`);
  exemptStudentsCache = exemptStudents;
  return exemptStudents;
}

function isExemptStudent(email) {
  if (!exemptStudentsCache) { // Initialize cache if not already done by other functions
    exemptStudentsCache = getExemptStudentsSet();
  }
  return exemptStudentsCache.has(String(email).toLowerCase().trim());
}

function extractNameFromEmail(email) {
  const parts = email.split('@')[0].split('.');
  return parts.map(p => p.charAt(0).toUpperCase() + p.slice(1)).join(' ');
}

// For ACTIVE Hall Passes - Now shows only date and start time
function generatePassHtml(bg, name, from, to, visualPassEndTime, visualDuration) {
  const destinationDisplay = destinationTitleCase(to);

  // Calculate the start time (which is the actual pass activation time for display purposes)
  const visualPassStartTime = new Date(visualPassEndTime.getTime() - visualDuration * 60000);

  // Format date and time for display
  const dateTimeFormatted = visualPassStartTime.toLocaleString([], {
    month: 'numeric',
    day: 'numeric',
    year: 'numeric',
    hour: '2-digit',
    minute: '2-digit'
  });

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
        <p style="margin: 12px 0; line-height: 1.6;"><strong>Time Start:</strong> ${dateTimeFormatted}</p>
        <p style="margin-top: 20px; text-align: center;">
          <a href="${END_PASS_FORM_URL}" style="display: inline-block; background-color: #f44336; color: white; padding: 10px 18px; text-decoration: none; border-radius: 5px; font-weight: bold; font-size: 16px;" target="_blank">End Pass</a>
        </p>
      </div>
    </div>
    <div style="text-align: center; margin-top: 25px; font-size: 12px; color: #888;">
      This is an automated email. Do not reply to this email.
    </div>
  </div>`;
}

// For ALL WAITLISTED Hall Passes
function generateWaitlistHtml(name, from, to, position, unlockTime) {
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

// Global variable (cached) for friends lookup
let friendsLookupCache = null;

/**
 * Loads the friends data from the "Friends" sheet and returns a lookup object.
 * The lookup object maps each email to a Set of their friends' emails.
 */
function getFriendsLookup() {
  if (friendsLookupCache) {
    return friendsLookupCache; // Return cached data if available
  }
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const friendsSheet = ss.getSheetByName(FRIENDS_SHEET_NAME);
  const friendsLookup = new Map();

  if (!friendsSheet) {
    Logger.log(`Warning: "Friends" sheet named "${FRIENDS_SHEET_NAME}" not found. Friend-based waitlisting will not be applied.`);
    friendsLookupCache = friendsLookup;
    return friendsLookup;
  }

  const friendsData = friendsSheet.getDataRange().getValues();
  for (let i = 0; i < friendsData.length; i++) {
    const email1 = String(friendsData[i][0]).toLowerCase().trim();
    const email2 = String(friendsData[i][1]).toLowerCase().trim();

    if (email1 && email2) {
      if (!friendsLookup.has(email1)) {
        friendsLookup.set(email1, new Set());
      }
      friendsLookup.get(email1).add(email2);

      if (!friendsLookup.has(email2)) {
        friendsLookup.set(email2, new Set());
      }
      friendsLookup.get(email2).add(email1);
    }
  }
  Logger.log(`Loaded ${friendsLookup.size} distinct students with friends.`);
  friendsLookupCache = friendsLookup;
  return friendsLookup;
}


function onFormSubmit(e) {
  Logger.log("onFormSubmit triggered for new submission.");
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    Logger.log(`Error: Main sheet "${SHEET_NAME}" not found.`);
    return;
  }
  const emailsSheet = ss.getSheetByName(EMAILS_SHEET_NAME);

  // --- Load all necessary lookup data once per execution ---
  const friendsLookup = getFriendsLookup();
  const exemptStudents = getExemptStudentsSet();

  let nameLookup = new Map();
  if (emailsSheet) {
    const emailsData = emailsSheet.getDataRange().getValues();
    emailsData.forEach(row => {
      const name = row[0];
      const email = String(row[1]).toLowerCase().trim();
      if (email) {
        nameLookup.set(email, name);
      }
    });
    Logger.log(`Loaded ${nameLookup.size} names from "${EMAILS_SHEET_NAME}".`);
  } else {
    Logger.log(`Warning: Sheet "${EMAILS_SHEET_NAME}" not found. Using email to extract name.`);
  }

  // --- Read all main sheet data once ---
  const allDataRange = sheet.getDataRange();
  const allData = allDataRange.getValues();
  const row = e.range.getRow(); // Get the row number of the submission (1-indexed)
  const rowIndex = row - 1; // Convert to 0-indexed for array access

  if (rowIndex < 0 || rowIndex >= allData.length || !allData[rowIndex] || allData[rowIndex].length < 5) {
    Logger.log('Error: Invalid row data from form submission or insufficient columns.');
    return;
  }

  // Get the submitted row's data
  const rowData = allData[rowIndex];
  const [timestamp, emailRaw, , roomNumber, destinationRaw] = rowData;
  const studentEmail = String(emailRaw).toLowerCase().trim();
  const destination = destinationRaw.toLowerCase();
  let name = nameLookup.has(studentEmail) ? nameLookup.get(studentEmail) : extractNameFromEmail(studentEmail);

  if (!nameLookup.has(studentEmail)) {
    Logger.log(`Name not found in "${EMAILS_SHEET_NAME}" for email: ${studentEmail}. Extracted "${name}" from email.`);
  }

  const timeSubmitted = new Date(timestamp);
  const today = timeSubmitted.toDateString();

  // Determine Time Block for the current submission
  const hour = timeSubmitted.getHours();
  const timeBlock = hour >= 8 && hour < 12 ? "AM" : hour >= 12 && hour <= 23 ? "PM" : "OUT";

  // --- Determine Pass Visual Duration based on destination ---
  let visualDuration = VISUAL_PASS_DURATION_MINUTES;
  if (destination.includes("nurse")) visualDuration = 8;
  else if (destination.includes("locker")) visualDuration = 5;
  else if (destination.includes("bathroom")) visualDuration = 5;
  else if (destination.includes("guidance")) visualDuration = 12;

  // --- Filter relevant passes for the current student and context ---
  const studentPastAndCurrentPassesForBlock = allData.filter((r, idx) => {
    if (idx === 0 || idx === rowIndex) return false; // Skip header and the current submission itself for status checks
    const rEmail = String(r[1]).toLowerCase().trim();
    const rTimestamp = new Date(r[0]);
    const rDate = rTimestamp.toDateString();
    const rHour = rTimestamp.getHours();
    const rTimeBlock = rHour >= 8 && rHour < 12 ? "AM" : rHour >= 12 && rHour <= 23 ? "PM" : "OUT";

    const isSameBlock = rDate === today && rTimeBlock === timeBlock;
    return isSameBlock && rEmail === studentEmail;
  });

  const studentActivePasses = studentPastAndCurrentPassesForBlock.filter(r => r[11] === "ACTIVE");
  const hasActivePass = studentActivePasses.length > 0;

  const studentWaitlistedPasses = studentPastAndCurrentPassesForBlock.filter(r => r[11] === "WAITLISTED");
  const hasWaitlistedPass = studentWaitlistedPasses.length > 0;

  const studentExpiredPasses = studentPastAndCurrentPassesForBlock.filter(r => r[11] === "EXPIRED");
  const usedPassesInBlock = studentExpiredPasses.length;

  Logger.log(`onFormSubmit checks for ${studentEmail} (Row ${row}):`);
  Logger.log(`  - Has Active Pass: ${hasActivePass}`);
  Logger.log(`  - Has Waitlisted Pass: ${hasWaitlistedPass}`);
  Logger.log(`  - Used Passes in Block (EXPIRED): ${usedPassesInBlock}`);
  Logger.log(`  - MAX_PASSES_PER_BLOCK: ${MAX_PASSES_PER_BLOCK}`);
  Logger.log(`  - Time Block: ${timeBlock}`);


  // --- Determine Status and Email Content ---
  let status = "REJECTED"; // Default status
  let emailBody = "";
  let activationTimestamp = ""; // To be stored in Column N
  let waitlistReasonForLog = "";

  if (timeBlock === "OUT") {
    emailBody = createHtml("red", `Hall passes are unavailable at this time.`);
    waitlistReasonForLog = "Outside operating hours";
    Logger.log(`  -> Status: REJECTED - ${waitlistReasonForLog}`);
  } else if (hasActivePass) {
    emailBody = createHtml("red", `You already have an active hall pass. Please wait until it is marked as used.`);
    waitlistReasonForLog = "Already active pass for this student";
    Logger.log(`  -> Status: REJECTED - ${waitlistReasonForLog}`);
  } else if (hasWaitlistedPass) {
    emailBody = createHtml("red", `You already have a pass request pending on the waitlist. Please wait for its activation or end your previous request.`);
    waitlistReasonForLog = "Already waitlisted pass for this student";
    Logger.log(`  -> Status: REJECTED - ${waitlistReasonForLog}`);
  } else if (usedPassesInBlock >= MAX_PASSES_PER_BLOCK) {
    emailBody = createHtml("red", `You have already reached your limit of ${MAX_PASSES_PER_BLOCK} hall passes this ${timeBlock === "AM" ? "morning" : "afternoon"}.`);
    waitlistReasonForLog = "Pass limit reached";
    Logger.log(`  -> Status: REJECTED - ${waitlistReasonForLog}`);
  } else if (exemptStudents.has(studentEmail)) {
    status = "ACTIVE";
    const visualExpiry = new Date(timeSubmitted.getTime() + EXEMPT_VISUAL_DURATION_MINUTES * 60000);
    emailBody = generatePassHtml("green", name, roomNumber, destination, visualExpiry, EXEMPT_VISUAL_DURATION_MINUTES);
    activationTimestamp = timeSubmitted; // Exempt passes are active immediately
    Logger.log(`  -> Status: ACTIVE (Exempt) - ${studentEmail} activated.`);
  } else {
    // Check for active passes in the classroom or among friends (considering all students)
    const studentFriends = friendsLookup.has(studentEmail) ? friendsLookup.get(studentEmail) : new Set();

    let friendHasActivePass = false;
    let nonExemptActivePassInClassroom = false;

    // Filter relevant passes (active, same day, same block, same classroom/friend) from *all* data
    const activePassesInCurrentBlock = allData.filter((r, idx) => {
        if (idx === 0 || idx === rowIndex) return false; // Skip header and the current submission
        const rStatus = r[11];
        const rTimestamp = new Date(r[0]);
        const rDate = rTimestamp.toDateString();
        const rHour = rTimestamp.getHours();
        const rTimeBlock = rHour >= 8 && rHour < 12 ? "AM" : rHour >= 12 && rHour <= 23 ? "PM" : "OUT";

        const isSameBlock = rDate === today && rTimeBlock === timeBlock;
        return isSameBlock && rStatus === "ACTIVE"; // Only care about ACTIVE passes for this check
    });

    if (studentFriends.size > 0) {
      friendHasActivePass = activePassesInCurrentBlock.some(r => studentFriends.has(String(r[1]).toLowerCase().trim()));
      if (friendHasActivePass) Logger.log(`  - Friend of ${studentEmail} has an active pass.`);
    }

    nonExemptActivePassInClassroom = activePassesInCurrentBlock.some(r => r[3] === roomNumber && !exemptStudents.has(String(r[1]).toLowerCase().trim()));
    if (nonExemptActivePassInClassroom) Logger.log(`  - Active pass in classroom ${roomNumber}.`);

    if (friendHasActivePass || nonExemptActivePassInClassroom) {
      status = "WAITLISTED";
      emailBody = generateWaitlistHtml(name, roomNumber, destination, 'N/A', new Date(Date.now() + 5 * 60 * 1000)); // Estimated activation time for email
      if (friendHasActivePass) {
        waitlistReasonForLog = "Friend has active pass";
      } else {
        waitlistReasonForLog = "Active pass in classroom";
      }
      Logger.log(`  -> Status: WAITLISTED - ${studentEmail} waitlisted. Reason: ${waitlistReasonForLog}`);
    } else {
      status = "ACTIVE";
      const visualExpiry = new Date(timeSubmitted.getTime() + visualDuration * 60000);
      emailBody = generatePassHtml("green", name, roomNumber, destination, visualExpiry, visualDuration);
      activationTimestamp = timeSubmitted; // Regular active passes are active immediately
      Logger.log(`  -> Status: ACTIVE - ${studentEmail} approved.`);
    }
  }

  // --- Prepare and perform batch update for the submitted row ---
  // Ensure the rowData array has enough elements for columns L, M, N (index 11, 12, 13)
  const targetCols = 14; // Columns A-N means 14 columns total (indices 0-13)
  while (rowData.length < targetCols) {
    rowData.push(""); // Add empty strings to fill up to the necessary column count
  }

  rowData[2] = name; // Column C - Name (index 2)
  rowData[11] = status; // Column L - Status (index 11)
  rowData[13] = activationTimestamp; // Column N - Activation Timestamp (index 13)

  // Update the single modified row in the sheet
  sheet.getRange(row, 1, 1, rowData.length).setValues([rowData]);
  Logger.log(`onFormSubmit: Row ${row} updated to Status: ${status}, Activation: ${activationTimestamp ? activationTimestamp.toLocaleString() : 'N/A'}`);

  // --- Send Email ---
  try {
    GmailApp.sendEmail('', EMAIL_SUBJECT, 'Your hall pass info', {
      htmlBody: emailBody,
      bcc: studentEmail
    });
    Logger.log(`onFormSubmit: Email sent to ${studentEmail} with status: ${status}.`);
  } catch (emailError) {
    Logger.log(`onFormSubmit: ERROR sending email to ${studentEmail}: ${emailError.message}`);
  }


  // --- Sort the sheet ---
  sortFormSheetNewestToOldest();
}

/**
 * Handles submissions from the "End Pass" Google Form.
 * Marks the corresponding active/waitlisted pass(es) as "EXPIRED" and calculates duration.
 */
function onEndPassFormSubmit(e) {
  Logger.log("onEndPassFormSubmit triggered.");

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const endPassSheet = ss.getSheetByName(END_PASS_SHEET_NAME);
  if (!endPassSheet) {
    Logger.log(`Error: End Pass sheet "${END_PASS_SHEET_NAME}" not found.`);
    return;
  }
  const endPassRow = e.range.getRow();
  const endPassRowData = endPassSheet.getRange(endPassRow, 1, 1, endPassSheet.getLastColumn()).getValues()[0];
  const [, studentEmailRaw] = endPassRowData; // Unpack directly
  const studentEmail = String(studentEmailRaw).toLowerCase().trim();
  Logger.log(`onEndPassFormSubmit: Processing end pass for ${studentEmail}.`);

  const formSheet = ss.getSheetByName(SHEET_NAME);
  if (!formSheet) {
    Logger.log(`Error: Main sheet "${SHEET_NAME}" not found.`);
    return;
  }

  const formSheetRange = formSheet.getDataRange();
  const formData = formSheetRange.getValues();
  const endPassTimestamp = new Date(endPassRowData[0]);

  let passesEndedCount = 0;

  // Process data in memory
  for (let i = 1; i < formData.length; i++) { // Start from 1 to skip header row
    const rowData = formData[i];
    const rowEmail = String(rowData[1]).toLowerCase().trim();
    const status = rowData[11]; // Column L - Status (0-indexed 11)
    const activationTimestampValue = rowData[13]; // Column N - Activation Timestamp (0-indexed 13)
    const originalSubmissionTime = new Date(rowData[0]);

    if (rowEmail === studentEmail && (status === "ACTIVE" || status === "WAITLISTED")) {
      Logger.log(`  - Found matching pass (row ${i + 1}) for ${studentEmail}. Current status: ${status}, Submission: ${originalSubmissionTime.toLocaleString()}.`);
      let startTime;
      if (activationTimestampValue instanceof Date) {
        startTime = activationTimestampValue;
        Logger.log(`    - Using Activation Timestamp (Col N): ${startTime.toLocaleString()}`);
      } else {
        startTime = originalSubmissionTime; // Fallback to submission time if activation is missing
        Logger.log(`    - Warning: Activation timestamp missing for row ${i + 1}. Using Submission Timestamp (Col A): ${startTime.toLocaleString()}.`);
      }

      const durationMillis = endPassTimestamp.getTime() - startTime.getTime();
      const totalSeconds = Math.floor(durationMillis / 1000);
      const minutes = Math.floor(totalSeconds / 60);
      const seconds = totalSeconds % 60;
      const formattedDuration = `${String(minutes).padStart(1, '0')}:${String(seconds).padStart(2, '0')}`;

      // Ensure rowData array has enough length to avoid errors when setting values
      const targetCols = Math.max(rowData.length, 14); // Need space for up to column N (index 13)
      while (rowData.length < targetCols) {
          rowData.push("");
      }

      // Mark the row in memory for update
      rowData[12] = formattedDuration; // Column M - Duration (0-indexed 12)
      rowData[11] = "EXPIRED"; // Column L - Status (0-indexed 11)
      passesEndedCount++;
      Logger.log(`  - Marking row ${i + 1} for ${studentEmail} as EXPIRED. Calculated Duration: ${formattedDuration}.`);
    }
  }

  // Perform batch update on the sheet
  if (passesEndedCount > 0) {
    formSheetRange.setValues(formData); // Write all modified data back to the sheet
    Logger.log(`Successfully updated ${passesEndedCount} passes for ${studentEmail} to EXPIRED.`);
  } else {
    Logger.log(`No active/waitlisted passes found for ${studentEmail} to end.`);
  }

  promoteWaitlist(); // Call promoteWaitlist after updates
  Logger.log("promoteWaitlist() called from onEndPassFormSubmit.");
}

/**
 * Marks active passes as expired if they exceed MAX_ACTIVE_DURATION_MINUTES.
 */
function markExpiredPasses() {
  Logger.log("markExpiredPasses triggered.");
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    Logger.log(`Error: Main sheet "${SHEET_NAME}" not found.`);
    return;
  }

  const dataRange = sheet.getDataRange();
  const data = dataRange.getValues();
  const now = new Date();
  let updatedCount = 0;

  for (let i = 1; i < data.length; i++) { // Start from 1 to skip header row
    const rowData = data[i];
    const status = rowData[11]; // Column L - Status (0-indexed 11)
    const activationTimestampValue = rowData[13]; // Column N - Activation Timestamp (0-indexed 13)
    const studentEmail = String(rowData[1]).toLowerCase().trim(); // Get email for logging

    if (status === "ACTIVE") {
      let startTime;
      if (activationTimestampValue instanceof Date) {
        startTime = activationTimestampValue;
      } else {
        startTime = new Date(rowData[0]); // Fallback to submission time
        Logger.log(`Warning: markExpiredPasses - Activation timestamp missing for row ${i + 1}. Using submission time: ${startTime.toLocaleString()}.`);
      }
      const expiryTime = new Date(startTime.getTime() + MAX_ACTIVE_DURATION_MINUTES * 60000);
      Logger.log(`markExpiredPasses: Checking pass for ${studentEmail} (row ${i + 1}). Activated: ${startTime.toLocaleString()}, Expires: ${expiryTime.toLocaleString()}. Current time: ${now.toLocaleString()}`);

      if (now > expiryTime) {
        // Ensure rowData array has enough length to avoid errors when setting values
        const targetCols = Math.max(rowData.length, 12); // Need space for column L (index 11)
        while (rowData.length < targetCols) {
            rowData.push("");
        }
        rowData[11] = "EXPIRED"; // Update status in memory
        updatedCount++;
        Logger.log(`markExpiredPasses: Pass for ${studentEmail} (row ${i + 1}) expired automatically.`);
      }
    }
  }

  // Perform batch update if any changes were made
  if (updatedCount > 0) {
    dataRange.setValues(data); // Write all modified data back to the sheet
    Logger.log(`markExpiredPasses: Updated ${updatedCount} passes to EXPIRED.`);
  } else {
    Logger.log('markExpiredPasses: No active passes to expire.');
  }
}

/**
 * Promotes waitlisted passes to active, enforcing a "one active pass per classroom" rule.
 */
function promoteWaitlist() {
  Logger.log("promoteWaitlist triggered (enforcing one active pass per classroom).");
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    Logger.log(`Error: Main sheet "${SHEET_NAME}" not found.`);
    return;
  }

  const dataRange = sheet.getDataRange();
  const data = dataRange.getValues(); // Read all data once
  const friendsLookup = getFriendsLookup(); // Load friends data
  const exemptStudents = getExemptStudentsSet(); // Load exempt students

  let promotedCount = 0;
  const updatesMade = new Set(); // To track emails of students that were promoted within this run

  // Track currently active students AND active rooms
  const currentlyActiveStudents = new Set();
  const currentlyActiveRooms = new Set(); // NEW: Track rooms that currently have an active pass (non-exempt)

  // Separate all passes into waitlisted and active categories
  const waitlistedPasses = [];
  // Store relevant info for active passes (room, email, isExempt)
  const allActivePasses = []; 

  for (let i = 1; i < data.length; i++) { // Skip header
    const rowData = data[i];
    const status = rowData[11]; // Column L (0-indexed 11)
    const email = String(rowData[1]).toLowerCase().trim();
    const room = rowData[3];
    const submissionTime = new Date(rowData[0]);

    if (status === "WAITLISTED") {
      waitlistedPasses.push({ rowIdx: i, data: rowData, email: email, submissionTime: submissionTime });
    } else if (status === "ACTIVE") {
      const isExempt = exemptStudents.has(email);
      allActivePasses.push({ room: room, email: email, isExempt: isExempt });
      currentlyActiveStudents.add(email);
      if (!isExempt) { // Only track non-exempt active passes for room availability
        currentlyActiveRooms.add(room); // Add the room if it has an active non-exempt pass
      }
    }
  }

  // Sort waitlisted passes by submission time (oldest first) for FIFO promotion
  waitlistedPasses.sort((a, b) => a.submissionTime.getTime() - b.submissionTime.getTime());
  Logger.log(`promoteWaitlist: Found ${waitlistedPasses.length} waitlisted passes, sorted by submission time.`);
  Logger.log(`promoteWaitlist: Currently active rooms (non-exempt): ${Array.from(currentlyActiveRooms).join(', ') || 'None'}`);

  // Iterate through sorted waitlisted passes to determine if they can be promoted
  for (const pass of waitlistedPasses) {
    const studentEmail = pass.email;
    const roomNumber = pass.data[3];
    const destination = pass.data[4].toLowerCase();
    const rowIdx = pass.rowIdx;
    const originalSubmissionTime = pass.submissionTime;

    Logger.log(`promoteWaitlist - Considering Row ${rowIdx + 1} for promotion. Email: ${studentEmail}, Room: ${roomNumber}, OriginalSubmission: ${originalSubmissionTime.toLocaleString()}`);

    // Skip if student already has an active pass (this check is already in onFormSubmit but good to keep here)
    if (currentlyActiveStudents.has(studentEmail)) {
        Logger.log(`  - Skipping promotion for ${studentEmail} (row ${rowIdx + 1}) - already has an active pass.`);
        continue;
    }

    // --- NEW CRITICAL CHECK: Is this room currently occupied by a non-exempt active pass? ---
    if (currentlyActiveRooms.has(roomNumber)) {
        Logger.log(`  - Skipping promotion for ${studentEmail} (row ${rowIdx + 1}) in Room ${roomNumber} - room already has an active pass.`);
        continue; // This room is occupied, cannot promote
    }
    // --- END NEW CRITICAL CHECK ---

    // Check if a friend currently has an active pass (using allActivePasses)
    let friendHasActivePass = false;
    const studentFriends = friendsLookup.has(studentEmail) ? friendsLookup.get(studentEmail) : new Set();
    if (studentFriends.size > 0) {
      friendHasActivePass = allActivePasses.some(active => studentFriends.has(active.email));
      if (friendHasActivePass) {
          Logger.log(`  - Skipping promotion for ${studentEmail} (row ${rowIdx + 1}) - a friend has an active pass.`);
          continue; // Cannot promote if a friend is active
      }
    }

    // If all conditions are met and this student hasn't been promoted yet within THIS run
    if (!updatesMade.has(studentEmail)) { // This prevents promoting the same student multiple times if they had multiple waitlisted entries
      // Ensure rowData array has enough length to avoid errors when setting values
      const targetCols = Math.max(pass.data.length, 14); // Need space for column N (index 13)
      while (pass.data.length < targetCols) {
          pass.data.push("");
      }

      // Update status and activation timestamp in memory
      pass.data[11] = "ACTIVE"; // Column L - Status
      const newActivationTime = new Date(); // This should be the current time of promotion
      pass.data[13] = newActivationTime; // Column N - Activation Timestamp

      let visualDuration = VISUAL_PASS_DURATION_MINUTES;
      if (destination.includes("nurse")) visualDuration = 8;
      else if (destination.includes("locker")) visualDuration = 5;
      else if (destination.includes("bathroom")) visualDuration = 5;
      else if (destination.includes("guidance")) visualDuration = 12;

      const visualExpiry = new Date(newActivationTime.getTime() + visualDuration * 60000); // Expiry based on NEW activation time
      const name = extractNameFromEmail(studentEmail);
      const html = generatePassHtml("green", name, roomNumber, destination, visualExpiry, visualDuration);

      try {
        GmailApp.sendEmail('', EMAIL_SUBJECT, "Your hall pass is now active", { htmlBody: html, bcc: studentEmail });
        Logger.log(`  - Student ${studentEmail} (row ${rowIdx + 1}) PROMOTED from waitlist to active. Email sent.`);
        Logger.log(`    - Original submission time (Col A): ${originalSubmissionTime.toLocaleString()}`);
        Logger.log(`    - New Activation time (Col N): ${newActivationTime.toLocaleString()}`);
      } catch (emailError) {
        Logger.log(`  - ERROR sending promotion email to ${studentEmail} (row ${rowIdx + 1}): ${emailError.message}`);
      }

      promotedCount++;
      updatesMade.add(studentEmail); // Mark this student as promoted within THIS run
      currentlyActiveStudents.add(studentEmail); // Add to active students set
      if (!exemptStudents.has(studentEmail)) { // NEW: Mark this room as active now
          currentlyActiveRooms.add(roomNumber);
          Logger.log(`    - Room ${roomNumber} is now considered active by ${studentEmail}.`);
      }
      data[rowIdx] = pass.data; // Update the main 'data' array with the modified 'pass.data'

      // IMPORTANT: Since we want only one per classroom, after promoting one,
      // we need to break out of the loop or continue to the next iteration ONLY IF
      // a different classroom is available.
      // The current logic promotes *all* eligible. To promote one *per classroom*
      // AND then stop trying for that classroom, we need to ensure the `currentlyActiveRooms`
      // set is updated and acts as a block. This is already happening.
      // The overall loop will continue to find *other* classrooms that are now open.
      // It will not promote another student for THIS `roomNumber` until the next run
      // of `promoteWaitlist` (after the current pass expires/ends).
      // If you literally mean "only ONE student promoted *across all classrooms* in one run",
      // you would need to add a `break` here: `if (promotedCount >= 1) break;`
      // However, "only one student *per classroom* has an active pass" implies
      // if Room A is open and Room B is open, you want to promote one for A and one for B.
      // The current logic with `currentlyActiveRooms.add(roomNumber)` will do this.
    } else {
        Logger.log(`  - Student ${studentEmail} (row ${rowIdx + 1}) NOT promoted. Conditions not met.`);
    }
  }

  // Perform batch update if any passes were promoted
  if (promotedCount > 0) {
    dataRange.setValues(data); // Write all modified data back to the sheet
    Logger.log(`promoteWaitlist: Promoted ${promotedCount} waitlisted passes and updated sheet.`);
  } else {
    Logger.log('promoteWaitlist: No waitlisted passes could be promoted at this time.');
  }
}

/**
 * Hides passes from previous days to keep the sheet tidy.
 */
function hidePreviousDaysPasses() {
  Logger.log("hidePreviousDaysPasses triggered.");
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const formSheet = ss.getSheetByName(SHEET_NAME);
  if (!formSheet) {
    Logger.log(`Error: Sheet "${SHEET_NAME}" not found for hiding rows.`);
    return;
  }
  const data = formSheet.getDataRange().getValues();
  const today = new Date();
  today.setHours(0, 0, 0, 0); // Set to start of today for comparison

  const rowsToHide = [];
  for (let i = 1; i < data.length; i++) { // Start from 1 to skip header
    const timestampValue = data[i][0];
    if (timestampValue instanceof Date) {
      const passDate = new Date(timestampValue);
      passDate.setHours(0, 0, 0, 0); // Set to start of its day for comparison
      if (passDate.getTime() < today.getTime()) {
        rowsToHide.push(i + 1); // Collect 1-indexed row numbers
      }
    }
  }

  let hiddenCount = 0;
  // Sort rows to hide in descending order to avoid index shifting issues when hiding
  rowsToHide.sort((a, b) => b - a);

  for (const rowNum of rowsToHide) {
    try {
      if (!formSheet.isRowHiddenByUser(rowNum)) { // Check if already hidden to avoid errors/redundancy
        formSheet.hideRows(rowNum);
        hiddenCount++;
        Logger.log(`hidePreviousDaysPasses: Hid row ${rowNum}.`);
      }
    } catch (e) {
      Logger.log(`hidePreviousDaysPasses: Could not hide row ${rowNum}: ${e.message}`);
    }
  }

  if (hiddenCount > 0) {
    Logger.log(`hidePreviousDaysPasses: Hidden ${hiddenCount} previous days' passes.`);
  } else {
    Logger.log('hidePreviousDaysPasses: No previous days\' passes to hide or all already hidden.');
  }
}

/**
 * Sorts the main form sheet by Timestamp (Column A) from newest to oldest.
 */
function sortFormSheetNewestToOldest() {
  Logger.log("sortFormSheetNewestToOldest triggered.");
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const formSheet = ss.getSheetByName(SHEET_NAME);
  if (!formSheet) {
    Logger.log(`Error: Sheet "${SHEET_NAME}" not found for sorting.`);
    return;
  }
  const lastRow = formSheet.getLastRow();
  if (lastRow < 2) {
    Logger.log('sortFormSheetNewestToOldest: No data to sort (less than 2 rows).');
    return;
  }
  // Sort by the first column (Timestamp), descending (newest first)
  formSheet.getRange(2, 1, lastRow - 1, formSheet.getLastColumn()).sort({ column: 1, ascending: false });
  Logger.log(`sortFormSheetNewestToOldest: Sheet "${SHEET_NAME}" sorted from newest to oldest.`);
}

/**
 * Handles manual edits on the spreadsheet to update pass status and send emails.
 * @param {GoogleAppsScript.Events.SheetsOnEdit} e The event object from an onEdit trigger.
 */
function onEdit(e) {
  Logger.log("onEdit triggered.");
  const range = e.range;
  const sheet = range.getSheet();

  // Ensure it's the main Hall Pass sheet
  if (sheet.getName() !== SHEET_NAME) {
    Logger.log(`onEdit: Edit not on target sheet "${SHEET_NAME}". Ignoring.`);
    return;
  }

  const editedColumn = range.getColumn(); // 1-indexed column number
  const editedRow = range.getRow(); // 1-indexed row number

  // Only proceed if it's not the header row and it's the Status column (Column L, index 12)
  if (editedRow > 1 && editedColumn === 12) { // Column L is the 12th column
    const newValue = e.value;
    const oldValue = e.oldValue;
    Logger.log(`onEdit: Status column (L) edited on row ${editedRow}. Old value: "${oldValue}", New value: "${newValue}".`);

    // Load necessary data for decision making
    const studentEmail = String(sheet.getRange(editedRow, 2).getValue()).toLowerCase().trim(); // Column B (Email)
    const roomNumber = sheet.getRange(editedRow, 4).getValue(); // Column D (Room Number)
    const destination = String(sheet.getRange(editedRow, 5).getValue()).toLowerCase().trim(); // Column E (Destination)
    const timestampSubmitted = new Date(sheet.getRange(editedRow, 1).getValue()); // Column A (Timestamp)

    // Load name lookup for email
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const emailsSheet = ss.getSheetByName(EMAILS_SHEET_NAME);
    let nameLookup = new Map();
    if (emailsSheet) {
      const emailsData = emailsSheet.getDataRange().getValues();
      emailsData.forEach(r => {
        const name = r[0];
        const email = String(r[1]).toLowerCase().trim();
        if (email) {
          nameLookup.set(email, name);
        }
      });
    }
    let name = nameLookup.has(studentEmail) ? nameLookup.get(studentEmail) : extractNameFromEmail(studentEmail);

    // Scenario: Manual promotion from WAITLISTED to ACTIVE
    if (oldValue === "WAITLISTED" && newValue === "ACTIVE") {
      Logger.log(`onEdit: Manual promotion detected for ${studentEmail} from WAITLISTED to ACTIVE on row ${editedRow}.`);

      // Set Activation Timestamp (Column N) to current time
      const activationTime = new Date();
      sheet.getRange(editedRow, 14).setValue(activationTime); // Column N is 14th column

      // Recalculate visual duration for the email
      let visualDuration = VISUAL_PASS_DURATION_MINUTES;
      if (destination.includes("nurse")) visualDuration = 8;
      else if (destination.includes("locker")) visualDuration = 5;
      else if (destination.includes("bathroom")) visualDuration = 5;
      else if (destination.includes("guidance")) visualDuration = 12;

      const visualExpiry = new Date(activationTime.getTime() + visualDuration * 60000);
      const html = generatePassHtml("green", name, roomNumber, destination, visualExpiry, visualDuration);

      try {
        GmailApp.sendEmail('', EMAIL_SUBJECT, "Your hall pass is now active", { htmlBody: html, bcc: studentEmail });
        Logger.log(`onEdit: Email sent for manual promotion of ${studentEmail}.`);
      } catch (emailError) {
        Logger.log(`onEdit: ERROR sending email for manual promotion to ${studentEmail}: ${emailError.message}`);
      }

      promoteWaitlist(); // This will re-evaluate the waitlist after this manual change
      Logger.log("promoteWaitlist() called from onEdit after manual status change.");

    } else if (oldValue === "ACTIVE" && newValue === "EXPIRED") {
        Logger.log(`onEdit: Manual expiration detected for ${studentEmail} from ACTIVE to EXPIRED on row ${editedRow}.`);

        // Get activation timestamp to calculate duration
        const activationTimestamp = sheet.getRange(editedRow, 14).getValue(); // Column N

        if (activationTimestamp instanceof Date) {
            const durationMillis = new Date().getTime() - activationTimestamp.getTime();
            const totalSeconds = Math.floor(durationMillis / 1000);
            const minutes = Math.floor(totalSeconds / 60);
            const seconds = totalSeconds % 60;
            const formattedDuration = `${String(minutes).padStart(1, '0')}:${String(seconds).padStart(2, '0')}`;
            sheet.getRange(editedRow, 13).setValue(formattedDuration); // Column M - Duration
            Logger.log(`onEdit: Duration ${formattedDuration} calculated for ${studentEmail} on manual expiration.`);
        } else {
            Logger.log(`onEdit: Could not calculate duration for ${studentEmail}. Activation timestamp (Col N) is missing or not a Date.`);
        }
        promoteWaitlist(); // A manual expiration might open a slot
        Logger.log("promoteWaitlist() called from onEdit after manual expiration.");
    } else {
        Logger.log(`onEdit: Status change from "${oldValue}" to "${newValue}" on row ${editedRow} for ${studentEmail} did not trigger specific action.`);
    }
  } else {
    Logger.log(`onEdit: Edit not in Status column (L) or is header row. Ignoring.`);
  }
}

// --- Setup Function (Run this once manually to initialize Script Properties) ---
function setup() {
  updateConstants();
  Logger.log("Setup complete. Remember to set up all necessary installable triggers.");
}
