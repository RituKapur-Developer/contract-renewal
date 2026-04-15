function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('Contract Reminder App');
}
function runDailyReminders() {
  const defaultEmails = 'user@gmail.com'; //userGmail
  sendRenewalRemindersUI(defaultEmails);
}
function updateContract(rowIndex, clientName, renewalDate, notes) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

  sheet.getRange(rowIndex, 1).setValue(clientName);
  sheet.getRange(rowIndex, 3).setValue(new Date(renewalDate));
  sheet.getRange(rowIndex, 5).setValue(notes);

  return "✅ Updated successfully";
}
function deleteContract(rowIndex) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  sheet.deleteRow(rowIndex);

  return "🗑️ Deleted successfully";
}
function getContracts(filterType) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const data = sheet.getDataRange().getValues();

  const today = new Date();
  today.setHours(0,0,0,0);

  let results = [];

  for (let i = 1; i < data.length; i++) {
    const clientName = data[i][0];
    const renewalDate = new Date(data[i][2]);
    const reminderSent = data[i][3];
    const notes = data[i][4];

    if (!renewalDate) continue;

    renewalDate.setHours(0,0,0,0);

    const diffDays = Math.floor((renewalDate - today) / (1000 * 60 * 60 * 24));

    let include = false;

    if (filterType === "ALL") include = true;
    if (filterType === "EXPIRING" && diffDays <= 30 && diffDays >= 0) include = true;
    if (filterType === "OVERDUE" && diffDays < 0) include = true;

    if (include) {
      results.push({
      rowIndex: i + 1, // VERY IMPORTANT
      clientName,
      renewalDate: renewalDate.toDateString(),
      diffDays,
      reminderSent: reminderSent ? "Yes" : "No",
      notes: notes || ""
    });
    }
  }

  return results;
}
function addContract(clientName, renewalDate, notes) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const originalDate = new Date(renewalDate);
  originalDate.setHours(0, 0, 0, 0);
  // Clone date (VERY IMPORTANT to avoid modifying original)
  const newDate = new Date(originalDate);
  newDate.setMonth(newDate.getMonth() + 11);
  // Fix edge case (e.g., Feb)
  if (newDate.getDate() < day) {
    newDate.setDate(0); // last day of previous month
  }
  newDate.setHours(0, 0, 0, 0);
  sheet.appendRow([
    clientName,
    originalDate,
    newDate,   // ✅ +11 months date
    "",        // reminder sent
    notes
  ]);

  return "✅ Contract added successfully!";
}
function sendRenewalRemindersUI(emails) {
  highlightExpiringContracts();

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const data = sheet.getDataRange().getValues();

  const email = emails || 'user@gmail.com';

  const today = new Date();
  today.setHours(0,0,0,0);

  let sentCount = 0;

  for (let i = 1; i < data.length; i++) {
    const clientName = data[i][0];
    const renewalDate = new Date(data[i][2]);
    const reminderSent = data[i][3];
    const notes = data[i][4];

    if (!renewalDate || reminderSent) continue;

    renewalDate.setHours(0,0,0,0);

    const diffDays = Math.floor((renewalDate - today) / (1000 * 60 * 60 * 24));

    if ([30, 7, 1, 0].includes(diffDays)) {

      const subject = `🔔 Contract Renewal Reminder (${diffDays} day${diffDays === 1 ? '' : 's'} left)`;

      const message = `
Hi,

This is a reminder that the contract with ${clientName} is due for renewal.

📅 Renewal Date: ${renewalDate.toDateString()}
⏳ Days Remaining: ${diffDays}

${notes ? "📝 Notes: " + notes : ""}

Please take necessary action.

Thanks
`;

      MailApp.sendEmail({
        to: email,
        subject: subject,
        body: message
      });

      sheet.getRange(i + 1, 4).setValue(new Date());
      sentCount++;
    }
  }

  return `✅ ${sentCount} reminder(s) sent successfully`;
}

function highlightExpiringContracts() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const data = sheet.getDataRange().getValues();

  const today = new Date();
  today.setHours(0,0,0,0);

  for (let i = 1; i < data.length; i++) {
    const renewalDate = new Date(data[i][2]);
    renewalDate.setHours(0,0,0,0);

    const diffDays = Math.floor((renewalDate - today) / (1000 * 60 * 60 * 24));

    const row = sheet.getRange(i + 1, 1, 1, sheet.getLastColumn());

    if (diffDays <= 30 && diffDays >= 0) {
      row.setBackground("#ffe6e6"); // light red
    } else {
      row.setBackground(null);
    }
  }
}
