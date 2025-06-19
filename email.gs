// Email configuration
const FROM_EMAIL = 'lavanyachawla24@gmail.com';
const EMAIL_SHEET_NAME = 'Sheet1';

function sendInterviewEmails() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(EMAIL_SHEET_NAME);
  if (!sheet) {
    throw new Error('Sheet not found');
  }

  // Get all data from the sheet
  const data = sheet.getDataRange().getValues();
  const headers = data[0];

  // Find required column indexes
  const statusColIndex = headers.indexOf('Status');
  const nameColIndex = 1; // Column B
  const emailColIndex = 7; // Column H
  const roleColIndex = 8; // Column I
  const dateColIndex = 14; // Column O
  const timeColIndex = 15; // Column P
  const locationColIndex = 16; // Column Q

  // Verify all required columns exist
  if (statusColIndex === -1) {
    throw new Error('Status column not found');
  }

  // Process each row starting from row 2 (skip header)
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const status = row[statusColIndex];

    // Check if status is "Interview Appointed"
    if (status === 'Interview Appointed' ) {
      const name = row[nameColIndex];
      const email = row[emailColIndex];
      const role = row[roleColIndex];
      const date = formatDate(row[dateColIndex]);
      const time = formatTime(row[timeColIndex]);
      const location = row[locationColIndex];

      // Skip if email is missing
      if (!email) {
        Logger.log(`Skipping row ${i + 1}: No email address found`);
        continue;
      }

      // Create email content
      const subject = 'Interview Confirmation Mail';
      const htmlBody = `
        <p>Dear ${name},</p>
        <p>As per our recent conversation, We are pleased to inform you that your CV has been shortlisted for the interview for the position of ${role} Intern at iTalent India Management Consultants Pvt. Ltd.</p>
        <p>We have scheduled a physical round of interview as per the details below:</p>
        <br>
        <p><strong>Date:</strong> ${date}</p>
        <p><strong>Time:</strong> ${time}</p>
        <p><strong>Company:</strong> iTalent India Management Consultants Pvt. Ltd. | Bizgrow Technology (OPC)</p>
        <p><strong>Meeting Point:</strong> ${location}</p>
        <br>
        <p>We request you to kindly confirm your availability for the interview by replying to this email and acknowledging the same.</p>
        <br>
        <p><strong>Note:</strong> Please carry an updated copy of your CV with you for the interview.</p>
        <p>Looking forward to your confirmation.</p>
        <p>Please contact me in case of any query.</p>
        <br>
        <p>Thanks & Regards,<br>
        Manisha Tekam</p>
      `;

      // Send email
      try {
        MailApp.sendEmail({
          to: email,
          from: FROM_EMAIL,
          subject: subject,
          htmlBody: htmlBody
        });
        Logger.log(`Email sent successfully to ${email}`);
          
        // Update status to "Mails Sent" in column M (13th column, 0-based index 12)
        sheet.getRange(i + 1, 13).setValue("Mails Sent");
        
      } catch (error) {
        Logger.log(`Failed to send email to ${email}: ${error.message}`);
      }
    }
  }
}

// Helper function to format date
function formatDate(date) {
  if (!date) return 'abc';
  if (typeof date === 'string') return date;
  return Utilities.formatDate(date, Session.getScriptTimeZone(), 'MMMM dd, yyyy');
}

// Helper function to format time
function formatTime(time) {
  if (!time) return 'abc';
  if (typeof time === 'string') return time;
  return Utilities.formatDate(time, Session.getScriptTimeZone(), 'hh:mm a');
}

// Add menu item to run the email sender
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Email Functions')
    .addItem('Send Interview Emails', 'sendInterviewEmails')
    .addToUi();
}