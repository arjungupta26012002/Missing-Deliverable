const MASTER_SPREADSHEET_ID = "sheetid"; 
const REF_SHEET_NAME = "RefID";
const SIGNATURE_IMAGE_FILE_ID = "imgid"; 
const MAIL_LOGS_SPREADSHEET_ID = "logsid"; 
const MAIL_LOGS_SHEET_NAME = "MailLogs"; 
const URL_REGEX = /^(https?|ftp):\/\/[^\s/$.?#].[^\s]*$/i; 

function doGet() {
  return HtmlService.createTemplateFromFile('index')
    .evaluate()
    .setTitle('Internship Submission Reminders'); 
}

function include(filename) {
  return HtmlService.createTemplateFromFile(filename).evaluate().getContent();
}

function getInternshipList() {
  Logger.log("Attempting to fetch internship list.");
  try {
    const ss = SpreadsheetApp.openById(MASTER_SPREADSHEET_ID);
    const masterSheet = ss.getSheetByName(REF_SHEET_NAME);

    if (!masterSheet) {
      throw new Error(`Sheet "${REF_SHEET_NAME}" not found in the master spreadsheet (ID: ${MASTER_SPREADSHEET_ID}).`);
    }

    const range = masterSheet.getDataRange();
    const values = range.getValues();

    const internshipData = values.slice(1).map(row => ({
      name: row[0],
      id: row[1]
    }));

    Logger.log(`Successfully fetched ${internshipData.length} internships.`);
    return internshipData;
  } catch (e) {
    Logger.log(`Error in getInternshipList: ${e.message}`);
    throw new Error(`Could not fetch internship list: ${e.message}. Please check the master spreadsheet ID and sheet name.`);
  }
}

function getCohortList(internshipId) {
  Logger.log(`Attempting to fetch cohort list for Internship ID: ${internshipId}`);
  if (!internshipId) {
    throw new Error("Internship ID is required to fetch cohorts.");
  }
  try {
    const ss = SpreadsheetApp.openById(internshipId);
    const sheets = ss.getSheets();
    const sheetNames = sheets.map(sheet => sheet.getName());
    Logger.log(`Successfully fetched ${sheetNames.length} cohorts for Internship ID: ${internshipId}`);
    return sheetNames;
  } catch (e) {
    Logger.log(`Error in getCohortList for ID ${internshipId}: ${e.message}`);
    throw new Error(`Could not fetch cohorts for the selected internship: ${e.message}. Please ensure the ID is correct and you have access.`);
  }
}

function parseDdMmToDateString(ddmmString) {

  if (typeof ddmmString !== 'string' || !/^\d{4}$/.test(ddmmString)) {
    Logger.log(`Invalid DDMM string format provided: "${ddmmString}". Expected a 4-digit number.`);
    return ddmmString; 
  }

  const day = parseInt(ddmmString.substring(0, 2), 10);
  const month = parseInt(ddmmString.substring(2, 4), 10) - 1; 

  const currentYear = new Date().getFullYear(); 

  const date = new Date(currentYear, month, day);

  if (isNaN(date.getTime()) || date.getDate() !== day || date.getMonth() !== month) {
    Logger.log(`Could not parse "${ddmmString}" into a valid date. Resulted in invalid date: ${date}`);
    return ddmmString; 
  }

  const options = { month: 'long', day: 'numeric' };
  let formattedDate = date.toLocaleDateString('en-US', options);

  const dayNum = date.getDate();
  let suffix = 'th';
  if (dayNum % 10 === 1 && dayNum % 100 !== 11) {
    suffix = 'st';
  } else if (dayNum % 10 === 2 && dayNum % 100 !== 12) {
    suffix = 'nd';
  } else if (dayNum % 10 === 3 && dayNum % 100 !== 13) {
    suffix = 'rd';
  }
  formattedDate = formattedDate.replace(/(\d+)$/, `$1${suffix}`);

  return formattedDate;
}

function formatDeadlineDate(dateString) {
  if (!dateString || typeof dateString !== 'string' || !/^\d{4}-\d{2}-\d{2}$/.test(dateString)) {
    Logger.log(`Invalid date string format provided for deadline: "${dateString}". Expected<\ctrl42>-MM-DD.`);
    return dateString; 
  }

  const parts = dateString.split('-');
  const year = parseInt(parts[0], 10);
  const month = parseInt(parts[1], 10) - 1; 
  const day = parseInt(parts[2], 10);

  const date = new Date(year, month, day);

  if (isNaN(date.getTime()) || date.getDate() !== day || date.getMonth() !== month || date.getFullYear() !== year) {
    Logger.log(`Could not parse "${dateString}" into a valid date for deadline formatting.`);
    return dateString; 
  }

  const options = { day: 'numeric', month: 'long', year: 'numeric' };
  return date.toLocaleDateString('en-US', options); 
}

function isValidEmail(email) {
  if (typeof email !== 'string' || email.trim() === '') {
    return false;
  }
  const emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
  return emailRegex.test(email.trim());
}

function generateEmailBody(data) {
  return `
    <div style="font-family: Arial, sans-serif; line-height: 1.6; color: #333;">
      <p>Dear ${data.internName},</p>
      <p>Hope you're doing well!</p>
      <p>This is a friendly reminder regarding your pending submissions for the **${data.internshipName} - ${data.cohortNumber}** internship program. The final deadline for all submissions is **${data.deadline}**.</p>
      <p>It looks like you still need to submit the following items:</p>
      <ul style="list-style-type: disc; margin-left: 20px;">
        ${data.missingItemsHtml}
      </ul>
      <p>Please make sure to complete these as soon as possible to ensure your successful completion of the internship.</p>
      <p>If you have already submitted these items, please disregard this email. If you believe this is an error or have any questions, please reply to this email.</p>
      <br>
      <p>Best regards,</p>
      <p>${data.yourName}</p>
      <img src="cid:signatureImage" alt="Signature" style="max-width: 200px; height: auto;">
    </div>
  `;
}

function sendPreviewToCoordinators(options) {
  Logger.log("Attempting to send preview mail with options: " + JSON.stringify(options));

  const {
    internshipId,
    internshipName,
    cohortName,
    deadline,
    coordinatorMails,
    feedbackFormLink,
    peerEvalFormLink,
    selfEvalFormLink,
    yourName
  } = options;

  if (!internshipId || !internshipName || !cohortName || !deadline || !coordinatorMails || !yourName) {
    throw new Error("Missing required parameters for sending preview mail. Coordinator emails and your name are also required for preview.");
  }

  if (feedbackFormLink && !URL_REGEX.test(feedbackFormLink)) Logger.log(`Warning: Invalid URL for Feedback Form in preview: ${feedbackFormLink}`);
  if (peerEvalFormLink && !URL_REGEX.test(peerEvalFormLink)) Logger.log(`Warning: Invalid URL for Peer Evaluation Form in preview: ${peerEvalFormLink}`);
  if (selfEvalFormLink && !URL_REGEX.test(selfEvalFormLink)) Logger.log(`Warning: Invalid URL for Self Evaluation Form in preview: ${selfEvalFormLink}`);

  const ss = SpreadsheetApp.openById(internshipId);
  const cohortSheet = ss.getSheetByName(cohortName);

  if (!cohortSheet) {
    throw new Error(`Cohort sheet "${cohortName}" not found in internship spreadsheet (ID: ${internshipId}).`);
  }

  const cohortNumberMatch = cohortName.match(/\b(\d{4})\b/);
  const cohortNumberRaw = cohortNumberMatch ? cohortNumberMatch[1] : cohortName;
  const cohortStartDateFormatted = parseDdMmToDateString(cohortNumberRaw);

  const formattedDeadline = formatDeadlineDate(deadline);

  const displayNamesMap = {
    'Submitted Feedback Form': 'Feedback Form',
    'Submitted Reflective Video': 'Reflective Video',
    'Submitted Peer Evaluation Form': 'Peer Evaluation Form',
    'Submitted Self Evaluation Form': 'Self Evaluation Form'
  };

  const submissionCheckItems = [
    { header: 'Submitted Feedback Form', link: feedbackFormLink },
    { header: 'Submitted Reflective Video', link: '' },
    { header: 'Submitted Peer Evaluation Form', link: peerEvalFormLink },
    { header: 'Submitted Self Evaluation Form', link: selfEvalFormLink }
  ];

  const headers = cohortSheet.getRange(1, 1, 1, cohortSheet.getLastColumn()).getValues()[0];
  headers.forEach(header => {
    if (typeof header === 'string' && header.toLowerCase().includes('week')) {
      submissionCheckItems.push({ header: header, link: '' });
      displayNamesMap[header] = header; 
    }
  });

  let sampleMissingItemsHtml = '';
  submissionCheckItems.forEach(item => {
    const displayName = displayNamesMap[item.header] || item.header; 
    if (item.link && item.link.trim() !== '') {
      sampleMissingItemsHtml += `<li><a href="${item.link}" target="_blank">${displayName}</a></li>`;
    } else {
      sampleMissingItemsHtml += `<li>${displayName}</li>`;
    }
  });

  const emailSubject = `Pending Submissions Reminder for ${cohortNumberRaw} ${internshipName}`; 
  const emailBody = generateEmailBody({
    internName: "Team", 
    internshipName,
    cohortNumber: cohortStartDateFormatted,
    deadline: formattedDeadline, 
    missingItemsHtml: sampleMissingItemsHtml,
    feedbackFormLink,
    peerEvalFormLink,
    selfEvalFormLink,
    yourName 
  });

  let signatureImageBlob = null;
  try {
    signatureImageBlob = DriveApp.getFileById(SIGNATURE_IMAGE_FILE_ID).getBlob();
  } catch (e) {
    Logger.log(`Error loading signature image (ID: ${SIGNATURE_IMAGE_FILE_ID}): ${e.message}. Email signature image might not appear in preview.`);
  }

  const recipients = coordinatorMails.split(',').map(e => e.trim()).filter(email => {
    if (isValidEmail(email)) {
      return true;
    } else {
      Logger.log(`Skipping invalid coordinator email for preview: "${email}"`);
      return false;
    }
  });

  if (recipients.length === 0) {
    throw new Error("No valid coordinator email addresses provided to send the preview to.");
  }

  const mailOptions = {
    to: recipients.join(','),
    subject: emailSubject,
    htmlBody: emailBody,
    inlineImages: signatureImageBlob ? { signatureImage: signatureImageBlob } : undefined
  };

  Logger.log(`Attempting to send PREVIEW mail.`);
  Logger.log(`Preview Mail TO: ${mailOptions.to}`);
  Logger.log(`Preview Mail Subject: ${mailOptions.subject}`);

  MailApp.sendEmail(mailOptions);
  Logger.log(`Preview mail successfully sent to: ${mailOptions.to}`);
  return "Preview email sent successfully to coordinators!";
}

function logMail(sentCount, mailSubject, senderEmail, internshipId, cohortName) { 
  try {
    const logSs = SpreadsheetApp.openById(MAIL_LOGS_SPREADSHEET_ID);
    const logSheet = logSs.getSheetByName(MAIL_LOGS_SHEET_NAME);

    if (!logSheet) {
      Logger.log(`Error: Mail Logs sheet "${MAIL_LOGS_SHEET_NAME}" not found in spreadsheet (ID: ${MAIL_LOGS_SPREADSHEET_ID}). Cannot log mail summary.`);
      return;
    }

    if (logSheet.getLastRow() === 0) {
      logSheet.appendRow(['Timestamp', 'Mails Sent Count', 'Mail Subject', 'Sender Email', 'Internship ID', 'Cohort Name']); 
    }

    const timestamp = new Date();
    const rowData = [
      timestamp,
      sentCount,
      mailSubject,
      senderEmail,
      internshipId, 
      cohortName 
    ];

    logSheet.appendRow(rowData);
    Logger.log(`Mail summary logged: ${sentCount} mails sent by ${senderEmail} for Internship ID ${internshipId}, Cohort ${cohortName}.`);
  } catch (e) {
    Logger.log(`Error logging mail summary activity: ${e.message}`);
  }
}

function countSubmittedWeeks(rowData, submissionColInfos) {
  let submittedWeeksCount = 0;
  submissionColInfos.forEach(col => {

    if (typeof col.header === 'string' && col.header.toLowerCase().includes('week')) {
      const status = String(rowData[col.index]).trim().toLowerCase();
      if (status === 'true' || status === 'yes') {
        submittedWeeksCount++;
      }
    }
  });
  return submittedWeeksCount;
}

function sendMailsToInterns(options) {
  Logger.log("Attempting to send mails to interns with options: " + JSON.stringify(options));

  const {
    internshipId,
    internshipName,
    cohortName,
    deadline,
    coordinatorMails,
    feedbackFormLink,
    peerEvalFormLink,
    selfEvalFormLink,
    yourName
  } = options;

  if (!internshipId || !internshipName || !cohortName || !deadline || !yourName) {
    throw new Error("Missing common required parameters for sending mail to interns.");
  }

  if (feedbackFormLink && !URL_REGEX.test(feedbackFormLink)) Logger.log(`Warning: Invalid URL for Feedback Form: ${feedbackFormLink}`);
  if (peerEvalFormLink && !URL_REGEX.test(peerEvalFormLink)) Logger.log(`Warning: Invalid URL for Peer Evaluation Form: ${peerEvalFormLink}`);
  if (selfEvalFormLink && !URL_REGEX.test(selfEvalFormLink)) Logger.log(`Warning: Invalid URL for Self Evaluation Form: ${selfEvalFormLink}`);

  let sentCount = 0;
  let totalProcessedInterns = 0; 
  let commonMailSubjectForLog = '';

  try {
    const ss = SpreadsheetApp.openById(internshipId);
    const cohortSheet = ss.getSheetByName(cohortName);

    if (!cohortSheet) {
      throw new Error(`Cohort sheet "${cohortName}" not found in internship spreadsheet (ID: ${internshipId}).`);
    }

    const headers = cohortSheet.getRange(1, 1, 1, cohortSheet.getLastColumn()).getValues()[0];
    const values = cohortSheet.getDataRange().getValues();
    const rawInternData = values.slice(1); 

    const internData = rawInternData.filter(row => {
      let filledCellsCount = 0;
      for (let i = 0; i < row.length; i++) {
        if (row[i] !== null && String(row[i]).trim() !== '') {
          filledCellsCount++;
        }
      }
      return filledCellsCount >= 5;
    });

    totalProcessedInterns = internData.length; 

    const nameCol = headers.indexOf('Full Name');
    const emailCol = headers.indexOf('Email Address');
    const statusCol = headers.findIndex(header => typeof header === 'string' && header.toLowerCase().includes('status'));

    const displayNamesMap = {
      'Submitted Feedback Form': 'Feedback Form',
      'Submitted Reflective Video': 'Reflective Video',
      'Submitted Peer Evaluation Form': 'Peer Evaluation Form',
      'Submitted Self Evaluation Form': 'Self Evaluation Form'
    };

    const formLinkMap = {
      'Submitted Feedback Form': feedbackFormLink,
      'Submitted Reflective Video': feedbackFormLink, 
      'Submitted Peer Evaluation Form': peerEvalFormLink,
      'Submitted Self Evaluation Form': selfEvalFormLink
    };

    const submissionCheckHeaders = [
      'Submitted Feedback Form',
      'Submitted Reflective Video',
      'Submitted Peer Evaluation Form',
      'Submitted Self Evaluation Form'
    ];
    headers.forEach(header => {
      if (typeof header === 'string' && header.toLowerCase().includes('week')) {
        submissionCheckHeaders.push(header);
        displayNamesMap[header] = header;
      }
    });

    const submissionColInfos = submissionCheckHeaders
      .map(headerName => ({ header: headerName, index: headers.indexOf(headerName) }))
      .filter(col => col.index !== -1);

    if (nameCol === -1 || emailCol === -1) {
      throw new Error("Essential columns 'Full Name' or 'Email Address' not found in the cohort sheet. Please ensure these headers exist.");
    }

    const cohortNumberMatch = cohortName.match(/\b(\d{4})\b/);
    const cohortNumberRaw = cohortNumberMatch ? cohortNumberMatch[1] : cohortName;
    const cohortStartDateFormatted = parseDdMmToDateString(cohortNumberRaw);
    const formattedDeadline = formatDeadlineDate(deadline);

    commonMailSubjectForLog = `Pending Submissions Reminder for ${cohortNumberRaw} ${internshipName}`;

    let signatureImageBlob = null;
    try {
      signatureImageBlob = DriveApp.getFileById(SIGNATURE_IMAGE_FILE_ID).getBlob();
    } catch (e) {
      Logger.log(`Error loading signature image (ID: ${SIGNATURE_IMAGE_FILE_ID}): ${e.message}. Email signature image might not appear.`);
    }

    let finalCcRecipients = [];
    const ownEmail = Session.getActiveUser().getEmail();
    if (isValidEmail(ownEmail)) {
      finalCcRecipients.push(ownEmail);
      Logger.log(`Added sender's email (${ownEmail}) to CC list.`);
    } else {
      Logger.log(`Warning: Sender's email (${ownEmail}) is invalid and will not be added to CC.`);
    }

    if (coordinatorMails) {
      const parsedCoordinatorMails = coordinatorMails.split(',').map(e => e.trim()).filter(e => {
        if (isValidEmail(e)) {
          return true;
        } else {
          Logger.log(`Skipping invalid coordinator CC email: "${e}"`);
          return false;
        }
      });
      finalCcRecipients = finalCcRecipients.concat(parsedCoordinatorMails);
    }

    internData.forEach((row, index) => {
      const internName = row[nameCol];
      const rawInternEmail = row[emailCol];
      const internEmail = typeof rawInternEmail === 'string' ? rawInternEmail.trim() : String(rawInternEmail).trim();

      if (!internName || !internEmail) {
        Logger.log(`Skipping intern record (index ${index} in filtered data): Missing intern name or email.`);

        return;
      }

      if (!isValidEmail(internEmail)) {
        Logger.log(`Skipping ${internName} (${internEmail}): Invalid email format found.`);
        return;
      }

      if (statusCol !== -1) {
        const statusValue = String(row[statusCol]).trim();
        if (statusValue !== '') {
          Logger.log(`Skipping ${internName} (${internEmail}): Status is "${statusValue}", not sending email.`);
          return;
        }
      }

      const submittedWeeks = countSubmittedWeeks(row, submissionColInfos);
      if (submittedWeeks < 2) {
        Logger.log(`Skipping ${internName} (${internEmail}): Only ${submittedWeeks} week(s) submitted. Requires at least 2.`);
        return;
      }

      let missingItemsHtml = '';

      submissionColInfos.forEach(col => {
        const status = String(row[col.index]).trim().toLowerCase();
        if (!(status === 'true' || status === 'yes')) {
          const displayName = displayNamesMap[col.header] || col.header;
          const itemLink = formLinkMap[col.header] || '';
          if (itemLink && itemLink.trim() !== '') {
            missingItemsHtml += `<li><a href="${itemLink}" target="_blank">${displayName}</a></li>`;
          } else {
            missingItemsHtml += `<li>${displayName}</li>`;
          }
        }
      });

      Logger.log(`Intern: ${internName}, Email: ${internEmail}, Missing Items: ${missingItemsHtml ? 'Yes' : 'No'}`);

      if (missingItemsHtml) { 
        const emailSubject = `Pending Submissions Reminder for ${cohortNumberRaw} ${internshipName}`;
        const emailBody = generateEmailBody({
          internName,
          internshipName,
          cohortNumber: cohortStartDateFormatted,
          deadline: formattedDeadline,
          missingItemsHtml,
          feedbackFormLink,
          peerEvalFormLink,
          selfEvalFormLink,
          yourName
        });

        const mailOptions = {
          to: internEmail,
          subject: emailSubject,
          htmlBody: emailBody,
          inlineImages: signatureImageBlob ? { signatureImage: signatureImageBlob } : undefined
        };

        if (finalCcRecipients.length > 0) {
          mailOptions.cc = finalCcRecipients.join(',');
        } else {
          delete mailOptions.cc;
        }

        MailApp.sendEmail(mailOptions);
        sentCount++;
        Logger.log(`Reminder email successfully sent to ${internName} (${internEmail}).`);

      } else {
        Logger.log(`Skipping ${internName} (${internEmail}): No pending submissions found.`);

      }
    });

    skippedCount = totalProcessedInterns - sentCount;
    Logger.log(`Mail sending complete. Sent: ${sentCount}, Skipped: ${skippedCount} (out of ${totalProcessedInterns} interns considered).`);

    if (sentCount > 0) {
      logMail(sentCount, commonMailSubjectForLog, ownEmail, internshipId, cohortName);
    } else {
      Logger.log("No mails were sent, skipping mail summary log.");
    }

    return { sentCount: sentCount, skippedCount: skippedCount };

  } catch (e) {
    Logger.log(`Critical error sending mails to interns: ${e.message}`);
    if (e.message.includes("Invalid email")) {
      throw new Error(`Failed to send mails: ${e.message}. Please check your email addresses in the sheet or coordinator email input.`);
    } else {
      throw new Error(`Failed to send mails: ${e.message}. See script logs for details.`);
    }
  }
}
