const COL_FINAL_EMAIL_STATUS = "Final Email Status";
const COL_FINAL_STATUS = "Final Status";

const COL_JOB_POSITION = "Which position are you applying for?";
const COL_EMAIL_ADDRESS = "Email Address";
const COL_EXTERNAL_FEEDBACK = "External Feedback";

const DEFAULT_FEEDBACK = "Understanding of technical concepts and/or experience not a fit";

const STATUS = {
  REJECTED: "Rejected",
  WAITLISTED: "Waitlisted"
};

function getCellValueByColumnName(row, colName, header) {
  const col = header.indexOf(colName);
  if (col != -1) {
    return row[col];
  }
}

function shouldSendStatus(row, header) {
  const finalStatus = getCellValueByColumnName(row, COL_FINAL_STATUS, header);
  if (finalStatus != STATUS.REJECTED && finalStatus != STATUS.WAITLISTED) {
    return false;
  }

  const emailStatus = getCellValueByColumnName(row, COL_FINAL_EMAIL_STATUS, header);
  if (emailStatus == "Sent") {
    return false;
  }
  return true;
}

function validate(candidateEmail, jobPosition) {
  if (candidateEmail == "") {
    return { status: false, err: "email not provided" };
  }

  if (jobPosition == "") {
    return { status: false, err: "position not provided" };
  }

  return { status: true, err: undefined }
}

function getRejectedHtmlBody(feedback = DEFAULT_FEEDBACK) {
  return `<p>Dear Candidate,</p>
<div>&nbsp;</div>
<div>Thank you very much for investing your time and effort to apply&nbsp;for an internship position at Gramoday.</div>
<div>&nbsp;</div>
<div>Unfortunately, at this time, we decided to proceed with our selection process with another candidate.&nbsp; We have the below feedback from our selection panel:</div>
<div>\"${feedback}\"</div>
<div>&nbsp;</div>
<div>Please follow our&nbsp;linkedin&nbsp;page for future opportunities :&nbsp;<a href="https://www.linkedin.com/company/agrilinks-technologies/" target="_blank" rel="noopener" data-saferedirecturl="https://www.google.com/url?q=https://www.linkedin.com/company/agrilinks-technologies/&amp;source=gmail&amp;ust=1642424959596000&amp;usg=AOvVaw0K-YGeMWceo55biB5PRhES">https://www.linkedin.com/<wbr />company/agrilinks-<wbr />technologies/</a><br /><br /></div>
<div>I wish you the best of luck in your future endeavors and hope we'll have a chance to meet again soon.</div>
</div>
<div>&nbsp;</div>
<div>Regards,</div>
<div>Gramoday Team</div>`;
}

function getWaitlistedHtmlBody(jobPosition) {
  return `<p>Dear Candidate,</p>
<div>&nbsp;</div>
<div>Thank you very much for investing your time and effort to apply&nbsp;for ${jobPosition} at Gramoday.</div>
<div><br />We really enjoyed meeting you, learning about your skills and experiences and having a really interesting conversation.</div>
<div><br />Unfortunately, at this time, we decided to proceed with our selection process with another candidate.<br /><br /></div>
<div>For now, we have kept your candidature as&nbsp;<strong>"<span class="il">waitlisted</span>"&nbsp;</strong>which means that in case we have an opening that better fits your profile, we will make sure to get in touch with you, and you will be automatically&nbsp;<strong>shortlisted</strong>.<br /><br /></div>
<div>I wish you the best of luck in your future endeavours and hope we'll have a chance to meet again soon.</div>
<div>&nbsp;</div>
<div>Regards,</div>
<div>Gramoday Team</div>`;
}

function sendEmail(email, jobPosition, candStatus, candFeedback) {
  const { status, err } = validate(email, jobPosition);
  if (!status) {
    Logger.log("Invalid data: ", err);
    return false;
  }

  MailApp.sendEmail({
    to: email,
    cc: INTERVIEWER_EMAIL,
    subject: `[Update] Gramoday - ${jobPosition}`,
    htmlBody: candStatus == STATUS.REJECTED ? getRejectedHtmlBody(candFeedback) : getWaitlistedHtmlBody(jobPosition)
  });
  return true;
}

function sendUpdate(row, header) {
  Logger.log("Sending updates to: %s", row);

  const email = getCellValueByColumnName(row, COL_EMAIL_ADDRESS, header);
  const jobPosition = getCellValueByColumnName(row, COL_JOB_POSITION, header);
  const status = getCellValueByColumnName(row, COL_FINAL_STATUS, header);
  const feedback = getCellValueByColumnName(row, COL_EXTERNAL_FEEDBACK, header);

  return sendEmail(email, jobPosition, status, feedback);
}

function updateStatusInSheet(sheet, row, header, statusCol, statusMsg) {
  const col = header.indexOf(statusCol);
  if (col != -1) {
    const range = sheet.getRange(row + 1, col + 1);
    range.setValue(statusMsg);
  }
}

function sendUpdates() {
  try {
    const sheet = SpreadsheetApp.getActiveSheet();
    const data = sheet.getDataRange().getValues();
    Logger.log("No of rows: %s", data.length);

    const header = data[0];
    for (let row = 1; row < data.length; ++row) {
      if (shouldSendStatus(data[row], header)) {
        const success = sendUpdate(data[row], header);
        if (success) {
          updateStatusInSheet(sheet, row, header, COL_FINAL_EMAIL_STATUS, "Sent");
        }
      }
    }
  } catch (err) {
    Logger.log("Error in sendUpdates: %s", err)
  }
}