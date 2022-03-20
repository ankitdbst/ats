const INTERVIEW_DURATION = 30; // in minutes
const INTERVIEW_INVITE_STATUS_COL = "Interview Invite Status"

function validate(candidateEmail, jobPosition, candidateSlot) {
  if (candidateEmail == "") {
    return { status: false, err: "email not provided" };
  }

  if (jobPosition == "") {
    return { status: false, err: "position not provided" };
  }

  if (candidateSlot == "") {
    return { status: false, err: "interview slot not provided" };
  };

  return { status: true, err: undefined }
}

function addMinutes(date, minutes) {
  return new Date(date.getTime() + minutes * 60000);
}

function parseInterviewSlot(candidateSlot) {
  const start = new Date(candidateSlot);
  const end = addMinutes(start, INTERVIEW_DURATION);

  Logger.log("Start Date: %s, End Date: %s", start, end);
  return { start: start.toISOString(), end: end.toISOString() };
}

function getHtmlDescription() {
  return "Dear Candidate,<br/><br />" +
    "We would like to have an interview over gmeet.<br/>" +
    "Please confirm your availability for this invite. Also, please reshare your resume on this thread.<br/><br />" +
    "Regards,<br />" +
    "Gramoday Team";
}

function prepareCalendarPayload(candidateEmail, jobPosition, candidateSlot) {
  const { start, end } = parseInterviewSlot(candidateSlot);

  const resource = {
    start: { dateTime: start },
    end: { dateTime: end },
    attendees: [{ email: candidateEmail }],
    conferenceData: {
      createRequest: {
        requestId: Utilities.getUuid(),
        conferenceSolutionKey: { type: "hangoutsMeet" },
      },
    },
    summary: `${jobPosition} | Gramoday Interview Invite`,
    description: getHtmlDescription(),
  };

  return resource;
}

function sendCalendarInvite(candidateEmail, jobPosition, candidateSlot) {
  Logger.log("Trying to send calendar invite to: %s for %s at %s", candidateEmail, jobPosition, candidateSlot);
  const { status, err } = validate(candidateEmail, jobPosition, candidateSlot);
  if (!status) {
    Logger.log("Error validating the invite data: %s, skipping", err);
    return false;
  }

  Logger.log("Sending invite...");

  const payload = prepareCalendarPayload(candidateEmail, jobPosition, candidateSlot);
  const calendarId = "primary";

  Logger.log("Calendar event payload: %s", payload);

  const res = Calendar.Events.insert(payload, calendarId, {
    conferenceDataVersion: 1,
    sendUpdates: "all",
  });
  console.log("Calendar invite status: %s", res.status);
  return true;
}

function getCellValueByColumnName(row, colName, header) {
  const col = header.indexOf(colName);
  if (col != -1) {
    return row[col];
  }
}

function shouldSendInvite(row, header) {
  const assignmentStatus = getCellValueByColumnName(row, "Assignment Status", header);
  if (assignmentStatus != "Passed") {
    return false;
  }

  const inviteStatus = getCellValueByColumnName(row, "Interview Invite Status", header);
  if (inviteStatus == "Sent") {
    return false;
  }
  const interviewSlot = getCellValueByColumnName(row, "Interview Slot", header);
  if (interviewSlot == "") {
    return false;
  }
  return true;
}

function sendInterviewInvite(row, header) {
  Logger.log("Sending invite to: %s", row);

  const email = getCellValueByColumnName(row, "Email Address", header);
  const jobPosition = getCellValueByColumnName(row, "Which position are you applying for?", header);
  const interviewSlot = getCellValueByColumnName(row, "Interview Slot", header);

  return sendCalendarInvite(email, jobPosition, interviewSlot);
}

function updateStatusInSheet(sheet, row, header, statusCol, statusMsg) {
  const col = header.indexOf(statusCol);
  if (col != -1) {
    const range = sheet.getRange(row + 1, col + 1);
    range.setValue(statusMsg);
  }
}

function sendInterviewInvites() {
  try {
    const sheet = SpreadsheetApp.getActiveSheet();
    const data = sheet.getDataRange().getValues();
    Logger.log("No of rows: %s", data.length);

    const header = data[0];
    for (let row = 1; row < data.length; ++row) {
      if (shouldSendInvite(data[row], header)) {
        const success = sendInterviewInvite(data[row], header);
        if (success) {
          updateStatusInSheet(sheet, row, header, INTERVIEW_INVITE_STATUS_COL, "Sent");
        }
      }
    }
  } catch (err) {
    Logger.log("Error in sendInterviewInvites: %s", err)
  }
}