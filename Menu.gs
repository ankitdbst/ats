function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('ATS')
    .addItem('Send Interview Invite(s)', 'sendInterviewInvites')
    .addItem('Send Update(s)', 'sendUpdates')
    .addToUi();
}
