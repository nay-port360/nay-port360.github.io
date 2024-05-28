Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) {
    // Initialization code goes here
  }
});

function openForm(event) {
  const item = Office.context.mailbox.item;

  // Get the email received date and time
  const receivedDateTime = item.dateTimeCreated;

  // Create a form URL with the received date and time as a query parameter
  const formUrl = `https://p360.test/operations/enquiries/create?receivedDateTime=${encodeURIComponent(receivedDateTime)}`;

  // Open the form URL in a dialog
  Office.context.ui.displayDialogAsync(formUrl, { height: 50, width: 50 }, function (result) {
    if (result.status === Office.AsyncResultStatus.Failed) {
      console.error('Failed to open the dialog:', result.error.message);
    }
  });

  // Complete the event
  event.completed();
}
