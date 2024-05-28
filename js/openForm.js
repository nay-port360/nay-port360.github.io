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

  // Redirect to the form URL
  window.open(formUrl, "_blank");

  // Complete the event
  event.completed();
}
