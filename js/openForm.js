Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) {
    // Initialization code goes here
  }
});

function openForm(event) {
  console.log('testing');

  // Complete the event
  event.completed();
}