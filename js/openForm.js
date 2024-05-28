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
  const formUrl = `https://www.port360.com/createForm?receivedDateTime=${encodeURIComponent(receivedDateTime)}`;

  // Get the current email as an EML file
  item.getAttachmentContentAsync(item.attachments[0].id, (result) => {
    if (result.status === Office.AsyncResultStatus.Succeeded) {
      const content = result.value.content; // Base64 encoded EML file
      const mimeType = result.value.format; // Should be "eml"

      // Create a Blob from the Base64 content
      const byteCharacters = atob(content);
      const byteNumbers = new Array(byteCharacters.length);
      for (let i = 0; i < byteCharacters.length; i++) {
        byteNumbers[i] = byteCharacters.charCodeAt(i);
      }
      const byteArray = new Uint8Array(byteNumbers);
      const blob = new Blob([byteArray], { type: mimeType });

      // Upload the Blob to your server (example using fetch)
      const formData = new FormData();
      formData.append('file', blob, 'email.eml');

      fetch(formUrl, {
        method: 'POST',
        body: formData
      }).then(response => response.json())
        .then(data => {
          // Handle success
          console.log('Form submitted successfully:', data);
        }).catch(error => {
          // Handle error
          console.error('Error submitting form:', error);
        });
    } else {
      console.error('Failed to get attachment content:', result.error);
    }
  });

  event.completed();
}
