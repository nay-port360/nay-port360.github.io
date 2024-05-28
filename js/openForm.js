// openForm.js

function openForm(event) {
  // Get the current message
  const item = Office.context.mailbox.item;

  // Get email details
  const subject = item.subject;
  const receivedDate = item.dateTimeReceived;

  // Get the email content (as a file)
  item.getAttachmentContentAsync((result) => {
    if (result.status === Office.AsyncResultStatus.Succeeded) {
      const attachment = result.value;

      // Prepare form data
      const formData = new FormData();
      formData.append('file', new Blob([attachment.content], { type: attachment.format }), `${subject}.eml`);
      formData.append('receivedDate', receivedDate);

      // Post form data to your website
      fetch('https://p360.test/operations/enquiries/create', {
        method: 'POST',
        body: formData,
      })
      .then(response => response.json())
      .then(data => {
        // Redirect to your form page
        window.open(`https://p360.test/operations/enquiries/create?receivedDate=${encodeURIComponent(receivedDate)}`, '_blank');
      })
      .catch(error => console.error('Error uploading email:', error));
    } else {
      console.error('Error getting attachment content:', result.error);
    }
  });
}

// Export the function
Office.initialize = function (reason) {
  Office.actions.associate('openForm', openForm);
};
