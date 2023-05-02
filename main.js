Office.initialize = function () {
  // Add a button to the ribbon for reporting phishing emails
  var phishingButton = Office.context.mailbox.item.addCustomButton({
    id: "reportPhishingButton",
    iconUrl: "https://officedev.github.io/Office-Add-in-samples/Samples/hello-world/outlook-hello-world/assets/icon-80.png",
    onClick: reportPhishing
  });
};

// Function to handle the click event for the phishing report button
function reportPhishing(event) {
  // Get the email subject and sender
  var emailSubject = Office.context.mailbox.item.subject;
  var emailSender = Office.context.mailbox.item.from.displayName;

  // Generate a unique report ID
  var reportID = Math.floor(Math.random() * 1000000);

  // Send the phishing report to your security team or email provider
  var reportBody = "Subject: " + emailSubject + "\n" +
                   "Sender: " + emailSender + "\n" +
                   "Report ID: " + reportID;
  Office.context.mailbox.makeEwsRequestAsync(
    "<CreateItem xmlns='http://schemas.microsoft.com/exchange/services/2006/messages'>" +
      "<Items>" +
        "<Message>" +
          "<Subject>Phishing Report - " + emailSubject + "</Subject>" +
          "<Body>" +
            "<BodyType>Text</BodyType>" +
            "<Text>" + reportBody + "</Text>" +
          "</Body>" +
          "<ToRecipients>" +
            "<Mailbox>" +
              "<EmailAddress>johann.boenewitz@supratix.com</EmailAddress>" +
            "</Mailbox>" +
          "</ToRecipients>" +
        "</Message>" +
      "</Items>" +
    "</CreateItem>",
    function (asyncResult) {
      if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
        // Display a success message to the user
        Office.context.mailbox.item.notificationMessages.addAsync(
          "phishingReportSuccess",
          {
            type: "informationalMessage",
            message: "Phishing email reported with ID " + reportID
          },
          function (result) {}
        );
      } else {
        // Display an error message to the user
        Office.context.mailbox.item.notificationMessages.addAsync(
          "phishingReportError",
          {
            type: "errorMessage",
            message: "Failed to report phishing email"
          },
          function (result) {}
        );
      }
    }
  );
}
