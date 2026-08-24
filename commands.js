/*
 * PDF Attachment Alert — Outlook Smart Alerts add-in
 *
 * Fires automatically whenever a user clicks Send. If any attachment has a
 * .pdf extension, it blocks the send and shows a warning with a "Send Anyway"
 * option, matching the same reminder the DLP Policy Tip shows in OWA:
 * "Please verify the customer on the attachment matches the customer you
 * are sending this to."
 */

Office.onReady(() => {
  // Nothing to initialize — event-based activation handles everything.
});

function checkForPdfAttachment(event) {
  const item = Office.context.mailbox.item;

  item.getAttachmentsAsync((result) => {
    if (result.status !== Office.AsyncResultStatus.Succeeded) {
      // If we can't read attachments for some reason, don't block sending.
      event.completed({ allowEvent: true });
      return;
    }

    const attachments = result.value || [];
    const hasPdf = attachments.some((a) => {
      const name = (a.name || "").toLowerCase();
      return name.endsWith(".pdf");
    });

    if (hasPdf) {
      event.completed({
        allowEvent: false,
        errorMessage:
          "Please verify the customer on the attachment matches the customer you are sending this to.",
        errorMessageMarkdown:
          "**Please verify the customer on the attachment matches the customer you are sending this to.**",
        cancelLabel: "Don't Send",
        sendModeOverride: Office.MailboxEnums.SendModeOverride.PromptUser,
      });
    } else {
      event.completed({ allowEvent: true });
    }
  });
}

// Register the function so the runtime can find it when OnMessageSend fires.
Office.actions.associate("checkForPdfAttachment", checkForPdfAttachment);
