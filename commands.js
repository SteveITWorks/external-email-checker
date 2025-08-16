Office.initialize = () => {};
Office.actions.associate("onMessageSendHandler", onMessageSendHandler);

async function onMessageSendHandler(event) {
  try {
    const item = Office.context.mailbox.item;

    const recipients = [
      ...(item.to || []),
      ...(item.cc || []),
      ...(item.bcc || []),
    ].map(r => r.emailAddress);

    // ✅ Change this to your real internal domain
    const internalDomain = "itworks.co.nz";

    const externalRecipients = recipients.filter(email =>
      !email.toLowerCase().endsWith("@" + internalDomain)
    );

    if (externalRecipients.length > 0) {
      Office.context.ui.displayDialogAsync(
        "https://steveitworks.github.io/outlook-external-checker/confirmation.html",
        { height: 40, width: 30, displayInIframe: true },
        result => {
          const dialog = result.value;
          dialog.addEventHandler(Office.EventType.DialogMessageReceived, msg => {
            if (msg.message === "confirm") {
              dialog.close();
              event.completed({ allowEvent: true });
            } else {
              dialog.close();
              event.completed({ allowEvent: false });
            }
          });
        }
      );
    } else {
      event.completed({ allowEvent: true });
    }
  } catch (err) {
    console.error(err);
    event.completed({ allowEvent: true }); // fallback
  }
}
