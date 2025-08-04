Office.initialize = () => {
  Office.actions.associate("onMessageSendHandler", onMessageSendHandler);
};

async function onMessageSendHandler(event) {
  try {
    const item = Office.context.mailbox.item;

    const [toRecipients, ccRecipients, bccRecipients] = await Promise.all([
      new Promise((resolve) => item.to.getAsync(resolve)),
      new Promise((resolve) => item.cc.getAsync(resolve)),
      new Promise((resolve) => item.bcc.getAsync(resolve)),
    ]);

    const allRecipients = [
      ...(toRecipients.value || []),
      ...(ccRecipients.value || []),
      ...(bccRecipients.value || []),
    ];

    const internalDomain = "itworks.co.nz";
    const externalRecipients = allRecipients.filter((recipient) => {
      const email = recipient.emailAddress?.toLowerCase() || "";
      return !email.endsWith("@" + internalDomain);
    });

    if (externalRecipients.length > 0) {
      Office.context.ui.displayDialogAsync(
        "https://steveitworks.github.io/external-email-checker/confirmation.html",
        { height: 40, width: 40, displayInIframe: true },
        (asyncResult) => {
          const dialog = asyncResult.value;
          dialog.addEventHandler(Office.EventType.DialogMessageReceived, (message) => {
            if (message.message === "cancel") {
              dialog.close();
              event.completed({ allowEvent: false });
            } else {
              dialog.close();
              event.completed({ allowEvent: true });
            }
          });
        }
      );
    } else {
      event.completed({ allowEvent: true });
    }
  } catch (e) {
    console.error("Error in send handler:", e);
    event.completed({ allowEvent: true });
  }
}
