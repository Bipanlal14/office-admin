Office.onReady(() => {
  console.log("OnSend handler loaded");
});

const INTERNAL_DOMAIN = "signatureclinic.co.uk";

const EXCLUDED_DOMAINS = [
  "trustedpartner.com",
  "nhs.uk"
];

const EXCLUDED_USERS = [
  "ceo@gmail.com"
];

function getDomain(email) {

  if (!email || !email.includes("@")) {
    return "";
  }

  return email.split("@")[1].toLowerCase();
}

function isExternal(email) {

  email = (email || "").toLowerCase().trim();

  if (!email) {
    return false;
  }

  const domain = getDomain(email);

  if (!domain) {
    return false;
  }

  if (domain === INTERNAL_DOMAIN) {
    return false;
  }

  if (EXCLUDED_DOMAINS.includes(domain)) {
    return false;
  }

  if (EXCLUDED_USERS.includes(email)) {
    return false;
  }

  return true;
}

function getRecipientsAsync(field) {

  return new Promise((resolve) => {

    field.getAsync((result) => {

      if (result.status === Office.AsyncResultStatus.Succeeded) {
        resolve(result.value || []);
      } else {
        resolve([]);
      }

    });

  });

}

async function onMessageSendHandler(event) {

  try {

    const item = Office.context.mailbox.item;

    const toRecipients = await getRecipientsAsync(item.to);
    const ccRecipients = await getRecipientsAsync(item.cc);
    const bccRecipients = await getRecipientsAsync(item.bcc);

    const allRecipients = [
      ...toRecipients,
      ...ccRecipients,
      ...bccRecipients
    ];

    const externalRecipients = allRecipients.filter(r =>
      isExternal(r.emailAddress)
    );

    if (externalRecipients.length > 0) {

      const recipientList = externalRecipients
        .map(r => r.emailAddress)
        .join(", ");

      Office.context.mailbox.item.notificationMessages.replaceAsync(
        "externalWarning",
        {
          type: Office.MailboxEnums.ItemNotificationMessageType.WarningMessage,
          message: `External recipients detected: ${recipientList}`,
          icon: "Icon.80x80",
          persistent: true
        }
      );

      event.completed({
        allowEvent: false,
        errorMessage:
          "External recipients detected. Please review recipients before sending."
      });

      return;
    }

    event.completed({
      allowEvent: true
    });

  } catch (error) {

    console.error("OnSend failure", error);

    event.completed({
      allowEvent: true
    });
  }
}

Office.actions.associate(
  "onMessageSendHandler",
  onMessageSendHandler
);
