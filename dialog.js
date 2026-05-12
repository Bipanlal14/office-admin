Office.onReady(() => {

  loadRecipients();

});

const INTERNAL_DOMAIN = "signatureclinic.co.uk";

const EXCLUDED_DOMAINS = [
  "trustedpartner.com",
  "nhs.uk"
];

const EXCLUDED_USERS = [
  "ceo@gmail.com"
];

async function loadRecipients() {

  const item = Office.context.mailbox.item;

  item.to.getAsync((result) => {

    if (result.status !== Office.AsyncResultStatus.Succeeded) {

      console.error("Failed to get recipients");

      return;
    }

    const recipients = result.value;

    const externalRecipients = recipients.filter(recipient => {

      const email = recipient.emailAddress.toLowerCase();

      const domain = email.split("@")[1];

      // Ignore internal users
      if (domain === INTERNAL_DOMAIN) {
        return false;
      }

      // Ignore excluded domains
      if (EXCLUDED_DOMAINS.includes(domain)) {
        return false;
      }

      // Ignore excluded users
      if (EXCLUDED_USERS.includes(email)) {
        return false;
      }

      return true;

    });

    renderRecipients(externalRecipients);

  });

}

function renderRecipients(recipients) {

  const emailList = document.getElementById("emailList");

  emailList.innerHTML = "";

  if (recipients.length === 0) {

    emailList.innerHTML = `
      <div class="safe-message">
        ✅ No external recipients detected
      </div>
    `;

    return;
  }

  recipients.forEach(recipient => {

    const div = document.createElement("div");

    div.className = "email-item";

    div.innerHTML = `
      ⚠ ${recipient.displayName} (${recipient.emailAddress})
    `;

    emailList.appendChild(div);

  });

}

function cancelSend() {

  Office.context.ui.messageParent("cancel");

}

function sendAnyway() {

  Office.context.ui.messageParent("send");

}