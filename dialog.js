Office.onReady((info) => {

  if (info.host === Office.HostType.Outlook) {

    refreshRecipients();

    // Auto refresh every 5 seconds
    setInterval(refreshRecipients, 5000);
  }
});

const CONFIG = {

  internalDomains: [
    "signatureclinic.co.uk"
  ],

  excludedDomains: [
    "trustedpartner.com",
    "nhs.uk"
  ],

  excludedUsers: [
    "ceo@gmail.com"
  ]
};

function getDomain(email) {

  if (!email || !email.includes("@")) {
    return "";
  }

  return email
    .split("@")[1]
    .toLowerCase()
    .trim();
}

function matchesAnyDomain(domain, domainList) {

  return domainList.some(d =>
    domain === d || domain.endsWith(`.${d}`)
  );
}

function isExternal(email) {

  if (!email || !email.includes("@")) {
    return false;
  }

  email = email.toLowerCase().trim();

  const domain = getDomain(email);

  if (!domain) {
    return false;
  }

  // Internal domains
  if (matchesAnyDomain(domain, CONFIG.internalDomains)) {
    return false;
  }

  // Trusted exclusions
  if (matchesAnyDomain(domain, CONFIG.excludedDomains)) {
    return false;
  }

  // Specific users
  if (CONFIG.excludedUsers.includes(email)) {
    return false;
  }

  return true;
}

async function getRecipients(field) {

  return new Promise((resolve) => {

    field.getAsync((result) => {

      if (
        result.status === Office.AsyncResultStatus.Succeeded
      ) {
        resolve(result.value || []);
      }
      else {
        resolve([]);
      }
    });
  });
}

async function refreshRecipients() {

  try {

    if (!Office.context || !Office.context.mailbox) {
      return;
    }

    const item = Office.context.mailbox.item;

    if (!item) {
      return;
    }

    const toRecipients = (await getRecipients(item.to)).map(r => ({
      ...r,
      type: "TO"
    }));

    const ccRecipients = (await getRecipients(item.cc)).map(r => ({
      ...r,
      type: "CC"
    }));

    const bccRecipients = (await getRecipients(item.bcc)).map(r => ({
      ...r,
      type: "BCC"
    }));

    const allRecipients = [
      ...toRecipients,
      ...ccRecipients,
      ...bccRecipients
    ];

    // Remove duplicates
    const uniqueRecipients = [];
    const seen = new Set();

    allRecipients.forEach(r => {

      const email = (r.emailAddress || "")
        .toLowerCase()
        .trim();

      if (!email || seen.has(email)) {
        return;
      }

      seen.add(email);

      uniqueRecipients.push(r);
    });

    const externalRecipients = uniqueRecipients.filter(r =>
      isExternal(r.emailAddress)
    );

    renderRecipients(externalRecipients);
  }
  catch (err) {

    console.error("Recipient refresh failed", err);
  }
}

function renderRecipients(recipients) {

  const emailList = document.getElementById("emailList");
  const summaryBox = document.getElementById("summaryBox");

  emailList.innerHTML = "";
  summaryBox.innerHTML = "";

  if (recipients.length === 0) {

    summaryBox.innerHTML = `
      <div class="safe-message">
        ✅ No external recipients detected
      </div>
    `;

    return;
  }

  const toCount = recipients.filter(r => r.type === "TO").length;
  const ccCount = recipients.filter(r => r.type === "CC").length;
  const bccCount = recipients.filter(r => r.type === "BCC").length;

  summaryBox.innerHTML = `
    <div class="summary-warning">
      ${recipients.length} external recipient(s) detected
      <br>
      TO: ${toCount} | CC: ${ccCount} | BCC: ${bccCount}
    </div>
  `;

  recipients.forEach(recipient => {

    const div = document.createElement("div");

    div.className = "email-item";

    const typeBadge = document.createElement("span");
    typeBadge.className = "recipient-type";
    typeBadge.textContent = recipient.type;

    const text = document.createElement("span");

    text.textContent =
      ` ${recipient.displayName || "Unknown"} (${recipient.emailAddress})`;

    div.appendChild(typeBadge);
    div.appendChild(text);

    emailList.appendChild(div);
  });
}

function closePane() {

  Office.context.ui.closeContainer();
}