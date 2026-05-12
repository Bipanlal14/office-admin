const externalRecipients = [
  "external.user@gmail.com",
  "client@yahoo.com",
  "partner@anothercompany.com"
];

window.onload = function () {

  const emailList = document.getElementById("emailList");

  externalRecipients.forEach(email => {

    const div = document.createElement("div");

    div.className = "email-item";

    div.textContent = email;

    emailList.appendChild(div);
  });
};

function cancelSend() {

  alert("Email sending cancelled.");
}

function sendAnyway() {

  alert("Email sent successfully.");
}