document.getElementById("leaveForm").addEventListener("submit", async function(e) {
  e.preventDefault();
  const startDate = document.getElementById("startDate").value;
  const endDate = document.getElementById("endDate").value;
  const isHalf = document.querySelector("input[name='dayType']:checked").value === "half";
  const halfType = isHalf ? document.querySelector("input[name='halfType']:checked").value : null;
  const absenceType = document.getElementById("absenceType").value;

  const user = Office.context.mailbox.userProfile.displayName;

  const event = {
    subject: `Absence - ${user} (${absenceType})`,
    body: {
      contentType: "HTML",
      content: `${user} est absent pour ${absenceType}.`
    },
    start: {
      dateTime: `${startDate}T08:00:00`,
      timeZone: "Europe/Paris"
    },
    end: {
      dateTime: `${endDate}T17:00:00`,
      timeZone: "Europe/Paris"
    },
    showAs: "away",
    isAllDay: !isHalf
  };

  // À ce stade, il faut authentifier avec MSAL et envoyer à l’API Microsoft Graph
  console.log("Événement prêt à être envoyé :", event);
});
