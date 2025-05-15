Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) {
    const mailbox = Office.context.mailbox;

    const displayName = mailbox.userDisplayName;
    const email = mailbox.userEmailAddress;

    document.getElementById("username").innerText = `Nom : ${displayName}`;
    document.getElementById("email").innerText = `Email : ${email}`;

    document.getElementById("leaveForm").addEventListener("submit", async (e) => {
      e.preventDefault();

      const type = document.getElementById("leaveType").value;
      const start = document.getElementById("startDate").value;
      const end = document.getElementById("endDate").value;

      const event = {
        subject: `[Congé - ${type}] ${displayName}`,
        body: {
          contentType: "HTML",
          content: `${displayName} sera absent du ${start} au ${end} (${type}).`
        },
        start: {
          dateTime: `${start}T09:00:00`,
          timeZone: "Europe/Paris"
        },
        end: {
          dateTime: `${end}T18:00:00`,
          timeZone: "Europe/Paris"
        },
        showAs: "OOF",
        location: {
          displayName: "Absence"
        },
        attendees: []
      };

      try {
        await Office.auth.getAccessTokenAsync({ allowSignInPrompt: true }, async (result) => {
          if (result.status === "succeeded") {
            const token = result.value;

            const response = await fetch("https://graph.microsoft.com/v1.0/users/lan_planning_conges@lanpark.eu/events", {
              method: "POST",
              headers: {
                "Authorization": `Bearer ${token}`,
                "Content-Type": "application/json"
              },
              body: JSON.stringify(event)
            });

            if (response.ok) {
              document.getElementById("status").innerText = "Congé ajouté avec succès !";
            } else {
              document.getElementById("status").innerText = "Erreur : " + (await response.text());
            }
          } else {
            document.getElementById("status").innerText = "Erreur d'authentification.";
          }
        });
      } catch (error) {
        document.getElementById("status").innerText = "Erreur : " + error.message;
      }
    });
  }
});
