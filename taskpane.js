(function () {
  Office.onReady(() => {
    document.getElementById("leaveForm").addEventListener("submit", async function (e) {
      e.preventDefault();
      const start = document.getElementById("startDate").value;
      const end = document.getElementById("endDate").value;
      const type = document.getElementById("leaveType").value;
      const subject = `Congé (${type})`;
      const body = `Congé de type ${type} du ${start} au ${end}`;

      Office.context.mailbox.displayNewAppointmentForm({
        subject: subject,
        start: new Date(start),
        end: new Date(end),
        body: body,
      });
    });
  });
})();
