const originalTextarea = document.getElementById("original");
const ignoreTextarea = document.getElementById("ignore");
const ignoringEmailsCount = document.getElementById("ignoring-emails-count");

function onIgnoreValueChange(text) {
  ignoringEmailsCount.innerHTML = text.split("\n").length;
  localStorage.setItem("ignoreAddresses", text);
}

function setIgnoreValue(text) {
  ignoreTextarea.value = text;
  onIgnoreValueChange(text);
}

setIgnoreValue(localStorage.getItem("ignoreAddresses"));

ignoreTextarea.addEventListener("input", (event) => {
  onIgnoreValueChange(event.target.value);
});

function addressesFromText(text) {
  if (text.trim() === "") return new Set();

  return new Set(
    text
      .trim()
      .split("\n")
      .map((line) => line.trim())
      .filter((line) => line)
  );
}

const buttonAddresses = document.getElementById("getAddressesButton");
buttonAddresses.addEventListener("click", () => {
  const original = addressesFromText(originalTextarea.value);
  const ignore = addressesFromText(ignoreTextarea.value);

  const finalAddresses = Array.from(original).filter(
    (address) => !ignore.has(address)
  );

  const batchSize = Number(document.getElementById("batchSize").value);
  const groups = finalAddresses.reduce(() => {
    const batches = [];
    for (let i = 0; i < finalAddresses.length; i += batchSize) {
      batches.push(finalAddresses.slice(i, i + batchSize));
    }
    return batches;
  }, []);

  navigator.clipboard.writeText(
    groups.map((batch) => batch.join(";")).join("\n")
  );
});

function getEmailsWithResponse(report) {
  return report.attendees
    .filter((attendee) => {
      return ["Accepted", "Declined", "Tentative"].includes(
        attendee.status.response
      );
    })
    .map((attendee) => attendee.address);
}

const fromTeamReport = document.getElementById("fromTeamReport");
fromTeamReport.addEventListener("click", () => {
  navigator.clipboard.readText().then((text) => {
    try {
      const emailsToIgnore = getEmailsWithResponse(JSON.parse(text));

      setIgnoreValue(emailsToIgnore.join("\n"));
    } catch (error) {
      alert(
        "Could not parse the report, please make sure you copied the correct content"
      );
      console.error("Error parsing the Teams report:", error);
    }
  });
});
