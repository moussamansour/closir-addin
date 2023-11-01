/* global document, Office */

Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) {
    document.getElementById("success-message").style.display = "none";
    if (Office.context.roamingSettings.get("closir_token39")) {
      document.getElementById("token").value = Office.context.roamingSettings.get("closir_token39");
      document.getElementById("saveButton").innerHTML = "Update";
    }
    document.getElementById("saveButton").onclick = saveToken;
  }
});
// Function to save the token
function saveToken() {
  var token = document.getElementById("token").value;
  // Save the token in add-on settings
  Office.context.mailbox.item.organizer.getAsync(function callback(result) {
    try {
      let email = result.value.emailAddress;
      const response = fetch("https://www.closir.com/api/public/token?requesterEmail=" + email, {
        method: "GET",
        headers: {
          "Content-Type": "application/json",
        },
      });
      response.then((data) => {
        data.json().then((result) => {
          if (result) {
            if (result["token"] == token) {
              document.getElementById("failed-message").style.display = "none";
              Office.context.roamingSettings.set("closir_token39", token);
              Office.context.roamingSettings.saveAsync(function (result) {
                if (result.status == Office.AsyncResultStatus.Succeeded) {
                  // Token saved successfully, provide feedback to the user
                  document.getElementById("saveButton").innerHTML = "Update";
                  document.getElementById("success-message").style.display = "block";
                  setTimeout(() => {
                    document.getElementById("success-message").style.display = "none";
                  }, 5000);
                  console.log("Token saved successfully!");
                } else {
                  // Handle error
                  console.log("Error saving token: " + result.error.message);
                }
              });
            } else {
              document.getElementById("failed-message").style.display = "block";
            }
          } else {
            document.getElementById("failed-message").style.display = "block";
          }
        });
      });
    } catch (e) {
      console.log(e);
      document.getElementById("failed-message").style.display = "block";
    }
  });
}
