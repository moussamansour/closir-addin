!function(){function e(){var e=document.getElementById("token").value;Office.context.mailbox.item.organizer.getAsync((function(t){try{var n=t.value.emailAddress;fetch("https://www.closir.com/api/public/token?requesterEmail="+n,{method:"GET",headers:{"Content-Type":"application/json"}}).then((function(t){t.json().then((function(t){t&&t.token==e?(document.getElementById("failed-message").style.display="none",Office.context.roamingSettings.set("closir_token39",e),Office.context.roamingSettings.saveAsync((function(e){e.status==Office.AsyncResultStatus.Succeeded?(document.getElementById("saveButton").innerHTML="Update",document.getElementById("success-message").style.display="block",setTimeout((function(){document.getElementById("success-message").style.display="none"}),5e3),console.log("Token saved successfully!")):console.log("Error saving token: "+e.error.message)}))):document.getElementById("failed-message").style.display="block"}))}))}catch(e){console.log(e),document.getElementById("failed-message").style.display="block"}}))}Office.onReady((function(t){t.host===Office.HostType.Outlook&&(document.getElementById("success-message").style.display="none",Office.context.roamingSettings.get("closir_token39")&&(document.getElementById("token").value=Office.context.roamingSettings.get("closir_token39"),document.getElementById("saveButton").innerHTML="Update"),document.getElementById("saveButton").onclick=e)}))}();
//# sourceMappingURL=settings.js.map