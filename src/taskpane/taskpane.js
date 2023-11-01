/* global document, Office */
import * as CryptoJS from "crypto-js";
var AWS = require("aws-sdk");
const config = require("../credentials/credentials");

const AWS_ACCESS_KEY_ID = config.AWS_ACCESS_KEY_ID;
const AWS_SECRET_ACCESS_KEY = config.AWS_SECRET_ACCESS_KEY;
const AWS_BUCKET = config.AWS_BUCKET;

AWS.config.update({
  accessKeyId: AWS_ACCESS_KEY_ID,
  secretAccessKey: AWS_SECRET_ACCESS_KEY,
});

let attachmentsIDs = [];
Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) {
    let itemId = "";
    if (Office.context.mailbox.item.itemId) {
      itemId = Office.context.mailbox.item.itemId;
      getSavedData(itemId);
    } else {
      Office.context.mailbox.item.saveAsync(function (result) {
        itemId = result.value;
        getSavedData(itemId);
      });
    }

    if (!Office.context.roamingSettings.get("closir_token39")) {
      document.getElementById("access_gained").style.display = "none";
      document.getElementById("item-subject").innerHTML =
        "Please add your Closir token in settings to be able to save your meetings onto Closir";
    } else {
      getTags();
      document.getElementById("approve_adding_meeting").onclick = getMeetingData;
      document.getElementById("host").onclick = showHost;
      document.getElementById("direct").onclick = hideHost;
    }
  }
});

function getSavedData(itemId) {
  if (Office.context.roamingSettings.get("meeting" + itemId)) {
    document.getElementById("approve_adding_meeting_text").innerText = "Update on Closir";
    document.getElementById("approve_adding_meeting").style.display = "none";
    document.getElementById("item-subject").style.display = "none";
    document.getElementById("meeting_saved").style.display = "block";
    let inputs = document.getElementsByTagName("INPUT");
    for (let j = 0; j < inputs.length; j++) {
      inputs[j].disabled = true;
    }
    document.getElementById("meeting_types").disabled = true;
  }
  if (Office.context.roamingSettings.get("meeting_type" + itemId)) {
    document.getElementById("meeting_types").value = Office.context.roamingSettings.get("meeting_type" + itemId);
  }
  if (Office.context.roamingSettings.get("host" + itemId)) {
    if (Office.context.roamingSettings.get("host" + itemId) == "host") {
      document.getElementById("host").checked = true;
      document.getElementById("host_name").style.display = "block";
      if (Office.context.roamingSettings.get("host_name" + itemId)) {
        document.getElementById("host_name").value = Office.context.roamingSettings.get("host_name" + itemId);
      }
    } else {
      document.getElementById("direct").checked = true;
    }
  } else {
    document.getElementById("direct").checked = true;
  }
}
async function showHost(value) {
  document.getElementById("host_name").style.display = "block";
}

async function hideHost(value) {
  document.getElementById("host_name").style.display = "none";
}

async function getMeetingData() {
  document.getElementById("loader").style.display = "block";
  let meeting_object = {};
  let mailboxItem = Office.context.mailbox.item;
  if (Office.MailboxEnums.ItemType.Appointment) {
    meeting_object["required_participants"] = [];
    meeting_object["optional_participants"] = [];
    if (mailboxItem.itemId) {
      meeting_object["meeting_name"] = mailboxItem.subject ? mailboxItem.subject : "";
      mailboxItem.body.getAsync("text", function (text) {
        meeting_object["meeting_notes"] = text.value ? text.value : "";
      });
      meeting_object["meeting_location"] = mailboxItem.location ? mailboxItem.location : "";
      meeting_object["address"] = mailboxItem.location ? mailboxItem.location : "";
      meeting_object["required_participants"] = mailboxItem.requiredAttendees ? mailboxItem.requiredAttendees : [];
      meeting_object["optional_participants"] = mailboxItem.optionalAttendees ? mailboxItem.optionalAttendees : [];
      meeting_object["end_of_meeting"] = mailboxItem.end ? formatDate(mailboxItem.end) : "";
      meeting_object["date_of_meeting"] = mailboxItem.start ? formatDate(mailboxItem.start) : "";
      meeting_object["time_zone"] = mailboxItem.start ? getTimeZone(mailboxItem.start) : "";
      fillMeetingsData(meeting_object);
    } else {
      const promises = [
        promiseAsyncFunction((callback) => mailboxItem.body.getAsync("text", callback)).then((note) => {
          meeting_object["meeting_notes"] = note;
        }),
        promiseAsyncFunction((callback) => mailboxItem.location.getAsync(callback)).then((location) => {
          meeting_object["meeting_location"] = location;
          meeting_object["address"] = location;
        }),
        promiseAsyncFunction((callback) => mailboxItem.subject.getAsync(callback)).then((subject) => {
          meeting_object["meeting_name"] = subject;
        }),
        promiseAsyncFunction((callback) => mailboxItem.requiredAttendees.getAsync(callback)).then(
          (requiredAttendees) => {
            meeting_object["required_participants"] = requiredAttendees;
          }
        ),
        promiseAsyncFunction((callback) => mailboxItem.optionalAttendees.getAsync(callback)).then(
          (optionalAttendees) => {
            meeting_object["optional_participants"] = optionalAttendees;
          }
        ),
        promiseAsyncFunction((callback) => mailboxItem.end.getAsync(callback)).then((end) => {
          meeting_object["end_of_meeting"] = formatDate(end);
        }),
        promiseAsyncFunction((callback) => mailboxItem.start.getAsync(callback)).then((start) => {
          meeting_object["date_of_meeting"] = formatDate(start);
          meeting_object["time_zone"] = getTimeZone(start);
        }),
      ];

      Promise.all(promises).then(() => {
        fillMeetingsData(meeting_object);
      });
    }
  } else {
    console.log("This feature is only available in an appointment context.");
  }
}

function fillMeetingsData(meeting_object) {
  let participants = [];
  meeting_object["slot_type"] = "meeting";
  let topics = document.getElementsByClassName("tagInput");
  let topics_weight = [];
  for (let i = 0; i < topics.length; i++) {
    if (topics[i].innerHTML != "0" && topics[i].style.display != "none" ) {
      topics_weight.push({ name: topics[i].id.replace("tag_", ""), weight: topics[i].innerHTML == "+" ? 1 : Number(topics[i].innerHTML), isSelected: true });
    }
  }
  if (topics_weight.length > 0) meeting_object["tags"] = JSON.parse(JSON.stringify(topics_weight));
  console.log(meeting_object);
  meeting_object["host"] = document.getElementById("host").checked ? "host" : "direct";
  if (document.getElementById("host").checked) meeting_object["host_name"] = document.getElementById("host_name").value;
  for (let i = 0; i < meeting_object["required_participants"].length; i++) {
    const attendee = meeting_object["required_participants"][i]["emailAddress"];
    participants.push(attendee);
  }
  for (let i = 0; i < meeting_object["optional_participants"].length; i++) {
    const attendee = meeting_object["optional_participants"][i]["emailAddress"];
    participants.push(attendee);
  }
  meeting_object["meeting_format"] = participants.length > 1 ? "Group" : "1-1";
  delete meeting_object["required_participants"];
  delete meeting_object["optional_participants"];
  let itemId = Office.context.mailbox.item.itemId;
  if (itemId == null || itemId == undefined) {
    Office.context.mailbox.item.saveAsync(function (result) {
      itemId = result.value;
      if (Office.context.roamingSettings.get("id" + itemId)) {
        //meeting_object["id"] = Office.context.roamingSettings.get("id" + itemId);
      }
    });
  }
  meeting_object["investor_recortds"] = participants.toString();
  getAttachments(meeting_object);
}

function promiseAsyncFunction(asyncFunction) {
  return new Promise((resolve, reject) => {
    asyncFunction(function callback(result) {
      if (result.status === "succeeded") {
        resolve(result.value);
      } else {
        reject(new Error("Async function failed"));
      }
    });
  });
}

function getAttachments(meeting_object) {
  const item = Office.context.mailbox.item;
  attachmentsIDs = [];
  if (item.attachments) {
    console.log(item.attachments);
    let counterFiles = item.attachments.length;
    if (counterFiles > 0) {
      for (let k = 0; k < item.attachments.length; k++) {
        let attachment = item.attachments[k];
        if (attachment.name == "open.url") {
          console.log("url attachment");
          counterFiles--;
        } else {
          try {
            item.getAttachmentContentAsync(attachment.id, (contentResult) => {
              console.log(attachment);
              console.log(contentResult);
              uploadFile(meeting_object, attachment, contentResult, counterFiles);
            });
          } catch (e) {
            counterFiles--;
            console.log(e);
          }
        }
      }
    } else {
      if (meeting_object["meeting_name"] == "") {
        document.getElementById("item-error").style.display = "block";
        document.getElementById("item-subject").style.display = "none";
        document.getElementById("loader").style.display = "none";
      } else {
        document.getElementById("item-error").style.display = "none";
        let formData = new FormData();
        let params = {};
        params = JSON.parse(JSON.stringify(meeting_object));
        for (let key in params) {
          if (params.hasOwnProperty(key)) {
            if (typeof params[key] == "object") {
              formData.append(key, JSON.stringify(params[key]));
            } else {
              formData.append(key, params[key]);
            }
          }
        }
        saveMeeting(formData);
      }
    }
  } else {
    const options = { asyncContext: { currentItem: item } };
    item.getAttachmentsAsync(options, callback);

    async function callback(result) {
      if (result.value.length > 0) {
        const attachmentPromises = result.value.map(async (attachment) => {
          return new Promise((resolve) => {
            result.asyncContext.currentItem.getAttachmentContentAsync(attachment.id, (contentResult) => {
              resolve({ attachment, contentResult });
            });
          });
        });

        const attachmentsWithContent = await Promise.all(attachmentPromises);
        let counterFiles = attachmentsWithContent.length;
        for (const { attachment, contentResult } of attachmentsWithContent) {
          switch (contentResult.value.format) {
            case Office.MailboxEnums.AttachmentContentFormat.Base64:
              uploadFile(meeting_object, attachment, contentResult, counterFiles);
              break;
            case Office.MailboxEnums.AttachmentContentFormat.Eml:
              // Handle email item attachment.
              break;
            case Office.MailboxEnums.AttachmentContentFormat.ICalendar:
              // Handle .icalendar attachment.
              break;
            case Office.MailboxEnums.AttachmentContentFormat.Url:
              // Handle cloud attachment.
              break;
            default:
              // Handle attachment formats that are not supported.
              break;
          }
        }
      } else {
        if (meeting_object["meeting_name"] == "") {
          document.getElementById("item-error").style.display = "block";
          document.getElementById("item-subject").style.display = "none";
          document.getElementById("loader").style.display = "none";
        } else {
          document.getElementById("item-error").style.display = "none";
          let formData = new FormData();
          let params = {};
          params = JSON.parse(JSON.stringify(meeting_object));
          for (let key in params) {
            if (params.hasOwnProperty(key)) {
              if (typeof params[key] == "object") {
                formData.append(key, JSON.stringify(params[key]));
              } else {
                formData.append(key, params[key]);
              }
            }
          }
          saveMeeting(formData);
        }
      }
    }
  }
}

function getTags() {
  if (Office.context.mailbox.item.itemId) {
    let itemId = Office.context.mailbox.item.itemId;
    processTags(itemId);
  } else {
    Office.context.mailbox.item.saveAsync(function (result) {
      let itemId = result.value;
      processTags(itemId);
    });
  }
}

function processTags(itemId) {
  let tags_saved = [];
  if (Office.context.roamingSettings.get("tags" + itemId)) {
    tags_saved = JSON.parse(Office.context.roamingSettings.get("tags" + itemId));
  }
  try {
    const response = fetch(
      "https://www.closir.com/api/public/investor-interest/company-meeting?ot=true&token=123456&requesterEmail=moussa.mansour99@outlook.com",
      {
        method: "GET",
        headers: {
          "Content-Type": "application/json",
        },
      }
    );
    response.then((data) => {
      data.json().then((result) => {
        if (result) {
          if (result.length == 0) document.getElementById("topics").style.display = "none";
          else {
            document.getElementById("topics").style.display = "block";
            for (let i = 0; i < result.length; i++) {
              let counter = 0;
              let container = document.createElement("div");
              container.className = "tagContainer";
              let tagName = document.createElement("p");
              tagName.innerHTML = result[i];
              tagName.className = "tagName";
              let plus = document.createElement("div");
              plus.className = "counterControl tagInput";
              plus.innerHTML = "+";
              plus.style.display = "none";
              plus.id = "tag_" + result[i];
              let minus = document.createElement("div");
              minus.className = "counterControl";
              minus.innerHTML = "-";
              minus.style.display = "none";

              for (let j = 0; j < tags_saved.length; j++) {
                if (tags_saved[j]["name"] == result[i]){
                  counter = tags_saved[j].weight;
                  plus.innerHTML = counter;
                  container.style.background = "#bb252e";
                  tagName.style.color = "#ffffff";
                  plus.style.display = "block";
                  plus.style.color = "#ffffff";
                  minus.style.display = "block";
                  minus.style.color = "#ffffff";
                }
              }
              container.append(minus);
              container.append(tagName);
              container.append(plus);
              //container.append(tagInput);
              tagName.addEventListener("click", function (e) {
                if (counter == 0) {
                  counter++;
                  plus.innerHTML = "+";
                  container.style.background = "#bb252e";
                  tagName.style.color = "#ffffff";
                  plus.style.display = "block";
                  plus.style.color = "#ffffff";
                  minus.style.display = "block";
                  minus.style.color = "#ffffff";
                } else {
                  counter = 0;
                  container.style.background = "#e6e6e6";
                  tagName.style.color = "#000000";
                  plus.style.display = "none";
                  plus.style.color = "#000000";
                  minus.style.display = "none";
                  minus.style.color = "#000000";
                }
              });
              plus.addEventListener("click", function (e) {
                counter++;
                plus.innerHTML = counter;
              });
              minus.addEventListener("click", function (e) {
                counter--;
                plus.innerHTML = counter;
                if (counter == 0) {
                  container.style.background = "#e6e6e6";
                  tagName.style.color = "#000000";
                  plus.style.display = "none";
                  plus.style.color = "#000000";
                  minus.style.display = "none";
                  minus.style.color = "#000000";
                }
              });
              document.getElementById("topics").append(container);
              if (Office.context.roamingSettings.get("meeting" + itemId)) {
                let inputs = document.getElementsByTagName("INPUT");
                for (let j = 0; j < inputs.length; j++) {
                  inputs[j].disabled = true;
                }
              }
            }
          }
        } else {
          document.getElementById("topics").style.display = "none";
        }
      });
    });
  } catch (e) {
    document.getElementById("topics").style.display = "none";
  }
}

function uploadFile(meeting_object, attachment, contentResult, counterFiles) {
  let decoded = atob(contentResult.value.content);
  const blobArray = new Uint8Array(decoded.length);
  for (let i = 0; i < decoded.length; i++) {
    blobArray[i] = decoded.charCodeAt(i);
  }
  const blob = new Blob([blobArray], { type: "application/octet-stream" });
  let s3 = new AWS.S3();
  let bucketName = AWS_BUCKET;
  let time = new Date().getTime();
  let ext = attachment.name.substr(attachment.name.lastIndexOf(".") + 1);
  let hashedName = CryptoJS.SHA512(attachment.name + time).toString();
  hashedName = hashedName + "." + ext;
  s3.upload(
    {
      Bucket: bucketName,
      Body: blob,
      Key: hashedName,
      ACL: "public-read",
    },
    function (err, data) {
      if (err) {
        console.log("Error uploading to S3:", err);
      } else {
        let params = {};

        params["file_extension"] = ext;
        params["aws_url"] = data.Location;
        params["aws_bucket"] = data.Bucket;
        params["file_name"] = attachment.name;
        params["hashed_filename"] = data.Key;
        params["file_info"] = "File info";
        params["mime_type"] = attachment.contentType;
        params["event_file_category"] = "Other";
        params["editable"] = false;

        let formData = new FormData();
        for (let key in params) {
          if (params.hasOwnProperty(key)) {
            formData.append(key, params[key]);
          }
        }
        try {
          const url = new URL(
            `https://www.closir.com/api/public/file?token=123456&requesterEmail=moussa.mansour99@outlook.com`
          );
          fetch(url, {
            method: "POST",
            body: formData,
          }).then((res) => {
            res.json().then((data) => {
              counterFiles--;
              attachmentsIDs.push(data.id);
              if (counterFiles == 0) {
                meeting_object["uploaded_files"] = attachmentsIDs.toString();
                console.log(meeting_object);
                if (meeting_object["meeting_name"] == "") {
                  document.getElementById("item-error").style.display = "block";
                  document.getElementById("loader").style.display = "none";
                  document.getElementById("item-subject").style.display = "none";
                } else {
                  document.getElementById("item-error").style.display = "none";
                  let formData = new FormData();
                  for (let key in meeting_object) {
                    if (meeting_object.hasOwnProperty(key)) {
                      if (typeof meeting_object[key] == "object")
                        formData.append(key, JSON.stringify(meeting_object[key]));
                      else formData.append(key, meeting_object[key]);
                    }
                  }
                  saveMeeting(formData);
                }
              }
            });
          });
        } catch (e) {
          console.log(e);
          let formData = new FormData();
          for (let key in meeting_object) {
            if (meeting_object.hasOwnProperty(key)) {
              if (typeof meeting_object[key] == "object") formData.append(key, JSON.stringify(meeting_object[key]));
              else formData.append(key, meeting_object[key]);
            }
          }
          saveMeeting(formData);
        }
      }
    }
  );
}

function saveMeeting(meeting_object) {
  let update = false;

  let itemId = Office.context.mailbox.item.itemId;
  if (itemId == null || itemId == undefined) {
    Office.context.mailbox.item.saveAsync(function (result) {
      itemId = result.value;
      saveMeetingCall(itemId, meeting_object);
    });
  } else {
    saveMeetingCall(itemId, meeting_object);
  }
}

function saveMeetingCall(itemId, meeting_object) {
  let update = false;
  try {
    const url = new URL(
      `https://www.closir.com/api/public/investor-interest/company-meeting?token=123456&requesterEmail=moussa.mansour99@outlook.com`
    );
    if (Office.context.roamingSettings.get("id" + itemId)) {
      update = true;
    }
    fetch(url, {
      //method: update ? "PATCH" : "POST",
      method: "POST",
      body: meeting_object,
    }).then((res) => {
      res.json().then((data) => {
        if (data) {
          Office.context.roamingSettings.set("meeting" + itemId, itemId);
          Office.context.roamingSettings.set("meeting_type" + itemId, data["meeting_type"]);
          Office.context.roamingSettings.set("host" + itemId, data["host"]);
          Office.context.roamingSettings.set("tags" + itemId, data["tags"]);
          Office.context.roamingSettings.set("id" + itemId, data["id"]);
          if (data["host_name"]) Office.context.roamingSettings.set("host_name" + itemId, data["host_name"]);

          Office.context.roamingSettings.saveAsync(function (result) {
            if (result.status == Office.AsyncResultStatus.Succeeded) {
              document.getElementById("approve_adding_meeting_text").innerText = "Update meeting on Closir";
              document.getElementById("approve_adding_meeting").style.display = "none";
              document.getElementById("item-subject").style.display = "none";
              document.getElementById("meeting_saved").style.display = "block";
              document.getElementById("meeting_saved_message").style.display = "block";
              setTimeout(() => {
                document.getElementById("meeting_saved_message").style.display = "none";
                document.getElementById("loader").style.display = "none";
                let inputs = document.getElementsByTagName("INPUT");
                for (let j = 0; j < inputs.length; j++) {
                  inputs[j].disabled = true;
                }
                document.getElementById("meeting_types").disabled = true;
              }, 5000);
              console.log("Token saved successfully!");
            } else {
              document.getElementById("loader").style.display = "none";
              document.getElementById("meeting_failed_message").style.display = "block";
              setTimeout(() => {
                document.getElementById("meeting_failed_message").style.display = "none";
              }, 6000);
              console.log("Error saving token: " + result.error.message);
            }
          });
        } else {
          document.getElementById("loader").style.display = "none";
          document.getElementById("meeting_failed_message").style.display = "block";
          setTimeout(() => {
            document.getElementById("meeting_failed_message").style.display = "none";
          }, 6000);
        }
      });
    });
  } catch (e) {
    document.getElementById("meeting_failed_message").style.display = "block";
    setTimeout(() => {
      document.getElementById("meeting_failed_message").style.display = "none";
    }, 6000);
    console.log(e);
  }
}

/* formatting functions */
function formatDate(date) {
  const year = date.getFullYear();
  const month = String(date.getMonth() + 1).padStart(2, "0"); // Months are 0-indexed
  const day = String(date.getDate()).padStart(2, "0");
  const hours = String(date.getHours()).padStart(2, "0");
  const minutes = String(date.getMinutes()).padStart(2, "0");
  const seconds = String(date.getSeconds()).padStart(2, "0");

  return `${year}-${month}-${day} ${hours}:${minutes}:${seconds}`;
}

function getTimeZone(date) {
  const offsetMinutes = date.getTimezoneOffset();
  const offsetHours = Math.abs(Math.floor(offsetMinutes / 60));
  const offsetMinutesPart = Math.abs(offsetMinutes % 60);
  const offsetSign = offsetMinutes < 0 ? "+" : "-";
  const formattedOffset = `GMT${offsetSign}${offsetHours.toString().padStart(2, "0")}:${offsetMinutesPart
    .toString()
    .padStart(2, "0")}`;

  return formattedOffset;
}
