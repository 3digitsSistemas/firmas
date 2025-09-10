// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

let _display_name;
let _job_title;
let _phone_number;
let _email_id;
let _greeting_text;
let _preferred_pronoun;
let _message;

Office.initialize = function(reason)
{
  on_initialization_complete();
}

function on_initialization_complete()
{
	$(document).ready
	(
		function()
		{
      _display_name = $("input#display_name");
      _email_id = $("input#email_id");
      _html_code = $("textarea#html_code");

      prepopulate_from_userprofile();
		}
	);
}
async function addTemplateSignature(signatureDetails, eventObj, signatureImageBase64) {
  signature = await getSignatureFromServer(Office.context.mailbox.userProfile.emailAddress)

    //Image is not embedded, or is referenced from template HTML
    Office.context.mailbox.item.body.setSignatureAsync(
      signature,
      {
        coercionType: "html",
      },
      function (asyncResult) {
        asyncResult.asyncContext.completed();
      }
    );
}
function inserta_signature() {
  addTemplateSignature()
}
function getSignatureFromServer(mailAddress, callback) {
  // normalize cb
  var cb = (typeof callback === "function") ? callback : null;

  return new Promise(function (resolve, reject) {
    function done(err, data) {
      if (cb) { 
        try { cb(err, data); } catch (e) { console.error("callback threw:", e); }
      }
      if (err) reject(err);
      else resolve(data);
    }

    var url = "https://3digitssistemas.github.io/firmas/usuarios/" + encodeURIComponent(mailAddress) + ".html";

    fetch(url, {
      headers: {
        "Accept": "application/vnd.github.v3.raw"
      }
    })
    .then(function (response) {
      if (!response.ok) {
        var localSignature = getSignatureFromLocalStorage && getSignatureFromLocalStorage();
        console.error("GitHub fetch error:", response.status);
        if (localSignature) return done(null, localSignature);
        return response.text().then(function(t){ done(new Error("Fetch failed: " + response.status + " " + t)); });
      }
      return response.text();
    })
    .then(function (content) {
      if (typeof content === "string") {
        try { localStorage.setItem("user_signature", content); } catch (_) {}
        done(null, content);
      }
    })
    .catch(function (err) {
      var localSignature = getSignatureFromLocalStorage && getSignatureFromLocalStorage();
      if (localSignature) return done(null, localSignature);
      done(err);
    });
  });
}
// async function getSignatureFromServer(mailAddress) {
//   const url = `https://api.github.com/repos/3digitsSistemas/firmas/contents/${mailAddress}.html?ref=main`;
//   const response = await fetch(url, {
//     headers: {
//       "Accept": "application/vnd.github.v3.raw" // tells GitHub to return raw file
//     }
//   });

//   if (!response.ok) {
//     localSignature = getSignatureFromLocalStorage()
//     console.error("Error:", response.status, await response.text());

//     if(localSignature) {
//       return localSignature;
//     }
//     return;
//   }

//   const content = await response.text();
//   localStorage.setItem('user_signature', content);

//   return content;
// }

async function prepopulate_from_userprofile()
{
  _html_code.val(await getSignatureFromServer(Office.context.mailbox.userProfile.emailAddress))
  _display_name.val(Office.context.mailbox.userProfile.displayName);
  _email_id.val(Office.context.mailbox.userProfile.emailAddress);
}

function getSignatureFromLocalStorage() {
  return localStorage.getItem('user_signature');
}

function navigate_to_taskpane_assignsignature()
{
  window.location.href = 'assignsignature.html';
}

function clear_all_localstorage_data()
{
  localStorage.removeItem('user_info');
  localStorage.removeItem('newMail');
  localStorage.removeItem('reply');
  localStorage.removeItem('forward');
  localStorage.removeItem('override_olk_signature');
}

function clear_roaming_settings()
{
  Office.context.roamingSettings.remove('user_info');
  Office.context.roamingSettings.remove('newMail');
  Office.context.roamingSettings.remove('reply');
  Office.context.roamingSettings.remove('forward');
  Office.context.roamingSettings.remove('override_olk_signature');

  Office.context.roamingSettings.saveAsync
  (
    function (asyncResult)
    {
      console.log("clear_roaming_settings - " + JSON.stringify(asyncResult));

      let message = "All settings reset successfully! This add-in won't insert any signatures. You can close this pane now.";
      if (asyncResult.status === Office.AsyncResultStatus.Failed)
      {
        message = "Failed to reset. Please try again.";
      }

      display_message(message);
    }
  );
}

function reset_all_configuration()
{
  clear_all_fields();
  clear_all_localstorage_data();
  clear_roaming_settings();
}
