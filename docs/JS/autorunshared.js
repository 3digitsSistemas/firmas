// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

// Contains code for event-based activation on Outlook on web, on Windows, and on Mac (new UI preview).

/**
 * Checks if signature exists.
 * If not, displays the information bar for user.
 * If exists, insert signature into new item (appointment or message).
 * @param {*} eventObj Office event object
 * @returns
*/
// function checkSignature(eventObj) {
//   console.log("event", eventObj)
//   let user_info_str = Office.context.roamingSettings.get("user_info");

//   if (!user_info_str) {
//     display_insight_infobar();
//   } else {
//     addTemplateSignature(eventObj)
//   }
// }
function checkSignature(eventObj) {
  try {
    // const user_info_str = Office.context.roamingSettings.get("user_info");

    // if (!user_info_str) {
    //   // No hay datos: muestra infobar y COMPLETA el evento
    //   display_insight_infobar("NO INFO STR");
    //   eventObj.completed();
    //   return;
    // }

    // Hay datos: intenta insertar firma
    addTemplateSignature(eventObj);
  } catch (e) {
    console.error("checkSignature error:", e);
    try { eventObj.completed(); } catch (_) {}
  }
}
/**
 * For Outlook on Windows and on Mac only. Insert signature into appointment or message.
 * Outlook on Windows and on Mac can use setSignatureAsync method on appointments and messages.
 * @param {*} compose_type The compose type (reply, forward, newMail)
 * @param {*} user_info Information details about the user
 * @param {*} eventObj Office event object
 */
async function insert_auto_signature(compose_type, user_info, eventObj) {
  let template_name = get_template_name(compose_type);
  let signature_info = get_signature_info(template_name, user_info);
  await addTemplateSignature(signature_info, eventObj);
}

/**
 * 
 * @param {*} signatureDetails object containing:
 *  "signature": The signature HTML of the template,
    "logoBase64": The base64 encoded logo image,
    "logoFileName": The filename of the logo image
 * @param {*} eventObj 
 * @param {*} signatureImageBase64 
 */
// async function addTemplateSignature(event) {
//   let signature = await getSignatureFromServer(Office.context.mailbox.userProfile.emailAddress)
//   //Image is not embedded, or is referenced from template HTML
//   Office.context.mailbox.item.body.setSignatureAsync(
//     signature,
//     {
//       coercionType: "html",
//       asyncContext: event
//     },
//     function (asyncResult) {
//       asyncResult.asyncContext.completed();
//     }
//   );
// }
// function addTemplateSignature(eventObj) {
//   try {
//     // 1) Obtén la firma (HTML)
//     const signatureHtml = getSignatureFromServer(
//       Office.context.mailbox.userProfile.emailAddress
//     );
//     // const signatureHtml = "hola"

//     // 2) Asegúrate de que la API existe en este host
//     const body = Office?.context?.mailbox?.item?.body;
//     if (!body || typeof body.setSignatureAsync !== "function") {
//       console.warn("setSignatureAsync no disponible en esta plataforma.");
//       eventObj.completed();
//       return;
//     }

//     // 3) Inserta la firma
//     body.setSignatureAsync(
//       signatureHtml,
//       { coercionType: "html" /*, append: false  <- si quieres añadir en vez de reemplazar */ },
//       function (asyncResult) {
//         if (asyncResult.status === Office.AsyncResultStatus.Failed) {
//           console.error("setSignatureAsync error:", asyncResult.error);
//         }
//         // 4) SIEMPRE completar el evento
//         eventObj.completed();
//       }
//     );
//   } catch (e) {
//     console.error("addTemplateSignature error:", e);
//     try { eventObj.completed(); } catch (_) {}
//   }
// }
function addTemplateSignature(eventObj) {
  display_insight_infobar("hola")
  try {
    var item = Office && Office.context && Office.context.mailbox && Office.context.mailbox.item;
    var body = item && item.body;

    if (!body || typeof body.setSignatureAsync !== "function") {
      console.warn("setSignatureAsync unavailable on this platform/item.");
      try { eventObj.completed(); } catch (_) {}
      return;
    }

    var email = Office.context.mailbox.userProfile.emailAddress;

    // Using the optional-callback getSignatureFromServer
    getSignatureFromServer(email, function (err, signatureHtml) {
      if (err || !signatureHtml) {
        console.error("Unable to get signature:", err || "empty");
        try { eventObj.completed(); } catch (_) {}
        return;
      }

      body.setSignatureAsync(
        signatureHtml,
        { coercionType: "html" /*, append: false */ },
        function (asyncResult) {
          if (asyncResult.status === Office.AsyncResultStatus.Failed) {
            console.error("setSignatureAsync error:", asyncResult.error);
          }
          try { eventObj.completed(); } catch (_) {}
        }
      );
    });
  } catch (e) {
    display_insight_infobar("error 1")
    console.error("addTemplateSignature exception:", e);
    try { eventObj.completed(); } catch (_) {}
  }
}

/**
 * Creates information bar to display when new message or appointment is created
 */
function display_insight_infobar(msg = "Please set your signature with the Office Add-ins sample.") {
  Office.context.mailbox.item.notificationMessages.addAsync("fd90eb33431b46f58a68720c36154b4a", {
    type: "insightMessage",
    message: msg,
    icon: "Icon.16x16",
    actions: [
      {
        actionType: "showTaskPane",
        actionText: "Set signatures",
        commandId: get_command_id(),
        contextData: "{''}",
      },
    ],
  });
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
      display_insight_infobar("catch err")
      var localSignature = getSignatureFromLocalStorage && getSignatureFromLocalStorage();
      if (localSignature) return done(null, localSignature);
      done(err);
    });
  });
}
/**
 * Gets template name (A,B,C) mapped based on the compose type
 * @param {*} compose_type The compose type (reply, forward, newMail)
 * @returns Name of the template to use for the compose type
 */
function get_template_name(compose_type) {
  if (compose_type === "reply") return Office.context.roamingSettings.get("reply");
  if (compose_type === "forward") return Office.context.roamingSettings.get("forward");
  return Office.context.roamingSettings.get("newMail");
}

/**
 * Gets HTML signature in requested template format for given user
 * @param {\} template_name Which template format to use (A,B,C)
 * @param {*} user_info Information details about the user
 * @returns HTML signature in requested template format
 */
function get_signature_info(template_name, user_info) {
  if (template_name === "templateB") return get_template_B_info(user_info);
  if (template_name === "templateC") return get_template_C_info(user_info);
  return get_template_A_info(user_info);
}

/**
 * Gets correct command id to match to item type (appointment or message)
 * @returns The command id
 */
function get_command_id() {
  if (Office.context.mailbox.item.itemType == "appointment") {
    return "MRCS_TpBtn1";
  }
  return "MRCS_TpBtn0";
}

/**
 * Gets HTML string for template A
 * Embeds the signature logo image into the HTML string
 * @param {*} user_info Information details about the user
 * @returns Object containing:
 *  "signature": The signature HTML of template A,
    "logoBase64": The base64 encoded logo image,
    "logoFileName": The filename of the logo image
 */
function get_template_A_info(user_info) {
  const logoFileName = "sample-logo.png";
  let str = "";
  if (is_valid_data(user_info.greeting)) {
    str += user_info.greeting + "<br/>";
  }

  str += "<table>";
  str += "<tr>";
  // Embed the logo using <img src='cid:...
  str +=
    "<td style='border-right: 1px solid #000000; padding-right: 5px;'><img src='cid:" +
    logoFileName +
    "' alt='MS Logo' width='24' height='24' /></td>";
  str += "<td style='padding-left: 5px;'>";
  str += "<strong>" + user_info.name + "</strong>";
  str += is_valid_data(user_info.pronoun) ? "&nbsp;" + user_info.pronoun : "";
  str += "<br/>";
  str += is_valid_data(user_info.job) ? user_info.job + "<br/>" : "";
  str += user_info.email + "<br/>";
  str += is_valid_data(user_info.phone) ? user_info.phone + "<br/>" : "";
  str += "</td>";
  str += "</tr>";
  str += "</table>";

  // return object with signature HTML, logo image base64 string, and filename to reference it with.
  return {
    signature: str,
    logoBase64:
      "iVBORw0KGgoAAAANSUhEUgAAACIAAAAiCAYAAAA6RwvCAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAJcEhZcwAADsQAAA7EAZUrDhsAAAEeSURBVFhHzdhBEoIwDIVh4EoeQJd6YrceQM+kvo5hQNokLymO/4aF0/ajlBl1fL4bEp0uj3K9XQ/lGi0MEcB3UdD0uVK1EEj7TIuGeBaKYCgIswCLcUMid8mMcUEiCMk71oRYE+Etsd4UD0aFeBBSFtOEMAgpg6lCIggpitlAMggpgllBeiAkFjNDeiIkBlMgeyAkL6Z6WJdlEJJnjvF4vje/BvRALNN23tyRXzVpd22dHSZtLhjMHemB8cxRINZZyGCssbL2vCN7YLwItHo0PTEMAm3OSA8Mi0DVw5rBRBCoCkERTBSBmhDEYDII5PqlZy1iZSGQuiOSZ6JW3rEuCIpgmDFuCGImZuEUBHkWiOweDUHaQhEE+pM/aobhBZaOpYLJeeeoAAAAAElFTkSuQmCC",
    logoFileName: logoFileName,
  };
}

/**
 * Gets HTML string for template B
 * References the signature logo image from the HTML
 * @param {*} user_info Information details about the user
 * @returns Object containing:
 *  "signature": The signature HTML of template B,
    "logoBase64": null since this template references the image and does not embed it ,
    "logoFileName": null since this template references the image and does not embed it
 */
function get_template_B_info(user_info) {
  let str = "";
  if (is_valid_data(user_info.greeting)) {
    str += user_info.greeting + "<br/>";
  }

  str += "<table>";
  str += "<tr>";
  // Reference the logo using a URI to the web server <img src='https://...
  str +=
    "<td style='border-right: 1px solid #000000; padding-right: 5px;'><img src='https://officedev.github.io/Office-Add-in-samples/Samples/outlook-set-signature/assets/sample-logo.png' alt='Logo' /></td>";
  str += "<td style='padding-left: 5px;'>";
  str += "<strong>" + user_info.name + "</strong>";
  str += is_valid_data(user_info.pronoun) ? "&nbsp;" + user_info.pronoun : "";
  str += "<br/>";
  str += user_info.email + "<br/>";
  str += is_valid_data(user_info.phone) ? user_info.phone + "<br/>" : "";
  str += "</td>";
  str += "</tr>";
  str += "</table>";

  return {
    signature: str,
    logoBase64: null,
    logoFileName: null,
  };
}

/**
 * Gets HTML string for template C
 * @param {*} user_info Information details about the user
 * @returns Object containing:
 *  "signature": The signature HTML of template C,
    "logoBase64": null since there is no image,
    "logoFileName": null since there is no image
 */
function get_template_C_info(user_info) {
  let str = "";
  if (is_valid_data(user_info.greeting)) {
    str += user_info.greeting + "<br/>";
  }

  str += user_info.name;

  return {
    signature: str,
    logoBase64: null,
    logoFileName: null,
  };
}

/**
 * Validates if str parameter contains text.
 * @param {*} str String to validate
 * @returns true if string is valid; otherwise, false.
 */
function is_valid_data(str) {
  return str !== null && str !== undefined && str !== "";
}

Office.actions.associate("checkSignature", checkSignature);
