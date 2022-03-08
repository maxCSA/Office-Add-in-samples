// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

// Contains code for event-based activation on both Outlook on web and Outlook on Windows.

/**
 * Checks if signature exists.
 * If not, displays the information bar for user.
 * If exists, insert signature into new item (appointment or message).
 * @param {*} eventObj Office event object
 * @returns
 */
function checkSignature(eventObj) {
  let user_info_str = Office.context.roamingSettings.get("user_info");
  if (!user_info_str) {
    display_insight_infobar();
  } else {
    let user_info = JSON.parse(user_info_str);

    if (Office.context.mailbox.item.getComposeTypeAsync) {
      //Find out if the compose type is "newEmail", "reply", or "forward" so that we can apply the correct template.
      Office.context.mailbox.item.getComposeTypeAsync(
        {
          asyncContext: {
            user_info: user_info,
            eventObj: eventObj,
          },
        },
        function (asyncResult) {
          if (asyncResult.status === "succeeded") {
            insert_auto_signature(
              asyncResult.value.composeType,
              asyncResult.asyncContext.user_info,
              asyncResult.asyncContext.eventObj
            );
          }
        }
      );
    } else {
      // Appointment item. Just use newMail pattern
      let user_info = JSON.parse(user_info_str);
      insert_auto_signature("newMail", user_info, eventObj);
    }
  }
}

/**
 * For Outlook on Windows only. Insert signature into appointment or message.
 * Outlook on Windows can use setSignatureAsync method on appointments and messages.
 * @param {*} compose_type The compose type (reply, forward, newMail)
 * @param {*} user_info Information details about the user
 * @param {*} eventObj Office event object
 */
function insert_auto_signature(compose_type, user_info, eventObj) {
  let template_name = get_template_name(compose_type);
  let signature_info = get_signature_info(template_name, user_info);
  addTemplateSignature(signature_info, eventObj);
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
function addTemplateSignature(signatureDetails, eventObj, signatureImageBase64) {
  if (is_valid_data(signatureDetails.logoBase64) === true) {
    //If a base64 image was passed we need to attach it.
    Office.context.mailbox.item.addFileAttachmentFromBase64Async(
      signatureDetails.logoBase64,
      signatureDetails.logoFileName,
      {
        isInline: true,
      },
      function (result) {
        //After image is attached, insert the signature
        Office.context.mailbox.item.body.setSignatureAsync(
          signatureDetails.signature,
          {
            coercionType: "html",
            asyncContext: eventObj,
          },
          function (asyncResult) {
            asyncResult.asyncContext.completed();
          }
        );
      }
    );
  } else {
    //Image is not embedded, or is referenced from template HTML
    Office.context.mailbox.item.body.setSignatureAsync(
      signatureDetails.signature,
      {
        coercionType: "html",
        asyncContext: eventObj,
      },
      function (asyncResult) {
        asyncResult.asyncContext.completed();
      }
    );
  }
}

/**
 * Creates information bar to display when new message or appointment is created
 */
function display_insight_infobar() {
  Office.context.mailbox.item.notificationMessages.addAsync("fd90eb33431b46f58a68720c36154b4a", {
    type: "insightMessage",
    message: "Please set your signature with the PnP sample add-in.",
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
    "' alt='MS Logo' width='218' height='76' /></td>";
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
      "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",
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
    "<td style='border-right: 1px solid #000000; padding-right: 5px;'><img src='https://github.com/maxCSA/Office-Add-in-samples/raw/main/Samples/outlook-set-signature/assets/sample-logo.png' alt='Logo' /></td>";
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
