/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office, console, setTimeout */

const EMAIL_TEMPLATES = {
  offer: {
    cc: "group@ship-around.com",
    intro: "Dear {name},<br>",
    body: "Please find attached:<br>",
    attachments: "{attachments}",
    note: "If you accept our offer, please note that the last page of our quotation is the proforma invoice.<br>",
    closing: "Looking forward to your order confirmation.",
  },
  acknowledge: {
    cc: "group@ship-around.com",
    intro: "Dear {name},<br>",
    body: "Thank you for reaching out to us.<br><br>We have logged your inquiry with reference SALE{lead}.<br>",
    note: "Please include the above reference in any future correspondence.<br>",
    closing:
      "We appreciate your interest and will get back to you shortly.<br><br>If you haven't already, please <a href='https://ship-around.com/register'>register</a> a free buyer account.<br><br>It only takes 5 minutes and will expedite processing your request.",
  },
  follow_up_1: {
    cc: "group@ship-around.com",
    intro: "Dear {name},<br>",
    body: "I am following up regarding our last quotation {quote_reference} for {quote_items}.<br><br>We would like to know if you are still interested in pursuing this order.<br>",
    note: "I have attached said quotation again for your perusal.<br>",
    closing:
      "Please let us know of your decision at your earliest convenience and if there is any way we can assist you further.<br><br>We appreciate your interest in Ship-Around for your procurement needs.",
  },
};

const DOCUMENT_TYPE_MAPPINGS = {
  Q202: "Quotation",
  DN202: "Delivery Note",
  PL202: "Packing List",
  INV202: "Invoice",
};

Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("acknowledge").onclick = acknowledgeRFQ;
    document.getElementById("prepare-quote-email").onclick = prepareQuoteEmail;
    document.getElementById("follow-up").onclick = followUp1;
  }
});

class EmailUtility {
  constructor(item) {
    this.item = item;
  }

  getEmailContent(templateType, replacements) {
    if (!EMAIL_TEMPLATES[templateType]) {
      throw new Error(`No template found for type: ${templateType}`);
    }

    const template = EMAIL_TEMPLATES[templateType];
    let content = "";

    for (const [key, text] of Object.entries(template)) {
      if (key === "cc") continue;

      let sectionContent = text;
      for (const [replaceKey, value] of Object.entries(replacements)) {
        sectionContent = sectionContent.replace(`{${replaceKey}}`, value);
      }
      content += sectionContent + "<br>";
    }

    return content;
  }

  async addSubject(prefix, prepend = true) {
    return new Promise((resolve, reject) => {
      this.item.subject.getAsync((result) => {
        if (result.status === Office.AsyncResultStatus.Failed) {
          reject(result.error);
        } else {
          let updatedSubject;
          if (prepend) {
            updatedSubject = prefix + result.value;
          } else {
            updatedSubject = prefix;
          }

          this.item.subject.setAsync(updatedSubject, (setResult) => {
            if (setResult.status === Office.AsyncResultStatus.Failed) {
              reject(setResult.error);
            } else {
              resolve();
            }
          });
        }
      });
    });
  }

  async addCC(emailAddress, replace = false) {
    return new Promise((resolve, reject) => {
      this.item.cc.getAsync((result) => {
        if (result.status === Office.AsyncResultStatus.Failed) {
          reject(result.error);
        } else {
          let updatedCC;
          if (replace) {
            updatedCC = [emailAddress];
          } else {
            const currentCC = result.value;
            if (!currentCC.includes(emailAddress)) {
              updatedCC = [...currentCC, emailAddress];
            } else {
              resolve();
              return;
            }
          }

          this.item.cc.setAsync(updatedCC, (setResult) => {
            if (setResult.status === Office.AsyncResultStatus.Failed) {
              reject(setResult.error);
            } else {
              resolve();
            }
          });
        }
      });
    });
  }

  async addBody(content) {
    return new Promise((resolve, reject) => {
      this.item.body.prependAsync(content, { coercionType: Office.CoercionType.Html }, (result) => {
        if (result.status === Office.AsyncResultStatus.Failed) {
          reject(result.error);
        } else {
          resolve();
        }
      });
    });
  }

  displayErrorInTaskpane(errorMessage) {
    const errorDiv = document.createElement("div");
    errorDiv.style.color = "red";
    errorDiv.textContent = errorMessage;
    document.body.appendChild(errorDiv);
  }

  getDocumentType(nameWithoutExtension) {
    for (const prefix in DOCUMENT_TYPE_MAPPINGS) {
      if (nameWithoutExtension.startsWith(prefix)) {
        return `${DOCUMENT_TYPE_MAPPINGS[prefix]} ${nameWithoutExtension}`;
      }
    }
    return nameWithoutExtension;
  }

  generateAttachmentTable(attachmentNames) {
    let attachmentTable = "";
    if (attachmentNames && attachmentNames.length > 0) {
      attachmentTable = '<table style="border-collapse: collapse;">';
      attachmentNames.forEach((name, index) => {
        attachmentTable += `<tr style="padding: 2px; background-color: #f5f5f5;"><td style="border: 1px solid; padding: 2px 4px;">${
          index + 1
        }</td><td style="border: 1px solid gray; padding: 2px 4px;">${name}</td></tr>`;
      });
      attachmentTable += "</table>";
    }
    return attachmentTable;
  }

  async listAttachments() {
    return new Promise((resolve, reject) => {
      this.item.getAttachmentsAsync((result) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          const attachments = result.value;
          const fileAttachments = attachments.filter(
            (attachment) =>
              attachment.attachmentType === Office.MailboxEnums.AttachmentType.File && !attachment.isInline
          );

          if (fileAttachments && fileAttachments.length > 0) {
            const attachmentNamesWithoutExtensions = fileAttachments.map((attachment) => {
              let nameWithoutExtension = attachment.name.split(".").slice(0, -1).join(".");
              nameWithoutExtension = this.getDocumentType(nameWithoutExtension);
              return this.capitalizeFirstLetter(nameWithoutExtension);
            });

            resolve(attachmentNamesWithoutExtensions);
          } else {
            console.log("The current message has no file attachments.");
            resolve([]);
          }
        } else {
          console.error("Failed to get attachments:", result.error);
          reject(result.error);
        }
      });
    });
  }

  capitalizeFirstLetter(string) {
    return string.charAt(0).toUpperCase() + string.slice(1);
  }
}

class Modal {
  constructor(modalId, inputDivIds, okButtonId, cancelButtonId) {
    this.modal = document.getElementById(modalId);
    this.allInputDivs = Array.from(this.modal.querySelectorAll("div[id$='InputDiv']"));
    this.inputDivs = inputDivIds.map((id) => document.getElementById(id));
    this.okButton = document.getElementById(okButtonId);
    this.cancelButton = document.getElementById(cancelButtonId);
    this.setupEventListeners();
  }

  setupEventListeners() {
    this.okButton.onclick = () => {
      const inputValues = this.inputDivs.map((div) => div.querySelector("input").value);
      this.resolve(inputValues);
      this.clearInputs();
      this.hide();
    };

    this.cancelButton.onclick = () => {
      this.reject(new Error("User cancelled the input."));
      this.hide();
    };
  }

  clearInputs() {
    this.inputDivs.forEach((div) => {
      const input = div.querySelector("input");
      if (input) {
        input.value = "";
      }
    });
  }

  show() {
    // Hide all input divs first
    this.allInputDivs.forEach((div) => (div.style.display = "none"));

    // Only show the specified input divs
    this.inputDivs.forEach((div) => (div.style.display = "block"));

    return new Promise((resolve, reject) => {
      this.modal.style.display = "block";
      this.resolve = resolve;
      this.reject = reject;

      // Use setTimeout to ensure the modal is fully rendered before setting focus
      setTimeout(() => {
        const firstInput = this.inputDivs[0].querySelector("input");
        if (firstInput) {
          firstInput.focus();
        }
      }, 100);
    });
  }

  hide() {
    this.modal.style.display = "none";
  }
}

export async function acknowledgeRFQ() {
  let emailUtility;
  try {
    // Get a reference to the current compose item
    const item = Office.context.mailbox.item;

    emailUtility = new EmailUtility(item);
    const modal = new Modal("inputModal", ["leadInputDiv", "nameInputDiv"], "modalOk", "modalCancel");

    // Show the modal and wait for the input
    const [lead, name] = await modal.show();

    // Use the modal input to prepend to the subject
    await emailUtility.addSubject(`[SALE${lead}] `);

    // Define the email address you want to add to CC
    const ccGroupAddress = EMAIL_TEMPLATES.acknowledge.cc;

    // Simply add the group handle to CC
    await emailUtility.addCC(ccGroupAddress);

    // Get the email content
    const emailContentToAdd = emailUtility.getEmailContent("acknowledge", {
      name: name,
      lead: lead,
    });

    // Use the addBody method to prepend the content
    await emailUtility.addBody(emailContentToAdd);
  } catch (error) {
    // Use the helper function to display the error in the taskpane
    emailUtility.displayErrorInTaskpane(`Error in acknowledgeRFQ: ${error.message}`);
  }
}

export async function prepareQuoteEmail() {
  let emailUtility;
  try {
    // Get a reference to the current compose item
    const item = Office.context.mailbox.item;

    emailUtility = new EmailUtility(item);
    const modal = new Modal("inputModal", ["nameInputDiv"], "modalOk", "modalCancel");

    // Show the modal and wait for the input
    const [name] = await modal.show();

    // Get the list of attachment names
    const attachmentNames = await emailUtility.listAttachments();

    // Extract just the attachment names without the prefix
    const quotationAttachments = attachmentNames
      .filter((name) => name.startsWith("Quotation Q202"))
      .map((name) => name.replace("Quotation ", ""));

    // Determine the subject prefix based on the number of Q202 attachments
    let subjectPrefix = "";
    if (quotationAttachments.length === 1) {
      subjectPrefix = `[Quotation ${quotationAttachments[0]}] `;
    } else if (quotationAttachments.length > 1) {
      subjectPrefix = `[Quotations ${quotationAttachments.join(", ")}] `;
    }

    // Use the addSubject method to prepend the prefix to the current subject
    if (subjectPrefix) {
      await emailUtility.addSubject(subjectPrefix);
    }

    // Define the email address you want to add to CC
    const ccGroupAddress = EMAIL_TEMPLATES.offer.cc;

    // Add the group email address to CC if it's not already there
    await emailUtility.addCC(ccGroupAddress);

    // Generate the attachment table
    const attachmentTable = emailUtility.generateAttachmentTable(attachmentNames);

    // Get the email content
    const emailContentToAdd = emailUtility.getEmailContent("offer", {
      name: name,
      attachments: attachmentTable,
    });

    // Use the addBody method to prepend the content
    await emailUtility.addBody(emailContentToAdd);
  } catch (error) {
    emailUtility.displayErrorInTaskpane(`Error in prepareQuoteEmail: ${error.message}`);
  }
}

export async function followUp1() {
  let emailUtility;
  try {
    // Get a reference to the current compose item
    const item = Office.context.mailbox.item;

    emailUtility = new EmailUtility(item);
    const modal = new Modal(
      "inputModal",
      ["leadInputDiv", "nameInputDiv", "quoteInputDiv", "itemsInputDiv"],
      "modalOk",
      "modalCancel"
    );

    // Show the modal and wait for the input
    const [lead, name, reference, items] = await modal.show();

    // Use the modal input to prepend to the subject
    await emailUtility.addSubject(`[SALE${lead}] `);

    // Define the email address you want to add to CC
    const ccGroupAddress = EMAIL_TEMPLATES.follow_up_1.cc;

    // Simply add the group handle to CC
    await emailUtility.addCC(ccGroupAddress);

    // Get the email content
    const emailContentToAdd = emailUtility.getEmailContent("follow_up_1", {
      name: name,
      quote_reference: reference,
      quote_items: items,
    });

    // Use the addBody method to prepend the content
    await emailUtility.addBody(emailContentToAdd);
  } catch (error) {
    // Use the helper function to display the error in the task pane
    emailUtility.displayErrorInTaskpane(`Error in followUp1: ${error.message}`);
  }
}
