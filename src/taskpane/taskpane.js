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
    closing:
      "We appreciate your interest in Ship-Around for your procurement needs and we are looking forward to your order confirmation.<br>",
    footnote:
      "If you haven't already, please <a href='https://ship-around.com/register'>register</a> a free buyer account.<br><br>It only takes 5 minutes and will expedite processing future requests.",
  },
  offer_order_existing_user: {
    cc: "group@ship-around.com",
    intro: "Dear {name},<br>",
    body: "Please find attached:<br>",
    attachments: "{attachments}",
    note: "We have already created a pending online order for you, to experience the future of online procurement.<br>",
    closing:
      "Our hybrid sales approach allows you to either buy online at already discounted item prices or proceed with attached quotation.<br><br>Visit the <a href='https://ship-around.com/my-account/orders/'>orders</a> page in your dashboard and checkout to confirm your order, and receive a proforma invoice at the discounted prices.<br><br>You can also click on the order link provided in the attached quote to navigate to the checkout page.<br>",
    footnote:
      "We appreciate your interest in Ship-Around for your procurement needs and we are looking forward to your online or offline order confirmation.",
  },
  purchase_order: {
    cc: "group@ship-around.com",
    intro: "Dear {name},<br>",
    body: "Please find attached:<br>",
    attachments: "{attachments}",
    note: "Looking forward to fulfilling this order.<br>",
    closing: "Thank you for being a valuable supplier of Ship-Around.<br>",
    footnote:
      "If you haven't already, please <a href='https://ship-around.com/register'>register</a> a free seller account.<br><br>It only takes 5 minutes and offers exposure to a worldwide online audience.<br><br>If you need more information setting up your online store, don't hesitate to contact us.",
  },
  proforma_invoice: {
    cc: "group@ship-around.com",
    intro: "Dear {name},<br>",
    body: "Thank you for your order confirmation<br><br>Please find attached:<br>",
    attachments: "{attachments}",
    note: "Thank you for choosing Ship-Around for your procurement needs.<br>",
    closing: "Looking forward to fulfilling your order.<br>",
    footnote:
      "If you haven't already, please <a href='https://ship-around.com/register'>register</a> a free buyer account.<br><br>It only takes 5 minutes and will expedite processing future requests.",
  },
  final_invoice: {
    cc: "group@ship-around.com",
    intro: "Dear {name},<br>",
    body: "Thank you for your order.<br><br>Please find attached:<br>",
    attachments: "{attachments}",
    note: "Thank you for choosing Ship-Around for your procurement needs.<br>",
    closing: "Looking forward to your next order.<br>",
    footnote:
      "If you haven't already, please <a href='https://ship-around.com/register'>register</a> a free buyer account.<br><br>It only takes 5 minutes and will expedite processing future requests.",
  },
  acknowledge: {
    cc: "group@ship-around.com",
    intro: "Dear {name},<br>",
    body: "Thank you for reaching out to us.<br><br>We have logged your inquiry with reference SALE{lead}.<br>",
    note: "Please include the above reference in any future correspondence.<br>",
    closing: "We appreciate your interest and will get back to you shortly.<br>",
    footnote:
      "If you haven't already, please <a href='https://ship-around.com/register'>register</a> a free buyer account.<br><br>It only takes 5 minutes and will expedite processing your request.",
  },
  follow_up: {
    cc: "group@ship-around.com",
    intro: "Dear {name},<br>",
    body: "I am following up regarding our last quotation {quote_reference} for {quote_items}.<br><br>We would like to know if you are still interested in pursuing this order.<br>",
    note: "I have attached said quotation again for your perusal.<br><br>Please let us know of your decision at your earliest convenience and if there is any way we can assist you further.<br>",
    closing: "We appreciate your interest in Ship-Around for your procurement needs.<br>",
    footnote:
      "If you haven't already, please <a href='https://ship-around.com/register'>register</a> a free buyer account.<br><br>It only takes 5 minutes and will expedite processing future requests.",
  },
  buyer_outreach: {
    cc: "info@ship-around.com",
    intro: "Dear {name},<br>",
    body: "Ship-Around extends a warm invitation to immerse yourself in our realm of digitalized procurement.<br><br>We recognise that adapting to new practices requires time and consideration. Hence, we present a hybrid approach — simply send us your inquiries, and we'll diligently source the best deals for you.<br><br>For a swifter, more streamlined procurement experience, delve into our <a href='https://ship-around.com/'><online marketplace</a>. Enjoy the benefits of a transparent system with no monthly subscriptions, hidden fees, or additional charges — only pay the displayed product price.<br>",
    note: "Our buyers reap the advantages of: <ol><li>Comprehensive product comparison</li><li>Detailed product listings</li><li>Efficient product location filtering</li></ol>",
    closing:
      "We'd be delighted to organize a brief call with you to explore how Ship-Around can transform your procurement processes. Are you available for a quick chat this week?",
  },
};

const DOCUMENT_TYPE_MAPPINGS = {
  Q202: "Quotation",
  DN202: "Delivery Note",
  PL202: "Packing List",
  INV202: "Invoice",
  PO202: "Purchase Order",
};

Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("acknowledge").onclick = acknowledgeRFQ;
    document.getElementById("prepare-quote-email").onclick = prepareQuoteEmail;
    document.getElementById("prepare-quote-with-order-email-existing-user").onclick =
      prepareQuoteWithOrderEmailUserExists;
    document.getElementById("prepare-po-email").onclick = preparePOEmail;
    document.getElementById("prepare-proforma-invoice-email").onclick = prepareProformaInvoiceEmail;
    document.getElementById("prepare-paid-invoice-email").onclick = prepareFinalInvoiceEmail;
    document.getElementById("follow-up").onclick = followUp;
    document.getElementById("buyer-outreach").onclick = buyerOutreachInitial;
    document.getElementById("get-message-id").onclick = getMessageID;
  }
});

class EmailUtility {
  constructor(item) {
    this.item = item;
  }

  capitalizeFirstLetter(string) {
    return string.charAt(0).toUpperCase() + string.slice(1);
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
        // Check if the key is "intro" and the replacement key is "name"
        if (key === "intro" && replaceKey === "name") {
          sectionContent = sectionContent.replace(`{${replaceKey}}`, this.capitalizeFirstLetter(value));
        } else sectionContent = sectionContent.replace(`{${replaceKey}}`, value);
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

  getMessageId() {
    return this.item.itemId;
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
    this.okButton.disabled = true;

    // Create an array to keep track of the input elements
    const inputElements = this.inputDivs.map((div) => div.querySelector("input"));

    inputElements.forEach((input) => {
      input.addEventListener("input", () => {
        // Check if all required inputs are filled
        const allInputsFilled = inputElements.every((input) => {
          if (input.hasAttribute("required")) {
            return input.value.trim() !== "";
          }
          return true;
        });

        // Enable or disable the "OK" button based on the condition
        this.okButton.disabled = !allInputsFilled;
      });
    });

    this.okButton.onclick = () => {
      const inputValues = inputElements.map((input) => input.value);

      this.resolve(inputValues);
      this.clearInputs();
      this.hide();
    };

    this.cancelButton.onclick = () => {
      this.reject(new Error("User cancelled the input."));
      this.clearInputs();
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

export async function getMessageID() {
  let emailUtility;
  try {
    // Get a reference to the current compose item
    const item = Office.context.mailbox.item;

    emailUtility = new EmailUtility(item);
    const messageId = await emailUtility.getMessageId();
    // Output the messageId somewhere in the task pane
    document.getElementById("messageIdOutput").textContent = "Message ID: " + messageId;

    // Optionally, use this ID to query more details via Microsoft Graph API
    //getMessageDetails(messageId);
  } catch (error) {
    // Use the helper function to display the error in the taskpane
    emailUtility.displayErrorInTaskpane(`Error in getMessageId: ${error.message}`);
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
    await emailUtility.addSubject(`[SALE${lead.trim()}] `);

    // Define the email address you want to add to CC
    const ccGroupAddress = EMAIL_TEMPLATES.acknowledge.cc;

    // Simply add the group handle to CC
    await emailUtility.addCC(ccGroupAddress);

    // Get the email content
    const emailContentToAdd = emailUtility.getEmailContent("acknowledge", {
      name: name.trim(),
      lead: lead.trim(),
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
      name: name.trim(),
      attachments: attachmentTable,
    });

    // Use the addBody method to prepend the content
    await emailUtility.addBody(emailContentToAdd);
  } catch (error) {
    emailUtility.displayErrorInTaskpane(`Error in prepareQuoteEmail: ${error.message}`);
  }
}

export async function prepareQuoteWithOrderEmailUserExists() {
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
    const ccGroupAddress = EMAIL_TEMPLATES.offer_order_existing_user.cc;

    // Add the group email address to CC if it's not already there
    await emailUtility.addCC(ccGroupAddress);

    // Generate the attachment table
    const attachmentTable = emailUtility.generateAttachmentTable(attachmentNames);

    // Get the email content
    const emailContentToAdd = emailUtility.getEmailContent("offer_order_existing_user", {
      name: name.trim(),
      attachments: attachmentTable,
    });

    // Use the addBody method to prepend the content
    await emailUtility.addBody(emailContentToAdd);
  } catch (error) {
    emailUtility.displayErrorInTaskpane(`Error in prepareQuoteEmail: ${error.message}`);
  }
}

export async function preparePOEmail() {
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
    const pOAttachments = attachmentNames
      .filter((name) => name.startsWith("Purchase Order PO202"))
      .map((name) => name.replace("Purchase Order ", ""));

    // Determine the subject prefix based on the number of Q202 attachments
    let subjectPrefix = "";
    if (pOAttachments.length === 1) {
      subjectPrefix = `[Purchase Order ${pOAttachments[0]}] `;
    } else if (pOAttachments.length > 1) {
      subjectPrefix = `[Purchase Orders ${pOAttachments.join(", ")}] `;
    }

    // Use the addSubject method to prepend the prefix to the current subject
    if (subjectPrefix) {
      await emailUtility.addSubject(subjectPrefix);
    }

    // Define the email address you want to add to CC
    const cCAddress = EMAIL_TEMPLATES.purchase_order.cc;

    // Add the group email address to CC if it's not already there
    await emailUtility.addCC(cCAddress);

    // Generate the attachment table
    const attachmentTable = emailUtility.generateAttachmentTable(attachmentNames);

    // Get the email content
    const emailContentToAdd = emailUtility.getEmailContent("purchase_order", {
      name: name.trim(),
      attachments: attachmentTable,
    });

    // Use the addBody method to prepend the content
    await emailUtility.addBody(emailContentToAdd);
  } catch (error) {
    emailUtility.displayErrorInTaskpane(`Error in preparePOEmail: ${error.message}`);
  }
}

export async function prepareProformaInvoiceEmail() {
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
    const invoiceAttachments = attachmentNames
      .filter((name) => name.startsWith("Invoice INV202"))
      .map((name) => name.replace("Invoice ", ""));

    // Determine the subject prefix based on the number of Q202 attachments
    let subjectPrefix = "";
    if (invoiceAttachments.length === 1) {
      subjectPrefix = `[Proforma Invoice ${invoiceAttachments[0]}] `;
    } else if (invoiceAttachments.length > 1) {
      subjectPrefix = `[Proforma Invoices ${invoiceAttachments.join(", ")}] `;
    }

    // Use the addSubject method to prepend the prefix to the current subject
    if (subjectPrefix) {
      await emailUtility.addSubject(subjectPrefix);
    }

    // Define the email address you want to add to CC
    const cCAddress = EMAIL_TEMPLATES.proforma_invoice.cc;

    // Add the group email address to CC if it's not already there
    await emailUtility.addCC(cCAddress);

    // Generate the attachment table
    const attachmentTable = emailUtility.generateAttachmentTable(attachmentNames);

    // Get the email content
    const emailContentToAdd = emailUtility.getEmailContent("proforma_invoice", {
      name: name.trim(),
      attachments: attachmentTable,
    });

    // Use the addBody method to prepend the content
    await emailUtility.addBody(emailContentToAdd);
  } catch (error) {
    emailUtility.displayErrorInTaskpane(`Error in prepareProformaInvoiceEmail: ${error.message}`);
  }
}

export async function prepareFinalInvoiceEmail() {
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
    const invoiceAttachments = attachmentNames
      .filter((name) => name.startsWith("Invoice INV202"))
      .map((name) => name.replace("Invoice ", ""));

    // Determine the subject prefix based on the number of Q202 attachments
    let subjectPrefix = "";
    if (invoiceAttachments.length === 1) {
      subjectPrefix = `[Paid Invoice ${invoiceAttachments[0]}] `;
    } else if (invoiceAttachments.length > 1) {
      subjectPrefix = `[Paid Invoices ${invoiceAttachments.join(", ")}] `;
    }

    // Use the addSubject method to prepend the prefix to the current subject
    if (subjectPrefix) {
      await emailUtility.addSubject(subjectPrefix);
    }

    // Define the email address you want to add to CC
    const cCAddress = EMAIL_TEMPLATES.final_invoice.cc;

    // Add the group email address to CC if it's not already there
    await emailUtility.addCC(cCAddress);

    // Generate the attachment table
    const attachmentTable = emailUtility.generateAttachmentTable(attachmentNames);

    // Get the email content
    const emailContentToAdd = emailUtility.getEmailContent("final_invoice", {
      name: name.trim(),
      attachments: attachmentTable,
    });

    // Use the addBody method to prepend the content
    await emailUtility.addBody(emailContentToAdd);
  } catch (error) {
    emailUtility.displayErrorInTaskpane(`Error in prepareFinalInvoiceEmail: ${error.message}`);
  }
}

export async function followUp() {
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
    await emailUtility.addSubject(`[SALE${lead.trim()}] `);

    // Define the email address you want to add to CC
    const ccGroupAddress = EMAIL_TEMPLATES.follow_up.cc;

    // Simply add the group handle to CC
    await emailUtility.addCC(ccGroupAddress);

    // Get the email content
    const emailContentToAdd = emailUtility.getEmailContent("follow_up", {
      name: name.trim(),
      quote_reference: reference.toUpperCase().trim(),
      quote_items: items.toLowerCase().trim(),
    });

    // Use the addBody method to prepend the content
    await emailUtility.addBody(emailContentToAdd);
  } catch (error) {
    // Use the helper function to display the error in the task pane
    emailUtility.displayErrorInTaskpane(`Error in followUp: ${error.message}`);
  }
}

export async function buyerOutreachInitial() {
  let emailUtility;
  try {
    // Get a reference to the current compose item
    const item = Office.context.mailbox.item;

    emailUtility = new EmailUtility(item);
    const modal = new Modal("inputModal", ["nameInputDiv"], "modalOk", "modalCancel");

    // Show the modal and wait for the input
    const [name] = await modal.show();

    // Use the modal input to prepend to the subject
    await emailUtility.addSubject(`Ship-Around introduction and meeting request`);

    // Define the email address you want to add to CC
    const ccAddress = EMAIL_TEMPLATES.buyer_outreach.cc;

    // Simply add CC
    await emailUtility.addCC(ccAddress);

    // Get the email content
    const emailContentToAdd = emailUtility.getEmailContent("buyer_outreach", {
      name: name.trim(),
    });

    // Use the addBody method to prepend the content
    await emailUtility.addBody(emailContentToAdd);
  } catch (error) {
    // Use the helper function to display the error in the taskpane
    emailUtility.displayErrorInTaskpane(`Error in acknowledgeRFQ: ${error.message}`);
  }
}
