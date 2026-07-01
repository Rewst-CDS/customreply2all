console.log("[SpecialReplyTools] Script file loaded and executing initial blocks.");

Office.onReady((info) => {
  console.log("[SpecialReplyTools] Office.onReady fired. Host:", info.host, " Platform:", info.platform);
  
  if (info.host === Office.HostType.Outlook) {
    console.log("[SpecialReplyTools] Host confirmed as Outlook. Forcing immediate function execution...");
    runSpecialReplyWorkflow();
  } else {
    console.warn("[SpecialReplyTools] Host type mismatch. Expected Outlook.");
  }
});

function runSpecialReplyWorkflow() {
  console.log("[SpecialReplyTools] ---> runSpecialReplyWorkflow triggered! <---");

  const item = Office.context.mailbox.item;
  if (!item) {
    console.error("[SpecialReplyTools] Failure: Office.context.mailbox.item is unavailable.");
    return;
  }

  console.log("[SpecialReplyTools] Requesting body content asynchronously...");
  
  // Explicitly requesting HTML format with an explicit callback handler to avoid drops in OWA
  item.body.getAsync(Office.CoercionType.Html, handleBodyAndCreateDraft);
}

function handleBodyAndCreateDraft(asyncResult) {
  console.log("[SpecialReplyTools] body.getAsync callback triggered. Status:", asyncResult.status);

  if (asyncResult.status === Office.AsyncResultStatus.Failed) {
    console.error("[SpecialReplyTools] Failed to retrieve message body content:", asyncResult.error);
    return;
  }

  const originalThreadBody = asyncResult.value;
  const item = Office.context.mailbox.item;

  try {
    const currentUserEmail = Office.context.mailbox.userProfile.emailAddress;
    let newToRecipients = [];
    let newCcRecipients = [];

    // 1. Map original sender to TO
    if (item.from) {
      newToRecipients.push({ displayName: item.from.displayName, emailAddress: item.from.emailAddress });
    }

    // 2. Map original TOs to CC (excluding self)
    if (item.to) {
      item.to.forEach((rcp) => {
        if (rcp.emailAddress.toLowerCase() !== currentUserEmail.toLowerCase()) {
          newCcRecipients.push({ displayName: rcp.displayName, emailAddress: rcp.emailAddress });
        }
      });
    }

    // 3. Keep original CCs in CC (excluding self)
    if (item.cc) {
      item.cc.forEach((rcp) => {
        if (rcp.emailAddress.toLowerCase() !== currentUserEmail.toLowerCase()) {
          newCcRecipients.push({ displayName: rcp.displayName, emailAddress: rcp.emailAddress });
        }
      });
    }

    const replySubject = item.subject.toLowerCase().startsWith("re:") ? item.subject : "RE: " + item.subject;

    // 4. Structure the historical header block
    const senderName = item.from ? item.from.displayName : "";
    const senderEmail = item.from ? item.from.emailAddress : "";
    
    const combinedHtmlBody = `
      <br><br>
      <div style="border:none;border-top:solid #B5B5B5 1.0pt;padding:3.0pt 0in 0in 0in;font-family:Calibri,sans-serif;font-size:11.0pt;">
        <b>From:</b> ${senderName} &lt;${senderEmail}&gt;<br>
        <b>Sent:</b> ${item.dateTimeCreated ? item.dateTimeCreated.toLocaleString() : ""}<br>
        <b>To:</b> ${item.to ? item.to.map(r => r.displayName || r.emailAddress).join("; ") : ""}<br>
        <b>Cc:</b> ${item.cc ? item.cc.map(r => r.displayName || r.emailAddress).join("; ") : ""}<br>
        <b>Subject:</b> ${item.subject}<br>
      </div>
      <br>
      ${originalThreadBody}
    `;

    console.log("[SpecialReplyTools] Launching message draft window...");
    
    Office.context.mailbox.displayNewMessageFormAsync(
      {
        toRecipients: newToRecipients,
        ccRecipients: newCcRecipients,
        subject: replySubject,
        htmlBody: combinedHtmlBody
      },
      function (formResult) {
        console.log("[SpecialReplyTools] Form callback status:", formResult.status);
        if (formResult.status === Office.AsyncResultStatus.Failed) {
          console.error("[SpecialReplyTools] Draft initialization failed:", formResult.error);
        } else {
          console.log("[SpecialReplyTools] Success: New draft window generated.");
        }
      }
    );

  } catch (runtimeError) {
    console.error("[SpecialReplyTools] Critical tracking block crash:", runtimeError);
  }
}