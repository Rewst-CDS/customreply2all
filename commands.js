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
  console.log("[SpecialReplyTools] Context item tracking:", item);

  if (!item) {
    console.error("[SpecialReplyTools] Failure: Office.context.mailbox.item is unavailable.");
    return;
  }

  // 1. Fetch the original email body as HTML asynchronously
  item.body.getAsync(Office.CoercionType.Html, function (bodyResult) {
    if (bodyResult.status === Office.AsyncResultStatus.Failed) {
      console.error("[SpecialReplyTools] Failed to retrieve original email body:", bodyResult.error);
      return;
    }

    try {
      const originalBody = bodyResult.value;
      const currentUserEmail = Office.context.mailbox.userProfile.emailAddress;
      console.log("[SpecialReplyTools] Current user profile: ", currentUserEmail);

      let newToRecipients = [];
      let newCcRecipients = [];

      // 2. Map original sender to TO
      if (item.from) {
        console.log("[SpecialReplyTools] Processing From field:", item.from.emailAddress);
        newToRecipients.push({ displayName: item.from.displayName, emailAddress: item.from.emailAddress });
      }

      // 3. Map original TOs to CC (excluding self)
      if (item.to) {
        item.to.forEach((rcp) => {
          if (rcp.emailAddress.toLowerCase() !== currentUserEmail.toLowerCase()) {
            newCcRecipients.push({ displayName: rcp.displayName, emailAddress: rcp.emailAddress });
          }
        });
      }

      // 4. Keep original CCs in CC (excluding self)
      if (item.cc) {
        item.cc.forEach((rcp) => {
          if (rcp.emailAddress.toLowerCase() !== currentUserEmail.toLowerCase()) {
            newCcRecipients.push({ displayName: rcp.displayName, emailAddress: rcp.emailAddress });
          }
        });
      }

      const replySubject = item.subject.toLowerCase().startsWith("re:") ? item.subject : "RE: " + item.subject;
      console.log("[SpecialReplyTools] Generated reply subject: ", replySubject);

      // 5. Construct a clean Outlook-style header block for the history thread
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
        ${originalBody}
      `;

      console.log("[SpecialReplyTools] Launching message draft window with history...");
      
      // 6. Spawn the custom reply window
      Office.context.mailbox.displayNewMessageFormAsync(
        {
          toRecipients: newToRecipients,
          ccRecipients: newCcRecipients,
          subject: replySubject,
          htmlBody: combinedHtmlBody
        },
        function (asyncResult) {
          console.log("[SpecialReplyTools] Form callback status:", asyncResult.status);
          if (asyncResult.status === Office.AsyncResultStatus.Failed) {
            console.error("[SpecialReplyTools] Draft initialization failed:", asyncResult.error);
          } else {
            console.log("[SpecialReplyTools] Success: New draft interface generated.");
          }
        }
      );

    } catch (runtimeError) {
      console.error("[SpecialReplyTools] Critical tracking block crash:", runtimeError);
    }
  });
}