console.log("[SpecialReplyTools] Script file loaded and executing initial blocks.");

// Register and execute immediately on initialization
Office.onReady((info) => {
  console.log("[SpecialReplyTools] Office.onReady fired. Host:", info.host, " Platform:", info.platform);
  
  if (info.host === Office.HostType.Outlook) {
    console.log("[SpecialReplyTools] Host confirmed as Outlook. Forcing immediate function execution...");
    
    // Trigger the code right away since the button click initializes the runtime environment
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

  try {
    const currentUserEmail = Office.context.mailbox.userProfile.emailAddress;
    console.log("[SpecialReplyTools] Current user profile: ", currentUserEmail);

    let newToRecipients = [];
    let newCcRecipients = [];

    // 1. Map original sender to TO
    if (item.from) {
      console.log("[SpecialReplyTools] Processing From field:", item.from.emailAddress);
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
    console.log("[SpecialReplyTools] Generated reply subject: ", replySubject);

    console.log("[SpecialReplyTools] Launching message draft window...");
    Office.context.mailbox.displayNewMessageFormAsync(
      {
        toRecipients: newToRecipients,
        ccRecipients: newCcRecipients,
        subject: replySubject,
        htmlBody: "<br><br>"
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
}