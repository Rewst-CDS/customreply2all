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

  try {
    const currentUserEmail = Office.context.mailbox.userProfile.emailAddress.toLowerCase();
    
    // We will build a list of additional recipients to inject directly into the CC line
    let extraCcRecipients = [];

    // Map original TOs to CC (excluding self)
    if (item.to) {
      item.to.forEach((rcp) => {
        if (rcp.emailAddress.toLowerCase() !== currentUserEmail) {
          extraCcRecipients.push({ 
            displayName: rcp.displayName, 
            emailAddress: rcp.emailAddress, 
            type: Office.MailboxEnums.RecipientType.Cc 
          });
        }
      });
    }

    // Keep original CCs in CC (excluding self)
    if (item.cc) {
      item.cc.forEach((rcp) => {
        if (rcp.emailAddress.toLowerCase() !== currentUserEmail) {
          extraCcRecipients.push({ 
            displayName: rcp.displayName, 
            emailAddress: rcp.emailAddress, 
            type: Office.MailboxEnums.RecipientType.Cc 
          });
        }
      });
    }

    console.log("[SpecialReplyTools] Launching native Reply All form with pre-computed extra recipients...");

    // Fire and forget: Outlook native engine handles the heavy lifting, no setTimeout required
    item.displayReplyAllFormAsync(
      { 
        htmlBody: "<br><br>",
        extraRecipients: extraCcRecipients
      },
      function (asyncResult) {
        console.log("[SpecialReplyTools] Native window generation status:", asyncResult.status);
        if (asyncResult.status === Office.AsyncResultStatus.Failed) {
          console.error("[SpecialReplyTools] Native window invocation failed:", asyncResult.error);
        } else {
          console.log("[SpecialReplyTools] Form generated successfully. Relinquishing runtime control.");
        }
      }
    );

  } catch (runtimeError) {
    console.error("[SpecialReplyTools] Critical tracking block crash:", runtimeError);
  }
}