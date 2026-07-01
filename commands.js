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
    const currentUserEmail = Office.context.mailbox.userProfile.emailAddress;
    
    let newToRecipients = [];
    let newCcRecipients = [];

    // 1. Calculate your custom recipient layouts directly from the stable read-pane item metadata
    if (item.from) {
      newToRecipients.push({ displayName: item.from.displayName, emailAddress: item.from.emailAddress });
    }

    if (item.to) {
      item.to.forEach((rcp) => {
        if (rcp.emailAddress.toLowerCase() !== currentUserEmail.toLowerCase()) {
          newCcRecipients.push({ displayName: rcp.displayName, emailAddress: rcp.emailAddress });
        }
      });
    }

    if (item.cc) {
      item.cc.forEach((rcp) => {
        if (rcp.emailAddress.toLowerCase() !== currentUserEmail.toLowerCase()) {
          newCcRecipients.push({ displayName: rcp.displayName, emailAddress: rcp.emailAddress });
        }
      });
    }

    console.log("[SpecialReplyTools] Executing native Reply All to preserve email history layout...");

    // 2. Launch using native ReplyAll form (forces thread content, attachments, and style blocks to load)
    item.displayReplyAllFormAsync(
      { htmlBody: "<br><br>" },
      function (asyncResult) {
        console.log("[SpecialReplyTools] Native reply window generated. Status:", asyncResult.status);
        
        if (asyncResult.status === Office.AsyncResultStatus.Failed) {
          console.error("[SpecialReplyTools] Native window invocation failed:", asyncResult.error);
          return;
        }

        console.log("[SpecialReplyTools] Injecting custom shuffled recipients into active window state...");

        // 3. Queue a tiny micro-task to execute right after window spin-up to scrub default recipients
        setTimeout(() => {
          try {
            const composeItem = Office.context.mailbox.item;
            
            if (composeItem && composeItem.to && composeItem.cc) {
              // Enforce your exact target constraints directly into the newly visible form fields
              composeItem.to.setAsync(newToRecipients);
              composeItem.cc.setAsync(newCcRecipients);
              console.log("[SpecialReplyTools] Recipients successfully scrubbed and updated.");
            } else {
              console.warn("[SpecialReplyTools] Active window compose context unavailable for field override.");
            }
          } catch (overrideError) {
            console.error("[SpecialReplyTools] Error rewriting fields inside the active draft container:", overrideError);
          }
        }, 300); // 300ms is standard for OWA window assembly stabilization
      }
    );

  } catch (runtimeError) {
    console.error("[SpecialReplyTools] Critical tracking block crash:", runtimeError);
  }
}