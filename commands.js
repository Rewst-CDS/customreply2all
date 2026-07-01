console.log("[SpecialReplyTools] Script file loaded.");

Office.onReady((info) => {
  console.log("[SpecialReplyTools] Office.onReady fired. Host:", info.host);
});

// 1. Function name matches the manifest <FunctionName> exactly
// 2. Accepts the 'event' parameter
function specialReplyToAll(event) {
  console.log("[SpecialReplyTools] ---> specialReplyToAll triggered by button! <---");

  const item = Office.context.mailbox.item;
  if (!item) {
    console.error("[SpecialReplyTools] Failure: item is unavailable.");
    event.completed();
    return;
  }

  const currentUserEmail = Office.context.mailbox.userProfile.emailAddress.toLowerCase();
  let extraCcRecipients = [];

  // Helper function to push recipients to CC
  const processRecipients = (recipients) => {
    if (recipients) {
      recipients.forEach((rcp) => {
        if (rcp.emailAddress.toLowerCase() !== currentUserEmail) {
          extraCcRecipients.push({ 
            displayName: rcp.displayName, 
            emailAddress: rcp.emailAddress, 
            type: Office.MailboxEnums.RecipientType.Cc 
          });
        }
      });
    }
  };

  // Map all original TOs and CCs to the new CC array
  processRecipients(item.to);
  processRecipients(item.cc);

  console.log("[SpecialReplyTools] Launching native Reply form with shuffled recipients...");

  // 3. Use displayReplyFormAsync (Reply to Sender) instead of ReplyAll
  // This natively puts the sender in 'To' and preserves the email body.
  item.displayReplyFormAsync(
    { 
      extraRecipients: extraCcRecipients
    },
    function (asyncResult) {
      if (asyncResult.status === Office.AsyncResultStatus.Failed) {
        console.error("[SpecialReplyTools] Native window invocation failed:", asyncResult.error);
      } else {
        console.log("[SpecialReplyTools] Form generated successfully.");
      }
      
      // 4. Signal to Outlook that the command has finished executing
      event.completed();
    }
  );
}

// 5. Register the function so the Outlook manifest can find and trigger it
Office.actions.associate("specialReplyToAll", specialReplyToAll);