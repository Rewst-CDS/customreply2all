console.log("[SpecialReplyTools] Shuffling module loaded.");

Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) {
    // Ensure we are executing inside a COMPOSE item context
    if (Office.context.mailbox.item.to) {
      executeRecipientShuffle();
    }
  }
});

function executeRecipientShuffle() {
  console.log("[SpecialReplyTools] ---> Shuffling active reply-all targets... <---");

  const currentDraft = Office.context.mailbox.item;
  const currentUserEmail = Office.context.mailbox.userProfile.emailAddress.toLowerCase();

  // 1. Fetch the recipients Outlook natively added to the TO line
  currentDraft.to.getAsync(function (asyncResult) {
    if (asyncResult.status === Office.AsyncResultStatus.Failed) {
      console.error("[SpecialReplyTools] Couldn't read current TO fields:", asyncResult.error);
      return;
    }

    const nativeToRecipients = asyncResult.value;
    let newToLine = [];
    let migratedCcLine = [];

    // 2. Separate out the original sender vs everyone else
    nativeToRecipients.forEach((rcp, index) => {
      const email = rcp.emailAddress.toLowerCase();
      
      // Keep yourself off all lines completely
      if (email === currentUserEmail) return;

      // Treat the very first recipient as the original sender (to stay in TO)
      if (index === 0) {
        newToLine.push({ displayName: rcp.displayName, emailAddress: rcp.emailAddress });
      } else {
        // Shunt all secondary TO recipients into the CC buffer
        migratedCcLine.push({ displayName: rcp.displayName, emailAddress: rcp.emailAddress });
      }
    });

    // 3. Overwrite the draft's TO line to contain ONLY the primary sender
    currentDraft.to.setAsync(newToLine, function (toSetResult) {
      if (toSetResult.status === Office.AsyncResultStatus.Succeeded) {
        console.log("[SpecialReplyTools] TO field trimmed down to primary sender.");
        
        // 4. Inject the rest of the team into the CC line safely
        if (migratedCcLine.length > 0) {
          currentDraft.cc.addAsync(migratedCcLine, function (ccAddResult) {
            if (ccAddResult.status === Office.AsyncResultStatus.Succeeded) {
              console.log("[SpecialReplyTools] Successfully shuffled remaining targets to CC.");
            } else {
              console.error("[SpecialReplyTools] CC injection failed:", ccAddResult.error);
            }
          });
        }
      } else {
        console.error("[SpecialReplyTools] Error rewriting TO row:", toSetResult.error);
      }
    });
  });
}