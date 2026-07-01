Office.onReady(function (info) {
    if (info.host === Office.HostType.Outlook) {
        console.log("[SpecialReplyTools] Office.onReady fired. Host: Outlook");
    }
});

console.log("[SpecialReplyTools] Script file loaded.");

function specialReplyToAll(event) {
    console.log("[SpecialReplyTools] ---> specialReplyToAll triggered by button! <---");

    const item = Office.context.mailbox.item;
    const currentUserEmail = Office.context.mailbox.userProfile.emailAddress.toLowerCase();
    const originalSender = item.from ? item.from.emailAddress : "";
    
    let toRecipients = [];
    let ccRecipients = [];

    if (originalSender && originalSender.toLowerCase() !== currentUserEmail) {
        toRecipients.push(item.from);
    }

    if (item.to) {
        item.to.forEach(function (recp) {
            if (recp.emailAddress.toLowerCase() !== currentUserEmail && recp.emailAddress.toLowerCase() !== originalSender.toLowerCase()) {
                ccRecipients.push(recp);
            }
        });
    }

    if (item.cc) {
        item.cc.forEach(function (recp) {
            if (recp.emailAddress.toLowerCase() !== currentUserEmail && recp.emailAddress.toLowerCase() !== originalSender.toLowerCase()) {
                ccRecipients.push(recp);
            }
        });
    }

    item.body.getAsync(Office.CoercionType.Html, function (asyncResult) {
        if (asyncResult.status === Office.AsyncResultStatus.Failed) {
            console.error("[SpecialReplyTools] Failed to get body:", asyncResult.error.message);
            if (event) event.completed();
            return;
        }

        const originalBody = asyncResult.value;
        const replyHeader = "<br><br><hr><b>From:</b> " + (item.from ? item.from.displayName : originalSender) + 
                            "<br><b>Date:</b> " + item.dateTimeCreated.toLocaleString() + 
                            "<br><b>Subject:</b> " + item.subject + "<br><br>";

        console.log("[SpecialReplyTools] Launching native Reply form with shuffled recipients...");

        Office.context.mailbox.displayNewMessageFormAsync({
            toRecipients: toRecipients,
            ccRecipients: ccRecipients,
            subject: (item.subject.toUpperCase().startsWith("RE:") ? item.subject : "RE: " + item.subject),
            htmlBody: replyHeader + originalBody
        }, function(formResult) {
            if (formResult.status === Office.AsyncResultStatus.Failed) {
                console.error("[SpecialReplyTools] Form failed:", formResult.error.message);
            } else {
                console.log("[SpecialReplyTools] Form generated successfully.");
            }
            
            if (event) event.completed();
        });
    });
}

Office.actions.associate("specialReplyToAll", specialReplyToAll);