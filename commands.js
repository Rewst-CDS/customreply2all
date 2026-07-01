Office.onReady(function (info) {
    console.log("[Add-in] Office.js initialization complete. Host: " + info.host + ", Platform: " + info.platform);
});

function customReplyAll(event) {
    console.log("[Add-in] customReplyAll button clicked. Execution started.");
    
    var item = Office.context.mailbox.item;
    if (!item) {
        console.error("[Add-in] Critical Error: Unable to access the current mailbox item.");
        if (event) event.completed();
        return;
    }

    var currentUser = Office.context.mailbox.userProfile.emailAddress.toLowerCase();
    console.log("[Add-in] Detected Current User email: " + currentUser);

    var toRecipients = [];
    var ccRecipientsMap = {};

    // 1. Set the original sender as the sole 'To' recipient
    if (item.from && item.from.emailAddress) {
        console.log("[Add-in] Step 1: Found original sender -> " + item.from.emailAddress);
        toRecipients.push(item.from.emailAddress);
        console.log("[Add-in] Current 'To' target array: ", toRecipients);
    } else {
        console.warn("[Add-in] Step 1 Warning: Original sender data is missing.");
    }

    // 2. Collect all other recipients from the original 'To' field
    if (item.to && item.to.length > 0) {
        console.log("[Add-in] Step 2: Extracting original 'To' recipients. Count: " + item.to.length);
        item.to.forEach(function (recipient, index) {
            if (recipient.emailAddress) {
                console.log("[Add-in] Processing original 'To' [" + index + "]: " + recipient.emailAddress);
                ccRecipientsMap[recipient.emailAddress.toLowerCase()] = recipient.emailAddress;
            }
        });
    } else {
        console.log("[Add-in] Step 2: Original 'To' field is empty.");
    }

    // 3. Collect all recipients from the original 'Cc' field
    if (item.cc && item.cc.length > 0) {
        console.log("[Add-in] Step 3: Extracting original 'Cc' recipients. Count: " + item.cc.length);
        item.cc.forEach(function (recipient, index) {
            if (recipient.emailAddress) {
                console.log("[Add-in] Processing original 'Cc' [" + index + "]: " + recipient.emailAddress);
                ccRecipientsMap[recipient.emailAddress.toLowerCase()] = recipient.emailAddress;
            }
        });
    } else {
        console.log("[Add-in] Step 3: Original 'Cc' field is empty.");
    }

    // 4. Clean up the CC list (remove the original sender and the current user replying)
    console.log("[Add-in] Step 4: Filtering out unwanted users from the CC map.");
    
    if (item.from && item.from.emailAddress) {
        var senderKey = item.from.emailAddress.toLowerCase();
        if (ccRecipientsMap[senderKey]) {
            console.log("[Add-in] Removing original sender (" + item.from.emailAddress + ") from CC map to prevent duplication.");
            delete ccRecipientsMap[senderKey];
        }
    }
    
    if (ccRecipientsMap[currentUser]) {
        console.log("[Add-in] Removing current user (" + currentUser + ") from CC map to avoid replying to self.");
        delete ccRecipientsMap[currentUser];
    }

    // Convert the unique CC map back into an array
    var ccRecipients = Object.keys(ccRecipientsMap).map(function (key) {
        return ccRecipientsMap[key];
    });
    console.log("[Add-in] Final mapped CC recipients array: ", ccRecipients);

    // 5. Format the subject line
    var subject = item.subject || "";
    console.log("[Add-in] Step 5: Analyzing original subject line -> \"" + subject + "\"");
    if (!subject.toLowerCase().startsWith("re:")) {
        subject = "Re: " + subject;
        console.log("[Add-in] Added 'Re:' prefix to subject line.");
    }
    console.log("[Add-in] Final Subject string: \"" + subject + "\"");

    // 6. Launch the new reply-all email window with the rearranged fields
    console.log("[Add-in] Step 6: Dispatching displayNewMessageFormAsync call to Outlook...");
    console.log("[Add-in] Final payload mapping - TO: ", toRecipients, " | CC: ", ccRecipients);

    Office.context.mailbox.displayNewMessageFormAsync({
        toRecipients: toRecipients,
        ccRecipients: ccRecipients,
        subject: subject,
        htmlBody: "<br/><br/>"
    }, function (asyncResult) {
        if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
            console.log("[Add-in] Success: Reply-all window generated successfully.");
        } else {
            console.error("[Add-in] Failure: displayNewMessageFormAsync failed with error: " + asyncResult.error.message);
        }

        // Check if the event object exists (required for standard UI command buttons)
        if (event) {
            console.log("[Add-in] Notifying Outlook platform that custom operation is complete via event.completed()");
            event.completed();
        } else {
            console.log("[Add-in] Context note: No event context passed to function (likely manual test call).");
        }
    });
}

// Associate the function name with the Action ID defined in your manifest
console.log("[Add-in] Attempting manifest action association for ID: 'customReplyAll'");
Office.actions.associate("customReplyAll", customReplyAll);
console.log("[Add-in] Action association registration parsed.");