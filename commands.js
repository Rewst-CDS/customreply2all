Office.onReady();

function onMessageComposeHandler(event) {
    console.log("[Add-in Event] Native draft detected. Evaluating launch trigger context...");
    var item = Office.context.mailbox.item;

    // Verify this window was specifically created via a 'Reply All' action
    item.getComposeTypeAsync(function (typeResult) {
        if (typeResult.status === Office.AsyncResultStatus.Failed) {
            console.error("[Add-in Event] Failed to identify compose type: " + typeResult.error.message);
            event.completed();
            return;
        }

        var composeType = typeResult.value.composeType;
        console.log("[Add-in Event] Detected Compose Action: " + composeType);

        if (composeType !== "replyAll") {
            console.log("[Add-in Event] Not a Reply All request. Exiting intercept cleanly.");
            event.completed();
            return;
        }

        console.log("[Add-in Event] Target met. Fetching current pre-populated recipient list...");
        
        // 1. Grab all recipients currently in the native 'To' field
        item.to.getAsync(function (toResult) {
            if (toResult.status === Office.AsyncResultStatus.Failed) {
                console.error("[Add-in Event] Failed to read current TO line: " + toResult.error.message);
                event.completed();
                return;
            }

            var nativeToRecipients = toResult.value || [];
            console.log("[Add-in Event] Native pre-populated TO array count: " + nativeToRecipients.length);

            if (nativeToRecipients.length === 0) {
                console.warn("[Add-in Event] The TO line is empty. Nothing to rearrange.");
                event.completed();
                return;
            }

            // Apply your assumption: the first entry is our target 'To' sender
            var targetSender = nativeToRecipients[0];
            console.log("[Add-in Event] ASSUMPTION: Target sender identified as: " + targetSender.emailAddress);

            // 2. Fetch all recipients currently in the native 'Cc' field
            item.cc.getAsync(function (ccResult) {
                var nativeCcRecipients = [];
                if (ccResult.status === Office.AsyncResultStatus.Succeeded) {
                    nativeCcRecipients = ccResult.value || [];
                }
                console.log("[Add-in Event] Native pre-populated CC array count: " + nativeCcRecipients.length);

                // 3. Build a map of secondary recipients (everyone else from TO + everyone from CC)
                var secondaryRecipientsMap = {};

                // Add the leftover elements from the original TO array
                for (var i = 1; i < nativeToRecipients.length; i++) {
                    var rec = nativeToRecipients[i];
                    if (rec.emailAddress) {
                        secondaryRecipientsMap[rec.emailAddress.toLowerCase()] = rec;
                    }
                }

                // Add elements from the original CC array
                nativeCcRecipients.forEach(function (rec) {
                    if (rec.emailAddress) {
                        secondaryRecipientsMap[rec.emailAddress.toLowerCase()] = rec;
                    }
                });

                // Safety check: ensure our designated target sender is completely stripped out of CC
                delete secondaryRecipientsMap[targetSender.emailAddress.toLowerCase()];

                // Flatten map down to a clean recipient object array
                var finalCcPayload = Object.keys(secondaryRecipientsMap).map(function (key) {
                    return {
                        displayName: secondaryRecipientsMap[key].displayName,
                        emailAddress: secondaryRecipientsMap[key].emailAddress
                    };
                });

                console.log("[Add-in Event] Final calculated CC payload array: ", finalCcPayload);

                // 4. Update the live compose window fields using Outlook setters
                console.log("[Add-in Event] Updating compose window. Isolate single recipient on TO...");
                item.to.setAsync([{ displayName: targetSender.displayName, emailAddress: targetSender.emailAddress }], function (setToResult) {
                    if (setToResult.status === Office.AsyncResultStatus.Succeeded) {
                        console.log("[Add-in Event] TO line updated successfully.");
                    } else {
                        console.error("[Add-in Event] Failed updating TO line: " + setToResult.error.message);
                    }

                    console.log("[Add-in Event] Updating compose window. Pushing remaining contacts to CC...");
                    item.cc.setAsync(finalCcPayload, function (setCcResult) {
                        if (setCcResult.status === Office.AsyncResultStatus.Succeeded) {
                            console.log("[Add-in Event] CC line updated successfully.");
                        } else {
                            console.error("[Add-in Event] Failed updating CC line: " + setCcResult.error.message);
                        }

                        // Complete execution loop
                        console.log("[Add-in Event] Interception rearrangement successfully executed.");
                        event.completed();
                    });
                });
            });
        });
    });
}

// Associate function name to manifest declaration
Office.actions.associate("onMessageComposeHandler", onMessageComposeHandler);
console.log("[Add-in Event] Background Event registration mapped to 'onMessageComposeHandler'");