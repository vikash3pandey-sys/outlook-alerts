/*
 * BIGGER WARNING LOGIC (Popup Dialog)
 */

Office.onReady();

var dialog; // Global dialog variable

function checkExternalRecipients(event) {
    var item = Office.context.mailbox.item;

    item.to.getAsync(function(result) {
        if (result.status !== Office.AsyncResultStatus.Succeeded) {
            event.completed({ allowEvent: true });
            return;
        }

        var recipients = result.value;
        var trustedDomains = ["paytm.com", "paytm.in"]; 
        var externalFound = false;

        // 1. Check Logic
        for (var i = 0; i < recipients.length; i++) {
            var email = recipients[i].emailAddress.toLowerCase();
            var isSafe = false;
            for (var j = 0; j < trustedDomains.length; j++) {
                if (email.indexOf("@" + trustedDomains[j]) > -1) {
                    isSafe = true; break;
                }
            }
            if (!isSafe) { externalFound = true; break; }
        }

        // 2. Decision
        if (!externalFound) {
            // Internal -> Send Silently
            event.completed({ allowEvent: true });
        } else {
            // External -> Open BIG POPUP Dialog
            var url = "https://vikash3pandey-sys.github.io/outlook-alerts/warning.html";

            // Open dialog (Width/Height are % of screen)
            Office.context.ui.displayDialogAsync(url, { height: 40, width: 30, displayInIframe: true },
                function (asyncResult) {
                    if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                        // If popup fails to open (blocker), block email safely
                        event.completed({ allowEvent: false, errorMessage: "⚠️ Security Check Failed: Could not open warning popup." });
                    } else {
                        // Dialog Opened Successfully
                        dialog = asyncResult.value;
                        
                        // Listen for button click from the popup
                        dialog.addEventHandler(Office.EventType.DialogMessageReceived, function(arg) {
                            dialog.close(); // Close popup
                            
                            if (arg.message === "allow") {
                                // User clicked "Send Anyway"
                                event.completed({ allowEvent: true });
                            } else {
                                // User clicked "Cancel"
                                event.completed({ allowEvent: false });
                            }
                        });
                    }
                }
            );
        }
    });
}
