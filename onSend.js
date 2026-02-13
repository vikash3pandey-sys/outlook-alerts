/*
 * Check Recipients On-Send
 */

Office.onReady();

function allowedToSend(event) {
    // 1. Get recipients
    Office.context.mailbox.item.to.getAsync(function(result) {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
            var recipients = result.value;
            var trustedDomains = ["paytm.com"]; 
            var externalFound = false;

            // 2. Simple Check
            for (var i = 0; i < recipients.length; i++) {
                var email = recipients[i].emailAddress.toLowerCase();
                var isSafe = false;
                
                for (var j = 0; j < trustedDomains.length; j++) {
                    if (email.indexOf("@" + trustedDomains[j]) > -1) {
                        isSafe = true;
                        break;
                    }
                }
                if (!isSafe) { externalFound = true; }
            }

            // 3. Block or Allow
            if (externalFound) {
                // BLOCK
                event.completed({ allowEvent: false, errorMessage: "⚠️ External recipients detected. Please check the 'To' list." });
            } else {
                // ALLOW
                event.completed({ allowEvent: true });
            }
        } else {
            // Fail safe: Allow if we can't read recipients
            event.completed({ allowEvent: true });
        }
    });
}
