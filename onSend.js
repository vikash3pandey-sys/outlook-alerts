/* onSend.js - Warning Only Logic */
Office.onReady();

function allowedToSend(event) {
    var item = Office.context.mailbox.item;
    item.to.getAsync(function(result) {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
            var recipients = result.value;
            // HARDCODED TRUSTED DOMAIN
            var trustedDomains = ["paytm.com"]; 
            var externalFound = false;

            for (var i = 0; i < recipients.length; i++) {
                var email = recipients[i].emailAddress.toLowerCase();
                var isSafe = false;
                for (var j = 0; j < trustedDomains.length; j++) {
                    if (email.indexOf("@" + trustedDomains[j]) > -1) {
                        isSafe = true; break;
                    }
                }
                if (!isSafe) { externalFound = true; }
            }

            if (externalFound) {
                // Check if we already showed the warning
                item.loadCustomPropertiesAsync(function (propResult) {
                    var props = propResult.value;
                    var alreadyWarned = props.get("WarningShown");

                    if (alreadyWarned) {
                        // User clicked Send AGAIN -> ALLOW
                        event.completed({ allowEvent: true });
                    } else {
                        // First time -> BLOCK & WARN
                        props.set("WarningShown", true);
                        props.saveAsync(function(saveResult) {
                            event.completed({ 
                                allowEvent: false, 
                                errorMessage: "⚠️ External recipients found. Click Send again to confirm." 
                            });
                        });
                    }
                });
            } else {
                event.completed({ allowEvent: true });
            }
        } else {
            event.completed({ allowEvent: true });
        }
    });
}
