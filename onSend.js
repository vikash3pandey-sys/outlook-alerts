Office.onReady();

function checkExternalRecipients(event) {
    var item = Office.context.mailbox.item;

    // 1. Guaranteed list of Trusted Domains
    var trustedDomains = [
        "paytmpayments.com", "paytmmoney.com", "paytminsurance.co.in", "paytmservices.com", "paytm.com", 
        "powerplay.today", "inapaq.com", "paytmmall.io", "cloud.paytm.com", "firstgames.id", "ticketnew.com", 
        "paytmmall.com", "paytmplay.com", "mobiquest.com", "fellowinfotech.com", "paytminsuretech.com", 
        "alpineinfocom.com", "firstgames.in", "first.games", "paytmfoundation.org", "paytmforbusiness.in", 
        "ps.paytm.com", "paytmcloud.in", "paytm.insure", "mypaytm.com", "paytm.business", "fincollect.in", 
        "creditmate.in", "gamepind.com", "insider.paytm.com", "pmltp.com", "finmate.tech", "cdo.paytm.com", 
        "paytmoffers.in", "paytmmloyal.com", "ocltp.com", "paytm.ca", "quarkinfocom.com", "pibpltp.com", 
        "paytmfirstgames.com", "paytmgic.com", "paytmwholesale.com", "paytmlabs.com", "info.paytmfirstgames.com", 
        "acumengame.com", "robustinfo.com", "one97.sg"
    ];

    item.to.getAsync(function(result) {
        if (result.status !== Office.AsyncResultStatus.Succeeded) {
            event.completed({ allowEvent: true });
            return;
        }

        var recipients = result.value;
        var externalEmails = []; 

        // 2. Scan Logic
        for (var i = 0; i < recipients.length; i++) {
            var email = recipients[i].emailAddress.toLowerCase();
            var isSafe = false;
            
            for (var j = 0; j < trustedDomains.length; j++) {
                // Checks if it ends with exactly "@domain.com" or a subdomain ".domain.com"
                if (email.endsWith("@" + trustedDomains[j]) || email.endsWith("." + trustedDomains[j])) {
                    isSafe = true; 
                    break;
                }
            }
            if (!isSafe) { 
                externalEmails.push(email); 
            }
        }

        // 3. Decision Logic (Native Outlook Banner)
        if (externalEmails.length === 0) {
            // Internal Only -> Send Silently
            event.completed({ allowEvent: true });
        } else {
            // External Found -> Use Native Outlook Banner instead of a Pop-up window
            item.loadCustomPropertiesAsync(function (propResult) {
                var props = propResult.value;
                var warningStatus = props.get("WarningBypass"); 

                if (warningStatus === "yes") {
                    // USER CLICKED SEND A SECOND TIME -> Let the email go through
                    props.remove("WarningBypass");
                    props.saveAsync(function() {
                         event.completed({ allowEvent: true });
                    });
                } else {
                    // FIRST CLICK -> Stop the send and show the native text banner at the top
                    props.set("WarningBypass", "yes");
                    props.saveAsync(function() {
                        
                        // Set the text for the native banner
                        var recWord = externalEmails.length === 1 ? "recipient" : "recipients";
                        var bannerText = "Confidentiality Warning: You are sending to " + externalEmails.length + " external " + recWord + ". Click Send again to confirm.";
                        
                        event.completed({ 
                            allowEvent: false, 
                            errorMessage: bannerText
                        });
                    });
                }
            });
        }
    });
}
