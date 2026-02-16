Office.onReady();

function checkExternalRecipients(event) {
    var item = Office.context.mailbox.item;

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

        for (var i = 0; i < recipients.length; i++) {
            var email = recipients[i].emailAddress.toLowerCase();
            var isSafe = false;
            
            for (var j = 0; j < trustedDomains.length; j++) {
                if (email.endsWith("@" + trustedDomains[j]) || email.endsWith("." + trustedDomains[j])) {
                    isSafe = true; break;
                }
            }
            if (!isSafe) { 
                externalEmails.push(email); 
            }
        }

        if (externalEmails.length === 0) {
            // Safe -> Send immediately
            event.completed({ allowEvent: true });
        } else {
            // External Found -> Check if user is clicking "Send" for the second time
            item.loadCustomPropertiesAsync(function (propResult) {
                var props = propResult.value;
                var warningStatus = props.get("WarningBypass"); 

                if (warningStatus === "yes") {
                    // USER CLICKED SEND TWICE -> "Send Anyway"
                    props.remove("WarningBypass");
                    props.saveAsync(function() {
                         event.completed({ allowEvent: true });
                    });
                } else {
                    // FIRST CLICK -> Show native warning banner inside the email
                    props.set("WarningBypass", "yes");
                    props.saveAsync(function() {
                        var recWord = externalEmails.length === 1 ? "recipient" : "recipients";
                        var bannerText = "Confidentiality Warning: You are sending to " + externalEmails.length + " external " + recWord + " outside Paytm. Click 'Send' again to send anyway.";
                        
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
