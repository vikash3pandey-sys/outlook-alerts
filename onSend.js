Office.onReady();

var dialog;

function checkExternalRecipients(event) {
    var item = Office.context.mailbox.item;

// 1. Load Dynamic Trusted Domains
    var savedDomains = Office.context.roamingSettings.get("TrustedDomains");
    var defaultDomains = [
        "paytmpayments.com", "paytmmoney.net", "paytminsurance.co.in", 
        "paytmservices.com", "paytm.com", "powerplay.today", "inapaq.com", 
        "paytmmall.io", "cloud.paytm.com", "firstgames.id", "ticketnew.com", 
        "paytmmall.com", "paytmplay.com", "mobiquest.com", "fellowinfotech.com", 
        "paytminsuretech.com", "alpineinfocom.com", "firstgames.in", "first.games", 
        "paytmfoundation.org", "paytmforbusiness.in", "ps.paytm.com", 
        "paytmcloud.in", "paytm.insure", "mypaytm.com", "paytm.business", 
        "fincollect.in", "creditmate.in", "gamepind.com", "insider.paytm.com", 
        "pmltp.com", "finmate.tech", "cdo.paytm.com", "paytmoffers.in", 
        "paytmmloyal.com", "ocltp.com", "paytm.ca", "quarkinfocom.com", 
        "pibpltp.com", "paytmfirstgames.com", "paytmgic.com", "paytmwholesale.com", 
        "paytmlabs.com", "info.paytmfirstgames.com", "acumengame.com", 
        "robustinfo.com", "one97.sg"
    ];
    var trustedDomains = savedDomains ? JSON.parse(savedDomains) : defaultDomains;

        // 2. Check logic
        for (var i = 0; i < recipients.length; i++) {
            var email = recipients[i].emailAddress.toLowerCase();
            var isSafe = false;
            for (var j = 0; j < trustedDomains.length; j++) {
                if (email.indexOf("@" + trustedDomains[j]) > -1) {
                    isSafe = true; break;
                }
            }
            if (!isSafe) { 
                externalEmails.push(email); // Save the bad email
            }
        }

        if (externalEmails.length === 0) {
            // Internal Only -> Send Silently
            event.completed({ allowEvent: true });
        } else {
            // External Found -> Prepare URL with the list of external emails
            var encodedEmails = encodeURIComponent(externalEmails.join(","));
            var url = "https://vikash3pandey-sys.github.io/outlook-alerts/warning.html?ext=" + encodedEmails;

            Office.context.ui.displayDialogAsync(url, { height: 50, width: 40, displayInIframe: true },
                function (asyncResult) {
                    if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                        event.completed({ allowEvent: false, errorMessage: "Security Check Failed." });
                    } else {
                        dialog = asyncResult.value;
                        
                        dialog.addEventHandler(Office.EventType.DialogMessageReceived, function(arg) {
                            dialog.close(); 
                            
                            if (arg.message === "allow") {
                                // SEND TO EVERYONE
                                event.completed({ allowEvent: true });
                            } else if (arg.message === "cancel") {
                                // CANCEL SEND
                                event.completed({ allowEvent: false });
                            } else if (arg.message === "remove_and_send") {
                                // SCRUB EXTERNALS AND SEND
                                removeExternalsAndSend(item, trustedDomains, event);
                            }
                        });
                    }
                }
            );
        }
    });
}

// Helper 1: Removes externals from "To" and "CC" and sends the email
function removeExternalsAndSend(item, trustedDomains, event) {
    item.to.getAsync(function(resTo) {
        var safeTo = filterSafe(resTo.value, trustedDomains);
        item.to.setAsync(safeTo, function() { 
            item.cc.getAsync(function(resCc) {
                var safeCc = filterSafe(resCc.value, trustedDomains);
                item.cc.setAsync(safeCc, function() {
                    event.completed({ allowEvent: true });
                });
            });
        });
    });
}

// Helper 2: Keeps only trusted emails
function filterSafe(recipients, trustedDomains) {
    if (!recipients) return [];
    var safeList = [];
    
    for (var i = 0; i < recipients.length; i++) {
        var email = recipients[i].emailAddress.toLowerCase();
        var isSafe = false;
        for (var j = 0; j < trustedDomains.length; j++) {
            if (email.indexOf("@" + trustedDomains[j]) > -1) {
                isSafe = true; break;
            }
        }
        if (isSafe) { safeList.push(recipients[i]); }
    }
    return safeList;
}
