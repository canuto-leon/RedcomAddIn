'use strict';

ExecuteOrDelayUntilScriptLoaded(initializePage, "sp.js");
var context;
function initializePage() {
    context = SP.ClientContext.get_current();
}

function sendEmail(from, to, body, subject) {
    var site = context.get_site();
    context.load(site);
    context.executeQueryAsync(function (s, a) {
        var siteurl = site.get_url();
        alert(siteurl);
        var urlTemplate = siteurl + "/_api/SP.Utilities.Utility.SendEmail";
        $.ajax({
            contentType: 'application/json',
            url: urlTemplate,
            type: "POST",
            data: JSON.stringify({
                'properties': {
                    '__metadata': { 'type': 'SP.Utilities.EmailProperties' },
                    'From': from,
                    'To': { 'results': [to] },
                    'Body': body,
                    'Subject': subject
                }
            }
          ),
            headers: {
                "Accept": "application/json;odata=verbose",
                "content-type": "application/json;odata=verbose",
                "X-RequestDigest": $("#__REQUESTDIGEST").val()
            },
            success: function (data) {
                alert("Correo enviado");
            },
            error: function (xhr, ajaxOptions, thrownError) {
                alert(xhr.status);
                alert(thrownError);
            }
        });
    });
}