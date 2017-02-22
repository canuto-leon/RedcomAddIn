SP.SOD.executeFunc('sp.js', 'SP.ClientContext', initializePageCorreo);
var web;
var currentUser;
function initializePageCorreo() {
    //Get the URI decoded URLs.
    hostweburl = decodeURIComponent(getQueryStringParameter("SPHostUrl"));
    appweburl = decodeURIComponent(getQueryStringParameter("SPAppWebUrl"));
    var scriptbase = hostweburl + "/_layouts/15/";
    $.getScript(scriptbase + "SP.RequestExecutor.js", createExecutorAndContextCorreo);
}
function createExecutorAndContextCorreo() {
    context = new SP.ClientContext(appweburl);
    factory = new SP.ProxyWebRequestExecutorFactory(appweburl);
    context.set_webRequestExecutorFactory(factory);
    appContextSite = new SP.AppContextSite(context, hostweburl);
    web = appContextSite.get_web();
    context.executeQueryAsync(onGetWebSuccess, onGetWebFail);
    function onGetWebSuccess() {
        currentUser = web.get_currentUser();
        context.load(currentUser);
        context.executeQueryAsync(onGetUserSuccess, onGetUserFail);
    }
    function onGetWebFail(sender, args) {
        alert('Failed to get web. Error:' + args.get_message());
    }
    function onGetUserSuccess() {
        $('#para').val(currentUser.get_email());
        $('#botonEnviar').show();
    }
    function onGetUserFail(sender, args) {
        alert('Failed to get user. Error:' + args.get_message());
    }
}
function enviaCorreo() {
    var mail = $('#para').val();
    var url = getQueryStringParameter('urlItem');
    var subject = $("#asunto").val();
    var body = $("#cuerpo").val() + "\n\n" + url;
    sendMail(mail, subject, body);
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
            }),
            headers: {
                "Accept": "application/json;odata=verbose",
                "content-type": "application/json;odata=verbose",
                "X-RequestDigest": $("#__REQUESTDIGEST").val()
            },
            success: function (data) {
                alert("Correo enviado");
                window.location.replace(document.referrer);
            },
            error: function (xhr, ajaxOptions, thrownError) {
                alert(xhr.status);
                alert(thrownError);
            }
        });
    });
}
function sendMail(toList, subject, mailContent) {
    var email = currentUser.get_email();
    var restUrl = hostweburl + "/_api/SP.Utilities.Utility.SendEmail", restHeaders = {
        "Accept": "application/json;odata=verbose",
        "content-type": "application/json;odata=verbose",
        "X-RequestDigest": $("#__REQUESTDIGEST").val()
    }, mailObject = {
        'properties': {
            '__metadata': {
                'type': 'SP.Utilities.EmailProperties'
            },
            'From': email,
            'To': {
                'results': toList
            },
            'Subject': subject,
            'Body': mailContent,
        }
    };
    return $.ajax({
        contentType: "application/json",
        url: restUrl,
        type: "POST",
        data: mailObject,
        headers: restHeaders,
        success: onSuccessMailSent,
        error: onFailMailSent
    });
    function onSuccessMailSent() {
        alert("Mail enviado");
    }
    function onFailMailSent(xhr, ajaxOptions, thrownError) {
        alert(xhr.responseText);
        alert(thrownError);
    }
}
//# sourceMappingURL=Correo.js.map