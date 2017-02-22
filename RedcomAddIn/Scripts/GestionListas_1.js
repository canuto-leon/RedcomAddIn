var hostweburl;
var appweburl;
var context;
var web;
var oLists;

SP.SOD.executeFunc('sp.js', 'SP.ClientContext', initializePage);
// Load the required SharePoint libraries
function initializePage() {
    //Get the URI decoded URLs.
    hostweburl =
        decodeURIComponent(
            getQueryStringParameter("SPHostUrl")
    );
    appweburl =
        decodeURIComponent(
            getQueryStringParameter("SPAppWebUrl")
    );


    var scriptbase = hostweburl + "/_layouts/15/";
    $.getScript(scriptbase + "SP.RequestExecutor.js", getListsFromWeb);
}
function getListsFromWeb()
{
    context = new SP.ClientContext(appweburl);
    factory = new SP.ProxyWebRequestExecutorFactory(appweburl);
    context.set_webRequestExecutorFactory(factory);
    appContextSite = new SP.AppContextSite(context, hostweburl);

    web = appContextSite.get_web();
    context.load(web);

    context.executeQueryAsync(onGetWebSuccess, onGetWebFail);

    function onGetWebSuccess() {
        $('#nombreSitio').innerHTML = web.get_title();
        dameListas();
    }

    function onGetWebFail(sender, args) {
        alert('Failed to get web. Error:' + args.get_message());
    }
}


// Function to retrieve a query string value.
// For production purposes you may want to use
//  a library to handle the query string.
function getQueryStringParameter(paramToRetrieve) {
    var params =
        document.URL.split("?")[1].split("&");
    var strParams = "";
    for (var i = 0; i < params.length; i = i + 1) {
        var singleParam = params[i].split("=");
        if (singleParam[0] == paramToRetrieve)
            return singleParam[1];
    }
}

function dameListas() {
    oLists = web.get_lists();
    context.load(oLists, "Include(Title, DefaultViewUrl)");
    context.executeQueryAsync(onGetListsSuccess, onGetListsFail);

    function onGetListsSuccess() {
        var listEnumerator = oLists.getEnumerator();

        while (listEnumerator.moveNext()) {
            var oList = listEnumerator.get_current();
            anadeFila(oList.get_title(), oList.get_defaultViewUrl());
        }
        redimensiona();
    }

    function onGetListsFail(sender, args) {
        alert('Failed to get Lists. Error:' + args.get_message());
    }
}

function redimensiona()
{
    var table = $("#tablaListas");
    var height = "outerHeight" in table ? table.outerHeight() : table.offsetHeight;
    var width = "outerWidth" in table ? table.outerWidth() : table.offsetWidth;
    var senderId = getQueryStringParameter('SenderId');
    if (inIframe)
    {
        resizeMessage = '<message senderId={Sender_ID}>resize({Width}, {Height})</message>';
        resizeMessage = resizeMessage.replace("{Sender_ID}", senderId);
        resizeMessage = resizeMessage.replace("{Height}", height + 100);
        resizeMessage = resizeMessage.replace("{Width}", width + 30);
        window.parent.postMessage(resizeMessage, "*");
    }

    function inIframe() {
        try {
            return window.self !== window.top;
        } catch (e) {
            return true;
        }
    }
}

function anadeFila(titulo, url) {
    var tabla = document.getElementById("tablaListas");
    var row = tabla.insertRow(tabla.rows.length);
    row.insertCell().innerText = titulo;
    var cell = row.insertCell();
    cell.innerHTML = "<a href='" + url + "' target='_blank'>Enlace</a>";
}

function getUrlVars() {
    var vars = [], hash;
    var hashes = window.location.href.slice(window.location.href.indexOf('?') + 1).split('&');
    for (var i = 0; i < hashes.length; i++) {
        hash = hashes[i].split('=');
        vars.push(hash[0]);
        vars[hash[0]] = hash[1];
    }
    return vars;
}


