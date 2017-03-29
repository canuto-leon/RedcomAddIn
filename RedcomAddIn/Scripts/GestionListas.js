///<reference path="typings/sharepoint/SharePoint.d.ts" /> 
///<reference path="typings/jquery/jquery.d.ts" />
var web;
var site;
var oLists;
var mostrarOcultas = false;
SP.SOD.executeFunc('sp.js', 'SP.ClientContext', cargaComunesListas);
function cargaComunesListas() {
    SP.SOD.executeFunc('Comunes.js', 'getQueryStringParameter', initializePageLists);
}
function initializePageLists() {
    //Get the URI decoded URLs.
    hostweburl = decodeURIComponent(getQueryStringParameter("SPHostUrl"));
    appweburl = decodeURIComponent(getQueryStringParameter("SPAppWebUrl"));
    mostrarOcultas = getQueryStringParameter("MostrarListas") === 'true';
    var scriptbase = hostweburl + "/_layouts/15/";
    $.getScript(scriptbase + "SP.RequestExecutor.js", createExecutorAndContextListas);
}
function createExecutorAndContextListas() {
    context = new SP.ClientContext(appweburl);
    factory = new SP.ProxyWebRequestExecutorFactory(appweburl);
    context.set_webRequestExecutorFactory(factory);
    appContextSite = new SP.AppContextSite(context, hostweburl);
    getSiteId();
}
function getSiteId() {
    site = appContextSite.get_site();
    context.load(site);
    context.executeQueryAsync(onGetSiteSuccess, onGetSiteFail);
    function onGetSiteSuccess() {
        $('#idColeccion')[0].innerHTML = site.get_id();
        getListsFromWeb();
    }
    function onGetSiteFail(sender, args) {
        alert('Failed to get site. Error:' + args.get_message());
    }
}
function getListsFromWeb() {
    web = appContextSite.get_web();
    context.load(web);
    context.executeQueryAsync(onGetWebSuccess, onGetWebFail);
    function onGetWebSuccess() {
        $('#nombreSitio')[0].innerHTML = web.get_title();
        dameListas();
    }
    function onGetWebFail(sender, args) {
        alert('Failed to get web. Error:' + args.get_message());
    }
}
function dameListas() {
    oLists = web.get_lists();
    context.load(oLists, "Include(Title, DefaultViewUrl, Hidden, Id)");
    context.executeQueryAsync(onGetListsSuccess, onGetListsFail);
    function onGetListsSuccess() {
        var listEnumerator = oLists.getEnumerator();
        while (listEnumerator.moveNext()) {
            var oList = listEnumerator.get_current();
            if (!oList.get_hidden() || mostrarOcultas)
                anadeFila(oList.get_title(), oList.get_defaultViewUrl(), oList.get_id());
        }
        redimensiona();
    }
    function onGetListsFail(sender, args) {
        alert('Failed to get Lists. Error:' + args.get_message());
    }
}
function anadeFila(titulo, url, id) {
    var tabla = document.getElementById("tablaListas");
    var row = tabla.insertRow(tabla.rows.length);
    var cell = row.insertCell();
    cell.innerHTML = "<a href='" + url + "' target='_blank'>" + titulo + "</a>";
    cell = row.insertCell();
    cell.innerHTML = "<img src='../Images/items25x25.png'>";
    cell.style.textAlign = "center";
    cell.addEventListener("click", function () { redireccionaItems(id); });
}
function redireccionaItems(id) {
    var query = document.URL.split("?")[1];
    window.location.replace("Items.aspx?" + query + "&IdLista=" + id);
}
function redimensiona() {
    var table = $("#tablaListas")[0];
    var height = table.offsetHeight;
    var width = table.offsetWidth;
    var senderId = getQueryStringParameter('SenderId');
    if (inIframe) {
        var resizeMessage = '<message senderId={Sender_ID}>resize({Width}, {Height})</message>';
        resizeMessage = resizeMessage.replace("{Sender_ID}", senderId);
        resizeMessage = resizeMessage.replace("{Height}", String(height + 100));
        resizeMessage = resizeMessage.replace("{Width}", String(width + 30));
        window.parent.postMessage(resizeMessage, "*");
    }
}
//# sourceMappingURL=GestionListas.js.map