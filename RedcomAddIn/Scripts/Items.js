///<reference path="typings/sharepoint/SharePoint.d.ts" /> 
///<reference path="typings/jquery/jquery.d.ts" />
var web;
var oList;
var items;
SP.SOD.executeFunc('sp.js', 'SP.ClientContext', initializePageItems);
function initializePageItems() {
    //Get the URI decoded URLs.
    hostweburl = decodeURIComponent(getQueryStringParameter("SPHostUrl"));
    appweburl = decodeURIComponent(getQueryStringParameter("SPAppWebUrl"));
    var scriptbase = hostweburl + "/_layouts/15/";
    $.getScript(scriptbase + "SP.RequestExecutor.js", createExecutorAndContextItems);
}
function createExecutorAndContextItems() {
    context = new SP.ClientContext(appweburl);
    factory = new SP.ProxyWebRequestExecutorFactory(appweburl);
    context.set_webRequestExecutorFactory(factory);
    appContextSite = new SP.AppContextSite(context, hostweburl);
    web = appContextSite.get_web();
    context.load(web);
    context.executeQueryAsync(onGetWebSuccess, onGetWebFail);
    function onGetWebSuccess() {
        dameLista();
    }
    function onGetWebFail(sender, args) {
        alert('Failed to get web. Error:' + args.get_message());
    }
}
function dameLista() {
    oList = web.get_lists().getById(getQueryStringParameter("IdLista"));
    context.load(oList, ['Title', 'DefaultDisplayFormUrl']);
    context.executeQueryAsync(onGetListSuccess, onGetListFail);
    function onGetListSuccess() {
        loadItems();
    }
    function onGetListFail(sender, args) {
        alert('Failed to get list. Error:' + args.get_message());
    }
}
function loadItems() {
    items = oList.getItems(new SP.CamlQuery());
    context.load(items, 'Include(DisplayName, ServerRedirectedEmbedUrl, FieldValuesAsText)');
    context.executeQueryAsync(onGetItemsSuccess, onGetItemsFail);
    function onGetItemsSuccess() {
        $('#nombreLista')[0].innerHTML = oList.get_title();
        var itemsEnumerator = items.getEnumerator();
        while (itemsEnumerator.moveNext()) {
            var oItem = itemsEnumerator.get_current();
            anadeFilaItems(oItem);
        }
        redimensionaItems();
    }
    function onGetItemsFail(sender, args) {
        alert('Failed to get items. Error:' + args.get_message());
    }
}
function redimensionaItems() {
    var table = $("#tablaItems")[0];
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
function anadeFilaItems(oItem) {
    var fieldValues = oItem.get_fieldValuesAsText().get_fieldValues();
    var id = fieldValues.ID;
    var titulo = oItem.get_displayName();
    var url = hostweburl + oList.get_defaultDisplayFormUrl() + "?Id=" + id;
    var mail = fieldValues.Created_x0020_By.split('|')[2];
    var tabla = document.getElementById("tablaItems");
    var row = tabla.insertRow(tabla.rows.length);
    var cell = row.insertCell();
    cell.innerHTML = "<a href='" + url + "' target='_blank'>" + titulo + "</a>";
    //cell = row.insertCell();
    //cell.innerHTML = "<img src='../Images/email25x25.png'>";
    //cell.style.textAlign = "center";
    //cell.addEventListener("click", function () { redireccionaCorreo(url, mail); });
    cell = row.insertCell();
    cell.innerHTML = "<img src='../Images/eliminar25x25.png'>";
    cell.style.textAlign = "center";
    cell.addEventListener("click", function () { eliminarItem(oItem); });
}
function redireccionaCorreo(url, mail) {
    var query = document.URL.split("?")[1];
    window.location.replace("Correo.aspx?" + query + "&mail=" + mail + "&urlItem=" + url);
}
function eliminarItem(oItem) {
    oItem.deleteObject();
    context.executeQueryAsync(onDeleteItemSuccess, onDeleteItemFail);
    function onDeleteItemSuccess() {
        alert("Elemento eliminado.");
        document.location.href = document.location.href;
    }
    function onDeleteItemFail(sender, args) {
        alert('Failed to get items. Error:' + args.get_message());
    }
}
//# sourceMappingURL=Items.js.map