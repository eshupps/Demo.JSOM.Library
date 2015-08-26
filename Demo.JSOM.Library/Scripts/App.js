'use strict';
var appweburl;
var hostweburl;
var appContext;
var hostContext;


$(document).ready(function () {
    hostweburl =
            decodeURIComponent(
                getQueryStringParameter("SPHostUrl")
        );
    appweburl =
        decodeURIComponent(
            getQueryStringParameter("SPAppWebUrl")
    );

    var scriptbase = hostweburl + "/_layouts/15/";

    $.getScript(scriptbase + "SP.Runtime.js",
        function () {
            $.getScript(scriptbase + "SP.js",
                function () {
                    $.getScript(scriptbase + "SP.RequestExecutor.js", getUserName);
            });
        }
    );
});

function getUserName() {
    appContext = new SP.ClientContext(appweburl);
    var factory = new SP.ProxyWebRequestExecutorFactory(appweburl);
    appContext.set_webRequestExecutorFactory(factory);
    hostContext = new SP.AppContextSite(appContext, hostweburl);
    
    var user = appContext.get_web().get_currentUser();
    appContext.load(user);
    appContext.executeQueryAsync(
        function () {
            $('#welcome').html('Welcome,  ' + user.get_title());
        },
        function (sender, args) {
            alert(args.get_message())
        }
    );
}

function createLibrary() {
    $("#message").attr("class","waiting").html("Working on it...");

    var web = hostContext.get_web();
    var lists = web.get_lists();
    var title = $("#libraryName").val();
    var listCreationInfo = new SP.ListCreationInformation();
    listCreationInfo.set_title(title);
    listCreationInfo.set_templateType(SP.ListTemplateType.documentLibrary);
    lists.add(listCreationInfo);

    appContext.load(lists);
    appContext.executeQueryAsync(
        function () {
            $("#message").attr("class","success").html("Library creation succeeded.");
        },
        function (sender, args) {
            $("#message").attr("class","error").html("Library creation failed.<br/>" + args.get_message());
        }
    );
}

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
