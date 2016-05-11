'use strict';

var context = SP.ClientContext.get_current();
var appWebUrl = decodeURIComponent(getQueryStringParameter("SPAppWebUrl"));
var hostWebUrl = decodeURIComponent(getQueryStringParameter("SPHostUrl"));
var formDigest;

function getQueryStringParameter(requestedParameter) {
    var param = document.URL.split("?")[1].split("&");

    for (var i = 0; i < param.length; i++) {
        var actual = param[i].split("=");
        if (actual[0] == requestedParameter) {
            return actual[1];
        }
    }
}

//1) Creamos el objeto PeopleManager en el contexto actual.
function createManager() {

    $.ajax({
        url: appWebUrl + "/_api/sp.userprofiles.peoplemanager",
        type: "GET",
        headers: { "accept": "application/json;odata=verbose" },
        success: function (data) {
            alert("PeopleManager creado.");
        },
        error: function (xhr) {
            alert(xhr.responseText);
        }
    });
}

//Obtendremos todas las propiedades del usuario.
function getActualUserProperties() {

    $.ajax({
        url: appWebUrl + "/_api/SP.UserProfiles.PeopleManager/GetMyProperties",
        type: "GET",
        headers: { "accept": "application/json;odata=verbose" },
        success: function (data) {
            var html = "<ul>";
            $.each(data.d.UserProfileProperties.results, function(index, prop) {
                html += "<li>" + prop.Key + " : " + prop.Value + "</li>";
            });
            html += "</ul>";
            $('#res').html("");
            $('#res').append(html);
        },
        error: function (xhr) {
            alert(xhr.responseText);
        }
    });
}

//
function getOtherUserProperties() {

    var splitted = $('#txtUser').val().split('@');
    var user = splitted[0];
    var dominio = splitted[1];

    $.ajax({
        url: appWebUrl + "/_api/SP.UserProfiles.PeopleManager/GetUserProfilePropertyFor(accountName=@v,propertyName='PreferredName')?@v='i%3A0%23.f%7Cmembership%7C"+user+"%40"+dominio+"'",
        type: "GET",
        headers: { "accept": "application/json;odata=verbose" },
        success: function (data) {
            var preferredName = data.d.GetUserProfilePropertyFor;
            var html = "<a href='#' onclick='follow(" +'"'+preferredName+'"'+ ")'> Seguir a: " + preferredName + "</a>"; 
            $('#res').html("");
            $('#res').append(html);
        },
        error: function (xhr) {
            alert(xhr.responseText.toString());
        }
    });
}

function follow(name) {

    var splitted = $('#txtUser').val().split('@');
    var user = splitted[0];
    var dominio = splitted[1];
    
    $.ajax({
        url: "/_api/sp.userprofiles.peoplemanager/follow(@v)?@v='i%3A0%23.f%7Cmembership%7C"+user+"%40"+dominio+"'",
        type: "POST",
    headers: { "X-RequestDigest":  formDigest},
    success: function(data) {
        alert("Has empezado a seguir a " + name);
    },
    error: function(xhr) {
        alert(xhr.responseText);
    }
});
    
}

var getFormDigest = function() {
    $.ajax({
        url: appWebUrl + "/_api/contextinfo",
        type: "POST",
        contentType: "application/json;odata=verbose",
        headers: {
            'accept':'application/json;odata=verbose'
        },
        success:function(data) {
            formDigest = data.d.GetContextWebInformation.FormDigestValue;
        },
        error:function(xhr) {
            alert(xhr.responseText);
        },
        async:false
    });
}();

$(document).ready(function () {
    createManager();
    getActualUserProperties();
    getFormDigest();
});
