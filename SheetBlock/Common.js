﻿"use strict";
Office.initialize = function (reason) {
};
var app = angular.module('myApp', []);
app.directive('onFinishRender', function ($timeout) {
    return {
        restrict: 'A',
        link: function (scope, element, attr) {
            if (scope.$last === true) {
                $timeout(function () {
                    scope.$emit(attr.onFinishRender);

                });
            }
        }
    }
});
//var DeploymentHost = "https://realdocs.pronetcre.com/";
var DeploymentHost = "https://localhost:44391/";
var messageBanner;
var BaseURI = "https://dakota-wcf.ezsoftco.com/WCFEzapptJsonService.svc/";
// Error Handling Region
$(document).ready(function () {
    var element = document.querySelector('.ms-MessageBanner');
    messageBanner = new fabric.MessageBanner(element);
    messageBanner.hideBanner();

});
function hideErrorMessage() {

    setTimeout(function () {
        messageBanner.hideBanner();
    }, 2000);
}
// Helper function for treating errors
function errorHandler(error) {
    showNotification("Error", error);
}
// Helper function for displaying notifications
function showNotification(header, content) {
    $("#notificationHeader").text(header);
    $("#notificationBody").text(content);
    messageBanner.showBanner();
    messageBanner.toggleExpansion();
    hideErrorMessage();
}
function FormatParams(params) {
    var par = "";
    for (var i = 1; i < params.length; i++) {
        if (i != params.length - 1)
            par += params[i] + "/";
        else
            par += params[i];

    }
    return par;
}
function AnyEmpty() {
    for (var i = 0; i < arguments.length; i++)
        if (arguments[i].trim() == "") return true;
    return false;
}
function Redirect(q) {
    window.location.href = q;
}
function getQueryStringValue(key) {
    return decodeURIComponent(window.location.search.replace(new RegExp("^(?:.*[&\\?]" + encodeURIComponent(key).replace(/[\.\+\*]/g, "\\$&") + "(?:\\=([^&]*))?)?.*$", "i"), "$1"));
}
Storage.prototype.setObj = function (key, obj) {
    return this.setItem(key, JSON.stringify(obj))
}
Storage.prototype.getObj = function (key) {
    return JSON.parse(this.getItem(key))
}

app.service('AngularServices', ['$http', function ($http) {
    var API = {
        Request: function (EndPoint) {
            return $http.get(BaseURI + EndPoint + "/" + FormatParams(arguments))
                .then(function (response) {
                    return response.data;
                }).catch(errorHandler);
        }
    };

    return API;
 }]);



