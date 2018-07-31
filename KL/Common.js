"use strict";
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
//var DeploymentHost = "https://amiraelmahdaly.github.io/ezappt/";
var DeploymentHost = "https://localhost:44391/";
var messageBanner;
var BaseURI = //"https://private-98aeb-wordaddin.apiary-mock.com/";
    //"https://private-be06fe-wordaddin.apiary-mock.com/";
    "https://private-9dc67-wordaddin.apiary-mock.com/word/";
   
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
        GET: function (EndPoint) {

            return $http(
                {
                    method: 'GET',
                    url: BaseURI + EndPoint + "/" + FormatParams(arguments),
                    headers: {
                      
                        'Cache-Control': 'no-cache',
                        'Pragma': 'no-cache'
                    }
                })
                .then(function (response) {
                    return response.data;
                }).catch(errorHandler);
        }
        ,
        POST: function (EndPoint, body, token) {
            var settings = {
                method: 'POST',
                url: BaseURI + EndPoint,
                data: body,
                headers: {
                    'Authorization': token,
                    'Content-Type': 'application/json'

                }
            };
            return $http(settings)
                .then(function (response) {
                    return response.data;
                }).catch(errorHandler);
        }
        ,
        PUT: function (EndPoint, body, token, docId) {
            return $http({
                method: 'PUT',
                url: BaseURI + EndPoint + "/" + docId,
                data: body,
                headers: {
                    'Authorization': token,
                    'Content-Type': 'application/json'

                }
                })
                .then(function (response) {
                    return response.data;
                }).catch(errorHandler);
        }



    };

    return API;
}]);



