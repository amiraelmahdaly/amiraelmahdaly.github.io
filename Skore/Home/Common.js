"use strict";
Office.initialize = function (reason) {
};
var app = angular.module('myApp', []);
app.config(['$httpProvider', function ($httpProvider) {
    $httpProvider.defaults.withCredentials = true;
}]);

app.directive('onFinishRender', function ($timeout) {
    return {
        restrict: 'A',
        link: function (scope, element, attr) {
            if (scope.$last === true) {
                $timeout(function () {
                    scope.$emit('ngRepeatFinished');


                });
            }
        }
    };
});
//var DeploymentHost = "";
var DeploymentHost = "https://localhost:44353/Home/";
var messageBanner;
var subDomain = "app";
var BaseURI = "https://" + subDomain +".spekit.co/api/";
   
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
    return error;

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
};
Storage.prototype.getObj = function (key) {
    return JSON.parse(this.getItem(key))
};
function removeObj(myObjects, prop, valu) {
    return myObjects.filter(function (val) {
        return val[prop] !== valu;
    });

}
function TrackEvent(eventName, email, eventObj) {
    eventObj.Timestamp = (new Date()).toUTCString();
    eventObj.user = email;
    analytics.track("Outlook Extension:" +  eventName, eventObj);
}
app.service('AngularServices', ['$http', function ($http) {
    var API = {
        GET: function (EndPoint) {

            return $http(
                {
                    method: 'GET',
                    url: BaseURI + EndPoint,
                   
                })
                .then(function (response) {
                    return response;
                }).catch(function (response) {
                    return response;
                });
        }
        ,
        POST: function (EndPoint, body, headers) {
            var settings = {
                method: 'POST',
                url: BaseURI + EndPoint,
                data: body,
                headers: headers
            };
            return $http(settings)
                .then(function (response) {
                    return response;
                }).catch(function (response) {
                    return response;
                });
        }
   



    };

    return API;
}]);



