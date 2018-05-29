(function () {
  "use strict";

  var messageBanner;

  // The Office initialize function must be run each time a new page is loaded.
  Office.initialize = function (reason) {
    $(document).ready(function () {
      var element = document.querySelector('.ms-MessageBanner');
      messageBanner = new fabric.MessageBanner(element);
      messageBanner.hideBanner();
      loadProps();
    });
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
    app.controller('myCtrl', function ($scope, $http, $compile) {



       
        $scope.Password = "";
        $scope.Email = "";

        var URI = "https://dakota-wcf.ezsoftco.com/WCFEzapptJsonService.svc/ValidateUserPassword/"
        $("#btnLogin").click(Login);
        function Login() {
            Request($scope.Email, $scope.Password).then(function (data) {
                switch (data.ValidateUserPasswordResult.CodeData) {
                    case -1:
                        showNotification("Notification", "There’s no staff");
                        break;
                    case -2:
                        showNotification("Notification", "Incorrect Password");
                        break;
                    case 0:
                        showNotification("Notification", "Successful Login");
                        break;
                    
                    default:
                }
            });
        }
        function Request() {
            return $http.get(URI + FormatParams(arguments))
                .then(function (response) {
                    return response.data;
                }).catch(errorHandler);
        }
        function FormatParams(params) {
            var par = "";
            for (var i = 0; i < params.length; i++) {
                if (i != params.length - 1)
                    par += params[i] + "/"
                else
                    par += params[i];

            }
            return par;
        }
        // Error Handling Region
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

    });

})();