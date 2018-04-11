/// <reference path="/Scripts/FabricUI/MessageBanner.js" />


(function () {
    "use strict";

    var messageBanner;

    // The initialize function must be run each time a new page is loaded.
    Office.initialize = function (reason) {

    };
    
    $(document).ready(function () {
        $("#mainCon").attr("style", "display:block;");
        $("#header").hide();
        $("#accordion").hide();
        $("#grid-row1").hide();
        var element = document.querySelector('.ms-MessageBanner');
        messageBanner = new fabric.MessageBanner(element);
        messageBanner.hideBanner();
   

    });

    var app = angular.module('myApp', ['ngAnimate', 'ngSanitize', 'ui.bootstrap']);
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
    app.directive('attrs', function() {
        return {
            link: function(scope, element, attrs) {
                var attrs = angular.copy(scope.$eval(attrs.attrs));
                element.attr(attrs).html(attrs.html);
            }
        };
    });
    app.filter('customArray', function ($filter) {
        return function (list, arrayFilter, element) {
            if (arrayFilter) {
                return $filter("filter")(list, function (listItem) {
                    return arrayFilter.indexOf(listItem[element]) != -1;
                });
            }
        };
    });
    app.controller('myCtrl', function ($scope, $http, $compile) {

        //initializations
        $scope.Wells = [];
        $scope.Wellbores = [];
        $scope.UIDWell = -1;



        // Functions 

        $scope.USerCredentials = { UserName: "", Password: "" };
        var TokenDialog;
        var TokenDialogUrl = DeploymentHost + "TokenDialog.html";

        function SlideToggle() {
            $(this).next().slideToggle("fast");
            $(this).toggleClass("active");
            $(".active").not($(this)).next().slideUp("fast");
            $(".active").not($(this)).removeClass("active");
            var id = $(this).attr('data-uidWell').replace(".", "\\.");
         // showNotification("html" ,$("#" + id).parent().html());
            if ($("[data-uidwell=" + id + "]").hasClass("active"))
                $(this).next().children().children('div:first').show();
          
        }
        $scope.$on('ngRepeatFinished', function (ngRepeatFinishedEvent) {
            


            $('.accordion-toggle').click(function () {
              
                $(".wellboreCon").hide();
                setTimeout(SlideToggle.bind(this), 0);
                var id = $(this).attr('data-uidWell');
                $scope.GetWellbores(id);



            });
            $("#grid-row1").hide();
            $("#accordion").show();
            $("#header").show();
           
       })
        function ShowTokenDialog() {
            Office.context.ui.displayDialogAsync(TokenDialogUrl, { height: 27, width: 22,displayInIframe: true },
                function (asyncResult) {
                    TokenDialog = asyncResult.value;
                    TokenDialog.addEventHandler(Office.EventType.DialogMessageReceived, processtokenDialogMessage);
                   // TokenDialog.addEventHandler(Office.EventType.DialogEventReceived, TokenDialogClosed);
                }
            );
        }
        function sleep(miliseconds) {
            var currentTime = new Date().getTime();

            while (currentTime + miliseconds >= new Date().getTime()) {
            }
        }
        function processtokenDialogMessage(arg) {
            var MessageObj = JSON.parse(arg.message);
            $scope.USerCredentials.UserName = MessageObj.UserName;
            $scope.USerCredentials.Password = MessageObj.Password;
            TokenDialog.close();
            RedirectIfTagged();
            $("#grid-row1").show();
            $scope.GetWells();

        }

       

        angular.element(document).ready(function () {
            $scope.Initial();


        });


    
  
        function RedirectIfTagged() {
            Word.run(function (context) {
                context.document.properties.load("comments");
                return context.sync().then(function () {
                    if (context.document.properties.comments != "") {
                        window.location.href = context.document.properties.comments;
                    }
                });
            })
.catch(function (error) {
    console.log('Error: ' + JSON.stringify(error));
    if (error instanceof OfficeExtension.Error) {
        console.log('Debug info: ' + JSON.stringify(error.debugInfo));
    }
});
        }
     
        $scope.GetWells = function () {
            $http.get(GetURI("well",'<wells xmlns="http://www.witsml.org/schemas/131" version="1.3.1.1"> <well uid=""><name/></well></wells>'),
                {
                    headers: GetHeader($scope.USerCredentials.UserName, $scope.USerCredentials.Password)
                })
                .then(function (response) {
                 $scope.Wells = GetJson(response.data.response.result).wells.well;
                 
                }).catch(function (e) {
                    throw e;
                    console.log(e);
                });

        }
        $scope.GetWellbores = function (id) {
            //if (!$('#accordion').child('h3').hasClass('ui-state-active')) return;
            var ideditted = id.replace(".", "\\.");
            if ($("#" + ideditted).hasClass("active")) return;
      
            $scope.UIDWell = id;
           // $("#" + $scope.UIDWell).next("div").first().text("");
            $http.get(GetURI("wellbore", '<wellbores xmlns="http://www.witsml.org/schemas/131" version="1.3.1.1"><wellbore uidWell="' + $scope.UIDWell+ '" uid=""><name /></wellbore></wellbores>'),
              {
                  headers: GetHeader($scope.USerCredentials.UserName, $scope.USerCredentials.Password)
              })
              .then(function (response) {
                  $scope.Wellbores = [];
                  if ($.isArray(GetJson(response.data.response.result).wellbores.wellbore))
                      $scope.Wellbores = GetJson(response.data.response.result).wellbores.wellbore;
                  else
                      $scope.Wellbores.push(GetJson(response.data.response.result).wellbores.wellbore)

                  $(".grid-row2").hide();
                  $(".wellboreCon").show();
              }).catch(function (data) {
                  console.log(data);
                  $(".grid-row2").hide();
              });
        }
      
       $scope.Initial = function () {

           ShowTokenDialog();
          
        }

    });

    //$$(Helper function for treating errors, $loc_script_taskpane_home_js_comment34$)$$
    function errorHandler(error) {
        // $$(Always be sure to catch any accumulated errors that bubble up from the Word.run execution., $loc_script_taskpane_home_js_comment35$)$$
        showNotification("Error:", error);
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
    }



  

    // Helper function for displaying notifications
    function showNotification(header, content) {
        $("#notificationHeader").text(header);
        $("#notificationBody").text(content);
        messageBanner.showBanner();
        messageBanner.toggleExpansion();
    }
})();
