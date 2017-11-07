/// <reference path="F:\Upwork\Microsoft Add-in\CrossPlatformProNet\CrossPlatformProNetWeb\Home.html" />
/// <reference path="F:\Upwork\Microsoft Add-in\CrossPlatformProNet\CrossPlatformProNetWeb\Home.html" />
/// <reference path="/Scripts/FabricUI/MessageBanner.js" />


(function () {
    "use strict";



    // The initialize function must be run each time a new page is loaded.
    Office.initialize = function (reason) {
      
    };

  



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
    app.controller('myCtrl', function ($scope, $http, $compile) {

        //initializations
        $scope.UserLibraries = [];
        $scope.headerOpened = false;
        $scope.LibrariesDocuments = [];
        $scope.CurrentSelectedText = "";
        $scope.CName = "";
        $scope.UserToken = "";
        $scope.LibraryID = "";
        $scope.oneAtATime = true;
        $scope.status = {
            isCustomHeaderOpen: false,
            isFirstOpen: true,
            isFirstDisabled: false
        };

        //functions


       var dialog;
        function ShowTokenDialog() {


            // Office.context.ui.displayDialogAsync('https://localhost:44380/Dialog.html', { height: 30, width: 20 });


            Office.context.ui.displayDialogAsync('https://amiraelmahdaly.github.io/Dialog/Dialog.html', { height: 35, width: 20 },
                function (asyncResult) {
                    dialog = asyncResult.value;
                    dialog.addEventHandler(Office.EventType.DialogMessageReceived, processMessage);
                }
            );


        }

        
        $scope.$on('ngRepeatFinished', function (ngRepeatFinishedEvent) {
            $('#accordion').find('.accordion-toggle').click(function () {

                //Expand or collapse this panel
                $(this).next().slideToggle('fast');
                $(this).children().toggleClass("active");
                //Hide the other panels
                $(".accordion-content").not($(this).next()).slideUp('fast');
                $(".accordion-toggle").not($(this)).children().removeClass("active");
            });
        });
        $scope.SaveCurrentDoc = function (text, name) {
            
            if ($scope.CName == "") $scope.CName = name;

            if ($scope.CName == name) {
                $scope.CurrentSelectedText = text;
                $scope.headerOpened = !$scope.headerOpened;
                document.getElementById("btnInsert").disabled = !$scope.headerOpened;
            }
            else {
                $scope.CurrentSelectedText = text;

                $scope.CName = name;
                $scope.headerOpened = true;
                document.getElementById("btnInsert").disabled = false;

                

            }


        }
       /* $scope. insertText = function(text) {
            Word.run(function (context) {

               
                // Create a proxy object for the document body.
                var body = context.document.body;

                // Queue a command to insert text at the end of the document body.
                body.insertText(text, Word.InsertLocation.end);

                // Synchronize the document state by executing the queued commands,
                // and return a promise to indicate task completion.
                return context.sync().then(function () {
                    console.log('Added a quote from a Chinese proverb.');
                });
            })
            .catch(function (error) {
                console.log('Error: ' + JSON.stringify(error));
                if (error instanceof OfficeExtension.Error) {
                    console.log('Debug info: ' + JSON.stringify(error.debugInfo));
                }
            });
        }*/
  $scope. insertText = function(text) {
            Word.run(function (context) {

                // Create a proxy object for the document.
                var thisDocument = context.document;

                // Queue a command to get the current selection.
                // Create a proxy range object for the selection.
                var range = thisDocument.getSelection();

                // Queue a command to replace the selected text.
                range.insertText(text, Word.InsertLocation.replace);

                // Synchronize the document state by executing the queued commands,
                // and return a promise to indicate task completion.
                return context.sync().then(function () {
                    console.log('Added a quote from Ralph Waldo Emerson.');
                });
            })
            .catch(function (error) {
                console.log('Error: ' + JSON.stringify(error));
                if (error instanceof OfficeExtension.Error) {
                    console.log('Debug info: ' + JSON.stringify(error.debugInfo));
                }
            });
        }


        function processMessage(arg) {
            dialog.close();
            $scope.Initial();

        }
        angular.element(document).ready(function () {
            $scope.Initial();


        });
        $scope.Initial = function () {
            //document.getElementById("btnGetToken").disabled = true;
            $scope.UserLibraries = [];
            
            $scope.LibrariesDocuments = [];
            if (localStorage.getItem("userToken") == null || localStorage.getItem("userToken").trim() == "")
                ShowTokenDialog();
            else {
                $scope.UserToken = localStorage.getItem("userToken");
                $scope.CheckUserToken();
               
                //$scope.GetUserLibraries();
            }

        }
        $scope.GetUserLibraries = function () {


            var uri = "https://test-landlord.assistantbroker.com/ajax/language_library_ajax.php?action=get_user_libraries&token=" + $scope.UserToken;
            $http.get(uri,
              {
                  'Content-Type': 'application/json',
                  'Accept': 'application/json'
              })
              .then
            (
                  function (response) {
                      console.log(response);
                     

                      $scope.UserLibraries = response.data.data;
           
                    

                      console.log($scope.UserLibraries);
                  },
                  function (response) {
                      // $scope.UserLibraries = null;
                  }
               );

        }
        $scope.GetLibrariesDocuments = function () {
            //document.getElementById("selectLib").removeChild(document.getElementById("sel"));
            document.getElementById("btnInsert").disabled = true;
            var uri = "https://test-landlord.assistantbroker.com/ajax/language_library_ajax.php?action=get_data_for_onlyoffice_plugin&token=" + $scope.UserToken + "&library_id=" + $scope.LibraryID;
            $http.get(uri,
              {
                  'Content-Type': 'application/json',
                  'Accept': 'application/json'
              })
              .then
            (
                  function (response) {
                      $scope.LibrariesDocuments = response.data.snippets;
                      console.log($scope.LibrariesDocuments);
                     
                  },
                  function (response) {
                      $scope.LibrariesDocuments = null;
                      console.log($scope.LibrariesDocuments);
                  }
               );

        }
        $scope.RemoveToken = function () {
            localStorage.removeItem("userToken");
            $scope.Initial();
        }
        $scope.CheckUserToken = function () {
            var uri = "https://test-landlord.assistantbroker.com/ajax/language_library_ajax.php?action=get_user_libraries&token=" + $scope.UserToken;
            $http.get(uri,
              {
                  'Content-Type': 'application/json',
                  'Accept': 'application/json'
              })
              .then
            (
                  function (response) {
                      console.log(response);
                      if (response.data.message == "Invalid token" || $scope.UserToken.trim()=="") {
                          ShowTokenDialog();
                          }


                      
                      else {
                          $scope.UserLibraries = [];
                          $scope.LibrariesDocuments = [];
                       //   var optionID = document.getElementById("selectLib").options[document.getElementById("selectLib").selectedIndex].id;
                          $scope.GetUserLibraries();
                          //if(optionID != null)
                          //document.getElementById(optionID).selected = true;

                          $scope.LibraryID = -1200;
                         $scope.GetLibrariesDocuments();
                         

                      }
                      console.log($scope.UserLibraries);
                  },
                  function (response) {
                      // $scope.UserLibraries = null;
                  }
               );

        }

        $("#selectLib").change(function () {
            $scope.LibraryID = $(this).children(":selected").attr("id");
            if (document.getElementById("selectLib").contains(document.getElementById("sel")))
            document.getElementById("selectLib").removeChild(document.getElementById("sel"));

            $scope.GetLibrariesDocuments();

        });
        $scope.btnRefresh = function () {

            if (!document.getElementById("selectLib").contains(document.getElementById("sel")))
            $('#selectLib').prepend($('<option id="sel">select</option>'));

            $scope.CheckUserToken();
        }

    });


 
})();


