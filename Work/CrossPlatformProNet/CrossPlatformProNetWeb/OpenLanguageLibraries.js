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
        $scope.UserLibraries = [];
        $scope.headerOpened = false;
        $scope.LibrariesDocuments = [];
        $scope.DocsAssociatedLibraries = null;
        $scope.CurrentSelectedText = "";
        $scope.CName = "";
        $scope.LibraryID = -1200;
        $scope.oneAtATime = true;
        $scope.dial = "";
        $scope.status = {
            isCustomHeaderOpen: false,
            isFirstOpen: true,
            isFirstDisabled: false
        };
        var initialLoad = true;
        var AdminDialog;
        var SwitchAccountDialog;
        var TokenDialog;
       
        //var DeploymentHost = "https://realdocs.pronetcre.com/";
        var DeploymentHost = "https://localhost:44380/";
        var SwitchAccountDialogUrl = DeploymentHost + "SwitchAccount.html";
        var AdminDialogUrl = DeploymentHost + "AdminDialog.html";
        var TokenDialogUrl = DeploymentHost + "TokenDialog.html";
        var DialogOpened = localStorage.getItem("DialogOpened");



        //  ---- Dialogs ---- ///

        //Admin Dialog
        function ShowAdminDialog() {
            if (!DialogOpened) {
                Office.context.ui.displayDialogAsync(AdminDialogUrl, { height: 15, width: 25, displayInIframe: false },
                    function (asyncResult) {
                        AdminDialog = asyncResult.value;
                        AdminDialog.addEventHandler(Office.EventType.DialogMessageReceived, processAdminDialogMessage);
                        AdminDialog.addEventHandler(Office.EventType.DialogEventReceived, AdminDialogClosed);
                    }
                );
                localStorage.setItem("DialogOpened", true);
            }


        }
        function processAdminDialogMessage(arg) {
            AdminDialog.close();
            localStorage.setItem("DialogOpened", false);
            $scope.LibraryID = -1200;
            $scope.Initial();
        }
        function AdminDialogClosed(arg) {
            $scope.Refresh();
            localStorage.setItem("DialogOpened", false);
        }

        //Switch Acccount Dialog
        function ShowSwitchAccountDialog() {
            if (DialogOpened == "false" || DialogOpened === null) {
                Office.context.ui.displayDialogAsync(SwitchAccountDialogUrl, { height: 25, width: 25, displayInIframe: false },
                    function (asyncResult) {
                        SwitchAccountDialog = asyncResult.value;
                        SwitchAccountDialog.addEventHandler(Office.EventType.DialogMessageReceived, ProcessSwitchAccountDialogMessage);
                        SwitchAccountDialog.addEventHandler(Office.EventType.DialogEventReceived, SwitchAccountDialogClosed);



                    }
                );
                localStorage.setItem("DialogOpened", true);
            }


        }
        function ProcessSwitchAccountDialogMessage(arg) {
            SwitchAccountDialog.close();
            $scope.LibraryID = -1200;
            localStorage.setItem("DialogOpened", false);
            $scope.Refresh();

        }
        function SwitchAccountDialogClosed(arg) {
            localStorage.setItem("DialogOpened", false);
        }

        // Token Dialog
        function ShowTokenDialog() {
            if (DialogOpened == "false" || DialogOpened === null) {
                Office.context.ui.displayDialogAsync(TokenDialogUrl, { height: 35, width: 20 },
                    function (asyncResult) {
                        TokenDialog = asyncResult.value;
                        TokenDialog.addEventHandler(Office.EventType.DialogMessageReceived, processtokenDialogMessage);
                        TokenDialog.addEventHandler(Office.EventType.DialogEventReceived, TokenDialogClosed);


                    }
                );
                localStorage.setItem("DialogOpened", true);
            }


        }
        function processtokenDialogMessage(arg) {
            TokenDialog.close();
            localStorage.setItem("DialogOpened", false);
            $scope.Initial();
        }
        function TokenDialogClosed(arg) {
            localStorage.setItem("DialogOpened", false);
        }




        function ShowAccountName() {
            $("#spnAccountName").text(localStorage.getItem("accountName"));
        }
        $scope.$on('ngRepeatFinished', function (ngRepeatFinishedEvent) {


            // $("#selectLib option[id='" + $scope.LibraryID + "']").attr("selected", "selected");

            $('#accordion').find('.accordion-toggle').click(function () {

                //Expand or collapse this panel
                $(this).next().slideToggle('fast');
                $(this).children().toggleClass("active");
                //Hide the other panels
                $(".accordion-content").not($(this).next()).slideUp('fast');
                $(".accordion-toggle").not($(this)).children().removeClass("active");


            });
        });
        $scope.$on('ngRepeatFinished1', function (ngRepeatFinishedEvent) {
            $("#selectLib option[id='" + $scope.LibraryID + "']").attr("selected", "selected");
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
        $scope.insertText = function (text) {
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
        angular.element(document).ready(function () {
            if (localStorage.getItem("userToken") === null) $(".fadeMe").css("display","flex");
            else $scope.Initial();
            setTimeout(function () {
                initialLoad = false;
            }, 3000);
        });
        $scope.Initial = function () {
            localStorage.setItem("DialogOpened", false);
            $scope.UserLibraries.length = 0;
            $scope.LibrariesDocuments.length = 0;
            ShowAccountName();
            $scope.$apply();
            $scope.PrependSelectOption();
            var checkToken = $scope.CheckUserToken();
            checkToken.then(function (response) {
                if (response.result == false) {
                    if (!initialLoad)
                        ShowTokenDialog();
                }
                else {
                    $(".fadeMe").css("display", "none");
                    $("#loader").css("display", "block");
                    $scope.PopulateDocsAssociatedLibsIDS();
                    var getlibs = $scope.GetUserLibraries();
                    getlibs.then(function (response) {
                        for (var i = 0; i < response.data.length; i++) {
                            if (response.data[i].account_id == localStorage.getItem("accountID")) {
                                $scope.UserLibraries = response.data[i].user_libraries;
                                break;
                            }


                        }
                        if ($scope.DocsAssociatedLibraries != null) {

                            $scope.UserLibraries = $scope.UserLibraries.filter(function (emp) {
                                return $scope.DocsAssociatedLibraries.indexOf(emp.id) !== -1;
                            });

                        }
                        $("#loader").css("display", "none");

                    });
                  
                }
            });
        }

        $scope.PopulateDocsAssociatedLibsIDS = function () {
            Word.run(function (context) {
                context.document.properties.load("comments");
                return context.sync().then(function () {
                    if (context.document.properties.comments != "")
                        $scope.DocsAssociatedLibraries = JSON.parse(context.document.properties.comments)
                    else
                        $scope.DocsAssociatedLibraries = null;
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
        $("#btnRefresh").click(function (e) {
            if (e.shiftKey) {
                ShowAdminDialog();
            }
            else {
                $scope.Refresh();
            }
        });
        $("#btnSwitch").click(function () {
            ShowSwitchAccountDialog();
        });
        $scope.LibraryExists = function () {
            for (var i = 0; i < $scope.UserLibraries.length; i++)
                if ($scope.UserLibraries[i].id == $scope.LibraryID) return true;
            return false;
        }
        $scope.GetUserLibraries = function () {
            var uri = localStorage.getItem("hostName") + "/ajax/language_library_ajax.php?action=get_user_libraries_by_account&token=" + localStorage.getItem("userToken");
            return $http.get(uri).then(function (response) {
                return response.data;
            });
        }
        $scope.GetLibrariesDocuments = function () {
            var uri = localStorage.getItem("hostName") + "/ajax/language_library_ajax.php?action=get_data_for_onlyoffice_plugin&token=" + localStorage.getItem("userToken") + "&library_id=" + $scope.LibraryID;
            return $http.get(uri).then(function (response) {
                return response.data;
            });
        }
        $scope.CheckUserToken = function () {
            var uri = "https://auth.pronetcre.com/authentication/ajax/manage_tokens.php?action=get_token_hostname&token=" + localStorage.getItem("userToken");
            return $http.get(uri).then(function (response) {
                return response.data;
            });


        }
        function Insertdocument(doc, DocsAssociatedLibraries) {
            Word.run(function (context) {
                var myNewDoc = context.application.createDocument(doc);
                if (DocsAssociatedLibraries != null)
                    myNewDoc.properties.comments = JSON.stringify(DocsAssociatedLibraries);
                context.load(myNewDoc);
                return context.sync().then(function () {
                    myNewDoc.open();
                });
            })
.catch(function (error) {
    console.log('Error: ' + JSON.stringify(error));
    if (error instanceof OfficeExtension.Error) {
        console.log('Debug info: ' + JSON.stringify(error.debugInfo));
    }
});


        }
        $("#selectLib").change(function () {
            $("#loader").css("display", "block");
            $scope.LibraryID = $(this).children(":selected").attr("id");
            if (document.getElementById("selectLib").contains(document.getElementById("sel")))
                document.getElementById("selectLib").removeChild(document.getElementById("sel"));
            var getLibsDocs = $scope.GetLibrariesDocuments();
            getLibsDocs.then(function (response) {
                document.getElementById("btnInsert").disabled = true;
                $scope.LibrariesDocuments = response.snippets;
                $("#loader").css("display", "none");
            });


        });
        $("#btnInsertDoc").click(function () {
            Insertdocument();
        });
        $("#btnToken").click(function () {
            ShowTokenDialog();
        });
        $scope.PrependSelectOption = function () {
            if (!document.getElementById("selectLib").contains(document.getElementById("sel")))
                $('#selectLib').prepend($('<option id="sel">select</option>'));
        }
        $scope.Refresh = function () {
            ShowAccountName();
            var checkToken = $scope.CheckUserToken();
            checkToken.then(function (response) {
                if (response.result == false)
                    ShowTokenDialog();
                else {
                    $(".fadeMe").css("display", "none");
                    $("#loader").css("display", "block");
                    $scope.PopulateDocsAssociatedLibsIDS();
                    var getlibs = $scope.GetUserLibraries();
                    getlibs.then(function (response) {
                        for (var i = 0; i < response.data.length; i++) {
                            if (response.data[i].account_id == localStorage.getItem("accountID")) {
                                $scope.UserLibraries = response.data[i].user_libraries;
                                break;
                            }
                        }
                        if ($scope.DocsAssociatedLibraries != null) {
                            $scope.UserLibraries = $scope.UserLibraries.filter(function (emp) {
                                return $scope.DocsAssociatedLibraries.indexOf(emp.id) !== -1;
                            });
                        }
                        if (!$scope.LibraryExists())
                            $scope.PrependSelectOption();
                        var getLibsDocs = $scope.GetLibrariesDocuments();
                        getLibsDocs.then(function (response) {
                            document.getElementById("btnInsert").disabled = true;
                            $scope.LibrariesDocuments = response.snippets;
                            $("#loader").css("display", "none");
                        });




                    });
                }
            });


        }

    });



})();


