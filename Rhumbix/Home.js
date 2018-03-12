/// <reference path="/Scripts/FabricUI/MessageBanner.js" />

(function () {
    "use strict";
    var messageBanner;


    // The initialize function must be run each time a new page is loaded.
    Office.initialize = function (reason) {
        $(document).ready(function () {
            // Initialize the FabricUI notification mechanism and hide it
            var element = document.querySelector('.ms-MessageBanner');
            messageBanner = new fabric.MessageBanner(element);
            messageBanner.hideBanner();

            // Add a click event handler for the highlight button.

        });
    }

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
    app.directive('attrs', function () {
        return {
            link: function (scope, element, attrs) {
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

        // Data Objects
        $scope.Projects = [];
        $scope.TimeKeepingEntries = [];
        $scope.Absences = [];
   
        // the default Page Size (page_size query Param)
        var defaultPageSize = 200;
        var BaseURI = "https://rc.rhumbix.com/public_api/v2/";

        $scope.Initial = function () {

            GetProjects();
           
      
           
        }

        // to be used 
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
        angular.element(document).ready(function () {
            $scope.Initial();
        });


        //Event Handlers

        $("#btnRetrieve").click(function () {
            

            if ($("#datepicker1").val() == "" || $("#datepicker2").val() == "")
            {
                showNotification("Please choose dates!");
                return;
            }
            switch ($('#SelEntries').find(":selected").attr("id")) {
                case "optAbsences":
                    GetAbsencesAndExport($("#datepicker1").val(), $("#datepicker2").val());
                    break;
                case "otpTimeEntries":
                    if ($('#selProjects').find(":selected").attr("id") == "optUnselected")
                        showNotification("Please Choose a project First");
                    else
                        GetTimeKeepingEntriesAndExport($("#datepicker1").val(), $("#datepicker2").val(), $('#selProjects').find(":selected").attr("id"));
                    break;

                default:
                    showNotification("Please choose Entry Type!")
                    break;




            }

        });




        // Services

        // Push Entries Recursively into dataOBJ and export it using exportFN Method
        function GetAndExportService(URI, dataOBJ, job_number, exportFN, sheetName, tableName) {
            $http.get(URI,
                {
                    headers: { "x-api-key": "nTkrUJUcCp47VeIKJWNmG52ByfQ8Hbk26iUwFwVZ" }
                })
                .then(function (response) {
                    for (var i = 0; i < response.data.results.length; i++) {
                        dataOBJ.push(response.data.results[i]);
                    }
                    if (response.data.next != null)
                        GetAndExportService(response.data.next, dataOBJ, job_number, exportFN);
                    else
                        if (exportFN != null) exportFN(sheetName, job_number, tableName, dataOBJ);
                }).catch(function (e) {
                    errorHandler(e);
                });
        }

        // Helpers & Wrappers To Call the Generic Service Method 
        function GetProjects() {
            // Initialization before calling the service

            GetAndExportService(BaseURI + "projects/?page_size=" + defaultPageSize, $scope.Projects, "", null, "", "");
        }
        function GetTimeKeepingEntriesAndExport(start_date, end_date, job_number) {
            // Initialization before calling the service
            $scope.TimeKeepingEntries = [];
            GetAndExportService(BaseURI + "timekeeping_entries/?start_date=" + start_date + "&end_date=" + end_date + "&page_size=" + defaultPageSize + "&job_number=" + job_number, $scope.TimeKeepingEntries, job_number, ExportEntries, "Time Entries", "TimeEntriesTable");
        }
        function GetAbsencesAndExport(start_date, end_date) {
            // Initialization before calling the service
            $scope.Absences = [];
            GetAndExportService(BaseURI + "absences/?start_date=" + start_date + "&end_date=" + end_date + "&page_size=" + defaultPageSize, $scope.Absences, "", ExportEntries, "Absences", "AbsencesTable");
        }

        // Exporting
        // Generic Entries Export.
        function ExportEntries(sheetName, job_number, tableName, Entries) {
            Excel.run(function (context) {
                if (Entries.length == 0)
                {
                    showNotification("No Entries available")
                    return context.sync();
                }
                // WorkSheet Naming with/without Job Number
                var WorkSheetName = (job_number != "") ? job_number + "-" + sheetName : sheetName;
                // adding worksheet
                var sheet = context.workbook.worksheets.add(WorkSheetName);
                // Get Entry Property Names to be the Table Columns
                var Columns = Object.getOwnPropertyNames(Entries[0]);
                // Dynamically Assign Columns (A1:X1), get X
                var EntriesTable = sheet.tables.add("A1:" + String.fromCharCode(65 + Columns.length - 1) + "1", true /*hasHeaders*/);
                EntriesTable.name = tableName;
                // Adding Columns to the table
                EntriesTable.getHeaderRowRange().values = [Columns];
                // Getting All Entries Rows
                var rows = Entries.map(function (item) {
                    var it = [];
                    for (var i = 0; i < Columns.length; i++) {
                        it.push(item[Columns[i]]);
                    }
                    return it;
                });
                // Adding Rows to the table
                EntriesTable.rows.add(null, rows);
                if (Office.context.requirements.isSetSupported("ExcelApi", 1.2)) {
                    // Autofit
                    sheet.getUsedRange().format.autofitColumns();
                    sheet.getUsedRange().format.autofitRows();
                }
                // Activating the Sheet
                sheet.activate();
                return context.sync();
            }).catch(errorHandler);
        }

        function hideErrorMessage() {
            setTimeout(function () {
                messageBanner.hideBanner();
            }, 2000);
        }
        // Helper function for treating errors
        function errorHandler(error) {
            // Always be sure to catch any accumulated errors that bubble up from the Excel.run execution
            switch (error.code) {
                case "ItemAlreadyExists":
                    showNotification("Sheet name already exists");
                   
                    break;
                default:
                    showNotification("Error", error);
                    break;
            }
        
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
            hideErrorMessage();
        }

    
        // Initialize Dt Pickers
        $(function () {
            $("#datepicker1").datepicker({
                minDate: "-30d",
                maxDate: "0d",
                dateFormat: "yy-mm-dd"
        });
    });
        $(function () {
            $("#datepicker2").datepicker({
                minDate: "0d",
                maxDate: "0d",
                defaultDate: "0d",
                dateFormat: "yy-mm-dd"

            });
        });
    });





})();
