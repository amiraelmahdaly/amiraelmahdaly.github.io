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
        $scope.ShiftExtras = [];
        $scope.TimeKeepingEntries = [];
        $scope.Absences = [];
        $scope.Notes = [];

        // the default Page Size (page_size query Param)
        var defaultPageSize = 200;
        var BaseURI = "https://rc.rhumbix.com/public_api/v2/";

        $scope.Initial = function () {
            //  GetProjects();
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


            if ($("#datepicker1").val().trim() == "" || $("#datepicker2").val().trim() == "") {
                showNotification("Please choose dates!");
                return;
            }

            if ($('#SelEntries').find(":selected").attr("id") == "optAbsences")
                GetAbsencesAndExport($("#datepicker1").val(), $("#datepicker2").val());
            else
                if ($('#selProjects').find(":selected").attr("id") == "optUnselected")
                    showNotification("Please Choose a project First");
                else
                    switch ($('#SelEntries').find(":selected").attr("id")) {
                        case "otpTimeEntries":
                            GetTimeKeepingEntriesAndExport($("#datepicker1").val(), $("#datepicker2").val(), $('#selProjects').find(":selected").attr("id"));
                            break;

                        case "optNotes":
                            GetNotesEntriesAndExport($("#datepicker1").val(), $("#datepicker2").val(), $('#selProjects').find(":selected").attr("id"));
                            break;

                        case "optShiftExtras":
                            GetShiftExtrasEntriesAndExport($("#datepicker1").val(), $("#datepicker2").val(), $('#selProjects').find(":selected").attr("id"));
                            break;

                        default:
                            showNotification("Please choose Entry Type!")
                            break;






                    }
     

        });
        $("#btnValidate").click(function () {
            GetProjects();
        });


        // Services

        // Push Entries Recursively into dataOBJ and export it using exportFN Method
        function GetAndExportService(URI, dataOBJ, job_number, exportFN, sheetName, tableName , groupBy) {
            $http.get(URI,
                {
                    headers: { "x-api-key": $("#txtApiKey").val() }
                    //"nTkrUJUcCp47VeIKJWNmG52ByfQ8Hbk26iUwFwVZ"
                })
                .then(function (response) {
                    for (var i = 0; i < response.data.results.length; i++) {
                        dataOBJ.push(response.data.results[i]);
                    }
                    if (response.data.next != null)
                        GetAndExportService(response.data.next, dataOBJ, job_number, exportFN);
                    else
                        if (exportFN != null) {
                            if (groupBy == null)
                                exportFN(sheetName, job_number, tableName, dataOBJ);
                            else {
                                var newDataObj = GroupBy(dataOBJ, groupBy);
                                for (var key in newDataObj) {
                                    exportFN(sheetName+"-"+key, "", tableName, newDataObj[key]);
                                }
                            }

                        }
                        
                }).catch(function (e) {
                    errorHandler(e);
                });
        }

        // Helpers & Wrappers To Call the Generic Service Method 
        function GetProjects() {
            // Initialization before calling the service
            $scope.Projects = [];
            GetAndExportService(BaseURI + "projects/?page_size=" + defaultPageSize, $scope.Projects, "", null, "", "",null);
        }
        function GetTimeKeepingEntriesAndExport(start_date, end_date, job_number) {
            // Initialization before calling the service
            $scope.TimeKeepingEntries = [];
            GetAndExportService(BaseURI + "timekeeping_entries/?start_date=" + start_date + "&end_date=" + end_date + "&page_size=" + defaultPageSize + "&job_number=" + job_number, $scope.TimeKeepingEntries, job_number, ExportEntries, "Time Entries", "TimeEntriesTable",null);
        }
        function GetAbsencesAndExport(start_date, end_date) {
            // Initialization before calling the service
            $scope.Absences = [];
            GetAndExportService(BaseURI + "absences/?start_date=" + start_date + "&end_date=" + end_date + "&page_size=" + defaultPageSize, $scope.Absences, "", ExportEntries, "Absences", "AbsencesTable",null);
        }
        function GetNotesEntriesAndExport(start_date, end_date, job_number) {
            // Initialization before calling the service
            $scope.Notes = [];
            GetAndExportService(BaseURI + "notes/?start_date=" + start_date + "&end_date=" + end_date + "&page_size=" + defaultPageSize + "&job_number=" + job_number, $scope.Notes, job_number, ExportEntries, "Notes Entries", "NotesEntriesTable",null);
        }
        function GetShiftExtrasEntriesAndExport(start_date, end_date, job_number) {
            // Initialization before calling the service
            $scope.ShiftExtras = [];
            GetAndExportService(BaseURI + "shift_extra_entries/?start_date=" + start_date + "&end_date=" + end_date + "&page_size=" + defaultPageSize + "&job_number=" + job_number, $scope.ShiftExtras, job_number, ExportEntries, "Shift Extras", "ShiftExtrasTable","entry_name");
        }
        function toType(obj) {
            return ({}).toString.call(obj).match(/\s([a-zA-Z]+)/)[1].toLowerCase()
        }

        // group 
        function GroupBy(arr, property) {
            return arr.reduce(function (memo, x) {
                if (!memo[x[property]]) { memo[x[property]] = []; }
                memo[x[property]].push(x);
                return memo;
            }, {});
        }

        // Exporting
        // Generic Entries Export.
        function ExportEntries(sheetName, job_number, tableName, Entries) {
            Excel.run(function (context) {
                if (Entries.length == 0) {
                    showNotification("No Entries available")
                    return context.sync();
                }

                var SheetNameEscapeChars = { '/': '', '\\': '', '<': '', '>': '', '&': '', '\'': '', '"': '' };
                var JsonObjectsReplaceChars = { '{': '', '}': '', '"': '' };
                job_number = job_number.replace(/[/\\&<>'"]/g, function (m) { return SheetNameEscapeChars[m] });



                // WorkSheet Naming with/without Job Number
                var WorkSheetName = (job_number != "") ? job_number + "-" + sheetName : sheetName;
              
                // adding worksheet
                var sheet = context.workbook.worksheets.add(WorkSheetName);
                // Get Entry Property Names to be the Table Columns
                var oldColumns = Object.getOwnPropertyNames(Entries[0]);
                //editing headers to contain second level keys
                var Columns = oldColumns;
                for (var i = 0; i < oldColumns.length; i++) {
                    if (toType(Entries[0][oldColumns[i]]) == "object")
                        Columns = oldColumns.slice(0, i).concat(Object.getOwnPropertyNames(Entries[0][oldColumns[i]])).concat(oldColumns.slice(i + 1));

                }


                // Dynamically Assign Columns (A1:X1), get X
                var EntriesTable = sheet.tables.add("A1:" + String.fromCharCode(65 + Columns.length - 1) + "1", true /*hasHeaders*/);
                // EntriesTable.name = tableName;
                // Adding Columns to the table
                EntriesTable.getHeaderRowRange().values = [Columns];
                // Getting All Entries Rows
                var rows = Entries.map(
                    function (item) {
                        var it = [];
                        for (var i = 0; i < oldColumns.length; i++) {
                            switch (toType(item[oldColumns[i]])) {
                                case "array":
                                    it.push(item[oldColumns[i]].toString());
                                    break;
                                case "object":
                                    //handling nested object
                                    var arr = $.map(item[oldColumns[i]], function (el) { return el; })
                                    it = it.concat(arr);
                                    break;
                                default:
                                    it.push(item[oldColumns[i]]);
                                    break;
                            }
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


        // Error Handling Region
        // Hiding Error Message
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
            $("#datepicker2").datepicker("setDate", "0d");
        });

      
    });





})();
