﻿/// <reference path="/Scripts/FabricUI/MessageBanner.js" />

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
        $scope.CompanyForms = [];
        $scope.Employees = [];
        $scope.CostEntries = [];



        // data binding Objects for controls
        var FreeUserEntries = {
            //  SelectEntryOption: {  text: "Select Entry", value: "", selected: true, disabled:true },
            TimeEntriesOption: { id: "otpTimeEntries", text: "Time Entries" },
            AbsensesOption: { id: "optAbsences", text: "Absences" },
            NotesOption: { id: "optNotes", text: "Notes" },
            ShiftExtrasOption: { id: "optShiftExtras", text: "Shift Extras" }
        };
        var PremiumUserFeatures = {

            TimeEntriesOption: { id: "otpTimeEntries", text: "Time Entries" },
            AbsensesOption: { id: "optAbsences", text: "Absences" },
            NotesOption: { id: "optNotes", text: "Notes" },
            ShiftExtrasOption: { id: "optShiftExtras", text: "Shift Extras" },
            CompanyFormsOption: { id: "optCompanyForms", text: "Project Forms" },
            EmployeesOption: { id: "OptEmployees", text: "Employees" },
            CostEntriesOption: { id: "optCostEntries", text: "Cost Codes" }


        };

        var FreePremium = "Free";
        var ApiKey = "";
        // the default Page Size (page_size query Param)
        var defaultPageSize = 200;
        var BaseURI = "https://rc.rhumbix.com/public_api/v2/";

        //Event Handlers
        angular.element(document).ready(function () {
            $scope.Initial();
        });
        $scope.Initial = function () {
            ApiKey = getQueryStringValue("ApiKey");
            GetProjects();
            PopulateEntryTypes();

        }
        $("#btnRetrieve").click(function () {
            var selectedEntryType = $('#SelEntries').find(":selected").attr("id");
            var date1 = $("#datepicker1").val();
            var date2 = $("#datepicker2").val();
            var jobIDS = $('#selProjects option:selected').map(function () {
                return this.id
            }).get();
                

            if (selectedEntryType == "OptEmployees")
                GetEmployeesEntriesAndExport();

          else if (date1.trim() == "" || date2.trim() == "") {
                showNotification("Please choose dates!");
                return;
            }

           else if (selectedEntryType == "optAbsences")
                GetAbsencesAndExport(date1, date2);
           else if (jobIDS.length == 0)
               showNotification("Please Choose Project")
            else
                    switch (selectedEntryType) {
                        case "otpTimeEntries":
                            GetTimeKeepingEntriesAndExport(date1, date2, jobIDS);
                            break;
                        case "optNotes":
                            GetNotesEntriesAndExport(date1, date2, jobIDS);
                            break;
                        case "optShiftExtras":
                            GetShiftExtrasEntriesAndExport(date1, date2, jobIDS);
                            break;
                        case "optCompanyForms":
                            GetCompanyFormsEntriesAndExport(date1, date2, jobIDS);
                            break;
                        case "optCostEntries":
                            GetCostEntriesAndExport(jobIDS);
                            break;
                        default:
                            showNotification("Please choose Entry Type!")
                            break;

                    
                }

        });
       

        // Push Entries Recursively into dataOBJ and export it using exportFN Method
        function GetAndExportService(URI, dataOBJ, job_number, exportFN, sheetName, tableName, groupBy) {
            $http.get(URI,
                {
                    headers: { "x-api-key": ApiKey }
                    //"nTkrUJUcCp47VeIKJWNmG52ByfQ8Hbk26iUwFwVZ"
                })
                .then(function (response) {
                    for (var i = 0; i < response.data.results.length; i++) {
                        dataOBJ.push(response.data.results[i]);
                    }
                    if (response.data.next != null)
                        GetAndExportService(response.data.next, dataOBJ, job_number, exportFN, sheetName, tableName, groupBy);
                    else if (dataOBJ.length == 0)
                            showNotification("No Entries Available");
                        else if (exportFN != null) {
                                if (groupBy == null)
                                    exportFN(sheetName, job_number, tableName, dataOBJ);
                                else {
                                    var newDataObj = GroupBy(dataOBJ, groupBy);
                                    for (var key in newDataObj) {
                                        exportFN(sheetName + "-" + key, "", tableName, newDataObj[key]);
                                    }

                                }
                        }
                })
                .catch(function (e) {
                    errorHandler(e);
                });
        }

        // Helpers & Wrappers To Call the Generic Service Method 
        function GetProjects() {
            // Initialization before calling the service
            $scope.Projects = [];
            GetAndExportService(BaseURI + "projects/?page_size=" + defaultPageSize, $scope.Projects, "", null, "", "", null);
        }
        function GetTimeKeepingEntriesAndExport(start_date, end_date, job_numbers) {
        // Initialization before calling the service
        for (var i = 0; i < job_numbers.length; i++) {
            $scope.TimeKeepingEntries = [];
            GetAndExportService(BaseURI + "timekeeping_entries/?start_date=" + start_date + "&end_date=" + end_date + "&page_size=" + defaultPageSize + "&job_number=" + encodeURIComponent(job_numbers[i]), $scope.TimeKeepingEntries, job_numbers[i], ExportEntries, "Time Entries", "TimeEntriesTable", null);
        }
    }
        function GetAbsencesAndExport(start_date, end_date) {
            // Initialization before calling the service
            $scope.Absences = [];
            GetAndExportService(BaseURI + "absences/?start_date=" + start_date + "&end_date=" + end_date + "&page_size=" + defaultPageSize, $scope.Absences, "", ExportEntries, "Absences", "AbsencesTable", null);
        }
        function GetNotesEntriesAndExport(start_date, end_date, job_numbers) {
            // Initialization before calling the service
            for (var i = 0; i < job_numbers.length; i++) {
                $scope.Notes = [];
                GetAndExportService(BaseURI + "notes/?start_date=" + start_date + "&end_date=" + end_date + "&page_size=" + defaultPageSize + "&job_number=" + encodeURIComponent(job_numbers[i]), $scope.Notes, job_numbers[i], ExportEntries, "Notes Entries", "NotesEntriesTable", null);
            }
            }
        function GetShiftExtrasEntriesAndExport(start_date, end_date, job_numbers) {

            // Initialization before calling the service
            for (var i = 0; i < job_numbers.length; i++) {
                $scope.ShiftExtras = [];
                GetAndExportService(BaseURI + "shift_extra_entries/?start_date=" + start_date + "&end_date=" + end_date + "&page_size=" + defaultPageSize + "&job_number=" + encodeURIComponent(job_numbers[i]), $scope.ShiftExtras, job_numbers[i], ExportEntries, "Shift Extras", "ShiftExtrasTable", "entry_name");

            }
           }
        function GetCompanyFormsEntriesAndExport(created_start_date, created_end_date, job_numbers) {

            // Initialization before calling the service
            for (var i = 0; i < job_numbers.length; i++) {
                $scope.CompanyForms = [];
                GetAndExportService(BaseURI + "project_entries/?created_start_date=" + created_start_date + "&created_end_date=" + created_end_date + "&page_size=" + defaultPageSize + "&job_number=" + encodeURIComponent(job_numbers[i]), $scope.CompanyForms, job_numbers[i], ExportEntries, "PF", "CompanyFormsTable", "schema_id");

            }
             }
        function GetEmployeesEntriesAndExport() {
            $scope.Employees = [];
            GetAndExportService(BaseURI + "employees/?page_size=" + defaultPageSize, $scope.Employees, "", ExportEntries, "Employees", "EmployeesTable", null);

        }
        function GetCostEntriesAndExport(job_numbers) {
            // Initialization before calling the service
            for (var i = 0; i < job_numbers.length; i++) {
                $scope.CostEntries = [];
                GetAndExportService(BaseURI + "cost_codes/?page_size=" + defaultPageSize + "&job_number=" + encodeURIComponent(job_numbers[i]), $scope.CostEntries, job_numbers[i], ExportEntries, "Cost Codes", "CostCodesTable", null);

            }
           
        }

        function toType(obj) {
            return ({}).toString.call(obj).match(/\s([a-zA-Z]+)/)[1].toLowerCase()
        }
        function PopulateSelect(id, Options) {
            $.each(Options, function (i, item) {
                $('#' + id).append($('<option>', {
                    value: item.value,
                    text: item.text,
                    id: item.id,
                    disabled: item.disabled,
                    selected: item.selected
                }));
            });
        }
        function PopulateEntryTypes() {
            FreePremium = getQueryStringValue("UserType");
            if (FreePremium == "Free") {
                PopulateSelect("SelEntries", FreeUserEntries);
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
            }
            else {
                $('#selProjects').attr('multiple', 'multiple');
                $('#selProjects').attr('size', '9');
                PopulateSelect("SelEntries", PremiumUserFeatures);
                // Edit dates also here

                $(function () {
                    var dateFormat = "yy-mm-dd",
                      from = $("#datepicker1")
                        .datepicker({
                            maxDate: "0d",
                            dateFormat: "yy-mm-dd"
                        })
                        .on("change", function () {
                            to.datepicker("option", "minDate", getDate(this));
                            var k = getDate(this);
                            var msecsIn90ADay = 86400000 * 90;
                            var endDate = new Date(getDate(this).getTime() + msecsIn90ADay);
                            var now = new Date();
                            now.setHours(0, 0, 0, 0);
                            if (endDate < now)
                                to.datepicker("option", "maxDate", endDate);
                            else
                                to.datepicker("option", "maxDate", now);

                        }),
                      to = $("#datepicker2").datepicker({
                          maxDate: "0d",
                          dateFormat: "yy-mm-dd"
                      })
                      .on("change", function () {
                          from.datepicker("option", "maxDate", getDate(this));
                          var msecsIn90ADay = 86400000 * 90;
                          var startDate = new Date(getDate(this).getTime() - msecsIn90ADay);
                          var now = new Date();
                          now.setHours(0, 0, 0, 0);
                          from.datepicker("option", "minDate", startDate);

                      });

                    function getDate(element) {
                        var date;
                        try {
                            date = $.datepicker.parseDate(dateFormat, element.value);
                        } catch (error) {
                            date = null;
                        }

                        return date;
                    }
                });
            }

        }
        function getQueryStringValue(key) {
            return decodeURIComponent(window.location.search.replace(new RegExp("^(?:.*[&\\?]" + encodeURIComponent(key).replace(/[\.\+\*]/g, "\\$&") + "(?:\\=([^&]*))?)?.*$", "i"), "$1"));
        }
        function GroupBy(arr, property) {
            return arr.reduce(function (memo, x) {
                if (!memo[x[property]]) { memo[x[property]] = []; }
                if (noNull(x))
                    memo[x[property]].push(x);
                return memo;
            }, {});
        }
        function noNull(target) {
            for (var member in target) {
                if (target[member] == null)
                    return false;
            }
            return true;
        }

        // Exporting To Excel
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
                WorkSheetName = WorkSheetName.replace(/[/\\&<>'"]/g, function (m) { return SheetNameEscapeChars[m] });
                // adding worksheet
                var sheet = context.workbook.worksheets.add(WorkSheetName.substring(0, 31));
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
                                    var arr = $.map(item[oldColumns[i]], function (el) {
                                        if (toType(el) == "object")
                                            return JSON.stringify(el);
                                        else if (el == null)
                                            return "";
                                        else
                                            return el;
                                    })
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
    });

})();
