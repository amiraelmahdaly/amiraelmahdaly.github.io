
var myCtrl = ['$scope', 'AngularServices', function ($scope, AngularServices) {

    var rawToken = "";
    var restId = '';
    var restUrl = '';
    var clicked = '';
    $scope.staffID = getQueryStringValue("staffID");
    $scope.userID = getQueryStringValue("userID");
    $scope.userName = getQueryStringValue("userName");
    var grouped = [];
    var categories = {
        '1': 'Active',
        '2': 'No show',
        '3': 'Cancelled',
        '4': 'Completed',
        '5': 'Arrived',
        '6': 'In progress',
        '7': 'Fraud'
    };
    $scope.allAppts = [];
    $scope.allSyncAppts = [];
    $scope.allSyncAvs = [];
    var arrSyncIds = [];
    $scope.pickedDateAppts = [];
    var editApptDialog;
    var editApptDialogUrl = DeploymentHost + "editAppt.html?staffID=" + $scope.staffID + "&userID=" + $scope.userID;
    var editApptDialogUrlStringified = "";
    var CalendarID;
    //AQMkADAwATMwMAItZDY0Ny01MGIzLTAwAi0wMAoARgAAA74W_N2a1ZNEhXH55hIt994HAI-GgP4QBltAvyL41da7HZ4AAAIBBgAAAI-GgP4QBltAvyL41da7HZ4AAfIJe8sAAAA=
    Office.initialize = function (reason) {
        $(document).ready(function () {
            //loadRestDetails();

            getAllAppts();
            $("#btnSyncAppt").click(function () {
                clicked = 'appt';
                AngularServices.GET("GetSyncItems", $scope.staffID).then(function (data) {
                    $scope.allSyncAppts = data.GetSyncItemsResult;
                    if ($scope.allSyncAppts.length > 0) {
                        loadRestDetails();
                    }
                    else
                        showNotification("No Appointments to sync");
                });

            });
            $("#btnSyncAppt").click();
            $("#btnSyncAv").click(function () {
                clicked = 'Av';
                AngularServices.GET("EzapptAvailableDates", $scope.staffID).then(function (data) {
                    $scope.allSyncAvs = data.EzapptAvailableDatesResult;
                    if ($scope.allSyncAvs.length > 0) {
                        loadRestDetails();
                    }
                    else
                        showNotification("No Avialable Times to sync");
                });

            });
            $("#datepicker1").datepicker({
                defaultDate: "0d",
                dateFormat: "m/d/yy",
                onSelect: function () {
                    getAllAppts();

                }
            });
            $("#datepicker1").datepicker("setDate", "0d");
        });
    };

    function outlookSync() {
        GetAllCalendars();
    }


    function RenameCalendar() {
        $.ajax({
            url: restUrl + 'calendars/' + CalendarID,
            method: "PATCH",
            data: '{ "Name": "Ez' + new Date().valueOf() + '" }',
            headers: {
                "Content-Type": "application/json",
                'Authorization': 'Bearer ' + rawToken
            }
        }).done(function (item) {
            CalendarID = item.Id;
            DeleteCalendar();
        }).fail(errorHandler);


    }
    function GetAllCalendars() {
        $.ajax({
            url: restUrl + 'calendars',
            method: "GET",
            headers: {
                "Content-Type": "application/json",
                'Authorization': 'Bearer ' + rawToken
            }
        }).done(function (item) {
            if (CalendarExists(item.value)) {

                var k = 0;
                CreateEvent(k);

            }
            else {
                CreateEzapptCalendar();
            }

        }).fail(errorHandler);

    }
    function CalendarExists(Calendars) {
        for (var i = 0; i < Calendars.length; i++) {
            if (Calendars[i].Name.trim() === "Ezappt") {
                CalendarID = Calendars[i].Id;
                return true;
            }
        }
        return false;

    }
    function DeleteCalendar() {
        $.ajax({
            url: restUrl + 'calendars/' + CalendarID,
            method: "DELETE",
            headers: {
                "Content-Type": "application/json",
                'Authorization': 'Bearer ' + rawToken
            }
        }).done(function (item) {
            console.log(item);
            CreateEzapptCalendar();
        }).fail(errorHandler);
    }
    function DeleteCalendarPrem(ID) {
        $.ajax({
            url: restUrl + '/deletedItems/' + ID,
            method: "DELETE",
            headers: {
                "Content-Type": "application/json",
                'Authorization': 'Bearer ' + rawToken
            }
        }).done(function (item) {
            CreateEzapptCalendar();
        }).fail(errorHandler);
    }
    function CreateEzapptCalendar() {
        $.ajax({
            url: restUrl + 'calendars',
            method: "POST",
            data: '{ "Name": "Ezappt" }',
            headers: {
                "Content-Type": "application/json",
                'Authorization': 'Bearer ' + rawToken
            }
        }).done(function (item) {
            CalendarID = item.Id;

            var k = 0;
            CreateEvent(k);

        }).fail(errorHandler);

    }
    function CreateEvent(k) {
        var apptSynced;
        if (clicked === 'appt')
             apptSynced = '{"Subject": "' + $scope.allSyncAppts[k].client + "  " + $scope.allSyncAppts[k].service + "service  " + categories[$scope.allSyncAppts[k].category] + '", "Categories": ["Purple category"],"Body": {"ContentType": "HTML","Content": ""},"Start": {"DateTime": "' + new Date($scope.allSyncAppts[k].dtStart).toISOString() + '","TimeZone": "Pacific Standard Time"},"End": {"DateTime": "' + new Date($scope.allSyncAppts[k].dtEnd).toISOString() + '","TimeZone": "Pacific Standard Time"},"Attendees": []}';
        else
            apptSynced = '{"Subject": "", "Categories": ["Free"],"Body": {"ContentType": "HTML","Content": ""},"Start": {"DateTime": "' + new Date($scope.allSyncAvs[k].Schedule_Start_Time).toISOString() + '","TimeZone": "Pacific Standard Time"},"End": {"DateTime": "' + new Date($scope.allSyncAvs[k].Schedule_End_Time).toISOString() + '","TimeZone": "Pacific Standard Time"},"Attendees": []}';
        $.ajax({
            'url': restUrl + '/events',
            // 'url': restUrl + 'calendars/' + CalendarID + '/events',
            'type': "POST",
            'data': apptSynced,
            'headers': {
                "Content-Type": "application/json",
                'Authorization': 'Bearer ' + rawToken
            }
        }).done(function (item) {
            console.log(item);
            if (clicked === 'appt') {
                arrSyncIds.push($scope.allSyncAppts[k].appointmentid);
                k++;
                if (k < $scope.allSyncAppts.length)
                    CreateEvent(k);
                else {
                    AngularServices.POST("RemoveApptSyncItemsFromDb", { "apptIdsJson": JSON.stringify(arrSyncIds) }).then(function (data) {

                        showNotification("outlook Appointments sync completed");
                    });
                }
            }
            else {
                k++;
                if (k < 5/*$scope.allSyncAvs.length*/)
                    CreateEvent(k);
                else 
                    showNotification("outlook Available Times sync completed");
            }

        }).fail(errorHandler);


    }



    function loadRestDetails() {




        if (Office.context.mailbox.diagnostics.hostName !== 'OutlookIOS') {
            restId = Office.context.mailbox.convertToRestId(
                Office.context.mailbox.item.itemId,
                Office.MailboxEnums.RestVersion.Beta
            );
        } else {
            restId = Office.context.mailbox.item.itemId;
        }
        restUrl = Office.context.mailbox.restUrl + '/v2.0/me/';


        Office.context.mailbox.getCallbackTokenAsync({ isRest: true }, function (result) {
            if (result.status === "succeeded") {
                rawToken = result.value;
                outlookSync();

            } else {
                rawToken = 'error';
            }
        });





    }
    function ShowEditApptDialog() {

        Office.context.ui.displayDialogAsync(editApptDialogUrlStringified, { height: 70, width: 60, displayInIframe: true },
            function (asyncResult) {
                editApptDialog = asyncResult.value;
                editApptDialog.addEventHandler(Office.EventType.DialogEventReceived, editApptDialogClosed);


            }
        );



    }
    function editApptDialogClosed(arg) {
        getAllAppts();
    }
    function getAllAppts() {
        AngularServices.GET("GetAppointments", $scope.staffID).then(function (data) {
            $scope.allAppts = data.GetAppointmentsResult;
            getPickedAppts($("#datepicker1").val());
            $scope.$applyAsync();
        });
    }


    function getPickedAppts(date) {
        $scope.pickedDateAppts = $scope.allAppts.filter(function (value) { return value.dtStart.indexOf(date) >= 0 })
    }
    $scope.$on('ngRepeatFinished', function (ngRepeatFinishedEvent) {
        $(".clickable-row").click(function () {
            var appointmentID = Number($(this).attr("id"));
            var appt = $scope.allAppts.filter(function (obj) {
                return obj.appointmentid == appointmentID;
            });
            var x = "";
            editApptDialogUrlStringified = editApptDialogUrl + "&appt=" + encodeURIComponent(JSON.stringify(appt[0]));
            ShowEditApptDialog();

        });
    });

}];

app.controller("myCtrl", myCtrl);