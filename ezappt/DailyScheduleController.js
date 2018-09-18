
var myCtrl = ['$scope', 'AngularServices', function ($scope, AngularServices) {

    var rawToken = "";
    var restId = '';
    var restUrl = '';
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
            $("#btnSync").click(function () {

                if ($scope.allAppts.length > 0) {
                    loadRestDetails();
                }
                else
                    showNotification("No Appointments to sync");
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


        $.ajax({
            'url': restUrl + '/events',
            // 'url': restUrl + 'calendars/' + CalendarID + '/events',
            'type': "POST",
            'data': '{"Subject": "' + $scope.allAppts[k].client + "  " + $scope.allAppts[k].service + "service  " + categories[$scope.allAppts[k].category] + '", "Categories": ["Purple category"],"Body": {"ContentType": "HTML","Content": ""},"Start": {"DateTime": "' + new Date($scope.allAppts[k].dtStart).toISOString() + '","TimeZone": "Pacific Standard Time"},"End": {"DateTime": "' + new Date($scope.allAppts[k].dtEnd).toISOString() + '","TimeZone": "Pacific Standard Time"},"Attendees": []}',
            'headers': {
                "Content-Type": "application/json",
                'Authorization': 'Bearer ' + rawToken
            }
        }).done(function (item) {
            console.log(item);
            AngularServices.POST("RemoveApptSyncItemsFromDb", { "apptIdsJson": JSON.stringify([$scope.allAppts[k].appointmentid]) }).then(function (data) {
                k++;
                if (k < $scope.allAppts.length)
                    CreateEvent(k);
                else
                    showNotification("outlook sync completed");
            });
            
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

        Office.context.ui.displayDialogAsync(editApptDialogUrlStringified, { height: 60, width: 60, displayInIframe: true },
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