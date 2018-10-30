
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
    $scope.allSyncAppts = [];
    $scope.allSyncAvs = [];
    var arrSyncIds = [];
    $scope.pickedDateAppts = [];
    var editApptDialog;
    var editApptDialogUrl = DeploymentHost + "editAppt.html?staffID=" + $scope.staffID + "&userID=" + $scope.userID;
    var editAvTimesDialog;
    var editAvTimesDialogUrl = DeploymentHost + "AvTimes.html?staffID=" + $scope.staffID + "&userID=" + $scope.userID;
    var editApptDialogUrlStringified = "";
    var CalendarID;
    var avTimes = [];
    Office.initialize = function (reason) {
        $(document).ready(function () {

            getAllAppts();

            // Hook Controls with events and configure controls.
            $("#btnSync").click(loadRestDetails);
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
                checkForEzappt();
            } else {
                rawToken = 'error';
            }
        });
    }
    function checkForEzappt() {
        $.ajax({
            url: restUrl + 'calendars',
            method: "GET",
            headers: {
                "Content-Type": "application/json",
                'Authorization': 'Bearer ' + rawToken
            }
        }).done(function (item) {
            if (CalendarExists(item.value)) {
                //calendar exists
                getSyncItems();
            }
            else {
                //no calendar
                ShowAvTimesDialog();
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
    function ShowAvTimesDialog() {
        Office.context.ui.displayDialogAsync(editAvTimesDialogUrl, { height: 70, width: 60, displayInIframe: true },
            function (asyncResult) {
                editAvTimesDialog = asyncResult.value;
                editAvTimesDialog.addEventHandler(Office.EventType.DialogMessageReceived, editAvTimesDialogMessageReceived);
            }
        );
    }
    function editAvTimesDialogMessageReceived(arg) {
        var avTimes = JSON.parse(arg.message);
        editAvTimesDialog.close();
        showNotification("Please wait until ezappt calendar is created.");
        CreateEzapptCalendar(avTimes);
    }
    function CreateEzapptCalendar(av) {
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
            avTimes = getEvents(new Date(av.StartDate), new Date(av.EndDate), Number(av.StartTime.replace(':30', '.5').replace(':00', '')), Number(av.EndTime.replace(':30', '.5').replace(':00', '')));
            var k = 0;
            CreateFreeTime(k);

        }).fail(errorHandler);

    }
    function CreateFreeTime(k) {
        $.ajax({
            //'url': restUrl + '/events',
            'url': restUrl + 'calendars/' + CalendarID + '/events',
            'type': "POST",
            'data': avTimes[k],
            'headers': {
                "Content-Type": "application/json",
                'Authorization': 'Bearer ' + rawToken
            }
        }).done(function (item) {
            console.log(item);
            k++;
            if (k < avTimes.length)
                CreateFreeTime(k);
            else {
                showNotification("Please wait while Syncing Appointments.");
                getSyncItems();
            }
        }).fail(errorHandler);


    }
    function DeleteEvents(events, n) {

        if (events.length == 0) {
            //showNotification("All events are now deleted");
            //create event here
            CreateEvent(n);
        }
        else {
            $.ajax({
                url: restUrl + 'events/' + (events.pop()).Id,
                method: "DELETE",
                headers: {
                    "Content-Type": "application/json",
                    'Authorization': 'Bearer ' + rawToken
                }
            }).done(function (item) {
                DeleteEvents(events, n);
            }).fail(errorHandler);
        }
    }
    function getAndDeleteEvents(dtStart, dtEnd, i) {
        $.ajax({
            url: restUrl + 'calendars/' + CalendarID + '/calendarview?startDateTime=' + dtStart + '&endDateTime=' + dtEnd + '&$select=id',
            method: "GET",
            headers: {
                "Content-Type": "application/json",
                'Authorization': 'Bearer ' + rawToken
            }
        }).done(function (item) {
            DeleteEvents(item.value, i);
        }).fail(errorHandler);
    }
    function getSyncItems() {
        AngularServices.GET("GetSyncItems", $scope.staffID).then(function (data) {
            $scope.allSyncAppts = data.GetSyncItemsResult;
            if ($scope.allSyncAppts.length > 0) {
                var k = 0;
                //new Date(new Date(dt).setMinutes(dt.getMinutes() + 30)).toISOString()
                getAndDeleteEvents(new Date(new Date(new Date($scope.allSyncAppts[k].dtStart)).setMinutes(new Date($scope.allSyncAppts[k].dtStart).getMinutes() + 30)).toISOString(), new Date($scope.allSyncAppts[k].dtEnd).toISOString(), k);

            }
            else
                showNotification("No Appointments to sync");
        });
    }
    function CreateEvent(k) {
        apptSynced = '{"Subject": "' + $scope.allSyncAppts[k].client + "  " + $scope.allSyncAppts[k].service + "service  " + categories[$scope.allSyncAppts[k].category] + '", "Categories": ["Purple category"],"Body": {"ContentType": "HTML","Content": ""},"Start": {"DateTime": "' + new Date($scope.allSyncAppts[k].dtStart).toISOString() + '","TimeZone": "Pacific Standard Time"},"End": {"DateTime": "' + new Date($scope.allSyncAppts[k].dtEnd).toISOString() + '","TimeZone": "Pacific Standard Time"},"Attendees": []}';
        $.ajax({
            //'url': restUrl + '/events',
            'url': restUrl + 'calendars/' + CalendarID + '/events',
            'type': "POST",
            'data': apptSynced,
            'headers': {
                "Content-Type": "application/json",
                'Authorization': 'Bearer ' + rawToken
            }
        }).done(function (item) {
            console.log(item);
            arrSyncIds.push($scope.allSyncAppts[k].appointmentid);
            k++;
            if (k < $scope.allSyncAppts.length)
                getAndDeleteEvents(new Date(new Date(new Date($scope.allSyncAppts[k].dtStart)).setMinutes(new Date($scope.allSyncAppts[k].dtStart).getMinutes() + 30)).toISOString(), new Date($scope.allSyncAppts[k].dtEnd).toISOString(), k);
            else {
                AngularServices.POST("RemoveApptSyncItemsFromDb", { "apptIdsJson": JSON.stringify(arrSyncIds) }).then(function (data) {

                showNotification("outlook Appointments sync completed.");
                });

            }

        }).fail(errorHandler);


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
    // Returns an array of dates between the two dates
    function getEvents(startDate, endDate, startTime, endTime) {
        var Dates = getDates(startDate, endDate);
        var Events = [];
        var dt;

        for (var i = 0; i < Dates.length; i++) {
            dt = new Date(Dates[i].toString());
            dt.setHours(dt.getHours() + startTime);
            for (var j = 0; j < (endTime - startTime) * 2; j++) {
                var event = '{"Subject": "", "Categories": ["Free"],"Body": {"ContentType": "HTML","Content": ""},"Start": {"DateTime": "' + new Date(dt).toISOString() + '","TimeZone": "Pacific Standard Time"},"End": {"DateTime": "' + new Date(new Date(dt).setMinutes(dt.getMinutes() + 30)).toISOString() + '","TimeZone": "Pacific Standard Time"},"Attendees": []}';
                Events.push(event);
                dt.setMinutes(dt.getMinutes() + 30);
            }

        }
        return Events;
    }
    function getDates(startDate, endDate) {
        var dates = [],
            currentDate = startDate,
            addDays = function (days) {
                var date = new Date(this.valueOf());
                date.setDate(date.getDate() + days);
                return date;
            };
        while (currentDate <= endDate) {
            dates.push(currentDate);
            currentDate = addDays.call(currentDate, 1);
        }
        return dates;
    };
    $scope.$on('ngRepeatFinished', function (ngRepeatFinishedEvent) {
        $(".clickable-row").click(function () {
            var appointmentID = Number($(this).attr("id"));
            var appt = $scope.allAppts.filter(function (obj) {
                return obj.appointmentid == appointmentID;
            });
            editApptDialogUrlStringified = editApptDialogUrl + "&appt=" + encodeURIComponent(JSON.stringify(appt[0]));
            ShowEditApptDialog();

        });
    });

}];

app.controller("myCtrl", myCtrl);