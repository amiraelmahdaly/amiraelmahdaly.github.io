
var myCtrl = ['$scope', 'AngularServices', function ($scope, AngularServices) {


    // Variables 
    $scope.lastAppointment = {};
    $scope.ClientInfo = {};
    $scope.staffID = getQueryStringValue("staffID");
    $scope.userID = getQueryStringValue("userID");
    $scope.clientName = getQueryStringValue("clientName");
    $scope.clientID = getQueryStringValue("clientID");
    $scope.userName = getQueryStringValue("userName");
    $scope.services = [];
    $scope.locations = [];
    $scope.currentDate = "";
    $scope.serviceName = '';
    $scope.serviceTime = '';
    $scope.dateID = '';
    $scope.time = "";
    $scope.notes = "";
    // Event Handlers
    $(document).ready(function () {
        $("#tags").val($scope.clientName);
        AngularServices.GET("GetAllClients").then(function (data) {
            FillAutoCompleteWidget(data.GetAllClientsResult);
            

        });
        AngularServices.GET("GetSatffServices", $scope.staffID).then(function (data) {
            $scope.services = data.GetSatffServicesResult;
        });
        AngularServices.GET("GetAllStaffLocations", $scope.staffID).then(function (data) {
            $scope.locations = data.GetAllStaffLocationsResult;
        });
        $('#datepick').datepicker({
            inline: true,
            dateFormat: "mm-dd-yy"
        });
        $("#btnSearch").click(function () {
            $scope.currentDate = $('#datepick').val();
            $scope.serviceName = $("#services").find(":selected").val();
            $scope.serviceTime = $scope.services[$("#services").find(":selected").index()-1].serviceTime;
            $scope.serviceID = $("#services").find(":selected").attr("id");
            $scope.locationID = $("#locations").find(":selected").attr("id");
            //if ($scope.locationID == undefined)
            //    $scope.locationID = 20010;
            AngularServices.GET("GetAvailableHoursByDate", $scope.currentDate, $scope.staffID, $scope.locationID).then(function (data) {
                $scope.dateID = data.GetAvailableHoursByDateResult.Date_ID;
            });
            AngularServices.GET("GetAvailableHoursByDate1", $scope.currentDate, $scope.staffID, $scope.locationID, $scope.serviceID).then(function (data) {
                $scope.times = data.GetAvailableHoursByDate1Result;
                $("#date").removeAttr("style");
            });
        });
        $("#btnCancel").click(function () {
            $("#confirm").css("display", "none");
            $("#date").removeAttr("style");
        });

        $("#btnConfirm").click(function () {
            var minsToAdd = Number($scope.serviceTime);
            var endTime = new Date(new Date("1970/01/01 " + $scope.time).getTime() + minsToAdd * 60000).toLocaleTimeString('en-UK', { hour: '2-digit', minute: '2-digit', hour12: true });
            var appt = { "DateID": $scope.dateID, "appointmentid": 0, "category": 1, "client": "", "clientid": $scope.clientID, "dtEnd": $scope.currentDate +" " + endTime, "dtStart": $scope.currentDate + " " + $scope.time, "isEzapptAppointment": true, "location": "", "locationid": $scope.locationID, "notes": $scope.notes, "service": "", "serviceid": $scope.serviceID }
            AngularServices.POST("SetAppointment",
                {
                    "appointmentJson": appt,
                    "staffID": $scope.staffID,
                    "userID": $scope.userID

                }).then(function (data) {
                    $("#confirm").css("display", "none");
                    $("#date").css("display", "none");
                    showNotification("Notification", "Booked Successfully");
                    $scope.notes = "";
                });
            
        });
    });
    $scope.apptDetails = function () {
        $scope.time = this.time;
        $("#date").css("display", "none");
        $("#confirm").removeAttr("style");
    }
    function FillAutoCompleteWidget(Clients) {
        $('#tags').autocomplete({
            source: function (request, response) {
                var re = $.ui.autocomplete.escapeRegex(request.term);
                var matcher = new RegExp("^" + re, "i");
                response($.grep(($.map(Clients, function (c, i) {
                    return {
                        label: c.lastName + "," + c.firstName,
                        value: c.lastName + "," + c.firstName,
                        id: c.clientID
                    };
                })), function (item) {
                    return matcher.test(item.label);
                }))

            },
            select: function (event, ui) {
                $("#tags").val(ui.item.label);
                $scope.clientID = ui.item.id;
                return false;
            }

        });
    }






}];

app.controller("myCtrl", myCtrl);





