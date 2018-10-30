
var myCtrl = ['$scope', 'AngularServices', function ($scope, AngularServices) {

    //https://anoka-wcf.ezsoftco.com/WCFEzapptJsonService.svc/CreateSchedule?staffSchedJson={"StaffId": 102, "StartDate": "10/04/2018", "EndDate": "10/04/2018", "StartTime": "09:00", "EndTime": "20:00", "LocationId": 1, "SelectedDays": "2,3,4,5,6", }
    $scope.locations = [];
    var staffID = getQueryStringValue("staffID");
    var userID = getQueryStringValue("userID");
    // Event Handlers
    $(document).ready(function () {
        $(function () {
            $('input[name="datetimes"]').daterangepicker({
                timePicker: true,
                timePicker24Hour: true,
                timePickerIncrement: 30,
                locale: {
                    format: 'MM-DD-YYYY HH:mm'
                },
            });
        });

        AngularServices.GET("GetAllStaffLocations", staffID).then(function (data) {
            $scope.locations = data.GetAllStaffLocationsResult;
        });
        $("#btnSave").click(saveAvTimes);
    });

    function saveAvTimes() {
        var SD = $('#datetimes').data('daterangepicker').startDate.format('MM-DD-YYYY');
        var ED = $('#datetimes').data('daterangepicker').endDate.format('MM-DD-YYYY');
        var ST = $('#datetimes').data('daterangepicker').startDate.format('HH:mm');
        var ET = $('#datetimes').data('daterangepicker').endDate.format('HH:mm');
        var loc = Number($("#locations").find(":selected").attr("id"));

        AngularServices.POST("CreateSchedule",
            {
                "staffSchedJson": { "StaffId": staffID, "StartDate": SD, "EndDate": ED, "StartTime": ST, "EndTime": ET, "LocationId": loc, "SelectedDays": "2,3,4,5,6" }
            }).then(function (data) {
                showNotification("Saved Successfully.");
                Office.context.ui.messageParent(JSON.stringify({ "StartDate": SD, "EndDate": ED, "StartTime": ST, "EndTime": ET }));

            });
    }


}];

app.controller("myCtrl", myCtrl);





