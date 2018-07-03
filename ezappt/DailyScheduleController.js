
var myCtrl = ['$scope', 'AngularServices', function ($scope, AngularServices) {

    var staffID = getQueryStringValue("staffID");
    var userID = getQueryStringValue("userID");
    var grouped = [];
    $scope.allAppts = [];
    $scope.pickedDateAppts = [];
    var editApptDialog;
    var editApptDialogUrl = DeploymentHost + "editAppt.html?staffID=" + staffID + "&userID=" + userID;
    var editApptDialogUrlStringified = "";
    $(document).ready(function () {
        $("#datepicker1").datepicker({
            defaultDate: "0d",
            dateFormat: "m/d/yy",
            onSelect: function () {
                getAllAppts();
                
            }
        });
        $("#datepicker1").datepicker("setDate", "0d");
    });
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
        AngularServices.GET("GetAppointments", staffID).then(function (data) {
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