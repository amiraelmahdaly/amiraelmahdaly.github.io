
var myCtrl = ['$scope', 'AngularServices', function ($scope, AngularServices) {

    var staffID = getQueryStringValue("staffID");
    var grouped = [];
    $scope.allAppts = [];
    $scope.pickedDateAppts = [];

    $(document).ready(function () {
        getAllAppts();
        $("#datepicker1").datepicker({
            defaultDate: "0d",
            dateFormat: "m/d/yy",
            onSelect: function (date) {
                getPickedAppts(date);
                $scope.$applyAsync();
            }
        });
        $("#datepicker1").datepicker("setDate", "0d");
    });

    function getAllAppts() {
        AngularServices.Request("GetAppointments", staffID).then(function (data) {
            $scope.allAppts = data.GetAppointmentsResult;
            getPickedAppts($("#datepicker1").val());
            $scope.$applyAsync();
        });
    }

    function getPickedAppts(date) {
        $scope.pickedDateAppts = $scope.allAppts.filter(function (value) { return value.dtStart.indexOf(date) >= 0 })
    }
    function GroupBy(arr, property) {
        return arr.reduce(function (memo, x) {
            //var date = Date.parse(x[property].substring(6, x[property].length - 2));
            var date = eval("new " + x[property].substring(1, x[property].length - 7)+")");
            var day = date.getMonth() + 1 + "/" + date.getDate() + "/" + date.getFullYear();
            var day1 = date.toLocaleDateString().toString();
            var startTime = date.toLocaleTimeString().substring(-3);
            if (!memo[day]) { memo[day] = []; }
            if (noNull(x)) {
                x.start = startTime;
                memo[day].push(x);
            }
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

}];

app.controller("myCtrl", myCtrl);