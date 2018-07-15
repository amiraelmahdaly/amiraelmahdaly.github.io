var myCtrl = ['$scope', 'AngularServices', function ($scope, AngularServices) {
    $scope.password = "IntelligentAssociate";
    $scope.userName = "Nischal";


    $("#btnLogin").click(Login);
    function Login() {

        if (AnyEmpty($scope.userName, $scope.password)) {
            showNotification("Notification", "Please Supply User Name & Password.");
            return;
        }

        AngularServices.POST("login", {
            "userid": $scope.userName,
            "Password": $scope.password
        }).then(function (data) {
            Redirect("Main.html?userName=" + data.user_name + "&token = " + data.auth_token);
        });

        //AngularServices.POST("SetAppointment",
        //    {
        //    "appointmentJson": { "DateID": 4924, "appointmentid": 185, "category": 1, "client": "Test Sagar", "clientid": 116, "dtEnd": "6/25/2018 10:00:00", "dtStart": "6/22/2018 8:00:00", "isEzapptAppointment": true, "location": "Team A", "locationid": 20010, "notes": "", "service": "Assessment", "serviceid": 20053 },
        //    "staffID": 105,
        //    "userID":122

        //}).then(function (data) {
        //    var x = "";
        //});
    }




}];

app.controller("myCtrl", myCtrl);
