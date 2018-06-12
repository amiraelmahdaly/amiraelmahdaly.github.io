var myCtrl = ['$scope', 'AngularServices', function ($scope, AngularServices) {
    $scope.Password = "";
    $scope.Email = "";

   
    $("#btnLogin").click(Login);
    function Login() {
        AngularServices.Request("ValidateUserPassword",$scope.Email, $scope.Password).then(function (data) {
            switch (data.ValidateUserPasswordResult.CodeData) {
                case -1:
                    showNotification("Notification", "There’s no staff");
                    break;
                case -2:
                    showNotification("Notification", "Incorrect Password");
                    break;
                case 0:
                    Redirect("DailySchedule.html?staffID=" + data.ValidateUserPasswordResult.staffID)
                    break;

                default:
            }
        });
    }


   
}];

app.controller("myCtrl", myCtrl);
