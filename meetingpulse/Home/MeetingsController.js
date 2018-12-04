var myCtrl = ['$scope', 'AngularServices', function ($scope, AngularServices) {

    angular.element(document).ready(function () {
        GetMeetings();

    });

    $scope.Meetings = [];
    $("#btnLogout").click(function () {
        SaveUser(null);
        Redirect("Login.html");
    });
    function GetMeetings() {
        var User = getCurrentUser();
        var headers = {
            "Content-Type": "application/json",
            "Accept": "application/json",
            "Authorization": "Bearer " + User.Token 
        };
        var data = {
            "email": User.Email,
            "password": User.Password
        };

        AngularServices.GET("meetings", headers).
            then(function (response) {
                switch (response.status) {
                    case 200:
                        $scope.Meetings = response.data.result;
                        break;
                    case 401:
                        AngularServices.RenewTokenOrLogout(GetMeetings);
                        break;
                    default:
                        Redirect("Login.html");
                        break;
                }
            }
            );
    }
}];

app.controller("myCtrl", myCtrl);
