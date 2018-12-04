var myCtrl = ['$scope', 'AngularServices', function ($scope, AngularServices) {
    Office.initialize = function (reason) {

    };
    angular.element(document).ready(function () {
        $scope.meetingID = getQueryStringValue("meetingID");
        $scope.BaseURL = BaseURL + "broadcast/" + $scope.meetingID +"/";
        GetPolls($scope.meetingID);
    });

    $("#btnMeeting").click(function () {
        Redirect("Meetings.html");
    });
    $scope.Polls = [];

    function GetPolls(pollID) {
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
        AngularServices.GET("meetings/" + pollID + "/polls/", headers).
            then(function (response) {
                switch (response.status) {
                    case 200:
                        $scope.Polls = response.data.result;
                        break;
                    case 401:
                        AngularServices.RenewTokenOrLogout(GetPolls(pollID));
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
