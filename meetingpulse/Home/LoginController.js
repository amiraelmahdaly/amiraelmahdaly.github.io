var myCtrl = ['$scope', 'AngularServices', function ($scope, AngularServices) {

    $scope.Email = "";
    $scope.Password = "";
    //$("#btnLogin").click(Login);
      
    Office.initialize = function (reason) {

    };
    $scope.Login = function () {

      
        var data = {
            "email": $scope.Email,
            "password": $scope.Password
        };
        var headers = {
            "Content-Type": "application/json",
            "Accept": "application/json"
        }

        AngularServices.POST("auth", data, headers).
            then(function (response) {
                switch (response.status) {
                    case 401:
                        document.getElementById("error").innerText = "Invalid credentials. Please check your email and password and retry"
                        break;
                    case 200:
                        SaveUser({
                            "Email": $scope.Email,
                            "Password": $scope.Password,
                            "Token": response.data.result.token
                        });
                        Redirect("Meetings.html");
                        break;
                    default:
      
                        break;
                }
            }

            );
    }
   
      

     

}];

app.controller("myCtrl", myCtrl);






