var myCtrl = ['$scope', 'AngularServices', function ($scope, AngularServices) {
    var LoginWithSalesForceDialog;
   
    $scope.Email = ""; //"admin@spekit.co";
    $scope.Password = "";//"!@#Spekit123";

    var headers = {
        'Content-Type': undefined,
        'Cache-Control': 'no-cache'
       
    };
    //$("#btnClearLC").click(function () {

    //    localStorage.clear();
    //});

    $("#btnLoginWithSalesForceSandBox").click(function () {
        LoginWithSalesForce("https://" + subDomain +".spekit.co/api/auth/login/salesforce?sandbox=true");
    });
    $("#btnLoginWithSalesForce").click(function () {
        LoginWithSalesForce("https://" + subDomain + ".spekit.co/api/auth/login/salesforce");
    });
    $("#btnLogin").click(Login);

    function ShowLoginWithSalesForceDialog(url) {

        Office.context.ui.displayDialogAsync(url, { height: 60, width: 60, displayInIframe: true },
            function (asyncResult) {
                LoginWithSalesForceDialog = asyncResult.value;
                LoginWithSalesForceDialog.addEventHandler(Office.EventType.DialogEventReceived, DialogClosed);


            }
        );



    }
    function DialogClosed(arg) {
        Login();
    }
    function LoginWithSalesForce(url) {
        ShowLoginWithSalesForceDialog(url);

    }
    function Login() {


        var data = new FormData();
        data.append('username', $scope.Email);
        data.append('password', $scope.Password);
     
        AngularServices.POST("auth/login", data, headers).
            then(function (response) {
                switch (response.status) {
                    case 200:
                        TrackEvent("Logged in", $scope.Email, {});
                        Redirect("Main.html?user=" + $scope.Email);
                        break;
                    case 403:
                        if (!($scope.Email.trim() == "" && $scope.Password.trim() == ""))
                        showNotification("Error", "You are not Authorized");
                        break;
                    default:
                }
        
            });


        //AngularServices.GET("auth/logout").
        //    then(function (data) {
        //        console.log(data);
        //    }).catch(function (data) {
        //        console.log(data);
        //    });



    }

    $(document).ready(function () {
        Login();
    });


}];

app.controller("myCtrl", myCtrl);
