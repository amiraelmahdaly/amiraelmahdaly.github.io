var myCtrl = ['$scope', 'AngularServices', function ($scope, AngularServices) {
    $scope.userName = getQueryStringValue("userName");
    var token = getQueryStringValue("token");





}];

app.controller("myCtrl", myCtrl);
