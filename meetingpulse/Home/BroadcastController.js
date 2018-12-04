var myCtrl = ['$scope', 'AngularServices', '$sce', function ($scope, AngularServices, $sce) {
    
        var Link = decodeURIComponent(getQueryStringValue("BroadcastLink"));
        $scope.BroadcastLink = $sce.trustAsResourceUrl(Link);
    Office.initialize = function (reason) {
        Office.context.document.settings.set('BroadcastLink', Link);
        Office.context.document.settings.saveAsync(function (asyncResult) {
            if (asyncResult.status == Office.AsyncResultStatus.Failed) {
                console.log('Settings save failed. Error: ' + asyncResult.error.message);
            } else {
                console.log('Settings saved.');
            }
        });
    };

   
}];

app.controller("myCtrl", myCtrl);






