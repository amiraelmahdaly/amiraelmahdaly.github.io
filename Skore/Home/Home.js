var myCtrl = ['$scope', 'AngularServices', function ($scope, AngularServices) {

    // The initialize function must be run each time a new page is loaded
    Office.initialize = function (reason) {};

    var editDialog;
    var uploadDialog;
    var editDialogUrl = DeploymentHost + "Edit.html";
    var uploadDialogUrl = DeploymentHost + "Upload.html";
    $scope.txt = '';
    $(document).ready(function () {
        $("#editBtn").click(function () {
            Office.context.ui.displayDialogAsync(editDialogUrl + '?txt=' + $scope.txt, { height: 50, width: 50, displayInIframe: true },
                function (asyncResult) {
                    editDialog = asyncResult.value;
                    editDialog.addEventHandler(Office.EventType.DialogMessageReceived, processEditMessage);
                }
            );
        });

        $("#uploadBtn").click(function () {
            showDialogs(uploadDialogUrl, uploadDialog);
        });
    });

    function showDialogs(url,dialog) {
        Office.context.ui.displayDialogAsync(url, { height: 50, width: 50, displayInIframe: true },
            function (asyncResult) {
                dialog = asyncResult.value;
                dialog.addEventHandler(Office.EventType.DialogEventReceived, processMessage);
            }
        );
    }
    function processEditMessage(arg) {
        $scope.txt = JSON.parse(arg.message).val;
        $scope.$applyAsync();
        editDialog.close();
    }
   
}];

app.controller("myCtrl", myCtrl);
