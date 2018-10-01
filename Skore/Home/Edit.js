var myCtrl = ['$scope', 'AngularServices', function ($scope, AngularServices) {

    // The initialize function must be run each time a new page is loaded
    Office.initialize = function (reason) {

    };
    $(document).ready(function () {
        $("#saveBtn").click(function () {
            //Office.context.ui.messageParent($("#textEditor").val());
            Office.context.ui.messageParent(JSON.stringify( { "val": $("#textEditor").val() }));
        });
    });


}];

app.controller("myCtrl", myCtrl);
