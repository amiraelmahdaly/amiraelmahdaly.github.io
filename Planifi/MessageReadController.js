var myCtrl = ['$scope', 'AngularServices', function ($scope, AngularServices) {
    $scope.charCount = '';
    $scope.Projects = [];
    $scope.ProjectPhases = [];
    var serviceRequest = {
        attachmentToken: '',
        ewsUrl: '',
        attachments: []
    };
    Office.initialize = function (reason) {
        angular.element(document).ready(function () {
            $("#btnSave").click(function () {
                showNotification("err0");
            });
            GetProjects();
            $("#project").change(function () {
                GetProjectPhases($("#project option:selected").attr("id"));
            });

            serviceRequest.ewsUrl = Office.context.mailbox.ewsUrl;
            getAttachmentToken();

        });
    };

    function getAttachmentToken() {
        if (serviceRequest.attachmentToken === "") {
            Office.context.mailbox.getCallbackTokenAsync(attachmentTokenCallback);
        }
    }


    function attachmentTokenCallback(asyncResult, userContext) {
        if (asyncResult.status === "succeeded") {
            // Cache the result from the server.
            serviceRequest.attachmentToken = asyncResult.value;
            serviceRequest.state = 3;
            makeServiceRequest();
        } else {
            showNotification("Error", "Could not get callback token: " + asyncResult.error.message);
        }
    }

    function makeServiceRequest() {
        // Format the attachment details for sending.
        for (var i = 0; i < Office.context.mailbox.item.attachments.length; i++) {
            serviceRequest.attachments[i] = JSON.parse(JSON.stringify(Office.context.mailbox.item.attachments[i]));
        }
        console.log(serviceRequest);
    }

    function GetProjects() {
        AngularServices.GET("analyzerapi/Dispatcher/Query?incType=GetShortProjectsQuery").
            then(function (data) {
                $scope.Projects = data.data;
                GetProjectPhases($scope.Projects[0].Value);

            });
    }
    function GetProjectPhases(ID) {
        AngularServices.GET("analyzerapi/Dispatcher/Query?incType=GetShortProjectsWPhaseQuery&WBS1=" + ID).
            then(function (data) {
                $scope.ProjectPhases = data.data;

            });
    }



}];

app.controller("myCtrl", myCtrl);