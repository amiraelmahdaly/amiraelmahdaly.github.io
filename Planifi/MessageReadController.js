var myCtrl = ['$scope', 'AngularServices', function ($scope, AngularServices) {
    $scope.charCount = '';
    $scope.Projects = [];
    var rawToken = "";
    $scope.ProjectPhases = [];
    var serviceRequest = {
        attachmentToken: '',
        ewsUrl: '',
        attachments: [],
        body:''
    };
 
   

  

    $(document).ready(function () {
        GetProjects();
        $("#project").change(function () {
            GetProjectPhases($("#project option:selected").attr("id"));
        });



       

        Office.initialize = function (reason) {
           serviceRequest.ewsUrl = Office.context.mailbox.ewsUrl;
            getAttachmentToken();
            getBodyText();
       
        };
     

    });
  

  

    function getAttachmentToken() {
        if (serviceRequest.attachmentToken.trim() === "")
        {
            Office.context.mailbox.getCallbackTokenAsync(attachmentTokenCallback);
        }
    }

    function getBodyText() {
        var _item = Office.context.mailbox.item;
        var body = _item.body;

        // Get the body asynchronous as text
        body.getAsync(Office.CoercionType.Text, function (asyncResult) {
            if (asyncResult.status !== Office.AsyncResultStatus.Succeeded) {
                // TODO: Handle error
               // showNotification("error", asyncResult.error.message);
               
            }
            else {
                serviceRequest.body = asyncResult.value;
               // showNotification("success", asyncResult.value);
                
            }
        });
    }


    function attachmentTokenCallback(asyncResult, userContext) {
        if (asyncResult.status === "succeeded") {
            // Cache the result from the server.
             serviceRequest.attachmentToken = asyncResult.value;
            token = asyncResult.value;
           // serviceRequest.state = 3;
           // makeServiceRequest();
        } else {
         //   showNotification("Error", "Could not get callback token: " + asyncResult.error.message);
        }
    }

    $scope.makeServiceRequest = function () {

        if (Office.context.mailbox.item.displayReplyForm != undefined) {
            // Format the attachment details for sending.
            for (var i = 0; i < Office.context.mailbox.item.attachments.length; i++) {
                serviceRequest.attachments[i] = JSON.parse(JSON.stringify(Office.context.mailbox.item.attachments[i]));
            }


            // serviceRequest.attachmentToken = "eyJhbGciOiJSUzI1NiIsImtpZCI6IjA2MDBGOUY2NzQ2MjA3MzdFNzM0MDRFMjg3QzQ1QTgxOENCN0NFQjgiLCJ4NXQiOiJCZ0Q1OW5SaUJ6Zm5OQVRpaDhSYWdZeTN6cmciLCJ0eXAiOiJKV1QifQ.eyJuYW1laWQiOiJkNjY4OGZkYS1lZmRkLTQzOGYtODE2Ni0yMjg3ZDJlZjQ0YjNAODRkZjllN2YtZTlmNi00MGFmLWI0MzUtYWFhYWFhYWFhYWFhIiwidmVyIjoiRXhjaGFuZ2UuQ2FsbGJhY2suVjEiLCJhcHBjdHhzZW5kZXIiOiJodHRwczovL2xvY2FsaG9zdDo0NDM3OC9NZXNzYWdlUmVhZC5odG1sQDg0ZGY5ZTdmLWU5ZjYtNDBhZi1iNDM1LWFhYWFhYWFhYWFhYSIsImFwcGN0eCI6IntcIm9pZFwiOlwiMDAwMzAwMDAtZDY0Ny01MGIzLTAwMDAtMDAwMDAwMDAwMDAwXCIsXCJwdWlkXCI6XCIwMDAzMDAwMEQ2NDc1MEIzXCIsXCJzbXRwXCI6XCJhbWlyYS5lbG1haGRhbHlAb3V0bG9vay5jb21cIixcImNpZFwiOlwiQjI5NzEyMDIxRDdBMDUxQlwiLFwic2NvcGVcIjpcIlBhcmVudEl0ZW1JZDpBUU1rQURBd0FUTXdNQUl0WkRZME55MDFNR0l6TFRBd0FpMHdNQW9BUmdBQUE3NFcrTjJhMVpORWhYSDU1aEl0OTk0SEFJL0dnUDRRQmx0QXZ5TDQxZGE3SFo0QUFBSUJEQUFBQUkvR2dQNFFCbHRBdnlMNDFkYTdIWjRBQWZKakpzZ0FBQUE9XCJ9IiwibmJmIjoxNTM2NzkxNTA0LCJleHAiOjE1MzY3OTE4MDQsImlzcyI6IjAwMDAwMDAyLTAwMDAtMGZmMS1jZTAwLTAwMDAwMDAwMDAwMEA4NGRmOWU3Zi1lOWY2LTQwYWYtYjQzNS1hYWFhYWFhYWFhYWEiLCJhdWQiOiIwMDAwMDAwMi0wMDAwLTBmZjEtY2UwMC0wMDAwMDAwMDAwMDAvb3V0bG9vay5vZmZpY2UzNjUuY29tQDg0ZGY5ZTdmLWU5ZjYtNDBhZi1iNDM1LWFhYWFhYWFhYWFhYSJ9.cxcTP_zI-pQVEo65-UeHXckEAezpBJqO8qs2C6ulQnQOoVl3x6xEcF_55TfgW3vMRxVe6T3IEaGl7GHGUTRU11geU1hhMuIzLFdNjs0BrjcllbEe60ik90ruP8XAAmN-r0VPsbQoqe5bTWtOjHPxBX1vy5fRSgDn0DWrwbsU5W_ic-cpCwjmf6ECXI7hYPbM2yqISuEjWL4suY11yfnvgcCVwrv1AHYFF9ZkPi3I9auXcGg6IsNkK1aKqih_Lg5ts92MD0Oqd1lEbu6JVOP9EIxpIeSICG4r64r_aAuOvHyoKS7QmfTzxS4Jdo1A2YGGtkhUpKZIWZJwtWquedAMzQ";
            AngularServices.POST("EmailAPI/API/Messages/UploadMessageAndToken", serviceRequest).
                then(function (response) {
                    console.log("success");
                });
        }
        else {
            Office.context.mailbox.item.bcc.addAsync(['jj@planifi.net']);
            ProcessSubject();
          

         

        }
       
    };


  

        

    function getSubject() {
        Office.context.mailbox.item.subject.getAsync(
            function (asyncResult) {
                if (asyncResult.status == Office.AsyncResultStatus.Failed) {
                    //write(asyncResult.error.message);
                }
                else {
                    // Successfully got the subject, display it.
                    setSubject(asyncResult.value);
                }
            });
    }
    function ProcessSubject() {
        getSubject();
    }
    function setSubject(oldSubject) {
     

        var subject = oldSubject + " #Project:" + $('#project').find(":selected").text() + ";Phase:" + $('#phase').find(":selected").text() + "#"

        Office.context.mailbox.item.subject.setAsync(
            subject,
            { asyncContext: { var1: 1, var2: 2 } },
            function (asyncResult) {
                if (asyncResult.status == Office.AsyncResultStatus.Failed) {
                    write(asyncResult.error.message);
                }
                else {
                    // Successfully set the subject.
                    // Do whatever appropriate for your scenario
                    // using the arguments var1 and var2 as applicable.
                }
            });
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