(function () {
  "use strict";

  var messageBanner;

  // The Office initialize function must be run each time a new page is loaded.
  Office.initialize = function (reason) {

  };


    var app = angular.module('myApp', []);
    app.directive('onFinishRender', function ($timeout) {
        return {
            restrict: 'A',
            link: function (scope, element, attr) {
                if (scope.$last === true) {
                    $timeout(function () {
                        scope.$emit(attr.onFinishRender);

                    });
                }
            }
        }
    });
    app.controller('myCtrl', function ($scope, $http, $compile) {

        var rawToken = "";
        var restId = '';
        var restUrl = '';
        var mail = "";

        $("#btnTest").click(function(){
            loadRestDetails(event);
        });
        $(document).ready(function () {
            var element = document.querySelector('.ms-MessageBanner');
            messageBanner = new fabric.MessageBanner(element);
            messageBanner.hideBanner();
           // loadRestDetails(event);
        });

        //guage
        var opts = {
            lines: 10, // The number of lines to draw
            angle: 0.05, // The length of each line
            lineWidth: 0.4, // The line thickness
            pointer: {
                length: 0.5, // The radius of the inner circle
                strokeWidth: 0.055, // The rotation offset
                color: '#000000' // Fill color
            },
            staticZones: [
                { strokeStyle: "#F03E3E", min: 0, max: 495 }, // Red from 100 to 130
                { strokeStyle: "#FFFFFF", min: 495, max: 499 }, // White Separator
                { strokeStyle: "#FFDD00", min: 500, max: 625 }, // Yellow
                { strokeStyle: "#FFFFFF", min: 625, max: 629 }, // White Separator
                { strokeStyle: "#30B32D", min: 630, max: 800 }, // Green

            ],
            staticLabels: {
                font: "14px open-sans", // Specifies font
                labels: [0, 250, 500, 630, 800], // Print labels at these values
                color: "#000000", // Optional: Label text color
                fractionDigits: 0 // Optional: Numerical precision. 0=round off.
            },
            limitMax: 'false', // If true, the pointer will not go past the end of the gauge
            percentColors: [[0.25, "#FF0000"], [0.50, "#FFFF00"], [1.0, "#009900"]],
            fontSize: 400
        };

        var target = $("#guageCanvas")[0];
        var g = new Gauge(target).setOptions(opts); // create sexy gauge!
            g.maxValue = 800; // set max gauge value
            g.animationSpeed = 20; // set animation speed (32 is default value)
            g.set(300); // set actual value
      
        //end guage
        function loadRestDetails(event) {
           

            mail = Office.context.mailbox.userProfile.emailAddress;
            if (Office.context.mailbox.diagnostics.hostName !== 'OutlookIOS') {
                restId = Office.context.mailbox.convertToRestId(
                    Office.context.mailbox.item.itemId,
                    Office.MailboxEnums.RestVersion.Beta
                );
            } else {
                restId = Office.context.mailbox.item.itemId;
            }
            restUrl = Office.context.mailbox.restUrl + '/v2.0/me/';
            Office.context.mailbox.getCallbackTokenAsync({ isRest: true }, function (result) {
                if (result.status === "succeeded") {
                    rawToken = result.value;
    
                    getItemHeadersViaRest(event);
                } else {
                    rawToken = 'error';
                }
            });

        }
        function sendApiData(recipient, headerMessage) {
            var settings = {
                "async": true,
                "crossDomain": true,
                "url": "https://dev-services.trustsecurenow.com/GonePhishing/",
                "method": "POST",
                "headers": {
                    "Content-Type": "application/json",
                },
                "processData": false,
                "data": "{\r\n\"apiKey\" : \"23424-324234-XX3453-ASDA89\",\r\n\"type\" : \"PhishingButtonPressed\",\r\n\"recipient\" : \"" + recipient + "\",\r\n\"headerFound\" : " + (headerMessage !== "notFound" ? "\"true\", \r\n\"headerMessage\" : \"" + headerMessage + "\"" : "\"false\"") + "}"
            }
            var x = "";
            $.ajax(settings).done(function (response) {
                console.log(response);
            }).fail(function (error) {
                console.log(error);
            });
        }
        function getItemHeadersViaRest(event) {

            $.ajax({
                url: restUrl + 'messages/' + restId + '?$select = InternetMessageHeaders',
                method: "GET",
                headers: {
                    "Content-Type": "application/json",
                    'Authorization': 'Bearer ' + rawToken
                }
            }).done(function (item) {
                var headersNames = item.InternetMessageHeaders.map(function (el) {
                    return el.Name.toLowerCase();
                });
                var index = $.inArray("x-sn-email-phishing", headersNames);
                if (index == -1) {
                    sendApiData(mail, "notFound");
                    forwardMessage(item.Id);
                }
                else {
                    Office.context.mailbox.item.notificationMessages.addAsync("progress", {
                        type: "progressIndicator",
                        message: "Good Job, you caught a phish 👍"
                    });
                    sendApiData(mail, item.InternetMessageHeaders[index].Value);

                }
                CreatePhishingFolder(item.Id, event);
            }).fail(function (error) {
                console.log(error);
                Office.context.mailbox.item.notificationMessages.addAsync("error", {
                    type: "errorMessage",
                    message: error.responseText
                });
            });
        }
        function forwardMessage(messageId) {
            var settings = {
                "url": restUrl + 'messages/' + messageId + '/forward',
                "type": "POST",
                "headers": {
                    "Content-Type": "application/json",
                    "Authorization": "Bearer " + rawToken
                },
                "data": "{ 'ToRecipients': [{'EmailAddress': {'Address': 'amira@codeafterdark.com'}}]}"

            }

            $.ajax(settings).done(function (response) {
                console.log(response);
            }).fail(function (error) {
                console.log(error);
            });
        }
        function CreatePhishingFolder(messageId, event) {
            Office.context.mailbox.item.notificationMessages.addAsync("progress", {
                type: "progressIndicator",
                message: "Add-in is moving the message to phishing folder."
            });
            var settings = {
                "url": restUrl + 'mailFolders',
                "type": "POST",
                "headers": {
                    "Content-Type": "application/json",
                    "Authorization": "Bearer " + rawToken
                },
                "data": "{'DisplayName': 'Phishing'}"
            }

            $.ajax(settings).done(function (response) {
                MoveMessageToFolder(messageId, response.Id, event);
            }).fail(function (error) {
                if (error.status == 409) {
                    var url = restUrl + 'mailFolders'
                    MoveIfExist(url, messageId, event);
                }
                else
                    Office.context.mailbox.item.notificationMessages.addAsync("error", {
                        type: "errorMessage",
                        message: error.responseText
                    });
            });
        }
        function MoveMessageToFolder(messageId, folderId, event) {

            var settings = {
                "url": restUrl + 'messages/' + messageId + '/move',
                "type": "POST",
                "headers": {
                    "Content-Type": "application/json",
                    "Authorization": "Bearer " + rawToken
                },
                "data": "{'DestinationId':'" + folderId + "'}"
            }

            $.ajax(settings).done(function (response) {
            }).fail(function (error) {
                console.log(error);

                Office.context.mailbox.item.notificationMessages.addAsync("error", {
                    type: "errorMessage",
                    message: error.responseText
                });
            });
        }

        function MoveIfExist(url, messageId, event) {

            var settings = {
                "url": url,
                "headers": {
                    "Content-Type": "application/json",
                    "Authorization": "Bearer " + rawToken
                }
            }

            $.ajax(settings).done(function (response) {
                var foldersNames = response.value.map(function (el) {
                    return el.DisplayName;
                });
                if ($.inArray("Phishing", foldersNames) == -1)
                    MoveIfExist(response['@odata.nextLink'], messageId, event);
                else {
                    var folderId = response.value[$.inArray("Phishing", foldersNames)].Id;
                    MoveMessageToFolder(messageId, folderId, event);
                }
            }).fail(function (error) {
                console.log(error);

                Office.context.mailbox.item.notificationMessages.addAsync("error", {
                    type: "errorMessage",
                    message: error.responseText
                });
            });
        }


        // Error Handling Region
        function hideErrorMessage() {

            setTimeout(function () {
                messageBanner.hideBanner();
            }, 2000);
        }
        // Helper function for treating errors
        function errorHandler(error) {
            // Always be sure to catch any accumulated errors that bubble up from the Excel.run execution
            switch (error.code) {
                case "case1":
                    showNotification("error1", "");
                    break;
                default:
                    showNotification("Error", error);
                    break;
            }

            console.log("Error: " + error);
            if (error instanceof OfficeExtension.Error) {
                console.log("Debug info: " + JSON.stringify(error.debugInfo));
            }
        }
        // Helper function for displaying notifications
        function showNotification(header, content) {
            $("#notificationHeader").text(header);
            $("#notificationBody").text(content);
            messageBanner.showBanner();
            messageBanner.toggleExpansion();
            hideErrorMessage();
        }

    });

})();