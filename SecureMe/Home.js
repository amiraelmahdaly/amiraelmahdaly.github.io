(function () {
  "use strict";

  var messageBanner;

  // The Office initialize function must be run each time a new page is loaded.
  


    var app = angular.module('myApp', ['ngSanitize']);
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


        var mail = "";
        var rawToken = "";
        var restId = '';
        
        var restUrl = '';

        var currentMessageID = "";


        Office.initialize = function (reason) {
            $(document).ready(function () {
                var element = document.querySelector('.ms-MessageBanner');
                messageBanner = new fabric.MessageBanner(element);
                messageBanner.hideBanner();
                loadRestDetails();

            });
        };

        function loadRestDetails() {



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

                    getItemHeadersViaRest();
                } else {
                    rawToken = 'error';
                }
            });





        }
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
                // "data": "{\r\n\"apiKey\" : \"23424-324234-XX3453-ASDA89\",\r\n\"type\" : \"PhishingButtonPressed\",\r\n\"recipient\" : \"" + recipient + "\",\r\n\"headerFound\" : \"true\", \r\n\"headerMessage\" : \"" + headerMessage + "\"" + "}"
                "data": "{\r\n\"apiKey\" : \"23424-324234-XX3453-ASDA89\",\r\n\"type\" : \"PhishingButtonPressed\",\r\n\"recipient\" : \"" + recipient + "\",\r\n\"headerFound\" : " + (headerMessage !== "notFound" ? "\"true\", \r\n\"headerMessage\" : \"" + headerMessage + "\"" : "\"false\"") + "}"
            }
            var x = "";
            $.ajax(settings).done(function (response) {
                console.log(response);
                PopualateData(response);


                $("#spinnerCon").css("display", "none");
            }).fail(errorHandler)
        }

        $("#btnForward").click(function () {
            forwardMessage(currentMessageID, $scope.response.forwardEmail);
        });
        $("#btnMoveToPhishing").click(function () {
           CreatePhishingFolder(currentMessageID);

        });
        function PopualateData(response) {
            response = JSON.parse(response);
            $scope.response = {
                "logoURL": response.LogoURL,
                "essScore": response.EssScore,
                "sections": response.Response,
                "enableForward": (response.EnableForward == "true") ? true : false,
                "forwardEmail": response.ForwardEmail
            };

            var HtmlSections="";
            var sections = $scope.response.sections;
            for (var i = 0; i < sections.length; i++) {
                switch (sections[i].Type.toLowerCase()) {

                    case "seperator":
                        HtmlSections += GenerateSeparator(sections[i].Properties.Lines);
                        break;
                    case "text":
                        HtmlSections += GenerateText(sections[i].Properties.Text, sections[i].Properties.Justify);
                        break;

                    case "link":
                        HtmlSections += GenerateLink(sections[i].Properties.Href, sections[i].Properties.Text, sections[i].Properties.Justify);
                        break;


                    default:
                }
            }
            $("#sections").html(HtmlSections);
            g.set($scope.response.essScore);
            $scope.$apply();
        
        }
        function Request(settings, done, fail) {

            $.ajax(settings).done(done).fail(fail);
        }
        function getItemHeadersViaRest() {

            $.ajax({
                url: restUrl + 'messages/' + restId + '?$select = InternetMessageHeaders',
                method: "GET",
                headers: {
                    "Content-Type": "application/json",
                    'Authorization': 'Bearer ' + rawToken
                }
            }).done(function (item) {
                currentMessageID = item.Id;
                var headersNames = item.InternetMessageHeaders.map(function (el) {
                    return el.Name.toLowerCase();
                });
                var index = $.inArray("x-sn-email-phishing", headersNames);
                if (index == -1) {
                   
                    sendApiData(mail, "notFound");
                  
                }
                else {
                    // ShowBarNotification("progress", "progressIndicator", "Good Job, you caught a phish 👍");
                    showNotification("Message","Good Job, you caught a phish 👍");
                    //event.completed();
                    sendApiData(mail, item.InternetMessageHeaders[index].Value);

                }

                
                  
                }).fail(errorHandler)
        }
        function forwardMessage(messageId,address) {
            var settings = {
                "url": restUrl + 'messages/' + messageId + '/forward',
                "type": "POST",
                "headers": {
                    "Content-Type": "application/json",
                    "Authorization": "Bearer " + rawToken
                },
                "data": "{ 'ToRecipients': [{'EmailAddress': {'Address': '" + address +"'}}]}"

            }

            $.ajax(settings).done(function (response) {
                showNotification("Message","Message is forwarded successfully.")
                console.log(response);
            }).fail(errorHandler)
        }
        function CreatePhishingFolder(messageId) {
            //ShowBarNotification("progress", "progressIndicator", "Add-in is moving the message to phishing folder.");
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
               MoveMessageToFolder(messageId, response.Id);
              
            }).fail(function (error) {
                if (error.status == 409) {
                    var url = restUrl + 'mailFolders'
                    MoveIfExist(url, messageId);
                }
                else
                    //ShowBarNotification("error", "errorMessage", error.responseText);
                    showNotification("error", error.responseText);
            });
        }
        function MoveMessageToFolder(messageId, folderId) {

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
            }).fail(errorHandler)
        }

        function MoveIfExist(url, messageId) {

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
                    MoveIfExist(response['@odata.nextLink'], messageId);
                else {
                    var folderId = response.value[$.inArray("Phishing", foldersNames)].Id;
                    MoveMessageToFolder(messageId, folderId);
                }
            }).fail(errorHandler)
        }


        function GenerateLink(href, text, justify) {

            return "<a href='" + href + "' class='text-" + justify + "' > " + text + "</a > ";
        }

        function GenerateImage(src, alt, justify) {
            return "<img src='" + src + "' alt='" + alt + "' class='text-" + justify + "'/>";
        }

        function GenerateText(text, justify) {
            return "<span class='text-" + justify + "' >" + text + "</span>"; 
        }

        function GenerateSeparator(lines){ 
            var Brs = "";
            for (var i = 0; i < lines; i++) {
                Brs += "<br/>";
            }
            return Brs;
        }
        function ShowBarNotification(key,type,message) {

            Office.context.mailbox.item.notificationMessages.addAsync(key
                , {
                    type: type,
                    message: message
            });
        }
        function BarErrorHandler(error) {
            ShowBarNotification("error", "errorMessage", error.responseText);
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
            showNotification("Error", error.responseText);
         
          
        }
        // Helper function for displaying notifications
        function showNotification(header, content) {
            $("#notificationHeader").text(header);
            $("#notificationBody").text(content);
            messageBanner.showBanner();
            messageBanner.toggleExpansion();
            hideErrorMessage();
        }

        //Loader
        var canvas = document.getElementById('spinner');
        var context = canvas.getContext('2d');
        var start = new Date();
        var lines = 16,
            cW = context.canvas.width,
            cH = context.canvas.height;

        var draw = function () {
            var rotation = parseInt(((new Date() - start) / 1000) * lines) / lines;
            context.save();
            context.clearRect(0, 0, cW, cH);
            context.translate(cW / 2, cH / 2);
            context.rotate(Math.PI * 2 * rotation);
            for (var i = 0; i < lines; i++) {

                context.beginPath();
                context.rotate(Math.PI * 2 / lines);
                context.moveTo(cW / 10, 0);
                context.lineTo(cW / 4, 0);
                context.lineWidth = cW / 30;
                context.strokeStyle = "rgba(0,0,0," + i / lines + ")";
                context.stroke();
            }
            context.restore();
        };
        window.setInterval(draw, 1000 / 30);
    });

})();