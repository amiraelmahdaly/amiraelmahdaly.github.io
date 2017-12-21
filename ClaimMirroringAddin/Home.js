/// <reference path="/Scripts/FabricUI/MessageBanner.js" />


(function () {
    "use strict";

   // var messageBanner;

    // The initialize function must be run each time a new page is loaded.
    Office.initialize = function (reason) {
        $(document).ready(function () {

           
                $(".button").click(function (event) {
                    $(this).addClass("onclic", validate(event.target.id));
                });

                function validate(id) {
                    setTimeout(function () {
                        $("#" + id).removeClass("onclic");
                        $("#" + id).addClass("validate", callback(id));
                    }, 2250);
                }
                function callback(id) {
                    setTimeout(function () {
                        $("#" + id).removeClass("validate");
                    }, 1250);
                }
           
            // Initialize the FabricUI notification mechanism and hide it
            //var element = document.querySelector('.ms-MessageBanner');
            //messageBanner = new fabric.MessageBanner(element);
            //messageBanner.hideBanner();

        });
    };

  
    // Helper function for displaying notifications
    //function showNotification(header, content) {
    //    $("#notificationHeader").text(header);
    //    $("#notificationBody").text(content);
    //    messageBanner.showBanner();
    //    messageBanner.toggleExpansion();
    //}
})();
