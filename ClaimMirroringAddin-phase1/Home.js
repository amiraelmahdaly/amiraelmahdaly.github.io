

(function () {
    "use strict";


    // The initialize function must be run each time a new page is loaded.
    Office.initialize = function (reason) {
      
    };


    $(document).ready(function () {


        $(".button").click(function (event) {
            $(this).addClass("onclic", validate(event.target.id));
           
        });


        function validate(id) {
            setTimeout(function () {
                $("#" + id).removeClass("onclic");
                $("#" + id).addClass("validate", callback(id));
            }, 500);
        }
        function callback(id) {
            setTimeout(function () {
                $("#" + id).removeClass("validate");
            }, 500);
        }

    });

})();
