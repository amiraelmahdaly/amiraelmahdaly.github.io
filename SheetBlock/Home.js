Office.initialize = function (reason) {
};
$(document).ready(function () {
    var element = document.querySelector('.ms-MessageBanner');
    messageBanner = new fabric.MessageBanner(element);
    messageBanner.hideBanner();
    localStorage.setItem("email", "");
    Office.context.document.getFilePropertiesAsync(function (asyncResult) {
        var fileUrl = asyncResult.value.url;

        if (fileUrl == "") {
            localStorage.setItem("docStatus", "notSaved");
            localStorage.setItem("fileName", "");

        }
        else {
            var filename = fileUrl.split("/");
            filename = filename[filename.length - 1];
            localStorage.setItem("fileName", filename);
            localStorage.setItem("docStatus", "notValidated");
        }
    });
});
function openNav() {
    document.getElementById("mySidenav").style.width = "250px";
}

function closeNav() {
    document.getElementById("mySidenav").style.width = "0";
}
function hideErrorMessage() {

    setTimeout(function () {
        messageBanner.hideBanner();
    }, 2000);
}
// Helper function for treating errors
function errorHandler(error) {
    showNotification("Error", error);
}
// Helper function for displaying notifications
function showNotification(header, content) {
    $("#notificationHeader").text(header);
    $("#notificationBody").text(content);
    messageBanner.showBanner();
    messageBanner.toggleExpansion();
    hideErrorMessage();
}


// create the module and name it myApp
var myApp = angular.module('myApp', ['ngRoute']);

// configure our routes
myApp.config(function ($routeProvider) {
    $routeProvider

        // route for the home page
        .when('/', {
            templateUrl: 'Views/login.html',
            controller: 'loginController'
        })
        .when('/login', {
            templateUrl: 'Views/login.html',
            controller: 'loginController'
        })
        .when('/changePlan', {
            templateUrl: 'Views/changePlan.html',
            controller: 'changePlanController'
        })

        // route for the about page
        .when('/signup', {
            templateUrl: 'Views/signup.html',
            controller: 'signupController'
        })
        .when('/main', {
            templateUrl: 'Views/main.html',
            controller: 'mainController'
        })
        .when('/certification', {
            templateUrl: 'Views/certification.html',
            controller: 'certificationController'
        })
        .when('/certificationDetails', {
            templateUrl: 'Views/certificationDetails.html',
            controller: 'certificationDetailsController'
        })
        .when('/myAccount', {
            templateUrl: 'Views/myAccount.html',
            controller: 'myAccountController'
        })
        .when('/save', {
            templateUrl: 'Views/save.html',
            controller: 'saveController'
        })
        .when('/share', {
            templateUrl: 'Views/share.html',
            controller: 'shareController'
        })
        .when('/shareCompleted', {
            templateUrl: 'Views/shareCompleted.html',
            controller: 'shareCompletedController'
        })
        .when('/validation', {
            templateUrl: 'Views/validation.html',
            controller: 'validationController'
        })

        // route for the contact page
        .when('/signupEmail', {
            templateUrl: 'Views/signupEmail.html',
            controller: 'signupEmailController'
        });
});

//Create global service 



// create the controller and inject Angular's $scope
myApp.controller('loginController', function ($scope) {
    // create a message to display in our view
    var userName = "admin";
    var password = "admin";
    var input = $("#password");

    // Execute a function when the user releases a key on the keyboard
    input.keyup(function (event) {
        // Cancel the default action, if needed
        event.preventDefault();
        // Number 13 is the "Enter" key on the keyboard
        if (event.keyCode === 13) {
            // Trigger the button element with a click
            $("#signIn").click();
        }
    });

    $("#signIn").click(function () {
        if ($("#userName").val() === userName && $("#password").val() === password)
            window.location.href = '#main';
        else
            showNotification("Invalid username or password", "");
    });

});

myApp.controller('signupController', function ($scope) {
    $("#register").click(function () {
        localStorage.setItem("email", $("#email").val());
        jQuery.validator.setDefaults({
            debug: true,
            success: "valid"
        });
        var form = $("#form");
        form.validate();
        if (form.valid())
            window.location.href = '#signupEmail'
    });

});

myApp.controller('signupEmailController', function ($scope) {
    $("#emailTxt").text(localStorage.getItem("email") + ".");
});


myApp.controller('mainController', function ($scope) {
    localStorage.setItem("page", "");
    Office.context.document.getFilePropertiesAsync(function (asyncResult) {
        var fileUrl = asyncResult.value.url;

        if (fileUrl != "") {
            var filename = fileUrl.split('\\');
            filename = filename[filename.length - 1];
            localStorage.setItem("fileName", filename);
            $("#fileName").text(localStorage.getItem("fileName"));
        }
    });

    switch (localStorage.getItem("docStatus")) {


        case "notCertified":
            $("#certify").removeAttr("disabled");
            break;
        default:
            $("#certify").attr("disabled", "disabled");
            break;
    }
    $("#certify").click(function () {
        if (!($("#certify").is(":disabled"))) {
            window.location.href = "#certification";
            localStorage.setItem("docStatus", "certified");
        }
    });


    $("#validate").click(function () {
        switch (localStorage.getItem("docStatus")) {
            case "notSaved":
                window.location.href = '#save';
                break;
            case "notValidated":
                window.location.href = '#validation';
                break;
            case "notCertified":
                window.location.href = '#validation';
                break;
            case "certified":
                window.location.href = '#certificationDetails';
                localStorage.setItem("path", "main");
                break;
            default:
                window.location.href = '#validation';
                break;
        }
    });
});

myApp.controller('certificationController', function ($scope) {
    setTimeout(function () {
        $("#Loader").css("display", "none");
        $("#display").css("display", "block");
        $("#title").text("Certification Completed");
    }, 2000);

    $("#next").click(function () {
        window.location.href = "#share";
    });
});

myApp.controller('certificationDetailsController', function ($scope) {
    $("#path").attr("href", "#" + localStorage.getItem("path"));

});

myApp.controller('saveController', function ($scope) {
    $("#yesSave").click(function () {

        Word.run(function (context) {

            // Create a proxy object for the document.
            var thisDocument = context.document;

            // Queue a commmand to load the document save state (on the saved property).
            context.load(thisDocument, 'saved');

            // Synchronize the document state by executing the queued commands, 
            // and return a promise to indicate task completion.
            return context.sync().then(function () {

                if (thisDocument.saved === false) {
                    // Queue a command to save this document.
                    localStorage.setItem("docStatus", "notValidated");
                    thisDocument.save();
                    window.location.href = '#validation'
                }
            });
        })
            .catch(function (error) {
                console.log("Error: " + JSON.stringify(error));
                if (error instanceof OfficeExtension.Error) {
                    console.log("Debug info: " + JSON.stringify(error.debugInfo));
                }
            });
    });
    $("#noSave").click(function () {
        window.location.href = '#main'

    });
});

myApp.controller('shareController', function ($scope) {
    $("#send").click(function () {
        window.location.href = '#shareCompleted';
    });
    $("#done").click(function () {
        window.location.href = '#main';
    });
});

myApp.controller('shareCompletedController', function ($scope) {
    $("#done").click(function () {
        window.location.href = '#main';
    });
});


myApp.directive('showtab',
    function () {
        return {
            link: function (scope, element, attrs) {
                element.click(function (e) {
                    e.preventDefault();
                    $(element).tab('show');
                });
            }
        };
    });
myApp.controller('myAccountController', function ($scope) {
    var today = new Date();
    var dd = today.getDate();
    var mm = today.getMonth() + 1; //January is 0!
    var yyyy = today.getFullYear();

    if (dd < 10) {
        dd = '0' + dd
    }

    if (mm < 10) {
        mm = '0' + mm
    }

    today = mm + '/' + dd + '/' + yyyy;
    if (localStorage.getItem("docStatus") == "certified") {

        $("#date").text(today);
        $("#hash").text("0xcctr..");
        $("#yestab").text("Yes");
        $("#docName").text(localStorage.getItem("fileName"));
    }

    $("#next").click(function () {
        window.location.href = '#changePlan';
    });
    function activaTab(tab) {
        $('.nav-tabs a[href="#' + tab + '"]').tab('show');
    };
    if (localStorage.getItem("page") == "plan") {
        activaTab("plan")
    }
    $("#clickRow").click(function () {
        if (localStorage.getItem("docStatus") == "certified")
        {
            window.location.href = '#certificationDetails';
            localStorage.setItem("path", "myAccount");

        }
    });
});

myApp.controller('validationController', function ($scope) {
    setTimeout(function () {
        $("#Loader").css("display", "none");
        $("#display").css("display", "block");
    }, 2000);

    $("#yesValidate").click(function () {
        window.location.href = '#certification';
        localStorage.setItem("docStatus", "certified");

    });
    $("#noValidate").click(function () {
        window.location.href = '#main';
        localStorage.setItem("docStatus", "notCertified");

    });
});


myApp.controller('changePlanController', function ($scope) {
    localStorage.setItem("page", "plan");
    $("#cancel").click(function () {
        window.location.href = '#myAccount';
    });
});
