$(document).ready(function () {
    var element = document.querySelector('.ms-MessageBanner');
    messageBanner = new fabric.MessageBanner(element);
    messageBanner.hideBanner();

});

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

        // route for the about page
        .when('/signup', {
            templateUrl: 'Views/signup.html',
            controller: 'signupController'
        })
        .when('/main', {
            templateUrl: 'Views/main.html',
            controller: 'mainController'
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
    $("#emailTxt").text(localStorage.getItem("email") + ".");
});