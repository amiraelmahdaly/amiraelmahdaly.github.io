
var myCtrl = ['$scope', 'AngularServices', function ($scope, AngularServices) {


    // Variables 
    $scope.lastAppointment = {};
    $scope.ClientInfo = {};

    // Event Handlers
    $(document).ready(function () {
        AngularServices.GET("GetAllClients").then(function (data) {
            FillAutoCompleteWidget(data.GetAllClientsResult);
        });
    });

    function FillAutoCompleteWidget(Clients) {
        $('#tags').autocomplete({
                source: function (request, response) {
                    var re = $.ui.autocomplete.escapeRegex(request.term);
                    var matcher = new RegExp("^" + re, "i");
                    response($.grep(($.map(Clients, function (c, i) {
                        return {
                            label: c.lastName + "," + c.firstName,
                            value: c.lastName + "," + c.firstName,
                            id: c.clientID
                        };
                    })), function (item) {
                        return matcher.test(item.label);
                    }))

            },
                select: function (event, ui) {
             $("#tags").val(ui.item.label); 
                    AngularServices.GET("GetClientForm", ui.item.id).then(function (data) {
                  
                        $scope.ClientInfo = data.GetClientFormResult;
                    });
                    AngularServices.GET("GetClientLastAppointmet", ui.item.id).then(function (data) {
                        $scope.lastAppointment = data.GetClientLastAppointmetResult;
                        $("#clientInfo").css("display","block")
                     //   $scope.$applyAsync();
                    });
                
                return false;
            }

            });
    }
 





}];

app.controller("myCtrl", myCtrl);


 

   
