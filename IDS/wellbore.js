
/// <reference path="/Scripts/FabricUI/MessageBanner.js" />


(function () {
    "use strict";


    // The initialize function must be run each time a new page is loaded.
    Office.initialize = function (reason) {
       
    };


    $(document).ready(function () {
        $("#mainCon").attr("style", "display:block;");
        $("#header").hide();
        $("#wellboreCon").hide();
        $("#grid-row1").show();
        TagDocument();
    });


    function TagDocument() {
        Word.run(function (context) {
            context.document.properties.comments = "wellbore.html?" + document.URL.split("wellbore.html?")[1];
            return context.sync().then(function () {
            });
        })
    .catch(function (error) {
        console.log('Error: ' + JSON.stringify(error));
        if (error instanceof OfficeExtension.Error) {
            console.log('Debug info: ' + JSON.stringify(error.debugInfo));
        }
    });
    }

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

        //initializations
        
        $scope.Well = { title: getQueryStringValue("wellTitle"), uid: getQueryStringValue("wellID") };
        $scope.Wellbore = { title: getQueryStringValue("wellboreTitle"), uid: getQueryStringValue("wellboreID") };
        var UserCredentials = {UserName:getQueryStringValue("UserName"), Password: getQueryStringValue("Password") };
        $scope.BHARUNS = [];
        $scope.Surveys = [];
        $scope.Tabular = [];

        // Functions 

       



        angular.element(document).ready(function () {
            $scope.Initial();



        });

        $scope.$on('ngRepeatFinished', function (ngRepeatFinishedEvent) {
        
            $(".btnInsertBHA").unbind().click(function () {
                $("#grid-row1").show();

                var id = $(this).attr('id');
                $scope.GetAndInsertTabular(id);
            });
            $(".btnInsertSurvey").unbind().click(function () {
                $("#grid-row1").show();

                var id = $(this).attr('id');
                InsertSurveysTable(id)
            });

            $("#btnInsertWellData").unbind().click(function () {
                $("#grid-row1").show();
                //InsertWellDataTable();
                InsertWellDataTable();
            });

            $("#btnInsertWellboreData").unbind().click(function () {
                $("#grid-row1").show();

                InsertWellboreDataTable();
            });


        });

        $scope.Initial = function () {
            
            $scope.GetSurveys();
            $scope.GetBHARUN();
            $scope.GetWellData();
            $scope.GetWellboreData();

        }
       
     
        //function InsertSurvey(i) {
        //    InsertSurveysTable(i);
       
        //}
        //function InsertBHA(i) {
        //    InsertBhaRunsTable(i);
            
        //}
    
        function InsertSurveysTable(surveyID) {
            Word.run(function (ctx) {
                var entry=[];
        
                var keys = ["md", "incl", "azi", "tvd", "vertSect", "dispNs", "dispEw", "dls", "typeTrajStation"];
                var survs = [["MD", "Inc", "Azi", "TVD", "VS", "N/S", "E/W", "Dogleg", "Tooltype"]];
                var firstTajectoryStation = $scope.Surveys[surveyID].trajectoryStation;

                for (var i = 0; i < firstTajectoryStation.length ; i++) {
                    entry = [];
                    for (var j = 0; j < survs[0].length; j++) {

                        if (firstTajectoryStation[i].hasOwnProperty(keys[j])) {
                            if (isNumber(firstTajectoryStation[i][keys[j]]["#text"]))
                            
                                entry.push(parseFloat(Math.round(firstTajectoryStation[i][keys[j]]["#text"] * 100) / 100).toFixed(2));
                            
                            else
                                entry.push(firstTajectoryStation[i][keys[j]]["#text"]);

                        }

                        else
                            entry.push("");
                    }
                        survs.push(entry);
             
                }
              

                //var fruitsNonuniform = [["Apple", "red"], ["Banana", "yellow", "long", "mushy"], ["Pear", "green", "oblong"]];
                //var fruitsUnderfilled = [["Apple", "red", "", ""], ["Banana", "yellow", "long", "mushy"], ["Pear", "green", "oblong", ""]];

                // number of rows to insert, number of columns, insert location , and finally the values which is the array itself.
                var table = ctx.document.body.insertTable(survs.length, survs[0].length, "end", survs);
                //    ctx.document.body.insertTable()
                ctx.load(table);
                return ctx.sync().then(function () {
                    table.style = "Grid Table 4 - Accent 5";
                    table.distributeColumns();
                    $("#grid-row1").hide();


                }).catch(function (e) {
                    console.log(e.message);

                });
            });

        }
 
        function PropExists(prop,arr) {
            for (var i = 0; i < arr.length; i++) {
                if (prop == arr[i]) return true;
            }
            return false;
        }
        $scope.GetSurveys = function () {
            $http.get(GetURI("trajectory", '<trajectorys xmlns="http://www.witsml.org/schemas/131" version="1.3.1.1"><trajectory uidWell="'+ $scope.Well.uid +'" uidWellbore="'+ $scope.Wellbore.uid +'" uid="" ><nameWell/><nameWellbore/> <name/><objectGrowing/><mdMn uom="" datum=""/><mdMx uom="" datum=""/><magDeclUsed uom=""/><aziVertSect uom=""/><dispNsVertSectOrig uom=""/><dispEwVertSectOrig uom=""/><trajectoryStation uid=""><dTimStn/><typeTrajStation/> <typeSurveyTool/><md uom=""/><tvd uom=""/><incl uom=""/><azi uom=""/> <dispNs uom=""/><dispEw uom=""/><vertSect uom=""/><dls uom=""/><rateTurn uom=""/><rateBuild uom=""/></trajectoryStation> </trajectory></trajectorys>'),
                {
                    headers: GetHeader(UserCredentials.UserName, UserCredentials.Password)
                })
                .then(function (response) {
                    $scope.Surveys = [];
                    if (GetJson(response.data.response.result).trajectorys.hasOwnProperty("trajectory")) {
                        if ($.isArray(GetJson(response.data.response.result).trajectorys.trajectory))
                            $scope.Surveys = GetJson(response.data.response.result).trajectorys.trajectory;
                        else
                            $scope.Surveys.push(GetJson(response.data.response.result).trajectorys.trajectory)

                    }

                    console.log("done");
                 

                }).catch(function (e) {
                    throw e;
                    console.log(e);
                });

        }
        $scope.GetBHARUN = function () {
            $http.get(GetURI("bhaRun", '<bhaRuns xmlns="http://www.witsml.org/schemas/131" version="1.3.1.1"> <bhaRun uidWell="' + $scope.Well.uid + '" uidWellbore="' + $scope.Wellbore.uid + '" uid=""> <nameWell/> <nameWellbore/> <name/> <tubular uidRef=""/> <dTimStart/> <dTimStop/> <numBitRun/> <numStringRun/> <objectiveBha/> <drillingParams uid=""> <eTimOpBit uom=""/> <mdHoleStart uom=""/> <mdHoleStop uom=""/> <tubular uidRef=""/> <tqOnBotAv uom=""/> <tqOnBotMx uom=""/> <tqOnBotMn uom=""/> <tqOffBotAv uom=""/> <tqDhAv uom=""/> <wtMud uom=""/> <flowratePump uom=""/> <distHold uom=""/> <rpmAv uom=""/> <rpmMx uom=""/> <rpmMn uom=""/> <rpmAvDh uom=""/> <ropAv uom=""/> <ropMx uom=""/> <ropMn uom=""/> <wobAv uom=""/> <wobMx uom=""/> <wobMn uom=""/> <wobAvDh uom=""/> <aziTop uom=""/> <aziBottom uom=""/> <inclStart uom=""/> <inclMx uom=""/> <inclMn uom=""/> <inclStop uom=""/> <presPumpAv uom=""/> <flowrateBit uom=""/> <comments/> </drillingParams> </bhaRun> </bhaRuns> '),
                {
                    headers: GetHeader(UserCredentials.UserName, UserCredentials.Password)
                })
                .then(function (response) {
                    $scope.BHARUNS = [];
                    if (GetJson(response.data.response.result).bhaRuns.hasOwnProperty("bhaRun"))
                    $scope.BHARUNS = GetJson(response.data.response.result).bhaRuns.bhaRun;
                    console.log("BHA Runs is populated")
                    $("#wellboreCon").show();
                    $("#header").show();
                    $("#grid-row1").hide();
                

                }).catch(function (e) {
                    throw e;
                    console.log(e);
                });

        }
        $scope.GetWellData = function () {
            $http.get(GetURI("well", '<wells version="1.3.1.1" xmlns="http://www.witsml.org/schemas/131"> <well uid="' + $scope.Well.uid + '"> <name/> <nameLegal/> <numLicense/> <numGovt/> <dTimLicense/> <field/> <country/> <state/> <county/> <region/> <block/> <timeZone/> <operator/> <operatorDiv/> <pcInterest uom=""/> <numAPI/> <wellDatum uid="" defaultMeasuredDepth="" defaultVerticalDepth="" defaultElevation=""> <name/> <code/> <elevation uom="" datum=""/> </wellDatum> <groundElevation uom="" datum=""/> <waterDepth uom="" datum=""/> <wellLocation uid=""> <wellCRS uidRef=""/> <easting uom=""/> <northing uom=""/> <description/> </wellLocation> </well> </wells>'),
                {
                    headers: GetHeader(UserCredentials.UserName, UserCredentials.Password)
                })
                .then(function (response) {
                    $scope.WellData = GetJson(response.data.response.result).wells.well;
                   
                    console.log("done");


                }).catch(function (e) {
                    throw e;
                    console.log(e);
                });

        }

        

        function InsertWellDataTable() {
            Word.run(function (ctx) {
                var entry = [];

                var keys = ["name", "numLicense", "country", "block", "timeZone"];
                var survs = [["Well Data:",""],["Name"],["License Number"],["Country"],["Block"],["Time Zone"]];
                var wellData = $scope.WellData;

  
                for (var i = 0; i < keys.length; i++) {
                    if (wellData.hasOwnProperty(keys[i]))
                        survs[i + 1].push(wellData[keys[i]]["#text"]);
                    else
                        survs[i + 1].push("");
                }

                


                //var fruitsNonuniform = [["Apple", "red"], ["Banana", "yellow", "long", "mushy"], ["Pear", "green", "oblong"]];
                //var fruitsUnderfilled = [["Apple", "red", "", ""], ["Banana", "yellow", "long", "mushy"], ["Pear", "green", "oblong", ""]];

                // number of rows to insert, number of columns, insert location , and finally the values which is the array itself.
                var table = ctx.document.body.insertTable(survs.length, survs[0].length, "end", survs);
                //    ctx.document.body.insertTable()
                ctx.load(table);
                return ctx.sync().then(function () {
                    table.style = "Grid Table 4 - Accent 5";
                    table.distributeColumns();
                    $("#grid-row1").hide();


                }).catch(function (e) {
                    console.log(e.message);

                });
            });

        }

        $scope.GetWellboreData = function () {
            $http.get(GetURI("wellbore", '<wellbores  version="1.3.1.1" xmlns="http://www.witsml.org/schemas/131"> <wellbore uidWell="' + $scope.Well.uid + '" uid="' + $scope.Wellbore.uid + '" > <nameWell/> <name/> <parentWellbore uidRef=""/> <suffixAPI/> <numGovt/> <statusWellbore/> <purposeWellbore/> <typeWellbore/> <mdCurrent uom="" datum=""/> <tvdCurrent uom="" datum=""/> <mdKickoff uom="" datum=""/> <tvdKickoff uom="" datum=""/> <mdPlanned uom="" datum=""/> <tvdPlanned uom="" datum=""/> </wellbore> </wellbores> '),
                {
                    headers: GetHeader(UserCredentials.UserName, UserCredentials.Password)
                })
                .then(function (response) {
                    $scope.WellboreData = GetJson(response.data.response.result).wellbores.wellbore;

                    console.log("done");


                }).catch(function (e) {
                    throw e;
                    console.log(e);
                });

        }



        function InsertWellboreDataTable() {
            Word.run(function (ctx) {
                var entry = [];

                var keys = ["name", "nameWell", "statusWellbore", "purposeWellbore", "typeWellbore","mdPlanned","tvdPlanned"];
                var survs = [["Wellbore Data:", ""], ["Name"], ["Well Name"], ["Status"], ["Purpose"], ["Type"], ["md (Planned) (m)"], ["tvd (Planned) (m)"]];
                var wellboreData = $scope.WellboreData;


                for (var i = 0; i < keys.length; i++) {
                    if (wellboreData.hasOwnProperty(keys[i]))
                        survs[i + 1].push(wellboreData[keys[i]]["#text"]);
                    else
                        survs[i + 1].push("");
                }



               

                //var fruitsNonuniform = [["Apple", "red"], ["Banana", "yellow", "long", "mushy"], ["Pear", "green", "oblong"]];
                //var fruitsUnderfilled = [["Apple", "red", "", ""], ["Banana", "yellow", "long", "mushy"], ["Pear", "green", "oblong", ""]];

                // number of rows to insert, number of columns, insert location , and finally the values which is the array itself.
                var table = ctx.document.body.insertTable(survs.length, survs[0].length, "end", survs);
             //   var table = context.document.getSelection().insertTable(survs.length, survs[0].length, "afer", survs);
               
                //    ctx.document.body.insertTable()
                ctx.load(table);
                return ctx.sync().then(function () {
               

                    table.style = "Grid Table 4 - Accent 5";
                    table.distributeColumns();
                    $("#grid-row1").hide();


                }).catch(function (e) {
                    console.log(e.message);

                });
            });

        }
        $scope.GetAndInsertTabular = function (bhaID) {
           // var bhaID = "2c90c0c062aee68a0162afb381fb00ee";
          
            $http.get(GetURI("tubular", '<tubulars xmlns="http://www.witsml.org/schemas/131" version="1.3.1.1"> <tubular uidWell="' + $scope.Well.uid + '" uidWellbore="' + $scope.Wellbore.uid + '" uid="' + bhaID + '"> <nameWell/> <nameWellbore/> <name/> <typeTubularAssy/> <tubularComponent uid=""> <typeTubularComp/> <sequence/> <description/> <id uom=""/> <od uom=""/> <odMx uom=""/> <len uom=""/> <numJointStand/> <wtPerLen uom=""/> <grade/> <vendor/> <model/> <bitRecord> <numBit/> <diaBit uom=""/> <manufacturer/> <typeBit/> <codeIADC/> <condFinalInner/> <condFinalOuter/> <condFinalDull/> <condFinalLocation/> <condFinalBearing/> <condFinalGauge/> <condFinalOther/> <condFinalReason/> </bitRecord> <nozzle uid=""> <index/> <diaNozzle uom=""/> </nozzle> </tubularComponent> </tubular> </tubulars> '),
                {
                    headers: GetHeader(UserCredentials.UserName, UserCredentials.Password)
                })
                .then(function (response) {
                    $scope.Tabular = [];
                    if ($.isArray(GetJson(response.data.response.result).tubulars.tubular.tubularComponent))
                        $scope.Tabular = GetJson(response.data.response.result).tubulars.tubular.tubularComponent;
                    else
                        $scope.Tabular.push(GetJson(response.data.response.result).tubulars.tubular.tubularComponent)
                    Word.run(function (ctx) {
                        var entry = [];

                        var keys = ["description", "numJointStand", "len", "od", "id", "typeTubularComp", "vendor"];
                        var survs = [["Component", "Qty", "Length", "OD", "ID", "Type", "Company"]];
                        for (var i = 0; i < $scope.Tabular.length ; i++) {
                            entry = [];
                            for (var j = 0; j < survs[0].length; j++) {
                                if ($scope.Tabular[i].hasOwnProperty(keys[j])) {
                                    if (isNumber($scope.Tabular[i][keys[j]]["#text"]))
                                        entry.push(parseFloat(Math.round($scope.Tabular[i][keys[j]]["#text"] * 100) / 100).toFixed(2));

                                    else
                                        entry.push($scope.Tabular[i][keys[j]]["#text"]);

                                }

                                else
                                    entry.push("");
                            }
                            survs.push(entry);

                        }

                        //var fruitsNonuniform = [["Apple", "red"], ["Banana", "yellow", "long", "mushy"], ["Pear", "green", "oblong"]];
                        //var fruitsUnderfilled = [["Apple", "red", "", ""], ["Banana", "yellow", "long", "mushy"], ["Pear", "green", "oblong", ""]];

                        // number of rows to insert, number of columns, insert location , and finally the values which is the array itself.
                        var table = ctx.document.body.insertTable(survs.length, survs[0].length, "end", survs);
                        //    ctx.document.body.insertTable()
                        ctx.load(table);
                        return ctx.sync().then(function () {
                            table.style = "Grid Table 4 - Accent 5";
                            table.distributeColumns();
                            $("#grid-row1").hide();


                        }).catch(function (e) {
                            console.log(e.message);

                        });
                    });

                    console.log("done");


                }).catch(function (e) {
                    throw e;
                    console.log(e);
                });

        }

       
    
    });

})();
