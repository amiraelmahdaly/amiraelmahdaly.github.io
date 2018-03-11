/// <reference path="/Scripts/FabricUI/MessageBanner.js" />


(function () {
    "use strict";

    // var messageBanner;

    // The initialize function must be run each time a new page is loaded.
    Office.initialize = function (reason) {
      
    };


    $(document).ready(function () {

        function AddPreamble() {
            Word.run(function (context) {
                var pars = context.document.getSelection().paragraphs;
                pars.load();
                return context.sync().then(function () {
                 
                    Claim.Preamble.Style =pars.items[0].style;
                    Claim.Preamble.LineSpacing = pars.items[0].lineSpacing;
             
                    return context.sync();
                })
            }).catch(function (error) {
                console.log(error);
                if (error instanceof OfficeExtension.Error) {
                    console.log("Debug info: " + JSON.stringify(error.debugInfo));
                }
            });


        }
        function AddStep() {
            Word.run(function (context) {
                var IsStart = false;
                var pars = context.document.getSelection().paragraphs;
                pars.load();
                return context.sync().then(function () {
                    for (var i = 0; i < pars.items.length; i++) {
                        if (i == 0) IsStart = true;
                        else IsStart = false;

                        Claim.Steps.push({
                            "Text": pars.items[i].text, "Style": pars.items[i].style, "LineSpacing": pars.items[i].lineSpacing, "IsStart": IsStart,
                            "FirstLineIndent" : pars.items[i].firstLineIndent, "LeftIndent" : pars.items[i].leftIndent

                        });
                    }

                    return context.sync();
                })
            }).catch(function (error) {
                console.log(error);
                if (error instanceof OfficeExtension.Error) {
                    console.log("Debug info: " + JSON.stringify(error.debugInfo));
                }
            });

        }
        function RemoveSelectivelyIfExists(text) {
            if (text.split(" ")[0].toLowerCase() == "selectively")
                return text.substring(text.indexOf(" ") + 1, text.length);
            else return text;
        }
        function RemoveIng(text) {
            //text = RemoveSelectivelyIfExists(text);
            var verb = "";
            var comma = "";
            var selectively = false;
            if (text.split(" ")[0].toLowerCase() == "selectively"){
                verb = text.split(" ")[1];
                selectively = true;
            }
            else {
                verb = text.split(" ")[0];
                selectively = false;
            }

            if (verb.slice(-1) == ",") comma = ",";
            var verbWithoutIng = GetInfinitiveVerbFromDB(verb.toLowerCase());

            if(verbWithoutIng == null)
                verbWithoutIng = nlp.verb(verb).conjugate().infinitive;
            verbWithoutIng = verbWithoutIng + comma + " ";
            if (!selectively) {
                var textWithoutVerb = text.substring(text.indexOf(" ") + 1, text.length);
                return verbWithoutIng + textWithoutVerb;
            }
            else if(selectively) {
                var textWithoutVerb = text.substring(text.indexOf(" ") + verb.length+2, text.length);
                return text.split(" ")[0] +" " + verbWithoutIng + textWithoutVerb;
            }
        }


        function trimVerb(verb) {
            if (verb.slice(-1) == ",") return verb.slice(0, -1);
            return verb;
        }
        function GetInfinitiveVerbFromDB(verb) {

            verb = trimVerb(verb);
            for (var i = 0; i < VerbsDB.length; i++) {
                
                if (VerbsDB[i].INGForm == verb)
                    return VerbsDB[i].InfinitiveForm;
                return null;
            }
        }
        function ProcessStep(type,text) {

            switch (type) {
                case "Device":
                    return RemoveIng(text);
                    break;
                case "Apparatus":
                    return "means for " + text;
                    break;
                case "CRM":
                    return RemoveIng(text);
                    break;
              
            }

        }
        var Claim = {
            "Preambles" : {
                "DevicePreamble": [
                    { "Text" : "A device for wireless communication, comprising:",
                        "LineSpacing" : 24,
                        "Style" : "Heading 1,Claim title",
                        "FirstLineIndent" : 0 },
                     { "Text" : "memory; and",
                         "LineSpacing" : 24,
                         "Style" : "Heading 2,Claim body",
                         "FirstLineIndent" : 36 }, 
                         {
                             "Text": "one or more processors coupled to the memory, the memory and the one or more processors configured to:",
                             "LineSpacing": 24,
                             "Style": "Heading 2,Claim body",
                             "FirstLineIndent": 36
                         }],
                "ApparatusPreamble": [{
                    "Text": "An apparatus for wireless communication, comprising:",
                    "LineSpacing": 24,
                    "Style": "Heading 1,Claim title",
                    "FirstLineIndent": 0
                }],
                "CRMPreamble": [{
                    "Text": "A non-transitory computer-readable medium storing one or more instructions for wireless communication, the one or more instructions comprising: ",
                    "LineSpacing": 24,
                    "Style": "Heading 1,Claim title",
                    "FirstLineIndent": 0
                },
                         {
                             "Text": "one or more instructions that, when executed by one or more processors of a device, cause the one or more processors to:",
                             "LineSpacing": 24,
                             "Style": "Heading 2,Claim body",
                             "FirstLineIndent": 36
                         }]
            },
            "Steps": []
        };

        var VerbsDB = [{
            "INGForm": "receiving",
            "InfinitiveForm" : "receive"
        },
        {
            "INGForm": "comparing",
            "InfinitiveForm": "compare"
        }];
        var nlp = window.nlp_compromise;

      
        

    function AddPreambleToWord(Preamble) {
        Word.run(function (context) {
         
            var par;
         
            for (var i = 0; i < Preamble.length; i++) {
      
                par = context.document.body.insertParagraph(Preamble[i].Text, Word.InsertLocation.end);
                par.style = Preamble[i].Style;
                par.firstLineIndent = Preamble[i].FirstLineIndent;
                par.lineSpacing = Preamble[i].LineSpacing;
   

            }
            return context.sync().then(function () {

                return context.sync();
            })
        }).catch(function (error) {
            console.log(error);
            if (error instanceof OfficeExtension.Error) {
                console.log("Debug info: " + JSON.stringify(error.debugInfo));
            }
        });

    }
    function InsertPreamble(type)
    {
          
        switch (type) {
            case "Device":
                AddPreambleToWord(Claim.Preambles.DevicePreamble);
                break;
            case "Apparatus":
                AddPreambleToWord(Claim.Preambles.ApparatusPreamble);
                break;
            case "CRM":
                AddPreambleToWord(Claim.Preambles.CRMPreamble)
                break;




        }

    }






    $(".button").click(function (event) {
        $(this).addClass("onclic", validate(event.target.id));
        if (event.target.id == "btnPreamble") AddPreamble();
        else if (event.target.id == "btnStep") AddStep();
        else if (event.target.id == "btnGenerate") GenerateClaim();
    });


    function GenerateClaim(){
        var step = "";
        var text = "";
        var par;
        var IndentDisplacement = 35;
        Word.run(function (context) {
            if ($("#rdDevice").is(":checked")) {
                for (var i = 0; i < Claim.Preambles.DevicePreamble.length; i++) {
                    par = context.document.body.insertParagraph(Claim.Preambles.DevicePreamble[i].Text, Word.InsertLocation.end);
                    par.style = Claim.Preambles.DevicePreamble[i].Style;
                    par.firstLineIndent = Claim.Preambles.DevicePreamble[i].FirstLineIndent;
                    par.lineSpacing = Claim.Preambles.DevicePreamble[i].LineSpacing;
                }
                for (var i = 0; i < Claim.Steps.length; i++) {
                    text = Claim.Steps[i].Text;
                    if (Claim.Steps[i].IsStart)
                        text = ProcessStep("Device",text);
                    step = context.document.body.insertParagraph(text, Word.InsertLocation.end);
                    step.style = Claim.Steps[i].Style;
                    step.lineSpacing = Claim.Steps[i].LineSpacing;
                    step.leftIndent = Claim.Steps[i].LeftIndent + IndentDisplacement;
                    step.firstLineIndent = Claim.Steps[i].FirstLineIndent;
                    
                    
                }
            }

            else if ($("#rdAppartus").is(":checked")) {
                for (var i = 0; i < Claim.Preambles.ApparatusPreamble.length; i++) {
                    par = context.document.body.insertParagraph(Claim.Preambles.ApparatusPreamble[i].Text, Word.InsertLocation.end);
                    par.style = Claim.Preambles.ApparatusPreamble[i].Style;
                    par.firstLineIndent = Claim.Preambles.ApparatusPreamble[i].FirstLineIndent;
                    par.lineSpacing = Claim.Preambles.ApparatusPreamble[i].LineSpacing;
                }
                for (var i = 0; i < Claim.Steps.length; i++) {
                    text = Claim.Steps[i].Text;
                    if (Claim.Steps[i].IsStart)
                        text = ProcessStep("Apparatus", text);
                    step = context.document.body.insertParagraph(text, Word.InsertLocation.end);
                    step.style = Claim.Steps[i].Style;
                    step.lineSpacing = Claim.Steps[i].LineSpacing;
                    step.leftIndent = Claim.Steps[i].LeftIndent;
                    step.firstLineIndent = Claim.Steps[i].FirstLineIndent;


                }
            }
             
            else if ($("#rdCRM").is(":checked")) {
                for (var i = 0; i < Claim.Preambles.CRMPreamble.length; i++) {
                    par = context.document.body.insertParagraph(Claim.Preambles.CRMPreamble[i].Text, Word.InsertLocation.end);
                    par.style = Claim.Preambles.CRMPreamble[i].Style;
                    par.firstLineIndent = Claim.Preambles.CRMPreamble[i].FirstLineIndent;
                    par.lineSpacing = Claim.Preambles.CRMPreamble[i].LineSpacing;
                }
                for (var i = 0; i < Claim.Steps.length; i++) {
                    text = Claim.Steps[i].Text;
                    if (Claim.Steps[i].IsStart)
                        text = ProcessStep("CRM", text);
                    step = context.document.body.insertParagraph(text, Word.InsertLocation.end);
                    step.style = Claim.Steps[i].Style;
                    step.lineSpacing = Claim.Steps[i].LineSpacing;
                    step.leftIndent = Claim.Steps[i].LeftIndent + IndentDisplacement;
                    step.firstLineIndent = Claim.Steps[i].FirstLineIndent;

                }
            }


            
            return context.sync().then(function () {
               
                return context.sync();
            })
        }).catch(function (error) {
            console.log(error);
            if (error instanceof OfficeExtension.Error) {
                console.log("Debug info: " + JSON.stringify(error.debugInfo));
            }
        });

        
    }
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

    // Initialize the FabricUI notification mechanism and hide it
    //var element = document.querySelector('.ms-MessageBanner');
    //messageBanner = new fabric.MessageBanner(element);
    //messageBanner.hideBanner();

});

 
  

  
// Helper function for displaying notifications
//function showNotification(header, content) {
//    $("#notificationHeader").text(header);
//    $("#notificationBody").text(content);
//    messageBanner.showBanner();
//    messageBanner.toggleExpansion();
//}
})();
