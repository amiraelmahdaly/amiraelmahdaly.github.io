/// <reference path="/Scripts/FabricUI/MessageBanner.js" />


(function () {
    "use strict";

   // var messageBanner;

    // The initialize function must be run each time a new page is loaded.
    Office.initialize = function (reason) {
      
    };


    $(document).ready(function () {
        var element = document.querySelector('.ms-MessageBanner');
        var messageBanner = new fabric.MessageBanner(element);
        messageBanner.hideBanner();
        showNotification("Hello", "Welcome");



        function LogError(error)
        {
            $("#errors").text(error);
        }
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
                showNotification("error", error);
                LogError(error);
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
                showNotification("error", error);
                LogError(error);


                console.log(error);
                if (error instanceof OfficeExtension.Error) {
                    console.log("Debug info: " + JSON.stringify(error.debugInfo));
                }
            });

        }
        function RemoveIng(text) {
            var verbWithoutIng = text.split(" ")[0].slice(0, -3) + " ";
            var textWithoutVerb = text.substring(text.indexOf(" ") + 1, text.length);
            return verbWithoutIng + textWithoutVerb;
                  
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
        function showNotification(header, content) {
            $("#notificationHeader").text(header);
            $("#notificationBody").text(content);
            messageBanner.showBanner();
            messageBanner.toggleExpansion();
        }
        var Claim = {
            "Preamble": { "Style": "", "LineSpacing": "" },
            "Steps": []
        };


        var DevicePreamble = "A device for wireless communication, comprising: memory; and one or more processors coupled to the memory, the memory and the one or more processors configured to:";
        var ApparatusPreamble = "An apparatus for wireless communication, comprising:";
        var CRMPreamble = "A non-transitory computer-readable medium storing one or more instructions for wireless communication, the one or more instructions comprising: one or more instructions that, when executed by one or more processors of a device, cause the one or more processors to:";









        $(".button").click(function (event) {
            $(this).addClass("onclic", validate(event.target.id));
            if (event.target.id == "btnPreamble") AddPreamble();
            else if (event.target.id == "btnStep") AddStep();
            else if (event.target.id == "btnGenerate") GenerateClaim();
        });


        function GenerateClaim(){
          
            Word.run(function (context) {
                if ($("#rdDevice").is(":checked")) {
                    var preamble = context.document.body.insertParagraph(DevicePreamble, Word.InsertLocation.end);
                    preamble.style = Claim.Preamble.Style;
                    preamble.lineSpacing = Claim.Preamble.LineSpacing;
                    var step = "";
                    var text = "";
                    for (var i = 0; i < Claim.Steps.length; i++) {
                        text = Claim.Steps[i].Text;
                        if (Claim.Steps[i].IsStart)
                           text = ProcessStep("Device",text);
                        step = context.document.body.insertParagraph(text, Word.InsertLocation.end);
                        step.style = Claim.Steps[i].Style;
                        step.lineSpacing = Claim.Steps[i].LineSpacing;
                        step.leftIndent = Claim.Steps[i].LeftIndent;
                        step.firstLineIndent = Claim.Steps[i].FirstLineIndent;

                    }
                }

               else if ($("#rdAppartus").is(":checked")) {
                    var preamble = context.document.body.insertParagraph(ApparatusPreamble, Word.InsertLocation.end);
                    preamble.style = Claim.Preamble.Style;
                    preamble.lineSpacing = Claim.Preamble.LineSpacing;
                    var step = "";
                    var text = "";
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
                   var preamble = context.document.body.insertParagraph(CRMPreamble, Word.InsertLocation.end);
                   preamble.style = Claim.Preamble.Style;
                   preamble.lineSpacing = Claim.Preamble.LineSpacing;
                   var step = "";
                   var text = "";
                   for (var i = 0; i < Claim.Steps.length; i++) {
                       text = Claim.Steps[i].Text;
                       if (Claim.Steps[i].IsStart)
                           text = ProcessStep("CRM", text);
                       step = context.document.body.insertParagraph(text, Word.InsertLocation.end);
                       step.style = Claim.Steps[i].Style;
                       step.lineSpacing = Claim.Steps[i].LineSpacing;
                       step.leftIndent = Claim.Steps[i].LeftIndent;
                       step.firstLineIndent = Claim.Steps[i].FirstLineIndent;

                   }
               }

            
                return context.sync().then(function () {
               
                    return context.sync();
                })
            }).catch(function (error) {
                console.log(error);
                LogError(error);

                showNotification("error", error);

                if (error instanceof OfficeExtension.Error) {
                    console.log("Debug info: " + JSON.stringify(error.debugInfo));
                }
            });

        
        }
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

        
    

    });

 
  

  
     //Helper function for displaying notifications
  
})();
