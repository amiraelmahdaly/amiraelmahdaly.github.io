/// <reference path="/Scripts/FabricUI/MessageBanner.js" />


(function () {
    "use strict";

    // var messageBanner;

    // The initialize function must be run each time a new page is loaded.
    Office.initialize = function (reason) {

    };


    $(document).ready(function () {
        // new code
        function getHtml() {
            Office.context.document.getSelectedDataAsync(Office.CoercionType.Html,
                { valueFormat: "unformatted", filterType: "all" },
                function (asyncResult) {
                    var error = asyncResult.error;
                    if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                        write(error.name + ": " + error.message);
                    }
                    else {
                        // Get selected data.
                        //return asyncResult.value;
                        Claim.Steps.push({
                            "html": asyncResult.value

                        });

                        write('Selected data is ' + asyncResult.value);
                    }
                });
        }
        function compineHtml(html1, html2) {
            var el1 = $('<div></div>');
            el1.html(html1);
            var el2 = $('<div></div>');
            el2.html(html2);
            var x = $('.WordSection1', el2).html().replace('<p class="MsoNormal">&nbsp;</p>', '');
            $('.WordSection1', el1).append(x);
            return el1.html();
        }
        function processStepHtml(type,html) {
            var el = $('<div></div>');
            el.html(html);
            var text = $('h2:first-child', el).text();
                switch (type) {
                    case "Device":
                        var newText = ProcessStep(type, text);
                        $('.WordSection1', el).children('h2').each(function () {
                            var margin = (parseInt($(this).css("margin-left")) / 96 + 0.5).toString() + "in";
                            $(this).css("margin-left", margin);
                        });
                        break;
                    case "Apparatus":
                        var newText = ProcessStep(type, text);
                        break;
                    case "CRM":
                        var newText = ProcessStep(type, text);
                        $('.WordSection1', el).children('h2').each(function () {
                            var margin = (parseInt($(this).css("margin-left")) / 96 + 0.5).toString() + "in";
                            $(this).css("margin-left", margin);
                        });
                        break;
            }
                $('h2:first-child', el).text(newText);
                return el.html();
        }
        function writeHtmlData(html) {
            Office.context.document.setSelectedDataAsync(html, { coercionType: Office.CoercionType.Html }, function (asyncResult) {
                if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                    write('Error: ' + asyncResult.error.message);
                }
            });
        }
        function write(message) {
            document.getElementById('message').innerText += message;
        }
        function AddStep() {
            getHtml();
        }
        function RemoveIng(text) {
            var verb = "";
            var comma = "";
            var selectively = false;
            if (text.split(" ")[0].toLowerCase() == "selectively") {
                verb = text.split(" ")[1];
                selectively = true;
            }
            else {
                verb = text.split(" ")[0];
                selectively = false;
            }

            if (verb.slice(-1) == ",") comma = ",";
            var verbWithoutIng = GetInfinitiveVerbFromDB(verb.toLowerCase());

            if (verbWithoutIng == null)
                verbWithoutIng = nlp.verb(verb).conjugate().infinitive;
            verbWithoutIng = verbWithoutIng + comma + " ";
            if (!selectively) {
                var textWithoutVerb = text.substring(text.indexOf(" ") + 1, text.length);
                return verbWithoutIng + textWithoutVerb;
            }
            else if (selectively) {
                var textWithoutVerb = text.substring(text.indexOf(" ") + verb.length + 2, text.length);
                return text.split(" ")[0] + " " + verbWithoutIng + textWithoutVerb;
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
        function ProcessStep(type, text) {

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
            "Preambles": {
                "DevicePreamble": "<html> <head> <meta http-equiv=Content-Type content='text/html; charset=windows-1256'> <meta name=Generator content='Microsoft Word 15 (filtered)'> <style> <!-- /* Font Definitions */ @font-face {font-family:'Cambria Math'; panose-1:2 4 5 3 5 4 6 3 2 4;} @font-face {font-family:Calibri; panose-1:2 15 5 2 2 2 4 3 2 4;} /* Style Definitions */ p.MsoNormal, li.MsoNormal, div.MsoNormal {margin-top:0in; margin-right:0in; margin-bottom:8.0pt; margin-left:0in; line-height:107%; font-size:11.0pt; font-family:'Calibri','sans-serif';} h1 {mso-style-name:'Heading 1\,Claim title'; mso-style-link:'Heading 1 Char\,Claim title Char'; margin:0in; margin-bottom:.0001pt; text-indent:0in; line-height:200%; font-size:12.0pt; font-family:'Times New Roman','serif'; font-weight:normal;} h1.CxSpFirst {mso-style-name:'Heading 1\,Claim titleCxSpFirst'; mso-style-link:'Heading 1 Char\,Claim title Char'; margin:0in; margin-bottom:.0001pt; text-indent:0in; line-height:200%; font-size:12.0pt; font-family:'Times New Roman','serif'; font-weight:normal;} h1.CxSpMiddle {mso-style-name:'Heading 1\,Claim titleCxSpMiddle'; mso-style-link:'Heading 1 Char\,Claim title Char'; margin:0in; margin-bottom:.0001pt; text-indent:0in; line-height:200%; font-size:12.0pt; font-family:'Times New Roman','serif'; font-weight:normal;} h1.CxSpLast {mso-style-name:'Heading 1\,Claim titleCxSpLast'; mso-style-link:'Heading 1 Char\,Claim title Char'; margin:0in; margin-bottom:.0001pt; text-indent:0in; line-height:200%; font-size:12.0pt; font-family:'Times New Roman','serif'; font-weight:normal;} h2 {mso-style-name:'Heading 2\,Claim body'; mso-style-link:'Heading 2 Char\,Claim body Char'; margin:0in; margin-bottom:.0001pt; text-indent:.5in; line-height:200%; font-size:12.0pt; font-family:'Times New Roman','serif'; font-weight:normal;} h2.CxSpFirst {mso-style-name:'Heading 2\,Claim bodyCxSpFirst'; mso-style-link:'Heading 2 Char\,Claim body Char'; margin:0in; margin-bottom:.0001pt; text-indent:.5in; line-height:200%; font-size:12.0pt; font-family:'Times New Roman','serif'; font-weight:normal;} h2.CxSpMiddle {mso-style-name:'Heading 2\,Claim bodyCxSpMiddle'; mso-style-link:'Heading 2 Char\,Claim body Char'; margin:0in; margin-bottom:.0001pt; text-indent:.5in; line-height:200%; font-size:12.0pt; font-family:'Times New Roman','serif'; font-weight:normal;} h2.CxSpLast {mso-style-name:'Heading 2\,Claim bodyCxSpLast'; mso-style-link:'Heading 2 Char\,Claim body Char'; margin:0in; margin-bottom:.0001pt; text-indent:.5in; line-height:200%; font-size:12.0pt; font-family:'Times New Roman','serif'; font-weight:normal;} p.MsoHeader, li.MsoHeader, div.MsoHeader {mso-style-link:'Header Char'; margin:0in; margin-bottom:.0001pt; font-size:11.0pt; font-family:'Calibri','sans-serif';} p.MsoFooter, li.MsoFooter, div.MsoFooter {mso-style-link:'Footer Char'; margin:0in; margin-bottom:.0001pt; font-size:11.0pt; font-family:'Calibri','sans-serif';} p.MsoListParagraph, li.MsoListParagraph, div.MsoListParagraph {margin-top:0in; margin-right:0in; margin-bottom:8.0pt; margin-left:.5in; line-height:107%; font-size:11.0pt; font-family:'Calibri','sans-serif';} p.MsoListParagraphCxSpFirst, li.MsoListParagraphCxSpFirst, div.MsoListParagraphCxSpFirst {margin-top:0in; margin-right:0in; margin-bottom:0in; margin-left:.5in; margin-bottom:.0001pt; line-height:107%; font-size:11.0pt; font-family:'Calibri','sans-serif';} p.MsoListParagraphCxSpMiddle, li.MsoListParagraphCxSpMiddle, div.MsoListParagraphCxSpMiddle {margin-top:0in; margin-right:0in; margin-bottom:0in; margin-left:.5in; margin-bottom:.0001pt; line-height:107%; font-size:11.0pt; font-family:'Calibri','sans-serif';} p.MsoListParagraphCxSpLast, li.MsoListParagraphCxSpLast, div.MsoListParagraphCxSpLast {margin-top:0in; margin-right:0in; margin-bottom:8.0pt; margin-left:.5in; line-height:107%; font-size:11.0pt; font-family:'Calibri','sans-serif';} span.Heading1Char {mso-style-name:'Heading 1 Char\,Claim title Char'; mso-style-link:'Heading 1\,Claim title'; font-family:'Times New Roman','serif';} span.Heading2Char {mso-style-name:'Heading 2 Char\,Claim body Char'; mso-style-link:'Heading 2\,Claim body'; font-family:'Times New Roman','serif';} span.HeaderChar {mso-style-name:'Header Char'; mso-style-link:Header;} span.FooterChar {mso-style-name:'Footer Char'; mso-style-link:Footer;} .MsoChpDefault {font-family:'Calibri','sans-serif';} .MsoPapDefault {margin-bottom:8.0pt; line-height:107%;} /* Page Definitions */ @page WordSection1 {size:8.5in 11.0in; margin:1.0in 1.25in 1.0in 1.25in;} div.WordSection1 {page:WordSection1;} /* List Definitions */ ol {margin-bottom:0in;} ul {margin-bottom:0in;} --> </style> </head> <body lang=EN-US> <div class=WordSection1> <h1 style='margin-left:0in;text-indent:0in'><span dir=LTR></span>A device for wireless communication, comprising:</h1> <h2>memory; and</h2> <h2>one or more processors coupled to the memory, the memory and the one or more processors configured to:</h2>  </div> </body> </html>",
                "ApparatusPreamble": "<html> <head> <meta http-equiv=Content-Type content='text/html; charset=windows-1256'> <meta name=Generator content='Microsoft Word 15 (filtered)'> <style> <!-- /* Font Definitions */ @font-face {font-family:Calibri; panose-1:2 15 5 2 2 2 4 3 2 4;} /* Style Definitions */ p.MsoNormal, li.MsoNormal, div.MsoNormal {margin-top:0in; margin-right:0in; margin-bottom:8.0pt; margin-left:0in; line-height:107%; font-size:11.0pt; font-family:'Calibri','sans-serif';} h1 {mso-style-name:'Heading 1\,Claim title'; mso-style-link:'Heading 1 Char\,Claim title Char'; margin:0in; margin-bottom:.0001pt; text-indent:0in; line-height:200%; font-size:12.0pt; font-family:'Times New Roman','serif'; font-weight:normal;} h1.CxSpFirst {mso-style-name:'Heading 1\,Claim titleCxSpFirst'; mso-style-link:'Heading 1 Char\,Claim title Char'; margin:0in; margin-bottom:.0001pt; text-indent:0in; line-height:200%; font-size:12.0pt; font-family:'Times New Roman','serif'; font-weight:normal;} h1.CxSpMiddle {mso-style-name:'Heading 1\,Claim titleCxSpMiddle'; mso-style-link:'Heading 1 Char\,Claim title Char'; margin:0in; margin-bottom:.0001pt; text-indent:0in; line-height:200%; font-size:12.0pt; font-family:'Times New Roman','serif'; font-weight:normal;} h1.CxSpLast {mso-style-name:'Heading 1\,Claim titleCxSpLast'; mso-style-link:'Heading 1 Char\,Claim title Char'; margin:0in; margin-bottom:.0001pt; text-indent:0in; line-height:200%; font-size:12.0pt; font-family:'Times New Roman','serif'; font-weight:normal;} h2 {mso-style-name:'Heading 2\,Claim body'; mso-style-link:'Heading 2 Char\,Claim body Char'; margin:0in; margin-bottom:.0001pt; text-indent:.5in; line-height:200%; font-size:12.0pt; font-family:'Times New Roman','serif'; font-weight:normal;} h2.CxSpFirst {mso-style-name:'Heading 2\,Claim bodyCxSpFirst'; mso-style-link:'Heading 2 Char\,Claim body Char'; margin:0in; margin-bottom:.0001pt; text-indent:.5in; line-height:200%; font-size:12.0pt; font-family:'Times New Roman','serif'; font-weight:normal;} h2.CxSpMiddle {mso-style-name:'Heading 2\,Claim bodyCxSpMiddle'; mso-style-link:'Heading 2 Char\,Claim body Char'; margin:0in; margin-bottom:.0001pt; text-indent:.5in; line-height:200%; font-size:12.0pt; font-family:'Times New Roman','serif'; font-weight:normal;} h2.CxSpLast {mso-style-name:'Heading 2\,Claim bodyCxSpLast'; mso-style-link:'Heading 2 Char\,Claim body Char'; margin:0in; margin-bottom:.0001pt; text-indent:.5in; line-height:200%; font-size:12.0pt; font-family:'Times New Roman','serif'; font-weight:normal;} p.MsoHeader, li.MsoHeader, div.MsoHeader {mso-style-link:'Header Char'; margin:0in; margin-bottom:.0001pt; font-size:11.0pt; font-family:'Calibri','sans-serif';} p.MsoFooter, li.MsoFooter, div.MsoFooter {mso-style-link:'Footer Char'; margin:0in; margin-bottom:.0001pt; font-size:11.0pt; font-family:'Calibri','sans-serif';} p.MsoListParagraph, li.MsoListParagraph, div.MsoListParagraph {margin-top:0in; margin-right:0in; margin-bottom:8.0pt; margin-left:.5in; line-height:107%; font-size:11.0pt; font-family:'Calibri','sans-serif';} p.MsoListParagraphCxSpFirst, li.MsoListParagraphCxSpFirst, div.MsoListParagraphCxSpFirst {margin-top:0in; margin-right:0in; margin-bottom:0in; margin-left:.5in; margin-bottom:.0001pt; line-height:107%; font-size:11.0pt; font-family:'Calibri','sans-serif';} p.MsoListParagraphCxSpMiddle, li.MsoListParagraphCxSpMiddle, div.MsoListParagraphCxSpMiddle {margin-top:0in; margin-right:0in; margin-bottom:0in; margin-left:.5in; margin-bottom:.0001pt; line-height:107%; font-size:11.0pt; font-family:'Calibri','sans-serif';} p.MsoListParagraphCxSpLast, li.MsoListParagraphCxSpLast, div.MsoListParagraphCxSpLast {margin-top:0in; margin-right:0in; margin-bottom:8.0pt; margin-left:.5in; line-height:107%; font-size:11.0pt; font-family:'Calibri','sans-serif';} span.Heading1Char {mso-style-name:'Heading 1 Char\,Claim title Char'; mso-style-link:'Heading 1\,Claim title'; font-family:'Times New Roman','serif';} span.Heading2Char {mso-style-name:'Heading 2 Char\,Claim body Char'; mso-style-link:'Heading 2\,Claim body'; font-family:'Times New Roman','serif';} span.HeaderChar {mso-style-name:'Header Char'; mso-style-link:Header;} span.FooterChar {mso-style-name:'Footer Char'; mso-style-link:Footer;} .MsoChpDefault {font-family:'Calibri','sans-serif';} .MsoPapDefault {margin-bottom:8.0pt; line-height:107%;} /* Page Definitions */ @page WordSection1 {size:8.5in 11.0in; margin:1.0in 1.25in 1.0in 1.25in;} div.WordSection1 {page:WordSection1;} /* List Definitions */ ol {margin-bottom:0in;} ul {margin-bottom:0in;} --> </style> </head> <body lang=EN-US> <div class=WordSection1> <h1 style='margin-left:0in;text-indent:0in'><span dir=LTR></span>An apparatus for wireless communication, comprising:</h1> </div> </body> </html>",
                "CRMPreamble": "<html> <head> <meta http-equiv=Content-Type content='text/html; charset=windows-1256'> <meta name=Generator content='Microsoft Word 15 (filtered)'> <style> <!-- /* Font Definitions */ @font-face {font-family:Calibri; panose-1:2 15 5 2 2 2 4 3 2 4;} /* Style Definitions */ p.MsoNormal, li.MsoNormal, div.MsoNormal {margin-top:0in; margin-right:0in; margin-bottom:8.0pt; margin-left:0in; line-height:107%; font-size:11.0pt; font-family:'Calibri','sans-serif';} h1 {mso-style-name:'Heading 1\,Claim title'; mso-style-link:'Heading 1 Char\,Claim title Char'; margin:0in; margin-bottom:.0001pt; text-indent:0in; line-height:200%; font-size:12.0pt; font-family:'Times New Roman','serif'; font-weight:normal;} h1.CxSpFirst {mso-style-name:'Heading 1\,Claim titleCxSpFirst'; mso-style-link:'Heading 1 Char\,Claim title Char'; margin:0in; margin-bottom:.0001pt; text-indent:0in; line-height:200%; font-size:12.0pt; font-family:'Times New Roman','serif'; font-weight:normal;} h1.CxSpMiddle {mso-style-name:'Heading 1\,Claim titleCxSpMiddle'; mso-style-link:'Heading 1 Char\,Claim title Char'; margin:0in; margin-bottom:.0001pt; text-indent:0in; line-height:200%; font-size:12.0pt; font-family:'Times New Roman','serif'; font-weight:normal;} h1.CxSpLast {mso-style-name:'Heading 1\,Claim titleCxSpLast'; mso-style-link:'Heading 1 Char\,Claim title Char'; margin:0in; margin-bottom:.0001pt; text-indent:0in; line-height:200%; font-size:12.0pt; font-family:'Times New Roman','serif'; font-weight:normal;} h2 {mso-style-name:'Heading 2\,Claim body'; mso-style-link:'Heading 2 Char\,Claim body Char'; margin:0in; margin-bottom:.0001pt; text-indent:.5in; line-height:200%; font-size:12.0pt; font-family:'Times New Roman','serif'; font-weight:normal;} h2.CxSpFirst {mso-style-name:'Heading 2\,Claim bodyCxSpFirst'; mso-style-link:'Heading 2 Char\,Claim body Char'; margin:0in; margin-bottom:.0001pt; text-indent:.5in; line-height:200%; font-size:12.0pt; font-family:'Times New Roman','serif'; font-weight:normal;} h2.CxSpMiddle {mso-style-name:'Heading 2\,Claim bodyCxSpMiddle'; mso-style-link:'Heading 2 Char\,Claim body Char'; margin:0in; margin-bottom:.0001pt; text-indent:.5in; line-height:200%; font-size:12.0pt; font-family:'Times New Roman','serif'; font-weight:normal;} h2.CxSpLast {mso-style-name:'Heading 2\,Claim bodyCxSpLast'; mso-style-link:'Heading 2 Char\,Claim body Char'; margin:0in; margin-bottom:.0001pt; text-indent:.5in; line-height:200%; font-size:12.0pt; font-family:'Times New Roman','serif'; font-weight:normal;} p.MsoHeader, li.MsoHeader, div.MsoHeader {mso-style-link:'Header Char'; margin:0in; margin-bottom:.0001pt; font-size:11.0pt; font-family:'Calibri','sans-serif';} p.MsoFooter, li.MsoFooter, div.MsoFooter {mso-style-link:'Footer Char'; margin:0in; margin-bottom:.0001pt; font-size:11.0pt; font-family:'Calibri','sans-serif';} p.MsoListParagraph, li.MsoListParagraph, div.MsoListParagraph {margin-top:0in; margin-right:0in; margin-bottom:8.0pt; margin-left:.5in; line-height:107%; font-size:11.0pt; font-family:'Calibri','sans-serif';} p.MsoListParagraphCxSpFirst, li.MsoListParagraphCxSpFirst, div.MsoListParagraphCxSpFirst {margin-top:0in; margin-right:0in; margin-bottom:0in; margin-left:.5in; margin-bottom:.0001pt; line-height:107%; font-size:11.0pt; font-family:'Calibri','sans-serif';} p.MsoListParagraphCxSpMiddle, li.MsoListParagraphCxSpMiddle, div.MsoListParagraphCxSpMiddle {margin-top:0in; margin-right:0in; margin-bottom:0in; margin-left:.5in; margin-bottom:.0001pt; line-height:107%; font-size:11.0pt; font-family:'Calibri','sans-serif';} p.MsoListParagraphCxSpLast, li.MsoListParagraphCxSpLast, div.MsoListParagraphCxSpLast {margin-top:0in; margin-right:0in; margin-bottom:8.0pt; margin-left:.5in; line-height:107%; font-size:11.0pt; font-family:'Calibri','sans-serif';} span.Heading1Char {mso-style-name:'Heading 1 Char\,Claim title Char'; mso-style-link:'Heading 1\,Claim title'; font-family:'Times New Roman','serif';} span.Heading2Char {mso-style-name:'Heading 2 Char\,Claim body Char'; mso-style-link:'Heading 2\,Claim body'; font-family:'Times New Roman','serif';} span.HeaderChar {mso-style-name:'Header Char'; mso-style-link:Header;} span.FooterChar {mso-style-name:'Footer Char'; mso-style-link:Footer;} .MsoChpDefault {font-family:'Calibri','sans-serif';} .MsoPapDefault {margin-bottom:8.0pt; line-height:107%;} /* Page Definitions */ @page WordSection1 {size:8.5in 11.0in; margin:1.0in 1.25in 1.0in 1.25in;} div.WordSection1 {page:WordSection1;} /* List Definitions */ ol {margin-bottom:0in;} ul {margin-bottom:0in;} --> </style> </head> <body lang=EN-US> <div class=WordSection1> <h1 style='margin-left:0in;text-indent:0in'><span dir=LTR></span>A non-transitory computer-readable medium storing one or more instructions for wireless communication, the one or more instructions comprising: </h1> <h2>one or more instructions that, when executed by one or more processors of a device, cause the one or more processors to:</h2> </div> </body> </html>"
            },
            "Steps": []
        };
        var VerbsDB = [{
            "INGForm": "receiving",
            "InfinitiveForm": "receive"
        },
        {
            "INGForm": "comparing",
            "InfinitiveForm": "compare"
        }];
        var nlp = window.nlp_compromise;
        $(".button").click(function (event) {
            $(this).addClass("onclic", validate(event.target.id));
            if (event.target.id == "btnStep") AddStep();
            else if (event.target.id == "btnGenerate") GenerateClaim();
        });
        function GenerateClaim() {
           
          
                if ($("#rdDevice").is(":checked")) {
                    var html = Claim.Preambles.DevicePreamble;
                    for (var i = 0; i < Claim.Steps.length; i++) {
                        var html = compineHtml(html, processStepHtml("Device", Claim.Steps[i].html));
                    }
                    writeHtmlData('<html><head><meta http-equiv=Content-Type content="text/html; charset=windows-1256">' + html + '<p class="MsoNormal">&nbsp;</p></body></html>');
                }

                else if ($("#rdAppartus").is(":checked")) {
                    var html = Claim.Preambles.ApparatusPreamble;
                    for (var i = 0; i < Claim.Steps.length; i++) {
                        var html = compineHtml(html, processStepHtml("Apparatus", Claim.Steps[i].html));
                    }
                    writeHtmlData('<html><head><meta http-equiv=Content-Type content="text/html; charset=windows-1256">' + html + '<p class="MsoNormal">&nbsp;</p></body></html>');
                }

                else if ($("#rdCRM").is(":checked")) {
                    var html = Claim.Preambles.CRMPreamble;
                    for (var i = 0; i < Claim.Steps.length; i++) {
                        var html = compineHtml(html, processStepHtml("CRM", Claim.Steps[i].html));
                    }
                    writeHtmlData('<html><head><meta http-equiv=Content-Type content="text/html; charset=windows-1256">' + html + '<p class="MsoNormal">&nbsp;</p></body></html>');
                }
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

    });
})();
