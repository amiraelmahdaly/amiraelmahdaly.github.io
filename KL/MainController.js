var myCtrl = ['$scope', 'AngularServices', function ($scope, AngularServices) {
    $scope.userName = decodeURIComponent(getQueryStringValue("userName"));
    var token = decodeURIComponent( getQueryStringValue("token")); 
    $scope.config = {};
    $scope.file = {};
    var url = Office.context.document.url;
    var number = url.substring(
        url.lastIndexOf("_") + 1,
        url.lastIndexOf(".")
    );
    var rangeTxtArr = [];
    var bodyTxtArr = [];
    var startParaID = "";
    var endParaID = "";
    var modifiers = [];
    var endIndex = "";
    var startIndex = "";
    var modID = 0;
    function getBodyPara() {
        Word.run(function (context) {
            // Create a proxy object for the document body.
            var body = context.document.body.paragraphs;
            body.load();
            return context.sync().then(function () {
                for (var i = 0; i < body.items.length; i++) {
                    bodyTxtArr.push(body.items[i].text)
                }
            })
            return context.sync();
        })
            .catch(errorHandler);
    }
    function getRangePara(txt, id, type) {
        rangeTxtArr = [];
        Word.run(function (context) {
            // Create a proxy object for the document body.
            var range = context.document.getSelection().paragraphs;
            var selection = context.document.getSelection();
            range.load();
            selection.load();
            return context.sync().then(function () {
                var rangeTxt = selection.text.replace(//g, '');
                for (var i = 0; i < range.items.length; i++) {
                    rangeTxtArr.push(range.items[i].text.replace(//g, ''))
                }
                var paraTxt = rangeTxtArr.join("\r");
                startIndex = paraTxt.indexOf(rangeTxt);
                endIndex = startIndex + rangeTxt.length - 1;
                getIndecies(bodyTxtArr, rangeTxtArr);
                if (type == "modifier")
                    addModifier(txt, id);
                else
                    buildObject();

            })

            return context.sync();
        })
            .catch(errorHandler);
    }
    function buildObject() {
        Office.context.document.getSelectedDataAsync(Office.CoercionType.Ooxml,
            function (asyncResult) {
                var ooxml = asyncResult.value;
                var parser = new DOMParser();
                var xmlDoc = parser.parseFromString(ooxml, "text/xml");
                var arr = xmlDoc.getElementsByTagName('w:comment');
                var newMods = [];
                for (var i = 0; i < arr.length; i++) {
                    var obj = modifiers[arr[i].textContent.substring(arr[i].textContent.indexOf(":") + 1, arr[i].textContent.indexOf(","))];
                    if (obj != null || obj != undefined) { 
                    delete obj.modifier_id;
                    newMods.push(obj);
                }
                }
                var body = [
                    {
                        "core_name": $("#Cores").find(":selected").text(),
                        "core_segment":
                            {
                                "start_para_id": startParaID,
                                "end_para_id": endParaID,
                                "start_index": startIndex,
                                "end_index": endIndex
                            },
                        "modifiers": newMods
                    }
                ];

               
                AngularServices.POST("anno/" + number, body, token).then(function (data) {

                    insertComment(data[0].comment);
                });
            });
    }
    function getIndecies(arr1, arr2) {
        startParaID = "";
        endParaID = "";
        for (var i = 0; i < arr1.length; i++) {
            if (arr1[i] == arr2[0]) {
                startParaID = i;
                endParaID = i + arr2.length - 1;
                return;


            }
        }
    }
    function addModifier(txt, id) {
        modifiers.push({
            "modifier_name": txt,
            "modifier_id": id,
            "modifier_segment": {
                "start_para_id": startParaID,
                "end_para_id": endParaID,
                "start_index": startIndex,
                "end_index": endIndex
            }
        })
        var x = 0;
    }
    function insertComment(addedComment) {
        Office.context.document.getSelectedDataAsync(Office.CoercionType.Ooxml,
            function (asyncResult) {
                var ooxml = asyncResult.value;
                var parser = new DOMParser();
                var xmlDoc = parser.parseFromString(ooxml, "text/xml");
                var arr = xmlDoc.getElementsByTagName('w:comments');
                var x = arr.length;
                if (x == 0) {
                    // no comments
                    ooxml = ooxml.replace('<w:r>', '<w:commentRangeStart w:id="0"/><w:r>');
                    var num = ooxml.lastIndexOf('</w:r>');
                    ooxml = ooxml.slice(0, num) + ooxml.slice(num).replace('</w:r>', '</w:r><w:commentRangeEnd w:id="' + n + '"/><w:r><w:rPr><w:rStyle w:val="CommentReference"/></w:rPr><w:commentReference w:id="' + n + '"/></w:r>');
                    ooxml = ooxml.replace('<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/webSettings" Target="webSettings.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/><Relationship Id="rId5" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/><Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/fontTable" Target="fontTable.xml"/></Relationships>', '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/webSettings" Target="webSettings.xml"/><Relationship Id="rId7" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/><Relationship Id="rId5" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/fontTable" Target="fontTable.xml"/><Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" Target="comments.xml"/></Relationships>')
                    ooxml = ooxml.replace('</w:styles></pkg:xmlData></pkg:part></pkg:package>', '<w:style w:type="character" w:styleId="CommentReference"><w:name w:val="annotation reference"/><w:basedOn w:val="DefaultParagraphFont"/><w:uiPriority w:val="99"/><w:semiHidden/><w:unhideWhenUsed/><w:rsid w:val="00833CF3"/><w:rPr><w:sz w:val="16"/><w:szCs w:val="16"/></w:rPr></w:style><w:style w:type="paragraph" w:styleId="CommentText"><w:name w:val="annotation text"/><w:basedOn w:val="Normal"/><w:link w:val="CommentTextChar"/><w:uiPriority w:val="99"/><w:semiHidden/><w:unhideWhenUsed/><w:rsid w:val="00833CF3"/><w:pPr><w:spacing w:line="240" w:lineRule="auto"/></w:pPr><w:rPr><w:sz w:val="20"/><w:szCs w:val="20"/></w:rPr></w:style><w:style w:type="character" w:customStyle="1" w:styleId="CommentTextChar"><w:name w:val="Comment Text Char"/><w:basedOn w:val="DefaultParagraphFont"/><w:link w:val="CommentText"/><w:uiPriority w:val="99"/><w:semiHidden/><w:rsid w:val="00833CF3"/><w:rPr><w:sz w:val="20"/><w:szCs w:val="20"/></w:rPr></w:style><w:style w:type="paragraph" w:styleId="CommentSubject"><w:name w:val="annotation subject"/><w:basedOn w:val="CommentText"/><w:next w:val="CommentText"/><w:link w:val="CommentSubjectChar"/><w:uiPriority w:val="99"/><w:semiHidden/><w:unhideWhenUsed/><w:rsid w:val="00833CF3"/><w:rPr><w:b/><w:bCs/></w:rPr></w:style><w:style w:type="character" w:customStyle="1" w:styleId="CommentSubjectChar"><w:name w:val="Comment Subject Char"/><w:basedOn w:val="CommentTextChar"/><w:link w:val="CommentSubject"/><w:uiPriority w:val="99"/><w:semiHidden/><w:rsid w:val="00833CF3"/><w:rPr><w:b/><w:bCs/><w:sz w:val="20"/><w:szCs w:val="20"/></w:rPr></w:style><w:style w:type="paragraph" w:styleId="BalloonText"><w:name w:val="Balloon Text"/><w:basedOn w:val="Normal"/><w:link w:val="BalloonTextChar"/><w:uiPriority w:val="99"/><w:semiHidden/><w:unhideWhenUsed/><w:rsid w:val="00833CF3"/><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr><w:rPr><w:rFonts w:ascii="Segoe UI" w:hAnsi="Segoe UI" w:cs="Segoe UI"/><w:sz w:val="18"/><w:szCs w:val="18"/></w:rPr></w:style><w:style w:type="character" w:customStyle="1" w:styleId="BalloonTextChar"><w:name w:val="Balloon Text Char"/><w:basedOn w:val="DefaultParagraphFont"/><w:link w:val="BalloonText"/><w:uiPriority w:val="99"/><w:semiHidden/><w:rsid w:val="00833CF3"/><w:rPr><w:rFonts w:ascii="Segoe UI" w:hAnsi="Segoe UI" w:cs="Segoe UI"/><w:sz w:val="18"/><w:szCs w:val="18"/></w:rPr></w:style></w:styles></pkg:xmlData></pkg:part><pkg:part pkg:name="/word/comments.xml" pkg:contentType="application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"><pkg:xmlData><w:comments xmlns:wpc="http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas" xmlns:cx="http://schemas.microsoft.com/office/drawing/2014/chartex" xmlns:cx1="http://schemas.microsoft.com/office/drawing/2015/9/8/chartex" xmlns:cx2="http://schemas.microsoft.com/office/drawing/2015/10/21/chartex" xmlns:cx3="http://schemas.microsoft.com/office/drawing/2016/5/9/chartex" xmlns:cx4="http://schemas.microsoft.com/office/drawing/2016/5/10/chartex" xmlns:cx5="http://schemas.microsoft.com/office/drawing/2016/5/11/chartex" xmlns:cx6="http://schemas.microsoft.com/office/drawing/2016/5/12/chartex" xmlns:cx7="http://schemas.microsoft.com/office/drawing/2016/5/13/chartex" xmlns:cx8="http://schemas.microsoft.com/office/drawing/2016/5/14/chartex" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:aink="http://schemas.microsoft.com/office/drawing/2016/ink" xmlns:am3d="http://schemas.microsoft.com/office/drawing/2017/model3d" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:wp14="http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:w10="urn:schemas-microsoft-com:office:word" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml" xmlns:w16cid="http://schemas.microsoft.com/office/word/2016/wordml/cid" xmlns:w16se="http://schemas.microsoft.com/office/word/2015/wordml/symex" xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup" xmlns:wpi="http://schemas.microsoft.com/office/word/2010/wordprocessingInk" xmlns:wne="http://schemas.microsoft.com/office/word/2006/wordml" xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" mc:Ignorable="w14 w15 w16se w16cid wp14"><w:comment w:id="0" w:author = "' + $scope.userName + '" w:date = "' + new Date().toISOString() + '" w:initials = "' + $scope.userName.match(/\b(\w)/g).join('') + '"> <w:p w:rsidR="002D485B" w:rsidRDefault="002D485B"><w:pPr><w:pStyle w:val="CommentText"/></w:pPr><w:r><w:rPr><w:rStyle w:val="CommentReference"/></w:rPr><w:annotationRef/></w:r><w:r><w:t>' + addedComment + '</w:t></w:r></w:p></w:comment></w:comments></pkg:xmlData></pkg:part></pkg:package>');


                }
                else {
                    // old comments
                    var n = xmlDoc.getElementsByTagName('w:comment').length;
                    var rd = xmlDoc.getElementsByTagName('w:p')[0].getAttribute('w:rsidRDefault');
                    ooxml = ooxml.replace('<w:r>', '<w:commentRangeStart w:id="' + n + '"/><w:r>');
                    var num1 = ooxml.lastIndexOf('</w:body>');
                    var num2 = ooxml.slice(0, num1).lastIndexOf('</w:p>');
                    ooxml = ooxml.slice(0, num1).slice(0, num2) + ooxml.slice(0, num1).slice(num2).replace('</w:p>', '<w:commentRangeEnd w:id="' + n + '"/><w:r><w:rPr><w:rStyle w:val="CommentReference"/></w:rPr><w:commentReference w:id="' + n + '"/></w:r></w:p>') + ooxml.slice(num1);
                    ooxml = ooxml.replace('</w:comment>', '</w:comment><w:comment w:id="' + n + '" w:author = "' + $scope.userName + '" w:date = "' + new Date().toISOString() + '" w:initials = "' + $scope.userName.match(/\b(\w)/g).join('') + '"> <w:p w:rsidR="' + rd + '" w:rsidRDefault="' + rd + '"><w:pPr><w:pStyle w:val="CommentText"/></w:pPr><w:r><w:rPr><w:rStyle w:val="CommentReference"/></w:rPr><w:annotationRef/></w:r><w:r><w:t>' + addedComment + '</w:t></w:r></w:p></w:comment>');

                }
                Office.context.document.setSelectedDataAsync(ooxml, { coercionType: Office.CoercionType.Ooxml }, function (asyncResult) {
                    if (asyncResult.status === "failed") {
                        console.log("Action failed with error: " + asyncResult.error.message);
                    }
                });
            });
    }
    $("#generateCommentBtn").click(function () {
        getRangePara("","","");
    });
    $(document).ready(function () {
        
        getBodyPara();
        AngularServices.GET("risk_policy", number).then(function (data) {
            $scope.file = data;
            console.log(data);
        });
        AngularServices.GET("config", number).then(function (data) {
            $scope.config = data;
        });
    });
    $("#saveBtn").click(function () {
        var body = {
            "client_name": $scope.config.client_name,
            "client_full_name": $scope.config.client_full_name,
            "client_address": $scope.config.client_address,
            "contract_type": $("#contractSel").val(),
            "body_highlights": $scope.config.body_highlights,
            "comment_highlights": $scope.config.comment_highlights,
            "comments": $scope.config.comments,
            "async_anno": $scope.config.async_anno
        }
        AngularServices.PUT("config",body,token,number).then(function (data) {
            console.log(data);
        });
    });
    $("#entity").change(function () {
        $(".tagSel").css("display", "none");
        $(".modifierSel").css("display", "none");
        $('.tagSel').prop('selectedIndex', 0);
        $("#" + $('#entity').find(":selected").text()).css("display", "block");
        if ($('#entity').find(":selected").text() == "Cores")
            $("#" + $('#Cores').find(":selected").text().split(' ').join('')).css("display", "block");
    });
    $("#Cores").change(function () {
        $(".modifierSel").css("display", "none");
        $("#" + $('#Cores').find(":selected").text().split(' ').join('')).css("display", "block");
    });
    $("#clause").change(function () {
        $(".clauses").css("display", "none");
        $("#" + $('#clause').find(":selected").text().split(' ').join('')).css("display", "block");
    });
    $scope.$on('ngRepeatFinished', function (ngRepeatFinishedEvent) {
        $("#" + $('#entity').find(":selected").text()).css("display", "block");
        $("#" + $('#Cores').find(":selected").text().split(' ').join('')).css("display", "block");
        $("#" + $scope.config.contract_type).attr("selected", "selected");
        $('#input_1_10 li').click(function () {
            var optIndex = $(this).index();
            var insertIndex = $("#clause").prop('selectedIndex') - 1;
            $("#clauseTxt").text($scope.config.insertions[insertIndex].options[optIndex]['option string']);
        });
        $("#addModifierBtn").click(function () {
            var modText = $(this).parent().prev().children().find(":selected").text();
            var addedComment = 'Modifier ID:' + modID + ', Entity Name:' + modText;
            insertComment(addedComment);
            getRangePara(modText, modID, "modifier");
            modID = modID + 1;
        });
    });

    
}];

app.controller("myCtrl", myCtrl);
