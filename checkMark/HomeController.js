var myCtrl = ['$scope', function ($scope, AngularServices) {
    $scope.Rows = [

        [
            {
                "Name": "C",
                "Comment": "check capitalization"
            }
            ,
            {
                "Name": "F",
                "Comment": "check for fragment"
            }
            ,
            {
                "Name": "P",
                "Comment": "Check punctuation"
            }
            ,
            {
                "Name": "S",
                "Comment": "Check spelling"

            }
            ,
            {
                "Name": "F",
                "Comment": "check for fragment"

            }
            ,
            {
                "Name": "T",
                "Comment": "Tense"

            }
            ,
            {
                "Name": "CS",
                "Comment": "Check for comma splice"

            }
            ,
            {
                "Name": "F",
                "Comment": "check for fragment"

            }
            ,
            {
                "Name": "RO",
                "Comment": "check for run-on sentence"

            }
            ,
            {
                "Name": "SV",
                "Comment": "Subject/Verb agreement"

            }
            ,
            {
                "Name": "¶",
                "Comment": "New paragraph needed"

            }]
        ,
        [{
            "Name": "cla",
            "Comment": "Clarify your idea/meaning"

        }
            ,
        {
            "Name": "det",
            "Comment": "detail needed"

        }
            ,
        {
            "Name": "dis",
            "Comment": "Discussion needed"

        }
            ,
        {
            "Name": "evi",
            "Comment": "Evidence needed"

        }
            ,
        {
            "Name": "rep",
            "Comment": "Repetitive"

        }
            ,
        {
            "Name": "phr",
            "Comment": "Rephrase"

        }
            ,
        {
            "Name": "spa",
            "Comment": "Spacing"

        }
            ,
        {
            "Name": "cit",
            "Comment": "Check citation"

        }]

    ];
    $scope.userName = "Check Mark";
    $scope.Apply = function () {
        localStorage.setObj("Rows", $scope.Rows);
       
    }
    
 
    $scope.Cancel = function () {
        $scope.Rows = localStorage.getObj("Rows");
    }
    $scope.AddComment = function () {
        var comment = this.col.Comment;
        insertComment(comment);
    }
    $("#toggle").click(function () {
        $("#config").toggle();
    });

    function insertComment(comment) {
        console.log(getHostInfo());
        if (PlatformIsWordOnline())
            insertCommentAtOnlineAPP(comment);
        else
            insertCommentAtDesktopApp(comment);
    }
    function insertCommentAtDesktopApp(addedComment) {
        Office.context.document.getSelectedDataAsync(Office.CoercionType.Ooxml,
            function (asyncResult) {
                var ooxml = asyncResult.value;
                var parser = new DOMParser();
                var xmlDoc = parser.parseFromString(ooxml, "text/xml");
                var arr = xmlDoc.getElementsByTagName('w:comments');
                var x = arr.length;
                if (x == 0) {
                    // no comments
                    if (ooxml.indexOf('<w:r ') > ooxml.indexOf('<w:r>'))
                        ooxml = ooxml.replace('<w:r ', '<w:commentRangeStart w:id="0"/><w:r ');
                    else
                        ooxml = ooxml.replace('<w:r>', '<w:commentRangeStart w:id="0"/><w:r>');
                    var num = ooxml.lastIndexOf('</w:r>');
                    ooxml = ooxml.slice(0, num) + ooxml.slice(num).replace('</w:r>', '</w:r><w:commentRangeEnd w:id="' + n + '"/><w:r><w:rPr><w:rStyle w:val="CommentReference"/></w:rPr><w:commentReference w:id="' + n + '"/></w:r>');
                    ooxml = ooxml.replace('<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/webSettings" Target="webSettings.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/><Relationship Id="rId5" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/><Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/fontTable" Target="fontTable.xml"/></Relationships>', '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/webSettings" Target="webSettings.xml"/><Relationship Id="rId7" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/><Relationship Id="rId5" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/fontTable" Target="fontTable.xml"/><Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" Target="comments.xml"/></Relationships>')
                    ooxml = ooxml.replace('</w:styles></pkg:xmlData></pkg:part>', '<w:style w:type="character" w:styleId="CommentReference"><w:name w:val="annotation reference"/><w:basedOn w:val="DefaultParagraphFont"/><w:uiPriority w:val="99"/><w:semiHidden/><w:unhideWhenUsed/><w:rsid w:val="00833CF3"/><w:rPr><w:sz w:val="16"/><w:szCs w:val="16"/></w:rPr></w:style><w:style w:type="paragraph" w:styleId="CommentText"><w:name w:val="annotation text"/><w:basedOn w:val="Normal"/><w:link w:val="CommentTextChar"/><w:uiPriority w:val="99"/><w:semiHidden/><w:unhideWhenUsed/><w:rsid w:val="00833CF3"/><w:pPr><w:spacing w:line="240" w:lineRule="auto"/></w:pPr><w:rPr><w:sz w:val="20"/><w:szCs w:val="20"/></w:rPr></w:style><w:style w:type="character" w:customStyle="1" w:styleId="CommentTextChar"><w:name w:val="Comment Text Char"/><w:basedOn w:val="DefaultParagraphFont"/><w:link w:val="CommentText"/><w:uiPriority w:val="99"/><w:semiHidden/><w:rsid w:val="00833CF3"/><w:rPr><w:sz w:val="20"/><w:szCs w:val="20"/></w:rPr></w:style><w:style w:type="paragraph" w:styleId="CommentSubject"><w:name w:val="annotation subject"/><w:basedOn w:val="CommentText"/><w:next w:val="CommentText"/><w:link w:val="CommentSubjectChar"/><w:uiPriority w:val="99"/><w:semiHidden/><w:unhideWhenUsed/><w:rsid w:val="00833CF3"/><w:rPr><w:b/><w:bCs/></w:rPr></w:style><w:style w:type="character" w:customStyle="1" w:styleId="CommentSubjectChar"><w:name w:val="Comment Subject Char"/><w:basedOn w:val="CommentTextChar"/><w:link w:val="CommentSubject"/><w:uiPriority w:val="99"/><w:semiHidden/><w:rsid w:val="00833CF3"/><w:rPr><w:b/><w:bCs/><w:sz w:val="20"/><w:szCs w:val="20"/></w:rPr></w:style><w:style w:type="paragraph" w:styleId="BalloonText"><w:name w:val="Balloon Text"/><w:basedOn w:val="Normal"/><w:link w:val="BalloonTextChar"/><w:uiPriority w:val="99"/><w:semiHidden/><w:unhideWhenUsed/><w:rsid w:val="00833CF3"/><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr><w:rPr><w:rFonts w:ascii="Segoe UI" w:hAnsi="Segoe UI" w:cs="Segoe UI"/><w:sz w:val="18"/><w:szCs w:val="18"/></w:rPr></w:style><w:style w:type="character" w:customStyle="1" w:styleId="BalloonTextChar"><w:name w:val="Balloon Text Char"/><w:basedOn w:val="DefaultParagraphFont"/><w:link w:val="BalloonText"/><w:uiPriority w:val="99"/><w:semiHidden/><w:rsid w:val="00833CF3"/><w:rPr><w:rFonts w:ascii="Segoe UI" w:hAnsi="Segoe UI" w:cs="Segoe UI"/><w:sz w:val="18"/><w:szCs w:val="18"/></w:rPr></w:style></w:styles></pkg:xmlData></pkg:part><pkg:part pkg:name="/word/comments.xml" pkg:contentType="application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"><pkg:xmlData><w:comments xmlns:wpc="http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas" xmlns:cx="http://schemas.microsoft.com/office/drawing/2014/chartex" xmlns:cx1="http://schemas.microsoft.com/office/drawing/2015/9/8/chartex" xmlns:cx2="http://schemas.microsoft.com/office/drawing/2015/10/21/chartex" xmlns:cx3="http://schemas.microsoft.com/office/drawing/2016/5/9/chartex" xmlns:cx4="http://schemas.microsoft.com/office/drawing/2016/5/10/chartex" xmlns:cx5="http://schemas.microsoft.com/office/drawing/2016/5/11/chartex" xmlns:cx6="http://schemas.microsoft.com/office/drawing/2016/5/12/chartex" xmlns:cx7="http://schemas.microsoft.com/office/drawing/2016/5/13/chartex" xmlns:cx8="http://schemas.microsoft.com/office/drawing/2016/5/14/chartex" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:aink="http://schemas.microsoft.com/office/drawing/2016/ink" xmlns:am3d="http://schemas.microsoft.com/office/drawing/2017/model3d" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:wp14="http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:w10="urn:schemas-microsoft-com:office:word" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml" xmlns:w16cid="http://schemas.microsoft.com/office/word/2016/wordml/cid" xmlns:w16se="http://schemas.microsoft.com/office/word/2015/wordml/symex" xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup" xmlns:wpi="http://schemas.microsoft.com/office/word/2010/wordprocessingInk" xmlns:wne="http://schemas.microsoft.com/office/word/2006/wordml" xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" mc:Ignorable="w14 w15 w16se w16cid wp14"><w:comment w:id="0" w:author = "' + $scope.userName + '" w:date = "' + new Date().toISOString() + '" w:initials = "' + $scope.userName.match(/\b(\w)/g).join('') + '"> <w:p w:rsidR="002D485B" w:rsidRDefault="002D485B"><w:pPr><w:pStyle w:val="CommentText"/></w:pPr><w:r><w:rPr><w:rStyle w:val="CommentReference"/></w:rPr><w:annotationRef/></w:r><w:r><w:t>' + addedComment + '</w:t></w:r></w:p></w:comment></w:comments></pkg:xmlData></pkg:part>');
                }
                else {
                    // old comments
                    var n = xmlDoc.getElementsByTagName('w:comment').length;
                    var rd = xmlDoc.getElementsByTagName('w:p')[0].getAttribute('w:rsidRDefault');
                    if (ooxml.indexOf('<w:r') == -1 || (ooxml.indexOf('<w:r ') != -1 && ooxml.indexOf('<w:r ') < ooxml.indexOf('<w:r>')))
                        ooxml = ooxml.replace('<w:r ', '<w:commentRangeStart w:id="' + n + '"/><w:r ');
                    else
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
    function insertCommentAtOnlineAPP(addedComment) {

        Office.context.document.getSelectedDataAsync(Office.CoercionType.Ooxml,
            function (asyncResult) {
                var ooxml = asyncResult.value;
                var parser = new DOMParser();
                var xmlDoc = parser.parseFromString(ooxml, "text/xml");
                var arr = xmlDoc.getElementsByTagName('w:comments');
                var x = arr.length;
                if (x == 0) {
                    // no comments


                    if (ooxml.indexOf('<w:r ') > ooxml.indexOf('<w:r>'))
                        ooxml = ooxml.replace('<w:r ', '<w:commentRangeStart w:id="0"/><w:r ');
                    else
                        ooxml = ooxml.replace('<w:r>', '<w:commentRangeStart w:id="0"/><w:r>');
                    var num = ooxml.lastIndexOf('</w:r>');
                    ooxml = ooxml.slice(0, num) + ooxml.slice(num).replace('</w:r>', '</w:r><w:commentRangeEnd w:id="0"/><w:r><w:rPr><w:rStyle w:val="CommentReference"/></w:rPr><w:commentReference w:id="0"/></w:r>');
                    //console.log(ooxml.indexOf('<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/webSettings" Target="webSettings.xml" /><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml" /><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml" /><Relationship Id="rId5" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml" /><Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/fontTable" Target="fontTable.xml" /></Relationships></pkg:xmlData></pkg:part>'));
                    ooxml = ooxml.replace('<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/webSettings" Target="webSettings.xml" /><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml" /><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml" /><Relationship Id="rId5" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml" /><Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/fontTable" Target="fontTable.xml" /></Relationships></pkg:xmlData></pkg:part>', '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"> <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/webSettings" Target="webSettings.xml" /> <Relationship Id="rId7" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml" /> <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml" /> <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml" /> <Relationship Id="rId6" Type="http://schemas.microsoft.com/office/2011/relationships/people" Target="people.xml" /> <Relationship Id="rId5" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/fontTable" Target="fontTable.xml" /> <Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" Target="comments.xml" /> </Relationships> </pkg:xmlData> </pkg:part> <pkg:part pkg:name="/word/comments.xml" pkg:contentType="application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"> <pkg:xmlData> <w:comments xmlns:wpc="http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas" xmlns:cx="http://schemas.microsoft.com/office/drawing/2014/chartex" xmlns:cx1="http://schemas.microsoft.com/office/drawing/2015/9/8/chartex" xmlns:cx2="http://schemas.microsoft.com/office/drawing/2015/10/21/chartex" xmlns:cx3="http://schemas.microsoft.com/office/drawing/2016/5/9/chartex" xmlns:cx4="http://schemas.microsoft.com/office/drawing/2016/5/10/chartex" xmlns:cx5="http://schemas.microsoft.com/office/drawing/2016/5/11/chartex" xmlns:cx6="http://schemas.microsoft.com/office/drawing/2016/5/12/chartex" xmlns:cx7="http://schemas.microsoft.com/office/drawing/2016/5/13/chartex" xmlns:cx8="http://schemas.microsoft.com/office/drawing/2016/5/14/chartex" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:aink="http://schemas.microsoft.com/office/drawing/2016/ink" xmlns:am3d="http://schemas.microsoft.com/office/drawing/2017/model3d" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:wp14="http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:w10="urn:schemas-microsoft-com:office:word" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml" xmlns:w16cid="http://schemas.microsoft.com/office/word/2016/wordml/cid" xmlns:w16se="http://schemas.microsoft.com/office/word/2015/wordml/symex" xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup" xmlns:wpi="http://schemas.microsoft.com/office/word/2010/wordprocessingInk" xmlns:wne="http://schemas.microsoft.com/office/word/2006/wordml" xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" mc:Ignorable="w14 w15 w16se w16cid wp14"> <w:comment w:id="0" w:author = "' + $scope.userName + '" w:date = "' + new Date().toISOString() + '" w:initials = "' + $scope.userName.match(/\b(\w)/g).join('') + '"> <w:p w:rsidR="002D485B" w:rsidRDefault="002D485B"><w:pPr><w:pStyle w:val="CommentText"/></w:pPr><w:r><w:rPr><w:rStyle w:val="CommentReference"/></w:rPr><w:annotationRef/></w:r><w:r><w:t xml:space="preserve">' + addedComment + '</w:t></w:r></w:p></w:comment></w:comments></pkg:xmlData></pkg:part>');
                    ooxml = ooxml.replace('</w:styles></pkg:xmlData></pkg:part>', '<w:style w:type="character" w:styleId="CommentReference"><w:name w:val="annotation reference"/><w:basedOn w:val="DefaultParagraphFont"/><w:uiPriority w:val="99"/><w:semiHidden/><w:unhideWhenUsed/><w:rsid w:val="00833CF3"/><w:rPr><w:sz w:val="16"/><w:szCs w:val="16"/></w:rPr></w:style><w:style w:type="paragraph" w:styleId="CommentText"><w:name w:val="annotation text"/><w:basedOn w:val="Normal"/><w:link w:val="CommentTextChar"/><w:uiPriority w:val="99"/><w:semiHidden/><w:unhideWhenUsed/><w:rsid w:val="00833CF3"/><w:pPr><w:spacing w:line="240" w:lineRule="auto"/></w:pPr><w:rPr><w:sz w:val="20"/><w:szCs w:val="20"/></w:rPr></w:style><w:style w:type="character" w:customStyle="1" w:styleId="CommentTextChar"><w:name w:val="Comment Text Char"/><w:basedOn w:val="DefaultParagraphFont"/><w:link w:val="CommentText"/><w:uiPriority w:val="99"/><w:semiHidden/><w:rsid w:val="00833CF3"/><w:rPr><w:sz w:val="20"/><w:szCs w:val="20"/></w:rPr></w:style><w:style w:type="paragraph" w:styleId="CommentSubject"><w:name w:val="annotation subject"/><w:basedOn w:val="CommentText"/><w:next w:val="CommentText"/><w:link w:val="CommentSubjectChar"/><w:uiPriority w:val="99"/><w:semiHidden/><w:unhideWhenUsed/><w:rsid w:val="00833CF3"/><w:rPr><w:b/><w:bCs/></w:rPr></w:style><w:style w:type="character" w:customStyle="1" w:styleId="CommentSubjectChar"><w:name w:val="Comment Subject Char"/><w:basedOn w:val="CommentTextChar"/><w:link w:val="CommentSubject"/><w:uiPriority w:val="99"/><w:semiHidden/><w:rsid w:val="00833CF3"/><w:rPr><w:b/><w:bCs/><w:sz w:val="20"/><w:szCs w:val="20"/></w:rPr></w:style><w:style w:type="paragraph" w:styleId="BalloonText"><w:name w:val="Balloon Text"/><w:basedOn w:val="Normal"/><w:link w:val="BalloonTextChar"/><w:uiPriority w:val="99"/><w:semiHidden/><w:unhideWhenUsed/><w:rsid w:val="00833CF3"/><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr><w:rPr><w:rFonts w:ascii="Segoe UI" w:hAnsi="Segoe UI" w:cs="Segoe UI"/><w:sz w:val="18"/><w:szCs w:val="18"/></w:rPr></w:style><w:style w:type="character" w:customStyle="1" w:styleId="BalloonTextChar"><w:name w:val="Balloon Text Char"/><w:basedOn w:val="DefaultParagraphFont"/><w:link w:val="BalloonText"/><w:uiPriority w:val="99"/><w:semiHidden/><w:rsid w:val="00833CF3"/><w:rPr><w:rFonts w:ascii="Segoe UI" w:hAnsi="Segoe UI" w:cs="Segoe UI"/><w:sz w:val="18"/><w:szCs w:val="18"/></w:rPr></w:style></w:styles></pkg:xmlData></pkg:part>');


                }
                else {
                    // old comments
                    var n = xmlDoc.getElementsByTagName('w:comment').length;
                    var rd = xmlDoc.getElementsByTagName('w:p')[0].getAttribute('w:rsidRDefault');
                    if (ooxml.indexOf('<w:r') == -1 || (ooxml.indexOf('<w:r ') != -1 && ooxml.indexOf('<w:r ') < ooxml.indexOf('<w:r>')))
                        ooxml = ooxml.replace('<w:r ', '<w:commentRangeStart w:id="' + n + '"/><w:r ');
                    else
                        ooxml = ooxml.replace('<w:r>', '<w:commentRangeStart w:id="' + n + '"/><w:r>');
                    var num1 = ooxml.lastIndexOf('</w:body>');
                    var num2 = ooxml.slice(0, num1).lastIndexOf('</w:p>');
                    ooxml = ooxml.slice(0, num1).slice(0, num2) + ooxml.slice(0, num1).slice(num2).replace('</w:p>', '<w:commentRangeEnd w:id="' + n + '"/><w:r><w:rPr><w:rStyle w:val="CommentReference"/></w:rPr><w:commentReference w:id="' + n + '"/></w:r></w:p>') + ooxml.slice(num1);
                    ooxml = ooxml.replace('</w:comment>', '</w:comment><w:comment w:id="' + n + '" w:author = "' + $scope.userName + '" w:date = "' + new Date().toISOString() + '" w:initials = "' + $scope.userName.match(/\b(\w)/g).join('') + '"> <w:p w:rsidR="' + rd + '" w:rsidRDefault="' + rd + '"><w:pPr><w:pStyle w:val="CommentText"/></w:pPr><w:r><w:rPr><w:rStyle w:val="CommentReference"/></w:rPr><w:annotationRef/></w:r><w:r><w:t xml:space="preserve">' + addedComment + '</w:t></w:r></w:p></w:comment>');

                }
                Office.context.document.setSelectedDataAsync(ooxml, { coercionType: Office.CoercionType.Ooxml }, function (asyncResult) {

                    if (asyncResult.status === "failed") {
                        console.log("Action failed with error: " + asyncResult.error.message);

                    }


                });
            });

    }
    $scope.AddRow  =   function () {
        $scope.Rows.push([]);
    }
    $scope.heighlightBtn = function (obj){
        var row = Number(obj.target.parentNode.parentNode.getAttribute("data-ng-row")) + 1;
        var col = Number(obj.target.parentNode.getAttribute("data-ng-col")) + 1;
        $(".buttons").css("border", "none");
        $("#btnsCon > :nth-child(" + row + ") :nth-child(" + col + ")").css("border", "2px solid red");
    }

    $scope.AddCtrls = function (obj) {
        var RowInd = obj.target.parentNode.parentNode.parentNode.getAttribute("data-ng-row");
        $scope.Rows[RowInd].push({ "Name": "", "Comment": "" });
    }
    $scope.DeleteRow = function (obj) {
        var RowInd = obj.target.parentNode.parentNode.parentNode.getAttribute("data-ng-row");
        $scope.Rows.splice(RowInd, 1);
    }
    $scope.DeleteCtrl = function (obj) {
        var RowInd = obj.target.parentNode.parentNode.parentNode.getAttribute("data-ng-row");
        var ColInd = obj.target.parentNode.parentNode.getAttribute("data-ng-col");
        $scope.Rows[RowInd].splice(ColInd, 1);
    }
    function getHostInfo() {
        var hostInfoValue = sessionStorage.getItem('hostInfoValue');

        if (hostInfoValue === null) return null;
        // Parse the value string (reference: office.debug.js)
        var items = hostInfoValue.split('$');
        if (!items[2]) {
            items = hostInfoValue.split('|');
        }

        var hostInfo = {
            type: items[0],
            platform: items[1],
            version: items[2],
            //culture: items[3] // Some platforms (i.e. Win32) returns a culture property
        };
        return hostInfo;
    }
    
    function PlatformIsWordOnline() {
        return getHostInfo().platform.toLowerCase() == "web";

    }

    $(document).ready(function () {
        $("#start").click(function () {
            $("#getStarted").css("display", "none");
        });
        var Rows = localStorage.getObj("Rows");
        if (Rows == null)
            localStorage.setObj("Rows",$scope.Rows);
        else
            $scope.Rows = Rows;
        $scope.$apply();
    });

}];

app.controller("myCtrl", myCtrl);

