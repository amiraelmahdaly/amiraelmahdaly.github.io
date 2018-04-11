var URI = "https://sandbox.idsdatanet.com/d2_omv_global/webservice/depotwebservice.html";
var DeploymentHost = "https://amiraelmahdaly.github.io/IDS/";
//var DeploymentHost = "https://localhost:44300/";

function GetHeader(UserName, Password) {
    return {
        "Authorization": "Basic " + window.btoa(UserName + ":" + Password),
        'Content-Type': 'application/x-www-form-urlencoded; charset=UTF-8'
    }
}
function GetURI(data_object, query_template) {
    var c = URI + "?type=witsml&version=1.3.1.1&data_object=" + data_object + "&query_template=" + encodeURIComponent(query_template);
    return c;
}
function xmlToJson(xml) {

    // Create the return object
    var obj = {};

    if (xml.nodeType == 1) { // element
        // do attributes
        if (xml.attributes.length > 0) {
            obj["@attributes"] = {};
            for (var j = 0; j < xml.attributes.length; j++) {
                var attribute = xml.attributes.item(j);
                obj["@attributes"][attribute.nodeName] = attribute.nodeValue;
            }
        }
    } else if (xml.nodeType == 3) { // text
        obj = xml.nodeValue;
    }

    // do children
    if (xml.hasChildNodes()) {
        for (var i = 0; i < xml.childNodes.length; i++) {
            var item = xml.childNodes.item(i);
            var nodeName = item.nodeName;
            if (typeof (obj[nodeName]) == "undefined") {
                obj[nodeName] = xmlToJson(item);
            } else {
                if (typeof (obj[nodeName].push) == "undefined") {
                    var old = obj[nodeName];
                    obj[nodeName] = [];
                    obj[nodeName].push(old);
                }
                obj[nodeName].push(xmlToJson(item));
            }
        }
    }
    return obj;
}
function GetJson(xml) {
    var edittedXml = xml.replace("\n", " ").replace('\"', '"');
    var xmlDOM = new DOMParser().parseFromString(edittedXml, 'text/xml');
    return xmlToJson(xmlDOM);
    // omit return 
    // check whether the wellbores property is an array ? 
    // if array then return xmlToJson(xmlDOM);
    // if not then just push the 
}
function getQueryStringValue(key) {
    return decodeURIComponent(window.location.search.replace(new RegExp("^(?:.*[&\\?]" + encodeURIComponent(key).replace(/[\.\+\*]/g, "\\$&") + "(?:\\=([^&]*))?)?.*$", "i"), "$1"));
}
function isNumber(n) {
    return !isNaN(parseFloat(n)) && isFinite(n);
}