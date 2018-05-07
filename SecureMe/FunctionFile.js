

(function () {

    Office.initialize = function (reason) {
        //If you need to initialize something you can do so here.


    };
})();

localStorage.removeItem("Headers");
var itemUrl = "";
var rawToken = "";
//var DeploymentHost = "https://realdocs.pronetcre.com/";
var DeploymentHost = "https://localhost:44392/";
var HeadersDialogUrl = DeploymentHost + "HeadersDialog.html";
    // The initialize function must be run each time a new page is loaded.
   

function loadRestDetails() {
        var restId = '';
        if (Office.context.mailbox.diagnostics.hostName !== 'OutlookIOS') {
            // Loaded in non-mobile context, so ID needs to be converted
            restId = Office.context.mailbox.convertToRestId(
                Office.context.mailbox.item.itemId,
                Office.MailboxEnums.RestVersion.Beta
            );
        } else {
            restId = Office.context.mailbox.item.itemId;
        }

        // Build the URL to the item
        //var itemUrl = Office.context.mailbox.restUrl + 
        itemUrl = Office.context.mailbox.restUrl  +
            '/v2.0/me/messages/' + restId;

        Office.context.mailbox.getCallbackTokenAsync({ isRest: true }, function (result) {
            if (result.status === "succeeded") {
                rawToken = result.value;
                getItemHeadersViaRest();
            } else {
                rawToken = 'error';
            }
        });
    }
function getItemHeadersViaRest() {
        $.ajax({
            url: itemUrl + '?$select = InternetMessageHeaders',
            dataType: 'json',
            headers: { 'Authorization': 'Bearer ' + rawToken }
        }).done(function (item) {
            var headersNames = item.InternetMessageHeaders.map(function (el) {
                return el.Name;
            });
            var index = $.inArray("X-Sn-Email-Phishing", headersNames);
            if (index !== -1)
                localStorage.setItem("Headers", "X-Sn-Email-Phishing = " + item.InternetMessageHeaders[index].Value);
            else
                localStorage.setItem("Headers", "Header not found.");
            ShowHeadersDialog();
            return;
        }).fail(function (error) {
            console.log(error);
        });
    }

    
function ShowHeadersDialog() {
    Office.context.ui.displayDialogAsync(HeadersDialogUrl, { height: 60, width: 25, displayInIframe: false });
    }

