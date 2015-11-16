/// <reference path="../App.js" />
var forwardingData;
var changeKey; //retrieved in getItemDataCallback, used for forwarding and deleting the email
(function () {
    "use strict";

    // The Office initialize function must be run each time a new page is loaded
    Office.initialize = function (reason) {
        $(document).ready(function () {
            app.initialize();

            $('#buttonForward').click(beginForwardAnEmail);
        });
    };

})();

function ForwardingData(to, subject, body, sourceid) {
    this.to = to;
    this.subject = subject;    
    this.body = body;    
    this.sourceid = sourceid;
};

function beginForwardAnEmail() {
    try {

        //Get a handle to the 'item' object for the active email 
        var item = Office.cast.item.toItemRead(Office.context.mailbox.item);
          
        //Get values from the page controls
        forwardingData = new ForwardingData($('#txtTo').prop('value'), $('#txtSubject').prop('value'), $('#txtBody').prop('value'), item.itemId);

        app.showNotification('Processing...', 'Forwarding email to ' + forwardingData.to + '...');

        //Call EWS to get the source item; then we need to make another EWS request to forward it
        var mailbox = Office.context.mailbox;
        mailbox.makeEwsRequestAsync(getItemDataRequest(item.itemId), getItemDataCallback);
    } catch (e) {
        app.showNotification('Error! ' + status, '[in beginForwardEmail()]:' + e);
    }
};

function getItemDataRequest(item_id) {
    var request;

    request = '<?xml version="1.0" encoding="utf-8"?>' +
        '<soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"' +
        '               xmlns:m="http://schemas.microsoft.com/exchange/services/2006/messages"' +
        '               xmlns:xsd="http://www.w3.org/2001/XMLSchema"' +
        '               xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/"' +
        '               xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types">' +
        '  <soap:Header>' +
        '    <RequestServerVersion Version="Exchange2013" xmlns="http://schemas.microsoft.com/exchange/services/2006/types" soap:mustUnderstand="0" />' +
        '  </soap:Header>' +
        '  <soap:Body>' +
        '    <GetItem' +
        '                xmlns="http://schemas.microsoft.com/exchange/services/2006/messages"' +
        '                xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types">' +
        '      <ItemShape>' +
        '        <t:BaseShape>IdOnly</t:BaseShape>' +
        '      </ItemShape>' +
        '      <ItemIds>' +
        '        <t:ItemId Id="' + item_id + '"/>' +
        '      </ItemIds>' +
        '    </GetItem>' +
        '  </soap:Body>' +
        '</soap:Envelope>';

    return request;
}

function getItemDataCallback(asyncResult) {
    if (asyncResult == null) {
        app.showNotification('Error!', '[in getItemDataCallback]: null result');
        return;
    }

    if (asyncResult.error != null) {
        app.showNotification('Error!', '[in getItemDataCallback]: ' + asyncResult.error.message);
    }
    else {

        var errorMsg;
        var prop = null;
        try {
            var response = $.parseXML(asyncResult.value);
            var responseDOM = $(response);

            if (responseDOM) {
                prop = responseDOM.filterNode("t:ItemId")[0];
            }

        } catch (e) {
            errorMsg = e;
        }
        if (!prop) {
            if (errorMsg) {
                app.showNotification('Error!', '[in getItemDataCallback]: Failed to retrieve item data (' + errorMsg + ')');
            }
            else { app.showNotification('Error!', '[in getItemDataCallback]: Failed to retrieve item data'); }

            return;
        }

        changeKey = prop.getAttribute("ChangeKey");
        // Now that we have a ChangeKey value, we can use EWS to forward the mail item.

        var addresses;
        var addressesSoap = '';
        
        addresses = forwardingData.to.split(",");
        
        // The following loop will build an XML fragment that we will insert into the SOAP message
        for (var address = 0; address < addresses.length; address++) {
            //Need to trim the addresses, as a leading space will cause an error!
            addressesSoap += "<t:Mailbox><t:EmailAddress>" + addresses[address].trim() + "</t:EmailAddress></t:Mailbox>";
        }

        //Now, forward the item

        app.showNotification('Processing...', 'Initializing forward to ' + forwardingData.to + '...');

        var mailbox = Office.context.mailbox;       
        mailbox.makeEwsRequestAsync(forwardItemRequest(forwardingData.subject, forwardingData.sourceid, addressesSoap, changeKey, forwardingData.body), forwardItemCallback);
    }
}

function forwardItemCallback(asyncResult) {

    if (asyncResult == null) {
        app.showNotification('Error!', '[in forwardItemCallback]: null result');
        return;
    }

    if (asyncResult.error != null) {
        app.showNotification('Error!', '[in forwardItemCallback]: ' + asyncResult.error.message);
    }
    else {

        var errorMsg;
        var prop = null;
        try {
            var response = $.parseXML(asyncResult.value);
            var responseDOM = $(response);

            if (responseDOM) {                
                prop = responseDOM.filterNode("m:ResponseCode")[0];
            }

        } catch (e) {
            errorMsg = e;
        }
        if (!prop) {
            if (errorMsg) {
                app.showNotification('Error!', '[in forwardItemCallback]: ' + errorMsg);
            }
            else { app.showNotification('Error!', '[in forwardItemCallback]: Failed to parse response'); }

            return;
        } else {
            //Verify forward result
            if (prop.textContent == "NoError") {
                app.showNotification("Success!", "The email has been forwarded.");                
            }
            else {
                app.showNotification('Error!', '[in forwardItemCallback]:' + prop.textContent);
            }
        }
    }
}

function forwardItemRequest(subject, item_id, recipients, changeKey, body) {
    var request;

    // The following string is a valid SOAP envelope and request for forwarding
    // a mail item. Note that we use the item_id value to specify the item we are interested in,
    // along with its ChangeKey value that we have just determined
    // We also provide the XML fragment to specify the recipient addresses
    
    request = '<?xml version="1.0" encoding="utf-8"?>' +
        '<soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"' +
        '               xmlns:m="http://schemas.microsoft.com/exchange/services/2006/messages"' +
        '               xmlns:xsd="http://www.w3.org/2001/XMLSchema"' +
        '               xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/"' +
        '               xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types">' +
        '  <soap:Header>' +
        '    <RequestServerVersion Version="Exchange2013" xmlns="http://schemas.microsoft.com/exchange/services/2006/types" soap:mustUnderstand="0" />' +
        '  </soap:Header>' +
        '  <soap:Body>' +
        '    <m:CreateItem MessageDisposition="SendAndSaveCopy">' +
        '      <m:Items>' +
        '        <t:ForwardItem>' +
        '         <t:Subject>' + subject + '</t:Subject>' +
        '          <t:ToRecipients>' + recipients + '</t:ToRecipients>' +
        '          <t:ReferenceItemId Id="' + item_id + '" ChangeKey="' + changeKey + '" />' +
        '          <t:NewBodyContent BodyType="Text">' + body + '</t:NewBodyContent>' +
        '        </t:ForwardItem>' +
        '      </m:Items>' +
        '    </m:CreateItem>' +
        '  </soap:Body>' +
        '</soap:Envelope>';

    return request;
}

// This function plug in filters nodes for the one that matches the given name.
// This sidesteps the issues in jquerys selector logic.
(function ($) {
    $.fn.filterNode = function (node) {
        return this.find("*").filter(function () {
            return this.nodeName === node;
        });
    };
})(jQuery);