'use strict';

(function () {

  var itemId;
  var ssoToken;
  var token;
  // The initialize function must be run each time a new page is loaded
  Office.initialize = function (reason) {
    var item = Office.context.mailbox.item;
    itemId = Office.context.mailbox.item.itemId;
    if (itemId === null || itemId == undefined) {
      Office.context.mailbox.item.saveAsync(function(result){
        itemId = result.value;
        console.log("ItemId is " + itemId);
      });
    }
    Office.context.mailbox.item.notificationMessages.addAsync("information", {
    type: "informationalMessage",
    message : "Request validated, Encryption is in progress.",
    icon : "icon-16",
    persistent: false
    });
    $(document).ready(function () {
      //loadItemProps();
      //getCallbackToken();
      getIdentityToken();
      //getAccessToken();
      //var url = new URI('./settings/dialog.html').absoluteTo(window.location).toString();
    //var dialogOptions = { width: 20, height: 40, displayInIframe: true };
      
      Office.context.ui.displayDialogAsync('https://localhost:3000/src/settings/dialog.html', {width: 20, height: 10, displayInIframe: true} ,
      function (asyncResult) {
          if (asyncResult.status === Office.AsyncResultStatus.Failed) {
              showNotification(asyncResult.error.code = ": " + asyncResult.error.message);
          } else {
              dialog = asyncResult.value;
              dialog.addEventHandler(Office.EventType.DialogMessageReceived, function (asyncResult) {
                if (asyncResult.type !== Microsoft.Office.WebExtension.EventType.DialogMessageReceived) {
                    // TODO: Handle unknown message.
                    return;
                }
                //dialog.close();
            });
          }
      });
      setTimeout(function(){
        encryptMail();
      }, 3000);

      
      setTimeout(function(){
        Office.context.ui.closeContainer();
      }, 12000);
      //setTimeout(function(){item.close()},6000);
      
    });
  };

  function processMessage(arg) {
    console.log("event handler for dialog");
    setTimeout(function(){dialog.close();},10000);
    
    dialog = null;
    // message processing code goes here;
}
  function getAccessToken() {
    console.log("access");
  Office.context.auth.getAccessTokenAsync(cbToken);
  console.log("no access"); 
}

function cbToken(asyncResult) {
  if (asyncResult.status === "succeeded") {
  token = asyncResult.value;
  console.log("access token : " + token);
}
else {
        if (result.error.code === 13003) {
            // SSO is not supported for domain user accounts, only
            // work or school (Office 365) or Microsoft Account IDs.
            console.log("not supported");
        } }
}

function getIdentityToken() {
  Office.context.mailbox.getUserIdentityTokenAsync(cbIdentity);
}

function cbIdentity(asyncResult) {
   ssoToken = asyncResult.value;
  console.log("identity token: " + ssoToken);
}

  function encryptMail() {
    
    encryptCurrentMessage();
 
  }

  // This function handles the click event of the sendNow button.
    // It retrieves the current mail item, so that we can get its itemId property
    // ans also get the MIME content
    // It also retrieves the mailbox, so that we can make an EWS request
    // to get more properties of the item. 
    function encryptCurrentMessage() {
        console.log("hello we are here");
        var tbody = $('.prop-table');
              tbody.append(makeTableRow("status", "encrypt"));
              $.support.cors = true;
        try{
          var item = {
              itemId: itemId
          }
          console.log("itemId inside " + item);
          var tbody = $('.prop-table');
              tbody.append(makeTableRow("status", "here"));
          $('#target').html('sending..');
          $.ajax({
            url: 'https://pooja.cres.c3s2.smtpi.com',
            type: 'post',
            headers: {
              "Authorization": "Bearer " + ssoToken
            },
            dataType: 'json',
            data: JSON.stringify(item),
            contentType: 'application/json',
            success: function (data) {
              var tbody = $('.prop-table');
              tbody.append(makeTableRow("successStatus", "success"));
            },
            error: function (data,errorCode, errorMessage) {
              var tbody = $('.prop-table');
              tbody.append(makeTableRow("errorStatus", errorMessage));
            //app.showNotification('Error: ' + errorCode + ' - ' + errorMessage);
             }
            
        });
        }
        catch (error) {
            //showNotification("Unspecified error.", err.Message);
            
        }
       
    }



    // This function is the callback for the getMailItemMimeContent method
    // in the getCurrentMessage function.
    // In brief, it first checks for an error repsonse, but if all is OK
    // t:ItemId element.
    // Recieves: mail message content as a Base64 MIME string
    function sendMessageCallback(content) {
        var toAddress = "poosuman@cisco.com";
        //var comment = $("#forward-comment").val();
        //if (comment == null || comment == '') {
        //   comment = "[user provided no comment]";
        //}
        try{
            console.log("hello");
            easyEws.sendPlainTextEmailWithAttachment("SPAM ALERT!!",
                                                     "Spam Reported",
                                                     toAddress,
                                                     "Email Attachment",
                                                     content,
                                                     successCallback,
                                                     showErrorCallback);
            console.log("here again");
        }
        catch (error) {
            //showNotification("Unspecified error.", err.Message);
        }
    }

    function copyToSpamFolder(item) {
      console.log("In copy to folder");
      debugger;
      var item = Office.context.mailbox.item;
      var itemId = item.itemId;
      var folderId = 'drafts';
      var soap = '<m:MoveItem>' +
      '                 <m:ToFolderId>' +
      '                    <t:DistinguishedFolderId Id="' + folderId + '"/>' +
      '                 </m:ToFolderId>' +
      '                 <m:ItemIds>' +
      '                    <t:ItemId Id="' + itemId +'"/>' +
      '                 </m:ItemIds>' +
      '           </m:MoveItem>';
      soap = getSoapHeader(soap);
      console.log("soap request",soap);
      Office.context.mailbox.makeEwsRequestAsync(soap, function (ewsResult) {
                if (ewsResult.status == "succeeded") {
                    console.log("success");
                    var xmlDoc = $.parseXML(ewsResult.value);
                    successCallback(xmlDoc);
                    //debugCallback(ewsResult.value); // return raw result
                } else {
                        console.log("error");
                        showErrorCallback("makeEwsRequestAsync failed.");
                        //debugCallback(ewsResult.value); // return raw result
                    
                }
            });
      
    }

    function getSoapHeader(request) {
            var result =
                '<?xml version="1.0" encoding="utf-8"?>' +
                '<soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"' +
                '               xmlns:xsd="http://www.w3.org/2001/XMLSchema"' +
                '               xmlns:m="http://schemas.microsoft.com/exchange/services/2006/messages"' +
                '               xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/"' +
                '               xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types">' +
                '   <soap:Header>' +
                '       <RequestServerVersion Version="Exchange2013" xmlns="http://schemas.microsoft.com/exchange/services/2006/types" soap:mustUnderstand="0" />' +
                '   </soap:Header>' +
                '   <soap:Body>' + request + '</soap:Body>' +
                '</soap:Envelope>';
            return result;
        };

    // This function is the callback for the easyEws sendPlainTextEmailWithAttachment
  // Recieves: a message that the result was successful.
    function successCallback(result) {

        // Get the table body element
        //var tbody = $('.prop-table');
        //tbody.append(makeTableRow("status", "successful"));
        //showNotification("Success", result); 
    
    }

    // This function will display errors that occur 
    // we use this as a callback for errors in easyEws
    function showErrorCallback(error) {
        
        // Get the table body element
        //var tbody = $('.prop-table');
        //tbody.append(makeTableRow("status", "error"));
    // Add a row to the table for each message property   
        //showNotification("Error", error);// .error.message);
    }

    function makeTableRow(name, value) {
    return $("<tr><td><strong>" + name + 
      "</strong></td><td class=\"prop-val\"><code>" +
      value + "</code></td></tr>");
  }


})();