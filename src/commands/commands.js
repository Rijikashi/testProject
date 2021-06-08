/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */
/* global global, Office, self, window */
Office.onReady(() => {
  // If needed, Office.js is ready to be called
});

/**
 * Shows a notification when the add-in command is executed.
 * @param event {Office.AddinCommands.Event}
 */
function action(event) {
  const message = {
    type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
    message: "Performed action.",
    icon: "Icon.80x80",
    persistent: true
  };

  // Show a notification message
  var item = Office.context.mailbox.item;
  
  verifySignature(item);
  item.to.getAsync(function(asyncResult){
    // var emails = asyncResult.value[0].emailAddress;
    var email_addresses = [];
    for (var i = 0; i < asyncResult.value.length; i++){
      email_addresses.push(asyncResult.value[i].emailAddress);
    }
    item.cc.getAsync(function(asyncResult){
      for (var i = 0; i < asyncResult.value.length; i++){
        email_addresses.push(asyncResult.value[i].emailAddress);
      }
      var result = verifyEmailAddress(email_addresses);
      if(result[0]){
        const errortMsg = {
          type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
          message: "ERROR: Multiple domains found: " + result[1].join(', '),
          icon: "Icon.80x80",
          persistent: true
        };
        item.notificationMessages.addAsync("action",errortMsg)
      }
      event.completed();
    });
  });

  // Be sure to indicate when the add-in command function is complete
  // Office.context.mailbox.item.notificationMessages.replaceAsync("action", message);
  // event.completed();
}

function verifyEmailAddress(arr) {
  // OVERVIEW: Parse through the array of email addresses
  // INPUT: arr - <array of str> : array of email addresses
  // OUTPUT: result - <array> 0: <bool> result, 1: <array of str> email domains if multiple are found
  var recipient_domains = {};
  var result = false;
  var domains = [];
  for (var i = 0; i < arr.length; i++) {
      var re = /\w*(\.com|\.net|\.org)/g;
      var domain = arr[i].match(re);
      if (!(domain in recipient_domains)) {
          recipient_domains[domain] = 1;
      }
      else {
          recipient_domains[domain]++;
      }
  }
  if ('novacoast.com' in recipient_domains) {
      delete recipient_domains['novacoast.com'];
  }
  if (Object.keys(recipient_domains).length > 1) {
      result = true;
      domains = Object.keys(recipient_domains);
  }
  return [result, domains];
}

function verifySignature(){
  // OVERVIEW: Parse the body of the email for signature

  var item = Office.context.mailbox.item;
  // retrieve body text
  item.body.getAsync(
    Office.CoercionType.Html,
    function(asyncResult){
      var inspect_body = verifySig(asyncResult.value);
      if(!inspect_body){
        const signature_message = {
          type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
          message: "Missing signature",
          icon: "Icon.80x80",
          persistent: true
        };
        item.notificationMessages.addAsync("signature", signature_message);
      }
    });
  }

  function verifySig(html){
      // OVERVIEW: Function to extract the html containing the signature, confirmed that there is signature
      // OUTOUT: true - missing signature, false - found signature
      var result = html.search(/\b(\w*Novacoast, Inc\w*)\b/g);
      if (result != -1){
          return true;
      }
      else{
          return false;
      } 
  }

// Necessary for the functionality of the code
function getGlobal() {
  return typeof self !== "undefined"
    ? self
    : typeof window !== "undefined"
    ? window
    : typeof global !== "undefined"
    ? global
    : undefined;
}

const g = getGlobal();

// the add-in command functions need to be available in global scope
g.action = action;
