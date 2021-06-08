/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office */

const Mustache = require("mustache");
const $ = require("jquery");
import 'bootstrap';
import './taskpane.scss';
import ko from 'knockout';
import 'webpack-jquery-ui/autocomplete';
import 'webpack-jquery-ui/css';
var currentTemplate = null;
var fieldValues = {};
var mustacheVars;

const escalationViewModel = function(){
    const self = this;
    self.emailTemplates = ko.observableArray([]);
    self.customers = ko.observableArray([]);
    self.selectedTemplate = ko.observable(null);
    self.selectedCustomer = ko.observable(null);
    self.templateFormSubmitted = ko.observable(false);
    self.getAutoCompleteObj = ko.computed(()=> {
        const options = {
            source: self.customers().map(cust => cust.customerName),
        }
        return options
    })
}

const escalationVM = new escalationViewModel();

Office.onReady(info => {
    if (info.host == Office.HostType.Outlook){
        ko.applyBindings(escalationVM);
        // document.getElementById("template-form").style.display = "none";
        // document.getElementById("form_btn").onclick = run2;
        // document.getElementById("customer_btn").onclick = run1;
        
        //get the customers from UNI 
        //TODO: Change this into API request
        const uniCustomerReponse = {
            "customers": [
              {
                "_id": "507f1f77bcf86cd799439011",
                "customerName": "Novacoast",
                "groupID": "507f1f77bcf86cd799439011",
                "webConsoleURL": [
                  "logrhythm.novacoast.com"
                ],
                "layoutID": "507f1f77bcf86cd799439011",
                "lastNotificationDate": "0001-01-01T00:00:00Z",
                "leads": [
                  {
                    "name": "John Doe",
                    "phone": "800-591-6964"
                  }
                ],
                "escalationContacts": [
                  {
                    "name": "John Doe",
                    "phone": "800-591-6964"
                  }
                ]
              }
            ],
            "total": 0
        }
        escalationVM.customers(uniCustomerReponse['customers']);

        // don't refresh page if the user presses enter while choosing template
        $('#escalation-email-form').submit(e => e.preventDefault())
        
        // display template upon change selection
        $('#email-template-list').change(() => {
            // selectedTemplate has a bindind in the html file
            currentTemplate = escalationVM.selectedTemplate()
            if (!currentTemplate) {
                escalationVM.templateFormSubmitted(false)
                return
            }
            console.log("selected a template");
            retrieveTemplate();
        });        
        
        $("#template-vars-form").submit((e)=>{
            e.preventDefault();
            formEmail();
        })
    }
});

ko.bindingHandlers.ko_autocomplete = {
	init(element) {
		$(element).autocomplete({
			select: (event, ui) => {
				const customer = ui.item.value
				// assign VM selectedCustomer value from option chosen in autocomplete
				const selCust = escalationVM.customers()
					.filter(cust => cust.customerName === customer)
				if (selCust.length > 0) {
					escalationVM.selectedCustomer(selCust[0])
					handleSelectedCustomer(selCust[0])
				}
			},
			change: (event, ui) => {
				if (ui.item === null) {
					escalationVM.selectedCustomer(null)
                    const currentCustomer = escalationVM.selectedCustomer()
					handleSelectedCustomer(currentCustomer)
				}
			},
		})
	},
	update(element, valueAccessor) {
		const options = valueAccessor()
		const valueUnwrapped = ko.unwrap(options)
		$(element).autocomplete('option', 'source', valueUnwrapped.source)
	},
}

function handleSelectedCustomer(customer) {
	if (!customer) {
		escalationVM.emailTemplates.removeAll()
		return
	}
	escalationVM.emailTemplates.removeAll()
	const customerID = customer._id

	// get customer's email templates
    escalationVM.emailTemplates.removeAll()

    //TODO: replace exampleTemplateReponse with UNI API call for templates
    const exampleTemplateResponse = {
            "_id": "507f1f77bcf86cd799439011",
            "name": "Default Incident",
            "toContacts": [
              {
                "name": "John Doe",
                "email": "name@customer.com"
              }
            ],
            "ccContacts": [
              {
                "name": "John Doe",
                "email": "name@customer.com"
              }
            ],
            "subject": "Alarm {{.Name}}",
            "bodyTemplate": {
              "_id": "507f1f77bcf86cd799439011",
              "name": "Standard",
              "content": "<html>\n  <head>\n    <style>\n      .ex {\n        font-size: 14px;\n      }\n    </style>\n  </head>\n  <body>\n    {{#var1 || var2}}\n    Case Reference\n    lorem ipsum {{var1}} {{var2}}\n    {{/var1 || var2}}\n    {{#field2}}\n    Event Details\n    lorem {{field2}} ipsum\n    {{/field2}}\n  </body>\n</html>\n"
            }
        }
    escalationVM.emailTemplates.push(exampleTemplateResponse);
	// api.getCustomerEmailTemplates(customerID, (templates) => {
	// 	templates.forEach(template => escalationVM.emailTemplates.push(template))
	// })
}

function retrieveTemplate(){
    // send the request here
    fetch("template-response-example.json")
    .then(response => response.json())
    .then(function(data){
        // form prompts:
        // const html = data["emailTemplates"][0]["bodyTemplate"]["content"];
        currentTemplate = data;
        // const templateParse = Mustache.parse(html).filter(v => v[0] === '#');
        // const firstLvlSectionNames = templateParse.map(x => x[1]);
        mustacheVars = parseJson(data);
        // formPrompt(mustacheVars);
        $('#template-form-body').empty()
        formPrompt(mustacheVars);
        $('#template-form-modal').modal('show')
    })
}


function formPrompt(mustacheVars){
    // OVERVIEW: Function to parse the html and prepares the second page in taskpane
    const formEntry = `
            <div class='row'>
                <div class='col-md-6'>
                    <div class="input-group mb-3">
                        <div class="input-group-prepend">
                            <div class="input-group-text">
                                <span for="subject">{{label}}</span>
                            </div>
                        </div>
                        <input type="text" class="form-control"
                        id="field-{{value}}" name="{{label}}" value="" >
                    </div>
                </div>
            </div>
            `;
    
    
    const template_fields = $("#template-form-body");
    // console.log(mustacheVars);
    Object.keys(mustacheVars).forEach(function(key) {
        mustacheVars[key].forEach((name) => {
            const field_name = name;
            // Add spaces between words
            const pretty_name = field_name.replace(/([a-z\d])([A-Z])/g, '$1 $2'); 
            const val = {
                label:pretty_name,
                value:field_name
            }
            const entryHTML = Mustache.render(formEntry,val);
            template_fields.append(entryHTML);
            fieldValues[field_name]= "field-"+field_name;
        });   
    });
}

function parseJson(json){
    //OVERVIEW: Function that parses the json data for template values
    const html = json["emailTemplates"][0]["bodyTemplate"]["content"];
    const templateParse = Mustache.parse(html).filter(v => v[0] === '#');
    const firstLvl = templateParse.map(x => x[1]);
    var mustacheVars = {};
    firstLvl.forEach((section) => {
		const variables = templateParse
			// looks at section name portion
			.filter(x => x[1] === section)
			// looks for sections within sections
			.map(x => x[4])[0].filter(y => y[0] === '#')
			// gets second level section names
			// which must be the same as the var within that section
			.map(z => z[1]);
		if (variables.length > 0) mustacheVars[section] = variables;
    });
    return mustacheVars;
}

function formEmail(){
    const data = currentTemplate;
    const html = data["emailTemplates"][0]["bodyTemplate"]["content"];
    const subject = data["emailTemplates"][0]["subject"];
    const cc_array = data["emailTemplates"][0]["ccContacts"];
    const to_array = data["emailTemplates"][0]["toContacts"];
    const template_values = pullValues();
    const completed_template = Mustache.render(html,template_values);
    bodyInjection(completed_template);
    subjectInjection(subject);
    recipientInjection(to_array,cc_array);
}

function pullValues(){
    // OVERVIEW: Function retrieves values from the prompts
    // console.log(fieldValues);

    //TODO: append values of the OR function
    const template_values = {};
    for (const key in fieldValues){
        const id = fieldValues[key];
        template_values[key] = document.getElementById(id).value;
    }
    Object.keys(mustacheVars).forEach((section) =>{
        if (mustacheVars[section].length >= 1){
            mustacheVars[section].forEach((key) =>{
                if (template_values[key]){
                    template_values[section] = true;
                }
            });
        }
    });
    return template_values;
}

function subjectInjection(subject){
    Office.context.mailbox.item.subject.setAsync(
        subject,
        { asyncContext: { var1: 1, var2: 2 } },
        function (asyncResult) {
            if (asyncResult.status == Office.AsyncResultStatus.Failed){
                write(asyncResult.error.message);
            }
            else {
                // Successfully set the subject.
                // Do whatever appropriate for your scenario
                // using the arguments var1 and var2 as applicable.
            }
        });
}

function bodyInjection(html){
    // Function that replaces the entire email body with message
    var item = Office.context.mailbox.item;
    item.body.getTypeAsync(
        function (result) {
            if (result.status == Office.AsyncResultStatus.Failed){
                write(result.error.message);
            }
            else {
                // Grab the text of email 
                // Successfully got the type of item body.
                // Set data of the appropriate type in body.
                if (result.value == Office.MailboxEnums.BodyType.Html) {
                    // Body is of HTML type.
                    // Specify HTML in the coercionType parameter
                    // of setSelectedDataAsync.
                    item.body.getAsync(
                        Office.CoercionType.Html,
                        function(asyncResult){
                          var inspect_body = verifyBody(asyncResult.value);
                          if(inspect_body){
                            //   extractSig(asyncResult.value);
                            // var newText = pullValues(asyncResult.value);
                            // item.body.setAsync(
                            //     newText,
                            //     { coercionType: Office.CoercionType.Html, 
                            //     asyncContext: { var3: 1, var4: 2 } },
                            //     function (asyncResult) {
                            //         if (asyncResult.status == 
                            //             Office.AsyncResultStatus.Failed){
                            //             write(asyncResult.error.message);
                            //         }
                            //         else {
                            //             // Successfully set data in item body.
                            //             // Do whatever appropriate for your scenario,
                            //             // using the arguments var3 and var4 as applicable.
                            //         }
                            //     });
                          }
                          else{
                            item.body.prependAsync(
                                html,
                                { coercionType: Office.CoercionType.Html, 
                                asyncContext: { var3: 1, var4: 2 } },
                                function (asyncResult) {
                                    if (asyncResult.status == 
                                        Office.AsyncResultStatus.Failed){
                                        write(asyncResult.error.message);
                                    }
                                    else {
                                        // Successfully set data in item body.
                                        // Do whatever appropriate for your scenario,
                                        // using the arguments var3 and var4 as applicable.
                                    }
                                });
                          }
                        });
                
                    
                }
                else {
                    // Body is of text type 
                    // NOTE: Currently does not appear to be relavent to the current implmentation of the app
                    const result_message = {
                        type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
                        message: "Text",
                        icon: "Icon.80x80",
                        persistent: true
                      };
                    item.notificationMessages.addAsync("result", result_message);
                    item.body.prependAsync(
                        html,
                        { coercionType: Office.CoercionType.Text, 
                            asyncContext: { var3: 1, var4: 2 } },
                        function (asyncResult) {
                            if (asyncResult.status == 
                                Office.AsyncResultStatus.Failed){
                                write(asyncResult.error.message);
                        }
                            else {
                                // Successfully set data in item body.
                                // Do whatever appropriate for your scenario,
                                // using the arguments var3 and var4 as applicable.
                            }
                         });
                }
            }
        });
}

function recipientInjection(to_array,cc_array){
    // OVERVIEW: Processing the arrays into Office JS format
    
    var to_recipients = [];
    var cc_recipients = [];
    for (var i = 0;i<to_array.length;i++){
        var entry = {};
        entry["emailAddress"] = to_array[i]["email"];
        entry["displayName"] = to_array[i]["name"];
        to_recipients.push(entry);
    }
    for (var i = 0;i<cc_array.length;i++){
        var entry = {};
        entry["emailAddress"] = cc_array[i]["email"];
        entry["displayName"] = cc_array[i]["name"];
        cc_recipients.push(entry);
    }
    Office.context.mailbox.item.to.setAsync(to_recipients, function(result) {
        if (result.error) {
            console.log(result.error);
        } else {
        }
    });
    Office.context.mailbox.item.cc.setAsync(cc_recipients, function(result) {
        if (result.error) {
            console.log(result.error);
        } else {
        }
    });
}

function verifyBody(html){
    // OVERVIEW: Function to verify if the user already formed a template
    // OUTPUT: True - template is found, false - template is not found
    // console.log(html)
    var result = html.search(/\b(\w*Case\w*)\b/g);
    // console.log("result: " + String(result))
    if (result != -1){
        return true;
    }
    else{
        return false;
    } 
}

