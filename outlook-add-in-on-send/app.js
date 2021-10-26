/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

    var mailboxItem;
    var _event;
    var _dialog;

    Office.initialize = function (reason) {
        mailboxItem = Office.context.mailbox.item;
    }

    function showDialog(event) {
        _event = event;
        Office.context.ui.displayDialogAsync("https://victorious-ocean-057516210.azurestaticapps.net/dialog.html", { height: 30, width: 30, displayInIframe: true, promptBeforeOpen: true }, dialogCallback);
    }

    function dialogCallback(asyncResult) {
        if (asyncResult.status === Office.AsyncResultStatus.Failed) {
            console.log("Error displaying dialog: " + asyncResult.error.code);
            _event.completed({ allowEvent: true });
        }
        else {
            _dialog = asyncResult.value;
            _dialog.addEventHandler(Office.EventType.DialogMessageReceived, messageHandler);
            _dialog.addEventHandler(Office.EventType.DialogEventReceived, eventHandler);
        }
    }

    function messageHandler(arg) {
        _dialog.close();
        _dialog = null;
    
        if (arg !== undefined && arg.message == 'prepend') {
            prependBody();
        } else {
            _event.completed({ allowEvent: true });
        }
    }

    function eventHandler(arg) {
        _dialog.close();
        _dialog = null;
    
        if (arg !== undefined) {
            console.log(arg.error + ': Error in dialog window');
        } else {
            console.log('Undefined error in dialog window');
        }
    
        _event.completed({ allowEvent: false });
    }

    function prependBody() {
        mailboxItem.body.prependAsync('<p>Hello World</p>', { coercionType: Office.CoercionType.Html }, prependBodyCallback);
    }

    function prependBodyCallback(asyncResult) {
        if (asyncResult.status == Office.AsyncResultStatus.Failed) {
            mailboxItem.notificationMessages.addAsync('NoSend', { type: 'errorMessage', message: 'Failed prepending the body upon send: ' + asyncResult.error });
            _event.completed({ allowEvent: false });
        } else {
            _event.completed({ allowEvent: true });
        }
    }

    // Entry point for Contoso Message Body Checker add-in before send is allowed.
    // <param name="event">MessageSend event is automatically passed by BlockOnSend code to the function specified in the manifest.</param>
    function validateBody(event) {
        mailboxItem.body.getAsync("html", { asyncContext: event }, checkBodyOnlyOnSendCallBack);
    }

    // Invoke by Contoso Subject and CC Checker add-in before send is allowed.
    // <param name="event">MessageSend event is automatically passed by BlockOnSend code to the function specified in the manifest.</param>
    function validateSubjectAndCC(event) {
        shouldChangeSubjectOnSend(event);
    }

    // Check if the subject should be changed. If it is already changed allow send. Otherwise change it.
    // <param name="event">MessageSend event passed from the calling function.</param>
    function shouldChangeSubjectOnSend(event) {
        mailboxItem.subject.getAsync(
            { asyncContext: event },
            function (asyncResult) {
                addCCOnSend(asyncResult.asyncContext);
                //console.log(asyncResult.value);
                // Match string.
                var checkSubject = (new RegExp(/\[Checked\]/)).test(asyncResult.value)
                // Add [Checked]: to subject line.
                subject = '[Checked]: ' + asyncResult.value;

                // Check if a string is blank, null or undefined.
                // If yes, block send and display information bar to notify sender to add a subject.
                if (asyncResult.value === null || (/^\s*$/).test(asyncResult.value)) {
                    mailboxItem.notificationMessages.addAsync('NoSend', { type: 'errorMessage', message: 'Please enter a subject for this email.' });
                    asyncResult.asyncContext.completed({ allowEvent: false });
                }
                else {
                    // If can't find a [Checked]: string match in subject, call subjectOnSendChange function.
                    if (!checkSubject) {
                        subjectOnSendChange(subject, asyncResult.asyncContext);
                        //console.log(checkSubject);
                    }
                    else {
                        // Allow send.
                        asyncResult.asyncContext.completed({ allowEvent: true });
                    }
                }

            }
          )
    }

    // Add a CC to the email.  In this example, CC contoso@contoso.onmicrosoft.com
    // <param name="event">MessageSend event passed from calling function</param>
    function addCCOnSend(event) {
        mailboxItem.cc.setAsync(['Contoso@contoso.onmicrosoft.com'], { asyncContext: event });        
    }

    // Check if the subject should be changed. If it is already changed allow send, otherwise change it.
    // <param name="subject">Subject to set.</param>
    // <param name="event">MessageSend event passed from the calling function.</param>
    function subjectOnSendChange(subject, event) {
        mailboxItem.subject.setAsync(
            subject,
            { asyncContext: event },
            function (asyncResult) {
                if (asyncResult.status == Office.AsyncResultStatus.Failed) {
                    mailboxItem.notificationMessages.addAsync('NoSend', { type: 'errorMessage', message: 'Unable to set the subject.' });

                    // Block send.
                    asyncResult.asyncContext.completed({ allowEvent: false });
                }
                else {
                    // Allow send.
                    asyncResult.asyncContext.completed({ allowEvent: true });
                }

            });
    }

    // Check if the body contains a specific set of blocked words. If it contains the blocked words, block email from being sent. Otherwise allows sending.
    // <param name="asyncResult">MessageSend event passed from the calling function.</param>
    function checkBodyOnlyOnSendCallBack(asyncResult) {
        var listOfBlockedWords = new Array("blockedword", "blockedword1", "blockedword2");
        var wordExpression = listOfBlockedWords.join('|');

        // \b to perform a "whole words only" search using a regular expression in the form of \bword\b.
        // i to perform case-insensitive search.
        var regexCheck = new RegExp('\\b(' + wordExpression + ')\\b', 'i');
        var checkBody = regexCheck.test(asyncResult.value);

        if (checkBody) {
            mailboxItem.notificationMessages.addAsync('NoSend', { type: 'errorMessage', message: 'Blocked words have been found in the body of this email. Please remove them.' });
            // Block send.
            asyncResult.asyncContext.completed({ allowEvent: false });
        }
        else {

            // Allow send.
            asyncResult.asyncContext.completed({ allowEvent: true });
        }
    }
