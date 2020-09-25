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


    function toggleProtection(args) {
        Excel.run(function (context) {

            // TODO1: Queue commands to reverse the protection status of the current worksheet.
            var sheet = context.workbook.worksheets.getActiveWorksheet();

            // TODO2: Queue command to load the sheet's "protection.protected" property from
            //        the document and re-synchronize the document and task pane.
            sheet.load('protection/protected');
            return context.sync()
                .then(
                    function () {
                        // TODO3: Move the queued toggle logic here.
                        if (sheet.protection.protected) {
                            sheet.protection.unprotect();
                        } else {
                            sheet.protection.protect();
                        }
                    }
                )
                // TODO4: Move the final call of `context.sync` here and ensure that it
                //        does not run until the toggle logic has been queued.
                .then(context.sync);
        })
            .catch(function (error) {
                console.log("Error: " + error);
                if (error instanceof OfficeExtension.Error) {
                    console.log("Debug info: " + JSON.stringify(error.debugInfo));
                }
            });
        args.completed();
    }

    // Show a notification message
    Office.context.mailbox.item.notificationMessages.replaceAsync("action", message);

    // Be sure to indicate when the add-in command function is complete
    event.completed();
}

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
g.toggleProtection = toggleProtection;