/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

// images references in the manifest
import "../../assets/icon-16.png";
import "../../assets/icon-32.png";
import "../../assets/icon-80.png";

/* global document, Office */

Office.onReady(info => {
    if (info.host === Office.HostType.Outlook) {
        document.getElementById("sideload-msg").style.display = "none";
        document.getElementById("app-body").style.display = "flex";
        document.getElementById("run").onclick = run;
    }
});

export async function run() {
    // Get a reference to the current message
    var item = Office.context.mailbox.item;
    appendLine("Subject", item.subject);
    appendLine("To", item.to[0].emailAddress);
    appendLine("RecipientType", item.to[0].recipientType);
    appendLine("DisplayName", item.to[0].displayName);
    appendLine("Body", JSON.stringify(item.body.));
}

function appendLine(label, content) {
    document.getElementById("item-subject").innerHTML += "<b>"+label+":</b> <br/>" + content +"<br/>";
}
