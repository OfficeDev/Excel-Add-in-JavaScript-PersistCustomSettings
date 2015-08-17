/*
* Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license at the bottom of this file.
*/

//// Add a dynamically-created div "toast" for displaying errors to the user.
var Toast = (function () {

    var Toast = "divToast",
        Close = "btnClose",
        Notice = "lblNotice",
        Output = "lblOutput";

    // Show the toast with the specified information.
    function showToast(title, message) {

        if (document.getElementById(Toast) == null) {
            createToast();
        }

        document.getElementById(Notice).innerText = title;
        document.getElementById(Output).innerText = message;

        // Provide implementations for solutions with and without jQuery.
        if ($) {
            $("#" + Toast).hide();
            $("#" + Toast).show("slow");
        }
        else {
            Toast.style.display = "none";
            Toast.style.display = "block";
        }
    }

    // Create the display for the toast.
    function createToast() {
        var divToast, lblClose, btnClose, divOutput;

        // Define an HTML element builder function.
        var nodeBuilder = function (elementType, options, text) {
            var newElement = document.createElement(elementType);
            for (var i = 0; i < options.length; i++) {
                newElement.setAttribute(options[i][0], options[i][1]);
            }
            if (text) { newElement.innerHTML = text; }
            return newElement;
        }

        // Create the container div.
        var toastStyle = "background-color:rgba(128, 128, 128, 0.77);" +
            "position:absolute;" +
            "bottom:0px;" +
            "width:90%;" +
            "text-align:center;" +
            "font-size:11pt;";
        divToast = nodeBuilder("div", [["style", toastStyle], ["id", Toast]]);

        // Create the close button as a div with a single span.
        var btnStyle = "text-align:right;" +
            "padding-right:10px;" +
            "font-size:10pt;" +
            "cursor:default";
        lblClose = nodeBuilder("div", [["style", btnStyle], ["id", Close]], "CLOSE ");
        btnClose = nodeBuilder("span", [["style", "cursor:pointer;"], ["onclick", "Toast.close()"]], "X");
        lblClose.appendChild(btnClose);

        // Create the div with the toast title and message.
        divOutput = nodeBuilder("div", [["id", "divOutput"], ["style", "margin-top:0px"]],
            "<span id='lblNotice' style='font-weight:bold;margin-top:0px;'></span><br />" +
            "<span id='lblOutput'></span>");

        // Add the child nodes to the toast div.
        divToast.appendChild(lblClose);
        divToast.appendChild(divOutput);

        // Add the toast div to the document body.
        document.body.appendChild(divToast);
    }

    // Close the toast.
    function close() {

        // Provide implementations for solutions with and without jQuery.
        if ($) {
            $("#" + Toast).hide("slow");
        }
        else {
            Toast.style.display = "none";
        }
    }

    return {
        showToast: showToast,
        close: close
    };

})();

// *********************************************************
//
// Excel-Add-in-JavaScript-PersistCustomSettings, https://github.com/OfficeDev/Excel-Add-in-JavaScript-PersistCustomSettings/
//
// Copyright (c) Microsoft Corporation
// All rights reserved.
//
// MIT License:
// Permission is hereby granted, free of charge, to any person obtaining
// a copy of this software and associated documentation files (the
// "Software"), to deal in the Software without restriction, including
// without limitation the rights to use, copy, modify, merge, publish,
// distribute, sublicense, and/or sell copies of the Software, and to
// permit persons to whom the Software is furnished to do so, subject to
// the following conditions:
//
// The above copyright notice and this permission notice shall be
// included in all copies or substantial portions of the Software.
//
// THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND,
// EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF
// MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND
// NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE
// LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION
// OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION
// WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
//
// *********************************************************