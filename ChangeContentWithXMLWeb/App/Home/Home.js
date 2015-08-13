/// <reference path="../App.js" />

// This function is run when the app is ready to start interacting with the host application
// It ensures the DOM is ready before adding click handlers to buttons
Office.initialize = function (reason) {
    $(document).ready(function () {

        // Wire up the click events of the two buttons in the WD_OpenXML_js.html page.
        $('#getOOXMLData').click(function () { getOOXML(); });
        $('#setOOXMLData').click(function () { setOOXML(); });
    });
};

// Variable to hold any Office Open XML
var currentOOXML = "";

function getOOXML() {
    // Get a reference to the <DIV> where we will write the status of our operation
    var report = document.getElementById("status");

    // Remove all nodes from the status <DIV> so we have a clean space to write to
    while (report.hasChildNodes()) {
        report.removeChild(report.lastChild);
    }

    // Now we can begin the process.
    // Step 1 is to call the getSelectedDataAsync method. The first parameter is the coercion
    // type (in our case Ooxml)
    // The second parameter is a set of options: The first option determines whether to include formating 
    // and the second determines whether to include only visible elements, or whether to include them all.
    // When the method returns, the function that is provided as the third parameter will run.
    Office.context.document.getSelectedDataAsync(
        Office.CoercionType.Ooxml,
        { valueFormat: Office.ValueFormat.Formatted, filterType: Office.FilterType.All },
        function (result) {
            // Get a referencene to our TEXTAREA element
            // that is located just below the buttons in the WD_OpenXML_js.html page.
            var textArea = document.getElementById("dataOOXML");
            if (result.status == "succeeded") {

                // If the getSelectedDataAsync call succeeded, then
                // result.value will return a valid chunk of OOXML, which we'll
                // hold in the currentOOXML variable.
                currentOOXML = result.value;

                // Now for some fun: We'll take the actual raw OOXML data and let the user
                // actually see it! We'll do this by putting the raw OOXML data in the TextArea.
                // (NOTE --- You probably wouldn't want to do this in a real App
                // but it is a great way for ************YOU AS A DEVELOPER********** 
                // to learn OOXML. In other words, simply make a document look
                // how you want it, select all or some of the document,
                // and then use this App so see what your selection looks like in OOXML!!)

                while (textArea.hasChildNodes()) {
                    textArea.removeChild(textArea.lastChild);
                }
                textArea.appendChild(document.createTextNode(currentOOXML));

                // Tell the user we succeeded
                report.innerText = "Got It --- Success!!";
            }
            else {
                // This runs if the getSelectedDataAsync method does not return a success flag
                currentOOXML = "";
                report.innerText = result.error.message;
            }
        });
}

function setOOXML() {
    // Get a reference to the <DIV> where we will write the outcome of our operation
    var report = document.getElementById("status");
    // Remove all nodes from the status <DIV> so we have a clean space to write to
    while (report.hasChildNodes()) {
        report.removeChild(report.lastChild);
    }

    // Check whether we have previosuly extracted OOXML
    if (currentOOXML != "") {

        // Call the setSelectedDataAsync, with parameters of:
        // 1. The Data to insert.
        // 2. The coercion type for that data.
        // 3. A callback function that lets us know if it succeeded.
        Office.context.document.setSelectedDataAsync(
            currentOOXML,
            { coercionType: Office.CoercionType.Ooxml },
            function (result) {
                // Update the report element
                if (result.status == "succeeded") {
                    report.innerText = "Set It --- Success!!";
                }
                else {
                    // This runs if the getSliceAsync method does not return a success flag
                    report.innerText = result.error.message;

                    // Clear the text area just so we don't give you the impression that there's
                    // valid OOXML waiting to be inserted... 
                    while (textArea.hasChildNodes()) {
                        textArea.removeChild(textArea.lastChild);
                    }
                }
            });
    }
    else {

        // If currentOOXML == "" then we should not even try to insert it, because
        // that is gauranteed to cause an exception, needlessly.
        report.innerText = "There is currently no OOXML to insert!"
            + " Please select some of your document and click [Get OOXML] first!";
    }
}

