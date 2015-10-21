/* Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. 
    See full license at the bottom of this file. */

/// <reference path="../App.js" />

(function () {
    "use strict";

    var sheetCopyNumber = 1;

	// The initialize function must be run each time a new page is loaded
	Office.initialize = function (reason) {
		$(document).ready(function () {
			app.initialize();

			$('#generate-template').button();
			$('#generate-template').click(generateTemplateRange);

			$('#create-csv').button();
			$('#create-csv').click(createCSVStream);
			$('#show-help').click(showHelp);
		});
	};


	function showHelp() {
	    window.open("HelpPop.html","mywindow","menubar=1,resizable=1,width=550,height=650");

	}
	function generateTemplateRange() {
	    // Run a batch operation against the Excel object model
	    Excel.run(function (ctx) {
	        // Run the queued-up commands, and return a promise to indicate task completion
	        // Create a proxy object for the active worksheet

	        var studentRoster = ctx.workbook.worksheets.add("_" + sheetCopyNumber);

            //Get user's service choice and build the cooresponding table
	        if ($("input[type='radio']:checked").val() == "Moodle") {
	            buildMoodleRange( studentRoster);
            }
	        else {
	            buildTeacherKitRange(studentRoster);
	        }
	        sheetCopyNumber++;

	        return ctx.sync();
	    }).catch(function (error) {
	        // Always be sure to catch any accumulated errors that bubble up from the Excel.run execution
	        app.showNotification("Error: " + error);
	        console.log("Error: " + error);
	        if (error instanceof OfficeExtension.Error) {
	            console.log("Debug info: " + JSON.stringify(error.debugInfo));
	        }
	    });
    }

    function buildMoodleRange( studentRoster) {

        // Create a proxy object for the active worksheet
         studentRoster.name = "MoodleRoster_" +sheetCopyNumber;

        // Queue a command to add a new table
        var table = studentRoster.tables.add('A1:D2', true);
        table.name = "moodelRosterTable_"+sheetCopyNumber;

        // Queue a command to get the newly added table
        table.getHeaderRowRange().values = [["ACTION", "ROLE", "USER ID NUMBER", "COURSE ID NUMBER"]];
        table.style = "TableStyleLight20";
    }

    function buildTeacherKitRange(studentRoster) {
        // Create a proxy object for the active worksheet
        studentRoster.name = "TeacherKitRoster";
  
        // Queue a command to add a new table
        var table = studentRoster.tables.add('A1:E2', true);
        table.name = "teacherKitRosterTable_" + sheetCopyNumber;

        // Queue a command to get the newly added table
        table.getHeaderRowRange().values = [["FIRST NAME", "LAST NAME", "EMAIL", "PARENTEMAIL", "PARENTPHONE"]];
        table.style = "TableStyleLight21";
    }
 
    function createCSVStream() {
        Excel.run(function (ctx) {
            var range = ctx.workbook.worksheets.getActiveWorksheet().getUsedRange();
            range.load("values");
            return ctx.sync()
                .then(function () {
                    var CSVString = "";

                    //Iterate the rows in the range
                    for (var i = 0; i < range.values.length; i++) {
                        var value = range.values[i];

                        //Iterate the cells in a row
                        for (var j = 0; j < value.length; j++) {
                            //Append a value and comma
                            CSVString = CSVString + value[j] + ",";
                        }

                        //strip the trailing ','
                        CSVString = CSVString.substr(0, CSVString.length - 1);

                        //append CRLF
                        CSVString = CSVString + "\r\n";
                    }
                    app.showNotification(CSVString);
                })
        });
    }
    
	
	/********************/
    /* Helper functions */
    /********************/

       // Helper for calls to the service. 
    function httpGetAsync(theUrl, callback)
    {
        var request = new XMLHttpRequest();
        request.open("GET", theUrl, true);
        request.onreadystatechange = function() { 
            if (request.readyState == 4 && request.status == 200)
                callback(request.responseText);
        }
        request.send(null);
    }
	
   // Helper that processes file names into an array. This is because the service returns
    // the file names as ["filename1.docx","filename2.docx","filename3.docx"].
    function processResponse(rawResponse) {
        
        // Remove quotes.
        rawResponse = rawResponse.replace(/"/g, "");
        
        // Remove opening brackets.
        rawResponse = rawResponse.replace("[", "");
        
        // Remove closing brackets.
        rawResponse = rawResponse.replace("]", "");
        
        // Return an array of file names.
        return rawResponse.split(',');
    }
    // Helper for calls to the service. 
    function httpGetAsync(theUrl, callback) {
        var request = new XMLHttpRequest();
        request.open("GET", theUrl, true);
        request.onreadystatechange = function () {
            if (request.readyState == 4 && request.status == 200)
                callback(request.responseText);
        }
        request.send(null);
    }
})();

/* 
Excel-Add-in-Generate-CSV, https://github.com/OfficeDev/Excel-Add-in-Generate-CSV

Copyright (c) Microsoft Corporation

All rights reserved.

MIT License:

Permission is hereby granted, free of charge, to any person obtaining a copy of this software and
associated documentation files (the "Software"), to deal in the Software without restriction, including
without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the
following conditions:

The above copyright notice and this permission notice shall be included in all copies or substantial
portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT
LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT
SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN
ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE
USE OR OTHER DEALINGS IN THE SOFTWARE.
*/