/* Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. 
    See full license at the bottom of this file. */


(function () {
    "use strict";

    var sheetCopyNumber = 1;
    var selectedService = "Moodle";
    var rosterName = "";
    var tableCreated = "false";

	// The initialize function must be run each time a new page is loaded
	Office.initialize = function (reason) {
		$(document).ready(function (){
		    app.initialize();

		    // If not using Excel 2016, return
		    if (!Office.context.requirements.isSetSupported('ExcelApi', '1.1')) {
		        app.showNotification("Need Office 2016 or greater", "Sorry, this add-in only works with newer versions of Excel.");
		        return;
		    }

			$('#generate-template').button();
			$('#generate-template').click(generate_templateClickHandler);
			$('#show-help').click(showHelp);
			$("#selectService").change(selectServiceHandler);
			$(".ms-Dropdown").Dropdown();
		});
	};

    /*******************************************/
    /* Change handler for service dropdown. Get the selected */
    /*  service value                                                            */
    /*******************************************/
	function selectServiceHandler() {
	    selectedService =$(this).val();
	}

	function generate_templateClickHandler() {

	    Excel.run(function (ctx) {
	        ctx.workbook.load("tables");
	        return ctx.sync().then(function () {
	            if (ctx.workbook.tables.count == 0) {
	                switch (selectedService) {
	                    case "Moodle":
	                        generateTemplateTable([["ACTION", "ROLE", "USER ID NUMBER", "COURSE ID NUMBER"]], [["add", "student", "usr-1", "econ 101"]])
	                        break;
	                    case "TeacherKit":
	                        generateTemplateTable([["FIRST NAME", "LAST NAME", "EMAIL", "PARENTEMAIL", "PARENTPHONE"]], [["Alex", "Dunsmuir", "alexd@patsoldemo6.com", "parent@home.com", "555-1212"]])
	                        break;
	                    case "MyClassroom":
	                        generateTemplateTable([["INSTRUCTOR", "STUDENT LAST NAME", "STUDENT FIRST NAME", "EMAIL", "PARENTEMAIL", "PARENTPHONE"]], [["Smith", "Dunsmuir", "Alex", "alexd@patsoldemo6.com", "parent@home.com", "555-1212"]])
	                        break;
	                }

	            }
	            else {
	                app.showNotification("Delete the existing table before creating a new one.");
	            }

	        })
	    });

	}

    /*******************************************/
    /* Open a pop-up window with the steps to export a csv */
    /*******************************************/
	function showHelp() {
	    var helpWindow = window.open("HelpPop.html", "mywindow", "menubar=1,resizable=1,width=800,height=850");
	}

    /*******************************************/
    /* Populate worksheet with students for chosen tool */
    /*******************************************/
	function generateTemplateTable(headerString, defaultTableValues) {
	    // Run a batch operation against the Excel object model
	    Excel.run(function (ctx) {
	        // Run the queued-up commands, and return a promise to indicate task completion

	        //Create a new worksheet for the selected service
	        var studentRoster = ctx.workbook.worksheets.getActiveWorksheet();
	        studentRoster.name = selectedService;
	        rosterName = selectedService;
	        buildRosterTable(studentRoster,  headerString);
	        return ctx.sync()
                //Run the batched commands
                .then(ctx.sync)
                    .then(function () {
                        //Fill the table created by the buildRosterRange function.
                        fillRoster(rosterName, defaultTableValues);
                        tableCreated = "true";
                        
                    });
	    }).catch(function (error) {
	        // Always be sure to catch any accumulated errors that bubble up from the Excel.run execution
	        app.showNotification("Error: " + error);
	        console.log("Error: " + error);
	        if (error instanceof OfficeExtension.Error) {
	            console.log("Debug info: " + JSON.stringify(error.debugInfo));
	        }
	    });
    }


    /*****************************************/
    /* Fill the roster table with fake student data        */
    /*****************************************/
	function fillRoster(rosterName, defaultValues) {

	    // Run a batch operation against the Excel object model
	    Excel.run(function (ctx) {
	        // Create a proxy object for the worksheets collection 
	        var worksheets = ctx.workbook.worksheets;
	        var table;
	        var headerRange;

	        // Queue a command to get the sheet with the name of the clicked button
	        var clickedSheet = ctx.workbook.worksheets.getActiveWorksheet();

            //add batch command to load the value of the worsheet.tables property
	        clickedSheet.load("tables");
	        //add batch command to load the value of the worsheet.tables property

            //Run the batched commands
	        return ctx.sync()
                .then(function () {
                    //Get a table from the returned tables property value
                    table = clickedSheet.tables.getItem(selectedService);

                    //add batch command to load the value of the table rows collection property
                    table.load("rows");
                })
                    //Run the batched commands
                    .then(ctx.sync)
                        .then(function () {

                            //Get the range of the loaded table header row
                            headerRange = table.getHeaderRowRange();

                            //Add a command to load the values of the header range
                            headerRange.load("values")
                            var tableRows = table.rows;
                            tableRows.add(null, defaultValues);

                        })

                        //Run the batched commands
                        .then(ctx.sync)
	        //Run the queued-up commands, and return a promise to indicate task completion
	        return ctx.sync();
	    })
		.catch(function (error) {
		    // Always be sure to catch any accumulated errors that bubble up from the Excel.run execution
		    app.showNotification("Error: " + error);
		    console.log("Error: " + error);
		    if (error instanceof OfficeExtension.Error) {
		        console.log("Debug info: " + JSON.stringify(error.debugInfo));
		    }
		});
	}

    /*****************************************/
    /* Create the roster table in the active worksheet */
    /*****************************************/
	function buildRosterTable(studentRoster, headerValues) {

        var tableRangeString =  "A1:" + columnName(headerValues[0].length-1) + "1";
        var table = studentRoster.tables.add(tableRangeString, true);
        table.name = rosterName;

        // Queue a command to get the newly added table
        table.getHeaderRowRange().values = headerValues;
        table.style = "TableStyleLight20";
    }
 
    /**
      * Returns the column name based on a zero-based column index.
      * For example, columnName(4) = 5th column = "E". Meanwhile, columnName(1000) = 1001st column = "ALM".
      * @param index Zero-based column index.
      * @returns {String} Locale-independent column name (e.g., a string comprised of one or more letters in the range "A:Z").
      */
	function columnName(index) {
	    if (typeof index !== 'number' || isNaN(index) || index < 0) {
	        throw new OfficeExtension.Error("InvalidArgument", "The parameter for Excel.columnName(x) must be positive and numeric.", [], { errorLocation: "Excel.Util.columnName" });
	    }
	    var letters = [];
	    while (index >= 0) {
	        letters.push(getSingleLetter(index % 26));
	        index = Math.floor(index / 26) - 1;
	    }
	    return letters.reverse().join('');
	    function getSingleLetter(zeroThrough25Index) {
	        return String.fromCharCode(zeroThrough25Index + 65); // ASCII code for "A" is 65
	    }
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