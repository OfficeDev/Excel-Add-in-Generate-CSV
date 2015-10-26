/* Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. 
    See full license at the bottom of this file. */


(function () {
    "use strict";

    var sheetCopyNumber = 1;
    var selectedService = "Moodle";
    var rosterName = "";

	// The initialize function must be run each time a new page is loaded
	Office.initialize = function (reason) {
		$(document).ready(function (){
			app.initialize();
			$('#generate-template').button();
			$('#generate-template').click(generateTemplateTable);
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

    /*******************************************/
    /* Open a pop-up window with the steps to export a csv */
    /*******************************************/
	function showHelp() {
	    var helpWindow = window.open("HelpPop.html", "mywindow", "menubar=1,resizable=1,width=800,height=850");
	}

    /*******************************************/
    /* Populate worksheet with students for chosen tool */
    /*******************************************/
	function generateTemplateTable() {
	    // Run a batch operation against the Excel object model
	    Excel.run(function (ctx) {
	        // Run the queued-up commands, and return a promise to indicate task completion
	        // Create a proxy object for the active worksheet

	        var studentRoster = ctx.workbook.worksheets.add("_" + sheetCopyNumber);
	        rosterName = selectedService + "Roster_" + sheetCopyNumber;


	        /******************************************************/
	        /* To add more columns to your roster table, add the column name 
            /*  to the 3rd param of buildRosterRange.     
            /*  The fillRoster method creates a row of "fake" student data 
            /* with a column value for each column name given in the 
            /* buildRoster method arguments. Be sure to add a case
            /* statement in fillRoster for your new column header
            /*
            /* TODO: Replace array of column headers with a data structure
            /* that contains column headers and default "fake" row values
	        /******************************************************/
	        if (selectedService == "Moodle") {
	            buildRosterTable(studentRoster, [["ACTION", "ROLE", "USER ID NUMBER", "COURSE ID NUMBER"]]);
            }
	        else {
	            buildRosterTable(studentRoster, [["FIRST NAME", "LAST NAME", "EMAIL", "PARENTEMAIL", "PARENTPHONE"]]);
            }

	        sheetCopyNumber++;

	        return ctx.sync().then(function () {

                //Fill the table created by the buildRosterRange function.
	            fillRoster(rosterName);
	            app.showNotification("Sheet created");
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
	function fillRoster(rosterName) {

	    // Run a batch operation against the Excel object model
	    Excel.run(function (ctx) {
	        // Create a proxy object for the worksheets collection 
	        var worksheets = ctx.workbook.worksheets;
	        var table;
	        var headerRange;

	        // Queue a command to get the sheet with the name of the clicked button
	        var clickedSheet = worksheets.getItem(rosterName);

            //add batch command to load the value of the worsheet.tables property
	        clickedSheet.load("tables");

            //Run the batched commands
	        return ctx.sync()
                .then(function () {

                    //Get a table from the returned tables property value
                    table = clickedSheet.tables.getItemAt(0);

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
                        })

                        //Run the batched commands
                        .then(ctx.sync)
                            .then(function () {

                                //loop through the loaded header range values
                                var headers = headerRange.values;
                                for (var i = 0; i < headers.length; i++) {
                                    var value = headers[i];
                                    for (var j = 0; j < value.length; j++) {

                                        switch (value[j]) {
                                            case "FIRST NAME":
                                                clickedSheet.getCell(1, j).values = "Alex";
                                                break;
                                            case "LAST NAME":
                                                clickedSheet.getCell(1, j).values = "Dunsmuir";
                                                break;
                                            case "EMAIL":
                                                clickedSheet.getCell(1, j).values = "adamd@patsoldemo6.com";
                                                break;
                                            case "PARENTEMAIL":
                                                clickedSheet.getCell(1, j).values = "parent@patsoldemo6.com";
                                                break;
                                            case "PARENTPHONE":
                                                clickedSheet.getCell(1, j).values = "555 111-2222";
                                                break;
                                            case "ACTION":
                                                clickedSheet.getCell(1, j).values = "add";
                                                break;
                                            case "ROLE":
                                                clickedSheet.getCell(1, j).values = "student";
                                                break;
                                            case "USER ID NUMBER":
                                                clickedSheet.getCell(1, j).values = "123a";
                                                break;
                                            case "COURSE ID NUMBER":
                                                clickedSheet.getCell(1, j).values = "econ 101";
                                                break;
                                        }
                                    }
                                }
                                // Queue a command to activate the clicked sheet
                                clickedSheet.activate();
                        })
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
    function buildRosterTable( studentRoster, headerValues) {

        // Create a proxy object for the active worksheet
        studentRoster.name = rosterName;

        var tableRangeString = "A1:";
        switch (headerValues[0].length) {
            case 1:
                tableRangeString += "A2";
                break;
            case 2:
                tableRangeString += "B2";
                break;
            case 3:
                tableRangeString += "C2";
                break;
            case 4:
                tableRangeString += "D2";
                break;
            case 5:
                tableRangeString += "E2";
                break;
            case 6:
                tableRangeString += "F2";
                break;
            case 7:
                tableRangeString += "G2";
                break;
            case 8:
                tableRangeString += "H2";
                break;
            case 9:
                tableRangeString += "I2";
                break;
            case 10:
                tableRangeString += "J2";
                break;
        }

        // Queue a command to add a new table
        var table = studentRoster.tables.add(tableRangeString, true);
        table.name = rosterName;

        // Queue a command to get the newly added table
        table.getHeaderRowRange().values = headerValues;
        table.style = "TableStyleLight20";
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