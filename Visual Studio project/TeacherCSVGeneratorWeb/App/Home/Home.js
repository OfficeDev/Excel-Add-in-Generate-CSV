/* Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. 
    See full license at the bottom of this file. */
(function () {
    "use strict";

    var selectedService = "Moodle";
    var rosterName = "";

    // The initialize function must be run each time a new page is loaded
    Office.initialize = function (reason) {
        $(document).ready(function () {
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

    /*********************************************************/
    /* Change handler for service dropdown. Get the selected */
    /*  service value                                                            */
    /*********************************************************/
    function selectServiceHandler() {
        selectedService = $(this).val();
    }

    function generate_templateClickHandler() {

        /***********************************************/
        /*Check for existing tables and then either 
        /*generate a new table or warn the user that
        /*there is an existing table that may have data
        /***********************************************/
        Excel.run(function (ctx) {
            //ctx.workbook.load("tables");
            var activeSheet = ctx.workbook.worksheets.getActiveWorksheet();
            ctx.load(activeSheet.tables, "name");
            return ctx.sync().then(function () {
                if (activeSheet.tables.count == 0) {
                    switch (selectedService) {
                        case "Moodle":
                            generateTemplateTable([["ACTION", "ROLE", "USER ID NUMBER", "COURSE ID NUMBER"]],
                                [["add", "student", "usr-1", "econ 101"]])
                            break;
                        case "TeacherKit":
                            generateTemplateTable([["FIRST NAME", "LAST NAME", "EMAIL", "PARENTEMAIL", "PARENTPHONE"]],
                                [["Alex", "Dunsmuir", "alexd@patsoldemo6.com", "parent@home.com", "555-1212"]])
                            break;
                        case "MyClassroom":
                            generateTemplateTable([["INSTRUCTOR", "STUDENT LAST NAME", "STUDENT FIRST NAME", "EMAIL", "PARENTEMAIL", "PARENTPHONE"]],
                                [["Smith", "Dunsmuir", "Alex", "alexd@patsoldemo6.com", "parent@home.com", "555-1212"]])
                            break;
                    }
                }
                else {
                    app.showNotification("Error", "Remove any existing student roster tables before adding a new one");
                }
            })
        }).catch(function (error) {
            app.showNotification("Error", "Something went wrong: " + error.message);
        });
    }

    /*******************************************************/
    /* Open a pop-up window with the steps to export a csv */
    /*******************************************************/
    function showHelp() {
        var helpWindow = window.open("HelpPop.html", "mywindow", "menubar=1,resizable=1,width=800,height=850");
    }

    /****************************************************/
    /* Populate worksheet with students for chosen tool */
    /****************************************************/
    function generateTemplateTable(headerString, defaultTableValues) {
        // Run a batch operation against the Excel object model
        Excel.run(function (ctx) {
            // Run the queued-up commands, and return a promise to indicate task completion

            //Create a new worksheet for the selected service
            var rosterWorksheet = ctx.workbook.worksheets.getActiveWorksheet();
            rosterWorksheet.name = selectedService;
            rosterName = selectedService;

            var tableRange = rosterWorksheet.getCell(0, 0).getBoundingRect(
                rosterWorksheet.getCell(0, headerString[0].length - 1))
  
            tableRange.load("address");

            //Run the batched commands
            return ctx.sync()
            .then(function () {
   
                //Build the table in the specified range
                var table = rosterWorksheet.tables.add(tableRange.address, true);
                table.name = rosterName;

                // Queue a command to get the newly added table
                table.getHeaderRowRange().values = headerString;
                table.style = "TableStyleLight20";
                table.rows.add(null, defaultTableValues)
            })
            .then(ctx.sync);
        }).catch(function (error) {
            // Always be sure to catch any accumulated errors that bubble up from the Excel.run execution
            app.showNotification("Error: " + error);
            console.log("Error: " + error);
            if (error instanceof OfficeExtension.Error) {
                console.log("Debug info: " + JSON.stringify(error.debugInfo));
            }
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