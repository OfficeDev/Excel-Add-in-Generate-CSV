/* Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. 
    See full license at the bottom of this file. */
(function () {
    "use strict";

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
            $('#generate-template').click(generateTemplateClickHandler);
            $('#show-help').click(showHelp);
            $(".ms-Dropdown").Dropdown();
        });
    };

    function generateTemplateClickHandler() {

        /***********************************************/
        /*Check for existing tables and then either 
        /*generate a new table or warn the user that
        /*there is an existing table that may have data
        /***********************************************/
        Excel.run(function (ctx) {
            var activeSheet = ctx.workbook.worksheets.getActiveWorksheet();
            ctx.load(activeSheet.tables, "name");
            return ctx.sync().then(function () {
                if (activeSheet.tables.count === 0) {
                    switch ($("#select-service").val()) {
                        case "Moodle":
                            return generateTemplateTable(ctx, activeSheet, [["ACTION", "ROLE", "USER ID NUMBER", "COURSE ID NUMBER"]],
                                [["add", "student", "usr-1", "econ 101"]]);
                        case "TeacherKit":
                            return generateTemplateTable(ctx, activeSheet, [["FIRST NAME", "LAST NAME", "EMAIL", "PARENTEMAIL", "PARENTPHONE"]],
                                [["Alex", "Dunsmuir", "alexd@patsoldemo6.com", "parent@home.com", "555-1212"]]);
                        case "MyClassroom":
                            return generateTemplateTable(ctx, activeSheet, [["INSTRUCTOR", "STUDENT LAST NAME", "STUDENT FIRST NAME", "EMAIL", "PARENTEMAIL", "PARENTPHONE"]],
                                [["Smith", "Dunsmuir", "Alex", "alexd@patsoldemo6.com", "parent@home.com", "555-1212"]]);
                    }
                }
                else {
                    app.showNotification("Error", "Remove any existing student roster tables before adding a new one");
                }
            });
        }).catch(function (error) {
            app.showNotification("Error", "Something went wrong: " + error);
            console.log("Error: " + error);
            if (error instanceof OfficeExtension.Error) {
                console.log("Debug info: " + JSON.stringify(error.debugInfo));
            }
        });
    }

    /*******************************************************/
    /* Open a pop-up window with the steps to export a csv */
    /*******************************************************/
    function showHelp() {
        window.open("HelpPop.html", "mywindow", "menubar=1,resizable=1,width=800,height=850");
    }

    /****************************************************/
    /* Populate worksheet with students for chosen tool */
    /****************************************************/
    function generateTemplateTable(ctx, activeSheet, headerString, defaultTableRow) {

        activeSheet.name = $("#select-service").val();

        var tableRange = activeSheet.getCell(0, 0).getBoundingRect(
            activeSheet.getCell(0, headerString[0].length - 1));

        tableRange.load("address");

        //Run the batched commands
        return ctx.sync()
        .then(function () {

            //Build the table in the specified range
            var table = activeSheet.tables.add(tableRange.address, true);
            table.name = $("#select-service").val();

            // Queue a command to get the newly added table
            table.getHeaderRowRange().values = headerString;
            table.style = "TableStyleLight20";
            table.rows.add(null, defaultTableRow);
        })
        .then(ctx.sync);

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