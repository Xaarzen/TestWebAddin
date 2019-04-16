(function () {
    "use strict";

    var cellToHighlight;
    var messageBanner;

    // The initialize function must be run each time a new page is loaded.
    Office.initialize = function (reason) {
        $(document).ready(function () {
            // Initialize the FabricUI notification mechanism and hide it
            var element = document.querySelector('.ms-MessageBanner');
            messageBanner = new fabric.MessageBanner(element);
            messageBanner.hideBanner();

            // If not using Excel 2016, use fallback logic.
            //if (!Office.context.requirements.isSetSupported('ExcelApi', '1.1')) {
            //    $("#template-description").text("This sample will display the value of the cells that you have selected in the spreadsheet.");
            //    $('#button-text').text("Display!");
            //    $('#button-desc').text("Display the selection");

            //    $('#highlight-button').click(displaySelectedCells);
            //    return;
            //}

            $("#template-description").text("This sample clear sheet, write values, highlight highest value and insert a new sheet.");
            $('#clearButton-text').text("Clear sheet.");
            $('#fillButton-text').text("Load sample data.");
            $('#highlightButton-text').text("Highlight largest value.");
            $('#newsheetButton-text').text("Add new sheet.");
            //$('#button-desc').text("Highlights the largest number.");

            //loadSampleData();

            // Add a click event handler for the highlight button.
            $('#fill-button').click(loadSampleData);
            $('#highlight-button').click(hightlightHighestValue);
            $('#clear-button').click(ClearSheet);
            $('#newsheet-button').click(AddNewSheet);

            InitializeEvents();
        });
    };

    function InitializeEvents() {
        Excel.run(function (ctx) {
            var worksheets = ctx.workbook.worksheets;
            worksheets.onActivated.add(handleSheetActivation);
            var activeSheet = ctx.workbook.worksheets.getActiveWorksheet();
            activeSheet.onSelectionChanged.add(handleSelectionChange);
            return ctx.sync()
                .then(function () {
                    console.log("Event handler successfully registered for onActivated event in the workbook.");
                    console.log("Event handler successfully registered for onSelectionChanged event in the active worksheet.");
                    UpdateLabelWithPosition();
                });
        })
            .catch(errorHandler);
    }

    function handleSheetActivation(event) {
        return Excel.run(function (context) {
            var worksheet = context.workbook.worksheets.getActiveWorksheet();
            worksheet.onSelectionChanged.add(handleSelectionChange);
            return context.sync()
                .then(function () {
                    UpdateLabelWithPosition();
                });
        }).catch(errorHandler);
    }

    function handleSelectionChange(event) {
        return Excel.run(function (ctx) {
            return ctx.sync()
                .then(function () {
                    UpdateLabelWithPosition();
                });
        }).catch(errorHandler);
    }

    function UpdateLabelWithPosition(event) {
        return Excel.run(function (ctx) {
            return ctx.sync()
                .then(function () {
                    var selection = ctx.workbook.getSelectedRange().load("address");
                    return ctx.sync()
                        .then(function () {
                            $('#position-label').text("Address of current selection: " + selection.address);
                        })
                });
        }).catch(errorHandler);
    }

    function loadSampleData() {
        //var values = [
        //    [Math.floor(Math.random() * 1000), Math.floor(Math.random() * 1000), Math.floor(Math.random() * 1000)],
        //    [Math.floor(Math.random() * 1000), Math.floor(Math.random() * 1000), Math.floor(Math.random() * 1000)],
        //    [Math.floor(Math.random() * 1000), Math.floor(Math.random() * 1000), Math.floor(Math.random() * 1000)]
        //];

        // Run a batch operation against the Excel object model
        Excel.run(function (ctx) {
            // Create a proxy object for the active sheet
            var selection = ctx.workbook.getSelectedRange().load("values, rowCount, columnCount, address, rowIndex, columnIndex");
            // Queue a command to write the sample data to the worksheet
            //sheet.getRange("B3:D5").values = values;
            return ctx.sync()
                .then(function () {
                    var totalRows = selection.rowCount;
                    var totalColumns = selection.columnCount;

                    var startingRow = selection.rowIndex;
                    var startingColumn = selection.columnIndex;

                    var arrayValues = new Array(totalRows);

                    for (var i = 0; i < arrayValues.length; i++) {
                        arrayValues[i] = new Array(totalColumns);
                    }

                    for (var rows = 0; rows < selection.rowCount; rows++) {
                        for (var columns = 0; columns < selection.columnCount; columns++) {
                            arrayValues[rows][columns] = (startingRow + rows + 1) * (startingColumn + columns + 1);
                        }
                    }
                    selection.values = arrayValues;
                })
                .then(ctx.sync);


            // Run the queued-up commands, and return a promise to indicate task completion
            return ctx.sync();
        })
            .catch(errorHandler);
    }

    function hightlightHighestValue() {
        // Run a batch operation against the Excel object model
        Excel.run(function (ctx) {
            // Create a proxy object for the selected range and load its properties
            var sourceRange = ctx.workbook.getSelectedRange().load("values, rowCount, columnCount");

            // Run the queued-up command, and return a promise to indicate task completion
            return ctx.sync()
                .then(function () {
                    var highestRow = 0;
                    var highestCol = 0;
                    var highestValue = sourceRange.values[0][0];

                    // Find the cell to highlight
                    for (var i = 0; i < sourceRange.rowCount; i++) {
                        for (var j = 0; j < sourceRange.columnCount; j++) {
                            if (!isNaN(sourceRange.values[i][j]) && sourceRange.values[i][j] > highestValue) {
                                highestRow = i;
                                highestCol = j;
                                highestValue = sourceRange.values[i][j];
                            }
                        }
                    }

                    cellToHighlight = sourceRange.getCell(highestRow, highestCol);
                    sourceRange.worksheet.getUsedRange().format.fill.clear();
                    sourceRange.worksheet.getUsedRange().format.font.bold = false;

                    // Highlight the cell
                    cellToHighlight.format.fill.color = "orange";
                    cellToHighlight.format.font.bold = true;
                })
                .then(ctx.sync);
        })
            .catch(errorHandler);
    }

    function ClearSheet() {
        Excel.run(function (ctx) {
            var sheet = ctx.workbook.worksheets.getActiveWorksheet().getUsedRange();
            sheet.load("name, values");

            // Run the queued-up command, and return a promise to indicate task completion
            return ctx.sync()
                .then(function () {
                    sheet.values = "";
                })
                .then(ctx.sync);
        })
            .catch(errorHandler);
    }
    function AddNewSheet(sheetNumber) {
        Excel.run(function (ctx) {
            var sheets = ctx.workbook.worksheets;
            sheets.load("items/name");

            var sheetName = "TestWorksheet";

            if (!isNaN(sheetNumber)) {
                sheetName = sheetName + sheetNumber;
            }

            var sheet = sheets.add(sheetName);
            sheet.load("name, position");

            // Run the queued-up command, and return a promise to indicate task completion
            return ctx.sync()
                .then(function () {
                    //sheet.values = "";
                })
                .then(ctx.sync);
        })
            .catch(errorHandler);
    }

    function displaySelectedCells() {
        Office.context.document.getSelectedDataAsync(Office.CoercionType.Text,
            function (result) {
                if (result.status === Office.AsyncResultStatus.Succeeded) {
                    showNotification('The selected text is:', '"' + result.value + '"');
                } else {
                    showNotification('Error', result.error.message);
                }
            });
    }

    // Helper function for treating errors
    function errorHandler(error) {
        // Always be sure to catch any accumulated errors that bubble up from the Excel.run execution
        showNotification("Error", error);
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
    }

    // Helper function for displaying notifications
    function showNotification(header, content) {
        $("#notification-header").text(header);
        $("#notification-body").text(content);
        messageBanner.showBanner();
        messageBanner.toggleExpansion();
    }
})();
