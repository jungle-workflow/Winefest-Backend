
function serveData(sheetName) {
    var output = toJSON(true, sheetName);
    return ContentService.createTextOutput(JSON.stringify({ data: output })).setMimeType(ContentService.MimeType.JSON);
}

function toJSON(log = true, sheetName = "Contest Registration") {
    //converts the specified sheet into separate arrays that are pushed to a json file
    let doc = SpreadsheetApp.getActiveSpreadsheet()

    let sheet = doc.getSheetByName(sheetName)
    let values = sheet.getDataRange().getValues();
    console.log(values)
    let output = [];

    for (let i = 1; i < values.length; i++) {
        let row = {}; // Create a new row object for each row 
        for (let j = 0; j < values[i].length; j++) {
            let key = undefined
            if (values[0][j].includes("General Information")) {
                key = "Name"
            } else if (values[0][j].includes("Please list the full names and types of product you will be sampling.")) {
                key = "Products"
            } else {
                key = values[0][j]
            }
            //console.log(`values[${i}][${j}]: ${values[i][j]}`)
            if (key) {
                //row[values[0][j]] = values[i][j]; // move this out of the if statement to capture all of the columns
                row[key] = values[i][j];
            }

        }
        output.push(row); // Push the row object to the output array 
        delete row;
    }

    log ? console.log(output) : undefined
    return output
}


/* function onOpen() {
    const ui = SpreadsheetApp.getUi();
    ui.createMenu('Contact Us Control Center')
        .addItem('Add New Member', 'showCheckboxes')
        .addItem('Remove Member', 'showRemoveMenu')
        .addItem('Refresh Permissions', 'permissionsOnEdit')
        .addToUi();
} */


function logError(rawSubsheet, msg = ["error"]) {
    rawSubsheet.appendRow(...msg)
}