
function serveData(sheetName) {
  var output = toJSON(true, sheetName);
  return ContentService.createTextOutput(JSON.stringify({ data: output })).setMimeType(ContentService.MimeType.JSON);
}

function toJSON(log = true, sheetName = "Wine Catalogue") {
  //converts the specified sheet into separate arrays that are pushed to a json file
  let doc = SpreadsheetApp.getActiveSpreadsheet()

  let sheet = doc.getSheetByName(sheetName)
  let values = sheet.getDataRange().getValues();
  log && console.log(values)
  let output = [];

  for (let i = 1; i < values.length; i++) {
    let row = {}; // Create a new row object for each row 
    for (let j = 0; j < values[i].length; j++) {
      let key = undefined
      key = values[0][j]
      //console.log(`values[${i}][${j}]: ${values[i][j]}`)
      if (key) {
        //row[values[0][j]] = values[i][j]; // move this out of the if statement to capture all of the columns
        row[key] = values[i][j];
      }
    }
    output.push(row); // Push the row object to the output array 
    /* delete row; */
  }

  log ? console.log(output) : undefined
  return output
}


function logError(rawSubsheet, msg = ["error"]) {
  rawSubsheet.appendRow(...msg)
}


function populateContactInfo() {
  /* 
      Add a method that can be run independently called populateContactInfo() that reads the "Data" column in each row of the sheet, "Raw Vendor Submissions", containing data in this schema
      {"action":"formSubmit","formData":{"email":"aaron.greider@gmail.com","phone":"9377167737"},"submit_time":"2026-09-22T12:21:50.615Z","changeLog":{"8f75ef71-69d6-488e-9617-2e2c68761dba":{"bottle":{"Created_Time":"2026-09-22T12:16:31.283Z","Modified_Time":"2026-09-22T12:17:02.303Z","Booth #":1788976331805,"Booth_Name":"Jungle Jim's International Market","Wine_Name":"Test!!","Price":4,"Region":"Japan","Is_VIP":"No","Wine_ID":"8f75ef71-69d6-488e-9617-2e2c68761dba","Submitter_Email":"agreider@junglejims.com","Distributor Name":"Test Distributor","Distributor Phone #":"","Distributor Email":"","Winery Name":"Test Winery","Winery Phone #":"","Winery Email":"","Booth_ID":1788976331805,"Submitter_Phone":9377167737},"type":"ADD"}}}
      it will then overwrite the data in each corresponding row in the sheet "Wine Catalogue". Match the rows by comparing Wine_ID field. If Wine_ID matches, then take the Email property of formData and copy the value to Submitter_Email in Wine Catalogue. Same with the phone property to the Submitter_Phone field.
  */

  // Reuse the shared spreadsheet context helper from Network Functions.
  const spreadsheetContext = getSpreadsheetContext();
  const rawSheet = spreadsheetContext.rawVendorsSheet;
  const catalogSheet = spreadsheetContext.wineCatalogueSheet;

  // Pull the current sheet contents so we can look up the relevant columns.
  const rawValues = rawSheet.getDataRange().getValues();
  const catalogValues = catalogSheet.getDataRange().getValues();

  // Find the columns we need in each sheet.
  const rawHeaders = rawValues[0] || [];
  const dataCol = rawHeaders.indexOf("Data");

  const catalogHeaders = catalogValues[0] || [];
  const wineIdCol = catalogHeaders.indexOf("Wine_ID");
  const emailCol = catalogHeaders.indexOf("Submitter_Email");
  const phoneCol = catalogHeaders.indexOf("Submitter_Phone");

  if (dataCol === -1 || wineIdCol === -1) {
    console.log("populateContactInfo: missing required columns", { dataCol, wineIdCol, emailCol, phoneCol });
    return;
  }

  // Loop through the raw submission rows and parse the JSON stored in the Data column.
  for (let i = 1; i < rawValues.length; i++) {
    const row = rawValues[i];
    if (!row || row.length <= dataCol || !row[dataCol]) {
      continue;
    }

    let payload;
    try {
      payload = JSON.parse(row[dataCol]);
      console.log("populateContactInfo: parsed submission row", i, payload && payload.formData ? payload.formData.email : "no email");
    } catch (e) {
      console.log("populateContactInfo: invalid JSON in row", i, e.message);
      continue;
    }

    // Pull the contact info from the form submission.
    const email = payload && payload.formData ? payload.formData.email : "";
    const phone = payload && payload.formData ? payload.formData.phone : "";
    const changes = payload && payload.changeLog ? Object.values(payload.changeLog) : [];
    console.log("populateContactInfo: processing changes", changes.length, { email, phone });

    // Check each wine change and match it to a row in the catalogue by Wine_ID.
    for (let j = 0; j < changes.length; j++) {
      const bottle = changes[j] && changes[j].bottle ? changes[j].bottle : null;
      if (!bottle || !bottle.Wine_ID) {
        continue;
      }

      for (let k = 1; k < catalogValues.length; k++) {
        const catalogRow = catalogValues[k];
        if (!catalogRow || !catalogRow[wineIdCol]) {
          continue;
        }

        // Update the matching wine row with the submitter email and phone.
        if (String(catalogRow[wineIdCol]) === String(bottle.Wine_ID)) {
          console.log("populateContactInfo: matched Wine_ID", bottle.Wine_ID, "at row", k + 1);
          if (emailCol !== -1) {
            catalogSheet.getRange(k + 1, emailCol + 1).setValue(email || "");
          }
          if (phoneCol !== -1) {
            catalogSheet.getRange(k + 1, phoneCol + 1).setValue(phone || "");
          }
        }
      }
    }
  }
}

const routes = {
  registerWine: "formSubmit",
  getBoothByNum: "getBoothByNum",
  orderWine: "buyerFormSubmit"
}

function getWinesByNum(boothNum = "10") {
  const catalogue = toJSON(false, "Wine Catalogue");
  return catalogue.filter((wine) => String(wine["Booth #"]) === String(boothNum));
}



function onOpen() {
    const ui = SpreadsheetApp.getUi();
    ui.createMenu('Wine Catalogue Functions')
        .addItem('Publish Catalogue to Github', 'publishWines')
        .addToUi();
}

function publishWines() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) throw new Error('Another publish is running');

  try {
    const json = toJSON(false);
    // Make sure we have a string, then Base64 encode as UTF-8
    const text = typeof json === 'string' ? json : JSON.stringify(json);
    const content = Utilities.base64Encode(text, Utilities.Charset.UTF_8);

    const token = PropertiesService.getScriptProperties().getProperty('GH_TOKEN');
    const url = 'https://api.github.com/repos/jungle-workflow/Winefest-Backend/contents/catalogue.json';
    const headers = {
      Authorization: 'Bearer ' + token,
      Accept: 'application/vnd.github+json',
      'X-GitHub-Api-Version': '2022-11-28',
    };

    // Get the existing file's SHA (required to overwrite)
    const existing = UrlFetchApp.fetch(url, { headers, muteHttpExceptions: true });
    const code = existing.getResponseCode();
    let sha;
    if (code === 200) {
      sha = JSON.parse(existing.getContentText()).sha;
    } else if (code !== 404) {
      throw new Error('GitHub GET failed: ' + code + ' ' + existing.getContentText());
    }

    const put = UrlFetchApp.fetch(url, {
      method: 'put',
      headers,
      contentType: 'application/json',
      payload: JSON.stringify({ message: 'Update catalogue', content, sha }),
      muteHttpExceptions: true,
    });
    const putCode = put.getResponseCode();
    if (putCode !== 200 && putCode !== 201) {
      throw new Error('GitHub PUT failed: ' + putCode + ' ' + put.getContentText());
    }

    // Clear the CDN cache (check the response while testing)
    const purge = UrlFetchApp.fetch(
      'https://purge.jsdelivr.net/gh/jungle-workflow/Winefest-Backend@main/catalogue.json',
      { muteHttpExceptions: true }
    );
    Logger.log('Purge: ' + purge.getResponseCode() + ' ' + purge.getContentText());
  } finally {
    lock.releaseLock();
  }
}