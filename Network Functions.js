function getSpreadsheetContext() {
  const spreadsheet = SpreadsheetApp.openById("1bBtgm6DoGjQP7IdzmSLtnTCor5FypMWRotrcclef9iY");
  return {
    spreadsheet,
    rawSubmissionsSheet: spreadsheet.getSheetByName("Raw Vendor Submissions"),
    rawSubmissionsValues: spreadsheet.getSheetByName("Raw Vendor Submissions").getDataRange().getValues(),
    wineCatalogueSheet: spreadsheet.getSheetByName("Wine Catalogue"),
    wineCatalogueValues: spreadsheet.getSheetByName("Wine Catalogue").getDataRange().getValues()
  };
}

// SERVE UP DATA
function doGet(e) {
  return serveData("Wine Catalogue")
}

// RECEIVE SUBMISSION AND WRITE TO SPREADSHEET
function doPost(e) {
  const body = JSON.parse(e.postData.contents);
  const spreadsheetContext = getSpreadsheetContext();
  let output, isSubmitted, isLogged;

  try {
    isSubmitted = logRaw(body, spreadsheetContext)
  } catch (error) {
    isSubmitted = { contents, error }
    logError(spreadsheetContext.rawSubmissionsSheet, [new Date.toISOString(), error, contents])
  }
  try {
    isLogged = updateCatalogue(body, spreadsheetContext)
  } catch (error) {
    isLogged = { contents, error }
    logError(spreadsheetContext.rawSubmissionsSheet, [new Date.toISOString(), error, contents])
  }
  // send a response back with the data that was supplied and a success indicator
  output = ContentService.createTextOutput(
    JSON.stringify({ isSubmitted, isLogged, eventObject: e })
  );
  output.setMimeType(ContentService.MimeType.JSON);
}

function logRaw(data, spreadsheetContext = getSpreadsheetContext()) {
  try {
    const sheet = spreadsheetContext.rawSubmissionsSheet || spreadsheetContext.spreadsheet.getSheetByName("Raw Vendor Submissions");
    sheet.appendRow([
      new Date(),
      JSON.stringify(data)]);
    return true
  } catch (e) {
    return false
  }
}


function updateCatalogue(data = {
  "action": "formSubmit",
  "formData": { "email": "agreider@junglejims.com" },
  "submit_time": "2026-09-09T14:25:02.181Z",
  "changeLog": {
    "1111": {
      "bottle": { "Created_Time": "2025-09-29T16:17:00.753Z", "Modified_Time": "2025-09-30T16:17:00.753Z", "Booth #": 1, "Booth_Name": "Tri-Vin Imports", "Wine_Name": "Goru Gold", "Price": 11.99, "Region": "Spain", "Is_VIP": "No", "Wine_ID": 1111, "Submitter Email Address": "tbockhold48@aol.com", "Distributor Name": "Euro Fine Wine", "Distributor Phone #": "Ego Bodegas", "Distributor Email": "tbockhold48@aol.com", "Winery Name": "Ego Bodegas", "Winery Phone #": 5136683086, "Winery Email": "tbockhold48@aol.com" },
      "type": "DELETE"
    }
  }
}, spreadsheetContext = getSpreadsheetContext()) {
  if (!data || !data.changeLog) {
    return false;
  }

  const sheet = spreadsheetContext.wineCatalogueSheet || spreadsheetContext.spreadsheet.getSheetByName("Wine Catalogue");

  const values = spreadsheetContext.wineCatalogueValues || sheet.getDataRange().getValues();
  const header = values.length ? values[0].slice() : [
    "Created_Time",
    "Modified_Time",
    "Booth #",
    "Booth_Name",
    "Wine_Name",
    "Price",
    "Region",
    "Is_VIP",
    "Wine_ID",
    "",
    "Submitter Email Address",
    "Distributor Name",
    "Distributor Phone #",
    "Distributor Email",
    "Winery Name",
    "Winery Phone #",
    "Winery Email",
    "Booth_ID",
  ];

  if (!header.some(function (key) {
    const columnName = String(key).trim();
    return columnName === "Booth_ID" || columnName === "Booth ID";
  })) {
    header.push("Booth_ID");
  }

  const catalogueMap = {};

  // Rebuild the current catalogue as a Wine_ID keyed map so we can merge updates cleanly.
  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    if (!row || !row.length) {
      continue;
    }

    const rowObject = {};
    header.forEach(function (key, idx) {
      rowObject[key] = row[idx] !== undefined ? row[idx] : "";
    });

    const wineId = rowObject.Wine_ID;
    if (wineId === undefined || wineId === null || wineId === "") {
      continue;
    }

    const wineKey = String(wineId);
    if (!catalogueMap[wineKey]) {
      catalogueMap[wineKey] = rowObject;
      continue;
    }

    const currentTime = new Date(catalogueMap[wineKey].Modified_Time || catalogueMap[wineKey].Created_Time || 0).getTime();
    //const incomingTime = new Date(rowObject.Modified_Time || rowObject.Created_Time || 0).getTime();
    const incomingTime = new Date(data.submit_time).getTime();

    if (incomingTime > currentTime) {
      catalogueMap[wineKey] = rowObject;
    }
  }

  const changeEntries = Object.values(data.changeLog || {});
  if (!changeEntries.length) {
    return false;
  }

  // Apply each incoming change to the map, honoring DELETE before newer CHANGE entries.
  changeEntries.forEach(function (entry) {
    const bottle = entry && entry.bottle ? entry.bottle : {};
    const type = (entry && entry.type) ? String(entry.type).toUpperCase() : "CHANGE";
    const wineId = bottle.Wine_ID;
    if (wineId === undefined || wineId === null || wineId === "") {
      return;
    }

    const wineKey = String(wineId);

    if (type === "DELETE") {
      console.log(`Deleting bottle with Wine_ID: ${wineKey}`);
      delete catalogueMap[wineKey];
      return;
    }

    const existingBottle = catalogueMap[wineKey];

    if (!existingBottle) {
      catalogueMap[wineKey] = {
        ...bottle,
        Modified_Time: new Date().toISOString()
      };
      return;
    }

    const existingTime = new Date(existingBottle.Modified_Time || existingBottle.Created_Time || 0).getTime();
    const incomingTime = new Date(data.submit_time || bottle.Modified_Time || bottle.Created_Time || 0).getTime();

    if (incomingTime >= existingTime) {
      catalogueMap[wineKey] = {
        ...existingBottle,
        ...bottle,
        Modified_Time: new Date(data.submit_time || Date.now()).toISOString()
      };
    }
  });

  // Flatten the merged map back into spreadsheet rows in the expected catalog schema.
  const finalRows = Object.values(catalogueMap).map(function (bottle) {
    const rowValues = {
      ...bottle,
      "Submitter Email Address": data.formData && data.formData.email
        ? data.formData.email
        : bottle["Submitter Email Address"] || "",
      Booth_ID: bottle.Booth_ID ?? "No Booth ID",
    };

    // Write by header name so Booth_ID cannot drift when columns are reordered.
    return header.map(function (key) {
      const columnName = String(key).trim() === "Booth ID" ? "Booth_ID" : String(key).trim();
      return rowValues[columnName] !== undefined && rowValues[columnName] !== null
        ? rowValues[columnName]
        : "";
    });
  });

  const allRows = [header].concat(finalRows);
  const firstRow = 1;
  const numRows = allRows.length;
  const numCols = header.length;

  sheet.clearContents();
  sheet.getRange(firstRow, 1, numRows, numCols).setValues(allRows);

  return true;
}