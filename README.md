const SHEET_NAME = "T-Shirt Submissions";

function doGet() {
  return jsonResponse({
    success: true,
    message: "T-Shirt submission app is running."
  });
}

function doPost(e) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = getOrCreateSheet(ss);

    const data = JSON.parse(e.postData.contents);

    const email = normalize(data.email);

    if (!email) {
      return jsonResponse({
        success: false,
        message: "Email is required."
      });
    }

    const now = new Date();
    const nowText = Utilities.formatDate(
      now,
      Session.getScriptTimeZone(),
      "MM/dd/yyyy hh:mm:ss a"
    );

    const rows = sheet.getDataRange().getValues();
    const headers = rows[0];

    const emailCol = headers.indexOf("Email");
    const rowIndex = findRowByEmail(rows, emailCol, email);

    const newOrder = {
      firstName: clean(data.firstName),
      lastName: clean(data.lastName),
      email: email,
      location: clean(data.location),
      size: clean(data.size),
      gender: clean(data.gender),
      originalSubmittedAt: nowText,
      lastUpdatedAt: nowText
    };

    if (rowIndex === -1) {
      const newRow = [
        newOrder.firstName,
        newOrder.lastName,
        newOrder.email,
        newOrder.location,
        newOrder.size,
        newOrder.gender,
        nowText,
        nowText,
        "New"
      ];

      sheet.appendRow(newRow);

      return jsonResponse({
        success: true,
        updated: false,
        previousOrder: null,
        currentOrder: newOrder,
        message: "New submission saved."
      });
    }

    const existingRow = rows[rowIndex];

    const previousOrder = {
      firstName: existingRow[0],
      lastName: existingRow[1],
      email: existingRow[2],
      location: existingRow[3],
      size: existingRow[4],
      gender: existingRow[5],
      originalSubmittedAt: existingRow[6],
      lastUpdatedAt: existingRow[7]
    };

    newOrder.originalSubmittedAt = existingRow[6] || nowText;
    newOrder.lastUpdatedAt = nowText;

    const sheetRowNumber = rowIndex + 1;

    sheet.getRange(sheetRowNumber, 1, 1, 9).setValues([[
      newOrder.firstName,
      newOrder.lastName,
      newOrder.email,
      newOrder.location,
      newOrder.size,
      newOrder.gender,
      newOrder.originalSubmittedAt,
      newOrder.lastUpdatedAt,
      "Updated"
    ]]);

    return jsonResponse({
      success: true,
      updated: true,
      previousOrder: previousOrder,
      currentOrder: newOrder,
      message: "Existing submission updated."
    });

  } catch (error) {
    return jsonResponse({
      success: false,
      message: error.toString()
    });
  }
}

function getOrCreateSheet(ss) {
  let sheet = ss.getSheetByName(SHEET_NAME);

  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
  }

  const headers = [
    "First Name",
    "Last Name",
    "Email",
    "Location",
    "Size",
    "Gender Fit",
    "Original Submitted At",
    "Last Updated At",
    "Status"
  ];

  const firstRow = sheet.getRange(1, 1, 1, headers.length).getValues()[0];
  const hasHeaders = firstRow.some(value => value !== "");

  if (!hasHeaders) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.setFrozenRows(1);
    sheet.autoResizeColumns(1, headers.length);
  }

  return sheet;
}

function findRowByEmail(rows, emailCol, email) {
  for (let i = 1; i < rows.length; i++) {
    const rowEmail = normalize(rows[i][emailCol]);
    if (rowEmail === email) {
      return i;
    }
  }

  return -1;
}

function clean(value) {
  return String(value || "").trim();
}

function normalize(value) {
  return String(value || "").trim().toLowerCase();
}

function jsonResponse(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}
