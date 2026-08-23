/**
 * CATERING UNIT MANAGEMENT — Vendor Details backend
 * ---------------------------------------------------
 * Deploy this as a Google Apps Script Web App. It:
 *   - Receives vendor form data + uploaded documents + generated PDF from
 *     vendor-details.html
 *   - Checks if the Aadhar Card No. already has a record (GET request)
 *   - Saves/replaces documents + PDF + a QR code (linking to the PDF) into a
 *     per-vendor subfolder (named by Aadhar number) in Google Drive
 *   - Appends a NEW row for a new Aadhar number, or UPDATES the existing row
 *     in place if that Aadhar number was already submitted before
 *     (Column A = Aadhar Card No., Column B = full record as JSON)
 *
 * ---- SETUP ----
 * 1. Go to https://script.google.com and create a New Project.
 * 2. Delete any starter code and paste this whole file in.
 * 3. Save the project (e.g. name it "Catering Vendor Backend").
 * 4. Click Deploy > New deployment.
 *      - Select type: "Web app"
 *      - Description: anything
 *      - Execute as: "Me"
 *      - Who has access: "Anyone"
 *    Click Deploy, and authorize the requested Drive/Sheets permissions.
 * 5. Copy the "Web app URL" you get after deploying.
 * 6. Open vendor-details.html and paste that URL into the
 *    APPS_SCRIPT_URL constant near the top of the <script> section.
 * 7. Whenever you edit this script, use Deploy > Manage deployments >
 *    Edit (pencil icon) > New version, so the same URL picks up changes.
 */

var SHEET_ID = '1LBAxObNxRTF2RVu8WlVMFw2UDq3C91Gztb9rRyP-rOY';
var SHEET_NAME = 'catering_data';
var DRIVE_FOLDER_ID = '186VQZH294gSTpl3VaNajAoraGseBM0Et';

function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);
    var aadhar = String(data.aadhar || '').replace(/[^0-9A-Za-z]/g, '');
    if (!aadhar) {
      throw new Error('Aadhar Card No. is missing — cannot file this submission.');
    }

    var ss = SpreadsheetApp.openById(SHEET_ID);
    var sheet = ss.getSheetByName(SHEET_NAME);
    if (!sheet) {
      sheet = ss.insertSheet(SHEET_NAME);
      sheet.appendRow(['Aadhar Card No.', 'Data (JSON)']);
    }

    // Look for an existing row for this Aadhar number (server is the source of truth)
    var existingRowIndex = -1;
    var previousRecord = null;
    var lastRow = sheet.getLastRow();
    if (lastRow >= 2) {
      var aadharCol = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
      for (var r = 0; r < aadharCol.length; r++) {
        if (String(aadharCol[r][0]) === aadhar) {
          existingRowIndex = r + 2; // +1 for header row, +1 for 1-based index
          break;
        }
      }
    }
    if (existingRowIndex > -1) {
      var prevJson = sheet.getRange(existingRowIndex, 2).getValue();
      try { previousRecord = JSON.parse(prevJson); } catch (parseErr) { previousRecord = null; }
    }

    // 1. Get (or create) this vendor's own subfolder, named by Aadhar number
    var parentFolder = DriveApp.getFolderById(DRIVE_FOLDER_ID);
    var userFolder;
    var existingFolders = parentFolder.getFoldersByName(aadhar);
    if (existingFolders.hasNext()) {
      userFolder = existingFolders.next();
    } else {
      userFolder = parentFolder.createFolder(aadhar);
    }

    // 2. Save/replace each newly-uploaded document; keep previous links for
    //    any document field that wasn't re-uploaded this time.
    var docLinks = (previousRecord && previousRecord.documents) ? previousRecord.documents : {};
    var files = data.files || [];
    for (var i = 0; i < files.length; i++) {
      var f = files[i];
      if (!f || !f.base64 || !f.field) continue;
      var bytes = Utilities.base64Decode(f.base64);
      var ext = extensionFor(f.filename, f.mimeType);
      var fileName = aadhar + '_' + f.field + ext;
      removeExistingFiles(userFolder, aadhar + '_' + f.field);
      var blob = Utilities.newBlob(bytes, f.mimeType || 'application/octet-stream', fileName);
      var savedFile = userFolder.createFile(blob);
      savedFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      docLinks[f.field] = savedFile.getUrl();
    }

    // 3. Save/replace the generated PDF report into the same folder
    var pdfLink = (previousRecord && previousRecord.pdfReport) ? previousRecord.pdfReport : '';
    if (data.pdf && data.pdf.base64) {
      var pdfBytes = Utilities.base64Decode(data.pdf.base64);
      var pdfName = aadhar + '_vendor-report.pdf';
      removeExistingFiles(userFolder, aadhar + '_vendor-report');
      var pdfBlob = Utilities.newBlob(pdfBytes, 'application/pdf', pdfName);
      var pdfFile = userFolder.createFile(pdfBlob);
      pdfFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      pdfLink = pdfFile.getUrl();
    }

    // 4. Generate a QR code pointing to the PDF report link, and save it too
    var qrLink = (previousRecord && previousRecord.qrCode) ? previousRecord.qrCode : '';
    if (pdfLink) {
      try {
        var qrApiUrl = 'https://api.qrserver.com/v1/create-qr-code/?size=300x300&data=' + encodeURIComponent(pdfLink);
        var qrResponse = UrlFetchApp.fetch(qrApiUrl, { muteHttpExceptions: true });
        if (qrResponse.getResponseCode() === 200) {
          var qrName = aadhar + '_qr.png';
          removeExistingFiles(userFolder, aadhar + '_qr');
          var qrBlob = qrResponse.getBlob().setName(qrName);
          var qrFile = userFolder.createFile(qrBlob);
          qrFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
          qrLink = qrFile.getUrl();
        }
      } catch (qrErr) {
        // QR generation failing shouldn't block the rest of the submission
      }
    }

    // 5. Build the full JSON record (form fields + document links + pdf + qr)
    var record = {};
    var formData = data.formData || {};
    for (var key in formData) { record[key] = formData[key]; }
    record.aadhar = aadhar;
    record.idNo = (previousRecord && previousRecord.idNo) ? previousRecord.idNo : (formData.idNo || '');
    record.documents = docLinks;
    record.pdfReport = pdfLink;
    record.qrCode = qrLink;
    record.submittedAt = new Date().toISOString();
    if (previousRecord) {
      record.firstSubmittedAt = previousRecord.firstSubmittedAt || previousRecord.submittedAt || record.submittedAt;
      record.updatedAt = new Date().toISOString();
    }

    // 6. Write to the Sheet — update the existing row in place, or append a new one
    if (existingRowIndex > -1) {
      sheet.getRange(existingRowIndex, 1, 1, 2).setValues([[aadhar, JSON.stringify(record)]]);
    } else {
      sheet.appendRow([aadhar, JSON.stringify(record)]);
    }

    return ContentService
      .createTextOutput(JSON.stringify({
        success: true,
        updated: existingRowIndex > -1,
        aadhar: aadhar,
        idNo: record.idNo,
        pdfLink: pdfLink,
        qrLink: qrLink,
        documents: docLinks
      }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ success: false, error: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * GET /exec?aadhar=XXXXXXXXXXXX
 * Looks up whether a record already exists for this Aadhar number.
 * Used by the form to prompt "already exists — load & edit?" before submit.
 */
function doGet(e) {
  var aadharParam = e && e.parameter ? e.parameter.aadhar : null;
  if (aadharParam) {
    var aadhar = String(aadharParam).replace(/[^0-9A-Za-z]/g, '');
    try {
      var ss = SpreadsheetApp.openById(SHEET_ID);
      var sheet = ss.getSheetByName(SHEET_NAME);
      if (sheet) {
        var lastRow = sheet.getLastRow();
        if (lastRow >= 2) {
          var values = sheet.getRange(2, 1, lastRow - 1, 2).getValues();
          for (var i = 0; i < values.length; i++) {
            if (String(values[i][0]) === aadhar) {
              var record = null;
              try { record = JSON.parse(values[i][1]); } catch (e2) { record = null; }
              return ContentService
                .createTextOutput(JSON.stringify({ exists: true, record: record }))
                .setMimeType(ContentService.MimeType.JSON);
            }
          }
        }
      }
      return ContentService
        .createTextOutput(JSON.stringify({ exists: false }))
        .setMimeType(ContentService.MimeType.JSON);
    } catch (err) {
      return ContentService
        .createTextOutput(JSON.stringify({ exists: false, error: err.message }))
        .setMimeType(ContentService.MimeType.JSON);
    }
  }
  return ContentService
    .createTextOutput(JSON.stringify({ status: 'Catering Vendor Backend is running.' }))
    .setMimeType(ContentService.MimeType.JSON);
}

/** Returns a lowercase file extension (with leading dot) from a filename, falling back to mimeType. */
function extensionFor(filename, mimeType) {
  if (filename) {
    var m = /\.([a-zA-Z0-9]+)$/.exec(filename);
    if (m) return '.' + m[1].toLowerCase();
  }
  var map = {
    'image/jpeg': '.jpg',
    'image/jpg': '.jpg',
    'image/png': '.png',
    'application/pdf': '.pdf'
  };
  return map[mimeType] || '';
}

/** Trashes any existing file in the folder whose name starts with namePrefix, so re-uploads replace instead of pile up. */
function removeExistingFiles(folder, namePrefix) {
  var it = folder.getFiles();
  while (it.hasNext()) {
    var f = it.next();
    if (f.getName().indexOf(namePrefix) === 0) {
      f.setTrashed(true);
    }
  }
}
