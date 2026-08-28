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
 *   - Serves a stable "scan to view" redirect link (?aadhar=XXXX&view=pdf)
 *     that always points at whatever the CURRENT vendor-report.pdf is for
 *     that Aadhar — this is what the QR code encodes, both the one saved
 *     into the Drive folder and the one embedded on the PDF itself.
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
 *
 * ---- IF THE QR CODE / _qr.png IS MISSING ----
 * QR generation calls an external site (api.qrserver.com) using
 * UrlFetchApp. The very first time this kind of call is added to a
 * project, Google needs a fresh permission grant ("Connect to an
 * external service") before it will work from the deployed Web App —
 * simply re-deploying is NOT enough if that consent was never given.
 * Fix: in the script editor, pick any function (e.g. doGet) in the
 * function dropdown and click "Run" once. Approve the permission
 * prompt that appears (it will now mention connecting to an external
 * service). Then go to Deploy > Manage deployments > edit (pencil) >
 * New version > Deploy, so the same URL picks up the authorized code.
 * This build also writes any QR failure reason into the saved record
 * (qrCodeError) and into the doPost response, so you can see exactly
 * what went wrong instead of it failing silently.
 */

var SHEET_ID = '1LBAxObNxRTF2RVu8WlVMFw2UDq3C91Gztb9rRyP-rOY';
var SHEET_NAME = 'catering_data';
var DRIVE_FOLDER_ID = '186VQZH294gSTpl3VaNajAoraGseBM0Et';

/** Strips everything except digits/letters, so "726655977066" and a Sheets-auto-converted
 *  number in the same cell always compare equal — this is what stops duplicate rows.
 *  Aadhar numbers are always 12 digits; if Sheets ever auto-converted a value to a Number
 *  and dropped a leading zero, this re-pads it so old and new values still match. */
function normalizeAadhar(v) {
  var s = String(v == null ? '' : v).replace(/[^0-9A-Za-z]/g, '');
  if (/^[0-9]+$/.test(s) && s.length > 0 && s.length < 12) {
    s = ('000000000000' + s).slice(-12);
  }
  return s;
}

/** Extracts the Drive file ID from a Drive sharing URL (as produced by
 *  File.getUrl()), e.g. ".../file/d/FILE_ID/view?..." or "...?id=FILE_ID". */
function extractDriveFileId(url) {
  if (!url) return null;
  var m = /\/d\/([a-zA-Z0-9_-]+)/.exec(url);
  if (m) return m[1];
  m = /[?&]id=([a-zA-Z0-9_-]+)/.exec(url);
  return m ? m[1] : null;
}

/** The stable "scan to view" link for a given Aadhar number. This is what
 *  every QR code (Drive-saved PNG and the one embedded in the PDF) encodes.
 *  It never changes even if the underlying PDF file is replaced/recreated,
 *  because doGet() below always looks up whatever the CURRENT record says. */
function stableViewUrl(aadhar) {
  return ScriptApp.getService().getUrl() + '?aadhar=' + encodeURIComponent(aadhar) + '&view=pdf';
}

/**
 * ===== ONE-TIME SETUP — RUN THIS MANUALLY ONCE =====
 * In the script editor, pick "runQrPermissionSetup" from the function
 * dropdown next to the ▶ Run button, then click Run. Google will show an
 * "Authorization required" popup — click through it (choose your account,
 * click "Advanced" > "Go to <project name> (unsafe)" if warned, since this
 * is your own script, then click "Allow"). This grants the
 * script.external_request permission that QR-code generation needs; simply
 * redeploying a new version does NOT grant it, only running a function
 * that calls an external URL does. After approving, check View > Logs (or
 * Executions) — it should say "Permission check HTTP status: 200". Then go
 * to Deploy > Manage deployments > edit (pencil) > New version > Deploy so
 * the live Web App picks up the now-authorized code. Safe to re-run any
 * time; it does not touch your Sheet or Drive data.
 */
function runQrPermissionSetup() {
  var res = UrlFetchApp.fetch('https://api.qrserver.com/v1/create-qr-code/?size=50x50&data=test', { muteHttpExceptions: true });
  Logger.log('Permission check HTTP status: ' + res.getResponseCode());
}

function doPost(e) {
  var lock = LockService.getScriptLock();
  try {
    var data = JSON.parse(e.postData.contents);
    var aadhar = normalizeAadhar(data.aadhar);
    if (!aadhar) {
      throw new Error('Aadhar Card No. is missing — cannot file this submission.');
    }

    // Only one submission at a time may search-and-write the sheet, so two near-simultaneous
    // requests for the same Aadhar can never both decide "not found" and both append a row.
    lock.waitLock(30000);

    var ss = SpreadsheetApp.openById(SHEET_ID);
    var sheet = ss.getSheetByName(SHEET_NAME);
    if (!sheet) {
      sheet = ss.insertSheet(SHEET_NAME);
      sheet.appendRow(['Aadhar Card No.', 'Data (JSON)']);
    }
    sheet.getRange('A:A').setNumberFormat('@'); // plain text, so numeric Aadhar values never get reformatted

    // Look for an existing row for this Aadhar number (server is the source of truth)
    var existingRowIndex = -1;
    var previousRecord = null;
    var lastRow = sheet.getLastRow();
    if (lastRow >= 2) {
      var aadharCol = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
      for (var r = 0; r < aadharCol.length; r++) {
        if (normalizeAadhar(aadharCol[r][0]) === aadhar) {
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

    // 4. Generate the QR code ONLY the very first time for this Aadhar, then
    //    NEVER touch it again — the same file (same Drive link) must stay
    //    valid for life, since this is what gets printed once on the vendor's
    //    physical ID card. The QR encodes a STABLE link (via stableViewUrl,
    //    based only on the Aadhar number) that always redirects to whatever
    //    the CURRENT pdfReport is — so even though the file itself never
    //    changes again, scanning it in future years will still open that
    //    year's latest uploaded document automatically.
    var qrLink = (previousRecord && previousRecord.qrCode) ? previousRecord.qrCode : '';
    var qrCodeError = (previousRecord && previousRecord.qrCodeError) ? previousRecord.qrCodeError : '';
    if (!qrLink) {
      try {
        var qrTargetUrl = stableViewUrl(aadhar);
        var qrApiUrl = 'https://api.qrserver.com/v1/create-qr-code/?size=300x300&margin=8&data=' + encodeURIComponent(qrTargetUrl);
        var qrResponse = UrlFetchApp.fetch(qrApiUrl, { muteHttpExceptions: true });
        var qrCode_ = qrResponse.getResponseCode();
        if (qrCode_ === 200) {
          var qrName = aadhar + '_qr.png';
          removeExistingFiles(userFolder, aadhar + '_qr');
          var qrBlob = qrResponse.getBlob().setName(qrName);
          var qrFile = userFolder.createFile(qrBlob);
          qrFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
          qrLink = qrFile.getUrl();
          qrCodeError = '';
        } else {
          qrCodeError = 'QR API returned HTTP ' + qrCode_;
        }
      } catch (qrErr) {
        // QR generation failing shouldn't block the rest of the submission,
        // but we keep the reason so it's visible instead of vanishing silently.
        qrCodeError = qrErr && qrErr.message ? qrErr.message : String(qrErr);
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
    record.qrViewUrl = stableViewUrl(aadhar);
    if (qrCodeError) { record.qrCodeError = qrCodeError; } else { delete record.qrCodeError; }
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
        qrViewUrl: record.qrViewUrl,
        qrCodeError: qrCodeError || undefined,
        documents: docLinks
      }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ success: false, error: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  } finally {
    try { lock.releaseLock(); } catch (releaseErr) { /* lock was never acquired — nothing to release */ }
  }
}

/**
 * GET /exec?list=1
 * Returns EVERY saved vendor record (one per Aadhar row) as a JSON array,
 * for the Vendor List page's table + on-screen report view. Document/photo
 * bytes are NOT included here (that would be far too heavy for a list) —
 * only the record fields, the Drive document/PDF links, and the QR link.
 */
function doGet_listAll() {
  try {
    var ss = SpreadsheetApp.openById(SHEET_ID);
    var sheet = ss.getSheetByName(SHEET_NAME);
    var vendors = [];
    if (sheet) {
      var lastRow = sheet.getLastRow();
      if (lastRow >= 2) {
        var values = sheet.getRange(2, 1, lastRow - 1, 2).getValues();
        for (var i = 0; i < values.length; i++) {
          var rec = null;
          try { rec = JSON.parse(values[i][1]); } catch (parseErr) { rec = null; }
          if (rec) vendors.push(rec);
        }
      }
    }
    return ContentService
      .createTextOutput(JSON.stringify({ success: true, vendors: vendors }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ success: false, error: err.message, vendors: [] }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * GET /exec?aadhar=XXXXXXXXXXXX&fetchDocs=1
 * Returns the raw base64 content + mime type of every document already
 * saved for this Aadhar (photo, signature, aadhar front/back, certificates).
 * Used by the form when EDITING an existing record: if the user doesn't
 * re-upload a document, the PDF still needs its actual bytes (not just the
 * Drive link) to draw it into the report — this is what supplies that.
 */
function doGet_fetchDocs(aadharParam) {
  var aadharF = normalizeAadhar(aadharParam);
  var out = {};
  try {
    var ss = SpreadsheetApp.openById(SHEET_ID);
    var sheet = ss.getSheetByName(SHEET_NAME);
    var rec = null;
    if (sheet) {
      var lastRow = sheet.getLastRow();
      if (lastRow >= 2) {
        var values = sheet.getRange(2, 1, lastRow - 1, 2).getValues();
        for (var i = 0; i < values.length; i++) {
          if (normalizeAadhar(values[i][0]) === aadharF) {
            try { rec = JSON.parse(values[i][1]); } catch (e3) { rec = null; }
            break;
          }
        }
      }
    }
    if (rec && rec.documents) {
      for (var field in rec.documents) {
        var fid = extractDriveFileId(rec.documents[field]);
        if (!fid) continue;
        try {
          var blob = DriveApp.getFileById(fid).getBlob();
          out[field] = {
            base64: Utilities.base64Encode(blob.getBytes()),
            mimeType: blob.getContentType() || 'application/octet-stream'
          };
        } catch (fileErr) {
          // Skip this one silently — the PDF will just show "NOT UPLOADED"
          // for that particular document instead of blocking everything else.
        }
      }
    }
    return ContentService
      .createTextOutput(JSON.stringify({ success: true, files: out }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ success: false, error: err.message, files: {} }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * GET /exec?aadhar=XXXXXXXXXXXX
 * Looks up whether a record already exists for this Aadhar number.
 * Used by the form to prompt "already exists — load & edit?" before submit.
 *
 * GET /exec?aadhar=XXXXXXXXXXXX&view=pdf
 * The stable "scan to view" link that every QR code encodes. Looks up the
 * CURRENT pdfReport link for this Aadhar and redirects the browser there,
 * so the link keeps working even if the underlying Drive file is replaced.
 */
function doGet(e) {
  var aadharParam = e && e.parameter ? e.parameter.aadhar : null;
  var debugParam = e && e.parameter ? e.parameter.debug : null;
  var viewParam = e && e.parameter ? e.parameter.view : null;
  var fetchDocsParam = e && e.parameter ? e.parameter.fetchDocs : null;
  var listParam = e && e.parameter ? e.parameter.list : null;

  if (listParam) {
    return doGet_listAll();
  }

  if (fetchDocsParam && aadharParam) {
    return doGet_fetchDocs(aadharParam);
  }

  if (viewParam === 'pdf' && aadharParam) {
    var aadharV = normalizeAadhar(aadharParam);
    var targetUrl = '';
    var notFoundMsg = '';
    try {
      var ss = SpreadsheetApp.openById(SHEET_ID);
      var sheet = ss.getSheetByName(SHEET_NAME);
      if (sheet) {
        var lastRow = sheet.getLastRow();
        if (lastRow >= 2) {
          var values = sheet.getRange(2, 1, lastRow - 1, 2).getValues();
          for (var i = 0; i < values.length; i++) {
            if (normalizeAadhar(values[i][0]) === aadharV) {
              var rec = null;
              try { rec = JSON.parse(values[i][1]); } catch (e2) { rec = null; }
              if (rec && rec.pdfReport) { targetUrl = rec.pdfReport; }
              break;
            }
          }
        }
      }
    } catch (lookupErr) {
      notFoundMsg = lookupErr.message;
    }

    if (targetUrl) {
      var safeUrl = JSON.stringify(targetUrl);
      var html = '<!DOCTYPE html><html><head><meta charset="utf-8">' +
        '<meta http-equiv="refresh" content="0; url=' + targetUrl + '">' +
        '<title>Opening report…</title></head><body style="font-family:sans-serif;text-align:center;padding:40px;">' +
        '<p>Opening the vendor report… if it does not open automatically, ' +
        '<a href="' + targetUrl + '">tap here</a>.</p>' +
        '<script>window.top.location.href = ' + safeUrl + ';</script>' +
        '</body></html>';
      return HtmlService.createHtmlOutput(html);
    }
    return HtmlService.createHtmlOutput(
      '<!DOCTYPE html><html><body style="font-family:sans-serif;text-align:center;padding:40px;">' +
      '<p>No report found yet for this Aadhar number' + (notFoundMsg ? (' (' + notFoundMsg + ')') : '') + '.</p>' +
      '</body></html>'
    );
  }

  if (debugParam) {
    try {
      var ss = SpreadsheetApp.openById(SHEET_ID);
      var sheet = ss.getSheetByName(SHEET_NAME);
      var rows = [];
      if (sheet) {
        var lastRow = sheet.getLastRow();
        if (lastRow >= 2) {
          var values = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
          for (var i = 0; i < values.length; i++) {
            rows.push({ row: i + 2, raw: values[i][0], rawType: typeof values[i][0], normalized: normalizeAadhar(values[i][0]) });
          }
        }
      }
      return ContentService
        .createTextOutput(JSON.stringify({ sheetFound: !!sheet, totalDataRows: rows.length, rows: rows }))
        .setMimeType(ContentService.MimeType.JSON);
    } catch (err) {
      return ContentService
        .createTextOutput(JSON.stringify({ error: err.message }))
        .setMimeType(ContentService.MimeType.JSON);
    }
  }

  if (aadharParam) {
    var aadhar = normalizeAadhar(aadharParam);
    try {
      var ss = SpreadsheetApp.openById(SHEET_ID);
      var sheet = ss.getSheetByName(SHEET_NAME);
      if (sheet) {
        var lastRow = sheet.getLastRow();
        if (lastRow >= 2) {
          var values = sheet.getRange(2, 1, lastRow - 1, 2).getValues();
          for (var i = 0; i < values.length; i++) {
            if (normalizeAadhar(values[i][0]) === aadhar) {
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
    .createTextOutput(JSON.stringify({ status: 'Catering Vendor Backend is running.', version: 'v2-qr-diagnostics' }))
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
