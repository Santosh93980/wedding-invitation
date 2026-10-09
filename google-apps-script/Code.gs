/**
 * Digital Innovation Tech - Wedding Invitation data backend
 *
 * Creates one Google Spreadsheet with two tabs:
 *   1) Contacts - enquiry/contact submissions
 *   2) Visitors - anonymous visit/technical analytics
 *
 * It can export the workbook as .xlsx and commit/update it in GitHub.
 *
 * Required Script Properties:
 *   GITHUB_TOKEN      = GitHub fine-grained token with Contents: Read and write
 *   GITHUB_REPO      = Santosh93980/wedding-invitation
 *
 * Optional:
 *   SPREADSHEET_ID   = existing Google Sheet ID
 */
const CONTACT_EMAIL = "palakollu.santosh16@gmail.com";
const XLSX_PATH = "data/wedding-website-data.xlsx";
const EXPORT_XLSX_TO_GITHUB = true;

function doGet(e) {
  return json_({ ok: true, service: "wedding-invitation-data" });
}

function doPost(e) {
  try {
    const data = e && e.parameter ? e.parameter : {};
    const action = String(data.action || "");
    const ss = getSpreadsheet_();
    setupSheets_(ss);

    if (action === "contact") {
      appendContact_(ss, data);
      const githubExport = EXPORT_XLSX_TO_GITHUB ? tryExportWorkbookToGitHub_(ss) : { exported: false, reason: "disabled" };
      return json_({ ok: true, saved: "contact", githubExport: githubExport });
    }

    if (action === "visit") {
      appendVisitor_(ss, data);
      updateDashboard_(ss);
      const visitorCount = ss.getSheetByName("Visitors").getLastRow() - 1;
      if (visitorCount > 0 && visitorCount % 10 === 0) {
        tryExportWorkbookToGitHub_(ss);
      }
      return json_({ ok: true, saved: "visitor" });
    }

    return json_({ ok: false, error: "Unknown action" });
  } catch (err) {
    console.error(err);
    return json_({ ok: false, error: String(err) });
  }
}

function getSpreadsheet_() {
  const props = PropertiesService.getScriptProperties();
  let id = props.getProperty("SPREADSHEET_ID");

  if (id) {
    try { return SpreadsheetApp.openById(id); } catch (e) {}
  }

  const ss = SpreadsheetApp.create("Digital Innovation Tech - Wedding Website Data");
  props.setProperty("SPREADSHEET_ID", ss.getId());
  return ss;
}

function setupSheets_(ss) {
  let contacts = ss.getSheetByName("Contacts");
  if (!contacts) contacts = ss.insertSheet("Contacts");
  if (contacts.getLastRow() === 0) {
    contacts.appendRow(["Timestamp","Name","Mobile","Email","Subject","Comments","Language","Page","Referrer","User Agent","Screen Size"]);
    contacts.setFrozenRows(1);
  }

  let visitors = ss.getSheetByName("Visitors");
  if (!visitors) visitors = ss.insertSheet("Visitors");
  if (visitors.getLastRow() === 0) {
    visitors.appendRow(["Timestamp","Page","Language","Referrer","Screen Size","User Agent","Visit ID"]);
    visitors.setFrozenRows(1);
  }

  updateDashboard_(ss);
}

function updateDashboard_(ss) {
  let dashboard = ss.getSheetByName("Dashboard");
  if (!dashboard) dashboard = ss.insertSheet("Dashboard", 0);

  dashboard.clear();
  dashboard.getRange("A1").setValue("Digital Innovation Tech - Wedding Website Dashboard");
  dashboard.getRange("A3").setValue("Total Page Visits");
  dashboard.getRange("B3").setFormula("=MAX(0,COUNTA(Visitors!A2:A))");
  dashboard.getRange("A4").setValue("Unique Browser Sessions");
  dashboard.getRange("B4").setFormula('=IFERROR(COUNTA(UNIQUE(FILTER(Visitors!G2:G,Visitors!G2:G<>""))),0)');
  dashboard.getRange("A5").setValue("Contact Enquiries");
  dashboard.getRange("B5").setFormula("=MAX(0,COUNTA(Contacts!A2:A))");
  dashboard.getRange("A7").setValue("Note");
  dashboard.getRange("B7").setValue("Visitors are anonymous technical visits. Names are collected only when a visitor submits the contact form.");
  dashboard.getRange("A1:B1").merge();
  dashboard.getRange("A1").setFontWeight("bold");
  dashboard.getRange("A3:A5").setFontWeight("bold");
  dashboard.autoResizeColumns(1, 2);
}

function appendContact_(ss, d) {
  const sheet = ss.getSheetByName("Contacts");
  sheet.appendRow([
    new Date(), clean_(d.name), clean_(d.mobile), clean_(d.email),
    clean_(d.subject), clean_(d.comments), clean_(d.language),
    clean_(d.page), clean_(d.referrer), clean_(d.userAgent), clean_(d.screen)
  ]);
  sheet.autoResizeColumns(1, 11);
}

function appendVisitor_(ss, d) {
  const sheet = ss.getSheetByName("Visitors");
  sheet.appendRow([
    new Date(), clean_(d.page), clean_(d.language), clean_(d.referrer),
    clean_(d.screen), clean_(d.userAgent), clean_(d.visitId)
  ]);
  sheet.autoResizeColumns(1, 7);
}

function clean_(value) {
  return String(value || "").slice(0, 4000);
}

function tryExportWorkbookToGitHub_(ss) {
  try {
    exportWorkbookToGitHub_(ss);
    return { exported: true };
  } catch (err) {
    console.warn("Google Sheet was saved, but optional GitHub XLSX export was skipped:", err);
    return { exported: false, reason: String(err) };
  }
}

function exportWorkbookToGitHub_(ss) {
  const props = PropertiesService.getScriptProperties();
  const token = props.getProperty("GITHUB_TOKEN");
  const repo = props.getProperty("GITHUB_REPO") || "Santosh93980/wedding-invitation";
  if (!token) throw new Error("Missing GITHUB_TOKEN Script Property.");

  SpreadsheetApp.flush();

  const exportUrl = "https://docs.google.com/spreadsheets/d/" + encodeURIComponent(ss.getId()) + "/export?format=xlsx";
  const blob = UrlFetchApp.fetch(exportUrl, {
    headers: { Authorization: "Bearer " + ScriptApp.getOAuthToken() },
    muteHttpExceptions: true
  }).getBlob();

  const base64 = Utilities.base64Encode(blob.getBytes());
  const apiUrl = "https://api.github.com/repos/" + repo + "/contents/" + XLSX_PATH;

  let existingSha = null;
  const getResponse = UrlFetchApp.fetch(apiUrl, {
    method: "get",
    headers: { Authorization: "Bearer " + token, Accept: "application/vnd.github+json" },
    muteHttpExceptions: true
  });
  if (getResponse.getResponseCode() === 200) {
    existingSha = JSON.parse(getResponse.getContentText()).sha || null;
  }

  const payload = {
    message: "Update wedding website data workbook",
    content: base64,
    branch: "main"
  };
  if (existingSha) payload.sha = existingSha;

  const response = UrlFetchApp.fetch(apiUrl, {
    method: "put",
    contentType: "application/json",
    headers: { Authorization: "Bearer " + token, Accept: "application/vnd.github+json" },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });

  const code = response.getResponseCode();
  if (code < 200 || code >= 300) {
    throw new Error("GitHub workbook update failed: " + code + " " + response.getContentText());
  }
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
