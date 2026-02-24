/**
 * Deploy as Web App:
 * Execute as: Me
 * Who has access: Anyone
 *
 * Optional: set a folder id to store uploaded certificate images.
 * Leave empty to save in Drive root.
 */
const RESULT_SHEET_NAME = "results";
const DRIVE_FOLDER_ID = "";
const SPREADSHEET_ID = "1d4O2p82cLkJ4m6cKcYa9zSJ11h-k_0I-W-IrY-cX_bA";
const HEADERS = [
  "submittedAt",
  "status",
  "score",
  "total",
  "name",
  "email",
  "paymentMethod",
  "paymentAccount",
  "wechat",
  "age",
  "gender",
  "nationality",
  "nativeLanguages",
  "learningLength",
  "proficiency",
  "highestHSK",
  "latestHSKDate",
  "motivation",
  "certificateFileName",
  "certificateImageUrl",
  "certificateDriveFileId",
  "certificateUploadError",
  "learningContextsJson",
  "learningContextsOther",
  "chinaVisitDetail",
  "answersJson"
];

function doGet() {
  return ContentService
    .createTextOutput(JSON.stringify({ ok: true, service: "pre-screening-upload" }))
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  try {
    const data = JSON.parse((e && e.postData && e.postData.contents) || "{}");
    const ss = getSpreadsheet_();
    const sheet = ss.getSheetByName(RESULT_SHEET_NAME) || ss.insertSheet(RESULT_SHEET_NAME);

    ensureHeader(sheet);

    const c = data.candidate || {};
    const upload = saveCertificateImageIfAny_(c, data.submittedAt);

    const rowMap = {
      submittedAt: data.submittedAt || "",
      status: data.status || "",
      score: data.score || "",
      total: data.total || "",
      name: c.name || "",
      email: c.email || "",
      paymentMethod: c.paymentMethod || "",
      paymentAccount: c.paymentAccount || "",
      wechat: c.wechat || "",
      age: c.age || "",
      gender: c.gender || "",
      nationality: c.nationality || "",
      nativeLanguages: c.nativeLanguages || "",
      learningLength: c.learningLength || "",
      proficiency: c.proficiency || "",
      highestHSK: c.highestHSK || "",
      latestHSKDate: c.latestHSKDate || "",
      motivation: c.motivation || "",
      certificateFileName: c.certificateInfo || "",
      certificateImageUrl: upload.url || "",
      certificateDriveFileId: upload.fileId || "",
      certificateUploadError: upload.error || "",
      learningContextsJson: JSON.stringify(c.learningContexts || []),
      learningContextsOther: c.learningContextsOther || "",
      chinaVisitDetail: c.chinaVisitDetail || "",
      answersJson: JSON.stringify(data.answers || [])
    };
    const row = HEADERS.map((key) => rowMap[key] || "");
    sheet.appendRow(row);

    return ContentService
      .createTextOutput(JSON.stringify({ ok: true, certificateImageUrl: upload.url || "" }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ ok: false, error: String(err) }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function ensureHeader(sheet) {
  if (!sheet) {
    throw new Error("Target sheet is not available.");
  }
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(HEADERS);
    return;
  }
  const current = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(String);
  if (current.join("||") === HEADERS.join("||")) return;
  sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
}

function getSpreadsheet_() {
  if (SPREADSHEET_ID) {
    return SpreadsheetApp.openById(SPREADSHEET_ID);
  }
  const active = SpreadsheetApp.getActiveSpreadsheet();
  if (!active) {
    throw new Error("No active spreadsheet. Please set SPREADSHEET_ID.");
  }
  return active;
}

function saveCertificateImageIfAny_(candidate, submittedAt) {
  try {
    const cert = candidate && candidate.certificateImage;
    if (!cert || !cert.base64) {
      return { url: "", fileId: "", error: "" };
    }

    const mimeType = cert.mimeType || "image/jpeg";
    const ext = mimeToExt_(mimeType);
    const safeName = sanitizeName_(cert.fileName || "certificate");
    const stamp = sanitizeName_(submittedAt || new Date().toISOString());
    const fileName = "Q12_" + stamp + "_" + safeName + "." + ext;

    const blob = Utilities.newBlob(Utilities.base64Decode(cert.base64), mimeType, fileName);
    const file = DRIVE_FOLDER_ID
      ? DriveApp.getFolderById(DRIVE_FOLDER_ID).createFile(blob)
      : DriveApp.createFile(blob);

    return {
      url: file.getUrl(),
      fileId: file.getId(),
      error: ""
    };
  } catch (err) {
    return {
      url: "",
      fileId: "",
      error: String(err)
    };
  }
}

function mimeToExt_(mimeType) {
  if (mimeType === "image/png") return "png";
  if (mimeType === "image/webp") return "webp";
  if (mimeType === "image/gif") return "gif";
  return "jpg";
}

function sanitizeName_(s) {
  return String(s || "")
    .replace(/[^\w\-.]+/g, "_")
    .replace(/^_+|_+$/g, "")
    .slice(0, 80) || "file";
}
