// ===== ID SUDAH DISET SESUAI PUNYA ANDA =====
const SPREADSHEET_ID = "1eFvsczVgzCr-YtuqSUwbC1vXOvlIlsMYk16GEDGdO6I";
const FOLDER_ID = "1h1ZSB5EBirCKXE3qq_do4hAHJQMGCBDN";

// ===== PIN ADMIN (ubah jika perlu) =====
const ADMIN_PIN = "123456";

// ===== UKURAN GAMBAR DI WORD (CM) =====
const IMG_WIDTH_CM = 5;
const IMG_HEIGHT_CM = 6;

function cmToPt_(cm) {
  return cm * 28.3464567; // 1 inch=72pt, 1 inch=2.54cm
}

// ==============================

function doGet(e) {
  const mode = (e && e.parameter && e.parameter.mode) ? String(e.parameter.mode) : "";
  const t = HtmlService.createTemplateFromFile("Index");
  t.MODE = mode;
  return t.evaluate().setTitle("Upload SPT Tahunan");
}

function ensureSheets_() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);

  // Sheet PEGAWAI
  let peg = ss.getSheetByName("PEGAWAI");
  if (!peg) {
    peg = ss.insertSheet("PEGAWAI");
    peg.appendRow(["Nama", "NIP/NIPPPK", "Jabatan"]);
  }

  // Sheet INPUT (timestamp masih boleh ada untuk arsip, tapi tidak ditampilkan di Word)
  let inp = ss.getSheetByName("INPUT");
  if (!inp) {
    inp = ss.insertSheet("INPUT");
    inp.appendRow(["Timestamp", "Nama", "NIP/NIPPPK", "Jabatan", "NamaFileSPT", "LinkFileDrive"]);
  }
}

function getPegawai() {
  ensureSheets_();

  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sh = ss.getSheetByName("PEGAWAI");
  const values = sh.getDataRange().getValues();
  values.shift(); // buang header

  return values
    .filter(r => r[0])
    .map(r => ({
      nama: String(r[0]),
      nipppk: String(r[1] || ""),
      jabatan: String(r[2] || "")
    }));
}

// data = { nama, nipppk, jabatan, fileName, base64 }
function submitInput(data) {
  ensureSheets_();

  if (!data?.nama) throw new Error("Nama belum dipilih.");
  if (!data?.base64) throw new Error("File SPT belum diupload.");

  const fileName = String(data.fileName || "SPT.jpg");
  const lower = fileName.toLowerCase();

  if (!lower.match(/\.(jpg|jpeg|png)$/)) throw new Error("File harus JPG/JPEG/PNG.");

  const mime = lower.endsWith(".png") ? "image/png" : "image/jpeg";

  // Simpan file upload ke Drive (folder Anda)
  const folder = DriveApp.getFolderById(FOLDER_ID);
  const bytes = Utilities.base64Decode(data.base64);
  const blob = Utilities.newBlob(bytes, mime, fileName);
  const file = folder.createFile(blob);

  // Agar link bisa dibuka admin; jika ingin lebih privat, ubah sesuai kebutuhan
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

  // Simpan data ke sheet INPUT
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sh = ss.getSheetByName("INPUT");
  sh.appendRow([
    new Date(),
    data.nama,
    data.nipppk,
    data.jabatan,
    fileName,
    file.getUrl()
  ]);

  return { ok: true, link: file.getUrl() };
}

function isAdmin(pin) {
  return String(pin || "") === ADMIN_PIN;
}

function generateWordDocx(pin) {
  if (!isAdmin(pin)) throw new Error("PIN admin salah.");

  ensureSheets_();

  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sh = ss.getSheetByName("INPUT");
  const values = sh.getDataRange().getValues();

  if (values.length <= 1) throw new Error("Belum ada data di tab INPUT.");

  const rows = values.slice(1);

  // 1) Buat Google Doc
  const doc = DocumentApp.create("Rekap SPT Tahunan");
  const body = doc.getBody();

  body.appendParagraph("REKAP UPLOAD SPT TAHUNAN").setBold(true);
  body.appendParagraph("");

  // 2) Buat tabel TANPA Timestamp
  const tableData = [];
  tableData.push(["No", "Nama", "NIP/NIPPPK", "Jabatan", "SPT Tahunan"]);

  rows.forEach((r, i) => {
    tableData.push([
      String(i + 1),
      String(r[1] || ""), // Nama
      String(r[2] || ""), // NIP/NIPPPK
      String(r[3] || ""), // Jabatan
      ""                  // gambar di sini
    ]);
  });

  const table = body.appendTable(tableData);

  // 3) Sisipkan gambar (ukuran 5cm x 6cm)
  const wPt = cmToPt_(IMG_WIDTH_CM);
  const hPt = cmToPt_(IMG_HEIGHT_CM);

  rows.forEach((r, i) => {
    const link = String(r[5] || ""); // LinkFileDrive
    if (!link) return;

    const fileId = extractDriveFileId_(link);
    if (!fileId) return;

    const imgBlob = DriveApp.getFileById(fileId).getBlob();
    const cell = table.getCell(i + 1, 4); // row i+1 (header), col 4 = SPT Tahunan
    cell.clear();

    const img = cell.appendParagraph("").appendInlineImage(imgBlob);
    img.setWidth(wPt);
    img.setHeight(hPt);
  });

  doc.saveAndClose();

  // 4) Export DOCX (cara stabil): UrlFetch Drive export endpoint
  const docxBlob = exportGoogleDocToDocxBlob_(doc.getId(), "Rekap_SPT_Tahunan.docx");

  // Simpan DOCX ke folder Drive Anda
  const outFile = DriveApp.getFolderById(FOLDER_ID).createFile(docxBlob);
  outFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

  return { ok: true, docxLink: outFile.getUrl() };
}

/**
 * Export Google Doc (Google-native) menjadi DOCX blob via Drive export endpoint.
 */
function exportGoogleDocToDocxBlob_(fileId, outName) {
  const url = "https://www.googleapis.com/drive/v3/files/" + encodeURIComponent(fileId) +
    "/export?mimeType=" + encodeURIComponent("application/vnd.openxmlformats-officedocument.wordprocessingml.document");

  const token = ScriptApp.getOAuthToken();
  const resp = UrlFetchApp.fetch(url, {
    method: "get",
    headers: { Authorization: "Bearer " + token },
    muteHttpExceptions: true
  });

  const code = resp.getResponseCode();
  if (code !== 200) {
    throw new Error("Gagal export DOCX. HTTP " + code + " | " + resp.getContentText().slice(0, 300));
  }

  const blob = resp.getBlob();
  blob.setName(outName || "Rekap_SPT_Tahunan.docx");
  return blob;
}

function extractDriveFileId_(url) {
  const m1 = url.match(/\/d\/([a-zA-Z0-9_-]{20,})/);
  if (m1) return m1[1];

  const m2 = url.match(/[?&]id=([a-zA-Z0-9_-]{20,})/);
  if (m2) return m2[1];

  const m3 = url.match(/[-\w]{25,}/);
  return m3 ? m3[0] : null;
}