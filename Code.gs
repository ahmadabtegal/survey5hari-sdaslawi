/**
 * API Dashboard Survei Pembelajaran 5 Hari Sekolah
 * Letakkan kode ini di Apps Script yang terhubung ke Google Sheets respons Form.
 *
 * Sheet default: "Form Responses 1"
 * Data identitas (nama santri/wali) TIDAK dikirim ke dashboard.
 */
const CONFIG = {
  SHEET_NAME: "Form Responses 1",
  HEADER_CLASS: "Kelas",
  HEADER_CHOICE: "Apakah Bapak/Ibu setuju dengan penerapan sistem pembelajaran 5 (lima) hari sekolah",
  HEADER_REASON: "Apa alasan atau pertimbangan Bapak/Ibu"
};

function doGet(e) {
  const payload = getDashboardData_();
  const callback = e && e.parameter ? e.parameter.callback : "";

  if (callback) {
    // JSONP dipakai agar dashboard GitHub Pages dapat mengambil data lintas domain.
    const safeCallback = String(callback).replace(/[^\w.$]/g, "");
    return ContentService
      .createTextOutput(safeCallback + "(" + JSON.stringify(payload) + ");")
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }

  return ContentService
    .createTextOutput(JSON.stringify(payload))
    .setMimeType(ContentService.MimeType.JSON);
}

function getDashboardData_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
  if (!sheet) {
    return { ok: false, message: "Sheet '" + CONFIG.SHEET_NAME + "' tidak ditemukan." };
  }

  const values = sheet.getDataRange().getDisplayValues();
  if (values.length < 2) {
    return {
      ok: true, updatedAt: new Date().toISOString(), total: 0,
      choices: emptyChoices_(), classes: [], reasons: []
    };
  }

  const headers = values[0].map(v => String(v || "").trim());
  const classIndex = findHeader_(headers, CONFIG.HEADER_CLASS, ["kelas"]);
  const choiceIndex = findHeader_(headers, CONFIG.HEADER_CHOICE, ["apakah bapak", "setuju", "5 (lima) hari"]);
  const reasonIndex = findHeader_(headers, CONFIG.HEADER_REASON, ["alasan", "pertimbangan"]);

  if (classIndex < 0 || choiceIndex < 0) {
    return {
      ok: false,
      message: "Kolom Kelas atau pertanyaan persetujuan tidak ditemukan. Periksa nama header pada baris pertama."
    };
  }

  const choices = emptyChoices_();
  const classMap = {};
  const reasons = [];
  let total = 0;

  values.slice(1).forEach(row => {
    const rawChoice = String(row[choiceIndex] || "").trim();
    if (!rawChoice) return;

    const normalized = normalizeChoice_(rawChoice);
    if (!normalized) return;

    total++;
    choices[normalized]++;

    const className = String(row[classIndex] || "Tidak diisi").trim() || "Tidak diisi";
    if (!classMap[className]) classMap[className] = emptyChoices_();
    classMap[className][normalized]++;

    const reason = reasonIndex >= 0 ? String(row[reasonIndex] || "").trim() : "";
    if (reason) reasons.push(reason);
  });

  const classes = Object.keys(classMap).sort((a, b) =>
    a.localeCompare(b, "id", { numeric: true, sensitivity: "base" })
  ).map(name => ({ name: name, total: sumChoices_(classMap[name]), choices: classMap[name] }));

  return {
    ok: true,
    updatedAt: new Date().toISOString(),
    total: total,
    choices: choices,
    classes: classes,
    reasons: reasons.slice(-100).reverse()
  };
}

function findHeader_(headers, preferred, fallbacks) {
  let i = headers.findIndex(h => h.toLowerCase().includes(preferred.toLowerCase()));
  if (i >= 0) return i;
  return headers.findIndex(h => fallbacks.every(part => h.toLowerCase().includes(part)));
}

function normalizeChoice_(value) {
  const v = value.toLowerCase();
  if (v.includes("sangat setuju")) return "Sangat Setuju";
  if (v.includes("kurang setuju")) return "Kurang Setuju";
  if (v.includes("tidak setuju")) return "Tidak Setuju";
  if (v.includes("setuju")) return "Setuju";
  return "";
}

function emptyChoices_() {
  return {
    "Sangat Setuju": 0,
    "Setuju": 0,
    "Kurang Setuju": 0,
    "Tidak Setuju": 0
  };
}

function sumChoices_(obj) {
  return Object.keys(obj).reduce((sum, key) => sum + obj[key], 0);
}
