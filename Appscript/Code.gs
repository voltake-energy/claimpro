const DEFAULT_SPREADSHEET_ID = "1i-OFalyNXcWirJFlbfAK-tYlrWPYrT8CLZBsrTFBU9c";
const DEFAULT_MAIN_FOLDER_ID = "15eDunv7n2GcC-MvExK9k1j9sqs_rnRtb";

// Pilihan B (lebih ketat) — tukar dua baris di atas jadi "", kemudian simpan
// ID sebenar dalam Project Settings > Script Properties:
//     SPREADSHEET_ID = 1i-OFaly...
//     MAIN_FOLDER_ID = 15eDunv7...
// Script Properties tak masuk repo, jadi ID langsung tak nampak di GitHub.
function cfg(key, fallback) {
const v = PropertiesService.getScriptProperties().getProperty(key);
return v && v.trim() ? v.trim() : fallback;
}
const SPREADSHEET_ID = cfg("SPREADSHEET_ID", DEFAULT_SPREADSHEET_ID);
const MAIN_FOLDER_ID = cfg("MAIN_FOLDER_ID", DEFAULT_MAIN_FOLDER_ID);

const SHEET_CLAIMS   = "Database Tuntutan";
const SHEET_STAFF    = "Senarai Pekerja";
const SHEET_SESSIONS = "Sessions";
const SHEET_AUDIT    = "Audit Log";

// Session TTL (Teknik #3)
const SESSION_TTL_HOURS  = 12;   // token kerja harian
const REMEMBER_TTL_DAYS  = 30;   // "kekal log masuk" pada peranti sendiri

// Rate limit login (Teknik #5)
const LOGIN_MAX_ATTEMPTS = 8;
const LOGIN_WINDOW_SECS  = 300;  // 5 minit

// Peraturan perniagaan — divalidate di SERVER, bukan hanya di borang (Teknik #4)
const ANNUAL_MEDICAL_LIMIT = 350;
const MEAL_RATE_SIANG = 10;
const MEAL_RATE_MALAM = 15;
const MEAL_RATE_BOTH  = 25;
const MAX_CLAIM_AMOUNT = 50000;         // guard input tak masuk akal
const MAX_FILES_PER_CLAIM = 10;
const MAX_FILE_BYTES = 10 * 1024 * 1024; // 10MB sefail

const ALLOWED_TYPES = [
"Perjalanan", "Makan & Minum", "Keraian",
"Penginapan", "Perubatan", "Lain-lain"
];
const ALLOWED_MIME = [
"application/pdf", "image/png", "image/jpeg", "image/jpg", "image/heic", "image/webp"
];

// PDPA (Teknik #7) — tempoh simpan rekod
const RETENTION_YEARS_CLAIMS = 7;   // selaras keperluan rekod LHDN
const RETENTION_DAYS_AUDIT   = 365;

// ==========================================================================
// CUSTOM MENU DI GOOGLE SHEETS (Sebelah Menu 'Help')
// ==========================================================================

function onOpen() {
try {
const ui = SpreadsheetApp.getUi();
ui.createMenu("⚙️ Sistem Admin")
.addItem("▶️ Jalankan Setup Sistem (setupSystem)", "menuRunSetupSystem")
.addItem("⏱️ Pasang Triggers Automatik", "menuRunInstallTriggers")
.addSeparator()
.addItem("🔑 Reset Kata Laluan Staf", "menuRunResetPassword")
.addItem("🧹 Bersihkan Sesi Tamat Tempoh", "menuRunPurgeSessions")
.addToUi();
} catch (err) {
console.error("onOpen error: " + err);
}
}

function menuRunSetupSystem() {
const ui = SpreadsheetApp.getUi();
const resp = ui.alert(
"Pengesahan Setup Sistem",
"Adakah anda pasti ingin menjalankan setupSystem? Ini akan mencipta pepper, struktur tab yang diperlukan dan melakukan hashing kata laluan.",
ui.ButtonSet.YES_NO
);
if (resp === ui.Button.YES) {
try {
const msg = setupSystem();
ui.alert("Berjaya", msg, ui.ButtonSet.OK);
} catch (e) {
ui.alert("Ralat", "Gagal menjalankan setupSystem: " + e.message, ui.ButtonSet.OK);
}
}
}

function menuRunInstallTriggers() {
const ui = SpreadsheetApp.getUi();
try {
const msg = installTriggers();
ui.alert("Berjaya", msg, ui.ButtonSet.OK);
} catch (e) {
ui.alert("Ralat", "Gagal memasang triggers: " + e.message, ui.ButtonSet.OK);
}
}

function menuRunResetPassword() {
const ui = SpreadsheetApp.getUi();
const idPrompt = ui.prompt("Reset Kata Laluan", "Masukkan ID Pekerja (cth: VES003):", ui.ButtonSet.OK_CANCEL);
if (idPrompt.getSelectedButton() !== ui.Button.OK) return;
const userId = idPrompt.getResponseText().trim();
if (!userId) {
ui.alert("ID Pekerja tidak boleh kosong.");
return;
}

const pwdPrompt = ui.prompt("Reset Kata Laluan", "Masukkan Kata Laluan Baharu Sementara untuk " + userId + ":", ui.ButtonSet.OK_CANCEL);
if (pwdPrompt.getSelectedButton() !== ui.Button.OK) return;
const newPwd = pwdPrompt.getResponseText().trim();
if (!newPwd) {
ui.alert("Kata laluan tidak boleh kosong.");
return;
}

try {
const msg = resetPassword(userId, newPwd);
ui.alert("Selesai", msg, ui.ButtonSet.OK);
} catch (e) {
ui.alert("Ralat", e.message, ui.ButtonSet.OK);
}
}

function menuRunPurgeSessions() {
const ui = SpreadsheetApp.getUi();
try {
purgeExpiredSessions();
ui.alert("Selesai", "Pembersihan token sesi tamat tempoh telah selesai.", ui.ButtonSet.OK);
} catch (e) {
ui.alert("Ralat", e.message, ui.ButtonSet.OK);
}
}

// ==========================================================================
// ENTRY POINT
// ==========================================================================

function doPost(e) {
const lock = LockService.getScriptLock();
try {
lock.waitLock(20000);
} catch (err) {
return jsonOut({ status: "error", message: "Pelayan sibuk. Cuba sekejap lagi." });
}

try {
if (!e || !e.postData || !e.postData.contents) {
return jsonOut({ status: "error", message: "Permintaan tidak sah." });
}

let body;
try {
  body = JSON.parse(e.postData.contents);
} catch (err) {
  return jsonOut({ status: "error", message: "Format permintaan tidak sah." });
}

const action = String(body.action || "").trim();

// --- Action TANPA token (satu-satunya yang dibenarkan) ---
if (action === "getStaffList") return handleGetStaffList(body);
if (action === "login")        return handleLogin(body);
if (action === "resume")       return handleResume(body);

// --- SEMUA action lain WAJIB ada session token yang sah (Teknik #3, #4) ---
switch (action) {
  case "logout":         return handleLogout(body);
  case "getData":        return handleGetData(body);
  case "submitClaim":    return handleSubmitClaim(body);
  case "editClaim":      return handleEditClaim(body);
  case "deleteClaim":    return handleDeleteClaim(body);
  case "updatePayment":  return handleUpdatePayment(body);
  case "changePassword": return handleChangePassword(body);
  default:
    return jsonOut({ status: "error", message: "Action tidak dikenali." });
}


} catch (error) {
console.error("doPost error: " + error.stack);
const msg = String(error.message || error);
const safe = /^(Unauthorized|Forbidden|Terlalu banyak|Input tidak sah)/.test(msg)
? msg
: "Ralat pelayan. Sila hubungi HR jika berulang.";
return jsonOut({ status: "error", message: safe });
} finally {
lock.releaseLock();
}
}

function doGet() {
return ContentService
.createTextOutput("ClaimPro HR backend aktif. Semua operasi melalui POST.")
.setMimeType(ContentService.MimeType.TEXT);
}

function jsonOut(obj) {
return ContentService.createTextOutput(JSON.stringify(obj))
.setMimeType(ContentService.MimeType.JSON);
}

// ==========================================================================
// TEKNIK #2 — PASSWORD: SHA-256 + SALT UNIK PER-USER
// ==========================================================================

function getPepper() {
const props = PropertiesService.getScriptProperties();
let pepper = props.getProperty("PWD_PEPPER");
if (!pepper) {
pepper = Utilities.getUuid() + Utilities.getUuid();
props.setProperty("PWD_PEPPER", pepper);
}
return pepper;
}

function generateSalt() {
return Utilities.getUuid().replace(/-/g, "");
}

function hashPassword(password, salt) {
const raw = Utilities.computeDigest(
Utilities.DigestAlgorithm.SHA_256,
String(password) + String(salt) + getPepper(),
Utilities.Charset.UTF_8
);
return raw.map(function (b) { return (b + 256) % 256; })
.map(function (b) { return b.toString(16).padStart(2, "0"); })
.join("");
}

function safeEquals(a, b) {
a = String(a || ""); b = String(b || "");
if (a.length !== b.length) return false;
let diff = 0;
for (let i = 0; i < a.length; i++) diff |= a.charCodeAt(i) ^ b.charCodeAt(i);
return diff === 0;
}

// ==========================================================================
// TEKNIK #3 — SESSION TOKEN + TTL (disimpan dalam tab 'Sessions')
// ==========================================================================

function getSessionSheet() {
const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
let sh = ss.getSheetByName(SHEET_SESSIONS);
if (!sh) {
sh = ss.insertSheet(SHEET_SESSIONS);
sh.appendRow(["Token", "UserId", "Nama", "Role", "Jenis", "Expiry", "Dicipta"]);
sh.getRange("A1:G1").setFontWeight("bold").setBackground("#FCE5CD");
sh.setFrozenRows(1);
}
return sh;
}

function createSession(user, kind) {
const sh = getSessionSheet();
const token = Utilities.getUuid() + "." + Utilities.getUuid();
const now = new Date();
const expiry = new Date(
kind === "remember"
? now.getTime() + REMEMBER_TTL_DAYS * 24 * 60 * 60 * 1000
: now.getTime() + SESSION_TTL_HOURS * 60 * 60 * 1000
);
sh.appendRow([token, user.id, user.name, user.role, kind, expiry, now]);
return token;
}

function findSessionRow(token) {
if (!token || String(token).length < 20) return null;
const sh = getSessionSheet();
const rows = sh.getDataRange().getValues();
for (let i = 1; i < rows.length; i++) {
if (safeEquals(rows[i][0], token)) {
return { rowIndex: i + 1, data: rows[i] };
}
}
return null;
}

function validateSession(token, kind) {
const hit = findSessionRow(token);
if (!hit) return null;

const row = hit.data;
const sessionKind = String(row[4] || "session");
if (kind && sessionKind !== kind) return null;

const expiry = new Date(row[5]);
if (isNaN(expiry.getTime()) || expiry.getTime() < Date.now()) {
getSessionSheet().deleteRow(hit.rowIndex);
return null;
}

return {
token: row[0],
id: String(row[1]),
name: String(row[2]),
role: String(row[3] || "staff").toLowerCase(),
kind: sessionKind
};
}

function revokeSession(token) {
const hit = findSessionRow(token);
if (hit) getSessionSheet().deleteRow(hit.rowIndex);
}

function revokeAllSessionsFor(userId) {
const sh = getSessionSheet();
const rows = sh.getDataRange().getValues();
for (let i = rows.length - 1; i >= 1; i--) {
if (String(rows[i][1]) === String(userId)) sh.deleteRow(i + 1);
}
}

function purgeExpiredSessions() {
const sh = getSessionSheet();
const rows = sh.getDataRange().getValues();
const now = Date.now();
let removed = 0;
for (let i = rows.length - 1; i >= 1; i--) {
const exp = new Date(rows[i][5]);
if (isNaN(exp.getTime()) || exp.getTime() < now) { sh.deleteRow(i + 1); removed++; }
}
console.log("purgeExpiredSessions: " + removed + " token dibuang.");
}

// ==========================================================================
// TEKNIK #4 — AUTH & ROLE CHECK DI SERVER
// ==========================================================================

function requireAuth(body, allowedRoles) {
const session = validateSession(body.token, "session");
if (!session) throw new Error("Unauthorized — sesi tamat, sila log masuk semula.");
if (allowedRoles && allowedRoles.indexOf(session.role) === -1) {
throw new Error("Forbidden — anda tiada kebenaran untuk tindakan ini.");
}
return session;
}

// ==========================================================================
// TEKNIK #5 — RATE LIMITING
// ==========================================================================

function checkRateLimit(bucket, identifier, maxAttempts, windowSecs) {
const cache = CacheService.getScriptCache();
const key = bucket + "_" + String(identifier).toLowerCase().slice(0, 80);
const count = Number(cache.get(key) || 0);
if (count >= maxAttempts) {
throw new Error("Terlalu banyak percubaan. Cuba lagi dalam beberapa minit.");
}
cache.put(key, String(count + 1), windowSecs);
}

function clearRateLimit(bucket, identifier) {
CacheService.getScriptCache()
.remove(bucket + "_" + String(identifier).toLowerCase().slice(0, 80));
}

// ==========================================================================
// TEKNIK #6 — SANITIZE INPUT / VALIDASI
// ==========================================================================

function sanitizeForSheet(value) {
const str = String(value === null || value === undefined ? "" : value);
if (/^[=+-@\t\r]/.test(str)) return "'" + str;
return str;
}

function cleanText(value, maxLen) {
let str = String(value === null || value === undefined ? "" : value);
str = str.replace(/[\u0000-\u001F\u007F]/g, " ").trim();
if (maxLen && str.length > maxLen) str = str.slice(0, maxLen);
return str;
}

function cleanFileName(value) {
let str = cleanText(value, 120);
str = str.replace(/[/\:*?"<>|]/g, "_");
return str || "resit";
}

function parseAmount(value) {
const n = parseFloat(String(value).replace(/,/g, ""));
if (isNaN(n) || !isFinite(n)) throw new Error("Input tidak sah: jumlah tuntutan.");
return Math.round(n * 100) / 100;
}

function parseIsoDate(value) {
const str = cleanText(value, 30);
if (!/^\d{4}-\d{2}-\d{2}$/.test(str)) throw new Error("Input tidak sah: format tarikh.");
const d = new Date(str + "T00:00:00");
if (isNaN(d.getTime())) throw new Error("Input tidak sah: tarikh.");
const now = new Date();
const maxFuture = new Date(now.getTime() + 24 * 60 * 60 * 1000);
const minPast = new Date(now.getFullYear() - 3, 0, 1);
if (d > maxFuture) throw new Error("Input tidak sah: tarikh belum berlaku.");
if (d < minPast) throw new Error("Input tidak sah: tarikh terlalu lama (lebih 3 tahun).");
return str;
}

// ==========================================================================
// TEKNIK #7 — AUDIT LOG (PDPA)
// ==========================================================================

function auditLog(userId, action, detail) {
try {
const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
let sh = ss.getSheetByName(SHEET_AUDIT);
if (!sh) {
sh = ss.insertSheet(SHEET_AUDIT);
sh.appendRow(["Masa", "ID Pengguna", "Tindakan", "Butiran"]);
sh.getRange("A1:D1").setFontWeight("bold").setBackground("#D9D2E9");
sh.setFrozenRows(1);
}
sh.appendRow([new Date(), sanitizeForSheet(userId), sanitizeForSheet(action),
sanitizeForSheet(cleanText(detail, 300))]);
} catch (err) {
console.error("auditLog gagal: " + err);
}
}

// ==========================================================================
// HELPER DATA
// ==========================================================================

function getStaffSheet() {
const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
const sh = ss.getSheetByName(SHEET_STAFF);
if (!sh) throw new Error("Tab '" + SHEET_STAFF + "' tidak dijumpai.");
return sh;
}

function getClaimsSheet() {
const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
let sh = ss.getSheetByName(SHEET_CLAIMS);
if (!sh) {
sh = ss.insertSheet(SHEET_CLAIMS);
sh.appendRow(["ID Tuntutan", "Nama Pekerja", "Tarikh", "Jenis", "Pax", "Keterangan",
"Jumlah (RM)", "Status", "Pautan Resit (Google Drive)",
"Tarikh/Masa Direkodkan", "Resit Pembayaran", "Tarikh Pembayaran",
"Jumlah Pembayaran"]);
sh.getRange("A1:M1").setFontWeight("bold").setBackground("#D9EAD3");
sh.setFrozenRows(1);
}
return sh;
}

function readStaff() {
const rows = getStaffSheet().getDataRange().getValues();
const out = [];
for (let i = 1; i < rows.length; i++) {
const id = String(rows[i][0] || "").trim();
const name = String(rows[i][1] || "").trim();
if (!id || !name) continue;
out.push({
rowIndex: i + 1,
id: id,
name: name,
passwordHash: String(rows[i][2] || "").trim(),
designation: String(rows[i][3] || "STAFF").trim(),
email: String(rows[i][4] || "").trim(),
salt: String(rows[i][5] || "").trim(),
pwdUpdated: rows[i][6] || ""
});
}
return out;
}

function roleOf(designation) {
return String(designation).toLowerCase() === "admin" ? "admin" : "staff";
}

function readClaims() {
const rows = getClaimsSheet().getDataRange().getValues();
const out = [];
for (let i = 1; i < rows.length; i++) {
const id = String(rows[i][0] || "").trim();
if (!id) continue;
const amount = Number(rows[i][6]) || 0;
let paid = rows[i][12];
paid = (paid === "" || paid === null || paid === undefined) ? NaN : Number(paid);
if (isNaN(paid)) {
const st = String(rows[i][7] || "").toLowerCase();
paid = (st === "telah dibayar" || st === "selesai" || st === "selesai dibayar") ? amount : 0;
}
out.push({
id: id,
name: String(rows[i][1] || "").trim(),
date: formatDateCell(rows[i][2]),
type: String(rows[i][3] || ""),
pax: Number(rows[i][4]) || 1,
desc: String(rows[i][5] || ""),
amount: amount,
status: String(rows[i][7] || "Menunggu"),
receiptUrl: String(rows[i][8] || ""),
timestamp: rows[i][9] ? new Date(rows[i][9]).getTime() : Date.now(),
payReceiptUrl: String(rows[i][10] || ""),
payDate: rows[i][11] ? String(rows[i][11]) : "",
paid: paid
});
}
return out;
}

function formatDateCell(cell) {
if (cell instanceof Date && !isNaN(cell.getTime())) {
return Utilities.formatDate(cell, Session.getScriptTimeZone(), "yyyy-MM-dd");
}
const s = String(cell || "").trim();
const m = s.match(/^(\d{4}-\d{2}-\d{2})/);
return m ? m[1] : s;
}

function generateClaimId(ownerName) {
  const now = new Date();
  const tz = Session.getScriptTimeZone();

  const datePart = Utilities.formatDate(now, tz, "yyMMdd");
  const timePart = Utilities.formatDate(now, tz, "HHmm");

  const firstName = String(ownerName || "")
    .trim()
    .split(/\s+/)[0]
    .toUpperCase();

  return "CLM-" + datePart + "-" + timePart + "-" + firstName;
}

function getStaffFolder(name) {
const mainFolder = DriveApp.getFolderById(MAIN_FOLDER_ID);
const safeName = cleanFileName(name);
const it = mainFolder.getFoldersByName(safeName);
return it.hasNext() ? it.next() : mainFolder.createFolder(safeName);
}

function saveFiles(fileArray, ownerName, prefix) {
if (!fileArray || !fileArray.length) return [];
if (fileArray.length > MAX_FILES_PER_CLAIM) {
throw new Error("Input tidak sah: maksimum " + MAX_FILES_PER_CLAIM + " fail satu tuntutan.");
}
const folder = getStaffFolder(ownerName);
const urls = [];
for (let i = 0; i < fileArray.length; i++) {
const f = fileArray[i] || {};
const mime = String(f.fileMimeType || "").toLowerCase();
if (ALLOWED_MIME.indexOf(mime) === -1) {
throw new Error("Input tidak sah: jenis fail tidak dibenarkan (" + mime + ").");
}
if (!f.fileBase64) throw new Error("Input tidak sah: fail kosong.");
const bytes = Utilities.base64Decode(f.fileBase64);
if (bytes.length > MAX_FILE_BYTES) {
throw new Error("Input tidak sah: saiz fail melebihi 10MB.");
}
const blob = Utilities.newBlob(bytes, mime, prefix + "_" + cleanFileName(f.fileName));
urls.push(folder.createFile(blob).getUrl());
}
return urls;
}

// ==========================================================================
// VALIDASI PERATURAN TUNTUTAN
// ==========================================================================

function validateClaimRules(claim, ownerName, excludeClaimId) {
if (ALLOWED_TYPES.indexOf(claim.type) === -1) {
throw new Error("Input tidak sah: jenis tuntutan.");
}
if (claim.amount <= 0 || claim.amount > MAX_CLAIM_AMOUNT) {
throw new Error("Input tidak sah: jumlah tuntutan di luar julat yang dibenarkan.");
}
if (!Number.isInteger(claim.pax) || claim.pax < 1 || claim.pax > 100) {
throw new Error("Input tidak sah: bilangan pax.");
}

// Had makan & minum
if (claim.type === "Makan & Minum") {
const siang = !!claim.mealSiang;
const malam = !!claim.mealMalam;
let rate = 0;
if (siang && malam) rate = MEAL_RATE_BOTH;
else if (malam) rate = MEAL_RATE_MALAM;
else if (siang) rate = MEAL_RATE_SIANG;
if (rate === 0) throw new Error("Input tidak sah: sila pilih sesi makan (Siang/Malam).");
const maxMeal = rate * claim.pax;
if (claim.amount > maxMeal + 0.001) {
throw new Error("Input tidak sah: melebihi had makan RM " + maxMeal.toFixed(2) + ".");
}
}

// Had perubatan tahunan
if (claim.type === "Perubatan") {
const year = new Date(claim.date + "T00:00:00").getFullYear();
const all = readClaims();
let used = 0;
all.forEach(function (c) {
if (excludeClaimId && c.id === excludeClaimId) return;
if (c.name !== ownerName) return;
if (c.type !== "Perubatan") return;
const cy = new Date(c.date + "T00:00:00").getFullYear();
if (cy === year) used += Number(c.amount) || 0;
});
const balance = ANNUAL_MEDICAL_LIMIT - used;
if (claim.amount > balance + 0.001) {
throw new Error("Input tidak sah: baki perubatan " + year + " tinggal RM "
+ Math.max(0, balance).toFixed(2) + ".");
}
}
}

// ==========================================================================
// HANDLERS
// ==========================================================================

function handleGetStaffList(body) {
checkRateLimit("stafflist", "global", 300, 300);
const staff = readStaff().map(function (s) {
return { id: s.id, name: s.name };
});
return jsonOut({ status: "success", staff: staff });
}

function handleLogin(body) {
const userId = cleanText(body.userId, 40);
const password = String(body.password || "");
if (!userId || !password) {
return jsonOut({ status: "error", message: "ID dan kata laluan diperlukan." });
}

checkRateLimit("login", userId, LOGIN_MAX_ATTEMPTS, LOGIN_WINDOW_SECS);

const staff = readStaff();
const user = staff.filter(function (s) { return s.id === userId; })[0];

const failMsg = "ID atau kata laluan tidak sah.";
if (!user) return jsonOut({ status: "error", message: failMsg });

if (!user.salt || !user.passwordHash || user.passwordHash.length !== 64) {
return jsonOut({
status: "error",
message: "Akaun belum dimigrasi. Hubungi HR untuk jalankan setupSystem()."
});
}

if (!safeEquals(hashPassword(password, user.salt), user.passwordHash)) {
auditLog(userId, "LOGIN_GAGAL", "Kata laluan salah");
return jsonOut({ status: "error", message: failMsg });
}

clearRateLimit("login", userId);

const sessionUser = { id: user.id, name: user.name, role: roleOf(user.designation) };
const token = createSession(sessionUser, "session");
const rememberToken = body.remember ? createSession(sessionUser, "remember") : "";

auditLog(user.id, "LOGIN", "Berjaya");

return jsonOut({
status: "success",
token: token,
rememberToken: rememberToken,
expiresIn: SESSION_TTL_HOURS * 3600,
user: sessionUser,
usingDefaultPassword: !user.pwdUpdated
});
}

function handleResume(body) {
const remembered = validateSession(body.rememberToken, "remember");
if (!remembered) {
return jsonOut({ status: "error", message: "Sesi tamat. Sila log masuk semula." });
}

const user = readStaff().filter(function (s) { return s.id === remembered.id; })[0];
if (!user) {
revokeSession(body.rememberToken);
return jsonOut({ status: "error", message: "Akaun tidak lagi aktif." });
}

const sessionUser = { id: user.id, name: user.name, role: roleOf(user.designation) };
const token = createSession(sessionUser, "session");

return jsonOut({
status: "success",
token: token,
expiresIn: SESSION_TTL_HOURS * 3600,
user: sessionUser,
usingDefaultPassword: !user.pwdUpdated
});
}

function handleLogout(body) {
if (body.token) {
const s = validateSession(body.token, "session");
if (s) auditLog(s.id, "LOGOUT", "");
revokeSession(body.token);
}
if (body.rememberToken) revokeSession(body.rememberToken);
return jsonOut({ status: "success" });
}

function handleGetData(body) {
const session = requireAuth(body);
const allClaims = readClaims();

const claims = session.role === "admin"
? allClaims
: allClaims.filter(function (c) { return c.name === session.name; });

const staff = session.role === "admin"
? readStaff().map(function (s) { return { id: s.id, name: s.name, role: roleOf(s.designation) }; })
: [{ id: session.id, name: session.name, role: session.role }];

return jsonOut({
status: "success",
user: { id: session.id, name: session.name, role: session.role },
claims: claims,
staff: staff,
config: {
annualMedicalLimit: ANNUAL_MEDICAL_LIMIT,
mealRates: { siang: MEAL_RATE_SIANG, malam: MEAL_RATE_MALAM, both: MEAL_RATE_BOTH },
allowedTypes: ALLOWED_TYPES
}
});
}

function handleSubmitClaim(body) {
const session = requireAuth(body);
checkRateLimit("submit", session.id, 30, 300);

const ownerName = session.name;

const claim = {
date: parseIsoDate(body.date),
type: cleanText(body.type, 40),
pax: parseInt(body.pax, 10),
desc: cleanText(body.desc, 300),
amount: parseAmount(body.amount),
mealSiang: !!body.mealSiang,
mealMalam: !!body.mealMalam
};
validateClaimRules(claim, ownerName, null);

const files = body.files || [];
if (!files.length) throw new Error("Input tidak sah: sekurang-kurangnya satu resit diperlukan.");
const fileUrls = saveFiles(files, ownerName, claim.date);

const sheet = getClaimsSheet();
const id = generateClaimId(ownerName);

sheet.appendRow([
sanitizeForSheet(id),
sanitizeForSheet(ownerName),
sanitizeForSheet(claim.date),
sanitizeForSheet(claim.type),
claim.pax,
sanitizeForSheet(claim.desc),
claim.amount,
"Menunggu",
sanitizeForSheet(fileUrls.join(", ")),
new Date(),
"", "", 0
]);

auditLog(session.id, "TUNTUTAN_BARU", id + " RM" + claim.amount.toFixed(2));

return jsonOut({
status: "success",
claim: {
id: id, name: ownerName, date: claim.date, type: claim.type, pax: claim.pax,
desc: claim.desc, amount: claim.amount, status: "Menunggu",
receiptUrl: fileUrls.join(", "), timestamp: Date.now(),
paid: 0, payReceiptUrl: "", payDate: ""
}
});
}

function handleEditClaim(body) {
const session = requireAuth(body);
const id = cleanText(body.id, 40);
const sheet = getClaimsSheet();
const rows = sheet.getDataRange().getValues();

let rowIndex = -1, existing = null;
for (let i = 1; i < rows.length; i++) {
if (String(rows[i][0]).trim() === id) { rowIndex = i + 1; existing = rows[i]; break; }
}
if (rowIndex === -1) return jsonOut({ status: "error", message: "Rekod tidak dijumpai." });

const ownerName = String(existing[1] || "").trim();
if (session.role !== "admin" && ownerName !== session.name) {
throw new Error("Forbidden — anda hanya boleh mengubah tuntutan sendiri.");
}
const alreadyPaid = Number(existing[12]) || 0;
if (alreadyPaid > 0 && session.role !== "admin") {
throw new Error("Forbidden — tuntutan yang telah dibayar tidak boleh diubah.");
}

const claim = {
date: parseIsoDate(body.date),
type: cleanText(body.type, 40),
pax: parseInt(body.pax, 10),
desc: cleanText(body.desc, 300),
amount: parseAmount(body.amount),
mealSiang: !!body.mealSiang,
mealMalam: !!body.mealMalam
};
validateClaimRules(claim, ownerName, id);

sheet.getRange(rowIndex, 3).setValue(sanitizeForSheet(claim.date));
sheet.getRange(rowIndex, 4).setValue(sanitizeForSheet(claim.type));
sheet.getRange(rowIndex, 5).setValue(claim.pax);
sheet.getRange(rowIndex, 6).setValue(sanitizeForSheet(claim.desc));
sheet.getRange(rowIndex, 7).setValue(claim.amount);

let fileUrlString = String(existing[8] || "");
if (body.files && body.files.length) {
const urls = saveFiles(body.files, ownerName, "Edit_" + claim.date);
fileUrlString = urls.join(", ");
sheet.getRange(rowIndex, 9).setValue(sanitizeForSheet(fileUrlString));
}

auditLog(session.id, "TUNTUTAN_EDIT", id);
return jsonOut({ status: "success", fileUrl: fileUrlString });
}

function handleDeleteClaim(body) {
const session = requireAuth(body);
const id = cleanText(body.id, 40);
const sheet = getClaimsSheet();
const rows = sheet.getDataRange().getValues();

for (let i = 1; i < rows.length; i++) {
if (String(rows[i][0]).trim() === id) {
const ownerName = String(rows[i][1] || "").trim();
const paid = Number(rows[i][12]) || 0;
if (session.role !== "admin" && ownerName !== session.name) {
throw new Error("Forbidden — anda hanya boleh memadam tuntutan sendiri.");
}
if (paid > 0 && session.role !== "admin") {
throw new Error("Forbidden — tuntutan yang telah dibayar tidak boleh dipadam.");
}
sheet.deleteRow(i + 1);
auditLog(session.id, "TUNTUTAN_PADAM", id + " (" + ownerName + ")");
return jsonOut({ status: "success" });
}
}
return jsonOut({ status: "error", message: "Rekod tidak dijumpai." });
}

function handleUpdatePayment(body) {
const session = requireAuth(body, ["admin"]);

const payeeName = cleanText(body.name, 120);
const updates = body.updates || [];
if (!payeeName || !updates.length) throw new Error("Input tidak sah: data pembayaran.");

const proof = body.fileBase64
? saveFiles([{ fileBase64: body.fileBase64, fileMimeType: body.fileMimeType, fileName: body.fileName }],
payeeName, "Bayaran")
: [];
if (!proof.length) throw new Error("Input tidak sah: bukti pembayaran diperlukan.");
const fileUrl = proof[0];

const sheet = getClaimsSheet();
const rows = sheet.getDataRange().getValues();
const paymentDateStr = new Date().toISOString();
let applied = 0;

for (let i = 1; i < rows.length; i++) {
const rowId = String(rows[i][0]).trim();
const match = updates.filter(function (u) { return String(u.id).trim() === rowId; })[0];
if (!match) continue;

const amount = Number(rows[i][6]) || 0;
const newPaid = parseAmount(match.paid);
if (newPaid < 0 || newPaid > amount + 0.001) {
  throw new Error("Input tidak sah: bayaran " + rowId + " melebihi jumlah tuntutan.");
}
if (String(rows[i][1] || "").trim() !== payeeName) {
  throw new Error("Input tidak sah: rekod " + rowId + " bukan milik " + payeeName + ".");
}

const status = newPaid >= amount - 0.001 ? "Telah Dibayar" : "Separa Dibayar";
const curReceipts = String(rows[i][10] || "");
const curDates = String(rows[i][11] || "");

sheet.getRange(i + 1, 8).setValue(status);
sheet.getRange(i + 1, 11).setValue(sanitizeForSheet(curReceipts ? curReceipts + ", " + fileUrl : fileUrl));
sheet.getRange(i + 1, 12).setValue(sanitizeForSheet(curDates ? curDates + ", " + paymentDateStr : paymentDateStr));
sheet.getRange(i + 1, 13).setValue(newPaid);
applied++;


}

auditLog(session.id, "BAYARAN", payeeName + " — " + applied + " rekod");
return jsonOut({ status: "success", fileUrl: fileUrl, applied: applied });
}

function handleChangePassword(body) {
const session = requireAuth(body);
checkRateLimit("chpwd", session.id, 10, 600);

const oldPwd = String(body.oldPassword || "");
const newPwd = String(body.newPassword || "");

if (newPwd.length < 8) {
return jsonOut({ status: "error", message: "Kata laluan baharu mesti sekurang-kurangnya 8 aksara." });
}
if (newPwd.length > 128) {
return jsonOut({ status: "error", message: "Kata laluan terlalu panjang." });
}
if (!/[A-Za-z]/.test(newPwd) || !/[0-9]/.test(newPwd)) {
return jsonOut({ status: "error", message: "Kata laluan mesti ada huruf dan nombor." });
}

const sheet = getStaffSheet();
const user = readStaff().filter(function (s) { return s.id === session.id; })[0];
if (!user) return jsonOut({ status: "error", message: "Akaun tidak dijumpai." });

if (!safeEquals(hashPassword(oldPwd, user.salt), user.passwordHash)) {
auditLog(session.id, "TUKAR_PWD_GAGAL", "Kata laluan lama salah");
return jsonOut({ status: "error", message: "Kata laluan semasa tidak tepat." });
}
if (safeEquals(hashPassword(newPwd, user.salt), user.passwordHash)) {
return jsonOut({ status: "error", message: "Kata laluan baharu tidak boleh sama dengan yang lama." });
}

const newSalt = generateSalt();
sheet.getRange(user.rowIndex, 3).setValue(hashPassword(newPwd, newSalt));
sheet.getRange(user.rowIndex, 6).setValue(newSalt);
sheet.getRange(user.rowIndex, 7).setValue(new Date());

revokeAllSessionsFor(session.id);
const freshToken = createSession(
{ id: user.id, name: user.name, role: roleOf(user.designation) }, "session"
);

auditLog(session.id, "TUKAR_PWD", "Berjaya");
return jsonOut({ status: "success", token: freshToken });
}

// ==========================================================================
// SETUP & MIGRATION
// ==========================================================================

function setupSystem() {
getPepper();

const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
const sheet = getStaffSheet();

if (String(sheet.getRange(1, 6).getValue()).trim() === "") {
sheet.getRange(1, 6).setValue("SALT").setFontWeight("bold");
}
if (String(sheet.getRange(1, 7).getValue()).trim() === "") {
sheet.getRange(1, 7).setValue("PWD_UPDATED").setFontWeight("bold");
}

const rows = sheet.getDataRange().getValues();
let migrated = 0, skipped = 0;

for (let i = 1; i < rows.length; i++) {
const id = String(rows[i][0] || "").trim();
const current = String(rows[i][2] || "").trim();
const salt = String(rows[i][5] || "").trim();
if (!id) continue;

if (salt && /^[0-9a-f]{64}$/.test(current)) { skipped++; continue; }

const plain = current || ("temp" + id);
const newSalt = generateSalt();
sheet.getRange(i + 1, 3).setValue(hashPassword(plain, newSalt));
sheet.getRange(i + 1, 6).setValue(newSalt);
migrated++;


}

getSessionSheet();
if (!ss.getSheetByName(SHEET_AUDIT)) auditLog("SYSTEM", "SETUP", "Audit log dicipta");

const msg = "setupSystem selesai. Di-hash: " + migrated + ", dilangkau: " + skipped + ".";
console.log(msg);
return msg;
}

function resetPassword(userId, newPlainPassword) {
const sheet = getStaffSheet();
const user = readStaff().filter(function (s) { return s.id === String(userId).trim(); })[0];
if (!user) throw new Error("ID " + userId + " tidak dijumpai.");
const salt = generateSalt();
sheet.getRange(user.rowIndex, 3).setValue(hashPassword(newPlainPassword, salt));
sheet.getRange(user.rowIndex, 6).setValue(salt);
sheet.getRange(user.rowIndex, 7).setValue("");
revokeAllSessionsFor(user.id);
auditLog("HR", "RESET_PWD", userId);
return "Kata laluan " + userId + " telah direset.";
}

function enforceRetention() {
const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
const claims = getClaimsSheet();
const cutoff = new Date();
cutoff.setFullYear(cutoff.getFullYear() - RETENTION_YEARS_CLAIMS);

let archive = ss.getSheetByName("Arkib Tuntutan");
const rows = claims.getDataRange().getValues();
let moved = 0;

for (let i = rows.length - 1; i >= 1; i--) {
const d = new Date(rows[i][2]);
if (!isNaN(d.getTime()) && d < cutoff) {
if (!archive) {
archive = ss.insertSheet("Arkib Tuntutan");
archive.appendRow(rows[0]);
}
archive.appendRow(rows[i]);
claims.deleteRow(i + 1);
moved++;
}
}

const audit = ss.getSheetByName(SHEET_AUDIT);
if (audit) {
const aRows = audit.getDataRange().getValues();
const aCut = Date.now() - RETENTION_DAYS_AUDIT * 24 * 60 * 60 * 1000;
for (let i = aRows.length - 1; i >= 1; i--) {
const t = new Date(aRows[i][0]);
if (!isNaN(t.getTime()) && t.getTime() < aCut) audit.deleteRow(i + 1);
}
}

purgeExpiredSessions();
console.log("enforceRetention: " + moved + " tuntutan diarkibkan.");
}

function installTriggers() {
ScriptApp.getProjectTriggers().forEach(function (t) {
if (t.getHandlerFunction() === "purgeExpiredSessions" ||
t.getHandlerFunction() === "enforceRetention") {
ScriptApp.deleteTrigger(t);
}
});
ScriptApp.newTrigger("purgeExpiredSessions").timeBased().everyHours(6).create();
ScriptApp.newTrigger("enforceRetention").timeBased().onMonthDay(1).atHour(3).create();
return "Trigger housekeeping dipasang.";
}
