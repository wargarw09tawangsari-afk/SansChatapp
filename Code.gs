/**
 * ============================================================================
 * SPREADSHEET ID: 1kmZa08cDx5ts4Ld819mLsVmBJ83J-vZPu89i59TU3VI
 * Sc-SansChat Backend Script (Realtime Chat History, Penggolongan Kolom C & Warga Management)
 * ============================================================================
 */
const SPREADSHEET_ID = "1kmZa08cDx5ts4Ld819mLsVmBJ83J-vZPu89i59TU3VI";

function doGet(e) {
  ensureDefaultSheets();
  const params = (e && e.parameter) ? e.parameter : {};
  if (params.action === 'getConfig') {
    const config = {
      helpSystemUrl: getHelpSystemUrl(),
      portalWebUrl: getWebAppsUrl(),
      sansChatUrl: getSansChatUrl()
    };
    return ContentService.createTextOutput(JSON.stringify({ success: true, config: config }))
      .setMimeType(ContentService.MimeType.JSON);
  }
  const page = String(params.page || "").trim().toLowerCase();
  const userParam = String(params.user || "").trim().toUpperCase();
  if (page === 'login') {
    const tmpl = HtmlService.createTemplateFromFile('login');
    tmpl.s3Url = getSansChatUrl();
    tmpl.s1Url = getHelpSystemUrl();
    return tmpl.evaluate()
      .setTitle('Sc-SansChat - Access Portal')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }
  if (page === 'utama') {
    const tmpl = HtmlService.createTemplateFromFile('utama');
    tmpl.initialUsername = userParam;
    tmpl.s3Url = getSansChatUrl();
    return tmpl.evaluate()
      .setTitle('Sc-SansChat - Utama')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }
  const idxTmpl = HtmlService.createTemplateFromFile('index');
  idxTmpl.s3Url = getSansChatUrl();
  return idxTmpl.evaluate()
    .setTitle('Sc-SansChat')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function doPost(e) {
  try {
    const payload = JSON.parse(e.postData.contents);
    const action = payload.action;
    let result = null;
    switch (action) {
      case 'loginUser':
        result = loginUser(payload.username, payload.password);
        break;
      case 'getInitialDashboardBundle':
        result = getInitialDashboardBundle(payload.username);
        break;
      case 'getUserProfileDetail':
        result = getUserProfileDetail(payload.username);
        break;
      case 'getAllWargaContacts':
        result = getAllWargaContacts();
        break;
      case 'updateProfilePicture':
        result = updateProfilePicture(payload.username, payload.base64Data);
        break;
      case 'updateGroupProfilePicture':
        result = updateGroupProfilePicture(payload.sheetName, payload.base64Data);
        break;
      case 'getAnnouncements':
        result = getAnnouncements();
        break;
      case 'getUserAssignedGroups':
        result = getUserAssignedGroups(payload.username);
        break;
      case 'getNamaGrupMapping':
        result = getNamaGrupMapping();
        break;
      case 'resolveGroupRoom':
        result = resolveGroupRoom(payload.groupName, payload.sheetName);
        break;
      case 'getGroupMembersByRole':
      case 'getGroupMembersByKolomC':
        result = getGroupMembersByKolomC(payload.groupName, payload.chatSheetName);
        break;
      case 'sendMessage':
        result = sendMessage(payload.chatSheetName, payload.username, payload.nama, payload.pesan, payload.tipe, payload.mediaUrl, payload.replyTo, payload.groupName);
        break;
      case 'getMessages':
        result = getMessages(payload.chatSheetName, payload.username, payload.groupName);
        break;
      case 'votePollOption':
        result = votePollOption(payload.chatSheetName, payload.rowNumber, payload.optionindex, payload.voterUsername, payload.groupName);
        break;
      case 'sendPrivateMessage':
        result = sendPrivateMessage(payload.username, payload.nama, payload.pesan, payload.tipe, payload.mediaUrl, payload.targetUsername, payload.replyTo);
        break;
      case 'getPrivateChatSheetRecords':
        result = getPrivateChatSheetRecords(payload.username);
        break;
      case 'hasUnreadPrivateMessages':
        result = hasUnreadPrivateMessages(payload.username);
        break;
      case 'markPrivateMessagesAsRead':
        markPrivateMessagesAsRead(payload.username);
        result = { success: true };
        break;
      case 'initiateCall':
        result = initiateCall(payload.targetUsername, payload.callerUsername, payload.callerName, payload.type, payload.roomUrl);
        break;
      case 'checkIncomingCall':
        result = checkIncomingCall(payload.username);
        break;
      case 'updateCallStatus':
        updateCallStatus(payload.username, payload.status);
        result = { success: true };
        break;
      default:
        result = { success: false, message: "Action tidak dikenali: " + action };
    }
    return ContentService.createTextOutput(JSON.stringify({ status: "ok", data: result }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ status: "error", message: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function getSpreadsheet() {
  return SpreadsheetApp.openById(SPREADSHEET_ID);
}

function columnLetterToIndex(colRef) {
  const clean = String(colRef || "").trim().toUpperCase();
  if (!clean) return -1;
  if (/^\d+$/.test(clean)) return parseInt(clean, 10);
  if (/^[A-Z]+$/.test(clean)) {
    let colIdx = 0;
    for (let i = 0; i < clean.length; i++) {
      colIdx = colIdx * 26 + (clean.charCodeAt(i) - 64);
    }
    return colIdx;
  }
  return -1;
}

function ensureDefaultSheets() {
  const ss = getSpreadsheet();
  if (!ss.getSheetByName("Users")) {
    const u = ss.insertSheet("Users");
    u.appendRow(["Username", "Nama", "Alamat", "Password", "Foto", "Role", "Grup RT37", "Grup RT38", "Grup RT39", "Grup RW09", "Grup Umum"]);
  }
  let ng = ss.getSheetByName("NamaGrup");
  if (!ng) {
    ng = ss.insertSheet("NamaGrup");
    ng.appendRow(["Nama Grup", "Sheet Chat", "Kolom", "Foto Grup", "Deskripsi"]);
    ng.appendRow(["Grup RT37", "ChatRT37", "G", "", "Grup komunikasi warga RT 37 RW 09"]);
    ng.appendRow(["Grup RT38", "ChatRT38", "H", "", "Grup komunikasi warga RT 38 RW 09"]);
    ng.appendRow(["Grup RT39", "ChatRT39", "I", "", "Grup komunikasi warga RT 39 RW 09"]);
    ng.appendRow(["Grup RW09", "ChatRW09", "J", "", "Grup komunikasi seluruh warga dan pengurus RW 09"]);
    ng.appendRow(["Grup Umum", "ChatUmum", "K", "", "Grup obrolan dan koordinasi umum seluruh warga SansChat RW 09"]);
  }
  if (!ss.getSheetByName("Informasi")) {
    const inf = ss.insertSheet("Informasi");
    inf.appendRow(["Isi Pengumuman", "Prioritas", "Tanggal"]);
    inf.appendRow(["Diberitahukan kepada seluruh warga RT 37, 38, dan 39 untuk hadir dalam kerja bakti kebersihan lingkungan pada hari Minggu pukul 07.00 WIB di balai warga.", "TINGGI", new Date()]);
  }
  if (!ss.getSheetByName("ChatPrivat")) {
    const cp = ss.insertSheet("ChatPrivat");
    cp.appendRow(["Timestamp", "Username", "Nama", "Pesan", "Tipe", "MediaUrl", "StatusPesan", "TujuanUsername", "ReplyTo"]);
  }
  if (!ss.getSheetByName("PanggilanAktif")) {
    const pa = ss.insertSheet("PanggilanAktif");
    pa.appendRow(["TargetUsername", "CallerUsername", "CallerName", "Type", "Status", "RoomUrl", "Timestamp"]);
  }
  if (!ss.getSheetByName("LogAktifitas")) {
    const logS = ss.insertSheet("LogAktifitas");
    logS.appendRow(["Timestamp", "Username", "Aktivitas", "Detail"]);
  }
  if (!ss.getSheetByName("SesiAktif")) {
    const sesiS = ss.insertSheet("SesiAktif");
    sesiS.appendRow(["Username", "Nama", "Role", "DeviceID", "WaktuLogin", "StatusSesi", "LastActive"]);
  }
  if (!ss.getSheetByName("NotifikasiAntrean")) {
    const notifS = ss.insertSheet("NotifikasiAntrean");
    notifS.appendRow(["NotifID", "TargetUsername", "SenderUsername", "SenderNama", "RoomName", "SheetSource", "TipeNotif", "RingkasanPesan", "StatusBaca", "Timestamp"]);
  }
}

function getHelpSystemUrl() {
  const sheet = getSpreadsheet().getSheetByName("Users");
  return sheet ? String(sheet.getRange("S1").getValue() || "").trim() : "";
}
function getWebAppsUrl() {
  const sheet = getSpreadsheet().getSheetByName("Users");
  return sheet ? String(sheet.getRange("S2").getValue() || "").trim() : "";
}
function getSansChatUrl() {
  const sheet = getSpreadsheet().getSheetByName("Users");
  return sheet ? String(sheet.getRange("S3").getValue() || "").trim() : "";
}
function getScriptUrl() {
  const s3 = getSansChatUrl();
  return (s3 && s3.startsWith("http")) ? s3 : ScriptApp.getService().getUrl();
}
function loginUser(username, password) {
  const ss = getSpreadsheet();
  const sheet = ss.getSheetByName("Users");
  if (!sheet || sheet.getLastRow() < 2) return { success: false, message: "Data Users masih kosong" };
  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 6).getValues();
  const cleanUser = String(username || "").trim().toUpperCase();
  const cleanPass = String(password || "").trim();
  for (let i = 0; i < data.length; i++) {
    const rowUser = String(data[i][0] || "").trim().toUpperCase();
    const rowPass = String(data[i][3] || "").trim();
    if (rowUser === cleanUser && rowPass === cleanPass) {
      return {
        success: true,
        redirectBaseUrl: getScriptUrl(),
        user: {
          username: rowUser,
          nama: String(data[i][1] || ""),
          alamat: String(data[i][2] || ""),
          foto: String(data[i][4] || ""),
          role: String(data[i][5] || ""),
          helpSystemUrl: getHelpSystemUrl(),
          webAppsUrl: getWebAppsUrl(),
          sansChatUrl: getScriptUrl()
        }
      };
    }
  }
  return { success: false, message: "Username atau Password Salah!" };
}
function getInitialDashboardBundle(username) {
  const cleanUser = String(username || "").trim().toUpperCase();
  return {
    profile: getUserProfileDetail(cleanUser),
    groups: getUserAssignedGroups(cleanUser),
    groupMapping: getNamaGrupMapping(),
    announcements: getAnnouncements()
  };
}
function getUserProfileDetail(username) {
  const sheet = getSpreadsheet().getSheetByName("Users");
  if (!sheet || sheet.getLastRow() < 2) return null;
  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 6).getValues();
  const cleanUser = String(username || "").trim().toUpperCase();
  for (let i = 0; i < data.length; i++) {
    if (String(data[i][0] || "").trim().toUpperCase() === cleanUser) {
      return {
        username: String(data[i][0] || "").trim().toUpperCase(),
        nama: String(data[i][1] || ""),
        alamat: String(data[i][2] || ""),
        foto: String(data[i][4] || ""),
        role: String(data[i][5] || ""),
        helpSystemUrl: getHelpSystemUrl(),
        webAppsUrl: getWebAppsUrl(),
        sansChatUrl: getScriptUrl()
      };
    }
  }
  return null;
}
function getAnnouncements() {
  const sheet = getSpreadsheet().getSheetByName("Informasi");
  if (!sheet || sheet.getLastRow() < 2) return [];
  return sheet.getRange(2, 2, sheet.getLastRow() - 1, 1).getValues().map(r => String(r[0] || "").trim()).filter(Boolean);
}
function getNamaGrupMapping() {
  const sheet = getSpreadsheet().getSheetByName("NamaGrup");
  if (!sheet || sheet.getLastRow() < 2) return [];
  const rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, 5).getValues();
  return rows.map(r => ({
    name: String(r[0] || "").trim(),
    sheetName: String(r[1] || "").trim(),
    kolom: String(r[2] || "").trim(),
    foto: String(r[3] || "").trim(),
    deskripsi: String(r[4] || "").trim()
  })).filter(g => g.name);
}
function getUserAssignedGroups(username) {
  const ss = getSpreadsheet();
  const userSheet = ss.getSheetByName("Users");
  const groupsMeta = getNamaGrupMapping();
  if (!userSheet || userSheet.getLastRow() < 2) return [];
  const cleanUser = String(username || "").trim().toUpperCase();
  const lastCol = Math.max(userSheet.getLastColumn(), 12);
  const rows = userSheet.getRange(2, 1, userSheet.getLastRow() - 1, lastCol).getValues();
  let matchedRow = null;
  for (let i = 0; i < rows.length; i++) {
    if (String(rows[i][0] || "").trim().toUpperCase() === cleanUser) { matchedRow = rows[i]; break; }
  }
  if (!matchedRow) return [];
  const assigned = [];
  groupsMeta.forEach(g => {
    const colIdx = columnLetterToIndex(g.kolom);
    const val = colIdx > 0 ? String(matchedRow[colIdx - 1] || "").trim().toUpperCase() : "AKTIF";
    if (val && val !== "TIDAK" && val !== "FALSE" && val !== "0") {
      assigned.push({ name: g.name, sheetName: g.sheetName || g.name.replace(/\s+/g, ""), foto: g.foto, lastMessage: "", lastTime: "" });
    }
  });
  return assigned;
}
function getGroupMembersByKolomC(groupName, chatSheetName) {
  const ss = getSpreadsheet();
  const ng = ss.getSheetByName("NamaGrup");
  const uSheet = ss.getSheetByName("Users");
  if (!ng || !uSheet || uSheet.getLastRow() < 2) return [];
  const ngRows = ng.getRange(2, 1, ng.getLastRow() - 1, 3).getValues();
  let targetCol = "";
  for (let i = 0; i < ngRows.length; i++) {
    if (String(ngRows[i][0] || "").trim().toLowerCase() === String(groupName || "").trim().toLowerCase() ||
        String(ngRows[i][1] || "").trim().toLowerCase() === String(chatSheetName || "").trim().toLowerCase()) {
      targetCol = String(ngRows[i][2] || "").trim();
      break;
    }
  }
  const colIdx = columnLetterToIndex(targetCol);
  if (colIdx < 1) return [];
  const uRows = uSheet.getRange(2, 1, uSheet.getLastRow() - 1, Math.max(uSheet.getLastColumn(), colIdx)).getValues();
  const members = [];
  uRows.forEach(r => {
    const flag = String(r[colIdx - 1] || "").trim().toUpperCase();
    if (flag && flag !== "TIDAK" && flag !== "FALSE" && flag !== "0") {
      members.push({ username: String(r[0] || "").trim(), nama: String(r[1] || "").trim(), role: String(r[5] || "").trim() });
    }
  });
  return members;
}
function sendMessage(chatSheetName, username, nama, pesan, tipe, mediaUrl, replyTo, groupName) {
  const ss = getSpreadsheet();
  const sName = String(chatSheetName || groupName || "ChatUmum").trim();
  let sh = ss.getSheetByName(sName);
  if (!sh) {
    sh = ss.insertSheet(sName);
    sh.appendRow(["Timestamp", "SenderUsername", "SenderNama", "Pesan", "Tipe", "MediaUrl", "ReplyTo", "ReadBy"]);
  }
  sh.appendRow([new Date(), String(username || "").toUpperCase(), nama || username, pesan || "", tipe || "text", mediaUrl || "", replyTo || "", ""]);
  return { success: true };
}
function getMessages(chatSheetName, username, groupName) {
  const ss = getSpreadsheet();
  const sName = String(chatSheetName || groupName || "ChatUmum").trim();
  const sh = ss.getSheetByName(sName);
  if (!sh || sh.getLastRow() < 2) return [];
  const rows = sh.getRange(2, 1, sh.getLastRow() - 1, Math.min(sh.getLastColumn(), 8)).getValues();
  return rows.map(r => ({
    timestamp: r[0] instanceof Date ? Utilities.formatDate(r[0], "Asia/Jakarta", "yyyy-MM-dd HH:mm:ss") : String(r[0] || ""),
    senderUsername: String(r[1] || ""),
    senderNama: String(r[2] || ""),
    pesan: String(r[3] || ""),
    tipe: String(r[4] || "text"),
    mediaUrl: String(r[5] || "")
  }));
}
function sendPrivateMessage(username, nama, pesan, tipe, mediaUrl, targetUsername, replyTo) {
  const ss = getSpreadsheet();
  let sh = ss.getSheetByName("ChatPrivat");
  if (!sh) {
    sh = ss.insertSheet("ChatPrivat");
    sh.appendRow(["Timestamp", "Username", "Nama", "Pesan", "Tipe", "MediaUrl", "StatusPesan", "TujuanUsername", "ReplyTo"]);
  }
  sh.appendRow([new Date(), String(username || "").toUpperCase(), nama || username, pesan || "", tipe || "text", mediaUrl || "", "UNREAD", String(targetUsername || "").toUpperCase(), replyTo || ""]);
  return { success: true };
}
function getPrivateChatSheetRecords(username) {
  const ss = getSpreadsheet();
  const sh = ss.getSheetByName("ChatPrivat");
  if (!sh || sh.getLastRow() < 2) return [];
  const cleanUser = String(username || "").trim().toUpperCase();
  const rows = sh.getRange(2, 1, sh.getLastRow() - 1, Math.min(sh.getLastColumn(), 9)).getValues();
  return rows.filter(r => {
    const u = String(r[1] || "").trim().toUpperCase();
    const t = String(r[7] || "").trim().toUpperCase();
    return u === cleanUser || t === cleanUser || cleanUser === "ADMINPUSAT";
  }).map(r => ({
    timestamp: r[0] instanceof Date ? Utilities.formatDate(r[0], "Asia/Jakarta", "yyyy-MM-dd HH:mm:ss") : String(r[0] || ""),
    pengirimUsername: String(r[1] || ""),
    nama: String(r[2] || ""),
    pesanD: String(r[3] || ""),
    pesanF: String(r[3] || "")
  }));
}
function markPrivateMessagesAsRead(username) {
  const ss = getSpreadsheet();
  const sh = ss.getSheetByName("ChatPrivat");
  if (!sh || sh.getLastRow() < 2) return;
  const cleanUser = String(username || "").trim().toUpperCase();
  const rows = sh.getRange(2, 1, sh.getLastRow() - 1, 8).getValues();
  for (let i = 0; i < rows.length; i++) {
    if (String(rows[i][7] || "").trim().toUpperCase() === cleanUser) {
      sh.getRange(i + 2, 7).setValue("READ");
    }
  }
}
