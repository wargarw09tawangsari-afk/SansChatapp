/**
 * ============================================================================
 * SPREADSHEET ID: 1kmZa08cDx5ts4Ld819mLsVmBJ83J-vZPu89i59TU3VI
 * Sc-SansChat Backend Script (Realtime Chat History, Penggolongan Kolom C,
 * Informasi Kolom B + Warna, SesiAktif, NotifikasiAntrean, LogAktifitas & Warga Management)
 * ============================================================================
 */
const SPREADSHEET_ID = "1kmZa08cDx5ts4Ld819mLsVmBJ83J-vZPu89i59TU3VI";
const TIMEZONE_WIB = "Asia/Jakarta";

function doGet(e) {
  ensureDefaultSheets();
  const params = (e && e.parameter) ? e.parameter : {};

  if (params.action === 'getConfig') {
    const config = {
      spreadsheetId: SPREADSHEET_ID,
      helpSystemUrl: getHelpSystemUrl(),
      portalWebUrl: getWebAppsUrl(),
      sansChatUrl: getSansChatUrl(),
      activeSheets: getActiveSheetsSummary()
    };
    return ContentService.createTextOutput(JSON.stringify({ success: true, config: config }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  if (params.action === 'getAnnouncements') {
    const items = getAnnouncements();
    return ContentService.createTextOutput(JSON.stringify({ status: "ok", data: items }))
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
    ensureDefaultSheets();
    const payload = JSON.parse(e.postData.contents);
    const action = payload.action;
    let result = null;

    switch (action) {
      case 'loginUser':
        result = loginUser(payload.username, payload.password, payload.deviceInfo);
        break;
      case 'logoutUser':
        result = logoutUserSession(payload.username, payload.detail);
        break;
      case 'heartbeatSession':
      case 'updateSessionActivity':
        result = heartbeatUserSession(payload.username, payload.statusSesi, payload.halaman, payload.deviceInfo);
        break;
      case 'logActivity':
        appendActivityLog(payload.username, payload.aktivitas, payload.detail, payload.halaman);
        result = { success: true };
        break;
      case 'getInitialDashboardBundle':
        result = getInitialDashboardBundle(payload.username, payload.halaman);
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
      case 'getActiveSessions':
        result = getActiveSessions();
        break;
      case 'getPendingNotifications':
        result = getPendingNotifications(payload.username);
        break;
      case 'markNotificationRead':
        result = markNotificationRead(payload.notifId, payload.username);
        break;
      case 'getUserAssignedGroups':
        result = getUserAssignedGroups(payload.username);
        break;
      case 'getNamaGrupMapping':
        result = getNamaGrupMapping();
        break;
      case 'resolveGroupRoom':
        result = resolveGroupRoom(payload.groupName, payload.sheetName, payload.username);
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

function formatWibTimestamp(dateObj) {
  const d = (dateObj instanceof Date) ? dateObj : new Date();
  return Utilities.formatDate(d, TIMEZONE_WIB, "yyyy-MM-dd HH:mm:ss 'WIB'");
}

function getActiveSheetsSummary() {
  const ss = getSpreadsheet();
  return ss.getSheets().map(sh => ({
    name: sh.getName(),
    isHidden: sh.isSheetHidden(),
    lastRow: sh.getLastRow(),
    lastColumn: sh.getLastColumn()
  }));
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

/**
 * Mencari indeks kolom (1-based) berdasarkan daftar nama header yang cocok
 * pada baris pertama sheet agar kompatibel dengan pengaturan kolom sheet yang sudah ada.
 */
function findColumnIndexByNames(headers, candidateNames, fallbackColIndex) {
  if (Array.isArray(headers) && headers.length > 0 && Array.isArray(candidateNames)) {
    const normalizedCandidates = candidateNames.map(n => String(n || "").trim().toLowerCase().replace(/[\s_-]+/g, ""));
    for (let c = 0; c < headers.length; c++) {
      const hNorm = String(headers[c] || "").trim().toLowerCase().replace(/[\s_-]+/g, "");
      if (!hNorm) continue;
      if (normalizedCandidates.includes(hNorm)) {
        return c + 1;
      }
    }
    for (let c = 0; c < headers.length; c++) {
      const hNorm = String(headers[c] || "").trim().toLowerCase().replace(/[\s_-]+/g, "");
      if (!hNorm) continue;
      for (let k = 0; k < normalizedCandidates.length; k++) {
        if (hNorm.includes(normalizedCandidates[k]) || normalizedCandidates[k].includes(hNorm)) {
          return c + 1;
        }
      }
    }
  }
  return fallbackColIndex || -1;
}

function getSheetHeaderAndRows(sheet, minCols) {
  const lastRow = sheet.getLastRow();
  const lastCol = Math.max(minCols || 1, sheet.getLastColumn());
  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  const rows = lastRow >= 2 ? sheet.getRange(2, 1, lastRow - 1, lastCol).getValues() : [];
  return { lastRow: lastRow, lastCol: lastCol, headers: headers, rows: rows };
}

/**
 * Memastikan seluruh sheet utama & sheet chat grup aktif (showSheet),
 * memiliki header lengkap, dan tersinkron dengan struktur aplikasi.
 */
function ensureDefaultSheets() {
  const ss = getSpreadsheet();

  let u = ss.getSheetByName("Users");
  if (!u) {
    u = ss.insertSheet("Users");
    u.appendRow(["Username", "Nama", "Alamat", "Password", "Foto", "Role", "Grup RT37", "Grup RT38", "Grup RT39", "Grup RW09", "Grup Umum"]);
  }
  if (u.isSheetHidden()) u.showSheet();

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
  if (ng.isSheetHidden()) ng.showSheet();

  if (ng.getLastRow() >= 2) {
    const rows = ng.getRange(2, 1, ng.getLastRow() - 1, 2).getValues();
    for (let i = 0; i < rows.length; i++) {
      const colA = String(rows[i][0] || "").trim();
      let colB = String(rows[i][1] || "").trim();
      if (colA && !colB) {
        colB = colA.replace(/^Grup\s*/i, "Chat").replace(/\s+/g, "");
        ng.getRange(i + 2, 2).setValue(colB);
      }
      if (colA && colB) {
        let chatSh = findExactSheetCaseInsensitive(ss, colB);
        if (!chatSh) {
          chatSh = ss.insertSheet(colB);
          chatSh.appendRow(["Timestamp", "SenderUsername", "SenderNama", "Pesan", "Tipe", "MediaUrl", "ReplyTo", "ReadBy"]);
        }
        if (chatSh.isSheetHidden()) chatSh.showSheet();
      }
    }
  }

  // 1. Sheet Informasi:
  //    Kolom A = Judul / Kategori Pengumuman
  //    Kolom B = Isi Pengumuman (Utama ditampilkan di dashboard)
  //    Kolom C = Prioritas (TINGGI / SEDANG / RENDAH / PENTING / DARURAT)
  //    Kolom D = Warna Teks (Hex seperti #fde047 atau nama warna)
  //    Kolom E = Warna Latar / Border (Opsional)
  //    Kolom F = Tanggal
  let inf = findExactSheetCaseInsensitive(ss, "Informasi");
  if (!inf) {
    inf = ss.insertSheet("Informasi");
    inf.appendRow(["Judul Informasi", "Isi Pengumuman", "Prioritas", "Warna Teks", "Warna Latar", "Tanggal"]);
    inf.appendRow([
      "Kerja Bakti Lingkungan",
      "Diberitahukan kepada seluruh warga RT 37, 38, dan 39 untuk hadir dalam kerja bakti kebersihan lingkungan pada hari Minggu pukul 07.00 WIB di balai warga.",
      "TINGGI",
      "#fde047",
      "#182229",
      new Date()
    ]);
    inf.appendRow([
      "Layanan Digital RW 09",
      "Gunakan menu Chat Privat untuk koordinasi langsung dengan pengurus RT/RW dan pantau pengumuman resmi RW 09 Cluster Sansiviera.",
      "SEDANG",
      "#6ee7b7",
      "#142621",
      new Date()
    ]);
  }
  if (inf.isSheetHidden()) inf.showSheet();

  // 2. Sheet ChatPrivat
  let cp = ss.getSheetByName("ChatPrivat");
  if (!cp) {
    cp = ss.insertSheet("ChatPrivat");
    cp.appendRow(["Timestamp", "Username", "Nama", "Pesan", "Tipe", "MediaUrl", "StatusPesan", "TujuanUsername", "ReplyTo"]);
  }
  if (cp.isSheetHidden()) cp.showSheet();

  // 3. Sheet PanggilanAktif
  let pa = ss.getSheetByName("PanggilanAktif");
  if (!pa) {
    pa = ss.insertSheet("PanggilanAktif");
    pa.appendRow(["TargetUsername", "CallerUsername", "CallerName", "Type", "Status", "RoomUrl", "Timestamp"]);
  }
  if (pa.isSheetHidden()) pa.showSheet();

  // 4. Sheet LogAktifitas
  let logS = findExactSheetCaseInsensitive(ss, "LogAktifitas") || findExactSheetCaseInsensitive(ss, "LogAktivitas");
  if (!logS) {
    logS = ss.insertSheet("LogAktifitas");
    logS.appendRow(["Timestamp", "Username", "Aktivitas", "Detail", "Halaman"]);
  }
  if (logS.isSheetHidden()) logS.showSheet();

  // 5. Sheet SesiAktif
  let sesiS = findExactSheetCaseInsensitive(ss, "SesiAktif");
  if (!sesiS) {
    sesiS = ss.insertSheet("SesiAktif");
    sesiS.appendRow(["Username", "Nama", "Role", "DeviceID", "WaktuLogin", "StatusSesi", "LastActive", "Halaman"]);
  }
  if (sesiS.isSheetHidden()) sesiS.showSheet();

  // 6. Sheet NotifikasiAntrean
  let notifS = findExactSheetCaseInsensitive(ss, "NotifikasiAntrean");
  if (!notifS) {
    notifS = ss.insertSheet("NotifikasiAntrean");
    notifS.appendRow(["NotifID", "TargetUsername", "SenderUsername", "SenderNama", "RoomName", "SheetSource", "TipeNotif", "RingkasanPesan", "StatusBaca", "Timestamp"]);
  }
  if (notifS.isSheetHidden()) notifS.showSheet();
}

function findExactSheetCaseInsensitive(ss, targetName) {
  const clean = String(targetName || "").trim();
  if (!clean) return null;
  const direct = ss.getSheetByName(clean);
  if (direct) return direct;
  const sheets = ss.getSheets();
  for (let i = 0; i < sheets.length; i++) {
    if (sheets[i].getName().trim().toUpperCase() === clean.toUpperCase()) {
      return sheets[i];
    }
  }
  return null;
}

function getLogSheet(ss) {
  let logSheet = findExactSheetCaseInsensitive(ss, "LogAktifitas") || findExactSheetCaseInsensitive(ss, "LogAktivitas");
  if (!logSheet) {
    logSheet = ss.insertSheet("LogAktifitas");
    logSheet.appendRow(["Timestamp", "Username", "Aktivitas", "Detail", "Halaman"]);
  }
  if (logSheet.isSheetHidden()) logSheet.showSheet();
  return logSheet;
}

function appendActivityLog(username, aktivitas, detail, halaman) {
  try {
    const ss = getSpreadsheet();
    const logSheet = getLogSheet(ss);
    const cleanUser = String(username || "SYSTEM").trim().toUpperCase();
    const now = new Date();
    const meta = getSheetHeaderAndRows(logSheet, 4);

    if (meta.lastRow === 0) {
      logSheet.appendRow(["Timestamp", "Username", "Aktivitas", "Detail", "Halaman"]);
    }

    const refreshed = getSheetHeaderAndRows(logSheet, 4);
    const colTime = findColumnIndexByNames(refreshed.headers, ["timestamp", "waktu", "tanggal", "time"], 1);
    const colUser = findColumnIndexByNames(refreshed.headers, ["username", "user", "idwarga", "pengguna"], 2);
    const colAct = findColumnIndexByNames(refreshed.headers, ["aktivitas", "activity", "aksi", "action", "event"], 3);
    const colDetail = findColumnIndexByNames(refreshed.headers, ["detail", "keterangan", "info", "deskripsi", "catatan"], 4);
    const colPage = findColumnIndexByNames(refreshed.headers, ["halaman", "page", "menu", "room", "sheet"], -1);

    const newRow = new Array(refreshed.lastCol).fill("");
    if (colTime >= 1) newRow[colTime - 1] = now;
    if (colUser >= 1) newRow[colUser - 1] = cleanUser;
    if (colAct >= 1) newRow[colAct - 1] = String(aktivitas || "");
    if (colDetail >= 1) newRow[colDetail - 1] = String(detail || "") + (halaman ? (" [" + halaman + "]") : "");
    if (colPage >= 1 && halaman) newRow[colPage - 1] = String(halaman || "");

    logSheet.appendRow(newRow);
  } catch (e) {}
}

function getSesiSheet(ss) {
  let sesiSheet = findExactSheetCaseInsensitive(ss, "SesiAktif");
  if (!sesiSheet) {
    sesiSheet = ss.insertSheet("SesiAktif");
    sesiSheet.appendRow(["Username", "Nama", "Role", "DeviceID", "WaktuLogin", "StatusSesi", "LastActive", "Halaman"]);
  }
  if (sesiSheet.isSheetHidden()) sesiSheet.showSheet();
  return sesiSheet;
}

function updateUserActiveSession(username, nama, role, statusSesi, isNewLogin, halaman, deviceInfo) {
  try {
    const ss = getSpreadsheet();
    const sesiSheet = getSesiSheet(ss);
    const cleanUser = String(username || "").trim().toUpperCase();
    if (!cleanUser) return;

    const now = new Date();
    if (sesiSheet.getLastRow() === 0) {
      sesiSheet.appendRow(["Username", "Nama", "Role", "DeviceID", "WaktuLogin", "StatusSesi", "LastActive", "Halaman"]);
    }

    const meta = getSheetHeaderAndRows(sesiSheet, 7);
    const colUser = findColumnIndexByNames(meta.headers, ["username", "user", "idwarga", "pengguna"], 1);
    const colNama = findColumnIndexByNames(meta.headers, ["nama", "name", "namawarga"], 2);
    const colRole = findColumnIndexByNames(meta.headers, ["role", "peran", "jabatan", "grup"], 3);
    const colDev = findColumnIndexByNames(meta.headers, ["deviceid", "device", "perangkat", "devicename"], 4);
    const colLogin = findColumnIndexByNames(meta.headers, ["waktulogin", "login", "waktumasuk", "logintime"], 5);
    const colStatus = findColumnIndexByNames(meta.headers, ["statussesi", "status", "online", "state"], 6);
    const colLast = findColumnIndexByNames(meta.headers, ["lastactive", "terakhiraktif", "lastseen", "update", "timestamp", "waktu"], 7);
    const colPage = findColumnIndexByNames(meta.headers, ["halaman", "page", "menu", "room", "aktivitas"], -1);

    for (let i = 0; i < meta.rows.length; i++) {
      const rowUser = String(meta.rows[i][colUser - 1] || "").trim().toUpperCase();
      if (rowUser === cleanUser) {
        const rowNum = i + 2;
        if (nama && colNama >= 1) sesiSheet.getRange(rowNum, colNama).setValue(nama);
        if (role && colRole >= 1) sesiSheet.getRange(rowNum, colRole).setValue(role);
        if (deviceInfo && colDev >= 1) sesiSheet.getRange(rowNum, colDev).setValue(deviceInfo);
        if (isNewLogin && colLogin >= 1) sesiSheet.getRange(rowNum, colLogin).setValue(now);
        if (colStatus >= 1) sesiSheet.getRange(rowNum, colStatus).setValue(statusSesi || "ONLINE");
        if (colLast >= 1) sesiSheet.getRange(rowNum, colLast).setValue(now);
        if (halaman && colPage >= 1) sesiSheet.getRange(rowNum, colPage).setValue(halaman);
        return;
      }
    }

    const newRow = new Array(meta.lastCol).fill("");
    if (colUser >= 1) newRow[colUser - 1] = cleanUser;
    if (colNama >= 1) newRow[colNama - 1] = nama || cleanUser;
    if (colRole >= 1) newRow[colRole - 1] = role || "";
    if (colDev >= 1) newRow[colDev - 1] = deviceInfo || ("PWA-" + cleanUser);
    if (colLogin >= 1) newRow[colLogin - 1] = now;
    if (colStatus >= 1) newRow[colStatus - 1] = statusSesi || "ONLINE";
    if (colLast >= 1) newRow[colLast - 1] = now;
    if (colPage >= 1 && halaman) newRow[colPage - 1] = String(halaman || "");
    sesiSheet.appendRow(newRow);
  } catch (e) {}
}

function heartbeatUserSession(username, statusSesi, halaman, deviceInfo) {
  const cleanUser = String(username || "").trim().toUpperCase();
  if (!cleanUser) return { success: false };
  const profile = getUserProfileDetail(cleanUser);
  updateUserActiveSession(
    cleanUser,
    profile ? profile.nama : cleanUser,
    profile ? profile.role : "",
    statusSesi || "ONLINE",
    false,
    halaman || "",
    deviceInfo || ""
  );
  if (halaman) {
    appendActivityLog(cleanUser, "PAGE_ACTIVE", "Aktif pada halaman: " + halaman, halaman);
  }
  return { success: true };
}

function logoutUserSession(username, detail) {
  const cleanUser = String(username || "").trim().toUpperCase();
  if (!cleanUser) return { success: false };
  updateUserActiveSession(cleanUser, "", "", "OFFLINE", false, "Login", "");
  appendActivityLog(cleanUser, "LOGOUT", detail || "Sesi diakhiri dari perangkat", "Logout");
  return { success: true };
}

function getActiveSessions() {
  const ss = getSpreadsheet();
  const sesiSheet = getSesiSheet(ss);
  const meta = getSheetHeaderAndRows(sesiSheet, 7);
  if (meta.rows.length === 0) return [];

  const colUser = findColumnIndexByNames(meta.headers, ["username", "user", "idwarga", "pengguna"], 1);
  const colNama = findColumnIndexByNames(meta.headers, ["nama", "name", "namawarga"], 2);
  const colRole = findColumnIndexByNames(meta.headers, ["role", "peran", "jabatan", "grup"], 3);
  const colDev = findColumnIndexByNames(meta.headers, ["deviceid", "device", "perangkat", "devicename"], 4);
  const colLogin = findColumnIndexByNames(meta.headers, ["waktulogin", "login", "waktumasuk", "logintime"], 5);
  const colStatus = findColumnIndexByNames(meta.headers, ["statussesi", "status", "online", "state"], 6);
  const colLast = findColumnIndexByNames(meta.headers, ["lastactive", "terakhiraktif", "lastseen", "update", "timestamp", "waktu"], 7);
  const colPage = findColumnIndexByNames(meta.headers, ["halaman", "page", "menu", "room", "aktivitas"], -1);

  return meta.rows.map(r => {
    const rawLogin = colLogin >= 1 ? r[colLogin - 1] : "";
    const rawLast = colLast >= 1 ? r[colLast - 1] : "";
    return {
      username: String((colUser >= 1 ? r[colUser - 1] : "") || "").trim().toUpperCase(),
      nama: String((colNama >= 1 ? r[colNama - 1] : "") || "").trim(),
      role: String((colRole >= 1 ? r[colRole - 1] : "") || "").trim(),
      deviceId: String((colDev >= 1 ? r[colDev - 1] : "") || "").trim(),
      waktuLogin: (rawLogin instanceof Date) ? rawLogin.toISOString() : String(rawLogin || ""),
      statusSesi: String((colStatus >= 1 ? r[colStatus - 1] : "OFFLINE") || "OFFLINE").trim().toUpperCase(),
      lastActive: (rawLast instanceof Date) ? rawLast.toISOString() : String(rawLast || ""),
      halaman: colPage >= 1 ? String(r[colPage - 1] || "").trim() : ""
    };
  }).filter(item => item.username);
}

function getNotifSheet(ss) {
  let notifSheet = findExactSheetCaseInsensitive(ss, "NotifikasiAntrean");
  if (!notifSheet) {
    notifSheet = ss.insertSheet("NotifikasiAntrean");
    notifSheet.appendRow(["NotifID", "TargetUsername", "SenderUsername", "SenderNama", "RoomName", "SheetSource", "TipeNotif", "RingkasanPesan", "StatusBaca", "Timestamp"]);
  }
  if (notifSheet.isSheetHidden()) notifSheet.showSheet();
  return notifSheet;
}

function enqueueNotification(targetUsername, senderUsername, senderNama, roomName, sheetSource, tipeNotif, ringkasanPesan) {
  try {
    const ss = getSpreadsheet();
    const notifSheet = getNotifSheet(ss);
    if (notifSheet.getLastRow() === 0) {
      notifSheet.appendRow(["NotifID", "TargetUsername", "SenderUsername", "SenderNama", "RoomName", "SheetSource", "TipeNotif", "RingkasanPesan", "StatusBaca", "Timestamp"]);
    }

    const meta = getSheetHeaderAndRows(notifSheet, 10);
    const colId = findColumnIndexByNames(meta.headers, ["notifid", "id", "idnotif", "notificationid"], 1);
    const colTarget = findColumnIndexByNames(meta.headers, ["targetusername", "target", "tujuanusername", "tujuan", "penerima", "username"], 2);
    const colSenderU = findColumnIndexByNames(meta.headers, ["senderusername", "pengirimusername", "pengirimid", "dari"], 3);
    const colSenderN = findColumnIndexByNames(meta.headers, ["sendernama", "pengirim", "namapengirim", "nama"], 4);
    const colRoom = findColumnIndexByNames(meta.headers, ["roomname", "room", "grup", "namagrup", "topik"], 5);
    const colSheet = findColumnIndexByNames(meta.headers, ["sheetsource", "sheet", "namasheet", "sumber"], 6);
    const colType = findColumnIndexByNames(meta.headers, ["tipenotif", "tipe", "jenis", "type", "kategori"], 7);
    const colMsg = findColumnIndexByNames(meta.headers, ["ringkasanpesan", "pesan", "isi", "ringkasan", "message", "teks"], 8);
    const colStatus = findColumnIndexByNames(meta.headers, ["statusbaca", "status", "dibaca", "readstatus", "state"], 9);
    const colTime = findColumnIndexByNames(meta.headers, ["timestamp", "waktu", "tanggal", "time"], 10);

    const cleanTarget = String(targetUsername || "ALL").trim().toUpperCase();
    const cleanSender = String(senderUsername || "SYSTEM").trim().toUpperCase();
    const notifId = "NTF-" + new Date().getTime() + "-" + Math.floor(Math.random() * 900 + 100);
    const shortMsg = String(ringkasanPesan || "").substring(0, 180);
    const now = new Date();

    const newRow = new Array(meta.lastCol).fill("");
    if (colId >= 1) newRow[colId - 1] = notifId;
    if (colTarget >= 1) newRow[colTarget - 1] = cleanTarget;
    if (colSenderU >= 1) newRow[colSenderU - 1] = cleanSender;
    if (colSenderN >= 1) newRow[colSenderN - 1] = senderNama || cleanSender;
    if (colRoom >= 1) newRow[colRoom - 1] = roomName || "";
    if (colSheet >= 1) newRow[colSheet - 1] = sheetSource || "";
    if (colType >= 1) newRow[colType - 1] = tipeNotif || "CHAT";
    if (colMsg >= 1) newRow[colMsg - 1] = shortMsg;
    if (colStatus >= 1) newRow[colStatus - 1] = "BELUM";
    if (colTime >= 1) newRow[colTime - 1] = now;

    notifSheet.appendRow(newRow);
  } catch (e) {}
}

function getPendingNotifications(username) {
  const ss = getSpreadsheet();
  const notifSheet = getNotifSheet(ss);
  const meta = getSheetHeaderAndRows(notifSheet, 10);
  if (meta.rows.length === 0) return [];

  const colId = findColumnIndexByNames(meta.headers, ["notifid", "id", "idnotif", "notificationid"], 1);
  const colTarget = findColumnIndexByNames(meta.headers, ["targetusername", "target", "tujuanusername", "tujuan", "penerima", "username"], 2);
  const colSenderU = findColumnIndexByNames(meta.headers, ["senderusername", "pengirimusername", "pengirimid", "dari"], 3);
  const colSenderN = findColumnIndexByNames(meta.headers, ["sendernama", "pengirim", "namapengirim", "nama"], 4);
  const colRoom = findColumnIndexByNames(meta.headers, ["roomname", "room", "grup", "namagrup", "topik"], 5);
  const colSheet = findColumnIndexByNames(meta.headers, ["sheetsource", "sheet", "namasheet", "sumber"], 6);
  const colType = findColumnIndexByNames(meta.headers, ["tipenotif", "tipe", "jenis", "type", "kategori"], 7);
  const colMsg = findColumnIndexByNames(meta.headers, ["ringkasanpesan", "pesan", "isi", "ringkasan", "message", "teks"], 8);
  const colStatus = findColumnIndexByNames(meta.headers, ["statusbaca", "status", "dibaca", "readstatus", "state"], 9);
  const colTime = findColumnIndexByNames(meta.headers, ["timestamp", "waktu", "tanggal", "time"], 10);

  const cleanUser = String(username || "").trim().toUpperCase();
  const result = [];

  for (let i = 0; i < meta.rows.length; i++) {
    const r = meta.rows[i];
    const target = String((colTarget >= 1 ? r[colTarget - 1] : "") || "").trim().toUpperCase();
    const sender = String((colSenderU >= 1 ? r[colSenderU - 1] : "") || "").trim().toUpperCase();
    const statusBaca = String((colStatus >= 1 ? r[colStatus - 1] : "BELUM") || "BELUM").trim().toUpperCase();
    const isAlreadyRead = (statusBaca === "SUDAH" || statusBaca === "DIBACA" || statusBaca === "READ" || statusBaca === "DONE" || statusBaca === "TERKIRIM");

    if (!isAlreadyRead && (target === cleanUser || (target === "ALL" && sender !== cleanUser))) {
      const rawTime = colTime >= 1 ? r[colTime - 1] : "";
      result.push({
        rowNumber: i + 2,
        notifId: String((colId >= 1 ? r[colId - 1] : "") || ""),
        targetUsername: target,
        senderUsername: sender,
        senderNama: String((colSenderN >= 1 ? r[colSenderN - 1] : sender) || sender),
        roomName: String((colRoom >= 1 ? r[colRoom - 1] : "") || ""),
        sheetSource: String((colSheet >= 1 ? r[colSheet - 1] : "") || ""),
        tipeNotif: String((colType >= 1 ? r[colType - 1] : "CHAT") || "CHAT"),
        ringkasanPesan: String((colMsg >= 1 ? r[colMsg - 1] : "") || ""),
        statusBaca: statusBaca,
        timestamp: (rawTime instanceof Date) ? rawTime.toISOString() : String(rawTime || "")
      });
    }
  }
  return result;
}

function markNotificationRead(notifId, username) {
  const ss = getSpreadsheet();
  const notifSheet = getNotifSheet(ss);
  const meta = getSheetHeaderAndRows(notifSheet, 9);
  if (meta.rows.length === 0) return { success: false };

  const colId = findColumnIndexByNames(meta.headers, ["notifid", "id", "idnotif", "notificationid"], 1);
  const colTarget = findColumnIndexByNames(meta.headers, ["targetusername", "target", "tujuanusername", "tujuan", "penerima", "username"], 2);
  const colStatus = findColumnIndexByNames(meta.headers, ["statusbaca", "status", "dibaca", "readstatus", "state"], 9);

  const cleanId = String(notifId || "").trim();
  const cleanUser = String(username || "").trim().toUpperCase();
  let updatedCount = 0;

  for (let i = 0; i < meta.rows.length; i++) {
    const rowId = String((colId >= 1 ? meta.rows[i][colId - 1] : "") || "").trim();
    const rowTarget = String((colTarget >= 1 ? meta.rows[i][colTarget - 1] : "") || "").trim().toUpperCase();
    if ((cleanId && rowId === cleanId) || (!cleanId && cleanUser && (rowTarget === cleanUser || rowTarget === "ALL"))) {
      if (colStatus >= 1) {
        notifSheet.getRange(i + 2, colStatus).setValue("SUDAH");
      }
      updatedCount++;
      if (cleanId) break;
    }
  }
  if (updatedCount > 0 && cleanUser) {
    appendActivityLog(cleanUser, "NOTIF_READ", "Menandai notifikasi antrean sebagai SUDAH (" + updatedCount + ")", "Utama");
  }
  return { success: true, updatedCount: updatedCount };
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

function loginUser(username, password, deviceInfo) {
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
      const nama = String(data[i][1] || "");
      const role = String(data[i][5] || "");
      appendActivityLog(cleanUser, "LOGIN_SUCCESS", "Berhasil login ke SansChat pada " + formatWibTimestamp(new Date()), "Login");
      updateUserActiveSession(cleanUser, nama, role, "ONLINE", true, "Utama", deviceInfo || ("PWA-" + cleanUser));
      return {
        success: true,
        redirectBaseUrl: getScriptUrl(),
        user: {
          username: rowUser,
          nama: nama,
          alamat: String(data[i][2] || ""),
          foto: String(data[i][4] || ""),
          role: role,
          helpSystemUrl: getHelpSystemUrl(),
          webAppsUrl: getWebAppsUrl(),
          sansChatUrl: getScriptUrl()
        }
      };
    }
  }
  appendActivityLog(cleanUser || "UNDEFINED", "LOGIN_FAILED", "Password/Username salah", "Login");
  return { success: false, message: "Username atau Password Salah!" };
}

function getInitialDashboardBundle(username, halaman) {
  const cleanUser = String(username || "").trim().toUpperCase();
  const profile = getUserProfileDetail(cleanUser);
  const pageLabel = String(halaman || "Utama").trim();
  if (cleanUser) {
    updateUserActiveSession(cleanUser, profile ? profile.nama : cleanUser, profile ? profile.role : "", "ONLINE", false, pageLabel);
    appendActivityLog(cleanUser, "OPEN_DASHBOARD", "Sinkronisasi dashboard & pengumuman (" + pageLabel + ")", pageLabel);
  }
  return {
    profile: profile,
    groups: getUserAssignedGroups(cleanUser),
    groupMapping: getNamaGrupMapping(),
    announcements: getAnnouncements(),
    activeSessions: getActiveSessions(),
    pendingNotifications: getPendingNotifications(cleanUser),
    hasUnreadPrivate: hasUnreadPrivateMessages(cleanUser),
    incomingCall: checkIncomingCall(cleanUser)
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

function getNamaGrupMapping() {
  const ss = getSpreadsheet();
  const groupSheet = ss.getSheetByName("NamaGrup");
  if (!groupSheet || groupSheet.getLastRow() < 2) return [];

  const numCols = Math.max(5, groupSheet.getLastColumn());
  const data = groupSheet.getRange(2, 1, groupSheet.getLastRow() - 1, numCols).getValues();
  let mappings = [];

  for (let i = 0; i < data.length; i++) {
    const colA = String(data[i][0] || "").trim();
    const colB = String(data[i][1] || "").trim();
    const colC = String(data[i][2] || "").trim();
    const colD = String(data[i][3] || "").trim();
    const colE = String(data[i][4] || "").trim();
    if (colA !== "") {
      const targetSheet = colB || colA.replace(/^Grup\s*/i, "Chat").replace(/\s+/g, "");
      mappings.push({
        rowIndex: i + 2,
        groupName: colA,
        sheetName: targetSheet,
        kolomConfig: colC,
        foto: colD,
        deskripsi: colE
      });
    }
  }
  return mappings;
}

function resolveGroupRoom(groupName, fallbackSheetName, username) {
  const ss = getSpreadsheet();
  const cleanGroup = String(groupName || "").trim();
  const cleanFallback = String(fallbackSheetName || "").trim();
  let resolvedSheetName = cleanFallback;
  let resolvedFoto = "";
  let resolvedKolom = "";

  const ng = ss.getSheetByName("NamaGrup");
  if (ng && ng.getLastRow() >= 2) {
    const numCols = Math.max(4, ng.getLastColumn());
    const ngData = ng.getRange(2, 1, ng.getLastRow() - 1, numCols).getValues();
    for (let i = 0; i < ngData.length; i++) {
      const colA = String(ngData[i][0] || "").trim();
      const colB = String(ngData[i][1] || "").trim();
      const colC = String(ngData[i][2] || "").trim();
      const colD = String(ngData[i][3] || "").trim();
      if (
        (cleanGroup && colA.toUpperCase() === cleanGroup.toUpperCase()) ||
        (cleanFallback && colB.toUpperCase() === cleanFallback.toUpperCase())
      ) {
        if (colB !== "") {
          resolvedSheetName = colB;
        }
        resolvedKolom = colC;
        resolvedFoto = colD;
        break;
      }
    }
  }

  if (!resolvedSheetName) {
    resolvedSheetName = cleanGroup.replace(/^Grup\s*/i, "Chat").replace(/\s+/g, "") || "ChatUmum";
  }

  let sheet = resolveChatSheet(ss, resolvedSheetName, cleanGroup);
  if (!sheet) {
    sheet = ss.insertSheet(resolvedSheetName);
    sheet.appendRow(["Timestamp", "SenderUsername", "SenderNama", "Pesan", "Tipe", "MediaUrl", "ReplyTo", "ReadBy"]);
    SpreadsheetApp.flush();
  }
  if (sheet.isSheetHidden()) sheet.showSheet();

  if (username) {
    const cleanUser = String(username).trim().toUpperCase();
    updateUserActiveSession(cleanUser, "", "", "ONLINE", false, cleanGroup || sheet.getName());
    appendActivityLog(cleanUser, "OPEN_ROOM", "Membuka room " + (cleanGroup || sheet.getName()) + " (Sheet: " + sheet.getName() + ")", cleanGroup || sheet.getName());
  }

  return {
    success: true,
    groupName: cleanGroup || resolvedSheetName,
    sheetName: sheet.getName(),
    kolomConfig: resolvedKolom,
    foto: resolvedFoto
  };
}

function getUserAssignedGroups(username) {
  const ss = getSpreadsheet();
  const userSheet = ss.getSheetByName("Users");
  const groupSheet = ss.getSheetByName("NamaGrup");
  if (!userSheet || !groupSheet || userSheet.getLastRow() < 2 || groupSheet.getLastRow() < 2) return [];

  const cleanUser = String(username || "").trim().toUpperCase();
  const userLastCol = Math.max(11, userSheet.getLastColumn());
  const userData = userSheet.getRange(2, 1, userSheet.getLastRow() - 1, userLastCol).getValues();

  let userRoles = [];
  let userRowData = null;
  for (let i = 0; i < userData.length; i++) {
    if (String(userData[i][0] || "").trim().toUpperCase() === cleanUser) {
      userRowData = userData[i];
      const roleStr = String(userData[i][5] || "").trim();
      if (roleStr !== "") {
        userRoles = roleStr.split(",").map(r => r.trim().toUpperCase()).filter(Boolean);
      }
      break;
    }
  }

  const groupNumCols = Math.max(5, groupSheet.getLastColumn());
  const groupData = groupSheet.getRange(2, 1, groupSheet.getLastRow() - 1, groupNumCols).getValues();
  let assignedRooms = [];

  for (let j = 0; j < groupData.length; j++) {
    const gName = String(groupData[j][0] || "").trim();
    const sName = String(groupData[j][1] || "").trim();
    const colCLetter = String(groupData[j][2] || "").trim();
    const gFoto = String(groupData[j][3] || "").trim();
    const gDesc = String(groupData[j][4] || "").trim();

    if (!gName) continue;

    let isMember = userRoles.includes(gName.toUpperCase());
    if (!isMember && userRowData && colCLetter) {
      const colIdx = columnLetterToIndex(colCLetter);
      if (colIdx >= 1 && colIdx <= userRowData.length) {
        const cellVal = String(userRowData[colIdx - 1] || "").trim().toUpperCase();
        if (cellVal === cleanUser) {
          isMember = true;
        }
      }
    }

    if (isMember) {
      const targetSheetName = sName || gName.replace(/^Grup\s*/i, "Chat").replace(/\s+/g, "");
      let lastMessage = "";
      let lastTime = "";
      let chatSheet = resolveChatSheet(ss, targetSheetName, gName);

      if (!chatSheet && targetSheetName) {
        chatSheet = ss.insertSheet(targetSheetName);
        chatSheet.appendRow(["Timestamp", "SenderUsername", "SenderNama", "Pesan", "Tipe", "MediaUrl", "ReplyTo", "ReadBy"]);
      }
      if (chatSheet && chatSheet.isSheetHidden()) chatSheet.showSheet();

      if (chatSheet && chatSheet.getLastRow() >= 2) {
        const lastRowVals = chatSheet.getRange(chatSheet.getLastRow(), 1, 1, 4).getValues()[0];
        lastTime = (lastRowVals[0] instanceof Date) ? lastRowVals[0].toISOString() : String(lastRowVals[0] || "");
        lastMessage = (lastRowVals[2] ? lastRowVals[2] + ": " : "") + String(lastRowVals[3] || "");
      }

      assignedRooms.push({
        name: gName,
        sheetName: chatSheet ? chatSheet.getName() : targetSheetName,
        kolomConfig: colCLetter,
        foto: gFoto,
        deskripsi: gDesc,
        lastMessage: lastMessage,
        lastTime: lastTime
      });
    }
  }
  return assignedRooms;
}

function getAllWargaContacts() {
  const sheet = getSpreadsheet().getSheetByName("Users");
  if (!sheet || sheet.getLastRow() < 2) return [];
  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 6).getValues();
  return data.map(r => ({
    username: String(r[0] || "").trim().toUpperCase(),
    nama: String(r[1] || "").trim(),
    alamat: String(r[2] || "").trim()
  })).filter(r => r.username && r.nama);
}

function updateProfilePicture(username, base64Data) {
  const sheet = getSpreadsheet().getSheetByName("Users");
  if (!sheet) return { success: false };
  const data = sheet.getDataRange().getValues();
  const cleanUser = String(username || "").trim().toUpperCase();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0] || "").trim().toUpperCase() === cleanUser) {
      sheet.getRange(i + 1, 5).setValue(base64Data);
      appendActivityLog(cleanUser, "UPDATE_AVATAR", "Memperbarui foto profil pengguna", "Utama");
      return { success: true };
    }
  }
  return { success: false };
}

function updateGroupProfilePicture(sheetName, base64Data) {
  const sheet = getSpreadsheet().getSheetByName("NamaGrup");
  if (!sheet) return { success: false };
  const data = sheet.getDataRange().getValues();
  const target = String(sheetName || "").trim().toUpperCase();
  for (let i = 1; i < data.length; i++) {
    const colA = String(data[i][0] || "").trim().toUpperCase();
    const colB = String(data[i][1] || "").trim().toUpperCase();
    if (colB === target || colA === target) {
      sheet.getRange(i + 1, 4).setValue(base64Data);
      appendActivityLog("ADMIN", "UPDATE_GROUP_AVATAR", "Memperbarui foto grup: " + target, target);
      return { success: true };
    }
  }
  return { success: false };
}

/**
 * Mengambil data pengumuman dari Spreadsheet 1kmZa08cDx5ts4Ld819mLsVmBJ83J-vZPu89i59TU3VI,
 * sheet "Informasi", dengan isi utama dari Kolom B (index 1).
 */
function getAnnouncements() {
  const ss = getSpreadsheet();
  let sheet = findExactSheetCaseInsensitive(ss, "Informasi");
  if (!sheet) {
    ensureDefaultSheets();
    sheet = findExactSheetCaseInsensitive(ss, "Informasi");
  }
  if (!sheet || sheet.getLastRow() < 2) return [];

  const numCols = Math.max(6, sheet.getLastColumn());
  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, numCols).getValues();
  const items = [];

  for (let i = 0; i < data.length; i++) {
    const judul = String(data[i][0] || "").trim();
    const isiKolomB = String(data[i][1] || "").trim();
    const prioritas = String(data[i][2] || "").trim().toUpperCase();
    const warnaTeks = String(data[i][3] || "").trim();
    const warnaLatar = String(data[i][4] || "").trim();
    const tanggalRaw = data[i][5];

    if (isiKolomB !== "" || judul !== "") {
      items.push({
        rowNumber: i + 2,
        judul: judul,
        teks: isiKolomB || judul,
        isiKolomB: isiKolomB,
        prioritas: prioritas,
        warnaTeks: warnaTeks,
        warnaLatar: warnaLatar,
        tanggal: (tanggalRaw instanceof Date) ? tanggalRaw.toISOString() : String(tanggalRaw || "")
      });
    }
  }
  return items;
}

function getGroupMembersByKolomC(groupName, chatSheetName) {
  const ss = getSpreadsheet();
  const userSheet = ss.getSheetByName("Users");
  const ngSheet = ss.getSheetByName("NamaGrup");
  if (!userSheet || userSheet.getLastRow() < 2) return [];

  const targetGroup = String(groupName || "").trim().toUpperCase();
  const targetSheet = String(chatSheetName || "").trim().toUpperCase();
  let colCRef = "";

  if (ngSheet && ngSheet.getLastRow() >= 2) {
    const ngData = ngSheet.getRange(2, 1, ngSheet.getLastRow() - 1, 3).getValues();
    for (let i = 0; i < ngData.length; i++) {
      const colA = String(ngData[i][0] || "").trim().toUpperCase();
      const colB = String(ngData[i][1] || "").trim().toUpperCase();
      const colC = String(ngData[i][2] || "").trim();
      if ((targetGroup && colA === targetGroup) || (targetSheet && colB === targetSheet)) {
        colCRef = colC;
        break;
      }
    }
  }

  const lastRow = userSheet.getLastRow();
  const lastCol = Math.max(11, userSheet.getLastColumn());
  const headers = userSheet.getRange(1, 1, 1, lastCol).getValues()[0];
  const userData = userSheet.getRange(2, 1, lastRow - 1, lastCol).getValues();

  const profileByUsername = {};
  for (let r = 0; r < userData.length; r++) {
    const uName = String(userData[r][0] || "").trim().toUpperCase();
    const nama = String(userData[r][1] || "").trim();
    const alamat = String(userData[r][2] || "").trim();
    const foto = String(userData[r][4] || "").trim();
    if (uName) {
      profileByUsername[uName] = {
        username: uName,
        nama: nama || uName,
        alamat: alamat || "-",
        foto: foto
      };
    }
  }

  let targetColIndex = columnLetterToIndex(colCRef);
  if (targetColIndex === -1 && colCRef) {
    for (let h = 0; h < headers.length; h++) {
      if (String(headers[h] || "").trim().toUpperCase() === colCRef.toUpperCase()) {
        targetColIndex = h + 1;
        break;
      }
    }
  }

  let members = [];
  const seen = {};

  if (targetColIndex >= 1 && targetColIndex <= lastCol) {
    for (let r = 0; r < userData.length; r++) {
      const cellVal = String(userData[r][targetColIndex - 1] || "").trim();
      if (!cellVal) continue;

      const tokens = cellVal.split(",").map(t => t.trim().toUpperCase()).filter(Boolean);
      for (let t = 0; t < tokens.length; t++) {
        const memberUser = tokens[t];
        if (seen[memberUser]) continue;
        seen[memberUser] = true;

        if (profileByUsername[memberUser]) {
          members.push(profileByUsername[memberUser]);
        } else {
          const rowFallbackUser = String(userData[r][0] || "").trim().toUpperCase();
          if (rowFallbackUser === memberUser) {
            members.push({
              username: memberUser,
              nama: String(userData[r][1] || memberUser).trim(),
              alamat: String(userData[r][2] || "-").trim(),
              foto: String(userData[r][4] || "").trim()
            });
          } else {
            members.push({
              username: memberUser,
              nama: memberUser,
              alamat: "-",
              foto: ""
            });
          }
        }
      }
    }
  }

  if (members.length === 0 && targetGroup) {
    for (let r = 0; r < userData.length; r++) {
      const usernameUser = String(userData[r][0] || "").trim().toUpperCase();
      const namaUser = String(userData[r][1] || "").trim();
      const alamatUser = String(userData[r][2] || "").trim();
      const fotoUser = String(userData[r][4] || "").trim();
      const roleCell = String(userData[r][5] || "").trim();
      if (usernameUser && namaUser && roleCell && !seen[usernameUser]) {
        const rolesArray = roleCell.split(",").map(item => item.trim().toUpperCase());
        if (rolesArray.includes(targetGroup)) {
          seen[usernameUser] = true;
          members.push({ username: usernameUser, nama: namaUser, alamat: alamatUser || "-", foto: fotoUser });
        }
      }
    }
  }

  return members;
}

function resolveChatSheet(ss, chatSheetName, groupNameHint) {
  const raw = String(chatSheetName || "").trim();
  const hint = String(groupNameHint || "").trim();

  const ng = ss.getSheetByName("NamaGrup");
  if (ng && ng.getLastRow() >= 2) {
    const ngData = ng.getRange(2, 1, ng.getLastRow() - 1, 2).getValues();
    for (let i = 0; i < ngData.length; i++) {
      const colA = String(ngData[i][0] || "").trim();
      const colB = String(ngData[i][1] || "").trim();
      if (
        (hint && colA.toUpperCase() === hint.toUpperCase()) ||
        (raw && colA.toUpperCase() === raw.toUpperCase()) ||
        (raw && colB.toUpperCase() === raw.toUpperCase())
      ) {
        if (colB) {
          return findExactSheetCaseInsensitive(ss, colB);
        }
      }
    }
  }

  if (raw) {
    const direct = findExactSheetCaseInsensitive(ss, raw);
    if (direct) return direct;
  }

  return null;
}

function sendMessage(chatSheetName, username, nama, pesan, tipe, mediaUrl, replyTo, groupName) {
  const lock = LockService.getScriptLock();
  lock.tryLock(10000);
  try {
    const ss = getSpreadsheet();
    const resolved = resolveGroupRoom(groupName, chatSheetName);
    const targetSheetName = resolved.sheetName || String(chatSheetName || "").trim();
    let sheet = resolveChatSheet(ss, targetSheetName, groupName);

    if (!sheet) {
      const safeName = targetSheetName || String(groupName || "ChatBaru").replace(/^Grup\s*/i, "Chat").replace(/\s+/g, "");
      sheet = ss.insertSheet(safeName);
      sheet.appendRow(["Timestamp", "SenderUsername", "SenderNama", "Pesan", "Tipe", "MediaUrl", "ReplyTo", "ReadBy"]);
    } else if (sheet.getLastRow() === 0) {
      sheet.appendRow(["Timestamp", "SenderUsername", "SenderNama", "Pesan", "Tipe", "MediaUrl", "ReplyTo", "ReadBy"]);
    }
    if (sheet.isSheetHidden()) sheet.showSheet();

    const cleanUser = String(username || "").trim().toUpperCase();
    const cleanSenderName = nama || cleanUser;
    const roomLabel = groupName || sheet.getName();
    sheet.appendRow([new Date(), cleanUser, cleanSenderName, pesan, tipe || "text", mediaUrl || "", replyTo || "", cleanUser]);
    updateUserActiveSession(cleanUser, cleanSenderName, "", "ONLINE", false, roomLabel);
    enqueueNotification("ALL", cleanUser, cleanSenderName, roomLabel, sheet.getName(), "GROUP_CHAT", pesan);
    appendActivityLog(cleanUser, "SEND_MESSAGE", "Kirim pesan ke room: " + roomLabel + " (" + sheet.getName() + ")", roomLabel);
    SpreadsheetApp.flush();
    return { success: true, sheetName: sheet.getName() };
  } finally {
    lock.releaseLock();
  }
}

function getMessages(chatSheetName, readerUsername, groupName) {
  const ss = getSpreadsheet();
  let sheet = resolveChatSheet(ss, chatSheetName, groupName);
  if (!sheet) {
    const resolved = resolveGroupRoom(groupName, chatSheetName);
    sheet = findExactSheetCaseInsensitive(ss, resolved.sheetName);
  }
  if (!sheet || sheet.getLastRow() < 2) return [];

  const lastRow = sheet.getLastRow();
  const numCols = Math.max(8, sheet.getLastColumn());
  const data = sheet.getRange(2, 1, lastRow - 1, numCols).getValues();
  const cleanReader = String(readerUsername || "").trim().toUpperCase();
  let msgs = [];

  for (let i = 0; i < data.length; i++) {
    let rawTime = data[i][0];
    let timeStr = (rawTime instanceof Date) ? rawTime.toISOString() : String(rawTime || "");
    let readByStr = String(data[i][7] || "");
    let readList = readByStr ? readByStr.split(",").map(s => s.trim().toUpperCase()).filter(Boolean) : [];

    if (cleanReader && !readList.includes(cleanReader)) {
      readList.push(cleanReader);
      sheet.getRange(i + 2, 8).setValue(readList.join(","));
    }

    msgs.push({
      rowNumber: i + 2,
      timestamp: timeStr,
      senderUsername: String(data[i][1] || ""),
      senderNama: String(data[i][2] || ""),
      pesan: String(data[i][3] || ""),
      tipe: String(data[i][4] || "text"),
      mediaUrl: String(data[i][5] || ""),
      replyTo: String(data[i][6] || ""),
      readCount: readList.length
    });
  }
  return msgs;
}

function votePollOption(chatSheetName, rowNumber, optionindex, voterUsername, groupName) {
  const ss = getSpreadsheet();
  const sheet = resolveChatSheet(ss, chatSheetName, groupName);
  if (!sheet || rowNumber < 2) return { success: false };
  const cell = sheet.getRange(rowNumber, 6);
  try {
    const pollData = JSON.parse(String(cell.getValue() || ""));
    if (!pollData.votes) pollData.votes = {};
    const cleanVoter = String(voterUsername || "").trim().toUpperCase();
    pollData.votes[cleanVoter] = optionindex;
    cell.setValue(JSON.stringify(pollData));
    appendActivityLog(cleanVoter, "VOTE_POLL", "Voting opsi #" + optionindex + " di " + (groupName || sheet.getName()), groupName || sheet.getName());
    SpreadsheetApp.flush();
    return { success: true };
  } catch (e) {
    return { success: false };
  }
}

function sendPrivateMessage(username, nama, pesan, tipe, mediaUrl, targetUsername, replyTo) {
  const lock = LockService.getScriptLock();
  lock.tryLock(10000);
  try {
    const ss = getSpreadsheet();
    let sheet = ss.getSheetByName("ChatPrivat");
    if (!sheet) {
      sheet = ss.insertSheet("ChatPrivat");
      sheet.appendRow(["Timestamp", "Username", "Nama", "Pesan", "Tipe", "MediaUrl", "StatusPesan", "TujuanUsername", "ReplyTo"]);
    }
    if (sheet.isSheetHidden()) sheet.showSheet();

    const cleanSender = String(username || "").trim().toUpperCase();
    const cleanTarget = String(targetUsername || cleanSender).trim().toUpperCase();
    const cleanSenderName = nama || cleanSender;

    sheet.appendRow([new Date(), cleanSender, cleanSenderName, pesan, tipe || "text", mediaUrl || "", "BELUM", cleanTarget, replyTo || ""]);
    updateUserActiveSession(cleanSender, cleanSenderName, "", "ONLINE", false, "Chat Privat");
    if (cleanTarget && cleanTarget !== cleanSender) {
      enqueueNotification(cleanTarget, cleanSender, cleanSenderName, "Chat Privat", "ChatPrivat", "PRIVATE_CHAT", pesan);
    }
    appendActivityLog(cleanSender, "SEND_PRIVATE", "Kirim pesan privat ke: " + cleanTarget, "Chat Privat");
    SpreadsheetApp.flush();
    return { success: true };
  } finally {
    lock.releaseLock();
  }
}

function getPrivateChatSheetRecords(username) {
  const sheet = getSpreadsheet().getSheetByName("ChatPrivat");
  if (!sheet || sheet.getLastRow() < 2) return [];
  const numCols = Math.max(9, sheet.getLastColumn());
  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, numCols).getValues();
  let msgs = [];
  const cleanUser = String(username || "").trim().toUpperCase();

  for (let i = 0; i < data.length; i++) {
    const rowSender = String(data[i][1] || "").trim().toUpperCase();
    const rowTarget = String(data[i][7] || rowSender).trim().toUpperCase();

    if (rowSender === cleanUser || rowTarget === cleanUser) {
      const pesanD = String(data[i][3] || "").trim();
      const pesanF = String(data[i][5] || "").trim();
      let rawTime = data[i][0];
      let timeStr = (rawTime instanceof Date) ? rawTime.toISOString() : String(rawTime || "");
      if (pesanD !== "" || pesanF !== "") {
        msgs.push({
          rowNumber: i + 2,
          timestamp: timeStr,
          pengirimUsername: rowSender,
          nama: String(data[i][2] || ""),
          pesanD: pesanD,
          pesanF: pesanF,
          tipe: String(data[i][4] || "text"),
          statusPesan: String(data[i][6] || ""),
          tujuanUsername: rowTarget,
          replyTo: String(data[i][8] || "")
        });
      }
    }
  }
  return msgs;
}

function hasUnreadPrivateMessages(username) {
  const sheet = getSpreadsheet().getSheetByName("ChatPrivat");
  if (!sheet || sheet.getLastRow() < 2) return false;
  const numCols = Math.max(8, sheet.getLastColumn());
  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, numCols).getValues();
  const uName = String(username || "").trim().toUpperCase();
  for (let i = 0; i < data.length; i++) {
    const rowSender = String(data[i][1] || "").trim().toUpperCase();
    const rowTarget = String(data[i][7] || rowSender).trim().toUpperCase();
    const status = String(data[i][6] || "").trim().toUpperCase();
    if (rowTarget === uName && rowSender !== uName && status !== "SUDAH") {
      return true;
    }
  }
  return false;
}

function markPrivateMessagesAsRead(username) {
  const sheet = getSpreadsheet().getSheetByName("ChatPrivat");
  if (!sheet || sheet.getLastRow() < 2) return;
  const numCols = Math.max(8, sheet.getLastColumn());
  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, numCols).getValues();
  const uName = String(username || "").trim().toUpperCase();
  for (let i = 0; i < data.length; i++) {
    const rowSender = String(data[i][1] || "").trim().toUpperCase();
    const rowTarget = String(data[i][7] || rowSender).trim().toUpperCase();
    if (rowTarget === uName && String(data[i][6] || "").trim().toUpperCase() !== "SUDAH") {
      sheet.getRange(i + 2, 7).setValue("SUDAH");
    }
  }
  markNotificationRead("", uName);
}

function initiateCall(targetUsername, callerUsername, callerName, type, roomUrl) {
  const ss = getSpreadsheet();
  let sheet = ss.getSheetByName("PanggilanAktif");
  if (!sheet) {
    sheet = ss.insertSheet("PanggilanAktif");
    sheet.appendRow(["TargetUsername", "CallerUsername", "CallerName", "Type", "Status", "RoomUrl", "Timestamp"]);
  }
  if (sheet.isSheetHidden()) sheet.showSheet();

  const cleanTarget = String(targetUsername || "ALL").trim().toUpperCase();
  const cleanCaller = String(callerUsername || "").trim().toUpperCase();
  const cleanCallerName = callerName || cleanCaller;
  const callType = type || "voice";

  sheet.appendRow([
    cleanTarget,
    cleanCaller,
    cleanCallerName,
    callType,
    "RINGING",
    roomUrl || "",
    new Date()
  ]);

  updateUserActiveSession(cleanCaller, cleanCallerName, "", "IN_CALL", false, "Panggilan");
  enqueueNotification(
    cleanTarget,
    cleanCaller,
    cleanCallerName,
    "Panggilan",
    "PanggilanAktif",
    callType === "video" ? "VIDEO_CALL" : "VOICE_CALL",
    "Panggilan masuk dari " + cleanCallerName
  );
  appendActivityLog(cleanCaller, "INITIATE_CALL", "Memulai " + callType + " call ke " + cleanTarget, "Panggilan");

  SpreadsheetApp.flush();
  return { success: true };
}

function checkIncomingCall(username) {
  const sheet = getSpreadsheet().getSheetByName("PanggilanAktif");
  if (!sheet || sheet.getLastRow() < 2) return null;
  const numCols = Math.max(7, sheet.getLastColumn());
  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, numCols).getValues();
  const uName = String(username || "").trim().toUpperCase();
  const now = Date.now();

  for (let i = data.length - 1; i >= 0; i--) {
    const target = String(data[i][0] || "").trim().toUpperCase();
    const caller = String(data[i][1] || "").trim().toUpperCase();
    const status = String(data[i][4] || "").trim().toUpperCase();
    const rawTime = data[i][6];
    const callTime = (rawTime instanceof Date) ? rawTime.getTime() : now;

    if ((target === uName || (target === "ALL" && caller !== uName)) && status === "RINGING" && (now - callTime) < 60000) {
      return {
        callerUsername: caller,
        callerName: String(data[i][2] || "Warga"),
        type: String(data[i][3] || "voice"),
        roomUrl: String(data[i][5] || "")
      };
    }
  }
  return null;
}

function updateCallStatus(username, newStatus) {
  const sheet = getSpreadsheet().getSheetByName("PanggilanAktif");
  if (!sheet || sheet.getLastRow() < 2) return;
  const data = sheet.getDataRange().getValues();
  const uName = String(username || "").trim().toUpperCase();
  for (let i = data.length - 1; i >= 1; i--) {
    const target = String(data[i][0] || "").trim().toUpperCase();
    const caller = String(data[i][1] || "").trim().toUpperCase();
    if (target === uName || caller === uName || target === "ALL") {
      sheet.getRange(i + 1, 5).setValue(newStatus);
      break;
    }
  }
  if (newStatus === "ENDED") {
    updateUserActiveSession(uName, "", "", "ONLINE", false, "Utama");
  }
  appendActivityLog(uName, "CALL_STATUS_" + String(newStatus || ""), "Status panggilan diubah menjadi " + newStatus, "Panggilan");
}
