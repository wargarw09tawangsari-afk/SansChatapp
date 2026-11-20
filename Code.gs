/**
 * ============================================================================
 * SPREADSHEET ID: 1kmZa08cDx5ts4Ld819mLsVmBJ83J-vZPu89i59TU3VI
 * Sc-SansChat Backend Script (Realtime Chat History, Penggolongan Kolom C,
 * Informasi Kolom B + Warna, SesiAktif, NotifikasiAntrean & Warga Management)
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
      sansChatUrl: getSansChatUrl(),
      activeSheets: getActiveSheetsSummary()
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
    ensureDefaultSheets();
    const payload = JSON.parse(e.postData.contents);
    const action = payload.action;
    let result = null;

    switch (action) {
      case 'loginUser':
        result = loginUser(payload.username, payload.password);
        break;
      case 'logoutUser':
        result = logoutUserSession(payload.username);
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
  let inf = ss.getSheetByName("Informasi");
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
  let logS = ss.getSheetByName("LogAktifitas");
  if (!logS) {
    logS = ss.insertSheet("LogAktifitas");
    logS.appendRow(["Timestamp", "Username", "Aktivitas", "Detail"]);
  }
  if (logS.isSheetHidden()) logS.showSheet();

  // 5. Sheet SesiAktif
  let sesiS = ss.getSheetByName("SesiAktif");
  if (!sesiS) {
    sesiS = ss.insertSheet("SesiAktif");
    sesiS.appendRow(["Username", "Nama", "Role", "DeviceID", "WaktuLogin", "StatusSesi", "LastActive"]);
  }
  if (sesiS.isSheetHidden()) sesiS.showSheet();

  // 6. Sheet NotifikasiAntrean
  let notifS = ss.getSheetByName("NotifikasiAntrean");
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

function appendActivityLog(username, aktivitas, detail) {
  try {
    const ss = getSpreadsheet();
    let logSheet = ss.getSheetByName("LogAktifitas");
    if (!logSheet) {
      logSheet = ss.insertSheet("LogAktifitas");
      logSheet.appendRow(["Timestamp", "Username", "Aktivitas", "Detail"]);
    }
    if (logSheet.isSheetHidden()) logSheet.showSheet();
    logSheet.appendRow([new Date(), String(username || "SYSTEM").toUpperCase(), aktivitas, detail]);
  } catch (e) {}
}

function updateUserActiveSession(username, nama, role, statusSesi, isNewLogin) {
  try {
    const ss = getSpreadsheet();
    let sesiSheet = ss.getSheetByName("SesiAktif");
    if (!sesiSheet) {
      sesiSheet = ss.insertSheet("SesiAktif");
      sesiSheet.appendRow(["Username", "Nama", "Role", "DeviceID", "WaktuLogin", "StatusSesi", "LastActive"]);
    }
    if (sesiSheet.isSheetHidden()) sesiSheet.showSheet();

    const cleanUser = String(username || "").trim().toUpperCase();
    const now = new Date();
    const lastRow = sesiSheet.getLastRow();
    if (lastRow >= 2) {
      const users = sesiSheet.getRange(2, 1, lastRow - 1, 1).getValues();
      for (let i = 0; i < users.length; i++) {
        if (String(users[i][0] || "").trim().toUpperCase() === cleanUser) {
          if (nama) sesiSheet.getRange(i + 2, 2).setValue(nama);
          if (role) sesiSheet.getRange(i + 2, 3).setValue(role);
          if (isNewLogin) sesiSheet.getRange(i + 2, 5).setValue(now);
          sesiSheet.getRange(i + 2, 6).setValue(statusSesi || "ONLINE");
          sesiSheet.getRange(i + 2, 7).setValue(now);
          return;
        }
      }
    }
    sesiSheet.appendRow([cleanUser, nama || cleanUser, role || "", "PWA-" + cleanUser, now, statusSesi || "ONLINE", now]);
  } catch (e) {}
}

function logoutUserSession(username) {
  const cleanUser = String(username || "").trim().toUpperCase();
  if (!cleanUser) return { success: false };
  updateUserActiveSession(cleanUser, "", "", "OFFLINE", false);
  appendActivityLog(cleanUser, "LOGOUT", "Sesi diakhiri dari perangkat");
  return { success: true };
}

function getActiveSessions() {
  const ss = getSpreadsheet();
  const sesiSheet = ss.getSheetByName("SesiAktif");
  if (!sesiSheet || sesiSheet.getLastRow() < 2) return [];
  const numCols = Math.max(7, sesiSheet.getLastColumn());
  const rows = sesiSheet.getRange(2, 1, sesiSheet.getLastRow() - 1, numCols).getValues();
  return rows.map(r => ({
    username: String(r[0] || "").trim().toUpperCase(),
    nama: String(r[1] || "").trim(),
    role: String(r[2] || "").trim(),
    deviceId: String(r[3] || "").trim(),
    waktuLogin: (r[4] instanceof Date) ? r[4].toISOString() : String(r[4] || ""),
    statusSesi: String(r[5] || "OFFLINE").trim().toUpperCase(),
    lastActive: (r[6] instanceof Date) ? r[6].toISOString() : String(r[6] || "")
  })).filter(item => item.username);
}

function enqueueNotification(targetUsername, senderUsername, senderNama, roomName, sheetSource, tipeNotif, ringkasanPesan) {
  try {
    const ss = getSpreadsheet();
    let notifSheet = ss.getSheetByName("NotifikasiAntrean");
    if (!notifSheet) {
      notifSheet = ss.insertSheet("NotifikasiAntrean");
      notifSheet.appendRow(["NotifID", "TargetUsername", "SenderUsername", "SenderNama", "RoomName", "SheetSource", "TipeNotif", "RingkasanPesan", "StatusBaca", "Timestamp"]);
    }
    if (notifSheet.isSheetHidden()) notifSheet.showSheet();

    const cleanTarget = String(targetUsername || "ALL").trim().toUpperCase();
    const cleanSender = String(senderUsername || "SYSTEM").trim().toUpperCase();
    const notifId = "NTF-" + new Date().getTime() + "-" + Math.floor(Math.random() * 900 + 100);
    const shortMsg = String(ringkasanPesan || "").substring(0, 180);

    notifSheet.appendRow([
      notifId,
      cleanTarget,
      cleanSender,
      senderNama || cleanSender,
      roomName || "",
      sheetSource || "",
      tipeNotif || "CHAT",
      shortMsg,
      "BELUM",
      new Date()
    ]);
  } catch (e) {}
}

function getPendingNotifications(username) {
  const ss = getSpreadsheet();
  const notifSheet = ss.getSheetByName("NotifikasiAntrean");
  if (!notifSheet || notifSheet.getLastRow() < 2) return [];
  const numCols = Math.max(10, notifSheet.getLastColumn());
  const rows = notifSheet.getRange(2, 1, notifSheet.getLastRow() - 1, numCols).getValues();
  const cleanUser = String(username || "").trim().toUpperCase();
  const result = [];

  for (let i = 0; i < rows.length; i++) {
    const target = String(rows[i][1] || "").trim().toUpperCase();
    const sender = String(rows[i][2] || "").trim().toUpperCase();
    const statusBaca = String(rows[i][8] || "BELUM").trim().toUpperCase();
    if (statusBaca !== "SUDAH" && (target === cleanUser || (target === "ALL" && sender !== cleanUser))) {
      result.push({
        rowNumber: i + 2,
        notifId: String(rows[i][0] || ""),
        targetUsername: target,
        senderUsername: sender,
        senderNama: String(rows[i][3] || sender),
        roomName: String(rows[i][4] || ""),
        sheetSource: String(rows[i][5] || ""),
        tipeNotif: String(rows[i][6] || "CHAT"),
        ringkasanPesan: String(rows[i][7] || ""),
        statusBaca: statusBaca,
        timestamp: (rows[i][9] instanceof Date) ? rows[i][9].toISOString() : String(rows[i][9] || "")
      });
    }
  }
  return result;
}

function markNotificationRead(notifId, username) {
  const ss = getSpreadsheet();
  const notifSheet = ss.getSheetByName("NotifikasiAntrean");
  if (!notifSheet || notifSheet.getLastRow() < 2) return { success: false };
  const numCols = Math.max(9, notifSheet.getLastColumn());
  const rows = notifSheet.getRange(2, 1, notifSheet.getLastRow() - 1, numCols).getValues();
  const cleanId = String(notifId || "").trim();
  const cleanUser = String(username || "").trim().toUpperCase();

  for (let i = 0; i < rows.length; i++) {
    const rowId = String(rows[i][0] || "").trim();
    const rowTarget = String(rows[i][1] || "").trim().toUpperCase();
    if ((cleanId && rowId === cleanId) || (!cleanId && cleanUser && (rowTarget === cleanUser || rowTarget === "ALL"))) {
      notifSheet.getRange(i + 2, 9).setValue("SUDAH");
      if (cleanId) break;
    }
  }
  return { success: true };
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
      const nama = String(data[i][1] || "");
      const role = String(data[i][5] || "");
      appendActivityLog(cleanUser, "LOGIN_SUCCESS", "Berhasil login ke SansChat");
      updateUserActiveSession(cleanUser, nama, role, "ONLINE", true);
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
  appendActivityLog(cleanUser || "UNDEFINED", "LOGIN_FAILED", "Password/Username salah");
  return { success: false, message: "Username atau Password Salah!" };
}

function getInitialDashboardBundle(username) {
  const cleanUser = String(username || "").trim().toUpperCase();
  const profile = getUserProfileDetail(cleanUser);
  if (cleanUser) {
    updateUserActiveSession(cleanUser, profile ? profile.nama : cleanUser, profile ? profile.role : "", "ONLINE", false);
  }
  return {
    profile: profile,
    groups: getUserAssignedGroups(cleanUser),
    groupMapping: getNamaGrupMapping(),
    announcements: getAnnouncements(),
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

function resolveGroupRoom(groupName, fallbackSheetName) {
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
      return { success: true };
    }
  }
  return { success: false };
}

function getAnnouncements() {
  const ss = getSpreadsheet();
  let sheet = ss.getSheetByName("Informasi");
  if (!sheet) {
    ensureDefaultSheets();
    sheet = ss.getSheetByName("Informasi");
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

    if (isiKolomB || judul) {
      items.push({
        judul: judul,
        teks: isiKolomB || judul,
        prioritas: prioritas,
        warnaTeks: warnaTeks,
        warnaLatar: warnaLatar
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
    sheet.appendRow([new Date(), cleanUser, cleanSenderName, pesan, tipe || "text", mediaUrl || "", replyTo || "", cleanUser]);
    updateUserActiveSession(cleanUser, cleanSenderName, "", "ONLINE", false);
    enqueueNotification("ALL", cleanUser, cleanSenderName, groupName || sheet.getName(), sheet.getName(), "GROUP_CHAT", pesan);
    appendActivityLog(cleanUser, "SEND_MESSAGE", "Kirim pesan ke room: " + sheet.getName());
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
    pollData.votes[String(voterUsername || "").trim().toUpperCase()] = optionindex;
    cell.setValue(JSON.stringify(pollData));
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
    updateUserActiveSession(cleanSender, cleanSenderName, "", "ONLINE", false);
    if (cleanTarget && cleanTarget !== cleanSender) {
      enqueueNotification(cleanTarget, cleanSender, cleanSenderName, "Chat Privat", "ChatPrivat", "PRIVATE_CHAT", pesan);
    }
    appendActivityLog(cleanSender, "SEND_PRIVATE", "Kirim pesan privat ke: " + cleanTarget);
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

  updateUserActiveSession(cleanCaller, cleanCallerName, "", "IN_CALL", false);
  enqueueNotification(
    cleanTarget,
    cleanCaller,
    cleanCallerName,
    "Panggilan",
    "PanggilanAktif",
    callType === "video" ? "VIDEO_CALL" : "VOICE_CALL",
    "Panggilan masuk dari " + cleanCallerName
  );

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
    updateUserActiveSession(uName, "", "", "ONLINE", false);
  }
}
