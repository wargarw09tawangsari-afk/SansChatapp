/**
 * ============================================================================
 * SPREADSHEET ID: 1kmZa08cDx5ts4Ld819mLsVmBJ83J-vZPu89i59TU3VI
 * Sc-SansChat Backend Script (Realtime Chat History, Penggolongan Kolom C & Warga Management)
 * Enhanced Support: Sheet Informasi, SesiAktif, NotifikasiAntrean
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
    ensureDefaultSheets();
    const payload = JSON.parse(e.postData.contents);
    const action = payload.action;
    let result = null;

    switch (action) {
      case 'loginUser':
        result = loginUser(payload.username, payload.password, payload.deviceId);
        break;
      case 'logoutUser':
        result = logoutUser(payload.username);
        break;
      case 'heartbeatSession':
        result = heartbeatSession(payload.username, payload.deviceId);
        break;
      case 'getActiveSessions':
        result = getActiveSessions(payload.username);
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
      case 'getInformasiItems':
        result = getInformasiItems(payload.username);
        break;
      case 'addInformasiItem':
        result = addInformasiItem(payload.username, payload.isiPengumuman, payload.prioritas, payload.targetRole);
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
      case 'enqueueNotification':
        result = enqueueNotification(
          payload.targetUsername,
          payload.senderUsername,
          payload.senderNama,
          payload.roomName,
          payload.sheetSource,
          payload.tipeNotif,
          payload.ringkasanPesan
        );
        break;
      case 'getPendingNotifications':
        result = getPendingNotifications(payload.username);
        break;
      case 'ackNotification':
        result = ackNotification(payload.notifId, payload.username, payload.statusBaca);
        break;
      case 'markAllNotificationsRead':
        result = markAllNotificationsRead(payload.username);
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

  if (ng.getLastRow() >= 2) {
    const rows = ng.getRange(2, 1, ng.getLastRow() - 1, 2).getValues();
    for (let i = 0; i < rows.length; i++) {
      const colA = String(rows[i][0] || "").trim();
      let colB = String(rows[i][1] || "").trim();
      if (colA && !colB) {
        colB = colA.replace(/^Grup\s*/i, "Chat").replace(/\s+/g, "");
        ng.getRange(i + 2, 2).setValue(colB);
      }
      if (colA && colB && !findExactSheetCaseInsensitive(ss, colB)) {
        const created = ss.insertSheet(colB);
        created.appendRow(["Timestamp", "SenderUsername", "SenderNama", "Pesan", "Tipe", "MediaUrl", "ReplyTo", "ReadBy"]);
      }
    }
  }

  let inf = ss.getSheetByName("Informasi");
  if (!inf) {
    inf = ss.insertSheet("Informasi");
    inf.appendRow(["Isi Pengumuman", "Prioritas", "Tanggal", "Penulis", "TargetRole", "StatusAktif"]);
    inf.appendRow([
      "Diberitahukan kepada seluruh warga RT 37, 38, dan 39 untuk hadir dalam kerja bakti kebersihan lingkungan pada hari Minggu pukul 07.00 WIB di balai warga.",
      "TINGGI",
      new Date(),
      "ADMINPUSAT",
      "SEMUA",
      "AKTIF"
    ]);
  } else if (inf.getLastRow() >= 1 && inf.getLastColumn() < 3) {
    inf.getRange(1, 1, 1, 3).setValues([["Isi Pengumuman", "Prioritas", "Tanggal"]]);
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

  let sesiS = ss.getSheetByName("SesiAktif");
  if (!sesiS) {
    sesiS = ss.insertSheet("SesiAktif");
    sesiS.appendRow(["Username", "Nama", "Role", "DeviceID", "WaktuLogin", "StatusSesi", "LastActive"]);
  } else if (sesiS.getLastRow() === 0) {
    sesiS.appendRow(["Username", "Nama", "Role", "DeviceID", "WaktuLogin", "StatusSesi", "LastActive"]);
  }

  let notifS = ss.getSheetByName("NotifikasiAntrean");
  if (!notifS) {
    notifS = ss.insertSheet("NotifikasiAntrean");
    notifS.appendRow(["NotifID", "TargetUsername", "SenderUsername", "SenderNama", "RoomName", "SheetSource", "TipeNotif", "RingkasanPesan", "StatusBaca", "Timestamp"]);
  } else if (notifS.getLastRow() === 0) {
    notifS.appendRow(["NotifID", "TargetUsername", "SenderUsername", "SenderNama", "RoomName", "SheetSource", "TipeNotif", "RingkasanPesan", "StatusBaca", "Timestamp"]);
  }
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
    logSheet.appendRow([new Date(), String(username || "SYSTEM").toUpperCase(), aktivitas, detail]);
  } catch (e) {}
}

function updateUserActiveSession(username, nama, role, statusSesi, deviceId) {
  const lock = LockService.getScriptLock();
  lock.tryLock(8000);
  try {
    const ss = getSpreadsheet();
    let sesiSheet = ss.getSheetByName("SesiAktif");
    if (!sesiSheet) {
      sesiSheet = ss.insertSheet("SesiAktif");
      sesiSheet.appendRow(["Username", "Nama", "Role", "DeviceID", "WaktuLogin", "StatusSesi", "LastActive"]);
    }
    const cleanUser = String(username || "").trim().toUpperCase();
    if (!cleanUser) return { success: false };

    const now = new Date();
    const cleanDevice = String(deviceId || ("PWA-" + cleanUser)).trim();
    const lastRow = sesiSheet.getLastRow();
    const statusUpper = String(statusSesi || "ONLINE").trim().toUpperCase();

    if (lastRow >= 2) {
      const numCols = Math.max(7, sesiSheet.getLastColumn());
      const rows = sesiSheet.getRange(2, 1, lastRow - 1, numCols).getValues();
      for (let i = 0; i < rows.length; i++) {
        const rowUser = String(rows[i][0] || "").trim().toUpperCase();
        if (rowUser === cleanUser) {
          const rowNum = i + 2;
          if (nama) sesiSheet.getRange(rowNum, 2).setValue(nama);
          if (role) sesiSheet.getRange(rowNum, 3).setValue(role);
          if (deviceId) sesiSheet.getRange(rowNum, 4).setValue(cleanDevice);
          if (statusUpper === "ONLINE") {
            sesiSheet.getRange(rowNum, 5).setValue(now);
          }
          sesiSheet.getRange(rowNum, 6).setValue(statusUpper);
          sesiSheet.getRange(rowNum, 7).setValue(now);
          SpreadsheetApp.flush();
          return { success: true, username: cleanUser, statusSesi: statusUpper };
        }
      }
    }

    sesiSheet.appendRow([
      cleanUser,
      nama || cleanUser,
      role || "WARGA",
      cleanDevice,
      now,
      statusUpper,
      now
    ]);
    SpreadsheetApp.flush();
    return { success: true, username: cleanUser, statusSesi: statusUpper };
  } finally {
    lock.releaseLock();
  }
}

function heartbeatSession(username, deviceId) {
  const cleanUser = String(username || "").trim().toUpperCase();
  if (!cleanUser) return { success: false };
  const profile = getUserProfileDetail(cleanUser);
  return updateUserActiveSession(
    cleanUser,
    profile ? profile.nama : cleanUser,
    profile ? profile.role : "",
    "ONLINE",
    deviceId
  );
}

function logoutUser(username) {
  const cleanUser = String(username || "").trim().toUpperCase();
  if (!cleanUser) return { success: false };
  const profile = getUserProfileDetail(cleanUser);
  updateUserActiveSession(
    cleanUser,
    profile ? profile.nama : cleanUser,
    profile ? profile.role : "",
    "OFFLINE",
    null
  );
  appendActivityLog(cleanUser, "LOGOUT", "Sesi ditandai OFFLINE pada SesiAktif");
  return { success: true };
}

function getActiveSessions(requestingUsername) {
  const ss = getSpreadsheet();
  const sesiSheet = ss.getSheetByName("SesiAktif");
  if (!sesiSheet || sesiSheet.getLastRow() < 2) return [];
  const numCols = Math.max(7, sesiSheet.getLastColumn());
  const data = sesiSheet.getRange(2, 1, sesiSheet.getLastRow() - 1, numCols).getValues();
  const nowMs = Date.now();
  const list = [];

  for (let i = 0; i < data.length; i++) {
    const u = String(data[i][0] || "").trim().toUpperCase();
    if (!u) continue;
    const waktuLogin = (data[i][4] instanceof Date) ? data[i][4].toISOString() : String(data[i][4] || "");
    const rawStatus = String(data[i][5] || "OFFLINE").trim().toUpperCase();
    const lastActiveDate = (data[i][6] instanceof Date) ? data[i][6] : null;
    const lastActiveIso = lastActiveDate ? lastActiveDate.toISOString() : String(data[i][6] || "");
    const isStale = lastActiveDate ? ((nowMs - lastActiveDate.getTime()) > 10 * 60 * 1000) : false;
    const effectiveStatus = (rawStatus === "ONLINE" && isStale) ? "IDLE" : rawStatus;

    list.push({
      username: u,
      nama: String(data[i][1] || u),
      role: String(data[i][2] || ""),
      deviceId: String(data[i][3] || ""),
      waktuLogin: waktuLogin,
      statusSesi: effectiveStatus,
      lastActive: lastActiveIso
    });
  }
  return list;
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

function loginUser(username, password, deviceId) {
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
      updateUserActiveSession(cleanUser, nama, role, "ONLINE", deviceId);
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
  if (cleanUser) {
    heartbeatSession(cleanUser);
  }
  return {
    profile: getUserProfileDetail(cleanUser),
    groups: getUserAssignedGroups(cleanUser),
    groupMapping: getNamaGrupMapping(),
    announcements: getAnnouncements(),
    informasiItems: getInformasiItems(cleanUser),
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
  const sheet = getSpreadsheet().getSheetByName("Informasi");
  if (!sheet || sheet.getLastRow() < 2) return [];
  const numCols = Math.max(3, sheet.getLastColumn());
  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, numCols).getValues();
  const priorityRank = { "DARURAT": 0, "URGENT": 0, "TINGGI": 1, "HIGH": 1, "NORMAL": 2, "SEDANG": 2, "RENDAH": 3, "LOW": 3 };
  const items = [];

  for (let i = 0; i < data.length; i++) {
    const isi = String(data[i][0] || "").trim();
    if (!isi) continue;
    const prioritas = String(data[i][1] || "NORMAL").trim().toUpperCase();
    const rawDate = data[i][2];
    const tanggalMs = (rawDate instanceof Date) ? rawDate.getTime() : 0;
    const statusAktif = (numCols >= 6) ? String(data[i][5] || "AKTIF").trim().toUpperCase() : "AKTIF";
    if (statusAktif === "NONAKTIF" || statusAktif === "ARSIP") continue;
    items.push({
      isi: isi,
      prioritas: prioritas,
      rank: (priorityRank[prioritas] !== undefined) ? priorityRank[prioritas] : 2,
      tanggalMs: tanggalMs,
      orderIdx: i
    });
  }

  items.sort((a, b) => {
    if (a.rank !== b.rank) return a.rank - b.rank;
    if (b.tanggalMs !== a.tanggalMs) return b.tanggalMs - a.tanggalMs;
    return b.orderIdx - a.orderIdx;
  });

  return items.map(it => it.isi);
}

function getInformasiItems(username) {
  const sheet = getSpreadsheet().getSheetByName("Informasi");
  if (!sheet || sheet.getLastRow() < 2) return [];
  const numCols = Math.max(3, sheet.getLastColumn());
  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, numCols).getValues();
  const result = [];
  for (let i = 0; i < data.length; i++) {
    const isi = String(data[i][0] || "").trim();
    if (!isi) continue;
    const rawDate = data[i][2];
    const tanggalStr = (rawDate instanceof Date) ? rawDate.toISOString() : String(rawDate || "");
    result.push({
      rowNumber: i + 2,
      isiPengumuman: isi,
      prioritas: String(data[i][1] || "NORMAL").trim().toUpperCase(),
      tanggal: tanggalStr,
      penulis: numCols >= 4 ? String(data[i][3] || "") : "",
      targetRole: numCols >= 5 ? String(data[i][4] || "SEMUA") : "SEMUA",
      statusAktif: numCols >= 6 ? String(data[i][5] || "AKTIF") : "AKTIF"
    });
  }
  return result.reverse();
}

function addInformasiItem(username, isiPengumuman, prioritas, targetRole) {
  const lock = LockService.getScriptLock();
  lock.tryLock(8000);
  try {
    const cleanText = String(isiPengumuman || "").trim();
    if (!cleanText) return { success: false, message: "Isi pengumuman kosong" };
    const ss = getSpreadsheet();
    let sheet = ss.getSheetByName("Informasi");
    if (!sheet) {
      sheet = ss.insertSheet("Informasi");
      sheet.appendRow(["Isi Pengumuman", "Prioritas", "Tanggal", "Penulis", "TargetRole", "StatusAktif"]);
    }
    const cleanUser = String(username || "ADMIN").trim().toUpperCase();
    const cleanPrio = String(prioritas || "NORMAL").trim().toUpperCase();
    const cleanTarget = String(targetRole || "SEMUA").trim().toUpperCase();
    const now = new Date();

    if (sheet.getLastColumn() >= 6 || sheet.getLastRow() === 1) {
      if (sheet.getLastColumn() < 6 && sheet.getLastRow() === 1) {
        sheet.getRange(1, 1, 1, 6).setValues([["Isi Pengumuman", "Prioritas", "Tanggal", "Penulis", "TargetRole", "StatusAktif"]]);
      }
      sheet.appendRow([cleanText, cleanPrio, now, cleanUser, cleanTarget, "AKTIF"]);
    } else {
      sheet.appendRow([cleanText, cleanPrio, now]);
    }

    appendActivityLog(cleanUser, "ADD_INFORMASI", "Menambah pengumuman (" + cleanPrio + ")");
    SpreadsheetApp.flush();
    return { success: true };
  } finally {
    lock.releaseLock();
  }
}

function enqueueNotification(targetUsername, senderUsername, senderNama, roomName, sheetSource, tipeNotif, ringkasanPesan) {
  const lock = LockService.getScriptLock();
  lock.tryLock(8000);
  try {
    const ss = getSpreadsheet();
    let notifSheet = ss.getSheetByName("NotifikasiAntrean");
    if (!notifSheet) {
      notifSheet = ss.insertSheet("NotifikasiAntrean");
      notifSheet.appendRow(["NotifID", "TargetUsername", "SenderUsername", "SenderNama", "RoomName", "SheetSource", "TipeNotif", "RingkasanPesan", "StatusBaca", "Timestamp"]);
    }
    const notifId = "NTF-" + Date.now() + "-" + Math.floor(Math.random() * 900 + 100);
    const now = new Date();
    const cleanSummary = String(ringkasanPesan || "").substring(0, 200);
    notifSheet.appendRow([
      notifId,
      String(targetUsername || "ALL").trim().toUpperCase(),
      String(senderUsername || "").trim().toUpperCase(),
      String(senderNama || senderUsername || "Warga").trim(),
      String(roomName || "").trim(),
      String(sheetSource || "").trim(),
      String(tipeNotif || "CHAT").trim().toUpperCase(),
      cleanSummary,
      "BELUM",
      now
    ]);
    SpreadsheetApp.flush();
    return { success: true, notifId: notifId };
  } finally {
    lock.releaseLock();
  }
}

function getPendingNotifications(username) {
  const ss = getSpreadsheet();
  const notifSheet = ss.getSheetByName("NotifikasiAntrean");
  if (!notifSheet || notifSheet.getLastRow() < 2) return [];
  const cleanUser = String(username || "").trim().toUpperCase();
  if (!cleanUser) return [];

  const numCols = Math.max(10, notifSheet.getLastColumn());
  const data = notifSheet.getRange(2, 1, notifSheet.getLastRow() - 1, numCols).getValues();
  const pending = [];

  for (let i = 0; i < data.length; i++) {
    const target = String(data[i][1] || "").trim().toUpperCase();
    const sender = String(data[i][2] || "").trim().toUpperCase();
    const statusBaca = String(data[i][8] || "BELUM").trim().toUpperCase();
    if (statusBaca !== "BELUM") continue;
    if (target === cleanUser || (target === "ALL" && sender !== cleanUser)) {
      const ts = (data[i][9] instanceof Date) ? data[i][9].toISOString() : String(data[i][9] || "");
      pending.push({
        rowNumber: i + 2,
        notifId: String(data[i][0] || ""),
        targetUsername: target,
        senderUsername: sender,
        senderNama: String(data[i][3] || sender),
        roomName: String(data[i][4] || ""),
        sheetSource: String(data[i][5] || ""),
        tipeNotif: String(data[i][6] || "CHAT"),
        ringkasanPesan: String(data[i][7] || ""),
        statusBaca: statusBaca,
        timestamp: ts
      });
    }
  }
  return pending;
}

function ackNotification(notifId, username, statusBaca) {
  const lock = LockService.getScriptLock();
  lock.tryLock(8000);
  try {
    const ss = getSpreadsheet();
    const notifSheet = ss.getSheetByName("NotifikasiAntrean");
    if (!notifSheet || notifSheet.getLastRow() < 2) return { success: false };
    const cleanId = String(notifId || "").trim();
    const cleanUser = String(username || "").trim().toUpperCase();
    const nextStatus = String(statusBaca || "SUDAH").trim().toUpperCase();
    const lastRow = notifSheet.getLastRow();
    const ids = notifSheet.getRange(2, 1, lastRow - 1, 2).getValues();
    let updated = 0;

    for (let i = 0; i < ids.length; i++) {
      const rowId = String(ids[i][0] || "").trim();
      const rowTarget = String(ids[i][1] || "").trim().toUpperCase();
      if ((cleanId && rowId === cleanId) || (!cleanId && cleanUser && rowTarget === cleanUser)) {
        notifSheet.getRange(i + 2, 9).setValue(nextStatus);
        updated++;
      }
    }
    SpreadsheetApp.flush();
    return { success: true, updated: updated };
  } finally {
    lock.releaseLock();
  }
}

function markAllNotificationsRead(username) {
  return ackNotification("", username, "SUDAH");
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

    const cleanUser = String(username || "").trim().toUpperCase();
    const cleanSenderNama = nama || cleanUser;
    sheet.appendRow([new Date(), cleanUser, cleanSenderNama, pesan, tipe || "text", mediaUrl || "", replyTo || "", cleanUser]);
    appendActivityLog(cleanUser, "SEND_MESSAGE", "Kirim pesan ke room: " + sheet.getName());
    enqueueNotification("ALL", cleanUser, cleanSenderNama, groupName || sheet.getName(), sheet.getName(), "GROUP_CHAT", pesan);
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
    const cleanSender = String(username || "").trim().toUpperCase();
    const cleanTarget = String(targetUsername || cleanSender).trim().toUpperCase();
    const cleanSenderNama = nama || cleanSender;
    sheet.appendRow([new Date(), cleanSender, cleanSenderNama, pesan, tipe || "text", mediaUrl || "", "BELUM", cleanTarget, replyTo || ""]);
    appendActivityLog(cleanSender, "SEND_PRIVATE", "Kirim pesan privat ke: " + cleanTarget);
    enqueueNotification(cleanTarget, cleanSender, cleanSenderNama, "Chat Privat", "ChatPrivat", "PRIVATE_CHAT", pesan);
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
}

function initiateCall(targetUsername, callerUsername, callerName, type, roomUrl) {
  const ss = getSpreadsheet();
  let sheet = ss.getSheetByName("PanggilanAktif");
  if (!sheet) {
    sheet = ss.insertSheet("PanggilanAktif");
    sheet.appendRow(["TargetUsername", "CallerUsername", "CallerName", "Type", "Status", "RoomUrl", "Timestamp"]);
  }
  const cleanTarget = String(targetUsername || "ALL").trim().toUpperCase();
  const cleanCaller = String(callerUsername || "").trim().toUpperCase();
  const cleanCallerName = callerName || callerUsername;
  const cleanType = type || "voice";
  const cleanUrl = roomUrl || "";

  sheet.appendRow([
    cleanTarget,
    cleanCaller,
    cleanCallerName,
    cleanType,
    "RINGING",
    cleanUrl,
    new Date()
  ]);
  enqueueNotification(
    cleanTarget,
    cleanCaller,
    cleanCallerName,
    cleanType === "video" ? "Video Call" : "Panggilan Suara",
    "PanggilanAktif",
    "CALL",
    (cleanType === "video" ? "Video Call dari " : "Panggilan Suara dari ") + cleanCallerName
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
}
