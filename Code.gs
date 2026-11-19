/**
 * ============================================================================
 * SPREADSHEET ID: 1kmZa08cDx5ts4Ld819mLsVmBJ83J-vZPu89i59TU3VI
 * Sc-SansChat Backend Script (Realtime Chat History & Warga Management)
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
      case 'getGroupMembersByRole':
        result = getGroupMembersByRole(payload.groupName);
        break;
      case 'sendMessage':
        result = sendMessage(payload.chatSheetName, payload.username, payload.nama, payload.pesan, payload.tipe, payload.mediaUrl, payload.replyTo);
        break;
      case 'getMessages':
        result = getMessages(payload.chatSheetName, payload.username);
        break;
      case 'votePollOption':
        result = votePollOption(payload.chatSheetName, payload.rowNumber, payload.optionindex, payload.voterUsername);
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

function ensureDefaultSheets() {
  const ss = getSpreadsheet();
  if (!ss.getSheetByName("Users")) {
    const u = ss.insertSheet("Users");
    u.appendRow(["Username", "Nama", "Alamat", "Password", "Foto", "Role"]);
  }
  if (!ss.getSheetByName("NamaGrup")) {
    const ng = ss.insertSheet("NamaGrup");
    ng.appendRow(["Nama Grup", "Sheet Chat", "Kolom", "Foto Grup"]);
    ng.appendRow(["Grup RT37", "ChatRT37", "G", ""]);
    ng.appendRow(["Grup RT38", "ChatRT38", "G", ""]);
    ng.appendRow(["Grup RT39", "ChatRT39", "G", ""]);
  }
  if (!ss.getSheetByName("Informasi")) {
    const inf = ss.insertSheet("Informasi");
    inf.appendRow(["Isi Pengumuman"]);
    inf.appendRow(["Diberitahukan kepada seluruh warga RT 37, 38, dan 39 untuk hadir dalam kerja bakti kebersihan lingkungan pada hari Minggu pukul 07.00 WIB di balai warga."]);
  }
  if (!ss.getSheetByName("ChatPrivat")) {
    const cp = ss.insertSheet("ChatPrivat");
    cp.appendRow(["Timestamp", "Username", "Nama", "Pesan", "Tipe", "MediaUrl", "StatusPesan", "TujuanUsername", "ReplyTo"]);
  }
  if (!ss.getSheetByName("PanggilanAktif")) {
    const pa = ss.insertSheet("PanggilanAktif");
    pa.appendRow(["TargetUsername", "CallerUsername", "CallerName", "Type", "Status", "RoomUrl", "Timestamp"]);
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
    announcements: getAnnouncements(),
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

function getUserAssignedGroups(username) {
  const ss = getSpreadsheet();
  const userSheet = ss.getSheetByName("Users");
  const groupSheet = ss.getSheetByName("NamaGrup");
  if (!userSheet || !groupSheet || userSheet.getLastRow() < 2) return [];

  const cleanUser = String(username || "").trim().toUpperCase();
  const userData = userSheet.getRange(2, 1, userSheet.getLastRow() - 1, 6).getValues();

  let userRoles = [];
  for (let i = 0; i < userData.length; i++) {
    if (String(userData[i][0] || "").trim().toUpperCase() === cleanUser) {
      const roleStr = String(userData[i][5] || "").trim();
      if (roleStr !== "") {
        userRoles = roleStr.split(",").map(r => r.trim().toUpperCase()).filter(Boolean);
      }
      break;
    }
  }
  if (userRoles.length === 0) return [];

  const groupData = groupSheet.getDataRange().getValues();
  let assignedRooms = [];
  for (let j = 1; j < groupData.length; j++) {
    const gName = String(groupData[j][0] || "").trim();
    const sName = String(groupData[j][1] || "").trim();
    const gFoto = String(groupData[j][3] || "").trim();
    if (gName !== "" && userRoles.includes(gName.toUpperCase())) {
      let lastMessage = "";
      let lastTime = "";
      const chatSheet = resolveChatSheet(ss, sName || gName);
      if (chatSheet && chatSheet.getLastRow() >= 2) {
        const lastRowVals = chatSheet.getRange(chatSheet.getLastRow(), 1, 1, 4).getValues()[0];
        lastTime = (lastRowVals[0] instanceof Date) ? lastRowVals[0].toISOString() : String(lastRowVals[0] || "");
        lastMessage = (lastRowVals[2] ? lastRowVals[2] + ": " : "") + String(lastRowVals[3] || "");
      }
      assignedRooms.push({
        name: gName,
        sheetName: sName || gName.replace(/^Grup\s*/i, "Chat").replace(/\s+/g, ""),
        foto: gFoto,
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
  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues();
  return data.map(r => String(r[0] || "").trim()).filter(Boolean);
}

function getGroupMembersByRole(groupName) {
  const userSheet = getSpreadsheet().getSheetByName("Users");
  if (!userSheet || userSheet.getLastRow() < 2) return [];

  const targetRole = String(groupName || "").trim().toUpperCase();
  const userData = userSheet.getRange(2, 1, userSheet.getLastRow() - 1, 6).getValues();
  let members = [];

  for (let r = 0; r < userData.length; r++) {
    const usernameUser = String(userData[r][0] || "").trim().toUpperCase();
    const namaUser = String(userData[r][1] || "").trim();
    const alamatUser = String(userData[r][2] || "").trim();
    const roleCell = String(userData[r][5] || "").trim();
    if (usernameUser && namaUser && roleCell) {
      const rolesArray = roleCell.split(",").map(item => item.trim().toUpperCase());
      if (rolesArray.includes(targetRole)) {
        members.push({ username: usernameUser, nama: namaUser, alamat: alamatUser });
      }
    }
  }
  return members;
}

function resolveChatSheet(ss, chatSheetName) {
  const raw = String(chatSheetName || "ChatRT37").trim();
  let sheet = ss.getSheetByName(raw);
  if (sheet) return sheet;

  // Fallback 1: Jika dikirim "Grup RT38", coba "ChatRT38"
  const altChat = raw.replace(/^Grup\s*/i, "Chat").replace(/\s+/g, "");
  sheet = ss.getSheetByName(altChat);
  if (sheet) return sheet;

  // Fallback 2: Cek tabel NamaGrup kolom A (Nama Grup) -> kolom B (Sheet Chat)
  const ng = ss.getSheetByName("NamaGrup");
  if (ng && ng.getLastRow() >= 2) {
    const ngData = ng.getRange(2, 1, ng.getLastRow() - 1, 2).getValues();
    for (let i = 0; i < ngData.length; i++) {
      const gName = String(ngData[i][0] || "").trim();
      const sName = String(ngData[i][1] || "").trim();
      if (gName.toUpperCase() === raw.toUpperCase() || sName.toUpperCase() === raw.toUpperCase()) {
        if (sName && ss.getSheetByName(sName)) return ss.getSheetByName(sName);
        if (gName && ss.getSheetByName(gName)) return ss.getSheetByName(gName);
      }
    }
  }

  // Fallback 3: Pencarian case-insensitive di seluruh tab sheet
  const allSheets = ss.getSheets();
  for (let j = 0; j < allSheets.length; j++) {
    if (allSheets[j].getName().trim().toUpperCase() === raw.toUpperCase()) {
      return allSheets[j];
    }
  }
  return null;
}

function sendMessage(chatSheetName, username, nama, pesan, tipe, mediaUrl, replyTo) {
  const lock = LockService.getScriptLock();
  lock.tryLock(10000);
  try {
    const ss = getSpreadsheet();
    const targetRoom = String(chatSheetName || "ChatRT37").trim();
    let sheet = resolveChatSheet(ss, targetRoom);
    if (!sheet) {
      const safeName = targetRoom.replace(/^Grup\s*/i, "Chat").replace(/\s+/g, "") || "ChatRT37";
      sheet = ss.insertSheet(safeName);
      sheet.appendRow(["Timestamp", "SenderUsername", "SenderNama", "Pesan", "Tipe", "MediaUrl", "ReplyTo", "ReadBy"]);
    }
    const cleanUser = String(username || "").trim().toUpperCase();
    sheet.appendRow([new Date(), cleanUser, nama || cleanUser, pesan, tipe || "text", mediaUrl || "", replyTo || "", cleanUser]);
    SpreadsheetApp.flush();
    return { success: true };
  } finally {
    lock.releaseLock();
  }
}

function getMessages(chatSheetName, readerUsername) {
  const ss = getSpreadsheet();
  const sheet = resolveChatSheet(ss, chatSheetName);
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

function votePollOption(chatSheetName, rowNumber, optionindex, voterUsername) {
  const ss = getSpreadsheet();
  const sheet = resolveChatSheet(ss, chatSheetName);
  if (!sheet || rowNumber < 2) return { success: false };
  const cell = sheet.getRange(rowNumber, 6);
  try {
    const pollData = JSON.parse(String(cell.getValue() || ""));
    if (!pollData.votes) pollData.votes = {};
    pollData.votes[String(voterUsername || "").trim().toUpperCase()] = optionindex;
    cell.setValue(JSON.stringify(pollData));
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
    sheet.appendRow([new Date(), cleanSender, nama || cleanSender, pesan, tipe || "text", mediaUrl || "", "BELUM", cleanTarget, replyTo || ""]);
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
  sheet.appendRow([
    String(targetUsername || "ALL").trim().toUpperCase(),
    String(callerUsername || "").trim().toUpperCase(),
    callerName || callerUsername,
    type || "voice",
    "RINGING",
    roomUrl || "",
    new Date()
  ]);
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