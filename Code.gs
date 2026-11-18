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
      case 'loginUserWithBundle':
        result = loginUserWithBundle(payload.username, payload.password);
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
        result = updateGroupProfilePicture(payload.room || payload.sheetName, payload.base64Data);
        break;
      case 'sendMessage':
        result = sendMessage(payload.room || payload.chatSheetName, payload.username, payload.nama || payload.username, payload.pesan || payload.message, payload.tipe || payload.type, payload.mediaUrl, payload.replyTo);
        break;
      case 'getMessages':
        result = getMessages(payload.room || payload.chatSheetName, payload.username);
        break;
      case 'votePollOption':
        result = votePollOption(payload.chatSheetName, payload.rowNumber, payload.optionindex, payload.voterUsername);
        break;
      case 'sendPrivateMessage':
        result = sendPrivateMessage(payload.username || payload.sender, payload.nama || payload.sender, payload.pesan || payload.message, payload.tipe || payload.type, payload.mediaUrl, payload.targetUsername || payload.receiver, payload.replyTo);
        break;
      case 'getPrivateChatSheetRecords':
      case 'getPrivateMessages':
        result = getPrivateChatSheetRecords(payload.username || payload.user1);
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
      case 'getAnnouncements':
        result = getAnnouncements();
        break;
      case 'postAnnouncement':
        result = postAnnouncement(payload.username, payload.title, payload.content);
        break;
      case 'getGroupMembersByRole':
        result = getGroupMembersByRole(payload.groupName);
        break;
      default:
        result = { success: false, message: 'Aksi tidak dikenal: ' + action };
        break;
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
  const required = {
    "Warga": ["Username", "Password", "Nama", "RT", "RW", "Role", "Telepon", "FotoProfile"],
    "NamaGrup": ["Nama Grup", "Sheet Chat", "Kolom", "Foto Grup"],
    "ChatUmum": ["Timestamp", "SenderUsername", "SenderNama", "Pesan", "Tipe", "MediaUrl", "ReplyTo", "ReadBy"],
    "ChatRT37": ["Timestamp", "SenderUsername", "SenderNama", "Pesan", "Tipe", "MediaUrl", "ReplyTo", "ReadBy"],
    "ChatRT38": ["Timestamp", "SenderUsername", "SenderNama", "Pesan", "Tipe", "MediaUrl", "ReplyTo", "ReadBy"],
    "ChatRT39": ["Timestamp", "SenderUsername", "SenderNama", "Pesan", "Tipe", "MediaUrl", "ReplyTo", "ReadBy"],
    "ChatRW09": ["Timestamp", "SenderUsername", "SenderNama", "Pesan", "Tipe", "MediaUrl", "ReplyTo", "ReadBy"],
    "ChatSinoman": ["Timestamp", "SenderUsername", "SenderNama", "Pesan", "Tipe", "MediaUrl", "ReplyTo", "ReadBy"],
    "ChatPrivat": ["Timestamp", "Username", "Nama", "Pesan", "Tipe", "MediaUrl", "StatusPesan", "TujuanUsername", "ReplyTo"],
    "PanggilanAktif": ["TargetUsername", "CallerUsername", "CallerName", "Type", "Status", "RoomUrl", "Timestamp"],
    "LogAktifitas": ["Timestamp", "Username", "Aktivitas", "Detail"],
    "Informasi": ["Isi Pengumuman"],
    "Pengumuman": ["Timestamp", "Sender", "Title", "Content"]
  };

  for (let sheetName in required) {
    let sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
      sheet = ss.insertSheet(sheetName);
      sheet.appendRow(required[sheetName]);
      if (sheetName === "Warga" && sheet.getLastRow() < 2) {
        sheet.appendRow(["ADMINPUSAT", "admin123", "Administrator Pusat", "01", "09", "ADMINPUSAT,PENGURUS", "+628123456789", ""]);
      }
      if (sheetName === "NamaGrup" && sheet.getLastRow() < 2) {
        sheet.appendRow(["Grup Umum", "ChatUmum", "G", ""]);
        sheet.appendRow(["Grup RT37", "ChatRT37", "G", ""]);
        sheet.appendRow(["Grup RT38", "ChatRT38", "G", ""]);
        sheet.appendRow(["Grup RT39", "ChatRT39", "G", ""]);
        sheet.appendRow(["Grup RW09", "ChatRW09", "G", ""]);
        sheet.appendRow(["Grup Sinoman", "ChatSinoman", "G", ""]);
      }
      if (sheetName === "Informasi" && sheet.getLastRow() < 2) {
        sheet.appendRow(["Diberitahukan kepada seluruh warga RW 09 untuk aktif memantau grup dan portal resmi."]);
      }
    }
  }
}

function logActivity(username, aktivitas, detail) {
  try {
    const ss = getSpreadsheet();
    const sheet = ss.getSheetByName("LogAktifitas");
    if (sheet) {
      sheet.appendRow([new Date().toLocaleString(), username || "SYSTEM", aktivitas, detail]);
    }
  } catch (e) {}
}

function getHelpSystemUrl() { return ScriptApp.getService().getUrl() + "?page=login"; }
function getWebAppsUrl() { return ScriptApp.getService().getUrl() + "?page=utama"; }
function getSansChatUrl() { return ScriptApp.getService().getUrl(); }

function loginUser(username, password) {
  const cleanUser = String(username || "").trim().toUpperCase();
  const cleanPass = String(password || "").trim();
  
  if (!cleanUser || !cleanPass) {
    return { success: false, message: "Username dan Password wajib diisi!" };
  }

  const ss = getSpreadsheet();
  const sheet = ss.getSheetByName("Warga");
  if (!sheet) return { success: false, message: "Sheet Warga tidak ditemukan." };

  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const uName = String(row[0] || "").trim().toUpperCase();
    const uPass = String(row[1] || "").trim();
    if (uName === cleanUser && uPass === cleanPass) {
      logActivity(cleanUser, "LOGIN_SUCCESS", "Berhasil masuk ke aplikasi SansChat");
      return {
        success: true,
        redirectBaseUrl: getSansChatUrl(),
        user: {
          username: row[0],
          nama: row[2],
          alamat: "RT " + (row[3] || "-") + " / RW " + (row[4] || "-"),
          role: row[5],
          telepon: row[6],
          foto: row[7] || "",
          helpSystemUrl: getHelpSystemUrl(),
          webAppsUrl: getWebAppsUrl(),
          sansChatUrl: getSansChatUrl()
        }
      };
    }
  }

  logActivity(cleanUser, "LOGIN_FAILED", "Kombinasi Username/Password salah");
  return { success: false, message: "Username atau Password salah!" };
}

function loginUserWithBundle(username, password) {
  const auth = loginUser(username, password);
  if (!auth || !auth.success) return auth;
  const cleanUser = auth.user.username;
  auth.bundle = getInitialDashboardBundle(cleanUser);
  return auth;
}

function getInitialDashboardBundle(username) {
  ensureDefaultSheets();
  const cleanUser = String(username || "").trim().toUpperCase();
  logActivity(cleanUser, "LOAD_DASHBOARD", "Memuat dashboard utama SansChat");
  return {
    success: true,
    profile: getUserProfileDetail(cleanUser),
    groups: getUserAssignedGroups(cleanUser),
    groupMappings: getGroupMappings(),
    announcements: getAnnouncements(),
    hasUnreadPrivate: hasUnreadPrivateMessages(cleanUser),
    incomingCall: checkIncomingCall(cleanUser)
  };
}

function getUserProfileDetail(username) {
  const cleanUser = String(username || "").trim().toUpperCase();
  const ss = getSpreadsheet();
  const sheet = ss.getSheetByName("Warga");
  if (!sheet) return null;

  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0] || "").trim().toUpperCase() === cleanUser) {
      return {
        username: data[i][0],
        nama: data[i][2],
        alamat: "RT " + (data[i][3] || "-") + " / RW " + (data[i][4] || "-"),
        role: data[i][5],
        telepon: data[i][6],
        foto: data[i][7] || "",
        helpSystemUrl: getHelpSystemUrl(),
        webAppsUrl: getWebAppsUrl(),
        sansChatUrl: getSansChatUrl()
      };
    }
  }
  return null;
}

function getAllWargaContacts() {
  const ss = getSpreadsheet();
  const sheet = ss.getSheetByName("Warga");
  if (!sheet || sheet.getLastRow() < 2) return [];

  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 8).getValues();
  return data.map(r => ({
    username: String(r[0] || "").trim().toUpperCase(),
    nama: String(r[2] || r[0] || "").trim(),
    alamat: "RT " + String(r[3] || "-") + " / RW " + String(r[4] || "-")
  })).filter(r => r.username && r.nama);
}

function getGroupMappings() {
  const ss = getSpreadsheet();
  const ng = ss.getSheetByName("NamaGrup");
  if (!ng || ng.getLastRow() < 2) return [];
  const data = ng.getRange(2, 1, ng.getLastRow() - 1, ng.getLastColumn()).getValues();
  const mappings = [];
  for (let i = 0; i < data.length; i++) {
    const colA = String(data[i][0] || "").trim();
    const colB = String(data[i][1] || "").trim();
    const colC = String(data[i][2] || "").trim();
    const colD = String(data[i][3] || "").trim();
    if (colA || colB) {
      mappings.push({
        groupName: colA,
        sheetName: colB || colA.replace(/^Grup\s*/i, "Chat").replace(/\s+/g, ""),
        colSetting: colC,
        foto: colD
      });
    }
  }
  return mappings;
}

function getUserAssignedGroups(username) {
  const ss = getSpreadsheet();
  const userSheet = ss.getSheetByName("Warga");
  const groupSheet = ss.getSheetByName("NamaGrup");
  if (!userSheet || !groupSheet) return [];

  const cleanUser = String(username || "").trim().toUpperCase();
  const userData = userSheet.getDataRange().getValues();

  let userRoles = [];
  for (let i = 1; i < userData.length; i++) {
    if (String(userData[i][0] || "").trim().toUpperCase() === cleanUser) {
      const roleStr = String(userData[i][5] || "").trim();
      if (roleStr !== "") {
        userRoles = roleStr.split(",").map(r => r.trim().toUpperCase()).filter(Boolean);
      }
      break;
    }
  }

  const isSuperAdmin = (cleanUser === "ADMINPUSAT" || userRoles.includes("ADMINPUSAT") || userRoles.includes("PENGURUS"));
  const groupData = groupSheet.getDataRange().getValues();
  let assignedRooms = [];

  for (let j = 1; j < groupData.length; j++) {
    const gName = String(groupData[j][0] || "").trim();
    const sName = String(groupData[j][1] || "").trim();
    const gColSetting = String(groupData[j][2] || "").trim();
    const gFoto = String(groupData[j][3] || "").trim();

    if (gName !== "" && (isSuperAdmin || userRoles.includes(gName.toUpperCase()) || (sName && userRoles.includes(sName.toUpperCase())))) {
      let lastMessage = "";
      let lastTime = "";
      const targetSheetName = lookupTargetSheetNameFromNamaGrup(ss, sName, gName);
      const chatSheet = ss.getSheetByName(targetSheetName);

      if (chatSheet && chatSheet.getLastRow() >= 2) {
        const lastRowVals = chatSheet.getRange(chatSheet.getLastRow(), 1, 1, 8).getValues()[0];
        lastTime = (lastRowVals[0] instanceof Date) ? lastRowVals[0].toISOString() : String(lastRowVals[0] || "");
        lastMessage = (lastRowVals[2] ? lastRowVals[2] + ": " : "") + String(lastRowVals[3] || "");
      }
      assignedRooms.push({
        name: gName,
        sheetName: targetSheetName,
        colSetting: gColSetting,
        foto: gFoto,
        lastMessage: lastMessage,
        lastTime: lastTime
      });
    }
  }
  return assignedRooms;
}

function updateProfilePicture(username, base64Data) {
  const sheet = getSpreadsheet().getSheetByName("Warga");
  if (!sheet) return { success: false };
  const data = sheet.getDataRange().getValues();
  const cleanUser = String(username || "").trim().toUpperCase();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0] || "").trim().toUpperCase() === cleanUser) {
      sheet.getRange(i + 1, 8).setValue(base64Data);
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
  if (!sheet || sheet.getLastRow() < 2) return ["Selamat datang di Portal Komunikasi Warga RW 09 Sansivera."];
  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues();
  return data.map(r => String(r[0] || "").trim()).filter(Boolean);
}

function postAnnouncement(username, title, content) {
  const ss = getSpreadsheet();
  let sheet = ss.getSheetByName("Pengumuman");
  if (!sheet) {
    sheet = ss.insertSheet("Pengumuman");
    sheet.appendRow(["Timestamp", "Sender", "Title", "Content"]);
  }
  sheet.appendRow([new Date().toLocaleString(), username, title, content]);
  return { success: true };
}

function getGroupMembersByRole(groupName) {
  const userSheet = getSpreadsheet().getSheetByName("Warga");
  if (!userSheet || userSheet.getLastRow() < 2) return [];

  const targetRole = String(groupName || "").trim().toUpperCase();
  const userData = userSheet.getRange(2, 1, userSheet.getLastRow() - 1, 8).getValues();
  let members = [];

  for (let r = 0; r < userData.length; r++) {
    const usernameUser = String(userData[r][0] || "").trim().toUpperCase();
    const namaUser = String(userData[r][2] || "").trim();
    const alamatUser = "RT " + String(userData[r][3] || "-") + " / RW " + String(userData[r][4] || "-");
    const roleCell = String(userData[r][5] || "").trim();
    if (usernameUser && namaUser) {
      const rolesArray = roleCell ? roleCell.split(",").map(item => item.trim().toUpperCase()) : [];
      if (rolesArray.includes(targetRole) || targetRole.includes("GRUP") || targetRole === "ADMINPUSAT" || true) {
        members.push({ username: usernameUser, nama: namaUser, alamat: alamatUser });
      }
    }
  }
  return members;
}

function lookupTargetSheetNameFromNamaGrup(ss, rawSheetOrGroup, fallbackGroupName) {
  const q1 = String(rawSheetOrGroup || "").trim();
  const q2 = String(fallbackGroupName || "").trim();
  const ng = ss.getSheetByName("NamaGrup");

  if (ng && ng.getLastRow() >= 2) {
    const ngData = ng.getRange(2, 1, ng.getLastRow() - 1, ng.getLastColumn()).getValues();
    for (let i = 0; i < ngData.length; i++) {
      const colA = String(ngData[i][0] || "").trim();
      const colB = String(ngData[i][1] || "").trim();

      if ((q1 && (colA.toUpperCase() === q1.toUpperCase() || colB.toUpperCase() === q1.toUpperCase())) ||
          (q2 && (colA.toUpperCase() === q2.toUpperCase() || colB.toUpperCase() === q2.toUpperCase()))) {
        if (colB && ss.getSheetByName(colB)) return colB;
        if (colA && ss.getSheetByName(colA)) return colA;
      }
    }
  }

  const directCandidate = q1 || q2;
  if (directCandidate && ss.getSheetByName(directCandidate)) return directCandidate;

  const altName = directCandidate.replace(/^Grup\s*/i, "Chat").replace(/\s+/g, "") || "ChatUmum";
  if (ss.getSheetByName(altName)) return altName;

  return directCandidate || "ChatUmum";
}

function resolveChatSheet(ss, chatSheetName, groupName) {
  const mappedSheetName = lookupTargetSheetNameFromNamaGrup(ss, chatSheetName, groupName);
  let sheet = ss.getSheetByName(mappedSheetName);
  if (sheet) return sheet;

  const safeName = mappedSheetName.replace(/^Grup\s*/i, "Chat").replace(/\s+/g, "") || "ChatUmum";
  sheet = ss.insertSheet(safeName);
  sheet.appendRow(["Timestamp", "SenderUsername", "SenderNama", "Pesan", "Tipe", "MediaUrl", "ReplyTo", "ReadBy"]);
  return sheet;
}

function sendMessage(chatSheetName, username, nama, pesan, tipe, mediaUrl, replyTo) {
  const lock = LockService.getScriptLock();
  lock.tryLock(10000);
  try {
    const ss = getSpreadsheet();
    const mappedName = lookupTargetSheetNameFromNamaGrup(ss, chatSheetName, "");
    let sheet = resolveChatSheet(ss, mappedName, "");
    const cleanUser = String(username || "").trim().toUpperCase();
    const cleanNama = String(nama || cleanUser).trim();
    sheet.appendRow([new Date(), cleanUser, cleanNama, pesan || "", tipe || "text", mediaUrl || "", replyTo || "", cleanUser]);
    SpreadsheetApp.flush();
    return { success: true, resolvedSheet: sheet.getName() };
  } finally {
    lock.releaseLock();
  }
}

function getMessages(chatSheetName, readerUsername) {
  const ss = getSpreadsheet();
  const sheet = resolveChatSheet(ss, chatSheetName, "");
  if (!sheet || sheet.getLastRow() < 2) return [];

  const lastRow = sheet.getLastRow();
  const numCols = Math.max(8, sheet.getLastColumn());
  const data = sheet.getRange(2, 1, lastRow - 1, numCols).getValues();
  const cleanReader = String(readerUsername || "").trim().toUpperCase();
  let msgs = [];

  for (let i = 0; i < data.length; i++) {
    let rawTime = data[i][0];
    let timeStr = (rawTime instanceof Date) ? rawTime.toISOString() : String(rawTime || "");
    let colBUser = String(data[i][1] || "").trim();
    let colCSender = String(data[i][2] || "").trim();
    let colDPesan = String(data[i][3] || "").trim();
    let colETipe = String(data[i][4] || "text").trim();
    let colFMedia = String(data[i][5] || "").trim();
    let colGReply = String(data[i][6] || "").trim();

    if (!colCSender && !colDPesan && !colFMedia) continue;

    let readByStr = String(data[i][7] || "");
    let readList = readByStr ? readByStr.split(",").map(s => s.trim().toUpperCase()).filter(Boolean) : [];

    if (cleanReader && !readList.includes(cleanReader)) {
      readList.push(cleanReader);
      sheet.getRange(i + 2, 8).setValue(readList.join(","));
    }

    msgs.push({
      rowNumber: i + 2,
      resolvedSheet: sheet.getName(),
      timestamp: timeStr,
      senderUsername: colBUser || colCSender.toUpperCase(),
      senderNama: colCSender || colBUser,
      pesan: colDPesan,
      tipe: colETipe || "text",
      mediaUrl: colFMedia,
      replyTo: colGReply,
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
    sheet.appendRow([new Date(), cleanSender, nama || cleanSender, pesan || "", tipe || "text", mediaUrl || "", "BELUM", cleanTarget, replyTo || ""]);
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