const SPREADSHEET_ID = "1kmZa08cDx5ts4Ld819mLsVmBJ83J-vZPu89i59TU3VI";
const UPLOAD_FOLDER_NAME = "SansChat_Media_Uploads_RW09";

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
    if (!e || !e.postData || !e.postData.contents) {
      return ContentService.createTextOutput(JSON.stringify({ status: "error", message: "Payload kosong" }))
        .setMimeType(ContentService.MimeType.JSON);
    }
    const payload = JSON.parse(e.postData.contents);
    const action = payload.action;
    let result = null;
    switch (action) {
      case 'loginUser': result = loginUser(payload.username, payload.password); break;
      case 'getInitialDashboardBundle': result = getInitialDashboardBundle(payload.username); break;
      case 'getUserProfileDetail': result = getUserProfileDetail(payload.username); break;
      case 'getAllWargaContacts': result = getAllWargaContacts(); break;
      case 'updateProfilePicture': result = updateProfilePicture(payload.username, payload.base64Data); break;
      case 'updateGroupProfilePicture': result = updateGroupProfilePicture(payload.sheetName, payload.base64Data); break;
      case 'uploadAttachmentToDrive': result = uploadAttachmentToDrive(payload.fileName, payload.mimeType, payload.base64Data, payload.username); break;
      case 'getAnnouncements': result = getAnnouncements(); break;
      case 'getUserAssignedGroups': result = getUserAssignedGroups(payload.username); break;
      case 'getNamaGrupMapping': result = getNamaGrupMapping(); break;
      case 'resolveGroupRoom': result = resolveGroupRoom(payload.groupName, payload.sheetName); break;
      case 'getGroupMembersByRole':
      case 'getGroupMembersByKolomC': result = getGroupMembersByKolomC(payload.groupName, payload.chatSheetName); break;
      case 'sendMessage': result = sendMessage(payload.chatSheetName, payload.username, payload.nama, payload.pesan, payload.tipe, payload.mediaUrl, payload.replyTo, payload.groupName, payload.fileName); break;
      case 'getMessages': result = getMessages(payload.chatSheetName, payload.username, payload.groupName); break;
      case 'votePollOption': result = votePollOption(payload.chatSheetName, payload.rowNumber, payload.optionindex, payload.voterUsername, payload.groupName); break;
      case 'sendPrivateMessage': result = sendPrivateMessage(payload.username, payload.nama, payload.pesan, payload.tipe, payload.mediaUrl, payload.targetUsername, payload.replyTo, payload.fileName); break;
      case 'getPrivateChatSheetRecords': result = getPrivateChatSheetRecords(payload.username); break;
      case 'hasUnreadPrivateMessages': result = hasUnreadPrivateMessages(payload.username); break;
      case 'markPrivateMessagesAsRead': markPrivateMessagesAsRead(payload.username); result = { success: true }; break;
      case 'initiateCall': result = initiateCall(payload.targetUsername, payload.callerUsername, payload.callerName, payload.type, payload.roomUrl); break;
      case 'checkIncomingCall': result = checkIncomingCall(payload.username); break;
      case 'updateCallStatus': updateCallStatus(payload.username, payload.status); result = { success: true }; break;
      default: result = { success: false, message: "Action tidak dikenali: " + action };
    }
    return ContentService.createTextOutput(JSON.stringify({ status: "ok", data: result }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ status: "error", message: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function getSpreadsheet() { return SpreadsheetApp.openById(SPREADSHEET_ID); }

function getOrCreateUploadFolder() {
  const folders = DriveApp.getFoldersByName(UPLOAD_FOLDER_NAME);
  if (folders.hasNext()) return folders.next();
  const folder = DriveApp.createFolder(UPLOAD_FOLDER_NAME);
  try { folder.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); } catch (e) {}
  return folder;
}

function uploadAttachmentToDrive(fileName, mimeType, base64Data, username) {
  try {
    if (!base64Data) return { success: false, message: "Data file kosong" };
    let cleanBase64 = String(base64Data);
    let detectedMime = String(mimeType || "").trim();
    if (cleanBase64.indexOf("base64,") !== -1) {
      const parts = cleanBase64.split("base64,");
      const metaPart = parts[0] || "";
      cleanBase64 = parts[1] || "";
      if (!detectedMime) {
        const m = metaPart.match(/data:([^;]+);/i);
        if (m && m[1]) detectedMime = m[1].trim();
      }
    }
    if (!detectedMime) detectedMime = "application/octet-stream";
    const decoded = Utilities.base64Decode(cleanBase64);
    const safeName = (username ? String(username).toUpperCase() + "_" : "") +
      Date.now() + "_" + String(fileName || "lampiran").replace(/[^a-zA-Z0-9._-]/g, "_");
    const blob = Utilities.newBlob(decoded, detectedMime, safeName);
    const folder = getOrCreateUploadFolder();
    const file = folder.createFile(blob);
    try { file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); } catch (e) {}
    const fileId = file.getId();
    const viewUrl = "https://drive.google.com/file/d/" + fileId + "/view?usp=sharing";
    const previewUrl = "https://drive.google.com/file/d/" + fileId + "/preview";
    const directImgUrl = detectedMime.startsWith("image/")
      ? "https://lh3.googleusercontent.com/d/" + fileId
      : previewUrl;
    return {
      success: true,
      fileId: fileId,
      fileName: file.getName(),
      mimeType: detectedMime,
      viewUrl: viewUrl,
      previewUrl: previewUrl,
      mediaUrl: directImgUrl
    };
  } catch (err) {
    return { success: false, message: "Gagal upload ke Drive: " + err.toString() };
  }
}

function ensureDefaultSheets() {
  const ss = getSpreadsheet();
  if (!ss.getSheetByName("Users")) {
    const u = ss.insertSheet("Users");
    u.appendRow(["Username", "Nama", "Alamat", "Password", "Foto", "Role", "Grup RT37", "Grup RT38", "Grup RT39", "Grup RW09", "Grup Umum"]);
  }
  if (!ss.getSheetByName("NamaGrup")) {
    const ng = ss.insertSheet("NamaGrup");
    ng.appendRow(["Nama Grup", "Sheet Chat", "Kolom", "Foto Grup", "Deskripsi", "Link"]);
    ng.appendRow(["Grup Umum", "ChatUmum", "K", "", "Grup obrolan umum warga RW 09", ""]);
  }
}

function getHelpSystemUrl() { const s = getSpreadsheet().getSheetByName("Users"); return s ? String(s.getRange("S1").getValue() || "").trim() : ""; }
function getWebAppsUrl() { const s = getSpreadsheet().getSheetByName("Users"); return s ? String(s.getRange("S2").getValue() || "").trim() : ""; }
function getSansChatUrl() { const s = getSpreadsheet().getSheetByName("Users"); return s ? String(s.getRange("S3").getValue() || "").trim() : ""; }
function getScriptUrl() { const s3 = getSansChatUrl(); return (s3 && s3.startsWith("http")) ? s3 : ScriptApp.getService().getUrl(); }

function loginUser(username, password) {
  const sheet = getSpreadsheet().getSheetByName("Users");
  if (!sheet || sheet.getLastRow() < 2) return { success: false, message: "Data Users masih kosong" };
  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 6).getValues();
  const cleanUser = String(username || "").trim().toUpperCase();
  const cleanPass = String(password || "").trim();
  for (let i = 0; i < data.length; i++) {
    if (String(data[i][0] || "").trim().toUpperCase() === cleanUser && String(data[i][3] || "").trim() === cleanPass) {
      return {
        success: true,
        user: {
          username: cleanUser,
          nama: String(data[i][1] || ""),
          alamat: String(data[i][2] || ""),
          foto: String(data[i][4] || ""),
          role: String(data[i][5] || "")
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
    groups: getUserAssignedGroups(cleanUser)
  };
}

function getUserProfileDetail(username) {
  const sheet = getSpreadsheet().getSheetByName("Users");
  if (!sheet || sheet.getLastRow() < 2) return null;
  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 6).getValues();
  const cleanUser = String(username || "").trim().toUpperCase();
  for (let i = 0; i < data.length; i++) {
    if (String(data[i][0] || "").trim().toUpperCase() === cleanUser) {
      return { username: cleanUser, nama: String(data[i][1] || ""), alamat: String(data[i][2] || ""), foto: String(data[i][4] || ""), role: String(data[i][5] || "") };
    }
  }
  return null;
}

function getUserAssignedGroups(username) {
  const ss = getSpreadsheet();
  const ng = ss.getSheetByName("NamaGrup");
  if (!ng || ng.getLastRow() < 2) return [{ name: "Grup Umum", sheetName: "ChatUmum", link: "" }];
  const lastCol = Math.max(2, ng.getLastColumn());
  const rows = ng.getRange(2, 1, ng.getLastRow() - 1, lastCol).getValues();
  return rows.filter(r => String(r[0] || "").trim() !== "").map(r => {
    let linkVal = "";
    for (let c = 2; c < r.length; c++) {
      const cellStr = String(r[c] || "").trim();
      if (/^https?:\/\//i.test(cellStr)) {
        linkVal = cellStr;
        break;
      }
    }
    return {
      name: String(r[0]).trim(),
      sheetName: String(r[1] || "ChatUmum").trim(),
      link: linkVal
    };
  });
}

function sendMessage(chatSheetName, username, nama, pesan, tipe, mediaUrl, replyTo, groupName, fileName) {
  const ss = getSpreadsheet();
  const safeSheet = String(chatSheetName || "ChatUmum").trim();
  let sheet = ss.getSheetByName(safeSheet) || ss.insertSheet(safeSheet);
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(["Timestamp", "SenderUsername", "SenderNama", "Pesan", "Tipe", "MediaUrl", "ReplyTo", "ReadBy"]);
  }
  let finalMedia = String(mediaUrl || "");
  if (finalMedia.startsWith("data:") && finalMedia.length > 45000) {
    const up = uploadAttachmentToDrive(fileName || ("upload_" + (tipe || "file")), tipe || "file", finalMedia, username);
    if (up && up.success && up.mediaUrl) finalMedia = up.mediaUrl;
  }
  sheet.appendRow([new Date(), String(username || "").toUpperCase(), nama || username, pesan, tipe || "text", finalMedia, replyTo || "", String(username || "").toUpperCase()]);
  return { success: true, mediaUrl: finalMedia };
}

function getMessages(chatSheetName, readerUsername, groupName) {
  const ss = getSpreadsheet();
  const sheet = ss.getSheetByName(String(chatSheetName || "ChatUmum").trim());
  if (!sheet || sheet.getLastRow() < 2) return [];
  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, Math.max(6, sheet.getLastColumn())).getValues();
  return data.map((r, idx) => ({
    rowNumber: idx + 2,
    timestamp: (r[0] instanceof Date) ? r[0].toISOString() : String(r[0] || ""),
    senderUsername: String(r[1] || ""),
    senderNama: String(r[2] || ""),
    pesan: String(r[3] || ""),
    tipe: String(r[4] || "text"),
    mediaUrl: String(r[5] || "")
  }));
}

function sendPrivateMessage(username, nama, pesan, tipe, mediaUrl, targetUsername, replyTo, fileName) {
  const ss = getSpreadsheet();
  let sheet = ss.getSheetByName("ChatPrivat") || ss.insertSheet("ChatPrivat");
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(["Timestamp", "Username", "Nama", "Pesan", "Tipe", "MediaUrl", "StatusPesan", "TujuanUsername", "ReplyTo"]);
  }
  let finalMedia = String(mediaUrl || "");
  if (finalMedia.startsWith("data:") && finalMedia.length > 45000) {
    const up = uploadAttachmentToDrive(fileName || ("privat_" + (tipe || "file")), tipe || "file", finalMedia, username);
    if (up && up.success && up.mediaUrl) finalMedia = up.mediaUrl;
  }
  sheet.appendRow([new Date(), String(username || "").toUpperCase(), nama || username, pesan, tipe || "text", finalMedia, "BELUM", String(targetUsername || username).toUpperCase(), replyTo || ""]);
  return { success: true, mediaUrl: finalMedia };
}

function getPrivateChatSheetRecords(username) {
  const sheet = getSpreadsheet().getSheetByName("ChatPrivat");
  if (!sheet || sheet.getLastRow() < 2) return [];
  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, Math.max(8, sheet.getLastColumn())).getValues();
  const cleanUser = String(username || "").trim().toUpperCase();
  return data.filter(r => String(r[1] || "").toUpperCase() === cleanUser || String(r[7] || "").toUpperCase() === cleanUser)
    .map((r, idx) => ({
      rowNumber: idx + 2,
      timestamp: (r[0] instanceof Date) ? r[0].toISOString() : String(r[0] || ""),
      pengirimUsername: String(r[1] || ""),
      nama: String(r[2] || ""),
      pesanD: String(r[3] || ""),
      pesanF: String(r[5] || ""),
      tipe: String(r[4] || "text")
    }));
}

function initiateCall(targetUsername, callerUsername, callerName, type, roomUrl) {
  const ss = getSpreadsheet();
  let sheet = ss.getSheetByName("PanggilanAktif") || ss.insertSheet("PanggilanAktif");
  if (sheet.getLastRow() === 0) sheet.appendRow(["TargetUsername", "CallerUsername", "CallerName", "Type", "Status", "RoomUrl", "Timestamp"]);
  sheet.appendRow([String(targetUsername || "ALL").toUpperCase(), String(callerUsername || "").toUpperCase(), callerName || callerUsername, type || "voice", "RINGING", roomUrl || "", new Date()]);
  return { success: true };
}
