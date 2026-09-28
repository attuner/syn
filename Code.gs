/**
 * Community Radio Backend - Google Apps Script
 * Hourly Named Slots + Chronological Sequence + Random Filler + Instant Push
 */

const FOLDER_NAME = "Community Radio";
const ADMIN_SECRET_KEY = "admin123";

function doPost(e) {
  try {
    const contents = JSON.parse(e.postData.contents);
    const action = contents.action;
    let response = { success: false, message: "Invalid action" };

    if (action === "register") response = handleRegister(contents);
    else if (action === "login") response = handleLogin(contents);
    else if (action === "uploadAudio") response = handleAudioUpload(contents);
    else if (action === "adminLogin") response = handleAdminLogin(contents);
    else if (action === "adminUpdateUserStatus") response = handleAdminUpdateUser(contents);
    else if (action === "adminDeleteTrack") response = handleDeleteTrack(contents);
    else if (action === "adminSaveSettings") response = handleSaveSettings(contents);
    else if (action === "heartbeat") response = handleHeartbeat(contents);
    else if (action === "saveHourlySchedule") response = handleSaveHourlySchedule(contents);
    else if (action === "adminPushTrack") response = handlePushTrack(contents);

    return ContentService.createTextOutput(JSON.stringify(response))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({ success: false, message: error.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function doGet(e) {
  const action = e.parameter.action;

  // DIRECT AUDIO STREAMER (Bypasses Google Drive 403 & CORS)
  if (action === "streamAudio") {
    try {
      const fileId = e.parameter.fileId;
      const file = DriveApp.getFileById(fileId);
      const mime = file.getMimeType() || "audio/mpeg";
      const bytes = file.getBlob().getBytes();
      const b64 = Utilities.base64Encode(bytes);
      const dataUri = "data:" + mime + ";base64," + b64;
      
      return ContentService.createTextOutput(JSON.stringify({ success: true, dataUri: dataUri }))
        .setMimeType(ContentService.MimeType.JSON);
    } catch(err) {
      return ContentService.createTextOutput(JSON.stringify({ success: false, error: err.toString() }))
        .setMimeType(ContentService.MimeType.JSON);
    }
  }

  let response = { success: false, message: "Invalid request" };

  if (action === "getStationData") {
    response = getStationData();
  } else if (action === "getAdminData") {
    if (e.parameter.key === ADMIN_SECRET_KEY) {
      response = getAdminData();
    } else {
      response = { success: false, message: "Unauthorized admin access" };
    }
  }

  return ContentService.createTextOutput(JSON.stringify(response))
    .setMimeType(ContentService.MimeType.JSON);
}

// ----------------- SHEET & DRIVE HELPERS -----------------

function getOrCreateFolder() {
  const folders = DriveApp.getFoldersByName(FOLDER_NAME);
  if (folders.hasNext()) return folders.next();
  const folder = DriveApp.createFolder(FOLDER_NAME);
  folder.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return folder;
}

function getSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  let usersSheet = ss.getSheetByName("Users");
  if (!usersSheet) {
    usersSheet = ss.insertSheet("Users");
    usersSheet.appendRow(["Username", "Mobile", "Password", "Status", "RegisteredAt"]);
  }

  let tracksSheet = ss.getSheetByName("Tracks");
  if (!tracksSheet) {
    tracksSheet = ss.insertSheet("Tracks");
    tracksSheet.appendRow(["Seq", "Title", "Description", "Contributor", "FileId", "StreamUrl", "CreatedAt"]);
  }

  let settingsSheet = ss.getSheetByName("Settings");
  if (!settingsSheet) {
    settingsSheet = ss.insertSheet("Settings");
    settingsSheet.appendRow(["Key", "Value"]);
    settingsSheet.appendRow(["station_name", "Community Voice FM"]);
    settingsSheet.appendRow(["broadcast_mode", "24/7 Live Broadcast"]);
    settingsSheet.appendRow(["seq_version", String(Date.now())]);
    settingsSheet.appendRow(["pushed_track", ""]);
    settingsSheet.appendRow(["push_version", ""]);
  }

  let listenersSheet = ss.getSheetByName("Listeners");
  if (!listenersSheet) {
    listenersSheet = ss.insertSheet("Listeners");
    listenersSheet.appendRow(["ListenerId", "Username", "Status", "LastSeen"]);
  }

  let hourlySheet = ss.getSheetByName("HourlySchedule");
  if (!hourlySheet) {
    hourlySheet = ss.insertSheet("HourlySchedule");
    hourlySheet.appendRow(["Hour", "SlotName", "TrackIdsJson"]);
  }

  return { usersSheet, tracksSheet, settingsSheet, listenersSheet, hourlySheet };
}

// ----------------- USER AUTHENTICATION -----------------

function handleRegister(data) {
  const { username, mobile, password } = data;
  if (!username || !mobile || !password) return { success: false, message: "All fields are required" };

  const { usersSheet } = getSheets();
  const rows = usersSheet.getDataRange().getValues();

  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][1]).trim() === String(mobile).trim()) {
      return { success: false, message: "Mobile number already registered" };
    }
  }

  usersSheet.appendRow([username.trim(), String(mobile).trim(), String(password).trim(), "pending", new Date().toISOString()]);
  return {
    success: true,
    message: "Registration successful. Contributor approval is pending.",
    user: { username: username.trim(), mobile: String(mobile).trim(), status: "pending" }
  };
}

function handleLogin(data) {
  const { mobile, password } = data;
  const { usersSheet } = getSheets();
  const rows = usersSheet.getDataRange().getValues();

  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][1]).trim() === String(mobile).trim() && String(rows[i][2]).trim() === String(password).trim()) {
      return {
        success: true,
        user: {
          username: rows[i][0],
          mobile: rows[i][1],
          status: rows[i][3]
        }
      };
    }
  }
  return { success: false, message: "Invalid mobile number or password" };
}

// ----------------- AUDIO UPLOAD -----------------
function handleAudioUpload(data) {
  const { mobile, username, title, description, base64File, fileName, mimeType } = data;
  const { usersSheet, tracksSheet } = getSheets();

  const userRows = usersSheet.getDataRange().getValues();
  let isApproved = false;
  for (let i = 1; i < userRows.length; i++) {
    if (String(userRows[i][1]).trim() === String(mobile).trim() && userRows[i][3] === "approved") {
      isApproved = true;
      break;
    }
  }

  if (!isApproved) {
    return { success: false, message: "User is not approved by Admin to upload audio." };
  }

  const folder = getOrCreateFolder();
  const decodedData = Utilities.base64Decode(base64File);
  const blob = Utilities.newBlob(decodedData, mimeType || "audio/mpeg", fileName || `${title}.mp3`);
  const file = folder.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

  const fileId = file.getId();
  const streamUrl = `https://drive.google.com/uc?export=download&id=${fileId}`;
  const createdAt = new Date().toISOString();

  const trackRows = tracksSheet.getDataRange().getValues();
  const nextSeq = trackRows.length; // chronological index

  tracksSheet.appendRow([nextSeq, title, description, username, fileId, streamUrl, createdAt]);

  return { success: true, message: "Audio uploaded successfully!", track: { seq: nextSeq, title, description, fileId, streamUrl, createdAt } };
}

// ----------------- HEARTBEAT & LISTENERS -----------------
function handleHeartbeat(data) {
  const { listenerId, username, isPlaying } = data;
  const { listenersSheet } = getSheets();
  const rows = listenersSheet.getDataRange().getValues();
  const now = Date.now();
  let found = false;

  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]).trim() === String(listenerId).trim()) {
      listenersSheet.getRange(i + 1, 2).setValue(username || "Guest Listener");
      listenersSheet.getRange(i + 1, 3).setValue(isPlaying ? "Listening Live" : "Idle");
      listenersSheet.getRange(i + 1, 4).setValue(now);
      found = true;
      break;
    }
  }

  if (!found) {
    listenersSheet.appendRow([listenerId, username || "Guest Listener", isPlaying ? "Listening Live" : "Idle", now]);
  }

  let activeCount = 0;
  const updatedRows = listenersSheet.getDataRange().getValues();
  for (let i = 1; i < updatedRows.length; i++) {
    const lastSeen = Number(updatedRows[i][3]) || 0;
    if (now - lastSeen < 50000 && updatedRows[i][2] === "Listening Live") {
      activeCount++;
    }
  }

  return { success: true, activeCount };
}

function getActiveListenersList() {
  const { listenersSheet } = getSheets();
  const rows = listenersSheet.getDataRange().getValues();
  const now = Date.now();
  const activeListeners = [];

  for (let i = 1; i < rows.length; i++) {
    const lastSeen = Number(rows[i][3]) || 0;
    if (now - lastSeen < 50000) {
      activeListeners.push({
        listenerId: rows[i][0],
        username: rows[i][1],
        status: rows[i][2],
        lastSeen: new Date(lastSeen).toLocaleTimeString()
      });
    }
  }
  return activeListeners;
}

// ----------------- HOURLY SCHEDULE & NAMES -----------------
function handleSaveHourlySchedule(data) {
  if (data.adminKey !== ADMIN_SECRET_KEY) return { success: false, message: "Unauthorized" };
  const { hourlySheet, settingsSheet } = getSheets();
  const schedules = data.schedules; // { "5": { name: "Morning Vibes", trackIds: [...] }, ... }

  hourlySheet.clearContents();
  hourlySheet.appendRow(["Hour", "SlotName", "TrackIdsJson"]);

  for (let h = 5; h <= 23; h++) {
    const slot = schedules[String(h)] || { name: "", trackIds: [] };
    hourlySheet.appendRow([h, slot.name || "", JSON.stringify(slot.trackIds || [])]);
  }

  const newVersion = String(Date.now());
  setSettingValue(settingsSheet, "seq_version", newVersion);

  return { success: true, message: "Hourly schedule and slot names saved successfully!", version: newVersion };
}

function getHourlySchedules() {
  const { hourlySheet } = getSheets();
  const rows = hourlySheet.getDataRange().getValues();
  const schedules = {};

  for (let i = 1; i < rows.length; i++) {
    if (rows[i][0] !== "") {
      try {
        schedules[String(rows[i][0])] = {
          name: rows[i][1] || "",
          trackIds: JSON.parse(rows[i][2] || "[]")
        };
      } catch(e) {
        schedules[String(rows[i][0])] = { name: rows[i][1] || "", trackIds: [] };
      }
    }
  }
  return schedules;
}

// ----------------- INSTANT PUSH AUDIO -----------------
function handlePushTrack(data) {
  if (data.adminKey !== ADMIN_SECRET_KEY) return { success: false, message: "Unauthorized" };
  const { settingsSheet } = getSheets();
  const fileId = data.fileId;
  const pushVer = String(Date.now());

  setSettingValue(settingsSheet, "pushed_track", fileId);
  setSettingValue(settingsSheet, "push_version", pushVer);

  return { success: true, message: "Audio pushed to all live listeners instantly!", pushVersion: pushVer };
}

// ----------------- BROADCAST DATA (Chronological) -----------------
function getStationData() {
  const { tracksSheet, settingsSheet } = getSheets();
  const trackRows = tracksSheet.getDataRange().getValues();
  const tracks = [];

  for (let i = 1; i < trackRows.length; i++) {
    if (trackRows[i][4]) {
      tracks.push({
        seq: Number(trackRows[i][0]) || i,
        title: trackRows[i][1],
        description: trackRows[i][2],
        contributor: trackRows[i][3],
        fileId: trackRows[i][4],
        streamUrl: trackRows[i][5],
        createdAt: trackRows[i][6] || ""
      });
    }
  }

  // 3) Sorted strictly in Chronological order (earliest upload first)
  tracks.sort((a, b) => new Date(a.createdAt || 0) - new Date(b.createdAt || 0));

  const settingsRows = settingsSheet.getDataRange().getValues();
  const settings = {};
  for (let i = 1; i < settingsRows.length; i++) {
    settings[settingsRows[i][0]] = settingsRows[i][1];
  }

  const hourly = getHourlySchedules();
  const activeListeners = getActiveListenersList();
  const liveCount = activeListeners.filter(l => l.status === "Listening Live").length;

  return { success: true, tracks, settings, hourly, liveListenersCount: liveCount };
}

// ----------------- ADMIN HANDLERS -----------------
function handleAdminLogin(data) {
  if (data.key === ADMIN_SECRET_KEY) return { success: true, message: "Authenticated" };
  return { success: false, message: "Invalid Admin Passcode" };
}

function getAdminData() {
  const { usersSheet } = getSheets();
  const station = getStationData();

  const userRows = usersSheet.getDataRange().getValues();
  const users = [];
  for (let i = 1; i < userRows.length; i++) {
    users.push({
      username: userRows[i][0],
      mobile: userRows[i][1],
      status: userRows[i][3],
      registeredAt: userRows[i][4]
    });
  }

  const activeListeners = getActiveListenersList();

  return {
    success: true,
    users,
    tracks: station.tracks,
    settings: station.settings,
    hourly: station.hourly,
    activeListeners
  };
}

function handleAdminUpdateUser(data) {
  if (data.adminKey !== ADMIN_SECRET_KEY) return { success: false, message: "Unauthorized" };
  const { mobile, status } = data;
  const { usersSheet } = getSheets();
  const rows = usersSheet.getDataRange().getValues();

  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][1]).trim() === String(mobile).trim()) {
      usersSheet.getRange(i + 1, 4).setValue(status);
      return { success: true, message: `User status changed to ${status}` };
    }
  }
  return { success: false, message: "User not found" };
}

function handleDeleteTrack(data) {
  if (data.adminKey !== ADMIN_SECRET_KEY) return { success: false, message: "Unauthorized" };
  const { tracksSheet, settingsSheet } = getSheets();
  const rows = tracksSheet.getDataRange().getValues();

  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][4]).trim() === String(data.fileId).trim()) {
      try { DriveApp.getFileById(data.fileId).setTrashed(true); } catch (err) {}
      tracksSheet.deleteRow(i + 1);
      setSettingValue(settingsSheet, "seq_version", String(Date.now()));
      return { success: true, message: "Track deleted successfully" };
    }
  }
  return { success: false, message: "Track not found" };
}

function handleSaveSettings(data) {
  if (data.adminKey !== ADMIN_SECRET_KEY) return { success: false, message: "Unauthorized" };
  const { settingsSheet } = getSheets();
  const s = data.settings;
  const rows = settingsSheet.getDataRange().getValues();

  for (let i = 1; i < rows.length; i++) {
    const key = rows[i][0];
    if (s[key] !== undefined) {
      settingsSheet.getRange(i + 1, 2).setValue(s[key]);
    }
  }
  return { success: true, message: "Settings saved" };
}

function setSettingValue(sheet, key, value) {
  const rows = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][0] === key) {
      sheet.getRange(i + 1, 2).setValue(value);
      return;
    }
  }
  sheet.appendRow([key, value]);
}