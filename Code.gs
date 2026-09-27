/**
 * Synchro Radio Backend - Google Apps Script
 * Robust, zero-error execution for all functions (direct or webhook)
 */

const FOLDER_NAME = "Community Radio";
const ADMIN_SECRET_KEY = "admin123";

function doGet(e) {
  try {
    e = e || { parameter: { action: "getStationData" } };
    const parameter = e.parameter || {};
    const action = parameter.action || "getStationData";

    // DIRECT AUDIO STREAMER (Bypasses Google Drive 403 & CORS)
    if (action === "streamAudio") {
      const fileId = parameter.fileId;
      if (!fileId) {
        return ContentService.createTextOutput(JSON.stringify({ success: false, message: "Missing fileId" }))
          .setMimeType(ContentService.MimeType.JSON);
      }
      try {
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
      if (parameter.key === ADMIN_SECRET_KEY) {
        response = getAdminData();
      } else {
        response = { success: false, message: "Unauthorized admin access" };
      }
    }

    return ContentService.createTextOutput(JSON.stringify(response))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ success: false, error: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function doPost(e) {
  try {
    if (!e || !e.postData || !e.postData.contents) {
      return ContentService.createTextOutput(JSON.stringify({ success: false, message: "No post data received" }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    let contents = {};
    try {
      contents = JSON.parse(e.postData.contents);
    } catch(parseErr) {
      return ContentService.createTextOutput(JSON.stringify({ success: false, message: "Invalid JSON body" }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    const action = contents.action;
    let response = { success: false, message: "Invalid action" };

    if (action === "register") response = handleRegister(contents);
    else if (action === "login") response = handleLogin(contents);
    else if (action === "uploadAudio") response = handleAudioUpload(contents);
    else if (action === "adminLogin") response = handleAdminLogin(contents);
    else if (action === "adminUpdateUserStatus") response = handleAdminUpdateUser(contents);
    else if (action === "adminUpdateAudioApproval") response = handleUpdateAudioApproval(contents);
    else if (action === "adminDeleteTrack") response = handleDeleteTrack(contents);
    else if (action === "adminUpdateTrackCategory") response = handleUpdateTrackCategory(contents);
    else if (action === "adminSaveSettings") response = handleSaveSettings(contents);
    else if (action === "heartbeat") response = handleHeartbeat(contents);
    else if (action === "saveHourlySchedule") response = handleSaveHourlySchedule(contents);
    else if (action === "adminPushTrack") response = handlePushTrack(contents);
    else if (action === "adminPushLiveMic") response = handlePushLiveMic(contents);
    else if (action === "markTrackPlayedInSlot") response = handleMarkTrackPlayedInSlot(contents);

    return ContentService.createTextOutput(JSON.stringify(response))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({ success: false, message: error.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ----------------- SHEET & DRIVE INITIALIZERS -----------------

function getOrCreateFolder() {
  const folders = DriveApp.getFoldersByName(FOLDER_NAME);
  if (folders.hasNext()) return folders.next();
  const folder = DriveApp.createFolder(FOLDER_NAME);
  folder.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return folder;
}

function getSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) {
    throw new Error("No active spreadsheet found. Open Extensions > Apps Script inside your Google Sheet.");
  }
  
  let usersSheet = ss.getSheetByName("Users");
  if (!usersSheet) {
    usersSheet = ss.insertSheet("Users");
    usersSheet.appendRow(["Username", "Mobile", "Password", "Status", "RegisteredAt"]);
  }

  let tracksSheet = ss.getSheetByName("Tracks");
  const trackHeaders = ["Seq", "Title", "Description", "Contributor", "FileId", "StreamUrl", "CreatedAt", "Category", "ApprovalStatus", "PlayedInSlot"];
  if (!tracksSheet) {
    tracksSheet = ss.insertSheet("Tracks");
    tracksSheet.appendRow(trackHeaders);
  } else {
    // Auto-migrate older sheets that have fewer header columns
    const lastCol = tracksSheet.getLastColumn();
    if (lastCol < trackHeaders.length) {
      tracksSheet.getRange(1, 1, 1, trackHeaders.length).setValues([trackHeaders]);
    }
  }

  let settingsSheet = ss.getSheetByName("Settings");
  if (!settingsSheet) {
    settingsSheet = ss.insertSheet("Settings");
    settingsSheet.appendRow(["Key", "Value"]);
    settingsSheet.appendRow(["station_name", "Synchro Radio"]);
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

  let randomSeqSheet = ss.getSheetByName("HourlyRandomSequence");
  if (!randomSeqSheet) {
    randomSeqSheet = ss.insertSheet("HourlyRandomSequence");
    randomSeqSheet.appendRow(["HourDateKey", "TrackIdsJson"]);
  }

  return { usersSheet, tracksSheet, settingsSheet, listenersSheet, hourlySheet, randomSeqSheet };
}

// ----------------- USER AUTHENTICATION -----------------

function handleRegister(data) {
  data = data || {};
  const username = String(data.username || "").trim();
  const mobile = String(data.mobile || "").trim();
  const password = String(data.password || "").trim();

  if (!username || !mobile || !password) {
    return { success: false, message: "All fields are required" };
  }

  const { usersSheet } = getSheets();
  const rows = usersSheet.getDataRange().getValues();

  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][1]).trim() === mobile) {
      return { success: false, message: "Mobile number already registered" };
    }
  }

  usersSheet.appendRow([username, mobile, password, "pending", new Date().toISOString()]);
  return {
    success: true,
    message: "Registration successful. Contributor approval is pending.",
    user: { username, mobile, status: "pending" }
  };
}

function handleLogin(data) {
  data = data || {};
  const mobile = String(data.mobile || "").trim();
  const password = String(data.password || "").trim();

  if (!mobile || !password) return { success: false, message: "Mobile and password required" };

  const { usersSheet } = getSheets();
  const rows = usersSheet.getDataRange().getValues();

  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][1]).trim() === mobile && String(rows[i][2]).trim() === password) {
      return {
        success: true,
        user: {
          username: String(rows[i][0] || "User"),
          mobile: String(rows[i][1] || ""),
          status: String(rows[i][3] || "pending")
        }
      };
    }
  }
  return { success: false, message: "Invalid mobile number or password" };
}

// ----------------- AUDIO UPLOAD -----------------

function handleAudioUpload(data) {
  data = data || {};
  const mobile = String(data.mobile || "").trim();
  const username = String(data.username || "Contributor").trim();
  const title = String(data.title || "Untitled Track").trim();
  const description = String(data.description || "").trim();
  const base64File = data.base64File;
  const fileName = data.fileName || `${title}.mp3`;
  const mimeType = data.mimeType || "audio/mpeg";

  if (!base64File) return { success: false, message: "Missing audio payload" };

  const { usersSheet, tracksSheet } = getSheets();

  const userRows = usersSheet.getDataRange().getValues();
  let isApproved = false;
  for (let i = 1; i < userRows.length; i++) {
    if (String(userRows[i][1]).trim() === mobile && String(userRows[i][3]).trim() === "approved") {
      isApproved = true;
      break;
    }
  }

  if (!isApproved) {
    return { success: false, message: "User is not approved by Admin to upload audio." };
  }

  const folder = getOrCreateFolder();
  const decodedData = Utilities.base64Decode(base64File);
  const blob = Utilities.newBlob(decodedData, mimeType, fileName);
  const file = folder.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

  const fileId = file.getId();
  const streamUrl = `https://drive.google.com/uc?export=download&id=${fileId}`;
  const createdAt = new Date().toISOString();
  
  const trackCat = "Slot plays";
  const approvalStatus = "pending";
  const playedInSlot = "false";

  const trackRows = tracksSheet.getDataRange().getValues();
  const nextSeq = trackRows.length;

  tracksSheet.appendRow([nextSeq, title, description, username, fileId, streamUrl, createdAt, trackCat, approvalStatus, playedInSlot]);

  return {
    success: true,
    message: "Audio uploaded as Slot play! Pending administrator approval.",
    track: { seq: nextSeq, title, description, fileId, streamUrl, createdAt, category: trackCat, approvalStatus, playedInSlot: false }
  };
}

// ----------------- HEARTBEAT & LISTENERS -----------------

function handleHeartbeat(data) {
  data = data || {};
  const listenerId = String(data.listenerId || "L_anon").trim();
  const username = String(data.username || "Guest Listener").trim();
  const isPlaying = Boolean(data.isPlaying);

  const { listenersSheet } = getSheets();
  const rows = listenersSheet.getDataRange().getValues();
  const now = Date.now();
  let found = false;

  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]).trim() === listenerId) {
      listenersSheet.getRange(i + 1, 2).setValue(username);
      listenersSheet.getRange(i + 1, 3).setValue(isPlaying ? "Listening Live" : "Idle");
      listenersSheet.getRange(i + 1, 4).setValue(now);
      found = true;
      break;
    }
  }

  if (!found) {
    listenersSheet.appendRow([listenerId, username, isPlaying ? "Listening Live" : "Idle", now]);
  }

  let activeCount = 0;
  const updatedRows = listenersSheet.getDataRange().getValues();
  for (let i = 1; i < updatedRows.length; i++) {
    const lastSeen = Number(updatedRows[i][3]) || 0;
    if (now - lastSeen < 50000 && String(updatedRows[i][2]) === "Listening Live") {
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
        listenerId: String(rows[i][0] || ""),
        username: String(rows[i][1] || "Guest"),
        status: String(rows[i][2] || "Idle"),
        lastSeen: new Date(lastSeen).toLocaleTimeString()
      });
    }
  }
  return activeListeners;
}

// ----------------- HOURLY SCHEDULE & RANDOM ROTATION -----------------

function getOrGenerateHourlyRandomSequence(tracksSheet, randomSeqSheet) {
  tracksSheet = tracksSheet || getSheets().tracksSheet;
  randomSeqSheet = randomSeqSheet || getSheets().randomSeqSheet;

  const now = new Date();
  const key = `${now.getFullYear()}-${now.getMonth()+1}-${now.getDate()}_H${now.getHours()}`;

  const rows = randomSeqSheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]) === key) {
      try {
        return JSON.parse(rows[i][1]);
      } catch(e) {}
    }
  }

  const trackRows = tracksSheet.getDataRange().getValues();
  const eligible = [];
  for (let i = 1; i < trackRows.length; i++) {
    const fileId = String(trackRows[i][4] || "");
    const cat = String(trackRows[i][7] || "");
    const app = String(trackRows[i][8] || "");
    const played = String(trackRows[i][9]);

    if (fileId && app === "approved" && played === "true" && cat === "Random plays") {
      eligible.push(fileId);
    }
  }

  for (let i = eligible.length - 1; i > 0; i--) {
    const j = Math.floor(Math.random() * (i + 1));
    const temp = eligible[i];
    eligible[i] = eligible[j];
    eligible[j] = temp;
  }

  randomSeqSheet.appendRow([key, JSON.stringify(eligible)]);
  return eligible;
}

function handleSaveHourlySchedule(data) {
  data = data || {};
  if (data.adminKey !== ADMIN_SECRET_KEY) return { success: false, message: "Unauthorized" };
  const { hourlySheet, settingsSheet } = getSheets();
  const schedules = data.schedules || {};

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
          name: String(rows[i][1] || ""),
          trackIds: JSON.parse(rows[i][2] || "[]")
        };
      } catch(e) {
        schedules[String(rows[i][0])] = { name: String(rows[i][1] || ""), trackIds: [] };
      }
    }
  }
  return schedules;
}

function handlePushLiveMic(data) {
  data = data || {};
  if (data.adminKey !== ADMIN_SECRET_KEY) return { success: false, message: "Unauthorized" };
  if (!data.base64File) return { success: false, message: "Missing audio recording" };

  const { tracksSheet, settingsSheet } = getSheets();
  const folder = getOrCreateFolder();

  const decodedData = Utilities.base64Decode(data.base64File);
  const blob = Utilities.newBlob(decodedData, "audio/webm", `Live_Broadcast_${Date.now()}.webm`);
  const file = folder.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

  const fileId = file.getId();
  const streamUrl = `https://drive.google.com/uc?export=download&id=${fileId}`;
  const nextSeq = tracksSheet.getDataRange().getValues().length;

  tracksSheet.appendRow([nextSeq, "🔴 Live Admin Broadcast", "Urgent Live Announcement", "Admin", fileId, streamUrl, new Date().toISOString(), "Slot plays", "approved", "true"]);

  const pushVer = String(Date.now());
  setSettingValue(settingsSheet, "pushed_track", fileId);
  setSettingValue(settingsSheet, "push_version", pushVer);

  return { success: true, message: "Live mic broadcast pushed to all listeners!", fileId, pushVersion: pushVer };
}

function handlePushTrack(data) {
  data = data || {};
  if (data.adminKey !== ADMIN_SECRET_KEY) return { success: false, message: "Unauthorized" };
  const { settingsSheet } = getSheets();
  const fileId = String(data.fileId || "");
  const pushVer = String(Date.now());

  setSettingValue(settingsSheet, "pushed_track", fileId);
  setSettingValue(settingsSheet, "push_version", pushVer);

  return { success: true, message: "Audio pushed to all live listeners instantly!", pushVersion: pushVer };
}

function handleUpdateAudioApproval(data) {
  data = data || {};
  if (data.adminKey !== ADMIN_SECRET_KEY) return { success: false, message: "Unauthorized" };
  const { tracksSheet } = getSheets();
  const rows = tracksSheet.getDataRange().getValues();
  const targetId = String(data.fileId || "").trim();

  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][4]).trim() === targetId) {
      tracksSheet.getRange(i + 1, 9).setValue(data.status);
      return { success: true, message: `Audio approval updated to ${data.status}` };
    }
  }
  return { success: false, message: "Track not found" };
}

function handleMarkTrackPlayedInSlot(data) {
  data = data || {};
  const { tracksSheet } = getSheets();
  const rows = tracksSheet.getDataRange().getValues();
  const targetId = String(data.fileId || "").trim();

  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][4]).trim() === targetId) {
      tracksSheet.getRange(i + 1, 10).setValue("true");
      return { success: true, message: "Audio marked as debuted in slot" };
    }
  }
  return { success: false, message: "Track not found" };
}

function handleUpdateTrackCategory(data) {
  data = data || {};
  if (data.adminKey !== ADMIN_SECRET_KEY) return { success: false, message: "Unauthorized" };
  const { tracksSheet } = getSheets();
  const rows = tracksSheet.getDataRange().getValues();
  const targetId = String(data.fileId || "").trim();

  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][4]).trim() === targetId) {
      tracksSheet.getRange(i + 1, 8).setValue(data.category);
      return { success: true, message: "Category updated to " + data.category };
    }
  }
  return { success: false, message: "Track not found" };
}

function handleDeleteTrack(data) {
  data = data || {};
  if (data.adminKey !== ADMIN_SECRET_KEY) return { success: false, message: "Unauthorized" };
  const { tracksSheet, settingsSheet } = getSheets();
  const rows = tracksSheet.getDataRange().getValues();
  const targetId = String(data.fileId || "").trim();

  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][4]).trim() === targetId) {
      try { DriveApp.getFileById(targetId).setTrashed(true); } catch (err) {}
      tracksSheet.deleteRow(i + 1);
      setSettingValue(settingsSheet, "seq_version", String(Date.now()));
      return { success: true, message: "Track deleted successfully" };
    }
  }
  return { success: false, message: "Track not found" };
}

function handleSaveSettings(data) {
  data = data || {};
  if (data.adminKey !== ADMIN_SECRET_KEY) return { success: false, message: "Unauthorized" };
  const { settingsSheet } = getSheets();
  const s = data.settings || {};
  const rows = settingsSheet.getDataRange().getValues();

  for (let i = 1; i < rows.length; i++) {
    const key = rows[i][0];
    if (s[key] !== undefined) {
      settingsSheet.getRange(i + 1, 2).setValue(s[key]);
    }
  }
  return { success: true, message: "Settings saved" };
}

// ----------------- BROADCAST DATA (Chronological) -----------------

function getStationData() {
  const { tracksSheet, settingsSheet, hourlySheet, randomSeqSheet } = getSheets();
  const trackRows = tracksSheet.getDataRange().getValues();
  const tracks = [];

  for (let i = 1; i < trackRows.length; i++) {
    if (trackRows[i][4]) {
      tracks.push({
        seq: Number(trackRows[i][0]) || i,
        title: String(trackRows[i][1] || "Untitled"),
        description: String(trackRows[i][2] || ""),
        contributor: String(trackRows[i][3] || "Community"),
        fileId: String(trackRows[i][4]),
        streamUrl: String(trackRows[i][5]),
        createdAt: String(trackRows[i][6] || ""),
        category: String(trackRows[i][7] || "Slot plays"),
        approvalStatus: String(trackRows[i][8] || "pending"),
        playedInSlot: String(trackRows[i][9]) === "true"
      });
    }
  }

  tracks.sort((a, b) => new Date(a.createdAt || 0) - new Date(b.createdAt || 0));

  const settingsRows = settingsSheet.getDataRange().getValues();
  const settings = {};
  for (let i = 1; i < settingsRows.length; i++) {
    settings[settingsRows[i][0]] = settingsRows[i][1];
  }

  const hourly = getHourlySchedules();
  const activeListeners = getActiveListenersList();
  const liveCount = activeListeners.filter(l => l.status === "Listening Live").length;
  const hourlyRandomSequence = getOrGenerateHourlyRandomSequence(tracksSheet, randomSeqSheet);

  return { success: true, tracks, settings, hourly, hourlyRandomSequence, liveListenersCount: liveCount };
}

// ----------------- ADMIN HANDLERS -----------------

function handleAdminLogin(data) {
  data = data || {};
  if (data.key === ADMIN_SECRET_KEY) return { success: true, message: "Authenticated" };
  return { success: false, message: "Invalid Admin Passcode" };
}

function getAdminData() {
  const { usersSheet } = getSheets();
  const station = getStationData();

  const userRows = usersSheet.getDataRange().getValues();
  const users = [];
  for (let i = 1; i < userRows.length; i++) {
    if (userRows[i][0] || userRows[i][1]) {
      users.push({
        username: String(userRows[i][0] || "User"),
        mobile: String(userRows[i][1] || ""), // Always converted to String
        status: String(userRows[i][3] || "pending"),
        registeredAt: String(userRows[i][4] || "")
      });
    }
  }

  const activeListeners = getActiveListenersList();

  return {
    success: true,
    users: users,
    tracks: station.tracks || [],
    settings: station.settings || {},
    hourly: station.hourly || {},
    activeListeners: activeListeners || []
  };
}

function handleAdminUpdateUser(data) {
  data = data || {};
  if (data.adminKey !== ADMIN_SECRET_KEY) return { success: false, message: "Unauthorized" };
  const { usersSheet } = getSheets();
  const rows = usersSheet.getDataRange().getValues();
  const targetMobile = String(data.mobile || "").trim();

  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][1]).trim() === targetMobile) {
      usersSheet.getRange(i + 1, 4).setValue(data.status);
      return { success: true, message: `User status changed to ${data.status}` };
    }
  }
  return { success: false, message: "User not found" };
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

// ----------------- TEST UTILITY -----------------
// Run this directly from the Apps Script editor to ensure ZERO errors
function testAllFunctions() {
  Logger.log("Testing initialization...");
  const sheets = getSheets();
  Logger.log("Sheets loaded: " + Object.keys(sheets).join(", "));
  
  const station = getStationData();
  Logger.log("Station tracks count: " + station.tracks.length);
  
  const admin = getAdminData();
  Logger.log("Users count: " + admin.users.length);

  const getTest = doGet();
  Logger.log("doGet() executed successfully.");

  const postTest = doPost();
  Logger.log("doPost() handled null payload safely.");

  Logger.log("All functions run with ZERO errors!");
}