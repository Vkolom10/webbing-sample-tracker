const SHEET_SAMPLES = "Samples";
const SHEET_COLORS = "Webbing Colors";
const SHEET_GROUPS = "Webbing Groups";
const DRIVE_FOLDER_NAME = "Webbing Sample Photos";

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    let sheet;
    if (data.type === "swatch") {
      sheet = ss.getSheetByName(SHEET_COLORS);
    } else if (data.type === "group") {
      sheet = ss.getSheetByName(SHEET_GROUPS);
    } else {
      sheet = ss.getSheetByName(SHEET_SAMPLES);
    }

    if (!sheet) {
      return jsonResponse({ success: false, error: "Sheet not found: " + data.type });
    }

    let photoUrl = data.photoUrl || "";

    if (data.photoBase64 && String(data.photoBase64).startsWith("data:")) {
      photoUrl = savePhotoToDrive(
        data.photoBase64,
        data.fileName || data.uid || "photo"
      );
    }

    if (data.action === "add") {
      if (data.type === "swatch") {
        sheet.appendRow([
          data.name || "",
          data.number || "",
          data.webbingType || "solid",
          data.uid || "",
          photoUrl
        ]);
      } else if (data.type === "group") {
        sheet.appendRow([
          data.name || "",
          data.note || "",
          data.uid || "",
          photoUrl
        ]);
      } else {
        sheet.appendRow([
          data.name || "",
          data.note || "",
          data.date || "",
          data.uid || "",
          photoUrl
        ]);
      }
    }

    if (data.action === "update") {
      updateRow(sheet, data, photoUrl);
    }

    if (data.action === "delete") {
      deleteRow(sheet, data.uid, data.type);
    }

    return jsonResponse({ success: true, photoUrl: photoUrl });

  } catch (err) {
    return jsonResponse({ success: false, error: err.toString() });
  }
}

function doGet(e) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const sampleSheet = ss.getSheetByName(SHEET_SAMPLES);
  const colorSheet = ss.getSheetByName(SHEET_COLORS);
  const groupSheet = ss.getSheetByName(SHEET_GROUPS);

  const samples = [];
  const colors = [];
  const groups = [];

  if (sampleSheet) {
    const rows = sampleSheet.getDataRange().getValues();
    for (let i = 1; i < rows.length; i++) {
      if (!rows[i][0]) continue;
      samples.push({
        name: rows[i][0],
        note: rows[i][1],
        date: rows[i][2],
        uid: rows[i][3],
        photoUrl: rows[i][4] || ""
      });
    }
  }

  if (colorSheet) {
    const rows = colorSheet.getDataRange().getValues();
    for (let i = 1; i < rows.length; i++) {
      if (!rows[i][0]) continue;
      colors.push({
        name: rows[i][0],
        number: rows[i][1],
        webbingType: rows[i][2] || "solid",
        uid: rows[i][3],
        photoUrl: rows[i][4] || ""
      });
    }
  }

  if (groupSheet) {
    const rows = groupSheet.getDataRange().getValues();
    for (let i = 1; i < rows.length; i++) {
      if (!rows[i][0]) continue;
      groups.push({
        name: rows[i][0],
        note: rows[i][1],
        uid: rows[i][2],
        photoUrl: rows[i][3] || ""
      });
    }
  }

  const output = {
    success: true,
    samples: samples,
    colors: colors,
    groups: groups
  };

  const callback = e && e.parameter && e.parameter.callback;
  if (callback) {
    return ContentService
      .createTextOutput(callback + "(" + JSON.stringify(output) + ")")
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }

  return jsonResponse(output);
}

function updateRow(sheet, data, photoUrl) {
  const rows = sheet.getDataRange().getValues();

  let uidCol;
  if (data.type === "swatch") {
    uidCol = 3;
  } else if (data.type === "group") {
    uidCol = 2;
  } else {
    uidCol = 3;
  }

  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][uidCol]) === String(data.uid)) {
      if (data.type === "swatch") {
        sheet.getRange(i + 1, 1, 1, 5).setValues([[
          data.name || "",
          data.number || "",
          data.webbingType || "solid",
          data.uid || "",
          photoUrl || rows[i][4] || ""
        ]]);
      } else if (data.type === "group") {
        sheet.getRange(i + 1, 1, 1, 4).setValues([[
          data.name || "",
          data.note || "",
          data.uid || "",
          photoUrl || rows[i][3] || ""
        ]]);
      } else {
        sheet.getRange(i + 1, 1, 1, 5).setValues([[
          data.name || "",
          data.note || "",
          data.date || "",
          data.uid || "",
          photoUrl || rows[i][4] || ""
        ]]);
      }
      return;
    }
  }

  // If update cannot find the row, add it instead.
  data.action = "add";
  doPost({ postData: { contents: JSON.stringify(data) } });
}

function deleteRow(sheet, uid, type) {
  const rows = sheet.getDataRange().getValues();

  let uidCol;
  if (type === "swatch") {
    uidCol = 3;
  } else if (type === "group") {
    uidCol = 2;
  } else {
    uidCol = 3;
  }

  for (let i = rows.length - 1; i >= 1; i--) {
    if (String(rows[i][uidCol]) === String(uid)) {
      sheet.deleteRow(i + 1);
      return;
    }
  }
}

function savePhotoToDrive(photoBase64, fileName) {
  const folder = getOrCreateFolder(DRIVE_FOLDER_NAME);
  const match = photoBase64.match(/^data:(.*?);base64,/);
  if (!match) return "";

  const contentType = match[1];
  const base64Data = photoBase64.replace(/^data:.*?;base64,/, "");
  const bytes = Utilities.base64Decode(base64Data);
  const extension = contentType.split("/")[1] || "jpg";
  const cleanName = String(fileName).replace(/[^\w\-]/g, "_");

  const blob = Utilities.newBlob(bytes, contentType, cleanName + "." + extension);
  const file = folder.createFile(blob);

  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

  return file.getUrl();
}

function getOrCreateFolder(folderName) {
  const folders = DriveApp.getFoldersByName(folderName);
  if (folders.hasNext()) return folders.next();
  return DriveApp.createFolder(folderName);
}

function jsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
