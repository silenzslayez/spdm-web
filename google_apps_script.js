function doGet(e) {
  return handleRequest(e);
}

function doPost(e) {
  return handleRequest(e);
}

function handleRequest(e) {
  // CORS Headers
  const headers = {
    "Access-Control-Allow-Origin": "*",
    "Access-Control-Allow-Methods": "GET, POST, OPTIONS",
    "Access-Control-Allow-Headers": "Content-Type",
  };
  
  if (e.postData && e.postData.contents) {
    try {
      const payload = JSON.parse(e.postData.contents);
      
      // Handle Photo Upload
      if (payload.action === 'uploadPhoto') {
        return handleUploadPhoto(payload);
      }
      
      // Handle Data Push from App -> Google Sheets
      if (payload.action === 'pushData') {
        return handlePushData(payload);
      }
      
      // Handle User Sync
      if (payload.action === 'pushUsers') {
        return handlePushUsers(payload);
      }
      
      // Handle Activity Log
      if (payload.action === 'logActivity') {
        return handleLogActivity(payload);
      }
      
    } catch (err) {
      return responseJSON({ success: false, message: err.toString() });
    }
  }

  // GET Request: Fetch all data
  try {
    const action = e.parameter.action;
    
    if (action === 'getUsers') {
      return handleGetUsers();
    }
    
    if (action === 'getLogs') {
      return handleGetLogs();
    }

    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
    const data = sheet.getDataRange().getValues();
    if (data.length < 2) return responseJSON([]);
    
    const sheetHeaders = data[0].map(h => String(h).toUpperCase().trim());
    
    // Cari index lajur berdasarkan format APDM asal atau format ringkas
    const idIdx = sheetHeaders.findIndex(h => h === "ID MURID" || h === "ID");
    const nameIdx = sheetHeaders.findIndex(h => h === "NAMA" || h === "NAMA MURID");
    const icIdx = sheetHeaders.findIndex(h => h === "NO. PENGENALAN" || h === "NO KAD PENGENALAN");
    const classIdx = sheetHeaders.findIndex(h => h === "NAMA KELAS" || h === "KELAS");
    const genderIdx = sheetHeaders.findIndex(h => h === "JANTINA");
    const p1NameIdx = sheetHeaders.findIndex(h => h === "PENJAGA 1" || h === "NAMA BAPA / PENJAGA 1" || h === "NAMA WARIS");
    const p1RelIdx = sheetHeaders.findIndex(h => h === "HUBUNGAN PENJAGA 1" || h === "HUBUNGAN");
    const p1PhoneIdx = sheetHeaders.findIndex(h => h === "NO. TEL. BIMBIT PENJAGA 1" || h === "NO TEL BAPA / PENJAGA 1" || h === "NO TEL WARIS");
    const p2NameIdx = sheetHeaders.findIndex(h => h === "PENJAGA 2" || h === "NAMA IBU / PENJAGA 2");
    const p2PhoneIdx = sheetHeaders.findIndex(h => h === "NO. TEL. BIMBIT PENJAGA 2" || h === "NO TEL IBU / PENJAGA 2");
    const address1Idx = sheetHeaders.findIndex(h => h === "ALAMAT 1" || h === "ALAMAT KEDIAMAN" || h === "ALAMAT");
    const address2Idx = sheetHeaders.findIndex(h => h === "ALAMAT 2");
    const address3Idx = sheetHeaders.findIndex(h => h === "ALAMAT 3");
    let hostelIdx = sheetHeaders.findIndex(h => h === "STATUS ASRAMA" || h === "ASRAMA");
    let religionIdx = sheetHeaders.findIndex(h => h === "AGAMA" || h === "AGAMA MURID");
    let teacherIdx = sheetHeaders.findIndex(h => h === "NAMA GURU KELAS" || h === "GURU KELAS");
    let photoIdx = sheetHeaders.findIndex(h => h === "GAMBAR");
    
    // Jika lajur GAMBAR tiada, cipta lajur baru di hujung
    if (photoIdx === -1) {
      photoIdx = sheetHeaders.length;
      sheet.getRange(1, photoIdx + 1).setValue("GAMBAR");
      sheetHeaders.push("GAMBAR");
    }
    
    const students = [];
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (!row[idIdx] && !row[nameIdx]) continue;
      
      let address = "";
      if (address1Idx > -1 && row[address1Idx]) address += row[address1Idx] + ", ";
      if (address2Idx > -1 && row[address2Idx]) address += row[address2Idx] + ", ";
      if (address3Idx > -1 && row[address3Idx]) address += row[address3Idx];
      address = address.replace(/,\s*$/, "").trim(); // Buang koma di hujung
      
      students.push({
        id: idIdx > -1 ? String(row[idIdx]) : `PEL-${i}`,
        name: nameIdx > -1 ? String(row[nameIdx]) : "Tanpa Nama",
        icNumber: icIdx > -1 ? String(row[icIdx]) : "",
        className: classIdx > -1 ? String(row[classIdx]) : "",
        gender: genderIdx > -1 ? String(row[genderIdx]) : "Lelaki",
        parentName: p1NameIdx > -1 ? String(row[p1NameIdx]) : "",
        parentRelation: p1RelIdx > -1 ? String(row[p1RelIdx]) : "Waris",
        parentPhone: p1PhoneIdx > -1 ? String(row[p1PhoneIdx]) : "",
        motherName: p2NameIdx > -1 ? String(row[p2NameIdx]) : "",
        motherPhone: p2PhoneIdx > -1 ? String(row[p2PhoneIdx]) : "",
        address: address,
        photoUrl: photoIdx > -1 ? String(row[photoIdx]) : "",
        hostelStatus: hostelIdx > -1 ? String(row[hostelIdx]) : "TIDAK",
        religion: religionIdx > -1 ? String(row[religionIdx]) : "Tiada Maklumat",
        teacherName: teacherIdx > -1 ? String(row[teacherIdx]) : ""
      });
    }
    
    return responseJSON(students);
  } catch (err) {
    return responseJSON({ error: err.toString() });
  }
}

function handleUploadPhoto(payload) {
  const studentId = payload.studentId;
  const base64Image = payload.base64Image;
  
  try {
    // Find or create SPDM Photos folder
    let folder;
    const folders = DriveApp.getFoldersByName("Gambar Murid SPDM");
    if (folders.hasNext()) {
        folder = folders.next();
    } else {
        folder = DriveApp.createFolder("Gambar Murid SPDM");
        folder.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    }
    
    // Convert base64 to blob
    const contentType = base64Image.split(';')[0].split(':')[1];
    const base64Data = base64Image.split(',')[1];
    const blob = Utilities.newBlob(Utilities.base64Decode(base64Data), contentType, studentId + ".jpg");
    
    // Check if file already exists (overwrite by trashing old one)
    const existingFiles = folder.getFilesByName(studentId + ".jpg");
    if (existingFiles.hasNext()) {
        existingFiles.next().setTrashed(true);
    }
    
    const file = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    const fileUrl = "https://drive.google.com/uc?export=view&id=" + file.getId();
    
    // Update the sheet's GAMBAR column
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    const data = sheet.getDataRange().getValues();
    const sheetHeaders = data[0].map(h => String(h).toUpperCase().trim());
    const idIdx = sheetHeaders.findIndex(h => h === "ID MURID" || h === "ID");
    let photoIdx = sheetHeaders.findIndex(h => h === "GAMBAR");
    
    if (photoIdx === -1) {
        photoIdx = sheetHeaders.length;
        sheet.getRange(1, photoIdx + 1).setValue("GAMBAR");
    }
    
    if (idIdx > -1) {
        const rowIndex = data.findIndex(row => String(row[idIdx]) === String(studentId));
        if (rowIndex > 0) {
            sheet.getRange(rowIndex + 1, photoIdx + 1).setValue(fileUrl);
        }
    }
    
    return responseJSON({ success: true, message: "Gambar berjaya disimpan", url: fileUrl });
  } catch (err) {
    return responseJSON({ success: false, message: err.toString() });
  }
}

function handlePushData(payload) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
    const data = sheet.getDataRange().getValues();
    const sheetHeaders = data[0].map(h => String(h).toUpperCase().trim());
    const idIdx = sheetHeaders.findIndex(h => h === "ID MURID" || h === "ID");
    
    if (idIdx === -1) return responseJSON({ success: false, message: "Lajur ID MURID tiada." });

    let updated = 0;
    const studentsToUpdate = payload.students || [];

    for (const student of studentsToUpdate) {
      const rowIndex = data.findIndex(row => String(row[idIdx]) === String(student.id));
      if (rowIndex > 0) {
        // Update maklumat asas
        const p1PhoneIdx = sheetHeaders.findIndex(h => h === "NO. TEL. BIMBIT PENJAGA 1" || h === "NO TEL BAPA / PENJAGA 1" || h === "NO TEL WARIS");
        const address1Idx = sheetHeaders.findIndex(h => h === "ALAMAT 1" || h === "ALAMAT KEDIAMAN" || h === "ALAMAT");
        const icIdx = sheetHeaders.findIndex(h => h === "NO. PENGENALAN" || h === "NO KAD PENGENALAN");
        const p1NameIdx = sheetHeaders.findIndex(h => h === "PENJAGA 1" || h === "NAMA BAPA / PENJAGA 1" || h === "NAMA WARIS");
        
        if (p1PhoneIdx > -1) sheet.getRange(rowIndex + 1, p1PhoneIdx + 1).setValue(student.parentPhone);
        if (address1Idx > -1) sheet.getRange(rowIndex + 1, address1Idx + 1).setValue(student.address);
        if (icIdx > -1) sheet.getRange(rowIndex + 1, icIdx + 1).setValue(student.icNumber);
        if (p1NameIdx > -1) sheet.getRange(rowIndex + 1, p1NameIdx + 1).setValue(student.parentName);
        updated++;
      } else {
        // Jika rekod tiada, cipta baris baru
        const newRow = new Array(sheetHeaders.length).fill("");
        newRow[idIdx] = student.id;
        
        const nameIdx = sheetHeaders.findIndex(h => h === "NAMA" || h === "NAMA MURID");
        const classIdx = sheetHeaders.findIndex(h => h === "NAMA KELAS" || h === "KELAS");
        const p1PhoneIdx = sheetHeaders.findIndex(h => h === "NO. TEL. BIMBIT PENJAGA 1" || h === "NO TEL BAPA / PENJAGA 1" || h === "NO TEL WARIS");
        const address1Idx = sheetHeaders.findIndex(h => h === "ALAMAT 1" || h === "ALAMAT KEDIAMAN" || h === "ALAMAT");
        const icIdx = sheetHeaders.findIndex(h => h === "NO. PENGENALAN" || h === "NO KAD PENGENALAN");
        const p1NameIdx = sheetHeaders.findIndex(h => h === "PENJAGA 1" || h === "NAMA BAPA / PENJAGA 1" || h === "NAMA WARIS");
        
        if (nameIdx > -1) newRow[nameIdx] = student.name;
        if (classIdx > -1) newRow[classIdx] = student.className;
        if (p1PhoneIdx > -1) newRow[p1PhoneIdx] = student.parentPhone;
        if (address1Idx > -1) newRow[address1Idx] = student.address;
        if (icIdx > -1) newRow[icIdx] = student.icNumber;
        if (p1NameIdx > -1) newRow[p1NameIdx] = student.parentName;
        
        sheet.appendRow(newRow);
        updated++;
      }
    }

    return responseJSON({ success: true, message: `${updated} rekod berjaya dihantar ke Google Sheets.` });
  } catch (err) {
    return responseJSON({ success: false, message: err.toString() });
  }
}

function getOrCreateUsersSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName("USERS");
  if (!sheet) {
    sheet = ss.insertSheet("USERS");
    sheet.appendRow(["ID", "NAME", "USERNAME", "PASSWORD", "ROLE", "ASSIGNEDCLASS"]);
  }
  return sheet;
}

function handleGetUsers() {
  try {
    const sheet = getOrCreateUsersSheet();
    const data = sheet.getDataRange().getValues();
    if (data.length < 2) return responseJSON([]);
    
    const users = [];
    for (let i = 1; i < data.length; i++) {
      users.push({
        id: String(data[i][0] || ""),
        name: String(data[i][1] || ""),
        username: String(data[i][2] || ""),
        password: String(data[i][3] || ""),
        role: String(data[i][4] || ""),
        assignedClass: String(data[i][5] || "")
      });
    }
    return responseJSON(users);
  } catch (err) {
    return responseJSON({ success: false, message: err.toString() });
  }
}

function handlePushUsers(payload) {
  try {
    const sheet = getOrCreateUsersSheet();
    const users = payload.users || [];
    
    if (sheet.getLastRow() > 1) {
      sheet.getRange(2, 1, sheet.getLastRow() - 1, 6).clearContent();
    }
    
    if (users.length > 0) {
      const rows = users.map(u => [
        u.id, u.name, u.username, u.password, u.role, u.assignedClass
      ]);
      sheet.getRange(2, 1, rows.length, 6).setValues(rows);
    }
    
    return responseJSON({ success: true, message: "Senarai pengguna berjaya diselaraskan." });
  } catch (err) {
    return responseJSON({ success: false, message: err.toString() });
  }
}

function responseJSON(data) {
  return ContentService.createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

function getOrCreateLogsSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName("LOGS");
  if (!sheet) {
    sheet = ss.insertSheet("LOGS");
    sheet.appendRow(["TIMESTAMP", "USERNAME", "ROLE", "ACTION", "DETAILS"]);
  }
  return sheet;
}

function handleLogActivity(payload) {
  try {
    const sheet = getOrCreateLogsSheet();
    const ts = payload.timestamp || new Date().toISOString();
    const username = payload.username || "System";
    const role = payload.role || "unknown";
    const action = payload.logAction || "UNKNOWN_ACTION";
    const details = payload.details || "";
    
    sheet.appendRow([ts, username, role, action, details]);
    
    return responseJSON({ success: true });
  } catch (err) {
    return responseJSON({ success: false, message: err.toString() });
  }
}

function handleGetLogs() {
  try {
    const sheet = getOrCreateLogsSheet();
    const data = sheet.getDataRange().getValues();
    if (data.length < 2) return responseJSON([]);
    
    const logs = [];
    for (let i = data.length - 1; i >= 1; i--) {
      logs.push({
        timestamp: String(data[i][0] || ""),
        username: String(data[i][1] || ""),
        role: String(data[i][2] || ""),
        action: String(data[i][3] || ""),
        details: String(data[i][4] || "")
      });
      if (logs.length >= 500) break;
    }
    return responseJSON(logs);
  } catch (err) {
    return responseJSON({ success: false, message: err.toString() });
  }
}
