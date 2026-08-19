const FOLDER_ID_DRIVE = "1cRFISaADgQrLsi6mBEAWtVcQcetmg2Ge";

function doGet(e) {
  const action = e.parameter.action;
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  if (action === "getStudents") {
    let sheet = ss.getSheetByName("Pelajar");
    if (!sheet) sheet = ss.getSheets()[0]; // Fallback ke sheet pertama
    
    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) return responseJSON([]);
    
    const headers = data[0].map(h => h.toString().trim().toUpperCase());
    
    // Cari index lajur berdasarkan format APDM asal atau format ringkas
    const idIdx = headers.findIndex(h => h === "ID MURID" || h === "ID");
    const nameIdx = headers.findIndex(h => h === "NAMA" || h === "NAMA MURID");
    const icIdx = headers.findIndex(h => h === "NO. PENGENALAN" || h === "NO KAD PENGENALAN");
    const classIdx = headers.findIndex(h => h === "NAMA KELAS" || h === "KELAS");
    const genderIdx = headers.findIndex(h => h === "JANTINA");
    const p1NameIdx = headers.findIndex(h => h === "PENJAGA 1" || h === "NAMA BAPA / PENJAGA 1" || h === "NAMA WARIS");
    const p1RelIdx = headers.findIndex(h => h === "HUBUNGAN PENJAGA 1" || h === "HUBUNGAN");
    const p1PhoneIdx = headers.findIndex(h => h === "NO. TEL. BIMBIT PENJAGA 1" || h === "NO TEL BAPA / PENJAGA 1" || h === "NO TEL WARIS");
    const p2NameIdx = headers.findIndex(h => h === "PENJAGA 2" || h === "NAMA IBU / PENJAGA 2");
    const p2PhoneIdx = headers.findIndex(h => h === "NO. TEL. BIMBIT PENJAGA 2" || h === "NO TEL IBU / PENJAGA 2");
    const address1Idx = headers.findIndex(h => h === "ALAMAT 1" || h === "ALAMAT KEDIAMAN" || h === "ALAMAT");
    const address2Idx = headers.findIndex(h => h === "ALAMAT 2");
    const address3Idx = headers.findIndex(h => h === "ALAMAT 3");
    let hostelIdx = headers.findIndex(h => h === "STATUS ASRAMA");
    let teacherIdx = headers.findIndex(h => h === "NAMA GURU KELAS" || h === "GURU KELAS");
    let photoIdx = headers.findIndex(h => h === "GAMBAR");
    
    // Jika lajur GAMBAR tiada, cipta lajur baru di hujung
    if (photoIdx === -1) {
      photoIdx = headers.length;
      sheet.getRange(1, photoIdx + 1).setValue("GAMBAR");
      headers.push("GAMBAR");
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
        teacherName: teacherIdx > -1 ? String(row[teacherIdx]) : ""
      });
    }
    
    return responseJSON(students);
  }
  
  return responseJSON({ status: "API Google Apps Script SIS Aktif (Mode Selamat)" });
}

function doPost(e) {
  try {
    const contents = JSON.parse(e.postData.contents);
    const action = contents.action;
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName("Pelajar");
    if (!sheet) sheet = ss.getSheets()[0];
    
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(h => h.toString().trim().toUpperCase());
    const idIdx = headers.findIndex(h => h === "ID MURID" || h === "ID");
    
    if (action === "uploadPhoto") {
      const folder = DriveApp.getFolderById(FOLDER_ID_DRIVE);
      const studentId = contents.studentId;
      const base64Data = contents.base64Image.split(",")[1];
      const blob = Utilities.newBlob(Utilities.base64Decode(base64Data), "image/jpeg", studentId + ".jpg");
      
      const file = folder.createFile(blob);
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      const photoUrl = "https://lh3.googleusercontent.com/d/" + file.getId();

      let photoIdx = headers.findIndex(h => h === "GAMBAR");
      if (photoIdx === -1) {
        photoIdx = headers.length;
        sheet.getRange(1, photoIdx + 1).setValue("GAMBAR");
      }

      if (idIdx > -1) {
        const data = sheet.getDataRange().getValues();
        for (let i = 1; i < data.length; i++) {
          if (String(data[i][idIdx]) === String(studentId)) {
            sheet.getRange(i + 1, photoIdx + 1).setValue(photoUrl);
            break;
          }
        }
      }
      return responseJSON({ success: true, photoUrl: photoUrl });
    }
    
    if (action === "syncAllData") {
       sheet.clear();
       sheet.appendRow(["ID MURID", "NO. PENGENALAN", "NAMA", "NAMA KELAS", "JANTINA", "PENJAGA 1", "NO. TEL. BIMBIT PENJAGA 1", "HUBUNGAN PENJAGA 1", "ALAMAT 1", "STATUS ASRAMA", "GAMBAR"]);
       
       contents.students.forEach(st => {
         sheet.appendRow([
           st.id, st.icNumber, st.name, st.className, st.gender, 
           st.parentName, st.parentPhone, st.parentRelation, 
           st.address, st.hostelStatus, st.photoUrl
         ]);
       });
       return responseJSON({ success: true, message: "Data berjaya dihantar ke Sheet secara pukal" });
    }
    
    if (action === "updateStudent") {
       const st = contents.student;
       if (!st || !st.id || idIdx === -1) return responseJSON({ error: "Data tidak sah atau ID tidak dijumpai" });
       
       const data = sheet.getDataRange().getValues();
       for (let i = 1; i < data.length; i++) {
          if (String(data[i][idIdx]) === String(st.id)) {
            const nameIdx = headers.findIndex(h => h === "NAMA" || h === "NAMA MURID");
            const classIdx = headers.findIndex(h => h === "NAMA KELAS" || h === "KELAS");
            const p1NameIdx = headers.findIndex(h => h === "PENJAGA 1" || h === "NAMA BAPA / PENJAGA 1" || h === "NAMA WARIS");
            const p1PhoneIdx = headers.findIndex(h => h === "NO. TEL. BIMBIT PENJAGA 1" || h === "NO TEL BAPA / PENJAGA 1" || h === "NO TEL WARIS");
            const icIdx = headers.findIndex(h => h === "NO. PENGENALAN" || h === "NO KAD PENGENALAN");
            const address1Idx = headers.findIndex(h => h === "ALAMAT 1" || h === "ALAMAT KEDIAMAN" || h === "ALAMAT");
            
            if (nameIdx > -1) sheet.getRange(i + 1, nameIdx + 1).setValue(st.name);
            if (classIdx > -1) sheet.getRange(i + 1, classIdx + 1).setValue(st.className);
            if (p1NameIdx > -1) sheet.getRange(i + 1, p1NameIdx + 1).setValue(st.parentName);
            if (p1PhoneIdx > -1) sheet.getRange(i + 1, p1PhoneIdx + 1).setValue(st.parentPhone);
            if (icIdx > -1) sheet.getRange(i + 1, icIdx + 1).setValue(st.icNumber);
            if (address1Idx > -1) sheet.getRange(i + 1, address1Idx + 1).setValue(st.address);
            
            return responseJSON({ success: true, message: "Pelajar dikemaskini" });
          }
       }
       return responseJSON({ error: "Pelajar tidak ditemui di Sheet" });
    }

    return responseJSON({ error: "Invalid action" });
  } catch (err) {
    return responseJSON({ error: err.toString() });
  }
}

function responseJSON(data) {
  return ContentService.createTextOutput(JSON.stringify(data)).setMimeType(ContentService.MimeType.JSON);
}
