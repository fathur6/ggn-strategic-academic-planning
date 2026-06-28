/**
 * PENGURUSAN INTEGRASI PTA 2026 - PPS
 * Google Sheets + Google Drive (Hierarchy: Sistem Kecil PPS > PTA-PSU 2026 > Lonjakan > Tindakan)
 */

const FOLDER_SISTEM_KECIL_ID = PropertiesService.getScriptProperties().getProperty('FOLDER_SISTEM_KECIL_ID') || '';
const FOLDER_PSU_ID = PropertiesService.getScriptProperties().getProperty('FOLDER_PSU_ID') || FOLDER_SISTEM_KECIL_ID;
const SPREADSHEET_ID = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID') || '';
const SHEET_NAME_STATUS = "Status_Semasa";
const SHEET_NAME_LOG = "Log_Kemaskini";
const SHEET_NAME_MAKLUMAT = "Maklumat_Lonjakan";
const SHEET_NAME_PSU = "Pelan_Strategik_Universiti";

function getSS() {
  return SpreadsheetApp.openById(SPREADSHEET_ID);
}

function doPost(e) {
  try {
    const payload = JSON.parse(e.postData.contents);
    if (payload.action === "processUpdate") return ContentService.createTextOutput(JSON.stringify(processUpdate(payload))).setMimeType(ContentService.MimeType.JSON);
    if (payload.action === "deleteEntry") return ContentService.createTextOutput(JSON.stringify(deleteEntryRecord(payload))).setMimeType(ContentService.MimeType.JSON);
    if (payload.action === "updateKpiInfo") return ContentService.createTextOutput(JSON.stringify(updateKpiInfo(payload))).setMimeType(ContentService.MimeType.JSON);
    if (payload.action === "deleteKpiInfo") return ContentService.createTextOutput(JSON.stringify(deleteKpiInfoRecord(payload))).setMimeType(ContentService.MimeType.JSON);
    if (payload.action === "processUpdatePSU") return ContentService.createTextOutput(JSON.stringify(processUpdatePSU(payload))).setMimeType(ContentService.MimeType.JSON);
    return ContentService.createTextOutput(JSON.stringify({success: false, message: "Aksi tidak dikenali."})).setMimeType(ContentService.MimeType.JSON);
  } catch(err) {
    return ContentService.createTextOutput(JSON.stringify({success: false, message: "Ralat HTTP: " + err.message})).setMimeType(ContentService.MimeType.JSON);
  }
}

function doGet(e) {
  if (e && e.parameter && e.parameter.action === "getKpiStatus") {
    return ContentService.createTextOutput(JSON.stringify({ success: true, data: getKpiStatus(), kpiInfo: getKpiInfoList() })).setMimeType(ContentService.MimeType.JSON);
  }
  if (e && e.parameter && e.parameter.action === "getPSUStatus") {
    return ContentService.createTextOutput(JSON.stringify({ success: true, data: getPelanStrategikData() })).setMimeType(ContentService.MimeType.JSON);
  }

  var userEmail = "";
  try { userEmail = Session.getActiveUser().getEmail(); } catch(ex) { userEmail = ""; }
  if (!userEmail) userEmail = "";

  var props = PropertiesService.getScriptProperties();
  if (!props.getProperty('PSU_SETUP_DONE')) {
    try { setupPelanStrategikUniversiti(); props.setProperty('PSU_SETUP_DONE', 'true'); } catch(e) {}
  }
  var adminListJson = props.getProperty('ADMIN_EMAILS');
  if (!adminListJson) {
    var defaultAdmins = ["SET_ADMIN_EMAILS_IN_SCRIPT_PROPERTIES"];
    adminListJson = JSON.stringify(defaultAdmins);
    try { props.setProperty('ADMIN_EMAILS', adminListJson); } catch(ex) {}
  }

  var template = HtmlService.createTemplateFromFile('index');
  template.userEmail = userEmail;
  template.adminListJson = adminListJson;
  template.PSU_DATA = JSON.stringify(getPelanStrategikData());

  return template.evaluate()
      .setTitle('Dashboard Pemantauan PTA 2026 - PPS')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1.0');
}

function updateKpiInfo(payload) {
  const ss = getSS();
  const sheet = ss.getSheetByName(SHEET_NAME_MAKLUMAT);
  if (!sheet) return { success: false, message: "Sheet Maklumat_Lonjakan belum dibina." };
  const info = payload.kpiInfo;
  const data = sheet.getDataRange().getValues();
  let rowIndex = -1;
  for (let i = 1; i < data.length; i++) { if (data[i][0] === info.idKpi) { rowIndex = i + 1; break; } }
  if (rowIndex !== -1) {
    sheet.getRange(rowIndex, 2).setValue(info.lonjakan); sheet.getRange(rowIndex, 3).setValue(info.initiative); sheet.getRange(rowIndex, 4).setValue(info.activity); sheet.getRange(rowIndex, 5).setValue(info.od); sheet.getRange(rowIndex, 6).setValue(info.pi); sheet.getRange(rowIndex, 7).setValue(info.target); sheet.getRange(rowIndex, 8).setValue(info.unit); sheet.getRange(rowIndex, 9).setValue(info.evidence); sheet.getRange(rowIndex, 10).setValue(info.q1_target); sheet.getRange(rowIndex, 11).setValue(info.q2_target); sheet.getRange(rowIndex, 12).setValue(info.q3_target); sheet.getRange(rowIndex, 13).setValue(info.q4_target);
  } else {
    sheet.appendRow([info.idKpi, info.lonjakan, info.initiative, info.activity, info.od, info.pi, info.target, info.unit, info.evidence, info.q1_target, info.q2_target, info.q3_target, info.q4_target]);
  }
  try { manageFolderStructure(info.idKpi); } catch(e) { console.error("Ralat mengurus folder: " + e.message); }
  return { success: true, message: "Maklumat KPI dan Folder berjaya dikemaskini." };
}

function deleteKpiInfoRecord(payload) {
  const ss = getSS(); const idKpi = payload.idKpi;
  try { 
    try { const folder = manageFolderStructure(idKpi); if (folder) folder.setTrashed(true); } catch(e) {}
    const sheetMaklumat = ss.getSheetByName(SHEET_NAME_MAKLUMAT); if (sheetMaklumat) { const dataM = sheetMaklumat.getDataRange().getValues(); for (let i = dataM.length - 1; i > 0; i--) { if (dataM[i][0] === idKpi) sheetMaklumat.deleteRow(i + 1); } }
    const sheetStatus = ss.getSheetByName(SHEET_NAME_STATUS); if (sheetStatus) { const dataS = sheetStatus.getDataRange().getValues(); for (let i = dataS.length - 1; i > 0; i--) { if (dataS[i][0] === idKpi) sheetStatus.deleteRow(i + 1); } }
    const sheetLog = ss.getSheetByName(SHEET_NAME_LOG); if (sheetLog) { const dataL = sheetLog.getDataRange().getValues(); for (let i = dataL.length - 1; i > 0; i--) { if (dataL[i][3] === idKpi) sheetLog.deleteRow(i + 1); } }
    return { success: true, message: "Tindakan dan rekod berkaitan berjaya dibuang." };
  } catch (err) { return { success: false, message: "Ralat membuang tindakan: " + err.toString() }; }
}

function getKpiInfoList() {
  const ss = getSS(); const sheet = ss.getSheetByName(SHEET_NAME_MAKLUMAT); if (!sheet) return [];
  const data = sheet.getDataRange().getValues(); let infoList = [];
  for (let i = 1; i < data.length; i++) { if (data[i][0]) infoList.push({ idKpi: data[i][0], lonjakan: data[i][1], initiative: data[i][2], activity: data[i][3], od: data[i][4], pi: data[i][5], target: data[i][6], unit: data[i][7], evidence: data[i][8], q1_target: data[i][9], q2_target: data[i][10], q3_target: data[i][11], q4_target: data[i][12] }); }
  return infoList;
}

function manageFolderStructure(idKpi) {
  if (!FOLDER_SISTEM_KECIL_ID) throw new Error("FOLDER_SISTEM_KECIL_ID not set in Script Properties");
  const details = getKpiDetails(idKpi); const safeActivity = details.activity.replace(/[\\/:\*\?"<>|]/g, '-').substring(0, 100);
  const rootSistemFolder = DriveApp.getFolderById(FOLDER_SISTEM_KECIL_ID); let ptaYearFolder = findOrCreateSubFolder(rootSistemFolder, "PTA-PSU 2026");
  const prefixMatch = details.lonjakan.match(/Lonjakan \d+/i); const lonjakanKey = prefixMatch ? prefixMatch[0] : details.lonjakan;
  let lonjakanFolder = findOrCreateSubFolder(ptaYearFolder, lonjakanKey, details.lonjakan);
  const expectedActivityName = idKpi + "_" + safeActivity; let activityFolder = null;
  const aFolders = lonjakanFolder.getFolders(); while (aFolders.hasNext()) { const folder = aFolders.next(); if (folder.getName().startsWith(idKpi)) { activityFolder = folder; if (folder.getName() !== expectedActivityName) folder.setName(expectedActivityName); break; } }
  if (!activityFolder) activityFolder = lonjakanFolder.createFolder(expectedActivityName);
  return activityFolder;
}

function findOrCreateSubFolder(parent, searchKey, fullNameIfNew) { const folders = parent.getFolders(); while (folders.hasNext()) { const f = folders.next(); if (f.getName().toLowerCase().includes(searchKey.toLowerCase())) return f; } return parent.createFolder(fullNameIfNew || searchKey); }

function getKpiDetails(idKpi) { const ss = getSS(); const sheet = ss.getSheetByName(SHEET_NAME_MAKLUMAT); if (sheet) { const data = sheet.getDataRange().getValues(); for (let i = 1; i < data.length; i++) { if (data[i][0] === idKpi) return { lonjakan: data[i][1], activity: data[i][3] }; } } return { lonjakan: "Lonjakan Baru", activity: "Aktiviti " + idKpi }; }

function processUpdate(payload) {
  const ss = getSS(); const sheetLog = ss.getSheetByName(SHEET_NAME_LOG); const sheetStatus = ss.getSheetByName(SHEET_NAME_STATUS);
  const entry = payload.entryData; const qData = payload.quarterData; let fileUrls = [], fileNames = [], hasNewFiles = false;
  try { const targetFolder = manageFolderStructure(payload.idKpi);
    if (payload.filesData && payload.filesData.length > 0) { hasNewFiles = true; payload.filesData.forEach((fileObj) => { const blob = Utilities.newBlob(Utilities.base64Decode(fileObj.base64), fileObj.type, fileObj.name); const file = targetFolder.createFile(blob); try { file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); } catch(e) {} fileUrls.push(file.getUrl()); fileNames.push(file.getName()); }); }
    const combinedUrls = fileUrls.join(", "), combinedNames = fileNames.join(", "); const dataLog = sheetLog.getDataRange().getValues(); let isUpdate = false;
    for (let i = 1; i < dataLog.length; i++) { if (dataLog[i][0] == entry.id) { sheetLog.getRange(i+1,2).setValue(new Date()); sheetLog.getRange(i+1,3).setValue(payload.userEmail); sheetLog.getRange(i+1,5).setValue(entry.nama); sheetLog.getRange(i+1,6).setValue(entry.tarikh); sheetLog.getRange(i+1,7).setValue(entry.kemajuan); sheetLog.getRange(i+1,8).setValue(entry.catatan); sheetLog.getRange(i+1,11).setValue(entry.fakulti||""); sheetLog.getRange(i+1,12).setValue(entry.tahap||""); sheetLog.getRange(i+1,13).setValue(entry.matrik||""); if (hasNewFiles) { sheetLog.getRange(i+1,9).setValue(combinedNames); sheetLog.getRange(i+1,10).setValue(combinedUrls); } isUpdate = true; break; } }
    if (!isUpdate) sheetLog.appendRow([entry.id, new Date(), payload.userEmail, payload.idKpi, entry.nama, entry.tarikh, entry.kemajuan, entry.catatan, combinedNames, combinedUrls, entry.fakulti||"", entry.tahap||"", entry.matrik||""]);
    updateTotalStatus(payload.idKpi, sheetLog, sheetStatus, hasNewFiles ? combinedUrls : null, qData);
    return { success: true, message: "Pencapaian dan fail berjaya disimpan." };
  } catch (err) { return { success: false, message: "Ralat Semasa Menyimpan: " + err.toString() }; }
}

function deleteEntryRecord(payload) { const ss = getSS(); const sheetLog = ss.getSheetByName(SHEET_NAME_LOG); const sheetStatus = ss.getSheetByName(SHEET_NAME_STATUS); try { const dataLog = sheetLog.getDataRange().getValues(); let rowIndexToDelete = -1; for (let i = 1; i < dataLog.length; i++) { if (dataLog[i][0] == payload.entryId) { rowIndexToDelete = i+1; break; } } if (rowIndexToDelete !== -1) { sheetLog.deleteRow(rowIndexToDelete); updateTotalStatus(payload.idKpi, sheetLog, sheetStatus, null, null); return { success: true, message: "Rekod berjaya dibuang." }; } else return { success: false, message: "Rekod tidak dijumpai." }; } catch (err) { return { success: false, message: "Ralat Pemadaman: " + err.toString() }; } }

function updateTotalStatus(idKpi, sheetLog, sheetStatus, latestFileUrls, quarterData) {
  const logs = sheetLog.getDataRange().getValues(); let total = 0; for (let i = 1; i < logs.length; i++) { if (logs[i][3] === idKpi) total += (parseFloat(logs[i][6]) || 0); }
  const statusData = sheetStatus.getDataRange().getValues(); let found = false;
  for (let i = 1; i < statusData.length; i++) { if (statusData[i][0] === idKpi) { found = true; const sasaran = parseFloat(statusData[i][3]) || 1; sheetStatus.getRange(i+1,6).setValue(total); sheetStatus.getRange(i+1,7).setValue(Math.min(100,(total/sasaran)*100)); if (latestFileUrls) sheetStatus.getRange(i+1,8).setValue(latestFileUrls.split(',')[0].trim()); if (quarterData) { sheetStatus.getRange(i+1,9).setValue(quarterData.q1_val||""); sheetStatus.getRange(i+1,10).setValue(quarterData.q1_status||""); sheetStatus.getRange(i+1,11).setValue(quarterData.q2_val||""); sheetStatus.getRange(i+1,12).setValue(quarterData.q2_status||""); sheetStatus.getRange(i+1,13).setValue(quarterData.q3_val||""); sheetStatus.getRange(i+1,14).setValue(quarterData.q3_status||""); sheetStatus.getRange(i+1,15).setValue(quarterData.q4_val||""); sheetStatus.getRange(i+1,16).setValue(quarterData.q4_status||""); sheetStatus.getRange(i+1,17).setValue(quarterData.kpi_tercapai||false); } break; } }
  if (!found) { let newRow = [idKpi,"","","","",total,0,latestFileUrls?latestFileUrls.split(',')[0].trim():""]; if (quarterData) newRow.push(quarterData.q1_val||"",quarterData.q1_status||"",quarterData.q2_val||"",quarterData.q2_status||"",quarterData.q3_val||"",quarterData.q3_status||"",quarterData.q4_val||"",quarterData.q4_status||"",quarterData.kpi_tercapai||false); else newRow.push("","","","","","","","",false); sheetStatus.appendRow(newRow); }
}

function getKpiStatus() {
  const ss = getSS(); const res = {}; const sheetStatus = ss.getSheetByName(SHEET_NAME_STATUS); const sheetLog = ss.getSheetByName(SHEET_NAME_LOG); if (!sheetStatus || !sheetLog) return res;
  const statusData = sheetStatus.getDataRange().getValues(); const logData = sheetLog.getDataRange().getValues();
  for (let i = 1; i < statusData.length; i++) { res[statusData[i][0]] = { current: statusData[i][5], quarterData: { q1_val: statusData[i][8]||"", q1_status: statusData[i][9]||"", q2_val: statusData[i][10]||"", q2_status: statusData[i][11]||"", q3_val: statusData[i][12]||"", q3_status: statusData[i][13]||"", q4_val: statusData[i][14]||"", q4_status: statusData[i][15]||"", kpi_tercapai: statusData[i][16]===true||statusData[i][16]==="TRUE" }, entries: [] }; }
  for (let i = 1; i < logData.length; i++) { 
    const idKpi = logData[i][3]; 
    if (res[idKpi]) {
      let safeDateStr = "";
      if (logData[i][5]) {
         try { safeDateStr = Utilities.formatDate(new Date(logData[i][5]),"GMT+8","yyyy-MM-dd"); }
         catch(e) { safeDateStr = logData[i][5]; } // Fallback jika tarikh kosong/rosak
      }
      res[idKpi].entries.push({ id: logData[i][0], nama: logData[i][4], tarikh: safeDateStr, kemajuan: logData[i][6], catatan: logData[i][7], fileName: logData[i][8], fileUrl: logData[i][9], fakulti: logData[i][10]||"", tahap: logData[i][11]||"", matrik: logData[i][12]||"" }); 
    } 
  }
  return res;
}

function setupMaklumatLonjakan() { const ss = getSS(); let sheet = ss.getSheetByName(SHEET_NAME_MAKLUMAT); if (!sheet) sheet = ss.insertSheet(SHEET_NAME_MAKLUMAT); const headers = ["ID KPI","Lonjakan","Inisiatif","Aktiviti","Definisi Operasi (OD)","Petunjuk Prestasi (PI)","Sasaran Keseluruhan","Unit","Bukti","Sasaran Q1","Sasaran Q2","Sasaran Q3","Sasaran Q4"]; sheet.getRange(1,1,1,headers.length).setValues([headers]).setBackground("#FBB03B").setFontColor("black").setFontWeight("bold"); if (sheet.getLastRow()===1) { const d = [["L7-KPI-01","Lonjakan 7: Keunggulan Global","Global Networking","Bilangan Pemeriksa Luar Antarabangsa","...",72,"penilai luar negara","Surat Pelantikan","18","36","54","72"],["L7-KPI-02","Lonjakan 7: Keunggulan Global","Global Networking","Bilangan Staf Akademik Antarabangsa sebagai Penyelia Bersama","...",200,"penyelia","Surat Pelantikan / Borang GS-04","50","100","150","200"],["L7-KPI-03","Lonjakan 7: Keunggulan Global","FlexS@Global","Dialog Antarabangsa","...",5,"kolaborator","Laporan Pelaksanaan Program","MoU","5 penceramah","International Dialogue","-"],["L7-KPI-10","Lonjakan 7: Keunggulan Global","Global university engagement","Enrollment of international students","...",1500,"Pelajar","Complete list","-","Promotion","750","1500"],["L8-KPI-04","Lonjakan 8: Pendidikan Fleksibel & PSH","Program Akademik Fleksibel","Kursus Fleksibel","...",1,"Course","Surat / Dokumen","-","Complete Course Plan","1 MOOC course","-"],["L8-KPI-05","Lonjakan 8: Pendidikan Fleksibel & PSH","Program PSH Bukan Formal","Program Bukan Formal","...",2,"program","Laporan Lengkap","-","-","1 program","2 program"],["L8-KPI-06","Lonjakan 8: Pendidikan Fleksibel & PSH","Micro-Credentials","Penawaran MC Formal","...",1,"Modul MC","Surat/Dokumen","-","Complete Modul Plan","1 Modul MC","-"],["L8-KPI-07","Lonjakan 8: Pendidikan Fleksibel & PSH","Program PSH Informal","Program PSH Informal","...",2,"program","Laporan Lengkap","-","1 program","-","2 program"],["L10-KPI-08","Lonjakan 10: Penyampaian Responsif & Dinamik","Nexus Learning Landscape","Pewujudan Green Learning Common Spaces","...",1,"Ruang","Laporan Penubuhan","-","Kertas Kerja","Kelulusan","1 Ruang"],["L10-KPI-09","Lonjakan 10: Penyampaian Responsif & Dinamik","E-Learning Platform","Pembangunan Fungsi KELIP","...",1,"Fungsi","Laporan Aktiviti","-","-","Fungsi Tambahan AI","-"]]; sheet.getRange(2,1,d.length,headers.length).setValues(d); } sheet.autoResizeColumns(1,headers.length); let sheetLog = ss.getSheetByName(SHEET_NAME_LOG); if (!sheetLog) sheetLog = ss.insertSheet(SHEET_NAME_LOG); const hdr = ["ID Entry","Timestamp","User Email","ID KPI","Nama Projek/Laporan","Tarikh","Kemajuan","Catatan","Nama Fail","URL Fail","Fakulti","Tahap Pengajian","No. Matrik"]; sheetLog.getRange(1,1,1,hdr.length).setValues([hdr]).setBackground("#002C5F").setFontColor("white").setFontWeight("bold"); return "Berjaya!"; }

function benarkanAksesDrive() { try { var root = DriveApp.getFolderById(FOLDER_SISTEM_KECIL_ID); var sub = findOrCreateSubFolder(root, "PTA-PSU 2026"); var test = sub.createFolder("Folder_Ujian_Padam"); test.setTrashed(true); Logger.log("BERJAYA: Susunan folder PTA-PSU 2026 sedia digunakan!"); } catch(e) { Logger.log("RALAT: " + e.message); throw e; } }

function setupPelanStrategikUniversiti() {
  const ss = getSS();
  let sheet = ss.getSheetByName(SHEET_NAME_PSU);
  if (!sheet) sheet = ss.insertSheet(SHEET_NAME_PSU);
  const headers = ["No","Strategi Pelaksanaan","Inisiatif","No Aktiviti","Aktiviti","Indikator Pengukuran","Sasaran KPI 2026 (Universiti)","Bilangan Sasaran KPI 2026 (PTj)","Status","Pencapaian Semasa","Bukti/URL","Catatan"];
  sheet.getRange(1,1,1,headers.length).setValues([headers]).setBackground("#002C5F").setFontColor("white").setFontWeight("bold");
  if (sheet.getLastRow() === 1) {
    const data = [
      ["PSU-01","SP1: Future Ready Curriculum","Curriculum Innovations (Flexible Academic Programmes)","1","Flexible Course (MOOC)","Number of courses","4 (3 MPU Courses, 1 RMC (PPS))","3 MPU Courses, 1 RMC (PPS)","","","",""],
      ["PSU-02","SP1: Future Ready Curriculum","","2","Apppointment of international academic staff - External Examiners","Percentage of international External Examiners (based on confirmed Viva-Voce candidates)","Faculty (except FUPL): 40%","40%","","","",""],
      ["PSU-03","SP1: Future Ready Curriculum","","2","Appointment of international academic staff - Co-Supervisor (PG)","Number of international academic staff as Co-Supervisor (PG)","200 international academic staff of Co-Supervisor","Contributed by any faculties except FUPL","","","",""],
      ["PSU-04","SP1: Future Ready Curriculum","Global Networking","3","Enrollment of international students in academic programmes","Number of international student enrollments (60% (900) from other countries than Indonesia)","1500","1500","","","",""],
      ["PSU-05","SP1: Future Ready Curriculum","Industry-based Learning","4","Industrial involvement in student assessment","Number of industry co-supervisors or External Examiners","10 industry co-supervisors or External Examiners","1 (Fakulti terlibat)","","","",""],
      ["PSU-06","SP2: Sustainable Learning Ecosystem","Establish a Green Learning Common Spaces that leverages AI-driven adaptive environment","5","Number of learning common spaces","2 lounge (KB and KGB), PWB:1, PPS:1","1 Green learning common- PWB","0","","","",""]
    ];
    sheet.getRange(2,1,data.length,headers.length).setValues(data);
  }
  sheet.autoResizeColumns(1, headers.length);
  return "Pelan Strategik Universiti sheet siap.";
}

function getPelanStrategikData() {
  const ss = getSS(); const sheet = ss.getSheetByName(SHEET_NAME_PSU);
  if (!sheet) return [];
  const data = sheet.getDataRange().getValues();
  const result = [];
  for (let i = 1; i < data.length; i++) {
    if (data[i][0]) {
      result.push({
        id: data[i][0], no: data[i][0], strategi: data[i][1]||"", inisiatif: data[i][2]||"",
        noAktiviti: data[i][3]||"", aktiviti: data[i][4]||"", indikator: data[i][5]||"",
        sasaranUniv: data[i][6]||"", sasaranPTj: data[i][7]||"",
        status: data[i][8]||"Belum Mula", pencapaian: data[i][9]||"", bukti: data[i][10]||"", catatan: data[i][11]||""
      });
    }
  }
  return result;
}

function processUpdatePSU(payload) {
  const ss = getSS(); const sheet = ss.getSheetByName(SHEET_NAME_PSU);
  if (!sheet) return { success: false, message: "Sheet PSU belum dibina." };
  const data = sheet.getDataRange().getValues();
  let rowIndex = -1;
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === payload.id) { rowIndex = i + 1; break; }
  }
  if (rowIndex !== -1) {
    if (payload.status !== undefined) sheet.getRange(rowIndex, 9).setValue(payload.status);
    if (payload.pencapaian !== undefined) sheet.getRange(rowIndex, 10).setValue(payload.pencapaian);
    if (payload.bukti !== undefined) sheet.getRange(rowIndex, 11).setValue(payload.bukti);
    if (payload.catatan !== undefined) sheet.getRange(rowIndex, 12).setValue(payload.catatan);
  } else {
    sheet.appendRow([payload.id, payload.strategi||"", payload.inisiatif||"", payload.noAktiviti||"", payload.aktiviti||"", payload.indikator||"", payload.sasaranUniv||"", payload.sasaranPTj||"", payload.status||"Belum Mula", payload.pencapaian||"", payload.bukti||"", payload.catatan||""]);
  }
  try { managePSUFolderStructure(payload.id); } catch(e) {}
  return { success: true, message: "Data PSU dikemaskini." };
}

function managePSUFolderStructure(idPSU) {
  if (!FOLDER_PSU_ID) throw new Error("FOLDER_PSU_ID not set");
  const rootFolder = DriveApp.getFolderById(FOLDER_PSU_ID);
  let psuFolder = findOrCreateSubFolder(rootFolder, "Pelan Strategik Universiti");
  let ptaYearFolder = findOrCreateSubFolder(psuFolder, "2026");
  const item = getPelanStrategikData().find(p => p.id === idPSU);
  const safeAktiviti = item ? item.aktiviti.replace(/[\\/:\*\?"<>|]/g, '-').substring(0, 80) : idPSU;
  let activityFolder = findOrCreateSubFolder(ptaYearFolder, idPSU + "_" + safeAktiviti);
  return activityFolder;
}