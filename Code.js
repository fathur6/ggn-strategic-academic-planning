/**
 * PENGURUSAN INTEGRASI PTA & PSU 2026 - PPS
 * Google Sheets + Google Drive (Hierarchy: Sistem Kecil PPS > PTA-PSU 2026 > Lonjakan > Tindakan)
 */

const FOLDER_SISTEM_KECIL_ID = PropertiesService.getScriptProperties().getProperty('FOLDER_SISTEM_KECIL_ID') || '';
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
    return ContentService.createTextOutput(JSON.stringify({success: false, message: "Aksi tidak dikenali."})).setMimeType(ContentService.MimeType.JSON);
  } catch(err) {
    return ContentService.createTextOutput(JSON.stringify({success: false, message: "Ralat HTTP: " + err.message})).setMimeType(ContentService.MimeType.JSON);
  }
}

function doGet(e) {
  if (e && e.parameter && e.parameter.action === "getKpiStatus") {
    return ContentService.createTextOutput(JSON.stringify({ success: true, data: getKpiStatus(), kpiInfo: getKpiInfoList() })).setMimeType(ContentService.MimeType.JSON);
  }

  return HtmlService.createHtmlOutputFromFile('index')
      .setTitle('Dashboard Pemantauan PTA 2026 - PPS')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1.0');
}

function getSessionData() {
  var userEmail = "";
  try { userEmail = Session.getActiveUser().getEmail(); } catch(ex) {}
  
  var props = PropertiesService.getScriptProperties();
  if (!props.getProperty('PSU_SETUP_DONE')) {
    try { setupPelanStrategikUniversiti(); props.setProperty('PSU_SETUP_DONE', 'true'); } catch(e) {}
  }
  
  var adminListJson = props.getProperty('ADMIN_EMAILS');
  var admins = ["SET_ADMIN_EMAILS_IN_SCRIPT_PROPERTIES"];
  if (adminListJson) {
    try { admins = JSON.parse(adminListJson); } catch(e) {}
  } else {
    props.setProperty('ADMIN_EMAILS', JSON.stringify(admins));
  }
  
  return { email: userEmail, admins: admins };
}

function updateKpiInfo(payload) {
  const ss = getSS();
  const info = payload.kpiInfo;
  const isPsu = info.idKpi && info.idKpi.startsWith("PSU");
  const sheet = isPsu ? ss.getSheetByName(SHEET_NAME_PSU) : ss.getSheetByName(SHEET_NAME_MAKLUMAT);
  if (!sheet) return { success: false, message: "Sheet rujukan belum dibina." };
  
  const data = sheet.getDataRange().getValues();
  let rowIndex = -1;
  for (let i = 1; i < data.length; i++) { if (data[i][0] === info.idKpi) { rowIndex = i + 1; break; } }
  
  if (rowIndex !== -1) {
    if (isPsu) {
      sheet.getRange(rowIndex, 2).setValue(info.lonjakan); sheet.getRange(rowIndex, 3).setValue(info.initiative); sheet.getRange(rowIndex, 5).setValue(info.activity); sheet.getRange(rowIndex, 6).setValue(info.pi); sheet.getRange(rowIndex, 8).setValue(info.target);
    } else {
      sheet.getRange(rowIndex, 2).setValue(info.lonjakan); sheet.getRange(rowIndex, 3).setValue(info.initiative); sheet.getRange(rowIndex, 4).setValue(info.activity); sheet.getRange(rowIndex, 5).setValue(info.od); sheet.getRange(rowIndex, 6).setValue(info.pi); sheet.getRange(rowIndex, 7).setValue(info.target); sheet.getRange(rowIndex, 8).setValue(info.unit); sheet.getRange(rowIndex, 9).setValue(info.evidence); sheet.getRange(rowIndex, 10).setValue(info.q1_target); sheet.getRange(rowIndex, 11).setValue(info.q2_target); sheet.getRange(rowIndex, 12).setValue(info.q3_target); sheet.getRange(rowIndex, 13).setValue(info.q4_target);
    }
  } else {
    if (isPsu) sheet.appendRow([info.idKpi, info.lonjakan, info.initiative, "", info.activity, info.pi, "", info.target, "Belum Mula", "0", "", ""]);
    else sheet.appendRow([info.idKpi, info.lonjakan, info.initiative, info.activity, info.od, info.pi, info.target, info.unit, info.evidence, info.q1_target, info.q2_target, info.q3_target, info.q4_target]);
  }
  
  try { manageFolderStructure(info.idKpi); } catch(e) {}
  return { success: true, message: "Maklumat berjaya dikemaskini." };
}

function deleteKpiInfoRecord(payload) {
  const ss = getSS(); const idKpi = payload.idKpi;
  try { 
    try { const folder = manageFolderStructure(idKpi); if (folder) folder.setTrashed(true); } catch(e) {}
    if (idKpi && idKpi.startsWith("PSU")) {
      const sheetPsu = ss.getSheetByName(SHEET_NAME_PSU); if (sheetPsu) { const dm = sheetPsu.getDataRange().getValues(); for (let i = dm.length - 1; i > 0; i--) { if (dm[i][0] === idKpi) sheetPsu.deleteRow(i + 1); } }
    } else {
      const sheetM = ss.getSheetByName(SHEET_NAME_MAKLUMAT); if (sheetM) { const dm = sheetM.getDataRange().getValues(); for (let i = dm.length - 1; i > 0; i--) { if (dm[i][0] === idKpi) sheetM.deleteRow(i + 1); } }
    }
    const sheetS = ss.getSheetByName(SHEET_NAME_STATUS); if (sheetS) { const ds = sheetS.getDataRange().getValues(); for (let i = ds.length - 1; i > 0; i--) { if (ds[i][0] === idKpi) sheetS.deleteRow(i + 1); } }
    const sheetL = ss.getSheetByName(SHEET_NAME_LOG); if (sheetL) { const dl = sheetL.getDataRange().getValues(); for (let i = dl.length - 1; i > 0; i--) { if (dl[i][3] === idKpi) sheetL.deleteRow(i + 1); } }
    return { success: true, message: "Tindakan dibuang." };
  } catch (err) { return { success: false, message: "Ralat: " + err.toString() }; }
}

function getKpiInfoList() {
  const ss = getSS(); const sheet = ss.getSheetByName(SHEET_NAME_MAKLUMAT); if (!sheet) return [];
  const data = sheet.getDataRange().getValues(); let infoList = [];
  for (let i = 1; i < data.length; i++) { if (data[i][0]) infoList.push({ idKpi: data[i][0], type: 'PTA', lonjakan: data[i][1], initiative: data[i][2], activity: data[i][3], od: data[i][4], pi: data[i][5], target: data[i][6], unit: data[i][7], evidence: data[i][8], q1_target: data[i][9], q2_target: data[i][10], q3_target: data[i][11], q4_target: data[i][12] }); }
  return infoList;
}

function getPelanStrategikData() {
  const ss = getSS(); const sheet = ss.getSheetByName(SHEET_NAME_PSU); if (!sheet) return [];
  const data = sheet.getDataRange().getValues(); const result = [];
  for (let i = 1; i < data.length; i++) {
    if (data[i][0]) {
      result.push({ idKpi: data[i][0], type: 'PSU', lonjakan: data[i][1]||"", initiative: data[i][2]||"", activity: data[i][4]||"", pi: data[i][5]||"", sasaranUniv: data[i][6]||"", target: data[i][7]||"", status: data[i][8]||"Belum Mula", pencapaian: data[i][9]||"", evidence: data[i][10]||"", od: data[i][11]||"" });
    }
  }
  return result;
}

function getKpiStatus() {
  const ss = getSS(); const res = {}; const sheetS = ss.getSheetByName(SHEET_NAME_STATUS); const sheetL = ss.getSheetByName(SHEET_NAME_LOG); if (!sheetS || !sheetL) return res;
  const sData = sheetS.getDataRange().getValues(); const lData = sheetL.getDataRange().getValues();
  for (let i = 1; i < sData.length; i++) { res[sData[i][0]] = { current: sData[i][5], quarterData: { q1_val: sData[i][8]||"", q1_status: sData[i][9]||"", q2_val: sData[i][10]||"", q2_status: sData[i][11]||"", q3_val: sData[i][12]||"", q3_status: sData[i][13]||"", q4_val: sData[i][14]||"", q4_status: sData[i][15]||"", kpi_tercapai: sData[i][16]===true||sData[i][16]==="TRUE" }, entries: [] }; }
  for (let i = 1; i < lData.length; i++) { 
    const idKpi = lData[i][3]; 
    if (res[idKpi]) {
      let sd = "";
      if (lData[i][5]) { try { sd = Utilities.formatDate(new Date(lData[i][5]),"GMT+8","yyyy-MM-dd"); } catch(e) { sd = lData[i][5]; } }
      res[idKpi].entries.push({ id: lData[i][0], nama: lData[i][4], tarikh: sd, kemajuan: lData[i][6], catatan: lData[i][7], fileName: lData[i][8], fileUrl: lData[i][9], fakulti: lData[i][10]||"", tahap: lData[i][11]||"", matrik: lData[i][12]||"" }); 
    } 
  }
  return res;
}

function manageFolderStructure(idKpi) {
  if (!FOLDER_SISTEM_KECIL_ID) throw new Error("FOLDER_SISTEM_KECIL_ID not set");
  const details = getKpiDetails(idKpi); const safeActivity = details.activity.replace(/[\\/:\*\?"<>|]/g, '-').substring(0, 100);
  const rootFolder = DriveApp.getFolderById(FOLDER_SISTEM_KECIL_ID); let yearFolder = findOrCreateSubFolder(rootFolder, "PTA-PSU 2026");
  let lonjKey = "Pelan Strategik Universiti";
  if (details.lonjakan && !details.lonjakan.includes("SP")) { const m = details.lonjakan.match(/Lonjakan \d+/i); lonjKey = m ? m[0] : details.lonjakan; }
  let lonjFolder = findOrCreateSubFolder(yearFolder, lonjKey, details.lonjakan);
  const expName = idKpi + "_" + safeActivity; let actFolder = null;
  const folders = lonjFolder.getFolders(); while (folders.hasNext()) { const f = folders.next(); if (f.getName().startsWith(idKpi)) { actFolder = f; if (f.getName() !== expName) f.setName(expName); break; } }
  if (!actFolder) actFolder = lonjFolder.createFolder(expName);
  return actFolder;
}

function findOrCreateSubFolder(parent, key, fullName) { const folders = parent.getFolders(); while (folders.hasNext()) { const f = folders.next(); if (f.getName().toLowerCase().includes(key.toLowerCase())) return f; } return parent.createFolder(fullName || key); }

function getKpiDetails(idKpi) { 
  const ss = getSS(); 
  let sheet = ss.getSheetByName(SHEET_NAME_MAKLUMAT); if (sheet) { const d = sheet.getDataRange().getValues(); for (let i = 1; i < d.length; i++) { if (d[i][0] === idKpi) return { lonjakan: d[i][1], activity: d[i][3] }; } } 
  sheet = ss.getSheetByName(SHEET_NAME_PSU); if (sheet) { const d = sheet.getDataRange().getValues(); for (let i = 1; i < d.length; i++) { if (d[i][0] === idKpi) return { lonjakan: d[i][1], activity: d[i][4] }; } } 
  return { lonjakan: "Aktiviti Baru", activity: "Aktiviti " + idKpi }; 
}

function processUpdate(payload) {
  const ss = getSS(); const sheetLog = ss.getSheetByName(SHEET_NAME_LOG); const sheetStatus = ss.getSheetByName(SHEET_NAME_STATUS);
  const entry = payload.entryData; const qData = payload.quarterData; let fileUrls = [], fileNames = [], hasNewFiles = false;
  try { const targetFolder = manageFolderStructure(payload.idKpi);
    if (payload.filesData && payload.filesData.length > 0) { hasNewFiles = true; payload.filesData.forEach((fo) => { const blob = Utilities.newBlob(Utilities.base64Decode(fo.base64), fo.type, fo.name); const file = targetFolder.createFile(blob); try { file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); } catch(e) {} fileUrls.push(file.getUrl()); fileNames.push(file.getName()); }); }
    const cUrls = fileUrls.join(", "), cNames = fileNames.join(", "); const logData = sheetLog.getDataRange().getValues(); let isUpdate = false;
    for (let i = 1; i < logData.length; i++) { if (logData[i][0] == entry.id) { sheetLog.getRange(i+1,2).setValue(new Date()); sheetLog.getRange(i+1,3).setValue(payload.userEmail); sheetLog.getRange(i+1,5).setValue(entry.nama); sheetLog.getRange(i+1,6).setValue(entry.tarikh); sheetLog.getRange(i+1,7).setValue(entry.kemajuan); sheetLog.getRange(i+1,8).setValue(entry.catatan); sheetLog.getRange(i+1,11).setValue(entry.fakulti||""); sheetLog.getRange(i+1,12).setValue(entry.tahap||""); sheetLog.getRange(i+1,13).setValue(entry.matrik||""); if (hasNewFiles) { sheetLog.getRange(i+1,9).setValue(cNames); sheetLog.getRange(i+1,10).setValue(cUrls); } isUpdate = true; break; } }
    if (!isUpdate) sheetLog.appendRow([entry.id, new Date(), payload.userEmail, payload.idKpi, entry.nama, entry.tarikh, entry.kemajuan, entry.catatan, cNames, cUrls, entry.fakulti||"", entry.tahap||"", entry.matrik||""]);
    updateTotalStatus(payload.idKpi, sheetLog, sheetStatus, hasNewFiles ? cUrls : null, qData);
    return { success: true, message: "Pencapaian dan fail berjaya disimpan." };
  } catch (err) { return { success: false, message: "Ralat: " + err.toString() }; }
}

function deleteEntryRecord(payload) { const ss = getSS(); const sheetLog = ss.getSheetByName(SHEET_NAME_LOG); const sheetStatus = ss.getSheetByName(SHEET_NAME_STATUS); try { const ld = sheetLog.getDataRange().getValues(); let ri = -1; for (let i = 1; i < ld.length; i++) { if (ld[i][0] == payload.entryId) { ri = i+1; break; } } if (ri !== -1) { sheetLog.deleteRow(ri); updateTotalStatus(payload.idKpi, sheetLog, sheetStatus, null, null); return { success: true, message: "Rekod dibuang." }; } else return { success: false, message: "Rekod tidak dijumpai." }; } catch (err) { return { success: false, message: "Ralat: " + err.toString() }; } }

function updateTotalStatus(idKpi, sheetLog, sheetStatus, latestFileUrls, quarterData) {
  const lData = sheetLog.getDataRange().getValues(); let total = 0; for (let i = 1; i < lData.length; i++) { if (lData[i][3] === idKpi) total += (parseFloat(lData[i][6]) || 0); }
  const sData = sheetStatus.getDataRange().getValues(); let found = false;
  for (let i = 1; i < sData.length; i++) { if (sData[i][0] === idKpi) { found = true; const sasaran = parseFloat(sData[i][3]) || 1; sheetStatus.getRange(i+1,6).setValue(total); sheetStatus.getRange(i+1,7).setValue(Math.min(100,(total/sasaran)*100)); if (latestFileUrls) sheetStatus.getRange(i+1,8).setValue(latestFileUrls.split(',')[0].trim()); if (quarterData) { sheetStatus.getRange(i+1,9).setValue(quarterData.q1_val||""); sheetStatus.getRange(i+1,10).setValue(quarterData.q1_status||""); sheetStatus.getRange(i+1,11).setValue(quarterData.q2_val||""); sheetStatus.getRange(i+1,12).setValue(quarterData.q2_status||""); sheetStatus.getRange(i+1,13).setValue(quarterData.q3_val||""); sheetStatus.getRange(i+1,14).setValue(quarterData.q3_status||""); sheetStatus.getRange(i+1,15).setValue(quarterData.q4_val||""); sheetStatus.getRange(i+1,16).setValue(quarterData.q4_status||""); sheetStatus.getRange(i+1,17).setValue(quarterData.kpi_tercapai||false); } break; } }
  if (!found) { let nr = [idKpi,"","","","",total,0,latestFileUrls?latestFileUrls.split(',')[0].trim():""]; if (quarterData) nr.push(quarterData.q1_val||"",quarterData.q1_status||"",quarterData.q2_val||"",quarterData.q2_status||"",quarterData.q3_val||"",quarterData.q3_status||"",quarterData.q4_val||"",quarterData.q4_status||"",quarterData.kpi_tercapai||false); else nr.push("","","","","","","","",false); sheetStatus.appendRow(nr); }
}

function setupPelanStrategikUniversiti() {
  const ss = getSS();
  let sheet = ss.getSheetByName(SHEET_NAME_PSU);
  if (!sheet) sheet = ss.insertSheet(SHEET_NAME_PSU);
  const headers = ["ID KPI","Strategi Pelaksanaan","Inisiatif","No Aktiviti","Aktiviti","Indikator Pengukuran","Sasaran KPI 2026 (Universiti)","Bilangan Sasaran KPI 2026 (PTj)","Status","Pencapaian Semasa","Bukti/URL","Catatan"];
  sheet.getRange(1,1,1,headers.length).setValues([headers]).setBackground("#002C5F").setFontColor("white").setFontWeight("bold");
  if (sheet.getLastRow() === 1) {
    const data = [
      ["PSU-SP1-01","SP1: Future Ready Curriculum","Curriculum Innovations (Flexible Academic Programmes)","1","Flexible Course (MOOC)","Number of courses","4 (3 MPU Courses, 1 RMC (PPS))","1","Belum Mula","0","",""],
      ["PSU-SP1-02","SP1: Future Ready Curriculum","Appointment of international academic staff","2","Appointment of international academic staff (External Examiners)","Percentage of international External Examiners (based on confirmed Viva -Voce candidates)","Faculty (except FUPL): 40%","40","Belum Mula","0","",""],
      ["PSU-03","SP1: Future Ready Curriculum","Appointment of international academic staff","2","Appointment of international academic staff (Co-Supervisor)","Number of international academic staff as Co-Supervisor (PG)","200 international academic staff of Co-Supervisor","200","Belum Mula","0","",""],
      ["PSU-SP1-04","SP1: Future Ready Curriculum","Global Networking","3","Enrollment of international students in academic programmes","Number of international student enrollments (60% (900) from other countries than Indonesia)","1500","1500","Belum Mula","0","",""],
      ["PSU-SP1-05","SP1: Future Ready Curriculum","Industry-based Learning","4","Industrial involvement in student assessment","Number of industry co-supervisors or External Examiners","10 industry co-supervisors or External Examiners","1","Belum Mula","0","",""],
      ["PSU-SP2-01","SP2: Sustainable Learning Ecosystem","Establish a Green Learning Common Spaces that leverages AI-driven adaptive environment","5","Number of learning common spaces","2 lounge (KB and KGB), PWB: 1, PPS: 1","1 Green learning common- PWB","0","Belum Mula","0","",""]
    ];
    sheet.getRange(2,1,data.length,headers.length).setValues(data);
  }
  sheet.autoResizeColumns(1, headers.length);
  return "Pelan Strategik Universiti siap.";
}
