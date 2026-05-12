const DATA_SHEET_ID = "1yUKNSoMFKyOa9WSMVt1ucmndXE9uveigVnvKzOT3wck";
const FACILITY_SHEET_ID = "1H35dwM2pQO6pdycz-HcOKXiVlL10GXZX9Vvme4CLzQE";

const SPECIAL_HOSPITALS = [
  "Abebech Gobena MCH Hospital",
  "ALERT Comprehensive Specialized Hospital",
  "Amanuel Mental Specialized Hospital",
  "Dagmawi Minilik Comprehensive Specialized Hospital",
  "Eka Kotebe General Hospital",
  "Gandhi Maternal And Child Health Specialty Center",
  "Ras Desta Damitew General Hospital",
  "St Paulo's Comprehensive Specialized Hospital",
  "St. Peter General Hospital",
  "Tikur Anbessa Comprehensive Specialized Hospital",
  "Tirunesh Beijing General Hospital",
  "Yekatit 12 Medical College General Hospital",
  "Zewditu Memorial General Hospital"
];

function formatW(z, w) {
  if (!z || !w) return w || '';
  let zF = z.toString().trim().replace(/\s+/g, '_');
  let wStr = w.toString().trim();
  let m = wStr.match(/\d+/);
  if (m) {
    let wNum = parseInt(m[0], 10);
    return zF + '_Woreda' + wNum;
  } else {
    let wClean = wStr.replace(/\s+/g, '_');
    if (wClean.toLowerCase().startsWith('woreda')) return zF + '_' + wClean;
    return zF + '_Woreda' + wClean;
  }
}

function doGet(e) {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Addis Ababa PHEM Weekly IDSR Data Entry Form')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0');
}

function doPost(e) {
  if (e && e.parameter && e.parameter.action === 'submit') return handleSubmit(e);
  if (e && e.parameter && e.parameter.action === 'editRecord') return editRecord(e);
  return ContentService.createTextOutput(JSON.stringify({status:'error'})).setMimeType(ContentService.MimeType.JSON);
}

function rpcSubmit(data) {
  try { return JSON.parse(handleSubmit({ parameter: data }).getContent()); } 
  catch(e) { return { status: 'error', message: e.toString() }; }
}

function rpcEdit(data) {
  try { return JSON.parse(editRecord({ parameter: data }).getContent()); } 
  catch(e) { return { status: 'error', message: e.toString() }; }
}

function rpcGetMetrics(woredaId, week, facilityName) {
  try { return JSON.parse(getMetrics(woredaId, week, facilityName).getContent()); } 
  catch(e) { return { status: 'error' }; }
}

function rpcGetSubmitted(week) {
  try { return JSON.parse(getSubmittedFacilities(week).getContent()); } 
  catch(e) { return { status: 'error', submitted: [] }; }
}

function rpcGetRecord(zone, woreda, facility, week, year) {
  try { return JSON.parse(getRecord(zone, woreda, facility, week, year).getContent()); } 
  catch(e) { return { status: 'error' }; }
}

function rpcGetWoredaData(zone, woreda, week, facility) {
  try { return JSON.parse(getWoredaData(zone, woreda, week, facility).getContent()); } 
  catch(e) { return { status: 'error' }; }
}

function rpcManualAutoPopulate(week, year, month) {
  try { 
    return JSON.parse(processAutoPopulate(week, year, month).getContent()); 
  } 
  catch(e) { 
    return { status: 'error', message: e.toString() }; 
  }
}

function appendTotalRow(rows, headers) {
  if (!rows || rows.length === 0) return rows;
  
  const skipHeaders = [
    'Region', 'Zone', 'Woreda', 'Health_Facility', 'Year', 'Month', 'Epi_Week', 
    'Other_1_Name', 'Other_2_Name', 'Other_3_Name', 
    'Initial Submission Time Stamp', 'Edited Submission Time Stamp', 
    'Completeness', 'Timeliness'
  ];
  
  let totalRow = new Array(headers.length).fill('');
  totalRow[0] = 'GRAND TOTAL';
  
  for (let c = 0; c < headers.length; c++) {
    if (skipHeaders.includes(headers[c])) continue;
    
    let sum = 0;
    let hasData = false;
    for (let r = 0; r < rows.length; r++) {
      let val = parseInt(rows[r][c], 10);
      if (!isNaN(val)) {
        sum += val;
        hasData = true;
      }
    }
    if (hasData) {
      totalRow[c] = sum;
    }
  }
  
  const expIdx = headers.indexOf('All_Total_sites_Expected_by_RHB');
  const repIdx = headers.indexOf('All_Total_sites_Reported');
  const timeIdx = headers.indexOf('All_Total_sites_Reported_Ontime');
  const compHeaderIdx = headers.indexOf('Completeness');
  const timeHeaderIdx = headers.indexOf('Timeliness');
  
  if (expIdx > -1 && repIdx > -1 && compHeaderIdx > -1 && totalRow[expIdx] > 0) {
      let comp = Math.round((totalRow[repIdx] / totalRow[expIdx]) * 100);
      totalRow[compHeaderIdx] = Math.min(100, comp); 
  }
  if (expIdx > -1 && timeIdx > -1 && timeHeaderIdx > -1 && totalRow[expIdx] > 0) {
      let time = Math.round((totalRow[timeIdx] / totalRow[expIdx]) * 100);
      totalRow[timeHeaderIdx] = Math.min(100, time);
  }

  rows.push(totalRow);
  return rows;
}

function getExportWeekCSV(week, year, zone, woreda, facility) {
  try {
    const sheet = SpreadsheetApp.openById(DATA_SHEET_ID).getSheetByName("Week " + week);
    if (!sheet) return "";
    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) return "";
    
    const headers = data[0];
    const specialMap = getSpecialHospitalMap();
    let filteredRows = [];
    let formattedSearchWoreda = formatW(zone, woreda);
    let isSpecial = facility && SPECIAL_HOSPITALS.includes(facility);

    for (let i = 1; i < data.length; i++) {
      let fName = data[i][3]?.toString().trim()||'';
      let z = data[i][1]?.toString().trim()||'';
      let w = data[i][2]?.toString().trim()||'';
      
      if (isSpecial) {
         if (fName === facility) {
            filteredRows.push([...data[i]]);
         }
      } else {
         if (specialMap[fName]) {
            z = specialMap[fName].zone;
            w = formatW(z, specialMap[fName].woreda);
         }
         if (z === zone && w === formattedSearchWoreda) {
           let rowCopy = [...data[i]];
           rowCopy[1] = z; 
           rowCopy[2] = w; 
           filteredRows.push(rowCopy);
         }
      }
    }
    
    if (filteredRows.length > 0) {
       appendTotalRow(filteredRows, headers);
       return [headers, ...filteredRows].map(r => r.join(',')).join('\n');
    }
    return "";
  } catch (e) { return ""; }
}

function getExportRangeCSV(startWeek, endWeek, year, zone, woreda, facility) {
  try {
    const ss = SpreadsheetApp.openById(DATA_SHEET_ID);
    let filteredRows = [];
    let headers = [];
    const specialMap = getSpecialHospitalMap();
    let formattedSearchWoreda = formatW(zone, woreda);
    let isSpecial = facility && SPECIAL_HOSPITALS.includes(facility);
    
    for (let w = parseInt(startWeek); w <= parseInt(endWeek); w++) {
      const sheet = ss.getSheetByName("Week " + w);
      if (sheet) {
        const data = sheet.getDataRange().getValues();
        if (headers.length === 0 && data.length > 0) headers = data[0]; 
        
        for (let i = 1; i < data.length; i++) {
          let fName = data[i][3]?.toString().trim()||'';
          let z = data[i][1]?.toString().trim()||'';
          let wor = data[i][2]?.toString().trim()||'';
          
          if (isSpecial) {
             if (fName === facility) {
                filteredRows.push([...data[i]]);
             }
          } else {
             if (specialMap[fName]) {
                 z = specialMap[fName].zone;
                 wor = formatW(z, specialMap[fName].woreda);
             }
             if (z === zone && wor === formattedSearchWoreda) {
               let rowCopy = [...data[i]];
               rowCopy[1] = z; 
               rowCopy[2] = wor; 
               filteredRows.push(rowCopy);
             }
          }
        }  
      }
    }
    
    if (filteredRows.length > 0 && headers.length > 0) {
      appendTotalRow(filteredRows, headers);
      return [headers, ...filteredRows].map(r => r.join(',')).join('\n');
    }
    return "";
  } catch (e) { return ""; }
}

function getStrictHeaders() {
  return [
    'Region', 'Zone', 'Woreda', 'Health_Facility', 'Year', 'Month', 'Epi_Week',
    'Total_Malaria_Confirmed_and_Clinical', 'TMalaria_OutP_Cases', 'TMalaria_InP_Cases', 'TMalaria_InP_Deaths', 'TMSuspected_Fever_Examined',
    'PosMalaria_RDT_or_Microscopy_PF_OutP_Cases', 'PosMalaria_RDT_or_Microscopy_PV_OutP_Cases', 'PosMalaria_RDT_or_Microscopy_Mixed_OutP_Cases',
    'Meningits_total_Cases', 'Meningitis_OutP_Cases', 'Meningitis_InP_Cases', 'Meningitis_InP_Deaths',
    'Dysentery_total_Cases', 'Dysentery_OutP_Cases', 'Dysentery_InP_Cases', 'Dysentery_InP_Deaths',
    'Scabies_total_Cases', 'Scabies_OutP_Cases', 'Scabies_InP_Cases', 'Scabies_InP_Death',
    'Relapsing_fever_total_Cases', 'Relapsing_fever_OutPatient_Cases', 'Relapsing_fever_InP_Cases', 'Relapsing_fever_InP_Deaths',
    'SAM_U5_Total_Cases', 'SAM_U5_OutP_Cases', 'SAM_U5_InP_Cases', 'SAM_InP_Deaths',
    'MAM_U5C_Total_Cases', 'MAM_U5C_OutP_Cases', 'MAM_U5C_InP_Cases',
    'MAM_PLW_Total_Cases', 'MAM_PLW_OutP_Cases', 'MAM_PLW_InP_Cases',
    'Diarrhea_with_dehydration_U5_Total_Cases', 'Diarrhea_with_dehydration_U5_OutP_Cases', 'Diarrhea_with_dehydration_U5_InP_Cases', 'Diarrhea_with_dehydration_U5__InP_Deaths',
    'Acute_jaundice_syndrome_within_14_days_of_illness_Total_Cases', 'Acute_jaundice_syndrome_within_14_days_of_illness_OutP_Cases', 'Acute_jaundice_syndrome_within_14_days_of_illness_InP_Cases', 'Acute_jaundice_syndrome_within_14_days_of_illness_InP_Deaths',
    'Sever_pneumonia_in_children_U5_Total_Cases', 'Sever_pneumonia_in_children_U5_OutP_Cases', 'Sever_pneumonia_in_children_U5_InP_Cases', 'Sever_pneumonia_in_children_U5_InP_Deaths',
    'Diabetic_Mellitus_new_Total_Cases', 'Diabetic_Mellitus_new_OutP_Cases', 'Diabetic_Mellitus_new_InP_cases', 'Diabetic_Mellitus_new_InP_Deaths',
    'HIV_new_Total_Cases', 'HIV_new_OutP_Cases', 'HIV_new_InP_Cases', 'HIV_new_InP_Deaths',
    'Hypertention_new_Total_Cases', 'Hypertention_new_OutP_Cases', 'Hypertention_new_InP_Cases', 'Hypertention_new_InP_Deaths',
    'Tuberculosis_new_Total_Cases', 'Tuberculosis_new_OutP_Cases', 'Tuberculosis_new_InP_Cases', 'Tuberculosis_new_InP_Deaths',
    'AFP_Polio_Cases', 'AFP_Polior_Deaths',
    'Anthrax_Cases', 'Anthrax_Deaths', 'Cholera_Cases', 'Cholera_Deaths',
    'Dracuncunculiasis_Guinea_worm_Cases', 'Dracunculiasis_Guinea_worm_Deaths',
    'Chikungunya_Total_Cases', 'Chikungunya_OutP_Cases', 'Chikungunya_InP_Cases', 'Chikungunya_InP_Deaths',
    'AEFI_Total_Cases', 'AEFI_OutP_Cases', 'AEFI_InP_Cases', 'AEFI_InP_Deaths',
    'Measles_Cases', 'Measles_Deaths', 'Neonatal_Tetanus_Cases', 'Neonatal_Tetanus_Deaths',
    'Human_influenza_caused_by_new_subtype_Cases', 'Human_influenza_caused_by_new_subtype_Deaths',
    'Suspected_rabies_exposure_Cases', 'Suspected_rabies_exposure_Deaths', 'Human_Rabies_Cases', 'Human_Rabies_Deaths',
    'Dengue_fever_Cases', 'Dengue_fever_Deaths', 'SARS_Cases', 'SARS_Deaths', 'Small_pox_Cases', 'Small_pox_Deaths',
    'Viral_hemorrhagic_fever_Cases', 'Viral_hemorrhagic_fever_Deaths', 'Yellow_fever_Cases', 'Yellow_fever_Deaths',
    'COVID_19_Total_Cases', 'COVID_19_OutP_Cases', 'COVID_19_InP_Cases', 'COVID_19_Deaths',
    'Monkeypox_virus_Cases', 'Monkeypox_virus_Deaths', 'Rift_Valley_Fever_Cases', 'Rift_Valley_Fever_Deaths',
    'Brucellosis_Cases', 'Brucellosis_Deaths',
    'Maternal_death', 'Perinatal_death',
    'Obstetric_fistula_Cases', 'Obstetric_fistula_Deaths',
    'Chemical_Poisoning_Cases', 'Chemical_Poisoning_Deaths',
    'Community_Notifications_Total', 'Community_Notifications_Case_Definition', 'Community_Notifications_30min',
    'Dog_Bite',
    'Other_1_Name', 'Other_1_Cases', 'Other_1_Deaths',
    'Other_2_Name', 'Other_2_Cases', 'Other_2_Deaths',
    'Other_3_Name', 'Other_3_Cases', 'Other_3_Deaths',
    'No_of_Gov_HPs_expected_by_RHB', 'No_of_Gov_HCs_expected_by_RHB', 'No_of_Gov_Hosps_expected_by_RHB',
    'No_of_NGOHF_expected_by_RHB', 'No_of_OtherHFs_expected_by_RHB',
    'No_of_Gov_HPs_Reported', 'No_of_Gov_HCs_Reported', 'No_of_Gov_Hosps_Reported',
    'No_of_NGOHFs_Reported', 'No_of_OthersHF_Reported',
    'No_of_Gov_HPs_Reported_Ontime', 'No_of_Gov_HCs_Reported_Ontime', 'No_of_Gov_Hosps_Reported_Ontime',
    'No_of_NGOHFs_Reported_Ontime', 'No_of_OthersHF_Reported_Ontime',
    'All_Total_sites_Reported', 'All_Total_sites_Expected_by_RHB', 'All_Total_sites_Reported_Ontime',
    'Total_Gov_sites_Reported', 'Completeness', 'Timeliness', 
    'Initial Submission Time Stamp', 'Edited Submission Time Stamp'
  ];
}

function getSpecialHospitalMap() {
  const facSheet = SpreadsheetApp.openById(FACILITY_SHEET_ID).getSheetByName('MFR Facility List');
  const facData = facSheet.getDataRange().getValues();
  const cz = facData[0].indexOf('Zone');
  const cw = facData[0].indexOf('Woreda');
  const cf = facData[0].indexOf('Health_Facility');
  
  let map = {};
  for(let i=1; i<facData.length; i++) {
     let fName = facData[i][cf]?.toString().trim();
     if (SPECIAL_HOSPITALS.includes(fName)) {
        map[fName] = { zone: facData[i][cz]?.toString().trim(), woreda: facData[i][cw]?.toString().trim() };
     }
  }
  return map;
}

function getMetrics(woredaId, week, facilityName) {
  try {
    const facSheet = SpreadsheetApp.openById(FACILITY_SHEET_ID).getSheetByName('MFR Facility List');
    const facData = facSheet.getDataRange().getValues();
    const fheaders = facData[0];
    const cid = fheaders.indexOf('Woreda ID');
    const cz = fheaders.indexOf('Zone');
    const cw = fheaders.indexOf('Woreda');
    const cf = fheaders.indexOf('Health_Facility');
    const cs = fheaders.indexOf('Operational Status');
    
    let expected = { totalExpected: 0 };
    let woredaName = "";
    let zoneName = "";
    let expectedFacs = [];

    const specialWoredas = ["1123", "1124", "1125", "1126", "1127", "1128", "1129", "1131", "1132", "1133", "1134", "1135"];
    let isSpecialWoreda = specialWoredas.includes(woredaId.toString().trim());
    
    for (let i = 1; i < facData.length; i++) {
      if (facData[i][cid]?.toString().trim() !== woredaId) continue;
      woredaName = facData[i][cw]?.toString().trim();
      zoneName = facData[i][cz]?.toString().trim();
      let facNameFromSheet = facData[i][cf]?.toString().trim();

      if (facData[i][cs]?.toString().trim().toLowerCase() !== 'operational') continue;

      if (facilityName && facNameFromSheet !== facilityName) continue;

      expected.totalExpected++;
      expectedFacs.push(facNameFromSheet);
    }
    
    if (isSpecialWoreda) {
        expected.totalExpected = 1;
    }

    let formattedWoreda = formatW(zoneName, woredaName);
    
    const dataSheet = SpreadsheetApp.openById(DATA_SHEET_ID);
    const sheet = dataSheet.getSheetByName("Week " + week);
    let reported = { totalReported: 0, totalOnTime: 0, govSitesReported: 0 };
    
    if (sheet) {
      const sdata = sheet.getDataRange().getValues();
      const shdr = sdata[0];
      const compIdx = shdr.indexOf('Completeness');
      const zIdx = shdr.indexOf('Zone');
      const wIdx = shdr.indexOf('Woreda');
      const fIdx = shdr.indexOf('Health_Facility');
      const govIdx = shdr.indexOf('Total_Gov_sites_Reported');
      const otIdx = shdr.indexOf('All_Total_sites_Reported_Ontime');

      let uniqueFacs = new Set();
      let govFacs = new Set();
      let onTimeFacs = new Set();

      for (let i = 1; i < sdata.length; i++) {
        // Skip rows with Completeness === 0 (auto-populated, not yet reported)
        if (compIdx > -1 && parseInt(sdata[i][compIdx]) === 0) continue;

        let fac = sdata[i][fIdx]?.toString().trim();
        let rowZone = sdata[i][zIdx]?.toString().trim();
        let rowWoreda = sdata[i][wIdx]?.toString().trim();
        
        let isMatch = false;
        if (SPECIAL_HOSPITALS.includes(fac)) {
          if (expectedFacs.includes(fac)) isMatch = true;
        } else {
          if (rowZone === zoneName && rowWoreda === formattedWoreda) isMatch = true;
        }
        
        if (facilityName && fac !== facilityName) isMatch = false;

        if (isMatch) {
          uniqueFacs.add(fac);
          if (parseInt(sdata[i][govIdx]) === 1) govFacs.add(fac);
          if (parseInt(sdata[i][otIdx]) === 1) onTimeFacs.add(fac);
        }
      }
      
      reported.totalReported = uniqueFacs.size;
      reported.govSitesReported = govFacs.size;
      reported.totalOnTime = onTimeFacs.size;
    }
    
    let completeness = expected.totalExpected > 0 ? Math.round((reported.totalReported / expected.totalExpected) * 100) : 0;
    let timeliness = expected.totalExpected > 0 ? Math.round((reported.totalOnTime / expected.totalExpected) * 100) : 0;
    
    completeness = Math.min(100, completeness);
    timeliness = Math.min(100, timeliness);

    return ContentService.createTextOutput(JSON.stringify({
      status:'success', expected, reported,
      completeness: completeness,
      timeliness: timeliness
    })).setMimeType(ContentService.MimeType.JSON);
  } catch (e) {
    return ContentService.createTextOutput(JSON.stringify({status:'error'})).setMimeType(ContentService.MimeType.JSON);
  }
}

// Compare function for sorting rows
function compareRows(a, b) {
  let facA = (a[3] || '').toString().trim();
  let facB = (b[3] || '').toString().trim();
  
  // Check if facilities are special hospitals
  let isSpecA = SPECIAL_HOSPITALS.includes(facA);
  let isSpecB = SPECIAL_HOSPITALS.includes(facB);
  
  // If both are special or both are non-special, sort by zone/woreda
  if (isSpecA === isSpecB) {
    let zoneA = (a[1] || '').toString().trim().toLowerCase();
    let zoneB = (b[1] || '').toString().trim().toLowerCase();
    if (zoneA < zoneB) return -1;
    if (zoneA > zoneB) return 1;
    
    let worA = (a[2] || '').toString().trim().toLowerCase();
    let worB = (b[2] || '').toString().trim().toLowerCase();
    
    let numA = parseInt((worA.match(/\d+/) || ['0'])[0], 10);
    let numB = parseInt((worB.match(/\d+/) || ['0'])[0], 10);
    let baseA = worA.replace(/\d+/g, '').trim();
    let baseB = worB.replace(/\d+/g, '').trim();
    
    if (baseA < baseB) return -1;
    if (baseA > baseB) return 1;
    if (numA !== numB) return numA - numB;
    
    // If both are special hospitals, sort alphabetically
    if (isSpecA) {
      if (facA < facB) return -1;
      if (facA > facB) return 1;
    }
    
    return 0;
  }
  
  // Non-special facilities come first, special hospitals go to the end
  return isSpecA ? 1 : -1;
}

function sortSheet(sheet) {
  if (sheet.getLastRow() <= 1) return;
  
  const range = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn());
  const values = range.getValues();
  
  // Check if sorting is needed
  let needsSort = false;
  for (let i = 0; i < values.length - 1; i++) {
    if (compareRows(values[i], values[i + 1]) > 0) {
      needsSort = true;
      break;
    }
  }
  
  if (needsSort) {
    values.sort(compareRows);
    range.setValues(values);
    
    // Re-apply formatting after sort
    const backgrounds = [];
    const fontStyles = [];
    const initTimeIdx = getStrictHeaders().indexOf('Initial Submission Time Stamp');
    
    for (let i = 0; i < values.length; i++) {
      let isAuto = values[i][initTimeIdx] === "AUTO-POPULATED";
      let bg = isAuto ? '#fee2e2' : null;
      let fs = isAuto ? 'italic' : 'normal';
      
      backgrounds.push(new Array(values[i].length).fill(bg));
      fontStyles.push(new Array(values[i].length).fill(fs));
    }
    
    range.setBackgrounds(backgrounds);
    range.setFontStyles(fontStyles);
  }
}

function handleSubmit(e) {
  try {
    const data = e.parameter;
    const week = data.week || '1', facility = data.facility || '', woredaId = data.woredaId || '', category = data.category || 'other';
    if (!facility || !woredaId) return ContentService.createTextOutput(JSON.stringify({status:'error',message:'Missing fields'})).setMimeType(ContentService.MimeType.JSON);
    
    const now = new Date();
    const day = now.getDay();
    const hour = now.getHours();
    const minutes = now.getMinutes();
    
    const isOnTime = (day === 1) || (day === 2 && (hour < 11 || (hour === 11 && minutes <= 59)));
    
    const ss = SpreadsheetApp.openById(DATA_SHEET_ID);
    const sheetName = "Week " + week;
    let sheet = ss.getSheetByName(sheetName);
    
    let isNewSheet = false;
    if (!sheet) {
      sheet = ss.insertSheet(sheetName);
      isNewSheet = true;
    } else if (sheet.getLastRow() === 0) {
      isNewSheet = true;
    }

    let formattedWoreda = formatW(data.zone, data.woreda);
    let searchZone = data.zone;
    let searchWoreda = formattedWoreda;
    if (SPECIAL_HOSPITALS.includes(facility)) {
      searchZone = facility;
      searchWoreda = facility;
    }
    
    if (!isNewSheet) {
      // Only read first 7 columns for duplicate check (fast)
      const lastRow = sheet.getLastRow();
      const existingData = sheet.getRange(1, 1, lastRow, 7).getValues();
      
      for (let i = 1; i < existingData.length; i++) {
        if ((existingData[i][1] || '').toString().trim() === searchZone &&
            (existingData[i][2] || '').toString().trim() === searchWoreda &&
            (existingData[i][3] || '').toString().trim() === facility &&
            (existingData[i][4] || '').toString().trim() == data.year &&
            (existingData[i][6] || '').toString().trim() == data.week) {
          
          // Found matching row - read full row to check if auto-populated
          const fullRow = sheet.getRange(i + 1, 1, 1, sheet.getLastColumn()).getValues()[0];
          const initTimeIdx = getStrictHeaders().indexOf('Initial Submission Time Stamp');
          let isAuto = fullRow[initTimeIdx] === "AUTO-POPULATED";
          
          if (isAuto) {
            let diseaseData = {};
            try { diseaseData = JSON.parse(data.jsonData || '{}'); } catch(e) {}
            const initialTime = Utilities.formatDate(new Date(), "GMT+3", "yyyy-MM-dd HH:mm:ss");
            const editedTime = "OVERWRITTEN AUTO";
            let rowData = Object.assign({}, data);
            if (SPECIAL_HOSPITALS.includes(facility)) {
              rowData.zone = facility;
              rowData.woreda = facility;
            } else {
              rowData.woreda = formattedWoreda;
            }
            const row = buildRowData(rowData, diseaseData, isOnTime, category, initialTime, editedTime);

            // Write the row (overwriting the auto-populated row)
            sheet.getRange(i + 1, 1, 1, row.length).setValues([row]);
            
            // Clear the red background since it's now submitted
            sheet.getRange(i + 1, 1, 1, row.length).setBackground(null).setFontStyle('normal');
            
            return ContentService.createTextOutput(JSON.stringify({status:'success',message:'Submitted for Week '+week})).setMimeType(ContentService.MimeType.JSON);
          }
          
          return ContentService.createTextOutput(JSON.stringify({status:'duplicate',message:'Already submitted for Week '+week})).setMimeType(ContentService.MimeType.JSON);
        }
      }
    }
    
    if (isNewSheet) {
      const headers = getStrictHeaders();
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#4f46e5').setFontColor('#ffffff').setFontSize(9);
      sheet.setFrozenRows(1);
      sheet.setFrozenColumns(4);
    }
    
    let diseaseData = {};
    try { diseaseData = JSON.parse(data.jsonData || '{}'); } catch(e) {}
    
    const initialTime = Utilities.formatDate(new Date(), "GMT+3", "yyyy-MM-dd HH:mm:ss");
    const editedTime = "";
    
    let rowData = Object.assign({}, data);
    if (SPECIAL_HOSPITALS.includes(facility)) {
      rowData.zone = facility;
      rowData.woreda = facility;
    } else {
      rowData.woreda = formattedWoreda;
    }

    const row = buildRowData(rowData, diseaseData, isOnTime, category, initialTime, editedTime);
    
    // Insert at correct sorted position
    if (sheet.getLastRow() > 0) {
      const allData = sheet.getDataRange().getValues();
      let insertPosition = allData.length + 1; // Default: append at end (1-indexed)
      
      // Find where this row should be inserted to maintain sort order
      for (let i = 1; i < allData.length; i++) {
        if (compareRows(row, allData[i]) < 0) {
          insertPosition = i + 1; // +1 because sheet rows are 1-indexed
          break;
        }
      }
      
      if (insertPosition <= allData.length) {
        // Insert at the correct position
        sheet.insertRowBefore(insertPosition);
        sheet.getRange(insertPosition, 1, 1, row.length).setValues([row]);
      } else {
        // Append at the end
        sheet.appendRow(row);
      }
    } else {
      sheet.appendRow(row);
    }
    
    return ContentService.createTextOutput(JSON.stringify({status:'success',message:'Submitted for Week '+week})).setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({status:'error',message:err.toString()})).setMimeType(ContentService.MimeType.JSON);
  }
}

function buildRowData(data, diseaseData, isOnTime, category, initialTime, editedTime) {
  const d = diseaseData;
  
  let cE = [0,0,0,0,0], cR = [0,0,0,0,0];
  let cIdx = category === 'gov_hp' ? 0 : category === 'gov_hc' ? 1 : category === 'gov_hosp' ? 2 : category === 'ngo' ? 3 : 4;
  cE[cIdx] = 1; 

  let isAutoPopulated = (initialTime === "AUTO-POPULATED");
  let reportedVal = isAutoPopulated ? 0 : 1;
  cR[cIdx] = reportedVal;
  
  let isGov = (category === 'gov_hc' || category === 'gov_hosp') ? 1 : 0;
  let ot = (isOnTime && reportedVal === 1) ? 1 : 0;
  let completeness = reportedVal * 100;
  let timeliness = ot ? 100 : 0;

  const row = [
    'Addis Ababa', 
    data.zone || '', data.woreda || '', data.facility || '',
    data.year || new Date().getFullYear(), data.month || '', data.week || '',
    getVal(d,'malaria_total'), getVal(d,'malaria_outp'), getVal(d,'malaria_inp'), getVal(d,'malaria_deaths'), getVal(d,'malaria_suspected'),
    getVal(d,'malaria_pf'), getVal(d,'malaria_pv'), getVal(d,'malaria_mixed'),
    getVal(d,'meningitis_total'), getVal(d,'meningitis_outp'), getVal(d,'meningitis_inp'), getVal(d,'meningitis_deaths'),
    getVal(d,'dysentery_total'), getVal(d,'dysentery_outp'), getVal(d,'dysentery_inp'), getVal(d,'dysentery_deaths'),
    getVal(d,'scabies_total'), getVal(d,'scabies_outp'), getVal(d,'scabies_inp'), getVal(d,'scabies_deaths'),
    getVal(d,'relapsing_total'), getVal(d,'relapsing_outp'), getVal(d,'relapsing_inp'), getVal(d,'relapsing_deaths'),
    getVal(d,'sam_u5_total'), getVal(d,'sam_u5_outp'), getVal(d,'sam_u5_inp'), getVal(d,'sam_u5_deaths'),
    getVal(d,'mam_u5c_total'), getVal(d,'mam_u5c_outp'), getVal(d,'mam_u5c_inp'),
    getVal(d,'mam_plw_total'), getVal(d,'mam_plw_outp'), getVal(d,'mam_plw_inp'),
    getVal(d,'diarrhea_total'), getVal(d,'diarrhea_outp'), getVal(d,'diarrhea_inp'), getVal(d,'diarrhea_deaths'),
    getVal(d,'jaundice_total'), getVal(d,'jaundice_outp'), getVal(d,'jaundice_inp'), getVal(d,'jaundice_deaths'),
    getVal(d,'pneumonia_total'), getVal(d,'pneumonia_outp'), getVal(d,'pneumonia_inp'), getVal(d,'pneumonia_deaths'),
    getVal(d,'diabetes_total'), getVal(d,'diabetes_outp'), getVal(d,'diabetes_inp'), getVal(d,'diabetes_deaths'),
    getVal(d,'hiv_total'), getVal(d,'hiv_outp'), getVal(d,'hiv_inp'), getVal(d,'hiv_deaths'),
    getVal(d,'hypertension_total'), getVal(d,'hypertension_outp'), getVal(d,'hypertension_inp'), getVal(d,'hypertension_deaths'),
    getVal(d,'tb_total'), getVal(d,'tb_outp'), getVal(d,'tb_inp'), getVal(d,'tb_deaths'),
    getVal(d,'afp_cases'), getVal(d,'afp_deaths'),
    getVal(d,'anthrax_cases'), getVal(d,'anthrax_deaths'), getVal(d,'cholera_cases'), getVal(d,'cholera_deaths'),
    getVal(d,'guinea_cases'), getVal(d,'guinea_deaths'),
    getVal(d,'chikungunya_total'), getVal(d,'chikungunya_outp'), getVal(d,'chikungunya_inp'), getVal(d,'chikungunya_deaths'),
    getVal(d,'aefi_total'), getVal(d,'aefi_outp'), getVal(d,'aefi_inp'), getVal(d,'aefi_deaths'),
    getVal(d,'measles_cases'), getVal(d,'measles_deaths'), getVal(d,'neo_tetanus_cases'), getVal(d,'neo_tetanus_deaths'),
    getVal(d,'flu_new_cases'), getVal(d,'flu_new_deaths'),
    getVal(d,'rabies_exp_cases'), getVal(d,'rabies_exp_deaths'), getVal(d,'human_rabies_cases'), getVal(d,'human_rabies_deaths'),
    getVal(d,'dengue_cases'), getVal(d,'dengue_deaths'), getVal(d,'sars_cases'), getVal(d,'sars_deaths'), getVal(d,'smallpox_cases'), getVal(d,'smallpox_deaths'),
    getVal(d,'vhf_cases'), getVal(d,'vhf_deaths'), getVal(d,'yellow_fever_cases'), getVal(d,'yellow_fever_deaths'),
    getVal(d,'covid19_total'), getVal(d,'covid19_outp'), getVal(d,'covid19_inp'), getVal(d,'covid19_deaths'),
    getVal(d,'monkeypox_cases'), getVal(d,'monkeypox_deaths'), getVal(d,'rift_valley_cases'), getVal(d,'rift_valley_deaths'),
    getVal(d,'brucellosis_cases'), getVal(d,'brucellosis_deaths'),
    getVal(d,'maternal_death'), getVal(d,'perinatal_death'),
    getVal(d,'fistula_cases'), getVal(d,'fistula_deaths'),
    getVal(d,'chemical_cases'), getVal(d,'chemical_deaths'),
    getVal(d,'community_total'), getVal(d,'community_caseDef'), getVal(d,'community_30min'),
    getVal(d,'dog_bite'),
    d.other1_name || '', getVal(d,'other1_cases'), getVal(d,'other1_deaths'),
    d.other2_name || '', getVal(d,'other2_cases'), getVal(d,'other2_deaths'),
    d.other3_name || '', getVal(d,'other3_cases'), getVal(d,'other3_deaths'),
    cE[0], cE[1], cE[2], cE[3], cE[4],
    cR[0], cR[1], cR[2], cR[3], cR[4],
    cR[0]*ot, cR[1]*ot, cR[2]*ot, cR[3]*ot, cR[4]*ot,
    reportedVal, 1, ot,
    isGov * reportedVal, completeness, timeliness,
    initialTime,  
    editedTime 
  ];
  return row;
}

function getVal(obj, key) { return parseInt(obj[key]) || 0; }

function getRecord(zone, woreda, facility, week, year) {
  try {
    const sheet = SpreadsheetApp.openById(DATA_SHEET_ID).getSheetByName("Week " + week);
    if (!sheet) return ContentService.createTextOutput(JSON.stringify({status:'error',message:'No record found'})).setMimeType(ContentService.MimeType.JSON);
    
    let searchWoreda = formatW(zone, woreda);
    if (SPECIAL_HOSPITALS.includes(facility)) {
      zone = facility;
      searchWoreda = facility;
    }

    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    for (let i = 1; i < data.length; i++) {
      if (data[i][1]?.toString().trim() === zone && data[i][2]?.toString().trim() === searchWoreda && data[i][3]?.toString().trim() === facility && data[i][4]?.toString().trim() == year && data[i][6]?.toString().trim() == week) {
        const record = {};
        headers.forEach((h, j) => { record[h] = data[i][j]; });
        return ContentService.createTextOutput(JSON.stringify({status:'success',record})).setMimeType(ContentService.MimeType.JSON);
      }
    }
    return ContentService.createTextOutput(JSON.stringify({status:'error',message:'Record not found'})).setMimeType(ContentService.MimeType.JSON);
  } catch (e) {
    return ContentService.createTextOutput(JSON.stringify({status:'error'})).setMimeType(ContentService.MimeType.JSON);
  }
}

function getWoredaData(zone, woreda, week, facility) {
  try {
    const sheet = SpreadsheetApp.openById(DATA_SHEET_ID).getSheetByName("Week " + week);
    if (!sheet) return ContentService.createTextOutput(JSON.stringify({status:'success', headers:[], rows:[]})).setMimeType(ContentService.MimeType.JSON);
    
    const data = sheet.getDataRange().getValues();
    const headers = data[0] || [];
    const rows = [];
    const specialMap = getSpecialHospitalMap();
    let formattedSearchWoreda = formatW(zone, woreda);
    let isSpecial = facility && SPECIAL_HOSPITALS.includes(facility);

    for (let i = 1; i < data.length; i++) {
      let fName = data[i][3]?.toString().trim()||'';
      let z = data[i][1]?.toString().trim()||'';
      let w = data[i][2]?.toString().trim()||'';
      
      if (isSpecial) {
         if (fName === facility) rows.push([...data[i]]);
      } else {
         if (specialMap[fName]) {
            z = specialMap[fName].zone;
            w = formatW(z, specialMap[fName].woreda);
         }
         if (z === zone && w === formattedSearchWoreda) {
           let rowCopy = [...data[i]];
           rowCopy[1] = z; 
           rowCopy[2] = w; 
           rows.push(rowCopy);
         }
      }
    }
    
    if (rows.length > 0) {
       appendTotalRow(rows, headers);
    }
    
    return ContentService.createTextOutput(JSON.stringify({status:'success', headers:headers, rows:rows})).setMimeType(ContentService.MimeType.JSON);
  } catch (e) {
    return ContentService.createTextOutput(JSON.stringify({status:'error'})).setMimeType(ContentService.MimeType.JSON);
  }
}

function editRecord(e) {
  try {
    const data = e.parameter;
    const week = data.week, facility = data.facility;
    const sheet = SpreadsheetApp.openById(DATA_SHEET_ID).getSheetByName("Week " + week);
    if (!sheet) return ContentService.createTextOutput(JSON.stringify({status:'error'})).setMimeType(ContentService.MimeType.JSON);
    
    const sdata = sheet.getDataRange().getValues();
    const hdr = sdata[0];
    const otIdx = hdr.indexOf('All_Total_sites_Reported_Ontime');
    const initTimeIdx = hdr.indexOf('Initial Submission Time Stamp'); 

    let formattedWoreda = formatW(data.zone, data.woreda);
    let searchZone = data.zone;
    let searchWoreda = formattedWoreda;
    if (SPECIAL_HOSPITALS.includes(facility)) {
      searchZone = facility;
      searchWoreda = facility;
    }

    for (let i = 1; i < sdata.length; i++) {
      if ((sdata[i][1] || '').toString().trim() === searchZone &&
          (sdata[i][2] || '').toString().trim() === searchWoreda &&
          (sdata[i][3] || '').toString().trim() === facility &&
          (sdata[i][4] || '').toString().trim() == data.year &&
          (sdata[i][6] || '').toString().trim() == data.week) {
        
        const wasOnTime = parseInt(sdata[i][otIdx]) === 1;
        const diseaseData = JSON.parse(data.jsonData || '{}');
        
        const existingInitialTime = (initTimeIdx !== -1 && sdata[i][initTimeIdx] !== "AUTO-POPULATED") ? sdata[i][initTimeIdx] : Utilities.formatDate(new Date(), "GMT+3", "yyyy-MM-dd HH:mm:ss");
        const editedTime = Utilities.formatDate(new Date(), "GMT+3", "yyyy-MM-dd HH:mm:ss");
        
        let rowData = Object.assign({}, data);
        if (SPECIAL_HOSPITALS.includes(facility)) {
          rowData.zone = facility;
          rowData.woreda = facility;
        } else {
          rowData.woreda = formattedWoreda;
        }

        const row = buildRowData(rowData, diseaseData, wasOnTime, data.category || 'other', existingInitialTime, editedTime);
        sheet.getRange(i + 1, 1, 1, row.length).setValues([row]);

        return ContentService.createTextOutput(JSON.stringify({status:'success',message:'Updated'})).setMimeType(ContentService.MimeType.JSON);
      }
    }
    return ContentService.createTextOutput(JSON.stringify({status:'error'})).setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({status:'error',message:err.toString()})).setMimeType(ContentService.MimeType.JSON);
  }
}

function getSubmittedFacilities(week) {
  try {
    const sheet = SpreadsheetApp.openById(DATA_SHEET_ID).getSheetByName("Week " + week);
    const submitted = [];
    if (sheet) {
      const data = sheet.getDataRange().getValues();
      const specialMap = getSpecialHospitalMap();
      const compIdx = data[0].indexOf('Completeness');

      for (let i = 1; i < data.length; i++) {
        if (compIdx > -1) {
           let compVal = parseInt(data[i][compIdx]) || 0;
           if (compVal === 0) continue; 
        }

        let fName = data[i][3]?.toString().trim()||'';
        let z = data[i][1]?.toString().trim()||'';
        let w = data[i][2]?.toString().trim()||'';
        
        if (specialMap[fName]) {
           z = specialMap[fName].zone;
           w = formatW(z, specialMap[fName].woreda);
        }
        
        submitted.push({
          zone: z,
          woreda: w,
          facility: fName,
          year: data[i][4]?.toString().trim()||'',
          week: data[i][6]?.toString().trim()||''
        });
      }
    }
    return ContentService.createTextOutput(JSON.stringify({status:'success',submitted})).setMimeType(ContentService.MimeType.JSON);
  } catch (e) {
    return ContentService.createTextOutput(JSON.stringify({status:'error',submitted:[]})).setMimeType(ContentService.MimeType.JSON);
  }
}

function processAutoPopulate(week, year, month) {
  try {
    const ss = SpreadsheetApp.openById(DATA_SHEET_ID);
    const sheetName = "Week " + week;
    let sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
       sheet = ss.insertSheet(sheetName);
       const headers = getStrictHeaders();
       sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
       sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#4f46e5').setFontColor('#ffffff').setFontSize(9);
       sheet.setFrozenRows(1);
       sheet.setFrozenColumns(4);
    }

    const existing = sheet.getDataRange().getValues();
    const headers = getStrictHeaders();
    
    const existingSet = new Set();
    const submittedSet = new Set();
    const compIdx = existing[0]?.indexOf('Completeness');

    if (existing.length > 1) {
        for(let i=1; i<existing.length; i++) {
            let z = (existing[i][1] || '').toString().trim();
            let w = (existing[i][2] || '').toString().trim();
            let f = (existing[i][3] || '').toString().trim();
            if(!f) continue;
            let key = z + '||' + w + '||' + f;
            existingSet.add(key);
            
            let compVal = 1; 
            if (compIdx > -1) compVal = parseInt(existing[i][compIdx]) || 0;
            if (compVal > 0) submittedSet.add(key);
        }
    }

    const facSheet = SpreadsheetApp.openById(FACILITY_SHEET_ID).getSheetByName('MFR Facility List');
    const facData = facSheet.getDataRange().getValues();
    const fheaders = facData[0];
    const cz = fheaders.indexOf('Zone');
    const cw = fheaders.indexOf('Woreda');
    const cf = fheaders.indexOf('Health_Facility');
    const ct = fheaders.indexOf('Facility Type');
    const co = fheaders.indexOf('Ownership');
    const cs = fheaders.indexOf('Operational Status');

    let rowsToAppend = [];

    for (let i = 1; i < facData.length; i++) {
        if ((facData[i][cs] || '').toString().trim().toLowerCase() !== 'operational') continue;
        let fName = (facData[i][cf] || '').toString().trim();
        if (!fName) continue;
        
        let zone = (facData[i][cz] || '').toString().trim();
        let woreda = (facData[i][cw] || '').toString().trim();
        let type = (facData[i][ct] || '').toString().trim();
        let owner = (facData[i][co] || '').toString().trim();
        
        let isSpecial = SPECIAL_HOSPITALS.includes(fName);
        let formattedW = formatW(zone, woreda);
        let finalZone = isSpecial ? fName : zone;
        let finalWoreda = isSpecial ? fName : formattedW;

        let key = finalZone + '||' + finalWoreda + '||' + fName;
        if (submittedSet.has(key)) continue;
        if (existingSet.has(key)) continue;

        let typeClean = (type || '').trim();
        let ownerClean = (owner || '').trim().toLowerCase();
        
        let exp_hp = 0, exp_hc = 0, exp_hosp = 0, exp_ngo = 0, exp_other = 0;
        
        if (typeClean === "Health Center" && ownerClean === "public/government") {
            exp_hc = 1;
        } else if (typeClean === "Hospital" && ownerClean === "public/government") {
            exp_hosp = 1;
        } else if (typeClean !== "Hospital" && typeClean !== "Health Center" && ownerClean === "private not for profit") {
            exp_ngo = 1;
        } else if (typeClean !== "Hospital" && typeClean !== "Health Center" && ownerClean !== "private not for profit" && ownerClean === "private for profit") {
            exp_other = 1;
        }

        let exp_total = exp_hp + exp_hc + exp_hosp + exp_ngo + exp_other;

        let row = new Array(headers.length).fill(0);
        
        row[headers.indexOf('Region')] = 'Addis Ababa';
        row[headers.indexOf('Zone')] = finalZone;
        row[headers.indexOf('Woreda')] = finalWoreda;
        row[headers.indexOf('Health_Facility')] = fName;
        row[headers.indexOf('Year')] = year;
        row[headers.indexOf('Month')] = month;
        row[headers.indexOf('Epi_Week')] = week;

        row[headers.indexOf('Other_1_Name')] = '';
        row[headers.indexOf('Other_2_Name')] = '';
        row[headers.indexOf('Other_3_Name')] = '';

        row[headers.indexOf('No_of_Gov_HPs_expected_by_RHB')] = exp_hp;
        row[headers.indexOf('No_of_Gov_HCs_expected_by_RHB')] = exp_hc;
        row[headers.indexOf('No_of_Gov_Hosps_expected_by_RHB')] = exp_hosp;
        row[headers.indexOf('No_of_NGOHF_expected_by_RHB')] = exp_ngo;
        row[headers.indexOf('No_of_OtherHFs_expected_by_RHB')] = exp_other;

        row[headers.indexOf('All_Total_sites_Expected_by_RHB')] = exp_total;

        row[headers.indexOf('Initial Submission Time Stamp')] = 'AUTO-POPULATED';
        row[headers.indexOf('Edited Submission Time Stamp')] = '';

        rowsToAppend.push(row);
    }

    if (rowsToAppend.length > 0) {
        let startRow = sheet.getLastRow() + 1;
        sheet.getRange(startRow, 1, rowsToAppend.length, rowsToAppend[0].length).setValues(rowsToAppend);
        
        // Apply red background + italic to auto-populated rows
        sheet.getRange(startRow, 1, rowsToAppend.length, rowsToAppend[0].length)
             .setBackground('#fee2e2')
             .setFontStyle('italic');
        
        // Sort entire sheet to maintain correct order
        sortSheet(sheet);
    }

    return ContentService.createTextOutput(JSON.stringify({ status: 'success', added: rowsToAppend.length })).setMimeType(ContentService.MimeType.JSON);
  } catch (e) {
    return ContentService.createTextOutput(JSON.stringify({ status: 'error', message: e.toString() })).setMimeType(ContentService.MimeType.JSON);
  }
}