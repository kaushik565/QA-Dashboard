// Core backend for Rejection Entry Form (multi-item + batch summary)
// Adjust sheet name or headers below as needed.

const RAW_DATA_SHEET = 'Raw Data';
const BRC_SHEET = 'BRC Status';
const BATCH_DETAILS_SHEET = 'Batch Details';
const BRC_TRACKER_SHEET = 'BRC Tracker';
const HEADERS = [
  'Sr No',
  'Date (DD/MM/YYYY)',
  'Shift',
  'Batch No',
  'Lot No',
  'Line',
  'Rejection stage',
  'Equipment ID',
  'Type of Rejections',
  'Cartridge Part',
  'Qty',
  'Verified By'
];
const BRC_HEADERS = ['Batch No','Status','Remarks','Updated On'];
const BATCH_DETAILS_HEADERS = [
  'Batch No',
  'Incident number',
  'Deviation number',
  'CA number',
  'PA number',
  'OOS number',
  'Updated On'
];
const BRC_TRACKER_HEADERS = [
  'Batch No',
  'BRC Submitted',
  'GDP Completed',
  'Updated On'
];

function doGet(e){
  ensureRawDataSheet();
  const view = e && e.parameter && e.parameter.view ? String(e.parameter.view).toLowerCase() : '';
  const file = view === 'batch' ? 'batch' : 'src';
  return HtmlService.createHtmlOutputFromFile(file)
    .setTitle(view==='batch' ? 'Batch Summary' : 'QA DASHBOARD')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// Ensures the Raw Data sheet exists with headers (non-destructive if already present)
function ensureRawDataSheet(){
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(RAW_DATA_SHEET);
  if(!sh){
    sh = ss.insertSheet(RAW_DATA_SHEET);
    sh.appendRow(HEADERS);
    sh.setFrozenRows(1);
    return;
  }
  if(sh.getLastRow() === 0){
    sh.appendRow(HEADERS);
    sh.setFrozenRows(1);
    return;
  }
  const first = sh.getRange(1,1,1,HEADERS.length).getValues()[0];
  const isBlank = first.every(val => String(val).trim() === '');
  if(isBlank){
    sh.getRange(1,1,1,HEADERS.length).setValues([HEADERS]);
    sh.setFrozenRows(1);
  }
}

function formatDateDDMMYYYY(iso){
  if(!iso) return '';
  const d = new Date(iso);
  if(isNaN(d)) return iso;
  const dd = ('0'+d.getDate()).slice(-2);
  const mm = ('0'+(d.getMonth()+1)).slice(-2);
  const yyyy = d.getFullYear();
  return dd+'/'+mm+'/'+yyyy;
}

// Accepts object with fields and RejectionDetails JSON array
function submitEntry(formObj){
  ensureRawDataSheet();
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(RAW_DATA_SHEET);
  const lock = LockService.getDocumentLock();
  lock.waitLock(20000);

  try{
    const itemsJson = formObj.RejectionDetails || '[]';
    let items;
    try{ items = JSON.parse(itemsJson); } catch(err){ items = []; }
    if(!items.length){ return {status:'empty', message:'No items to save'}; }

    const dateStr = formatDateDDMMYYYY(formObj.Date);
    const shift = formObj.Shift || '';
    const batch = formObj.BatchNumber || '';
    const lot = formObj.LOT_NO || '';
    const line = formObj.Line || '';
    const verified = formObj.VerifiedByName || '';

    const rows = items.map(it => [
      '',
      dateStr,
      shift,
      batch,
      lot,
      line,
      categoryLabel(it.category),
      it.equipment || 'NA',
      it.type || '',
      it.cartridgePart || '',
      Number(it.qty)||0,
      verified
    ]);

    const startRow = sh.getLastRow()+1;
    sh.getRange(startRow,1,rows.length,HEADERS.length).setValues(rows);
    rebuildSrNo(sh);
    return {status:'success', message: rows.length + ' rows added'};
  } catch(err){
    return {status:'error', message: err.message};
  } finally {
    lock.releaseLock();
  }
}

function categoryLabel(cat){
  const map = {VI1_Types:'VI-1',VI2_Types:'VI-2',VI3_Types:'VI-3',Vacuum_Types:'VACCUM REJECTIONS',VI4_Types:'VI-4'};
  return map[cat] || cat || '';
}

// Rebuild Sr No column (first column) with 1..n excluding header
function rebuildSrNo(sh){
  const last = sh.getLastRow();
  if(last < 2) return;
  const nums = Array.from({length:last-1}, (_,i)=>[i+1]);
  sh.getRange(2,1,last-1,1).setValues(nums);
}

// Aggregate totals per stage for a given Batch No
function getBatchSummary(batch){
  batch = (batch||'').trim();
  if(!batch) return {status:'error', message:'Batch required'};
  ensureRawDataSheet();
  const sh = SpreadsheetApp.getActive().getSheetByName(RAW_DATA_SHEET);
  const data = sh.getDataRange().getValues();
  if(data.length < 2) return {status:'error', message:'No data'};

  const stageTotals = {};
  let overall = 0;
  for(let i=1;i<data.length;i++){
    const row = data[i];
    const rowBatch = String(row[3]||'').trim();
    if(rowBatch !== batch) continue;
    const stage = String(row[6]||'').trim() || 'UNSPECIFIED';
    const qty = Number(row[10]||0);
    overall += qty;
    stageTotals[stage] = (stageTotals[stage] || 0) + qty;
  }
  if(overall === 0){
    return {status:'error', message:'Batch not found'};
  }
  const standard = ['VI-1','VI-2','VI-3','VACCUM REJECTIONS','VI-4'];
  standard.forEach(stage => {
    if(!stageTotals.hasOwnProperty(stage)) stageTotals[stage] = 0;
  });
  return {status:'success', batch:batch, totals:stageTotals, overall:overall};
}

function getLookupData(){
  try{
    const ss = SpreadsheetApp.getActive();
    const sheet = ss.getSheetByName('Lookups');
    if(!sheet){
      return defaultLookups();
    }
    const headers = sheet.getRange(1,1,1,sheet.getLastColumn()).getValues()[0];
    const lastRow = sheet.getLastRow();
    const rows = lastRow > 1 ? sheet.getRange(2,1,lastRow-1,sheet.getLastColumn()).getValues() : [];
    const obj = {};
    headers.forEach((header, idx) => {
      if(!header) return;
      const values = rows.map(r => String(r[idx]||'').trim()).filter(v => v);
      if(values.length) obj[header] = values;
    });
    const defaults = defaultLookups();
    Object.keys(defaults).forEach(key => {
      if(!obj[key] || !obj[key].length){
        obj[key] = defaults[key];
      }
    });
    return obj;
  } catch(err){
    return defaultLookups();
  }
}

/************************************************************
  safeSetup() - ONLY creates sheets & headers if missing.
  Does NOT delete existing data.
************************************************************/
function safeSetup(){
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // Raw Data sheet
  let raw = ss.getSheetByName("Raw Data");
  if (!raw){
    raw = ss.insertSheet("Raw Data");
  }
  ensureHeaders(raw, HEADERS);

  // Lookups sheet
  let look = ss.getSheetByName("Lookups");
  if (!look){
    look = ss.insertSheet("Lookups");
    look.appendRow([
      "Shift","LOT_NO","Line",
      "VI1_Types","VI2_Types","VI3_Types","Vacuum_Types","VI4_Types",
      "Equipment_LINE_A","Equipment_LINE_B","Equipment_LINE_C","Equipment_LINE_D","Equipment_LINE_E","Equipment_AUTOMATION_LINE",
      "CartridgePart","VerifiedByName"
    ]);
    // Populate only if empty below header
    if (look.getLastRow() === 1){
      const defaults = defaultLookups();
      const headers = look.getRange(1,1,1,look.getLastColumn()).getValues()[0];
      headers.forEach((h,idx)=>{
        const arr = defaults[h];
        if (arr && arr.length){
          look.getRange(2, idx+1, arr.length, 1).setValues(arr.map(v=>[v]));
        }
      });
    }
    look.setFrozenRows(1);
  }

  // Dashboard sheet
  if (!ss.getSheetByName("Dashboard")){
    ss.insertSheet("Dashboard").getRange("A1").setValue("Dashboard");
  }

  let brc = ss.getSheetByName(BRC_SHEET);
  if(!brc){
    brc = ss.insertSheet(BRC_SHEET);
    brc.appendRow(BRC_HEADERS);
    brc.setFrozenRows(1);
  } else {
    ensureHeaders(brc, BRC_HEADERS);
  }

  let batchDetails = ss.getSheetByName(BATCH_DETAILS_SHEET);
  if(!batchDetails){
    batchDetails = ss.insertSheet(BATCH_DETAILS_SHEET);
    batchDetails.appendRow(BATCH_DETAILS_HEADERS);
    batchDetails.setFrozenRows(1);
  } else {
    ensureHeaders(batchDetails, BATCH_DETAILS_HEADERS);
  }

  let tracker = ss.getSheetByName(BRC_TRACKER_SHEET);
  if(!tracker){
    tracker = ss.insertSheet(BRC_TRACKER_SHEET);
    tracker.appendRow(BRC_TRACKER_HEADERS);
    tracker.setFrozenRows(1);
  } else {
    ensureHeaders(tracker, BRC_TRACKER_HEADERS);
  }

  return "Safe setup complete (no data cleared).";
}

/************************************************************
  ensureHeaders(sheet, headerArray) - only adds if missing
************************************************************/
function ensureHeaders(sh, headers){
  if (sh.getLastRow() === 0){
    sh.appendRow(headers);
    sh.setFrozenRows(1);
    return;
  }
  const first = sh.getRange(1,1,1,headers.length).getValues()[0];
  const blank = first.every(val => String(val).trim()==='');
  if(blank){
    sh.getRange(1,1,1,headers.length).setValues([headers]);
    sh.setFrozenRows(1);
  }
}

function getBrcStatus(batch){
  batch = (batch||'').trim();
  if(!batch) return {status:'error', message:'Batch required'};
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(BRC_SHEET);
  if(!sh){
    return {status:'error', message:'BRC Status sheet not found'};
  }
  const data = sh.getDataRange().getValues();
  if(data.length < 2){
    return {status:'error', message:'No BRC records yet'};
  }
  const header = data[0];
  const batchCol = header.indexOf('Batch No');
  const statusCol = header.indexOf('Status');
  const remarksCol = header.indexOf('Remarks');
  const updatedCol = header.indexOf('Updated On');
  if(batchCol === -1 || statusCol === -1 || remarksCol === -1 || updatedCol === -1){
    return {status:'error', message:'BRC sheet headers missing'};
  }
  for(let i=1;i<data.length;i++){
    const row = data[i];
    if(String(row[batchCol]||'').trim().toUpperCase() === batch.toUpperCase()){
      return {
        status:'success',
        batch:batch,
        data:{
          status: row[statusCol] || 'NA',
          remarks: row[remarksCol] || '',
          updatedOn: row[updatedCol] || ''
        }
      };
    }
  }
  return {status:'error', message:'No BRC status recorded for this batch'};
}

function getBatchDetails(batch){
  batch = (batch||'').trim();
  if(!batch) return {status:'error', message:'Batch required'};
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(BATCH_DETAILS_SHEET);
  if(!sh){
    return {status:'error', message:'Batch Details sheet not found'};
  }
  const data = sh.getDataRange().getValues();
  if(data.length < 2){
    return {status:'error', message:'No details yet'};
  }
  const header = data[0];
  const idx = header.indexOf('Batch No');
  if(idx === -1){
    return {status:'error', message:'Batch Details headers missing'};
  }
  for(let i=1;i<data.length;i++){
    const row = data[i];
    if(String(row[idx]||'').trim().toUpperCase() === batch.toUpperCase()){
      return {
        status:'success',
        batch:batch,
        data:
          {
            incident: row[header.indexOf('Incident number')] || '',
            deviation: row[header.indexOf('Deviation number')] || '',
            ca: row[header.indexOf('CA number')] || '',
            pa: row[header.indexOf('PA number')] || '',
            oos: row[header.indexOf('OOS number')] || '',
            updatedOn: row[header.indexOf('Updated On')] || ''
          }
      };
    }
  }
  return {status:'error', message:'No batch details recorded yet'};
}

function saveBatchDetails(payload){
  const batch = (payload && payload.BatchNo || '').trim();
  if(!batch) return {status:'error', message:'Batch number required'};
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(BATCH_DETAILS_SHEET) || ss.insertSheet(BATCH_DETAILS_SHEET);
  ensureHeaders(sh, BATCH_DETAILS_HEADERS);
  const lock = LockService.getDocumentLock();
  lock.waitLock(20000);
  try{
    const normalized = { // defaults to empty string if undefined
      incident: payload.IncidentNumber || '',
      deviation: payload.DeviationNumber || '',
      ca: payload.CANumber || '',
      pa: payload.PANumber || '',
      oos: payload.OOSNumber || ''
    };
    const data = sh.getDataRange().getValues();
    let targetRow = -1;
    if(data.length > 1){
      const header = data[0];
      const batchIdx = header.indexOf('Batch No');
      for(let i=1;i<data.length;i++){
        if(String(data[i][batchIdx]||'').trim().toUpperCase() === batch.toUpperCase()){
          targetRow = i+1; // because data index 0 is header
          break;
        }
      }
    }
    const rowValues = [
      batch,
      normalized.incident,
      normalized.deviation,
      normalized.ca,
      normalized.pa,
      normalized.oos,
      new Date()
    ];
    if(targetRow === -1){
      sh.appendRow(rowValues);
    } else {
      sh.getRange(targetRow,1,1,rowValues.length).setValues([rowValues]);
    }
    return {status:'success', message:'Batch details saved'};
  } catch(err){
    return {status:'error', message: err.message};
  } finally {
    lock.releaseLock();
  }
}

function getBrcTracker(batch){
  batch = (batch||'').trim();
  if(!batch) return {status:'error', message:'Batch required'};
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(BRC_TRACKER_SHEET);
  if(!sh){
    return {status:'error', message:'BRC Tracker sheet not found'};
  }
  const data = sh.getDataRange().getValues();
  if(data.length < 2){
    return {status:'error', message:'No tracker entries yet'};
  }
  const header = data[0];
  const batchIdx = header.indexOf('Batch No');
  if(batchIdx === -1){
    return {status:'error', message:'BRC Tracker headers missing'};
  }
  for(let i=1;i<data.length;i++){
    const row = data[i];
    if(String(row[batchIdx]||'').trim().toUpperCase() === batch.toUpperCase()){
      return {
        status:'success',
        batch:batch,
        data:{
          submitted: row[header.indexOf('BRC Submitted')] === true || String(row[header.indexOf('BRC Submitted')]||'').toUpperCase()==='TRUE',
          gdp: row[header.indexOf('GDP Completed')] === true || String(row[header.indexOf('GDP Completed')]||'').toUpperCase()==='TRUE',
          updatedOn: row[header.indexOf('Updated On')] || ''
        }
      };
    }
  }
  return {status:'error', message:'No tracker data yet'};
}

function saveBrcTracker(payload){
  const batch = (payload && payload.BatchNo || '').trim();
  if(!batch) return {status:'error', message:'Batch number required'};
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(BRC_TRACKER_SHEET) || ss.insertSheet(BRC_TRACKER_SHEET);
  ensureHeaders(sh, BRC_TRACKER_HEADERS);
  const lock = LockService.getDocumentLock();
  lock.waitLock(20000);
  try{
    const data = sh.getDataRange().getValues();
    let targetRow = -1;
    if(data.length > 1){
      const header = data[0];
      const batchIdx = header.indexOf('Batch No');
      for(let i=1;i<data.length;i++){
        if(String(data[i][batchIdx]||'').trim().toUpperCase() === batch.toUpperCase()){
          targetRow = i+1;
          break;
        }
      }
    }
    const toBool = v => v === true || String(v||'').toUpperCase()==='TRUE';
    const rowValues = [
      batch,
      toBool(payload.BrcSubmitted),
      toBool(payload.GdpCompleted),
      new Date()
    ];
    if(targetRow === -1){
      sh.appendRow(rowValues);
    } else {
      sh.getRange(targetRow,1,1,rowValues.length).setValues([rowValues]);
    }
    return {status:'success', message:'BRC tracker updated'};
  } catch(err){
    return {status:'error', message: err.message};
  } finally {
    lock.releaseLock();
  }
}

function defaultLookups(){
  return {
    Shift:["A-SHIFT","B-SHIFT","C-SHIFT","GENERAL"],
    LOT_NO:["NA","LOT-1","LOT-2","LOT-3","LOT-4","LOT-5"],
    Line:["LINE-A","LINE-B","LINE-C","LINE-D","LINE-E","AUTOMATION LINE"],
    VI1_Types:["WEAK WELD","DUST WELD","IMPROPER WELD","AIR BUBBLES","DAMAGE","NARROW CHANNEL","ALIGNMENT ISSUE","QC TORQUE TEST","CHILD PARTS WELDED","HAIR WELD","DUMP REJECTIONS","OIL","BULGING","NA"],
    VI2_Types:["WEAK WELD","DUST WELD","IMPROPER WELD","AIR BUBBLES","DAMAGE","ALIGNMENT ISSUE","WHITE LINE","NA"],
    VI3_Types:["PEAL OFF","TEAR OFF","DAMAGE","DUST WELD","WELDING REJECTIONS","OVERMELT","NA"],
    Vacuum_Types:["EN 6 LEAK","EN 4 LEAK","SN LOW","SN HIGH","EN LOW","EN HIGH","QR REJECTIONS","CLOCKED ERROR","NA"],
    VI4_Types:["IMPROPER WELDING","WEAK WELDING","DUST WELD","ALIGNMENT ISSUE","QR REJECTIONS","DAMAGE","OVERMELT","NARROW CHANNEL","MATRIC REJECTIONS","SEALING REJECTIONS","AIR BUBBLES","WELDING REJECTIONS","NA"],
    Equipment_LINE_A:["EC/EQID/III-00631","EC/EQID/III-00163"],
    Equipment_LINE_B:["EC/EQID/III-00572","EC/EQID/III-00133"],
    Equipment_LINE_C:["EC/EQID/III-00069","EC/EQID/III-00101"],
    Equipment_LINE_D:["EC/EQID/III-00003","EC/EQID/III-00633"],
    Equipment_LINE_E:["EC/EQID/III-00036","EC/EQID/III-00632"],
    Equipment_AUTOMATION_LINE:["EC/EQID/III-00659"],
    CartridgePart:["NA","ELUTE SIDE","MATRIX SIDE","SAMPLE FILETER SIDE","DUMP SIDE","FILTER RODS SIDE","SAMPLE FILTER BOTTOM","IN CHANNEL","QR CODE"],
    VerifiedByName:["L R NAIDU","KAUSHIK","SAI KUMAR","SRINU","ROHINI","NARENDRA","RAJU","CHAKRADHAR"]
  };
}