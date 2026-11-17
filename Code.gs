/**
 * ZIET Mumbai Workshop Resource Repository — Central API + Google Drive Upload
 * Spreadsheet: 1zo-CnHZR5Pp88nCtwfuw5VSRfNi-vZcCJ2vlB6BE6hY
 * Coordinator folders parent: 1BdsPAFk5CCVHP161h8DhF7_oLg3a2_4T
 * Deploy as Web App: Execute as Me; access Anyone.
 */
const SPREADSHEET_ID = '1zo-CnHZR5Pp88nCtwfuw5VSRfNi-vZcCJ2vlB6BE6hY';
const SECONDARY_SPREADSHEET_ID = '1Aadm4x_qHczMzV6uF9F6NLYRbTy_pll1Ls2ncaguXzM';
const SHEET_NAME = 'Resources';
const COORDINATOR_ROOT_ID = '1BdsPAFk5CCVHP161h8DhF7_oLg3a2_4T';
const COORDINATORS = [
  'Mrs. Seema Mishra - TA Economics',
  'Mrs. Deepa Dilwari - TA English',
  'Mr. Gajendra Kapde - TA Chemistry',
  'Mr. Nilesh Thangan - TA Physics',
  'Dr. Rajesh Kumar Sharma - TA Library',
  'Anand Kr. Srivastava - TA (F&P)'
];
const CATEGORIES = ['Course Module','Course Design','Concept Map','Course Report','Additional Repositories'];
const MAX_UPLOAD_BYTES = 10 * 1024 * 1024;
const HEADERS = [
  'resourceId','workshopId','workshopName','day','date','session',
  'title','resourceType','url','description','speaker','sourceApp',
  'createdAt','updatedAt','status','coordinator','driveFileId',
  'driveFolderId','storageMode','fileName','mimeType','targetGroup','workshopFrom','workshopTo'
];

function doGet(e) {
  const action = (e && e.parameter && e.parameter.action) || 'list';
  if (action === 'health') return json_({success:true,service:'ZIET Workshop Resource API',version:2});
  if (action !== 'list') return json_({success:false,message:'Unsupported GET action'});
  const rows = readAll_();
  const callback = String((e && e.parameter && e.parameter.callback) || '').trim();
  const payload = {success:true,resources:rows};
  if (callback && /^[A-Za-z_$][0-9A-Za-z_$\\.]*$/.test(callback)) {
    return ContentService.createTextOutput(callback + '(' + JSON.stringify(payload) + ');')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return json_(payload);
}

function doPost(e) {
  let createdFile = null;
  try {
    let body = {};
    const raw = (e && e.postData && e.postData.contents) || '';
    if (raw) {
      try { body = JSON.parse(raw); }
      catch (jsonError) {
        const encoded = e && e.parameter && e.parameter.payload;
        if (encoded) body = JSON.parse(encoded);
        else throw jsonError;
      }
    } else if (e && e.parameter && e.parameter.payload) {
      body = JSON.parse(e.parameter.payload);
    }
    const action = String(body.action || '').toLowerCase();
    if (action === 'upload') {
      const resource = body.resource || body;
      const uploaded = saveDriveUpload_(body, resource);
      createdFile = uploaded.file;
      const record = Object.assign({}, resource, {
        url: uploaded.file.getUrl(),
        driveFileId: uploaded.file.getId(),
        driveFolderId: uploaded.folder.getId(),
        storageMode: 'google-drive-upload',
        fileName: uploaded.file.getName(),
        mimeType: uploaded.file.getMimeType()
      });
      const saved = upsert_(record, body.sourceApp || 'central-repository');
      return json_({success:true,resource:saved.resource,driveUrl:uploaded.file.getUrl(),
        fileId:uploaded.file.getId(),folderId:uploaded.folder.getId(),
        folderName:uploaded.folder.getName(),message:'File uploaded to Coordinator Google Drive folder and repository record saved'});
    }
    if (action === 'upsert' || action === 'save') {
      const resource = body.resource || body;
      return json_(upsert_(resource, body.sourceApp || resource.sourceApp || 'unknown'));
    }
    if (action === 'delete') {
      const resourceId = body.resourceId || (body.resource && body.resource.resourceId);
      if (!resourceId) throw new Error('resourceId is required');
      return json_(softDelete_(String(resourceId), body.sourceApp || 'unknown'));
    }
    if (action === 'list') return json_({success:true,resources:readAll_()});
    throw new Error('Unsupported action: ' + action);
  } catch (err) {
    if (createdFile) {
      try { createdFile.setTrashed(true); } catch (cleanupError) {}
    }
    return json_({success:false,message:String(err && err.message || err)});
  }
}


/**
 * Run this once from the Apps Script editor as the script owner.
 * It requests Drive access and creates any missing Coordinator/category folders.
 */
function authorizeDriveAccess() {
  const root = DriveApp.getFolderById(COORDINATOR_ROOT_ID);
  let created = 0;
  COORDINATORS.forEach(name => {
    const coordinatorFolder = getOrCreateChildFolder_(root, name);
    CATEGORIES.forEach(category => {
      const before = coordinatorFolder.getFoldersByName(category).hasNext();
      getOrCreateChildFolder_(coordinatorFolder, category);
      if (!before) created++;
    });
  });
  // Confirm both reporting spreadsheets are accessible and create the secondary Resources tab if needed.
  SpreadsheetApp.openById(SPREADSHEET_ID).getName();
  const secondaryName = SpreadsheetApp.openById(SECONDARY_SPREADSHEET_ID).getName();
  getSecondarySheet_();
  return 'Drive access authorised. Secondary reporting sheet ready: ' + secondaryName + '. Missing category folders created: ' + created;
}

function saveDriveUpload_(body, resource) {
  if (!resource.workshopId) throw new Error('Select a workshop before uploading');
  if (!resource.title) throw new Error('Document title is required');
  if (!resource.targetGroup || !resource.workshopFrom || !resource.workshopTo) throw new Error('Target Groups, Workshop From and Workshop To are required');
  if (String(resource.workshopTo) < String(resource.workshopFrom)) throw new Error('Workshop To cannot be earlier than Workshop From');
  if (COORDINATORS.indexOf(String(resource.coordinator || '')) === -1)
    throw new Error('Select a valid Coordinator');
  if (CATEGORIES.indexOf(String(resource.resourceType || '')) === -1)
    throw new Error('Drive upload is available for Course Module, Course Design, Concept Map and Course Report');
  const fileName = String(body.fileName || '').trim();
  const base64 = String(body.fileBase64 || '');
  if (!fileName || !base64) throw new Error('Choose a file to upload');
  const bytes = Utilities.base64Decode(base64);
  if (bytes.length > MAX_UPLOAD_BYTES) throw new Error('File exceeds the 10 MB upload limit');
  if (!bytes.length) throw new Error('The selected file is empty');

  const root = DriveApp.getFolderById(COORDINATOR_ROOT_ID);
  const coordinatorFolder = getOrCreateChildFolder_(root, String(resource.coordinator));
  const categoryFolder = getOrCreateChildFolder_(coordinatorFolder, String(resource.resourceType));
  const mimeType = String(body.mimeType || 'application/octet-stream');
  const blob = Utilities.newBlob(bytes, mimeType, fileName);
  const file = categoryFolder.createFile(blob);
  return {file:file,folder:categoryFolder};
}

function getOrCreateChildFolder_(parent, name) {
  const folders = parent.getFoldersByName(name);
  return folders.hasNext() ? folders.next() : parent.createFolder(name);
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function getSheet_() {
  return getSheetFromSpreadsheet_(SPREADSHEET_ID);
}

// Creates/uses a Resources tab in the secondary reporting spreadsheet without changing other tabs.
function getSecondarySheet_() {
  return getSheetFromSpreadsheet_(SECONDARY_SPREADSHEET_ID);
}

function getSheetFromSpreadsheet_(spreadsheetId) {
  const ss = SpreadsheetApp.openById(spreadsheetId);
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) sh = ss.insertSheet(SHEET_NAME);
  if (sh.getLastRow() === 0) {
    sh.getRange(1,1,1,HEADERS.length).setValues([HEADERS]);
    sh.setFrozenRows(1);
  } else {
    const current = sh.getRange(1,1,1,Math.max(sh.getLastColumn(),HEADERS.length)).getValues()[0];
    HEADERS.forEach(h => { if (current.indexOf(h) === -1) sh.getRange(1,sh.getLastColumn()+1).setValue(h); });
  }
  return sh;
}

function readAll_() {
  const sh = getSheet_();
  const values = sh.getDataRange().getValues();
  if (values.length < 2) return [];
  const headers = values[0].map(String);
  return values.slice(1).filter(row => row.some(v => v !== '')).map(row => {
    const obj = {};
    headers.forEach((h,i) => obj[h] = row[i] instanceof Date
      ? Utilities.formatDate(row[i], Session.getScriptTimeZone(), 'yyyy-MM-dd') : row[i]);
    return obj;
  });
}

function upsert_(input, sourceApp) {
  if (!input || !input.resourceId) throw new Error('resourceId is required');
  if (!input.workshopId) throw new Error('workshopId is required');
  if (!input.title || !input.url) throw new Error('title and url are required');
  if (!input.targetGroup || !input.workshopFrom || !input.workshopTo) throw new Error('Target Groups, Workshop From and Workshop To are required');
  if (String(input.workshopTo) < String(input.workshopFrom)) throw new Error('Workshop To cannot be earlier than Workshop From');
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    const now = new Date().toISOString();
    const primary = upsertIntoSheet_(getSheet_(), input, sourceApp, now);
    let secondary;
    if (SECONDARY_SPREADSHEET_ID && SECONDARY_SPREADSHEET_ID !== SPREADSHEET_ID) {
      secondary = upsertIntoSheet_(getSecondarySheet_(), input, sourceApp, now);
    }
    return {success:true,resource:primary,secondarySynced:!!secondary,
      message: secondary ? 'Resource saved to both spreadsheets' : 'Resource saved'};
  } finally { lock.releaseLock(); }
}

function upsertIntoSheet_(sh, input, sourceApp, now) {
  const data = sh.getDataRange().getValues();
  const headers = data[0].map(String);
  const idCol = headers.indexOf('resourceId');
  let rowNum = -1;
  for (let i=1;i<data.length;i++) if (String(data[i][idCol]) === String(input.resourceId)) { rowNum=i+1; break; }
  const existing = rowNum > 0 ? rowToObject_(headers,data[rowNum-1]) : {};
  const record = Object.assign({}, existing, input, {
    sourceApp: sourceApp,
    createdAt: existing.createdAt || input.createdAt || now,
    updatedAt: now,
    status: 'active'
  });
  const row = headers.map(h => record[h] !== undefined ? record[h] : '');
  if (rowNum > 0) sh.getRange(rowNum,1,1,headers.length).setValues([row]);
  else sh.appendRow(row);
  return record;
}

function softDelete_(resourceId, sourceApp) {
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    const primary = softDeleteInSheet_(getSheet_(), resourceId, sourceApp);
    let secondary;
    if (SECONDARY_SPREADSHEET_ID && SECONDARY_SPREADSHEET_ID !== SPREADSHEET_ID) {
      secondary = softDeleteInSheet_(getSecondarySheet_(), resourceId, sourceApp);
    }
    return {success:true,resourceId:resourceId,status:'deleted',secondarySynced:!!secondary};
  } finally { lock.releaseLock(); }
}

function softDeleteInSheet_(sh, resourceId, sourceApp) {
  const data = sh.getDataRange().getValues(), headers = data[0].map(String);
  const idCol = headers.indexOf('resourceId'), statusCol = headers.indexOf('status');
  for (let i=1;i<data.length;i++) {
    if (String(data[i][idCol]) === resourceId) {
      sh.getRange(i+1,statusCol+1).setValue('deleted');
      const updatedCol = headers.indexOf('updatedAt');
      if (updatedCol >= 0) sh.getRange(i+1,updatedCol+1).setValue(new Date().toISOString());
      const sourceCol = headers.indexOf('sourceApp');
      if (sourceCol >= 0) sh.getRange(i+1,sourceCol+1).setValue(sourceApp);
      return true;
    }
  }
  return false;
}

function rowToObject_(headers,row) {
  const out={}; headers.forEach((h,i)=>out[h]=row[i]); return out;
}
