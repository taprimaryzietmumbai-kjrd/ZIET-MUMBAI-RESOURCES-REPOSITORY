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
    // Secure spreadsheet access actions; existing public workshop API remains unchanged.
    if (String(action).startsWith('secure') || String(action).startsWith('admin')) {
      return json_(handleSecureAction_(action, body));
    }
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


/* ================= SECURE TA SPREADSHEET ACCESS =================
   Admin setup: run initializeSecurityAdmin() once from the Apps Script editor.
   This stores a salted password hash in Script Properties; the password is never
   stored in the spreadsheet or source code.
*/
const TA_REGISTER_SHEET = 'Authorised_TAs';
const TA_REQUESTS_SHEET = 'TA_Access_Requests';
const SECURE_SESSION_SECONDS = 6 * 60 * 60;
const OTP_SECONDS = 5 * 60;

function initializeSecurityAdmin() {
  const props = PropertiesService.getScriptProperties();
  if (props.getProperty('SECURE_ADMIN_HASH')) {
    throw new Error('Administrator is already configured. Do not rerun this setup unless you intend to reset the admin password.');
  }
  const password = Utilities.getUuid() + '-' + Utilities.getUuid();
  const salt = Utilities.getUuid() + Utilities.getUuid();
  props.setProperty('SECURE_ADMIN_SALT', salt);
  props.setProperty('SECURE_ADMIN_HASH', securityHash_(password, salt));
  const register = getSecurityRegister_();
  const requests = getSecurityRequests_();
  [register, requests].forEach(sh => {
    sh.hideSheet();
    const protection = sh.protect().setDescription('ZIET Secure Access — owner only');
    protection.setWarningOnly(false);
    const owner = Session.getEffectiveUser().getEmail();
    if (owner) {
      protection.getEditors().forEach(user => {
        if (user.getEmail() !== owner) protection.removeEditor(user);
      });
      protection.addEditor(owner);
    }
    if (protection.canDomainEdit()) protection.setDomainEdit(false);
  });
  Logger.log('INITIAL ADMIN PASSWORD (copy and keep privately; it is shown only during this setup): ' + password);
  return 'Administrator configured. Copy the INITIAL ADMIN PASSWORD from the execution log and keep it privately. Do not run this setup again.';
}


// Emergency credential rotation. Run from the Apps Script editor if the admin
// password may have been exposed. The new password is shown once in the execution log.
function resetSecurityAdminPassword() {
  const password = Utilities.getUuid() + '-' + Utilities.getUuid();
  const salt = Utilities.getUuid() + Utilities.getUuid();
  const props = PropertiesService.getScriptProperties();
  props.setProperty('SECURE_ADMIN_SALT', salt);
  props.setProperty('SECURE_ADMIN_HASH', securityHash_(password, salt));
  CacheService.getScriptCache().removeAll([]);
  Logger.log('NEW INITIAL ADMIN PASSWORD (copy and keep privately): ' + password);
  return 'Administrator password rotated. Copy the new password from the execution log.';
}

function handleSecureAction_(action, body) {
  switch (action) {
    case 'securerequestaccess': return secureRequestAccess_(body);
    case 'secureverifyotp': return secureVerifyOtp_(body);
    case 'securegetspreadsheet': return secureGetSpreadsheet_(body);
    case 'adminlogin': return adminLogin_(body);
    case 'adminlistrequests': requireAdmin_(body.adminToken); return {success:true,requests:listAccessRequests_(),tas:listAuthorisedTAs_()};
    case 'adminapproverequest': requireAdmin_(body.adminToken); return adminApproveRequest_(body.requestId);
    case 'adminrejectrequest': requireAdmin_(body.adminToken); return adminRejectRequest_(body.requestId);
    case 'adminsaveta': requireAdmin_(body.adminToken); return adminSaveTA_(body.ta || {});
    case 'admindeactivateta': requireAdmin_(body.adminToken); return adminDeactivateTA_(body.employeeCode);
    default: throw new Error('Unsupported secure action');
  }
}

function securityHash_(value, salt) {
  const bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, salt + '|' + value, Utilities.Charset.UTF_8);
  return Utilities.base64Encode(bytes);
}

function getSecurityRegister_() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sh = ss.getSheetByName(TA_REGISTER_SHEET);
  if (!sh) sh = ss.insertSheet(TA_REGISTER_SHEET);
  const headers = ['Employee Code','Name','DOB','DOJ','Registered Mobile','Status','Created At','Updated At'];
  if (sh.getLastRow() === 0) sh.getRange(1,1,1,headers.length).setValues([headers]);
  sh.setFrozenRows(1);
  return sh;
}

function getSecurityRequests_() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sh = ss.getSheetByName(TA_REQUESTS_SHEET);
  if (!sh) sh = ss.insertSheet(TA_REQUESTS_SHEET);
  const headers = ['Request ID','Employee Code','Name','Requested At','Status','Approved At','Rejected At'];
  if (sh.getLastRow() === 0) sh.getRange(1,1,1,headers.length).setValues([headers]);
  sh.setFrozenRows(1);
  return sh;
}

function normaliseCode_(v) { return String(v || '').trim().toUpperCase(); }
function normaliseDate_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  const s = String(v || '').trim();
  if (/^\d{4}-\d{2}-\d{2}$/.test(s)) return s;
  const m = s.match(/^(\d{1,2})[\/.-](\d{1,2})[\/.-](\d{4})$/);
  return m ? m[3] + '-' + ('0'+m[2]).slice(-2) + '-' + ('0'+m[1]).slice(-2) : s;
}

function secureRequestAccess_(b) {
  const code = normaliseCode_(b.employeeCode);
  const dob = normaliseDate_(b.dob), doj = normaliseDate_(b.doj);
  const last4 = String(b.mobileLast4 || '').replace(/\D/g,'').slice(-4);
  if (!code || !dob || !doj || !/^\d{4}$/.test(last4))
    throw new Error('Enter Employee Code, DOB, DOJ and the last four mobile digits.');
  const sh = getSecurityRegister_(), data = sh.getDataRange().getValues();
  if (data.length < 2) throw new Error('No authorised TA records are configured yet.');
  const h = data[0].map(String);
  const ix = {code:h.indexOf('Employee Code'),dob:h.indexOf('DOB'),doj:h.indexOf('DOJ'),mobile:h.indexOf('Registered Mobile'),status:h.indexOf('Status'),name:h.indexOf('Name')};
  const match = data.slice(1).find(r => normaliseCode_(r[ix.code]) === code &&
    normaliseDate_(r[ix.dob]) === dob && normaliseDate_(r[ix.doj]) === doj &&
    String(r[ix.mobile] || '').replace(/\D/g,'').slice(-4) === last4 &&
    String(r[ix.status] || '').toLowerCase() === 'active');
  if (!match) throw new Error('The details could not be verified. Please contact the administrator.');
  const requestId = 'TA-' + Utilities.getUuid().replace(/-/g,'').slice(0,12).toUpperCase();
  getSecurityRequests_().appendRow([requestId,code,String(match[ix.name]||''),new Date(),'Pending','','']);
  return {success:true,requestId:requestId,message:'Identity matched. Your request is pending administrator approval.'};
}

function adminLogin_(b) {
  const props = PropertiesService.getScriptProperties();
  const salt = props.getProperty('SECURE_ADMIN_SALT'), expected = props.getProperty('SECURE_ADMIN_HASH');
  if (!salt || !expected) throw new Error('Administrator setup is not complete. Run initializeSecurityAdmin() in Apps Script first.');
  const supplied = securityHash_(String(b.password || ''), salt);
  if (supplied !== expected) throw new Error('Invalid administrator password.');
  const token = Utilities.getUuid() + Utilities.getUuid();
  CacheService.getScriptCache().put('ADMIN:' + token, '1', SECURE_SESSION_SECONDS);
  return {success:true,adminToken:token,expiresIn:SECURE_SESSION_SECONDS};
}

function requireAdmin_(token) {
  if (!token || CacheService.getScriptCache().get('ADMIN:' + token) !== '1')
    throw new Error('Administrator session expired. Sign in again.');
}

function listAccessRequests_() {
  const sh = getSecurityRequests_(), v = sh.getDataRange().getValues();
  return v.slice(1).filter(r => r[0]).map(r => ({
    requestId:String(r[0]),employeeCode:String(r[1]),name:String(r[2]),
    requestedAt:r[3] instanceof Date ? r[3].toISOString() : String(r[3]||''),
    status:String(r[4]||'Pending')
  })).reverse();
}

function listAuthorisedTAs_() {
  const sh = getSecurityRegister_(), v = sh.getDataRange().getValues();
  return v.slice(1).filter(r => r[0]).map(r => ({
    employeeCode:String(r[0]),name:String(r[1]),dob:normaliseDate_(r[2]),
    doj:normaliseDate_(r[3]),mobile:String(r[4]),status:String(r[5]||'Active')
  }));
}

function findRequestRow_(requestId) {
  const sh = getSecurityRequests_(), v = sh.getDataRange().getValues();
  for (let i=1;i<v.length;i++) if (String(v[i][0]) === String(requestId)) return {sheet:sh,row:i+1,values:v[i]};
  throw new Error('Access request not found.');
}

function adminApproveRequest_(requestId) {
  const found = findRequestRow_(requestId);
  if (String(found.values[4]) !== 'Pending') throw new Error('This request is no longer pending.');
  const otp = String(Math.floor(100000 + Math.random() * 900000));
  const key = 'OTP:' + requestId;
  CacheService.getScriptCache().put(key, securityHash_(otp, requestId), OTP_SECONDS);
  found.sheet.getRange(found.row,5).setValue('Approved — OTP issued');
  found.sheet.getRange(found.row,6).setValue(new Date());
  return {success:true,requestId:requestId,otp:otp,expiresIn:OTP_SECONDS,
    message:'Share this one-time code with the TA through WhatsApp. It expires in five minutes.'};
}

function adminRejectRequest_(requestId) {
  const found = findRequestRow_(requestId);
  found.sheet.getRange(found.row,5).setValue('Rejected');
  found.sheet.getRange(found.row,7).setValue(new Date());
  CacheService.getScriptCache().remove('OTP:' + requestId);
  return {success:true,message:'Request rejected.'};
}

function secureVerifyOtp_(b) {
  const requestId = String(b.requestId || ''), otp = String(b.otp || '').trim();
  if (!requestId || !/^\d{6}$/.test(otp)) throw new Error('Enter the six-digit access code.');
  const cached = CacheService.getScriptCache().get('OTP:' + requestId);
  if (!cached || cached !== securityHash_(otp, requestId)) throw new Error('Invalid or expired access code. Request approval again.');
  CacheService.getScriptCache().remove('OTP:' + requestId);
  const found = findRequestRow_(requestId);
  if (String(found.values[4]) !== 'Approved — OTP issued') throw new Error('This request is not approved.');
  const session = Utilities.getUuid() + Utilities.getUuid();
  CacheService.getScriptCache().put('VIEW:' + session, String(found.values[1]), SECURE_SESSION_SECONDS);
  found.sheet.getRange(found.row,5).setValue('Access granted');
  return {success:true,sessionToken:session,expiresIn:SECURE_SESSION_SECONDS};
}

function secureGetSpreadsheet_(b) {
  const token = String(b.sessionToken || '');
  if (!token) throw new Error('Sign in to view spreadsheet data.');
  if (!CacheService.getScriptCache().get('VIEW:' + token)) throw new Error('Your viewer session has expired. Request access again.');
  const which = String(b.sheet || 'master');
  const id = which === 'compendium' ? SECONDARY_SPREADSHEET_ID : (which === 'master' ? SPREADSHEET_ID : '');
  if (!id) throw new Error('Unknown spreadsheet selection.');
  const ss = SpreadsheetApp.openById(id);
  const sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) return {success:true,spreadsheet:ss.getName(),headers:[],rows:[]};
  const values = sh.getDataRange().getDisplayValues();
  const maxRows = 500, maxCols = 40;
  const clipped = values.slice(0,maxRows).map(r => r.slice(0,maxCols));
  return {success:true,spreadsheet:ss.getName(),headers:clipped[0]||[],rows:clipped.slice(1),truncated:values.length>maxRows||values.some(r=>r.length>maxCols)};
}

function adminSaveTA_(ta) {
  const code = normaliseCode_(ta.employeeCode);
  const name = String(ta.name || '').trim();
  const dob = normaliseDate_(ta.dob), doj = normaliseDate_(ta.doj);
  const mobile = String(ta.mobile || '').replace(/\D/g,'');
  if (!code || !name || !dob || !doj || !/^\d{10}$/.test(mobile))
    throw new Error('Enter Employee Code, Name, DOB, DOJ and a valid 10-digit mobile number.');
  const sh = getSecurityRegister_(), v = sh.getDataRange().getValues(), h = v[0].map(String);
  const codeCol = h.indexOf('Employee Code'), nameCol=h.indexOf('Name');
  let row = -1;
  for(let i=1;i<v.length;i++) if(normaliseCode_(v[i][codeCol])===code){row=i+1;break;}
  const now = new Date();
  const record = [code,name,dob,doj,mobile,String(ta.status||'Active'),now,now];
  if(row>0) { record[6]=v[row-1][6]||now; sh.getRange(row,1,1,record.length).setValues([record]); }
  else sh.appendRow(record);
  return {success:true,message:'TA record saved.'};
}

function adminDeactivateTA_(employeeCode) {
  const code=normaliseCode_(employeeCode), sh=getSecurityRegister_(), v=sh.getDataRange().getValues();
  for(let i=1;i<v.length;i++) if(normaliseCode_(v[i][0])===code){sh.getRange(i+1,6).setValue('Inactive');sh.getRange(i+1,8).setValue(new Date());return {success:true,message:'TA access deactivated.'};}
  throw new Error('Employee Code not found.');
}
