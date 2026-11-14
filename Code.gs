/**
 * ZIET Mumbai Workshop Resource Repository — Central API
 * Spreadsheet: 1zo-CnHZR5Pp88nCtwfuw5VSRfNi-vZcCJ2vlB6BE6hY
 * Deploy as Web App: Execute as Me; access Anyone.
 */
const SPREADSHEET_ID = '1zo-CnHZR5Pp88nCtwfuw5VSRfNi-vZcCJ2vlB6BE6hY';
const SHEET_NAME = 'Resources';
const HEADERS = [
  'resourceId','workshopId','workshopName','day','date','session',
  'title','resourceType','url','description','speaker','sourceApp',
  'createdAt','updatedAt','status'
];

function doGet(e) {
  const action = (e && e.parameter && e.parameter.action) || 'list';
  if (action === 'health') return json_({success:true,service:'ZIET Workshop Resource API',version:1});
  if (action !== 'list') return json_({success:false,message:'Unsupported GET action'});
  const rows = readAll_();
  // JSONP supports cross-origin reads from GitHub Pages without exposing credentials.
  const callback = String((e && e.parameter && e.parameter.callback) || '').trim();
  const payload = {success:true,resources:rows};
  if (callback && /^[A-Za-z_$][0-9A-Za-z_$\.]*$/.test(callback)) {
    return ContentService.createTextOutput(callback + '(' + JSON.stringify(payload) + ');')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return json_(payload);
}

function doPost(e) {
  try {
    let body = {};
    const raw = (e && e.postData && e.postData.contents) || '';
    if (raw) {
      try { body = JSON.parse(raw); }
      catch (jsonError) {
        // HTML form POST is used by static GitHub Pages clients to avoid CORS preflight.
        const encoded = e && e.parameter && e.parameter.payload;
        if (encoded) body = JSON.parse(encoded);
        else throw jsonError;
      }
    } else if (e && e.parameter && e.parameter.payload) {
      body = JSON.parse(e.parameter.payload);
    }
    const action = String(body.action || '').toLowerCase();
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
    return json_({success:false,message:String(err && err.message || err)});
  }
}

function getSheet_() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) sh = ss.insertSheet(SHEET_NAME);
  if (sh.getLastRow() === 0) {
    sh.getRange(1,1,1,HEADERS.length).setValues([HEADERS]);
    sh.setFrozenRows(1);
  } else {
    const current = sh.getRange(1,1,1,Math.max(sh.getLastColumn(),HEADERS.length)).getValues()[0];
    HEADERS.forEach((h,i) => { if (current.indexOf(h) === -1) sh.getRange(1,sh.getLastColumn()+1).setValue(h); });
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
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    const sh = getSheet_();
    const data = sh.getDataRange().getValues();
    const headers = data[0].map(String);
    const idCol = headers.indexOf('resourceId');
    let rowNum = -1;
    for (let i=1;i<data.length;i++) if (String(data[i][idCol]) === String(input.resourceId)) { rowNum=i+1; break; }
    const now = new Date().toISOString();
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
    return {success:true,resource:record,message:'Resource saved'};
  } finally { lock.releaseLock(); }
}

function softDelete_(resourceId, sourceApp) {
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    const sh = getSheet_(), data = sh.getDataRange().getValues(), headers = data[0].map(String);
    const idCol = headers.indexOf('resourceId'), statusCol = headers.indexOf('status');
    for (let i=1;i<data.length;i++) {
      if (String(data[i][idCol]) === resourceId) {
        sh.getRange(i+1,statusCol+1).setValue('deleted');
        const updatedCol = headers.indexOf('updatedAt');
        if (updatedCol >= 0) sh.getRange(i+1,updatedCol+1).setValue(new Date().toISOString());
        const sourceCol = headers.indexOf('sourceApp');
        if (sourceCol >= 0) sh.getRange(i+1,sourceCol+1).setValue(sourceApp);
        return {success:true,resourceId:resourceId,status:'deleted'};
      }
    }
    return {success:true,resourceId:resourceId,status:'deleted',message:'Record was already absent'};
  } finally { lock.releaseLock(); }
}

function rowToObject_(headers,row) {
  const out={}; headers.forEach((h,i)=>out[h]=row[i]); return out;
}
function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
