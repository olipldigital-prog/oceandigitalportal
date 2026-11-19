/**
 * OCEAN ATTENDANCE PORTAL — Google Apps Script backend
 * Deploy as Web app: Execute as Me; Who has access: your organisation / authorised users.
 * Run setupPortal() once from the editor, then change the starter passwords immediately.
 */
const CFG = {
  SPREADSHEET_ID: '', // Leave blank to create/use the spreadsheet bound to this Apps Script project.
  PHOTO_FOLDER_ID: '', // Optional: existing private Drive folder ID for worker photos.
  TIMEZONE: 'Asia/Kolkata',
  SESSION_HOURS: 8,
  DEFAULT_ADMIN_ID: 'admin',
  DEFAULT_ADMIN_PASSWORD: 'Ocean@12345',
  DEFAULT_SECURITY_ID: 'security',
  DEFAULT_SECURITY_PASSWORD: 'Gate@12345'
};
const SHEETS = {
  Workers: ['Worker ID','Full Name','Vendor','Designation','Project','Project Address','DOB','Sex','Blood Group','ID Ref','Date of Issue','Validity Up To','EOHS Inducted On','Issuing Authority','Emergency Contact','Project Manager','Safety Manager','Emergency Contact Name','Photo Drive ID','Status','QR Active','Created At','Updated At'],
  Attendance: ['S.No','Date','Worker ID','Worker Name','Designation','Contractor Agency','Project Location','IN-Time','OUT-Time','Total Hrs','Overtime','Verified By','Last Action At'],
  Users: ['Role','User ID','Salt','Password Hash','Updated At'],
  Settings: ['Key','Value'],
  ActivityLog: ['Timestamp','Actor','Role','Action','Worker ID','Details']
};
function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('Ocean Attendance Portal')
    .addMetaTag('viewport','width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}
function setupPortal() {
  const ss = getSS_();
  Object.keys(SHEETS).forEach(name => {
    let sh = ss.getSheetByName(name);
    if (!sh) sh = ss.insertSheet(name);
    if (sh.getLastRow() === 0) sh.appendRow(SHEETS[name]);
    sh.setFrozenRows(1);
  });
  const users = sheet_('Users');
  if (users.getLastRow() < 2) {
    createUser_('admin', CFG.DEFAULT_ADMIN_ID, CFG.DEFAULT_ADMIN_PASSWORD);
    createUser_('security', CFG.DEFAULT_SECURITY_ID, CFG.DEFAULT_SECURITY_PASSWORD);
  }
  const defaults = {
    company: 'OCEAN LIFESPACES INDIA PVT LTD',
    address: 'Ocean One, MF-1, Industrial Estate, CIPET Hostel Road, Guindy, Chennai - 600032',
    workHours: 8,
    projects: []
  };
  const sh = sheet_('Settings');
  const existing = readSettings_();
  Object.keys(defaults).forEach(k => { if (!(k in existing)) sh.appendRow([k, JSON.stringify(defaults[k])]); });
  return 'Setup complete. Spreadsheet: ' + ss.getUrl() + '. Initial admin: ' + CFG.DEFAULT_ADMIN_ID + ' / ' + CFG.DEFAULT_ADMIN_PASSWORD + '; security: ' + CFG.DEFAULT_SECURITY_ID + ' / ' + CFG.DEFAULT_SECURITY_PASSWORD + '. Change both passwords now.';
}
function getSS_() {
  if (CFG.SPREADSHEET_ID) return SpreadsheetApp.openById(CFG.SPREADSHEET_ID);
  const active = SpreadsheetApp.getActiveSpreadsheet();
  if (!active) throw new Error('Set CFG.SPREADSHEET_ID to a Google Sheet ID, then run setupPortal().');
  return active;
}
function sheet_(name) { const sh = getSS_().getSheetByName(name); if (!sh) throw new Error('Missing sheet '+name+'. Run setupPortal() first.'); return sh; }
function now_() { return new Date(); }
function dateKey_(d) { return Utilities.formatDate(d, CFG.TIMEZONE, 'yyyy-MM-dd'); }
function timeKey_(d) { return Utilities.formatDate(d, CFG.TIMEZONE, 'HH:mm:ss'); }
function iso_(v) {
  if (!v) return '';
  if (Object.prototype.toString.call(v)==='[object Date]' && !isNaN(v)) return Utilities.formatDate(v, CFG.TIMEZONE, 'yyyy-MM-dd');
  return String(v);
}
function createUser_(role, id, password) {
  const salt = Utilities.getUuid().replace(/-/g,'');
  sheet_('Users').appendRow([role,id,salt,hash_(password,salt),now_()]);
}
function hash_(password,salt) {
  const bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, salt + ':' + password, Utilities.Charset.UTF_8);
  return bytes.map(b => ('0'+((b+256)%256).toString(16)).slice(-2)).join('');
}
function rows_(name) {
  const sh=sheet_(name), vals=sh.getDataRange().getValues();
  if (vals.length<2) return [];
  const headers=vals[0];
  return vals.slice(1).map((r,i)=>{const o={_row:i+2};headers.forEach((h,j)=>o[h]=r[j]);return o;});
}
function apiLogin(role,userId,password) {
  role=String(role||'').toLowerCase(); userId=String(userId||'').trim();
  if (!['admin','security'].includes(role)) return {ok:false,message:'Invalid access role.'};
  const u=rows_('Users').find(x=>String(x.Role)===role && String(x['User ID'])===userId);
  if (!u || hash_(String(password||''),String(u.Salt))!==String(u['Password Hash'])) {
    log_('Unknown',role,'LOGIN_FAILED','',userId); return {ok:false,message:'Invalid User ID or password.'};
  }
  const token=Utilities.getUuid()+Utilities.getUuid();
  CacheService.getScriptCache().put('sess_'+token,JSON.stringify({role,userId,created:Date.now()}),CFG.SESSION_HOURS*3600);
  log_(userId,role,'LOGIN','', 'Successful login');
  return {ok:true,token,userId,role};
}
function auth_(token,allowed) {
  const raw=CacheService.getScriptCache().get('sess_'+String(token||''));
  if (!raw) throw new Error('Session expired. Please log in again.');
  const s=JSON.parse(raw);
  if (allowed && !allowed.includes(s.role)) throw new Error('You do not have permission for this action.');
  return s;
}
function apiGetInitData(token) {
  const s=auth_(token,['admin']);
  const workers=rows_('Workers').map(w=>({
    workerId:String(w['Worker ID']),name:String(w['Full Name']),vendor:String(w.Vendor),designation:String(w.Designation),
    project:String(w.Project),address:String(w['Project Address']),dob:iso_(w.DOB),sex:String(w.Sex),blood:String(w['Blood Group']),
    aadhaar:String(w['ID Ref']),issue:iso_(w['Date of Issue']),validity:iso_(w['Validity Up To']),induction:iso_(w['EOHS Inducted On']),
    authority:String(w['Issuing Authority']),emergency:String(w['Emergency Contact']),pm:String(w['Project Manager']),
    sm:String(w['Safety Manager']),emergencyName:String(w['Emergency Contact Name']),photoId:String(w['Photo Drive ID']),
    status:String(w.Status||'Active'),qrActive:String(w['QR Active']).toUpperCase()==='TRUE'
  }));
  const attendance=rows_('Attendance').map(a=>({
    serial:a['S.No'],date:iso_(a.Date),workerId:String(a['Worker ID']),workerName:String(a['Worker Name']),
    designation:String(a.Designation),vendor:String(a['Contractor Agency']),project:String(a['Project Location']),
    inTime:a['IN-Time'] instanceof Date?timeKey_(a['IN-Time']):String(a['IN-Time']||''),
    outTime:a['OUT-Time'] instanceof Date?timeKey_(a['OUT-Time']):String(a['OUT-Time']||''),
    totalHours:Number(a['Total Hrs']||0),overtime:Number(a.Overtime||0),verifiedBy:String(a['Verified By']||'')
  }));
  return {ok:true,workers,attendance,settings:readSettings_()};
}
function apiGetWorkerPhoto(token,workerId) {
  auth_(token,['admin','security']);
  const w=rows_('Workers').find(x=>String(x['Worker ID'])===String(workerId));
  if (!w || !w['Photo Drive ID']) return {ok:true,photoData:''};
  const f=DriveApp.getFileById(String(w['Photo Drive ID']));
  return {ok:true,photoData:'data:'+f.getMimeType()+';base64,'+Utilities.base64Encode(f.getBlob().getBytes())};
}
function apiSaveWorker(token,worker) {
  const s=auth_(token,['admin']);
  if (!worker || !worker.workerId || !worker.name || !worker.vendor || !worker.designation || !worker.project) return {ok:false,message:'Complete all required worker details.'};
  const id=String(worker.workerId).trim();
  const sh=sheet_('Workers'), all=rows_('Workers'), found=all.find(x=>String(x['Worker ID'])===id);
  let photoId=found?String(found['Photo Drive ID']||''):'';
  if (worker.photoData && worker.photoData.indexOf('base64,')>-1) {
    const match=worker.photoData.match(/^data:(image\/[a-zA-Z0-9.+-]+);base64,(.+)$/);
    if (!match) return {ok:false,message:'Invalid photo format.'};
    const bytes=Utilities.base64Decode(match[2]);
    if (bytes.length>2*1024*1024) return {ok:false,message:'Photo is too large. Upload a compressed image under 2 MB.'};
    const blob=Utilities.newBlob(bytes,match[1],id.replace(/[^a-zA-Z0-9_-]/g,'_')+'.jpg');
    const folder=getPhotoFolder_();
    const file=photoId?DriveApp.getFileById(photoId):folder.createFile(blob);
    if (photoId) { const replacement=folder.createFile(blob); file.setTrashed(true); photoId=replacement.getId(); }
    else photoId=file.getId();
  }
  const vals=[id,worker.name,worker.vendor,worker.designation,worker.project,worker.address||'',worker.dob||'',worker.sex||'',worker.blood||'',worker.aadhaar||'',worker.issue||'',worker.validity||'',worker.induction||'',worker.authority||'Project HR',worker.emergency||'',worker.pm||'',worker.sm||'',worker.emergencyName||'',photoId,'Active',true,found?found['Created At']:now_(),now_()];
  if (found) sh.getRange(found._row,1,1,vals.length).setValues([vals]); else sh.appendRow(vals);
  log_(s.userId,s.role,'WORKER_SAVE',id,worker.name);
  return {ok:true,message:'Worker saved. QR code is active.'};
}
function getPhotoFolder_() {
  if (CFG.PHOTO_FOLDER_ID) return DriveApp.getFolderById(CFG.PHOTO_FOLDER_ID);
  const props=PropertiesService.getScriptProperties();
  let id=props.getProperty('OCEAN_PHOTO_FOLDER');
  if (id) { try{return DriveApp.getFolderById(id);}catch(e){} }
  const f=DriveApp.createFolder('Ocean Attendance Worker Photos - Private');
  props.setProperty('OCEAN_PHOTO_FOLDER',f.getId()); return f;
}
function apiSetWorkerStatus(token,workerId,status) {
  const s=auth_(token,['admin']); const w=rows_('Workers').find(x=>String(x['Worker ID'])===String(workerId));
  if (!w) return {ok:false,message:'Worker not found.'};
  const active=status==='Active'; const sh=sheet_('Workers');
  sh.getRange(w._row,20,1,2).setValues([[active?'Active':'Inactive',active]]);
  sh.getRange(w._row,23).setValue(now_());
  log_(s.userId,s.role,'QR_'+(active?'ACTIVATE':'DEACTIVATE'),workerId,'');
  return {ok:true};
}
function apiRecordScan(token,workerId,requestedAction,verifiedBy) {
  const s=auth_(token,['admin','security']); workerId=String(workerId||'').trim();
  const w=rows_('Workers').find(x=>String(x['Worker ID'])===workerId);
  if (!w || String(w.Status)!=='Active' || String(w['QR Active']).toUpperCase()!=='TRUE') return {ok:false,message:'Invalid, inactive or deactivated worker QR ID.'};
  const lock=LockService.getScriptLock(); lock.waitLock(10000);
  try {
    const sh=sheet_('Attendance'), rows=rows_('Attendance'), today=dateKey_(now_());
    const open=rows.filter(a=>String(a['Worker ID'])===workerId && iso_(a.Date)===today && !a['OUT-Time']).sort((a,b)=>b._row-a._row)[0];
    let action=String(requestedAction||'AUTO').toUpperCase();
    if (action==='AUTO') action=open?'OUT':'IN';
    if (action==='IN' && open) return {ok:false,message:'Worker is already clocked IN. Clock OUT first.'};
    if (action==='OUT' && !open) return {ok:false,message:'No open clock-IN record found for today.'};
    const now=now_(), t=timeKey_(now);
    let rowNumber;
    if (action==='IN') {
      sh.appendRow([sh.getLastRow(),today,workerId,w['Full Name'],w.Designation,w.Vendor,w.Project,t,'',0,0,verifiedBy||s.userId,now]);
      rowNumber=sh.getLastRow();
    } else {
      rowNumber=open._row;
      const start=open['IN-Time'] instanceof Date?open['IN-Time']:parseTimeToday_(today,String(open['IN-Time']));
      const hours=Math.max(0,(now.getTime()-start.getTime())/3600000);
      const workHours=Number(readSettings_().workHours||8), ot=Math.max(0,hours-workHours);
      sh.getRange(rowNumber,9,1,5).setValues([[t,Number(hours.toFixed(2)),Number(ot.toFixed(2)),verifiedBy||s.userId,now]]);
    }
    log_(s.userId,s.role,'CLOCK_'+action,workerId,w['Full Name']);
    return {ok:true,action,workerName:String(w['Full Name']),date:today,time:t,message:w['Full Name']+' clock '+(action==='IN'?'IN':'OUT')+' recorded at '+t,worker:{workerId,name:w['Full Name'],designation:w.Designation,vendor:w.Vendor,project:w.Project},photoData:photoData_(w)};
  } finally { lock.releaseLock(); }
}
function parseTimeToday_(day,t) { return Utilities.parseDate(day+' '+t, CFG.TIMEZONE, 'yyyy-MM-dd HH:mm:ss'); }
function photoData_(w) {
  if (!w['Photo Drive ID']) return '';
  try { const f=DriveApp.getFileById(String(w['Photo Drive ID'])); return 'data:'+f.getMimeType()+';base64,'+Utilities.base64Encode(f.getBlob().getBytes()); } catch(e) { return ''; }
}
function apiUpdateCredentials(token,payload) {
  const s=auth_(token,['admin']);
  if (!payload || !payload.newUserId || !payload.newPassword || String(payload.newPassword).length<8) return {ok:false,message:'Enter a user ID and password of at least 8 characters.'};
  const admin=rows_('Users').find(x=>String(x.Role)==='admin');
  if (!admin || hash_(String(payload.currentAdminPassword||''),String(admin.Salt))!==String(admin['Password Hash'])) return {ok:false,message:'Current Admin password is incorrect.'};
  const target=String(payload.target);
  if (!['admin','security'].includes(target)) return {ok:false,message:'Invalid account.'};
  const sh=sheet_('Users'),u=rows_('Users').find(x=>String(x.Role)===target);
  const salt=Utilities.getUuid().replace(/-/g,'');
  sh.getRange(u._row,2,1,4).setValues([[String(payload.newUserId).trim(),salt,hash_(String(payload.newPassword),salt),now_()]]);
  log_(s.userId,s.role,'CREDENTIAL_CHANGE','',target);
  return {ok:true};
}
function readSettings_() {
  const out={}; rows_('Settings').forEach(r=>{try{out[String(r.Key)]=JSON.parse(String(r.Value));}catch(e){out[String(r.Key)]=r.Value;}});
  return out;
}
function apiUpdateSettings(token,settings) {
  const s=auth_(token,['admin']); const sh=sheet_('Settings'), old=rows_('Settings');
  const obj={company:settings.company||'',address:settings.address||'',workHours:Number(settings.workHours||8),projects:settings.projects||[]};
  Object.keys(obj).forEach(k=>{const row=old.find(x=>String(x.Key)===k);if(row)sh.getRange(row._row,2).setValue(JSON.stringify(obj[k]));else sh.appendRow([k,JSON.stringify(obj[k])]);});
  log_(s.userId,s.role,'SETTINGS_UPDATE','','Updated portal settings'); return {ok:true};
}
function log_(actor,role,action,workerId,details) {
  try { sheet_('ActivityLog').appendRow([now_(),actor,role,action,workerId,details]); } catch(e) {}
}
