/*
 * ORAI FIELDWORK APP — SERVER (Google Apps Script)
 * EDD-I Orai, DVVNL. Developed by Nagendra Singh.
 *
 * FIRST-TIME SETUP
 *  1. Create a new Google Sheet named "Orai Fieldwork Data".
 *  2. Extensions > Apps Script. Delete the sample code, paste this whole file, Save.
 *     (Updating an existing server: paste the new code, Save, Run setupOrai once,
 *      then Deploy > Manage deployments > Edit (pencil) > Version: New version > Deploy.
 *      The /exec link stays the same.)
 *  3. Choose the function "setupOrai" at the top and click Run (allow the permissions).
 *     It creates the Users, Hierarchy, Roles and Mobile Data sheets and the admin login.
 *  4. Deploy > New deployment > type "Web app":
 *       Execute as: Me      Who has access: Anyone
 *     Copy the Web app URL (ends in /exec) and send it to Claude for the app.
 *  5. Log in to the app as admin, then CHANGE THE PASSWORD (Admin > Reset Password).
 *
 * Divisions / subdivisions / substations: edit the "Hierarchy" sheet.
 * Consumer master: upload from the app (Admin > Master Data). Files go to the
 *   Drive folder "Orai Fieldwork Master Data"; the "Masters" sheet lists them.
 * Roles and their scope: edit the "Roles" sheet
 *   (scope = DIVISION, SUBDIVISION, SUBSTATION_MULTI or SUBSTATION_SINGLE).
 */

const SHEET_NAME = 'Mobile Data';
const USERS_SHEET = 'Users';
const HIERARCHY_SHEET = 'Hierarchy';
const ROLES_SHEET = 'Roles';

/* Admin created by setupOrai(). Change the password after first login. */
const ADMIN_USER_ID = 'admin';
const ADMIN_NAME = 'Nagendra Singh (SDO)';
const ADMIN_PASSWORD = '';   /* only used if the Users sheet does not exist yet: set one before first setup */
const ORAI_DIVISION = 'EDD-I Orai';

/* Users sheet columns (1-based) */
const U = {ID:1, NAME:2, HASH:3, ACTIVE:4, TOKEN:5, ADMIN:6, ROLE:7, DIVISION:8, SUBDIVISION:9, SUBSTATION:10};
const USER_HEADERS = ['User ID','User Name','Password Hash','Active','Auth Token','Admin',
                      'Role','Division','Subdivision','Substation'];

const SERVER_VERSION = 'orai-server-6 (disconnection report)';

function doGet() {
  return json_({success:true, message:'Orai Fieldwork API is running', version:SERVER_VERSION});
}

/* Admin > Master Data asks this first, so it can say exactly what is missing. */
function serverInfo_(body) {
  if (!requireAdmin_(body)) return json_({success:false, code:'AUTH', message:'Admin access required.'});
  try {
    const f = getMasterFolder_();
    getMastersSheet_();
    getDailySheet_();
    return json_({success:true, version:SERVER_VERSION, drive_ok:true, folder:f.getName(), daily:dailySummary_()});
  } catch (e) {
    return json_({success:false, version:SERVER_VERSION, drive_ok:false,
      message:'Google Drive access is not allowed yet: open Apps Script, run setupOrai and allow the permissions. (' + e + ')'});
  }
}

function doPost(e) {
  try {
    const body = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    const a = body.action;

    if (a === 'login') return login_(body);
    if (a === 'check_auth') return checkAuth_(body);
    if (a === 'check_upload_status') return checkUploadStatus_(body);
    if (a === 'upload') return upload_(body);
    if (a === 'sync_upload_status') return syncUploadStatus_(body);

    if (a === 'master_lists') return masterLists_(body);
    if (a === 'list_users') return listUsers_(body);
    if (a === 'create_user') return createUser_(body);
    if (a === 'set_user_active') return setUserActive_(body);
    if (a === 'reset_password') return resetPassword_(body);
    if (a === 'dashboard') return dashboard_(body);
    if (a === 'dashboard_filters') return dashboard_filters_(body);
    if (a === 'export_report') return exportReport_(body);
    if (a === 'get_record') return getRecord_(body);
    if (a === 'update_record') return updateRecord_(body);

    if (a === 'server_info') return serverInfo_(body);
    if (a === 'master_check') return masterCheck_(body);
    if (a === 'master_download') return masterDownload_(body);
    if (a === 'master_download_part') return masterDownloadPart_(body);
    if (a === 'master_updates') return masterUpdatesDaily_(body);
    if (a === 'disconnection_report') return disconnectionReport_(body);
    if (a === 'daily_info') return dailyInfo_(body);
    if (a === 'daily_update_upload') return dailyUpdateUpload_(body);
    if (a === 'daily_update_get') return dailyUpdateGet_(body);
    if (a === 'master_list') return masterList_(body);
    if (a === 'master_upload_start') return masterUploadStart_(body);
    if (a === 'master_upload_chunk') return masterUploadChunk_(body);
    if (a === 'master_upload_finish') return masterUploadFinish_(body);
    if (a === 'master_set_active') return masterSetActive_(body);

    return json_({success:false, message:'This feature is not set up on the Orai server yet (' + a + ').'});
  } catch (err) {
    return json_({success:false, message:String(err)});
  }
}

/* ---------- ONE-TIME SETUP ---------- */

function setupOrai() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  getUsersSheet_();

  let h = ss.getSheetByName(HIERARCHY_SHEET);
  if (!h) {
    h = ss.insertSheet(HIERARCHY_SHEET);
    h.getRange(1,1,7,3).setValues([
      ['Division','Subdivision','Substation'],
      [ORAI_DIVISION,'Town-I (Kalpi Road)','Rahiya'],
      [ORAI_DIVISION,'Town-I (Kalpi Road)','Kalpi Road'],
      [ORAI_DIVISION,'Town-I (Kalpi Road)','Orai Kalpi Road'],
      [ORAI_DIVISION,'Town-II','Konch Naka'],
      [ORAI_DIVISION,'Town-II','Susheel Nagar'],
      [ORAI_DIVISION,'Town-II','Karmer Road']
    ]);
    h.setFrozenRows(1);
  }

  let r = ss.getSheetByName(ROLES_SHEET);
  if (!r) {
    r = ss.insertSheet(ROLES_SHEET);
    r.getRange(1,1,5,2).setValues([
      ['Role','Scope'],
      ['EE','DIVISION'],
      ['SDO','SUBDIVISION'],
      ['JE','SUBSTATION_MULTI'],
      ['Surveyor','SUBSTATION_SINGLE']
    ]);
    r.setFrozenRows(1);
  }

  const m = ss.getSheetByName(SHEET_NAME) || ss.insertSheet(SHEET_NAME);
  ensureHeaders_(m);

  getMastersSheet_();
  getMasterFolder_();

  Logger.log('Setup done (Users, Hierarchy, Roles, Masters, Drive folder).');
}

/* ---------- USERS ---------- */

function getUsersSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(USERS_SHEET);

  if (!sheet) {
    sheet = ss.insertSheet(USERS_SHEET);
    sheet.getRange(1,1,1,USER_HEADERS.length).setValues([USER_HEADERS]);
    if (!ADMIN_PASSWORD) throw new Error('Set ADMIN_PASSWORD at the top of the script before the first setupOrai.');
    sheet.getRange(2,1,1,USER_HEADERS.length).setValues([[
      ADMIN_USER_ID, ADMIN_NAME, sha256_(ADMIN_PASSWORD), 'YES', newToken_(), 'YES',
      'Admin', ORAI_DIVISION, '', ''
    ]]);
    sheet.setFrozenRows(1);
    sheet.hideColumns(U.HASH);
    sheet.hideColumns(U.TOKEN);
    return sheet;
  }

  if (sheet.getLastColumn() < USER_HEADERS.length) {
    sheet.getRange(1,1,1,USER_HEADERS.length).setValues([USER_HEADERS]);
  }
  return sheet;
}

function newToken_() {
  return Utilities.getUuid() + Utilities.getUuid();
}

function userFromRow_(row) {
  return {
    user_id: String(row[U.ID-1] || '').trim(),
    user_name: String(row[U.NAME-1] || '').trim(),
    is_admin: String(row[U.ADMIN-1] || '').trim().toUpperCase() === 'YES',
    role: String(row[U.ROLE-1] || '').trim(),
    division: String(row[U.DIVISION-1] || '').trim(),
    subdivision: String(row[U.SUBDIVISION-1] || '').trim(),
    substation: String(row[U.SUBSTATION-1] || '').trim()
  };
}

function login_(body) {
  const userId = String(body.user_id || '').trim();
  const password = String(body.password || '');
  if (!userId || !password) {
    return json_({success:false, message:'User ID and password are required.'});
  }

  const sheet = getUsersSheet_();
  const data = sheet.getDataRange().getValues();
  const hash = sha256_(password);

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (String(row[U.ID-1] || '').trim().toLowerCase() === userId.toLowerCase() &&
        String(row[U.ACTIVE-1] || '').trim().toUpperCase() === 'YES' &&
        String(row[U.HASH-1] || '').trim() === hash) {

      let token = String(row[U.TOKEN-1] || '').trim();
      if (!token) {
        token = newToken_();
        sheet.getRange(i + 1, U.TOKEN).setValue(token);
      }
      const u = userFromRow_(row);
      return json_(Object.assign({success:true, token:token}, u));
    }
  }
  return json_({success:false, message:'Invalid User ID or Password.'});
}

function checkAuth_(body) {
  const s = getUserByToken_(String(body.token || ''));
  if (!s) return json_({success:false, code:'AUTH', message:'Login invalid or user disabled.'});
  return json_(Object.assign({success:true}, s));
}

function getUserByToken_(token) {
  if (!token) return null;
  const data = getUsersSheet_().getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (String(row[U.ACTIVE-1] || '').trim().toUpperCase() === 'YES' &&
        String(row[U.TOKEN-1] || '').trim() === token) {
      return userFromRow_(row);
    }
  }
  return null;
}

function listUsers_(body) {
  if (!requireAdmin_(body)) return json_({success:false, code:'AUTH', message:'Admin access required.'});
  const data = getUsersSheet_().getDataRange().getValues();
  const users = [];
  for (let i = 1; i < data.length; i++) {
    const u = userFromRow_(data[i]);
    if (!u.user_id) continue;
    u.active = String(data[i][U.ACTIVE-1] || '');
    users.push(u);
  }
  return json_({success:true, users:users});
}

function createUser_(body) {
  if (!requireAdmin_(body)) return json_({success:false, code:'AUTH', message:'Admin access required.'});

  const userId = String(body.user_id || '').trim();
  const userName = String(body.user_name || '').trim();
  const password = String(body.password || '');
  const role = String(body.role || '').trim();
  const division = String(body.division || '').trim();
  const subdivision = String(body.subdivision || '').trim();
  const substations = Array.isArray(body.substations) ? body.substations.map(String) : [];

  if (!userId || !userName || !password) {
    return json_({success:false, message:'User ID, User Name and Password are required.'});
  }
  if (password.length < 6) {
    return json_({success:false, message:'Password must be at least 6 characters.'});
  }

  const sheet = getUsersSheet_();
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][U.ID-1] || '').trim().toLowerCase() === userId.toLowerCase()) {
      return json_({success:false, message:'User ID already exists.'});
    }
  }

  sheet.appendRow([userId, userName, sha256_(password), 'YES', newToken_(), 'NO',
                   role, division, subdivision, substations.join(', ')]);
  return json_({success:true, message:'User created.'});
}

function resetPassword_(body) {
  if (!requireAdmin_(body)) return json_({success:false, code:'AUTH', message:'Admin access required.'});

  const userId = String(body.user_id || '').trim();
  const pwd = String(body.new_password || '');
  if (pwd.length < 6) return json_({success:false, message:'Password must be at least 6 characters.'});

  const sheet = getUsersSheet_();
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][U.ID-1] || '').trim().toLowerCase() === userId.toLowerCase()) {
      sheet.getRange(i + 1, U.HASH).setValue(sha256_(pwd));
      sheet.getRange(i + 1, U.TOKEN).setValue(newToken_());   /* logs that user out everywhere */
      return json_({success:true, message:'Password reset for ' + userId + '.'});
    }
  }
  return json_({success:false, message:'User not found.'});
}

/* ---------- ORGANIZATION (for Admin > Create User) ---------- */

function masterLists_(body) {
  if (!requireAdmin_(body)) return json_({success:false, code:'AUTH', message:'Admin access required.'});

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const h = ss.getSheetByName(HIERARCHY_SHEET);
  const r = ss.getSheetByName(ROLES_SHEET);
  if (!h || !r) return json_({success:false, message:'Run setupOrai once in Apps Script first.'});

  const divisions = [], divisionMap = {}, subdivisionMap = {};
  h.getDataRange().getValues().slice(1).forEach(function(row) {
    const d = String(row[0] || '').trim(), s = String(row[1] || '').trim(), st = String(row[2] || '').trim();
    if (!d) return;
    if (!divisionMap[d]) { divisionMap[d] = {subdivisions:[], substations:[]}; divisions.push(d); }
    if (s && divisionMap[d].subdivisions.indexOf(s) === -1) divisionMap[d].subdivisions.push(s);
    if (st && divisionMap[d].substations.indexOf(st) === -1) divisionMap[d].substations.push(st);
    if (s) {
      const key = d + '||' + s;
      subdivisionMap[key] = subdivisionMap[key] || [];
      if (st && subdivisionMap[key].indexOf(st) === -1) subdivisionMap[key].push(st);
    }
  });

  const roles = [], roleScope = {};
  r.getDataRange().getValues().slice(1).forEach(function(row) {
    const role = String(row[0] || '').trim(), scope = String(row[1] || '').trim().toUpperCase();
    if (!role) return;
    roles.push(role);
    roleScope[role] = scope;
  });

  return json_({success:true, divisions:divisions, division_map:divisionMap,
                subdivision_map:subdivisionMap, roles:roles, role_scope:roleScope});
}


/* =========================================================
 * MASTER DATA (consumer master XLSX) — upload, assign, serve
 * =========================================================
 * Files are kept in the Drive folder "Orai Fieldwork Master Data".
 * The "Masters" sheet lists every upload; the row with Active = YES
 * for a scope is what phones download. Scope = ALL (every user) or a
 * subdivision name from the Hierarchy sheet (only that subdivision).
 * A phone that already has a master picks up a newer one by itself
 * ("Check for Master Data update" or the automatic check).
 */
const MASTERS_SHEET = 'Masters';
const MASTER_FOLDER_NAME = 'Orai Fieldwork Master Data';
const MASTER_HEADERS = ['Master File ID','File Name','File Size','Scope','Active',
                        'Uploaded At','Uploaded By','Consumers','Billing Month','Report Date','Notes','Parts','Base ID'];
const M = {ID:1, NAME:2, SIZE:3, SCOPE:4, ACTIVE:5, AT:6, BY:7, COUNT:8, BMONTH:9, RDATE:10, NOTES:11, PARTS:12, BASE:13};
/* A master is kept in Drive as its upload parts (Base64 text, ~4 MB each) inside one folder,
   and phones download it part by part. The server never has to hold the whole file, so a
   full-division master (30 MB+) works within Apps Script limits. */

function getMastersSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(MASTERS_SHEET);
  if (!sh) {
    sh = ss.insertSheet(MASTERS_SHEET);
    sh.getRange(1,1,1,MASTER_HEADERS.length).setValues([MASTER_HEADERS]);
    sh.setFrozenRows(1);
  } else if (sh.getLastColumn() < MASTER_HEADERS.length) {
    sh.getRange(1,1,1,MASTER_HEADERS.length).setValues([MASTER_HEADERS]);
  }
  return sh;
}

function getMasterFolder_() {
  const props = PropertiesService.getScriptProperties();
  const id = props.getProperty('MASTER_FOLDER_ID');
  if (id) {
    try { return DriveApp.getFolderById(id); } catch (e) {}
  }
  const it = DriveApp.getFoldersByName(MASTER_FOLDER_NAME);
  const folder = it.hasNext() ? it.next() : DriveApp.createFolder(MASTER_FOLDER_NAME);
  props.setProperty('MASTER_FOLDER_ID', folder.getId());
  return folder;
}

function masterRows_() {
  const data = getMastersSheet_().getDataRange().getValues();
  const rows = [];
  for (let i = 1; i < data.length; i++) {
    const r = data[i];
    if (!String(r[M.ID-1] || '').trim()) continue;
    rows.push({
      row: i + 1,
      master_file_id: String(r[M.ID-1]).trim(),
      file_name: String(r[M.NAME-1] || ''),
      file_size: Number(r[M.SIZE-1]) || 0,
      scope: String(r[M.SCOPE-1] || 'ALL').trim() || 'ALL',
      active: String(r[M.ACTIVE-1] || '').trim().toUpperCase() === 'YES',
      last_updated: r[M.AT-1] instanceof Date ? r[M.AT-1].toISOString() : String(r[M.AT-1] || ''),
      uploaded_by: String(r[M.BY-1] || ''),
      consumers: Number(r[M.COUNT-1]) || 0,
      billing_month: r[M.BMONTH-1] instanceof Date ? Utilities.formatDate(r[M.BMONTH-1], Session.getScriptTimeZone(), 'yyyy-MM')
                                                    : String(r[M.BMONTH-1] || '').replace(/^'/, ''),
      report_date: r[M.RDATE-1] instanceof Date ? Utilities.formatDate(r[M.RDATE-1], Session.getScriptTimeZone(), 'yyyy-MM-dd')
                                                : String(r[M.RDATE-1] || '').replace(/^'/, ''),
      base_id: String(r[M.BASE-1] || '').replace(/^'/, ''),
      notes: String(r[M.NOTES-1] || ''),
      parts: Number(r[M.PARTS-1]) || 0
    });
  }
  return rows;
}

/* The active master for this user: their own subdivision first, else ALL. */
function assignedMaster_(session) {
  const rows = masterRows_().filter(function(m) { return m.active; });
  const sub = String(session.subdivision || '').trim().toLowerCase();
  const own = sub ? rows.filter(function(m) { return m.scope.toLowerCase() === sub; }) : [];
  const all = rows.filter(function(m) { return m.scope.toUpperCase() === 'ALL'; });
  const pick = own.length ? own : all;
  return pick.length ? pick[pick.length - 1] : null;
}

function masterCheck_(body) {
  const s = getUserByToken_(String(body.token || ''));
  if (!s) return json_({success:false, code:'AUTH', message:'Login invalid or user disabled.'});
  const m = assignedMaster_(s);
  if (!m) {
    return json_({success:false, message:'No Master Data has been uploaded yet. Admin: Admin > Master Data > Upload. ' +
                  '(मास्टर डेटा अभी अपलोड नहीं हुआ है)'});
  }
  return json_({success:true, master_file_id:m.master_file_id, file_name:m.file_name,
                file_size:m.file_size, last_updated:m.last_updated, scope:m.scope, parts:m.parts,
                billing_month:m.billing_month, report_date:m.report_date});
}

function masterDownload_(body) {
  const s = getUserByToken_(String(body.token || ''));
  if (!s) return json_({success:false, code:'AUTH', message:'Login invalid or user disabled.'});

  const id = String(body.master_file_id || '').trim();
  /* Only files listed in the Masters sheet can be downloaded. */
  const m = masterRows_().filter(function(x) { return x.master_file_id === id; })[0];
  if (!m) return json_({success:false, message:'This Master Data file is not available.'});
  if (m.parts) return json_({success:false, message:'This master is stored in parts. Please reopen the app to update it.'});

  const bytes = DriveApp.getFileById(id).getBlob().getBytes();
  return json_({success:true, total_size:bytes.length, data:Utilities.base64Encode(bytes)});
}

function partName_(i) { return 'part_' + ('00000' + i).slice(-5) + '.b64'; }

/* One part of a parted master, as the Base64 text it was uploaded as. */
function masterDownloadPart_(body) {
  const s = getUserByToken_(String(body.token || ''));
  if (!s) return json_({success:false, code:'AUTH', message:'Login invalid or user disabled.'});
  const id = String(body.master_file_id || '').trim();
  const m = masterRows_().filter(function(x) { return x.master_file_id === id; })[0];
  if (!m || !m.parts) return json_({success:false, message:'This Master Data file is not available.'});
  const i = Number(body.index);
  if (!(i >= 0 && i < m.parts)) return json_({success:false, message:'Bad part number.'});
  const it = DriveApp.getFolderById(id).getFilesByName(partName_(i));
  if (!it.hasNext()) return json_({success:false, message:'Master part ' + (i + 1) + ' is missing on the server. Upload the master again.'});
  return json_({success:true, index:i, parts:m.parts, data:it.next().getBlob().getDataAsString()});
}

/* Phones ask this now and then; they compare file ids themselves. */
function masterUpdates_(body) {
  const s = getUserByToken_(String(body.token || ''));
  if (!s) return json_({success:false, code:'AUTH', message:'Login invalid or user disabled.'});
  return json_({success:true, available:!!assignedMaster_(s), master_changed:true});
}

/* ---------- Admin: list ---------- */
function masterList_(body) {
  if (!requireAdmin_(body)) return json_({success:false, code:'AUTH', message:'Admin access required.'});
  const scopes = ['ALL'];
  const h = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(HIERARCHY_SHEET);
  if (h) h.getDataRange().getValues().slice(1).forEach(function(r) {
    const sub = String(r[1] || '').trim();
    if (sub && scopes.indexOf(sub) === -1) scopes.push(sub);
  });
  const masters = masterRows_().reverse().slice(0, 20);
  return json_({success:true, scopes:scopes, masters:masters, daily:dailySummary_()});
}

/* ---------- Admin: upload in parts ----------
 * The app cuts the file into parts whose raw size is a multiple of 3 bytes,
 * so the Base64 parts can simply be joined back together.               */
function masterUploadStart_(body) {
  if (!requireAdmin_(body)) return json_({success:false, code:'AUTH', message:'Admin access required.'});
  const folder = getMasterFolder_().createFolder('_upload_' + Date.now());
  return json_({success:true, upload_id:folder.getId()});
}

function masterUploadChunk_(body) {
  if (!requireAdmin_(body)) return json_({success:false, code:'AUTH', message:'Admin access required.'});
  const folder = uploadFolder_(String(body.upload_id || ''));
  if (!folder) return json_({success:false, message:'Upload not found. Please start again.'});
  const idx = Number(body.index);
  if (!(idx >= 0)) return json_({success:false, message:'Bad part number.'});
  const name = partName_(idx);
  const old = folder.getFilesByName(name);
  while (old.hasNext()) old.next().setTrashed(true);       /* a retried part replaces the first try */
  folder.createFile(name, String(body.data || ''), MimeType.PLAIN_TEXT);
  return json_({success:true, index:idx});
}

function masterUploadFinish_(body) {
  const admin = requireAdmin_(body);
  if (!admin) return json_({success:false, code:'AUTH', message:'Admin access required.'});
  const folder = uploadFolder_(String(body.upload_id || ''));
  if (!folder) return json_({success:false, message:'Upload not found. Please start again.'});

  const parts = Number(body.parts) || 0;
  if (parts < 1) return json_({success:false, message:'No parts were uploaded.'});
  let size = 0;
  for (let i = 0; i < parts; i++) {
    const it = folder.getFilesByName(partName_(i));
    if (!it.hasNext()) return json_({success:false, message:'Part ' + (i + 1) + ' of ' + parts + ' is missing. Please upload again.'});
    const len = it.next().getSize();                     /* Base64 text length */
    size += len / 4 * 3;
  }
  /* Base64 padding makes the estimate up to 2 bytes larger per part. */
  const total = Number(body.total_size);
  if (Math.abs(size - total) > 2 * parts) {
    return json_({success:false, message:'File arrived incomplete (about ' + Math.round(size) + ' of ' + total + ' bytes). Please upload again.'});
  }

  const scope = String(body.scope || 'ALL').trim() || 'ALL';
  const stamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd_HHmm');
  const fileName = String(body.file_name || 'master.xlsx').replace(/[\\\/:*?"<>|]/g, '_');
  folder.setName(stamp + '_' + scope.replace(/[^A-Za-z0-9-]+/g, '_') + '_' + fileName);

  const sh = getMastersSheet_();
  const old = masterRows_().filter(function(m) { return m.scope.toLowerCase() === scope.toLowerCase(); });
  sh.appendRow([folder.getId(), fileName, total, scope, 'YES', new Date(), admin.user_id,
                Number(body.consumers) || '', "'" + String(body.billing_month || ''), "'" + String(body.report_date || ''),
                String(body.notes || ''), parts, "'" + String(body.base_id || '')]);

  /* The new master replaces the old one(s) of the same scope: their rows are removed
     and their Drive files go to the Drive trash (recoverable there for 30 days). */
  const removed = removeMasters_(old);
  cleanStaleUploads_(folder.getId());

  return json_({success:true, master_file_id:folder.getId(), removed:removed,
                message:'Master Data uploaded for ' + (scope === 'ALL' ? 'all users' : scope) +
                        ' (' + parts + ' parts).' + (removed ? ' Old master removed.' : '') +
                        ' Phones will load it on their next update check.'});
}

function trashDriveItem_(id) {
  try { DriveApp.getFolderById(id).setTrashed(true); return true; } catch (e) {}
  try { DriveApp.getFileById(id).setTrashed(true); return true; } catch (e) {}
  return false;
}

/* Delete rows bottom-up so row numbers stay right, and trash their Drive files. */
function removeMasters_(list) {
  if (!list.length) return 0;
  const sh = getMastersSheet_();
  list.slice().sort(function(a, b) { return b.row - a.row; }).forEach(function(m) {
    trashDriveItem_(m.master_file_id);
    sh.deleteRow(m.row);
  });
  /* daily updates built for those masters are of no use any more */
  const ids = list.map(function(m) { return m.master_file_id; });
  removeDailyRows_(dailyRows_().filter(function(d) { return ids.indexOf(d.master_file_id) !== -1; }));
  return list.length;
}

/* Upload folders left behind by failed / abandoned uploads (older than 6 hours). */
function cleanStaleUploads_(keepId) {
  const cutoff = Date.now() - 6 * 3600 * 1000;
  const it = getMasterFolder_().getFolders();
  while (it.hasNext()) {
    const f = it.next();
    if (f.getId() !== keepId && /^_upload_\d+$/.test(f.getName()) && Number(f.getName().slice(8)) < cutoff) {
      f.setTrashed(true);
    }
  }
}

/* Run once from Apps Script if you want to clear old masters now (keeps the active one per scope). */
function removeOldMastersNow() {
  const old = masterRows_().filter(function(m) { return !m.active; });
  Logger.log('Removed ' + removeMasters_(old) + ' old master(s).');
  cleanStaleUploads_('');
}

/* Admin can switch an older master back on (e.g. a wrong file was uploaded). */
function masterSetActive_(body) {
  if (!requireAdmin_(body)) return json_({success:false, code:'AUTH', message:'Admin access required.'});
  const id = String(body.master_file_id || '').trim();
  const rows = masterRows_();
  const m = rows.filter(function(x) { return x.master_file_id === id; })[0];
  if (!m) return json_({success:false, message:'Master not found.'});
  const sh = getMastersSheet_();
  rows.forEach(function(x) {
    if (x.scope.toLowerCase() === m.scope.toLowerCase()) sh.getRange(x.row, M.ACTIVE).setValue(x.master_file_id === id ? 'YES' : 'NO');
  });
  return json_({success:true, message:m.file_name + ' is now the active master for ' + m.scope + '.'});
}

function uploadFolder_(id) {
  if (!id) return null;
  try {
    const f = DriveApp.getFolderById(id);
    const parents = f.getParents();
    if (!parents.hasNext() || parents.next().getId() !== getMasterFolder_().getId()) return null;
    return f;
  } catch (e) { return null; }
}

/* =========================================================
 * DAILY BILLED / UNBILLED UPDATES
 * =========================================================
 * Every morning the PC job (Orai_Daily_Update.ps1) reads the newest
 * BILLED_ / UNBILLED_ CCB pair, builds a small gzip "change file"
 * (only bill date, MR source, units, reading, payments per consumer)
 * and uploads it here. Phones that have the active master download the
 * newest change file on their next sync and apply it on top of the
 * master, so Progress > Billing shows this month's billed / unbilled.
 * Only the newest few files are kept; older ones go to the Drive trash.
 */
const DAILY_SHEET = 'Daily Updates';
const DAILY_HEADERS = ['Seq','Tag','Report Date','Billing Month','Base ID','Master File ID','File ID','Size',
                       'Billed','Unbilled','Consumers','Uploaded At','Uploaded By'];
const D = {SEQ:1, TAG:2, RDATE:3, BMONTH:4, BASE:5, MASTER:6, FILE:7, SIZE:8, BILLED:9, UNBILLED:10, COUNT:11, AT:12, BY:13};
const DAILY_KEEP = 3;

function getDailySheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(DAILY_SHEET);
  if (!sh) {
    sh = ss.insertSheet(DAILY_SHEET);
    sh.getRange(1,1,1,DAILY_HEADERS.length).setValues([DAILY_HEADERS]);
    sh.setFrozenRows(1);
  }
  return sh;
}

function getDailyFolder_() {
  const parent = getMasterFolder_();
  const it = parent.getFoldersByName('Daily Updates');
  return it.hasNext() ? it.next() : parent.createFolder('Daily Updates');
}

function txt_(v) {                          /* keep "2026-11" as text, not a date */
  if (v instanceof Date) return Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  return String(v == null ? '' : v).replace(/^'/, '');
}

function dailyRows_() {
  const data = getDailySheet_().getDataRange().getValues();
  const rows = [];
  for (let i = 1; i < data.length; i++) {
    const r = data[i];
    if (!r[D.SEQ-1]) continue;
    rows.push({row:i + 1, seq:Number(r[D.SEQ-1]), tag:txt_(r[D.TAG-1]), report_date:txt_(r[D.RDATE-1]),
      billing_month:txt_(r[D.BMONTH-1]).slice(0, 7), base_id:txt_(r[D.BASE-1]), master_file_id:txt_(r[D.MASTER-1]),
      file_id:txt_(r[D.FILE-1]), size:Number(r[D.SIZE-1]) || 0, billed:Number(r[D.BILLED-1]) || 0,
      unbilled:Number(r[D.UNBILLED-1]) || 0, consumers:Number(r[D.COUNT-1]) || 0,
      at:r[D.AT-1] instanceof Date ? r[D.AT-1].toISOString() : txt_(r[D.AT-1]), by:txt_(r[D.BY-1])});
  }
  return rows;
}

function latestDailyFor_(masterId) {
  const rows = dailyRows_().filter(function(d) { return d.master_file_id === masterId; });
  rows.sort(function(a, b) { return b.seq - a.seq; });
  return rows[0] || null;
}

/* PC job asks this first: which master the update must be built for. */
function dailyInfo_(body) {
  if (!requireAdmin_(body)) return json_({success:false, code:'AUTH', message:'Admin access required.'});
  const masters = masterRows_().filter(function(m) { return m.active; }).map(function(m) {
    const d = latestDailyFor_(m.master_file_id);
    return {scope:m.scope, master_file_id:m.master_file_id, base_id:m.base_id, file_name:m.file_name,
            last_tag:d ? d.tag : '', last_report_date:d ? d.report_date : ''};
  });
  return json_({success:true, version:SERVER_VERSION, masters:masters});
}

function dailyUpdateUpload_(body) {
  const admin = requireAdmin_(body);
  if (!admin) return json_({success:false, code:'AUTH', message:'Admin access required.'});
  const masterId = String(body.master_file_id || '');
  const m = masterRows_().filter(function(x) { return x.active && x.master_file_id === masterId; })[0];
  if (!m) return json_({success:false, message:'The master changed since this update was built. Run the update again.'});
  const data = String(body.data || '');
  if (!data) return json_({success:false, message:'Empty update.'});

  const tag = String(body.tag || '');
  const file = getDailyFolder_().createFile('daily_' + tag + '_' + Date.now() + '.b64', data, MimeType.PLAIN_TEXT);
  const sh = getDailySheet_();
  const all = dailyRows_();
  const seq = all.reduce(function(a, d) { return Math.max(a, d.seq); }, 0) + 1;
  sh.appendRow([seq, "'" + tag, "'" + String(body.report_date || ''), "'" + String(body.billing_month || ''),
                "'" + m.base_id, m.master_file_id, file.getId(), Number(body.size) || 0,
                Number(body.billed) || 0, Number(body.unbilled) || 0, Number(body.consumers) || 0, new Date(), admin.user_id]);

  /* keep only the newest DAILY_KEEP updates (older files to the Drive trash) */
  const old = dailyRows_().sort(function(a, b) { return b.seq - a.seq; }).slice(DAILY_KEEP);
  removeDailyRows_(old);

  return json_({success:true, seq:seq, message:'Daily update ' + tag + ' saved (billed ' + (Number(body.billed) || 0) +
                ', unbilled ' + (Number(body.unbilled) || 0) + '). Phones load it on their next sync.'});
}

function removeDailyRows_(list) {
  if (!list.length) return 0;
  const sh = getDailySheet_();
  list.slice().sort(function(a, b) { return b.row - a.row; }).forEach(function(d) {
    try { DriveApp.getFileById(d.file_id).setTrashed(true); } catch (e) {}
    sh.deleteRow(d.row);
  });
  return list.length;
}

/* Phone: the newest change file for its master (Base64 of the gzip). */
function dailyUpdateGet_(body) {
  const s = getUserByToken_(String(body.token || ''));
  if (!s) return json_({success:false, code:'AUTH', message:'Login invalid or user disabled.'});
  const m = assignedMaster_(s);
  const d = m ? dailyRows_().filter(function(x) { return x.seq === Number(body.seq) && x.master_file_id === m.master_file_id; })[0] : null;
  if (!d) return json_({success:false, message:'This daily update is no longer available.'});
  return json_({success:true, seq:d.seq, data:DriveApp.getFileById(d.file_id).getBlob().getDataAsString()});
}

/* Replaces the old masterUpdates_: tells the phone about a new master or the newest daily update. */
function masterUpdatesDaily_(body) {
  const s = getUserByToken_(String(body.token || ''));
  if (!s) return json_({success:false, code:'AUTH', message:'Login invalid or user disabled.'});
  const m = assignedMaster_(s);
  if (!m) return json_({success:true, available:false});
  const phoneFile = String(body.master_file_id || '');
  if (phoneFile ? phoneFile !== m.master_file_id : String(body.base_id || '') !== m.base_id) {
    return json_({success:true, available:true, master_changed:true});
  }
  const d = latestDailyFor_(m.master_file_id);
  if (!d) return json_({success:true, available:true, up_to_date:true});
  if (d.seq > (Number(body.seq) || 0)) {
    return json_({success:true, available:true, report_date:d.report_date, billing_month:d.billing_month,
                  downloads:[{seq:d.seq, from_seq:0, server:true, size:d.size}]});
  }
  return json_({success:true, available:true, up_to_date:true, report_date:d.report_date});
}

function dailySummary_() {
  const rows = dailyRows_().sort(function(a, b) { return b.seq - a.seq; });
  return rows.map(function(d) {
    return {seq:d.seq, tag:d.tag, report_date:d.report_date, billing_month:d.billing_month, billed:d.billed,
            unbilled:d.unbilled, consumers:d.consumers, size:d.size, at:d.at, master_file_id:d.master_file_id};
  });
}

/* =========================================================
 * DISCONNECTION / RECHECK / PHONE CALL ENTRIES + DISCONNECTION REPORT
 * =========================================================
 * Field entries arrive through the same 'upload' action as surveys,
 * marked with activity_type. Each type has its own sheet; one row per
 * visit, never merged with surveys. 'disconnection_report' feeds the
 * Dashboard PDF reports (complete report, time given / paid / disconnected
 * but not yet paid).
 */
const ACTIVITY = {
  DISCONNECTION: {sheet:'Disconnections', id:'disconnection_id'},
  RECHECK:       {sheet:'Rechecks',       id:'recheck_id'},
  PHONE_CALLING: {sheet:'Phone Calls',    id:'calling_id'}
};
const ACT_HEADERS = ['Entry ID','Visit Date','Created At','User ID','User Name','User Subdivision','User Substation',
  'Account ID','Consumer Name','Father Name','Address','Supply Type','Load','SDO Code','Village','Status',
  'Payment Mode','Paid Amount','Committed Date','Outstanding','Meter Status','Current Reading','Mobile',
  'Mobile Refused','House Condition','Remarks','Latitude','Longitude','GPS Accuracy','Data Up To','Uploaded At','Full Entry (JSON)'];
const A = {}; ACT_HEADERS.forEach(function(h, i) { A[h] = i; });

function activityType_(r) {
  const t = String((r && r.activity_type) || '').trim().toUpperCase();
  return ACTIVITY[t] ? t : '';
}

function getActivitySheet_(type) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const name = ACTIVITY[type].sheet;
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.getRange(1,1,1,ACT_HEADERS.length).setValues([ACT_HEADERS]);
    sh.setFrozenRows(1);
  }
  return sh;
}

function istDate_(iso) {
  const d = new Date(iso);
  if (isNaN(d)) return '';
  return Utilities.formatDate(d, 'Asia/Kolkata', 'yyyy-MM-dd');
}

function activityIds_(type) {
  const sh = getActivitySheet_(type);
  const n = sh.getLastRow() - 1;
  const ids = new Set();
  if (n > 0) sh.getRange(2, 1, n, 1).getValues().forEach(function(r) { if (r[0]) ids.add(String(r[0]).trim()); });
  return ids;
}

/* Save activity entries; returns {uploaded, duplicates}. */
function uploadActivities_(session, records) {
  const out = {uploaded:0, duplicates:0};
  const byType = {};
  records.forEach(function(r) { const t = activityType_(r); (byType[t] = byType[t] || []).push(r); });
  const users = getUsersSheet_().getDataRange().getValues();
  let sub = '', subst = '';
  for (let i = 1; i < users.length; i++) {
    if (String(users[i][U.ID-1]).trim().toLowerCase() === String(session.user_id).toLowerCase()) {
      sub = String(users[i][U.SUBDIVISION-1] || ''); subst = String(users[i][U.SUBSTATION-1] || ''); break;
    }
  }
  Object.keys(byType).forEach(function(type) {
    const sh = getActivitySheet_(type);
    const ids = activityIds_(type);
    const rows = [];
    byType[type].forEach(function(r) {
      const id = String(r[ACTIVITY[type].id] || '').trim();
      if (!id) return;
      if (ids.has(id)) { out.duplicates++; return; }
      ids.add(id);
      let json = JSON.stringify(r);
      if (json.length > 45000) json = json.slice(0, 45000);
      const status = r.disconnection_status || r.recheck_status || r.call_status || r.calling_status || r.status || '';
      const row = new Array(ACT_HEADERS.length).fill('');
      const put = function(h, v) { row[A[h]] = v == null ? '' : v; };
      put('Entry ID', id); put('Visit Date', istDate_(r.created_at) || istDate_(new Date().toISOString()));
      put('Created At', r.created_at); put('User ID', session.user_id); put('User Name', session.user_name || r.user_name);
      put('User Subdivision', sub); put('User Substation', subst);
      put('Account ID', "'" + String(r.account_id || '')); put('Consumer Name', r.consumer_name); put('Father Name', r.father_name);
      put('Address', r.address); put('Supply Type', "'" + String(r.supply_type || '')); put('Load', r.load);
      put('SDO Code', r.sdo_code); put('Village', r.feeded_village_name || r.village); put('Status', status);
      put('Payment Mode', r.payment_mode); put('Paid Amount', r.paid_amount); put('Committed Date', r.committed_payment_date);
      put('Outstanding', r.outstanding); put('Meter Status', r.meter_status); put('Current Reading', r.current_reading);
      put('Mobile', "'" + String(r.mobile_number || '')); put('Mobile Refused', r.mobile_refused);
      put('House Condition', r.house_condition); put('Remarks', r.remarks); put('Latitude', r.latitude);
      put('Longitude', r.longitude); put('GPS Accuracy', r.gps_accuracy); put('Data Up To', r.master_data_up_to);
      put('Uploaded At', new Date()); put('Full Entry (JSON)', json);
      rows.push(row);
    });
    if (rows.length) sh.getRange(sh.getLastRow() + 1, 1, rows.length, ACT_HEADERS.length).setValues(rows);
    out.uploaded += rows.length;
  });
  return out;
}

/* Wraps the survey upload: activity entries go to their own sheets. */
function upload_(body) {
  const session = getUserByToken_(String(body.token || ''));
  if (!session) return json_({success:false, code:'AUTH', message:'Login invalid or user disabled. Please login again.'});
  const all = body.records || [];
  const acts = all.filter(function(r) { return activityType_(r); });
  const surveys = all.filter(function(r) { return !activityType_(r); });
  let uploaded = 0, duplicates = 0;
  if (acts.length) {
    const lock = LockService.getScriptLock();
    lock.waitLock(30000);
    try { const a = uploadActivities_(session, acts); uploaded += a.uploaded; duplicates += a.duplicates; }
    finally { lock.releaseLock(); }
  }
  if (surveys.length) {
    const s = unwrap_(uploadSurveys_(Object.assign({}, body, {records:surveys})));
    if (!s.success) return json_(s);
    uploaded += Number(s.uploaded) || 0; duplicates += Number(s.duplicates) || 0;
  }
  return json_({success:true, uploaded:uploaded, duplicates:duplicates});
}

function unwrap_(out) {
  return out && typeof out.getContent === 'function' ? JSON.parse(out.getContent()) : out;
}

/* Wraps the survey status check: also reports activity entries already on the server. */
function syncUploadStatus_(body) {
  const session = getUserByToken_(String(body.token || ''));
  if (!session) return json_({success:false, code:'AUTH', message:'Login invalid or user disabled.'});
  const all = body.records || [];
  const acts = all.filter(function(r) { return activityType_(r); });
  const surveys = all.filter(function(r) { return !activityType_(r); });
  const keys = [];
  const idCache = {};
  acts.forEach(function(r) {
    const t = activityType_(r);
    const id = String(r[ACTIVITY[t].id] || '').trim();
    if (!idCache[t]) idCache[t] = activityIds_(t);
    if (id && idCache[t].has(id)) keys.push(t + '|' + id);
  });
  if (surveys.length) {
    const s = unwrap_(syncSurveyStatus_(Object.assign({}, body, {records:surveys})));
    if (!s.success) return json_(s);
    (s.present_keys || []).forEach(function(k) { keys.push(k); });
  }
  return json_({success:true, present_keys:keys});
}

/* Dashboard PDF: disconnection entries between two dates (yyyy-MM-dd, IST visit date). */
function disconnectionReport_(body) {
  const s = getUserByToken_(String(body.token || ''));
  if (!s) return json_({success:false, code:'AUTH', message:'Login invalid or user disabled.'});
  const from = String(body.from_date || ''), to = String(body.to_date || '');
  if (!/^\d{4}-\d{2}-\d{2}$/.test(from) || !/^\d{4}-\d{2}-\d{2}$/.test(to)) {
    return json_({success:false, message:'Choose the report dates.'});
  }
  const scope = (function() {
    if (s.is_admin) return 'ALL';
    const r = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ROLES_SHEET);
    let sc = '';
    if (r) r.getDataRange().getValues().slice(1).forEach(function(x) {
      if (String(x[0]).trim().toLowerCase() === String(s.role || '').toLowerCase()) sc = String(x[1]).trim().toUpperCase();
    });
    return sc === 'DIVISION' ? 'ALL' : sc === 'SUBDIVISION' ? 'SUBDIVISION' : 'OWN';
  })();
  const subs = String(body.substation || '').split('||').map(function(x) { return x.trim().toLowerCase(); }).filter(String);

  const data = getActivitySheet_('DISCONNECTION').getDataRange().getValues();
  const rows = [];
  for (let i = 1; i < data.length; i++) {
    const r = data[i];
    let date = r[A['Visit Date']];
    date = date instanceof Date ? Utilities.formatDate(date, 'Asia/Kolkata', 'yyyy-MM-dd') : String(date || '').slice(0, 10);
    if (!date || date < from || date > to) continue;
    const userId = String(r[A['User ID']] || '');
    if (scope === 'OWN' && userId.toLowerCase() !== String(s.user_id).toLowerCase()) continue;
    if (scope === 'SUBDIVISION' && String(r[A['User Subdivision']] || '').toLowerCase() !== String(s.subdivision || '').toLowerCase()) continue;
    const userSubst = String(r[A['User Substation']] || '');
    if (subs.length) {
      const mine = userSubst.split(',').map(function(x) { return x.trim().toLowerCase(); });
      if (!mine.some(function(x) { return subs.indexOf(x) !== -1; })) continue;
    }
    const v = function(h) { const x = r[A[h]]; return x instanceof Date ? Utilities.formatDate(x, 'Asia/Kolkata', 'yyyy-MM-dd') : String(x == null ? '' : x).replace(/^'/, ''); };
    rows.push({
      id:v('Entry ID'), date:date, account_id:v('Account ID'), consumer_name:v('Consumer Name'), father_name:v('Father Name'),
      address:v('Address'), supply_type:v('Supply Type'), load:v('Load'), mobile:v('Mobile'), mobile_refused:v('Mobile Refused'),
      latitude:v('Latitude'), longitude:v('Longitude'), outstanding:v('Outstanding'), visit_day_outstanding:'',
      phone_data_up_to:v('Data Up To'), status:String(v('Status')).toUpperCase(), committed_date:v('Committed Date'),
      payment_mode:v('Payment Mode'), paid_amount:v('Paid Amount'), user_id:userId, user_name:v('User Name'),
      user_substation:userSubst, meter_status:v('Meter Status'), remarks:v('Remarks')
    });
  }
  return json_({success:true, rows:rows, scope:scope});
}

function checkUploadStatus_(body) {
  const session = getUserByToken_(String(body.token || ''));
  if (!session) return json_({success:false, message:'Login invalid or user disabled.'});

  const ss=SpreadsheetApp.getActiveSpreadsheet();
  const sheet=ss.getSheetByName(SHEET_NAME);
  if (!sheet || sheet.getLastRow() < 2) return json_({success:true,uploaded_keys:[]});

  const values=sheet.getDataRange().getValues();
  const uploadedKeys=new Set();

  for(let i=1;i<values.length;i++){
    const user=String(values[i][0]||'').trim().toLowerCase();
    const account=String(values[i][2]||'').trim();
    const surveyType=String(values[i][15]||'EXISTING CONSUMER').trim().toUpperCase();
    const surveyId=String(values[i][18]||'').trim();

    if(surveyType === 'CONNECTION NOT IN DATABASE'){
      if(user && surveyId) uploadedKeys.add(user+'|SURVEY|'+surveyId);
    }else if(user && account){
      uploadedKeys.add(user+'|ACCOUNT|'+account);
    }
  }

  const found=[];
  for(const r of (body.records||[])){
    const user=session.user_id.trim().toLowerCase();
    const key=(String(r.survey_type||'').toUpperCase()==='CONNECTION NOT IN DATABASE')
      ? user+'|SURVEY|'+String(r.survey_id||'').trim()
      : user+'|ACCOUNT|'+String(r.account_id||'').trim();
    if(uploadedKeys.has(key)) found.push(key);
  }
  return json_({success:true,uploaded_keys:found});
}

function uploadSurveys_(body) {
  const session = getUserByToken_(String(body.token || ''));
  if (!session) {
    return json_({success:false, code:'AUTH', message:'Login invalid or user disabled. Please login again.'});
  }

  const records=body.records||[];
  if(!records.length) return json_({success:true,uploaded:0,duplicates:0});

  const ss=SpreadsheetApp.getActiveSpreadsheet();
  let sheet=ss.getSheetByName(SHEET_NAME);
  if(!sheet) sheet=ss.insertSheet(SHEET_NAME);
  ensureHeaders_(sheet);

  const lock=LockService.getScriptLock();
  lock.waitLock(30000);

  try{
    const values=sheet.getDataRange().getValues();
    const existingKeys=new Set();

    // IMPORTANT: this key is only for preventing the exact same user's
    // already-uploaded survey from being uploaded again. Account IDs from
    // different users are still allowed.
    for(let i=1;i<values.length;i++){
      const user=String(values[i][0]||'').trim().toLowerCase();
      const account=String(values[i][2]||'').trim();
      const type=String(values[i][15]||'EXISTING CONSUMER').trim().toUpperCase();
      const surveyId=String(values[i][18]||'').trim();
      if(type==='CONNECTION NOT IN DATABASE'){
        if(user && surveyId) existingKeys.add(user+'|SURVEY|'+surveyId);
      }else if(user && account){
        existingKeys.add(user+'|ACCOUNT|'+account);
      }
    }

    const rows=[];
    let duplicates=0;

    for(const r of records){
      const type=String(r.survey_type||'EXISTING CONSUMER').trim().toUpperCase();
      const user=String(session.user_id||'').trim().toLowerCase();
      const account=String(r.account_id||'').trim();
      const surveyId=String(r.survey_id||'').trim();

      const key=type==='CONNECTION NOT IN DATABASE'
        ? user+'|SURVEY|'+surveyId
        : user+'|ACCOUNT|'+account;

      if(key && existingKeys.has(key)){
        duplicates++;
        continue;
      }

      rows.push([
        session.user_id,
        session.user_name,
        account,
        r.consumer_name||'',
        r.father_name||'',
        r.address||'',
        r.supply_type||'',
        r.load||'',
        r.sdo_code||'',
        r.village||'',
        r.meter_condition||'',
        r.ac_installed||'',
        r.mobile_number||'',
        r.created_at||'',
        '', // Local Record ID no longer used as server identity
        r.survey_type||'EXISTING CONSUMER',
        r.survey_status||'CONSUMER FOUND',
        r.connection_type||'',
        surveyId,
        r.house_locked||'NO',
        r.mobile_refused||'NO',
        r.meter_installed||'',
        r.meter_number||'',
        r.current_reading||'',
        r.outstanding||'',
        r.connected_load||'',
        r.remarks||''
      ]);

      if(key) existingKeys.add(key);
    }

    if(rows.length){
      sheet.getRange(sheet.getLastRow()+1,1,rows.length,rows[0].length).setValues(rows);
    }

    return json_({success:true,uploaded:rows.length,duplicates:duplicates});
  }finally{
    lock.releaseLock();
  }
}


function syncSurveyStatus_(body) {
  const session = getUserByToken_(String(body.token || ''));
  if (!session) {
    return json_({success:false, code:'AUTH', message:'Login invalid or user disabled.'});
  }

  const requested = Array.isArray(body.records) ? body.records : [];
  if (!requested.length) return json_({success:true, present_keys:[]});

  const ss=SpreadsheetApp.getActiveSpreadsheet();
  const sheet=ss.getSheetByName(SHEET_NAME);

  if (!sheet || sheet.getLastRow() < 2) {
    return json_({success:true, present_keys:[]});
  }

  ensureHeaders_(sheet);

  const values=sheet.getDataRange().getValues();
  const presentKeys=new Set();

  // Sheet columns:
  // 0 User ID, 2 Account ID, 15 Survey Type, 18 Survey ID.
  for(let i=1;i<values.length;i++){
    const user=String(values[i][0]||'').trim().toLowerCase();
    const account=String(values[i][2]||'').trim();
    const type=String(values[i][15]||'EXISTING CONSUMER').trim().toUpperCase();
    const surveyId=String(values[i][18]||'').trim();

    if(type === 'CONNECTION NOT IN DATABASE'){
      if(user && surveyId){
        presentKeys.add(user+'|SURVEY|'+surveyId);
      }
    }else if(user && account){
      presentKeys.add(user+'|ACCOUNT|'+account);
    }
  }

  const found=[];
  for(const r of requested){
    const user=String(r.user_id || '').trim().toLowerCase();
    const account=String(r.account_id || '').trim();
    const type=String(r.survey_type || 'EXISTING CONSUMER').trim().toUpperCase();
    const surveyId=String(r.survey_id || '').trim();

    const key=(type === 'CONNECTION NOT IN DATABASE')
      ? user+'|SURVEY|'+surveyId
      : user+'|ACCOUNT|'+account;

    if(user && ((type === 'CONNECTION NOT IN DATABASE' && surveyId) ||
                (type !== 'CONNECTION NOT IN DATABASE' && account)) &&
       presentKeys.has(key)){
      found.push(key);
    }
  }

  return json_({success:true, present_keys:found});
}

function setUserActive_(body) {
  const admin = requireAdmin_(body);
  if (!admin) return json_({success:false, code:'AUTH', message:'Admin access required.'});

  const userId = String(body.user_id || '').trim();
  const active = String(body.active || '').toUpperCase() === 'YES' ? 'YES' : 'NO';

  if (userId.toLowerCase() === admin.user_id.toLowerCase() && active === 'NO') {
    return json_({success:false, message:'You cannot disable your own admin account.'});
  }

  const sheet = getUsersSheet_();
  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0] || '').trim().toLowerCase() === userId.toLowerCase()) {
      sheet.getRange(i + 1, 4).setValue(active);
      return json_({success:true});
    }
  }

  return json_({success:false, message:'User not found.'});
}

function requireAdmin_(body) {
  const session = getUserByToken_(String(body.token || ''));
  if (!session || !session.is_admin) return null;
  return session;
}


function dashboard_(body) {
  const admin = requireAdmin_(body);
  if (!admin) return json_({success:false, code:'AUTH', message:'Admin access required.'});

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  if (!sheet || sheet.getLastRow() < 2) {
    return json_({
      success:true,total:0,user_count:0,village_count:0,
      user_wise:[],village_wise:[],meter_condition:[],ac_installed:[]
    });
  }

  const fromDate = String(body.from_date || '').trim();
  const toDate = String(body.to_date || '').trim();
  const filterUser = String(body.user_id || '').trim();
  const filterVillage = String(body.village || '').trim();

  const values=sheet.getDataRange().getValues();
  const users={}, villages={}, meters={}, acs={};
  let total=0;

  for(let i=1;i<values.length;i++){
    const rowUserId=String(values[i][0]||'').trim();
    const rowUserName=String(values[i][1]||'').trim();
    const rowVillage=String(values[i][9]||'').trim();
    const rowMeter=String(values[i][10]||'').trim();
    const rowAc=String(values[i][11]||'').trim();
    const rawCreated=values[i][13];

    if (filterUser && rowUserId !== filterUser) continue;
    if (filterVillage && rowVillage !== filterVillage) continue;

    if (fromDate || toDate) {
      const d = rawCreated instanceof Date ? rawCreated : new Date(String(rawCreated||''));
      if (isNaN(d.getTime())) continue;
      const day = Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd');
      if (fromDate && day < fromDate) continue;
      if (toDate && day > toDate) continue;
    }

    const user = rowUserName || rowUserId || 'Unknown';
    const village = rowVillage || 'Unknown';
    const meter = rowMeter || 'Unknown';
    const ac = rowAc || 'Unknown';

    users[user]=(users[user]||0)+1;
    villages[village]=(villages[village]||0)+1;
    meters[meter]=(meters[meter]||0)+1;
    acs[ac]=(acs[ac]||0)+1;
    total++;
  }

  const toList_=obj=>Object.keys(obj)
    .map(k=>({name:k,count:obj[k]}))
    .sort((a,b)=>b.count-a.count);

  return json_({
    success:true,
    total:total,
    user_count:Object.keys(users).length,
    village_count:Object.keys(villages).length,
    user_wise:toList_(users),
    village_wise:toList_(villages),
    meter_condition:toList_(meters),
    ac_installed:toList_(acs)
  });
}

function dashboard_filters_(body) {
  const admin = requireAdmin_(body);
  if (!admin) return json_({success:false, code:'AUTH', message:'Admin access required.'});

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  const userMap = {};
  const villages = {};

  if (sheet && sheet.getLastRow() >= 2) {
    const values=sheet.getDataRange().getValues();
    for(let i=1;i<values.length;i++){
      const userId=String(values[i][0]||'').trim();
      const userName=String(values[i][1]||'').trim();
      const village=String(values[i][9]||'').trim();
      if(userId) userMap[userId]=userName || userId;
      if(village) villages[village]=true;
    }
  }

  const users=Object.keys(userMap)
    .map(id=>({user_id:id,user_name:userMap[id]}))
    .sort((a,b)=>a.user_name.localeCompare(b.user_name));

  return json_({
    success:true,
    users:users,
    villages:Object.keys(villages).sort((a,b)=>a.localeCompare(b))
  });
}


function exportReport_(body) {
  const admin = requireAdmin_(body);
  if (!admin) return json_({success:false, code:'AUTH', message:'Admin access required.'});

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  if (!sheet || sheet.getLastRow() < 2) {
    return json_({success:true, rows:[]});
  }

  const fromDate = String(body.from_date || '').trim();
  const toDate = String(body.to_date || '').trim();
  const filterUser = String(body.user_id || '').trim();
  const filterVillage = String(body.village || '').trim();

  const values = sheet.getDataRange().getValues();
  const rows = [];

  for (let i=1; i<values.length; i++) {
    const rowUserId=String(values[i][0]||'').trim();
    const rowUserName=String(values[i][1]||'').trim();
    const rowVillage=String(values[i][9]||'').trim();
    const rawCreated=values[i][13];

    if (filterUser && rowUserId !== filterUser) continue;
    if (filterVillage && rowVillage !== filterVillage) continue;

    if (fromDate || toDate) {
      const d = rawCreated instanceof Date ? rawCreated : new Date(String(rawCreated||''));
      if (isNaN(d.getTime())) continue;
      const day = Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd');
      if (fromDate && day < fromDate) continue;
      if (toDate && day > toDate) continue;
    }

    rows.push({
      user_id:rowUserId,
      user_name:rowUserName,
      account_id:String(values[i][2]||''),
      consumer_name:String(values[i][3]||''),
      father_name:String(values[i][4]||''),
      address:String(values[i][5]||''),
      supply_type:String(values[i][6]||''),
      load:String(values[i][7]||''),
      sdo_code:String(values[i][8]||''),
      village:rowVillage,
      meter_condition:String(values[i][10]||''),
      ac_installed:String(values[i][11]||''),
      mobile_number:String(values[i][12]||''),
      created_at: rawCreated instanceof Date
        ? Utilities.formatDate(rawCreated, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss')
        : String(rawCreated||''),
      survey_type:String(values[i][15]||''),
      survey_status:String(values[i][16]||''),
      connection_type:String(values[i][17]||''),
      survey_id:String(values[i][18]||''),
      house_locked:String(values[i][19]||''),
      mobile_refused:String(values[i][20]||''),
      meter_installed:String(values[i][21]||''),
      meter_number:String(values[i][22]||''),
      current_reading:String(values[i][23]||''),
      outstanding:String(values[i][24]||''),
      connected_load:String(values[i][25]||''),
      remarks:String(values[i][26]||'')
    });
  }

  return json_({success:true, rows:rows});
}


function getRecord_(body) {
  const admin = requireAdmin_(body);
  if (!admin) return json_({success:false, code:'AUTH', message:'Admin access required.'});

  const accountId=String(body.account_id||'').trim();
  if (!accountId) return json_({success:false,message:'Account ID required.'});

  const sheet=SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  if (!sheet || sheet.getLastRow()<2) {
    return json_({success:true,found:false});
  }

  const values=sheet.getDataRange().getValues();

  for(let i=1;i<values.length;i++){
    if(String(values[i][2]||'').trim()===accountId){
      return json_({
        success:true,
        found:true,
        record:rowToRecord_(values[i])
      });
    }
  }

  return json_({success:true,found:false});
}

function updateRecord_(body) {
  const admin = requireAdmin_(body);
  if (!admin) return json_({success:false, code:'AUTH', message:'Admin access required.'});

  const accountId=String(body.account_id||'').trim();
  const village=String(body.village||'').trim();
  const meter=String(body.meter_condition||'').trim().toUpperCase();
  const ac=String(body.ac_installed||'').trim().toUpperCase();
  const mobile=String(body.mobile_number||'').replace(/\D/g,'');

  if (!accountId) return json_({success:false,message:'Account ID required.'});
  if (!village) return json_({success:false,message:'Village is required.'});
  if (!['OK','DAMAGED','NOT INSTALLED'].includes(meter)) {
    return json_({success:false,message:'Invalid meter condition.'});
  }
  if (!['YES','NO'].includes(ac)) {
    return json_({success:false,message:'Invalid AC Installed value.'});
  }
  if (!/^\d{10}$/.test(mobile)) {
    return json_({success:false,message:'Invalid mobile number.'});
  }

  const sheet=SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  if (!sheet || sheet.getLastRow()<2) {
    return json_({success:false,message:'Record not found.'});
  }

  const values=sheet.getDataRange().getValues();

  for(let i=1;i<values.length;i++){
    if(String(values[i][2]||'').trim()===accountId){
      // V12/V13 columns:
      // J Village, K Meter Condition, L AC Installed, M Mobile Number
      sheet.getRange(i+1,10).setValue(village);
      sheet.getRange(i+1,11).setValue(meter);
      sheet.getRange(i+1,12).setValue(ac);
      sheet.getRange(i+1,13).setValue(mobile);

      // Audit columns are added at the end without changing existing data.
      const auditCols=ensureCorrectionAuditColumns_(sheet);
      const now=new Date();
      sheet.getRange(i+1,auditCols.updatedBy).setValue(admin.user_id);
      sheet.getRange(i+1,auditCols.updatedName).setValue(admin.user_name);
      sheet.getRange(i+1,auditCols.updatedAt).setValue(now);

      const updatedValues=sheet.getRange(i+1,1,1,sheet.getLastColumn()).getValues()[0];

      return json_({
        success:true,
        updated_by:admin.user_name || admin.user_id,
        updated_at:Utilities.formatDate(now,Session.getScriptTimeZone(),'yyyy-MM-dd HH:mm:ss'),
        record:rowToRecord_(updatedValues)
      });
    }
  }

  return json_({success:false,message:'Record not found.'});
}

function ensureCorrectionAuditColumns_(sheet) {
  const required=[
    'Last Corrected By User ID',
    'Last Corrected By Name',
    'Last Corrected At'
  ];

  const lastCol=sheet.getLastColumn();
  let headers=sheet.getRange(1,1,1,lastCol).getValues()[0];

  const positions={};
  for(const h of required){
    let idx=headers.indexOf(h);
    if(idx===-1){
      sheet.getRange(1,sheet.getLastColumn()+1).setValue(h);
      headers=sheet.getRange(1,1,1,sheet.getLastColumn()).getValues()[0];
      idx=headers.indexOf(h);
    }
    positions[h]=idx+1;
  }

  return {
    updatedBy:positions[required[0]],
    updatedName:positions[required[1]],
    updatedAt:positions[required[2]]
  };
}

function rowToRecord_(row) {
  return {
    user_id:String(row[0]||''),
    user_name:String(row[1]||''),
    account_id:String(row[2]||''),
    consumer_name:String(row[3]||''),
    father_name:String(row[4]||''),
    address:String(row[5]||''),
    supply_type:String(row[6]||''),
    load:String(row[7]||''),
    sdo_code:String(row[8]||''),
    village:String(row[9]||''),
    meter_condition:String(row[10]||''),
    ac_installed:String(row[11]||''),
    mobile_number:String(row[12]||''),
    created_at:row[13] instanceof Date
      ? Utilities.formatDate(row[13],Session.getScriptTimeZone(),'yyyy-MM-dd HH:mm:ss')
      : String(row[13]||''),
    local_record_id:String(row[14]||''),
    survey_type:String(row[15]||''),
    survey_status:String(row[16]||''),
    connection_type:String(row[17]||''),
    survey_id:String(row[18]||''),
    house_locked:String(row[19]||''),
    mobile_refused:String(row[20]||''),
    meter_installed:String(row[21]||''),
    meter_number:String(row[22]||''),
    current_reading:String(row[23]||''),
    outstanding:String(row[24]||''),
    connected_load:String(row[25]||''),
    remarks:String(row[26]||'')
  };
}

function ensureHeaders_(sheet) {
  const headers = [
    'User ID','User Name','Account ID','Consumer Name','Father/Husband',
    'Address','Supply Type','Load','SDO Code','Village','Meter Condition',
    'AC Installed','Mobile Number','Created At','Local Record ID',
    'Survey Type','Survey Status','Connection Type','Survey ID',
    'House Locked','Mobile Refused','Meter Installed','Meter Number',
    'Current Reading','Outstanding','Connected Load','Remarks'
  ];

  if (sheet.getLastRow() === 0) {
    sheet.getRange(1,1,1,headers.length).setValues([headers]);
    return;
  }

  const lastCol=sheet.getLastColumn();
  let current=sheet.getRange(1,1,1,lastCol).getValues()[0];

  for (let i=0; i<headers.length; i++) {
    if (current.indexOf(headers[i]) === -1) {
      sheet.getRange(1,sheet.getLastColumn()+1).setValue(headers[i]);
      current=sheet.getRange(1,1,1,sheet.getLastColumn()).getValues()[0];
    }
  }
}

function sha256_(text) {
  const bytes = Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_256,
    text,
    Utilities.Charset.UTF_8
  );
  return bytes.map(function(b) {
    const v = b < 0 ? b + 256 : b;
    return ('0' + v.toString(16)).slice(-2);
  }).join('');
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
