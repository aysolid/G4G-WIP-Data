/**
 * Gaming for Good (G4G) Data Collection System v2.0
 * Authenticated via Username/Password
 */

// ============================================================================
// CONFIGURATION & CONSTANTS
// ============================================================================

const SHEET_NAMES = {
  CONFIG: 'CONFIG',
  INSTRUMENTS: 'INSTRUMENTS',
  PARTICIPANTS: 'PARTICIPANTS',
  COMPLETIONS: 'COMPLETIONS',
  USERS: 'USERS',
  AUDIT_LOG: 'AUDIT_LOG'
};

const SITES = ['UGA', 'MIZZOU'];
const ROLES = ['Admin', 'ProjectLead', 'SiteLead', 'Facilitator', 'Viewer'];
const STATUSES = ['Active', 'Withdrawn', 'Completed'];

// 19 mandatory instruments
const INSTRUMENTS_DATA = [
  { id: 'consent', name: 'Consent form', order: 1 },
  { id: 'assent', name: 'Assent Form', order: 2 },
  { id: 'pretest', name: 'Pre-test', order: 3 },
  { id: 'l1_journal', name: 'Lesson 1 Journal', order: 4 },
  { id: 'l2_journal', name: 'Lesson 2 Journal', order: 5 },
  { id: 'l3_journal', name: 'Lesson 3 Journal', order: 6 },
  { id: 'l3_worksheet', name: 'Lesson 3 - Design a Game Worksheet', order: 7 },
  { id: 'l4_journal', name: 'Lesson 4 Journal', order: 8 },
  { id: 'l4_worksheet', name: 'Lesson 4 - Paper Prototyping Worksheet', order: 9 },
  { id: 'l5_journal', name: 'Lesson 5 Journal', order: 10 },
  { id: 'l5_worksheet', name: 'Lesson 5 - Debugging Worksheet', order: 11 },
  { id: 'l6_journal', name: 'Lesson 6 Journal', order: 12 },
  { id: 'l6_worksheet', name: 'Lesson 6 - Game Refinement Worksheet', order: 13 },
  { id: 'l7_journal', name: 'Lesson 7 Journal', order: 14 },
  { id: 'l7_worksheet', name: 'Lesson 7 - Playtesting Feedback Guide', order: 15 },
  { id: 'l8_journal', name: 'Lesson 8 Journal', order: 16 },
  { id: 'posttest', name: 'Post - Test', order: 17 },
  { id: 'participant_feedback', name: 'Participant Feedback Survey', order: 18 },
  { id: 'parent_satisfaction', name: 'Parent Satisfaction Survey', order: 19 }
];

// ============================================================================
// CORE & AUTHENTICATION
// ============================================================================

function doGet(e) {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('G4G Data Collection System')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/**
 * Authenticate user and return session token
 */
function login(username, password) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAMES.USERS);
  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    // Check credentials (simple text match for prototype - use hashing in production)
    if (data[i][0] == username && data[i][1] == password && data[i][4] === true) {
      const user = {
        username: data[i][0],
        role: data[i][2],
        site_scope: data[i][3],
        can_admin: data[i][2] === 'Admin' || data[i][2] === 'ProjectLead',
        can_enroll: ['Admin', 'ProjectLead', 'SiteLead', 'Facilitator'].includes(data[i][2]),
        can_view_names: ['Admin', 'ProjectLead', 'SiteLead'].includes(data[i][2])
      };
      
      // Generate Token
      const token = Utilities.getUuid();
      CacheService.getScriptCache().put(token, JSON.stringify(user), 21600); // 6 hours
      
      return { success: true, token: token, user: user };
    }
  }
  
  throw new Error('Invalid username or password');
}

/**
 * Verify token and retrieve user object
 */
function getUserFromToken(token) {
  if (!token) throw new Error('Session invalid. Please login again.');
  const userJson = CacheService.getScriptCache().get(token);
  if (!userJson) throw new Error('Session expired. Please login again.');
  return JSON.parse(userJson);
}

// ============================================================================
// INITIALIZATION
// ============================================================================

function initApp() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  createConfigSheet(ss);
  createInstrumentsSheet(ss);
  createParticipantsSheet(ss);
  createCompletionsSheet(ss);
  createUsersSheet(ss);
  createAuditLogSheet(ss);
  
  return 'System Initialized. Default Admin: admin / g4g2024';
}

function createUsersSheet(ss) {
  let sheet = ss.getSheetByName(SHEET_NAMES.USERS);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAMES.USERS);
    // New Structure: Username, Password, Role, Site, Active
    sheet.appendRow(['username', 'password', 'role', 'site_scope', 'active']);
    sheet.getRange(1, 1, 1, 5).setFontWeight('bold').setBackground('#4285f4').setFontColor('#ffffff');
    // Default Admin
    sheet.appendRow(['admin', 'g4g2024', 'Admin', 'ALL', true]);
    sheet.setFrozenRows(1);
  }
}

// Helper functions for other sheets (Config, Instruments, etc.) remain mostly same
function createConfigSheet(ss) { /* ... same as before ... */ 
  let sheet = ss.getSheetByName(SHEET_NAMES.CONFIG);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAMES.CONFIG);
    sheet.appendRow(['key', 'value']);
    sheet.getRange(1, 1, 1, 2).setFontWeight('bold').setBackground('#4285f4').setFontColor('#ffffff');
    const configs = [['app_name', 'G4G Data Hub'], ['version', '2.0']];
    configs.forEach(c => sheet.appendRow(c));
  }
}
function createInstrumentsSheet(ss) {
  let sheet = ss.getSheetByName(SHEET_NAMES.INSTRUMENTS);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAMES.INSTRUMENTS);
    sheet.appendRow(['instrument_id', 'instrument_name', 'sort_order', 'default_url', 'active']);
    sheet.getRange(1, 1, 1, 5).setFontWeight('bold').setBackground('#4285f4').setFontColor('#ffffff');
    INSTRUMENTS_DATA.forEach(inst => sheet.appendRow([inst.id, inst.name, inst.order, '', true]));
  }
}
function createParticipantsSheet(ss) {
  let sheet = ss.getSheetByName(SHEET_NAMES.PARTICIPANTS);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAMES.PARTICIPANTS);
    // Added created_by_username
    const headers = ['participant_id', 'site', 'cohort', 'enroll_date', 'created_by', 'participant_name', 'status', 'notes', 'completion_percent', 'missing_count', 'last_updated'];
    sheet.appendRow(headers);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#4285f4').setFontColor('#ffffff');
  }
}
function createCompletionsSheet(ss) {
  let sheet = ss.getSheetByName(SHEET_NAMES.COMPLETIONS);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAMES.COMPLETIONS);
    const headers = ['participant_id', 'instrument_id', 'instrument_name', 'is_complete', 'completed_at', 'completed_by', 'response_ref', 'link', 'notes'];
    sheet.appendRow(headers);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#4285f4').setFontColor('#ffffff');
  }
}
function createAuditLogSheet(ss) {
  let sheet = ss.getSheetByName(SHEET_NAMES.AUDIT_LOG);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAMES.AUDIT_LOG);
    sheet.appendRow(['timestamp', 'username', 'action', 'target_id', 'details']);
    sheet.getRange(1, 1, 1, 5).setFontWeight('bold').setBackground('#4285f4').setFontColor('#ffffff');
  }
}

// ============================================================================
// DATA ACCESS & LOGIC
// ============================================================================

function canAccessSite(user, site) {
  return user.site_scope === 'ALL' || user.site_scope === site;
}

// --- DASHBOARD ---
function getDashboardStats(token, filters) {
  const user = getUserFromToken(token); // Verify Auth
  const participants = getParticipantsLogic(user, filters);

  const totalParticipants = participants.length;
  const fullyComplete = participants.filter(p => p.completion_percent === 100).length;
  const fullyCompletePercent = totalParticipants > 0 ? Math.round((fullyComplete / totalParticipants) * 100) : 0;
  const totalCompletion = participants.reduce((sum, p) => sum + (p.completion_percent || 0), 0);
  const avgCompletion = totalParticipants > 0 ? Math.round(totalCompletion / totalParticipants) : 0;
  const totalMissing = participants.reduce((sum, p) => sum + (p.missing_count || 0), 0);

  // Missing analysis
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const completionsData = ss.getSheetByName(SHEET_NAMES.COMPLETIONS).getDataRange().getValues();
  const participantIds = new Set(participants.map(p => p.participant_id));
  const missingByInstrument = {};

  for (let i = 1; i < completionsData.length; i++) {
    if (participantIds.has(completionsData[i][0]) && !completionsData[i][3]) {
      const name = completionsData[i][2];
      missingByInstrument[name] = (missingByInstrument[name] || 0) + 1;
    }
  }

  const topMissing = Object.entries(missingByInstrument)
    .map(([name, count]) => ({ instrument_name: name, missing_count: count }))
    .sort((a, b) => b.missing_count - a.missing_count).slice(0, 5);

  const cohorts = [...new Set(participants.map(p => p.cohort).filter(c => c))].sort();

  return {
    stats: { total_participants: totalParticipants, fully_complete: fullyComplete, fully_complete_percent: fullyCompletePercent, avg_completion: avgCompletion, total_missing: totalMissing, top_missing: topMissing },
    participants: participants,
    cohorts: cohorts,
    user: user
  };
}

function getParticipantsLogic(user, filters) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const data = ss.getSheetByName(SHEET_NAMES.PARTICIPANTS).getDataRange().getValues();
  const participants = [];
  const showNames = user.can_view_names;

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (!canAccessSite(user, row[1])) continue;
    
    if (filters) {
      if (filters.site && filters.site !== 'ALL' && row[1] !== filters.site) continue;
      if (filters.cohort && row[2] !== filters.cohort) continue;
      if (filters.status && row[6] !== filters.status) continue;
      if (filters.search) {
        if (!String(row[0]).toLowerCase().includes(filters.search.toLowerCase())) continue;
      }
    }

    participants.push({
      participant_id: row[0], site: row[1], cohort: row[2], enroll_date: row[3],
      created_by: row[4], participant_name: showNames ? row[5] : '***',
      status: row[6], notes: row[7], completion_percent: row[8], missing_count: row[9], last_updated: row[10]
    });
  }
  return participants;
}

// --- PARTICIPANT MANAGEMENT ---

function createParticipant(token, data) {
  const user = getUserFromToken(token);
  if (!user.can_enroll) throw new Error('Permission denied');
  if (!canAccessSite(user, data.site)) throw new Error('Site access denied');
  if (!data.participant_name) throw new Error('Participant Name is required');

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Generate ID
  const pSheet = ss.getSheetByName(SHEET_NAMES.PARTICIPANTS);
  const pData = pSheet.getDataRange().getValues();
  let maxNum = 0;
  const prefix = data.site + '-';
  for (let i = 1; i < pData.length; i++) {
    const id = String(pData[i][0]);
    if (id.startsWith(prefix)) {
      const num = parseInt(id.replace(prefix, ''));
      if (!isNaN(num) && num > maxNum) maxNum = num;
    }
  }
  const pid = prefix + (maxNum + 1).toString().padStart(3, '0');
  const now = new Date();

  pSheet.appendRow([
    pid, data.site, data.cohort, now, user.username, 
    data.participant_name, 'Active', data.notes, 0, 19, now
  ]);

  // Completions
  const cSheet = ss.getSheetByName(SHEET_NAMES.COMPLETIONS);
  const instruments = getInstruments(); // Helper below
  const rows = instruments.map(inst => [pid, inst.instrument_id, inst.instrument_name, false, '', '', '', inst.default_url || '', '']);
  if (rows.length > 0) cSheet.getRange(cSheet.getLastRow()+1, 1, rows.length, rows[0].length).setValues(rows);

  logAudit(user.username, 'CREATE_PARTICIPANT', pid, {site: data.site});
  return pid;
}

function getParticipantDetail(token, participantId) {
  const user = getUserFromToken(token);
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const pData = ss.getSheetByName(SHEET_NAMES.PARTICIPANTS).getDataRange().getValues();
  
  let p = null;
  for(let i=1; i<pData.length; i++) {
    if(pData[i][0] === participantId) {
      if(!canAccessSite(user, pData[i][1])) throw new Error('Access denied');
      p = {
        participant_id: pData[i][0], site: pData[i][1], cohort: pData[i][2],
        enroll_date: pData[i][3], created_by: pData[i][4],
        participant_name: user.can_view_names ? pData[i][5] : '***',
        status: pData[i][6], notes: pData[i][7], completion_percent: pData[i][8],
        missing_count: pData[i][9], last_updated: pData[i][10]
      };
      break;
    }
  }
  if(!p) throw new Error('Participant not found');

  const cData = ss.getSheetByName(SHEET_NAMES.COMPLETIONS).getDataRange().getValues();
  const completions = [];
  for(let i=1; i<cData.length; i++) {
    if(cData[i][0] === participantId) {
      completions.push({
        instrument_id: cData[i][1], instrument_name: cData[i][2],
        is_complete: cData[i][3], completed_at: cData[i][4], completed_by: cData[i][5],
        response_ref: cData[i][6], link: cData[i][7], notes: cData[i][8]
      });
    }
  }
  
  // Sort
  const instruments = getInstruments();
  const order = {}; instruments.forEach(i => order[i.instrument_id] = i.sort_order);
  completions.sort((a,b) => (order[a.instrument_id]||999) - (order[b.instrument_id]||999));

  return { participant: p, completions: completions, user: user };
}

function updateCompletion(token, participantId, instrumentId, data) {
  const user = getUserFromToken(token);
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Verify access first
  const pSheet = ss.getSheetByName(SHEET_NAMES.PARTICIPANTS);
  const pData = pSheet.getDataRange().getValues();
  let valid = false;
  for(let i=1; i<pData.length; i++) {
    if(pData[i][0] === participantId && canAccessSite(user, pData[i][1])) { valid = true; break; }
  }
  if(!valid) throw new Error('Access denied');

  const cSheet = ss.getSheetByName(SHEET_NAMES.COMPLETIONS);
  const cData = cSheet.getDataRange().getValues();
  for(let i=1; i<cData.length; i++) {
    if(cData[i][0] === participantId && cData[i][1] === instrumentId) {
      const row = i+1;
      const now = new Date();
      cSheet.getRange(row, 4).setValue(data.is_complete);
      cSheet.getRange(row, 5).setValue(data.is_complete ? now : '');
      cSheet.getRange(row, 6).setValue(data.is_complete ? user.username : '');
      cSheet.getRange(row, 7).setValue(data.response_ref || '');
      cSheet.getRange(row, 9).setValue(data.notes || '');
      break;
    }
  }
  recalculateCompletion(participantId);
  return true;
}

function updateParticipantStatus(token, participantId, status) {
  const user = getUserFromToken(token);
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const pSheet = ss.getSheetByName(SHEET_NAMES.PARTICIPANTS);
  const data = pSheet.getDataRange().getValues();
  
  for(let i=1; i<data.length; i++) {
    if(data[i][0] === participantId) {
      if(!canAccessSite(user, data[i][1])) throw new Error('Access denied');
      pSheet.getRange(i+1, 7).setValue(status);
      pSheet.getRange(i+1, 11).setValue(new Date());
      logAudit(user.username, 'UPDATE_STATUS', participantId, {status: status});
      return true;
    }
  }
}

// --- ADMIN FUNCTIONS ---

function getUsers(token) {
  const user = getUserFromToken(token);
  if (!user.can_admin) throw new Error('Unauthorized');
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const data = ss.getSheetByName(SHEET_NAMES.USERS).getDataRange().getValues();
  const users = [];
  for(let i=1; i<data.length; i++) {
    users.push({ username: data[i][0], password: data[i][1], role: data[i][2], site_scope: data[i][3], active: data[i][4] });
  }
  return users;
}

function saveUser(token, userData) {
  const user = getUserFromToken(token);
  if (!user.can_admin) throw new Error('Unauthorized');
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAMES.USERS);
  const data = sheet.getDataRange().getValues();
  
  // Update existing
  for(let i=1; i<data.length; i++) {
    if(data[i][0] === userData.username) {
      sheet.getRange(i+1, 2).setValue(userData.password);
      sheet.getRange(i+1, 3).setValue(userData.role);
      sheet.getRange(i+1, 4).setValue(userData.site_scope);
      sheet.getRange(i+1, 5).setValue(userData.active);
      return true;
    }
  }
  // Create new
  sheet.appendRow([userData.username, userData.password, userData.role, userData.site_scope, userData.active]);
  logAudit(user.username, 'CREATE_USER', userData.username, {});
  return true;
}

// --- HELPERS ---
function getInstruments() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const data = ss.getSheetByName(SHEET_NAMES.INSTRUMENTS).getDataRange().getValues();
  return data.slice(1).filter(r => r[4]).map(r => ({instrument_id: r[0], instrument_name: r[1], sort_order: r[2], default_url: r[3]}));
}

function recalculateCompletion(participantId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const cData = ss.getSheetByName(SHEET_NAMES.COMPLETIONS).getDataRange().getValues();
  let total=0, complete=0;
  for(let i=1; i<cData.length; i++) {
    if(cData[i][0] === participantId) {
      total++;
      if(cData[i][3]) complete++;
    }
  }
  const pct = total ? Math.round((complete/total)*100) : 0;
  
  const pSheet = ss.getSheetByName(SHEET_NAMES.PARTICIPANTS);
  const pData = pSheet.getDataRange().getValues();
  for(let i=1; i<pData.length; i++) {
    if(pData[i][0] === participantId) {
      pSheet.getRange(i+1, 9).setValue(pct);
      pSheet.getRange(i+1, 10).setValue(total-complete);
      pSheet.getRange(i+1, 11).setValue(new Date());
      break;
    }
  }
}

function logAudit(username, action, target, details) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.getSheetByName(SHEET_NAMES.AUDIT_LOG).appendRow([new Date(), username, action, target, JSON.stringify(details)]);
}