/**
 * Gaming for Good (G4G) Data Collection System
 * Cross-site research operations tracker for NSF-funded project
 *
 * Platform: Google Apps Script + Google Sheets
 * Sites: UGA and University of Missouri
 * Purpose: Track completion of 19 mandatory research instruments per participant
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

// 19 mandatory instruments in exact order
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
// ENTRY POINT & INITIALIZATION
// ============================================================================

/**
 * Entry point for web app
 * Checks authorization and serves appropriate HTML
 */
function doGet(e) {
  try {
    const userEmail = Session.getActiveUser().getEmail();
    const user = getUserByEmail(userEmail);

    // Check if user is authorized
    if (!user || !user.active) {
      return HtmlService.createTemplateFromFile('Unauthorized')
        .evaluate()
        .setTitle('G4G Data Collection System')
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
    }

    // Serve main application
    const template = HtmlService.createTemplateFromFile('Index');
    template.user = user;
    template.appName = getConfig('app_name') || 'Gaming4Good Data Collection System';

    return template.evaluate()
      .setTitle('G4G Data Collection System')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);

  } catch (error) {
    Logger.log('Error in doGet: ' + error.toString());
    return HtmlService.createHtmlOutput('<h1>Error loading application</h1><p>' + error.toString() + '</p>');
  }
}

/**
 * Include HTML files (for templates and styles)
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/**
 * Initialize application - create sheets and seed data
 * Run this once after creating the spreadsheet
 */
function initApp() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const userEmail = Session.getActiveUser().getEmail();

  Logger.log('Initializing G4G Data Collection System...');

  // Create CONFIG sheet
  createConfigSheet(ss);

  // Create INSTRUMENTS sheet
  createInstrumentsSheet(ss);

  // Create PARTICIPANTS sheet
  createParticipantsSheet(ss);

  // Create COMPLETIONS sheet
  createCompletionsSheet(ss);

  // Create USERS sheet
  createUsersSheet(ss, userEmail);

  // Create AUDIT_LOG sheet
  createAuditLogSheet(ss);

  Logger.log('Initialization complete!');

  return 'Application initialized successfully! Deploy as web app to use.';
}

/**
 * Create CONFIG sheet with default settings
 */
function createConfigSheet(ss) {
  let sheet = ss.getSheetByName(SHEET_NAMES.CONFIG);

  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAMES.CONFIG);
    sheet.appendRow(['key', 'value']);
    sheet.getRange(1, 1, 1, 2).setFontWeight('bold').setBackground('#4285f4').setFontColor('#ffffff');

    const configs = [
      ['app_name', 'Gaming4Good Data Collection System'],
      ['version', '1.0'],
      ['sites', SITES.join(',')],
      ['participant_id_format', '{SITE}-{NUMBER}']
    ];

    configs.forEach(config => sheet.appendRow(config));

    Logger.log('Created CONFIG sheet');
  }
}

/**
 * Create INSTRUMENTS sheet with 19 mandatory items
 */
function createInstrumentsSheet(ss) {
  let sheet = ss.getSheetByName(SHEET_NAMES.INSTRUMENTS);

  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAMES.INSTRUMENTS);
    sheet.appendRow(['instrument_id', 'instrument_name', 'sort_order', 'default_url', 'active']);
    sheet.getRange(1, 1, 1, 5).setFontWeight('bold').setBackground('#4285f4').setFontColor('#ffffff');

    INSTRUMENTS_DATA.forEach(inst => {
      sheet.appendRow([
        inst.id,
        inst.name,
        inst.order,
        '', // default_url - to be filled by admins
        true
      ]);
    });

    sheet.setFrozenRows(1);
    Logger.log('Created INSTRUMENTS sheet with 19 items');
  }
}

/**
 * Create PARTICIPANTS sheet
 */
function createParticipantsSheet(ss) {
  let sheet = ss.getSheetByName(SHEET_NAMES.PARTICIPANTS);

  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAMES.PARTICIPANTS);
    const headers = [
      'participant_id', 'site', 'cohort', 'enroll_date', 'created_by',
      'participant_name', 'status', 'notes', 'completion_percent',
      'missing_count', 'last_updated'
    ];
    sheet.appendRow(headers);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#4285f4').setFontColor('#ffffff');
    sheet.setFrozenRows(1);

    Logger.log('Created PARTICIPANTS sheet');
  }
}

/**
 * Create COMPLETIONS sheet (normalized)
 */
function createCompletionsSheet(ss) {
  let sheet = ss.getSheetByName(SHEET_NAMES.COMPLETIONS);

  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAMES.COMPLETIONS);
    const headers = [
      'participant_id', 'instrument_id', 'instrument_name', 'is_complete',
      'completed_at', 'completed_by', 'response_ref', 'link', 'notes'
    ];
    sheet.appendRow(headers);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#4285f4').setFontColor('#ffffff');
    sheet.setFrozenRows(1);

    Logger.log('Created COMPLETIONS sheet');
  }
}

/**
 * Create USERS sheet and add script owner as Admin
 */
function createUsersSheet(ss, ownerEmail) {
  let sheet = ss.getSheetByName(SHEET_NAMES.USERS);

  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAMES.USERS);
    sheet.appendRow(['email', 'role', 'site_scope', 'active']);
    sheet.getRange(1, 1, 1, 4).setFontWeight('bold').setBackground('#4285f4').setFontColor('#ffffff');

    // Add script owner as Admin
    sheet.appendRow([ownerEmail, 'Admin', 'ALL', true]);

    sheet.setFrozenRows(1);
    Logger.log('Created USERS sheet and added ' + ownerEmail + ' as Admin');
  }
}

/**
 * Create AUDIT_LOG sheet
 */
function createAuditLogSheet(ss) {
  let sheet = ss.getSheetByName(SHEET_NAMES.AUDIT_LOG);

  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAMES.AUDIT_LOG);
    const headers = ['timestamp', 'user_email', 'action', 'participant_id', 'instrument_id', 'details_json'];
    sheet.appendRow(headers);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#4285f4').setFontColor('#ffffff');
    sheet.setFrozenRows(1);

    Logger.log('Created AUDIT_LOG sheet');
  }
}

// ============================================================================
// AUTHENTICATION & AUTHORIZATION
// ============================================================================

/**
 * Get current user information
 */
function getCurrentUser() {
  const email = Session.getActiveUser().getEmail();
  return getUserByEmail(email);
}

/**
 * Get user by email from USERS sheet
 */
function getUserByEmail(email) {
  const cache = CacheService.getScriptCache();
  const cacheKey = 'user_' + email;
  const cached = cache.get(cacheKey);

  if (cached) {
    return JSON.parse(cached);
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAMES.USERS);
  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === email && data[i][3] === true) {
      const user = {
        email: data[i][0],
        role: data[i][1],
        site_scope: data[i][2],
        active: data[i][3]
      };

      cache.put(cacheKey, JSON.stringify(user), 300); // Cache for 5 minutes
      return user;
    }
  }

  return null;
}

/**
 * Check if user can access a specific site
 */
function canAccessSite(user, site) {
  if (!user) return false;
  if (user.site_scope === 'ALL') return true;
  return user.site_scope === site;
}

/**
 * Check if user can view participant names (PII)
 */
function canViewNames(user) {
  if (!user) return false;
  return ['Admin', 'ProjectLead', 'SiteLead'].includes(user.role);
}

/**
 * Check if user can manage users
 */
function canManageUsers(user) {
  if (!user) return false;
  return user.role === 'Admin';
}

/**
 * Check if user can create participants
 */
function canCreateParticipants(user) {
  if (!user) return false;
  return ['Admin', 'ProjectLead', 'SiteLead', 'Facilitator'].includes(user.role);
}

/**
 * Check if user can edit completions
 */
function canEditCompletions(user) {
  if (!user) return false;
  return ['Admin', 'ProjectLead', 'SiteLead', 'Facilitator'].includes(user.role);
}

// ============================================================================
// CONFIGURATION FUNCTIONS
// ============================================================================

/**
 * Get configuration value
 */
function getConfig(key) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAMES.CONFIG);
  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === key) {
      return data[i][1];
    }
  }

  return null;
}

/**
 * Set configuration value
 */
function setConfig(key, value) {
  const user = getCurrentUser();
  if (!canManageUsers(user)) {
    throw new Error('Unauthorized');
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAMES.CONFIG);
  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === key) {
      sheet.getRange(i + 1, 2).setValue(value);
      return true;
    }
  }

  // Key not found, append new row
  sheet.appendRow([key, value]);
  return true;
}

// ============================================================================
// INSTRUMENT FUNCTIONS
// ============================================================================

/**
 * Get all active instruments (cached)
 */
function getInstruments() {
  const cache = CacheService.getScriptCache();
  const cached = cache.get('instruments');

  if (cached) {
    return JSON.parse(cached);
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAMES.INSTRUMENTS);
  const data = sheet.getDataRange().getValues();

  const instruments = [];
  for (let i = 1; i < data.length; i++) {
    if (data[i][4] === true) { // active only
      instruments.push({
        instrument_id: data[i][0],
        instrument_name: data[i][1],
        sort_order: data[i][2],
        default_url: data[i][3],
        active: data[i][4]
      });
    }
  }

  // Sort by sort_order
  instruments.sort((a, b) => a.sort_order - b.sort_order);

  cache.put('instruments', JSON.stringify(instruments), 600); // Cache for 10 minutes
  return instruments;
}

/**
 * Update instrument URLs (Admin only)
 */
function updateInstrument(instrumentId, defaultUrl) {
  const user = getCurrentUser();
  if (!canManageUsers(user)) {
    throw new Error('Unauthorized');
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAMES.INSTRUMENTS);
  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === instrumentId) {
      sheet.getRange(i + 1, 4).setValue(defaultUrl);

      // Clear cache
      CacheService.getScriptCache().remove('instruments');

      logAudit('UPDATE_INSTRUMENT', {
        instrument_id: instrumentId,
        default_url: defaultUrl
      });

      return true;
    }
  }

  return false;
}

// ============================================================================
// PARTICIPANT FUNCTIONS
// ============================================================================

/**
 * Generate unique participant ID
 */
function generateParticipantId(site) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAMES.PARTICIPANTS);
  const data = sheet.getDataRange().getValues();

  // Find highest number for this site
  let maxNum = 0;
  const prefix = site + '-';

  for (let i = 1; i < data.length; i++) {
    const id = data[i][0];
    if (id && id.toString().startsWith(prefix)) {
      const num = parseInt(id.toString().replace(prefix, ''));
      if (!isNaN(num) && num > maxNum) {
        maxNum = num;
      }
    }
  }

  const nextNum = (maxNum + 1).toString().padStart(3, '0');
  return prefix + nextNum;
}

/**
 * Create new participant and completion records
 */
function createParticipant(data) {
  const user = getCurrentUser();

  if (!canCreateParticipants(user)) {
    throw new Error('You do not have permission to create participants');
  }

  if (!canAccessSite(user, data.site)) {
    throw new Error('You cannot create participants for ' + data.site);
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const participantId = generateParticipantId(data.site);
  const now = new Date();

  // Insert participant record
  const participantsSheet = ss.getSheetByName(SHEET_NAMES.PARTICIPANTS);
  participantsSheet.appendRow([
    participantId,
    data.site,
    data.cohort || '',
    now,
    user.email,
    data.participant_name || '',
    'Active',
    data.notes || '',
    0, // completion_percent
    19, // missing_count (all 19 items)
    now
  ]);

  // Create 19 completion records
  const completionsSheet = ss.getSheetByName(SHEET_NAMES.COMPLETIONS);
  const instruments = getInstruments();

  const completionRows = instruments.map(inst => [
    participantId,
    inst.instrument_id,
    inst.instrument_name,
    false, // is_complete
    '', // completed_at
    '', // completed_by
    '', // response_ref
    inst.default_url || '', // link
    '' // notes
  ]);

  if (completionRows.length > 0) {
    completionsSheet.getRange(
      completionsSheet.getLastRow() + 1,
      1,
      completionRows.length,
      completionRows[0].length
    ).setValues(completionRows);
  }

  logAudit('CREATE_PARTICIPANT', {
    participant_id: participantId,
    site: data.site,
    cohort: data.cohort
  });

  return participantId;
}

/**
 * Get participants list with filters
 */
function getParticipants(filters) {
  const user = getCurrentUser();
  if (!user) throw new Error('Unauthorized');

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAMES.PARTICIPANTS);
  const data = sheet.getDataRange().getValues();

  const participants = [];
  const showNames = canViewNames(user);

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const site = row[1];

    // Check site access
    if (!canAccessSite(user, site)) continue;

    // Apply filters
    if (filters) {
      if (filters.site && filters.site !== 'ALL' && site !== filters.site) continue;
      if (filters.cohort && row[2] !== filters.cohort) continue;
      if (filters.status && row[6] !== filters.status) continue;
      if (filters.search) {
        const searchLower = filters.search.toLowerCase();
        const participantId = (row[0] || '').toString().toLowerCase();
        if (!participantId.includes(searchLower)) continue;
      }
    }

    participants.push({
      participant_id: row[0],
      site: row[1],
      cohort: row[2],
      enroll_date: row[3],
      created_by: row[4],
      participant_name: showNames ? row[5] : '',
      status: row[6],
      notes: row[7],
      completion_percent: row[8],
      missing_count: row[9],
      last_updated: row[10]
    });
  }

  return participants;
}

/**
 * Get participant detail with completion checklist
 */
function getParticipantDetail(participantId) {
  const user = getCurrentUser();
  if (!user) throw new Error('Unauthorized');

  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // Get participant info
  const participantsSheet = ss.getSheetByName(SHEET_NAMES.PARTICIPANTS);
  const participantData = participantsSheet.getDataRange().getValues();

  let participant = null;
  for (let i = 1; i < participantData.length; i++) {
    if (participantData[i][0] === participantId) {
      const site = participantData[i][1];

      // Check site access
      if (!canAccessSite(user, site)) {
        throw new Error('You do not have access to this participant');
      }

      const showNames = canViewNames(user);

      participant = {
        participant_id: participantData[i][0],
        site: participantData[i][1],
        cohort: participantData[i][2],
        enroll_date: participantData[i][3],
        created_by: participantData[i][4],
        participant_name: showNames ? participantData[i][5] : '',
        status: participantData[i][6],
        notes: participantData[i][7],
        completion_percent: participantData[i][8],
        missing_count: participantData[i][9],
        last_updated: participantData[i][10]
      };
      break;
    }
  }

  if (!participant) {
    throw new Error('Participant not found');
  }

  // Get completions
  const completionsSheet = ss.getSheetByName(SHEET_NAMES.COMPLETIONS);
  const completionsData = completionsSheet.getDataRange().getValues();

  const completions = [];
  for (let i = 1; i < completionsData.length; i++) {
    if (completionsData[i][0] === participantId) {
      completions.push({
        participant_id: completionsData[i][0],
        instrument_id: completionsData[i][1],
        instrument_name: completionsData[i][2],
        is_complete: completionsData[i][3],
        completed_at: completionsData[i][4],
        completed_by: completionsData[i][5],
        response_ref: completionsData[i][6],
        link: completionsData[i][7],
        notes: completionsData[i][8]
      });
    }
  }

  // Sort by instrument order
  const instruments = getInstruments();
  const instrumentOrder = {};
  instruments.forEach(inst => {
    instrumentOrder[inst.instrument_id] = inst.sort_order;
  });

  completions.sort((a, b) => {
    return (instrumentOrder[a.instrument_id] || 999) - (instrumentOrder[b.instrument_id] || 999);
  });

  return {
    participant: participant,
    completions: completions
  };
}

/**
 * Update completion status
 */
function updateCompletion(participantId, instrumentId, completionData) {
  const user = getCurrentUser();

  if (!canEditCompletions(user)) {
    throw new Error('You do not have permission to edit completions');
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // Check site access
  const participantsSheet = ss.getSheetByName(SHEET_NAMES.PARTICIPANTS);
  const participantData = participantsSheet.getDataRange().getValues();
  let participantSite = null;

  for (let i = 1; i < participantData.length; i++) {
    if (participantData[i][0] === participantId) {
      participantSite = participantData[i][1];
      break;
    }
  }

  if (!participantSite || !canAccessSite(user, participantSite)) {
    throw new Error('You do not have access to this participant');
  }

  // Update completion record
  const completionsSheet = ss.getSheetByName(SHEET_NAMES.COMPLETIONS);
  const completionsData = completionsSheet.getDataRange().getValues();

  for (let i = 1; i < completionsData.length; i++) {
    if (completionsData[i][0] === participantId && completionsData[i][1] === instrumentId) {
      const now = new Date();

      completionsSheet.getRange(i + 1, 4).setValue(completionData.is_complete);
      completionsSheet.getRange(i + 1, 5).setValue(completionData.is_complete ? now : '');
      completionsSheet.getRange(i + 1, 6).setValue(completionData.is_complete ? user.email : '');
      completionsSheet.getRange(i + 1, 7).setValue(completionData.response_ref || '');
      completionsSheet.getRange(i + 1, 9).setValue(completionData.notes || '');

      break;
    }
  }

  // Recalculate completion stats
  recalculateCompletion(participantId);

  logAudit('UPDATE_COMPLETION', {
    participant_id: participantId,
    instrument_id: instrumentId,
    is_complete: completionData.is_complete
  });

  return true;
}

/**
 * Recalculate completion percentage and missing count for a participant
 */
function recalculateCompletion(participantId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // Count completions
  const completionsSheet = ss.getSheetByName(SHEET_NAMES.COMPLETIONS);
  const completionsData = completionsSheet.getDataRange().getValues();

  let totalCount = 0;
  let completedCount = 0;

  for (let i = 1; i < completionsData.length; i++) {
    if (completionsData[i][0] === participantId) {
      totalCount++;
      if (completionsData[i][3] === true) {
        completedCount++;
      }
    }
  }

  const completionPercent = totalCount > 0 ? Math.round((completedCount / totalCount) * 100) : 0;
  const missingCount = totalCount - completedCount;

  // Update participant record
  const participantsSheet = ss.getSheetByName(SHEET_NAMES.PARTICIPANTS);
  const participantData = participantsSheet.getDataRange().getValues();

  for (let i = 1; i < participantData.length; i++) {
    if (participantData[i][0] === participantId) {
      participantsSheet.getRange(i + 1, 9).setValue(completionPercent);
      participantsSheet.getRange(i + 1, 10).setValue(missingCount);
      participantsSheet.getRange(i + 1, 11).setValue(new Date());
      break;
    }
  }

  return { completion_percent: completionPercent, missing_count: missingCount };
}

/**
 * Recalculate all participants (maintenance task)
 */
function recalculateAllCompletions() {
  const user = getCurrentUser();
  if (!canManageUsers(user)) {
    throw new Error('Unauthorized');
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const participantsSheet = ss.getSheetByName(SHEET_NAMES.PARTICIPANTS);
  const participantData = participantsSheet.getDataRange().getValues();

  let count = 0;
  for (let i = 1; i < participantData.length; i++) {
    const participantId = participantData[i][0];
    recalculateCompletion(participantId);
    count++;
  }

  logAudit('RECALCULATE_ALL', { count: count });

  return 'Recalculated ' + count + ' participants';
}

/**
 * Update participant status
 */
function updateParticipantStatus(participantId, status, notes) {
  const user = getCurrentUser();

  if (!canEditCompletions(user)) {
    throw new Error('Unauthorized');
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const participantsSheet = ss.getSheetByName(SHEET_NAMES.PARTICIPANTS);
  const participantData = participantsSheet.getDataRange().getValues();

  for (let i = 1; i < participantData.length; i++) {
    if (participantData[i][0] === participantId) {
      const site = participantData[i][1];

      if (!canAccessSite(user, site)) {
        throw new Error('Unauthorized');
      }

      participantsSheet.getRange(i + 1, 7).setValue(status);
      if (notes !== undefined) {
        participantsSheet.getRange(i + 1, 8).setValue(notes);
      }
      participantsSheet.getRange(i + 1, 11).setValue(new Date());

      logAudit('UPDATE_PARTICIPANT_STATUS', {
        participant_id: participantId,
        status: status
      });

      return true;
    }
  }

  throw new Error('Participant not found');
}

// ============================================================================
// DASHBOARD & STATISTICS
// ============================================================================

/**
 * Get dashboard statistics
 */
function getDashboardStats(filters) {
  const user = getCurrentUser();
  if (!user) throw new Error('Unauthorized');

  const participants = getParticipants(filters);

  // Calculate statistics
  const totalParticipants = participants.length;
  const fullyComplete = participants.filter(p => p.completion_percent === 100).length;
  const fullyCompletePercent = totalParticipants > 0 ?
    Math.round((fullyComplete / totalParticipants) * 100) : 0;

  const totalCompletion = participants.reduce((sum, p) => sum + (p.completion_percent || 0), 0);
  const avgCompletion = totalParticipants > 0 ?
    Math.round(totalCompletion / totalParticipants) : 0;

  const totalMissing = participants.reduce((sum, p) => sum + (p.missing_count || 0), 0);

  // Find top missing instruments
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const completionsSheet = ss.getSheetByName(SHEET_NAMES.COMPLETIONS);
  const completionsData = completionsSheet.getDataRange().getValues();

  const participantIds = new Set(participants.map(p => p.participant_id));
  const missingByInstrument = {};

  for (let i = 1; i < completionsData.length; i++) {
    const participantId = completionsData[i][0];
    const instrumentName = completionsData[i][2];
    const isComplete = completionsData[i][3];

    if (participantIds.has(participantId) && !isComplete) {
      missingByInstrument[instrumentName] = (missingByInstrument[instrumentName] || 0) + 1;
    }
  }

  const topMissing = Object.entries(missingByInstrument)
    .map(([name, count]) => ({ instrument_name: name, missing_count: count }))
    .sort((a, b) => b.missing_count - a.missing_count)
    .slice(0, 5);

  // Get unique cohorts for filter
  const cohorts = [...new Set(participants.map(p => p.cohort).filter(c => c))].sort();

  return {
    stats: {
      total_participants: totalParticipants,
      fully_complete: fullyComplete,
      fully_complete_percent: fullyCompletePercent,
      avg_completion: avgCompletion,
      total_missing: totalMissing,
      top_missing: topMissing
    },
    participants: participants,
    cohorts: cohorts,
    user: {
      email: user.email,
      role: user.role,
      site_scope: user.site_scope,
      can_view_names: canViewNames(user),
      can_create: canCreateParticipants(user),
      can_edit: canEditCompletions(user),
      can_admin: canManageUsers(user)
    }
  };
}

// ============================================================================
// USER MANAGEMENT
// ============================================================================

/**
 * Get all users (Admin only)
 */
function getUsers() {
  const user = getCurrentUser();
  if (!canManageUsers(user)) {
    throw new Error('Unauthorized');
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAMES.USERS);
  const data = sheet.getDataRange().getValues();

  const users = [];
  for (let i = 1; i < data.length; i++) {
    users.push({
      email: data[i][0],
      role: data[i][1],
      site_scope: data[i][2],
      active: data[i][3]
    });
  }

  return users;
}

/**
 * Save or update user (Admin only)
 */
function saveUser(userData) {
  const user = getCurrentUser();
  if (!canManageUsers(user)) {
    throw new Error('Unauthorized');
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAMES.USERS);
  const data = sheet.getDataRange().getValues();

  // Check if user exists
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === userData.email) {
      // Update existing user
      sheet.getRange(i + 1, 2).setValue(userData.role);
      sheet.getRange(i + 1, 3).setValue(userData.site_scope);
      sheet.getRange(i + 1, 4).setValue(userData.active);

      // Clear cache
      CacheService.getScriptCache().remove('user_' + userData.email);

      logAudit('UPDATE_USER', { email: userData.email, role: userData.role });
      return true;
    }
  }

  // Add new user
  sheet.appendRow([
    userData.email,
    userData.role,
    userData.site_scope,
    userData.active !== false // default to true
  ]);

  logAudit('CREATE_USER', { email: userData.email, role: userData.role });
  return true;
}

/**
 * Delete user (Admin only)
 */
function deleteUser(email) {
  const user = getCurrentUser();
  if (!canManageUsers(user)) {
    throw new Error('Unauthorized');
  }

  // Prevent deleting self
  if (email === user.email) {
    throw new Error('Cannot delete your own account');
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAMES.USERS);
  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === email) {
      sheet.deleteRow(i + 1);
      CacheService.getScriptCache().remove('user_' + email);
      logAudit('DELETE_USER', { email: email });
      return true;
    }
  }

  return false;
}

// ============================================================================
// EXPORT FUNCTIONS
// ============================================================================

/**
 * Export participants to CSV
 */
function exportParticipants(filters) {
  const user = getCurrentUser();
  if (!user) throw new Error('Unauthorized');

  const participants = getParticipants(filters);
  const showNames = canViewNames(user);

  // Build CSV
  let csv = 'participant_id,site,cohort,enroll_date,status,completion_percent,missing_count,last_updated';
  if (showNames) {
    csv = 'participant_id,site,cohort,enroll_date,participant_name,status,completion_percent,missing_count,last_updated';
  }
  csv += '\n';

  participants.forEach(p => {
    const row = [
      p.participant_id,
      p.site,
      p.cohort,
      p.enroll_date,
      showNames ? p.participant_name : null,
      p.status,
      p.completion_percent,
      p.missing_count,
      p.last_updated
    ].filter(v => v !== null);

    csv += row.map(v => '"' + (v || '').toString().replace(/"/g, '""') + '"').join(',') + '\n';
  });

  logAudit('EXPORT_PARTICIPANTS', { count: participants.length });

  return csv;
}

/**
 * Export completions to CSV
 */
function exportCompletions(filters) {
  const user = getCurrentUser();
  if (!user) throw new Error('Unauthorized');

  const participants = getParticipants(filters);
  const participantIds = new Set(participants.map(p => p.participant_id));

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const completionsSheet = ss.getSheetByName(SHEET_NAMES.COMPLETIONS);
  const completionsData = completionsSheet.getDataRange().getValues();

  let csv = 'participant_id,instrument_name,is_complete,completed_at,completed_by,response_ref\n';

  for (let i = 1; i < completionsData.length; i++) {
    const participantId = completionsData[i][0];

    if (participantIds.has(participantId)) {
      const row = [
        completionsData[i][0], // participant_id
        completionsData[i][2], // instrument_name
        completionsData[i][3] ? 'Yes' : 'No', // is_complete
        completionsData[i][4], // completed_at
        completionsData[i][5], // completed_by
        completionsData[i][6]  // response_ref
      ];

      csv += row.map(v => '"' + (v || '').toString().replace(/"/g, '""') + '"').join(',') + '\n';
    }
  }

  logAudit('EXPORT_COMPLETIONS', { count: participantIds.size });

  return csv;
}

/**
 * Export single participant checklist
 */
function exportParticipantChecklist(participantId) {
  const detail = getParticipantDetail(participantId);

  let csv = 'instrument_name,is_complete,completed_at,completed_by,response_ref,link\n';

  detail.completions.forEach(c => {
    const row = [
      c.instrument_name,
      c.is_complete ? 'Yes' : 'No',
      c.completed_at,
      c.completed_by,
      c.response_ref,
      c.link
    ];

    csv += row.map(v => '"' + (v || '').toString().replace(/"/g, '""') + '"').join(',') + '\n';
  });

  logAudit('EXPORT_PARTICIPANT_CHECKLIST', { participant_id: participantId });

  return csv;
}

// ============================================================================
// AUDIT LOG
// ============================================================================

/**
 * Log audit event
 */
function logAudit(action, details) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_NAMES.AUDIT_LOG);
    const user = getCurrentUser();

    sheet.appendRow([
      new Date(),
      user ? user.email : 'system',
      action,
      details.participant_id || '',
      details.instrument_id || '',
      JSON.stringify(details)
    ]);
  } catch (error) {
    Logger.log('Error logging audit: ' + error.toString());
  }
}

/**
 * Get recent audit logs (Admin only)
 */
function getAuditLogs(limit) {
  const user = getCurrentUser();
  if (!canManageUsers(user)) {
    throw new Error('Unauthorized');
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAMES.AUDIT_LOG);
  const data = sheet.getDataRange().getValues();

  const logs = [];
  const startRow = Math.max(1, data.length - (limit || 100));

  for (let i = data.length - 1; i >= startRow; i--) {
    logs.push({
      timestamp: data[i][0],
      user_email: data[i][1],
      action: data[i][2],
      participant_id: data[i][3],
      instrument_id: data[i][4],
      details_json: data[i][5]
    });
  }

  return logs;
}
