/**
 * Citimotors AI CRM — Google Apps Script Web App
 * Deploy: Extensions > Apps Script (sa parehong Google Sheet) > paste this file >
 *   Deploy > New deployment > type "Web app" > Execute as "Me" > Who has access "Anyone" >
 *   Deploy, then copy the Web App URL into the CRM's Sync (QR) tab ("Write endpoint").
 *
 * If updating an EXISTING deployment (recommended, so the URL stays the same):
 *   Deploy > Manage deployments > (pencil icon on your active deployment) >
 *   Version: "New version" > Deploy.
 *
 * Handles:
 *  - READS (doGet): ?action=leads and ?action=conversations&leadId=X — these feed the
 *    CRM's Pipeline and AI Conversations pages. Without this, the dashboard always shows
 *    "Walang lead na makita" no matter what, because it has nothing to fetch.
 *  - WRITES (doPost): 'updateLead' (pipeline stage changes, form saves, etc.) and
 *    'setTakeover' (the AI/Human toggle switch in the CRM — Pipeline card AND AI Conversations
 *    panel both call this same action).
 *
 * IMPORTANT: Whenever Human Takeover changes (via either 'updateLead' with a 'Human Takeover'
 * field, or via 'setTakeover'), this ALSO updates the "Chat Control" tab's "Assigned To"
 * column (matched by Facebook ID) — that's the value the n8n workflow's "Check Assigned To"
 * node actually reads to decide whether the AI should keep auto-replying. Without this, the
 * CRM toggle only changes a label in the Leads tab and never actually pauses the bot.
 */
const LEADS_SHEET_NAME = 'Leads';
const LEAD_ID_COLUMN = 'Lead ID';
const CONVERSATIONS_SHEET_NAME = 'Conversations';
const CHAT_CONTROL_SHEET_NAME = 'Chat Control';
const CHAT_CONTROL_FB_COLUMN = 'Facebook ID';
const CHAT_CONTROL_ASSIGNED_COLUMN = 'Assigned To';
// Must match the value n8n's "Check Assigned To" / "Check Assigned To1" nodes compare against.
const HUMAN_TAKEOVER_ASSIGNEE = 'Romeo';

// ============================== READS (doGet) ==============================

function doGet(e) {
  const action = e.parameter.action;

  if (action === 'leads') {
    return jsonOutput(getSheetAsObjects(LEADS_SHEET_NAME));
  }

  if (action === 'conversations') {
    const leadId = e.parameter.leadId || '';
    const rows = getSheetAsObjects(CONVERSATIONS_SHEET_NAME);
    const filtered = rows.filter(r =>
      String(r['Lead ID']) === String(leadId) ||
      String(r['Facebook ID']) === String(leadId)
    );
    return jsonOutput(filtered);
  }

  return jsonOutput({ ok: false, error: 'Unknown action' });
}

// Generic header-row -> array-of-objects reader. Works for any sheet as long as
// row 1 has column headers — doesn't care about column order, so it stays correct
// even if columns get reordered or new ones get added later.
function getSheetAsObjects(sheetName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) return [];
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return []; // header row only (or completely empty) = no data yet

  const headers = data[0];
  const rows = [];
  for (let r = 1; r < data.length; r++) {
    const row = data[r];
    const isBlank = row.every(cell => cell === '' || cell === null);
    if (isBlank) continue; // skip fully empty rows
    const obj = {};
    headers.forEach((h, i) => { obj[h] = row[i]; });
    rows.push(obj);
  }
  return rows;
}

function jsonOutput(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

// ============================== WRITES (doPost) ==============================

function doPost(e) {
  const payload = JSON.parse(e.postData.contents);

  if (payload.action === 'updateLead') {
    updateLead(payload.leadId, payload.fields);
    return jsonOutput({ ok: true });
  }

  if (payload.action === 'setTakeover') {
    setTakeover(payload.facebookId, payload.mode);
    return jsonOutput({ ok: true });
  }

  return jsonOutput({ ok: false, error: 'Unknown action' });
}

function updateLead(leadId, fields) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(LEADS_SHEET_NAME);
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const idColIdx = headers.indexOf(LEAD_ID_COLUMN);
  if (idColIdx === -1) throw new Error('Lead ID column not found');

  for (let r = 1; r < data.length; r++) {
    if (String(data[r][idColIdx]) === String(leadId)) {
      Object.keys(fields).forEach(fieldKey => {
        const colIdx = headers.indexOf(fieldKey);
        if (colIdx !== -1) sheet.getRange(r + 1, colIdx + 1).setValue(fields[fieldKey]);
      });
      const lastActivityIdx = headers.indexOf('Last Activity');
      if (lastActivityIdx !== -1) sheet.getRange(r + 1, lastActivityIdx + 1).setValue(new Date().toISOString());

      // If Human Takeover changed, propagate it to Chat Control so the bot itself
      // actually stops/resumes replying (not just a label change in the CRM).
      if (Object.prototype.hasOwnProperty.call(fields, 'Human Takeover')) {
        const fbIdx = headers.indexOf('Facebook ID');
        const facebookId = fbIdx !== -1 ? String(data[r][fbIdx]) : '';
        const isTakeover = String(fields['Human Takeover']).toLowerCase() === 'true' ||
                            String(fields['Human Takeover']).toLowerCase() === 'yes';
        if (facebookId) {
          syncChatControlAssignedTo(ss, facebookId, isTakeover ? HUMAN_TAKEOVER_ASSIGNEE : '');
        }
      }
      break;
    }
  }
}

// Handles the CRM's AI/Human toggle switch (Pipeline card AND AI Conversations panel
// both call this same 'setTakeover' action via DataService.setHumanTakeover()).
function setTakeover(facebookId, mode) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const isHuman = String(mode).toLowerCase() === 'human';

  // This is the write that actually matters to n8n — Chat Control's "Assigned To"
  // must read exactly "Romeo" for the bot to pause.
  syncChatControlAssignedTo(ss, facebookId, isHuman ? HUMAN_TAKEOVER_ASSIGNEE : '');

  // Also reflect the change on the Leads sheet itself, so the CRM's own lead list/kanban
  // shows the correct AI/Human badge without needing a manual refresh of that row.
  const sheet = ss.getSheetByName(LEADS_SHEET_NAME);
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const fbColIdx = headers.indexOf('Facebook ID');
  if (fbColIdx === -1) return;

  for (let r = 1; r < data.length; r++) {
    if (String(data[r][fbColIdx]) === String(facebookId)) {
      const setIfColumnExists = (col, val) => {
        const idx = headers.indexOf(col);
        if (idx !== -1) sheet.getRange(r + 1, idx + 1).setValue(val);
      };
      setIfColumnExists('Human Takeover', isHuman);
      setIfColumnExists('Assigned To', isHuman ? 'Agent' : 'AI');
      setIfColumnExists('AI Status', isHuman ? 'Paused' : 'Active');
      setIfColumnExists('Last Activity', new Date().toISOString());
      break;
    }
  }
}

function syncChatControlAssignedTo(ss, facebookId, assignedToValue) {
  const ccSheet = ss.getSheetByName(CHAT_CONTROL_SHEET_NAME);
  if (!ccSheet) return; // sheet not set up yet — skip silently
  const ccData = ccSheet.getDataRange().getValues();
  const ccHeaders = ccData[0];
  const fbColIdx = ccHeaders.indexOf(CHAT_CONTROL_FB_COLUMN);
  let assignedColIdx = ccHeaders.indexOf(CHAT_CONTROL_ASSIGNED_COLUMN);
  if (fbColIdx === -1) return;

  // Add "Assigned To" column if it doesn't exist yet.
  if (assignedColIdx === -1) {
    assignedColIdx = ccHeaders.length;
    ccSheet.getRange(1, assignedColIdx + 1).setValue(CHAT_CONTROL_ASSIGNED_COLUMN);
  }

  for (let r = 1; r < ccData.length; r++) {
    if (String(ccData[r][fbColIdx]) === String(facebookId)) {
      ccSheet.getRange(r + 1, assignedColIdx + 1).setValue(assignedToValue);
      const lastUpdateIdx = ccHeaders.indexOf('Last Update');
      if (lastUpdateIdx !== -1) ccSheet.getRange(r + 1, lastUpdateIdx + 1).setValue(new Date().toISOString());
      return;
    }
  }
  // No existing Chat Control row for this customer yet — create one.
  const newRow = [];
  newRow[fbColIdx] = facebookId;
  newRow[assignedColIdx] = assignedToValue;
  const lastUpdateIdx = ccHeaders.indexOf('Last Update');
  if (lastUpdateIdx !== -1) newRow[lastUpdateIdx] = new Date().toISOString();
  ccSheet.appendRow(newRow);
}
