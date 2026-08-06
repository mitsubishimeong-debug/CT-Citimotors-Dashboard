/**
 * Citimotors AI CRM — Google Apps Script Web App
 * Deploy: Extensions > Apps Script (sa parehong Google Sheet) > paste this file >
 *   Deploy > New deployment > type "Web app" > Execute as "Me" > Who has access "Anyone" >
 *   Deploy, then copy the Web App URL into the CRM's Sync (QR) tab ("Write endpoint").
 *
 * Handles: updating a lead's Pipeline Stage / AI Status / Human Takeover / Assigned To
 * from the CRM (drag-and-drop pipeline, human-takeover toggle) so the Google Sheet used
 * by the n8n Messenger workflow always reflects what's shown in the CRM.
 *
 * IMPORTANT: When "Human Takeover" is toggled from the CRM, this ALSO updates the
 * "Chat Control" tab's "Assigned To" column (matched by Facebook ID) — that's the
 * value the n8n workflow's "Check Assigned To" node actually reads to decide whether
 * the AI should keep auto-replying. Without this, the CRM toggle only changed a label
 * in the Leads tab and never actually paused the bot.
 */
const LEADS_SHEET_NAME = 'Leads';
const LEAD_ID_COLUMN = 'Lead ID';
const CHAT_CONTROL_SHEET_NAME = 'Chat Control';
const CHAT_CONTROL_FB_COLUMN = 'Facebook ID';
const CHAT_CONTROL_ASSIGNED_COLUMN = 'Assigned To';
// Must match the value n8n's "Check Assigned To" / "Check Assigned To1" nodes compare against.
const HUMAN_TAKEOVER_ASSIGNEE = 'Romeo';

function doPost(e) {
  const payload = JSON.parse(e.postData.contents);
  if (payload.action === 'updateLead') {
    updateLead(payload.leadId, payload.fields);
    return ContentService.createTextOutput(JSON.stringify({ ok: true })).setMimeType(ContentService.MimeType.JSON);
  }
  return ContentService.createTextOutput(JSON.stringify({ ok: false, error: 'Unknown action' })).setMimeType(ContentService.MimeType.JSON);
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
