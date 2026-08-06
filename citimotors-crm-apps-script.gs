/**
 * Citimotors AI CRM — Google Apps Script Web App
 * Deploy: Extensions > Apps Script (sa parehong Google Sheet) > paste this file >
 *   Deploy > New deployment > type "Web app" > Execute as "Me" > Who has access "Anyone" >
 *   Deploy, then copy the Web App URL into the CRM's Sync (QR) tab ("Write endpoint").
 *
 * Handles: updating a lead's Pipeline Stage / AI Status / Human Takeover / Assigned To
 * from the CRM (drag-and-drop pipeline, human-takeover toggle) so the Google Sheet used
 * by the n8n Messenger workflow always reflects what's shown in the CRM.
 */
const LEADS_SHEET_NAME = 'Leads';
const LEAD_ID_COLUMN = 'Lead ID';

function doPost(e) {
  const payload = JSON.parse(e.postData.contents);
  if (payload.action === 'updateLead') {
    updateLead(payload.leadId, payload.fields);
    return ContentService.createTextOutput(JSON.stringify({ ok: true })).setMimeType(ContentService.MimeType.JSON);
  }
  return ContentService.createTextOutput(JSON.stringify({ ok: false, error: 'Unknown action' })).setMimeType(ContentService.MimeType.JSON);
}

function updateLead(leadId, fields) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(LEADS_SHEET_NAME);
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
      break;
    }
  }
}
