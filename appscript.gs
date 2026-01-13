// Ogilvy X Hogarth Tracker 2026 - Google Apps Script Backend
// Google Sheet: https://docs.google.com/spreadsheets/d/1XktaJJ-KuOur8fW5SeEWdBAbape27KRGhKnFFS8FBWA/edit

const SHEET_ID = '1XktaJJ-KuOur8fW5SeEWdBAbape27KRGhKnFFS8FBWA';
const SHEET_NAME = 'Awards';

function getSheet() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    // Set up headers
    sheet.getRange(1, 1, 1, 8).setValues([[
      'ID', 'Project', 'Client', 'Producer', 'Job', 'Date of Award', 'Amount', 'Date Recorded'
    ]]);
    sheet.getRange(1, 1, 1, 8).setFontWeight('bold');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function generateId() {
  return Utilities.getUuid();
}

function doGet(e) {
  try {
    const sheet = getSheet();
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const entries = [];
    
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (row[0]) { // If ID exists
        entries.push({
          id: row[0],
          project: row[1] || '',
          client: row[2] || '',
          producer: row[3] || '',
          job: row[4] || '',
          date: row[5] ? Utilities.formatDate(new Date(row[5]), Session.getScriptTimeZone(), 'yyyy-MM-dd') : '',
          amount: row[6] || 0,
          dateOfRecording: row[7] ? Utilities.formatDate(new Date(row[7]), Session.getScriptTimeZone(), 'yyyy-MM-dd') : ''
        });
      }
    }
    
    return ContentService.createTextOutput(JSON.stringify({ entries: entries }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({ error: error.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const action = data.action;
    const payload = data.payload;
    
    const sheet = getSheet();
    
    if (action === 'add') {
      const id = generateId();
      const dateOfRecording = payload.dateOfRecording || Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
      sheet.appendRow([
        id,
        payload.project,
        payload.client,
        payload.producer,
        payload.job,
        payload.date,
        payload.amount,
        dateOfRecording
      ]);
      return ContentService.createTextOutput(JSON.stringify({ success: true, id: id }))
        .setMimeType(ContentService.MimeType.JSON);
    }
    
    if (action === 'update') {
      const dataRange = sheet.getDataRange();
      const values = dataRange.getValues();
      const idCol = 0;
      
      for (let i = 1; i < values.length; i++) {
        if (values[i][idCol] === payload.id) {
          sheet.getRange(i + 1, 2, 1, 6).setValues([[
            payload.project,
            payload.client,
            payload.producer,
            payload.job,
            payload.date,
            payload.amount
          ]]);
          if (payload.dateOfRecording) {
            sheet.getRange(i + 1, 8).setValue(payload.dateOfRecording);
          }
          return ContentService.createTextOutput(JSON.stringify({ success: true }))
            .setMimeType(ContentService.MimeType.JSON);
        }
      }
      return ContentService.createTextOutput(JSON.stringify({ error: 'Entry not found' }))
        .setMimeType(ContentService.MimeType.JSON);
    }
    
    if (action === 'delete') {
      const dataRange = sheet.getDataRange();
      const values = dataRange.getValues();
      const idCol = 0;
      
      for (let i = 1; i < values.length; i++) {
        if (values[i][idCol] === payload.id) {
          sheet.deleteRow(i + 1);
          return ContentService.createTextOutput(JSON.stringify({ success: true }))
            .setMimeType(ContentService.MimeType.JSON);
        }
      }
      return ContentService.createTextOutput(JSON.stringify({ error: 'Entry not found' }))
        .setMimeType(ContentService.MimeType.JSON);
    }
    
    if (action === 'clear') {
      const lastRow = sheet.getLastRow();
      if (lastRow > 1) {
        sheet.deleteRows(2, lastRow - 1);
      }
      return ContentService.createTextOutput(JSON.stringify({ success: true }))
        .setMimeType(ContentService.MimeType.JSON);
    }
    
    return ContentService.createTextOutput(JSON.stringify({ error: 'Unknown action' }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({ error: error.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}
