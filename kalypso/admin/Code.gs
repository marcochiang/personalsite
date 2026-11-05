// ══════════════════════════════════════════════════════════════
// Kalypso Menu Admin — Google Apps Script (Code.gs)
// ══════════════════════════════════════════════════════════════
// Deploy as Web App: Execute as "Me", Access "Anyone"
// Copy the deployment URL into kalypso/index.html and
// kalypso/admin/index.html (the API_URL constant).
// ══════════════════════════════════════════════════════════════

const ADMIN_PASSWORD = 'CHANGE_ME'; // Set your admin password here
const SHEET_ID = '';                // Set your Google Sheet ID here

// ── Helpers ──

function getSheet(name) {
  return SpreadsheetApp.openById(SHEET_ID).getSheetByName(name);
}

function sheetToObjects(sheet) {
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return [];
  const headers = data[0];
  return data.slice(1).map(row => {
    const obj = {};
    headers.forEach((h, i) => { obj[h] = row[i]; });
    return obj;
  });
}

function findRowById(sheet, id) {
  const data = sheet.getDataRange().getValues();
  const idCol = data[0].indexOf('id');
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][idCol]) === String(id)) return i + 1; // 1-indexed
  }
  return -1;
}

function getHeaders(sheet) {
  return sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
}

function jsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// ── GET endpoint ──

function doGet(e) {
  const action = (e && e.parameter && e.parameter.action) || 'getAll';

  if (action === 'getAll') {
    return jsonResponse(getAllNested());
  }
  if (action === 'getRaw') {
    return jsonResponse(getRawTables());
  }
  return jsonResponse({ error: 'Unknown action' });
}

// Returns data nested like the original JS structure
function getAllNested() {
  const cats = sheetToObjects(getSheet('Categories')).sort((a, b) => a.sort_order - b.sort_order);
  const items = sheetToObjects(getSheet('Items')).sort((a, b) => a.sort_order - b.sort_order);
  const cards = sheetToObjects(getSheet('CuriosityCards')).sort((a, b) => a.sort_order - b.sort_order);

  const categories = cats.map(cat => ({
    id: cat.id,
    name: cat.name,
    items: items
      .filter(i => i.category_id === cat.id)
      .map(i => ({ id: i.id, title: i.title, desc: i.desc }))
  }));

  const curiosityCards = cards.map(c => c.text);

  return { categories, curiosityCards };
}

// Returns flat tables for admin dashboard
function getRawTables() {
  return {
    categories: sheetToObjects(getSheet('Categories')),
    items: sheetToObjects(getSheet('Items')),
    curiosityCards: sheetToObjects(getSheet('CuriosityCards')),
  };
}

// ── POST endpoint ──

function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);
    const { action, sheet: sheetName, password, data } = body;

    // Verify action
    if (action === 'verify') {
      return jsonResponse({ success: password === ADMIN_PASSWORD });
    }

    // All other actions require password
    if (password !== ADMIN_PASSWORD) {
      return jsonResponse({ error: 'Unauthorized' });
    }

    if (!sheetName) {
      return jsonResponse({ error: 'Missing sheet name' });
    }

    const sheet = getSheet(sheetName);
    if (!sheet) {
      return jsonResponse({ error: 'Sheet not found: ' + sheetName });
    }

    switch (action) {
      case 'add':    return jsonResponse(addRow(sheet, data));
      case 'update': return jsonResponse(updateRow(sheet, data));
      case 'delete': return jsonResponse(deleteRow(sheet, data));
      case 'reorder': return jsonResponse(reorderRows(sheet, data));
      default:       return jsonResponse({ error: 'Unknown action: ' + action });
    }
  } catch (err) {
    return jsonResponse({ error: err.toString() });
  }
}

function addRow(sheet, data) {
  const headers = getHeaders(sheet);
  const row = headers.map(h => data[h] !== undefined ? data[h] : '');
  sheet.appendRow(row);
  return { success: true };
}

function updateRow(sheet, data) {
  const rowNum = findRowById(sheet, data.id);
  if (rowNum === -1) return { error: 'Row not found: ' + data.id };

  const headers = getHeaders(sheet);
  const currentRow = sheet.getRange(rowNum, 1, 1, headers.length).getValues()[0];

  headers.forEach((h, i) => {
    if (data[h] !== undefined && data[h] !== null) {
      currentRow[i] = data[h];
    }
  });

  sheet.getRange(rowNum, 1, 1, headers.length).setValues([currentRow]);
  return { success: true };
}

function deleteRow(sheet, data) {
  const rowNum = findRowById(sheet, data.id);
  if (rowNum === -1) return { error: 'Row not found: ' + data.id };
  sheet.deleteRow(rowNum);
  return { success: true };
}

function reorderRows(sheet, data) {
  const order = data.order; // [{ id, sort_order }, ...]
  order.forEach(entry => {
    const rowNum = findRowById(sheet, entry.id);
    if (rowNum === -1) return;
    const headers = getHeaders(sheet);
    const sortCol = headers.indexOf('sort_order') + 1;
    if (sortCol > 0) {
      sheet.getRange(rowNum, sortCol).setValue(entry.sort_order);
    }
  });
  return { success: true };
}
