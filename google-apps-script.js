// ============================================================
// BEEMAS — Google Apps Script Backend
// Paste into Apps Script editor
// Deploy: Execute as Me, Access: Anyone
// ============================================================

function doGet(e) {
  var action   = e.parameter.action;
  var callback = e.parameter.callback;
  var result;

  if      (action === 'register')            result = registerPlayer(e.parameter.name, e.parameter.phone);
  else if (action === 'updateSession')       result = updateSession(e.parameter.phone, e.parameter.total, e.parameter.failed, e.parameter.success, e.parameter.prize);
  else if (action === 'getPrizes')           result = getPrizes();
  else if (action === 'claimPrize')          result = claimPrize(e.parameter.prizeId, e.parameter.phone);
  else if (action === 'getPlayers')          result = getPlayers();
  else if (action === 'getSettings')         result = getSettings();
  else if (action === 'setPlayRestriction')  result = setPlayRestriction(e.parameter.value);
  else if (action === 'setPrizeQty')         result = setPrizeQty(e.parameter.prizeId, e.parameter.qty);
  else result = { status: 'ok', message: 'Beemas API running' };

  var output = JSON.stringify(result);
  if (callback) {
    return ContentService.createTextOutput(callback + '(' + output + ')')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService.createTextOutput(output)
    .setMimeType(ContentService.MimeType.JSON);
}

// ── helpers ──────────────────────────────────────────────────────

function getMaxPlays() {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Settings');
    if (!sheet) return 0;
    var data = sheet.getDataRange().getValues();
    for (var i = 0; i < data.length; i++) {
      if (data[i][0] === 'play_restriction') return Number(data[i][1]) || 0;
    }
    return 0;
  } catch(e) { return 0; }
}

function countSessions(phone, data) {
  var count = 0;
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][4]) === String(phone)) count++;
  }
  return count;
}

// ── player functions ──────────────────────────────────────────────

function registerPlayer(name, phone) {
  try {
    var ss    = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('Players') || ss.insertSheet('Players');

    if (sheet.getLastRow() === 0) {
      sheet.appendRow(['Timestamp','Date','Time','Name','Phone','Total Attempts','Failed','Success','Prize Won']);
      sheet.getRange(1,1,1,9).setFontWeight('bold').setBackground('#e91e8c').setFontColor('#fff0f8');
      sheet.setColumnWidths(1, 9, 110);
    }

    var maxPlays  = getMaxPlays();
    var allData   = sheet.getLastRow() > 1 ? sheet.getDataRange().getValues() : [[]];
    var playCount = countSessions(phone, allData);

    if (maxPlays > 0 && playCount >= maxPlays) {
      return { status: 'restricted', playCount: playCount, maxPlays: maxPlays };
    }

    var now     = new Date();
    var dateStr = Utilities.formatDate(now, Session.getScriptTimeZone(), 'dd/MM/yyyy');
    var timeStr = Utilities.formatDate(now, Session.getScriptTimeZone(), 'hh:mm a');

    sheet.appendRow([now.toISOString(), dateStr, timeStr, name, phone, 0, 0, 0, 'No Gift']);

    return { status: 'ok', playCount: playCount + 1, maxPlays: maxPlays };
  } catch (err) { return { status: 'error', message: err.toString() }; }
}

function updateSession(phone, total, failed, success, prize) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Players');
    if (!sheet) return { status: 'error', message: 'Players sheet not found' };
    var data = sheet.getDataRange().getValues();
    for (var i = data.length - 1; i >= 1; i--) {
      if (String(data[i][4]) === String(phone)) {
        sheet.getRange(i+1, 6).setValue(Number(total)   || 0);
        sheet.getRange(i+1, 7).setValue(Number(failed)  || 0);
        sheet.getRange(i+1, 8).setValue(Number(success) || 0);
        if (prize) sheet.getRange(i+1, 9).setValue(prize);
        return { status: 'ok' };
      }
    }
    return { status: 'error', message: 'Session not found' };
  } catch (err) { return { status: 'error', message: err.toString() }; }
}

function getPlayers() {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Players');
    if (!sheet) return { players: [] };
    var data = sheet.getDataRange().getValues();
    return { players: data.slice(1) };
  } catch (err) { return { status: 'error', message: err.toString() }; }
}

// ── prizes ────────────────────────────────────────────────────────

function getPrizes() {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Prizes');
    if (!sheet) return { prizes: [] };
    var data = sheet.getDataRange().getValues();
    var prizes = [];
    for (var i = 1; i < data.length; i++) {
      prizes.push({ id: data[i][0], icon: data[i][1], name: data[i][2], quantity: Number(data[i][3]) });
    }
    return { prizes: prizes };
  } catch (err) { return { status: 'error', message: err.toString() }; }
}

function claimPrize(prizeId, phone) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Prizes');
    if (!sheet) return { status: 'error', message: 'Prizes sheet not found' };
    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] == prizeId) {
        var qty = Number(data[i][3]);
        if (qty <= 0) return { status: 'out', message: 'Prize out of stock.' };
        sheet.getRange(i+1, 4).setValue(qty - 1);
        return { status: 'ok', prize: data[i][2] };
      }
    }
    return { status: 'error', message: 'Prize not found.' };
  } catch (err) { return { status: 'error', message: err.toString() }; }
}

function setPrizeQty(prizeId, qty) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Prizes');
    if (!sheet) return { status: 'error', message: 'Prizes sheet not found' };
    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] == prizeId) {
        sheet.getRange(i+1, 4).setValue(Number(qty));
        return { status: 'ok' };
      }
    }
    return { status: 'error', message: 'Prize not found' };
  } catch (err) { return { status: 'error', message: err.toString() }; }
}

// ── settings ──────────────────────────────────────────────────────

function getSettings() {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Settings');
    if (!sheet) return { play_restriction: 0 };
    var data = sheet.getDataRange().getValues();
    var out  = { play_restriction: 0 };
    for (var i = 0; i < data.length; i++) {
      if (data[i][0] === 'play_restriction') out.play_restriction = Number(data[i][1]) || 0;
    }
    return out;
  } catch(e) { return { play_restriction: 0 }; }
}

function setPlayRestriction(value) {
  try {
    var ss    = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('Settings') || ss.insertSheet('Settings');
    var data  = sheet.getLastRow() > 0 ? sheet.getDataRange().getValues() : [];
    for (var i = 0; i < data.length; i++) {
      if (data[i][0] === 'play_restriction') {
        sheet.getRange(i+1, 2).setValue(Number(value) || 0);
        return { status: 'ok' };
      }
    }
    sheet.appendRow(['play_restriction', Number(value) || 0]);
    return { status: 'ok' };
  } catch(e) { return { status: 'error', message: e.toString() }; }
}

// ============================================================
// RUN ONCE — in Apps Script editor: run setupPrizesSheet first,
// then run setupSettingsSheet
// ============================================================

function setupPrizesSheet(ss) {
  if (!ss) ss = SpreadsheetApp.getActiveSpreadsheet();
  var existing = ss.getSheetByName('Prizes');
  if (existing) ss.deleteSheet(existing);
  var sheet = ss.insertSheet('Prizes');
  sheet.appendRow(['ID','Icon','Prize Name','Quantity']);
  sheet.getRange(1,1,1,4).setFontWeight('bold').setBackground('#e91e8c').setFontColor('#fff0f8');
  [
    ['P1','🍦','₹30 Off Ice Cream',  999],
    ['P2','🍔','50% Off Burger',      999],
    ['P3','🍕','25% Off Pizza',       999],
    ['P4','🍛','50% Off Biryani',     999],
    ['P5','🥛','Free Badam Milk',     999],
  ].forEach(function(r){ sheet.appendRow(r); });
  sheet.setColumnWidths(1, 4, 140);
  Logger.log('Prizes sheet created — 5 prizes.');
}

function setupSettingsSheet(ss) {
  if (!ss) ss = SpreadsheetApp.getActiveSpreadsheet();
  var existing = ss.getSheetByName('Settings');
  if (existing) ss.deleteSheet(existing);
  var sheet = ss.insertSheet('Settings');
  sheet.appendRow(['Key','Value']);
  sheet.getRange(1,1,1,2).setFontWeight('bold').setBackground('#e91e8c').setFontColor('#fff0f8');
  sheet.appendRow(['play_restriction', 0]);
  sheet.setColumnWidth(1, 160); sheet.setColumnWidth(2, 100);
  Logger.log('Settings sheet created. play_restriction = 0 (unlimited).');
}
