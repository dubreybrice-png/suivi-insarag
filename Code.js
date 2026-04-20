/* ═══════════════════════════════════════════════════
   SUIVI INSARAG – Code.js
   ═══════════════════════════════════════════════════ */

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Suivi INSARAG')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/* ── helpers ── */
function getSS() {
  return SpreadsheetApp.getActiveSpreadsheet();
}

function ensureRulesSheet() {
  var ss = getSS();
  var sh = ss.getSheetByName('règle médicale');
  if (sh) return sh;
  sh = ss.insertSheet('règle médicale');
  var headers = ['Vaccin / Évènement', 'Type', 'Délai rappel (années)', 'Description'];
  sh.getRange(1, 1, 1, 4).setValues([headers]).setFontWeight('bold');
  var data = [
    ['DTP 25',        'age',      0,  'Obligatoire à 25 ans'],
    ['DTP 45',        'age',      0,  'Obligatoire à 45 ans'],
    ['DTP 65',        'age',      0,  'Obligatoire à 65 ans'],
    ['HEP B',         'presence', 0,  'Doit être OK'],
    ['HEP A 1',       'date',     0,  'Première dose hépatite A'],
    ['HEP A 2',       'rappel',   1,  'Rappel 6-12 mois après dose 1 – on prévoit à 1 an'],
    ['Fièvre jaune 1','date',     0,  'Première dose fièvre jaune'],
    ['Fièvre jaune 2','rappel',  10,  'Rappel 10 ans après dose 1'],
    ['Typhoïde',      'rappel',   3,  'Rappel tous les 3 ans'],
    ['Méningo ACYW135','rappel',  5,  'Rappel tous les 5 ans'],
    ['ROR',           'presence', 0,  'Doit être OK'],
    ['Grippe',        'rappel',   1,  'Rappel tous les ans'],
    ['BCG',           'presence', 0,  'Doit être OK ou manque'],
    ['Groupe sanguin','presence', 0,  'Doit être renseigné'],
    ['VMA',           'rappel',   1,  'Valide 1 an – alerte à 11 mois'],
    ['ECG de la VMA', 'date',     0,  'Date ECG'],
    ['Pano dentaire', 'rappel',  10,  'Rappel tous les 10 ans'],
    ['COVID (3 doses)','presence',0,  'Doit être OK ou manque']
  ];
  sh.getRange(2, 1, data.length, 4).setValues(data);
  sh.autoResizeColumns(1, 4);
  return sh;
}

function ensureCommentsSheet() {
  var ss = getSS();
  var sh = ss.getSheetByName('Commentaires');
  if (sh) return sh;
  sh = ss.insertSheet('Commentaires');
  sh.getRange(1, 1, 1, 4).setValues([['Agent', 'Commentaire', 'Date', 'Validé']]).setFontWeight('bold');
  return sh;
}

/* ── lecture données ── */
function getData() {
  var ss = getSS();
  var sh = ss.getSheetByName('Données');
  if (!sh) return { agents: [], rules: [] };
  var data = sh.getDataRange().getValues();
  // Row labels are in col A (index 0), agents start col B (index 1)
  var labels = [];
  for (var r = 0; r < data.length; r++) labels.push(String(data[r][0]).trim());

  // Build label→row map (1-based for Sheets)
  var labelRows = {};
  for (var r = 0; r < labels.length; r++) { if (labels[r]) labelRows[labels[r]] = r + 1; }

  var agents = [];
  for (var c = 1; c < data[0].length; c++) {
    var name = String(data[0][c] || '').trim();
    if (!name) continue;
    var agent = { name: name, col: c + 1, fields: {} };
    for (var r = 1; r < data.length; r++) {
      var label = labels[r];
      var val = data[r][c];
      if (val instanceof Date) {
        agent.fields[label] = Utilities.formatDate(val, 'Europe/Paris', 'dd/MM/yyyy');
      } else {
        agent.fields[label] = String(val || '').trim();
      }
    }
    agents.push(agent);
  }

  // rules
  var rsh = ensureRulesSheet();
  var rdata = rsh.getDataRange().getValues();
  var rules = [];
  for (var r = 1; r < rdata.length; r++) {
    rules.push({
      name: String(rdata[r][0]).trim(),
      type: String(rdata[r][1]).trim(),
      delay: Number(rdata[r][2]) || 0,
      desc: String(rdata[r][3] || '').trim()
    });
  }

  // comments
  var csh = ensureCommentsSheet();
  var cdata = csh.getDataRange().getValues();
  var comments = {};
  for (var r = 1; r < cdata.length; r++) {
    var ag = String(cdata[r][0]).trim();
    if (!ag) continue;
    if (!comments[ag]) comments[ag] = [];
    comments[ag].push({
      row: r + 1,
      text: String(cdata[r][1] || ''),
      date: cdata[r][2] instanceof Date ? Utilities.formatDate(cdata[r][2], 'Europe/Paris', 'dd/MM/yyyy HH:mm') : String(cdata[r][2] || ''),
      validated: String(cdata[r][3]).toLowerCase() === 'true' || cdata[r][3] === true
    });
  }

  return { agents: agents, rules: rules, comments: comments, labelRows: labelRows };
}

/* ── mise à jour cellule ── */
function updateCell(col, rowLabel, newValue) {
  var ss = getSS();
  var sh = ss.getSheetByName('Données');
  if (!sh) throw new Error('Onglet Données introuvable');
  var data = sh.getDataRange().getValues();
  var targetRow = -1;
  for (var r = 0; r < data.length; r++) {
    if (String(data[r][0]).trim() === rowLabel) { targetRow = r + 1; break; }
  }
  if (targetRow < 0) throw new Error('Ligne "' + rowLabel + '" introuvable');
  // Parse dd/MM/yyyy as Date
  var m = String(newValue).match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
  var val = m ? new Date(parseInt(m[3]), parseInt(m[2])-1, parseInt(m[1])) : newValue;
  sh.getRange(targetRow, col).setValue(val);
  return true;
}

/* ── sauvegarde règles ── */
function saveRules(rulesArray) {
  var sh = ensureRulesSheet();
  // Clear data rows
  var last = sh.getLastRow();
  if (last > 1) sh.getRange(2, 1, last - 1, 4).clearContent();
  if (rulesArray.length > 0) {
    var rows = rulesArray.map(function(r) { return [r.name, r.type, r.delay, r.desc]; });
    sh.getRange(2, 1, rows.length, 4).setValues(rows);
  }
  return true;
}

/* ── commentaires ── */
function addComment(agentName, text) {
  var sh = ensureCommentsSheet();
  sh.appendRow([agentName, text, new Date(), false]);
  return true;
}

function validateComment(row) {
  var sh = ensureCommentsSheet();
  sh.getRange(row, 4).setValue(true);
  return true;
}

function deleteComment(row) {
  var sh = ensureCommentsSheet();
  sh.deleteRow(row);
  return true;
}
