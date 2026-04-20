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
  var headers = ['Vaccin / Évènement', 'Type', 'Délai rappel (mois)', 'Description'];
  sh.getRange(1, 1, 1, 4).setValues([headers]).setFontWeight('bold');
  var data = [
    ['DTP 25',        'age',      0,  'Obligatoire à 25 ans'],
    ['DTP 45',        'age',      0,  'Obligatoire à 45 ans'],
    ['DTP 65',        'age',      0,  'Obligatoire à 65 ans'],
    ['HEP B',         'presence', 0,  'Doit être OK'],
    ['HEP A 1',       'date',     0,  'Première dose hépatite A'],
    ['HEP A 2',       'rappel',  12,  'Rappel 6-12 mois après dose 1'],
    ['Fièvre jaune 1','date',     0,  'Première dose fièvre jaune'],
    ['Fièvre jaune 2','rappel', 120,  'Rappel 10 ans après dose 1'],
    ['Typhoïde',      'rappel',  36,  'Rappel tous les 3 ans'],
    ['Méningo ACYW135','rappel', 60,  'Rappel tous les 5 ans'],
    ['ROR',           'presence', 0,  'Doit être OK'],
    ['Grippe',        'rappel',  12,  'Rappel tous les ans'],
    ['BCG',           'presence', 0,  'Doit être OK ou manque'],
    ['Groupe sanguin','presence', 0,  'Doit être renseigné'],
    ['VMA',           'rappel',  12,  'Valide 1 an – alerte à 11 mois'],
    ['ECG de la VMA', 'date',     0,  'Date ECG'],
    ['Pano dentaire', 'rappel', 120,  'Rappel tous les 10 ans'],
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

/* ═══════════════════════════════════════════════════
   BILAN PAR MAIL – tous les 15 jours, vendredi 15h30
   ═══════════════════════════════════════════════════ */

/** Installe le trigger toutes les 2 semaines vendredi 15h30.
 *  À exécuter UNE SEULE FOIS manuellement. */
function installBilanTrigger() {
  // Supprime les anciens triggers de cette fonction
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === 'sendBilanEmail') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('sendBilanEmail')
    .timeBased()
    .onWeekDay(ScriptApp.WeekDay.FRIDAY)
    .atHour(15)
    .nearMinute(30)
    .everyWeeks(2)
    .inTimezone('Europe/Paris')
    .create();
  Logger.log('Trigger bilan INSARAG installé : vendredi 15h30, toutes les 2 semaines.');
}

/* ── helpers serveur ── */
function _parseDate(s) {
  if (!s) return null;
  if (s instanceof Date) return isNaN(s.getTime()) ? null : s;
  s = String(s).trim();
  var m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
  if (m) return new Date(parseInt(m[3]), parseInt(m[2]) - 1, parseInt(m[1]));
  var d = new Date(s);
  return isNaN(d.getTime()) ? null : d;
}
function _addMonths(d, m) { var r = new Date(d); r.setMonth(r.getMonth() + m); return r; }
function _age(bd) {
  if (!bd) return null;
  var t = new Date(), a = t.getFullYear() - bd.getFullYear(), m = t.getMonth() - bd.getMonth();
  if (m < 0 || (m === 0 && t.getDate() < bd.getDate())) a--;
  return a;
}
function _fmtDate(d) { return d ? ('0' + d.getDate()).slice(-2) + '/' + ('0' + (d.getMonth() + 1)).slice(-2) + '/' + d.getFullYear() : ''; }

/* ── analyse serveur (miroir de analyzeAgent côté client) ── */
function _analyzeServer() {
  var d = getData();
  var rulesMap = {};
  d.rules.forEach(function(r) { rulesMap[r.name] = r; });
  var now = new Date();
  var retards = [], upcoming = [];

  d.agents.forEach(function(agent) {
    var f = agent.fields;
    var bd = _parseDate(f['Date naissance']), agentAge = _age(bd);

    // DTP 25/45/65
    [25, 45, 65].forEach(function(a) {
      var key = 'DTP ' + a, rule = rulesMap[key];
      if (!rule || agentAge === null || agentAge < a) return;
      var val = f[key], da = _parseDate(val);
      if (!da && (!val || val === '' || val === '0' || val === 'FALSE'))
        retards.push({ agent: agent.name, label: key, info: 'MANQUANT – obligatoire à ' + a + ' ans' });
    });

    // Présence simple : HEP B, ROR, BCG, Groupe sanguin, COVID
    [{ key: 'HEP B' }, { key: 'ROR' }, { key: 'BCG' }, { key: 'COVID (3 doses)', label: 'COVID' }].forEach(function(item) {
      var v = String(f[item.key] || '').toLowerCase();
      if (v !== 'ok') retards.push({ agent: agent.name, label: item.label || item.key, info: 'MANQUANT' });
    });
    (function() {
      var v = String(f['Groupe sanguin'] || '').trim();
      if (!v || v === '0' || v.toLowerCase() === 'false')
        retards.push({ agent: agent.name, label: 'Groupe sanguin', info: 'MANQUANT' });
    })();

    // HEP A 1
    (function() {
      var d1 = _parseDate(f['HEP A 1']);
      if (!d1) retards.push({ agent: agent.name, label: 'HEP A 1', info: 'MANQUANT' });
    })();

    // Fièvre jaune 1
    (function() {
      var d1 = _parseDate(f['Fièvre jaune 1']);
      if (!d1) retards.push({ agent: agent.name, label: 'Fièvre jaune 1', info: 'MANQUANT' });
    })();

    // Rappels avec délai en mois
    var rappels = [
      { key: 'HEP A 2', src: 'HEP A 1', def: 12 },
      { key: 'Fièvre jaune 2', src: 'Fièvre jaune 1', def: 120 },
      { key: 'Typhoïde', src: 'Typhoïde', def: 36, self: true },
      { key: 'Méningo ACYW135', src: 'Méningo ACYW135', def: 60, self: true, label: 'Méningo' },
      { key: 'Grippe', src: 'Grippe', def: 12, self: true, okVal: true },
      { key: 'VMA', src: 'VMA', def: 12, self: true },
      { key: 'Pano dentaire', src: 'Pano dentaire', def: 120, self: true }
    ];
    rappels.forEach(function(rp) {
      var rule = rulesMap[rp.key], delayM = rule ? rule.delay : rp.def;
      var lbl = rp.label || rp.key;
      if (rp.okVal) {
        var v = String(f[rp.key] || '').toLowerCase();
        if (v === 'ok') return;
      }
      var srcDate = _parseDate(f[rp.src]) || _parseDate(f[rp.src + ' ']);
      if (!srcDate) {
        if (rp.self) retards.push({ agent: agent.name, label: lbl, info: 'MANQUANT' });
        return;
      }
      if (!rp.self) {
        var d2 = _parseDate(f[rp.key]);
        if (d2) return; // déjà fait
      }
      var due = _addMonths(srcDate, delayM);
      var nextMonth = _addMonths(now, 1);
      if (now > due) {
        retards.push({ agent: agent.name, label: lbl, info: 'EN RETARD depuis le ' + _fmtDate(due) });
      } else if (due <= nextMonth) {
        upcoming.push({ agent: agent.name, label: lbl, info: 'Échéance le ' + _fmtDate(due) });
      }
    });
  });

  return { retards: retards, upcoming: upcoming };
}

/* ── envoi du mail bilan ── */
function sendBilanEmail() {
  var result = _analyzeServer();
  var retards = result.retards, upcoming = result.upcoming;
  var dest = ['cecile.verges@sdis66.fr', 'florian.bois@sdis66.fr', 'eve.laparra@sdis66.fr', 'brice.dubrey@sdis66.fr'];

  var subject = '📋 Bilan Suivi INSARAG – ' + _fmtDate(new Date());

  var html = '<div style="font-family:Arial,sans-serif;max-width:700px;">';
  html += '<h2 style="color:#1565c0;">📋 Bilan du Suivi INSARAG</h2>';
  html += '<p style="color:#555;">Rapport automatique généré le <b>' + _fmtDate(new Date()) + '</b></p>';

  // Retards
  html += '<h3 style="color:#d50000;">🔴 Examens en retard : ' + retards.length + '</h3>';
  if (retards.length > 0) {
    html += '<table style="border-collapse:collapse;width:100%;font-size:13px;">';
    html += '<tr style="background:#d50000;color:#fff;"><th style="padding:6px 10px;text-align:left;">Agent</th><th style="padding:6px 10px;text-align:left;">Examen</th><th style="padding:6px 10px;text-align:left;">Détail</th></tr>';
    retards.forEach(function(r, i) {
      var bg = i % 2 === 0 ? '#fff5f5' : '#ffffff';
      html += '<tr style="background:' + bg + ';"><td style="padding:5px 10px;">' + r.agent + '</td><td style="padding:5px 10px;">' + r.label + '</td><td style="padding:5px 10px;">' + r.info + '</td></tr>';
    });
    html += '</table>';
  } else {
    html += '<p style="color:#0277bd;">✅ Aucun examen en retard !</p>';
  }

  // À venir dans le mois
  html += '<h3 style="color:#ff6f00;">🟠 Échéances dans le mois à venir : ' + upcoming.length + '</h3>';
  if (upcoming.length > 0) {
    html += '<table style="border-collapse:collapse;width:100%;font-size:13px;">';
    html += '<tr style="background:#ff6f00;color:#fff;"><th style="padding:6px 10px;text-align:left;">Agent</th><th style="padding:6px 10px;text-align:left;">Examen</th><th style="padding:6px 10px;text-align:left;">Détail</th></tr>';
    upcoming.forEach(function(r, i) {
      var bg = i % 2 === 0 ? '#fff8e1' : '#ffffff';
      html += '<tr style="background:' + bg + ';"><td style="padding:5px 10px;">' + r.agent + '</td><td style="padding:5px 10px;">' + r.label + '</td><td style="padding:5px 10px;">' + r.info + '</td></tr>';
    });
    html += '</table>';
  } else {
    html += '<p style="color:#0277bd;">✅ Aucune échéance prévue dans le mois.</p>';
  }

  html += '<hr style="margin-top:20px;border:none;border-top:1px solid #ddd;">';
  html += '<p style="color:#999;font-size:11px;">Ce mail est envoyé automatiquement par le Suivi INSARAG tous les 15 jours.</p>';
  html += '</div>';

  MailApp.sendEmail({
    to: dest.join(','),
    subject: subject,
    htmlBody: html
  });

  Logger.log('Bilan INSARAG envoyé à ' + dest.join(', '));
}
