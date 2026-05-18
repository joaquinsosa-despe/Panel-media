// https://script.google.com/a/macros/despegar.com/s/AKfycbwu2hQiExuzPkzDayrkxPluho3kaiDzc9Aw7AbPnBNSYmY84BlBDAHOpJo3HRj6C7MPvA/exec 
// ============================================================
//  DESPEGAR MEDIA REVENUE — AUTO GOOGLE SLIDES
//  Waterfall vertical + Paleta corporativa + Tendencia dinámica
// ============================================================


// -----------------------------------------
// CONFIGURACIÓN — EDITÁ ESTOS VALORES
// -----------------------------------------
const CONFIG = {
  TARGET_MONTH : "March",
  TARGET_YEAR  : "2026",
  DATASET_SHEET: "dataset",
  HELPER_SHEET : "_DATA_HELPER",
  PRES_TITLE   : "Media Sales Despegar",
  FIELD_COLUMNS: "Actuals + RR",
  FIELD_LINE   : "Budget + Forecast + Breakthrough FY26",
};


const LOB_MAP = {
  "B2C-SITE"       : "B2C",
  "B2C-APP"        : "B2C",
  "B2C-CallCenter" : "B2C",
  "B2C-OFF"        : "B2C",
  "B2B"            : "B2B",
  "B2B2C"          : "B2B2C",
  "S/C"            : "Acuerdos sin canal",
};
const LOB_ORDER = ["B2C", "B2B", "B2B2C", "Acuerdos sin canal"];


const MONTH_ORDER = [
  "January","February","March","April","May","June",
  "July","August","September","October","November","December"
];


const C = {
  purple      : "#6B35A8",
  purpleDark  : "#4A1F7A",
  purpleLight : "#EDE7F6",
  green       : "#1E8449",
  greenLight  : "#E9F7EF",
  red         : "#C0392B",
  redLight    : "#FDEDEC",
  blue        : "#1565C0",
  gray        : "#9E9E9E",
  grayLight   : "#F5F5F5",
  white       : "#FFFFFF",
  black       : "#1A1A1A",
  subText     : "#666666",
};

// =============================================================================
// HUB TASKS — Sección de tareas pendientes por equipo
// =============================================================================

var HUB_EDITORS = [
  'matias.m.sanchez@despegar.com',
  'joaquin.sosa@despegar.com'
];

var TEAM_EMAILS = {
  media : [],
  b2b   : [],
  risk  : []
};

function _getTeamForUser_(email) {
  if (!email) return null;
  var lc = email.toLowerCase();
  if (HUB_EDITORS.map(function(e){ return e.toLowerCase(); }).indexOf(lc) !== -1) {
    return 'all';
  }
  var keys = Object.keys(AREA_CONFIG);
  for (var i = 0; i < keys.length; i++) {
    var areaId = keys[i];
    var members = (AREA_CONFIG[areaId].team || []);
    for (var j = 0; j < members.length; j++) {
      var m = members[j];
      if (m.email && m.email.toLowerCase() === lc) return areaId;
    }
  }
  var teamKeys = Object.keys(TEAM_EMAILS);
  for (var t = 0; t < teamKeys.length; t++) {
    var tid = teamKeys[t];
    var emails = TEAM_EMAILS[tid].map(function(e){ return e.toLowerCase(); });
    if (emails.indexOf(lc) !== -1) return tid;
  }
  return null;
}

function _ensureHubTasksSheet_() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID_V2);
  var sh = ss.getSheetByName('Hub_Tasks');
  if (!sh) {
    sh = ss.insertSheet('Hub_Tasks');
    sh.appendRow(['id','team','title','description','owner','status','createdAt','updatedAt']);
    sh.setFrozenRows(1);
  }
  return sh;
}

function getHubTasks() {
  var email = Session.getActiveUser().getEmail();
  var team  = _getTeamForUser_(email);
  if (!team) throw new Error('Sin acceso: no pertenecés a ningún equipo configurado.');
  var sh   = _ensureHubTasksSheet_();
  var data = sh.getDataRange().getValues();
  if (data.length <= 1) return [];
  var headers = data[0];
  var idIdx   = headers.indexOf('id');
  var teamIdx = headers.indexOf('team');
  var rows = data.slice(1).filter(function(r){ return r[idIdx]; });
  if (team !== 'all') {
    rows = rows.filter(function(r){ return r[teamIdx] === team; });
  }
  return rows.map(function(r){
    var obj = {};
    headers.forEach(function(h, i){ obj[h] = r[i]; });
    return obj;
  });
}

function saveHubTask(payload) {
  var email = Session.getActiveUser().getEmail();
  var team  = _getTeamForUser_(email);
  if (!team) throw new Error('Sin acceso.');
  var targetTeam = payload.team || team;
  if (team !== 'all' && targetTeam !== team) throw new Error('No podés editar tareas de otro equipo.');
  var sh      = _ensureHubTasksSheet_();
  var data    = sh.getDataRange().getValues();
  var headers = data[0];
  var idIdx   = headers.indexOf('id');
  var now     = new Date().toISOString();
  if (payload.id) {
    for (var i = 1; i < data.length; i++) {
      if (data[i][idIdx] === payload.id) {
        headers.forEach(function(h, col){
          if (h === 'updatedAt')           sh.getRange(i+1, col+1).setValue(now);
          else if (payload[h] !== undefined) sh.getRange(i+1, col+1).setValue(payload[h]);
        });
        return { ok: true, id: payload.id };
      }
    }
  }
  var newId = Utilities.getUuid();
  var row   = headers.map(function(h){
    if (h === 'id')        return newId;
    if (h === 'team')      return targetTeam;
    if (h === 'createdAt') return now;
    if (h === 'updatedAt') return now;
    return payload[h] !== undefined ? payload[h] : '';
  });
  sh.appendRow(row);
  return { ok: true, id: newId };
}

function deleteHubTask(taskId) {
  var email = Session.getActiveUser().getEmail();
  var team  = _getTeamForUser_(email);
  if (!team) throw new Error('Sin acceso.');
  var sh      = _ensureHubTasksSheet_();
  var data    = sh.getDataRange().getValues();
  var headers = data[0];
  var idIdx   = headers.indexOf('id');
  var teamIdx = headers.indexOf('team');
  for (var i = 1; i < data.length; i++) {
    if (data[i][idIdx] === taskId) {
      if (team !== 'all' && data[i][teamIdx] !== team) throw new Error('Sin acceso a esta tarea.');
      sh.deleteRow(i + 1);
      return { ok: true };
    }
  }
  throw new Error('Tarea no encontrada.');
}
