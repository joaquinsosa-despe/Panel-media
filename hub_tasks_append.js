
// =============================================================================
// HUB TASKS — Sección de tareas pendientes por equipo
// =============================================================================

var HUB_EDITORS = [
  'matias.m.sanchez@despegar.com',
  'joaquin.sosa@despegar.com'
];

// Emails por equipo (para usuarios que no tienen email en AREA_CONFIG.team)
var TEAM_EMAILS = {
  media : [],
  b2b   : [],
  risk  : []
};

/**
 * Devuelve el equipo del usuario: 'media' | 'b2b' | 'risk' | 'all' | null
 */
function _getTeamForUser_(email) {
  if (!email) return null;
  var lc = email.toLowerCase();

  // Editores globales
  if (HUB_EDITORS.map(function(e){ return e.toLowerCase(); }).indexOf(lc) !== -1) {
    return 'all';
  }

  // Buscar en AREA_CONFIG.team (cada miembro puede tener campo email)
  var keys = Object.keys(AREA_CONFIG);
  for (var i = 0; i < keys.length; i++) {
    var areaId = keys[i];
    var members = (AREA_CONFIG[areaId].team || []);
    for (var j = 0; j < members.length; j++) {
      var m = members[j];
      if (m.email && m.email.toLowerCase() === lc) return areaId;
    }
  }

  // Buscar en TEAM_EMAILS fallback
  var teamKeys = Object.keys(TEAM_EMAILS);
  for (var t = 0; t < teamKeys.length; t++) {
    var tid = teamKeys[t];
    var emails = TEAM_EMAILS[tid].map(function(e){ return e.toLowerCase(); });
    if (emails.indexOf(lc) !== -1) return tid;
  }

  return null;
}

/**
 * Asegura que exista la hoja Hub_Tasks con los headers correctos
 */
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

/**
 * Devuelve las tareas del equipo del usuario (o todas si es editor global)
 */
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

  var rows = data.slice(1).filter(function(r){ return r[idIdx]; }); // excluir filas vacías

  if (team !== 'all') {
    rows = rows.filter(function(r){ return r[teamIdx] === team; });
  }

  return rows.map(function(r){
    var obj = {};
    headers.forEach(function(h, i){ obj[h] = r[i]; });
    return obj;
  });
}

/**
 * Crea o actualiza una tarea
 */
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
    // Actualizar fila existente
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

  // Nueva fila
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

/**
 * Elimina una tarea por id
 */
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
