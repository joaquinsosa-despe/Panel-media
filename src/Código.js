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

// =============================================================================
// HUB SECTIONS — configuración de secciones del Hub
// =============================================================================

var AREA_CONFIG = {};

function _ensureHubSectionsSheet_() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID_V2);
  var sh = ss.getSheetByName('Hub_Sections');
  if (!sh) {
    sh = ss.insertSheet('Hub_Sections');
    sh.appendRow(['areaId','label','team','order','color','icon']);
    sh.setFrozenRows(1);
  }
  return sh;
}

function getHubSections() {
  var sh   = _ensureHubSectionsSheet_();
  var data = sh.getDataRange().getValues();
  if (data.length <= 1) return [];
  var headers = data[0];
  return data.slice(1).map(function(r){
    var obj = {};
    headers.forEach(function(h, i){ obj[h] = r[i]; });
    return obj;
  });
}

function saveHubSection(payload) {
  var email = Session.getActiveUser().getEmail();
  var lc    = email.toLowerCase();
  if (HUB_EDITORS.map(function(e){ return e.toLowerCase(); }).indexOf(lc) === -1) {
    throw new Error('Solo los editores pueden modificar secciones.');
  }
  var sh      = _ensureHubSectionsSheet_();
  var data    = sh.getDataRange().getValues();
  var headers = data[0];
  var idIdx   = headers.indexOf('areaId');
  if (payload.areaId) {
    for (var i = 1; i < data.length; i++) {
      if (data[i][idIdx] === payload.areaId) {
        headers.forEach(function(h, col){
          if (payload[h] !== undefined) sh.getRange(i+1, col+1).setValue(payload[h]);
        });
        return { ok: true, areaId: payload.areaId };
      }
    }
  }
  var row = headers.map(function(h){ return payload[h] !== undefined ? payload[h] : ''; });
  sh.appendRow(row);
  return { ok: true, areaId: payload.areaId };
}

function deleteHubSection(areaId) {
  var email = Session.getActiveUser().getEmail();
  var lc    = email.toLowerCase();
  if (HUB_EDITORS.map(function(e){ return e.toLowerCase(); }).indexOf(lc) === -1) {
    throw new Error('Solo los editores pueden eliminar secciones.');
  }
  var sh      = _ensureHubSectionsSheet_();
  var data    = sh.getDataRange().getValues();
  var headers = data[0];
  var idIdx   = headers.indexOf('areaId');
  for (var i = 1; i < data.length; i++) {
    if (data[i][idIdx] === areaId) {
      sh.deleteRow(i + 1);
      return { ok: true };
    }
  }
  throw new Error('Sección no encontrada.');
}

// =============================================================================
// HUB MEMBERS — configuración de miembros por área
// =============================================================================

function _ensureHubMembersSheet_() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID_V2);
  var sh = ss.getSheetByName('Hub_Members');
  if (!sh) {
    sh = ss.insertSheet('Hub_Members');
    sh.appendRow(['areaId','email','name','role']);
    sh.setFrozenRows(1);
  }
  return sh;
}

function getHubMembers() {
  var sh   = _ensureHubMembersSheet_();
  var data = sh.getDataRange().getValues();
  if (data.length <= 1) return [];
  var headers = data[0];
  return data.slice(1).map(function(r){
    var obj = {};
    headers.forEach(function(h, i){ obj[h] = r[i]; });
    return obj;
  });
}

function saveHubMember(payload) {
  var email = Session.getActiveUser().getEmail();
  var lc    = email.toLowerCase();
  if (HUB_EDITORS.map(function(e){ return e.toLowerCase(); }).indexOf(lc) === -1) {
    throw new Error('Solo los editores pueden modificar miembros.');
  }
  var sh      = _ensureHubMembersSheet_();
  var data    = sh.getDataRange().getValues();
  var headers = data[0];
  var areaIdx  = headers.indexOf('areaId');
  var emailIdx = headers.indexOf('email');
  if (payload.areaId && payload.email) {
    for (var i = 1; i < data.length; i++) {
      if (data[i][areaIdx] === payload.areaId && data[i][emailIdx] === payload.email) {
        headers.forEach(function(h, col){
          if (payload[h] !== undefined) sh.getRange(i+1, col+1).setValue(payload[h]);
        });
        return { ok: true };
      }
    }
  }
  var row = headers.map(function(h){ return payload[h] !== undefined ? payload[h] : ''; });
  sh.appendRow(row);
  return { ok: true };
}

function deleteHubMember(areaId, memberEmail) {
  var email = Session.getActiveUser().getEmail();
  var lc    = email.toLowerCase();
  if (HUB_EDITORS.map(function(e){ return e.toLowerCase(); }).indexOf(lc) === -1) {
    throw new Error('Solo los editores pueden eliminar miembros.');
  }
  var sh      = _ensureHubMembersSheet_();
  var data    = sh.getDataRange().getValues();
  var headers = data[0];
  var areaIdx  = headers.indexOf('areaId');
  var emailIdx = headers.indexOf('email');
  for (var i = 1; i < data.length; i++) {
    if (data[i][areaIdx] === areaId && data[i][emailIdx] === memberEmail) {
      sh.deleteRow(i + 1);
      return { ok: true };
    }
  }
  throw new Error('Miembro no encontrado.');
}

// =============================================================================
// HUB LINKS — links rápidos por área
// =============================================================================

function _ensureHubLinksSheet_() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID_V2);
  var sh = ss.getSheetByName('Hub_Links');
  if (!sh) {
    sh = ss.insertSheet('Hub_Links');
    sh.appendRow(['id','areaId','label','url','icon','order']);
    sh.setFrozenRows(1);
  }
  return sh;
}

function getHubLinks() {
  var sh   = _ensureHubLinksSheet_();
  var data = sh.getDataRange().getValues();
  if (data.length <= 1) return [];
  var headers = data[0];
  return data.slice(1).map(function(r){
    var obj = {};
    headers.forEach(function(h, i){ obj[h] = r[i]; });
    return obj;
  });
}

function saveHubLink(payload) {
  var email = Session.getActiveUser().getEmail();
  var team  = _getTeamForUser_(email);
  if (!team) throw new Error('Sin acceso.');
  var sh      = _ensureHubLinksSheet_();
  var data    = sh.getDataRange().getValues();
  var headers = data[0];
  var idIdx   = headers.indexOf('id');
  var now     = new Date().toISOString();
  if (payload.id) {
    for (var i = 1; i < data.length; i++) {
      if (data[i][idIdx] === payload.id) {
        headers.forEach(function(h, col){
          if (payload[h] !== undefined) sh.getRange(i+1, col+1).setValue(payload[h]);
        });
        return { ok: true, id: payload.id };
      }
    }
  }
  var newId = Utilities.getUuid();
  var row   = headers.map(function(h){
    if (h === 'id') return newId;
    return payload[h] !== undefined ? payload[h] : '';
  });
  sh.appendRow(row);
  return { ok: true, id: newId };
}

function deleteHubLink(linkId) {
  var email = Session.getActiveUser().getEmail();
  var team  = _getTeamForUser_(email);
  if (!team) throw new Error('Sin acceso.');
  var sh      = _ensureHubLinksSheet_();
  var data    = sh.getDataRange().getValues();
  var headers = data[0];
  var idIdx   = headers.indexOf('id');
  for (var i = 1; i < data.length; i++) {
    if (data[i][idIdx] === linkId) {
      sh.deleteRow(i + 1);
      return { ok: true };
    }
  }
  throw new Error('Link no encontrado.');
}

// =============================================================================
// SPREADSHEET IDs
// =============================================================================
var SPREADSHEET_ID    = '1BxiMVs0XRA5nFMdKvBdBZjgmUUqptlbs74OgVE2upms'; // original
var SPREADSHEET_ID_V2 = '1HCZUV6ixNfmNFGOGxKfhFdx5jSv7OOTXCvRqeJJ5C_s'; // v2

// =============================================================================
// doGet — entry point
// =============================================================================
function doGet(e) {
  return HtmlService
    .createTemplateFromFile('Hub')
    .evaluate()
    .setTitle('Media Hub — Despegar')
    .addMetaTag('viewport','width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// =============================================================================
// include — helper para incluir archivos HTML parciales
// =============================================================================
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// =============================================================================
// DATA FETCH — funciones de lectura de datos
// =============================================================================

function getSheetData() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID_V2);
  var sh = ss.getSheetByName(CONFIG.DATASET_SHEET);
  if (!sh) throw new Error('Sheet "' + CONFIG.DATASET_SHEET + '" not found');
  return sh.getDataRange().getValues();
}

function getHelperData() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID_V2);
  var sh = ss.getSheetByName(CONFIG.HELPER_SHEET);
  if (!sh) return [];
  return sh.getDataRange().getValues();
}

// =============================================================================
// WATERFALL DATA — datos para el gráfico de cascada
// =============================================================================

function getWaterfallData() {
  var rows = getSheetData();
  if (rows.length < 2) return { labels: [], values: [], colors: [], total: 0 };

  var headers = rows[0].map(function(h){ return String(h).trim(); });
  var mIdx    = headers.indexOf("Month");
  var yIdx    = headers.indexOf("Year");
  var lobIdx  = headers.indexOf("LOB");
  var valIdx  = headers.indexOf(CONFIG.FIELD_COLUMNS);

  if ([mIdx, yIdx, lobIdx, valIdx].indexOf(-1) !== -1) {
    throw new Error("Columnas requeridas no encontradas. Headers: " + headers.join(", "));
  }

  var targetMonth = CONFIG.TARGET_MONTH;
  var targetYear  = String(CONFIG.TARGET_YEAR);

  // Agrupar por LOB mapeado
  var lobTotals = {};
  LOB_ORDER.forEach(function(l){ lobTotals[l] = 0; });

  rows.slice(1).forEach(function(r){
    if (String(r[mIdx]).trim() !== targetMonth) return;
    if (String(r[yIdx]).trim() !== targetYear)  return;
    var raw    = String(r[lobIdx]).trim();
    var mapped = LOB_MAP[raw];
    if (!mapped) return;
    var val = parseFloat(String(r[valIdx]).replace(/,/g, '')) || 0;
    lobTotals[mapped] = (lobTotals[mapped] || 0) + val;
  });

  var labels = [], values = [], colors = [];
  var palette = ["#6B35A8","#1E8449","#1565C0","#E67E22"];
  LOB_ORDER.forEach(function(lob, i){
    labels.push(lob);
    values.push(Math.round(lobTotals[lob]));
    colors.push(palette[i % palette.length]);
  });

  var total = values.reduce(function(a,b){ return a+b; }, 0);
  return { labels: labels, values: values, colors: colors, total: total };
}

// =============================================================================
// TREND DATA — datos para el gráfico de tendencia
// =============================================================================

function getTrendData() {
  var rows = getSheetData();
  if (rows.length < 2) return { months: [], actuals: [], budget: [] };

  var headers = rows[0].map(function(h){ return String(h).trim(); });
  var mIdx    = headers.indexOf("Month");
  var yIdx    = headers.indexOf("Year");
  var aIdx    = headers.indexOf(CONFIG.FIELD_COLUMNS);
  var bIdx    = headers.indexOf(CONFIG.FIELD_LINE);

  if ([mIdx, yIdx, aIdx, bIdx].indexOf(-1) !== -1) {
    return { months: [], actuals: [], budget: [] };
  }

  var targetYear = String(CONFIG.TARGET_YEAR);
  var byMonth    = {};

  MONTH_ORDER.forEach(function(m){ byMonth[m] = { actuals: 0, budget: 0 }; });

  rows.slice(1).forEach(function(r){
    if (String(r[yIdx]).trim() !== targetYear) return;
    var month = String(r[mIdx]).trim();
    if (!byMonth[month]) return;
    byMonth[month].actuals += parseFloat(String(r[aIdx]).replace(/,/g,'')) || 0;
    byMonth[month].budget  += parseFloat(String(r[bIdx]).replace(/,/g,'')) || 0;
  });

  var months = [], actuals = [], budget = [];
  MONTH_ORDER.forEach(function(m){
    months.push(m.substr(0,3));
    actuals.push(Math.round(byMonth[m].actuals));
    budget.push(Math.round(byMonth[m].budget));
  });

  return { months: months, actuals: actuals, budget: budget };
}

// =============================================================================
// SLIDES GENERATION — generación de presentación
// =============================================================================

function generateSlides() {
  var wfData    = getWaterfallData();
  var trendData = getTrendData();

  var pres  = SlidesApp.create(CONFIG.PRES_TITLE + " — " + CONFIG.TARGET_MONTH + " " + CONFIG.TARGET_YEAR);
  var slide = pres.getSlides()[0];
  slide.getBackground().setSolidFill(C.white);

  // Título
  var titleBox = slide.insertTextBox(
    CONFIG.PRES_TITLE + "\n" + CONFIG.TARGET_MONTH + " " + CONFIG.TARGET_YEAR,
    0, 0, 720, 60
  );
  var titleStyle = titleBox.getText().getTextStyle();
  titleStyle.setFontSize(22).setBold(true).setForegroundColor(C.purpleDark);
  titleBox.getFill().setSolidFill(C.purpleLight);

  // Waterfall
  _drawWaterfall_(slide, wfData);

  // Trend
  _drawTrend_(slide, trendData);

  return pres.getUrl();
}

function _drawWaterfall_(slide, data) {
  var x = 20, y = 70, w = 320, h = 280;
  var n = data.values.length;
  if (n === 0) return;

  var maxVal = Math.max.apply(null, data.values.map(Math.abs)) * 1.2 || 1;
  var barW   = (w - 20) / n - 8;
  var scaleY = h / maxVal;

  var running = 0;
  data.values.forEach(function(val, i){
    var barH  = Math.abs(val) * scaleY;
    var baseY = y + h - running * scaleY - barH;
    if (val < 0) baseY = y + h - running * scaleY;

    var shape = slide.insertShape(SlidesApp.ShapeType.RECTANGLE,
      x + i * (barW + 8) + 10, baseY, barW, barH);
    shape.getFill().setSolidFill(data.colors[i]);
    shape.getBorder().setTransparent();

    var lbl = slide.insertTextBox(
      data.labels[i] + "\n" + _fmt_(val),
      x + i * (barW + 8) + 10, baseY - 30, barW, 28
    );
    var ls = lbl.getText().getTextStyle();
    ls.setFontSize(8).setBold(true).setForegroundColor(C.black);
    lbl.getText().getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);

    running += val;
  });

  // Total
  var totalBox = slide.insertTextBox(
    "TOTAL: " + _fmt_(data.total),
    x, y + h + 5, w, 20
  );
  var ts = totalBox.getText().getTextStyle();
  ts.setFontSize(10).setBold(true).setForegroundColor(C.purpleDark);
}

function _drawTrend_(slide, data) {
  var x = 360, y = 70, w = 340, h = 280;
  var n = data.months.length;
  if (n === 0) return;

  var allVals = data.actuals.concat(data.budget);
  var maxVal  = Math.max.apply(null, allVals) * 1.1 || 1;
  var scaleY  = h / maxVal;
  var stepX   = w / (n - 1 || 1);

  // Línea actuals
  for (var i = 0; i < n - 1; i++){
    var x1 = x + i * stepX;
    var y1 = y + h - data.actuals[i] * scaleY;
    var x2 = x + (i+1) * stepX;
    var y2 = y + h - data.actuals[i+1] * scaleY;
    var line = slide.insertLine(SlidesApp.LineCategory.STRAIGHT,
      x1, y1, x2, y2);
    line.getLineFill().setSolidFill(C.purple);
    line.setWeight(2.5);
  }

  // Línea budget
  for (var j = 0; j < n - 1; j++){
    var bx1 = x + j * stepX;
    var by1 = y + h - data.budget[j] * scaleY;
    var bx2 = x + (j+1) * stepX;
    var by2 = y + h - data.budget[j+1] * scaleY;
    var bline = slide.insertLine(SlidesApp.LineCategory.STRAIGHT,
      bx1, by1, bx2, by2);
    bline.getLineFill().setSolidFill(C.gray);
    bline.setWeight(1.5);
  }

  // Labels eje X
  for (var k = 0; k < n; k++){
    var lx = x + k * stepX - 15;
    var lb = slide.insertTextBox(data.months[k], lx, y + h + 2, 30, 12);
    lb.getText().getTextStyle().setFontSize(7).setForegroundColor(C.subText);
  }
}

function _fmt_(n) {
  if (Math.abs(n) >= 1e6) return (n/1e6).toFixed(1) + "M";
  if (Math.abs(n) >= 1e3) return (n/1e3).toFixed(0) + "K";
  return String(n);
}

// =============================================================================
// PANEL V2 — Funciones del panel de control
// =============================================================================

function getPanelData() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID_V2);
  return _buildPanelPayload_(ss);
}

function _buildPanelPayload_(ss) {
  var result = {
    revenue   : _getRevenueSummary_(ss),
    lob       : _getLobBreakdown_(ss),
    trend     : _getTrendSummary_(ss),
    topDeals  : _getTopDeals_(ss),
    timestamp : new Date().toISOString()
  };
  return result;
}

function _getRevenueSummary_(ss) {
  var sh = ss.getSheetByName(CONFIG.DATASET_SHEET);
  if (!sh) return {};
  var rows    = sh.getDataRange().getValues();
  var headers = rows[0].map(function(h){ return String(h).trim(); });
  var mIdx    = headers.indexOf("Month");
  var yIdx    = headers.indexOf("Year");
  var aIdx    = headers.indexOf(CONFIG.FIELD_COLUMNS);
  var bIdx    = headers.indexOf(CONFIG.FIELD_LINE);

  var actual = 0, budget = 0;
  rows.slice(1).forEach(function(r){
    if (String(r[mIdx]).trim() !== CONFIG.TARGET_MONTH) return;
    if (String(r[yIdx]).trim() !== String(CONFIG.TARGET_YEAR)) return;
    actual += parseFloat(String(r[aIdx]).replace(/,/g,'')) || 0;
    budget += parseFloat(String(r[bIdx]).replace(/,/g,'')) || 0;
  });

  return {
    actual : Math.round(actual),
    budget : Math.round(budget),
    pct    : budget ? Math.round(actual / budget * 100) : 0
  };
}

function _getLobBreakdown_(ss) {
  var sh = ss.getSheetByName(CONFIG.DATASET_SHEET);
  if (!sh) return [];
  var rows    = sh.getDataRange().getValues();
  var headers = rows[0].map(function(h){ return String(h).trim(); });
  var mIdx    = headers.indexOf("Month");
  var yIdx    = headers.indexOf("Year");
  var lobIdx  = headers.indexOf("LOB");
  var aIdx    = headers.indexOf(CONFIG.FIELD_COLUMNS);

  var totals = {};
  LOB_ORDER.forEach(function(l){ totals[l] = 0; });

  rows.slice(1).forEach(function(r){
    if (String(r[mIdx]).trim() !== CONFIG.TARGET_MONTH) return;
    if (String(r[yIdx]).trim() !== String(CONFIG.TARGET_YEAR)) return;
    var mapped = LOB_MAP[String(r[lobIdx]).trim()];
    if (!mapped) return;
    totals[mapped] += parseFloat(String(r[aIdx]).replace(/,/g,'')) || 0;
  });

  return LOB_ORDER.map(function(lob){
    return { lob: lob, value: Math.round(totals[lob]) };
  });
}

function _getTrendSummary_(ss) {
  var sh = ss.getSheetByName(CONFIG.DATASET_SHEET);
  if (!sh) return { months: [], actuals: [], budget: [] };
  var rows    = sh.getDataRange().getValues();
  var headers = rows[0].map(function(h){ return String(h).trim(); });
  var mIdx    = headers.indexOf("Month");
  var yIdx    = headers.indexOf("Year");
  var aIdx    = headers.indexOf(CONFIG.FIELD_COLUMNS);
  var bIdx    = headers.indexOf(CONFIG.FIELD_LINE);

  var byMonth = {};
  MONTH_ORDER.forEach(function(m){ byMonth[m] = { a: 0, b: 0 }; });

  rows.slice(1).forEach(function(r){
    if (String(r[yIdx]).trim() !== String(CONFIG.TARGET_YEAR)) return;
    var m = String(r[mIdx]).trim();
    if (!byMonth[m]) return;
    byMonth[m].a += parseFloat(String(r[aIdx]).replace(/,/g,'')) || 0;
    byMonth[m].b += parseFloat(String(r[bIdx]).replace(/,/g,'')) || 0;
  });

  var months = [], actuals = [], budget = [];
  MONTH_ORDER.forEach(function(m){
    months.push(m.substr(0,3));
    actuals.push(Math.round(byMonth[m].a));
    budget.push(Math.round(byMonth[m].b));
  });
  return { months: months, actuals: actuals, budget: budget };
}

function _getTopDeals_(ss) {
  var sh = ss.getSheetByName(CONFIG.DATASET_SHEET);
  if (!sh) return [];
  var rows    = sh.getDataRange().getValues();
  var headers = rows[0].map(function(h){ return String(h).trim(); });
  var mIdx    = headers.indexOf("Month");
  var yIdx    = headers.indexOf("Year");
  var aIdx    = headers.indexOf(CONFIG.FIELD_COLUMNS);
  var advIdx  = headers.indexOf("Advertiser");
  if (advIdx === -1) return [];

  var deals = [];
  rows.slice(1).forEach(function(r){
    if (String(r[mIdx]).trim() !== CONFIG.TARGET_MONTH) return;
    if (String(r[yIdx]).trim() !== String(CONFIG.TARGET_YEAR)) return;
    var val = parseFloat(String(r[aIdx]).replace(/,/g,'')) || 0;
    if (val > 0) deals.push({ name: String(r[advIdx]).trim(), value: Math.round(val) });
  });

  deals.sort(function(a,b){ return b.value - a.value; });
  return deals.slice(0, 10);
}

// =============================================================================
// PANEL DATA V3 — datos enriquecidos para el Hub
// =============================================================================

function getPanelDataV3() {
  var email   = Session.getActiveUser().getEmail();
  var isAdmin = HUB_EDITORS.map(function(e){ return e.toLowerCase(); })
                  .indexOf(email.toLowerCase()) !== -1;

  // Secciones desde la hoja Hub_Sections
  var sectionsData = getHubSections();
  var membersData  = getHubMembers();
  var linksData    = getHubLinks();
  var tasksData    = [];
  try { tasksData = getHubTasks(); } catch(e) {}

  // Reconstruir AREA_CONFIG dinámicamente
  AREA_CONFIG = {};
  sectionsData.forEach(function(s){
    if (!s.areaId) return;
    AREA_CONFIG[s.areaId] = {
      label : s.label  || s.areaId,
      team  : [],
      color : s.color  || '#6B35A8',
      icon  : s.icon   || '',
      order : s.order  || 0
    };
  });
  membersData.forEach(function(m){
    if (!m.areaId || !AREA_CONFIG[m.areaId]) return;
    AREA_CONFIG[m.areaId].team.push({ email: m.email, name: m.name, role: m.role });
  });

  // Construir secciones para el frontend
  var sections = Object.keys(AREA_CONFIG).map(function(areaId){
    var cfg   = AREA_CONFIG[areaId];
    var links = linksData.filter(function(l){ return l.areaId === areaId; })
                  .sort(function(a,b){ return a.order - b.order; });
    var tasks = tasksData.filter(function(t){ return t.team === areaId; });
    return {
      id     : areaId,
      label  : cfg.label,
      color  : cfg.color,
      icon   : cfg.icon,
      order  : cfg.order,
      team   : cfg.team,
      links  : links,
      tasks  : tasks
    };
  });

  sections.sort(function(a, b){ return a.order - b.order; });
  return { sections: sections, user: { email: email, isAdmin: isAdmin } };
}