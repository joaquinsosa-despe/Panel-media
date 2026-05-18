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


// -----------------------------------------
// MENÚ
// -----------------------------------------
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("📊 Dashboard")
    .addItem("▶ Generar Google Slides", "generateGoogleSlides")
    .addSeparator()
    .addItem("🗑 Eliminar hoja auxiliar", "deleteHelperSheet")
    .addItem("🔄 Resetear presentación",  "resetPresentation")
    .addToUi();
}


// -----------------------------------------
// FUNCIÓN PRINCIPAL
// -----------------------------------------
function generateGoogleSlides() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();


  toast("📖 Leyendo dataset...");
  const dsSheet = ss.getSheetByName(CONFIG.DATASET_SHEET);
  if (!dsSheet) { ui.alert("❌ No se encontró la hoja 'dataset'."); return; }


  const allData = readDataset(dsSheet);
  if (!allData || allData.length === 0) return;


  const qLabel = getQuarter(CONFIG.TARGET_MONTH);


  const byMonth = allData.filter(
    r => r.month === CONFIG.TARGET_MONTH && r.year === CONFIG.TARGET_YEAR
  );
  const byQ = allData.filter(
    r => r.year === CONFIG.TARGET_YEAR && getQuarter(r.month) === qLabel
  );


  if (byMonth.length === 0) {
    ui.alert(`⚠️ Sin datos para ${CONFIG.TARGET_MONTH} ${CONFIG.TARGET_YEAR}.\nVerificá el CONFIG.`);
    return;
  }


  toast("🔢 Calculando métricas...");
  // Leer GB para calcular tasas (usa las funciones del Web App ya definidas)
  const gbMes = buildGBByCountry_(ss, CONFIG.TARGET_MONTH, CONFIG.TARGET_YEAR);
  const gbQ   = buildGBForQuarter_(ss, CONFIG.TARGET_MONTH, CONFIG.TARGET_YEAR);
  const tables = buildAllTables(byMonth, byQ, gbMes, gbQ);


  toast("📊 Generando gráficos...");
  const helper = buildHelperSheet(ss, tables, allData);
  SpreadsheetApp.flush();
  Utilities.sleep(5000);


  const rawCharts = helper.getCharts();
  if (rawCharts.length === 0) {
    ui.alert("❌ No se generaron gráficos. Intentá de nuevo.");
    return;
  }


  const charts = rawCharts.sort((a, b) =>
    a.getContainerInfo().getAnchorRow() - b.getContainerInfo().getAnchorRow()
  );


  toast("🎨 Construyendo Google Slides...");
  const pres = buildPresentation(charts, tables, allData);


  const url = pres.getUrl();
  toast("✅ ¡Listo!");
  ui.alert("✅ Google Slides generado",
    `Presentación creada exitosamente.\n\n🔗 ${url}`, ui.ButtonSet.OK);
}


// -----------------------------------------
// LEER DATASET
// -----------------------------------------
function readDataset(sheet) {
  const raw=sheet.getDataRange().getValues();
  const headers=raw[0].map(h=>String(h).trim());
  const idx={year:headers.indexOf("Year"),month:headers.indexOf("Month"),country:headers.indexOf("Pais"),channel:headers.indexOf("Lob-Channel"),concept:headers.indexOf("Concepto"),cluster:headers.indexOf("Cluster"),partner:headers.indexOf("Partner"),actuals:headers.indexOf("Actuals + RR"),budgetFY26:headers.indexOf("Budget + Forecast + Breakthrough FY26"),budgetFY27:headers.indexOf("Budget FY27")};
  const REQUIRED={"Year":idx.year,"Month":idx.month,"Pais":idx.country,"Lob-Channel":idx.channel,"Concepto":idx.concept,"Cluster":idx.cluster,"Actuals + RR":idx.actuals};
  const missing=Object.entries(REQUIRED).filter(([,v])=>v===-1).map(([name])=>name);
  if(missing.length>0){SpreadsheetApp.getUi().alert("\u26a0\ufe0f Columnas no encontradas:\n"+missing.join("\n")+"\n\nVerific\u00e1 los nombres.");return[];}
  return raw.slice(1).map(r=>{
    const ch=String(r[idx.channel]||"").trim(),mo=String(r[idx.month]||"").trim(),yr=String(r[idx.year]||"").trim();
    const fyInf=getFYInfo_(mo,yr);
    const budget=(fyInf.fy>=2027&&idx.budgetFY27!==-1)?parseNum_(r[idx.budgetFY27]):(idx.budgetFY26!==-1?parseNum_(r[idx.budgetFY26]):0);
    return{year:yr,quarter:"Q"+fyInf.q,month:mo,country:String(r[idx.country]||"").trim(),channel:ch,lob:LOB_MAP[ch]||"S/C",concept:String(r[idx.concept]||"").trim(),cluster:String(r[idx.cluster]||"").trim(),partner:idx.partner!==-1?String(r[idx.partner]||"").trim():"",actuals:parseNum_(r[idx.actuals]),budget};
  }).filter(r=>r.month&&r.month!=="Month"&&r.year&&r.year!=="Year");
}


function parseNum_(v) {
  if (typeof v === "number") return v;
  if (typeof v === "string") return parseFloat(v.replace(",", ".")) || 0;
  return 0;
}


function getQuarter(month) {
  const map = {
    January:"Q1", February:"Q1", March:"Q1",
    April:"Q2",   May:"Q2",      June:"Q2",
    July:"Q3",    August:"Q3",   September:"Q3",
    October:"Q4", November:"Q4", December:"Q4",
  };
  return map[month] || "Q1";
}


// -----------------------------------------
// CONSTRUIR TODAS LAS TABLAS
// -----------------------------------------
function buildAllTables(byMonth, byQ) {


  function groupBy(data, key) {
    const map = {};
    data.forEach(r => {
      const k = r[key] || "N/A";
      if (!map[k]) map[k] = { actuals: 0, budget: 0 };
      map[k].actuals += r.actuals;
      map[k].budget  += r.budget;
    });
    return map;
  }


  function toRows(map) {
    const rows = Object.entries(map).map(([k, v]) => {
      const delta = v.actuals - v.budget;
      const pct   = v.budget !== 0 ? (delta / v.budget) * 100 : 0;
      return [k, v.budget, v.actuals, delta, pct];
    });
    rows.sort((a, b) => b[2] - a[2]);
    const tAct = rows.reduce((s, r) => s + r[2], 0);
    const tBud = rows.reduce((s, r) => s + r[1], 0);
    const tDel = tAct - tBud;
    const tPct = tBud !== 0 ? (tDel / tBud) * 100 : 0;
    rows.push(["TOTAL", tBud, tAct, tDel, tPct]);
    return rows;
  }


  function buildSection(data) {
    const lobBreak = {};
    LOB_ORDER.forEach(lob => {
      const d = data.filter(r => r.lob === lob);
      lobBreak[lob] = {
        pais   : toRows(groupBy(d, "country")),
        cluster: toRows(groupBy(d, "cluster")),
      };
    });
    return {
      pais   : toRows(groupBy(data, "country")),
      cluster: toRows(groupBy(data, "cluster")),
      concept: toRows(groupBy(data, "concept")),
      lob    : toRows(groupBy(data, "lob")),
      lobBreak,
    };
  }


  return { mes: buildSection(byMonth), q: buildSection(byQ) };
}


function buildWaterfallData(rows) {
  const dataRows = rows.slice(0, -1);
  const totalRow = rows[rows.length - 1];
  const totalBud = totalRow[1] / 1e6;
  const totalAct = totalRow[2] / 1e6;


  const labelBud = activeBudgetLabel_();
  const labelAct = displayName(CONFIG.FIELD_COLUMNS);


  const out = [[
    "Categoría","Base","Positivo","PosLbl",
    "Negativo","NegLbl","Budget","BudgetLbl","Actuals","ActualsLbl"
  ]];


  out.push([labelBud, 0, 0, "", 0, "", totalBud, `${totalBud.toFixed(1)}M`, 0, ""]);


  let cum = totalBud;
  dataRows.forEach(r => {
    const delta = r[3] / 1e6;
    if (delta >= 0) {
      out.push([r[0], cum, delta, `+${delta.toFixed(1)}M`, 0, "", 0, "", 0, ""]);
      cum += delta;
    } else {
      const abs = Math.abs(delta);
      cum += delta;
      out.push([r[0], cum, 0, "", abs, `-${abs.toFixed(1)}M`, 0, "", 0, ""]);
    }
  });


  out.push([labelAct, 0, 0, "", 0, "", 0, "", totalAct, `${totalAct.toFixed(1)}M`]);


  return out;
}


// -----------------------------------------
// TREND DATA
// -----------------------------------------
function buildTrendData(allData, lobFilter) {
  // Mostrar FY anterior + FY actual (excluir FY más antiguos)
  const targetFY  = getFYInfo_(CONFIG.TARGET_MONTH, CONFIG.TARGET_YEAR).fy;
  const minFY     = targetFY - 1;  // FY anterior
  const filtered = lobFilter ? allData.filter(r => r.lob === lobFilter) : allData;
  const trendMap = {};


  filtered.forEach(r => {
    const yr = parseInt(r.year);
    const mi = MONTH_ORDER.indexOf(r.month);
    if (mi === -1) return;
    const rowFY = getFYInfo_(r.month, r.year).fy;
    if (rowFY < minFY) return;  // excluir FY más antiguos
    const key = `${yr}-${String(mi + 1).padStart(2, "0")}`;
    const lbl = `${r.month.slice(0, 3)}'${String(yr).slice(2)}`;
    if (!trendMap[key]) trendMap[key] = { lbl, actuals: 0, budget: 0 };
    trendMap[key].actuals += r.actuals;
    trendMap[key].budget  += r.budget;
  });


  const rows = [[
    "Mes",
    displayName(CONFIG.FIELD_COLUMNS), "ActualsLbl",
    activeBudgetLabel_(),    "BudgetLbl"
  ]];


  Object.entries(trendMap)
    .sort((a, b) => a[0].localeCompare(b[0]))
    .forEach(([, v]) => {
      const actM = v.actuals / 1e6;
      const budM = v.budget  / 1e6;
      rows.push([v.lbl, actM, `$${actM.toFixed(1)}M`, budM, `$${budM.toFixed(1)}M`]);
    });


  return rows;
}


// -----------------------------------------
// CREAR HOJA AUXILIAR + GRÁFICOS
// -----------------------------------------
function buildHelperSheet(ss, tables, allData) {
  let helper = ss.getSheetByName(CONFIG.HELPER_SHEET);
  if (helper) {
    helper.getCharts().forEach(c => helper.removeChart(c));
    helper.clearContents();
  } else {
    helper = ss.insertSheet(CONFIG.HELPER_SHEET);
  }


  const { mes, q } = tables;
  const Q = getQuarter(CONFIG.TARGET_MONTH);
  const M = CONFIG.TARGET_MONTH;
  const Y = CONFIG.TARGET_YEAR;


  let row = 1;
  const wfDefs    = [];
  const trendDefs = [];


  function writeWaterfall(rows, title) {
    const wfData = buildWaterfallData(rows);
    helper.getRange(row, 1, wfData.length, 10).setValues(wfData);
    wfDefs.push({ row, rows: wfData.length, title });
    row += wfData.length + 3;
  }


  function writeTrend(data, title) {
    helper.getRange(row, 1, data.length, 5).setValues(data);
    trendDefs.push({ row, rows: data.length, title });
    row += data.length + 3;
  }


  writeWaterfall(mes.pais,    `🌍 ${M} ${Y} — Bridge por País`);
  writeWaterfall(mes.cluster, `🎯 ${M} ${Y} — Bridge por Cluster`);
  writeWaterfall(mes.concept, `💡 ${M} ${Y} — Bridge por Concepto`);
  writeWaterfall(q.pais,      `🌍 ${Q} ${Y} — Bridge por País`);
  writeWaterfall(q.cluster,   `🎯 ${Q} ${Y} — Bridge por Cluster`);
  writeWaterfall(q.concept,   `💡 ${Q} ${Y} — Bridge por Concepto`);
  writeWaterfall(mes.lob,     `📊 ${M} ${Y} — Bridge por LOB`);
  writeWaterfall(q.lob,       `📊 ${Q} ${Y} — Bridge por LOB`);


  LOB_ORDER.forEach(lob => {
    writeWaterfall(mes.lobBreak[lob].pais,    `🌍 ${lob} ${M} — Bridge País`);
    writeWaterfall(mes.lobBreak[lob].cluster, `🎯 ${lob} ${M} — Bridge Cluster`);
  });
  LOB_ORDER.forEach(lob => {
    writeWaterfall(q.lobBreak[lob].pais,    `🌍 ${lob} ${Q} — Bridge País`);
    writeWaterfall(q.lobBreak[lob].cluster, `🎯 ${lob} ${Q} — Bridge Cluster`);
  });


  writeTrend(buildTrendData(allData, null),
    `📈 Tendencia Global — ${displayName(CONFIG.FIELD_COLUMNS)} vs ${activeBudgetLabel_()}`);
  LOB_ORDER.forEach(lob =>
    writeTrend(buildTrendData(allData, lob),
      `📈 Tendencia ${lob} — ${displayName(CONFIG.FIELD_COLUMNS)} vs ${activeBudgetLabel_()}`)
  );


  SpreadsheetApp.flush();


  wfDefs.forEach(def => {
    try {
      helper.insertChart(
        helper.newChart()
          .setChartType(Charts.ChartType.COLUMN)
          .setOption("isStacked", true)
          .setOption("title", "")
          .setOption("legend", { position: "bottom", textStyle: { fontSize: 9, color: C.black } })
          .setOption("hAxis", { textStyle: { fontSize: 8, color: C.black }, slantedText: true, slantedTextAngle: 30 })
          .setOption("vAxis", { format: '#0.0"M"', textStyle: { fontSize: 8, color: C.black }, gridlines: { color: "#EEEEEE" }, minValue: 0 })
          .setOption("series", {
            0: { color: C.white, visibleInLegend: false },
            1: { color: C.green, visibleInLegend: true,  labelInLegend: "Favorable" },
            2: { color: C.red,   visibleInLegend: true,  labelInLegend: "Desfavorable" },
            3: { color: C.gray,  visibleInLegend: true,  labelInLegend: activeBudgetLabel_() },
            4: { color: C.blue,  visibleInLegend: true,  labelInLegend: displayName(CONFIG.FIELD_COLUMNS) },
          })
          .setOption("annotations.textStyle", { fontSize: 9, bold: true, color: "#000000", auraColor: "none" })
          .setOption("annotations.alwaysOutside", false)
          .setOption("annotations.stem.length", 0)
          .setOption("annotations.stem.color", "#FFFFFF")
          .setOption("tooltip.trigger", "focus")
          .setOption("backgroundColor", C.white)
          .setOption("chartArea.left",  55)
          .setOption("chartArea.top",   25)
          .setOption("chartArea.width", "80%")
          .addRange(helper.getRange(def.row, 1, def.rows, 10))
          .setNumHeaders(1)
          .setPosition(def.row, 12, 0, 0)
          .build()
      );
    } catch (e) { Logger.log(`Error waterfall "${def.title}": ${e.message}`); }
  });


  trendDefs.forEach(def => {
    try {
      helper.insertChart(
        helper.newChart()
          .setChartType(Charts.ChartType.COMBO)
          .setOption("seriesType", "bars")
          .setOption("title", "")
          .setOption("legend", { position: "bottom", textStyle: { fontSize: 9, color: C.black } })
          .setOption("hAxis", { textStyle: { fontSize: 8, color: C.black }, slantedText: true, slantedTextAngle: 45 })
          .setOption("vAxis", { title: "$M", format: '"$"#0.0"M"', textStyle: { fontSize: 8, color: C.black }, gridlines: { color: "#EEEEEE" } })
          .setOption("series", {
            0: { type: "bars", color: C.blue,   labelInLegend: displayName(CONFIG.FIELD_COLUMNS) },
            1: { type: "line", color: C.purple, lineWidth: 2, pointSize: 5, lineDashStyle: [6, 3], labelInLegend: activeBudgetLabel_() },
          })
          .setOption("annotations.textStyle", { fontSize: 8, bold: true, color: C.black })
          .setOption("annotations.alwaysOutside", true)
          .setOption("annotations.stem.length", 6)
          .setOption("annotations.stem.color", "#AAAAAA")
          .setOption("tooltip.trigger", "focus")
          .setOption("backgroundColor", C.white)
          .setOption("chartArea.left",  60)
          .setOption("chartArea.top",   25)
          .setOption("chartArea.width", "82%")
          .addRange(helper.getRange(def.row, 1, def.rows, 5))
          .setNumHeaders(1)
          .setPosition(def.row, 12, 0, 0)
          .build()
      );
    } catch (e) { Logger.log(`Error trend "${def.title}": ${e.message}`); }
  });


  helper.hideSheet();
  SpreadsheetApp.flush();
  return helper;
}


// -----------------------------------------
// CONSTRUIR PRESENTACIÓN
// -----------------------------------------
function buildPresentation(charts, tables, allData) {
  const Q         = getQuarter(CONFIG.TARGET_MONTH);
  const M         = CONFIG.TARGET_MONTH;
  const Y         = CONFIG.TARGET_YEAR;
  const presTitle = `${CONFIG.PRES_TITLE} — ${M} + ${Q} ${Y} vs Forecast`;


  let pres;
  const props   = PropertiesService.getScriptProperties();
  const savedId = props.getProperty("PRESENTATION_ID");


  if (savedId) {
    try {
      pres = SlidesApp.openById(savedId);
      pres.getSlides().forEach(s => pres.removeSlide(s));
      pres.setName(presTitle);
    } catch (e) {
      pres = SlidesApp.create(presTitle);
      props.setProperty("PRESENTATION_ID", pres.getId());
    }
  } else {
    pres = SlidesApp.create(presTitle);
    props.setProperty("PRESENTATION_ID", pres.getId());
  }


  const { mes, q } = tables;


  addCoverSlide(pres, M, Y, Q, mes, q);
  // Calcular etiqueta del inicio del trend (inicio del FY anterior)
  const trendFYInfo   = getFYInfo_(M, Y);
  const trendStartYr  = String(trendFYInfo.fy - 2);           // año cal. inicio FY anterior
  const trendStartLbl = `Abr '${trendStartYr.slice(-2)}`;    // ej: "Abr '24"
  const trendFYRange  = `FY${String(trendFYInfo.fy-1).slice(-2)}–FY${String(trendFYInfo.fy).slice(-2)}`;

  addTrendSlide(pres, charts[24],
    `📈 Tendencia Global — ${trendFYRange}`,
    `${displayName(CONFIG.FIELD_COLUMNS)} vs ${activeBudgetLabel_()} · desde ${trendStartLbl}`,
    null, allData);


  addBridgeSlide(pres, charts[0], `🌍 ${M} ${Y} — Bridge por País`,     mes.pais,    "País");
  addBridgeSlide(pres, charts[1], `🎯 ${M} ${Y} — Bridge por Cluster`,  mes.cluster, "Cluster");
  addBridgeSlide(pres, charts[2], `💡 ${M} ${Y} — Bridge por Concepto`, mes.concept, "Concepto");
  addBridgeSlide(pres, charts[3], `🌍 ${Q} ${Y} — Bridge por País`,     q.pais,      "País");
  addBridgeSlide(pres, charts[4], `🎯 ${Q} ${Y} — Bridge por Cluster`,  q.cluster,   "Cluster");
  addBridgeSlide(pres, charts[5], `💡 ${Q} ${Y} — Bridge por Concepto`, q.concept,   "Concepto");
  addBridgeSlide(pres, charts[6], `📊 ${M} ${Y} — Bridge por LOB`,      mes.lob,     "LOB");


  let ci = 8;
  LOB_ORDER.forEach(lob => {
    addBridgeSlide(pres, charts[ci],
      `🌍 ${lob} ${M} — Bridge por País`,    mes.lobBreak[lob].pais,    "País");
    addBridgeSlide(pres, charts[ci + 1],
      `🎯 ${lob} ${M} — Bridge por Cluster`, mes.lobBreak[lob].cluster, "Cluster");
    ci += 2;
  });


  addBridgeSlide(pres, charts[7], `📊 ${Q} ${Y} — Bridge por LOB`, q.lob, "LOB");


  LOB_ORDER.forEach(lob => {
    addBridgeSlide(pres, charts[ci],
      `🌍 ${lob} ${Q} — Bridge por País`,    q.lobBreak[lob].pais,    "País");
    addBridgeSlide(pres, charts[ci + 1],
      `🎯 ${lob} ${Q} — Bridge por Cluster`, q.lobBreak[lob].cluster, "Cluster");
    ci += 2;
  });


  LOB_ORDER.forEach((lob, i) => {
    addTrendSlide(pres, charts[25 + i],
      `📈 Tendencia ${lob} — ${trendFYRange}`,
      `${displayName(CONFIG.FIELD_COLUMNS)} vs ${activeBudgetLabel_()} · ${lob} · desde ${trendStartLbl}`,
      lob, allData);
  });


  return pres;
}


// -----------------------------------------
// SLIDES: PORTADA
// -----------------------------------------
function addCoverSlide(pres, M, Y, Q, mes, q) {
  const slide = pres.appendSlide();
  const W = pres.getPageWidth();
  const H = pres.getPageHeight();


  slide.getBackground().setSolidFill(C.white);
  addRect(slide, 0, 0, W, H * 0.38, C.purple);
  addRect(slide, 0, H * 0.38, W, 4, C.purpleDark);


  applyText(slide.insertTextBox("DESPEGAR", W - 160, 18, 140, 30), 13, true, C.white, "RIGHT");
  applyText(slide.insertTextBox("Media Sales\nDespegar", 50, H * 0.05, W - 100, H * 0.20), 30, true, C.white, "LEFT");
  applyText(slide.insertTextBox(`${M} + ${Q} ${Y}  vs  Forecast\nAnálisis Bridge Waterfall`, 50, H * 0.26, W - 100, 60), 14, false, "#D1B3F0", "LEFT");


  const totMes = getTotal(mes.pais);
  const totQ   = getTotal(q.pais);


  const kpis = [
    { label: `${displayName(CONFIG.FIELD_COLUMNS)} ${M}`, value: fmtM1(totMes.act), color: C.blue },
    { label: `${activeBudgetLabel_()} ${M}`,    value: fmtM1(totMes.bud), color: C.gray },
    { label: `Delta ${M}`, value: `${totMes.del >= 0 ? "+" : ""}${fmtM1(totMes.del)}`,
      sub: `${totMes.del >= 0 ? "+" : ""}${totMes.pct.toFixed(1)}%`, color: totMes.del >= 0 ? C.green : C.red },
    { label: `${displayName(CONFIG.FIELD_COLUMNS)} ${Q}`, value: fmtM1(totQ.act), color: C.blue },
    { label: `${activeBudgetLabel_()} ${Q}`,    value: fmtM1(totQ.bud), color: C.gray },
    { label: `Delta ${Q}`, value: `${totQ.del >= 0 ? "+" : ""}${fmtM1(totQ.del)}`,
      sub: `${totQ.del >= 0 ? "+" : ""}${totQ.pct.toFixed(1)}%`, color: totQ.del >= 0 ? C.green : C.red },
  ];


  const cW = (W - 80) / 3, cH = 84, cardY = H * 0.44;
  kpis.forEach((k, i) => {
    const col = i % 3, row_ = Math.floor(i / 3);
    const x = 40 + col * (cW + 8), y = cardY + row_ * (cH + 10);
    addRect(slide, x, y, cW, cH, C.grayLight);
    addRect(slide, x, y, 4, cH, k.color);
    applyText(slide.insertTextBox(k.value, x + 8, y + 8, cW - 12, 36), 18, true, k.color, "LEFT");
    if (k.sub) applyText(slide.insertTextBox(k.sub, x + 8, y + 44, cW - 12, 20), 11, false, k.color, "LEFT");
    applyText(slide.insertTextBox(k.label, x + 8, y + 64, cW - 12, 16), 8, false, C.subText, "LEFT");
  });


  applyText(slide.insertTextBox("Generado automáticamente · Fuente: dataset", 0, H - 22, W, 18), 7, false, C.subText, "CENTER");
}


// -----------------------------------------
// SLIDES: BRIDGE
// -----------------------------------------
function addBridgeSlide(pres, embeddedChart, title, tableData, categoryLabel) {
  const slide = pres.appendSlide();
  const W = pres.getPageWidth(), H = pres.getPageHeight();


  slide.getBackground().setSolidFill(C.white);
  addRect(slide, 0, 0, W, 58, C.purple);
  addRect(slide, 0, 58, W, 3, C.purpleDark);
  applyText(slide.insertTextBox(title, 15, 10, W * 0.80, 38), 16, true, C.white, "LEFT");
  applyText(slide.insertTextBox(`${CONFIG.TARGET_MONTH} · ${CONFIG.TARGET_YEAR}`, W - 160, 15, 148, 28), 10, false, "#D1B3F0", "RIGHT");


  const chartW = W * 0.62;
  if (embeddedChart) {
    try { slide.insertSheetsChart(embeddedChart, 5, 66, chartW, H - 88); }
    catch (e) {
      try { slide.insertImage(embeddedChart.getAs("image/png"), 5, 66, chartW, H - 88); }
      catch (e2) { Logger.log(`No se pudo insertar gráfico: ${e2.message}`); }
    }
  }


  addBridgeTable(slide, tableData, categoryLabel, chartW + 12, 66, W - chartW - 20, H - 88);
  applyText(slide.insertTextBox(`Media Sales Despegar  ·  ${CONFIG.TARGET_MONTH} ${CONFIG.TARGET_YEAR}  ·  Fuente: dataset`, 0, H - 18, W, 16), 6, false, C.subText, "CENTER");
}


// -----------------------------------------
// SLIDES: TENDENCIA
// -----------------------------------------
function addTrendSlide(pres, embeddedChart, title, subtitle, lobFilter, allData) {
  const slide = pres.appendSlide();
  const W = pres.getPageWidth(), H = pres.getPageHeight();


  slide.getBackground().setSolidFill(C.white);
  addRect(slide, 0, 0, W, 58, C.purple);
  addRect(slide, 0, 58, W, 3, C.purpleDark);
  applyText(slide.insertTextBox(title, 15, 8, W * 0.82, 32), 16, true, C.white, "LEFT");
  applyText(slide.insertTextBox(subtitle, 15, 38, W * 0.82, 18), 9, false, "#D1B3F0", "LEFT");


  const chartW = W * 0.62;
  if (embeddedChart) {
    try { slide.insertSheetsChart(embeddedChart, 5, 66, chartW, H - 88); }
    catch (e) {
      try { slide.insertImage(embeddedChart.getAs("image/png"), 5, 66, chartW, H - 88); }
      catch (e2) { Logger.log(`No se pudo insertar trend: ${e2.message}`); }
    }
  }


  const trendRows = buildTrendData(allData, lobFilter).slice(1);
  addTrendTable(slide, trendRows, chartW + 12, 66, W - chartW - 20, H - 88);
  applyText(slide.insertTextBox(`Media Sales Despegar  ·  ${CONFIG.TARGET_MONTH} ${CONFIG.TARGET_YEAR}  ·  Fuente: dataset`, 0, H - 18, W, 16), 6, false, C.subText, "CENTER");
}


// -----------------------------------------
// TABLAS EN SLIDES
// -----------------------------------------
function addBridgeTable(slide, rows, catLabel, x, y, w, h) {
  const headers = [catLabel, activeBudgetLabel_(), displayName(CONFIG.FIELD_COLUMNS), "Delta", "Pct"];
  const allRows = [headers, ...rows];
  const rowH    = Math.min(20, h / (allRows.length + 1));
  const colW    = [w*0.27, w*0.17, w*0.18, w*0.19, w*0.19];


  allRows.forEach((r, i) => {
    const ry      = y + i * rowH;
    const isHead  = i === 0;
    const isTotal = !isHead && String(r[0]).toUpperCase() === "TOTAL";
    const isEven  = !isHead && i % 2 === 0;
    const bg = isHead ? C.purple : isTotal ? C.purpleLight : isEven ? C.grayLight : C.white;


    addRect(slide, x, ry, w, rowH - 1, bg);
    if (isTotal) addRect(slide, x, ry, 3, rowH - 1, C.purple);


    const vals = isHead ? r : [
      r[0], fmtM1(r[1]), fmtM1(r[2]),
      `${r[3] >= 0 ? "+" : ""}${fmtM1(r[3])}`,
      `${r[4] >= 0 ? "+" : ""}${r[4].toFixed(1)}%`,
    ];


    let cx = x;
    vals.forEach((v, ci) => {
      const cell  = slide.insertTextBox(String(v), cx + 2, ry + 2, colW[ci] - 3, rowH - 4);
      const fsize = isHead ? 7 : isTotal ? 7.5 : 6.5;
      const bold  = isHead || isTotal;
      const color = isHead  ? C.white
                  : isTotal ? C.purpleDark
                  : ci === 3 ? (r[3] >= 0 ? C.green : C.red)
                  : ci === 4 ? (r[4] >= 0 ? C.green : C.red)
                  : C.black;
      applyText(cell, fsize, bold, color, ci === 0 ? "LEFT" : "CENTER");
      cx += colW[ci];
    });
  });
}


function addTrendTable(slide,trendRows,x,y,w,h){
  if(!trendRows||trendRows.length===0)return;
  const headers=["Mes",activeBudgetLabel_(),"Actuals","\u0394","\u0394%"];
  const HEAD_H=20,MIN_ROW=14,bodyH=h-HEAD_H;
  const maxRows=Math.floor(bodyH/MIN_ROW);
  const visRows=trendRows.length>maxRows?trendRows.slice(-maxRows):trendRows;
  const rowH=Math.max(MIN_ROW,Math.min(18,bodyH/Math.max(visRows.length,1)));
  const colW=[w*0.13,w*0.20,w*0.20,w*0.22,w*0.25];
  addRect(slide,x,y,w,HEAD_H-1,C.purple);
  let cx=x;
  headers.forEach((v,ci)=>{
    applyText(slide.insertTextBox(String(v),cx+2,y+2,colW[ci]-3,HEAD_H-4),6.5,true,C.white,ci===0?"LEFT":"CENTER");
    cx+=colW[ci];
  });
  visRows.forEach((r,i)=>{
    const ry=y+HEAD_H+i*rowH;
    const act=parseFloat(r[1]),bud=parseFloat(r[3]);
    const del=act-bud,pct=bud!==0?(del/bud*100):0,isPos=del>=0;
    addRect(slide,x,ry,w,rowH-1,i%2===0?C.grayLight:C.white);
    const vals=[r[0],fmtM1(bud*1e6),fmtM1(act*1e6),`${isPos?"+":""}${fmtM1(del*1e6)}`,`${pct>=0?"+":""}${pct.toFixed(1)}%`];
    let cx=x;
    vals.forEach((v,ci)=>{
      applyText(slide.insertTextBox(String(v),cx+2,ry+1,colW[ci]-3,rowH-2),7,false,(ci===3||ci===4)?(isPos?C.green:C.red):C.black,ci===0?"LEFT":"CENTER");
      cx+=colW[ci];
    });
  });
}


// -----------------------------------------
// UTILIDADES COMPARTIDAS (Slides)
// -----------------------------------------
function displayName(field) {
  return field
    .replace("Budget + Forecast + Breakthrough", "Budget + FC1 + Breakthrough")
    .replace("Budget + Forecast", "Budget + FC1");
}



function activeBudgetLabel_() {
  var fy=getFYInfo_(CONFIG.TARGET_MONTH,CONFIG.TARGET_YEAR).fy;
  return fy>=2027?'Budget FY27':displayName(CONFIG.FIELD_LINE);
}

function getTotal(rows) {
  const t = rows[rows.length - 1];
  const act = t[2], bud = t[1], del = act - bud;
  return { act, bud, del, pct: bud !== 0 ? (del / bud) * 100 : 0 };
}


function addRect(slide, x, y, w, h, color) {
  const s = slide.insertShape(SlidesApp.ShapeType.RECTANGLE, x, y, w, h);
  s.getFill().setSolidFill(color);
  s.getBorder().setTransparent();
  return s;
}


function applyText(shape, size, bold, color, align) {
  shape.getBorder().setTransparent();
  shape.getFill().setTransparent();
  shape.getText().getTextStyle().setFontSize(size).setBold(bold).setForegroundColor(color);
  const al = align === "CENTER" ? SlidesApp.ParagraphAlignment.CENTER
           : align === "RIGHT"  ? SlidesApp.ParagraphAlignment.END
           :                      SlidesApp.ParagraphAlignment.START;
  shape.getText().getParagraphStyle().setParagraphAlignment(al);
}


function fmtM1(v) {
  const abs = Math.abs(v) / 1e6;
  return `${v < 0 ? "-" : ""}$${abs.toFixed(1)}M`;
}


function toast(msg) {
  SpreadsheetApp.getActiveSpreadsheet().toast(msg, "📊 Dashboard", 5);
}


function deleteHelperSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ws = ss.getSheetByName(CONFIG.HELPER_SHEET);
  if (ws) { ss.deleteSheet(ws); toast("🗑 Hoja auxiliar eliminada."); }
}


function resetPresentation() {
  PropertiesService.getScriptProperties().deleteProperty("PRESENTATION_ID");
  toast("🔄 Reseteado. La próxima ejecución crea una nueva presentación.");
}




// ============================================================
//  PARTE 2 — WEB APP DASHBOARD
// ============================================================


function doGet(e) {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('Media Sales Despegar — Dashboard')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}


// -- Normalización de headers ---------------------------------
// "País" = "Pais" = "PAIS" = "pais" etc.
function norm_(h) {
  return String(h || '').trim()
    .replace(/[\u00c1\u00e1\u00c0\u00e0\u00c2\u00e2\u00c4\u00e4]/g, 'a')
    .replace(/[\u00c9\u00e9\u00c8\u00e8\u00ca\u00ea\u00cb\u00eb]/g, 'e')
    .replace(/[\u00cd\u00ed\u00cc\u00ec\u00ce\u00ee\u00cf\u00ef]/g, 'i')
    .replace(/[\u00d3\u00f3\u00d2\u00f2\u00d4\u00f4\u00d6\u00f6]/g, 'o')
    .replace(/[\u00da\u00fa\u00d9\u00f9\u00db\u00fb\u00dc\u00fc]/g, 'u')
    .replace(/[\u00d1\u00f1]/g, 'n')
    .toLowerCase();
}


function col_(headers, candidates) {
  var nh = headers.map(norm_);
  for (var i = 0; i < candidates.length; i++) {
    var idx = nh.indexOf(norm_(candidates[i]));
    if (idx !== -1) return idx;
  }
  return -1;
}


// -- Datos principales ----------------------------------------
function getDashboardData(targetMonth, targetYear, groupBy, filterPartners, filterClusters) {
  try {
    groupBy        = groupBy        || 'M';
    filterPartners = (Array.isArray(filterPartners) ? filterPartners : []).map(function(s){return String(s).trim();});
    filterClusters = (Array.isArray(filterClusters) ? filterClusters : []).map(function(s){return String(s).trim();});

    var ss      = SpreadsheetApp.getActiveSpreadsheet();
    var dsSheet = ss.getSheetByName('dataset');
    if (!dsSheet) return { error: "No se encontró 'dataset'." };


    var raw     = dsSheet.getDataRange().getValues();
    var headers = raw[0].map(function(h) { return String(h).trim(); });


    var MONTH_ORD_FY = ['January','February','March','April','May','June',
                     'July','August','September','October','November','December'];

    // FY27 = desde April 2026 en adelante
    function isFY27_(mo, yr) {
      if (Number(yr) > 2026) return true;
      if (Number(yr) < 2026) return false;
      return MONTH_ORD_FY.indexOf(mo) >= MONTH_ORD_FY.indexOf('April');
    }

    var idx = {
      year   : col_(headers, ['Date - Year',    'Year',    'Año', 'Ano']),
      quarter: col_(headers, ['Date - Quarter',  'Quarter', 'Trimestre']),
      month  : col_(headers, ['Date - Month',    'Month',   'Mes']),
      country: col_(headers, ['País', 'Pais', 'Country']),
      channel: col_(headers, ['Lob-Channel', 'Lob Channel', 'Canal']),
      concept: col_(headers, ['Concepto', 'Concept']),
      partner: col_(headers, ['Partner']),
      cluster: col_(headers, ['Cluster']),
      arr    : col_(headers, ['Actuals + RR', 'Actuals+RR']),
      bf26   : col_(headers, ['Budget + Forecast FY26',    'Budget + Forecast',    'Budget+Forecast']),
      bfb26  : col_(headers, ['Budget + Forecast + Breakthrough FY26',
                               'Budget + Forecast + Breakthrough', 'Budget+Forecast+Breakthrough']),
      bf27   : col_(headers, ['Budget FY27']),
    };

    // missing: sólo los campos obligatorios
    var missing = ['year','month','country','arr'].filter(function(k){ return idx[k]===-1; });


    if (missing.length > 0) {
      return { error: 'Cols no encontradas en dataset: ' + missing.join(', ') };
    }


    var MONTH_ORD = ['January','February','March','April','May','June',
                     'July','August','September','October','November','December'];


    var allRows = raw.slice(1).map(function(r) {
      var mo  = String(r[idx.month]  || '').trim();
      var yr  = String(r[idx.year]   || '').trim();
      var fy7 = isFY27_(mo, yr);
      var bf26v  = idx.bf26  >= 0 ? parseWA_(r[idx.bf26])  : 0;
      var bfb26v = idx.bfb26 >= 0 ? parseWA_(r[idx.bfb26]) : 0;
      var bf27v  = idx.bf27  >= 0 ? parseWA_(r[idx.bf27])  : 0;
      return {
        year   : yr,
        quarter: String(r[idx.quarter] || '').trim(),
        month  : mo,
        country: String(r[idx.country] || '').trim(),
        channel: String(r[idx.channel] || '').trim(),
        concept: String(r[idx.concept] || '').trim(),
        partner: idx.partner >= 0 ? String(r[idx.partner] || '').trim() : '',
        cluster: String(r[idx.cluster] || '').trim(),
        arr    : parseWA_(r[idx.arr]),
        bf     : fy7 ? bf27v  : bf26v,
        bfb    : fy7 ? bf27v  : bfb26v,
        fy27   : fy7,
      };
    }).filter(function(r) {
      return r.month && r.month !== 'Month' &&
             r.year  && r.year  !== 'Year'  &&
             r.country && r.country !== '';
    });


    var periods;
    if (groupBy === 'Q')      periods = getQPeriods_(allRows);
    else if (groupBy === 'H') periods = getHPeriods_(allRows);
    else                      periods = getPeriods_(allRows, MONTH_ORD);

    // Auto-detect period format mismatch (e.g. "March" passed when groupBy='Q')
    var _needReset = !targetMonth || !targetYear || targetMonth==='' || targetYear==='';
    if (!_needReset && groupBy === 'Q' && !/^Q\d/.test(String(targetMonth))) _needReset = true;
    if (!_needReset && groupBy === 'H' && !/^H\d/.test(String(targetMonth))) _needReset = true;
    if (_needReset) {
      var last = periods[periods.length - 1];
      targetMonth = last.month;
      targetYear  = last.year;
    }
    targetYear  = String(targetYear).trim();
    targetMonth = String(targetMonth).trim();

    // Filas base del período (sin filtros de partner/cluster)
    var basePeriodRows;
    if (groupBy === 'Q') {
      var _tQ = parseInt(String(targetMonth).replace('Q','')), _tFY = parseInt(targetYear);
      basePeriodRows = allRows.filter(function(r) {
        var fi = getFYInfo_(r.month, r.year); return fi.fy === _tFY && fi.q === _tQ;
      });
    } else if (groupBy === 'H') {
      var _tH = parseInt(String(targetMonth).replace('H','')), _tFY2 = parseInt(targetYear);
      basePeriodRows = allRows.filter(function(r) {
        var fi = getFYInfo_(r.month, r.year); return fi.fy === _tFY2 && fi.h === _tH;
      });
    } else {
      basePeriodRows = allRows.filter(function(r) {
        return r.month === targetMonth && r.year === targetYear;
      });
    }
    if (basePeriodRows.length === 0) {
      return { error: 'Sin datos para ' + targetMonth + ' ' + targetYear };
    }

    // Listas de opciones para filtros (desde el período sin filtrar)
    var partnerSet = {}, clusterSet = {};
    basePeriodRows.forEach(function(r) {
      if (r.partner && r.partner.trim()) partnerSet[r.partner.trim()] = true;
      if (r.cluster && r.cluster.trim()) clusterSet[r.cluster.trim()] = true;
    });
    var partnerList = Object.keys(partnerSet).sort();
    var clusterList = Object.keys(clusterSet).sort();

    // Aplicar filtros de partner y cluster
    var filtered = basePeriodRows;
    if (filterPartners.length > 0) {
      filtered = filtered.filter(function(r) { return filterPartners.indexOf(r.partner) !== -1; });
    }
    if (filterClusters.length > 0) {
      filtered = filtered.filter(function(r) { return filterClusters.indexOf(r.cluster) !== -1; });
    }
    if (filtered.length === 0) {
      return { error: 'Sin datos para los filtros seleccionados.' };
    }

    // Etiqueta de período para la UI
    var periodLabel;
    if (groupBy === 'Q')      periodLabel = 'Q' + parseInt(targetMonth.replace('Q','')) + ' FY' + String(targetYear).slice(-2);
    else if (groupBy === 'H') periodLabel = 'H' + parseInt(targetMonth.replace('H','')) + ' FY' + String(targetYear).slice(-2);
    else                      periodLabel = targetMonth + ' ' + targetYear;


    // GB por país — suma todos los canales de data_gb
    // S/C no existe en data_gb → nunca entra en el denominador
    var gbByCountry = buildGBByCountry_(ss, targetMonth, targetYear);
    var hasGB       = gbByCountry !== null;


    var rowsGB = filtered.map(function(r) {
      var gb = hasGB ? (gbByCountry[r.country] || { gbArr:0, gbBf:0, gbBfb:0 })
                     : { gbArr:0, gbBf:0, gbBfb:0 };
      return {
        year: r.year, quarter: r.quarter, month: r.month,
        country: r.country, channel: r.channel,
        concept: r.concept, partner: r.partner, cluster: r.cluster,
        arr: r.arr, bf: r.bf, bfb: r.bfb,
        gbArr: gb.gbArr, gbBf: gb.gbBf, gbBfb: gb.gbBfb,
      };
    });


    var seenC = {}, countries = [];
    rowsGB.forEach(function(r) {
      if (!seenC[r.country] && r.country) {
        seenC[r.country] = true;
        countries.push(r.country);
      }
    });
    countries.sort(function(a, b) {
      var sA = rowsGB.filter(function(r){ return r.country === a; }).reduce(function(s,r){ return s+r.arr; }, 0);
      var sB = rowsGB.filter(function(r){ return r.country === b; }).reduce(function(s,r){ return s+r.arr; }, 0);
      return sB - sA;
    });


    var general = {
      byCountry: aggBy_(rowsGB, 'country', gbByCountry),
      byCluster: aggBy_(rowsGB, 'cluster', gbByCountry),
      byConcept: aggBy_(rowsGB, 'concept',  gbByCountry),
    };
    var byCountry = {};
    countries.forEach(function(c) {
      var rc = rowsGB.filter(function(r) { return r.country === c; });
      byCountry[c] = {
        byCluster: aggBy_(rc, 'cluster', gbByCountry),
        byConcept: aggBy_(rc, 'concept',  gbByCountry),
      };
    });


    var totArr = rowsGB.reduce(function(s,r){ return s+r.arr; }, 0);
    var totBf  = rowsGB.reduce(function(s,r){ return s+r.bf;  }, 0);
    var totBfb = rowsGB.reduce(function(s,r){ return s+r.bfb; }, 0);
    var totGB  = sumCountryGBs_(countries, gbByCountry);


    var evolutivo = buildEvolutivoFromRows_(ss, allRows);
    var yoyData   = buildYoYData_(allRows, targetMonth, targetYear, gbByCountry, groupBy);

    var periodIsFY27 = isFY27_(targetMonth, targetYear);

    return {
      targetMonth    : targetMonth,
      targetYear     : targetYear,
      periodLabel    : periodLabel,
      groupBy        : groupBy,
      isFY27         : periodIsFY27,
      bfLabel        : periodIsFY27 ? 'Budget FY27' : 'Budget+FC FY26',
      bfbLabel       : periodIsFY27 ? 'Budget FY27' : 'Breakthrough FY26',
      countries      : countries,
      periods        : periods,
      hasGB          : hasGB,
      partnerList    : partnerList,
      clusterList    : clusterList,
      filterPartners : filterPartners,
      filterClusters : filterClusters,
      totals: {
        arr: totArr, bf: totBf, bfb: totBfb,
        gbArr: totGB.gbArr, gbBf: totGB.gbBf, gbBfb: totGB.gbBfb,
      },
      general   : general,
      byCountry : byCountry,
      evolutivo : evolutivo,
      yoy       : yoyData,
    };


  } catch(e) {
    return { error: 'Error: ' + e.message };
  }
}


// -- GB por país (suma todos los canales) ---------------------
function buildGBByCountry_(ss, targetMonth, targetYear) {
  var gbSheet = ss.getSheetByName('data_gb');
  if (!gbSheet) return null;


  var raw     = gbSheet.getDataRange().getValues();
  var headers = raw[0].map(function(h) { return String(h).trim(); });


  var iYear    = col_(headers, ['Año', 'Ano', 'Year']);
  var iMonth   = col_(headers, ['Month', 'Mes']);
  var iCountry = col_(headers, ['País', 'Pais', 'Country']);
  var iGbArr   = col_(headers, ['GB Actuals + RR', 'GB Actuals+RR']);
  var iGbBf    = col_(headers, ['GB Budget + Forecast', 'GB Budget+Forecast']);
  var iGbBfb   = col_(headers, ['GB Budget + Forecast + Breakthrough', 'GB Budget+Forecast+Breakthrough']);


  if (iYear === -1 || iMonth === -1 || iCountry === -1) return null;


  var tyStr = String(targetYear).trim();
  var tmStr = String(targetMonth).trim();
  var result = {};


  for (var i = 1; i < raw.length; i++) {
    var r = raw[i];
    if (String(r[iYear]).trim()  !== tyStr) continue;
    if (String(r[iMonth]).trim() !== tmStr) continue;
    var country = String(r[iCountry] || '').trim();
    if (!country || country === 'nan' || country === 'undefined') continue;
    if (!result[country]) result[country] = { gbArr:0, gbBf:0, gbBfb:0 };
    result[country].gbArr += parseWA_(iGbArr !== -1 ? r[iGbArr] : 0);
    result[country].gbBf  += parseWA_(iGbBf  !== -1 ? r[iGbBf]  : 0);
    result[country].gbBfb += parseWA_(iGbBfb !== -1 ? r[iGbBfb] : 0);
  }


  if (Object.keys(result).length === 0) return null;
  return result;
}



function buildGBForQuarter_(ss, targetMonth, targetYear) {
  var gbSheet = ss.getSheetByName('data_gb');
  if (!gbSheet) return null;
  var Q_MONTHS = {January:['January','February','March'],February:['January','February','March'],March:['January','February','March'],April:['April','May','June'],May:['April','May','June'],June:['April','May','June'],July:['July','August','September'],August:['July','August','September'],September:['July','August','September'],October:['October','November','December'],November:['October','November','December'],December:['October','November','December']};
  var qMonths=Q_MONTHS[targetMonth]||[targetMonth];
  var raw=gbSheet.getDataRange().getValues();
  var headers=raw[0].map(function(h){return String(h).trim();});
  var iYear=col_(headers,['Año','Ano','Year']),iMonth=col_(headers,['Month','Mes']),iCountry=col_(headers,['País','Pais','Country']);
  var iGbArr=col_(headers,['GB Actuals + RR','GB Actuals+RR']),iGbBf=col_(headers,['GB Budget + Forecast','GB Budget+Forecast']),iGbBfb=col_(headers,['GB Budget + Forecast + Breakthrough','GB Budget+Forecast+Breakthrough']);
  if(iYear===-1||iMonth===-1||iCountry===-1)return null;
  var tyStr=String(targetYear).trim(),result={};
  for(var i=1;i<raw.length;i++){
    var r=raw[i];
    if(String(r[iYear]).trim()!==tyStr)continue;
    if(qMonths.indexOf(String(r[iMonth]).trim())===-1)continue;
    var country=String(r[iCountry]||'').trim();
    if(!country)continue;
    if(!result[country])result[country]={gbArr:0,gbBf:0,gbBfb:0};
    result[country].gbArr+=parseWA_(iGbArr!==-1?r[iGbArr]:0);
    result[country].gbBf+=parseWA_(iGbBf!==-1?r[iGbBf]:0);
    result[country].gbBfb+=parseWA_(iGbBfb!==-1?r[iGbBfb]:0);
  }
  return Object.keys(result).length?result:null;
}

function sumCountryGBs_(countries, gbByCountry) {
  if (!gbByCountry) return { gbArr:0, gbBf:0, gbBfb:0 };
  var gbArr = 0, gbBf = 0, gbBfb = 0;
  countries.forEach(function(c) {
    var gb = gbByCountry[c] || { gbArr:0, gbBf:0, gbBfb:0 };
    gbArr += gb.gbArr; gbBf += gb.gbBf; gbBfb += gb.gbBfb;
  });
  return { gbArr:gbArr, gbBf:gbBf, gbBfb:gbBfb };
}


function aggBy_(rows, field, gbByCountry) {
  var groups = {};
  rows.forEach(function(r) {
    var k = r[field] || 'N/A';
    if (!groups[k]) groups[k] = [];
    groups[k].push(r);
  });


  var result = Object.keys(groups).map(function(label) {
    var grp = groups[label];
    var arr = grp.reduce(function(s,r){ return s+r.arr; }, 0);
    var bf  = grp.reduce(function(s,r){ return s+r.bf;  }, 0);
    var bfb = grp.reduce(function(s,r){ return s+r.bfb; }, 0);
    var seenP = {}, gbArr = 0, gbBf = 0, gbBfb = 0;
    grp.forEach(function(r) {
      if (seenP[r.country]) return;
      seenP[r.country] = true;
      var gb = gbByCountry ? (gbByCountry[r.country] || { gbArr:0, gbBf:0, gbBfb:0 })
                           : { gbArr:0, gbBf:0, gbBfb:0 };
      gbArr += gb.gbArr; gbBf += gb.gbBf; gbBfb += gb.gbBfb;
    });
    // Partners — Pareto 80/20 sobre A+RR del contexto actual.
    // Partners fuera del corte → agrupados como "Others <ClusterNombre>".
    var partnerMap = {};
    grp.forEach(function(r) {
      var pk = r.partner || '';
      if (!pk) return;
      if (!partnerMap[pk]) partnerMap[pk] = { label: pk, arr: 0, bf: 0, bfb: 0, clData: {} };
      partnerMap[pk].arr += r.arr;
      partnerMap[pk].bf  += r.bf;
      partnerMap[pk].bfb += r.bfb;
      var cl = String(r.cluster || '').trim();
      if (!partnerMap[pk].clData[cl]) partnerMap[pk].clData[cl] = { arr:0, bf:0, bfb:0 };
      partnerMap[pk].clData[cl].arr += r.arr;
      partnerMap[pk].clData[cl].bf  += r.bf;
      partnerMap[pk].clData[cl].bfb += r.bfb;
    });
    var allP = Object.keys(partnerMap)
      .map(function(k) { return partnerMap[k]; })
      .filter(function(p) { return p.arr > 0; })
      .sort(function(a, b) { return b.arr - a.arr; });
    var totP = allP.reduce(function(s, p) { return s + p.arr; }, 0);
    var cum = 0, cut = allP.length;
    for (var pi = 0; pi < allP.length; pi++) {
      cum += allP[pi].arr;
      if (totP > 0 && cum / totP >= 0.8) { cut = pi + 1; break; }
    }
    var top80  = allP.slice(0, cut);
    var rest20 = allP.slice(cut);
    // Agrupar rest20 por nombre real del cluster: "Others <Cluster>"
    var othBk = {};
    rest20.forEach(function(p) {
      Object.keys(p.clData).forEach(function(cl) {
        var bk = 'Others' + (cl ? ' ' + cl : '');
        if (!othBk[bk]) othBk[bk] = { arr:0, bf:0, bfb:0 };
        othBk[bk].arr += p.clData[cl].arr;
        othBk[bk].bf  += p.clData[cl].bf;
        othBk[bk].bfb += p.clData[cl].bfb;
      });
    });
    var partners = top80.map(function(p) {
      return { label: p.label, cluster: '', arr: p.arr, bf: p.bf, bfb: p.bfb };
    });
    Object.keys(othBk)
      .map(function(bk) { return { label:bk, cluster:'', arr:othBk[bk].arr, bf:othBk[bk].bf, bfb:othBk[bk].bfb }; })
      .sort(function(a,b) { return b.arr - a.arr; })
      .forEach(function(oe) { if (oe.arr > 0) partners.push(oe); });

    return {
      label: label, isTotal: false,
      arr: arr, bf: bf, bfb: bfb,
      dBf : arr - bf,  dBfb: arr - bfb,
      pctBf : bf  ? (arr - bf)  / bf  * 100 : 0,
      pctBfb: bfb ? (arr - bfb) / bfb * 100 : 0,
      rateArr: gbArr ? arr / gbArr * 100 : 0,
      rateBf : gbBf  ? bf  / gbBf  * 100 : 0,
      rateBfb: gbBfb ? bfb / gbBfb * 100 : 0,
      gbArr: gbArr, gbBf: gbBf, gbBfb: gbBfb,
      partners: partners,
    };
  });


  result.sort(function(a, b) { return b.arr - a.arr; });


  var tArr = result.reduce(function(s,r){ return s+r.arr; }, 0);
  var tBf  = result.reduce(function(s,r){ return s+r.bf;  }, 0);
  var tBfb = result.reduce(function(s,r){ return s+r.bfb; }, 0);
  var allC = Object.keys(rows.reduce(function(acc,r){ acc[r.country]=1; return acc; }, {}));
  var tGB  = sumCountryGBs_(allC, gbByCountry);


  result.push({
    label: 'TOTAL', isTotal: true,
    arr: tArr, bf: tBf, bfb: tBfb,
    dBf : tArr - tBf,  dBfb: tArr - tBfb,
    pctBf : tBf  ? (tArr - tBf)  / tBf  * 100 : 0,
    pctBfb: tBfb ? (tArr - tBfb) / tBfb * 100 : 0,
    rateArr: tGB.gbArr ? tArr / tGB.gbArr * 100 : 0,
    rateBf : tGB.gbBf  ? tBf  / tGB.gbBf  * 100 : 0,
    rateBfb: tGB.gbBfb ? tBfb / tGB.gbBfb * 100 : 0,
    gbArr: tGB.gbArr, gbBf: tGB.gbBf, gbBfb: tGB.gbBfb,
  });


  return result;
}


function getPeriods_(rows, monthOrder) {
  var set = {};
  rows.forEach(function(r) {
    var k = r.year + '__' + r.month;
    if (!set[k]) set[k] = { year: r.year, month: r.month };
  });
  return Object.keys(set).map(function(k) { return set[k]; }).sort(function(a, b) {
    if (a.year !== b.year) return Number(a.year) - Number(b.year);
    return monthOrder.indexOf(a.month) - monthOrder.indexOf(b.month);
  });
}


// CRÍTICO: typeof NaN === 'number' en JS → siempre usar isNaN()
function parseWA_(v) {
  if (v === null || v === undefined || v === '') return 0;
  if (typeof v === 'number') return isNaN(v) ? 0 : v;
  var p = parseFloat(String(v).replace(',', '.'));
  return isNaN(p) ? 0 : p;
}


// -- Debug y test ---------------------------------------------
function debugGBRates() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var gbSheet = ss.getSheetByName('data_gb');
  if (!gbSheet) { Logger.log('ERROR: data_gb no encontrada'); return; }
  var raw = gbSheet.getDataRange().getValues();
  Logger.log('Headers: ' + JSON.stringify(raw[0]));
  var gb = buildGBByCountry_(ss, 'March', '2026');
  Logger.log('gbByCountry: ' + JSON.stringify(gb));
  if (gb) {
    var totArr = 0, totBf = 0;
    Object.keys(gb).forEach(function(c) { totArr += gb[c].gbArr; totBf += gb[c].gbBf; });
    Logger.log('totGbArr=' + totArr + '  (esperado ~645016800)');
    Logger.log('totGbBf='  + totBf  + '  (esperado ~577395300)');
  }
  SpreadsheetApp.getActiveSpreadsheet().toast('Ver Registros', 'Debug OK', 10);
}


function testDashboard() {
  var result = getDashboardData('March', '2026');
  if (result.error) { Logger.log('ERROR: ' + result.error); return; }
  var t = result.totals;
  Logger.log('hasGB   : ' + result.hasGB);
  Logger.log('gbArr   : ' + t.gbArr + '  (esperado ~645016800)');
  Logger.log('gbBf    : ' + t.gbBf  + '  (esperado ~577395300)');
  Logger.log('Rate A+RR: ' + (t.gbArr ? (t.arr/t.gbArr*100).toFixed(4) : 0) + '%  (esperado 0.6352%)');
  Logger.log('Rate B+F : ' + (t.gbBf  ? (t.bf /t.gbBf *100).toFixed(4) : 0) + '%  (esperado 0.4505%)');
  SpreadsheetApp.getActiveSpreadsheet().toast('Ver Registros', 'Test OK', 10);
}


// ─────────────────────────────────────────────────────────────────────
// buildEvolutivoFromRows_  — computa evolutivo histórico.
// Recibe allRows (ya en memoria en getDashboardData) + ss para leer
// data_gb UNA sola vez. Sin loops de I/O adicionales.
// ─────────────────────────────────────────────────────────────────────
function buildEvolutivoFromRows_(ss, allRows) {
  var MONTH_ORD = ['January','February','March','April','May','June',
                   'July','August','September','October','November','December'];
  var BUCKETS = ['MR','Overs','Others'];

  function bucket_(concept) {
    var s = norm_(String(concept || ''));
    if (s.indexOf('produccion') !== -1 || s.indexOf('monto fijo') !== -1 ||
        s === 'mr' || s.indexOf('facturacion') !== -1 || s.indexOf('factura') !== -1)
      return 'MR';
    if (s.indexOf('over') !== -1 || s.indexOf('commission') !== -1 ||
        s.indexOf('comision') !== -1)
      return 'Overs';
    return 'Others';
  }

  // Acumular desde allRows (ya en memoria, sin I/O)
  var pMap = {};
  allRows.forEach(function(r) {
    if (Number(r.year) < 2024) return;
    var pk = r.year + '__' + r.month;
    if (!pMap[pk]) {
      var bk = {};
      BUCKETS.forEach(function(b) { bk[b] = { arr:0, bf:0, bfb:0 }; });
      pMap[pk] = { year:r.year, month:r.month, totArr:0, totBf:0, totBfb:0, bk:bk };
    }
    var e = pMap[pk];
    var b = bucket_(r.concept);
    e.totArr += r.arr;  e.totBf += r.bf;  e.totBfb += r.bfb;
    e.bk[b].arr += r.arr;  e.bk[b].bf += r.bf;  e.bk[b].bfb += r.bfb;
  });

  if (!Object.keys(pMap).length) return null;

  // Leer data_gb UNA vez para todos los períodos
  var gbMap = {};
  var gbS = ss.getSheetByName('data_gb');
  if (gbS) {
    var gbVals = gbS.getDataRange().getValues();
    var hgb   = gbVals[0].map(function(h) { return String(h).trim(); });
    var giY   = col_(hgb, ['Año','Ano','Year']);
    var giM   = col_(hgb, ['Month','Mes']);
    var giC   = col_(hgb, ['País','Pais','Country']);
    var giArr = col_(hgb, ['GB Actuals + RR','GB Actuals+RR']);
    var giBf  = col_(hgb, ['GB Budget + Forecast','GB Budget+Forecast']);
    var giBfb = col_(hgb, ['GB Budget + Forecast + Breakthrough','GB Budget+Forecast+Breakthrough']);
    if (giY >= 0 && giM >= 0) {
      var seen = {};
      for (var i = 1; i < gbVals.length; i++) {
        var gr = gbVals[i];
        var gy = String(gr[giY] || '').trim();
        var gm = String(gr[giM] || '').trim();
        var gc = giC >= 0 ? String(gr[giC] || '').trim() : '__';
        if (!gy || !gm) continue;
        var pk  = gy + '__' + gm;
        var ck  = pk + '__' + gc;
        if (seen[ck]) continue;
        seen[ck] = true;
        if (!gbMap[pk]) gbMap[pk] = { gbArr:0, gbBf:0, gbBfb:0 };
        if (giArr >= 0) gbMap[pk].gbArr += parseWA_(gr[giArr]);
        if (giBf  >= 0) gbMap[pk].gbBf  += parseWA_(gr[giBf]);
        if (giBfb >= 0) gbMap[pk].gbBfb += parseWA_(gr[giBfb]);
      }
    }
  }

  var hasGB = Object.keys(gbMap).some(function(k) { return gbMap[k].gbArr > 0; });
  var abbr  = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];

  function rate_(v, g) { return (g && g > 0) ? v / g * 100 : null; }

  var sorted = Object.keys(pMap).sort(function(a, b) {
    var pa = pMap[a], pb = pMap[b];
    if (pa.year !== pb.year) return Number(pa.year) - Number(pb.year);
    return MONTH_ORD.indexOf(pa.month) - MONTH_ORD.indexOf(pb.month);
  });

  var periods = sorted.map(function(pk) {
    var e  = pMap[pk];
    var gb = gbMap[pk] || { gbArr:0, gbBf:0, gbBfb:0 };
    var mi = MONTH_ORD.indexOf(e.month);
    var lbl = (mi >= 0 ? abbr[mi] : e.month.slice(0,3)) + "'" + e.year.slice(2);
    var bkOut = {};
    BUCKETS.forEach(function(b) {
      var bv = e.bk[b];
      bkOut[b] = {
        arr: bv.arr, bf: bv.bf, bfb: bv.bfb,
        rateArr: rate_(bv.arr, gb.gbArr),
        rateBf:  rate_(bv.bf,  gb.gbBf),
        rateBfb: rate_(bv.bfb, gb.gbBfb)
      };
    });
    return {
      year: e.year, month: e.month, label: lbl,
      totArr: e.totArr, totBf: e.totBf, totBfb: e.totBfb,
      gbArr: gb.gbArr,
      rateArr: rate_(e.totArr, gb.gbArr),
      rateBf:  rate_(e.totBf,  gb.gbBf),
      rateBfb: rate_(e.totBfb, gb.gbBfb),
      bk: bkOut
    };
  });

  return { periods: periods, hasGB: hasGB };
}

// ─────────────────────────────────────────────────────────────────────
// buildYoYData_
// CY = targetMonth/targetYear   PY = mismo mes, año anterior.
// 80/20 de partners: calculado sobre CY.  Resto → "Others <Cluster>".
// ─────────────────────────────────────────────────────────────────────
function buildYoYData_(allRows, targetMonth, targetYear, gbByCountry, groupBy) {
  var cyYear = String(targetYear).trim();
  var pyYear = String(Number(cyYear) - 1);
  var mo     = String(targetMonth).trim();
  var grpBy  = groupBy || 'M';
  var cyRows, pyRows;
  if (grpBy === 'Q') {
    var _cQ=parseInt(mo.replace('Q','')), _cFY=parseInt(cyYear), _pFY=parseInt(pyYear);
    cyRows=allRows.filter(function(r){var fi=getFYInfo_(r.month,r.year);return fi.fy===_cFY&&fi.q===_cQ;});
    pyRows=allRows.filter(function(r){var fi=getFYInfo_(r.month,r.year);return fi.fy===_pFY&&fi.q===_cQ;});
  } else if (grpBy === 'H') {
    var _cH=parseInt(mo.replace('H','')), _cFY2=parseInt(cyYear), _pFY2=parseInt(pyYear);
    cyRows=allRows.filter(function(r){var fi=getFYInfo_(r.month,r.year);return fi.fy===_cFY2&&fi.h===_cH;});
    pyRows=allRows.filter(function(r){var fi=getFYInfo_(r.month,r.year);return fi.fy===_pFY2&&fi.h===_cH;});
  } else {
    cyRows=allRows.filter(function(r){return r.month===mo&&r.year===cyYear;});
    pyRows=allRows.filter(function(r){return r.month===mo&&r.year===pyYear;});
  }

  function yoyAggBy_(cyR, pyR, field) {
    var cyMap = {};
    cyR.forEach(function(r) {
      var k = r[field] || 'N/A';
      if (!cyMap[k]) cyMap[k] = [];
      cyMap[k].push(r);
    });
    var pyTotByKey = {}, pyPartnersByKey = {};
    pyR.forEach(function(r) {
      var k  = r[field] || 'N/A';
      var pk = r.partner || '__';
      pyTotByKey[k] = (pyTotByKey[k] || 0) + r.arr;
      if (!pyPartnersByKey[k]) pyPartnersByKey[k] = {};
      pyPartnersByKey[k][pk] = (pyPartnersByKey[k][pk] || 0) + r.arr;
    });

    var result = Object.keys(cyMap).map(function(label) {
      var grp   = cyMap[label];
      var cyArr = grp.reduce(function(s, r) { return s + r.arr; }, 0);
      var pyArr = pyTotByKey[label] || 0;
      var seenP = {}, gbArr = 0;
      grp.forEach(function(r) {
        if (seenP[r.country]) return;
        seenP[r.country] = true;
        var gb = gbByCountry ? (gbByCountry[r.country] || { gbArr:0 }) : { gbArr:0 };
        gbArr += gb.gbArr;
      });

      // 80/20 on CY
      var pMap = {};
      grp.forEach(function(r) {
        var pk = r.partner || ''; if (!pk) return;
        if (!pMap[pk]) pMap[pk] = { label:pk, arr:0, clArr:{} };
        pMap[pk].arr += r.arr;
        var cl = String(r.cluster || '').trim();
        pMap[pk].clArr[cl] = (pMap[pk].clArr[cl] || 0) + r.arr;
      });
      var allP = Object.keys(pMap).map(function(k){return pMap[k];})
        .filter(function(p){return p.arr>0;}).sort(function(a,b){return b.arr-a.arr;});
      var totP = allP.reduce(function(s,p){return s+p.arr;},0);
      var cum=0, cut=allP.length;
      for(var pi=0;pi<allP.length;pi++){
        cum+=allP[pi].arr;
        if(totP>0&&cum/totP>=0.8){cut=pi+1;break;}
      }
      var top80=allP.slice(0,cut), rest20=allP.slice(cut);
      var pyLookup = pyPartnersByKey[label] || {};

      // Others buckets por cluster real
      var othBkCY={}, othBkPY={};
      rest20.forEach(function(p){
        var pPY=pyLookup[p.label]||0;
        Object.keys(p.clArr).forEach(function(cl){
          var bk='Others'+(cl?' '+cl:'');
          othBkCY[bk]=(othBkCY[bk]||0)+p.clArr[cl];
        });
        if(pPY>0) othBkPY['Others']=(othBkPY['Others']||0)+pPY;
      });

      var partners=top80.map(function(p){
        var pyCurr=pyLookup[p.label]||0, d=p.arr-pyCurr;
        return{label:p.label,cluster:'',cyArr:p.arr,pyArr:pyCurr,delta:d,pct:pyCurr?d/pyCurr*100:0};
      });
      var allBk=Object.keys(othBkCY).concat(Object.keys(othBkPY))
        .filter(function(bk,i,a){return a.indexOf(bk)===i;});
      allBk.map(function(bk){
        var cyV=othBkCY[bk]||0,pyV=othBkPY[bk]||0,d=cyV-pyV;
        return{label:bk,cluster:'',cyArr:cyV,pyArr:pyV,delta:d,pct:pyV?d/pyV*100:0};
      }).filter(function(o){return o.cyArr>0||o.pyArr>0;})
        .sort(function(a,b){return b.cyArr-a.cyArr;})
        .forEach(function(o){partners.push(o);});

      var delta=cyArr-pyArr;
      return{label:label,isTotal:false,cyArr:cyArr,pyArr:pyArr,
        delta:delta,pct:pyArr?delta/pyArr*100:0,gbArr:gbArr,partners:partners};
    });
    result.sort(function(a,b){return b.cyArr-a.cyArr;});
    Object.keys(pyTotByKey).forEach(function(k){
      if(!cyMap[k]&&pyTotByKey[k]>0){
        var pyV=pyTotByKey[k];
        result.push({label:k,isTotal:false,cyArr:0,pyArr:pyV,delta:-pyV,pct:-100,gbArr:0,partners:[]});
      }
    });
    var tCY=result.reduce(function(s,r){return s+r.cyArr;},0);
    var tPY=result.reduce(function(s,r){return s+r.pyArr;},0);
    var tD=tCY-tPY;
    result.push({label:'TOTAL',isTotal:true,cyArr:tCY,pyArr:tPY,
      delta:tD,pct:tPY?tD/tPY*100:0,gbArr:0,partners:[]});
    return result;
  }

  var seenC={}, countries=[];
  cyRows.forEach(function(r){if(!seenC[r.country]&&r.country){seenC[r.country]=true;countries.push(r.country);}});
  var byCountry_detail={};
  countries.forEach(function(c){
    var cyCR=cyRows.filter(function(r){return r.country===c;});
    var pyCR=pyRows.filter(function(r){return r.country===c;});
    byCountry_detail[c]={
      byCluster:yoyAggBy_(cyCR,pyCR,'cluster'),
      byConcept:yoyAggBy_(cyCR,pyCR,'concept'),
    };
  });
  // ── Partner Bridge: top gainers & losers YoY ──────────────────────
  var _pCY = {}, _pPY = {};
  cyRows.forEach(function(r) {
    if (!r.partner || !r.partner.trim()) return;
    _pCY[r.partner] = (_pCY[r.partner] || 0) + r.arr;
  });
  pyRows.forEach(function(r) {
    if (!r.partner || !r.partner.trim()) return;
    _pPY[r.partner] = (_pPY[r.partner] || 0) + r.arr;
  });
  var _allKeys = {};
  Object.keys(_pCY).forEach(function(k){ _allKeys[k]=true; });
  Object.keys(_pPY).forEach(function(k){ _allKeys[k]=true; });

  var _allP = Object.keys(_allKeys).map(function(pk) {
    var cy = _pCY[pk] || 0, py = _pPY[pk] || 0, d = cy - py;
    var status = d >= 0
      ? (py === 0 ? 'Sin ref' : 'Incrementó')
      : (cy === 0 ? 'Sin inv. en el período' : 'Disminuyó');
    return { label: pk, cyArr: cy, pyArr: py, delta: d, pct: py ? d/py*100 : 0, status: status };
  });

  var _pos = _allP.filter(function(p){ return p.delta > 0; })
               .sort(function(a,b){ return b.delta - a.delta; });
  var _neg = _allP.filter(function(p){ return p.delta < 0; })
               .sort(function(a,b){ return a.delta - b.delta; });

  function _othRow(arr, lbl) {
    if (!arr.length) return null;
    var cy=arr.reduce(function(s,p){return s+p.cyArr;},0);
    var py=arr.reduce(function(s,p){return s+p.pyArr;},0);
    var d=cy-py;
    return { label:lbl, cyArr:cy, pyArr:py, delta:d, pct:py?d/py*100:0, status:'' };
  }

  var posTop10  = _pos.slice(0,10), posRest = _pos.slice(10);
  var negTop10  = _neg.slice(0,10), negRest = _neg.slice(10);
  var posOth    = _othRow(posRest, 'Others +');
  var negOth    = _othRow(negRest, 'Others −');
  var posList   = posTop10.concat(posOth ? [posOth] : []);
  var negList   = negTop10.concat(negOth ? [negOth] : []);

  var partnerBridge = {
    positive       : posList,
    negative       : negList,
    positiveOthers : posRest,
    negativeOthers : negRest,
    totalPositive  : _pos.reduce(function(s,p){return s+p.delta;},0),
    totalNegative  : _neg.reduce(function(s,p){return s+p.delta;},0),
    totalDelta     : _allP.reduce(function(s,p){return s+p.delta;},0),
  };

  return{
    cyYear:cyYear, pyYear:pyYear, month:mo, hasPY:pyRows.length>0,
    byCountry:yoyAggBy_(cyRows,pyRows,'country'),
    byCluster:yoyAggBy_(cyRows,pyRows,'cluster'),
    byConcept:yoyAggBy_(cyRows,pyRows,'concept'),
    byCountry_detail:byCountry_detail,
    partnerBridge:partnerBridge,
  };
}


// ─────────────────────────────────────────────────────────────────────
// FY HELPERS  (FY convention: FY26 = April 2025 – March 2026)
// April–Dec of year Y → FY(Y+1)   |   Jan–Mar of year Y → FY(Y)
// ─────────────────────────────────────────────────────────────────────
function getFYInfo_(month, year) {
  var MO = ['January','February','March','April','May','June',
            'July','August','September','October','November','December'];
  var mi = MO.indexOf(String(month || '').trim());
  var yr = parseInt(year);
  if (mi < 0 || isNaN(yr)) return { fy: yr||0, q:1, h:1, fyLabel:'FY?' };
  var fy    = mi >= 3 ? yr + 1 : yr;
  var qInFY = Math.floor(((mi - 3 + 12) % 12) / 3) + 1;
  return { fy: fy, q: qInFY, h: qInFY <= 2 ? 1 : 2, fyLabel: 'FY' + String(fy).slice(-2) };
}

function getQPeriods_(rows) {
  var set = {};
  rows.forEach(function(r) {
    if (!r.month || !r.year) return;
    var fi = getFYInfo_(r.month, r.year);
    var k  = fi.fy + '__Q' + fi.q;
    if (!set[k]) set[k] = { fy: fi.fy, q: fi.q };
  });
  return Object.keys(set)
    .map(function(k){ return set[k]; })
    .sort(function(a,b){ return a.fy!==b.fy ? a.fy-b.fy : a.q-b.q; })
    .map(function(p){
      return { month:'Q'+p.q, year:String(p.fy),
               label:'Q'+p.q+' FY'+String(p.fy).slice(-2) };
    });
}

function getHPeriods_(rows) {
  var set = {};
  rows.forEach(function(r) {
    if (!r.month || !r.year) return;
    var fi = getFYInfo_(r.month, r.year);
    var k  = fi.fy + '__H' + fi.h;
    if (!set[k]) set[k] = { fy: fi.fy, h: fi.h };
  });
  return Object.keys(set)
    .map(function(k){ return set[k]; })
    .sort(function(a,b){ return a.fy!==b.fy ? a.fy-b.fy : a.h-b.h; })
    .map(function(p){
      return { month:'H'+p.h, year:String(p.fy),
               label:'H'+p.h+' FY'+String(p.fy).slice(-2) };
    });
}
