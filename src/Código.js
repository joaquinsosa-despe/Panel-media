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
  // Helper para buscar header con fallback (compat con datasets viejos)
  function findCol(...candidates){
    for(const c of candidates){
      const i = headers.indexOf(c);
      if(i !== -1) return i;
    }
    return -1;
  }
  const idx={
    year   : findCol("Year","Date - Year","A\u00f1o","Ano"),
    month  : findCol("Month","Date - Month","Mes"),
    country: findCol("Pais","Pa\u00eds","Country"),
    channel: findCol("Lob-Channel","Lob Channel","Canal"),
    concept: findCol("Concepto","Concept"),
    partner: findCol("Partner"),
    cluster: findCol("Cluster"),
    producto: findCol("Producto_agrupado","Producto agrupado","Producto"),
    actuals: findCol("Actuals + RR","Actuals+RR"),
    budget : findCol("Budget","Budget + Forecast + Breakthrough FY26","Budget + Forecast FY26","Budget + Forecast","Budget+Forecast"),
    rr     : findCol("RR","Rolling Reforecast","Reforecast")
  };
  const REQUIRED={"Year":idx.year,"Month":idx.month,"Pais":idx.country,"Lob-Channel":idx.channel,"Concepto":idx.concept,"Cluster":idx.cluster,"Actuals + RR":idx.actuals};
  const missing=Object.entries(REQUIRED).filter(([,v])=>v===-1).map(([name])=>name);
  if(missing.length>0){SpreadsheetApp.getUi().alert("\u26a0\ufe0f Columnas no encontradas:\n"+missing.join("\n")+"\n\nVerific\u00e1 los nombres.");return[];}
  return raw.slice(1).map(r=>{
    const ch=String(r[idx.channel]||"").trim(),mo=String(r[idx.month]||"").trim(),yr=String(r[idx.year]||"").trim();
    const fyInf=getFYInfo_(mo,yr);
    const budget=idx.budget!==-1?parseNum_(r[idx.budget]):0;
    return{
      year:yr,
      quarter:"Q"+fyInf.q,
      month:mo,
      country:String(r[idx.country]||"").trim(),
      channel:ch,
      lob:LOB_MAP[ch]||"S/C",
      concept:String(r[idx.concept]||"").trim(),
      cluster:String(r[idx.cluster]||"").trim(),
      partner:idx.partner!==-1?String(r[idx.partner]||"").trim():"",
      producto:idx.producto!==-1?String(r[idx.producto]||"").trim():"",
      actuals:parseNum_(r[idx.actuals]),
      budget,
      rr:idx.rr!==-1?parseNum_(r[idx.rr]):0
    };
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


// ─────────────────────────────────────────────────────────────
//  ROUTER (Hub + Dashboards)
// ─────────────────────────────────────────────────────────────
// Cada área tiene su propia hoja + (eventualmente) su propio Apps Script.
// 'media' está integrado en este mismo Apps Script. Para B2B y Riesgos, una
// vez deployados, completar `url` con la URL pública del web app correspondiente.
//
// Convención: si url está vacío => el botón muestra "Próximamente" y no navega.

// Sr Manager que supervisa las 3 áreas. El equipo de cada área reporta directo
// (excepto B2B/B2B2C que tiene un Manager intermedio).
var SR_MANAGER = { name: 'Matias M. Sanchez Berardo', role: 'Sr Manager' };

var AREA_CONFIG = {
  media: {
    name: 'Media Sales',
    tagline: 'Revenue · Forecast · Pareto · Comentarios',
    icon: 'i-trend',
    url: '?view=media',  // mismo Apps Script
    team: [
      { name: 'Joaquín Sosa Beláustegui' },
      { name: 'Nicolas R. Occhipinti'    }
    ]
  },
  b2b: {
    name: 'B2B & B2B2C',
    tagline: 'Wholesale & partnership channels',
    icon: 'i-users',
    url: '',  // TODO: pegar URL del web app cuando esté deployado
    team: [
      { name: 'Diego Bracco',           role: 'Manager' },
      { name: 'Gregorio J. Minetti'    },
      { name: 'Thiago Harari'          },
      { name: 'Delfina Santos'         },
      { name: 'Antonella A. Di Franco' }
    ]
  },
  risk: {
    name: 'Riesgos',
    tagline: 'Risk management & analytics',
    icon: 'i-shield',
    url: '',  // TODO: pegar URL del web app cuando esté deployado
    team: [
      { name: 'Francisco Maldonado'  },
      { name: 'Amos Denenberg Korob' }
    ]
  }
};

function doGet(e) {
  var view = (e && e.parameter && e.parameter.view) || '';
  var selfUrl = '';
  try{ selfUrl = ScriptApp.getService().getUrl() || ''; }catch(err){ selfUrl = ''; }

  if(view === 'media' || view === 'dashboard'){
    var dashTpl = HtmlService.createTemplateFromFile('Index');
    dashTpl.WEB_APP_URL = selfUrl;
    return dashTpl.evaluate()
      .setTitle('Media Sales Despegar — Dashboard')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }
  // Default: Hub
  var tpl = HtmlService.createTemplateFromFile('Hub');
  tpl.AREA_CONFIG_JSON  = JSON.stringify(AREA_CONFIG);
  tpl.SR_MANAGER_JSON   = JSON.stringify(SR_MANAGER);
  tpl.HUB_CATALOG_JSON  = JSON.stringify(getHubCatalog());
  // URL completa del web app (necesaria para navegar afuera del iframe de Apps Script)
  tpl.WEB_APP_URL = selfUrl;
  return tpl.evaluate()
    .setTitle('Despegar — Hub Sr Manager')
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
function getDashboardData(targetMonth, targetYear, groupBy, filterPartners, filterClusters, filterPos, filterLobs, filterProductos) {
  try {
    groupBy         = groupBy         || 'M';
    filterPartners  = (Array.isArray(filterPartners)  ? filterPartners  : []).map(function(s){return String(s).trim();});
    filterClusters  = (Array.isArray(filterClusters)  ? filterClusters  : []).map(function(s){return String(s).trim();});
    filterPos       = (Array.isArray(filterPos)       ? filterPos       : []).map(function(s){return String(s).trim();});
    filterLobs      = (Array.isArray(filterLobs)      ? filterLobs      : []).map(function(s){return String(s).trim();});
    filterProductos = (Array.isArray(filterProductos) ? filterProductos : []).map(function(s){return String(s).trim();});

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
      year    : col_(headers, ['Date - Year',    'Year',    'Año', 'Ano']),
      quarter : col_(headers, ['Date - Quarter',  'Quarter', 'Trimestre']),
      month   : col_(headers, ['Date - Month',    'Month',   'Mes']),
      country : col_(headers, ['País', 'Pais', 'Country']),
      channel : col_(headers, ['Lob-Channel', 'Lob Channel', 'Canal']),
      concept : col_(headers, ['Concepto', 'Concept']),
      partner : col_(headers, ['Partner']),
      cluster : col_(headers, ['Cluster']),
      producto: col_(headers, ['Producto_agrupado', 'Producto agrupado', 'Producto']),
      arr     : col_(headers, ['Actuals + RR', 'Actuals+RR']),
      // Nueva columna única "Budget" (reemplaza B+F FY26, B+F+B FY26 y Budget FY27)
      budget  : col_(headers, ['Budget','Budget + Forecast + Breakthrough FY26','Budget + Forecast FY26','Budget + Forecast','Budget+Forecast','Budget FY27']),
      rr      : col_(headers, ['RR', 'Rolling Reforecast', 'Reforecast'])
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
      var budgetV = idx.budget >= 0 ? parseWA_(r[idx.budget]) : 0;
      var rrV     = idx.rr     >= 0 ? parseWA_(r[idx.rr])     : 0;
      return {
        year    : yr,
        quarter : String(r[idx.quarter] || '').trim(),
        month   : mo,
        country : String(r[idx.country] || '').trim(),
        channel : String(r[idx.channel] || '').trim(),
        concept : String(r[idx.concept] || '').trim(),
        partner : idx.partner  >= 0 ? String(r[idx.partner]  || '').trim() : '',
        cluster : String(r[idx.cluster] || '').trim(),
        producto: idx.producto >= 0 ? String(r[idx.producto] || '').trim() : '',
        arr     : parseWA_(r[idx.arr]),
        // Compat: el dashboard sigue leyendo bf/bfb. Apuntan al Budget único.
        bf      : budgetV,
        bfb     : budgetV,
        budget  : budgetV,
        rr      : rrV,
        fy27    : fy7
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
      // Default = mes-año actual si está en el dataset; si no, el último disponible.
      var _today = new Date();
      var _todayMonth = MONTH_ORD[_today.getMonth()];
      var _todayYear  = String(_today.getFullYear());
      var _found = null;
      if (groupBy === 'M') {
        for (var _pi = 0; _pi < periods.length; _pi++) {
          if (periods[_pi].month === _todayMonth && periods[_pi].year === _todayYear) {
            _found = periods[_pi];
            break;
          }
        }
      }
      if (_found) {
        targetMonth = _found.month;
        targetYear  = _found.year;
      } else {
        var last = periods[periods.length - 1];
        targetMonth = last.month;
        targetYear  = last.year;
      }
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
    var partnerSet = {}, clusterSet = {}, posSet = {}, lobSet = {}, productoSet = {};
    basePeriodRows.forEach(function(r) {
      if (r.partner  && r.partner.trim())  partnerSet[r.partner.trim()]   = true;
      if (r.cluster  && r.cluster.trim())  clusterSet[r.cluster.trim()]   = true;
      if (r.country  && r.country.trim())  posSet[r.country.trim()]       = true;
      // LOB derivado del channel (B2C-SITE, B2C-APP, ... → B2C, etc.)
      var lob = LOB_MAP[r.channel] || (r.channel ? 'S/C' : '');
      if (lob) lobSet[lob] = true;
      if (r.producto && r.producto.trim()) productoSet[r.producto.trim()] = true;
    });
    var partnerList   = Object.keys(partnerSet).sort();
    var clusterList   = Object.keys(clusterSet).sort();
    var posList       = Object.keys(posSet).sort();
    var lobList       = Object.keys(lobSet).sort();
    var productoList  = Object.keys(productoSet).sort();

    // Aplicar filtros (todos en AND)
    var filtered = basePeriodRows;
    if (filterPartners.length > 0) {
      filtered = filtered.filter(function(r) { return filterPartners.indexOf(r.partner) !== -1; });
    }
    if (filterClusters.length > 0) {
      filtered = filtered.filter(function(r) { return filterClusters.indexOf(r.cluster) !== -1; });
    }
    if (filterPos.length > 0) {
      filtered = filtered.filter(function(r) { return filterPos.indexOf(r.country) !== -1; });
    }
    if (filterLobs.length > 0) {
      filtered = filtered.filter(function(r) {
        var lob = LOB_MAP[r.channel] || (r.channel ? 'S/C' : '');
        return filterLobs.indexOf(lob) !== -1;
      });
    }
    if (filterProductos.length > 0) {
      filtered = filtered.filter(function(r) { return filterProductos.indexOf(r.producto) !== -1; });
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
        producto: r.producto || '',
        arr: r.arr, bf: r.bf, bfb: r.bfb,
        budget: r.budget != null ? r.budget : r.bf,
        rr: r.rr || 0,
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
    var totRR  = rowsGB.reduce(function(s,r){ return s+(r.rr||0); }, 0);
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
      bfLabel        : 'Budget',
      bfbLabel       : 'Budget',
      countries      : countries,
      periods        : periods,
      hasGB          : hasGB,
      partnerList     : partnerList,
      clusterList     : clusterList,
      posList         : posList,
      lobList         : lobList,
      productoList    : productoList,
      filterPartners  : filterPartners,
      filterClusters  : filterClusters,
      filterPos       : filterPos,
      filterLobs      : filterLobs,
      filterProductos : filterProductos,
      totals: {
        arr: totArr, bf: totBf, bfb: totBfb,
        budget: totBf, rr: totRR,
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
    var rr  = grp.reduce(function(s,r){ return s+(r.rr||0); }, 0);
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
      arr: arr, bf: bf, bfb: bfb, budget: bf, rr: rr,
      dBf : arr - bf,  dBfb: arr - bfb,
      pctBf : bf  ? (arr - bf)  / bf  * 100 : 0,
      pctBfb: bfb ? (arr - bfb) / bfb * 100 : 0,
      // Variance vs RR (forecast original): mide cuánto le pegamos al RR
      dRR  : arr - rr,
      pctRR: rr ? (arr - rr) / rr * 100 : 0,
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
  var tRR  = result.reduce(function(s,r){ return s+(r.rr||0); }, 0);
  var allC = Object.keys(rows.reduce(function(acc,r){ acc[r.country]=1; return acc; }, {}));
  var tGB  = sumCountryGBs_(allC, gbByCountry);


  result.push({
    label: 'TOTAL', isTotal: true,
    arr: tArr, bf: tBf, bfb: tBfb, budget: tBf, rr: tRR,
    dBf : tArr - tBf,  dBfb: tArr - tBfb,
    pctBf : tBf  ? (tArr - tBf)  / tBf  * 100 : 0,
    pctBfb: tBfb ? (tArr - tBfb) / tBfb * 100 : 0,
    dRR  : tArr - tRR,
    pctRR: tRR ? (tArr - tRR) / tRR * 100 : 0,
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

/**
 * ─────────────────────────────────────────────────────────────
 *  Media Sales Dashboard — Code.gs (ampliado v2)
 * ─────────────────────────────────────────────────────────────
 *  Este archivo contiene SOLO las funciones nuevas que pidió el frontend v2.
 *  Pegar al final de tu Code.gs existente. Las funciones existentes
 *  (doGet, getDashboardData, etc.) se mantienen tal cual.
 *
 *  Hojas requeridas en el Sheet
 *    SPREADSHEET_ID: 1CrXabTy4dP6mgD0Rbd2kxotThdtVyMXXnVS_cI44xaA
 *
 *    - "comentarios"   (existe; se inicializa con encabezados si está vacía)
 *    - "suscripciones" (se crea si no existe)
 *
 *  Estructura de la hoja "comentarios"
 *    A: timestamp        Date
 *    B: author           string (email)
 *    C: targetType       'row' | 'cell' | 'general'
 *    D: targetKey        ej. "country:Argentina" | "cluster:MR" | "partner:Booking.com"
 *                            | "country:Argentina|cluster:MR" (combinaciones)
 *    E: period           "Oct 2025" (o vacío para nota permanente)
 *    F: text             contenido (texto libre)
 *    G: pinned           TRUE/FALSE  → pinned se muestra siempre arriba del row
 *    H: resolved         TRUE/FALSE  → resolved oculta de la vista por defecto
 *    I: parentId         string      → para hilos: el row(_RID) padre
 *    J: id               string      → uuid (lo genera la función)
 *
 *  Estructura de la hoja "suscripciones"
 *    A: timestamp | B: email | C: schedule (weekly|daily|monthly)
 *    D: day (1=Lun..7=Dom, sólo weekly) | E: hour (0-23) | F: active (TRUE/FALSE)
 *    G: filters (JSON string)
 * ─────────────────────────────────────────────────────────────
 */

var SPREADSHEET_ID_V2 = '1CrXabTy4dP6mgD0Rbd2kxotThdtVyMXXnVS_cI44xaA';
var SHEET_COMMENTS    = 'comentarios';
var SHEET_SUBS        = 'suscripciones';
var SHEET_DIRECTORY   = 'Directory';  // emails @despegar.com para @menciones
var SHEET_HUB_CATALOG = 'Hub_Catalog'; // catálogo de secciones + dashboards del portal

// Lista semilla de admins (bootstrap). En runtime se persiste en ScriptProperties
// (clave ADMIN_EMAILS_JSON) y se administra desde el frontend.
var ADMIN_EMAILS = [
  'joaquin.sosa@despegar.com'
];
var ADMIN_PROPS_KEY = 'ADMIN_EMAILS_JSON';

// ── helpers ──────────────────────────────────────────────────
function _ss_(){ return SpreadsheetApp.openById(SPREADSHEET_ID_V2); }
function _userEmail_(){
  try{ return Session.getActiveUser().getEmail() || Session.getEffectiveUser().getEmail() || 'anon'; }
  catch(e){ return 'anon'; }
}
function _uuid_(){ return Utilities.getUuid(); }

function _getAdminList_(){
  try{
    var raw = PropertiesService.getScriptProperties().getProperty(ADMIN_PROPS_KEY);
    if(raw){
      var parsed = JSON.parse(raw);
      if(Array.isArray(parsed) && parsed.length){
        return parsed.map(function(e){ return String(e).trim().toLowerCase(); })
                     .filter(function(e){ return e; });
      }
    }
  }catch(e){ /* fallthrough a la semilla */ }
  // Semilla: persistir la lista hardcodeada en la primera ejecución
  var seed = ADMIN_EMAILS.map(function(e){ return String(e).trim().toLowerCase(); })
                        .filter(function(e){ return e; });
  try{ PropertiesService.getScriptProperties().setProperty(ADMIN_PROPS_KEY, JSON.stringify(seed)); }catch(e){}
  return seed;
}
function _setAdminList_(arr){
  var clean = (arr||[]).map(function(e){ return String(e).trim().toLowerCase(); })
                      .filter(function(e){ return e; });
  // dedupe
  var seen = {}, out = [];
  clean.forEach(function(e){ if(!seen[e]){ seen[e]=1; out.push(e); } });
  PropertiesService.getScriptProperties().setProperty(ADMIN_PROPS_KEY, JSON.stringify(out));
  return out;
}
function _isAdmin_(email){
  if(!email) return false;
  var e = String(email).toLowerCase();
  return _getAdminList_().indexOf(e) !== -1;
}
function _requireAdmin_(){
  var email = _userEmail_();
  if(!_isAdmin_(email)) throw new Error('No autorizado: se requiere ser admin');
  return email;
}
function _validEmail_(s){
  return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(String(s||'').trim());
}

/**
 * Devuelve info del usuario actual para el frontend.
 * Llamado desde el HTML: google.script.run.getCurrentUserInfo()
 */
function getCurrentUserInfo(){
  var email = _userEmail_();
  return { email: email, isAdmin: _isAdmin_(email) };
}

/** Lista de admins. Sólo admins pueden consultarla. */
function listAdmins(){
  _requireAdmin_();
  return { ok:true, admins: _getAdminList_() };
}

/** Agrega un nuevo admin por email. Sólo admins. */
function addAdmin(email){
  var caller = _requireAdmin_();
  var e = String(email||'').trim().toLowerCase();
  if(!_validEmail_(e)) return { ok:false, reason:'Email inválido' };
  var current = _getAdminList_();
  if(current.indexOf(e) !== -1) return { ok:false, reason:'Ese email ya es admin' };
  current.push(e);
  var saved = _setAdminList_(current);
  return { ok:true, admins: saved, addedBy: caller };
}

/** Quita un admin por email. Sólo admins. No permite quitar al último admin. */
function removeAdmin(email){
  var caller = _requireAdmin_();
  var e = String(email||'').trim().toLowerCase();
  var current = _getAdminList_();
  var idx = current.indexOf(e);
  if(idx === -1) return { ok:false, reason:'Ese email no es admin' };
  if(current.length <= 1) return { ok:false, reason:'No se puede quitar al último admin' };
  current.splice(idx, 1);
  var saved = _setAdminList_(current);
  return { ok:true, admins: saved, removedBy: caller };
}

function _ensureCommentsSheet_(){
  var ss = _ss_();
  var sh = ss.getSheetByName(SHEET_COMMENTS);
  if(!sh){ sh = ss.insertSheet(SHEET_COMMENTS); }
  // Headers (only set if first row is empty)
  if(sh.getLastRow() === 0){
    sh.getRange(1,1,1,11).setValues([[
      'timestamp','author','targetType','targetKey','period','text','pinned','resolved','parentId','id','filters'
    ]]);
    sh.getRange(1,1,1,11).setFontWeight('bold').setBackground('#1a1d23').setFontColor('#fff');
    sh.setFrozenRows(1);
  } else if(sh.getLastColumn() < 11){
    // Migración: agregar columna 'filters' a sheets viejas (comentarios pre-existentes quedan sin filters)
    sh.getRange(1, 11).setValue('filters')
      .setFontWeight('bold').setBackground('#1a1d23').setFontColor('#fff');
  }
  // CRÍTICO: la columna period (E) debe ser texto plano para que Sheets no
  // auto-convierta "May 2026" → Date object. Idempotente: setear cada vez es seguro.
  try { sh.getRange('E:E').setNumberFormat('@'); } catch(_){}
  return sh;
}

// Convierte el valor de la columna 'period' a string limpio. Si Sheets lo guardó
// como Date object (auto-conversión vieja), reconstruye "Month YYYY".
function _normalizePeriod_(val){
  if(val == null) return '';
  if(Object.prototype.toString.call(val) === '[object Date]' && !isNaN(val.getTime())){
    var months = ['January','February','March','April','May','June',
                  'July','August','September','October','November','December'];
    return months[val.getMonth()] + ' ' + val.getFullYear();
  }
  return String(val);
}

function _ensureSubsSheet_(){
  var ss = _ss_();
  var sh = ss.getSheetByName(SHEET_SUBS);
  if(!sh){
    sh = ss.insertSheet(SHEET_SUBS);
    sh.getRange(1,1,1,7).setValues([[
      'timestamp','email','schedule','day','hour','active','filters'
    ]]);
    sh.getRange(1,1,1,7).setFontWeight('bold').setBackground('#1a1d23').setFontColor('#fff');
    sh.setFrozenRows(1);
  }
  return sh;
}

// ═════════════════════════════════════════════════════════════
//  FEATURE 3 + 10 — Anotaciones / Comentarios por celda o fila
// ═════════════════════════════════════════════════════════════

/**
 * Guarda un nuevo comentario / anotación / respuesta.
 * Llamado desde el HTML: google.script.run.saveComment({...})
 *
 * payload = {
 *   targetType: 'row' | 'cell' | 'general',
 *   targetKey:  string  (ej "country:Argentina|cluster:MR"),
 *   period:     string  ("Oct 2025") o '',
 *   text:       string,
 *   pinned:     boolean,
 *   parentId:   string | null  (responde a un hilo existente)
 * }
 */
function saveComment(payload){
  if(!payload || !payload.text){ throw new Error('Texto vacío'); }
  var sh = _ensureCommentsSheet_();
  var id = _uuid_();
  // filters: objeto serializado en JSON. null/undefined → '' (comentarios sin contexto de filtros).
  var filtersStr = '';
  if(payload.filters && typeof payload.filters === 'object'){
    try { filtersStr = JSON.stringify(payload.filters).slice(0, 4000); }
    catch(e){ filtersStr = ''; }
  }
  var author = _userEmail_();
  var periodStr = String(payload.period||'');
  sh.appendRow([
    new Date(),
    author,
    String(payload.targetType||'row'),
    String(payload.targetKey||''),
    periodStr,
    String(payload.text).slice(0,2000),
    !!payload.pinned,
    false,
    String(payload.parentId||''),
    id,
    filtersStr
  ]);
  // DEFENSIVO: Sheets puede auto-convertir "May 2026" → Date al hacer appendRow
  // aunque la columna esté en formato '@'. Forzamos la celda de period (col 5)
  // a texto plano y reescribimos el valor explícitamente como string.
  try {
    var newRow = sh.getLastRow();
    sh.getRange(newRow, 5).setNumberFormat('@').setValue(periodStr);
  } catch(_){}

  // Resolver destinatarios y mandar mail (best-effort, no rompe el save).
  try {
    var directory = getDirectory();
    var mentioned = _parseMentions_(payload.text, directory);
    var recipients = {};
    mentioned.forEach(function(e){ recipients[e] = true; });
    // Si es un reply, sumar autores del thread (mismo targetKey)
    if(payload.parentId){
      _threadAuthors_(String(payload.targetKey||''), String(payload.parentId)).forEach(function(e){
        recipients[e] = true;
      });
    }
    // No mandar mail al propio autor
    delete recipients[String(author||'').toLowerCase()];
    var list = Object.keys(recipients);
    if(list.length > 0){
      _sendCommentMail_(list, payload, id);
    }
  } catch(e){
    try { console.error('[saveComment mail]', e && e.message || e); } catch(_){}
  }

  return { ok:true, id:id };
}

/**
 * Devuelve TODOS los comentarios indexados por targetKey.
 * El frontend recibe un mapa { "country:Argentina": [c1,c2,...], ... }
 * para hacer un join en cada row sin disparar N requests.
 */
function getCommentsIndex(){
  var sh = _ensureCommentsSheet_();
  var last = sh.getLastRow();
  if(last < 2) return { byKey:{}, count:0 };
  var lastCol = Math.max(10, sh.getLastColumn()); // soporta sheets con o sin la col 11 (filters)
  var data = sh.getRange(2,1,last-1,lastCol).getValues();
  var byKey = {};
  var count = 0;
  data.forEach(function(r){
    var resolved = r[7] === true || r[7] === 'TRUE';
    if(resolved) return;
    var key = String(r[3]||'');
    if(!byKey[key]) byKey[key] = [];
    // filters: parseado de la columna 11 si existe
    var filters = null;
    var rawFilters = r.length > 10 ? r[10] : '';
    if(rawFilters){
      try { filters = JSON.parse(String(rawFilters)); }
      catch(e){ filters = null; }
    }
    byKey[key].push({
      id: String(r[9]||''),
      timestamp: r[0] ? r[0].toISOString() : '',
      author: String(r[1]||''),
      targetType: String(r[2]||''),
      period: _normalizePeriod_(r[4]),
      text: String(r[5]||''),
      pinned: r[6] === true || r[6] === 'TRUE',
      parentId: String(r[8]||''),
      filters: filters
    });
    count++;
  });
  // sort: pinned first, then chronological
  Object.keys(byKey).forEach(function(k){
    byKey[k].sort(function(a,b){
      if(a.pinned && !b.pinned) return -1;
      if(!a.pinned && b.pinned) return 1;
      return a.timestamp.localeCompare(b.timestamp);
    });
  });
  return { byKey:byKey, count:count };
}

/** Marca como resuelto un comentario (no lo borra; queda historial). Sólo autor o admin. */
function resolveComment(id){
  var email = _userEmail_();
  var sh = _ensureCommentsSheet_();
  var last = sh.getLastRow();
  if(last < 2) return { ok:false };
  var data = sh.getRange(2,1,last-1,10).getValues();
  for(var i=0;i<data.length;i++){
    if(String(data[i][9]) === String(id)){
      if(String(data[i][1]) !== email && !_isAdmin_(email)){
        return { ok:false, reason:'Sólo el autor o un admin puede resolver' };
      }
      sh.getRange(i+2, 8).setValue(true);   // resolved column
      return { ok:true };
    }
  }
  return { ok:false };
}

/** Borra un comentario (autor propio o admin). */
function deleteComment(id){
  var email = _userEmail_();
  var sh = _ensureCommentsSheet_();
  var last = sh.getLastRow();
  if(last < 2) return { ok:false };
  var data = sh.getRange(2,1,last-1,10).getValues();
  for(var i=0;i<data.length;i++){
    if(String(data[i][9]) === String(id)){
      if(String(data[i][1]) !== email && !_isAdmin_(email)){
        return { ok:false, reason:'Sólo el autor o un admin puede borrar' };
      }
      sh.deleteRow(i+2);
      return { ok:true };
    }
  }
  return { ok:false };
}

// ═════════════════════════════════════════════════════════════
//  FEATURE 11 — Suscripciones por email
// ═════════════════════════════════════════════════════════════

/**
 * Da de alta una suscripción. Frontend:
 *   google.script.run.subscribeEmail({
 *     email:'x@despegar.com', schedule:'weekly', day:1, hour:9, filters:{...}
 *   })
 */
function subscribeEmail(payload){
  if(!payload || !payload.email) throw new Error('Email requerido');
  var sh = _ensureSubsSheet_();
  sh.appendRow([
    new Date(),
    String(payload.email).trim().toLowerCase(),
    String(payload.schedule||'weekly'),
    Number(payload.day||1),
    Number(payload.hour||9),
    true,
    JSON.stringify(payload.filters||{})
  ]);
  _ensureWeeklyTrigger_();
  return { ok:true };
}

function listSubscriptions(){
  var email = _userEmail_();
  var sh = _ensureSubsSheet_();
  var last = sh.getLastRow();
  if(last < 2) return [];
  var data = sh.getRange(2,1,last-1,7).getValues();
  return data.filter(function(r){return String(r[1]).toLowerCase() === email.toLowerCase();})
             .map(function(r){return {
               timestamp:r[0]?r[0].toISOString():'', email:r[1], schedule:r[2],
               day:r[3], hour:r[4], active:r[5], filters:r[6]
             };});
}

function unsubscribeEmail(email){
  var sh = _ensureSubsSheet_();
  var last = sh.getLastRow();
  if(last < 2) return { ok:false };
  var data = sh.getRange(2,2,last-1,1).getValues();
  for(var i=0;i<data.length;i++){
    if(String(data[i][0]).toLowerCase() === String(email).toLowerCase()){
      sh.getRange(i+2, 6).setValue(false);
    }
  }
  return { ok:true };
}

/**
 * Trigger time-driven: corre todas las mañanas y manda mail
 * a las suscripciones que matchean el día/hora actual.
 * IMPORTANTE: configurar un trigger horario (cada 1h) sobre esta función.
 */
function _ensureWeeklyTrigger_(){
  var triggers = ScriptApp.getProjectTriggers();
  var has = triggers.some(function(t){ return t.getHandlerFunction() === 'sendScheduledDigests'; });
  if(!has){
    ScriptApp.newTrigger('sendScheduledDigests')
      .timeBased().everyHours(1).create();
  }
}

function sendScheduledDigests(){
  var now = new Date();
  var hour = now.getHours();
  var dow = now.getDay() === 0 ? 7 : now.getDay();   // 1..7 (lun..dom)
  var sh = _ensureSubsSheet_();
  var last = sh.getLastRow();
  if(last < 2) return;
  var data = sh.getRange(2,1,last-1,7).getValues();
  data.forEach(function(r){
    var schedule=r[2], day=Number(r[3]||1), h=Number(r[4]||9), active=r[5]===true||r[5]==='TRUE';
    if(!active) return;
    var due=false;
    if(schedule==='daily'   && h===hour)                due = true;
    if(schedule==='weekly'  && h===hour && day===dow)   due = true;
    if(schedule==='monthly' && h===hour && now.getDate()===1) due = true;
    if(!due) return;
    try{
      _sendDigestEmail_(String(r[1]), JSON.parse(r[6]||'{}'));
    }catch(e){
      Logger.log('Falló envío a '+r[1]+': '+e);
    }
  });
}

function _sendDigestEmail_(email, filters){
  // Toma el último mes cerrado por defecto. Reutiliza tu pipeline existente.
  var d = getDashboardData('', '', 'M', filters.partners||[], filters.clusters||[]);
  if(d.error) return;
  var t = d.totals;
  var dBf = (t.arr||0) - (t.bf||0);
  var pBf = t.bf ? (dBf/t.bf*100) : 0;
  var fmtM = function(n){ return '$'+((n||0)/1e6).toFixed(2)+'M'; };
  var url = 'https://script.google.com/macros/s/__YOUR_DEPLOYMENT_ID__/exec';
  var html = '<div style="font-family:Inter,Arial,sans-serif;background:#0B0D10;color:#E6E8EB;padding:24px;border-radius:12px">'
    + '<h2 style="margin:0 0 8px;color:#fff">Media Sales — '+d.periodLabel+'</h2>'
    + '<p style="color:#8B92A0;margin:0 0 18px">Resumen automático</p>'
    + '<table style="border-collapse:collapse;width:100%;margin-bottom:18px">'
    +   '<tr>'
    +     '<td style="padding:14px;background:#15181D;border:1px solid #232830;border-radius:8px;width:33%">'
    +       '<div style="color:#8B92A0;font-size:11px;text-transform:uppercase">Actuals + RR</div>'
    +       '<div style="font-size:22px;font-weight:600;color:#fff">'+fmtM(t.arr)+'</div>'
    +     '</td>'
    +     '<td style="padding:14px;background:#15181D;border:1px solid #232830;border-radius:8px;width:33%">'
    +       '<div style="color:#8B92A0;font-size:11px;text-transform:uppercase">'+(d.bfLabel||'B+F')+'</div>'
    +       '<div style="font-size:22px;font-weight:600;color:#fff">'+fmtM(t.bf)+'</div>'
    +     '</td>'
    +     '<td style="padding:14px;background:#15181D;border:1px solid #232830;border-radius:8px;width:33%">'
    +       '<div style="color:#8B92A0;font-size:11px;text-transform:uppercase">Δ vs '+(d.bfLabel||'B+F')+'</div>'
    +       '<div style="font-size:22px;font-weight:600;color:'+(dBf>=0?'#3ECF8E':'#F87171')+'">'+(dBf>=0?'+':'')+fmtM(dBf)+' ('+(pBf>=0?'+':'')+Math.round(pBf)+'%)</div>'
    +     '</td>'
    +   '</tr>'
    + '</table>'
    + '<a href="'+url+'" style="display:inline-block;background:#7C7FFF;color:#fff;text-decoration:none;padding:10px 18px;border-radius:8px;font-weight:600">Ver dashboard completo →</a>'
    + '</div>';
  MailApp.sendEmail({
    to: email,
    subject: '[Media Sales] '+d.periodLabel+' — '+(dBf>=0?'+':'')+fmtM(dBf)+' vs '+(d.bfLabel||'B+F'),
    htmlBody: html,
    name: 'Media Sales Dashboard'
  });
}

// ═════════════════════════════════════════════════════════════
//  FEATURE 2 — Sparklines (serie histórica por dimensión)
// ═════════════════════════════════════════════════════════════

/**
 * Devuelve series temporales de los últimos N meses para CADA label
 * de country/cluster/concept/partner. El frontend renderea sparklines inline.
 *
 *   { months:['May 25','Jun 25',...], byCountry:{Argentina:[...]}, byCluster:{...},
 *     byConcept:{...}, byPartner:{...} }
 *
 * IMPLEMENTACIÓN: Esto necesita que tu pipeline actual exponga datos
 * históricos. Lo más simple es llamar a tu función existente
 * `_buildEvolutivoData_(referenceMonth, lookback)` y reagregarlo.
 *
 * Si tu Code.gs ya tiene `getEvolutivoData()` o equivalente, podés reusar
 * ese pivot pero con groupBy por country/cluster/concept en lugar de bucket.
 *
 * Acá dejo una versión que asume que vas a poder llamar a `getDashboardData`
 * mes a mes hacia atrás. Si tu dataset es chico (~6-12 meses x 6 países x 5
 * clusters = pocos miles de filas) es viable; si no, conviene hacer una query
 * más eficiente directo sobre la planilla pivot.
 */
function getSparklines(referenceMonth, referenceYear, lookback){
  lookback = lookback || 6;
  var out = {
    months: [],
    byCountry: {},
    byCluster: {},
    byConcept: {},
    byPartner: {}
  };
  // Construir lista de meses hacia atrás
  var MONTHS_ES = ['Ene','Feb','Mar','Abr','May','Jun','Jul','Ago','Sep','Oct','Nov','Dic'];
  var refIdx = MONTHS_ES.indexOf(referenceMonth);
  if(refIdx < 0) refIdx = new Date().getMonth();
  var refY = Number(referenceYear) || new Date().getFullYear();

  var queue = [];
  for(var i=lookback-1; i>=0; i--){
    var idx = refIdx - i;
    var y = refY;
    while(idx < 0){ idx += 12; y -= 1; }
    queue.push({ month: MONTHS_ES[idx], year: String(y), label: MONTHS_ES[idx]+' '+String(y).slice(2) });
  }
  out.months = queue.map(function(q){ return q.label; });

  // Helper: acumula valores en un map por label
  function pushVal(map, label, val){
    if(!map[label]) map[label] = new Array(queue.length).fill(0);
    return map;
  }

  queue.forEach(function(q, i){
    var d;
    try{ d = getDashboardData(q.month, q.year, 'M', [], []); }
    catch(e){ return; }
    if(!d || d.error) return;

    // Country
    (d.general && d.general.byCountry || []).forEach(function(r){
      if(r.isTotal) return;
      pushVal(out.byCountry, r.label);
      out.byCountry[r.label][i] = r.arr || 0;
    });
    (d.general && d.general.byCluster || []).forEach(function(r){
      if(r.isTotal) return;
      pushVal(out.byCluster, r.label);
      out.byCluster[r.label][i] = r.arr || 0;
    });
    (d.general && d.general.byConcept || []).forEach(function(r){
      if(r.isTotal) return;
      pushVal(out.byConcept, r.label);
      out.byConcept[r.label][i] = r.arr || 0;
    });
    // Partners (top 30 por mes alcanza para sparklines)
    (d.general && d.general.byCountry || []).forEach(function(r){
      (r.partners || []).slice(0, 30).forEach(function(p){
        if(p.label.indexOf('Others')===0) return;
        pushVal(out.byPartner, p.label);
        out.byPartner[p.label][i] += (p.arr || 0);
      });
    });
  });

  return out;
}

// ═════════════════════════════════════════════════════════════
//  FEATURE 6 — Pareto / curva de concentración
// ═════════════════════════════════════════════════════════════

/**
 * Calcula la curva de Pareto para partners de un período.
 *
 * @param {string} referenceMonth  Mes (ej "March")
 * @param {string} referenceYear   Año (ej "2026")
 * @param {string|boolean} mode    'cy' (default) | 'compare' | 'py'
 *                                 - 'cy':      solo período actual
 *                                 - 'compare': CY + PY (bars duales)
 *                                 - 'py':      rankear por PY (ver qué partners bajaron / churnearon)
 *                                 Por compat: true=compare, false/null='cy'
 */
function getPareto(referenceMonth, referenceYear, mode){
  // Normalizar mode (compat con boolean)
  if(mode === true)  mode = 'compare';
  if(!mode || mode === false) mode = 'cy';
  if(['cy','compare','py'].indexOf(mode) === -1) mode = 'cy';

  function aggregatePartners_(month, year){
    var d = getDashboardData(month, year, 'M', [], []);
    if(!d || d.error) return null;
    var agg = {};
    (d.general && d.general.byCountry || []).forEach(function(r){
      (r.partners||[]).forEach(function(p){
        if(p.label.indexOf('Others')===0) return;
        agg[p.label] = (agg[p.label]||0) + (p.arr||0);
      });
    });
    return { agg:agg, periodLabel: d.periodLabel || (month+' '+year), month:d.targetMonth||month, year:d.targetYear||year };
  }

  function buildRankedList_(agg){
    var arr = Object.keys(agg).map(function(k){return {label:k, arr:agg[k]};});
    arr.sort(function(a,b){return b.arr - a.arr;});
    var total = arr.reduce(function(s,r){return s + r.arr;}, 0);
    var cum = 0, minFor80 = null;
    arr = arr.map(function(r,i){
      cum += r.arr;
      var cumPct = total ? cum/total*100 : 0;
      if(minFor80===null && cumPct>=80) minFor80 = i+1;
      return { label:r.label, arr:r.arr, cumArr:cum, cumPct:cumPct, rank:i+1 };
    });
    return {
      partners: arr,
      total: total,
      top5Pct:  arr.slice(0,5).reduce(function(s,r){return s+r.arr;},0)  / (total||1) * 100,
      top20Pct: arr.slice(0,20).reduce(function(s,r){return s+r.arr;},0) / (total||1) * 100,
      minPartnersFor80: minFor80 || arr.length
    };
  }

  var cy = aggregatePartners_(referenceMonth, referenceYear);
  if(!cy) return { partners:[], total:0, mode:mode };

  var pyYear = String(Number(referenceYear) - 1);
  var py = (mode === 'compare' || mode === 'py')
    ? aggregatePartners_(referenceMonth, pyYear)
    : null;

  var cyRanked = buildRankedList_(cy.agg);
  var pyRanked = py ? buildRankedList_(py.agg) : null;

  // Elegir el ranking primario según el modo
  // - 'cy' / 'compare': ranking por CY
  // - 'py':             ranking por PY (para ver quién bajó o churneó)
  var primary = (mode === 'py' && pyRanked) ? pyRanked : cyRanked;
  var secondaryAgg = (mode === 'py') ? cy.agg : (py ? py.agg : null);
  var primaryAgg   = (mode === 'py') ? py.agg : cy.agg;

  // Top 30 del primario
  var top30 = primary.partners.slice(0, 30);

  // Si compare o py: anotar cada partner con su contraparte
  if(mode === 'compare' || mode === 'py'){
    top30 = top30.map(function(p){
      var secVal = (secondaryAgg && secondaryAgg[p.label]) || 0;
      if(mode === 'py'){
        // p.arr = PY value (primario); secVal = CY value
        var delta = secVal - (p.arr||0);                // CY - PY (positivo = subió este año)
        var pct   = (p.arr||0) ? (delta / (p.arr||1) * 100) : (secVal ? 100 : 0);
        return Object.assign({}, p, {
          pyArr: p.arr,    // alias para frontend (val de PY)
          cyArr: secVal,   // valor de CY
          delta: delta,
          pct: pct
        });
      } else {
        // mode 'compare': p.arr = CY (primario); secVal = PY
        var d2 = (p.arr||0) - secVal;
        var pp2 = secVal ? (d2 / secVal * 100) : (p.arr ? 100 : 0);
        return Object.assign({}, p, { pyArr: secVal, delta: d2, pct: pp2 });
      }
    });
  }

  var result = {
    mode: mode,
    partners: top30,
    total: primary.total,
    top5Pct: primary.top5Pct,
    top20Pct: primary.top20Pct,
    minPartnersFor80: primary.minPartnersFor80,
    cyLabel: cy.periodLabel,
    pyLabel: py ? py.periodLabel : (referenceMonth+' '+pyYear),
    primaryLabel: (mode === 'py') ? (py?py.periodLabel:('Año anterior')) : cy.periodLabel
  };

  if(pyRanked){
    result.pyTotal      = pyRanked.total;
    result.pyTop5Pct    = pyRanked.top5Pct;
    result.pyTop20Pct   = pyRanked.top20Pct;
    result.cyTotalForPY = cyRanked.total;
    result.totalDelta   = (mode === 'py')
      ? (cyRanked.total - pyRanked.total)
      : (cyRanked.total - pyRanked.total);
    result.totalPct     = pyRanked.total ? (result.totalDelta / pyRanked.total * 100) : 0;
    result.hasPY        = true;
  } else {
    result.hasPY = false;
  }

  return result;
}

// ═════════════════════════════════════════════════════════════
//  FEATURE 7 — Cohortes de partners (nuevos / existentes / churneados)
// ═════════════════════════════════════════════════════════════

/**
 * Clasifica cada partner del mes actual según su antigüedad:
 *   - 'new'      : primera aparición en últimos 6 meses
 *   - 'existing' : activo hace 12+ meses (continuo)
 *   - 'returning': inactivo 3-12 meses y volvió
 *   - 'churned'  : activo en PY-mismo-mes pero no en CY
 *
 * Devuelve agregados de A+RR por cohorte para entender de dónde viene el crecimiento.
 */
function getPartnerCohorts(referenceMonth, referenceYear){
  var ref = getDashboardData(referenceMonth, referenceYear, 'M', [], []);
  if(!ref || ref.error) return { cohorts:{}, error:ref && ref.error };

  // Construye índice partner→primera aparición + última aparición en lookback de 18 meses
  var spark = getSparklines(referenceMonth, referenceYear, 18);
  var partners = spark.byPartner || {};
  var months = spark.months || [];
  var lastIdx = months.length - 1;     // current month
  var pyIdx   = Math.max(0, lastIdx - 12);

  var cohorts = {
    new:       { count:0, arr:0, partners:[] },
    existing:  { count:0, arr:0, partners:[] },
    returning: { count:0, arr:0, partners:[] },
    churned:   { count:0, arr:0, partners:[] }
  };

  Object.keys(partners).forEach(function(p){
    var series = partners[p] || [];
    var cy = series[lastIdx] || 0;
    var py = series[pyIdx]   || 0;
    var firstActiveIdx = series.findIndex(function(v){return v > 0;});
    var lastActiveBefore = -1;
    for(var i=lastIdx-1; i>=0; i--){ if(series[i]>0){ lastActiveBefore = i; break; } }
    var ageMonths = firstActiveIdx >= 0 ? (lastIdx - firstActiveIdx) : 0;

    var cohort;
    if(cy === 0 && py > 0)                 cohort = 'churned';
    else if(cy > 0 && firstActiveIdx >= lastIdx - 6)  cohort = 'new';
    else if(cy > 0 && lastActiveBefore >= 0 && lastIdx - lastActiveBefore > 3) cohort = 'returning';
    else if(cy > 0)                        cohort = 'existing';
    else return;

    cohorts[cohort].count++;
    cohorts[cohort].arr += (cohort === 'churned' ? py : cy);
    cohorts[cohort].partners.push({ label:p, arr:(cohort==='churned'?py:cy), ageMonths:ageMonths });
  });

  // Sort partner lists by arr desc, top 10
  Object.keys(cohorts).forEach(function(k){
    cohorts[k].partners.sort(function(a,b){return b.arr - a.arr;});
    cohorts[k].partners = cohorts[k].partners.slice(0, 10);
  });

  return { cohorts:cohorts, months:months };
}

// ═════════════════════════════════════════════════════════════
//  FEATURE 13 — Comparación libre entre 2 períodos arbitrarios
// ═════════════════════════════════════════════════════════════

/**
 * Devuelve un dataset con A+RR de dos períodos para comparar.
 * Frontend:
 *   google.script.run.compareTwo({a:{month:'Oct',year:'2025'}, b:{month:'Sep',year:'2025'}})
 *
 * Retorna estructura similar a `data.yoy` para reutilizar el render existente.
 */
function compareTwo(payload){
  var a = payload.a, b = payload.b;
  var dA = getDashboardData(a.month, a.year, 'M', payload.partners||[], payload.clusters||[]);
  var dB = getDashboardData(b.month, b.year, 'M', payload.partners||[], payload.clusters||[]);
  if(dA.error) throw new Error('Período A: '+dA.error);
  if(dB.error) throw new Error('Período B: '+dB.error);

  function joinRows(rowsA, rowsB){
    var mapB = {};
    (rowsB||[]).forEach(function(r){ mapB[r.label] = r; });
    var joined = (rowsA||[]).filter(function(r){return !r.isTotal;}).map(function(rA){
      var rB = mapB[rA.label] || {arr:0};
      return {
        label: rA.label,
        cyArr: rA.arr,
        pyArr: rB.arr,
        delta: (rA.arr||0) - (rB.arr||0),
        pct:   rB.arr ? ((rA.arr - rB.arr)/rB.arr*100) : 0,
        partners: rA.partners || []
      };
    });
    var tot = {cyArr:0,pyArr:0};
    joined.forEach(function(r){tot.cyArr+=r.cyArr; tot.pyArr+=r.pyArr;});
    joined.push({
      label:'TOTAL', cyArr:tot.cyArr, pyArr:tot.pyArr,
      delta: tot.cyArr - tot.pyArr,
      pct: tot.pyArr ? (tot.cyArr-tot.pyArr)/tot.pyArr*100 : 0,
      isTotal:true, partners:[]
    });
    return joined;
  }

  return {
    aLabel: a.month+' '+a.year,
    bLabel: b.month+' '+b.year,
    cyYear: a.year, pyYear: b.year, month: a.month,
    byCountry: joinRows(dA.general.byCountry, dB.general.byCountry),
    byCluster: joinRows(dA.general.byCluster, dB.general.byCluster),
    byConcept: joinRows(dA.general.byConcept, dB.general.byConcept)
  };
}

// ═════════════════════════════════════════════════════════════
//  FEATURE 1 — Risk view (alertas automáticas)
// ═════════════════════════════════════════════════════════════

/**
 * Devuelve top N rows con mayor desviación negativa vs B+F.
 * El frontend hoy ya tiene los datos para hacerlo en JS (más responsivo),
 * así que esta función es opcional. La dejo por si querés mover la lógica
 * al backend (mismo umbral en todos lados).
 *
 *   { red: [...], amber:[...], green:[...] }
 */
function getRiskView(referenceMonth, referenceYear){
  var d = getDashboardData(referenceMonth, referenceYear, 'M', [], []);
  if(d.error) return { error:d.error };
  var rows = [];
  (d.general.byCountry||[]).forEach(function(r){ if(!r.isTotal) rows.push({dim:'country',  label:r.label, arr:r.arr, bf:r.bf, delta:(r.arr-r.bf), pct:r.bf?(r.arr-r.bf)/r.bf*100:0}); });
  (d.general.byCluster||[]).forEach(function(r){ if(!r.isTotal) rows.push({dim:'cluster',  label:r.label, arr:r.arr, bf:r.bf, delta:(r.arr-r.bf), pct:r.bf?(r.arr-r.bf)/r.bf*100:0}); });
  (d.general.byConcept||[]).forEach(function(r){ if(!r.isTotal) rows.push({dim:'concept',  label:r.label, arr:r.arr, bf:r.bf, delta:(r.arr-r.bf), pct:r.bf?(r.arr-r.bf)/r.bf*100:0}); });
  rows.sort(function(a,b){return Math.abs(b.delta)-Math.abs(a.delta);});
  var red = rows.filter(function(r){return r.pct < -10;}).slice(0,5);
  var green = rows.filter(function(r){return r.pct > 10;}).slice(0,5);
  return { red:red, green:green, periodLabel:d.periodLabel };
}

// ═════════════════════════════════════════════════════════════
//  PIVOT EXPLORER — análisis ad-hoc con filas/cols/medidas/filtros
// ═════════════════════════════════════════════════════════════

/**
 * Schema del dataset principal. Define qué campos son dimensiones,
 * qué campos son medidas y qué medidas computadas se ofrecen.
 * Mantener sincronizado con los headers reales de la hoja `dataset`.
 */
var PIVOT_SCHEMA = {
  fields: [
    { id:'year',     label:'Año',          type:'dim',  headers:['Date - Year','Year','Año','Ano'] },
    { id:'quarter',  label:'Trimestre',    type:'dim',  headers:['Date - Quarter','Quarter','Trimestre'] },
    { id:'month',    label:'Mes',          type:'dim',  headers:['Date - Month','Month','Mes'], order:['January','February','March','April','May','June','July','August','September','October','November','December'] },
    { id:'country',  label:'País',         type:'dim',  headers:['País','Pais','Country'] },
    { id:'channel',  label:'LoB / Canal',  type:'dim',  headers:['Lob-Channel','Lob Channel','Canal'] },
    { id:'concept',  label:'Concepto',     type:'dim',  headers:['Concepto','Concept'] },
    { id:'partner',  label:'Partner',      type:'dim',  headers:['Partner'] },
    { id:'cluster',  label:'Cluster',      type:'dim',  headers:['Cluster'] },
    { id:'producto', label:'Producto',     type:'dim',  headers:['Producto_agrupado','Producto agrupado','Producto'] }
  ],
  measures: [
    { id:'arr',    label:'Actuals + RR', headers:['Actuals + RR','Actuals+RR'],                                                                                              agg:'sum', fmt:'usd' },
    { id:'budget', label:'Budget',       headers:['Budget','Budget + Forecast + Breakthrough FY26','Budget + Forecast FY26','Budget + Forecast','Budget+Forecast','Budget FY27'], agg:'sum', fmt:'usd' },
    { id:'rr',     label:'RR',           headers:['RR','Rolling Reforecast','Reforecast'],                                                                                   agg:'sum', fmt:'usd' }
  ],
  computed: [
    { id:'delta_budget', label:'Δ vs Budget',  deps:['arr','budget'], fmt:'usd', formula:function(m){ return (m.arr||0) - (m.budget||0); } },
    { id:'pct_budget',   label:'Δ% vs Budget', deps:['arr','budget'], fmt:'pct', formula:function(m){ return (m.budget||0) ? ((m.arr - m.budget) / m.budget * 100) : 0; } },
    // Variance vs RR: para meses cerrados, mide cuánto le pegamos al forecast original
    { id:'delta_rr',     label:'Δ vs RR',      deps:['arr','rr'],     fmt:'usd', formula:function(m){ return (m.arr||0) - (m.rr||0); } },
    { id:'pct_rr',       label:'Δ% vs RR',     deps:['arr','rr'],     fmt:'pct', formula:function(m){ return (m.rr||0) ? ((m.arr - m.rr) / m.rr * 100) : 0; } }
  ]
};

/**
 * Devuelve el schema al frontend (sin las funciones, sólo descripciones).
 */
function getPivotSchema(){
  return {
    fields: PIVOT_SCHEMA.fields.map(function(f){ return { id:f.id, label:f.label, type:f.type }; }),
    measures: PIVOT_SCHEMA.measures.map(function(m){ return { id:m.id, label:m.label, fmt:m.fmt }; }),
    computed: PIVOT_SCHEMA.computed.map(function(c){ return { id:c.id, label:c.label, fmt:c.fmt, deps:c.deps }; })
  };
}

/** Resuelve la columna del header dado una lista de candidatos (reusa norm_/col_). */
function _resolveColIdx_(headers, candidates){
  if(typeof col_ === 'function') return col_(headers, candidates);
  // fallback
  for(var i=0;i<candidates.length;i++){
    var idx = headers.indexOf(candidates[i]);
    if(idx !== -1) return idx;
  }
  return -1;
}

/** Lee el dataset crudo y devuelve { headers, rows, fieldIdx, measureIdx }. */
function _loadDatasetCached_(){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName('dataset');
  if(!sh) throw new Error("No se encontró la hoja 'dataset'");
  var raw = sh.getDataRange().getValues();
  if(!raw || raw.length < 2) return { headers:[], rows:[], fieldIdx:{}, measureIdx:{} };
  var headers = raw[0].map(function(h){ return String(h).trim(); });
  var fieldIdx = {};
  PIVOT_SCHEMA.fields.forEach(function(f){
    fieldIdx[f.id] = _resolveColIdx_(headers, f.headers);
  });
  var measureIdx = {};
  PIVOT_SCHEMA.measures.forEach(function(m){
    measureIdx[m.id] = _resolveColIdx_(headers, m.headers);
  });
  return { headers:headers, raw:raw, fieldIdx:fieldIdx, measureIdx:measureIdx };
}

/**
 * Devuelve los valores distintos (ordenados) de un campo dimensión.
 * Útil para poblar los multiselect de filtros.
 *   google.script.run.getDistinctValues('country')
 */
function getDistinctValues(fieldId){
  var ds = _loadDatasetCached_();
  var idx = ds.fieldIdx[fieldId];
  if(idx == null || idx < 0) return { ok:false, reason:'Campo no encontrado: '+fieldId };
  var fieldDef = PIVOT_SCHEMA.fields.filter(function(f){ return f.id===fieldId; })[0];
  var seen = {};
  for(var i=1;i<ds.raw.length;i++){
    var v = ds.raw[i][idx];
    if(v === '' || v == null) continue;
    var s = String(v).trim();
    if(!s) continue;
    seen[s] = 1;
  }
  var values = Object.keys(seen);
  if(fieldDef && fieldDef.order){
    var order = fieldDef.order;
    values.sort(function(a,b){
      var ia = order.indexOf(a), ib = order.indexOf(b);
      if(ia === -1 && ib === -1) return a.localeCompare(b);
      if(ia === -1) return 1;
      if(ib === -1) return -1;
      return ia - ib;
    });
  } else if(fieldId === 'year'){
    values.sort(function(a,b){ return Number(b) - Number(a); }); // más reciente primero
  } else {
    values.sort(function(a,b){ return a.localeCompare(b); });
  }
  return { ok:true, values: values };
}

/**
 * Ejecuta una pivot sobre el dataset.
 *
 * config = {
 *   rows:     ['country','channel'],
 *   cols:     ['quarter']           // opcional; vacío = sin columnas
 *   measures: ['arr','bf26','delta_bf','pct_bf'],
 *   filters:  { year:['2026'], country:['Argentina','Brasil'] },  // multiselect por campo
 *   limit:    5000   // opcional, corta filas para queries pesadas
 * }
 *
 * return {
 *   ok, rowDims, colDims, measures (con metadata fmt/label),
 *   rowKeys:  [[v1,v2],...],
 *   colKeys:  [[v1],...],
 *   cells:    { 'r1|r2': { 'c1': { measureId:value,... }, ... }, ... },
 *   rowTotals:{ 'r1|r2': { measureId:value, ... } },
 *   colTotals:{ 'c1':    { measureId:value, ... } },
 *   grandTotal:{ measureId:value, ... },
 *   rowCount, datasetSize
 * }
 */
function runPivot(config){
  config = config || {};
  var rows     = (config.rows || []).filter(Boolean);
  var cols     = (config.cols || []).filter(Boolean);
  var measures = (config.measures || []).filter(Boolean);
  var filters  = config.filters || {};
  // Límite generoso para datasets en crecimiento. Apps Script tiene 6 min de ejecución;
  // con 1M filas + suma simple estamos cómodos. El bottleneck real es getDataRange().getValues().
  var limit    = config.limit || 1000000;

  if(!measures.length) return { ok:false, reason:'Elegí al menos una medida.' };

  var ds = _loadDatasetCached_();

  // Resolver field defs y validaciones
  function fieldDef(id){ return PIVOT_SCHEMA.fields.filter(function(f){return f.id===id;})[0]; }
  function measureDef(id){
    return PIVOT_SCHEMA.measures.filter(function(m){return m.id===id;})[0]
        || PIVOT_SCHEMA.computed.filter(function(c){return c.id===id;})[0];
  }
  function isComputed(id){ return PIVOT_SCHEMA.computed.some(function(c){return c.id===id;}); }

  // Validar filas/cols
  var rowIdxs = rows.map(function(id){
    var i = ds.fieldIdx[id];
    if(i == null || i < 0) throw new Error('Campo de filas no encontrado: '+id);
    return i;
  });
  var colIdxs = cols.map(function(id){
    var i = ds.fieldIdx[id];
    if(i == null || i < 0) throw new Error('Campo de columnas no encontrado: '+id);
    return i;
  });

  // Pre-resolver índices de medidas base requeridas (incluye deps de computadas)
  var baseMeasuresNeeded = {};
  measures.forEach(function(mid){
    if(isComputed(mid)){
      var c = measureDef(mid);
      (c.deps||[]).forEach(function(d){ baseMeasuresNeeded[d] = 1; });
    } else {
      baseMeasuresNeeded[mid] = 1;
    }
  });
  var baseMeasureList = Object.keys(baseMeasuresNeeded);

  // Pre-resolver índices de filtros
  var filterDefs = [];
  Object.keys(filters).forEach(function(fid){
    var vals = filters[fid];
    if(!Array.isArray(vals) || !vals.length) return;
    var i = ds.fieldIdx[fid];
    if(i == null || i < 0) return;
    var setLc = {};
    vals.forEach(function(v){ setLc[String(v).toLowerCase()] = 1; });
    filterDefs.push({ idx:i, set:setLc });
  });

  // Agregación en una sola pasada
  var cells = {};            // rowKey -> colKey -> { baseMeasureId: sum }
  var rowTotals = {};        // rowKey -> { baseMeasureId: sum }
  var colTotals = {};        // colKey -> { baseMeasureId: sum }
  var grand = {};            // baseMeasureId: sum
  var rowKeysSeen = {};      // rowKey -> [val1,val2]
  var colKeysSeen = {};      // colKey -> [val1]
  var rowCount = 0;

  baseMeasureList.forEach(function(m){ grand[m] = 0; });

  var raw = ds.raw;
  var rawLen = raw.length;
  for(var r=1;r<rawLen && rowCount<limit; r++){
    var row = raw[r];

    // Aplicar filtros
    var passes = true;
    for(var fi=0; fi<filterDefs.length; fi++){
      var fd = filterDefs[fi];
      var v = String(row[fd.idx]||'').toLowerCase();
      if(!fd.set[v]){ passes = false; break; }
    }
    if(!passes) continue;

    // Construir keys
    var rowParts = rowIdxs.map(function(i){ var v = row[i]; return v==null?'':String(v).trim(); });
    var colParts = colIdxs.map(function(i){ var v = row[i]; return v==null?'':String(v).trim(); });
    // Si una dim está vacía, mostrar como "(vacío)"
    rowParts = rowParts.map(function(p){ return p === '' ? '(vacío)' : p; });
    colParts = colParts.map(function(p){ return p === '' ? '(vacío)' : p; });

    var rowKey = rowParts.join('|');
    var colKey = colParts.join('|');
    if(!rowKeysSeen[rowKey]) rowKeysSeen[rowKey] = rowParts;
    if(cols.length && !colKeysSeen[colKey]) colKeysSeen[colKey] = colParts;

    // Acumular medidas base
    if(!cells[rowKey]) cells[rowKey] = {};
    if(!cells[rowKey][colKey]){
      cells[rowKey][colKey] = {};
      baseMeasureList.forEach(function(m){ cells[rowKey][colKey][m] = 0; });
    }
    if(!rowTotals[rowKey]){
      rowTotals[rowKey] = {};
      baseMeasureList.forEach(function(m){ rowTotals[rowKey][m] = 0; });
    }
    if(cols.length){
      if(!colTotals[colKey]){
        colTotals[colKey] = {};
        baseMeasureList.forEach(function(m){ colTotals[colKey][m] = 0; });
      }
    }

    baseMeasureList.forEach(function(m){
      var mIdx = ds.measureIdx[m];
      if(mIdx == null || mIdx < 0) return;
      var v = parseWA_(row[mIdx]);
      cells[rowKey][colKey][m] += v;
      rowTotals[rowKey][m] += v;
      if(cols.length) colTotals[colKey][m] += v;
      grand[m] += v;
    });

    rowCount++;
  }

  // Aplicar computadas
  function applyComputed(targetObj){
    measures.forEach(function(mid){
      if(!isComputed(mid)) return;
      var c = measureDef(mid);
      try{
        targetObj[mid] = c.formula(targetObj);
      }catch(e){
        targetObj[mid] = null;
      }
    });
  }
  Object.keys(cells).forEach(function(rk){
    Object.keys(cells[rk]).forEach(function(ck){ applyComputed(cells[rk][ck]); });
  });
  Object.keys(rowTotals).forEach(function(rk){ applyComputed(rowTotals[rk]); });
  Object.keys(colTotals).forEach(function(ck){ applyComputed(colTotals[ck]); });
  applyComputed(grand);

  // Ordenar keys
  function sortKeys(seenMap, dimIds){
    var keys = Object.keys(seenMap);
    var parts = keys.map(function(k){ return seenMap[k]; });
    // Buscar orden por la primera dim que tenga `order` definido o usar magnitud de la primera medida
    keys.sort(function(a,b){
      var pa = seenMap[a], pb = seenMap[b];
      for(var d=0; d<dimIds.length; d++){
        var def = fieldDef(dimIds[d]);
        var va = pa[d], vb = pb[d];
        if(def && def.order){
          var ia = def.order.indexOf(va), ib = def.order.indexOf(vb);
          if(ia === -1 && ib === -1){ var cmp = String(va).localeCompare(String(vb)); if(cmp) return cmp; continue; }
          if(ia === -1) return 1;
          if(ib === -1) return -1;
          if(ia !== ib) return ia - ib;
          continue;
        }
        if(dimIds[d] === 'year'){
          var na = Number(va), nb = Number(vb);
          if(na !== nb) return nb - na; // año desc
          continue;
        }
        var c2 = String(va).localeCompare(String(vb));
        if(c2) return c2;
      }
      return 0;
    });
    return keys.map(function(k){ return seenMap[k]; });
  }

  var rowKeys = sortKeys(rowKeysSeen, rows);
  var colKeys = cols.length ? sortKeys(colKeysSeen, cols) : [];

  return {
    ok: true,
    rowDims: rows.map(function(id){ var f = fieldDef(id); return { id:id, label:f?f.label:id }; }),
    colDims: cols.map(function(id){ var f = fieldDef(id); return { id:id, label:f?f.label:id }; }),
    measures: measures.map(function(mid){
      var def = measureDef(mid);
      return { id:mid, label:def?def.label:mid, fmt:def?def.fmt:'usd', computed:isComputed(mid) };
    }),
    rowKeys: rowKeys,
    colKeys: colKeys,
    cells: cells,
    rowTotals: rowTotals,
    colTotals: colTotals,
    grandTotal: grand,
    rowCount: rowCount,
    datasetSize: rawLen - 1
  };
}

// ═════════════════════════════════════════════════════════════
//  VISTAS GUARDADAS (Saved Views) — Explorer y otros tabs
// ═════════════════════════════════════════════════════════════
//  Hoja: 'vistas_guardadas'
//  Columnas: id | name | scope | config_json | description | created_by | created_at | updated_at
//  Acceso: listar/aplicar = cualquier usuario logueado. Guardar/borrar = admin.

var SHEET_VIEWS = 'vistas_guardadas';

function _ensureViewsSheet_(){
  var ss = _ss_();
  var sh = ss.getSheetByName(SHEET_VIEWS);
  if(!sh){ sh = ss.insertSheet(SHEET_VIEWS); }
  if(sh.getLastRow() === 0){
    sh.getRange(1,1,1,8).setValues([[
      'id','name','scope','config_json','description','created_by','created_at','updated_at'
    ]]);
    sh.getRange(1,1,1,8).setFontWeight('bold').setBackground('#1a1d23').setFontColor('#fff');
    sh.setFrozenRows(1);
  }
  return sh;
}

/**
 * Lista vistas guardadas, opcionalmente filtradas por scope (ej. 'explorer').
 *   google.script.run.listViews('explorer')
 */
function listViews(scope){
  var sh = _ensureViewsSheet_();
  var last = sh.getLastRow();
  if(last < 2) return { ok:true, views:[] };
  var data = sh.getRange(2,1,last-1,8).getValues();
  var out = [];
  for(var i=0;i<data.length;i++){
    var r = data[i];
    if(scope && String(r[2]) !== String(scope)) continue;
    out.push({
      id: String(r[0]||''),
      name: String(r[1]||''),
      scope: String(r[2]||''),
      description: String(r[4]||''),
      createdBy: String(r[5]||''),
      createdAt: r[6] ? new Date(r[6]).toISOString() : '',
      updatedAt: r[7] ? new Date(r[7]).toISOString() : ''
    });
  }
  out.sort(function(a,b){ return (b.updatedAt||'').localeCompare(a.updatedAt||''); });
  return { ok:true, views: out };
}

/**
 * Devuelve la config completa (parseada) de una vista por id.
 *   google.script.run.loadView('uuid')
 */
function loadView(id){
  var sh = _ensureViewsSheet_();
  var last = sh.getLastRow();
  if(last < 2) return { ok:false, reason:'No hay vistas' };
  var data = sh.getRange(2,1,last-1,8).getValues();
  for(var i=0;i<data.length;i++){
    if(String(data[i][0]) === String(id)){
      try{
        return {
          ok: true,
          id: String(data[i][0]),
          name: String(data[i][1]||''),
          scope: String(data[i][2]||''),
          config: JSON.parse(String(data[i][3]||'{}')),
          description: String(data[i][4]||''),
          createdBy: String(data[i][5]||''),
          createdAt: data[i][6] ? new Date(data[i][6]).toISOString() : '',
          updatedAt: data[i][7] ? new Date(data[i][7]).toISOString() : ''
        };
      }catch(e){
        return { ok:false, reason:'Config corrupta: '+e };
      }
    }
  }
  return { ok:false, reason:'No encontrada' };
}

/**
 * Guarda o actualiza una vista. Sólo admins.
 *
 * payload = {
 *   id: 'uuid|null',  // null = crear; con id = actualizar (sólo si existe)
 *   name: 'Mi cut mensual',
 *   scope: 'explorer',
 *   config: {...},    // objeto serializable
 *   description: 'opcional'
 * }
 */
function saveView(payload){
  var caller = _requireAdmin_();
  if(!payload || !payload.name || !payload.scope || !payload.config){
    return { ok:false, reason:'Faltan campos (name/scope/config)' };
  }
  var sh = _ensureViewsSheet_();
  var name   = String(payload.name).slice(0,100);
  var scope  = String(payload.scope).slice(0,32);
  var desc   = String(payload.description||'').slice(0,500);
  var cfg    = JSON.stringify(payload.config).slice(0,40000);
  var now    = new Date();

  // Update si trae id existente
  if(payload.id){
    var last = sh.getLastRow();
    if(last >= 2){
      var data = sh.getRange(2,1,last-1,8).getValues();
      for(var i=0;i<data.length;i++){
        if(String(data[i][0]) === String(payload.id)){
          sh.getRange(i+2,2).setValue(name);
          sh.getRange(i+2,3).setValue(scope);
          sh.getRange(i+2,4).setValue(cfg);
          sh.getRange(i+2,5).setValue(desc);
          sh.getRange(i+2,8).setValue(now);
          return { ok:true, id:String(payload.id), updated:true };
        }
      }
    }
  }
  // Insert
  var id = _uuid_();
  sh.appendRow([id, name, scope, cfg, desc, caller, now, now]);
  return { ok:true, id:id, created:true };
}

/** Borra una vista. Sólo admins. */
function deleteView(id){
  _requireAdmin_();
  var sh = _ensureViewsSheet_();
  var last = sh.getLastRow();
  if(last < 2) return { ok:false, reason:'No hay vistas' };
  var data = sh.getRange(2,1,last-1,1).getValues();
  for(var i=0;i<data.length;i++){
    if(String(data[i][0]) === String(id)){
      sh.deleteRow(i+2);
      return { ok:true };
    }
  }
  return { ok:false, reason:'No encontrada' };
}

// ═════════════════════════════════════════════════════════════
//  Test util
// ═════════════════════════════════════════════════════════════
function testV2(){
  Logger.log('User: ' + _userEmail_());
  Logger.log('Comments index: ' + JSON.stringify(getCommentsIndex()).slice(0,400));
  Logger.log('Pareto: ' + JSON.stringify(getPareto('', '')).slice(0,400));
  Logger.log('Risk: ' + JSON.stringify(getRiskView('', '')).slice(0,400));
}


// ═════════════════════════════════════════════════════════════
//  TOQAN AI — Integración con Toqan (https://work.toqan.ai)
//  API asincrónica: POST /create_conversation → poll POST /get_answer
//  Docs: https://toqan-api.readme.io
//  Auth: header X-Api-Key (NO "Authorization: Bearer")
// ═════════════════════════════════════════════════════════════

const TOQAN_BASE_URL = 'https://api.coco.prod.toqan.ai';

/**
 * PASO 1 — Ejecutar UNA sola vez para guardar la API key.
 * Editá la constante TOQAN_API_KEY abajo con tu key real antes de correr.
 */
function guardarAPIKeyToqan() {
  const TOQAN_API_KEY = 'sk_PEGAR_TU_KEY_ACA';
  if (TOQAN_API_KEY === 'sk_PEGAR_TU_KEY_ACA') {
    throw new Error('❌ Editá la función guardarAPIKeyToqan() y pegá tu API key real antes de correrla.');
  }
  PropertiesService.getScriptProperties().setProperty('TOQAN_API_KEY', TOQAN_API_KEY);
  Logger.log('✅ API Key guardada en Script Properties (TOQAN_API_KEY).');
}

/**
 * PASO 2 — Smoke test. Manda un "hola" a Toqan y loguea la respuesta.
 * Si esto anda, el resto (chat widget) es directo.
 * Ver logs en: Apps Script → Executions → testToqan.
 */
function testToqan() {
  const apikey = PropertiesService.getScriptProperties().getProperty('TOQAN_API_KEY');
  if (!apikey) {
    Logger.log('❌ Falta API key. Corré guardarAPIKeyToqan() primero.');
    return;
  }

  Logger.log('🔄 Creando conversación con Toqan...');
  const conv = toqanCreateConversation_(apikey, 'Respondé con UNA sola frase corta para confirmar que la integración funciona.');
  Logger.log(`✓ conversation_id=${conv.conversation_id}  request_id=${conv.request_id}`);

  Logger.log('🔄 Esperando respuesta (poll cada 3s, máx ~2 min)...');
  const result = toqanPollUntilComplete_(apikey, conv.conversation_id, conv.request_id, { maxAttempts: 40, intervalMs: 3000 });

  Logger.log('✅ Status: ' + result.status);
  Logger.log('💬 Respuesta completa (JSON):');
  Logger.log(JSON.stringify(result, null, 2));
}

/**
 * Diagnóstico — corré esto si testToqan() falla.
 * Muestra: si la key está seteada, qué devuelve el endpoint y todos los headers de respuesta.
 * NO loguea la key completa por seguridad.
 */
function diagnosticarToqan() {
  const apikey = PropertiesService.getScriptProperties().getProperty('TOQAN_API_KEY');

  Logger.log('═══ DIAGNÓSTICO TOQAN ═══');
  if (!apikey) {
    Logger.log('❌ No hay API key en Script Properties. Corré guardarAPIKeyToqan() primero.');
    return;
  }
  Logger.log(`✓ API key seteada — longitud=${apikey.length}, primeros 4=${apikey.slice(0,4)}, últimos 4=${apikey.slice(-4)}`);
  Logger.log(`✓ URL base: ${TOQAN_BASE_URL}`);

  Logger.log('');
  Logger.log('--- Test: POST /create_conversation con X-Api-Key ---');
  _toqanDiagRequest_(`${TOQAN_BASE_URL}/create_conversation`, 'post', { 'X-Api-Key': apikey, 'Content-Type': 'application/json' }, JSON.stringify({ user_message: 'ping' }));
}

function _toqanDiagRequest_(url, method, headers, payload) {
  try {
    const opts = { method: method, headers: headers, muteHttpExceptions: true };
    if (payload) opts.payload = payload;
    const res = UrlFetchApp.fetch(url, opts);
    Logger.log(`  URL: ${url}`);
    Logger.log(`  Status: ${res.getResponseCode()}`);
    const respHeaders = res.getAllHeaders();
    Logger.log(`  Response headers: ${JSON.stringify(respHeaders)}`);
    const body = res.getContentText();
    Logger.log(`  Body (primeros 500 chars): ${body.slice(0, 500)}`);
  } catch (e) {
    Logger.log(`  ❌ Excepción: ${e.message}`);
  }
}

// --------- Helpers privados ---------

function toqanCreateConversation_(apikey, userMessage) {
  const payload = { user_message: userMessage };
  const res = UrlFetchApp.fetch(`${TOQAN_BASE_URL}/create_conversation`, {
    method: 'post',
    headers: { 'X-Api-Key': apikey, 'Content-Type': 'application/json' },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });
  const code = res.getResponseCode();
  const body = res.getContentText();
  if (code !== 200) throw new Error(`Toqan createConversation falló [${code}]: ${body}`);
  return JSON.parse(body);
}

function toqanContinueConversation_(apikey, conversationId, userMessage) {
  const payload = { conversation_id: conversationId, user_message: userMessage };
  const res = UrlFetchApp.fetch(`${TOQAN_BASE_URL}/continue_conversation`, {
    method: 'post',
    headers: { 'X-Api-Key': apikey, 'Content-Type': 'application/json' },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });
  const code = res.getResponseCode();
  const body = res.getContentText();
  if (code !== 200) throw new Error(`Toqan continueConversation falló [${code}]: ${body}`);
  return JSON.parse(body);
}

function toqanGetAnswer_(apikey, conversationId, requestId) {
  const payload = { conversation_id: conversationId, request_id: requestId };
  const res = UrlFetchApp.fetch(`${TOQAN_BASE_URL}/get_answer`, {
    method: 'post',
    headers: { 'X-Api-Key': apikey, 'Content-Type': 'application/json' },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });
  const code = res.getResponseCode();
  const body = res.getContentText();
  if (code !== 200) throw new Error(`Toqan getAnswer falló [${code}]: ${body}`);
  return JSON.parse(body);
}

function toqanPollUntilComplete_(apikey, conversationId, requestId, opts) {
  const maxAttempts = (opts && opts.maxAttempts) || 40;
  const intervalMs  = (opts && opts.intervalMs)  || 3000;
  for (let i = 0; i < maxAttempts; i++) {
    const data = toqanGetAnswer_(apikey, conversationId, requestId);
    if (data.status === 'finished') return data;
    if (data.status === 'error')    throw new Error('Toqan request error: ' + JSON.stringify(data));
    Logger.log(`  ⏳ intento ${i+1}/${maxAttempts} — status=${data.status}`);
    Utilities.sleep(intervalMs);
  }
  throw new Error(`Timeout: Toqan no respondió en ${(maxAttempts * intervalMs) / 1000}s`);
}


// ═════════════════════════════════════════════════════════════
//  CUSTOM TOOL ENDPOINT (consumido desde work.toqan.ai)
// ═════════════════════════════════════════════════════════════
//
// Toqan llama acá vía POST con JSON body:
//   { "auth": "<TOOL_SHARED_SECRET>", "tool": "queryMediaSales", "params": {...} }
//
// Respuesta:
//   { "ok": true,  "tool": "...", "data": {...} }
//   { "ok": false, "error": "..." }
//
// Nota: Apps Script Web Apps no leen headers HTTP, por eso el secret va en
// el body en vez del header. Web Apps siempre devuelven 200, así que el
// status real va en el campo "ok".
// ═════════════════════════════════════════════════════════════

const TOOL_DIMENSIONS = ['year','quarter','month','country','cluster','lob','channel','concept','partner','producto'];

function doPost(e) {
  try {
    if (!e || !e.postData || !e.postData.contents) {
      return _toolJson_({ ok: false, error: 'Empty body' });
    }
    let req;
    try { req = JSON.parse(e.postData.contents); }
    catch (err) { return _toolJson_({ ok: false, error: 'Invalid JSON body' }); }

    const expected = PropertiesService.getScriptProperties().getProperty('TOOL_SHARED_SECRET');
    if (!expected) return _toolJson_({ ok: false, error: 'Server not configured — corré generarSharedSecret()' });
    if (req.auth !== expected) return _toolJson_({ ok: false, error: 'Unauthorized' });

    const tool   = req.tool || '';
    const params = req.params || {};

    switch (tool) {
      case 'queryMediaSales':
        return _toolJson_({ ok: true, tool: tool, data: tool_queryMediaSales_(params) });
      case 'listDimensions':
        return _toolJson_({ ok: true, tool: tool, data: tool_listDimensions_() });
      case 'ping':
        return _toolJson_({ ok: true, tool: tool, data: { pong: true, ts: new Date().toISOString() } });
      default:
        return _toolJson_({ ok: false, error: 'Unknown tool: "' + tool + '". Disponibles: queryMediaSales, listDimensions, ping' });
    }
  } catch (err) {
    return _toolJson_({ ok: false, error: String(err && err.message || err), stack: err && err.stack ? String(err.stack).slice(0, 500) : null });
  }
}

function _toolJson_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * Genera un shared secret nuevo y lo guarda en Script Properties.
 * Copialo del log y pegalo en la config del Custom Tool en Toqan.
 * Si lo volvés a correr, REEMPLAZA el anterior (y rompe el tool en Toqan hasta actualizarlo).
 */
function generarSharedSecret() {
  const secret = Utilities.getUuid().replace(/-/g, '') + Utilities.getUuid().replace(/-/g, '');
  PropertiesService.getScriptProperties().setProperty('TOOL_SHARED_SECRET', secret);
  Logger.log('✅ Shared secret generado y guardado como Script Property TOOL_SHARED_SECRET.');
  Logger.log('🔑 COPIÁ ESTE VALOR (lo vas a pegar en la config del Custom Tool en Toqan):');
  Logger.log('');
  Logger.log(secret);
  Logger.log('');
  Logger.log('⚠️ NO lo compartas en chats ni lo pegues en el código. Vive solo en Script Properties y en Toqan.');
}

/**
 * Test local — simula un POST de Toqan llamando a doPost() directamente.
 * Útil para verificar que el endpoint funciona ANTES de exponerlo como Web App.
 */
function testTool() {
  const secret = PropertiesService.getScriptProperties().getProperty('TOOL_SHARED_SECRET');
  if (!secret) {
    Logger.log('❌ Falta TOOL_SHARED_SECRET. Corré generarSharedSecret() primero.');
    return;
  }

  Logger.log('═══ Test 1: ping ═══');
  _runFakePost_({ auth: secret, tool: 'ping' });

  Logger.log('');
  Logger.log('═══ Test 2: listDimensions (qué valores tiene cada dimensión) ═══');
  _runFakePost_({ auth: secret, tool: 'listDimensions' });

  Logger.log('');
  Logger.log('═══ Test 3: queryMediaSales — Brasil vs Budget, Q1 2026 ═══');
  _runFakePost_({
    auth: secret,
    tool: 'queryMediaSales',
    params: {
      group_by: 'country',
      filters: { country: 'Brasil', year: '2026', quarter: 'Q1' }
    }
  });

  Logger.log('');
  Logger.log('═══ Test 4: queryMediaSales — Top 5 clusters por desvío vs RR (mes actual del CONFIG) ═══');
  _runFakePost_({
    auth: secret,
    tool: 'queryMediaSales',
    params: {
      group_by: 'cluster',
      filters: { year: CONFIG.TARGET_YEAR, month: CONFIG.TARGET_MONTH },
      sort_by: 'vs_rr',
      sort_dir: 'desc',
      limit: 5
    }
  });

  Logger.log('');
  Logger.log('═══ Test 5: auth inválida (debe rechazar) ═══');
  _runFakePost_({ auth: 'wrong', tool: 'ping' });
}

function _runFakePost_(bodyObj) {
  const fakeEvent = { postData: { contents: JSON.stringify(bodyObj) } };
  const resp = doPost(fakeEvent);
  const txt = resp.getContent();
  // Truncar respuestas muy largas para no romper logs
  Logger.log(txt.length > 4000 ? txt.slice(0, 4000) + '... [truncado, total ' + txt.length + ' chars]' : txt);
}

// --------- Implementación de las tools ---------

/**
 * Agrega el dataset por una o más dimensiones, con filtros opcionales.
 * params:
 *   group_by   (string | string[])    — dimensiones a agrupar. Requerido.
 *   filters    (object, opcional)     — { country, cluster, lob, month, ... }, valores string o array
 *   sort_by    (string, opcional)     — actuals|budget|rr|vs_budget|vs_budget_pct|vs_rr|vs_rr_pct
 *   sort_dir   (string, opcional)     — "asc" | "desc" (default "desc")
 *   limit      (number, opcional)     — top N
 */
function tool_queryMediaSales_(params) {
  const ss = _ss_();
  const sheet = ss.getSheetByName(CONFIG.DATASET_SHEET);
  if (!sheet) throw new Error("No se encontró la hoja '" + CONFIG.DATASET_SHEET + "'");
  const rows = readDataset(sheet);
  if (!rows || rows.length === 0) throw new Error('Dataset vacío o sin columnas requeridas');

  // group_by
  let groupBy = params.group_by;
  if (!groupBy) throw new Error('Falta param "group_by". Dimensiones válidas: ' + TOOL_DIMENSIONS.join(', '));
  if (typeof groupBy === 'string') groupBy = [groupBy];
  if (!Array.isArray(groupBy)) throw new Error('"group_by" debe ser string o array de strings');
  const invalidDims = groupBy.filter(g => TOOL_DIMENSIONS.indexOf(g) === -1);
  if (invalidDims.length > 0) throw new Error('group_by tiene dimensiones inválidas: ' + invalidDims.join(', ') + '. Válidas: ' + TOOL_DIMENSIONS.join(', '));

  // filters (case-insensitive en valores string)
  const filters = params.filters || {};
  let filtered = rows;
  Object.keys(filters).forEach(key => {
    if (TOOL_DIMENSIONS.indexOf(key) === -1) return; // ignorar filtros desconocidos
    const val = filters[key];
    if (val == null || val === '') return;
    const values = (Array.isArray(val) ? val : [val]).map(v => String(v).toLowerCase());
    filtered = filtered.filter(r => values.indexOf(String(r[key]).toLowerCase()) !== -1);
  });

  // aggregate
  const bucket = {};
  filtered.forEach(r => {
    const key = groupBy.map(g => r[g]).join('||');
    if (!bucket[key]) {
      bucket[key] = {};
      groupBy.forEach(g => { bucket[key][g] = r[g]; });
      bucket[key].actuals = 0;
      bucket[key].budget  = 0;
      bucket[key].rr      = 0;
    }
    bucket[key].actuals += r.actuals || 0;
    bucket[key].budget  += r.budget  || 0;
    bucket[key].rr      += r.rr      || 0;
  });

  // compute deviations + round
  let out = Object.keys(bucket).map(k => {
    const row = bucket[k];
    row.vs_budget = row.actuals - row.budget;
    row.vs_budget_pct = row.budget !== 0 ? (row.actuals - row.budget) / row.budget : null;
    row.vs_rr = row.actuals - row.rr;
    row.vs_rr_pct = row.rr !== 0 ? (row.actuals - row.rr) / row.rr : null;
    ['actuals','budget','rr','vs_budget','vs_rr'].forEach(m => { row[m] = Math.round(row[m]); });
    ['vs_budget_pct','vs_rr_pct'].forEach(m => { if (row[m] != null) row[m] = Math.round(row[m] * 10000) / 10000; });
    return row;
  });

  // sort
  const sortBy  = params.sort_by;
  const sortDir = (params.sort_dir || 'desc').toString().toLowerCase() === 'asc' ? 1 : -1;
  if (sortBy) {
    out.sort((a, b) => {
      let av = a[sortBy], bv = b[sortBy];
      if (av == null) av = sortDir === 1 ?  Infinity : -Infinity;
      if (bv == null) bv = sortDir === 1 ?  Infinity : -Infinity;
      return (av - bv) * sortDir;
    });
  }

  // limit
  const limit = parseInt(params.limit, 10);
  if (limit > 0) out = out.slice(0, limit);

  // totals across filtered dataset (sin agrupar)
  const totals = filtered.reduce((acc, r) => {
    acc.actuals += r.actuals || 0;
    acc.budget  += r.budget  || 0;
    acc.rr      += r.rr      || 0;
    return acc;
  }, { actuals: 0, budget: 0, rr: 0 });
  totals.vs_budget = totals.actuals - totals.budget;
  totals.vs_budget_pct = totals.budget !== 0 ? (totals.actuals - totals.budget) / totals.budget : null;
  totals.vs_rr = totals.actuals - totals.rr;
  totals.vs_rr_pct = totals.rr !== 0 ? (totals.actuals - totals.rr) / totals.rr : null;
  ['actuals','budget','rr','vs_budget','vs_rr'].forEach(m => { totals[m] = Math.round(totals[m]); });
  ['vs_budget_pct','vs_rr_pct'].forEach(m => { if (totals[m] != null) totals[m] = Math.round(totals[m] * 10000) / 10000; });

  return {
    group_by: groupBy,
    filters: filters,
    sort_by: sortBy || null,
    sort_dir: sortDir === 1 ? 'asc' : 'desc',
    rows: out,
    row_count: out.length,
    rows_aggregated_from: filtered.length,
    totals_filtered: totals
  };
}

/**
 * Devuelve los valores únicos de cada dimensión del dataset.
 * Útil para que la IA sepa qué países/clusters/meses existen antes de filtrar.
 */
function tool_listDimensions_() {
  const ss = _ss_();
  const sheet = ss.getSheetByName(CONFIG.DATASET_SHEET);
  if (!sheet) throw new Error("No se encontró la hoja '" + CONFIG.DATASET_SHEET + "'");
  const rows = readDataset(sheet);

  const dims = {};
  TOOL_DIMENSIONS.forEach(d => { dims[d] = {}; });
  rows.forEach(r => {
    TOOL_DIMENSIONS.forEach(d => {
      const v = r[d];
      if (v != null && v !== '') dims[d][v] = true;
    });
  });
  const out = {};
  TOOL_DIMENSIONS.forEach(d => {
    out[d] = Object.keys(dims[d]).sort();
  });
  return { dimensions: out, total_rows: rows.length };
}


// ═════════════════════════════════════════════════════════════
//  DIRECTORY — Lista de usuarios @despegar.com para @menciones
// ═════════════════════════════════════════════════════════════
//
// Sheet "Directory" con columnas: email | nombre
// La mantiene el usuario manualmente. Sirve para:
//  (a) typeahead en el textarea de comentarios
//  (b) resolver @<local-part> → email completo al mandar mail
// ═════════════════════════════════════════════════════════════

function _ensureDirectorySheet_(){
  var ss = _ss_();
  var sh = ss.getSheetByName(SHEET_DIRECTORY);
  if(!sh){
    sh = ss.insertSheet(SHEET_DIRECTORY);
    sh.getRange(1,1,1,2).setValues([['email','name']]);
    sh.getRange(1,1,1,2).setFontWeight('bold').setBackground('#1a1d23').setFontColor('#fff');
    sh.setFrozenRows(1);
    // Seed con el equipo actual de AREA_CONFIG (sin emails — solo nombres como referencia)
    // El usuario completa los emails manualmente.
    var seedRows = [];
    Object.keys(AREA_CONFIG).forEach(function(k){
      (AREA_CONFIG[k].team || []).forEach(function(p){
        if(p.name) seedRows.push(['', p.name]);
      });
    });
    if(SR_MANAGER && SR_MANAGER.name) seedRows.push(['', SR_MANAGER.name]);
    if(seedRows.length){
      sh.getRange(2, 1, seedRows.length, 2).setValues(seedRows);
    }
  }
  return sh;
}

/**
 * Llamado desde frontend: google.script.run.getDirectory()
 * Devuelve [{email, name}, ...] con todas las filas válidas (email no vacío).
 */
function getDirectory(){
  var sh = _ensureDirectorySheet_();
  var last = sh.getLastRow();
  if(last < 2) return { users: [] };
  var data = sh.getRange(2, 1, last-1, 2).getValues();
  var users = [];
  data.forEach(function(r){
    var email = String(r[0]||'').trim().toLowerCase();
    var name  = String(r[1]||'').trim();
    if(!email) return; // saltar filas sin email
    if(!_validEmail_(email)) return;
    users.push({ email: email, name: name || email.split('@')[0] });
  });
  return { users: users };
}

/**
 * Parsea menciones del texto. Patrones soportados:
 *  - @usuario.nombre@despegar.com     → email completo
 *  - @usuario.nombre                  → resuelve via Directory (busca email que empiece con eso)
 * Devuelve array de emails únicos (lowercase).
 */
function _parseMentions_(text, directory){
  var out = {};
  var t = String(text||'');
  // 1) Emails completos
  var emailRe = /@([a-z0-9._-]+@[a-z0-9.-]+\.[a-z]{2,})/gi;
  var m;
  while((m = emailRe.exec(t)) !== null){
    out[m[1].toLowerCase()] = true;
  }
  // 2) Locales (sin @domain) → resolver via directory
  // Eliminamos primero los emails completos del texto para no doble-matchear
  var t2 = t.replace(emailRe, '');
  var localRe = /@([a-z0-9._-]+)/gi;
  var users = (directory && directory.users) || [];
  var byLocal = {};
  users.forEach(function(u){
    var local = u.email.split('@')[0];
    byLocal[local] = u.email;
  });
  while((m = localRe.exec(t2)) !== null){
    var resolved = byLocal[m[1].toLowerCase()];
    if(resolved) out[resolved] = true;
  }
  return Object.keys(out);
}

/**
 * Manda mail con el comentario a una lista de destinatarios.
 * Construye un HTML body con: autor, texto, tab, período, filtros, link al dashboard.
 */
function _sendCommentMail_(recipients, payload, commentId){
  if(!recipients || recipients.length === 0) return;
  try {
    var author = _userEmail_();
    var period = String(payload.period||'');
    var text   = String(payload.text||'');
    var filters = payload.filters || {};
    var tab     = filters.tab || '';
    var webAppUrl = '';
    try { webAppUrl = ScriptApp.getService().getUrl() || ''; } catch(e){}
    // Construir link al dashboard de Media
    var link = webAppUrl ? (webAppUrl + '?view=media') : '';

    var filterBadges = '';
    var FORDER = ['pos','lob','cluster','partner','producto'];
    var FLABEL = { pos:'Pos', lob:'Lob', cluster:'Cluster', partner:'Partner', producto:'Producto' };
    FORDER.forEach(function(k){
      var v = filters[k];
      if(Array.isArray(v) && v.length > 0){
        filterBadges += '<span style="display:inline-block;background:#EDE7F6;border:1px solid #C4B5FD;color:#5B21B6;padding:2px 8px;border-radius:4px;font-size:11px;margin:0 4px 4px 0;">'+
          '<strong>'+FLABEL[k]+':</strong> '+_escHtml_(v.join(', '))+
        '</span>';
      }
    });

    var html =
      '<div style="font-family:Arial,sans-serif;color:#1f2937;max-width:640px">'+
        '<div style="background:#6366F1;color:#fff;padding:16px 20px;border-radius:8px 8px 0 0">'+
          '<div style="font-size:11px;letter-spacing:0.08em;opacity:0.85;text-transform:uppercase">Media Sales · Comentario nuevo</div>'+
          '<div style="font-size:16px;font-weight:600;margin-top:4px">'+_escHtml_(author)+' te mencionó</div>'+
        '</div>'+
        '<div style="background:#ffffff;border:1px solid #E5E7EB;border-top:0;padding:18px 20px;border-radius:0 0 8px 8px">'+
          '<div style="font-size:14px;line-height:1.55;white-space:pre-wrap;color:#111827;background:#F9FAFB;border-left:3px solid #6366F1;padding:12px 14px;border-radius:4px;margin-bottom:14px">'+
            _escHtml_(text)+
          '</div>'+
          '<table style="font-size:12.5px;color:#374151;line-height:1.6"><tbody>'+
            (tab    ? '<tr><td style="padding-right:10px;color:#6B7280">Tab:</td><td><strong>'+_escHtml_(tab)+'</strong></td></tr>' : '')+
            (period ? '<tr><td style="padding-right:10px;color:#6B7280">Período:</td><td><strong>'+_escHtml_(period)+'</strong></td></tr>' : '')+
          '</tbody></table>'+
          (filterBadges ? '<div style="margin-top:10px"><div style="font-size:11px;color:#6B7280;margin-bottom:4px">Filtros activos al comentar:</div>'+filterBadges+'</div>' : '')+
          (link ? '<div style="margin-top:18px"><a href="'+_escHtml_(link)+'" style="background:#6366F1;color:#fff;text-decoration:none;padding:9px 16px;border-radius:6px;font-size:13px;font-weight:600;display:inline-block">Abrir dashboard</a></div>' : '')+
        '</div>'+
        '<div style="font-size:10.5px;color:#9CA3AF;padding:10px 4px 0">'+
          'Este mail se generó automáticamente desde Panel-media. Para responder, hacelo en el dashboard.'+
        '</div>'+
      '</div>';

    var subject = '[Media Sales] '+author+' te mencionó en un comentario';

    MailApp.sendEmail({
      to: recipients.join(','),
      subject: subject,
      htmlBody: html,
      name: 'Panel-media'
    });
  } catch(e){
    // No queremos romper el save si el mail falla. Solo logueamos.
    try { console.error('[sendCommentMail]', e && e.message || e); } catch(_){}
  }
}

function _escHtml_(s){
  return String(s==null?'':s).replace(/[&<>"']/g, function(c){
    return ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'})[c];
  });
}

/**
 * Devuelve los emails únicos de autores de un hilo dado (comentarios con el mismo
 * targetKey y que sean parte del thread por parentId).
 */
function _threadAuthors_(targetKey, parentId){
  if(!parentId || !targetKey) return [];
  try {
    var sh = _ensureCommentsSheet_();
    var last = sh.getLastRow();
    if(last < 2) return [];
    var data = sh.getRange(2,1,last-1,Math.max(10, sh.getLastColumn())).getValues();
    var authors = {};
    // Recorrer el hilo: cualquier comment con mismo targetKey cuya id sea parentId
    // o cuyo parentId apunte al thread root. Para simplicidad, agarramos TODOS los authors
    // del targetKey (un thread es por-key en el modelo actual).
    data.forEach(function(r){
      var key = String(r[3]||'');
      if(key !== targetKey) return;
      var author = String(r[1]||'').toLowerCase();
      var resolved = r[7] === true || r[7] === 'TRUE';
      if(author && !resolved) authors[author] = true;
    });
    return Object.keys(authors);
  } catch(e){
    return [];
  }
}


// ═════════════════════════════════════════════════════════════
//  HUB CATALOG — Estructura del portal Finance (sections + dashboards)
// ═════════════════════════════════════════════════════════════
//
// Sheet "Hub_Catalog": una fila por dashboard. Sección se repite por fila.
// Para sections sin dashboards: una fila con section_name y dashboard_name vacío.
//
// Columnas:
//   A: section_name      C: section_order   E: dashboard_desc   G: dashboard_icon   I: scope
//   B: section_icon      D: dashboard_name  F: dashboard_url    H: dashboard_order
//
// scope: '' (todos) | 'admin' (solo admins)
// ═════════════════════════════════════════════════════════════

// Seed inicial. Replica la estructura visible en la plataforma de Finance.
// El usuario completa URLs/dashboards faltantes editando el sheet directamente.
var HUB_SEED_ = [
  // section_name, section_icon, section_order, dashboard_name, dashboard_desc, dashboard_url, dashboard_icon, dashboard_order, scope
  ['Highlights',            'i-trend',   1, '', '', '', '', 0, ''],
  ['Tactic',                'i-target',  2, '', '', '', '', 0, ''],
  ['Consolidated Results',  'i-trend',   3, 'MRM',                     'Monthly Revenue Management',      '', 'i-trend', 1, ''],
  ['Consolidated Results',  'i-trend',   3, 'Group P&L',               'P&L del grupo',                    '', 'i-trend', 2, ''],
  ['Consolidated Results',  'i-trend',   3, 'Net Income',              '',                                 '', 'i-trend', 3, ''],
  ['Consolidated Results',  'i-trend',   3, 'Cash Flow',               '',                                 '', 'i-trend', 4, ''],
  ['Consolidated Results',  'i-trend',   3, 'Industry & Market Share', '',                                 '', 'i-trend', 5, ''],
  ['Consolidated Results',  'i-trend',   3, 'Exposure',                '',                                 '', 'i-trend', 6, ''],
  ['Consolidated Results',  'i-trend',   3, 'Installments',            '',                                 '', 'i-trend', 7, ''],
  ['Consolidated Results',  'i-trend',   3, 'P&L by Legal Entity',     '',                                 '', 'i-trend', 8, ''],
  ['B2C',                   'i-users',   4, '', '', '', '', 0, ''],
  ['B2C OKRs',              'i-target',  5, '', '', '', '', 0, ''],
  ['B2B & B2B2C',           'i-users',   6, '', '', '', '', 0, ''],
  ['Media Sales',           'i-trend',   7, 'Panel completo',          'Revenue · Forecast · Pareto · Comentarios', '?view=media', 'i-trend', 1, ''],
  ['Travel Partners',       'i-plane',   8, '', '', '', '', 0, ''],
  ['Riesgos',               'i-shield',  9, '', '', '', '', 0, '']
];

function _ensureHubCatalogSheet_(){
  var ss = _ss_();
  var sh = ss.getSheetByName(SHEET_HUB_CATALOG);
  if(!sh){
    sh = ss.insertSheet(SHEET_HUB_CATALOG);
    var headers = ['section_name','section_icon','section_order','dashboard_name','dashboard_desc','dashboard_url','dashboard_icon','dashboard_order','scope'];
    sh.getRange(1,1,1,headers.length).setValues([headers])
      .setFontWeight('bold').setBackground('#1a1d23').setFontColor('#fff');
    sh.setFrozenRows(1);
    if(HUB_SEED_.length){
      sh.getRange(2, 1, HUB_SEED_.length, 9).setValues(HUB_SEED_);
    }
    // Anchos de columna sugeridos (best-effort)
    try {
      sh.setColumnWidth(1, 180); sh.setColumnWidth(2, 110); sh.setColumnWidth(3, 110);
      sh.setColumnWidth(4, 200); sh.setColumnWidth(5, 240); sh.setColumnWidth(6, 240);
      sh.setColumnWidth(7, 110); sh.setColumnWidth(8, 110); sh.setColumnWidth(9, 90);
    } catch(_){}
  }
  return sh;
}

/**
 * Devuelve la estructura del Hub para el frontend:
 *   { sections: [{ name, icon, order, dashboards: [...] }], user: { email, isAdmin } }
 * Filtra dashboards con scope='admin' si el usuario no es admin.
 * Las secciones siempre se devuelven aunque no tengan dashboards visibles
 * (para que el sidebar muestre el header con "Sin dashboards" o "Próximamente").
 */
function getHubCatalog(){
  var email = _userEmail_();
  var isAdmin = _isAdmin_(email);
  var sh = _ensureHubCatalogSheet_();
  var last = sh.getLastRow();
  if(last < 2){
    return { sections: [], user: { email: email, isAdmin: isAdmin } };
  }
  var data = sh.getRange(2, 1, last-1, 9).getValues();
  var bySection = {};
  data.forEach(function(r){
    var sec = String(r[0]||'').trim();
    if(!sec) return;
    var scope = String(r[8]||'').trim().toLowerCase();
    if(!bySection[sec]){
      bySection[sec] = {
        name: sec,
        icon: String(r[1]||'i-trend').trim() || 'i-trend',
        order: Number(r[2]) || 0,
        dashboards: []
      };
    }
    var dn = String(r[3]||'').trim();
    if(!dn) return; // fila placeholder de sección sin dashboard
    if(scope === 'admin' && !isAdmin) return; // ocultar a no-admins
    bySection[sec].dashboards.push({
      name: dn,
      description: String(r[4]||''),
      url: String(r[5]||''),
      icon: String(r[6]||'i-trend').trim() || 'i-trend',
      order: Number(r[7]) || 0,
      scope: scope
    });
  });
  var sections = Object.keys(bySection).map(function(k){ return bySection[k]; });
  sections.sort(function(a,b){ return a.order - b.order; });
  sections.forEach(function(s){
    s.dashboards.sort(function(a,b){ return a.order - b.order; });
  });
  return { sections: sections, user: { email: email, isAdmin: isAdmin } };
}
