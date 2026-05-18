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
    year   : findCol("Year","Date - Year","Año","Ano"),
    month  : findCol("Month","Date - Month","Mes"),
    country: findCol("Pais","País","Country"),
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
  if(missing.length>0){SpreadsheetApp.getUi().alert("⚠️ Columnas no encontradas:\n"+missing.join("\n")+"\n\nVerificá los nombres.");return[];}
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
  const headers=["Mes",activeBudgetLabel_(),"Actuals","Δ","Δ%"];
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
