
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
