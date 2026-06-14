// ══════════════════════════════════════════════════════════════════
// INSTRUMENTO DE DIAGNÓSTICO FINANCIERO DIGITAL
// Cavedatos / Bida Global — Programa de Finanzas Digitales para Cámaras
//
// INSTRUCCIONES DE INSTALACIÓN:
// 1. Abre tu Google Sheet "Madurez Digital Camaras"
// 2. Menú → Extensiones → Apps Script
// 3. Borra el código que aparece y pega TODO este archivo
// 4. Guarda (Ctrl+S) con el nombre "Diagnóstico Financiero Digital"
// 5. Menú → Implementar → Nueva implementación
// 6. Tipo: Aplicación web
//    Ejecutar como: Yo (tu cuenta)
//    Quién tiene acceso: Cualquier persona (para que el HTML pueda enviar)
// 7. Haz clic en "Implementar"
// 8. Copia la URL que aparece — esa URL va en el archivo HTML
// ══════════════════════════════════════════════════════════════════

const SHEET_NAME = "Diagnósticos";
const SPREADSHEET_ID = "11dvBRfyQzOZRdcGx72a6ZG2KffokJqnEL3c8KQ8WFNY";
const SHEET_HEADER = [
  "Fecha y hora",
  "Cámara",
  "Sector",
  "N° Agremiados",
  "Equipo financiero (personas)",
  "ERP actual",
  "Suite ofimática",
  "Tenant M365 activo",
  "Usa SharePoint",
  // Procedimientos (5)
  "Proc 1 — Nombre",
  "Proc 1 — Pasos",
  "Proc 1 — Tiempo/ciclo",
  "Proc 1 — Frecuencia",
  "Proc 1 — Herramienta",
  "Proc 1 — Dolor principal",
  "Proc 2 — Nombre",
  "Proc 2 — Pasos",
  "Proc 2 — Tiempo/ciclo",
  "Proc 2 — Frecuencia",
  "Proc 2 — Herramienta",
  "Proc 2 — Dolor principal",
  "Proc 3 — Nombre",
  "Proc 3 — Pasos",
  "Proc 3 — Tiempo/ciclo",
  "Proc 3 — Frecuencia",
  "Proc 3 — Herramienta",
  "Proc 3 — Dolor principal",
  "Proc 4 — Nombre",
  "Proc 4 — Pasos",
  "Proc 4 — Tiempo/ciclo",
  "Proc 4 — Frecuencia",
  "Proc 4 — Herramienta",
  "Proc 4 — Dolor principal",
  "Proc 5 — Nombre",
  "Proc 5 — Pasos",
  "Proc 5 — Tiempo/ciclo",
  "Proc 5 — Frecuencia",
  "Proc 5 — Herramienta",
  "Proc 5 — Dolor principal",
  // Índice de digitalización
  "IDX — Facturación y cobros (1-5)",
  "IDX — Registro contable (1-5)",
  "IDX — Conciliaciones bancarias (1-5)",
  "IDX — Reportes gerenciales (1-5)",
  "IDX — Presupuesto y forecast (1-5)",
  "IDX — Almacenamiento y archivo (1-5)",
  "IDX — Comunicación financiera interna (1-5)",
  "IDX — Promedio general",
  // IA
  "Usa IA hoy",
  "Herramienta de IA",
  "Para qué la usa",
  "Frecuencia de uso IA",
  "Experiencia negativa con IA",
  "Detalle experiencia negativa",
  // Prioridades
  "Prioridades de transformación (top 3)",
  "Transformación prioritaria (libre)",
  // Resumen consolidado
  "Resumen consolidado"
];

// ── Punto de entrada principal ────────────────────────────────────
function doPost(e) {
  try {
    const raw = e.postData.contents || e.parameter.data || "{}";
    const data = JSON.parse(raw);
    escribirFila(data);
    return ContentService
      .createTextOutput(JSON.stringify({ status: "ok", mensaje: "Diagnóstico registrado correctamente." }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ status: "error", mensaje: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ── Necesario para pruebas desde el navegador ─────────────────────
function doGet(e) {
  return ContentService
    .createTextOutput(JSON.stringify({ status: "ok", mensaje: "El script está activo y listo para recibir diagnósticos." }))
    .setMimeType(ContentService.MimeType.JSON);
}

// ── Escribe la fila en el Sheet ───────────────────────────────────
function escribirFila(data) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sheet = ss.getSheetByName(SHEET_NAME);

  // Crear la hoja si no existe
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    // Encabezados
    sheet.appendRow(SHEET_HEADER);
    // Formato de encabezados
    const headerRange = sheet.getRange(1, 1, 1, SHEET_HEADER.length);
    headerRange.setBackground("#0D1F3C");
    headerRange.setFontColor("#FFFFFF");
    headerRange.setFontWeight("bold");
    headerRange.setFontSize(10);
    sheet.setFrozenRows(1);
    // Ancho de columnas clave
    sheet.setColumnWidth(1, 140);  // Fecha
    sheet.setColumnWidth(2, 200);  // Cámara
    sheet.setColumnWidth(SHEET_HEADER.length, 400); // Resumen consolidado
  }

  // Construir resumen consolidado
  const procs = data.procedimientos || [];
  const idx = data.indice || {};
  const idxVals = Object.values(idx).filter(v => v > 0);
  const promedio = idxVals.length
    ? (idxVals.reduce((a,b)=>a+b,0)/idxVals.length).toFixed(1)
    : "—";

  const resumen = buildResumen(data, promedio);

  // Armar la fila
  const fila = [
    new Date().toLocaleString("es-VE", { timeZone: "America/Caracas" }),
    data.camara || "",
    data.sector || "",
    data.asociados || "",
    data.personas || "",
    data.erp || "",
    data.suite || "",
    data.tenant || "",
    data.sharepoint || "",
    // Proc 1-5
    ...[0,1,2,3,4].flatMap(i => [
      procs[i]?.nombre    || "",
      procs[i]?.pasos     || "",
      procs[i]?.tiempo    || "",
      procs[i]?.frecuencia || "",
      procs[i]?.herramienta || "",
      procs[i]?.dolor     || ""
    ]),
    // Índice
    idx["Facturación y cobros"]              || 0,
    idx["Registro contable"]                 || 0,
    idx["Conciliaciones bancarias"]          || 0,
    idx["Reportes gerenciales"]              || 0,
    idx["Presupuesto y forecast"]            || 0,
    idx["Almacenamiento y archivo"]          || 0,
    idx["Comunicación financiera interna"]   || 0,
    promedio,
    // IA
    data.usaIA        || "",
    data.herramientaIA || "",
    data.paraQueIA    || "",
    data.frecuenciaIA || "",
    data.expNegativa  || "",
    data.detalleNeg   || "",
    // Prioridades
    (data.prioridades || []).join(" | "),
    data.transformacionLibre || "",
    // Resumen
    resumen
  ];

  sheet.appendRow(fila);

  // Formato de la nueva fila
  const lastRow = sheet.getLastRow();
  const rowRange = sheet.getRange(lastRow, 1, 1, SHEET_HEADER.length);
  rowRange.setFontSize(10);
  rowRange.setVerticalAlignment("top");
  rowRange.setWrap(true);

  // Color alternado por fila
  if (lastRow % 2 === 0) {
    rowRange.setBackground("#F4F7FB");
  }

  // Altura de fila automática
  sheet.setRowHeight(lastRow, 60);
}

// ── Construye el texto de resumen consolidado ─────────────────────
function buildResumen(data, promedio) {
  const procs = (data.procedimientos || []).filter(p => p.nombre);
  const prior = data.prioridades || [];

  let txt = "";
  txt += `CÁMARA: ${data.camara || "—"} | Sector: ${data.sector || "—"} | `;
  txt += `Agremiados: ${data.asociados || "—"} | Equipo financiero: ${data.personas || "—"} personas\n`;
  txt += `Suite: ${data.suite || "—"} | ERP: ${data.erp || "ninguno"} | `;
  txt += `Tenant M365: ${data.tenant || "—"} | SharePoint: ${data.sharepoint || "—"}\n\n`;

  txt += `PROCEDIMIENTOS MEDULARES (${procs.length}):\n`;
  procs.forEach((p, i) => {
    txt += `${i+1}. ${p.nombre} — ${p.pasos||"?"} pasos — ${p.tiempo||"?"}/ciclo — ${p.frecuencia||"?"} — Herramienta: ${p.herramienta||"?"} — Dolor: ${p.dolor||"—"}\n`;
  });

  txt += `\nÍNDICE DE DIGITALIZACIÓN — Promedio: ${promedio}/5\n`;
  const idx = data.indice || {};
  Object.entries(idx).forEach(([area, val]) => {
    const barras = "█".repeat(val) + "░".repeat(5-val);
    txt += `${area}: ${barras} ${val}/5\n`;
  });

  txt += `\nUSO DE IA: ${data.usaIA||"—"} | Herramienta: ${data.herramientaIA||"—"} | Frecuencia: ${data.frecuenciaIA||"—"}\n`;
  if (data.paraQueIA) txt += `Para qué: ${data.paraQueIA}\n`;

  txt += `\nPRIORIDADES (top 3):\n`;
  prior.forEach((p, i) => { txt += `${i+1}. ${p}\n`; });

  if (data.transformacionLibre) {
    txt += `\nTRANSFORMACIÓN PRIORITARIA:\n${data.transformacionLibre}`;
  }

  return txt;
}
