/**
 * AUTOMATIZACIÓN FLUJO MALLS - Parque Arauco
 */

const CONFIG = {
  SHEET_BBDD: 'BBDD',
  SHEET_PPTO: 'presupuesto',
  SHEET_PPT_DIA: 'ppto_por_día',
  SHEET_FLUJO_DIA: 'flujo_por_día',
  SHEET_RESUMEN: 'Resumen',

  COL: {
    DATE: 0, YEAR: 1, MALL: 2, ENTRANCE: 3,
    FLUJO: 4, SEMANA: 5, DIA: 6, MES: 7, FECHA_HOM: 8, AISLAR: 9
  },

  ROW_DIAS_START: 12,
};

// CONSTANTES GLOBALES DE ESTILO
const COLOR_HEADER_BG = '#1F4E78';
const COLOR_HEADER_TXT = '#FFFFFF';
const COLOR_VERDE_BG = '#E2EFDA';
const COLOR_VERDE_TXT = '#375623';
const COLOR_ROJO_BG = '#FCE4D6';
const COLOR_ROJO_TXT = '#C00000';

// ==================================================
// FUNCIÓN PRINCIPAL VINCULADA AL BOTÓN
// ==================================================
function actualizarTodo() {
  generarPresupuestoModulo();
  generarResumen();
  generarFlujoDia();
  generarAccesosModulo();
  generarTablasSecundariasYGrafico();
  aplicarEstiloVisual();
  generarInsights();
  crearOActualizarTablaDinamicaAccesos();

  SpreadsheetApp.getActive().toast('¡Todo actualizado con éxito!', 'Automatización', 3);
}

// ==================================================
// HELPERS COMPARTIDOS
// ==================================================
function parseDateHelper(val) {
  if (!val) return null;
  if (val instanceof Date) {
    return { year: val.getFullYear(), month: val.getMonth() + 1, day: val.getDate() };
  }
  const str = String(val).trim();
  const parts = str.split(/[-/]/);
  if (parts.length === 3) {
    if (parts[0].length === 4) {
      return { year: parseInt(parts[0], 10), month: parseInt(parts[1], 10), day: parseInt(parts[2], 10) };
    } else if (parts[2].length === 4) {
      return { year: parseInt(parts[2], 10), month: parseInt(parts[1], 10), day: parseInt(parts[0], 10) };
    }
  }
  return null;
}

function parseFlujoValHelper(val) {
  if (typeof val === 'number') return val;
  if (!val) return 0;
  let strVal = String(val).trim().replace(/\./g, '').replace(',', '.');
  return parseFloat(strVal) || 0;
}

/**
 * HELPER DE AISLAMIENTO:
 * Lee de la hoja BBDD el bloque N:P para construir un Set con la clave 'MALL_YYYY-MM-DD' de días a excluir.
 */
function obtenerSetDiasAislar(sheetBBDD) {
  const setAislar = new Set();
  const lastRow = sheetBBDD.getLastRow();
  if (lastRow < 2) return setAislar;

  const dataAislar = sheetBBDD.getRange(2, 14, lastRow - 1, 2).getValues(); // O:P (O=Mall/Prop, P=Fecha)

  dataAislar.forEach(r => {
    const mall = String(r[0]).trim().toUpperCase();
    const fechaVal = r[1];
    
    if (mall && fechaVal) {
      const parsed = parseDateHelper(fechaVal);
      if (parsed) {
        const mm = String(parsed.month).padStart(2, '0');
        const dd = String(parsed.day).padStart(2, '0');
        const clave = `${mall}_${parsed.year}-${mm}-${dd}`;
        setAislar.add(clave);
      }
    }
  });

  return setAislar;
}
/**
 * HELPER DE MESES:
 * Lee el valor de una celda y devuelve un arreglo numérico con los meses indicados (ej. "7, 8, 9" -> [7, 8, 9]).
 */
function obtenerListaMeses(valorCelda) {
  if (!valorCelda) return [];
  
  // Si la celda es un objeto Date por auto-interpretación de Sheets
  if (valorCelda instanceof Date) {
    return [valorCelda.getMonth() + 1];
  }
  
  const texto = String(valorCelda).trim();
  if (!texto) return [];

  // Si contiene comas (ej: "7, 8, 9")
  if (texto.includes(',')) {
    return texto.split(',')
      .map(m => Number(m.trim()))
      .filter(m => !isNaN(m) && m > 0 && m <= 12);
  }

  // Si contiene guion para rangos (ej: "7-9")
  if (texto.includes('-')) {
    const partes = texto.split('-').map(m => Number(m.trim()));
    if (partes.length === 2 && !isNaN(partes[0]) && !isNaN(partes[1])) {
      const lista = [];
      for (let m = partes[0]; m <= partes[1]; m++) lista.push(m);
      return lista;
    }
  }

  // Si es un solo número
  const num = Number(texto);
  return (isNaN(num) || num <= 0 || num > 12) ? [] : [num];
}
// ==================================================
// 1. MÓDULO PRESUPUESTO POR DÍA 
// ==================================================
function generarPresupuestoModulo() {
  const ss = SpreadsheetApp.getActive();
  const sheetPptDia = ss.getSheetByName(CONFIG.SHEET_PPT_DIA);
  const sheetPptoBase = ss.getSheetByName(CONFIG.SHEET_PPTO);
  const bbdd = ss.getSheetByName(CONFIG.SHEET_BBDD);

  if (!sheetPptDia || !sheetPptoBase || !bbdd) return;

  const targetYear = parseInt(sheetPptDia.getRange("B2").getValue(), 10);
  const targetMonth = parseInt(sheetPptDia.getRange("B3").getValue(), 10);

  if (isNaN(targetYear) || isNaN(targetMonth)) return;

  const setAislar = obtenerSetDiasAislar(bbdd);

  // Obtener días máximos reportados en la BBDD por Mall
  const dbData = bbdd.getDataRange().getValues();
  dbData.shift();

  const maxDiaCalendarioPorMall = {};
  dbData.forEach(r => {
    const monthVal = Number(r[CONFIG.COL.MES]);
    const yearVal = Number(r[CONFIG.COL.YEAR]);
    const mall = String(r[CONFIG.COL.MALL]).trim().toUpperCase();
    const rawDate = r[CONFIG.COL.DATE];
    const parsed = parseDateHelper(rawDate);

    if (monthVal === targetMonth && yearVal === targetYear && parsed && parsed.day) {
      const mm = String(parsed.month).padStart(2, '0');
      const dd = String(parsed.day).padStart(2, '0');
      const claveAislar = `${mall}_${parsed.year}-${mm}-${dd}`;

      if (!setAislar.has(claveAislar)) {
        if (!maxDiaCalendarioPorMall[mall] || parsed.day > maxDiaCalendarioPorMall[mall]) {
          maxDiaCalendarioPorMall[mall] = parsed.day;
        }
      }
    }
  });

  const maxDayToFill = new Date(targetYear, targetMonth, 0).getDate();
  const lastRowBase = sheetPptoBase.getLastRow();
  if (lastRowBase < 2) return;

  // Lectura de las 3 columnas planas en 'presupuesto' (A: Fecha, B: Mall, C: Flujo)
  const baseData = sheetPptoBase.getRange(2, 1, lastRowBase - 1, 3).getValues();
  const mapDataDia = {};

  for (let i = 0; i < baseData.length; i++) {
    const dateVal = baseData[i][0];
    const mall = String(baseData[i][1]).trim().toUpperCase();
    const flujoFinal = parseFlujoValHelper(baseData[i][2]);

    if (!mall || !dateVal) continue;

    const parsedDate = parseDateHelper(dateVal);
    if (parsedDate) {
      const { year: yearVal, month: monthVal, day: dayVal } = parsedDate;
      const mm = String(monthVal).padStart(2, '0');
      const dd = String(dayVal).padStart(2, '0');
      const claveAislar = `${mall}_${yearVal}-${mm}-${dd}`;

      if (!setAislar.has(claveAislar)) {
        if (yearVal === targetYear && monthVal === targetMonth && dayVal <= maxDayToFill) {
          if (!mapDataDia[mall]) mapDataDia[mall] = {};
          mapDataDia[mall][dayVal] = (mapDataDia[mall][dayVal] || 0) + flujoFinal;
        }
      }
    }
  }

  // Generación de la tabla ppto_por_día
  const lastRowPpt = sheetPptDia.getLastRow();
  let mallList = [];
  if (lastRowPpt >= 7) {
    mallList = sheetPptDia.getRange(7, 1, lastRowPpt - 6, 1).getValues().map(r => String(r[0]).trim().toUpperCase());
  }
  const filteredMalls = mallList.filter(m => m !== "" && m !== "TOTAL GENERAL");

  const maxCols = sheetPptDia.getMaxColumns();
  if (maxCols > 2 && lastRowPpt >= 6) {
    sheetPptDia.getRange(6, 2, Math.max(lastRowPpt - 5, 25), maxCols - 1).clearContent();
  }

  const dateHeaders = [];
  for (let d = 1; d <= maxDayToFill; d++) {
    const formattedDay = String(d).padStart(2, '0');
    const formattedMonth = String(targetMonth).padStart(2, '0');
    dateHeaders.push(`${formattedDay}-${formattedMonth}-${targetYear}`);
  }
  dateHeaders.push("Total hasta hoy");
  dateHeaders.push("Días contados"); // <- NUEVA COLUMNA DE ENCABEZADO

  if (dateHeaders.length > 0) {
    sheetPptDia.getRange(6, 2, 1, dateHeaders.length)
      .setValues([dateHeaders])
      .setFontWeight("bold")
      .setFontColor(COLOR_HEADER_TXT)
      .setBackground(COLOR_HEADER_BG)
      .setHorizontalAlignment("center");

    const outputMatrix = [];
    const dailyTotals = new Array(dateHeaders.length).fill(0);

    filteredMalls.forEach(mall => {
      const rowValues = [];
      let totalHastaHoyMall = 0;
      const maxDiaMall = maxDiaCalendarioPorMall[mall] || 0;

      for (let d = 1; d <= maxDayToFill; d++) {
        const val = (mapDataDia[mall] && mapDataDia[mall][d]) ? Math.round(mapDataDia[mall][d]) : 0;
        rowValues.push(val);
        dailyTotals[d - 1] += val;

        if (d <= maxDiaMall) {
          totalHastaHoyMall += val;
        }
      }

      rowValues.push(totalHastaHoyMall);
      rowValues.push(maxDiaMall); // <- INSERTA LA CANTIDAD DE DÍAS CONTADOS (ej: 30 o 31)

      dailyTotals[dateHeaders.length - 2] += totalHastaHoyMall; // Suma acumulada para "Total hasta hoy"
      // La columna "Días contados" no se suma en la fila total (se deja en 0/vacía)

      outputMatrix.push(rowValues);
    });

    if (outputMatrix.length > 0) {
      sheetPptDia.getRange(7, 2, outputMatrix.length, dateHeaders.length)
        .setValues(outputMatrix)
        .setNumberFormat("#,##0");

      const totalRowIndex = 7 + filteredMalls.length;
      sheetPptDia.getRange(totalRowIndex, 1).setValue("Total general").setFontWeight("bold");
      
      // Ajustar fila de totales
      const totalRowValues = [...dailyTotals];
      totalRowValues[dateHeaders.length - 1] = ""; // Dejar vacía la celda del total para "Días contados"

      sheetPptDia.getRange(totalRowIndex, 2, 1, dateHeaders.length)
        .setValues([totalRowValues])
        .setFontWeight("bold")
        .setNumberFormat("#,##0");

      const colTotalIdx = 1 + dateHeaders.length - 1; // Columna "Total hasta hoy"
      sheetPptDia.getRange(6, colTotalIdx, outputMatrix.length + 2, 1).setFontWeight("bold");
    }
  }

  SpreadsheetApp.flush();
}

// ==================================================
// 2. MÓDULO RESUMEN 
// ==================================================
function generarResumen() {
  const ss = SpreadsheetApp.getActive();
  const bbdd = ss.getSheetByName(CONFIG.SHEET_BBDD);
  const pptoBase = ss.getSheetByName(CONFIG.SHEET_PPTO);
  const out = ss.getSheetByName(CONFIG.SHEET_RESUMEN);

  if (!bbdd || !pptoBase || !out) return;

  const mes = Number(out.getRange('B4').getValue());
  if (!mes) return;

  const setAislar = obtenerSetDiasAislar(bbdd);

  const values = out.getDataRange().getValues();
  let headerRow = -1;
  for (let i = 0; i < values.length; i++) {
    if (String(values[i][0]).trim().toLowerCase() === 'id_property') { headerRow = i; break; }
  }
  if (headerRow === -1) return;

  out.getRange(headerRow + 1, 1, 1, 7)
    .setFontWeight("bold")
    .setFontColor(COLOR_HEADER_TXT)
    .setBackground(COLOR_HEADER_BG)
    .setHorizontalAlignment("center");

  const malls = [];
  let i = headerRow + 1;
  while (i < values.length) {
    const cell = String(values[i][0]).trim();
    if (!cell || cell.toLowerCase() === 'total general') break;
    malls.push(cell);
    i++;
  }
  const totalRowIdx = i;

  const data = bbdd.getDataRange().getValues();
  data.shift();

  const years = [...new Set(data.map(r => Number(r[CONFIG.COL.YEAR])))].sort((a, b) => a - b);
  const anioAnt = years[years.length - 2] || 2025;
  const anioAct = years[years.length - 1] || 2026;

  const actAcc = {};
  const diasValidosPorMall = {};
  const maxDiaCalendarioPorMall = {};

  data.forEach(r => {
    if (Number(r[CONFIG.COL.MES]) !== mes) return;
    if (Number(r[CONFIG.COL.YEAR]) !== anioAct) return;

    const mall = String(r[CONFIG.COL.MALL]).trim().toUpperCase();
    const flujo = Number(r[CONFIG.COL.FLUJO]) || 0;
    const rawDate = r[CONFIG.COL.DATE];
    const parsed = parseDateHelper(rawDate);
    if (!parsed || !parsed.day) return;

    const mm = String(parsed.month).padStart(2, '0');
    const dd = String(parsed.day).padStart(2, '0');
    const claveAislar = `${mall}_${parsed.year}-${mm}-${dd}`;

    if (setAislar.has(claveAislar)) return;

    if (!diasValidosPorMall[mall]) diasValidosPorMall[mall] = new Set();
    diasValidosPorMall[mall].add(parsed.day);

    if (!maxDiaCalendarioPorMall[mall] || parsed.day > maxDiaCalendarioPorMall[mall]) {
      maxDiaCalendarioPorMall[mall] = parsed.day;
    }

    actAcc[mall] = (actAcc[mall] || 0) + flujo;
  });

  const antAcc = {};
  data.forEach(r => {
    if (Number(r[CONFIG.COL.MES]) !== mes) return;
    if (Number(r[CONFIG.COL.YEAR]) !== anioAnt) return;

    const mall = String(r[CONFIG.COL.MALL]).trim().toUpperCase();
    const rawDate = r[CONFIG.COL.DATE];
    const parsed = parseDateHelper(rawDate);
    if (!parsed || !parsed.day) return;

    const mm = String(parsed.month).padStart(2, '0');
    const dd = String(parsed.day).padStart(2, '0');
    const claveAislar = `${mall}_${parsed.year}-${mm}-${dd}`;

    if (setAislar.has(claveAislar)) return;

    const maxDia = maxDiaCalendarioPorMall[mall] || 0;
    if (parsed.day > maxDia) return;

    antAcc[mall] = (antAcc[mall] || 0) + (Number(r[CONFIG.COL.FLUJO]) || 0);
  });

  // NUEVA LÓGICA: Sumar Presupuesto acumulado hasta el día reportado según la nueva estructura A:C
  const pptoByMall = {};
  const lastRowPpto = pptoBase.getLastRow();
  if (lastRowPpto >= 2) {
    const pptoRaw = pptoBase.getRange(2, 1, lastRowPpto - 1, 3).getValues();
    pptoRaw.forEach(r => {
      const dateVal = r[0];
      const mall = String(r[1]).trim().toUpperCase();
      const flujoFinal = parseFlujoValHelper(r[2]);
      if (!mall || !dateVal) return;

      const parsedDate = parseDateHelper(dateVal);
      if (!parsedDate || parsedDate.year !== anioAct || parsedDate.month !== mes) return;

      const mm = String(parsedDate.month).padStart(2, '0');
      const dd = String(parsedDate.day).padStart(2, '0');
      const claveAislar = `${mall}_${parsedDate.year}-${mm}-${dd}`;

      if (setAislar.has(claveAislar)) return;

      const maxDia = maxDiaCalendarioPorMall[mall] || 0;
      if (parsedDate.day > maxDia) return;

      pptoByMall[mall] = (pptoByMall[mall] || 0) + flujoFinal;
    });
  }

  out.getRange(headerRow + 1, 2, 1, 2).setValues([[anioAnt, anioAct]]);

  limpiarFormatosCondicionalesColumna(out, `D${headerRow + 2}:D${headerRow + 2 + malls.length}`);
  limpiarFormatosCondicionalesColumna(out, `F${headerRow + 2}:F${headerRow + 2 + malls.length}`);

  const totalFilas = malls.length + 1;
  out.getRange(headerRow + 2, 2, totalFilas, 6).setBackground(null).setFontColor(null);

  let sumAnt = 0, sumAct = 0, sumPpto = 0;
  malls.forEach((mall, idx) => {
    const mallKey = mall.toUpperCase();
    const rowNum = headerRow + 2 + idx;
    const vAnt = antAcc[mallKey] || 0;
    const vAct = actAcc[mallKey] || 0;
    const varAA = vAnt ? (vAct - vAnt) / vAnt : '';
    const pptoVal = pptoByMall[mallKey] || 0;
    const varPpto = pptoVal ? (vAct - pptoVal) / pptoVal : '';
    const dias = maxDiaCalendarioPorMall[mallKey] || 0;

    out.getRange(rowNum, 2, 1, 6).setValues([[vAnt, vAct, varAA, pptoVal, varPpto, dias]]);
    out.getRange(rowNum, 2, 1, 2).setNumberFormat('#,##0');
    out.getRange(rowNum, 4).setNumberFormat('0.0%');
    out.getRange(rowNum, 5).setNumberFormat('#,##0');
    out.getRange(rowNum, 6).setNumberFormat('0.0%');

    sumAnt += vAnt; sumAct += vAct; sumPpto += pptoVal;
  });

  if (totalRowIdx < values.length && String(values[totalRowIdx][0]).trim().toLowerCase() === 'total general') {
    const rowNum = totalRowIdx + 1;
    const varAA = sumAnt ? (sumAct - sumAnt) / sumAnt : '';
    const varPpto = sumPpto ? (sumAct - sumPpto) / sumPpto : '';

    out.getRange(rowNum, 2, 1, 6).setValues([[sumAnt, sumAct, varAA, sumPpto, varPpto, '']]);
    out.getRange(rowNum, 2, 1, 2).setNumberFormat('#,##0');
    out.getRange(rowNum, 4).setNumberFormat('0.0%');
    out.getRange(rowNum, 5).setNumberFormat('#,##0');
    out.getRange(rowNum, 6).setNumberFormat('0.0%');
  }

  const rangeVarAA = out.getRange(headerRow + 2, 4, totalFilas, 1);
  const rangeVarPpto = out.getRange(headerRow + 2, 6, totalFilas, 1);
  aplicarFormatoCondicionalVariacion(out, rangeVarAA);
  aplicarFormatoCondicionalVariacion(out, rangeVarPpto);

  // Llama a la función de accesos enviándole los datos de día máximo y años
  generarAccesosModulo();

  SpreadsheetApp.flush();
}

// ==================================================
// 3. MÓDULO FLUJO POR DÍA
// ==================================================
function generarFlujoDia() {
  const ss = SpreadsheetApp.getActive();
  const sheetFlujo = ss.getSheetByName(CONFIG.SHEET_FLUJO_DIA);
  const bbdd = ss.getSheetByName(CONFIG.SHEET_BBDD);

  if (!sheetFlujo || !bbdd) return;

  const targetYear = parseInt(sheetFlujo.getRange("B2").getValue(), 10);
  const targetMonth = parseInt(sheetFlujo.getRange("B3").getValue(), 10);

  if (isNaN(targetYear) || isNaN(targetMonth)) return;

  const setAislar = obtenerSetDiasAislar(bbdd);

  const data = bbdd.getDataRange().getValues();
  data.shift();

  const mapFlujo = {};
  const maxDiaCalendarioPorMall = {};

  data.forEach(r => {
    const year = Number(r[CONFIG.COL.YEAR]);
    const month = Number(r[CONFIG.COL.MES]);
    const mall = String(r[CONFIG.COL.MALL]).trim().toUpperCase();
    const rawDate = r[CONFIG.COL.DATE];
    const flujo = Number(r[CONFIG.COL.FLUJO]) || 0;

    if (year === targetYear && month === targetMonth && mall) {
      const parsed = parseDateHelper(rawDate);
      if (parsed && parsed.day) {
        const mm = String(parsed.month).padStart(2, '0');
        const dd = String(parsed.day).padStart(2, '0');
        const claveAislar = `${mall}_${parsed.year}-${mm}-${dd}`;

        if (!setAislar.has(claveAislar)) {
          if (!mapFlujo[mall]) mapFlujo[mall] = {};
          mapFlujo[mall][parsed.day] = (mapFlujo[mall][parsed.day] || 0) + flujo;

          if (!maxDiaCalendarioPorMall[mall] || parsed.day > maxDiaCalendarioPorMall[mall]) {
            maxDiaCalendarioPorMall[mall] = parsed.day;
          }
        }
      }
    }
  });

  const maxDayToFill = new Date(targetYear, targetMonth, 0).getDate();
  const lastRow = sheetFlujo.getLastRow();
  let mallList = [];

  if (lastRow >= 7) {
    mallList = sheetFlujo.getRange(7, 1, lastRow - 6, 1).getValues().map(r => String(r[0]).trim().toUpperCase());
  }
  const filteredMalls = mallList.filter(m => m !== "" && m !== "TOTAL GENERAL");

  if (filteredMalls.length === 0) return;

  const maxCols = sheetFlujo.getMaxColumns();
  if (maxCols > 1 && lastRow >= 6) {
    const rangeToClear = sheetFlujo.getRange(6, 2, Math.max(lastRow - 5, 25), maxCols - 1);
    rangeToClear.clearContent();
    rangeToClear.setBackground(null);
  }

  const dateHeaders = [];
  for (let d = 1; d <= maxDayToFill; d++) {
    const formattedDay = String(d).padStart(2, '0');
    const formattedMonth = String(targetMonth).padStart(2, '0');
    dateHeaders.push(`${formattedDay}-${formattedMonth}-${targetYear}`);
  }
  dateHeaders.push("Total hasta hoy");

  if (dateHeaders.length > 0) {
    const totalColumns = dateHeaders.length;

    const headerRange = sheetFlujo.getRange(6, 2, 1, totalColumns);
    headerRange
      .setValues([dateHeaders])
      .setFontWeight("bold")
      .setFontColor(COLOR_HEADER_TXT)
      .setBackground(COLOR_HEADER_BG)
      .setHorizontalAlignment("center")
      .setVerticalAlignment("middle");

    const outputMatrix = [];
    const dailyTotals = new Array(totalColumns).fill(0);

    filteredMalls.forEach(mall => {
      const rowValues = [];
      let totalHastaHoyMall = 0;
      const maxDiaMall = maxDiaCalendarioPorMall[mall] || 0;

      for (let d = 1; d <= maxDayToFill; d++) {
        const val = (mapFlujo[mall] && mapFlujo[mall][d] !== undefined) ? Math.round(mapFlujo[mall][d]) : 0;
        rowValues.push(val);
        dailyTotals[d - 1] += val;

        if (d <= maxDiaMall) {
          totalHastaHoyMall += val;
        }
      }

      rowValues.push(totalHastaHoyMall);
      dailyTotals[totalColumns - 1] += totalHastaHoyMall;

      outputMatrix.push(rowValues);
    });

    if (outputMatrix.length > 0) {
      const dataRange = sheetFlujo.getRange(7, 2, outputMatrix.length, totalColumns);
      dataRange
        .setValues(outputMatrix)
        .setNumberFormat("#,##0")
        .setHorizontalAlignment("right");

      const totalRowIndex = 7 + filteredMalls.length;
      sheetFlujo.getRange(totalRowIndex, 1).setValue("Total general").setFontWeight("bold");
      
      const totalRowRange = sheetFlujo.getRange(totalRowIndex, 2, 1, dailyTotals.length);
      totalRowRange
        .setValues([dailyTotals])
        .setFontWeight("bold")
        .setNumberFormat("#,##0")
        .setHorizontalAlignment("right");

      const colTotalIdx = 1 + totalColumns;
      sheetFlujo.getRange(6, colTotalIdx, outputMatrix.length + 2, 1)
        .setFontWeight("bold");
    }
  }

  SpreadsheetApp.flush();
}

// ==================================================
// 4. MÓDULO ACCESOS, TABLAS Y GRÁFICOS
// ==================================================
function generarAccesosModulo(maxDiaCalendarioPorMall = {}, anioAnt = 2025, anioAct = 2026) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const bbdd = ss.getSheetByName(CONFIG.SHEET_BBDD);
  const sheetResumen = ss.getSheetByName(CONFIG.SHEET_RESUMEN);
  if (!bbdd || !sheetResumen) return;

  const mesFiltro = Number(sheetResumen.getRange('B4').getValue());
  if (!mesFiltro || isNaN(mesFiltro)) return;

  const setAislar = obtenerSetDiasAislar(bbdd);
  const data = bbdd.getDataRange().getValues();
  data.shift();

  // 1. Obtener los años reales presentes en la BBDD
  const years = [...new Set(data.map(r => Number(r[CONFIG.COL.YEAR])))].filter(y => !isNaN(y) && y > 0).sort((a, b) => a - b);
  if (years.length >= 2) {
    anioAnt = years[years.length - 2];
    anioAct = years[years.length - 1];
  }

  // 2. RE-CALCULAR maxDia de forma directa y garantizada para el año actual (2026)
  const maxDiasPropios = {};
  data.forEach(r => {
    if (Number(r[CONFIG.COL.MES]) !== mesFiltro) return;
    if (Number(r[CONFIG.COL.YEAR]) !== anioAct) return;

    const mall = String(r[CONFIG.COL.MALL]).trim().toUpperCase();
    const parsed = parseDateHelper(r[CONFIG.COL.DATE]);
    if (!parsed || !parsed.day) return;

    const mm = String(parsed.month).padStart(2, '0');
    const dd = String(parsed.day).padStart(2, '0');
    const claveAislar = `${mall}_${parsed.year}-${mm}-${dd}`;

    if (setAislar.has(claveAislar)) return;

    if (!maxDiasPropios[mall] || parsed.day > maxDiasPropios[mall]) {
      maxDiasPropios[mall] = parsed.day;
    }
  });

  // 3. Mapeo inicial de Accesos
  const accesosMap = {};
  const ordenMalls = [];
  const ordenAccesos = {};

  data.forEach(r => {
    const mall = String(r[CONFIG.COL.MALL]).trim();
    let acceso = String(r[CONFIG.COL.ENTRANCE]).replace(/[\r\n]+/g, " ").replace(/\s+/g, " ").trim();
    if (!mall || !acceso) return;

    if (!accesosMap[mall]) {
      accesosMap[mall] = {};
      ordenMalls.push(mall);
      ordenAccesos[mall] = [];
    }

    if (!accesosMap[mall][acceso]) {
      accesosMap[mall][acceso] = { [anioAnt]: 0, [anioAct]: 0 };
      ordenAccesos[mall].push(acceso);
    }
  });

  // 4. Acumulación aplicando CORTE RIGUROSO por maxDia
  data.forEach(r => {
    const monthVal = Number(r[CONFIG.COL.MES]);
    if (monthVal !== mesFiltro) return;

    const mall = String(r[CONFIG.COL.MALL]).trim();
    const mallKey = mall.toUpperCase();
    let acceso = String(r[CONFIG.COL.ENTRANCE]).replace(/[\r\n]+/g, " ").replace(/\s+/g, " ").trim();
    
    const parsedDate = parseDateHelper(r[CONFIG.COL.DATE]);
    if (!parsedDate || !parsedDate.day) return;

    const mm = String(parsedDate.month).padStart(2, '0');
    const dd = String(parsedDate.day).padStart(2, '0');
    const claveAislar = `${mallKey}_${parsedDate.year}-${mm}-${dd}`;

    if (setAislar.has(claveAislar)) return;

    // AQUÍ APLICAMOS EL CORTE: Si el día de 2025 o 2026 supera el máximo día que existe registrado en 2026, SE IGNORA.
    const limiteDia = maxDiasPropios[mallKey] || 0;
    if (limiteDia > 0 && parsedDate.day > limiteDia) return;

    const yearVal = parsedDate.year;
    const flujo = parseFlujoValHelper(r[CONFIG.COL.FLUJO]);

    if (accesosMap[mall] && accesosMap[mall][acceso]) {
      if (yearVal === anioAnt || yearVal === anioAct) {
        accesosMap[mall][acceso][yearVal] += flujo;
      }
    }
  });

  // 5. Salida de datos
  const output = [];
  ordenMalls.forEach(mall => {
    ordenAccesos[mall].forEach(acceso => {
      const vAnt = Math.round(accesosMap[mall][acceso][anioAnt] || 0);
      const vAct = Math.round(accesosMap[mall][acceso][anioAct] || 0);
      
      let variacion = 0;
      if (vAnt > 0) {
        variacion = (vAct - vAnt) / vAnt;
      } else if (vAnt === 0 && vAct > 0) {
        variacion = 1;
      }

      output.push([mall, acceso, vAnt, vAct, variacion]);
    });
  });

  const startCol = 10; // Columna J
  const maxRows = Math.max(sheetResumen.getLastRow(), 2);
  sheetResumen.getRange(1, startCol, maxRows, 5).clearContent();

  limpiarFormatosCondicionalesColumna(sheetResumen, 'N2:N' + (output.length + 1));

  // Encabezados
  sheetResumen.getRange(1, startCol, 1, 5)
    .setValues([["Mall", "Acceso / Entrada", String(anioAnt), String(anioAct), "Variación"]])
    .setFontWeight("bold")
    .setFontColor(COLOR_HEADER_TXT)
    .setBackground(COLOR_HEADER_BG)
    .setHorizontalAlignment("center");

  if (output.length > 0) {
    const rangeData = sheetResumen.getRange(2, startCol, output.length, 5);
    rangeData.setValues(output);
    
    sheetResumen.getRange(2, startCol + 2, output.length, 2).setNumberFormat("#,##0");
    const rangeVariacion = sheetResumen.getRange(2, startCol + 4, output.length, 1);
    rangeVariacion.setNumberFormat("0.0%");

    aplicarFormatoCondicionalVariacion(sheetResumen, rangeVariacion);
  }

  SpreadsheetApp.flush();
}

// ==================================================
// 5. TABLA MES Y DÍAS + GRÁFICO (CON HERENCIA DE FILTROS)
// ==================================================
function generarTablasSecundariasYGrafico() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const bbdd = ss.getSheetByName(CONFIG.SHEET_BBDD);
  const sheet = ss.getSheetByName(CONFIG.SHEET_RESUMEN);
  if (!bbdd || !sheet) return;

  const setAislar = obtenerSetDiasAislar(bbdd);

  limpiarFormatosCondicionalesColumna(sheet, 'T13:T25');
  limpiarFormatosCondicionalesColumna(sheet, 'Z13:Z45');

  const data = bbdd.getDataRange().getValues();
  data.shift();

  // --- FILTROS DE MES Y HERENCIA AUTOMÁTICA ---
  const entradaMes = String(sheet.getRange('R5').getValue()).trim().toLowerCase();
  const mallMes = String(sheet.getRange('R6').getValue()).trim().toLowerCase();

  // NUEVA LÓGICA CON HELPER:
  const mesesTablaDias = obtenerListaMeses(sheet.getRange('X6').getValue());
  const setMesesDias = new Set(mesesTablaDias.length > 0 ? mesesTablaDias : [8]);

  const entradaDias = entradaMes;
  const mallDias = mallMes;

  const resumenMes = {};
  for (let m = 1; m <= 12; m++) resumenMes[m] = { 2025: 0, 2026: 0 };

  // Estructuras para guardar flujos por clave "MES-DIA" (ej: "8-15", "9-01")
  const dias2025 = {};
  const dias2026 = {};

  data.forEach(r => {
    const mallVal = String(r[CONFIG.COL.MALL]).trim().toLowerCase();
    const entradaVal = String(r[CONFIG.COL.ENTRANCE]).replace(/[\r\n]+/g, " ").replace(/\s+/g, " ").trim().toLowerCase();
    const monthVal = Number(r[CONFIG.COL.MES]);

    const parsedDate = parseDateHelper(r[CONFIG.COL.DATE]);
    if (!parsedDate) return;

    const mm = String(parsedDate.month).padStart(2, '0');
    const dd = String(parsedDate.day).padStart(2, '0');
    const claveAislar = `${mallVal.toUpperCase()}_${parsedDate.year}-${mm}-${dd}`;

    if (setAislar.has(claveAislar)) return;

    const yearVal = parsedDate.year;
    const flujo = parseFlujoValHelper(r[CONFIG.COL.FLUJO]);

    // A) TABLA MES
const matchMallMes = !mallMes || mallVal === mallMes;
const matchEntradaMes = !entradaMes || entradaMes === "todos" || entradaVal === entradaMes;

if (matchMallMes && matchEntradaMes) {
  if ((yearVal === 2025 || yearVal === 2026) && monthVal >= 1 && monthVal <= 12) {
    resumenMes[monthVal][yearVal] += flujo;
  }
}

    // B) TABLA DÍAS
    const matchMesDias = setMesesDias.has(monthVal);
    const matchMallDias = !mallDias || mallVal === mallDias;
    const matchEntradaDias = !entradaDias || entradaMes === "todos" || entradaVal === entradaDias;

    // Para 2025, permitimos también cargar datos si el mes es el mes siguiente (mNum + 1)
    // para cubrir los días de fin de mes que homologan al inicio del mes siguiente.
    const esMesContiguo2025 = Array.from(setMesesDias).some(mNum => {
      const mesSiguiente = (mNum % 12) + 1;
      return monthVal === mNum || monthVal === mesSiguiente;
    });

    if (matchMallDias && matchEntradaDias) {
      if (parsedDate.day >= 1 && parsedDate.day <= 31) {
        const claveDia = `${parsedDate.month}-${parsedDate.day}`;

        if (yearVal === 2025 && esMesContiguo2025) {
          dias2025[claveDia] = (dias2025[claveDia] || 0) + flujo;
        } else if (yearVal === 2026 && matchMesDias) {
          dias2026[claveDia] = (dias2026[claveDia] || 0) + flujo;
        }
      }
    }
  });

  // FORMATO ENCABEZADOS TABLAS SECUNDARIAS
  sheet.getRange('Q12:T12')
    .setFontWeight("bold")
    .setFontColor(COLOR_HEADER_TXT)
    .setBackground(COLOR_HEADER_BG)
    .setHorizontalAlignment("center");

  sheet.getRange('W12:Z12')
    .setFontWeight("bold")
    .setFontColor(COLOR_HEADER_TXT)
    .setBackground(COLOR_HEADER_BG)
    .setHorizontalAlignment("center");

  // 1. IMPRIMIR TABLA MES (Q:T)
  const outMes = [];
  for (let m = 1; m <= 12; m++) {
    const v2025 = Math.round(resumenMes[m][2025]);
    const v2026 = Math.round(resumenMes[m][2026]);
    const varPct = v2025 ? (v2026 - v2025) / v2025 : (v2026 > 0 ? 1 : 0);
    outMes.push([m, v2025, v2026, varPct]);
  }

  const rangeMesTotal = sheet.getRange('Q13:T25');
  rangeMesTotal.clearContent().setBackground(null);
  sheet.getRange(13, 17, outMes.length, 4).setValues(outMes);
  sheet.getRange(13, 18, outMes.length, 2).setNumberFormat("#,##0");

  const rangeVarMes = sheet.getRange(13, 20, outMes.length, 1);
  rangeVarMes.setNumberFormat("0.0%");
  aplicarFormatoCondicionalVariacion(sheet, rangeVarMes);

  // 2. IMPRIMIR TABLA DÍAS MULTI-MES (W:Z)
  const outDias = [];
  const listaMeses = Array.from(setMesesDias).sort((a, b) => a - b);

  listaMeses.forEach(mNum => {
    const mmStr = String(mNum).padStart(2, '0');
    
    // Obtiene los días REALES del mes (ej: 30 para Septiembre, 31 para Octubre)
    const diasEnElMes = new Date(2026, mNum, 0).getDate();
    
    for (let d = 1; d <= diasEnElMes; d++) { // <--- CORREGIDO AQUÍ
      const ddStr = String(d).padStart(2, '0');
      const fecha2026Str = `${ddStr}-${mmStr}-2026`;
      const claveDiaActual = `${mNum}-${d}`;
      
      // Homologación por día de la semana (resta 364 días)
      const fechaActualObj = new Date(2026, mNum - 1, d);
      const fechaHomologada2025 = new Date(fechaActualObj.getTime() - (364 * 24 * 60 * 60 * 1000));

      const mesHomologado = fechaHomologada2025.getMonth() + 1;
      const diaHomologado = fechaHomologada2025.getDate();
      const claveHomologada = `${mesHomologado}-${diaHomologado}`;

      const v2025 = Math.round(dias2025[claveHomologada] || 0);
      const tieneDatos2026 = (dias2026[claveDiaActual] !== undefined && dias2026[claveDiaActual] !== null);
      const v2026 = tieneDatos2026 ? Math.round(dias2026[claveDiaActual]) : "";

      let varPct = -1;
      if (v2026 !== "") {
        varPct = v2025 ? (v2026 - v2025) / v2025 : (v2026 > 0 ? 1 : 0);
      }

      outDias.push([fecha2026Str, v2025, v2026, varPct]);
    }
  });

  // Limpiar rango extendido para soportar más de 31 días (ej: 60 días para 2 meses)
  const rangeDiasTotal = sheet.getRange('W13:Z100');
  rangeDiasTotal.clearContent().setBackground(null);

  if (outDias.length > 0) {
    sheet.getRange(13, 23, outDias.length, 4).setValues(outDias);
    sheet.getRange(13, 24, outDias.length, 2).setNumberFormat("#,##0");

    const rangeVarDias = sheet.getRange(13, 26, outDias.length, 1);
    rangeVarDias.setNumberFormat("0.0%");
    aplicarFormatoCondicionalVariacion(sheet, rangeVarDias);
  }

  // 3. REUBICAR GRÁFICO
  const charts = sheet.getCharts();
  charts.forEach(c => sheet.removeChart(c));

  if (outDias.length > 0) {
    const rangeGrafico = sheet.getRange(12, 23, outDias.length + 1, 3);
    const chart = sheet.newChart()
      .setChartType(Charts.ChartType.LINE)
      .addRange(rangeGrafico)
      .setPosition(3, 28, 0, 0)
      .setOption('title', 'Tendencia de Flujo Diario (2025 vs 2026)')
      .setOption('width', 600)
      .setOption('height', 350)
      .build();

    sheet.insertChart(chart);
  }

  SpreadsheetApp.flush();
}

/**
 * APLICA EL FORMATO CONDICIONAL UNIFICADO A CUALQUIER RANGO DE VARIACIÓN (%)
 */
function aplicarFormatoCondicionalVariacion(sheet, range) {
  if (!sheet || !range) return;

  const rules = sheet.getConditionalFormatRules();

  const reglaVerde = SpreadsheetApp.newConditionalFormatRule()
    .whenNumberGreaterThanOrEqualTo(0)
    .setBackground(COLOR_VERDE_BG)
    .setFontColor(COLOR_VERDE_TXT)
    .setRanges([range])
    .build();

  const reglaRoja = SpreadsheetApp.newConditionalFormatRule()
    .whenNumberLessThan(0)
    .setBackground(COLOR_ROJO_BG)
    .setFontColor(COLOR_ROJO_TXT)
    .setRanges([range])
    .build();

  rules.push(reglaVerde);
  rules.push(reglaRoja);

  sheet.setConditionalFormatRules(rules);
}

function limpiarFormatosCondicionalesColumna(sheet, rangoA1) {
  const range = sheet.getRange(rangoA1);
  const targetCol = range.getColumn();
  const rules = sheet.getConditionalFormatRules();
  
  const filteredRules = rules.filter(rule => {
    return !rule.getRanges().some(r => {
      const col = r.getColumn();
      const lastCol = col + r.getNumColumns() - 1;
      return targetCol >= col && targetCol <= lastCol;
    });
  });
  
  sheet.setConditionalFormatRules(filteredRules);
}

function aplicarEstiloVisual() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetResumen = ss.getSheetByName(CONFIG.SHEET_RESUMEN);
  
  if (sheetResumen) {
    // 1. Agregamos 'B3' a las celdas que se borran y limpian
    const celdasEliminadas = ['B3', 'B5', 'R7', 'X5', 'X7', 'X8', 'X9'];
    celdasEliminadas.forEach(celda => {
      sheetResumen.getRange(celda)
        .clearContent()
        .clearFormat()
        .clearNote();
    });

    // 2. Mantenemos el color del filtro solo para B4 y los demás
    const COLOR_FILTER_BG = '#FFF2CC';
    const COLOR_FILTER_TXT = '#7F6000';
    const rangosFiltros = ['B4', 'R5:R6', 'X6'];

    rangosFiltros.forEach(rangoA1 => {
      const rng = sheetResumen.getRange(rangoA1);
      rng.setBackground(COLOR_FILTER_BG)
         .setFontColor(COLOR_FILTER_TXT)
         .setFontWeight('bold');
    });
  }
}

// ==================================================
// 6. FUNCIÓN GENERAR INSIGHTS (CON BARRERA DE RECENCIA - 2 MESES)
// ==================================================
function generarInsights() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const bbdd = ss.getSheetByName(CONFIG.SHEET_BBDD);
  const sheetResumen = ss.getSheetByName(CONFIG.SHEET_RESUMEN);
  let sheetInsights = ss.getSheetByName('Insights');

  if (!bbdd || !sheetResumen) return;
  if (!sheetInsights) sheetInsights = ss.insertSheet('Insights');

  const mesFiltro = Number(sheetResumen.getRange('B4').getValue());
  const dataBBDD = bbdd.getDataRange().getValues();
  dataBBDD.shift();

  const years = [...new Set(dataBBDD.map(r => Number(r[CONFIG.COL.YEAR])))].sort((a, b) => a - b);
  const anioAct = years[years.length - 1] || 2026;

  if (!mesFiltro || isNaN(mesFiltro)) return;

  const setAislar = obtenerSetDiasAislar(bbdd);

  // Definir recencia para considerar contadores vigentes (mes actual + 2 anteriores)
  const mesesRecientes = [];
  for (let m = 0; m <= 2; m++) {
    let mVal = mesFiltro - m;
    if (mVal > 0) mesesRecientes.push(mVal);
  }

  // Identificar el período base para comparar el promedio esperado
  const mesAnteriorVal = mesFiltro > 1 ? mesFiltro - 1 : 12;
  const anioMesAnterior = mesFiltro > 1 ? anioAct : anioAct - 1;

  // 1. IDENTIFICAR ACCESOS ACTIVOS Y CALCULAR PROMEDIO HISTÓRICO ESPERADO
  const accesosActivos = new Set();
  const sumaMesAnt = {};
  const diasMesAnt = {};
  const mapaPromedioHistorico = {};

  dataBBDD.forEach(r => {
    const yearVal = Number(r[CONFIG.COL.YEAR]);
    const monthVal = Number(r[CONFIG.COL.MES]);
    const flujo = parseFlujoValHelper(r[CONFIG.COL.FLUJO]);
    const mall = String(r[CONFIG.COL.MALL]).trim();
    let acceso = String(r[CONFIG.COL.ENTRANCE]).replace(/[\r\n]+/g, " ").replace(/\s+/g, " ").trim();

    if (!mall || !acceso) return;
    const llave = `${mall} - ${acceso}`;

    if (flujo > 0) {
      if ((yearVal === anioAct && mesesRecientes.includes(monthVal)) || 
          (yearVal === (anioAct - 1) && monthVal === mesFiltro)) {
        accesosActivos.add(llave);
      }
    }

    // Acumular datos del mes anterior (o año pasado) como base de comparación
    if ((yearVal === anioMesAnterior && monthVal === mesAnteriorVal) || 
        (yearVal === (anioAct - 1) && monthVal === mesFiltro)) {
      if (flujo > 0) {
        sumaMesAnt[llave] = (sumaMesAnt[llave] || 0) + flujo;
        diasMesAnt[llave] = (diasMesAnt[llave] || 0) + 1;
      }
    }
  });

  // Calcular la media diaria histórica esperada
  Object.keys(sumaMesAnt).forEach(llave => {
    if (diasMesAnt[llave] > 0) {
      mapaPromedioHistorico[llave] = sumaMesAnt[llave] / diasMesAnt[llave];
    }
  });

  // 2. PROCESAR DATOS DEL MES ACTUAL A EVALUAR
  let maxDiaMes = 0;
  const mapaUltimoDiaAcceso = {};
  const mapaFlujoDiarioPorAcceso = {};

  dataBBDD.forEach(r => {
    const yearVal = Number(r[CONFIG.COL.YEAR]);
    const monthVal = Number(r[CONFIG.COL.MES]);

    if (yearVal !== anioAct || monthVal !== mesFiltro) return;

    const mall = String(r[CONFIG.COL.MALL]).trim();
    let acceso = String(r[CONFIG.COL.ENTRANCE]).replace(/[\r\n]+/g, " ").replace(/\s+/g, " ").trim();
    if (!mall || !acceso) return;

    const llaveAcceso = `${mall} - ${acceso}`;
    if (!accesosActivos.has(llaveAcceso)) return;

    const parsedDate = parseDateHelper(r[CONFIG.COL.DATE]);
    if (!parsedDate) return;

    const mm = String(parsedDate.month).padStart(2, '0');
    const dd = String(parsedDate.day).padStart(2, '0');
    const claveAislar = `${mall.toUpperCase()}_${parsedDate.year}-${mm}-${dd}`;

    if (setAislar.has(claveAislar)) return;

    const flujo = parseFlujoValHelper(r[CONFIG.COL.FLUJO]);

    if (!mapaFlujoDiarioPorAcceso[llaveAcceso]) {
      mapaFlujoDiarioPorAcceso[llaveAcceso] = {};
    }
    mapaFlujoDiarioPorAcceso[llaveAcceso][parsedDate.day] = (mapaFlujoDiarioPorAcceso[llaveAcceso][parsedDate.day] || 0) + flujo;

    if (parsedDate.day > maxDiaMes) {
      maxDiaMes = parsedDate.day;
    }

    if (!mapaUltimoDiaAcceso[llaveAcceso] || parsedDate.day > mapaUltimoDiaAcceso[llaveAcceso]) {
      mapaUltimoDiaAcceso[llaveAcceso] = parsedDate.day;
    }
  });

  // LIMPIEZA HOJA
  sheetInsights.clear();
  sheetInsights.getRange('A1:Z100').setFontWeight('normal').setBackground(null).setFontColor(null);

  sheetInsights.getRange('A1').setValue('DIAGNÓSTICO Y ALERTAS OPERATIVAS').setFontWeight('bold').setFontSize(14);
  sheetInsights.getRange('A2').setValue(`Período analizado: ${mesFiltro}/${anioAct}`).setFontSize(10);

  // ENCABEZADOS
  sheetInsights.getRange('A4').setValue('🔴 CONTADORES DESCONECTADOS / SIN DATOS EN EL MES')
    .setFontWeight('bold').setFontColor('#C00000');
  sheetInsights.getRange('A5:C5')
    .setValues([['Acceso / Entrada', 'Detalle', 'Días Desfasados / Estado']])
    .setFontWeight('bold').setBackground('#FCE4D6').setFontColor('#C00000');

  sheetInsights.getRange('E4').setValue('⚠️ ACCESOS CON ANOMALÍAS O FLUJOS BAJOS')
    .setFontWeight('bold').setFontColor('#B25900');
  sheetInsights.getRange('E5:G5')
    .setValues([['Acceso / Entrada', 'Flujo Último Día / Promedio Actual', 'Diagnóstico']])
    .setFontWeight('bold').setBackground('#FFF2CC').setFontColor('#7F6000');

  sheetInsights.getRange('I4').setValue('🟢 ACCESOS CON PICOS ANORMALMENTE ALTOS')
    .setFontWeight('bold').setFontColor('#375623');
  sheetInsights.getRange('I5:K5')
    .setValues([['Acceso / Entrada', 'Flujo Último Día', 'Variación vs Promedio']])
    .setFontWeight('bold').setBackground('#E2EFDA').setFontColor('#375623');

  const desfasados = [];
  const bajos = [];
  const altos = [];

  // 3. CLASIFICACIÓN DE ACCESOS
  accesosActivos.forEach(llave => {
    const ultimoDia = mapaUltimoDiaAcceso[llave];
    const flujos = mapaFlujoDiarioPorAcceso[llave];

    if (!ultimoDia) {
      desfasados.push([
        llave,
        `Sin ningún registro de datos en el mes ${mesFiltro}`,
        `${maxDiaMes > 0 ? maxDiaMes : 'Todo el'} día(s) sin datos`
      ]);
    } 
    else if (ultimoDia < maxDiaMes) {
      const diasDesfasados = maxDiaMes - ultimoDia;
      desfasados.push([
        llave,
        `Sin datos desde el día ${ultimoDia}`,
        `${diasDesfasados} día(s) desfasado`
      ]);
    } 
    else {
      const diasConDatos = Object.keys(flujos).map(Number);
      const totalDias = diasConDatos.length;
      if (totalDias === 0) return;

      const valoresFlujo = diasConDatos.map(d => flujos[d]);
      const sumaTotal = valoresFlujo.reduce((acc, v) => acc + v, 0);
      const promedioMesActual = sumaTotal / totalDias;
      const flujoUltimoDia = flujos[maxDiaMes] || 0;

      // Promedio histórico esperado (si no existe, usa la media del mes actual)
      const promedioHistorico = mapaPromedioHistorico[llave] || promedioMesActual;

      // AQUÍ SE AGREGA LA DETECCIÓN DE DÍAS INTERMITENTES CRÍTICOS
      // -------------------------------------------------------------
      const diasAnormalmenteBajos = diasConDatos.filter(d => {
        const val = flujos[d];
        return promedioHistorico > 500 && val < (promedioHistorico * 0.20);
      });

      // EVALUACIÓN DE REGLAS PARA ALERTAS AMARILLAS ⚠️
      
      // 1. Si el último día reportó cero
      if (flujoUltimoDia === 0) {
        bajos.push([llave, 0, 'Reporta flujo cero en el último día']);
      } 
      // 2. Si hubo uno o varios días intermitentes con caída severa (<20% del histórico)
      else if (diasAnormalmenteBajos.length > 0) {
        bajos.push([
          llave, 
          `${diasAnormalmenteBajos.length} día(s) crítico(s)`, 
          `Caída puntual (<20% del histórico ~${Math.round(promedioHistorico)}) en día(s): ${diasAnormalmenteBajos.join(', ')}`
        ]);
      }
      // 3. Si todo el mes en promedio viene muy caído vs su histórico (ej: Rampa N2)
      else if (promedioHistorico > 500 && promedioMesActual < (promedioHistorico * 0.20)) {
        const pctCaida = Math.round((1 - (promedioMesActual / promedioHistorico)) * 100);
        bajos.push([
          llave, 
          `Prom. ${Math.round(promedioMesActual)}/día`, 
          `Caída crítica de -${pctCaida}% vs histórico (~${Math.round(promedioHistorico)}/día)`
        ]);
      }
      // 4. Si solo el último día registró caída inusual vs su histórico
      else if (promedioHistorico > 100 && flujoUltimoDia < (promedioHistorico * 0.30)) {
        const varPct = Math.round(((flujoUltimoDia - promedioHistorico) / promedioHistorico) * 100);
        bajos.push([llave, flujoUltimoDia, `Flujo en último día inusualmente bajo (${varPct}% vs prom. histórico)`]);
      }

      // EVALUACIÓN DE PICOS ALTOS 🟢
      if (promedioHistorico <= 100 && flujoUltimoDia >= 500) {
        altos.push([llave, flujoUltimoDia, `Pico inusual: de histórico ~${Math.round(promedioHistorico)} subió a ${flujoUltimoDia}`]);
      }
      else if (promedioHistorico > 200 && flujoUltimoDia > (promedioHistorico * 1.50)) {
        const varPct = Math.round(((flujoUltimoDia - promedioHistorico) / promedioHistorico) * 100);
        altos.push([llave, flujoUltimoDia, `+${varPct}% sobre el promedio histórico`]);
      }
    }  
  });

  // 4. IMPRIMIR DATOS
  if (desfasados.length > 0) {
    sheetInsights.getRange(6, 1, desfasados.length, 3).setValues(desfasados);
  } else {
    sheetInsights.getRange('A6').setValue('Sin novedades');
  }

  if (bajos.length > 0) {
    sheetInsights.getRange(6, 5, bajos.length, 3).setValues(bajos);
  } else {
    sheetInsights.getRange('E6').setValue('Sin novedades');
  }

  if (altos.length > 0) {
    sheetInsights.getRange(6, 9, altos.length, 3).setValues(altos);
  } else {
    sheetInsights.getRange('I6').setValue('Sin novedades');
  }

  sheetInsights.autoResizeColumns(1, 11);
  SpreadsheetApp.flush();
}

// ==================================================
// 7. CREAR/ACTUALIZAR TABLA DINÁMICA 
// ==================================================
function crearOActualizarTablaDinamicaAccesos() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const bbddSheet = ss.getSheetByName(CONFIG.SHEET_BBDD);

  if (!bbddSheet) {
    SpreadsheetApp.getUi().alert("No se encontró la hoja de BBDD especificada en CONFIG.");
    return;
  }

  let sheetPivot = ss.getSheetByName('Vista_Diaria_Accesos');
  if (!sheetPivot) {
    sheetPivot = ss.insertSheet('Vista_Diaria_Accesos');
  } else {
    // Si ya existe un filtro activo en la hoja, lo removemos antes de limpiar para evitar errores
    if (sheetPivot.getFilter()) {
      sheetPivot.getFilter().remove();
    }
    sheetPivot.clear();
  }

  const sourceRange = bbddSheet.getDataRange();

  // 1. Crear la tabla dinámica directamente en la celda A1
  const pivotTable = sheetPivot.getRange('A4').createPivotTable(sourceRange);

  // Filas: Mall y Entrada (desactivando encabezados vacíos extra)
  pivotTable.addRowGroup(CONFIG.COL.MALL + 1).showTotals(true);
  pivotTable.addRowGroup(CONFIG.COL.ENTRANCE + 1).showTotals(false);

  // Columnas: Fecha
  pivotTable.addColumnGroup(CONFIG.COL.DATE + 1);

  // Valores: Suma de Flujo
  pivotTable.addPivotValue(CONFIG.COL.FLUJO + 1, SpreadsheetApp.PivotTableSummarizeFunction.SUM);

  // Filtro general por año 2026 sobre la BBDD
  try {
    const colFechaIndex = CONFIG.COL.DATE + 1;
    const filterCriteria = SpreadsheetApp.newFilterCriteria()
      .whenTextContains('2026')
      .build();
    pivotTable.addFilter(colFechaIndex, filterCriteria);
  } catch (e) {
    try {
      const colFechaIndex = CONFIG.COL.DATE + 1;
      const filterCriteria = SpreadsheetApp.newFilterCriteria()
        .whenDateBetween(new Date(2026, 0, 1), new Date(2026, 11, 31, 23, 59, 59))
        .build();
      pivotTable.addFilter(colFechaIndex, filterCriteria);
    } catch (err) {
      console.log("No se pudo aplicar el filtro de fecha en la Pivot: " + err.toString());
    }
  }

  SpreadsheetApp.flush();

  // 2. CORRECCIÓN DE POSICIÓN DE FILA:
  // Si la primera fila quedó totalmente vacía tras renderizar la Pivot, la eliminamos
  const primeraFilaValores = sheetPivot.getRange(1, 1, 1, sheetPivot.getLastColumn()).getValues()[0];
  const estaVacia = primeraFilaValores.every(celda => celda === "" || celda === null);

  if (estaVacia && sheetPivot.getLastRow() > 1) {
    sheetPivot.deleteRow(1);
  }

  SpreadsheetApp.flush();
}
// ==================================================
// Función doPost
// ==================================================
 function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);
    var ss = SpreadsheetApp.getActiveSpreadsheet();

    // 1. ESCRIBIR EN HOJA BBDD (Sin la columna id_platform, de A a J)
    var sheetBBDD = ss.getSheetByName("BBDD") || ss.insertSheet("BBDD");
    sheetBBDD.clearContents();

    var datosMatriz = data.hoja_principal;

    if (datosMatriz && datosMatriz.length > 0) {
      var numFilasPeticion = datosMatriz.length;
      var numColsPeticion = datosMatriz[0].length;
      var numFilasActuales = sheetBBDD.getMaxRows();

      // Expandir filas si el volumen de datos lo requiere
      if (numFilasActuales < numFilasPeticion) {
        sheetBBDD.insertRowsAfter(numFilasActuales, numFilasPeticion - numFilasActuales);
      }

      sheetBBDD.getRange(1, 1, numFilasPeticion, numColsPeticion).setValues(datosMatriz);
    }

    // 2. ESCRIBIR EN HOJA Estimaciones_SOM
    var sheetSOM = ss.getSheetByName("Estimaciones_SOM") || ss.insertSheet("Estimaciones_SOM");
    sheetSOM.clearContents();

    var datosSOM = data.hoja_som;

    if (datosSOM && datosSOM.length > 0) {
      var numFilasSOM = datosSOM.length;
      var numColsSOM = datosSOM[0].length;
      var numFilasActualesSOM = sheetSOM.getMaxRows();

      if (numFilasActualesSOM < numFilasSOM) {
        sheetSOM.insertRowsAfter(numFilasActualesSOM, numFilasSOM - numFilasActualesSOM);
      }

      sheetSOM.getRange(1, 1, numFilasSOM, numColsSOM).setValues(datosSOM);
    }

    // 3. EJECUTAR RECÁLCULO DE TABLEROS
    actualizarTodo();

    return ContentService.createTextOutput("OK").setMimeType(ContentService.MimeType.TEXT);

  } catch (error) {
    return ContentService.createTextOutput("Error: " + error.toString()).setMimeType(ContentService.MimeType.TEXT);
  }
}
