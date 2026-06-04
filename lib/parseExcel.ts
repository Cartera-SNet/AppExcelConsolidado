import ExcelJS from 'exceljs';

export interface ParsedRecord {
  institucion: string;
  nit: string;
  cantidad: number;
  valor: number;
  saldo: number;
  dpto: string;
  all_dates: string[];   // ISO strings of every radicacion date found
  min_fecha: Date | null;
  max_fecha: Date | null;
  archivo: string;
}

// ── Definitive IPS map ───────────────────────────────────────────────────────
// Source of truth: CONSOLIDADOS.xlsx names + Mapeo_IPS_Empresas.xlsx NITs/DPTOs
// Key = NIT digits only (no spaces, dots, dashes)
// nit = formatted as it appears in the CONSOLIDADOS output ("900267064 2")
// ────────────────────────────────────────────────────────────────────────────
const IPS_MAP: Record<string, { nombre: string; nit: string; dpto: string }> = {
  '9006005509': { nombre: 'INVERSIONES MÉDICAS BARU',                          nit: '900600550 9', dpto: 'BOLIVAR'         },
  '9009548004': { nombre: 'CENTRO MEDICO Y DE REHABILITACION BARU',            nit: '900954800 4', dpto: 'BOLIVAR'         },
  '9008270653': { nombre: 'CENTRO DE DIAGNOSTICO E IMAGENES BAHIA',            nit: '900827065 3', dpto: 'MAGDALENA'       },
  '9002670642': { nombre: 'INVERSIONES AZALUD S.A.S',                          nit: '900267064 2', dpto: 'MAGDALENA'       },
  '9008265097': { nombre: 'RUC MAGDALENA',                                     nit: '900826509 7', dpto: 'MAGDALENA'       },
  '9006577310': { nombre: 'CENTRO MEDICO Y DE REHABILITACION BAHIA',           nit: '900657731 0', dpto: 'MAGDALENA'       },
  '9005133065': { nombre: 'FUNDACIÓN MARIA REINA',                             nit: '900513306 5', dpto: 'SUCRE'           },
  '9002573336': { nombre: 'ODONTOTRANS',                                       nit: '900257333 6', dpto: 'VALLE DEL CAUCA' },
  '9010812818': { nombre: 'URGETRAUMA',                                        nit: '901081281 8', dpto: 'VALLE DEL CAUCA' },
  '9007924171': { nombre: 'RED DE URGENCIAS DE LA COSTA PACIFICA',             nit: '900792417 1', dpto: 'VALLE DEL CAUCA' },
  '9000027800': { nombre: 'FUNDACION CAMPBELL',                                nit: '900002780 0', dpto: 'ATLANTICO'       },
  '9006313616': { nombre: 'INVERSIONES MÉDICAS VALLE SALUD S.A.S',             nit: '900631361 6', dpto: 'VALLE DEL CAUCA' },
  '9008473829': { nombre: 'CENTRO MEDICO Y REHABILITACION VALLE SALUD S.A.S.', nit: '900847382 9', dpto: 'VALLE DEL CAUCA' },
  '9011497576': { nombre: 'CVO UNIDAD MÉDICA DE TRAUMA DEL VALLE S.A.S.',      nit: '901149757 6', dpto: 'VALLE DEL CAUCA' },
  '9009007541': { nombre: 'CLINICA VALLE SALUD SAN FERNANDO S.A.S',            nit: '900900754 1', dpto: 'VALLE DEL CAUCA' },
  '9004698828': { nombre: 'CENTRO MEDICO SERVISALUD INTEGRAL IPS S.A.S.',      nit: '900469882 8', dpto: 'VALLE DEL CAUCA' },
  '9015238689': { nombre: 'MOVID IPS S.A.S',                                   nit: '901523868 9', dpto: 'ATLANTICO'       },
  '9005585950': { nombre: 'FUNDACION MEDICA CAMPBELL',                         nit: '900558595 0', dpto: 'ATLANTICO'       },
  '8020243290': { nombre: 'RED DE URGENCIA DE LA COSTA',                       nit: '802024329 0', dpto: 'ATLANTICO'       },
};

// ── Helpers ──────────────────────────────────────────────────────────────────

function onlyDigits(s: string): string {
  return s.replace(/\D/g, '');
}

function lookupByNIT(rawNIT: string): { nombre: string; nit: string; dpto: string } | null {
  const d = onlyDigits(rawNIT);
  return IPS_MAP[d] ?? null;
}

/** Normalise text for label matching: remove accents, lowercase, collapse spaces */
function norm(v: unknown): string {
  if (v === null || v === undefined) return '';
  return String(v)
    .normalize('NFD').replace(/[\u0300-\u036f]/g, '')
    .replace(/\s+/g, ' ').trim().toLowerCase();
}

/**
 * Read a cell value safely, unwrapping formula objects { formula, result }
 * that ExcelJS returns when the cell contains a formula.
 */
function cv(ws: ExcelJS.Worksheet, row: number, col: number): unknown {
  try {
    const v = ws.getCell(row, col).value;
    if (v === null || v === undefined) return null;
    if (typeof v === 'object' && 'result' in (v as object)) {
      return (v as { result?: unknown }).result ?? null;
    }
    return v;
  } catch { return null; }
}

function numVal(ws: ExcelJS.Worksheet, row: number, col: number): number {
  const v = cv(ws, row, col);
  return typeof v === 'number' ? v : 0;
}

function strVal(ws: ExcelJS.Worksheet, row: number, col: number): string {
  const v = cv(ws, row, col);
  return v !== null ? String(v).trim() : '';
}

/**
 * Scan a region for a cell whose normalised text contains `needle`.
 * Returns { row, col } (1-based) or null.
 */
function findLabel(
  ws: ExcelJS.Worksheet,
  needle: string,
  maxRow = 15,
  maxCol = 10
): { row: number; col: number } | null {
  for (let r = 1; r <= maxRow; r++) {
    for (let c = 1; c <= maxCol; c++) {
      if (norm(cv(ws, r, c)).includes(needle)) return { row: r, col: c };
    }
  }
  return null;
}

/**
 * After finding a label cell, find the first positive number to the right
 * within `radius` columns, then one row below.
 */
function numNear(
  ws: ExcelJS.Worksheet,
  row: number,
  col: number,
  radius = 5
): number {
  for (let c = col + 1; c <= col + radius; c++) {
    const v = numVal(ws, row, c);
    if (v > 0) return v;
  }
  for (let r = row + 1; r <= row + 2; r++) {
    for (let c = col; c <= col + radius; c++) {
      const v = numVal(ws, r, c);
      if (v > 0) return v;
    }
  }
  return 0;
}

// ── FORMAT A: La Previsora / Seguros / standard format ───────────────────────
//
// Summary block (rows 2-5 typically):
//   col E: "Cantidad Facturas"   col G: value   (Campbell, Azalud)
//
// Data header row (identified by "Número de reclamo"):
//   col A: NIT | col B: Razón social | col D: Número reclamo
//   col E: Fecha de radicación | col F: Valor neto | col G: Valor saldo
//
function parseFormatA(ws: ExcelJS.Worksheet, filename: string): ParsedRecord {

  // ── 1. Summary totals: scan by label ─────────────────────────────────────
  const lblCant  = findLabel(ws, 'cantidad');
  const lblValor = findLabel(ws, 'valor cobrado');
  const lblSaldo = findLabel(ws, 'saldo requerido');

  let cantidad = lblCant  ? numNear(ws, lblCant.row,  lblCant.col)  : 0;
  let valor    = lblValor ? numNear(ws, lblValor.row, lblValor.col) : 0;
  let saldo    = lblSaldo ? numNear(ws, lblSaldo.row, lblSaldo.col) : 0;

  // ── 2. Data header row ────────────────────────────────────────────────────
  let headerRow = -1;
  const cols: Record<string, number> = {};

  outer:
  for (let r = 1; r <= 15; r++) {
    for (let c = 1; c <= 10; c++) {
      const t = norm(cv(ws, r, c));
      if (t.includes('numero de reclamo') || t.includes('fecha de radicacion')) {
        headerRow = r;
        break outer;
      }
    }
  }

  if (headerRow > 0) {
    for (let c = 1; c <= 10; c++) {
      const t = norm(cv(ws, headerRow, c));
      if (t.includes('nit del prestador'))   cols.nit        = c;
      if (t.includes('razon social'))        cols.institucion = c;
      if (t.includes('numero de reclamo'))   cols.reclamo    = c;
      if (t.includes('fecha de radicacion')) cols.fecha      = c;
      if (t.includes('valor neto'))          cols.valor      = c;
      if (t.includes('valor saldo'))         cols.saldo      = c;
    }
  }

  // ── 3. Data rows: NIT, name, dates, fallback totals ───────────────────────
  let rawNIT = '', institucionName = '';
  const dates: Date[] = [];
  let dCant = 0, dValor = 0, dSaldo = 0;

  if (headerRow > 0) {
    for (let r = headerRow + 1; r <= ws.rowCount; r++) {
      const reclamo = cols.reclamo ? cv(ws, r, cols.reclamo) : null;
      const fecha   = cols.fecha   ? cv(ws, r, cols.fecha)   : null;
      const vCell   = cols.valor   ? numVal(ws, r, cols.valor) : 0;
      const sCell   = cols.saldo   ? numVal(ws, r, cols.saldo) : 0;

      if (!reclamo && !fecha && !vCell && !sCell) continue;

      if (!rawNIT && cols.nit)        rawNIT          = strVal(ws, r, cols.nit);
      if (!institucionName && cols.institucion) institucionName = strVal(ws, r, cols.institucion);
      if (reclamo) dCant++;
      dValor += vCell;
      dSaldo += sCell;
      if (fecha instanceof Date) dates.push(fecha);
    }
  }

  // Prefer summary block values; fall back to row-sum
  if (cantidad === 0) cantidad = dCant;
  if (valor    === 0) valor    = dValor;
  if (saldo    === 0) saldo    = dSaldo;

  // ── 4. Resolve IPS identity from NIT ─────────────────────────────────────
  if (!rawNIT) rawNIT = filename;
  const ips = lookupByNIT(rawNIT);

  return {
    institucion: (ips?.nombre ?? institucionName) || filename,
    nit:         ips?.nit    ?? rawNIT,
    cantidad, valor, saldo,
    dpto:        ips?.dpto   ?? '',
    all_dates:   dates.map(d => d.toISOString()),
    min_fecha:   dates.length ? new Date(Math.min(...dates.map(d => d.getTime()))) : null,
    max_fecha:   dates.length ? new Date(Math.max(...dates.map(d => d.getTime()))) : null,
    archivo:     filename,
  };
}

// ── FORMAT B: "Estado de Cuenta" style ───────────────────────────────────────
//
// Row 1 col A: " ESTADO DE CUENTA ... NIT 900827065-3"
// Summary block:
//   [5,6]="Cantidad facturas"  [5,8]=93     (label col F, value col H)
//   [6,6]="Valor Cobrado"      [6,8]=269183800
//   [7,6]="Saldo Requerido..."  [7,8]=175556504
//
// Data header row 9:
//   col A: No. | col B: Nro factura | col E: Fecha Aviso | col F: Fecha Egreso
//   col G: Valor Cobrado | col H: Saldo Pendiente
//
function parseFormatB(ws: ExcelJS.Worksheet, filename: string): ParsedRecord {

  // ── 1. NIT from title ─────────────────────────────────────────────────────
  const title = strVal(ws, 1, 1);
  const nitMatch = title.match(/NIT\s*([\d.\-]+)/i);
  const rawNIT   = nitMatch ? nitMatch[1] : '';

  // ── 2. Summary totals by label scan ──────────────────────────────────────
  const lblCant  = findLabel(ws, 'cantidad facturas', 10, 10);
  const lblValor = findLabel(ws, 'valor cobrado',     10, 10);
  const lblSaldo = findLabel(ws, 'saldo requerido',   10, 10);

  let cantidad = lblCant  ? numNear(ws, lblCant.row,  lblCant.col)  : 0;
  let valor    = lblValor ? numNear(ws, lblValor.row, lblValor.col) : 0;
  let saldo    = lblSaldo ? numNear(ws, lblSaldo.row, lblSaldo.col) : 0;

  // ── 3. Data header row (has "Nro factura") ────────────────────────────────
  let headerRow = -1;
  let colFecha = 5, colValorD = 7, colSaldoD = 8;

  outer:
  for (let r = 1; r <= 15; r++) {
    for (let c = 1; c <= 8; c++) {
      if (norm(cv(ws, r, c)).includes('nro factura')) {
        headerRow = r;
        break outer;
      }
    }
  }

  if (headerRow > 0) {
    for (let c = 1; c <= 10; c++) {
      const t = norm(cv(ws, headerRow, c));
      if (t.includes('fecha aviso') || t.includes('fecha egreso')) colFecha   = c;
      if (t.includes('valor cobrado'))   colValorD = c;
      if (t.includes('saldo pendiente')) colSaldoD = c;
    }
  }

  // ── 4. Data rows: dates and fallback totals ───────────────────────────────
  const dates: Date[] = [];
  let dValor = 0, dSaldo = 0, dCant = 0;

  if (headerRow > 0) {
    for (let r = headerRow + 1; r <= ws.rowCount; r++) {
      const fecha  = cv(ws, r, colFecha);
      const vCell  = numVal(ws, r, colValorD);
      const sCell  = numVal(ws, r, colSaldoD);
      if (!fecha && !vCell && !sCell) continue;
      if (fecha instanceof Date) dates.push(fecha);
      if (vCell) { dValor += vCell; dCant++; }
      if (sCell) dSaldo += sCell;
    }
  }

  if (cantidad === 0) cantidad = dCant;
  if (valor    === 0) valor    = dValor;
  if (saldo    === 0) saldo    = dSaldo;

  // ── 5. Resolve IPS identity ───────────────────────────────────────────────
  const ips = lookupByNIT(rawNIT);

  return {
    institucion: ips?.nombre ?? filename,
    nit:         ips?.nit    ?? rawNIT,
    cantidad, valor, saldo,
    dpto:        ips?.dpto   ?? '',
    all_dates:   dates.map(d => d.toISOString()),
    min_fecha:   dates.length ? new Date(Math.min(...dates.map(d => d.getTime()))) : null,
    max_fecha:   dates.length ? new Date(Math.max(...dates.map(d => d.getTime()))) : null,
    archivo:     filename,
  };
}

// ── Format auto-detection ─────────────────────────────────────────────────────
function detectFormat(ws: ExcelJS.Worksheet): 'A' | 'B' {
  // Format B has "ESTADO DE CUENTA" in cell A1
  if (norm(cv(ws, 1, 1)).includes('estado de cuenta')) return 'B';
  // Format B also has "Nro factura" in its data header
  for (let r = 1; r <= 15; r++)
    for (let c = 1; c <= 8; c++)
      if (norm(cv(ws, r, c)).includes('nro factura')) return 'B';
  // Format A has "Número de reclamo" or "Fecha de radicación" in its header
  return 'A';
}

// ── Public API ────────────────────────────────────────────────────────────────
// eslint-disable-next-line @typescript-eslint/no-explicit-any
export async function parseExcelFile(buffer: any, filename: string): Promise<ParsedRecord> {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load(buffer);

  // Prefer the first non-"PENDIENTE" sheet
  let ws = wb.worksheets[0];
  for (const sheet of wb.worksheets) {
    if (!sheet.name.toLowerCase().includes('pendiente')) { ws = sheet; break; }
  }

  return detectFormat(ws) === 'B'
    ? parseFormatB(ws, filename)
    : parseFormatA(ws, filename);
}
