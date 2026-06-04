import ExcelJS from 'exceljs';

export interface ParsedRecord {
  institucion: string;
  nit: string;
  cantidad: number;
  valor: number;
  saldo: number;
  dpto: string;
  tiene_2026: boolean;
  min_fecha: Date | null;
  max_fecha: Date | null;
  archivo: string;
}

// ── IPS lookup ───────────────────────────────────────────────────────────────
const IPS_MAP: Record<string, { nombre: string; dpto: string }> = {
  '9008270653': { nombre: 'CENTRO DE DIAGNOSTICO E IMAGENES BAHIA',     dpto: 'MAGDALENA' },
  '9000027800': { nombre: 'FUNDACION CAMPBELL',                          dpto: 'ATLANTICO' },
  '9002670642': { nombre: 'INVERSIONES AZALUD S.A.S',                    dpto: 'MAGDALENA' },
  '9006005509': { nombre: 'INVERSIONES MÉDICAS BARU',                    dpto: 'BOLIVAR' },
  '9009548004': { nombre: 'CENTRO MEDICO Y DE REHABILITACION BARU',      dpto: 'BOLIVAR' },
  '9008265097': { nombre: 'RUC MAGDALENA',                               dpto: 'MAGDALENA' },
  '9006577310': { nombre: 'CENTRO MEDICO Y DE REHABILITACION BAHIA',     dpto: 'MAGDALENA' },
  '9005133065': { nombre: 'FUNDACIÓN MARIA REINA',                       dpto: 'SUCRE' },
  '9002573336': { nombre: 'ODONTOTRANS',                                 dpto: 'VALLE DEL CAUCA' },
  '9010812818': { nombre: 'URGETRAUMA',                                  dpto: 'VALLE DEL CAUCA' },
  '9007924171': { nombre: 'RED DE URGENCIAS DE LA COSTA PACIFICA',       dpto: 'VALLE DEL CAUCA' },
  '9006313616': { nombre: 'INVERSIONES MÉDICAS VALLE SALUD S.A.S',       dpto: 'VALLE DEL CAUCA' },
  '9008473829': { nombre: 'CENTRO MEDICO Y REHABILITACION VALLE SALUD',  dpto: 'VALLE DEL CAUCA' },
  '9011497576': { nombre: 'CVO UNIDAD MÉDICA DE TRAUMA DEL VALLE',       dpto: 'VALLE DEL CAUCA' },
  '9009007541': { nombre: 'CLINICA VALLE SALUD SAN FERNANDO',            dpto: 'VALLE DEL CAUCA' },
  '9004698828': { nombre: 'CENTRO MEDICO SERVISALUD INTEGRAL IPS',       dpto: 'VALLE DEL CAUCA' },
  '9015238689': { nombre: 'MOVID IPS S.A.S',                             dpto: 'ATLANTICO' },
  '9005585950': { nombre: 'FUNDACION MEDICA CAMPBELL',                   dpto: 'ATLANTICO' },
  '8020243290': { nombre: 'RED DE URGENCIA DE LA COSTA',                 dpto: 'ATLANTICO' },
};

function nitDigits(nit: string): string {
  return nit.replace(/\D/g, '');
}

function lookupIPS(nit: string): { nombre: string; dpto: string } | null {
  const d = nitDigits(nit);
  if (IPS_MAP[d]) return IPS_MAP[d];
  for (const key of Object.keys(IPS_MAP)) {
    if (nitDigits(key) === d) return IPS_MAP[key];
  }
  return null;
}

function formatNIT(raw: string): string {
  const d = nitDigits(raw);
  if (d.length >= 9) return `${d.slice(0, -1)} ${d.slice(-1)}`;
  return raw.trim();
}

function norm(v: unknown): string {
  if (v === null || v === undefined) return '';
  return String(v)
    .normalize('NFD').replace(/[\u0300-\u036f]/g, '')
    .replace(/\s+/g, ' ').trim().toLowerCase();
}

// Safe cell read: returns the raw value at (row, col), never throws.
// If the cell contains a formula, returns the cached result.
function cellVal(ws: ExcelJS.Worksheet, row: number, col: number): unknown {
  try {
    const v = ws.getCell(row, col).value;
    if (v === null || v === undefined) return null;
    if (typeof v === "object" && "result" in (v as object)) {
      return (v as { result?: unknown }).result ?? null;
    }
    return v;
  } catch { return null; }
}

function numAt(ws: ExcelJS.Worksheet, row: number, col: number): number | null {
  const v = cellVal(ws, row, col);
  return typeof v === "number" ? v : null;
}

// Scan up to maxRow × maxCol looking for a cell whose text includes `needle`.
// Returns { row, col } or null.
function findCell(
  ws: ExcelJS.Worksheet,
  needle: string,
  maxRow = 15,
  maxCol = 10
): { row: number; col: number } | null {
  for (let r = 1; r <= maxRow; r++) {
    for (let c = 1; c <= maxCol; c++) {
      if (norm(cellVal(ws, r, c)).includes(needle)) return { row: r, col: c };
    }
  }
  return null;
}

// After finding a label cell, search right and down for the first numeric value
function findNumericNear(
  ws: ExcelJS.Worksheet,
  labelRow: number,
  labelCol: number,
  searchRadius = 4
): number {
  // Search same row to the right first
  for (let c = labelCol + 1; c <= labelCol + searchRadius; c++) {
    const v = numAt(ws, labelRow, c);
    if (v !== null && v > 0) return v;
  }
  // Search rows below same column area
  for (let r = labelRow + 1; r <= labelRow + 2; r++) {
    for (let c = labelCol; c <= labelCol + searchRadius; c++) {
      const v = numAt(ws, r, c);
      if (v !== null && v > 0) return v;
    }
  }
  return 0;
}

// ── FORMAT A: La Previsora / Seguros style ───────────────────────────────────
// Summary block rows 2-4, data header row ~6 with "Número de reclamo"
function parseFormatA(ws: ExcelJS.Worksheet, filename: string): ParsedRecord {
  // 1) Find summary values by locating label cells directly
  let cantidad = 0, valor = 0, saldo = 0;

  const lblCantidad = findCell(ws, 'cantidad', 10, 8);
  const lblValor    = findCell(ws, 'valor cobrado', 10, 8);
  const lblSaldo    = findCell(ws, 'saldo requerido', 10, 8);

  if (lblCantidad) cantidad = findNumericNear(ws, lblCantidad.row, lblCantidad.col);
  if (lblValor)    valor    = findNumericNear(ws, lblValor.row,    lblValor.col);
  if (lblSaldo)    saldo    = findNumericNear(ws, lblSaldo.row,    lblSaldo.col);

  // 2) Find data header row (has "numero de reclamo" or "fecha de radicacion")
  let headerRow = -1;
  const cols: Record<string, number> = {};

  for (let r = 1; r <= 15; r++) {
    for (let c = 1; c <= 8; c++) {
      const t = norm(cellVal(ws, r, c));
      if (t.includes('numero de reclamo') || t.includes('fecha de radicacion')) {
        headerRow = r;
        break;
      }
    }
    if (headerRow > 0) break;
  }

  if (headerRow > 0) {
    // Map column positions from header row
    for (let c = 1; c <= 10; c++) {
      const t = norm(cellVal(ws, headerRow, c));
      if (t.includes('nit del prestador'))    cols.nit = c;
      if (t.includes('razon social'))         cols.institucion = c;
      if (t.includes('numero de reclamo'))    cols.reclamo = c;
      if (t.includes('fecha de radicacion'))  cols.fecha = c;
      if (t.includes('valor neto'))           cols.valor = c;
      if (t.includes('valor saldo'))          cols.saldo = c;
    }
  }

  // 3) Read data rows to get NIT, institution, dates and fallback totals
  let nitRaw = '', institucionName = '';
  const dates: Date[] = [];
  let dataCantidad = 0, dataValor = 0, dataSaldo = 0;

  if (headerRow > 0) {
    for (let r = headerRow + 1; r <= ws.rowCount; r++) {
      const reclamoVal = cols.reclamo ? cellVal(ws, r, cols.reclamo) : null;
      const fechaVal   = cols.fecha   ? cellVal(ws, r, cols.fecha)   : null;
      const valorCell  = cols.valor   ? cellVal(ws, r, cols.valor)   : null;
      const saldoCell  = cols.saldo   ? cellVal(ws, r, cols.saldo)   : null;
      const nitCell    = cols.nit     ? cellVal(ws, r, cols.nit)     : null;
      const instCell   = cols.institucion ? cellVal(ws, r, cols.institucion) : null;

      if (!reclamoVal && !fechaVal && !valorCell && !saldoCell) continue;

      if (!nitRaw && nitCell)          nitRaw          = String(nitCell);
      if (!institucionName && instCell) institucionName = String(instCell);
      if (reclamoVal)                  dataCantidad++;
      if (typeof valorCell === 'number') dataValor += valorCell;
      if (typeof saldoCell === 'number') dataSaldo += saldoCell;
      if (fechaVal instanceof Date)    dates.push(fechaVal);
    }
  }

  // Use summary totals if found, otherwise fall back to data row totals
  if (cantidad === 0) cantidad = dataCantidad;
  if (valor === 0)    valor    = dataValor;
  if (saldo === 0)    saldo    = dataSaldo;

  if (!nitRaw) nitRaw = filename;
  const ipsInfo = lookupIPS(nitRaw);

  return {
    institucion: (ipsInfo?.nombre ?? institucionName) || filename,
    nit:         formatNIT(nitRaw),
    cantidad, valor, saldo,
    dpto:        ipsInfo?.dpto ?? '',
    tiene_2026:  dates.some(d => d.getFullYear() === 2026),
    min_fecha:   dates.length ? new Date(Math.min(...dates.map(d => d.getTime()))) : null,
    max_fecha:   dates.length ? new Date(Math.max(...dates.map(d => d.getTime()))) : null,
    archivo:     filename,
  };
}

// ── FORMAT B: "Estado de Cuenta" style (CENTRO DE DIAGNOSTICO) ──────────────
// Title row 1 has NIT. Summary rows 5-7: labels in col F (6), values in col H (8).
function parseFormatB(ws: ExcelJS.Worksheet, filename: string): ParsedRecord {
  // Extract NIT from title
  let nitRaw = '';
  const titleVal = String(cellVal(ws, 1, 1) ?? '');
  const nitMatch = titleVal.match(/NIT\s*([\d.\-]+)/i);
  if (nitMatch) nitRaw = nitMatch[1];

  // Read summary by locating labels anywhere in first 10 rows
  let cantidad = 0, valor = 0, saldo = 0;

  const lblCantidad = findCell(ws, 'cantidad facturas', 10, 10);
  const lblValor    = findCell(ws, 'valor cobrado',     10, 10);
  const lblSaldo    = findCell(ws, 'saldo requerido',   10, 10);

  if (lblCantidad) cantidad = findNumericNear(ws, lblCantidad.row, lblCantidad.col);
  if (lblValor)    valor    = findNumericNear(ws, lblValor.row,    lblValor.col);
  if (lblSaldo)    saldo    = findNumericNear(ws, lblSaldo.row,    lblSaldo.col);

  // Find data header (has "Nro factura" or "Valor Cobrado" in the detail table)
  let headerRow = -1;
  let colFecha = 5, colValorD = 7, colSaldoD = 8;

  for (let r = 1; r <= 15; r++) {
    for (let c = 1; c <= 8; c++) {
      if (norm(cellVal(ws, r, c)).includes('nro factura')) {
        headerRow = r;
        break;
      }
    }
    if (headerRow > 0) break;
  }

  if (headerRow > 0) {
    for (let c = 1; c <= 10; c++) {
      const t = norm(cellVal(ws, headerRow, c));
      if (t.includes('fecha aviso') || t.includes('fecha egreso')) colFecha   = c;
      if (t.includes('valor cobrado'))   colValorD = c;
      if (t.includes('saldo pendiente')) colSaldoD = c;
    }
  }

  // Collect dates and fallback totals from data rows
  const dates: Date[] = [];
  let dataValor = 0, dataSaldo = 0, dataCantidad = 0;

  if (headerRow > 0) {
    for (let r = headerRow + 1; r <= ws.rowCount; r++) {
      const fechaVal  = cellVal(ws, r, colFecha);
      const valorCell = numAt(ws, r, colValorD);
      const saldoCell = numAt(ws, r, colSaldoD);
      if (!fechaVal && !valorCell && !saldoCell) continue;
      if (fechaVal instanceof Date) dates.push(fechaVal);
      if (valorCell) { dataValor += valorCell; dataCantidad++; }
      if (saldoCell) dataSaldo += saldoCell;
    }
  }

  if (cantidad === 0) cantidad = dataCantidad;
  if (valor === 0)    valor    = dataValor;
  if (saldo === 0)    saldo    = dataSaldo;

  if (!nitRaw) nitRaw = filename;
  const ipsInfo = lookupIPS(nitRaw);

  return {
    institucion: ipsInfo?.nombre ?? filename,
    nit:         formatNIT(nitRaw),
    cantidad, valor, saldo,
    dpto:        ipsInfo?.dpto ?? 'MAGDALENA',
    tiene_2026:  dates.some(d => d.getFullYear() === 2026),
    min_fecha:   dates.length ? new Date(Math.min(...dates.map(d => d.getTime()))) : null,
    max_fecha:   dates.length ? new Date(Math.max(...dates.map(d => d.getTime()))) : null,
    archivo:     filename,
  };
}

// ── Format detection ─────────────────────────────────────────────────────────
function detectFormat(ws: ExcelJS.Worksheet): 'A' | 'B' {
  const title = norm(cellVal(ws, 1, 1));
  if (title.includes('estado de cuenta')) return 'B';

  for (let r = 1; r <= 15; r++) {
    for (let c = 1; c <= 8; c++) {
      const t = norm(cellVal(ws, r, c));
      if (t.includes('numero de reclamo') || t.includes('fecha de radicacion')) return 'A';
      if (t.includes('nro factura')) return 'B';
    }
  }
  return 'A';
}

// ── Public API ───────────────────────────────────────────────────────────────
// eslint-disable-next-line @typescript-eslint/no-explicit-any
export async function parseExcelFile(buffer: any, filename: string): Promise<ParsedRecord> {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load(buffer);

  // Prefer first non-PENDIENTE sheet
  let ws = wb.worksheets[0];
  for (const sheet of wb.worksheets) {
    if (!sheet.name.toLowerCase().includes('pendiente')) { ws = sheet; break; }
  }

  const fmt = detectFormat(ws);
  return fmt === 'B' ? parseFormatB(ws, filename) : parseFormatA(ws, filename);
}
