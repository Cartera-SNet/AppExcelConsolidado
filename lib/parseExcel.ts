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

// ─────────────────────────────────────────────────────────────────────────────
// NIT lookup table from Mapeo_IPS_Empresas.xlsx
// Key: normalised NIT digits → { institucion, dpto }
// ─────────────────────────────────────────────────────────────────────────────
const IPS_MAP: Record<string, { nombre: string; dpto: string }> = {
  '9008270653': { nombre: 'CENTRO DE DIAGNOSTICO E IMAGENES BAHIA', dpto: 'MAGDALENA' },
  '9002780 0':  { nombre: 'FUNDACION CAMPBELL', dpto: 'ATLANTICO' },
  '9000027800': { nombre: 'FUNDACION CAMPBELL', dpto: 'ATLANTICO' },
  '9002670642': { nombre: 'INVERSIONES AZALUD S.A.S', dpto: 'MAGDALENA' },
  '9006005509': { nombre: 'INVERSIONES MÉDICAS BARU', dpto: 'BOLIVAR' },
  '9009548004': { nombre: 'CENTRO MEDICO Y DE REHABILITACION BARU', dpto: 'BOLIVAR' },
  '9008265097': { nombre: 'RUC MAGDALENA', dpto: 'MAGDALENA' },
  '9006577310': { nombre: 'CENTRO MEDICO Y DE REHABILITACION BAHIA', dpto: 'MAGDALENA' },
  '9005133065': { nombre: 'FUNDACIÓN MARIA REINA', dpto: 'SUCRE' },
  '9002573336': { nombre: 'ODONTOTRANS', dpto: 'VALLE DEL CAUCA' },
  '9010812818': { nombre: 'URGETRAUMA', dpto: 'VALLE DEL CAUCA' },
  '9007924171': { nombre: 'RED DE URGENCIAS DE LA COSTA PACIFICA', dpto: 'VALLE DEL CAUCA' },
  '9006313616': { nombre: 'INVERSIONES MÉDICAS VALLE SALUD S.A.S', dpto: 'VALLE DEL CAUCA' },
  '9008473829': { nombre: 'CENTRO MEDICO Y REHABILITACION VALLE SALUD', dpto: 'VALLE DEL CAUCA' },
  '9011497576': { nombre: 'CVO UNIDAD MÉDICA DE TRAUMA DEL VALLE', dpto: 'VALLE DEL CAUCA' },
  '9009007541': { nombre: 'CLINICA VALLE SALUD SAN FERNANDO', dpto: 'VALLE DEL CAUCA' },
  '9004698828': { nombre: 'CENTRO MEDICO SERVISALUD INTEGRAL IPS', dpto: 'VALLE DEL CAUCA' },
  '9015238689': { nombre: 'MOVID IPS S.A.S', dpto: 'ATLANTICO' },
  '9005585950': { nombre: 'FUNDACION MEDICA CAMPBELL', dpto: 'ATLANTICO' },
  '8020243290': { nombre: 'RED DE URGENCIA DE LA COSTA', dpto: 'ATLANTICO' },
};

/** Strip a NIT string to only its digits */
function nitDigits(nit: string): string {
  return nit.replace(/\D/g, '');
}

/** Look up IPS metadata by NIT */
function lookupIPS(nit: string): { nombre: string; dpto: string } | null {
  const digits = nitDigits(nit);
  // Try exact match first
  if (IPS_MAP[digits]) return IPS_MAP[digits];
  // Try any key that starts with the same leading digits (lenient)
  for (const key of Object.keys(IPS_MAP)) {
    const kd = nitDigits(key);
    if (kd === digits) return IPS_MAP[key];
  }
  return null;
}

/** Format NIT with space before check digit: "900267064 2" */
function formatNIT(raw: string): string {
  const digits = nitDigits(raw);
  if (digits.length >= 10) return `${digits.slice(0, -1)} ${digits.slice(-1)}`;
  if (digits.length === 9) return `${digits.slice(0, -1)} ${digits.slice(-1)}`;
  return raw.trim();
}

// ─────────────────────────────────────────────────────────────────────────────
// Normalise text: remove accents, lowercase, collapse spaces
// ─────────────────────────────────────────────────────────────────────────────
function norm(v: unknown): string {
  if (v === null || v === undefined) return '';
  return String(v)
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .replace(/\s+/g, ' ')
    .trim()
    .toLowerCase();
}

// ─────────────────────────────────────────────────────────────────────────────
// FORMAT A — "La Previsora / Seguros" style
//   Row 6 header: NIT del prestador | Razón social | Ramo | Número de reclamo |
//                 Fecha de radicación | Valor neto | Valor saldo
// ─────────────────────────────────────────────────────────────────────────────
function parseFormatA(ws: ExcelJS.Worksheet, filename: string): ParsedRecord {
  // Find the header row
  let headerRow = -1;
  const cols: Record<string, number> = {};

  for (let r = 1; r <= Math.min(ws.rowCount, 15); r++) {
    const row = ws.getRow(r);
    const vals: string[] = [];
    row.eachCell({ includeEmpty: true }, (cell) => vals.push(norm(cell.value)));

    if (vals.some(c => c.includes('numero de reclamo') || c.includes('numero de reclamo'))) {
      headerRow = r;
      vals.forEach((cell, idx) => {
        const c = idx + 1;
        if (cell.includes('nit del prestador'))               cols.nit = c;
        else if (cell.includes('razon social'))               cols.institucion = c;
        else if (cell.includes('numero de reclamo'))          cols.reclamo = c;
        else if (cell.includes('fecha de radicacion') || cell.includes('fecha de radicaci')) cols.fecha = c;
        else if (cell.includes('valor neto'))                 cols.valor = c;
        else if (cell.includes('valor saldo'))                cols.saldo = c;
      });
      break;
    }
  }

  if (headerRow < 0) throw new Error('Formato A: no se encontró la fila de encabezado.');

  // Read summary row (rows 2–4: Cantidad, Valor, Saldo)
  let cantidad = 0, valor = 0, saldo = 0;
  let nitRaw = '';
  let institucionName = '';

  // Try to grab totals from the header block (rows 2–4)
  for (let r = 1; r < headerRow; r++) {
    const row = ws.getRow(r);
    const cells: unknown[] = [];
    row.eachCell({ includeEmpty: true }, (cell) => cells.push(cell.value));
    const labels = cells.map(norm);

    labels.forEach((l, i) => {
      const val = cells[i + 1];
      if (l.includes('cantidad') && typeof val === 'number') cantidad = val;
      if (l.includes('valor cobrado') && typeof val === 'number') valor = val;
      if (l.includes('saldo requerido') && typeof val === 'number') saldo = val;
    });
  }

  // Collect dates and get NIT/institution from data rows
  const dates: Date[] = [];
  let dataCantidad = 0, dataValor = 0, dataSaldo = 0;

  for (let r = headerRow + 1; r <= ws.rowCount; r++) {
    const reclamoVal = ws.getCell(r, cols.reclamo)?.value;
    const fechaVal   = ws.getCell(r, cols.fecha)?.value;
    const valorCell  = ws.getCell(r, cols.valor)?.value;
    const saldoCell  = ws.getCell(r, cols.saldo)?.value;
    const nitCell    = ws.getCell(r, cols.nit)?.value;
    const instCell   = ws.getCell(r, cols.institucion)?.value;

    if ([reclamoVal, fechaVal, valorCell, saldoCell].every(v => v === null || v === '')) continue;

    if (!nitRaw && nitCell) nitRaw = String(nitCell);
    if (!institucionName && instCell) institucionName = String(instCell);
    if (reclamoVal !== null && reclamoVal !== '') dataCantidad++;
    if (typeof valorCell === 'number') dataValor += valorCell;
    if (typeof saldoCell === 'number') dataSaldo += saldoCell;
    if (fechaVal instanceof Date) dates.push(fechaVal);
  }

  // Prefer summary totals if available
  if (cantidad === 0) cantidad = dataCantidad;
  if (valor === 0)    valor    = dataValor;
  if (saldo === 0)    saldo    = dataSaldo;

  // NIT: try summary block if not found in data
  if (!nitRaw) nitRaw = filename;

  const nitFormatted = formatNIT(nitRaw);
  const ipsInfo = lookupIPS(nitRaw);

  return {
    institucion: (ipsInfo?.nombre ?? institucionName) || filename,
    nit: nitFormatted,
    cantidad,
    valor,
    saldo,
    dpto: ipsInfo?.dpto ?? '',
    tiene_2026: dates.some(d => d.getFullYear() === 2026),
    min_fecha: dates.length ? new Date(Math.min(...dates.map(d => d.getTime()))) : null,
    max_fecha: dates.length ? new Date(Math.max(...dates.map(d => d.getTime()))) : null,
    archivo: filename,
  };
}

// ─────────────────────────────────────────────────────────────────────────────
// FORMAT B — "Estado de Cuenta" style (CENTRO DE DIAGNOSTICO)
//   Row 1: title with NIT
//   Row 5: Fecha de corte | ... | Cantidad facturas | | N
//   Row 9: No. | Nro factura | Amparo | Nombre | Fecha Aviso | Fecha Egreso |
//           Valor Cobrado | Saldo Pendiente
// ─────────────────────────────────────────────────────────────────────────────
function parseFormatB(ws: ExcelJS.Worksheet, filename: string): ParsedRecord {
  let cantidad = 0, valor = 0, saldo = 0;
  let nitRaw = '';

  // Extract NIT from title row (row 1)
  const titleVal = String(ws.getCell(1, 1).value ?? '');
  const nitMatch = titleVal.match(/NIT\s*([\d.\-]+)/i);
  if (nitMatch) nitRaw = nitMatch[1];

  // Read summary block
  for (let r = 1; r <= 10; r++) {
    const row = ws.getRow(r);
    const cells: unknown[] = [];
    row.eachCell({ includeEmpty: true }, (cell) => cells.push(cell.value));
    const labels = cells.map(norm);

    labels.forEach((l, i) => {
      const val = cells[i + 1];  // cell after label
      // also check cells[i+2] (label is sometimes 2 cols before value)
      const val2 = cells[i + 2];
      if (l.includes('cantidad facturas')) {
        if (typeof val === 'number') cantidad = val;
        else if (typeof val2 === 'number') cantidad = val2;
      }
      if (l.includes('valor cobrado')) {
        if (typeof val === 'number') valor = val;
        else if (typeof val2 === 'number') valor = val2;
      }
      if (l.includes('saldo requerido')) {
        if (typeof val === 'number') saldo = val;
        else if (typeof val2 === 'number') saldo = val2;
      }
    });
  }

  // Find data header row: "No. | Nro factura | ..."
  let headerRow = -1;
  let colValor = 7, colSaldo = 8, colFecha = 5;

  for (let r = 1; r <= 15; r++) {
    const row = ws.getRow(r);
    const vals: string[] = [];
    row.eachCell({ includeEmpty: true }, (cell) => vals.push(norm(cell.value)));
    if (vals.some(c => c.includes('nro factura') || c.includes('valor cobrado'))) {
      headerRow = r;
      vals.forEach((v, i) => {
        if (v.includes('valor cobrado'))  colValor = i + 1;
        if (v.includes('saldo pendiente')) colSaldo = i + 1;
        if (v.includes('fecha aviso') || v.includes('fecha egreso')) colFecha = i + 1;
      });
      break;
    }
  }

  const dates: Date[] = [];
  if (headerRow > 0) {
    for (let r = headerRow + 1; r <= ws.rowCount; r++) {
      const fechaVal = ws.getCell(r, colFecha)?.value;
      if (fechaVal instanceof Date) dates.push(fechaVal);
    }
  }

  const nitFormatted = formatNIT(nitRaw || filename);
  const ipsInfo = lookupIPS(nitRaw);

  return {
    institucion: ipsInfo?.nombre ?? filename,
    nit: nitFormatted,
    cantidad,
    valor,
    saldo,
    dpto: ipsInfo?.dpto ?? 'MAGDALENA',
    tiene_2026: dates.some(d => d.getFullYear() === 2026),
    min_fecha: dates.length ? new Date(Math.min(...dates.map(d => d.getTime()))) : null,
    max_fecha: dates.length ? new Date(Math.max(...dates.map(d => d.getTime()))) : null,
    archivo: filename,
  };
}

// ─────────────────────────────────────────────────────────────────────────────
// Auto-detect format and parse
// ─────────────────────────────────────────────────────────────────────────────
function detectFormat(ws: ExcelJS.Worksheet): 'A' | 'B' {
  for (let r = 1; r <= Math.min(ws.rowCount, 15); r++) {
    const row = ws.getRow(r);
    const vals: string[] = [];
    row.eachCell({ includeEmpty: true }, (cell) => vals.push(norm(cell.value)));
    if (vals.some(c => c.includes('numero de reclamo') || c.includes('fecha de radicacion'))) return 'A';
    if (vals.some(c => c.includes('nro factura') || c.includes('estado de cuenta'))) return 'B';
  }
  // Check title row for "ESTADO DE CUENTA"
  const title = norm(ws.getCell(1, 1).value);
  if (title.includes('estado de cuenta')) return 'B';
  return 'A';
}

// ─────────────────────────────────────────────────────────────────────────────
// Public API
// ─────────────────────────────────────────────────────────────────────────────
// eslint-disable-next-line @typescript-eslint/no-explicit-any
export async function parseExcelFile(buffer: any, filename: string): Promise<ParsedRecord> {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load(buffer);

  // Use first sheet, or prefer non-"PENDIENTE" sheet
  let ws = wb.worksheets[0];
  for (const sheet of wb.worksheets) {
    const n = sheet.name.toLowerCase();
    if (!n.includes('pendiente')) { ws = sheet; break; }
  }

  const fmt = detectFormat(ws);
  if (fmt === 'B') return parseFormatB(ws, filename);
  return parseFormatA(ws, filename);
}
