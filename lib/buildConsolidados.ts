import ExcelJS from 'exceljs';
import { ParsedRecord } from './parseExcel';

const NUM_FMT = '_ * #,##0_ ;_ * \\-#,##0_ ;_ * "-"_ ;_ @_ ';
const FONT_NAME = 'Tahoma';

function cell(ws: ExcelJS.Worksheet, row: number, col: number) {
  return ws.getCell(row, col);
}

function applyHeaderStyle(c: ExcelJS.Cell) {
  c.font = { name: FONT_NAME, size: 10, bold: true };
  c.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
  c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFD9D9D9' } };
  c.border = {
    top:    { style: 'thin' }, bottom: { style: 'thin' },
    left:   { style: 'thin' }, right:  { style: 'thin' },
  };
}

function applyDataStyle(c: ExcelJS.Cell, align: 'left' | 'center' | 'right') {
  c.font = { name: FONT_NAME, size: 10 };
  c.alignment = { horizontal: align };
  c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFFFFF' } };
  c.border = {
    top:    { style: 'thin' }, bottom: { style: 'thin' },
    left:   { style: 'thin' }, right:  { style: 'thin' },
  };
}

export async function buildConsolidados(
  records: ParsedRecord[],
  title1 = 'CARTERA GENERAL CARTERAS GRUPO CAMPBELL',
  title2 = 'CARTERA SEGUROS DEL ESTADO',
): Promise<Buffer> {
  const wb = new ExcelJS.Workbook();
  wb.creator = 'Consolidados App';
  const ws = wb.addWorksheet('RESUMEN');

  // ── Column widths (matches CONSOLIDADOS exactly) ─────────────────────────
  ws.getColumn('A').width = 4.29;
  ws.getColumn('B').width = 49.0;
  ws.getColumn('C').width = 16.57;
  ws.getColumn('D').width = 16.0;
  ws.getColumn('E').width = 19.14;
  ws.getColumn('F').width = 17.86;
  ws.getColumn('G').width = 15.86;

  // ── Row heights ───────────────────────────────────────────────────────────
  ws.getRow(5).height = 6.75;
  ws.getRow(7).height = 25.5;

  // ── Row 2: title 1 (merged B2:G2) ────────────────────────────────────────
  ws.mergeCells('B2:G2');
  const t1 = cell(ws, 2, 2);
  t1.value = title1;
  t1.font  = { name: FONT_NAME, size: 10, bold: true };
  t1.alignment = { horizontal: 'center' };

  // ── Row 3: title 2 (merged B3:G3) ────────────────────────────────────────
  ws.mergeCells('B3:G3');
  const t2 = cell(ws, 3, 2);
  t2.value = title2;
  t2.font  = { name: FONT_NAME, size: 10, bold: true };
  t2.alignment = { horizontal: 'center' };

  // ── Row 4: blank merged B4:F4 ────────────────────────────────────────────
  ws.mergeCells('B4:F4');

  // ── Row 6: totals row (D6, E6, F6) ───────────────────────────────────────
  const totalCantidad = records.reduce((s, r) => s + r.cantidad, 0);
  const totalValor    = records.reduce((s, r) => s + r.valor, 0);
  const totalSaldo    = records.reduce((s, r) => s + r.saldo, 0);

  const d6 = cell(ws, 6, 4);
  d6.value  = totalCantidad;
  d6.font   = { name: FONT_NAME, size: 10, bold: true };
  d6.numFmt = NUM_FMT;

  const e6 = cell(ws, 6, 5);
  e6.value  = totalValor;
  e6.font   = { name: FONT_NAME, size: 10, bold: true };
  e6.numFmt = NUM_FMT;

  const f6 = cell(ws, 6, 6);
  f6.value  = totalSaldo;
  f6.font   = { name: FONT_NAME, size: 10, bold: true };
  f6.numFmt = NUM_FMT;

  // ── Row 7: column headers ─────────────────────────────────────────────────
  const headers = ['INSTITUCIÓN', 'NIT', 'CANTIDAD DE FACTURAS', 'VALOR FACTURA', 'CARTERA ACTIVA', 'DPTO'];
  headers.forEach((h, i) => {
    const c = cell(ws, 7, i + 2);  // B=2 … G=7
    c.value = h;
    applyHeaderStyle(c);
    if (i >= 2 && i <= 4) c.numFmt = NUM_FMT;
  });

  // ── Data rows starting at row 8 ───────────────────────────────────────────
  records.forEach((rec, idx) => {
    const r = idx + 8;
    ws.getRow(r).height = 15;

    const bCell = cell(ws, r, 2);
    bCell.value = rec.institucion;
    applyDataStyle(bCell, 'left');

    const cCell = cell(ws, r, 3);
    cCell.value = rec.nit;
    applyDataStyle(cCell, 'center');

    const dCell = cell(ws, r, 4);
    dCell.value  = rec.cantidad;
    dCell.numFmt = NUM_FMT;
    applyDataStyle(dCell, 'left');

    const eCell = cell(ws, r, 5);
    eCell.value  = rec.valor;
    eCell.numFmt = NUM_FMT;
    applyDataStyle(eCell, 'left');

    const fCell = cell(ws, r, 6);
    fCell.value  = rec.saldo;
    fCell.numFmt = NUM_FMT;
    applyDataStyle(fCell, 'left');

    const gCell = cell(ws, r, 7);
    gCell.value = rec.dpto;
    applyDataStyle(gCell, 'left');
  });

  const arrayBuffer = await wb.xlsx.writeBuffer();
  return Buffer.from(arrayBuffer);
}
