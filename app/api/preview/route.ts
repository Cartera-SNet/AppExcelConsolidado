import { NextRequest, NextResponse } from 'next/server';
import { parseExcelFile } from '@/lib/parseExcel';

export const runtime = 'nodejs';
export const maxDuration = 60;

function fmtDate(d: Date | null): string | null {
  if (!d) return null;
  return d.toLocaleDateString('es-CO', { day: '2-digit', month: '2-digit', year: 'numeric' });
}

export async function POST(req: NextRequest) {
  try {
    const formData = await req.formData();
    const files = formData.getAll('files') as File[];

    const records = [];
    const errors: string[] = [];

    for (const file of files) {
      if (!file.name.toLowerCase().endsWith('.xlsx')) {
        errors.push(`Archivo no permitido: ${file.name}`);
        continue;
      }
      try {
        const buffer = Buffer.from(await file.arrayBuffer());
        const rec = await parseExcelFile(buffer, file.name);
        records.push({
          ...rec,
          min_fecha: fmtDate(rec.min_fecha),
          max_fecha: fmtDate(rec.max_fecha),
        });
      } catch (err: unknown) {
        errors.push(`${file.name}: ${err instanceof Error ? err.message : String(err)}`);
      }
    }

    return NextResponse.json({ records, errors });
  } catch (err) {
    console.error(err);
    return NextResponse.json({ error: 'Error interno' }, { status: 500 });
  }
}
