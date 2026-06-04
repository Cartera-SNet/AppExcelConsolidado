import { NextRequest, NextResponse } from 'next/server';
import { parseExcelFile } from '@/lib/parseExcel';
import { buildConsolidados } from '@/lib/buildConsolidados';

export const runtime = 'nodejs';
export const maxDuration = 60;

export async function POST(req: NextRequest) {
  try {
    const formData = await req.formData();
    const files = formData.getAll('files') as File[];
    const title1 = (formData.get('title1') as string | null) ?? 'CARTERA GENERAL CARTERAS GRUPO CAMPBELL';
    const title2 = (formData.get('title2') as string | null) ?? 'CARTERA SEGUROS DEL ESTADO';

    if (!files || files.length === 0) {
      return NextResponse.json({ error: 'No se recibieron archivos.' }, { status: 400 });
    }

    const records = [];
    const errors: string[] = [];

    for (const file of files) {
      if (!file.name.toLowerCase().endsWith('.xlsx')) {
        errors.push(`Archivo no permitido: ${file.name}`);
        continue;
      }
      try {
        const buffer = Buffer.from(await file.arrayBuffer());
        const record = await parseExcelFile(buffer, file.name);
        records.push(record);
      } catch (err: unknown) {
        errors.push(`${file.name}: ${err instanceof Error ? err.message : String(err)}`);
      }
    }

    if (records.length === 0) {
      return NextResponse.json(
        { error: 'No se pudo procesar ningún archivo.', details: errors },
        { status: 422 }
      );
    }

    const xlsxBuffer = await buildConsolidados(records, title1, title2);

    return new NextResponse(xlsxBuffer as unknown as BodyInit, {
      status: 200,
      headers: {
        'Content-Type': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        'Content-Disposition': 'attachment; filename="CONSOLIDADOS.xlsx"',
        'X-Errors': JSON.stringify(errors),
        'X-Records': String(records.length),
      },
    });
  } catch (err: unknown) {
    console.error(err);
    return NextResponse.json({ error: 'Error interno del servidor.' }, { status: 500 });
  }
}

// Also expose a preview endpoint (GET with JSON body not standard, so POST with action=preview)
export async function GET() {
  return NextResponse.json({ status: 'ok', version: '2.0' });
}
