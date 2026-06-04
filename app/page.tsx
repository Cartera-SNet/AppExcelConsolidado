'use client';

import { useState, useRef, useCallback, DragEvent } from 'react';

interface RecordRow {
  institucion: string;
  nit: string;
  cantidad: number;
  valor: number;
  saldo: number;
  dpto: string;
  tiene_2026: boolean;
  min_fecha: string | null;
  max_fecha: string | null;
  archivo: string;
}

function fmtNum(n: number) {
  return Math.round(n).toLocaleString('es-CO');
}

export default function HomePage() {
  const [files, setFiles]             = useState<File[]>([]);
  const [records, setRecords]         = useState<RecordRow[]>([]);
  const [downloadUrl, setDownloadUrl] = useState<string | null>(null);
  const [errors, setErrors]           = useState<string[]>([]);
  const [loading, setLoading]         = useState(false);
  const [dragOver, setDragOver]       = useState(false);
  const [title1, setTitle1]           = useState('CARTERA GENERAL CARTERAS GRUPO CAMPBELL');
  const [title2, setTitle2]           = useState('CARTERA SEGUROS DEL ESTADO');

  const inputRef = useRef<HTMLInputElement>(null);

  const addFiles = useCallback((incoming: FileList | null) => {
    if (!incoming) return;
    const xlsx = Array.from(incoming).filter(f => f.name.toLowerCase().endsWith('.xlsx'));
    setFiles(prev => {
      const existing = new Set(prev.map(f => f.name));
      return [...prev, ...xlsx.filter(f => !existing.has(f.name))];
    });
  }, []);

  const onDrop = (e: DragEvent) => {
    e.preventDefault();
    setDragOver(false);
    addFiles(e.dataTransfer.files);
  };

  const removeFile = (name: string) =>
    setFiles(prev => prev.filter(f => f.name !== name));

  const handleSubmit = async () => {
    if (!files.length) return;
    setLoading(true);
    setErrors([]);
    setRecords([]);
    if (downloadUrl) URL.revokeObjectURL(downloadUrl);
    setDownloadUrl(null);

    try {
      const fd = new FormData();
      files.forEach(f => fd.append('files', f));
      fd.append('title1', title1);
      fd.append('title2', title2);

      // 1) Get preview JSON
      const previewRes = await fetch('/api/preview', { method: 'POST', body: fd });
      if (previewRes.ok) {
        const data = await previewRes.json();
        setRecords(data.records ?? []);
        const errs: string[] = data.errors ?? [];
        if (errs.length) setErrors(errs);
      }

      // 2) Get the Excel file
      const fd2 = new FormData();
      files.forEach(f => fd2.append('files', f));
      fd2.append('title1', title1);
      fd2.append('title2', title2);
      const res = await fetch('/api/process', { method: 'POST', body: fd2 });

      if (!res.ok) {
        const body = await res.json().catch(() => ({}));
        setErrors(prev => [...prev, body.error ?? 'Error al procesar.']);
      } else {
        const blob = await res.blob();
        setDownloadUrl(URL.createObjectURL(blob));
      }
    } catch (err) {
      setErrors(['Error de red: ' + String(err)]);
    } finally {
      setLoading(false);
    }
  };

  const totalCantidad = records.reduce((s, r) => s + r.cantidad, 0);
  const totalValor    = records.reduce((s, r) => s + r.valor, 0);
  const totalSaldo    = records.reduce((s, r) => s + r.saldo, 0);

  return (
    <div className="page">
      <header className="site-header">
        <div>
          <h1>Consolidados de Cartera</h1>
          <p>Carga archivos Excel por IPS y genera el CONSOLIDADOS en formato exacto</p>
        </div>
      </header>

      <div className="panel">
        <h2>Archivos Excel (.xlsx)</h2>

        <label
          className={`dropzone${dragOver ? ' drag-over' : ''}`}
          onDragOver={e => { e.preventDefault(); setDragOver(true); }}
          onDragLeave={() => setDragOver(false)}
          onDrop={onDrop}
          onClick={() => inputRef.current?.click()}
        >
          <svg width="36" height="36" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.5">
            <path d="M4 16v2a2 2 0 002 2h12a2 2 0 002-2v-2M12 12V4m0 0L8 8m4-4l4 4" strokeLinecap="round" strokeLinejoin="round"/>
          </svg>
          <strong>Arrastra archivos aquí o haz clic para seleccionar</strong>
          <small>Soporta: formato La Previsora / Seguros y formato Estado de Cuenta</small>
        </label>
        <input ref={inputRef} type="file" accept=".xlsx" multiple
          onChange={e => addFiles(e.target.files)} />

        {files.length > 0 && (
          <div className="file-count">
            {files.length} archivo(s) — haz clic en ✕ para quitar:
            <ul style={{ marginTop: 6, paddingLeft: 18 }}>
              {files.map(f => (
                <li key={f.name} style={{ fontSize: '.83rem', display: 'flex', alignItems: 'center', gap: 8, marginBottom: 3 }}>
                  📄 {f.name}
                  <button onClick={() => removeFile(f.name)}
                    style={{ background: 'none', border: 'none', cursor: 'pointer', color: '#991b1b', fontWeight: 700 }}>
                    ✕
                  </button>
                </li>
              ))}
            </ul>
          </div>
        )}

        <div className="titles-grid">
          <div className="field">
            <label>Título línea 1</label>
            <input value={title1} onChange={e => setTitle1(e.target.value)} />
          </div>
          <div className="field">
            <label>Título línea 2</label>
            <input value={title2} onChange={e => setTitle2(e.target.value)} />
          </div>
        </div>

        <button className="btn-primary" onClick={handleSubmit} disabled={loading || !files.length}>
          {loading
            ? <><div className="spinner" /> Procesando…</>
            : '▶ Generar CONSOLIDADOS'}
        </button>
      </div>

      {errors.length > 0 && (
        <div className="alert alert-error">
          <strong>Advertencias / errores:</strong>
          <ul>{errors.map((e, i) => <li key={i}>{e}</li>)}</ul>
        </div>
      )}

      {downloadUrl && (
        <div className="download-block">
          <a href={downloadUrl} download="CONSOLIDADOS.xlsx">
            <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
              <path d="M12 15V3m0 12l-4-4m4 4l4-4M2 17v2a2 2 0 002 2h16a2 2 0 002-2v-2" strokeLinecap="round" strokeLinejoin="round"/>
            </svg>
            Descargar CONSOLIDADOS.xlsx
          </a>
          <span className="download-info">✓ {records.length} IPS procesadas</span>
        </div>
      )}

      {records.length > 0 && (
        <div className="panel">
          <h2>Vista previa — hoja RESUMEN</h2>

          <div className="table-wrap">
            <table className="summary-table">
              <tbody>
                <tr className="totals-row">
                  <td></td><td></td>
                  <td>{fmtNum(totalCantidad)}</td>
                  <td>{fmtNum(totalValor)}</td>
                  <td>{fmtNum(totalSaldo)}</td>
                  <td></td>
                </tr>
              </tbody>
              <thead>
                <tr>
                  <th>INSTITUCIÓN</th>
                  <th>NIT</th>
                  <th>CANTIDAD DE<br/>FACTURAS</th>
                  <th>VALOR FACTURA</th>
                  <th>CARTERA ACTIVA</th>
                  <th>DPTO</th>
                </tr>
              </thead>
              <tbody>
                {records.map((rec, i) => (
                  <tr key={i} style={{ background: i % 2 === 1 ? 'var(--row-alt)' : undefined }}>
                    <td>{rec.institucion}</td>
                    <td className="center">{rec.nit}</td>
                    <td className="right">{fmtNum(rec.cantidad)}</td>
                    <td className="right">{fmtNum(rec.valor)}</td>
                    <td className="right">{fmtNum(rec.saldo)}</td>
                    <td>{rec.dpto}</td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>

          <h2 style={{ marginTop: 24 }}>Validación de fechas de radicación</h2>
          <table className="check-table">
            <thead>
              <tr><th>Archivo</th><th>Rango de fechas</th><th>¿Hay 2026?</th></tr>
            </thead>
            <tbody>
              {records.map((rec, i) => (
                <tr key={i}>
                  <td>{rec.archivo}</td>
                  <td style={{ fontSize: '.83rem' }}>
                    {rec.min_fecha && rec.max_fecha
                      ? `${rec.min_fecha} a ${rec.max_fecha}`
                      : 'Sin fechas detectadas'}
                  </td>
                  <td>
                    <span className={`badge ${rec.tiene_2026 ? 'badge-yes' : 'badge-no'}`}>
                      {rec.tiene_2026 ? 'Sí' : 'No'}
                    </span>
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      )}
    </div>
  );
}
