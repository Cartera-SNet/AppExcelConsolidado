'use client';

import { useState, useRef, useCallback, DragEvent } from 'react';
import Image from 'next/image';

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
  const [showAdvanced, setShowAdvanced] = useState(false);
  const [cartera, setCartera]         = useState('CARTERA GENERAL CARTERAS GRUPO CAMPBELL');
  const [aseguradora, setAseguradora] = useState('CARTERA SEGUROS DEL ESTADO');

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

  const clearAll = () => {
    setFiles([]);
    setRecords([]);
    setErrors([]);
    if (downloadUrl) URL.revokeObjectURL(downloadUrl);
    setDownloadUrl(null);
  };

  const handleSubmit = async () => {
    if (!files.length) return;
    setLoading(true);
    setErrors([]);
    setRecords([]);
    if (downloadUrl) URL.revokeObjectURL(downloadUrl);
    setDownloadUrl(null);

    try {
      const makeForm = () => {
        const fd = new FormData();
        files.forEach(f => fd.append('files', f));
        fd.append('title1', cartera);
        fd.append('title2', aseguradora);
        return fd;
      };

      const [previewRes, processRes] = await Promise.all([
        fetch('/api/preview', { method: 'POST', body: makeForm() }),
        fetch('/api/process', { method: 'POST', body: makeForm() }),
      ]);

      if (previewRes.ok) {
        const data = await previewRes.json();
        setRecords(data.records ?? []);
        const errs: string[] = data.errors ?? [];
        if (errs.length) setErrors(errs);
      }

      if (!processRes.ok) {
        const body = await processRes.json().catch(() => ({}));
        setErrors(prev => [...prev, body.error ?? 'Error al procesar.']);
      } else {
        const blob = await processRes.blob();
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
  const con2026       = records.filter(r => r.tiene_2026).length;

  return (
    <div className="page">

      {/* ── Header ── */}
      <header className="site-header">
        <div className="site-header-left">
          <Image src="/logo.svg" alt="Robot Excel" width={52} height={52} className="header-logo" />
          <div>
            <h1>Consolidados de Cartera</h1>
            <p>Carga los Excel de cada IPS y genera el archivo consolidado listo para descargar</p>
          </div>
        </div>
        <span className="header-badge">La Previsora · Estado de Cuenta</span>
      </header>

      {/* ── Upload panel ── */}
      <div className="panel">
        <div className="panel-title">
          <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><path d="M21 15v4a2 2 0 01-2 2H5a2 2 0 01-2-2v-4M17 8l-5-5-5 5M12 3v12"/></svg>
          Paso 1 — Carga los archivos Excel de las IPS
        </div>

        <label
          className={`dropzone${dragOver ? ' drag-over' : ''}`}
          onDragOver={e => { e.preventDefault(); setDragOver(true); }}
          onDragLeave={() => setDragOver(false)}
          onDrop={onDrop}
          onClick={() => inputRef.current?.click()}
        >
          <div className="dropzone-icon">
            <svg width="24" height="24" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.8" strokeLinecap="round" strokeLinejoin="round"><path d="M14 2H6a2 2 0 00-2 2v16a2 2 0 002 2h12a2 2 0 002-2V8z"/><polyline points="14 2 14 8 20 8"/><line x1="12" y1="18" x2="12" y2="12"/><line x1="9" y1="15" x2="15" y2="15"/></svg>
          </div>
          <strong>Arrastra los archivos .xlsx aquí</strong>
          <small>o haz clic para seleccionarlos desde tu computador</small>
        </label>
        <input ref={inputRef} type="file" accept=".xlsx" multiple
          onChange={e => addFiles(e.target.files)} />

        {files.length > 0 && (
          <div className="file-list">
            {files.map(f => (
              <div key={f.name} className="file-chip">
                <svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.5" strokeLinecap="round" strokeLinejoin="round"><path d="M14 2H6a2 2 0 00-2 2v16a2 2 0 002 2h12a2 2 0 002-2V8z"/><polyline points="14 2 14 8 20 8"/></svg>
                <span title={f.name}>{f.name}</span>
                <button onClick={() => setFiles(prev => prev.filter(x => x.name !== f.name))} title="Quitar">✕</button>
              </div>
            ))}
          </div>
        )}

        {/* Advanced: rename titles */}
        <div className="advanced-toggle" onClick={() => setShowAdvanced(v => !v)}>
          <svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" style={{ transform: showAdvanced ? 'rotate(90deg)' : 'none', transition: 'transform .2s' }}><polyline points="9 18 15 12 9 6"/></svg>
          Opciones avanzadas — cambiar nombre del reporte en el Excel
        </div>

        {showAdvanced && (
          <div className="titles-grid">
            <div className="field">
              <label>Nombre del grupo / cartera</label>
              <input
                value={cartera}
                onChange={e => setCartera(e.target.value)}
                placeholder="Ej: CARTERA GRUPO CAMPBELL"
              />
              <span className="field-hint">Aparece en la línea 1 del Excel generado</span>
            </div>
            <div className="field">
              <label>Nombre de la aseguradora</label>
              <input
                value={aseguradora}
                onChange={e => setAseguradora(e.target.value)}
                placeholder="Ej: SEGUROS DEL ESTADO"
              />
              <span className="field-hint">Aparece en la línea 2 del Excel generado</span>
            </div>
          </div>
        )}

        <div className="actions-row">
          <button className="btn-primary" onClick={handleSubmit} disabled={loading || !files.length}>
            {loading ? (
              <><div className="spinner" /> Procesando archivos…</>
            ) : (
              <>
                <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.2" strokeLinecap="round" strokeLinejoin="round"><polyline points="22 12 18 12 15 21 9 3 6 12 2 12"/></svg>
                Generar CONSOLIDADOS
              </>
            )}
          </button>
          {(files.length > 0 || records.length > 0) && !loading && (
            <button className="btn-ghost" onClick={clearAll}>Limpiar todo</button>
          )}
        </div>
      </div>

      {/* ── Errors ── */}
      {errors.length > 0 && (
        <div className="alert alert-error">
          <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" style={{flexShrink:0}}><circle cx="12" cy="12" r="10"/><line x1="12" y1="8" x2="12" y2="12"/><line x1="12" y1="16" x2="12.01" y2="16"/></svg>
          <div>
            <strong>Algunos archivos tuvieron problemas</strong>
            <ul>{errors.map((e, i) => <li key={i}>{e}</li>)}</ul>
          </div>
        </div>
      )}

      {/* ── Stats + Download ── */}
      {records.length > 0 && (
        <>
          <div className="stats-strip">
            <div className="stat-card">
              <span className="stat-label">IPS procesadas</span>
              <span className="stat-value">{records.length}</span>
              <span className="stat-sub">{con2026 > 0 ? `${con2026} con fechas en 2026` : 'Sin fechas en 2026'}</span>
            </div>
            <div className="stat-card">
              <span className="stat-label">Total facturas</span>
              <span className="stat-value">{fmtNum(totalCantidad)}</span>
              <span className="stat-sub">registros consolidados</span>
            </div>
            <div className="stat-card">
              <span className="stat-label">Cartera activa</span>
              <span className="stat-value">$ {fmtNum(totalSaldo)}</span>
              <span className="stat-sub">de $ {fmtNum(totalValor)} facturado</span>
            </div>
          </div>

          {downloadUrl && (
            <div className="download-banner">
              <div className="download-banner-left">
                <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.5" strokeLinecap="round" strokeLinejoin="round"><polyline points="20 6 9 17 4 12"/></svg>
                ¡Listo! Tu archivo CONSOLIDADOS está generado con {records.length} IPS
              </div>
              <a className="download-btn" href={downloadUrl} download="CONSOLIDADOS.xlsx">
                <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.2" strokeLinecap="round" strokeLinejoin="round"><path d="M12 15V3m0 12l-4-4m4 4l4-4M2 17v2a2 2 0 002 2h16a2 2 0 002-2v-2"/></svg>
                Descargar CONSOLIDADOS.xlsx
              </a>
            </div>
          )}
        </>
      )}

      {/* ── Preview table ── */}
      {records.length > 0 && (
        <div className="panel">
          <div className="panel-title">
            <svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><rect x="3" y="3" width="18" height="18" rx="2"/><line x1="3" y1="9" x2="21" y2="9"/><line x1="3" y1="15" x2="21" y2="15"/><line x1="9" y1="9" x2="9" y2="21"/></svg>
            Paso 2 — Revisa el resumen antes de descargar
          </div>

          <div className="table-wrap">
            <table className="summary-table">
              <thead>
                <tr className="totals-row">
                  <td></td><td></td>
                  <td>{fmtNum(totalCantidad)}</td>
                  <td>{fmtNum(totalValor)}</td>
                  <td>{fmtNum(totalSaldo)}</td>
                  <td></td>
                </tr>
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
                  <tr key={i}>
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

          <div className="section-divider">Validación de fechas de radicación</div>

          <table className="check-table">
            <thead>
              <tr>
                <th>Archivo</th>
                <th>Rango de fechas</th>
                <th>¿Contiene fechas del 2026?</th>
              </tr>
            </thead>
            <tbody>
              {records.map((rec, i) => (
                <tr key={i}>
                  <td style={{ fontSize: '.8rem', color: 'var(--muted)' }}>{rec.archivo}</td>
                  <td>{rec.min_fecha && rec.max_fecha
                    ? `${rec.min_fecha} — ${rec.max_fecha}`
                    : <span style={{ color: 'var(--muted)' }}>Sin fechas detectadas</span>}
                  </td>
                  <td>
                    <span className={`badge ${rec.tiene_2026 ? 'badge-yes' : 'badge-no'}`}>
                      {rec.tiene_2026 ? '⚠ Sí, hay fechas 2026' : '✓ No'}
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
