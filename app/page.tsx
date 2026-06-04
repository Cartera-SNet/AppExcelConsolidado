'use client';

import { useState, useRef, useCallback, DragEvent, useMemo } from 'react';
import Image from 'next/image';

interface RecordRow {
  institucion: string;
  nit: string;
  cantidad: number;
  valor: number;
  saldo: number;
  dpto: string;
  all_dates: string[];      // ISO strings
  min_fecha: string | null;
  max_fecha: string | null;
  archivo: string;
}

// ── Date filter operators ────────────────────────────────────────────────────
type Operator = 'any_after' | 'any_before' | 'all_after' | 'all_before' | 'any_in_year' | 'any_outside_year';

const OPERATORS: { value: Operator; label: string; needsDate: boolean; needsYear: boolean }[] = [
  { value: 'any_after',       label: 'Alguna fecha ES POSTERIOR a',   needsDate: true,  needsYear: false },
  { value: 'any_before',      label: 'Alguna fecha ES ANTERIOR a',    needsDate: true,  needsYear: false },
  { value: 'all_after',       label: 'TODAS las fechas son después de', needsDate: true,  needsYear: false },
  { value: 'all_before',      label: 'TODAS las fechas son antes de',   needsDate: true,  needsYear: false },
  { value: 'any_in_year',     label: 'Alguna fecha está EN el año',    needsDate: false, needsYear: true  },
  { value: 'any_outside_year',label: 'Alguna fecha FUERA del año',     needsDate: false, needsYear: true  },
];

function applyFilter(dates: string[], op: Operator, refDate: string, refYear: string): boolean {
  if (!dates.length) return false;
  const parsed = dates.map(d => new Date(d));

  if (op === 'any_in_year') {
    const y = parseInt(refYear);
    return parsed.some(d => d.getFullYear() === y);
  }
  if (op === 'any_outside_year') {
    const y = parseInt(refYear);
    return parsed.some(d => d.getFullYear() !== y);
  }

  const ref = new Date(refDate + 'T00:00:00');
  if (isNaN(ref.getTime())) return false;

  if (op === 'any_after')  return parsed.some(d => d > ref);
  if (op === 'any_before') return parsed.some(d => d < ref);
  if (op === 'all_after')  return parsed.every(d => d > ref);
  if (op === 'all_before') return parsed.every(d => d < ref);
  return false;
}

function fmtNum(n: number) {
  return Math.round(n).toLocaleString('es-CO');
}

// ── Icons ────────────────────────────────────────────────────────────────────
const IconUpload = () => (
  <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
    <path d="M21 15v4a2 2 0 01-2 2H5a2 2 0 01-2-2v-4M17 8l-5-5-5 5M12 3v12"/>
  </svg>
);
const IconFile = () => (
  <svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.5" strokeLinecap="round" strokeLinejoin="round">
    <path d="M14 2H6a2 2 0 00-2 2v16a2 2 0 002 2h12a2 2 0 002-2V8z"/><polyline points="14 2 14 8 20 8"/>
  </svg>
);
const IconDownload = () => (
  <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.2" strokeLinecap="round" strokeLinejoin="round">
    <path d="M12 15V3m0 12l-4-4m4 4l4-4M2 17v2a2 2 0 002 2h16a2 2 0 002-2v-2"/>
  </svg>
);
const IconCheck = () => (
  <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.5" strokeLinecap="round" strokeLinejoin="round">
    <polyline points="20 6 9 17 4 12"/>
  </svg>
);
const IconCalendar = () => (
  <svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
    <rect x="3" y="4" width="18" height="18" rx="2"/><line x1="16" y1="2" x2="16" y2="6"/><line x1="8" y1="2" x2="8" y2="6"/><line x1="3" y1="10" x2="21" y2="10"/>
  </svg>
);
const IconTable = () => (
  <svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
    <rect x="3" y="3" width="18" height="18" rx="2"/><line x1="3" y1="9" x2="21" y2="9"/><line x1="3" y1="15" x2="21" y2="15"/><line x1="9" y1="9" x2="9" y2="21"/>
  </svg>
);
const IconSettings = () => (
  <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
    <circle cx="12" cy="12" r="3"/><path d="M19.4 15a1.65 1.65 0 00.33 1.82l.06.06a2 2 0 010 2.83 2 2 0 01-2.83 0l-.06-.06a1.65 1.65 0 00-1.82-.33 1.65 1.65 0 00-1 1.51V21a2 2 0 01-2 2 2 2 0 01-2-2v-.09A1.65 1.65 0 009 19.4a1.65 1.65 0 00-1.82.33l-.06.06a2 2 0 01-2.83-2.83l.06-.06A1.65 1.65 0 004.68 15a1.65 1.65 0 00-1.51-1H3a2 2 0 01-2-2 2 2 0 012-2h.09A1.65 1.65 0 004.6 9a1.65 1.65 0 00-.33-1.82l-.06-.06a2 2 0 012.83-2.83l.06.06A1.65 1.65 0 009 4.68a1.65 1.65 0 001-1.51V3a2 2 0 012-2 2 2 0 012 2v.09a1.65 1.65 0 001 1.51 1.65 1.65 0 001.82-.33l.06-.06a2 2 0 012.83 2.83l-.06.06A1.65 1.65 0 0019.4 9a1.65 1.65 0 001.51 1H21a2 2 0 012 2 2 2 0 01-2 2h-.09a1.65 1.65 0 00-1.51 1z"/>
  </svg>
);

// ── Main component ────────────────────────────────────────────────────────────
export default function HomePage() {
  const [files, setFiles]               = useState<File[]>([]);
  const [records, setRecords]           = useState<RecordRow[]>([]);
  const [downloadUrl, setDownloadUrl]   = useState<string | null>(null);
  const [errors, setErrors]             = useState<string[]>([]);
  const [loading, setLoading]           = useState(false);
  const [dragOver, setDragOver]         = useState(false);
  const [showAdvanced, setShowAdvanced] = useState(false);
  const [cartera, setCartera]           = useState('CARTERA GENERAL CARTERAS GRUPO CAMPBELL');
  const [aseguradora, setAseguradora]   = useState('CARTERA SEGUROS DEL ESTADO');

  // Date filter state
  const [filterOp, setFilterOp]       = useState<Operator>('any_in_year');
  const [filterDate, setFilterDate]   = useState('2025-12-31');
  const [filterYear, setFilterYear]   = useState('2026');

  const inputRef = useRef<HTMLInputElement>(null);

  const addFiles = useCallback((incoming: FileList | null) => {
    if (!incoming) return;
    const xlsx = Array.from(incoming).filter(f => f.name.toLowerCase().endsWith('.xlsx'));
    setFiles(prev => {
      const existing = new Set(prev.map(f => f.name));
      return [...prev, ...xlsx.filter(f => !existing.has(f.name))];
    });
  }, []);

  const onDrop = (e: DragEvent) => { e.preventDefault(); setDragOver(false); addFiles(e.dataTransfer.files); };

  const clearAll = () => {
    setFiles([]); setRecords([]); setErrors([]);
    if (downloadUrl) URL.revokeObjectURL(downloadUrl);
    setDownloadUrl(null);
  };

  const handleSubmit = async () => {
    if (!files.length) return;
    setLoading(true); setErrors([]); setRecords([]);
    if (downloadUrl) URL.revokeObjectURL(downloadUrl);
    setDownloadUrl(null);

    const makeForm = () => {
      const fd = new FormData();
      files.forEach(f => fd.append('files', f));
      fd.append('title1', cartera);
      fd.append('title2', aseguradora);
      return fd;
    };

    try {
      const [previewRes, processRes] = await Promise.all([
        fetch('/api/preview', { method: 'POST', body: makeForm() }),
        fetch('/api/process', { method: 'POST', body: makeForm() }),
      ]);

      if (previewRes.ok) {
        const data = await previewRes.json();
        setRecords(data.records ?? []);
        if ((data.errors ?? []).length) setErrors(data.errors);
      }

      if (!processRes.ok) {
        const body = await processRes.json().catch(() => ({}));
        setErrors(prev => [...prev, body.error ?? 'Error al procesar.']);
      } else {
        setDownloadUrl(URL.createObjectURL(await processRes.blob()));
      }
    } catch (err) {
      setErrors(['Error de red: ' + String(err)]);
    } finally {
      setLoading(false);
    }
  };

  // ── Computed filter results ──────────────────────────────────────────────
  const currentOp = OPERATORS.find(o => o.value === filterOp)!;

  const filterResults = useMemo(() => {
    return records.map(rec => ({
      ...rec,
      cumple: applyFilter(rec.all_dates, filterOp, filterDate, filterYear),
      fechas_que_cumplen: rec.all_dates
        .filter(d => {
          const parsed = [d];
          return applyFilter(parsed, filterOp, filterDate, filterYear);
        })
        .length,
    }));
  }, [records, filterOp, filterDate, filterYear]);

  const totalCumple  = filterResults.filter(r => r.cumple).length;
  const totalCantidad = records.reduce((s, r) => s + r.cantidad, 0);
  const totalValor    = records.reduce((s, r) => s + r.valor, 0);
  const totalSaldo    = records.reduce((s, r) => s + r.saldo, 0);

  return (
    <div className="page">

      {/* Header */}
      <header className="site-header">
        <div className="site-header-left">
          <Image src="/logo.svg" alt="Robot Excel" width={52} height={52} className="header-logo" />
          <div>
            <h1>Consolidados de Cartera</h1>
            <p>Carga los Excel de cada IPS y genera el archivo consolidado listo para descargar</p>
          </div>
        </div>
        <span className="header-badge">Estado de Cuenta</span>
      </header>

      {/* Upload panel */}
      <div className="panel">
        <div className="panel-title"><IconUpload />Paso 1 — Carga los archivos Excel de las IPS</div>

        <label
          className={`dropzone${dragOver ? ' drag-over' : ''}`}
          onDragOver={e => { e.preventDefault(); setDragOver(true); }}
          onDragLeave={() => setDragOver(false)}
          onDrop={onDrop}
          onClick={() => inputRef.current?.click()}
        >
          <div className="dropzone-icon">
            <svg width="24" height="24" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.8" strokeLinecap="round" strokeLinejoin="round">
              <path d="M14 2H6a2 2 0 00-2 2v16a2 2 0 002 2h12a2 2 0 002-2V8z"/><polyline points="14 2 14 8 20 8"/><line x1="12" y1="18" x2="12" y2="12"/><line x1="9" y1="15" x2="15" y2="15"/>
            </svg>
          </div>
          <strong>Arrastra los archivos .xlsx aquí</strong>
          <small>o haz clic para seleccionarlos desde tu computador</small>
        </label>
        <input ref={inputRef} type="file" accept=".xlsx" multiple onChange={e => addFiles(e.target.files)} />

        {files.length > 0 && (
          <div className="file-list">
            {files.map(f => (
              <div key={f.name} className="file-chip">
                <IconFile />
                <span title={f.name}>{f.name}</span>
                <button onClick={() => setFiles(prev => prev.filter(x => x.name !== f.name))} title="Quitar">✕</button>
              </div>
            ))}
          </div>
        )}

        <div className="advanced-toggle" onClick={() => setShowAdvanced(v => !v)}>
          <svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"
            style={{ transform: showAdvanced ? 'rotate(90deg)' : 'none', transition: 'transform .2s' }}>
            <polyline points="9 18 15 12 9 6"/>
          </svg>
          Opciones avanzadas — cambiar nombre del reporte en el Excel
        </div>

        {showAdvanced && (
          <div className="titles-grid">
            <div className="field">
              <label>Nombre del grupo / cartera</label>
              <input value={cartera} onChange={e => setCartera(e.target.value)} placeholder="Ej: CARTERA GRUPO CAMPBELL" />
              <span className="field-hint">Aparece en la línea 1 del Excel generado</span>
            </div>
            <div className="field">
              <label>Nombre de la aseguradora</label>
              <input value={aseguradora} onChange={e => setAseguradora(e.target.value)} placeholder="Ej: SEGUROS DEL ESTADO" />
              <span className="field-hint">Aparece en la línea 2 del Excel generado</span>
            </div>
          </div>
        )}

        <div className="actions-row">
          <button className="btn-primary" onClick={handleSubmit} disabled={loading || !files.length}>
            {loading ? <><div className="spinner" />Procesando…</> : <><IconSettings />Generar CONSOLIDADOS</>}
          </button>
          {(files.length > 0 || records.length > 0) && !loading && (
            <button className="btn-ghost" onClick={clearAll}>Limpiar todo</button>
          )}
        </div>
      </div>

      {/* Errors */}
      {errors.length > 0 && (
        <div className="alert alert-error">
          <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" style={{flexShrink:0}}><circle cx="12" cy="12" r="10"/><line x1="12" y1="8" x2="12" y2="12"/><line x1="12" y1="16" x2="12.01" y2="16"/></svg>
          <div><strong>Algunos archivos tuvieron problemas</strong><ul>{errors.map((e, i) => <li key={i}>{e}</li>)}</ul></div>
        </div>
      )}

      {/* Stats */}
      {records.length > 0 && (
        <>
          <div className="stats-strip">
            <div className="stat-card">
              <span className="stat-label">IPS procesadas</span>
              <span className="stat-value">{records.length}</span>
              <span className="stat-sub">{totalCumple > 0 ? `${totalCumple} cumplen el filtro` : 'Ninguna cumple el filtro'}</span>
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
              <div className="download-banner-left"><IconCheck />¡Listo! Tu archivo CONSOLIDADOS está generado con {records.length} IPS</div>
              <a className="download-btn" href={downloadUrl} download="CONSOLIDADOS.xlsx"><IconDownload />Descargar CONSOLIDADOS.xlsx</a>
            </div>
          )}
        </>
      )}

      {/* Results panel */}
      {records.length > 0 && (
        <div className="panel">
          <div className="panel-title"><IconTable />Paso 2 — Revisa el resumen antes de descargar</div>

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
                  <th>INSTITUCIÓN</th><th>NIT</th>
                  <th>CANTIDAD DE<br/>FACTURAS</th><th>VALOR FACTURA</th><th>CARTERA ACTIVA</th><th>DPTO</th>
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

          {/* ── Date filter section ── */}
          <div className="section-divider"><IconCalendar />Validación de fechas de radicación</div>

          {/* Filter builder */}
          <div className="filter-card">
            <div className="filter-card-title">¿Qué quieres verificar?</div>
            <div className="filter-row">
              <select
                className="filter-select"
                value={filterOp}
                onChange={e => setFilterOp(e.target.value as Operator)}
              >
                {OPERATORS.map(op => (
                  <option key={op.value} value={op.value}>{op.label}</option>
                ))}
              </select>

              {currentOp.needsDate && (
                <input
                  type="date"
                  className="filter-date"
                  value={filterDate}
                  onChange={e => setFilterDate(e.target.value)}
                />
              )}
              {currentOp.needsYear && (
                <input
                  type="number"
                  className="filter-year"
                  value={filterYear}
                  min="2000" max="2099"
                  onChange={e => setFilterYear(e.target.value)}
                  placeholder="Año"
                />
              )}
            </div>

            <div className="filter-summary">
              {totalCumple > 0
                ? <span className="filter-badge filter-badge-warn">⚠ {totalCumple} de {records.length} IPS cumplen esta condición</span>
                : <span className="filter-badge filter-badge-ok">✓ Ninguna IPS cumple esta condición</span>
              }
            </div>
          </div>

          {/* Results table */}
          <table className="check-table">
            <thead>
              <tr>
                <th>Archivo</th>
                <th>Rango de fechas</th>
                <th>Fechas que cumplen</th>
                <th>¿Cumple condición?</th>
              </tr>
            </thead>
            <tbody>
              {filterResults.map((rec, i) => (
                <tr key={i} style={{ background: rec.cumple ? '#fffbeb' : undefined }}>
                  <td style={{ fontSize: '.8rem', color: 'var(--muted)' }}>{rec.archivo}</td>
                  <td style={{ fontSize: '.83rem' }}>
                    {rec.min_fecha && rec.max_fecha
                      ? `${rec.min_fecha} — ${rec.max_fecha}`
                      : <span style={{ color: 'var(--muted)' }}>Sin fechas</span>}
                  </td>
                  <td className="center" style={{ fontSize: '.83rem' }}>
                    {rec.all_dates.length > 0
                      ? `${rec.fechas_que_cumplen} de ${rec.all_dates.length}`
                      : '—'}
                  </td>
                  <td>
                    <span className={`badge ${rec.cumple ? 'badge-yes' : 'badge-no'}`}>
                      {rec.cumple ? '⚠ Sí' : '✓ No'}
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
