/**
 * AnalysisPanel — Painel Pro de Análise recriado em React / Recharts.
 *
 * Lógica 100% fiel ao Analise Dados V6.html:
 *  • Corporate (índice 4 linhas skip): serial=1, model=3, location=6
 *  • SDS:  serial=1, model=5, lastUpdate=10, hostname=2, ip=3, counter=8, manufacturer=4
 *  • NDD:  serial=5, lastCounterDate=7, counterReference=8, accountedPages=9,
 *           pagesDifference=10, percentageDifference=11, hostname=4
 */
import React, { useState, useMemo, useRef, useCallback } from 'react';
import { PieChart, Pie, Cell, Tooltip, ResponsiveContainer, BarChart, Bar, XAxis, YAxis, CartesianGrid } from 'recharts';
import { readRawExcel } from '../services/excelService';

// ── Column index map ─────────────────────────────────────────────────────────

const COL = {
  corp: { serial: 1, model: 3, location: 6 },
  sds:  { serial: 1, model: 5, lastUpdate: 10, hostname: 2, ip: 3, counter: 8, manufacturer: 4 },
  ndd:  { serial: 5, lastCounterDate: 7, counterReference: 8, accountedPages: 9, percentageDifference: 11, hostname: 4 },
};

// ── Helpers ──────────────────────────────────────────────────────────────────

const normalizeSerial = (v: any): string =>
  String(v || '').replace(/[^a-zA-Z0-9]/g, '').toUpperCase();

const cleanNumber = (v: any): number => {
  if (typeof v === 'number') return v;
  if (typeof v !== 'string') return 0;
  return parseFloat(v.replace(/\./g, '').replace(',', '.')) || 0;
};

const parseUpdateDate = (v: any): Date | null => {
  if (!v) return null;
  if (typeof v === 'number') return new Date(Math.round((v - 25569) * 864e5));
  if (v instanceof Date) return v;
  const m = String(v).match(/(\d{1,2})[\/.-](\d{1,2})[\/.-](\d{4})/);
  if (m) return new Date(+m[3], +m[2] - 1, +m[1]);
  const d = new Date(v);
  return isNaN(d.getTime()) ? null : d;
};

function findByPrefix(map: Map<string, any[]>, prefix: string): any[] | null {
  if (!prefix) return null;
  if (map.has(prefix)) return map.get(prefix)!;
  for (const [k, v] of map.entries()) {
    if (k.startsWith(prefix) || prefix.startsWith(k)) return v;
  }
  return null;
}

// ── Types ────────────────────────────────────────────────────────────────────

interface LocationDetail {
  total: number; monitored: number; backups: number; billed: number;
}
interface ModelDetail {
  total: number; monitored: number; billed: number;
  backups: number; active: number; monitoredActive: number;
  monitoringPercentage: string;
}
interface CompRow {
  serial: string; corporateLocation: string;
  sdsModel: string; sdsHostname: string; sdsIp: string;
  sdsLastUpdate: string; sdsCounter: any;
  nddAccountedPages?: number; nddCounterReference?: number;
  percentageDifference?: string;
  status: 'ok' | 'Atenção' | 'Erro'; days?: number;
}

// ── Chart colors ─────────────────────────────────────────────────────────────

const CHART_COLORS = {
  ok:          '#10B981',
  attention:   '#F59E0B',
  error:       '#EF4444',
  mismatch:    '#2563EB',
  unmonitored: '#6B7280',
  backup:      '#9CA3AF',
};
const MFR_COLORS = [
  '#E6007E','#9C27B0','#673AB7','#3F51B5','#2196F3',
  '#03A9F4','#00BCD4','#009688','#4CAF50','#8BC34A',
];

// ── Tooltip ──────────────────────────────────────────────────────────────────

const ChartTip = ({ active, payload }: any) => {
  if (!active || !payload?.length) return null;
  return (
    <div className="bg-white border border-gray-200 rounded-lg shadow-lg px-3 py-2">
      <p className="text-xs font-bold text-gray-800">{payload[0].name}</p>
      <p className="text-xs text-gray-500">{Number(payload[0].value).toLocaleString('pt-BR')}</p>
    </div>
  );
};

// ── Status badge ─────────────────────────────────────────────────────────────

const StatusBadge: React.FC<{ status: string }> = ({ status }) => {
  const cls =
    status === 'ok' || status === 'Monitorado'  ? 'bg-emerald-100 text-emerald-800' :
    status === 'Atenção'                         ? 'bg-amber-100   text-amber-800'   :
    status === 'Erro'                            ? 'bg-red-100     text-red-800'     :
                                                   'bg-gray-100    text-gray-700';
  return (
    <span className={`inline-flex px-2 py-0.5 rounded-full text-[10px] font-bold uppercase ${cls}`}>
      {status}
    </span>
  );
};

// ── KPI card ─────────────────────────────────────────────────────────────────

interface KpiProps {
  label: string; value: number | string; color?: string;
  onClick?: () => void; active?: boolean;
}
const KpiCard: React.FC<KpiProps> = ({ label, value, color = 'text-gray-800', onClick, active }) => (
  <div
    onClick={onClick}
    className={`bg-white rounded-xl border-2 p-4 text-center transition-all duration-200
      ${onClick ? 'cursor-pointer hover:-translate-y-0.5 hover:shadow-md' : ''}
      ${active ? 'border-pink-500 shadow-[0_0_0_3px_rgba(236,72,153,0.2)]' : 'border-gray-200'}`}
  >
    <div className={`text-2xl font-extrabold ${color}`}>{value}</div>
    <div className="text-xs text-gray-500 mt-0.5 leading-tight">{label}</div>
  </div>
);

// ── Sortable table ────────────────────────────────────────────────────────────

interface ColDef { key: string; label: string; className?: string; render?: (row: any) => React.ReactNode; }

const SortableTable: React.FC<{
  cols: ColDef[]; rows: any[];
  onRowClick?: (row: any) => void;
  activeRowKey?: string; activeRowValue?: any;
  footer?: React.ReactNode;
  emptyText?: string;
}> = ({ cols, rows, onRowClick, activeRowKey, activeRowValue, footer, emptyText = 'Nenhum item.' }) => {
  const [sortKey, setSortKey] = useState('');
  const [sortDir, setSortDir] = useState<'asc'|'desc'>('asc');

  const sorted = useMemo(() => {
    if (!sortKey) return rows;
    return [...rows].sort((a, b) => {
      const va = a[sortKey]; const vb = b[sortKey];
      const na = parseFloat(String(va)); const nb = parseFloat(String(vb));
      if (!isNaN(na) && !isNaN(nb)) return sortDir === 'asc' ? na - nb : nb - na;
      return sortDir === 'asc' ? String(va).localeCompare(String(vb)) : String(vb).localeCompare(String(va));
    });
  }, [rows, sortKey, sortDir]);

  const toggleSort = (key: string) => {
    if (sortKey === key) setSortDir(d => d === 'asc' ? 'desc' : 'asc');
    else { setSortKey(key); setSortDir('asc'); }
  };

  if (!rows.length) return <p className="text-gray-400 text-xs p-4 italic">{emptyText}</p>;

  return (
    <div className="border border-gray-200 rounded-xl overflow-hidden">
      <div className="overflow-auto max-h-80 custom-scrollbar">
        <table className="w-full text-xs border-collapse">
          <thead>
            <tr className="bg-gray-50 sticky top-0 z-10">
              {cols.map(c => (
                <th
                  key={c.key}
                  onClick={() => toggleSort(c.key)}
                  className={`px-3 py-2.5 text-left font-bold text-gray-500 uppercase text-[10px] tracking-wide cursor-pointer select-none whitespace-nowrap border-b border-gray-200 hover:bg-gray-100 ${c.className || ''}`}
                >
                  {c.label}
                  {sortKey === c.key && <span className="ml-1">{sortDir === 'asc' ? '↑' : '↓'}</span>}
                </th>
              ))}
            </tr>
          </thead>
          <tbody className="divide-y divide-gray-100 bg-white">
            {sorted.map((row, i) => {
              const isActive = activeRowKey && row[activeRowKey] === activeRowValue;
              return (
                <tr
                  key={i}
                  onClick={() => onRowClick?.(row)}
                  className={`transition-colors
                    ${onRowClick ? 'cursor-pointer hover:bg-gray-50' : ''}
                    ${isActive ? 'bg-pink-50 shadow-[inset_3px_0_0_0_#E6007E]' : ''}`}
                >
                  {cols.map(c => (
                    <td key={c.key} className={`px-3 py-2 whitespace-nowrap ${c.className || ''}`}>
                      {c.render ? c.render(row) : row[c.key] ?? '-'}
                    </td>
                  ))}
                </tr>
              );
            })}
          </tbody>
          {footer && (
            <tfoot className="bg-gray-50 font-semibold text-xs border-t-2 border-gray-200">
              {footer}
            </tfoot>
          )}
        </table>
      </div>
    </div>
  );
};

// ── Section card ─────────────────────────────────────────────────────────────

const SectionCard: React.FC<{ title: string; children: React.ReactNode; className?: string; action?: React.ReactNode }> = ({ title, children, className = '', action }) => (
  <div className={`bg-white rounded-xl border border-gray-200 shadow-sm ${className}`}>
    <div className="px-5 py-4 border-b border-gray-100 flex items-center justify-between">
      <h3 className="text-sm font-bold text-gray-700">{title}</h3>
      {action}
    </div>
    <div className="p-5">{children}</div>
  </div>
);

// ── Upload card ───────────────────────────────────────────────────────────────

interface FileInputProps {
  label: string; hint: string; accept: string;
  file: File | null; onFile: (f: File | null) => void;
  required?: boolean; color?: string;
}

const FileInputCard: React.FC<FileInputProps> = ({ label, hint, accept, file, onFile, required, color = 'pink' }) => {
  const ref = useRef<HTMLInputElement>(null);
  const colorMap: Record<string, string> = {
    pink:   'border-pink-200 bg-pink-50 hover:border-pink-400 text-pink-700',
    blue:   'border-blue-200 bg-blue-50 hover:border-blue-400 text-blue-700',
    teal:   'border-teal-200 bg-teal-50 hover:border-teal-400 text-teal-700',
  };
  const cls = colorMap[color] || colorMap.pink;
  return (
    <div>
      <div className="flex items-center gap-1.5 mb-1.5">
        <span className="text-xs font-bold text-gray-700">{label}</span>
        {required && <span className="text-[9px] text-red-500 font-bold uppercase">obrigatório</span>}
      </div>
      <button
        type="button"
        onClick={() => ref.current?.click()}
        className={`w-full border-2 border-dashed rounded-xl px-4 py-3 text-left transition-all ${cls}`}
      >
        <div className="flex items-center gap-2">
          {file ? (
            <>
              <svg className="w-4 h-4 text-emerald-500 flex-shrink-0" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                <path strokeLinecap="round" strokeLinejoin="round" strokeWidth="2" d="M9 12l2 2 4-4m6 2a9 9 0 11-18 0 9 9 0 0118 0z"/>
              </svg>
              <span className="text-xs font-semibold text-emerald-700 truncate">{file.name}</span>
              <button type="button" onClick={e => { e.stopPropagation(); onFile(null); if (ref.current) ref.current.value = ''; }}
                className="ml-auto text-gray-400 hover:text-red-500 flex-shrink-0">
                <svg className="w-3.5 h-3.5" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                  <path strokeLinecap="round" strokeLinejoin="round" strokeWidth="2" d="M6 18L18 6M6 6l12 12"/>
                </svg>
              </button>
            </>
          ) : (
            <>
              <svg className="w-4 h-4 flex-shrink-0" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                <path strokeLinecap="round" strokeLinejoin="round" strokeWidth="2" d="M7 16a4 4 0 01-.88-7.903A5 5 0 1115.9 6L16 6a5 5 0 011 9.9M15 13l-3-3m0 0l-3 3m3-3v12"/>
              </svg>
              <span className="text-xs">{hint}</span>
            </>
          )}
        </div>
      </button>
      <input ref={ref} type="file" accept={accept} className="hidden"
        onChange={e => onFile(e.target.files?.[0] || null)} />
    </div>
  );
};

// ── Main component ────────────────────────────────────────────────────────────

export const AnalysisPanel: React.FC = () => {

  // ── File state ──────────────────────────────────────────────────────────────
  const [corpFile,  setCorpFile]  = useState<File | null>(null);
  const [sdsFile,   setSdsFile]   = useState<File | null>(null);
  const [nddFile,   setNddFile]   = useState<File | null>(null);

  // ── Data state ──────────────────────────────────────────────────────────────
  const [corpRaw,   setCorpRaw]   = useState<any[][]>([]);
  const [sdsMap,    setSdsMap]    = useState<Map<string, any[]>>(new Map());
  const [nddMap,    setNddMap]    = useState<Map<string, any[]>>(new Map());

  // ── UI state ────────────────────────────────────────────────────────────────
  const [loading,   setLoading]   = useState(false);
  const [error,     setError]     = useState('');
  const [analyzed,  setAnalyzed]  = useState(false);

  // ── Settings ─────────────────────────────────────────────────────────────────
  const [attThreshold, setAttThreshold] = useState(10);
  const [errThreshold, setErrThreshold] = useState(20);
  const [backupCounts, setBackupCounts] = useState<Record<string, number>>({});
  const [excludedModels, setExcludedModels] = useState<Set<string>>(new Set());

  // ── Filters ──────────────────────────────────────────────────────────────────
  const [activeFilter,   setActiveFilter]   = useState<'attention'|'error'|'mismatch'|null>(null);
  const [locationFilter, setLocationFilter] = useState<string | null>(null);
  const [modelFilter,    setModelFilter]    = useState<string | null>(null);

  // ── Detail tab ───────────────────────────────────────────────────────────────
  const [detailTab, setDetailTab] = useState<string>('comparison');

  // ── Analysis engine ──────────────────────────────────────────────────────────

  const analysis = useMemo(() => {
    if (!corpRaw.length || !sdsMap.size) return null;

    const now = new Date();

    // Apply model exclusions
    let baseCorp = corpRaw.filter(r => !excludedModels.has(String(r[COL.corp.model] || '')));

    // Pre-calculate serial sets for attention/error/mismatch (over ALL active items)
    const computeBackupSerials = (data: any[][]): Set<string> => {
      const backupSet = new Set<string>();
      const models = [...new Set(data.map(r => String(r[COL.corp.model] || '')))];
      models.forEach(m => {
        const devs = data.filter(r => String(r[COL.corp.model] || '') === m);
        let toAssign = backupCounts[m] || 0;
        devs.forEach(d => {
          if (toAssign > 0) { backupSet.add(normalizeSerial(d[COL.corp.serial])); toAssign--; }
        });
      });
      return backupSet;
    };

    const allBackupSerials = computeBackupSerials(baseCorp);
    const allActive = baseCorp.filter(r => !allBackupSerials.has(normalizeSerial(r[COL.corp.serial])));

    const sets = { attention: new Set<string>(), error: new Set<string>(), mismatch: new Set<string>() };

    allActive.forEach(cr => {
      const cs = normalizeSerial(cr[COL.corp.serial]);
      if (!cs) return;
      const si = findByPrefix(sdsMap, cs);
      if (si) {
        const d = parseUpdateDate(si[COL.sds.lastUpdate]);
        if (d) {
          const days = (now.getTime() - d.getTime()) / 86400000;
          if (days > 7) sets.attention.add(cs);
        }
        const ni = findByPrefix(nddMap, cs);
        if (ni) {
          let pct = cleanNumber(ni[COL.ndd.percentageDifference]);
          if (isNaN(pct) || pct === 0) {
            const ref = cleanNumber(ni[COL.ndd.counterReference]);
            const acc = cleanNumber(ni[COL.ndd.accountedPages]);
            pct = ref > 0 ? Math.abs(ref - acc) / ref * 100 : acc > 0 ? 100 : 0;
          }
          if (pct > errThreshold) { sets.error.add(cs); sets.mismatch.add(cs); }
          else if (pct > attThreshold) { sets.attention.add(cs); sets.mismatch.add(cs); }
        }
      }
    });

    // Uncontracted in error set
    sdsMap.forEach((_, sk) => {
      const inCorp = corpRaw.some(cr => {
        const cs = normalizeSerial(cr[COL.corp.serial]);
        return cs && (sk.startsWith(cs) || cs.startsWith(sk));
      });
      if (!inCorp) sets.error.add(sk);
    });

    // Apply filters
    let filtered = baseCorp;
    if (activeFilter) {
      const fs = sets[activeFilter];
      filtered = filtered.filter(r => fs.has(normalizeSerial(r[COL.corp.serial])));
    }
    if (locationFilter) filtered = filtered.filter(r => String(r[COL.corp.location] || '') === locationFilter);
    if (modelFilter)    filtered = filtered.filter(r => String(r[COL.corp.model]    || '') === modelFilter);

    const filteredBackupSerials = computeBackupSerials(filtered);
    const filteredActive = filtered.filter(r => !filteredBackupSerials.has(normalizeSerial(r[COL.corp.serial])));

    // Build detail lists
    const unmonitored: any[]  = [];
    const lateReports: any[]  = [];
    const comparison: CompRow[] = [];
    const locationDetails: Record<string, LocationDetail> = {};
    const modelDetails:    Record<string, Omit<ModelDetail, 'backups'|'active'|'monitoredActive'|'monitoringPercentage'>> = {};
    const manufacturerDetails: Record<string, number> = {};

    // Location/model pass over ALL filtered (including backups)
    filtered.forEach(cr => {
      const cs   = normalizeSerial(cr[COL.corp.serial]);
      const loc  = String(cr[COL.corp.location] || 'N/A');
      const mdl  = String(cr[COL.corp.model]    || 'N/A');
      const si   = findByPrefix(sdsMap, cs);
      const ni   = findByPrefix(nddMap, cs);
      const isBackup = filteredBackupSerials.has(cs);

      if (!locationDetails[loc]) locationDetails[loc] = { total: 0, monitored: 0, backups: 0, billed: 0 };
      locationDetails[loc].total++;
      if (si) locationDetails[loc].monitored++;
      if (isBackup) locationDetails[loc].backups++;
      if (ni) locationDetails[loc].billed++;

      if (!modelDetails[mdl]) modelDetails[mdl] = { total: 0, monitored: 0, billed: 0 };
      modelDetails[mdl].total++;
      if (si) modelDetails[mdl].monitored++;
      if (ni) modelDetails[mdl].billed++;

      if (si) {
        const mfr = String(si[COL.sds.manufacturer] || 'N/A');
        manufacturerDetails[mfr] = (manufacturerDetails[mfr] || 0) + 1;
      }
    });

    // Active-only pass for comparison/unmonitored
    filteredActive.forEach(cr => {
      const cs = normalizeSerial(cr[COL.corp.serial]);
      if (!cs) return;
      const si = findByPrefix(sdsMap, cs);
      if (si) {
        let lastStr = si[COL.sds.lastUpdate];
        let days = Infinity;
        const d = parseUpdateDate(lastStr);
        if (d) {
          days = (now.getTime() - d.getTime()) / 86400000;
          lastStr = d.toLocaleDateString('pt-BR');
        }
        const row: CompRow = {
          serial: cs, corporateLocation: String(cr[COL.corp.location] || '-'),
          sdsModel: String(si[COL.sds.model] || '-'), sdsHostname: String(si[COL.sds.hostname] || '-'),
          sdsIp: String(si[COL.sds.ip] || '-'), sdsLastUpdate: String(lastStr || '-'),
          sdsCounter: si[COL.sds.counter], status: 'ok',
        };
        if (days > 7) { row.status = 'Atenção'; row.days = Math.floor(days); lateReports.push({ ...row }); }
        comparison.push(row);
      } else {
        unmonitored.push({ serial: cs, model: String(cr[COL.corp.model] || '-'), location: String(cr[COL.corp.location] || '-') });
      }
    });

    // NDD enrichment
    const nddOnly: any[] = [];
    const issues:  any[] = [];
    if (nddMap.size > 0) {
      nddMap.forEach((ni, nSerial) => {
        const compItem = comparison.find(c => nSerial.startsWith(c.serial) || c.serial.startsWith(nSerial));
        if (compItem) {
          const ref = cleanNumber(ni[COL.ndd.counterReference]);
          const acc = cleanNumber(ni[COL.ndd.accountedPages]);
          let pct   = cleanNumber(ni[COL.ndd.percentageDifference]);
          if (isNaN(pct) || pct === 0) pct = ref > 0 ? Math.abs(ref - acc) / ref * 100 : acc > 0 ? 100 : 0;
          compItem.nddAccountedPages  = acc;
          compItem.nddCounterReference = ref;
          compItem.percentageDifference = pct.toFixed(2);
          if (pct > errThreshold) { issues.push({ serial: nSerial, pct: pct.toFixed(2), ref, acc, type: 'Erro' }); compItem.status = 'Erro'; }
          else if (pct > attThreshold) { issues.push({ serial: nSerial, pct: pct.toFixed(2), ref, acc, type: 'Atenção' }); if (compItem.status !== 'Erro') compItem.status = 'Atenção'; }
        } else {
          const activeSerials = new Set(filteredActive.map(r => normalizeSerial(r[COL.corp.serial])));
          if (activeSerials.has(nSerial)) {
            nddOnly.push({ serial: nSerial, lastCounter: ni[COL.ndd.counterReference], lastDate: ni[COL.ndd.lastCounterDate], hostname: ni[COL.ndd.hostname] });
          }
        }
      });
    }

    // Uncontracted
    const uncontracted: any[] = [];
    sdsMap.forEach((si, sk) => {
      const found = corpRaw.some(cr => {
        const cs = normalizeSerial(cr[COL.corp.serial]);
        return cs && (sk.startsWith(cs) || cs.startsWith(sk));
      });
      if (!found) uncontracted.push({ serial: sk, model: String(si[COL.sds.model] || '-'), ip: String(si[COL.sds.ip] || '-') });
    });

    // Model details with backup adjustment
    const modelFinal: Record<string, ModelDetail> = {};
    Object.entries(modelDetails).forEach(([mdl, d]) => {
      const backups = backupCounts[mdl] || 0;
      const active  = Math.max(d.total - backups, 0);
      // Count monitored backups (backups that appear in SDS)
      const monitoredBackups = baseCorp.filter(r =>
        String(r[COL.corp.model] || '') === mdl &&
        allBackupSerials.has(normalizeSerial(r[COL.corp.serial])) &&
        findByPrefix(sdsMap, normalizeSerial(r[COL.corp.serial]))
      ).length;
      const monitoredActive = Math.max(d.monitored - monitoredBackups, 0);
      modelFinal[mdl] = {
        ...d, backups, active, monitoredActive,
        monitoringPercentage: active > 0 ? ((monitoredActive / active) * 100).toFixed(2) : '0.00',
      };
    });

    // Summary
    const monitoredOk = comparison.filter(c => c.status === 'ok').length;
    const otherAttention = sets.attention.size - [...sets.mismatch].filter(s => sets.attention.has(s)).length;
    const otherError     = sets.error.size     - [...sets.mismatch].filter(s => sets.error.has(s)).length;

    const monPct = allActive.length > 0
      ? (activeFilter
          ? (sets[activeFilter].size / allActive.length * 100)
          : (comparison.length / allActive.length * 100)
        ).toFixed(2)
      : '0.00';

    return {
      summary: {
        total: filtered.length, backup: filteredBackupSerials.size,
        monitored: comparison.length, unmonitored: unmonitored.length,
        attention: sets.attention.size, error: sets.error.size, mismatch: sets.mismatch.size,
        monitoredOk, otherAttention, otherError, monPct,
      },
      unmonitored, lateReports, comparison, nddOnly, issues, uncontracted,
      locationDetails, modelFinal, manufacturerDetails,
      sets,
      hasNdd: nddMap.size > 0,
    };
  }, [corpRaw, sdsMap, nddMap, backupCounts, excludedModels, activeFilter, locationFilter, modelFilter, attThreshold, errThreshold]);

  // ── Handlers ─────────────────────────────────────────────────────────────────

  const handleAnalyze = useCallback(async () => {
    if (!corpFile || !sdsFile) return;
    setLoading(true); setError(''); setAnalyzed(false);
    try {
      const [cr, sr, nr] = await Promise.all([
        readRawExcel(corpFile, 4),
        readRawExcel(sdsFile,  1),
        nddFile ? readRawExcel(nddFile, 1) : Promise.resolve([]),
      ]);
      setCorpRaw(cr);
      setSdsMap(new Map(sr.map(r => [normalizeSerial(r[COL.sds.serial]), r])));
      setNddMap(new Map(nr.map(r => [normalizeSerial(r[COL.ndd.serial]), r])));
      setBackupCounts({}); setExcludedModels(new Set());
      setActiveFilter(null); setLocationFilter(null); setModelFilter(null);
      setAnalyzed(true);
    } catch (e: any) {
      setError(e.message || 'Erro ao processar arquivos.');
    } finally {
      setLoading(false);
    }
  }, [corpFile, sdsFile, nddFile]);

  const clearFilter = () => { setActiveFilter(null); setLocationFilter(null); setModelFilter(null); };

  const handleFilterCard = (f: 'attention'|'error'|'mismatch') => {
    setLocationFilter(null); setModelFilter(null);
    setActiveFilter(prev => prev === f ? null : f);
  };

  // ── Unique models for device management ──────────────────────────────────────
  const corpModels = useMemo(() => [...new Set(corpRaw.map(r => String(r[COL.corp.model] || '')))].filter(Boolean).sort(), [corpRaw]);

  const canAnalyze = !!corpFile && !!sdsFile && !loading;

  // ── Donut center render ───────────────────────────────────────────────────────
  const RADIAN = Math.PI / 180;
  const renderLabel = ({ cx, cy, midAngle, innerRadius, outerRadius, percent }: any) => {
    if (percent < 0.06) return null;
    const r = innerRadius + (outerRadius - innerRadius) * 0.5;
    return <text x={cx + r * Math.cos(-midAngle * RADIAN)} y={cy + r * Math.sin(-midAngle * RADIAN)}
      fill="white" textAnchor="middle" dominantBaseline="central" style={{ fontSize: 10, fontWeight: 700 }}>
      {`${(percent * 100).toFixed(0)}%`}
    </text>;
  };

  // ── Render ────────────────────────────────────────────────────────────────────

  return (
    <div className="flex-grow overflow-auto custom-scrollbar bg-gray-50">
      <div className="p-4 max-w-screen-2xl mx-auto space-y-4">

        {/* ── Upload card ── */}
        <div className="bg-white rounded-xl border border-gray-200 shadow-sm p-5">
          <div className="mb-4">
            <h2 className="text-base font-extrabold text-gray-800">
              <span className="bg-gradient-to-r from-pink-600 to-purple-600 bg-clip-text text-transparent">
                Painel Pro — Análise de Contrato
              </span>
            </h2>
            <p className="text-xs text-gray-500 mt-0.5">
              Cruzamento de Corporate × SDS × NDD por número de série. Corporate e SDS são obrigatórios.
            </p>
          </div>

          <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-4">
            <FileInputCard label="1. Corporate (Contrato)"  hint="Selecionar .xls/.xlsx/.csv" accept=".csv,.xlsx,.xlsb,.xls" file={corpFile} onFile={setCorpFile} required color="pink" />
            <FileInputCard label="2. SDS (Monitoramento)"   hint="Selecionar .csv/.xlsx"      accept=".csv,.xlsx,.xlsb,.xls" file={sdsFile}  onFile={setSdsFile}  required color="blue" />
            <FileInputCard label="3. NDD (Bilhetagem)"      hint="Selecionar .csv (opcional)" accept=".csv,.xlsx,.xlsb,.xls" file={nddFile}  onFile={setNddFile}  color="teal" />

            <div className="flex flex-col gap-3">
              <div className="flex gap-3">
                <div className="flex-1">
                  <label className="text-[10px] font-bold text-gray-500 uppercase block mb-1">Atenção (%)</label>
                  <input type="number" min={0} max={100} value={attThreshold}
                    onChange={e => setAttThreshold(+e.target.value)}
                    className="w-full border border-gray-300 rounded-lg px-2 py-1.5 text-sm focus:outline-none focus:ring-2 focus:ring-pink-300" />
                </div>
                <div className="flex-1">
                  <label className="text-[10px] font-bold text-gray-500 uppercase block mb-1">Problema (%)</label>
                  <input type="number" min={0} max={100} value={errThreshold}
                    onChange={e => setErrThreshold(+e.target.value)}
                    className="w-full border border-gray-300 rounded-lg px-2 py-1.5 text-sm focus:outline-none focus:ring-2 focus:ring-pink-300" />
                </div>
              </div>
              <button
                disabled={!canAnalyze}
                onClick={handleAnalyze}
                className="mt-auto w-full py-2.5 px-4 rounded-xl font-bold text-white text-sm transition-all
                  bg-gradient-to-r from-pink-600 to-purple-600 hover:from-pink-700 hover:to-purple-700
                  disabled:opacity-40 disabled:cursor-not-allowed shadow-sm"
              >
                {loading ? 'Analisando...' : 'Analisar Dados'}
              </button>
            </div>
          </div>

          {error && <p className="mt-3 text-xs text-red-600 font-semibold bg-red-50 border border-red-200 rounded-lg px-3 py-2">{error}</p>}
        </div>

        {/* ── Results ── */}
        {analyzed && analysis && (
          <>
            {/* KPI Summary */}
            <div className="grid grid-cols-2 sm:grid-cols-4 lg:grid-cols-8 gap-3">
              <KpiCard label="Total Contrato"     value={analysis.summary.total}       color="text-gray-800" />
              <KpiCard label="Em Backup"          value={analysis.summary.backup}      color="text-gray-500" />
              <KpiCard label="Monitorados"        value={analysis.summary.monitored}   color="text-emerald-600" />
              <KpiCard label="Não Monitorados"    value={analysis.summary.unmonitored} color="text-gray-600" />
              <KpiCard label="Atenção"            value={analysis.summary.attention}   color="text-amber-600"
                onClick={() => handleFilterCard('attention')} active={activeFilter === 'attention'} />
              <KpiCard label="Problemas"          value={analysis.summary.error}       color="text-red-600"
                onClick={() => handleFilterCard('error')} active={activeFilter === 'error'} />
              <KpiCard label="Diverg. Contadores" value={analysis.summary.mismatch}   color="text-blue-600"
                onClick={() => handleFilterCard('mismatch')} active={activeFilter === 'mismatch'} />
              <KpiCard
                label={activeFilter === 'attention' ? '% em Atenção' : activeFilter === 'error' ? '% c/ Problemas' : '% Monitoramento'}
                value={`${analysis.summary.monPct.replace('.', ',')}%`}
                color="text-transparent"
              />
            </div>

            {(activeFilter || locationFilter || modelFilter) && (
              <div className="flex items-center gap-2 bg-pink-50 border border-pink-200 rounded-xl px-4 py-2">
                <svg className="w-3.5 h-3.5 text-pink-600 flex-shrink-0" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                  <path strokeLinecap="round" strokeLinejoin="round" strokeWidth="2" d="M3 4a1 1 0 011-1h16a1 1 0 011 1v2a1 1 0 01-.293.707L13 13.414V19a1 1 0 01-.553.894l-4 2A1 1 0 017 21v-7.586L3.293 6.707A1 1 0 013 6V4z"/>
                </svg>
                <span className="text-xs text-pink-700 font-semibold flex-grow">
                  Filtro ativo:&nbsp;
                  {activeFilter && <span className="capitalize">{activeFilter}</span>}
                  {locationFilter && <span>Localidade: {locationFilter}</span>}
                  {modelFilter && <span>Modelo: {modelFilter}</span>}
                </span>
                <button onClick={clearFilter} className="text-xs font-bold text-pink-600 hover:text-pink-800 bg-white px-2 py-0.5 rounded-lg border border-pink-200">
                  Limpar filtro
                </button>
              </div>
            )}

            {/* Charts row */}
            <div className="grid grid-cols-1 lg:grid-cols-2 gap-4">
              <SectionCard title="Composição do Parque">
                <div className="relative h-64">
                  <ResponsiveContainer width="100%" height="100%">
                    <PieChart>
                      <Pie
                        data={[
                          { name: 'Monitorados OK',       value: analysis.summary.monitoredOk,   fill: CHART_COLORS.ok },
                          { name: 'Atenção (Outros)',     value: analysis.summary.otherAttention, fill: CHART_COLORS.attention },
                          { name: 'Problemas (Outros)',   value: analysis.summary.otherError,    fill: CHART_COLORS.error },
                          { name: 'Diverg. Contadores',   value: analysis.summary.mismatch,      fill: CHART_COLORS.mismatch },
                          { name: 'Não Monitorados',      value: analysis.summary.unmonitored,   fill: CHART_COLORS.unmonitored },
                          { name: 'Backup',               value: analysis.summary.backup,        fill: CHART_COLORS.backup },
                        ].filter(d => d.value > 0)}
                        cx="40%" cy="50%" innerRadius={55} outerRadius={88}
                        dataKey="value" labelLine={false} label={renderLabel}
                      >
                        {[CHART_COLORS.ok, CHART_COLORS.attention, CHART_COLORS.error, CHART_COLORS.mismatch, CHART_COLORS.unmonitored, CHART_COLORS.backup]
                          .map((c, i) => <Cell key={i} fill={c} />)}
                      </Pie>
                      <Tooltip content={<ChartTip />} />
                    </PieChart>
                  </ResponsiveContainer>
                  <div className="absolute inset-0 flex flex-col items-start justify-center pl-[calc(40%-44px)] pointer-events-none">
                    <span className="text-2xl font-extrabold text-gray-800 leading-none">{analysis.summary.total}</span>
                    <span className="text-[9px] text-gray-400 mt-0.5">equipamentos</span>
                  </div>
                  {/* Legend */}
                  <div className="absolute right-0 top-1/2 -translate-y-1/2 flex flex-col gap-1.5">
                    {[
                      { l: 'Monitorados OK',     c: CHART_COLORS.ok,          v: analysis.summary.monitoredOk },
                      { l: 'Atenção',            c: CHART_COLORS.attention,   v: analysis.summary.otherAttention },
                      { l: 'Problemas',          c: CHART_COLORS.error,       v: analysis.summary.otherError },
                      { l: 'Diverg.',            c: CHART_COLORS.mismatch,    v: analysis.summary.mismatch },
                      { l: 'Não Monitorados',    c: CHART_COLORS.unmonitored, v: analysis.summary.unmonitored },
                      { l: 'Backup',             c: CHART_COLORS.backup,      v: analysis.summary.backup },
                    ].filter(d => d.v > 0).map(d => (
                      <div key={d.l} className="flex items-center gap-1.5">
                        <div className="w-2.5 h-2.5 rounded-full flex-shrink-0" style={{ backgroundColor: d.c }} />
                        <span className="text-[10px] text-gray-600">{d.l} ({d.v})</span>
                      </div>
                    ))}
                  </div>
                </div>
              </SectionCard>

              <SectionCard title="Dispositivos por Fabricante">
                {Object.keys(analysis.manufacturerDetails).length > 0 ? (
                  <div className="h-64 flex justify-center">
                    <ResponsiveContainer width="90%" height="100%">
                      <PieChart>
                        <Pie data={Object.entries(analysis.manufacturerDetails).map(([name, value]) => ({ name, value }))}
                          cx="40%" cy="50%" outerRadius={88} dataKey="value" label={renderLabel} labelLine={false}>
                          {Object.keys(analysis.manufacturerDetails).map((_, i) => <Cell key={i} fill={MFR_COLORS[i % MFR_COLORS.length]} />)}
                        </Pie>
                        <Tooltip content={<ChartTip />} />
                      </PieChart>
                    </ResponsiveContainer>
                  </div>
                ) : (
                  <p className="text-xs text-gray-400 italic p-4">Nenhum dado de fabricante disponível.</p>
                )}
              </SectionCard>
            </div>

            {/* Device management */}
            <SectionCard title="Gerenciamento de Modelos">
              <p className="text-xs text-gray-500 mb-3">Configure backups por modelo e exclua modelos da análise.</p>
              <div className="border border-gray-200 rounded-xl overflow-hidden">
                <div className="overflow-auto max-h-60 custom-scrollbar">
                  <table className="w-full text-xs border-collapse">
                    <thead>
                      <tr className="bg-gray-50 sticky top-0">
                        <th className="px-3 py-2.5 text-left text-[10px] font-bold text-gray-500 uppercase border-b border-gray-200">Modelo</th>
                        <th className="px-3 py-2.5 text-left text-[10px] font-bold text-gray-500 uppercase border-b border-gray-200">Qtd. Total</th>
                        <th className="px-3 py-2.5 text-left text-[10px] font-bold text-gray-500 uppercase border-b border-gray-200">Backups</th>
                        <th className="px-3 py-2.5 text-center text-[10px] font-bold text-gray-500 uppercase border-b border-gray-200">Incluir</th>
                      </tr>
                    </thead>
                    <tbody className="divide-y divide-gray-100 bg-white">
                      {corpModels.map(m => {
                        const total = corpRaw.filter(r => String(r[COL.corp.model] || '') === m).length;
                        const included = !excludedModels.has(m);
                        return (
                          <tr key={m} className="hover:bg-gray-50">
                            <td className="px-3 py-2 font-medium text-gray-800">{m}</td>
                            <td className="px-3 py-2 text-gray-600">{total}</td>
                            <td className="px-3 py-2">
                              <input type="number" min={0} max={total}
                                value={backupCounts[m] || 0}
                                onChange={e => {
                                  let v = Math.max(0, Math.min(+e.target.value, total));
                                  setBackupCounts(prev => ({ ...prev, [m]: v }));
                                }}
                                className="w-20 border border-gray-300 rounded-lg px-2 py-1 text-xs focus:outline-none focus:ring-1 focus:ring-pink-300" />
                            </td>
                            <td className="px-3 py-2 text-center">
                              <input type="checkbox" checked={included}
                                onChange={e => setExcludedModels(prev => {
                                  const s = new Set(prev);
                                  e.target.checked ? s.delete(m) : s.add(m);
                                  return s;
                                })}
                                className="w-4 h-4 rounded text-pink-600 border-gray-300 focus:ring-pink-400" />
                            </td>
                          </tr>
                        );
                      })}
                    </tbody>
                  </table>
                </div>
              </div>
            </SectionCard>

            {/* Location + Model reports side by side */}
            <div className="grid grid-cols-1 lg:grid-cols-2 gap-4">
              <SectionCard title="Dispositivos por Localidade">
                <SortableTable
                  cols={[
                    { key: 'location',   label: 'Localidade', className: 'max-w-[180px] truncate font-medium text-gray-800' },
                    { key: 'total',      label: 'Total',      className: 'text-center font-bold' },
                    { key: 'monitored',  label: 'Monitorados',className: 'text-center text-emerald-600 font-semibold',
                      render: r => <span className="text-emerald-600 font-semibold">{r.monitored}</span> },
                    { key: 'unmonitored',label: 'Não Monit.', className: 'text-center text-gray-600 font-semibold' },
                    { key: 'backups',    label: 'Backups',    className: 'text-center text-gray-400' },
                    { key: 'billed',     label: 'Bilhetagem', className: 'text-center',
                      render: r => analysis.hasNdd
                        ? <span className="text-blue-600 font-semibold">{r.billed}</span>
                        : <span className="text-gray-400">N/A</span> },
                  ]}
                  rows={(Object.entries(analysis.locationDetails) as [string, LocationDetail][])
                    .map(([location, d]) => ({ location, ...d, unmonitored: d.total - d.monitored }))
                    .sort((a, b) => b.total - a.total)}
                  onRowClick={r => { setActiveFilter(null); setModelFilter(null); setLocationFilter(prev => prev === r.location ? null : r.location); }}
                  activeRowKey="location" activeRowValue={locationFilter}
                />
              </SectionCard>

              <SectionCard title="Análise por Modelo">
                {(() => {
                  const rows = (Object.entries(analysis.modelFinal) as [string, ModelDetail][])
                    .map(([model, d]) => ({ model, ...d }))
                    .sort((a, b) => b.total - a.total);
                  const totWithBackups    = { total: rows.reduce((s,r)=>s+r.total,0),  monitored: rows.reduce((s,r)=>s+r.monitored,0) };
                  const totWithoutBackups = { total: rows.reduce((s,r)=>s+r.active,0), monitored: rows.reduce((s,r)=>s+r.monitoredActive,0) };
                  return (
                    <SortableTable
                      cols={[
                        { key: 'model', label: 'Modelo', className: 'max-w-[160px] truncate font-medium text-gray-800' },
                        { key: 'total', label: 'Total', className: 'text-center font-bold' },
                        { key: 'monitored', label: 'Monitorados', className: 'text-center',
                          render: r => <span className="text-emerald-600 font-semibold">{r.monitored}</span> },
                        { key: 'unmonitored', label: 'Não Monit.', className: 'text-center text-gray-600',
                          render: r => r.total - r.monitored },
                        { key: 'monitoringPercentage', label: '% Mon. (Ativos)', className: 'text-center text-gray-700' },
                        { key: 'billed', label: 'Bilhetagem', className: 'text-center',
                          render: r => analysis.hasNdd
                            ? <span className="text-blue-600 font-semibold">{r.billed}</span>
                            : <span className="text-gray-400">N/A</span> },
                      ]}
                      rows={rows}
                      onRowClick={r => { setActiveFilter(null); setLocationFilter(null); setModelFilter(prev => prev === r.model ? null : r.model); }}
                      activeRowKey="model" activeRowValue={modelFilter}
                      footer={
                        <>
                          <tr>
                            <td className="px-3 py-2">Total (Sem Backups)</td>
                            <td className="px-3 py-2 text-center">{totWithoutBackups.total}</td>
                            <td className="px-3 py-2 text-center text-emerald-600">{totWithoutBackups.monitored}</td>
                            <td className="px-3 py-2 text-center">{totWithoutBackups.total - totWithoutBackups.monitored}</td>
                            <td className="px-3 py-2 text-center">
                              {totWithoutBackups.total > 0 ? ((totWithoutBackups.monitored/totWithoutBackups.total)*100).toFixed(2) : '0.00'}%
                            </td>
                            <td className="px-3 py-2 text-center">-</td>
                          </tr>
                          <tr>
                            <td className="px-3 py-2">Total (Com Backups)</td>
                            <td className="px-3 py-2 text-center">{totWithBackups.total}</td>
                            <td className="px-3 py-2 text-center text-emerald-600">{totWithBackups.monitored}</td>
                            <td className="px-3 py-2 text-center">{totWithBackups.total - totWithBackups.monitored}</td>
                            <td className="px-3 py-2 text-center">
                              {totWithBackups.total > 0 ? ((totWithBackups.monitored/totWithBackups.total)*100).toFixed(2) : '0.00'}%
                            </td>
                            <td className="px-3 py-2 text-center">-</td>
                          </tr>
                        </>
                      }
                    />
                  );
                })()}
              </SectionCard>
            </div>

            {/* Detail tabs */}
            <SectionCard title="Análise Detalhada (Dispositivos Ativos)">
              {/* Tab buttons */}
              <div className="flex flex-wrap gap-1.5 border-b border-gray-200 mb-4 pb-2">
                {[
                  { id: 'comparison',   label: 'Comparativo Geral' },
                  { id: 'late',         label: `Sem Comunicação +7d (${analysis.lateReports.length})` },
                  { id: 'unmonitored',  label: `Não Monitorados (${analysis.unmonitored.length})` },
                  { id: 'uncontracted', label: `Fora de Contrato (${analysis.uncontracted.length})` },
                  { id: 'issues',       label: `Problemas (${analysis.issues.length})` },
                  ...(analysis.hasNdd ? [{ id: 'nddonly', label: `Apenas Bilhetagem (${analysis.nddOnly.length})` }] : []),
                ].map(t => (
                  <button key={t.id} onClick={() => setDetailTab(t.id)}
                    className={`px-3 py-1.5 text-xs font-semibold rounded-lg border-2 transition-all
                      ${detailTab === t.id
                        ? 'text-pink-600 bg-pink-50 border-pink-400'
                        : 'text-gray-500 border-transparent hover:bg-gray-100'}`}>
                    {t.label}
                  </button>
                ))}
              </div>

              {/* Tab panels */}
              {detailTab === 'comparison' && (
                <SortableTable
                  cols={[
                    { key: 'serial',           label: 'Nº Série',        className: 'font-mono text-gray-800' },
                    { key: 'sdsModel',         label: 'Modelo (SDS)' },
                    { key: 'sdsHostname',      label: 'Hostname' },
                    { key: 'sdsIp',            label: 'IP' },
                    { key: 'corporateLocation',label: 'Localidade',      className: 'max-w-[160px] truncate' },
                    { key: 'sdsLastUpdate',    label: 'Últ. Monitor.' },
                    { key: 'sdsCounter',       label: 'Contador SDS' },
                    { key: 'nddAccountedPages',label: 'Págs. Bilhetadas', render: r => r.nddAccountedPages ?? '-' },
                    { key: 'nddCounterReference', label: 'Cont. Físico (NDD)', render: r => r.nddCounterReference ?? '-' },
                    { key: 'percentageDifference', label: 'Diferença %', render: r => r.percentageDifference ? `${r.percentageDifference}%` : '-' },
                    { key: 'status', label: 'Status', render: r => <StatusBadge status={r.status} /> },
                  ]}
                  rows={analysis.comparison}
                  emptyText="Nenhum dispositivo monitorado encontrado."
                />
              )}
              {detailTab === 'late' && (
                <SortableTable
                  cols={[
                    { key: 'serial',            label: 'Nº Série',   className: 'font-mono' },
                    { key: 'sdsModel',          label: 'Modelo' },
                    { key: 'corporateLocation', label: 'Localidade', className: 'max-w-[180px] truncate' },
                    { key: 'days',              label: 'Dias sem Comunicação',
                      render: r => <span className="text-amber-700 font-semibold">{r.days === Infinity || !r.days ? 'Indeterminado' : r.days}</span> },
                  ]}
                  rows={analysis.lateReports}
                  emptyText="Nenhum dispositivo sem comunicação."
                />
              )}
              {detailTab === 'unmonitored' && (
                <SortableTable
                  cols={[
                    { key: 'serial',   label: 'Nº Série',   className: 'font-mono' },
                    { key: 'model',    label: 'Modelo' },
                    { key: 'location', label: 'Localidade', className: 'max-w-[200px] truncate' },
                  ]}
                  rows={analysis.unmonitored}
                  emptyText="Todos os dispositivos estão monitorados."
                />
              )}
              {detailTab === 'uncontracted' && (
                <SortableTable
                  cols={[
                    { key: 'serial', label: 'Nº Série', className: 'font-mono' },
                    { key: 'model',  label: 'Modelo (SDS)' },
                    { key: 'ip',     label: 'IP' },
                  ]}
                  rows={analysis.uncontracted}
                  emptyText="Nenhum dispositivo fora de contrato."
                />
              )}
              {detailTab === 'issues' && (
                <SortableTable
                  cols={[
                    { key: 'serial', label: 'Nº Série', className: 'font-mono' },
                    { key: 'type',   label: 'Status',   render: r => <StatusBadge status={r.type} /> },
                    { key: 'pct',    label: 'Diferença %',   render: r => `${r.pct}%` },
                    { key: 'ref',    label: 'Contador Físico' },
                    { key: 'acc',    label: 'Págs. Bilhetadas' },
                  ]}
                  rows={analysis.issues}
                  emptyText="Nenhum problema de contador encontrado."
                />
              )}
              {detailTab === 'nddonly' && (
                <SortableTable
                  cols={[
                    { key: 'serial',      label: 'Nº Série',       className: 'font-mono' },
                    { key: 'hostname',    label: 'Host / IP' },
                    { key: 'lastDate',    label: 'Últ. Coleta' },
                    { key: 'lastCounter', label: 'Contador NDD' },
                  ]}
                  rows={analysis.nddOnly}
                  emptyText="Nenhum dispositivo apenas na bilhetagem."
                />
              )}
            </SectionCard>
          </>
        )}

      </div>
    </div>
  );
};
