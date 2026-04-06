import React, { useState, useMemo, useEffect } from 'react';
import * as XLSX from 'xlsx';
import {
  PieChart, Pie, Cell, Tooltip, ResponsiveContainer,
  BarChart, Bar, XAxis, YAxis, CartesianGrid,
} from 'recharts';
import { CepInvalidEntry, CepCorrectionEntry } from '../types';

// ── Types ─────────────────────────────────────────────────────────────────────

export interface StatEntry { name: string; count: number; }

export interface SerialDetail {
  sdsStatus: 'monitored' | 'alert' | 'notMonitored' | 'noData';
  nddStatus: 'monitored' | 'alert' | 'notMonitored' | 'noData';
  billingStatus: 'active' | 'noRecent' | 'never' | null;
  inContract: boolean;
  // Extended detail fields for table/export
  ip?: string;
  hostname?: string;
  logradouro?: string;
  bairro?: string;
  cidade?: string;
  uf?: string;
  cep?: string;
  modelo?: string;
  lastSdsUpdate?: string;
  lastNddUpdate?: string;
  billingStatusText?: string;
  counterValue?: number | null;
  filial?: string;
  site?: string;
  department?: string;
}

export interface LocationBreakdown {
  name: string;
  total: number;
  sds: { monitored: number; alert: number; notMonitored: number; noData: number };
  ndd: { monitored: number; alert: number; notMonitored: number; noData: number };
  billing: { active: number; noRecent: number; never: number };
  situacao: StatEntry[];
  serials: string[];
  serialDetails: Record<string, SerialDetail>;
}

export interface DashboardStats {
  total: number;
  sdsLoaded: boolean;
  nddLoaded: boolean;
  corpLoaded: boolean;
  sds: { monitored: number; alert: number; notMonitored: number; incomplete: number };
  ndd: { monitored: number; alert: number; notMonitored: number };
  billing: { active: number; noRecent: number; never: number };
  corp: { inContract: number; outOfContract: number; ativo: number; inativo: number };
  producing: StatEntry[];
  situacao: StatEntry[];
  tipo: StatEntry[];
  byCidade: StatEntry[];
  byContrato: StatEntry[];
  byModelo: StatEntry[];
  byUf: StatEntry[];
  cepStats: {
    total: number;
    valid: number;
    invalid: number;
    unchecked: number;
    invalidList: CepInvalidEntry[];
  } | null;
  connectionType: StatEntry[];
  locationsByContrato: LocationBreakdown[];
  locationsByCity: LocationBreakdown[];
  allSerialDetails: Record<string, SerialDetail>;
}

// ── Color palette ─────────────────────────────────────────────────────────────

const C = {
  green:   '#22c55e',
  yellow:  '#f59e0b',
  red:     '#ef4444',
  blue:    '#3b82f6',
  purple:  '#8b5cf6',
  slate:   '#94a3b8',
  teal:    '#14b8a6',
  indigo:  '#6366f1',
  orange:  '#f97316',
  pink:    '#ec4899',
};

const BAR_PALETTE = [
  C.blue, C.purple, C.teal, C.orange, C.pink,
  C.indigo, '#84cc16', '#06b6d4', '#d946ef', '#64748b',
];

// ── Helpers ───────────────────────────────────────────────────────────────────

const RADIAN = Math.PI / 180;

const renderPieLabel = ({ cx, cy, midAngle, innerRadius, outerRadius, percent }: any) => {
  if (percent < 0.07) return null;
  const r = innerRadius + (outerRadius - innerRadius) * 0.55;
  const x = cx + r * Math.cos(-midAngle * RADIAN);
  const y = cy + r * Math.sin(-midAngle * RADIAN);
  return (
    <text x={x} y={y} fill="white" textAnchor="middle" dominantBaseline="central"
      style={{ fontSize: 11, fontWeight: 700 }}>
      {`${(percent * 100).toFixed(0)}%`}
    </text>
  );
};

const PieTooltip = ({ active, payload }: any) => {
  if (!active || !payload?.length) return null;
  return (
    <div className="bg-white border border-gray-200 rounded-lg shadow-lg px-3 py-2">
      <p className="text-xs font-bold text-gray-800">{payload[0].name}</p>
      <p className="text-xs text-gray-500">{Number(payload[0].value).toLocaleString('pt-BR')} registros</p>
    </div>
  );
};

const BarTooltip = ({ active, payload, label }: any) => {
  if (!active || !payload?.length) return null;
  return (
    <div className="bg-white border border-gray-200 rounded-lg shadow-lg px-3 py-2">
      {label && <p className="text-xs text-gray-500 mb-1">{label}</p>}
      {payload.map((p: any, i: number) => (
        <p key={i} className="text-xs font-bold" style={{ color: p.fill }}>
          {p.name}: {Number(p.value).toLocaleString('pt-BR')}
        </p>
      ))}
    </div>
  );
};

// ── Health Bar (stacked 3-color progress) ─────────────────────────────────────

interface HealthBarProps {
  monitored: number;
  alert: number;
  notMonitored: number;
  noData: number;
  total: number;
  label?: string;
  showPct?: boolean;
}

const HealthBar: React.FC<HealthBarProps> = ({ monitored, alert, notMonitored, noData, total, label, showPct = true }) => {
  const known = total - noData;
  if (known <= 0) return (
    <div>
      {label && <div className="text-[9px] font-bold text-gray-400 mb-0.5">{label}</div>}
      <div className="h-2 rounded-full bg-gray-100 flex items-center px-1.5">
        <span className="text-[8px] text-gray-400 italic">sem dados</span>
      </div>
    </div>
  );
  const monPct  = monitored    / total * 100;
  const alertPct = alert       / total * 100;
  const notPct  = notMonitored / total * 100;
  const healthPct = Math.round(monitored / known * 100);
  return (
    <div>
      {(label || showPct) && (
        <div className="flex items-center justify-between mb-0.5">
          {label && <span className="text-[9px] font-bold text-gray-500">{label}</span>}
          {showPct && <span className="text-[9px] font-bold text-gray-700">{healthPct}% ok</span>}
        </div>
      )}
      <div className="h-2 rounded-full bg-gray-100 flex overflow-hidden">
        <div className="bg-emerald-500 transition-all duration-500" style={{ width: `${monPct}%` }} />
        <div className="bg-amber-400 transition-all duration-500"   style={{ width: `${alertPct}%` }} />
        <div className="bg-red-400 transition-all duration-500"     style={{ width: `${notPct}%` }} />
      </div>
    </div>
  );
};

// ── Donut with center overlay ─────────────────────────────────────────────────

interface DonutItem { name: string; value: number; color: string; }

interface DonutWithCenterProps {
  data: DonutItem[];
  centerValue: number | string;
  centerLabel: string;
  height?: number;
}

const DonutWithCenter: React.FC<DonutWithCenterProps> = ({ data, centerValue, centerLabel, height = 200 }) => (
  <div className="relative">
    <ResponsiveContainer width="100%" height={height}>
      <PieChart>
        <Pie data={data} cx="50%" cy="50%" innerRadius={52} outerRadius={82}
          dataKey="value" labelLine={false} label={renderPieLabel}>
          {data.map((d, i) => <Cell key={i} fill={d.color} />)}
        </Pie>
        <Tooltip content={<PieTooltip />} />
      </PieChart>
    </ResponsiveContainer>
    <div className="absolute inset-0 flex items-center justify-center pointer-events-none">
      <div className="text-center">
        <div className="text-xl font-extrabold text-gray-800 leading-none">
          {typeof centerValue === 'number' ? centerValue.toLocaleString('pt-BR') : centerValue}
        </div>
        <div className="text-[9px] text-gray-400 mt-0.5 font-medium">{centerLabel}</div>
      </div>
    </div>
  </div>
);

// ── PieLegend ─────────────────────────────────────────────────────────────────

const PieLegend: React.FC<{ items: { name: string; color: string; value: number }[] }> = ({ items }) => (
  <div className="flex flex-wrap justify-center gap-x-4 gap-y-1 mt-2">
    {items.map(d => (
      <div key={d.name} className="flex items-center gap-1.5">
        <div className="w-2.5 h-2.5 rounded-full flex-shrink-0" style={{ backgroundColor: d.color }} />
        <span className="text-[10px] text-gray-600">{d.name} ({d.value.toLocaleString('pt-BR')})</span>
      </div>
    ))}
  </div>
);

// ── ChartCard ─────────────────────────────────────────────────────────────────

const ChartCard: React.FC<{ title: string; children: React.ReactNode; className?: string; action?: React.ReactNode }> = ({ title, children, className = '', action }) => (
  <div className={`bg-white rounded-xl border border-gray-200 shadow-sm flex flex-col ${className}`}>
    <div className="px-4 py-3 border-b border-gray-100 flex items-center justify-between flex-shrink-0">
      <h3 className="text-xs font-bold text-gray-500 uppercase tracking-wider">{title}</h3>
      {action}
    </div>
    <div className="p-4 flex-grow min-h-0">{children}</div>
  </div>
);

// ── KPI Card ──────────────────────────────────────────────────────────────────

interface KpiCardProps {
  label: string;
  value: number;
  sub?: string;
  bgClass: string;
  icon: React.ReactNode;
  healthBar?: HealthBarProps;
  onClick?: () => void;
}

const KpiCard: React.FC<KpiCardProps> = ({ label, value, sub, bgClass, icon, healthBar, onClick }) => (
  <div
    className={`rounded-xl p-4 flex flex-col shadow-sm ${bgClass} ${onClick ? 'cursor-pointer hover:brightness-110 hover:shadow-md active:scale-[.98] transition-all' : ''}`}
    onClick={onClick}
    role={onClick ? 'button' : undefined}
  >
    <div className="flex items-center gap-3 mb-2">
      <div className="flex-shrink-0 w-9 h-9 rounded-lg bg-white/20 flex items-center justify-center text-white">
        {icon}
      </div>
      <div className="min-w-0">
        <div className="text-2xl font-extrabold text-white leading-none">
          {value.toLocaleString('pt-BR')}
        </div>
        <div className="text-[10px] font-semibold text-white/90 leading-tight">{label}</div>
      </div>
    </div>
    {sub && <div className="text-[10px] text-white/60 mb-1.5">{sub}</div>}
    {healthBar && (
      <div className="mt-auto">
        <HealthBar {...healthBar} showPct={false} />
      </div>
    )}
  </div>
);

// ── Icons ─────────────────────────────────────────────────────────────────────

const Ico = {
  equip:   <svg className="w-4 h-4" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path strokeLinecap="round" strokeLinejoin="round" strokeWidth="2" d="M9 3H5a2 2 0 00-2 2v4m6-6h10a2 2 0 012 2v4M9 3v18m0 0h10a2 2 0 002-2V9M9 21H5a2 2 0 01-2-2V9m0 0h18"/></svg>,
  check:   <svg className="w-4 h-4" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path strokeLinecap="round" strokeLinejoin="round" strokeWidth="2" d="M9 12l2 2 4-4m6 2a9 9 0 11-18 0 9 9 0 0118 0z"/></svg>,
  warn:    <svg className="w-4 h-4" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path strokeLinecap="round" strokeLinejoin="round" strokeWidth="2" d="M12 9v2m0 4h.01m-6.938 4h13.856c1.54 0 2.502-1.667 1.732-3L13.732 4c-.77-1.333-2.694-1.333-3.464 0L3.34 16c-.77 1.333.192 3 1.732 3z"/></svg>,
  off:     <svg className="w-4 h-4" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path strokeLinecap="round" strokeLinejoin="round" strokeWidth="2" d="M18.364 18.364A9 9 0 005.636 5.636m12.728 12.728A9 9 0 015.636 5.636m12.728 12.728L5.636 5.636"/></svg>,
  bill:    <svg className="w-4 h-4" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path strokeLinecap="round" strokeLinejoin="round" strokeWidth="2" d="M9 7h6m0 10v-3m-3 3h.01M9 17h.01M9 14h.01M12 14h.01M15 11h.01M12 11h.01M9 11h.01M7 21h10a2 2 0 002-2V5a2 2 0 00-2-2H7a2 2 0 00-2 2v14a2 2 0 002 2z"/></svg>,
  star:    <svg className="w-4 h-4" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path strokeLinecap="round" strokeLinejoin="round" strokeWidth="2" d="M5 3v4M3 5h4M6 17v4m-2-2h4m5-16l2.286 6.857L21 12l-5.714 2.143L13 21l-2.286-6.857L5 12l5.714-2.143L13 3z"/></svg>,
  contract:<svg className="w-4 h-4" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path strokeLinecap="round" strokeLinejoin="round" strokeWidth="2" d="M9 12h6m-6 4h6m2 5H7a2 2 0 01-2-2V5a2 2 0 012-2h5.586a1 1 0 01.707.293l5.414 5.414a1 1 0 01.293.707V19a2 2 0 01-2 2z"/></svg>,
  map:     <svg className="w-4 h-4" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path strokeLinecap="round" strokeLinejoin="round" strokeWidth="2" d="M17.657 16.657L13.414 20.9a1.998 1.998 0 01-2.827 0l-4.244-4.243a8 8 0 1111.314 0z"/><path strokeLinecap="round" strokeLinejoin="round" strokeWidth="2" d="M15 11a3 3 0 11-6 0 3 3 0 016 0z"/></svg>,
  close:   <svg className="w-5 h-5" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path strokeLinecap="round" strokeLinejoin="round" strokeWidth="2" d="M6 18L18 6M6 6l12 12"/></svg>,
  arrow:   <svg className="w-3.5 h-3.5" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path strokeLinecap="round" strokeLinejoin="round" strokeWidth="2" d="M9 5l7 7-7 7"/></svg>,
};

// ── LocationCard ──────────────────────────────────────────────────────────────

interface LocationCardProps {
  loc: LocationBreakdown;
  sdsLoaded: boolean;
  nddLoaded: boolean;
  onClick: () => void;
}

const LocationCard: React.FC<LocationCardProps> = ({ loc, sdsLoaded, nddLoaded, onClick }) => {
  const sdsKnown = loc.total - loc.sds.noData;
  const sdsRate  = sdsKnown > 0 ? Math.round(loc.sds.monitored / sdsKnown * 100) : null;
  const nddKnown = loc.total - loc.ndd.noData;
  const nddRate  = nddKnown > 0 ? Math.round(loc.ndd.monitored / nddKnown * 100) : null;

  const billingTotal = loc.billing.active + loc.billing.noRecent + loc.billing.never;
  const billingRate  = billingTotal > 0 ? Math.round(loc.billing.active / billingTotal * 100) : null;

  const health = sdsRate ?? nddRate ?? 100;
  const borderColor = health >= 80 ? 'border-emerald-200 hover:border-emerald-400'
    : health >= 50 ? 'border-amber-200 hover:border-amber-400'
    : 'border-red-200 hover:border-red-400';

  const badgeColor = health >= 80 ? 'bg-emerald-50 text-emerald-700'
    : health >= 50 ? 'bg-amber-50 text-amber-700'
    : 'bg-red-50 text-red-700';

  return (
    <button
      onClick={onClick}
      className={`bg-white border-2 ${borderColor} rounded-xl p-4 text-left hover:shadow-md transition-all duration-200 group w-full`}
    >
      {/* Header */}
      <div className="flex items-start justify-between gap-2 mb-3">
        <div className="min-w-0">
          <div className="text-[9px] text-gray-400 font-bold uppercase tracking-wider mb-0.5">Contrato/Local</div>
          <div className="text-sm font-extrabold text-gray-800 truncate" title={loc.name}>{loc.name}</div>
        </div>
        <div className="flex-shrink-0 text-right">
          <div className="text-2xl font-extrabold text-blue-600 leading-none">{loc.total.toLocaleString('pt-BR')}</div>
          <div className="text-[9px] text-gray-400">equip.</div>
        </div>
      </div>

      {/* Health badge */}
      {(sdsRate !== null || nddRate !== null) && (
        <div className={`inline-flex items-center gap-1 px-2 py-0.5 rounded-full text-[10px] font-bold mb-2.5 ${badgeColor}`}>
          <div className={`w-1.5 h-1.5 rounded-full ${health >= 80 ? 'bg-emerald-500' : health >= 50 ? 'bg-amber-500' : 'bg-red-500'}`} />
          {health}% monitorado
        </div>
      )}

      {/* SDS health bar */}
      {sdsLoaded && (
        <div className="mb-2">
          <HealthBar
            label="SDS"
            {...loc.sds}
            total={loc.total}
            showPct={false}
          />
        </div>
      )}

      {/* NDD health bar */}
      {nddLoaded && (
        <div className="mb-2">
          <HealthBar
            label="MPS"
            {...loc.ndd}
            total={loc.total}
            showPct={false}
          />
        </div>
      )}

      {/* Billing indicator */}
      {nddLoaded && billingRate !== null && (
        <div className="flex items-center gap-1.5 mt-2 pt-2 border-t border-gray-100">
          <div className="text-[9px] text-gray-500 font-semibold">Bilhetagem:</div>
          <div className="flex-grow h-1.5 rounded-full bg-gray-100 flex overflow-hidden">
            <div className="bg-emerald-500" style={{ width: `${billingRate}%` }} />
            <div className="bg-amber-400" style={{ width: `${loc.billing.noRecent / billingTotal * 100}%` }} />
            <div className="bg-red-400"    style={{ width: `${loc.billing.never   / billingTotal * 100}%` }} />
          </div>
          <div className="text-[9px] font-bold text-gray-700">{billingRate}%</div>
        </div>
      )}

      {/* Footer hint */}
      <div className="flex items-center justify-end gap-0.5 mt-2 text-[9px] text-blue-500 font-semibold opacity-0 group-hover:opacity-100 transition-opacity">
        Ver detalhes {Ico.arrow}
      </div>
    </button>
  );
};

// ── FilterRow (clickable stat row for interactive filtering) ──────────────────

interface FilterRowProps {
  label: string;
  value: number;
  color: string;
  total: number;
  isActive: boolean;
  onClick: () => void;
}

const FilterRow: React.FC<FilterRowProps> = ({ label, value, color, total, isActive, onClick }) => (
  <button
    onClick={onClick}
    className={`flex items-center gap-2 w-full rounded-lg px-2 py-1.5 transition-all text-left border-2 ${
      isActive ? '' : 'border-transparent hover:bg-white/70'
    }`}
    style={isActive ? { borderColor: color, backgroundColor: color + '22' } : {}}
  >
    <div className="w-2.5 h-2.5 rounded-full flex-shrink-0" style={{ backgroundColor: color }} />
    <div className="text-xs text-gray-700 flex-grow">{label}</div>
    <div className="text-xs font-bold text-gray-800">{value.toLocaleString('pt-BR')}</div>
    <div className="text-[10px] text-gray-400 w-8 text-right">
      {total > 0 ? `${Math.round(value / total * 100)}%` : '–'}
    </div>
    {isActive && (
      <svg className="w-3.5 h-3.5 flex-shrink-0" fill="none" stroke="currentColor" viewBox="0 0 24 24" style={{ color }}>
        <path strokeLinecap="round" strokeLinejoin="round" strokeWidth="2.5" d="M5 13l4 4L19 7"/>
      </svg>
    )}
  </button>
);

// ── LocationDetailModal ───────────────────────────────────────────────────────

interface LocationDetailModalProps {
  loc: LocationBreakdown | null;
  sdsLoaded: boolean;
  nddLoaded: boolean;
  onClose: () => void;
}

const PAGE_SIZE = 50;

const SDS_LABEL: Record<string, string> = {
  monitored: 'Monitorado', alert: 'Alerta', notMonitored: 'Não Monitorado', noData: '-',
};
const NDD_LABEL: Record<string, string> = {
  monitored: 'Monitorado', alert: 'Alerta', notMonitored: 'Não Monitorado', noData: '-',
};

const LocationDetailModal: React.FC<LocationDetailModalProps> = ({ loc, sdsLoaded, nddLoaded, onClose }) => {
  const [sdsFilter, setSdsFilter]         = useState<'all' | 'monitored' | 'alert' | 'notMonitored'>('all');
  const [nddFilter, setNddFilter]         = useState<'all' | 'monitored' | 'alert' | 'notMonitored'>('all');
  const [billingFilter, setBillingFilter] = useState<'all' | 'active' | 'noRecent' | 'never'>('all');
  const [showSds, setShowSds]             = useState(true);
  const [showNdd, setShowNdd]             = useState(true);
  const [showBilling, setShowBilling]     = useState(true);
  const [search, setSearch]               = useState('');
  const [page, setPage]                   = useState(1);
  const [showDuplicatesOnly, setShowDuplicatesOnly] = useState(false);
  const [viewMode, setViewMode]           = useState<'chips' | 'table'>('chips');
  const [tableSortKey, setTableSortKey]   = useState('serie');
  const [tableSortDir, setTableSortDir]   = useState<'asc' | 'desc'>('asc');

  // Reset all state when the location changes
  useEffect(() => {
    setSdsFilter('all');
    setNddFilter('all');
    setBillingFilter('all');
    setShowSds(true);
    setShowNdd(true);
    setShowBilling(true);
    setSearch('');
    setPage(1);
    setShowDuplicatesOnly(false);
    setViewMode('chips');
  }, [loc?.name]);

  const uniqueSerials = useMemo(() => (loc ? [...new Set(loc.serials)] : []), [loc]);

  const duplicates = useMemo(() => {
    if (!loc) return [];
    const counts: Record<string, number> = {};
    loc.serials.forEach(s => { counts[s] = (counts[s] || 0) + 1; });
    return Object.keys(counts).filter(k => counts[k] > 1);
  }, [loc]);

  const numMissingSerials = loc ? loc.total - loc.serials.length : 0;

  const filteredSerials = useMemo(() => {
    if (!loc) return [];
    let list = uniqueSerials;
    if (showDuplicatesOnly) {
      list = duplicates;
    }
    return list.filter(serial => {
      const d = loc.serialDetails?.[serial];
      if (sdsFilter     !== 'all' && (!d || d.sdsStatus     !== sdsFilter))     return false;
      if (nddFilter     !== 'all' && (!d || d.nddStatus     !== nddFilter))     return false;
      if (billingFilter !== 'all' && (!d || d.billingStatus !== billingFilter)) return false;
      if (search && !serial.toLowerCase().includes(search.toLowerCase()))       return false;
      return true;
    });
  }, [loc, uniqueSerials, duplicates, showDuplicatesOnly, sdsFilter, nddFilter, billingFilter, search]);

  const totalPages = Math.max(1, Math.ceil(filteredSerials.length / PAGE_SIZE));
  const safePage   = Math.min(page, totalPages);
  const pageSerials = filteredSerials.slice((safePage - 1) * PAGE_SIZE, safePage * PAGE_SIZE);

  const hasFilters = sdsFilter !== 'all' || nddFilter !== 'all' || billingFilter !== 'all' || search !== '';

  const clearFilters = () => {
    setSdsFilter('all'); setNddFilter('all'); setBillingFilter('all');
    setSearch(''); setPage(1);
  };

  const sortedTableRows = useMemo(() => {
    if (!loc) return [];
    const rows = filteredSerials.map(s => ({ serial: s, ...(loc.serialDetails?.[s] || {}) }));
    return rows.sort((a, b) => {
      let va: string | number = '';
      let vb: string | number = '';
      switch (tableSortKey) {
        case 'serie':      va = a.serial;              vb = b.serial;              break;
        case 'modelo':     va = a.modelo || '';         vb = b.modelo || '';        break;
        case 'cidade':     va = a.cidade || '';         vb = b.cidade || '';        break;
        case 'uf':         va = a.uf || '';             vb = b.uf || '';            break;
        case 'bairro':     va = a.bairro || '';         vb = b.bairro || '';        break;
        case 'sds':        va = a.sdsStatus || '';      vb = b.sdsStatus || '';     break;
        case 'ndd':        va = a.nddStatus || '';      vb = b.nddStatus || '';     break;
        case 'billing':    va = a.billingStatus || '';  vb = b.billingStatus || ''; break;
        case 'counter':    va = a.counterValue ?? -1;   vb = b.counterValue ?? -1;  break;
        case 'lastSds':    va = a.lastSdsUpdate || '';  vb = b.lastSdsUpdate || ''; break;
        case 'lastNdd':    va = a.lastNddUpdate || '';  vb = b.lastNddUpdate || ''; break;
        case 'site':       va = a.site || '';           vb = b.site || '';          break;
        case 'department': va = a.department || '';     vb = b.department || '';    break;
        default: break;
      }
      if (va < vb) return tableSortDir === 'asc' ? -1 : 1;
      if (va > vb) return tableSortDir === 'asc' ? 1 : -1;
      return 0;
    });
  }, [filteredSerials, loc, tableSortKey, tableSortDir]);

  // Detect if Site / Department are populated for this location
  const hasSite = useMemo(() =>
    sortedTableRows.some(r => r.site && r.site.trim() !== ''),
  [sortedTableRows]);
  const hasDepartment = useMemo(() =>
    sortedTableRows.some(r => r.department && r.department.trim() !== ''),
  [sortedTableRows]);

  const handleTableSort = (key: string) => {
    if (tableSortKey === key) setTableSortDir(d => d === 'asc' ? 'desc' : 'asc');
    else { setTableSortKey(key); setTableSortDir('asc'); }
  };

  const exportData = (format: 'csv' | 'xlsx') => {
    if (!loc) return;
    const rows = sortedTableRows.map(r => {
      const base: Record<string, string | number> = {
        'Série':               r.serial,
        'Modelo':              r.modelo || '-',
        'Filial':              r.filial || '-',
        'Logradouro':          r.logradouro || '-',
        'Bairro':              r.bairro || '-',
        'Cidade':              r.cidade || '-',
        'UF':                  r.uf || '-',
        'CEP':                 r.cep || '-',
        'IP':                  r.ip || '-',
        'Hostname':            r.hostname || '-',
        'SDS':                 SDS_LABEL[r.sdsStatus] || '-',
        'Últ. Atualiz. SDS':   r.lastSdsUpdate || '-',
        'Contador':            r.counterValue != null ? r.counterValue : '-',
        'MPS':                 NDD_LABEL[r.nddStatus] || '-',
        'Últ. Atualiz. MPS':   r.lastNddUpdate || '-',
        'Bilhetagem':          r.billingStatusText || '-',
      };
      if (hasSite)       base['Site']       = r.site       || '-';
      if (hasDepartment) base['Departamento'] = r.department || '-';
      return base;
    });

    const fileName = `${loc.name.replace(/[^a-zA-Z0-9]/g, '_')}_localidade`;

    if (format === 'xlsx') {
      const ws = XLSX.utils.json_to_sheet(rows);
      const wb = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(wb, ws, 'Dispositivos');
      XLSX.writeFile(wb, `${fileName}.xlsx`);
    } else {
      const headers = Object.keys(rows[0] || {});
      const csvLines = [
        '\uFEFF' + headers.join(';'),
        ...rows.map(r => headers.map(h => {
          const v = String((r as any)[h] ?? '').replace(/"/g, '""');
          return v.includes(';') || v.includes('\n') ? `"${v}"` : v;
        }).join(';')),
      ];
      const blob = new Blob([csvLines.join('\n')], { type: 'text/csv;charset=utf-8;' });
      const url  = URL.createObjectURL(blob);
      const a    = document.createElement('a');
      a.href = url; a.download = `${fileName}.csv`; a.click();
      URL.revokeObjectURL(url);
    }
  };

  const toggleSds = (s: 'monitored' | 'alert' | 'notMonitored') => {
    setSdsFilter(p => p === s ? 'all' : s); setPage(1);
  };
  const toggleNdd = (s: 'monitored' | 'alert' | 'notMonitored') => {
    setNddFilter(p => p === s ? 'all' : s); setPage(1);
  };
  const toggleBilling = (s: 'active' | 'noRecent' | 'never') => {
    setBillingFilter(p => p === s ? 'all' : s); setPage(1);
  };

  // Pagination page number list (max 7 visible)
  const getPageNumbers = (cur: number, total: number): (number | null)[] => {
    if (total <= 7) return Array.from({ length: total }, (_, i) => i + 1);
    if (cur <= 4)            return [1, 2, 3, 4, 5, null, total];
    if (cur >= total - 3)    return [1, null, total - 4, total - 3, total - 2, total - 1, total];
    return [1, null, cur - 1, cur, cur + 1, null, total];
  };

  if (!loc) return null;

  const sdsData = [
    { name: 'Monitorado',     value: loc.sds.monitored,    color: C.green  },
    { name: 'Alerta',         value: loc.sds.alert,        color: C.yellow },
    { name: 'Não Monitorado', value: loc.sds.notMonitored, color: C.red    },
  ].filter(d => d.value > 0);

  const nddChartData = [
    { name: 'Monitorado',     value: loc.ndd.monitored,    color: C.green  },
    { name: 'Alerta',         value: loc.ndd.alert,        color: C.yellow },
    { name: 'Não Monitorado', value: loc.ndd.notMonitored, color: C.red    },
  ].filter(d => d.value > 0);

  const billingItems = [
    { key: 'active',   name: 'Ativa',           value: loc.billing.active,   color: C.green  },
    { key: 'noRecent', name: 'Sem Rec.',         value: loc.billing.noRecent, color: C.yellow },
    { key: 'never',    name: 'Nunca Bilhetado',  value: loc.billing.never,    color: C.red    },
  ].filter(d => d.value > 0);

  const billingTotal = loc.billing.active + loc.billing.noRecent + loc.billing.never;
  const sdsKnown = loc.total - loc.sds.noData;
  const nddKnown = loc.total - loc.ndd.noData;

  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center bg-black/40 backdrop-blur-sm p-4"
      onClick={e => { if (e.target === e.currentTarget) onClose(); }}>
      <div className={`bg-white rounded-2xl shadow-2xl w-full max-h-[92vh] flex flex-col animate-fade-in-up transition-all duration-300 ${viewMode === 'table' ? 'w-[96vw]' : 'max-w-3xl'}`}>

        {/* ── Header ── */}
        <div className="px-6 py-4 border-b border-gray-100 flex items-start justify-between flex-shrink-0">
          <div>
            <div className="text-[10px] text-gray-400 font-bold uppercase tracking-wider">Detalhes do Contrato / Local</div>
            <h2 className="text-xl font-extrabold text-gray-800 mt-0.5">{loc.name}</h2>
            <div className="flex items-center gap-3 mt-1 flex-wrap">
              <span className="text-sm text-gray-500">{loc.total.toLocaleString('pt-BR')} equipamentos</span>
              {uniqueSerials.length > 0 && (
                <span className="text-xs bg-blue-50 text-blue-700 px-2 py-0.5 rounded-full font-semibold">
                  {uniqueSerials.length} seriais únicos
                </span>
              )}
              {numMissingSerials > 0 && (
                <span className="text-xs bg-red-50 text-red-700 px-2 py-0.5 rounded-full font-bold shadow-sm" title="Equipamentos cujas linhas no relatório não possuem um número de série válido">
                  ⚠️ {numMissingSerials} sem identificação
                </span>
              )}
              {duplicates.length > 0 && (
                <button 
                  onClick={() => { setShowDuplicatesOnly(!showDuplicatesOnly); setPage(1); }}
                  className={`text-xs px-2 py-0.5 rounded-full font-bold shadow-sm transition-colors border ${
                    showDuplicatesOnly 
                      ? 'bg-amber-500 text-white border-amber-600' 
                      : 'bg-amber-50 text-amber-700 border-transparent hover:bg-amber-100'
                  }`}
                  title="Clique para filtrar a visualização pelos seriais duplicados"
                >
                  {showDuplicatesOnly ? 'Ocultar duplicados' : `⚠️ ${duplicates.length} duplicidades`}
                </button>
              )}
            </div>
          </div>
          <button onClick={onClose}
            className="text-gray-400 hover:text-gray-600 hover:bg-gray-100 rounded-lg p-1.5 transition-colors flex-shrink-0">
            {Ico.close}
          </button>
        </div>

        {/* ── Section-visibility checkboxes ── */}
        {(sdsLoaded || nddLoaded) && (
          <div className="px-6 py-2.5 border-b border-gray-100 bg-gray-50/60 flex items-center gap-5 flex-shrink-0 flex-wrap">
            <span className="text-[10px] font-bold text-gray-400 uppercase tracking-wider">Exibir:</span>
            {sdsLoaded && (
              <label className="flex items-center gap-1.5 cursor-pointer select-none">
                <input type="checkbox" checked={showSds}
                  onChange={e => { setShowSds(e.target.checked); if (!e.target.checked) setSdsFilter('all'); }}
                  className="w-3.5 h-3.5 accent-blue-600" />
                <span className="text-xs font-semibold text-blue-700">Monitoramento SDS</span>
              </label>
            )}
            {nddLoaded && (
              <label className="flex items-center gap-1.5 cursor-pointer select-none">
                <input type="checkbox" checked={showNdd}
                  onChange={e => { setShowNdd(e.target.checked); if (!e.target.checked) setNddFilter('all'); }}
                  className="w-3.5 h-3.5 accent-teal-600" />
                <span className="text-xs font-semibold text-teal-700">Monitoramento MPS</span>
              </label>
            )}
            {nddLoaded && (
              <label className="flex items-center gap-1.5 cursor-pointer select-none">
                <input type="checkbox" checked={showBilling}
                  onChange={e => { setShowBilling(e.target.checked); if (!e.target.checked) setBillingFilter('all'); }}
                  className="w-3.5 h-3.5 accent-purple-600" />
                <span className="text-xs font-semibold text-purple-700">Bilhetagem MPS</span>
              </label>
            )}
            {hasFilters && (
              <button onClick={clearFilters}
                className="ml-auto flex items-center gap-1 text-[10px] font-bold text-red-500 hover:text-red-700 transition-colors">
                <svg className="w-3 h-3" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                  <path strokeLinecap="round" strokeLinejoin="round" strokeWidth="2" d="M6 18L18 6M6 6l12 12"/>
                </svg>
                Limpar filtros
              </button>
            )}
          </div>
        )}

        {/* ── Scrollable body ── */}
        <div className="overflow-y-auto flex-grow px-6 py-4 space-y-5 custom-scrollbar">

          <div>

          {/* Monitoring charts */}
          {(sdsLoaded || nddLoaded) && (showSds || showNdd) && (
            <div className={`grid gap-4 ${(sdsLoaded && showSds) && (nddLoaded && showNdd) ? 'grid-cols-2' : 'grid-cols-1'}`}>

              {sdsLoaded && showSds && sdsData.length > 0 && (
                <div className="bg-blue-50/50 rounded-xl p-4 border border-blue-100">
                  <div className="text-xs font-bold text-blue-800 uppercase tracking-wide mb-1">Monitoramento SDS</div>
                  <div className="text-[9px] text-blue-400 mb-3">Clique em uma linha para filtrar os seriais</div>
                  <DonutWithCenter data={sdsData} centerValue={loc.total} centerLabel="equip." height={160} />
                  <div className="space-y-0.5 mt-3">
                    <FilterRow label="Monitorados"     value={loc.sds.monitored}    color={C.green}  total={sdsKnown} isActive={sdsFilter === 'monitored'}    onClick={() => toggleSds('monitored')} />
                    <FilterRow label="Alerta"          value={loc.sds.alert}        color={C.yellow} total={sdsKnown} isActive={sdsFilter === 'alert'}        onClick={() => toggleSds('alert')} />
                    <FilterRow label="Não Monitorados" value={loc.sds.notMonitored} color={C.red}    total={sdsKnown} isActive={sdsFilter === 'notMonitored'} onClick={() => toggleSds('notMonitored')} />
                  </div>
                </div>
              )}

              {nddLoaded && showNdd && nddChartData.length > 0 && (
                <div className="bg-teal-50/50 rounded-xl p-4 border border-teal-100">
                  <div className="text-xs font-bold text-teal-800 uppercase tracking-wide mb-1">Monitoramento MPS</div>
                  <div className="text-[9px] text-teal-400 mb-3">Clique em uma linha para filtrar os seriais</div>
                  <DonutWithCenter data={nddChartData} centerValue={loc.total} centerLabel="equip." height={160} />
                  <div className="space-y-0.5 mt-3">
                    <FilterRow label="Monitorados"     value={loc.ndd.monitored}    color={C.green}  total={nddKnown} isActive={nddFilter === 'monitored'}    onClick={() => toggleNdd('monitored')} />
                    <FilterRow label="Alerta"          value={loc.ndd.alert}        color={C.yellow} total={nddKnown} isActive={nddFilter === 'alert'}        onClick={() => toggleNdd('alert')} />
                    <FilterRow label="Não Monitorados" value={loc.ndd.notMonitored} color={C.red}    total={nddKnown} isActive={nddFilter === 'notMonitored'} onClick={() => toggleNdd('notMonitored')} />
                  </div>
                </div>
              )}
            </div>
          )}

          {/* Billing */}
          {nddLoaded && showBilling && billingTotal > 0 && (
            <div className="bg-gray-50 rounded-xl p-4 border border-gray-200">
              <div className="text-xs font-bold text-gray-600 uppercase tracking-wide mb-1">Bilhetagem MPS</div>
              <div className="text-[9px] text-gray-400 mb-3">Clique em um card para filtrar os seriais</div>
              <div className="flex gap-3 mb-3">
                {billingItems.map(d => (
                  <button key={d.key}
                    onClick={() => toggleBilling(d.key as 'active' | 'noRecent' | 'never')}
                    className="flex-1 rounded-lg py-2 px-3 text-center border-2 transition-all hover:opacity-90"
                    style={{
                      backgroundColor: d.color + '18',
                      borderColor: billingFilter === d.key ? d.color : 'transparent',
                      boxShadow: billingFilter === d.key ? `0 0 0 2px ${d.color}40` : 'none',
                    }}>
                    <div className="text-lg font-extrabold" style={{ color: d.color }}>
                      {d.value.toLocaleString('pt-BR')}
                    </div>
                    <div className="text-[9px] font-semibold text-gray-600">{d.name}</div>
                    <div className="text-[9px] text-gray-400">{Math.round(d.value / billingTotal * 100)}%</div>
                    {billingFilter === d.key && (
                      <div className="mt-1 text-[8px] font-bold" style={{ color: d.color }}>✓ Filtro ativo</div>
                    )}
                  </button>
                ))}
              </div>
              <HealthBar
                monitored={loc.billing.active}
                alert={loc.billing.noRecent}
                notMonitored={loc.billing.never}
                noData={loc.total - billingTotal}
                total={loc.total}
                showPct={false}
              />
            </div>
          )}

          {/* Situação */}
          {loc.situacao.length > 0 && (
            <div>
              <div className="text-xs font-bold text-gray-500 uppercase tracking-wide mb-3">Situação dos Equipamentos</div>
              <div className="space-y-2">
                {loc.situacao.map((s, i) => (
                  <div key={s.name} className="flex items-center gap-2">
                    <div className="text-xs text-gray-700 w-36 truncate flex-shrink-0" title={s.name}>{s.name}</div>
                    <div className="flex-grow h-3 rounded-full bg-gray-100 overflow-hidden">
                      <div
                        className="h-full rounded-full transition-all duration-500"
                        style={{ width: `${s.count / loc.total * 100}%`, backgroundColor: BAR_PALETTE[i % BAR_PALETTE.length] }}
                      />
                    </div>
                    <div className="text-xs font-bold text-gray-700 w-6 text-right flex-shrink-0">{s.count}</div>
                    <div className="text-[10px] text-gray-400 w-8 text-right flex-shrink-0">
                      {Math.round(s.count / loc.total * 100)}%
                    </div>
                  </div>
                ))}
              </div>
            </div>
          )}

          </div>

          {/* ── Serial list / detail table ── */}
          {uniqueSerials.length > 0 && (
            <div>
              {/* Header row: title + search + view toggle + export */}
              <div className="flex items-center justify-between gap-2 mb-3 flex-wrap gap-y-2">
                <div className="flex items-center gap-2 flex-wrap">
                  <div className="text-xs font-bold text-gray-500 uppercase tracking-wide">Dispositivos</div>
                  <span className="text-[10px] text-gray-400">
                    {hasFilters
                      ? `${filteredSerials.length} de ${uniqueSerials.length}`
                      : uniqueSerials.length.toLocaleString('pt-BR')}
                  </span>
                  {hasFilters && (
                    <span className="text-[9px] bg-blue-100 text-blue-700 px-1.5 py-0.5 rounded-full font-bold">
                      Filtros ativos
                    </span>
                  )}
                </div>
                <div className="flex items-center gap-2 flex-wrap">
                  {/* Search */}
                  <div className="relative">
                    <svg className="absolute left-2 top-1/2 -translate-y-1/2 w-3 h-3 text-gray-400" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                      <path strokeLinecap="round" strokeLinejoin="round" strokeWidth="2" d="M21 21l-6-6m2-5a7 7 0 11-14 0 7 7 0 0114 0z"/>
                    </svg>
                    <input
                      type="text"
                      value={search}
                      onChange={e => { setSearch(e.target.value); setPage(1); }}
                      placeholder="Buscar serial..."
                      className="pl-6 pr-6 py-1.5 text-xs border border-gray-200 rounded-lg focus:outline-none focus:ring-2 focus:ring-blue-200 w-36"
                    />
                    {search && (
                      <button onClick={() => { setSearch(''); setPage(1); }}
                        className="absolute right-2 top-1/2 -translate-y-1/2 text-gray-400 hover:text-gray-600">
                        <svg className="w-3 h-3" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                          <path strokeLinecap="round" strokeLinejoin="round" strokeWidth="2" d="M6 18L18 6M6 6l12 12"/>
                        </svg>
                      </button>
                    )}
                  </div>
                  {/* View toggle */}
                  <div className="flex rounded-lg overflow-hidden border border-gray-200 text-[10px] font-bold">
                    <button
                      onClick={() => setViewMode('chips')}
                      className={`px-2.5 py-1.5 transition-colors ${viewMode === 'chips' ? 'bg-blue-600 text-white' : 'bg-white text-gray-500 hover:bg-gray-50'}`}>
                      Seriais
                    </button>
                    <button
                      onClick={() => setViewMode('table')}
                      className={`px-2.5 py-1.5 transition-colors border-l border-gray-200 ${viewMode === 'table' ? 'bg-blue-600 text-white' : 'bg-white text-gray-500 hover:bg-gray-50'}`}>
                      Tabela
                    </button>
                  </div>
                  {/* Export buttons — only in table mode */}
                  {viewMode === 'table' && filteredSerials.length > 0 && (
                    <div className="flex gap-1">
                      <button
                        onClick={() => exportData('xlsx')}
                        className="flex items-center gap-1 px-2.5 py-1.5 text-[10px] font-bold bg-emerald-600 text-white rounded-lg hover:bg-emerald-700 transition-colors">
                        <svg className="w-3 h-3" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                          <path strokeLinecap="round" strokeLinejoin="round" strokeWidth="2" d="M4 16v1a3 3 0 003 3h10a3 3 0 003-3v-1m-4-4l-4 4m0 0l-4-4m4 4V4"/>
                        </svg>
                        XLSX
                      </button>
                      <button
                        onClick={() => exportData('csv')}
                        className="flex items-center gap-1 px-2.5 py-1.5 text-[10px] font-bold bg-gray-600 text-white rounded-lg hover:bg-gray-700 transition-colors">
                        <svg className="w-3 h-3" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                          <path strokeLinecap="round" strokeLinejoin="round" strokeWidth="2" d="M4 16v1a3 3 0 003 3h10a3 3 0 003-3v-1m-4-4l-4 4m0 0l-4-4m4 4V4"/>
                        </svg>
                        CSV
                      </button>
                    </div>
                  )}
                </div>
              </div>

              {filteredSerials.length === 0 ? (
                <div className="text-center py-8 text-xs text-gray-400 italic bg-gray-50 rounded-xl">
                  Nenhum serial corresponde aos filtros ativos.
                </div>
              ) : viewMode === 'chips' ? (
                <>
                  <div className="flex flex-wrap gap-1.5">
                    {pageSerials.map((s, i) => {
                      const isDup = duplicates.includes(s);
                      return (
                        <span key={i}
                          className={`px-2 py-0.5 rounded text-[10px] font-mono cursor-default relative transition-colors ${
                            isDup
                              ? 'bg-amber-100/70 text-amber-800 border border-amber-300 shadow-sm pr-6'
                              : 'bg-gray-100 hover:bg-gray-200 text-gray-700'
                          }`}>
                          {s}
                          {isDup && <span className="absolute -top-1.5 -right-1.5 bg-red-500 text-white text-[8px] font-bold px-1 rounded-full shadow-sm">x2+</span>}
                        </span>
                      );
                    })}
                  </div>
                  {totalPages > 1 && (
                    <div className="flex items-center justify-between mt-4 pt-3 border-t border-gray-100">
                      <button onClick={() => setPage(p => Math.max(1, p - 1))} disabled={safePage === 1}
                        className="px-3 py-1.5 text-xs font-bold text-gray-600 bg-gray-100 hover:bg-gray-200 rounded-lg disabled:opacity-40 disabled:cursor-not-allowed transition-colors">
                        ← Anterior
                      </button>
                      <div className="flex items-center gap-1">
                        {getPageNumbers(safePage, totalPages).map((p, i) =>
                          p === null ? (
                            <span key={`sep-${i}`} className="w-7 text-center text-xs text-gray-400">…</span>
                          ) : (
                            <button key={p} onClick={() => setPage(p)}
                              className={`w-7 h-7 text-xs font-bold rounded-lg transition-colors ${safePage === p ? 'bg-blue-600 text-white' : 'text-gray-500 hover:bg-gray-100'}`}>
                              {p}
                            </button>
                          )
                        )}
                      </div>
                      <button onClick={() => setPage(p => Math.min(totalPages, p + 1))} disabled={safePage === totalPages}
                        className="px-3 py-1.5 text-xs font-bold text-gray-600 bg-gray-100 hover:bg-gray-200 rounded-lg disabled:opacity-40 disabled:cursor-not-allowed transition-colors">
                        Próximo →
                      </button>
                    </div>
                  )}
                </>
              ) : (
                /* ── Detail table ── */
                <div className="rounded-xl border border-gray-200 shadow-sm"
                  style={{ overflowX: 'auto', overflowY: 'auto', maxHeight: '55vh' }}>
                  <table className="text-[11px] border-collapse" style={{ width: '100%', minWidth: 'max-content' }}>
                    <thead className="sticky top-0 z-20">
                      <tr className="bg-gray-50 border-b border-gray-200">
                        {[
                          { key: 'serie',      label: 'Série',       sticky: true  },
                          { key: 'modelo',     label: 'Modelo',      sticky: false },
                          { key: 'filial',     label: 'Filial',      sticky: false },
                          { key: 'logradouro', label: 'Logradouro',  sticky: false },
                          { key: 'bairro',     label: 'Bairro',      sticky: false },
                          { key: 'cidade',     label: 'Cidade',      sticky: false },
                          { key: 'uf',         label: 'UF',          sticky: false },
                          { key: 'cep',        label: 'CEP',         sticky: false },
                          { key: 'ip',         label: 'IP',          sticky: false },
                          { key: 'hostname',   label: 'Hostname',    sticky: false },
                          ...(hasSite       ? [{ key: 'site',       label: 'Site',        sticky: false }] : []),
                          ...(hasDepartment ? [{ key: 'department', label: 'Departamento', sticky: false }] : []),
                          { key: 'sds',        label: 'SDS',         sticky: false },
                          { key: 'lastSds',    label: 'Últ. SDS',    sticky: false },
                          { key: 'counter',    label: 'Contador',    sticky: false },
                          { key: 'ndd',        label: 'MPS',         sticky: false },
                          { key: 'lastNdd',    label: 'Últ. MPS',    sticky: false },
                          { key: 'billing',    label: 'Bilhetagem',  sticky: false },
                        ].map(col => (
                          <th
                            key={col.key}
                            onClick={() => handleTableSort(col.key)}
                            className={`px-3 py-2 text-left font-bold text-gray-600 uppercase tracking-wide cursor-pointer select-none whitespace-nowrap hover:bg-gray-100 transition-colors bg-gray-50 ${col.sticky ? 'sticky left-0 z-30 shadow-[2px_0_4px_-2px_rgba(0,0,0,0.1)]' : ''}`}>
                            <span className="flex items-center gap-1">
                              {col.label}
                              {tableSortKey === col.key
                                ? <span className="text-blue-500">{tableSortDir === 'asc' ? '↑' : '↓'}</span>
                                : <span className="text-gray-300">↕</span>
                              }
                            </span>
                          </th>
                        ))}
                      </tr>
                    </thead>
                    <tbody>
                      {sortedTableRows.map((r, i) => {
                        const sdsColor  = r.sdsStatus === 'monitored' ? 'bg-green-100 text-green-800' : r.sdsStatus === 'alert' ? 'bg-yellow-100 text-yellow-800' : r.sdsStatus === 'notMonitored' ? 'bg-red-100 text-red-800' : 'bg-gray-100 text-gray-500';
                        const nddColor  = r.nddStatus === 'monitored' ? 'bg-green-100 text-green-800' : r.nddStatus === 'alert' ? 'bg-yellow-100 text-yellow-800' : r.nddStatus === 'notMonitored' ? 'bg-red-100 text-red-800' : 'bg-gray-100 text-gray-500';
                        const billColor = r.billingStatus === 'active' ? 'bg-green-100 text-green-800' : r.billingStatus === 'noRecent' ? 'bg-yellow-100 text-yellow-800' : r.billingStatus === 'never' ? 'bg-red-100 text-red-800' : 'bg-gray-100 text-gray-400';
                        const isDup = duplicates.includes(r.serial);
                        return (
                          <tr key={r.serial} className={`border-b border-gray-100 hover:bg-blue-50/40 transition-colors ${i % 2 === 0 ? '' : 'bg-gray-50/30'}`}>
                            {/* Serial — sticky */}
                            <td className="px-3 py-2 sticky left-0 z-10 bg-white font-mono font-semibold text-gray-800 shadow-[2px_0_4px_-2px_rgba(0,0,0,0.08)] whitespace-nowrap">
                              <span className="flex items-center gap-1">
                                {r.serial}
                                {isDup && <span className="text-[8px] font-bold bg-amber-400 text-white px-1 rounded-full">x2+</span>}
                              </span>
                            </td>
                            <td className="px-3 py-2 whitespace-nowrap text-gray-700">{r.modelo || '-'}</td>
                            <td className="px-3 py-2 whitespace-nowrap text-gray-600">{r.filial || '-'}</td>
                            <td className="px-3 py-2" style={{ minWidth: 180 }}><div className="truncate text-gray-700" style={{ maxWidth: 220 }} title={r.logradouro}>{r.logradouro || '-'}</div></td>
                            <td className="px-3 py-2 whitespace-nowrap text-gray-600">{r.bairro || '-'}</td>
                            <td className="px-3 py-2 whitespace-nowrap text-gray-700 font-medium">{r.cidade || '-'}</td>
                            <td className="px-3 py-2 whitespace-nowrap text-gray-600">{r.uf || '-'}</td>
                            <td className="px-3 py-2 whitespace-nowrap font-mono text-gray-600">{r.cep || '-'}</td>
                            <td className="px-3 py-2 whitespace-nowrap font-mono text-gray-600">{r.ip || '-'}</td>
                            <td className="px-3 py-2 whitespace-nowrap text-gray-600" style={{ minWidth: 120 }}><div className="truncate" style={{ maxWidth: 160 }} title={r.hostname}>{r.hostname || '-'}</div></td>
                            {hasSite       && <td className="px-3 py-2 whitespace-nowrap text-gray-600">{r.site || '-'}</td>}
                            {hasDepartment && <td className="px-3 py-2 whitespace-nowrap text-gray-600">{r.department || '-'}</td>}
                            <td className="px-3 py-2 whitespace-nowrap"><span className={`text-[10px] font-bold px-1.5 py-0.5 rounded-full ${sdsColor}`}>{SDS_LABEL[r.sdsStatus] || '-'}</span></td>
                            <td className="px-3 py-2 whitespace-nowrap text-gray-600">{r.lastSdsUpdate || '-'}</td>
                            <td className="px-3 py-2 whitespace-nowrap text-right text-gray-700 font-mono">{r.counterValue != null ? r.counterValue.toLocaleString('pt-BR') : '-'}</td>
                            <td className="px-3 py-2 whitespace-nowrap"><span className={`text-[10px] font-bold px-1.5 py-0.5 rounded-full ${nddColor}`}>{NDD_LABEL[r.nddStatus] || '-'}</span></td>
                            <td className="px-3 py-2 whitespace-nowrap text-gray-600">{r.lastNddUpdate || '-'}</td>
                            <td className="px-3 py-2 whitespace-nowrap"><span className={`text-[10px] font-bold px-1.5 py-0.5 rounded-full ${billColor}`}>{r.billingStatusText || '-'}</span></td>
                          </tr>
                        );
                      })}
                    </tbody>
                  </table>
                </div>
              )}
            </div>
          )}
        </div>

        {/* ── Footer ── */}
        <div className="px-6 py-3 border-t border-gray-100 flex items-center justify-between flex-shrink-0">
          {hasFilters ? (
            <button onClick={clearFilters}
              className="text-xs font-bold text-red-500 hover:text-red-700 transition-colors">
              Limpar todos os filtros
            </button>
          ) : <div />}
          <button onClick={onClose}
            className="px-4 py-2 text-sm font-bold text-gray-600 bg-gray-100 hover:bg-gray-200 rounded-lg transition-colors">
            Fechar
          </button>
        </div>
      </div>
    </div>
  );
};

// ── CardFilterModal ───────────────────────────────────────────────────────────

type CardFilterKey =
  | 'total' | 'sdsMonitored' | 'sdsAlert' | 'sdsNotMonitored'
  | 'mpsMonitored' | 'mpsNotMonitored' | 'billingActive' | 'billingNoRecent' | 'billingNever'
  | 'inContract' | 'outOfContract';

const CARD_FILTER_LABEL: Record<CardFilterKey, string> = {
  total:           'Todos os Equipamentos',
  sdsMonitored:    'SDS Monitorados',
  sdsAlert:        'SDS em Alerta',
  sdsNotMonitored: 'SDS Não Monitorados',
  mpsMonitored:    'MPS Monitorados',
  mpsNotMonitored: 'MPS Não Monitorados',
  billingActive:   'Bilhetagem Ativa',
  billingNoRecent: 'Sem Bilhetagem Recente',
  billingNever:    'Nunca Bilhetado',
  inContract:      'Em Contrato',
  outOfContract:   'Fora do Contrato',
};

interface CardFilterModalProps {
  filterKey: CardFilterKey | null;
  allSerialDetails: Record<string, SerialDetail>;
  sdsLoaded: boolean;
  nddLoaded: boolean;
  onClose: () => void;
}

const CardFilterModal: React.FC<CardFilterModalProps> = ({ filterKey, allSerialDetails, sdsLoaded, nddLoaded, onClose }) => {
  const syntheticLoc = useMemo((): LocationBreakdown | null => {
    if (!filterKey) return null;

    const entries = Object.entries(allSerialDetails);
    const filtered = entries.filter(([, d]) => {
      switch (filterKey) {
        case 'total':           return true;
        case 'sdsMonitored':    return d.sdsStatus === 'monitored';
        case 'sdsAlert':        return d.sdsStatus === 'alert';
        case 'sdsNotMonitored': return d.sdsStatus === 'notMonitored';
        case 'mpsMonitored':    return d.nddStatus === 'monitored';
        case 'mpsNotMonitored': return d.nddStatus === 'notMonitored';
        case 'billingActive':   return d.billingStatus === 'active';
        case 'billingNoRecent': return d.billingStatus === 'noRecent';
        case 'billingNever':    return d.billingStatus === 'never';
        case 'inContract':      return d.inContract === true;
        case 'outOfContract':   return d.inContract === false;
        default: return false;
      }
    });

    const serials = filtered.map(([s]) => s);
    const serialDetails = Object.fromEntries(filtered);
    const sds  = { monitored: 0, alert: 0, notMonitored: 0, noData: 0 };
    const ndd  = { monitored: 0, alert: 0, notMonitored: 0, noData: 0 };
    const bill = { active: 0, noRecent: 0, never: 0 };

    filtered.forEach(([, d]) => {
      sds[d.sdsStatus]++;
      ndd[d.nddStatus]++;
      if (d.billingStatus) bill[d.billingStatus]++;
    });

    return {
      name: CARD_FILTER_LABEL[filterKey],
      total: serials.length,
      sds,
      ndd,
      billing: bill,
      situacao: [],
      serials,
      serialDetails,
    };
  }, [filterKey, allSerialDetails]);

  return (
    <LocationDetailModal
      loc={syntheticLoc}
      sdsLoaded={sdsLoaded}
      nddLoaded={nddLoaded}
      onClose={onClose}
    />
  );
};

// ── LocationTable ─────────────────────────────────────────────────────────────

interface LocationTableProps {
  locations: LocationBreakdown[];
  sdsLoaded: boolean;
  nddLoaded: boolean;
  onRowClick: (loc: LocationBreakdown) => void;
}

const LocationTable: React.FC<LocationTableProps> = ({ locations, sdsLoaded, nddLoaded, onRowClick }) => {
  const [search, setSearch]     = useState('');
  const [sortKey, setSortKey]   = useState('total');
  const [sortDir, setSortDir]   = useState<'asc' | 'desc'>('desc');

  const rows = useMemo(() => {
    const getVal = (loc: LocationBreakdown): number | string => {
      switch (sortKey) {
        case 'name':             return loc.name;
        case 'total':            return loc.total;
        case 'sds.monitored':    return loc.sds.monitored;
        case 'sds.alert':        return loc.sds.alert;
        case 'sds.notMon':       return loc.sds.notMonitored;
        case 'ndd.monitored':    return loc.ndd.monitored;
        case 'ndd.alert':        return loc.ndd.alert;
        case 'ndd.notMon':       return loc.ndd.notMonitored;
        case 'bill.active':      return loc.billing.active;
        case 'bill.noRecent':    return loc.billing.noRecent;
        case 'bill.never':       return loc.billing.never;
        default:                 return 0;
      }
    };
    const filtered = search.trim()
      ? locations.filter(l => l.name.toLowerCase().includes(search.toLowerCase()))
      : locations;
    return [...filtered].sort((a, b) => {
      const va = getVal(a), vb = getVal(b);
      if (typeof va === 'string') {
        const cmp = (va as string).localeCompare(vb as string, 'pt-BR');
        return sortDir === 'asc' ? cmp : -cmp;
      }
      return sortDir === 'asc'
        ? (va as number) - (vb as number)
        : (vb as number) - (va as number);
    });
  }, [locations, search, sortKey, sortDir]);

  const handleSort = (k: string) => {
    if (sortKey === k) setSortDir(d => d === 'asc' ? 'desc' : 'asc');
    else { setSortKey(k); setSortDir('desc'); }
  };

  const SortArrow = ({ k }: { k: string }) => (
    <span className="ml-0.5 opacity-50 text-[9px]">
      {sortKey === k ? (sortDir === 'asc' ? '▲' : '▼') : '⇅'}
    </span>
  );

  const Th: React.FC<{ label: string; k: string; cls?: string }> = ({ label, k, cls = '' }) => (
    <th onClick={() => handleSort(k)}
      className={`px-2 py-2 text-[10px] font-bold uppercase tracking-wide cursor-pointer select-none whitespace-nowrap transition-colors hover:opacity-80 ${cls}`}>
      {label}<SortArrow k={k} />
    </th>
  );

  // Colored count cell: shows value + percentage below, zeroes greyed out
  const NumCell = ({ val, color, total }: { val: number; color: string; total: number }) => {
    if (val === 0) return (
      <td className="px-2 py-2 text-center text-xs text-gray-300 font-semibold">0</td>
    );
    const pct = total > 0 ? Math.round(val / total * 100) : 0;
    return (
      <td className="px-2 py-2 text-center">
        <div className="text-xs font-bold leading-tight" style={{ color }}>
          {val.toLocaleString('pt-BR')}
        </div>
        <div className="text-[9px] text-gray-400 leading-tight">{pct}%</div>
      </td>
    );
  };

  // Small health indicator dot in the row
  const healthColor = (loc: LocationBreakdown) => {
    const sdsKnown = loc.total - loc.sds.noData;
    const nddKnown = loc.total - loc.ndd.noData;
    const rate = sdsKnown > 0 ? loc.sds.monitored / sdsKnown
               : nddKnown > 0 ? loc.ndd.monitored / nddKnown
               : null;
    if (rate === null) return C.slate;
    return rate >= 0.8 ? C.green : rate >= 0.5 ? C.yellow : C.red;
  };

  // Totals from visible rows
  const totals = useMemo(() => rows.reduce((acc, l) => ({
    total:           acc.total           + l.total,
    sdsMonitored:    acc.sdsMonitored    + l.sds.monitored,
    sdsAlert:        acc.sdsAlert        + l.sds.alert,
    sdsNotMon:       acc.sdsNotMon       + l.sds.notMonitored,
    nddMonitored:    acc.nddMonitored    + l.ndd.monitored,
    nddAlert:        acc.nddAlert        + l.ndd.alert,
    nddNotMon:       acc.nddNotMon       + l.ndd.notMonitored,
    billActive:      acc.billActive      + l.billing.active,
    billNoRecent:    acc.billNoRecent    + l.billing.noRecent,
    billNever:       acc.billNever       + l.billing.never,
  }), { total: 0, sdsMonitored: 0, sdsAlert: 0, sdsNotMon: 0,
        nddMonitored: 0, nddAlert: 0, nddNotMon: 0,
        billActive: 0, billNoRecent: 0, billNever: 0 }), [rows]);

  const colSpanTotal = 2
    + (sdsLoaded ? 3 : 0)
    + (nddLoaded ? 6 : 0);

  return (
    <div>
      {/* Search + count */}
      <div className="flex items-center gap-3 mb-3 flex-wrap">
        <div className="relative">
          <svg className="absolute left-2.5 top-1/2 -translate-y-1/2 w-3.5 h-3.5 text-gray-400 pointer-events-none"
            fill="none" stroke="currentColor" viewBox="0 0 24 24">
            <path strokeLinecap="round" strokeLinejoin="round" strokeWidth="2" d="M21 21l-6-6m2-5a7 7 0 11-14 0 7 7 0 0114 0z"/>
          </svg>
          <input type="text" value={search} onChange={e => setSearch(e.target.value)}
            placeholder="Filtrar localidade..."
            className="pl-8 pr-7 py-1.5 text-xs border border-gray-200 rounded-lg focus:outline-none focus:ring-2 focus:ring-blue-200 w-52" />
          {search && (
            <button onClick={() => setSearch('')}
              className="absolute right-2 top-1/2 -translate-y-1/2 text-gray-400 hover:text-gray-600">
              <svg className="w-3 h-3" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                <path strokeLinecap="round" strokeLinejoin="round" strokeWidth="2" d="M6 18L18 6M6 6l12 12"/>
              </svg>
            </button>
          )}
        </div>
        <span className="text-[10px] text-gray-400">
          {rows.length.toLocaleString('pt-BR')} {rows.length === 1 ? 'localidade' : 'localidades'}
          {search && ` de ${locations.length}`}
        </span>
      </div>

      {/* Table */}
      <div className="max-h-[500px] overflow-y-auto custom-scrollbar rounded-xl border border-gray-200 shadow-sm">
        <table className="w-full text-xs border-collapse">
          <thead className="sticky top-0 z-10">
            {/* Group header row */}
            <tr>
              <th colSpan={2}
                className="bg-gray-100 border-b border-r border-gray-200 px-3 py-1.5 text-left text-[9px] font-bold text-gray-400 uppercase tracking-wider" />
              {sdsLoaded && (
                <th colSpan={3}
                  className="bg-blue-100 border-b border-r border-blue-200 text-center text-[9px] font-bold text-blue-700 uppercase tracking-wider px-2 py-1.5">
                  Monitoramento SDS
                </th>
              )}
              {nddLoaded && (
                <th colSpan={3}
                  className="bg-teal-100 border-b border-r border-teal-200 text-center text-[9px] font-bold text-teal-700 uppercase tracking-wider px-2 py-1.5">
                  Monitoramento MPS
                </th>
              )}
              {nddLoaded && (
                <th colSpan={3}
                  className="bg-purple-100 border-b border-purple-200 text-center text-[9px] font-bold text-purple-700 uppercase tracking-wider px-2 py-1.5">
                  Bilhetagem MPS
                </th>
              )}
            </tr>
            {/* Column header row */}
            <tr className="bg-white shadow-sm">
              <Th label="Localidade" k="name"
                cls="text-left bg-gray-50 border-b border-r border-gray-200 pl-3 min-w-[180px] sticky left-0 z-20 shadow-[2px_0_4px_-2px_rgba(0,0,0,0.08)]" />
              <Th label="Total" k="total"
                cls="text-center bg-gray-50 border-b border-r border-gray-200 text-gray-600 min-w-[52px]" />
              {sdsLoaded && <>
                <Th label="Monit." k="sds.monitored"  cls="text-center bg-blue-50/70  border-b border-blue-100  text-emerald-700 min-w-[60px]" />
                <Th label="Alerta" k="sds.alert"      cls="text-center bg-blue-50/70  border-b border-blue-100  text-amber-700  min-w-[60px]" />
                <Th label="N.Mon." k="sds.notMon"     cls="text-center bg-blue-50/70  border-b border-r border-blue-200  text-red-700    min-w-[60px]" />
              </>}
              {nddLoaded && <>
                <Th label="Monit." k="ndd.monitored"  cls="text-center bg-teal-50/70  border-b border-teal-100  text-emerald-700 min-w-[60px]" />
                <Th label="Alerta" k="ndd.alert"      cls="text-center bg-teal-50/70  border-b border-teal-100  text-amber-700  min-w-[60px]" />
                <Th label="N.Mon." k="ndd.notMon"     cls="text-center bg-teal-50/70  border-b border-r border-teal-200  text-red-700    min-w-[60px]" />
              </>}
              {nddLoaded && <>
                <Th label="Ativa"    k="bill.active"   cls="text-center bg-purple-50/70 border-b border-purple-100 text-emerald-700 min-w-[60px]" />
                <Th label="Sem Rec." k="bill.noRecent" cls="text-center bg-purple-50/70 border-b border-purple-100 text-amber-700  min-w-[60px]" />
                <Th label="Nunca"    k="bill.never"    cls="text-center bg-purple-50/70 border-b border-purple-100 text-red-700    min-w-[60px]" />
              </>}
            </tr>
          </thead>

          <tbody>
            {rows.map((loc, i) => {
              const sdsKnown     = loc.total - loc.sds.noData;
              const nddKnown     = loc.total - loc.ndd.noData;
              const billingTotal = loc.billing.active + loc.billing.noRecent + loc.billing.never;
              const hColor       = healthColor(loc);
              return (
                <tr key={loc.name}
                  onClick={() => onRowClick(loc)}
                  className={`border-b border-gray-100 cursor-pointer transition-colors group hover:bg-blue-50/50 ${i % 2 === 0 ? 'bg-white' : 'bg-gray-50/40'}`}
                >
                  {/* Name cell — sticky */}
                  <td className={`px-3 py-2.5 border-r border-gray-100 sticky left-0 z-10 transition-colors ${i % 2 === 0 ? 'bg-white' : 'bg-gray-50/40'} group-hover:bg-blue-50/50 shadow-[2px_0_4px_-2px_rgba(0,0,0,0.05)]`}>
                    <div className="flex items-center gap-2">
                      <div className="w-2 h-2 rounded-full flex-shrink-0" style={{ backgroundColor: hColor }} />
                      <span className="font-semibold text-gray-800 truncate max-w-[200px]" title={loc.name}>
                        {loc.name}
                      </span>
                      <svg className="w-3 h-3 text-blue-400 ml-auto flex-shrink-0 opacity-0 group-hover:opacity-100 transition-opacity"
                        fill="none" stroke="currentColor" viewBox="0 0 24 24">
                        <path strokeLinecap="round" strokeLinejoin="round" strokeWidth="2" d="M10 6H6a2 2 0 00-2 2v10a2 2 0 002 2h10a2 2 0 002-2v-4M14 4h6m0 0v6m0-6L10 14"/>
                      </svg>
                    </div>
                  </td>
                  {/* Total */}
                  <td className="px-2 py-2.5 text-center border-r border-gray-100">
                    <span className="font-extrabold text-blue-600">{loc.total.toLocaleString('pt-BR')}</span>
                  </td>
                  {/* SDS */}
                  {sdsLoaded && <>
                    <NumCell val={loc.sds.monitored}    color={C.green}  total={sdsKnown} />
                    <NumCell val={loc.sds.alert}        color={C.yellow} total={sdsKnown} />
                    <NumCell val={loc.sds.notMonitored} color={C.red}    total={sdsKnown} />
                  </>}
                  {/* MPS */}
                  {nddLoaded && <>
                    <NumCell val={loc.ndd.monitored}    color={C.green}  total={nddKnown} />
                    <NumCell val={loc.ndd.alert}        color={C.yellow} total={nddKnown} />
                    <NumCell val={loc.ndd.notMonitored} color={C.red}    total={nddKnown} />
                  </>}
                  {/* Billing */}
                  {nddLoaded && <>
                    <NumCell val={loc.billing.active}   color={C.green}  total={billingTotal} />
                    <NumCell val={loc.billing.noRecent} color={C.yellow} total={billingTotal} />
                    <NumCell val={loc.billing.never}    color={C.red}    total={billingTotal} />
                  </>}
                </tr>
              );
            })}

            {rows.length === 0 && (
              <tr>
                <td colSpan={colSpanTotal}
                  className="text-center py-10 text-xs text-gray-400 italic">
                  Nenhuma localidade encontrada.
                </td>
              </tr>
            )}
          </tbody>

          {/* Totals footer */}
          {rows.length > 1 && (
            <tfoot className="sticky bottom-0 z-10">
              <tr className="bg-gray-100 border-t-2 border-gray-300 font-bold">
                <td className="px-3 py-2 text-xs font-bold text-gray-600 sticky left-0 bg-gray-100 border-r border-gray-200">
                  Total ({rows.length})
                </td>
                <td className="px-2 py-2 text-center text-xs font-extrabold text-blue-700 border-r border-gray-200">
                  {totals.total.toLocaleString('pt-BR')}
                </td>
                {sdsLoaded && <>
                  <td className="px-2 py-2 text-center text-xs font-bold" style={{ color: C.green  }}>{totals.sdsMonitored.toLocaleString('pt-BR')}</td>
                  <td className="px-2 py-2 text-center text-xs font-bold" style={{ color: C.yellow }}>{totals.sdsAlert.toLocaleString('pt-BR')}</td>
                  <td className="px-2 py-2 text-center text-xs font-bold border-r border-gray-200" style={{ color: C.red    }}>{totals.sdsNotMon.toLocaleString('pt-BR')}</td>
                </>}
                {nddLoaded && <>
                  <td className="px-2 py-2 text-center text-xs font-bold" style={{ color: C.green  }}>{totals.nddMonitored.toLocaleString('pt-BR')}</td>
                  <td className="px-2 py-2 text-center text-xs font-bold" style={{ color: C.yellow }}>{totals.nddAlert.toLocaleString('pt-BR')}</td>
                  <td className="px-2 py-2 text-center text-xs font-bold border-r border-gray-200" style={{ color: C.red    }}>{totals.nddNotMon.toLocaleString('pt-BR')}</td>
                  <td className="px-2 py-2 text-center text-xs font-bold" style={{ color: C.green  }}>{totals.billActive.toLocaleString('pt-BR')}</td>
                  <td className="px-2 py-2 text-center text-xs font-bold" style={{ color: C.yellow }}>{totals.billNoRecent.toLocaleString('pt-BR')}</td>
                  <td className="px-2 py-2 text-center text-xs font-bold" style={{ color: C.red    }}>{totals.billNever.toLocaleString('pt-BR')}</td>
                </>}
              </tr>
            </tfoot>
          )}
        </table>
      </div>
    </div>
  );
};

// ── Main Dashboard component ──────────────────────────────────────────────────

interface Props {
  stats: DashboardStats;
}

export const Dashboard: React.FC<Props> = ({ stats }) => {
  const [selectedLocation, setSelectedLocation] = useState<LocationBreakdown | null>(null);
  const [activeCardFilter, setActiveCardFilter] = useState<CardFilterKey | null>(null);
  const [locationGroupBy, setLocationGroupBy] = useState<'contrato' | 'cidade'>('contrato');
  const [locationView, setLocationView]         = useState<'cards' | 'table'>('cards');
  const [showAllLocations, setShowAllLocations] = useState(false);

  // ── Empty state ────────────────────────────────────────────────────────────
  if (stats.total === 0) {
    return (
      <div className="flex-grow flex items-center justify-center bg-gray-50">
        <div className="text-center">
          <div className="w-16 h-16 rounded-full bg-gray-100 flex items-center justify-center mx-auto mb-4">
            <svg className="w-8 h-8 text-gray-400" fill="none" stroke="currentColor" viewBox="0 0 24 24">
              <path strokeLinecap="round" strokeLinejoin="round" strokeWidth="1.5"
                d="M9 19v-6a2 2 0 00-2-2H5a2 2 0 00-2 2v6a2 2 0 002 2h2a2 2 0 002-2zm0 0V9a2 2 0 012-2h2a2 2 0 012 2v10m-6 0a2 2 0 002 2h2a2 2 0 002-2m0 0V5a2 2 0 012-2h2a2 2 0 012 2v14a2 2 0 01-2 2h-2a2 2 0 01-2-2z"/>
            </svg>
          </div>
          <h3 className="text-base font-bold text-gray-600 mb-1">Nenhum dado carregado</h3>
          <p className="text-sm text-gray-400">Carregue a pasta de relatórios para visualizar o dashboard.</p>
        </div>
      </div>
    );
  }

  // ── Computed values ────────────────────────────────────────────────────────

  const sdsChartData: DonutItem[] = [
    { name: 'Monitorado',        value: stats.sds.monitored,    color: C.green  },
    { name: 'Alerta',            value: stats.sds.alert,        color: C.yellow },
    { name: 'Não Monitorado',    value: stats.sds.notMonitored, color: C.red    },
    { name: 'Dados Incompletos', value: stats.sds.incomplete,   color: C.slate  },
  ].filter(d => d.value > 0);

  const nddChartData: DonutItem[] = [
    { name: 'Monitorado',     value: stats.ndd.monitored,    color: C.green  },
    { name: 'Alerta',         value: stats.ndd.alert,        color: C.yellow },
    { name: 'Não Monitorado', value: stats.ndd.notMonitored, color: C.red    },
  ].filter(d => d.value > 0);

  const billingChartData: DonutItem[] = [
    { name: 'Bilhetagem Ativa',       value: stats.billing.active,  color: C.green  },
    { name: 'Sem Bilhetagem Recente', value: stats.billing.noRecent, color: C.yellow },
    { name: 'Nunca Bilhetado',        value: stats.billing.never,   color: C.red    },
  ].filter(d => d.value > 0);

  const producingForPie: DonutItem[] = stats.producing.map((d, i) => ({
    name: d.name, value: d.count, color: BAR_PALETTE[i % BAR_PALETTE.length],
  }));

  const totalSds = stats.sds.monitored + stats.sds.alert + stats.sds.notMonitored + stats.sds.incomplete;
  const totalNdd = stats.ndd.monitored + stats.ndd.alert + stats.ndd.notMonitored;
  const billingTotal = stats.billing.active + stats.billing.noRecent + stats.billing.never;
  const pct = (n: number, t: number) => t > 0 ? `${Math.round(n / t * 100)}% de ${t} registros` : '';

  const locations = locationGroupBy === 'contrato' ? stats.locationsByContrato : stats.locationsByCity;
  const visibleLocations = showAllLocations ? locations : locations.slice(0, 18);

  // ── Render ─────────────────────────────────────────────────────────────────

  return (
    <>
      {/* Location detail modal */}
      <LocationDetailModal
        loc={selectedLocation}
        sdsLoaded={stats.sdsLoaded}
        nddLoaded={stats.nddLoaded}
        onClose={() => setSelectedLocation(null)}
      />

      {/* KPI card drill-down modal */}
      <CardFilterModal
        filterKey={activeCardFilter}
        allSerialDetails={stats.allSerialDetails}
        sdsLoaded={stats.sdsLoaded}
        nddLoaded={stats.nddLoaded}
        onClose={() => setActiveCardFilter(null)}
      />

      <div className="flex-grow overflow-auto custom-scrollbar bg-gray-50">
        <div className="p-4 max-w-screen-2xl mx-auto space-y-4">

          {/* ── Hint bar ── */}
          <div className="flex items-center gap-2 bg-blue-50 border border-blue-100 rounded-lg px-3 py-2 text-xs text-blue-700">
            <svg className="w-3.5 h-3.5 flex-shrink-0" fill="none" stroke="currentColor" viewBox="0 0 24 24">
              <path strokeLinecap="round" strokeLinejoin="round" strokeWidth="2" d="M13 16h-1v-4h-1m1-4h.01M21 12a9 9 0 11-18 0 9 9 0 0118 0z"/>
            </svg>
            Os indicadores refletem os&nbsp;<strong>{stats.total.toLocaleString('pt-BR')} registros</strong>&nbsp;do filtro atual — use a busca para detalhar por contrato, período ou qualquer campo.
          </div>

          {/* ── KPI cards ── */}
          <div className="grid grid-cols-2 sm:grid-cols-3 lg:grid-cols-4 xl:grid-cols-6 gap-3">
            <KpiCard
              label="Total Equipamentos"
              value={stats.total}
              sub="registros filtrados"
              bgClass="bg-gradient-to-br from-blue-600 to-indigo-700"
              icon={Ico.equip}
              onClick={() => setActiveCardFilter('total')}
            />

            {stats.sdsLoaded ? (
              <>
                <KpiCard
                  label="SDS Monitorados"
                  value={stats.sds.monitored}
                  sub={pct(stats.sds.monitored, totalSds)}
                  bgClass="bg-gradient-to-br from-emerald-500 to-green-600"
                  icon={Ico.check}
                  healthBar={{ monitored: stats.sds.monitored, alert: stats.sds.alert, notMonitored: stats.sds.notMonitored, noData: stats.sds.incomplete, total: totalSds }}
                  onClick={() => setActiveCardFilter('sdsMonitored')}
                />
                <KpiCard
                  label="SDS em Alerta"
                  value={stats.sds.alert}
                  sub="verificar urgência"
                  bgClass="bg-gradient-to-br from-amber-500 to-orange-600"
                  icon={Ico.warn}
                  healthBar={{ monitored: stats.sds.monitored, alert: stats.sds.alert, notMonitored: stats.sds.notMonitored, noData: stats.sds.incomplete, total: totalSds }}
                  onClick={() => setActiveCardFilter('sdsAlert')}
                />
                <KpiCard
                  label="SDS Não Monitorados"
                  value={stats.sds.notMonitored}
                  sub="requer atenção"
                  bgClass="bg-gradient-to-br from-red-500 to-rose-700"
                  icon={Ico.off}
                  healthBar={{ monitored: stats.sds.monitored, alert: stats.sds.alert, notMonitored: stats.sds.notMonitored, noData: stats.sds.incomplete, total: totalSds }}
                  onClick={() => setActiveCardFilter('sdsNotMonitored')}
                />
              </>
            ) : (
              <div className="col-span-2 sm:col-span-3 bg-white border border-dashed border-gray-200 rounded-xl p-4 flex items-center justify-center text-xs text-gray-400 italic">
                Carregue a Base SDS para ver o status de monitoramento
              </div>
            )}

            {stats.nddLoaded ? (
              <>
                <KpiCard
                  label="MPS Monitorados"
                  value={stats.ndd.monitored}
                  sub={pct(stats.ndd.monitored, totalNdd)}
                  bgClass="bg-gradient-to-br from-teal-500 to-cyan-600"
                  icon={Ico.star}
                  healthBar={{ monitored: stats.ndd.monitored, alert: stats.ndd.alert, notMonitored: stats.ndd.notMonitored, noData: 0, total: totalNdd }}
                  onClick={() => setActiveCardFilter('mpsMonitored')}
                />
                <KpiCard
                  label="MPS Não Monitorados"
                  value={stats.ndd.notMonitored}
                  sub="sem leitura recente"
                  bgClass="bg-gradient-to-br from-rose-500 to-pink-700"
                  icon={Ico.off}
                  healthBar={{ monitored: stats.ndd.monitored, alert: stats.ndd.alert, notMonitored: stats.ndd.notMonitored, noData: 0, total: totalNdd }}
                  onClick={() => setActiveCardFilter('mpsNotMonitored')}
                />
                <KpiCard
                  label="Bilhetagem Ativa"
                  value={stats.billing.active}
                  sub={pct(stats.billing.active, billingTotal)}
                  bgClass="bg-gradient-to-br from-violet-500 to-purple-700"
                  icon={Ico.bill}
                  healthBar={{ monitored: stats.billing.active, alert: stats.billing.noRecent, notMonitored: stats.billing.never, noData: stats.total - billingTotal, total: stats.total }}
                  onClick={() => setActiveCardFilter('billingActive')}
                />
              </>
            ) : (
              <div className="col-span-2 sm:col-span-3 bg-white border border-dashed border-gray-200 rounded-xl p-4 flex items-center justify-center text-xs text-gray-400 italic">
                Carregue a Base MPS para monitoramento NDD
              </div>
            )}

            {stats.corpLoaded ? (
              <>
                <KpiCard
                  label="Em Contrato"
                  value={stats.corp.inContract}
                  sub={pct(stats.corp.inContract, stats.total)}
                  bgClass="bg-gradient-to-br from-amber-500 to-yellow-600"
                  icon={Ico.contract}
                  onClick={() => setActiveCardFilter('inContract')}
                />
                <KpiCard
                  label="Fora do Contrato"
                  value={stats.corp.outOfContract}
                  sub="sem match no contrato"
                  bgClass="bg-gradient-to-br from-slate-500 to-gray-700"
                  icon={Ico.off}
                  onClick={() => setActiveCardFilter('outOfContract')}
                />
              </>
            ) : (
              <div className="col-span-2 bg-white border border-dashed border-gray-200 rounded-xl p-4 flex items-center justify-center text-xs text-gray-400 italic">
                Carregue o Contrato para ver equipamentos em contrato
              </div>
            )}
          </div>

          {/* ── Location / Contract Cards or Table ── */}
          {locations.length > 0 && (
            <ChartCard
              title={`Visão por ${locationGroupBy === 'contrato' ? 'Contrato' : 'Cidade'}`}
              action={
                <div className="flex items-center gap-2">
                  {/* Group-by toggle */}
                  <div className="flex items-center gap-0.5 bg-gray-100 rounded-lg p-0.5">
                    {(['contrato', 'cidade'] as const).map(g => (
                      <button key={g}
                        onClick={() => { setLocationGroupBy(g); setShowAllLocations(false); }}
                        className={`px-2 py-1 rounded-md text-[10px] font-bold transition-colors ${
                          locationGroupBy === g ? 'bg-white text-blue-700 shadow-sm' : 'text-gray-500 hover:text-gray-700'
                        }`}>
                        {g === 'contrato' ? 'Contrato' : 'Cidade'}
                      </button>
                    ))}
                  </div>
                  {/* View-mode toggle */}
                  <div className="flex items-center gap-0.5 bg-gray-100 rounded-lg p-0.5">
                    <button
                      onClick={() => setLocationView('cards')}
                      title="Visualização em cartões"
                      className={`px-2 py-1 rounded-md text-[10px] font-bold transition-colors flex items-center gap-1 ${
                        locationView === 'cards' ? 'bg-white text-blue-700 shadow-sm' : 'text-gray-500 hover:text-gray-700'
                      }`}>
                      <svg className="w-3 h-3" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                        <path strokeLinecap="round" strokeLinejoin="round" strokeWidth="2"
                          d="M4 6a2 2 0 012-2h2a2 2 0 012 2v2a2 2 0 01-2 2H6a2 2 0 01-2-2V6zm10 0a2 2 0 012-2h2a2 2 0 012 2v2a2 2 0 01-2 2h-2a2 2 0 01-2-2V6zM4 16a2 2 0 012-2h2a2 2 0 012 2v2a2 2 0 01-2 2H6a2 2 0 01-2-2v-2zm10 0a2 2 0 012-2h2a2 2 0 012 2v2a2 2 0 01-2 2h-2a2 2 0 01-2-2v-2z"/>
                      </svg>
                      Cartões
                    </button>
                    <button
                      onClick={() => setLocationView('table')}
                      title="Visualização em tabela"
                      className={`px-2 py-1 rounded-md text-[10px] font-bold transition-colors flex items-center gap-1 ${
                        locationView === 'table' ? 'bg-white text-blue-700 shadow-sm' : 'text-gray-500 hover:text-gray-700'
                      }`}>
                      <svg className="w-3 h-3" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                        <path strokeLinecap="round" strokeLinejoin="round" strokeWidth="2"
                          d="M3 10h18M3 14h18M10 4v16M3 6a2 2 0 012-2h14a2 2 0 012 2v12a2 2 0 01-2 2H5a2 2 0 01-2-2V6z"/>
                      </svg>
                      Tabela
                    </button>
                  </div>
                </div>
              }
            >
              {locationView === 'table' ? (
                <LocationTable
                  locations={locations}
                  sdsLoaded={stats.sdsLoaded}
                  nddLoaded={stats.nddLoaded}
                  onRowClick={loc => setSelectedLocation(loc)}
                />
              ) : (
                <>
                  <div className="grid grid-cols-2 sm:grid-cols-3 lg:grid-cols-4 xl:grid-cols-6 gap-3">
                    {visibleLocations.map(loc => (
                      <LocationCard
                        key={loc.name}
                        loc={loc}
                        sdsLoaded={stats.sdsLoaded}
                        nddLoaded={stats.nddLoaded}
                        onClick={() => setSelectedLocation(loc)}
                      />
                    ))}
                  </div>
                  {locations.length > 18 && (
                    <div className="flex items-center justify-center mt-4">
                      <button
                        onClick={() => setShowAllLocations(prev => !prev)}
                        className="text-xs font-bold text-blue-600 hover:text-blue-800 bg-blue-50 hover:bg-blue-100 px-4 py-1.5 rounded-lg transition-colors">
                        {showAllLocations
                          ? 'Ver menos'
                          : `Ver todos os ${locations.length} ${locationGroupBy === 'contrato' ? 'contratos' : 'cidades'}`}
                      </button>
                    </div>
                  )}
                </>
              )}
            </ChartCard>
          )}

          {/* ── Monitoring overview ── */}
          {(stats.sdsLoaded || stats.nddLoaded) && (
            <div className="grid grid-cols-1 md:grid-cols-2 xl:grid-cols-3 gap-4">

              {stats.sdsLoaded && sdsChartData.length > 0 && (
                <ChartCard title="Status de Monitoramento SDS">
                  <DonutWithCenter
                    data={sdsChartData}
                    centerValue={totalSds}
                    centerLabel="equipamentos"
                  />
                  <PieLegend items={sdsChartData} />
                </ChartCard>
              )}

              {stats.nddLoaded && nddChartData.length > 0 && (
                <ChartCard title="Status de Monitoramento MPS">
                  <DonutWithCenter
                    data={nddChartData}
                    centerValue={totalNdd}
                    centerLabel="equipamentos"
                  />
                  <PieLegend items={nddChartData} />
                </ChartCard>
              )}

              {stats.nddLoaded && billingChartData.length > 0 && (
                <ChartCard title="Status de Bilhetagem MPS">
                  <DonutWithCenter
                    data={billingChartData}
                    centerValue={billingTotal}
                    centerLabel="equipamentos"
                  />
                  <PieLegend items={billingChartData} />
                </ChartCard>
              )}
            </div>
          )}

          {/* ── Equipment analysis ── */}
          <div className="grid grid-cols-1 md:grid-cols-2 xl:grid-cols-3 gap-4">

            {producingForPie.length > 0 && (
              <ChartCard title="Equipamentos Produzindo">
                <DonutWithCenter
                  data={producingForPie}
                  centerValue={stats.total}
                  centerLabel="equipamentos"
                />
                <PieLegend items={producingForPie} />
              </ChartCard>
            )}

            {stats.situacao.length > 0 && (
              <ChartCard title="Situação do Equipamento">
                <ResponsiveContainer width="100%" height={220}>
                  <BarChart data={stats.situacao} margin={{ top: 4, right: 8, left: -24, bottom: 42 }}>
                    <CartesianGrid strokeDasharray="3 3" vertical={false} stroke="#f0f0f0" />
                    <XAxis dataKey="name" tick={{ fontSize: 9 }} angle={-35} textAnchor="end" interval={0} />
                    <YAxis tick={{ fontSize: 9 }} />
                    <Tooltip content={<BarTooltip />} />
                    <Bar dataKey="count" name="Equipamentos" radius={[4, 4, 0, 0]}>
                      {stats.situacao.map((_, i) => <Cell key={i} fill={BAR_PALETTE[i % BAR_PALETTE.length]} />)}
                    </Bar>
                  </BarChart>
                </ResponsiveContainer>
              </ChartCard>
            )}

            {stats.tipo.length > 0 && (
              <ChartCard title="Tipo de OS">
                <ResponsiveContainer width="100%" height={220}>
                  <BarChart data={stats.tipo} margin={{ top: 4, right: 8, left: -24, bottom: 42 }}>
                    <CartesianGrid strokeDasharray="3 3" vertical={false} stroke="#f0f0f0" />
                    <XAxis dataKey="name" tick={{ fontSize: 9 }} angle={-35} textAnchor="end" interval={0} />
                    <YAxis tick={{ fontSize: 9 }} />
                    <Tooltip content={<BarTooltip />} />
                    <Bar dataKey="count" name="Registros" radius={[4, 4, 0, 0]}>
                      {stats.tipo.map((_, i) => <Cell key={i} fill={BAR_PALETTE[(i + 3) % BAR_PALETTE.length]} />)}
                    </Bar>
                  </BarChart>
                </ResponsiveContainer>
              </ChartCard>
            )}
          </div>

          {/* ── Por Modelo e UF (Contrato) ── */}
          {stats.corpLoaded && (stats.byModelo.length > 0 || stats.byUf.length > 0) && (
            <div className="grid grid-cols-1 xl:grid-cols-2 gap-4">
              {stats.byModelo.length > 0 && (
                <ChartCard title="Top Modelos (Contrato)">
                  <ResponsiveContainer width="100%" height={Math.max(200, stats.byModelo.length * 28)}>
                    <BarChart data={stats.byModelo} layout="vertical" margin={{ top: 4, right: 30, left: 8, bottom: 4 }}>
                      <CartesianGrid strokeDasharray="3 3" horizontal={false} stroke="#f0f0f0" />
                      <XAxis type="number" tick={{ fontSize: 10 }} />
                      <YAxis type="category" dataKey="name" width={145} tick={{ fontSize: 9 }} />
                      <Tooltip content={<BarTooltip />} />
                      <Bar dataKey="count" name="Equipamentos" radius={[0, 4, 4, 0]}>
                        {stats.byModelo.map((_, i) => <Cell key={i} fill={BAR_PALETTE[i % BAR_PALETTE.length]} />)}
                      </Bar>
                    </BarChart>
                  </ResponsiveContainer>
                </ChartCard>
              )}
              {stats.byUf.length > 0 && (
                <ChartCard title="Por Estado / UF (Contrato)">
                  <ResponsiveContainer width="100%" height={Math.max(200, stats.byUf.length * 28)}>
                    <BarChart data={stats.byUf} layout="vertical" margin={{ top: 4, right: 30, left: 8, bottom: 4 }}>
                      <CartesianGrid strokeDasharray="3 3" horizontal={false} stroke="#f0f0f0" />
                      <XAxis type="number" tick={{ fontSize: 10 }} />
                      <YAxis type="category" dataKey="name" width={36} tick={{ fontSize: 10, fontWeight: 700 }} />
                      <Tooltip content={<BarTooltip />} />
                      <Bar dataKey="count" name="Equipamentos" fill={C.orange} radius={[0, 4, 4, 0]} />
                    </BarChart>
                  </ResponsiveContainer>
                </ChartCard>
              )}
            </div>
          )}

          {/* ── Top Contratos & Cidades ── */}
          <div className="grid grid-cols-1 xl:grid-cols-2 gap-4">
            {stats.byContrato.length > 0 && (
              <ChartCard title="Top 10 Contratos">
                <ResponsiveContainer width="100%" height={Math.max(200, stats.byContrato.length * 34)}>
                  <BarChart data={stats.byContrato} layout="vertical" margin={{ top: 4, right: 30, left: 8, bottom: 4 }}>
                    <CartesianGrid strokeDasharray="3 3" horizontal={false} stroke="#f0f0f0" />
                    <XAxis type="number" tick={{ fontSize: 10 }} />
                    <YAxis type="category" dataKey="name" width={95} tick={{ fontSize: 10 }} />
                    <Tooltip content={<BarTooltip />} />
                    <Bar dataKey="count" name="Registros" fill={C.blue} radius={[0, 4, 4, 0]} />
                  </BarChart>
                </ResponsiveContainer>
              </ChartCard>
            )}

            {stats.byCidade.length > 0 && (
              <ChartCard title="Top 10 Cidades">
                <ResponsiveContainer width="100%" height={Math.max(200, stats.byCidade.length * 34)}>
                  <BarChart data={stats.byCidade} layout="vertical" margin={{ top: 4, right: 30, left: 8, bottom: 4 }}>
                    <CartesianGrid strokeDasharray="3 3" horizontal={false} stroke="#f0f0f0" />
                    <XAxis type="number" tick={{ fontSize: 10 }} />
                    <YAxis type="category" dataKey="name" width={95} tick={{ fontSize: 10 }} />
                    <Tooltip content={<BarTooltip />} />
                    <Bar dataKey="count" name="Registros" fill={C.purple} radius={[0, 4, 4, 0]} />
                  </BarChart>
                </ResponsiveContainer>
              </ChartCard>
            )}
          </div>

          {/* ── MPS connection ── */}
          {stats.nddLoaded && stats.connectionType.length > 0 && (
            <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
              <ChartCard title="Tipo de Conexão MPS">
                <ResponsiveContainer width="100%" height={160}>
                  <BarChart data={stats.connectionType} margin={{ top: 4, right: 8, left: -24, bottom: 4 }}>
                    <CartesianGrid strokeDasharray="3 3" vertical={false} stroke="#f0f0f0" />
                    <XAxis dataKey="name" tick={{ fontSize: 10 }} />
                    <YAxis tick={{ fontSize: 10 }} />
                    <Tooltip content={<BarTooltip />} />
                    <Bar dataKey="count" name="Equipamentos" fill={C.teal} radius={[4, 4, 0, 0]} />
                  </BarChart>
                </ResponsiveContainer>
              </ChartCard>
            </div>
          )}

        </div>
      </div>
    </>
  );
};
