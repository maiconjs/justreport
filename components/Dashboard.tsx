import React, { useState, useMemo } from 'react';
import {
  PieChart, Pie, Cell, Tooltip, ResponsiveContainer,
  BarChart, Bar, XAxis, YAxis, CartesianGrid,
} from 'recharts';
import { CepInvalidEntry } from '../types';

// ── Types ─────────────────────────────────────────────────────────────────────

export interface StatEntry { name: string; count: number; }

export interface LocationBreakdown {
  name: string;
  total: number;
  sds: { monitored: number; alert: number; notMonitored: number; noData: number };
  ndd: { monitored: number; alert: number; notMonitored: number; noData: number };
  billing: { active: number; noRecent: number; never: number };
  situacao: StatEntry[];
  serials: string[];
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
}

const KpiCard: React.FC<KpiCardProps> = ({ label, value, sub, bgClass, icon, healthBar }) => (
  <div className={`rounded-xl p-4 flex flex-col shadow-sm ${bgClass}`}>
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

// ── LocationDetailModal ───────────────────────────────────────────────────────

interface LocationDetailModalProps {
  loc: LocationBreakdown | null;
  sdsLoaded: boolean;
  nddLoaded: boolean;
  onClose: () => void;
}

const LocationDetailModal: React.FC<LocationDetailModalProps> = ({ loc, sdsLoaded, nddLoaded, onClose }) => {
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

  const billingData = [
    { name: 'Ativa',           value: loc.billing.active,  color: C.green  },
    { name: 'Sem Rec.',        value: loc.billing.noRecent, color: C.yellow },
    { name: 'Nunca Bilhetado', value: loc.billing.never,   color: C.red    },
  ].filter(d => d.value > 0);

  const billingTotal = loc.billing.active + loc.billing.noRecent + loc.billing.never;
  const sdsKnown = loc.total - loc.sds.noData;
  const nddKnown = loc.total - loc.ndd.noData;

  const StatRow: React.FC<{ label: string; value: number; color: string; total: number }> = ({ label, value, color, total }) => (
    <div className="flex items-center gap-2">
      <div className="w-2 h-2 rounded-full flex-shrink-0" style={{ backgroundColor: color }} />
      <div className="text-xs text-gray-700 flex-grow">{label}</div>
      <div className="text-xs font-bold text-gray-800">{value.toLocaleString('pt-BR')}</div>
      <div className="text-[10px] text-gray-400 w-8 text-right">{total > 0 ? `${Math.round(value / total * 100)}%` : '–'}</div>
    </div>
  );

  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center bg-black/40 backdrop-blur-sm p-4"
      onClick={e => { if (e.target === e.currentTarget) onClose(); }}>
      <div className="bg-white rounded-2xl shadow-2xl w-full max-w-2xl max-h-[90vh] flex flex-col">

        {/* Header */}
        <div className="px-6 py-4 border-b border-gray-100 flex items-start justify-between">
          <div>
            <div className="text-[10px] text-gray-400 font-bold uppercase tracking-wider">Detalhes do Contrato / Local</div>
            <h2 className="text-xl font-extrabold text-gray-800 mt-0.5">{loc.name}</h2>
            <div className="flex items-center gap-3 mt-1">
              <span className="text-sm text-gray-500">{loc.total.toLocaleString('pt-BR')} equipamentos</span>
              {loc.serials.length > 0 && (
                <span className="text-xs bg-blue-50 text-blue-700 px-2 py-0.5 rounded-full font-semibold">
                  {loc.serials.length} seriais únicos
                </span>
              )}
            </div>
          </div>
          <button onClick={onClose}
            className="text-gray-400 hover:text-gray-600 hover:bg-gray-100 rounded-lg p-1.5 transition-colors">
            {Ico.close}
          </button>
        </div>

        {/* Scrollable body */}
        <div className="overflow-y-auto flex-grow px-6 py-4 space-y-5 custom-scrollbar">

          {/* Monitoring charts */}
          {(sdsLoaded || nddLoaded) && (
            <div className={`grid gap-4 ${sdsLoaded && nddLoaded ? 'grid-cols-2' : 'grid-cols-1'}`}>

              {sdsLoaded && sdsData.length > 0 && (
                <div className="bg-blue-50/50 rounded-xl p-4 border border-blue-100">
                  <div className="text-xs font-bold text-blue-800 uppercase tracking-wide mb-3">Monitoramento SDS</div>
                  <DonutWithCenter
                    data={sdsData}
                    centerValue={loc.total}
                    centerLabel="equip."
                    height={170}
                  />
                  <div className="space-y-1.5 mt-3">
                    <StatRow label="Monitorados"     value={loc.sds.monitored}    color={C.green}  total={sdsKnown} />
                    <StatRow label="Alerta"          value={loc.sds.alert}        color={C.yellow} total={sdsKnown} />
                    <StatRow label="Não Monitorados" value={loc.sds.notMonitored} color={C.red}    total={sdsKnown} />
                  </div>
                </div>
              )}

              {nddLoaded && nddChartData.length > 0 && (
                <div className="bg-teal-50/50 rounded-xl p-4 border border-teal-100">
                  <div className="text-xs font-bold text-teal-800 uppercase tracking-wide mb-3">Monitoramento MPS</div>
                  <DonutWithCenter
                    data={nddChartData}
                    centerValue={loc.total}
                    centerLabel="equip."
                    height={170}
                  />
                  <div className="space-y-1.5 mt-3">
                    <StatRow label="Monitorados"     value={loc.ndd.monitored}    color={C.green}  total={nddKnown} />
                    <StatRow label="Alerta"          value={loc.ndd.alert}        color={C.yellow} total={nddKnown} />
                    <StatRow label="Não Monitorados" value={loc.ndd.notMonitored} color={C.red}    total={nddKnown} />
                  </div>
                </div>
              )}
            </div>
          )}

          {/* Billing */}
          {nddLoaded && billingTotal > 0 && (
            <div className="bg-gray-50 rounded-xl p-4 border border-gray-200">
              <div className="text-xs font-bold text-gray-600 uppercase tracking-wide mb-3">Bilhetagem MPS</div>
              <div className="flex gap-3 mb-3">
                {billingData.map(d => (
                  <div key={d.name}
                    className="flex-1 rounded-lg py-2 px-3 text-center"
                    style={{ backgroundColor: d.color + '18', border: `1px solid ${d.color}30` }}>
                    <div className="text-lg font-extrabold" style={{ color: d.color }}>
                      {d.value.toLocaleString('pt-BR')}
                    </div>
                    <div className="text-[9px] font-semibold text-gray-600">{d.name}</div>
                    <div className="text-[9px] text-gray-400">{Math.round(d.value / billingTotal * 100)}%</div>
                  </div>
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

          {/* Equipment list */}
          {loc.serials.length > 0 && (
            <div>
              <div className="text-xs font-bold text-gray-500 uppercase tracking-wide mb-2">
                Seriais ({loc.serials.length})
              </div>
              <div className="flex flex-wrap gap-1.5 max-h-28 overflow-y-auto custom-scrollbar">
                {loc.serials.slice(0, 60).map((s, i) => (
                  <span key={i} className="px-2 py-0.5 bg-gray-100 text-gray-700 rounded text-[10px] font-mono">
                    {s}
                  </span>
                ))}
                {loc.serials.length > 60 && (
                  <span className="px-2 py-0.5 bg-blue-50 text-blue-600 rounded text-[10px] font-semibold">
                    +{loc.serials.length - 60} mais
                  </span>
                )}
              </div>
            </div>
          )}
        </div>

        {/* Footer */}
        <div className="px-6 py-3 border-t border-gray-100 flex justify-end">
          <button onClick={onClose}
            className="px-4 py-2 text-sm font-bold text-gray-600 bg-gray-100 hover:bg-gray-200 rounded-lg transition-colors">
            Fechar
          </button>
        </div>
      </div>
    </div>
  );
};

// ── CepQualitySection ─────────────────────────────────────────────────────────

interface CepQualityProps {
  cepStats: {
    total: number; valid: number; invalid: number; unchecked: number;
    invalidList: CepInvalidEntry[];
  };
}

const CepQualitySection: React.FC<CepQualityProps> = ({ cepStats }) => {
  const [expanded, setExpanded] = useState(false);
  const [ufFilter, setUfFilter] = useState('');

  const validPct  = cepStats.total > 0 ? Math.round(cepStats.valid    / cepStats.total * 100) : 0;
  const invalidPct = cepStats.total > 0 ? Math.round(cepStats.invalid / cepStats.total * 100) : 0;

  const filtered = useMemo(() =>
    ufFilter
      ? cepStats.invalidList.filter(e => e.uf === ufFilter)
      : cepStats.invalidList,
    [cepStats.invalidList, ufFilter]
  );

  const ufs = useMemo(() =>
    [...new Set(cepStats.invalidList.map(e => e.uf).filter(Boolean))].sort(),
    [cepStats.invalidList]
  );

  if (cepStats.unchecked > 0 && cepStats.valid === 0 && cepStats.invalid === 0) {
    return (
      <ChartCard title="Qualidade de Endereços (CEP)">
        <div className="flex items-center gap-2 text-xs text-gray-400 italic py-4 justify-center">
          <svg className="w-4 h-4 animate-spin text-amber-400" fill="none" viewBox="0 0 24 24">
            <circle className="opacity-25" cx="12" cy="12" r="10" stroke="currentColor" strokeWidth="4"/>
            <path className="opacity-75" fill="currentColor" d="M4 12a8 8 0 018-8v8z"/>
          </svg>
          Aguardando validação ViaCEP ({cepStats.unchecked} CEPs pendentes)...
        </div>
      </ChartCard>
    );
  }

  return (
    <ChartCard
      title="Qualidade de Endereços (CEP)"
      action={
        cepStats.invalid > 0 ? (
          <button
            onClick={() => setExpanded(e => !e)}
            className="text-[10px] font-bold text-red-600 bg-red-50 hover:bg-red-100 px-3 py-1 rounded-lg transition-colors"
          >
            {expanded ? 'Ocultar' : `Ver ${cepStats.invalid} inválidos`}
          </button>
        ) : undefined
      }
    >
      {/* Summary bar */}
      <div className="grid grid-cols-3 gap-3 mb-4">
        <div className="rounded-xl p-3 bg-emerald-50 border border-emerald-100 text-center">
          <div className="text-2xl font-extrabold text-emerald-700">{cepStats.valid.toLocaleString('pt-BR')}</div>
          <div className="text-[10px] font-semibold text-emerald-600 mt-0.5">CEPs Válidos</div>
          <div className="text-[10px] text-emerald-500">{validPct}%</div>
        </div>
        <div className="rounded-xl p-3 bg-red-50 border border-red-100 text-center">
          <div className="text-2xl font-extrabold text-red-700">{cepStats.invalid.toLocaleString('pt-BR')}</div>
          <div className="text-[10px] font-semibold text-red-600 mt-0.5">CEPs Inválidos</div>
          <div className="text-[10px] text-red-500">{invalidPct}%</div>
        </div>
        <div className="rounded-xl p-3 bg-gray-50 border border-gray-200 text-center">
          <div className="text-2xl font-extrabold text-gray-600">{cepStats.unchecked.toLocaleString('pt-BR')}</div>
          <div className="text-[10px] font-semibold text-gray-500 mt-0.5">Sem CEP</div>
          <div className="text-[10px] text-gray-400">não validados</div>
        </div>
      </div>

      {/* Progress bar */}
      <div className="h-2 rounded-full bg-gray-100 flex overflow-hidden mb-2">
        <div className="bg-emerald-500 transition-all duration-700" style={{ width: `${validPct}%` }} />
        <div className="bg-red-400 transition-all duration-700"    style={{ width: `${invalidPct}%` }} />
      </div>
      <div className="text-[10px] text-gray-400 mb-4">
        {cepStats.total} CEPs únicos verificados via ViaCEP
      </div>

      {/* Invalid list */}
      {expanded && cepStats.invalid > 0 && (
        <div>
          <div className="flex items-center gap-2 mb-3">
            <span className="text-xs font-bold text-gray-600">Filtrar por UF:</span>
            <select
              value={ufFilter}
              onChange={e => setUfFilter(e.target.value)}
              className="border rounded text-xs px-2 py-0.5 bg-gray-50"
            >
              <option value="">Todos ({cepStats.invalid})</option>
              {ufs.map(uf => (
                <option key={uf} value={uf}>
                  {uf} ({cepStats.invalidList.filter(e => e.uf === uf).length})
                </option>
              ))}
            </select>
          </div>
          <div className="overflow-x-auto">
            <table className="w-full text-xs border-collapse">
              <thead>
                <tr className="bg-red-50 text-red-800">
                  <th className="px-3 py-2 text-left font-bold border-b border-red-100 whitespace-nowrap">Série</th>
                  <th className="px-3 py-2 text-left font-bold border-b border-red-100 whitespace-nowrap">CEP</th>
                  <th className="px-3 py-2 text-left font-bold border-b border-red-100 whitespace-nowrap">UF</th>
                  <th className="px-3 py-2 text-left font-bold border-b border-red-100 whitespace-nowrap">Cidade</th>
                  <th className="px-3 py-2 text-left font-bold border-b border-red-100 whitespace-nowrap">Modelo</th>
                  <th className="px-3 py-2 text-left font-bold border-b border-red-100">Endereço Original</th>
                </tr>
              </thead>
              <tbody className="divide-y divide-red-50">
                {filtered.slice(0, 100).map((e, i) => (
                  <tr key={i} className="hover:bg-red-50/50 transition-colors">
                    <td className="px-3 py-1.5 font-mono text-gray-700">{e.serial}</td>
                    <td className="px-3 py-1.5 font-bold text-red-600">{e.cep}</td>
                    <td className="px-3 py-1.5 font-bold text-gray-700">{e.uf || '-'}</td>
                    <td className="px-3 py-1.5 text-gray-700">{e.cidade || '-'}</td>
                    <td className="px-3 py-1.5 text-gray-600">{e.modelo || '-'}</td>
                    <td className="px-3 py-1.5 text-gray-500 max-w-xs truncate" title={e.enderecoRaw}>{e.enderecoRaw}</td>
                  </tr>
                ))}
              </tbody>
            </table>
            {filtered.length > 100 && (
              <div className="text-xs text-gray-400 text-center mt-2 italic">
                Mostrando 100 de {filtered.length} registros
              </div>
            )}
          </div>
        </div>
      )}
    </ChartCard>
  );
};

// ── Main Dashboard component ──────────────────────────────────────────────────

interface Props {
  stats: DashboardStats;
}

export const Dashboard: React.FC<Props> = ({ stats }) => {
  const [selectedLocation, setSelectedLocation] = useState<LocationBreakdown | null>(null);
  const [locationGroupBy, setLocationGroupBy] = useState<'contrato' | 'cidade'>('contrato');
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
                />
                <KpiCard
                  label="SDS em Alerta"
                  value={stats.sds.alert}
                  sub="verificar urgência"
                  bgClass="bg-gradient-to-br from-amber-500 to-orange-600"
                  icon={Ico.warn}
                  healthBar={{ monitored: stats.sds.monitored, alert: stats.sds.alert, notMonitored: stats.sds.notMonitored, noData: stats.sds.incomplete, total: totalSds }}
                />
                <KpiCard
                  label="SDS Não Monitorados"
                  value={stats.sds.notMonitored}
                  sub="requer atenção"
                  bgClass="bg-gradient-to-br from-red-500 to-rose-700"
                  icon={Ico.off}
                  healthBar={{ monitored: stats.sds.monitored, alert: stats.sds.alert, notMonitored: stats.sds.notMonitored, noData: stats.sds.incomplete, total: totalSds }}
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
                />
                <KpiCard
                  label="MPS Não Monitorados"
                  value={stats.ndd.notMonitored}
                  sub="sem leitura recente"
                  bgClass="bg-gradient-to-br from-rose-500 to-pink-700"
                  icon={Ico.off}
                  healthBar={{ monitored: stats.ndd.monitored, alert: stats.ndd.alert, notMonitored: stats.ndd.notMonitored, noData: 0, total: totalNdd }}
                />
                <KpiCard
                  label="Bilhetagem Ativa"
                  value={stats.billing.active}
                  sub={pct(stats.billing.active, billingTotal)}
                  bgClass="bg-gradient-to-br from-violet-500 to-purple-700"
                  icon={Ico.bill}
                  healthBar={{ monitored: stats.billing.active, alert: stats.billing.noRecent, notMonitored: stats.billing.never, noData: stats.total - billingTotal, total: stats.total }}
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
                />
                <KpiCard
                  label="Fora do Contrato"
                  value={stats.corp.outOfContract}
                  sub="sem match no contrato"
                  bgClass="bg-gradient-to-br from-slate-500 to-gray-700"
                  icon={Ico.off}
                />
              </>
            ) : (
              <div className="col-span-2 bg-white border border-dashed border-gray-200 rounded-xl p-4 flex items-center justify-center text-xs text-gray-400 italic">
                Carregue o Contrato para ver equipamentos em contrato
              </div>
            )}
          </div>

          {/* ── Location / Contract Cards ── */}
          {locations.length > 0 && (
            <ChartCard
              title={`Visão por ${locationGroupBy === 'contrato' ? 'Contrato' : 'Cidade'}`}
              action={
                <div className="flex items-center gap-1 bg-gray-100 rounded-lg p-0.5">
                  {(['contrato', 'cidade'] as const).map(g => (
                    <button
                      key={g}
                      onClick={() => { setLocationGroupBy(g); setShowAllLocations(false); }}
                      className={`px-2 py-1 rounded-md text-[10px] font-bold transition-colors ${
                        locationGroupBy === g
                          ? 'bg-white text-blue-700 shadow-sm'
                          : 'text-gray-500 hover:text-gray-700'
                      }`}
                    >
                      {g === 'contrato' ? 'Contrato' : 'Cidade'}
                    </button>
                  ))}
                </div>
              }
            >
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
                    className="text-xs font-bold text-blue-600 hover:text-blue-800 bg-blue-50 hover:bg-blue-100 px-4 py-1.5 rounded-lg transition-colors"
                  >
                    {showAllLocations
                      ? 'Ver menos'
                      : `Ver todos os ${locations.length} ${locationGroupBy === 'contrato' ? 'contratos' : 'cidades'}`}
                  </button>
                </div>
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

          {/* ── CEP Quality Report ── */}
          {stats.cepStats && stats.cepStats.total > 0 && (
            <CepQualitySection cepStats={stats.cepStats} />
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
