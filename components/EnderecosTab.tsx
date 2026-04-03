import React, { useState, useMemo } from 'react';
import { CepInvalidEntry, CepCorrectionEntry } from '../types';

export interface EnderecosTabProps {
  cepStats: {
    total: number;
    valid: number;
    invalid: number;
    unchecked: number;
    invalidList: CepInvalidEntry[];
    correctedList: CepCorrectionEntry[];
  } | null;
}

const ITEMS_PER_PAGE = 50;

export const EnderecosTab: React.FC<EnderecosTabProps> = ({ cepStats }) => {
  const [expandedSection, setExpandedSection] = useState<'none' | 'invalid' | 'corrected'>('none');
  const [ufFilter, setUfFilter] = useState('');
  const [page, setPage] = useState(1);

  if (!cepStats) {
    return (
      <div className="flex-grow flex items-center justify-center bg-gray-50">
        <div className="text-center">
          <div className="w-16 h-16 rounded-full bg-gray-100 flex items-center justify-center mx-auto mb-4">
            <svg className="w-8 h-8 text-gray-400" fill="none" stroke="currentColor" viewBox="0 0 24 24">
              <path strokeLinecap="round" strokeLinejoin="round" strokeWidth="1.5"
                d="M17.657 16.657L13.414 20.9a1.998 1.998 0 01-2.827 0l-4.244-4.243a8 8 0 1111.314 0z"/>
              <path strokeLinecap="round" strokeLinejoin="round" strokeWidth="1.5" d="M15 11a3 3 0 11-6 0 3 3 0 016 0z"/>
            </svg>
          </div>
          <h3 className="text-base font-bold text-gray-600 mb-1">Nenhum endereço</h3>
          <p className="text-sm text-gray-400">Carregue o arquivo de base para avaliar as ferramentas de CEP.</p>
        </div>
      </div>
    );
  }

  const validPct  = cepStats.total > 0 ? Math.round((cepStats.valid / cepStats.total) * 100) : 0;
  const invalidPct = cepStats.total > 0 ? Math.round((cepStats.invalid / cepStats.total) * 100) : 0;

  const filteredInvalid = useMemo(() =>
    ufFilter
      ? cepStats.invalidList.filter(e => e.uf === ufFilter)
      : cepStats.invalidList,
  [cepStats.invalidList, ufFilter]);

  const ufs = useMemo(() =>
    [...new Set(cepStats.invalidList.map(e => e.uf).filter(Boolean))].sort(),
  [cepStats.invalidList]);

  const invalidPageItems = filteredInvalid.slice((page - 1) * ITEMS_PER_PAGE, page * ITEMS_PER_PAGE);
  const correctedPageItems = cepStats.correctedList.slice((page - 1) * ITEMS_PER_PAGE, page * ITEMS_PER_PAGE);

  const totalInvalidPages = Math.ceil(filteredInvalid.length / ITEMS_PER_PAGE);
  const totalCorrectedPages = Math.ceil(cepStats.correctedList.length / ITEMS_PER_PAGE);

  const handleSectionChange = (section: 'none' | 'invalid' | 'corrected') => {
    setExpandedSection(s => s === section ? 'none' : section);
    setPage(1);
  };

  const handlePageChange = (newPage: number) => {
    setPage(newPage);
  };

  if (cepStats.unchecked > 0 && cepStats.valid === 0 && cepStats.invalid === 0) {
    return (
      <div className="flex-grow p-4 max-w-screen-2xl mx-auto w-full">
        <div className="bg-white rounded-xl border border-gray-200 shadow-sm p-4 text-center text-xs text-gray-400 italic">
          <svg className="w-4 h-4 animate-spin text-amber-400 mr-2 inline-block" fill="none" viewBox="0 0 24 24">
            <circle className="opacity-25" cx="12" cy="12" r="10" stroke="currentColor" strokeWidth="4"/>
            <path className="opacity-75" fill="currentColor" d="M4 12a8 8 0 018-8v8z"/>
          </svg>
          Aguardando validação ViaCEP ({cepStats.unchecked} CEPs pendentes)...
        </div>
      </div>
    );
  }

  return (
    <div className="flex-grow overflow-auto custom-scrollbar bg-gray-50 p-4">
      <div className="max-w-screen-2xl mx-auto w-full space-y-4">
        <div className="bg-white rounded-xl border border-gray-200 shadow-sm flex flex-col">
          <div className="px-4 py-3 border-b border-gray-100 flex items-center justify-between flex-shrink-0">
            <h3 className="text-xs font-bold text-gray-500 uppercase tracking-wider">Qualidade de Endereços (CEP)</h3>
            <div className="flex gap-2">
              {cepStats.correctedList.length > 0 && (
                <button
                  onClick={() => handleSectionChange('corrected')}
                  className={`text-[10px] font-bold px-3 py-1 rounded-lg transition-colors ${
                    expandedSection === 'corrected' 
                      ? 'bg-blue-600 text-white' 
                      : 'text-blue-600 bg-blue-50 hover:bg-blue-100'
                  }`}
                >
                  {expandedSection === 'corrected' ? 'Ocultar Correções' : `Ver ${cepStats.correctedList.length} correções`}
                </button>
              )}
              {cepStats.invalid > 0 && (
                <button
                  onClick={() => handleSectionChange('invalid')}
                  className={`text-[10px] font-bold px-3 py-1 rounded-lg transition-colors ${
                    expandedSection === 'invalid' 
                      ? 'bg-red-600 text-white' 
                      : 'text-red-600 bg-red-50 hover:bg-red-100'
                  }`}
                >
                  {expandedSection === 'invalid' ? 'Ocultar Inválidos' : `Ver ${cepStats.invalid} inválidos`}
                </button>
              )}
            </div>
          </div>
          
          <div className="p-4 flex-grow min-h-0">
            {/* Summary bar */}
            <div className="grid grid-cols-4 gap-3 mb-4">
              <div className="rounded-xl p-3 bg-emerald-50 border border-emerald-100 text-center">
                <div className="text-2xl font-extrabold text-emerald-700">{cepStats.valid.toLocaleString('pt-BR')}</div>
                <div className="text-[10px] font-semibold text-emerald-600 mt-0.5">CEPs Válidos</div>
                <div className="text-[10px] text-emerald-500">{validPct}%</div>
              </div>
              <div className="rounded-xl p-3 bg-blue-50 border border-blue-100 text-center">
                <div className="text-2xl font-extrabold text-blue-700">{cepStats.correctedList.length.toLocaleString('pt-BR')}</div>
                <div className="text-[10px] font-semibold text-blue-600 mt-0.5">Corrigidos</div>
                <div className="text-[10px] text-blue-500">
                  {cepStats.total > 0 ? Math.round((cepStats.correctedList.length / cepStats.total) * 100) : 0}% auto-corrigidos
                </div>
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
            {expandedSection === 'invalid' && cepStats.invalid > 0 && (
              <div className="mt-4 pt-4 border-t border-gray-100">
                <div className="flex items-center gap-2 mb-3">
                  <span className="text-xs font-bold text-gray-600">Filtrar por UF:</span>
                  <select
                    value={ufFilter}
                    onChange={e => { setUfFilter(e.target.value); setPage(1); }}
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
                      {invalidPageItems.map((e, i) => (
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
                  
                  {filteredInvalid.length > ITEMS_PER_PAGE && (
                    <div className="flex items-center justify-between mt-4 bg-gray-50 p-2 rounded-lg border">
                      <span className="text-xs text-gray-500">
                        Mostrando {((page - 1) * ITEMS_PER_PAGE) + 1} a {Math.min(page * ITEMS_PER_PAGE, filteredInvalid.length)} de {filteredInvalid.length}
                      </span>
                      <div className="flex gap-1">
                        <button 
                          disabled={page === 1}
                          onClick={() => handlePageChange(page - 1)}
                          className="px-3 py-1 text-xs border rounded bg-white hover:bg-gray-100 disabled:opacity-50"
                        > Anterior </button>
                        <button 
                          disabled={page === totalInvalidPages}
                          onClick={() => handlePageChange(page + 1)}
                          className="px-3 py-1 text-xs border rounded bg-white hover:bg-gray-100 disabled:opacity-50"
                        > Próxima </button>
                      </div>
                    </div>
                  )}
                </div>
              </div>
            )}

            {/* Corrected list */}
            {expandedSection === 'corrected' && cepStats.correctedList.length > 0 && (
              <div className="overflow-x-auto mt-4 pt-4 border-t border-gray-100">
                <table className="w-full text-xs border-collapse">
                  <thead>
                    <tr className="bg-blue-50 text-blue-800">
                      <th className="px-3 py-2 text-left font-bold border-b border-blue-100 w-32 whitespace-nowrap">Série</th>
                      <th className="px-3 py-2 text-left font-bold border-b border-blue-100 w-32 whitespace-nowrap">CEP</th>
                      <th className="px-3 py-2 text-left font-bold border-b border-blue-100 w-1/2">Endereço Original (XLS)</th>
                      <th className="px-3 py-2 text-left font-bold border-b border-blue-100 w-1/2">Novo Endereço (ViaCEP)</th>
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-blue-50">
                    {correctedPageItems.map((e, i) => (
                      <tr key={i} className="hover:bg-blue-50/50 transition-colors">
                        <td className="px-3 py-2 font-mono text-gray-700">{e.serial}</td>
                        <td className="px-3 py-2 font-bold text-blue-600">{e.cep}</td>
                        <td className="px-3 py-2 text-red-600 line-through opacity-70" title={e.original}>{e.original}</td>
                        <td className="px-3 py-2 font-semibold text-emerald-700" title={e.corrected}>{e.corrected}</td>
                      </tr>
                    ))}
                  </tbody>
                </table>
                {cepStats.correctedList.length > ITEMS_PER_PAGE && (
                  <div className="flex items-center justify-between mt-4 bg-gray-50 p-2 rounded-lg border">
                    <span className="text-xs text-gray-500">
                      Mostrando {((page - 1) * ITEMS_PER_PAGE) + 1} a {Math.min(page * ITEMS_PER_PAGE, cepStats.correctedList.length)} de {cepStats.correctedList.length}
                    </span>
                    <div className="flex gap-1">
                      <button 
                        disabled={page === 1}
                        onClick={() => handlePageChange(page - 1)}
                        className="px-3 py-1 text-xs border rounded bg-white hover:bg-gray-100 disabled:opacity-50"
                      > Anterior </button>
                      <button 
                        disabled={page === totalCorrectedPages}
                        onClick={() => handlePageChange(page + 1)}
                        className="px-3 py-1 text-xs border rounded bg-white hover:bg-gray-100 disabled:opacity-50"
                      > Próxima </button>
                    </div>
                  </div>
                )}
              </div>
            )}
          </div>
        </div>
      </div>
    </div>
  );
};
