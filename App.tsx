import React, { useState, useEffect, useMemo, useRef } from 'react';
import * as XLSX from 'xlsx';
import { jsPDF } from 'jspdf';
import autoTable from 'jspdf-autotable';
import { ReportItem, SdsInfo, MapInfo, FilterState, MapColumnConfig } from './types';
import { readExcelFile, processReportData, parseMapWorkbook, processMapSheet, mapColumnsConfig } from './services/excelService';
import { parseDateRobust } from './utils/dateUtils';
import { Modal } from './components/Modal';
import { ProgressBar } from './components/ProgressBar';

const App: React.FC = () => {
  // State
  const [allData, setAllData] = useState<ReportItem[]>([]);
  const [sdsData, setSdsData] = useState<Map<string, { rawLastUpdate: Date | null, rawDetection: Date | null }>>(new Map());
  const [mapData, setMapData] = useState<Map<string, MapInfo>>(new Map());
  
  const [isProcessing, setIsProcessing] = useState(false);
  const [progress, setProgress] = useState(0);
  const [progressText, setProgressText] = useState('');

  // UI State
  const [sheetModalOpen, setSheetModalOpen] = useState(false);
  const [exportModalOpen, setExportModalOpen] = useState(false);
  const [mapWorkbook, setMapWorkbook] = useState<XLSX.WorkBook | null>(null);
  const [mapSheetNames, setMapSheetNames] = useState<string[]>([]);
  const [selectedSheet, setSelectedSheet] = useState<string | null>(null);
  const [manualHeaderRow, setManualHeaderRow] = useState<string>('');
  const [exportType, setExportType] = useState<'csv' | 'xlsx' | 'pdf' | null>(null);

  // Filter State
  const [filters, setFilters] = useState<FilterState>({
    search: '',
    searchField: 'all',
    startCreation: '',
    endCreation: '',
    startConclusion: '',
    endConclusion: '',
    alertDays: 7,
    offlineDays: 30,
    selectedTypes: [],
    selectedProds: [],
    selectedStatus: [],
    selectedSituacao: [],
    selectedConexao: [],
    selectedMon: []
  });

  const [currentPage, setCurrentPage] = useState(1);
  const itemsPerPage = 50;

  // File Inputs Refs
  const folderInputRef = useRef<HTMLInputElement>(null);
  const sdsInputRef = useRef<HTMLInputElement>(null);
  const mapInputRef = useRef<HTMLInputElement>(null);

  // --- Helpers ---

  const getSdsInfo = (serial: string): SdsInfo => {
    const empty = { status: '-', colorClass: '', lastUpdate: '-', detection: '-', rawLastUpdate: null, rawDetection: null } as SdsInfo;
    if (sdsData.size === 0 || !serial || serial === '-') return empty;
    
    const key = String(serial).trim().toUpperCase();
    const record = sdsData.get(key);
    
    if (!record) {
      return { ...empty, status: 'Não Monitorado', colorClass: 'bg-red-100 text-red-800 font-semibold' };
    }

    const { rawLastUpdate, rawDetection } = record;
    const lastUpdate = rawLastUpdate ? rawLastUpdate.toLocaleDateString('pt-BR') : 'N/A';
    const detection = rawDetection ? rawDetection.toLocaleDateString('pt-BR') : 'N/A';

    if (!rawLastUpdate) return { status: 'Dados Incompletos', colorClass: 'bg-yellow-100 text-yellow-800', lastUpdate, detection, rawLastUpdate, rawDetection };

    const now = new Date();
    const diffDays = Math.ceil(Math.abs(now.getTime() - rawLastUpdate.getTime()) / (1000 * 60 * 60 * 24));

    if (diffDays > filters.offlineDays) return { status: 'Não Monitorado', colorClass: 'bg-red-100 text-red-800 font-semibold', lastUpdate, detection, rawLastUpdate, rawDetection };
    if (diffDays > filters.alertDays) return { status: 'Alerta', colorClass: 'bg-yellow-100 text-yellow-800 font-semibold', lastUpdate, detection, rawLastUpdate, rawDetection };
    
    return { status: 'Monitorado', colorClass: 'bg-green-100 text-green-800 font-semibold', lastUpdate, detection, rawLastUpdate, rawDetection };
  };

  const getMapInfo = (serial: string): MapInfo => {
    if (mapData.size === 0 || !serial || serial === '-') {
      const empty: any = {};
      mapColumnsConfig.forEach(c => empty[c.key] = '-');
      return empty;
    }
    const key = String(serial).trim().toUpperCase();
    return mapData.get(key) || mapColumnsConfig.reduce((acc, col) => ({...acc, [col.key]: 'N/A'}), {} as MapInfo);
  };

  // --- Handlers ---

  const handleFolderSelect = async (e: React.ChangeEvent<HTMLInputElement>) => {
    if (!e.target.files || e.target.files.length === 0) return;
    
    setIsProcessing(true);
    setProgress(0);
    setProgressText("Lendo arquivos...");
    setAllData([]);

    const files = Array.from(e.target.files).filter((f: any) => f.name.match(/\.(xlsx|xls|csv)$/i)) as File[];
    let processed = 0;
    const newData: ReportItem[] = [];

    // Process in chunks to not freeze UI
    const chunkSize = 5;
    for (let i = 0; i < files.length; i += chunkSize) {
        const chunk = files.slice(i, i + chunkSize);
        await Promise.all(chunk.map(async (file: File) => {
            try {
                const raw = await readExcelFile(file);
                const processedRows = processReportData(raw, file.name);
                newData.push(...processedRows);
            } catch (err) {
                console.warn(`Erro lendo ${file.name}`, err);
            }
        }));
        processed += chunk.length;
        setProgress(Math.round((processed / files.length) * 100));
        setProgressText(`Processados ${processed} de ${files.length} arquivos...`);
        // Small delay to allow UI update
        await new Promise(r => setTimeout(r, 10));
    }

    setAllData(newData);
    setIsProcessing(false);
    setProgressText('');
  };

  const handleSdsSelect = async (e: React.ChangeEvent<HTMLInputElement>) => {
    if (!e.target.files || !e.target.files[0]) return;
    setIsProcessing(true);
    setProgressText("Lendo base SDS...");
    
    try {
        const raw = await readExcelFile(e.target.files[0]);
        const newSdsData = new Map();
        
        raw.forEach((row: any) => {
             // Helper to find keys case-insensitive
             const findVal = (keys: string[]) => {
                for (const k of keys) {
                    if (row[k]) return row[k];
                    const upper = Object.keys(row).find(rk => rk.toUpperCase() === k.toUpperCase());
                    if (upper) return row[upper];
                }
                return null;
             };

             const serial = findVal(['Número de série', 'Numero de serie', 'Serial', 'Nº Série']);
             const lastUpdate = findVal(['Ultima atualização', 'Última atualização', 'Last Update']);
             const detection = findVal(['Data de detecção', 'Data detecção', 'Detection Date']);

             if (serial) {
                 const key = String(serial).trim().toUpperCase();
                 newSdsData.set(key, {
                     rawLastUpdate: parseDateRobust(lastUpdate),
                     rawDetection: parseDateRobust(detection)
                 });
             }
        });
        setSdsData(newSdsData);
        alert(`Base SDS carregada: ${newSdsData.size} registros.`);
    } catch (err) {
        alert("Erro ao ler arquivo SDS");
        console.error(err);
    } finally {
        setIsProcessing(false);
        if(sdsInputRef.current) sdsInputRef.current.value = '';
    }
  };

  const handleMapSelect = async (e: React.ChangeEvent<HTMLInputElement>) => {
      if (!e.target.files || !e.target.files[0]) return;
      try {
          const wb = await parseMapWorkbook(e.target.files[0]);
          setMapWorkbook(wb);
          setMapSheetNames(wb.SheetNames);
          if (wb.SheetNames.length > 0) {
              setSelectedSheet(wb.SheetNames[0]);
              setSheetModalOpen(true);
          }
      } catch (err) {
          alert("Erro ao abrir arquivo de Mapa");
          console.error(err);
      } finally {
          if (mapInputRef.current) mapInputRef.current.value = '';
      }
  };

  const confirmMapSheet = async () => {
      if (!mapWorkbook || !selectedSheet) return;
      setSheetModalOpen(false);
      setIsProcessing(true);
      setProgress(0);
      setProgressText("Mapeando dados (Turbo)...");

      setTimeout(() => {
          try {
              const ws = mapWorkbook.Sheets[selectedSheet];
              const manualRow = manualHeaderRow ? parseInt(manualHeaderRow) : null;
              const data = processMapSheet(ws, manualRow);
              setMapData(data);
              alert(`Mapa carregado: ${data.size} registros.`);
          } catch (err: any) {
              alert(`Erro no processamento: ${err.message}`);
          } finally {
              setIsProcessing(false);
              setMapWorkbook(null);
          }
      }, 100);
  };

  const handleDedupe = () => {
      if (!window.confirm("Remover duplicados baseados na OS?")) return;
      const seen = new Set();
      const unique: ReportItem[] = [];
      allData.forEach(item => {
          let key = String(item.os).trim();
          if (!key || key === '-' || key === 'N/A') {
               key = `NO_OS|${item.serie}|${item.dataCriacao}|${item.tipo}`;
          }
          if (!seen.has(key)) {
              seen.add(key);
              unique.push(item);
          }
      });
      alert(`Removidos ${allData.length - unique.length} duplicados.`);
      setAllData(unique);
  };

  // --- Filtering Logic ---

  const uniqueValues = useMemo(() => {
      const sets = {
          tipo: new Set<string>(),
          prod: new Set<string>(),
          status: new Set<string>(),
          situacao: new Set<string>(),
          conexao: new Set<string>(),
          mon: new Set<string>()
      };
      allData.forEach(item => {
          if (item.tipo) sets.tipo.add(item.tipo);
          if (item.equipProduzindo) sets.prod.add(item.equipProduzindo);
          if (item.statusOs) sets.status.add(item.statusOs);
          if (item.situacaoEquip) sets.situacao.add(item.situacaoEquip);
          if (item.tipoConexao) sets.conexao.add(item.tipoConexao);
          sets.mon.add(getSdsInfo(item.serie).status);
      });
      return {
          tipo: Array.from(sets.tipo).sort(),
          prod: Array.from(sets.prod).sort(),
          status: Array.from(sets.status).sort(),
          situacao: Array.from(sets.situacao).sort(),
          conexao: Array.from(sets.conexao).sort(),
          mon: Array.from(sets.mon).sort()
      };
      // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [allData, sdsData]);

  const filteredData = useMemo(() => {
      let data = [...allData];
      
      // Date Filters
      const stripTime = (d: string) => d ? new Date(d + "T00:00:00").getTime() : null;
      const startC = stripTime(filters.startCreation);
      const endC = stripTime(filters.endCreation);
      const startF = stripTime(filters.startConclusion);
      const endF = stripTime(filters.endConclusion);

      data = data.filter(item => {
          if (startC || endC) {
             const t = item._rawCriacao ? new Date(item._rawCriacao).setHours(0,0,0,0) : null;
             if (t) {
                if (startC && t < startC) return false;
                if (endC && t > endC) return false;
             } else if (startC || endC) return false; // Filter active but no date
          }
          if (startF || endF) {
             const t = item._rawConclusao ? new Date(item._rawConclusao).setHours(0,0,0,0) : null;
             if (t) {
                if (startF && t < startF) return false;
                if (endF && t > endF) return false;
             } else if (startF || endF) return false;
          }

          if (filters.selectedTypes.length && !filters.selectedTypes.includes(item.tipo)) return false;
          if (filters.selectedProds.length && !filters.selectedProds.includes(item.equipProduzindo)) return false;
          if (filters.selectedStatus.length && !filters.selectedStatus.includes(item.statusOs)) return false;
          if (filters.selectedSituacao.length && !filters.selectedSituacao.includes(item.situacaoEquip)) return false;
          if (filters.selectedConexao.length && !filters.selectedConexao.includes(item.tipoConexao)) return false;
          
          if (filters.selectedMon.length) {
              const status = getSdsInfo(item.serie).status;
              if (!filters.selectedMon.includes(status)) return false;
          }

          if (filters.search) {
              const term = filters.search.toLowerCase();
              if (filters.searchField === 'all') {
                  const sds = getSdsInfo(item.serie);
                  const map = getMapInfo(item.serie);
                  // Construct a searchable string
                  const fullStr = [
                      Object.values(item).join(' '),
                      sds.status,
                      Object.values(map).join(' ')
                  ].join(' ').toLowerCase();
                  if (!fullStr.includes(term)) return false;
              } else {
                  const val = String(item[filters.searchField] || '').toLowerCase();
                  if (!val.includes(term)) return false;
              }
          }

          return true;
      });

      return data;
  }, [allData, filters, sdsData, mapData]);

  useEffect(() => setCurrentPage(1), [filteredData.length]);

  const paginatedData = useMemo(() => {
      const start = (currentPage - 1) * itemsPerPage;
      return filteredData.slice(start, start + itemsPerPage);
  }, [filteredData, currentPage]);

  const totalPages = Math.ceil(filteredData.length / itemsPerPage);

  // --- Export ---
  
  const [exportCols, setExportCols] = useState<string[]>([]);
  const allExportOptions = useMemo(() => {
      const standard = ['Monitoramento', 'Ultima Atualizacao', 'Data Deteccao', 'Data Criação', 'Data Conclusão', 'OS', 'ID OS Corp', 'Tipo', 'Status OS', 'Contrato', 'Nº Série', 'Situação Equip.', 'Produzindo', 'Tipo Conexão', 'IP', 'Hostname', 'Bairro', 'Cidade', 'Filial', 'Origem'];
      const mapCols = mapColumnsConfig.map(c => c.label);
      return [...standard, ...mapCols];
  }, []);

  const handleExportClick = (type: 'csv' | 'xlsx' | 'pdf') => {
      setExportType(type);
      setExportCols(allExportOptions); // Default select all
      setExportModalOpen(true);
  };

  const confirmExport = () => {
      setExportModalOpen(false);
      if (!exportType) return;

      const exportData = filteredData.map(row => {
          const sds = getSdsInfo(row.serie);
          const map = getMapInfo(row.serie);
          
          const fullRow: any = {
              'Monitoramento': sds.status,
              'Ultima Atualizacao': sds.lastUpdate,
              'Data Deteccao': sds.detection,
              'Data Criação': row.dataCriacao,
              'Data Conclusão': row.dataConclusao,
              'OS': row.os,
              'ID OS Corp': row.idOsCorp,
              'Tipo': row.tipo,
              'Status OS': row.statusOs,
              'Contrato': row.contrato,
              'Nº Série': row.serie,
              'Situação Equip.': row.situacaoEquip,
              'Produzindo': row.equipProduzindo,
              'Tipo Conexão': row.tipoConexao,
              'IP': row.ip,
              'Hostname': row.hostname,
              'Bairro': row.bairro,
              'Cidade': row.cidade,
              'Filial': row.filial,
              'Origem': row.origem
          };
          mapColumnsConfig.forEach(c => fullRow[c.label] = map[c.key]);

          // Filter columns
          const filteredRow: any = {};
          exportCols.forEach(col => {
              filteredRow[col] = fullRow[col] || '';
          });
          return filteredRow;
      });

      const fname = `Relatorio_${new Date().toLocaleDateString('pt-BR').replace(/\//g,'-')}`;

      if (exportType === 'pdf') {
          const doc = new jsPDF({ orientation: 'landscape', unit: 'mm', format: 'a4' });
          doc.setFontSize(8);
          doc.text("Relatório Exportado", 14, 10);
          const head = [Object.keys(exportData[0])];
          const body = exportData.map(Object.values);
          autoTable(doc, {
             head,
             body,
             startY: 15,
             styles: { fontSize: 5 },
             theme: 'grid'
          });
          doc.save(`${fname}.pdf`);
      } else {
          const ws = XLSX.utils.json_to_sheet(exportData);
          const wb = XLSX.utils.book_new();
          XLSX.utils.book_append_sheet(wb, ws, "Dados");
          XLSX.writeFile(wb, `${fname}.${exportType}`);
      }
  };

  // --- Components ---
  
  const FilterDropdown = ({ title, options, selected, onChange, color = 'gray' }: any) => {
      const [open, setOpen] = useState(false);
      return (
          <div className="relative inline-block">
              <button 
                 onClick={() => setOpen(!open)}
                 className={`flex items-center gap-1 px-2 py-1 rounded text-xs font-bold uppercase border ${selected.length ? 'bg-blue-100 border-blue-300 text-blue-700' : `bg-${color}-50 border-${color}-200 text-${color}-600`}`}
              >
                  {title} {selected.length > 0 && `(${selected.length})`}
                  <svg className="w-3 h-3" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path strokeLinecap="round" strokeLinejoin="round" strokeWidth="2" d="M19 9l-7 7-7-7"/></svg>
              </button>
              {open && (
                  <>
                  <div className="fixed inset-0 z-30" onClick={() => setOpen(false)}></div>
                  <div className="absolute z-40 mt-1 w-48 bg-white border border-gray-200 shadow-lg rounded p-2 max-h-60 overflow-y-auto">
                      {options.length === 0 && <div className="text-xs text-gray-400 italic">Vazio</div>}
                      {options.map((opt: string) => (
                          <label key={opt} className="flex items-center gap-2 p-1 hover:bg-gray-50 cursor-pointer">
                              <input 
                                  type="checkbox" 
                                  checked={selected.includes(opt)}
                                  onChange={(e) => {
                                      if(e.target.checked) onChange([...selected, opt]);
                                      else onChange(selected.filter((s: string) => s !== opt));
                                  }}
                                  className="rounded border-gray-300 text-blue-600 h-3 w-3"
                              />
                              <span className="text-xs text-gray-700 truncate" title={opt}>{opt}</span>
                          </label>
                      ))}
                  </div>
                  </>
              )}
          </div>
      );
  };

  return (
    <div className="h-full flex flex-col bg-white">
      {/* Header */}
      <div className="bg-white shadow-sm p-3 z-20 flex flex-col gap-2 border-b border-gray-200">
        <div className="flex flex-col md:flex-row justify-between items-center gap-2">
           <div className="flex items-center gap-4">
               <h1 className="text-xl font-extrabold text-gray-800 tracking-tight bg-gradient-to-r from-blue-600 to-purple-600 bg-clip-text text-transparent">
                   Just Report
               </h1>
               <div className="flex items-center bg-gray-100 rounded-md p-1 border border-gray-200 gap-1">
                   <button onClick={() => handleExportClick('csv')} disabled={!allData.length} className="px-2 py-1 text-xs font-bold bg-white border rounded hover:text-green-600 disabled:opacity-50 transition">CSV</button>
                   <button onClick={() => handleExportClick('xlsx')} disabled={!allData.length} className="px-2 py-1 text-xs font-bold bg-white border rounded hover:text-green-600 disabled:opacity-50 transition">XLSX</button>
                   <button onClick={() => handleExportClick('pdf')} disabled={!allData.length} className="px-2 py-1 text-xs font-bold bg-white border rounded hover:text-red-600 disabled:opacity-50 transition">PDF</button>
               </div>
               <button 
                  onClick={handleDedupe}
                  disabled={!allData.length}
                  className="flex items-center gap-1 px-3 py-1 text-xs font-bold text-amber-700 bg-amber-50 border border-amber-200 rounded hover:bg-amber-100 disabled:opacity-50 transition"
               >
                  <svg className="w-3 h-3" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path strokeLinecap="round" strokeLinejoin="round" strokeWidth="2" d="M19 7l-.867 12.142A2 2 0 0116.138 21H7.862a2 2 0 01-1.995-1.858L5 7m5 4v6m4-6v6m1-10V4a1 1 0 00-1-1h-4a1 1 0 00-1 1v3M4 7h16"/></svg>
                  Limpar Duplicados
               </button>
           </div>

           <div className="flex items-center gap-2 bg-blue-50 px-3 py-1.5 rounded-md border border-blue-100 shadow-sm">
               <span className="text-[10px] font-extrabold text-blue-700 uppercase tracking-wide">1. Pasta Relatórios</span>
               <input 
                 ref={folderInputRef}
                 type="file" 
                 webkitdirectory="" 
                 directory="" 
                 multiple 
                 onChange={handleFolderSelect}
                 className="text-xs text-gray-600 file:bg-blue-600 file:text-white file:border-0 file:rounded file:px-2 file:py-0.5 file:text-[10px] file:font-bold hover:file:bg-blue-700 cursor-pointer"
               />
           </div>
        </div>

        {/* Secondary Inputs */}
        <div className="flex flex-wrap items-center gap-3 bg-gray-50 p-2 rounded-lg border border-gray-200 shadow-inner">
           <div className="flex items-center gap-2 pr-3 border-r border-gray-300">
               <span className="text-[10px] font-bold text-gray-500 uppercase">2. Base SDS</span>
               <input ref={sdsInputRef} type="file" accept=".xlsx,.xls,.csv" onChange={handleSdsSelect} className="text-xs text-gray-500 w-40 file:bg-gray-200 file:text-gray-700 file:border-0 file:rounded file:px-2 file:py-0.5 file:text-[10px] file:font-bold hover:file:bg-gray-300"/>
           </div>
           <div className="flex items-center gap-2 pr-3 border-r border-gray-300">
               <span className="text-[10px] font-bold text-purple-700 uppercase">3. Mapa</span>
               <input ref={mapInputRef} type="file" accept=".xlsx,.xls,.xlsb" onChange={handleMapSelect} className="text-xs text-gray-500 w-40 file:bg-purple-100 file:text-purple-700 file:border-0 file:rounded file:px-2 file:py-0.5 file:text-[10px] file:font-bold hover:file:bg-purple-200"/>
           </div>
           
           <div className="flex items-center gap-2 text-xs ml-auto">
               <div className="flex items-center gap-1" title="Alert Threshold (Days)">
                   <div className="w-2 h-2 rounded-full bg-yellow-400"></div>
                   <input 
                     type="number" 
                     value={filters.alertDays}
                     onChange={e => setFilters({...filters, alertDays: Number(e.target.value)})}
                     className="w-8 h-5 text-center border rounded text-[10px]"
                   />
               </div>
               <div className="flex items-center gap-1" title="Offline Threshold (Days)">
                   <div className="w-2 h-2 rounded-full bg-red-500"></div>
                   <input 
                     type="number" 
                     value={filters.offlineDays}
                     onChange={e => setFilters({...filters, offlineDays: Number(e.target.value)})}
                     className="w-8 h-5 text-center border rounded text-[10px]"
                   />
               </div>
               <div className="font-mono font-bold text-blue-700 ml-2">
                   {filteredData.length} registros
               </div>
           </div>
        </div>
        
        <ProgressBar visible={isProcessing} progress={progress} text={progressText} />

        {/* Filters Toolbar */}
        <div className="flex flex-wrap items-end gap-2 pt-1">
            {/* Date Filters */}
            <div className="flex items-center gap-1 border rounded p-1 h-[32px]">
                <span className="text-[9px] font-bold text-blue-600 uppercase px-1 border-r">Criação</span>
                <input type="date" className="text-[10px] outline-none bg-transparent w-20" onChange={e => setFilters({...filters, startCreation: e.target.value})} />
                <span className="text-gray-300">-</span>
                <input type="date" className="text-[10px] outline-none bg-transparent w-20" onChange={e => setFilters({...filters, endCreation: e.target.value})} />
            </div>
            <div className="flex items-center gap-1 border rounded p-1 h-[32px]">
                <span className="text-[9px] font-bold text-green-600 uppercase px-1 border-r">Conclusão</span>
                <input type="date" className="text-[10px] outline-none bg-transparent w-20" onChange={e => setFilters({...filters, startConclusion: e.target.value})} />
                <span className="text-gray-300">-</span>
                <input type="date" className="text-[10px] outline-none bg-transparent w-20" onChange={e => setFilters({...filters, endConclusion: e.target.value})} />
            </div>
            
            {/* Search */}
            <div className="flex flex-grow gap-1 h-[32px]">
                <select 
                  className="border rounded text-[10px] bg-gray-50 px-1 outline-none focus:border-blue-400"
                  value={filters.searchField}
                  onChange={e => setFilters({...filters, searchField: e.target.value})}
                >
                    <option value="all">Global</option>
                    <option value="serie">Série</option>
                    <option value="os">OS</option>
                    <option value="ip">IP</option>
                </select>
                <div className="relative flex-grow">
                    <input 
                      type="search" 
                      placeholder="Buscar..." 
                      className="w-full h-full border rounded pl-7 pr-2 text-xs outline-none focus:border-blue-500 focus:ring-1 focus:ring-blue-200 transition"
                      value={filters.search}
                      onChange={e => setFilters({...filters, search: e.target.value})}
                    />
                    <svg className="w-3 h-3 absolute left-2 top-1/2 -translate-y-1/2 text-gray-400" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path strokeLinecap="round" strokeLinejoin="round" strokeWidth="2" d="M21 21l-6-6m2-5a7 7 0 11-14 0 7 7 0 0114 0z"/></svg>
                </div>
            </div>
        </div>
      </div>

      {/* Table */}
      <div className="flex-grow overflow-auto custom-scrollbar relative">
          <table className="w-full border-collapse min-w-max">
              <thead className="bg-gray-50 sticky top-0 z-10 shadow-sm text-xs">
                  <tr>
                      {/* SDS Columns */}
                      <th className="px-3 py-2 text-left font-bold text-gray-800 border-b border-gray-200 bg-blue-50/80 backdrop-blur sticky left-0 z-20">
                          <div className="flex items-center justify-between gap-2">
                             <span>Monitoramento</span>
                             <FilterDropdown title="" options={uniqueValues.mon} selected={filters.selectedMon} onChange={(v: string[]) => setFilters({...filters, selectedMon: v})} color="blue"/>
                          </div>
                      </th>
                      <th className="px-3 py-2 text-left font-bold text-gray-800 border-b border-gray-200 bg-blue-50/80">Ult. Atualização</th>
                      <th className="px-3 py-2 text-left font-bold text-gray-800 border-b border-gray-200 bg-blue-50/80 border-r">Data Detecção</th>

                      {/* Map Columns */}
                      {mapColumnsConfig.map(col => (
                          <th key={col.key} className="px-3 py-2 text-left font-bold text-purple-900 border-b border-purple-100 bg-purple-50/80 whitespace-nowrap">
                              {col.label}
                          </th>
                      ))}

                      {/* Standard Columns */}
                      <th className="px-3 py-2 text-left font-bold text-gray-600 border-b bg-gray-50">Data Criação</th>
                      <th className="px-3 py-2 text-left font-bold text-gray-600 border-b bg-gray-50">Data Conclusão</th>
                      <th className="px-3 py-2 text-left font-bold text-gray-600 border-b bg-gray-50">OS</th>
                      <th className="px-3 py-2 text-left font-bold text-gray-600 border-b bg-gray-50">ID OS Corp</th>
                      
                      <th className="px-3 py-2 text-left font-bold text-gray-600 border-b bg-gray-50">
                          <div className="flex items-center justify-between gap-1">
                              <span>Tipo</span>
                              <FilterDropdown title="" options={uniqueValues.tipo} selected={filters.selectedTypes} onChange={(v: string[]) => setFilters({...filters, selectedTypes: v})} />
                          </div>
                      </th>
                      <th className="px-3 py-2 text-left font-bold text-gray-600 border-b bg-gray-50">
                          <div className="flex items-center justify-between gap-1">
                              <span>Status</span>
                              <FilterDropdown title="" options={uniqueValues.status} selected={filters.selectedStatus} onChange={(v: string[]) => setFilters({...filters, selectedStatus: v})} />
                          </div>
                      </th>
                      <th className="px-3 py-2 text-left font-bold text-gray-600 border-b bg-gray-50">Contrato</th>
                      <th className="px-3 py-2 text-left font-bold text-gray-600 border-b bg-gray-50">Série</th>
                      <th className="px-3 py-2 text-left font-bold text-gray-600 border-b bg-gray-50">
                          <div className="flex items-center justify-between gap-1">
                              <span>Situação</span>
                              <FilterDropdown title="" options={uniqueValues.situacao} selected={filters.selectedSituacao} onChange={(v: string[]) => setFilters({...filters, selectedSituacao: v})} />
                          </div>
                      </th>
                      <th className="px-3 py-2 text-left font-bold text-gray-600 border-b bg-gray-50">
                          <div className="flex items-center justify-between gap-1">
                              <span>Prod.</span>
                              <FilterDropdown title="" options={uniqueValues.prod} selected={filters.selectedProds} onChange={(v: string[]) => setFilters({...filters, selectedProds: v})} />
                          </div>
                      </th>
                      <th className="px-3 py-2 text-left font-bold text-gray-600 border-b bg-gray-50">
                          <div className="flex items-center justify-between gap-1">
                              <span>Conexão</span>
                              <FilterDropdown title="" options={uniqueValues.conexao} selected={filters.selectedConexao} onChange={(v: string[]) => setFilters({...filters, selectedConexao: v})} />
                          </div>
                      </th>
                      <th className="px-3 py-2 text-left font-bold text-gray-600 border-b bg-gray-50">IP</th>
                      <th className="px-3 py-2 text-left font-bold text-gray-600 border-b bg-gray-50">Hostname</th>
                      <th className="px-3 py-2 text-left font-bold text-gray-600 border-b bg-gray-50">Bairro</th>
                      <th className="px-3 py-2 text-left font-bold text-gray-600 border-b bg-gray-50">Cidade</th>
                      <th className="px-3 py-2 text-left font-bold text-gray-600 border-b bg-gray-50">Filial</th>
                      <th className="px-3 py-2 text-left font-bold text-gray-600 border-b bg-gray-50">Origem</th>
                  </tr>
              </thead>
              <tbody className="text-xs text-gray-700 bg-white divide-y divide-gray-100">
                  {paginatedData.length === 0 ? (
                      <tr><td colSpan={30} className="px-6 py-8 text-center text-gray-400 text-sm">Nenhum dado para exibir. Carregue relatórios para começar.</td></tr>
                  ) : (
                      paginatedData.map((row) => {
                          const sds = getSdsInfo(row.serie);
                          const map = getMapInfo(row.serie);
                          return (
                              <tr key={row.id} className="hover:bg-blue-50 transition-colors group">
                                  <td className={`px-3 py-2 whitespace-nowrap border-r border-gray-100 sticky left-0 group-hover:bg-blue-100 transition-colors ${sds.colorClass}`}>{sds.status}</td>
                                  <td className="px-3 py-2 whitespace-nowrap border-r border-gray-100">{sds.lastUpdate}</td>
                                  <td className="px-3 py-2 whitespace-nowrap border-r border-gray-200">{sds.detection}</td>
                                  
                                  {mapColumnsConfig.map(col => (
                                      <td key={col.key} className="px-3 py-2 whitespace-nowrap border-r border-purple-100 bg-purple-50/30 text-purple-900 group-hover:bg-purple-100/50" title={map[col.key]}>
                                          <div className="max-w-[200px] truncate">{map[col.key]}</div>
                                      </td>
                                  ))}

                                  <td className="px-3 py-2 whitespace-nowrap border-r border-gray-100">{row.dataCriacao}</td>
                                  <td className="px-3 py-2 whitespace-nowrap border-r border-gray-100">{row.dataConclusao}</td>
                                  <td className="px-3 py-2 whitespace-nowrap border-r border-gray-100">{row.os}</td>
                                  <td className="px-3 py-2 whitespace-nowrap border-r border-gray-100">{row.idOsCorp}</td>
                                  <td className="px-3 py-2 whitespace-nowrap border-r border-gray-100">{row.tipo}</td>
                                  <td className="px-3 py-2 whitespace-nowrap border-r border-gray-100">{row.statusOs}</td>
                                  <td className="px-3 py-2 whitespace-nowrap border-r border-gray-100">{row.contrato}</td>
                                  <td className="px-3 py-2 whitespace-nowrap border-r border-gray-100 font-mono">{row.serie}</td>
                                  <td className="px-3 py-2 whitespace-nowrap border-r border-gray-100">{row.situacaoEquip}</td>
                                  <td className="px-3 py-2 whitespace-nowrap border-r border-gray-100">{row.equipProduzindo}</td>
                                  <td className="px-3 py-2 whitespace-nowrap border-r border-gray-100">{row.tipoConexao}</td>
                                  <td className="px-3 py-2 whitespace-nowrap border-r border-gray-100">{row.ip}</td>
                                  <td className="px-3 py-2 whitespace-nowrap border-r border-gray-100">{row.hostname}</td>
                                  <td className="px-3 py-2 whitespace-nowrap border-r border-gray-100">{row.bairro}</td>
                                  <td className="px-3 py-2 whitespace-nowrap border-r border-gray-100">{row.cidade}</td>
                                  <td className="px-3 py-2 whitespace-nowrap border-r border-gray-100">{row.filial}</td>
                                  <td className="px-3 py-2 whitespace-nowrap">{row.origem}</td>
                              </tr>
                          );
                      })
                  )}
              </tbody>
          </table>
      </div>

      {/* Pagination */}
      <div className="p-2 border-t bg-gray-50 flex justify-between items-center z-20">
          <button 
             onClick={() => setCurrentPage(prev => Math.max(prev - 1, 1))}
             disabled={currentPage === 1}
             className="px-4 py-1 rounded bg-white border shadow-sm text-xs font-bold hover:bg-gray-100 disabled:opacity-50 transition"
          >
              Anterior
          </button>
          <span className="text-xs text-gray-600 font-medium">
              Página {currentPage} de {Math.max(totalPages, 1)}
          </span>
          <button 
             onClick={() => setCurrentPage(prev => Math.min(prev + 1, totalPages))}
             disabled={currentPage === totalPages || totalPages === 0}
             className="px-4 py-1 rounded bg-white border shadow-sm text-xs font-bold hover:bg-gray-100 disabled:opacity-50 transition"
          >
              Próximo
          </button>
      </div>

      {/* Map Sheet Selection Modal */}
      <Modal 
         isOpen={sheetModalOpen} 
         onClose={() => setSheetModalOpen(false)}
         title="Configurar Mapa"
         footer={
             <div className="flex gap-2">
                 <button onClick={() => setSheetModalOpen(false)} className="px-4 py-2 text-sm text-gray-600 hover:bg-gray-100 rounded transition">Cancelar</button>
                 <button onClick={confirmMapSheet} disabled={!selectedSheet} className="px-4 py-2 text-sm bg-blue-600 text-white rounded font-bold hover:bg-blue-700 disabled:opacity-50 transition">Confirmar</button>
             </div>
         }
      >
         <div className="flex flex-col gap-4">
             <div>
                 <label className="block text-xs font-bold text-gray-700 uppercase mb-2">Selecione a Aba:</label>
                 <div className="flex flex-col gap-1 max-h-40 overflow-y-auto border p-2 rounded bg-gray-50 custom-scrollbar">
                     {mapSheetNames.map(name => (
                         <button 
                            key={name}
                            onClick={() => setSelectedSheet(name)}
                            className={`text-left px-3 py-2 text-sm rounded border transition ${selectedSheet === name ? 'bg-blue-100 border-blue-300 text-blue-800 font-semibold' : 'bg-white border-gray-200 hover:bg-gray-100'}`}
                         >
                             {name}
                         </button>
                     ))}
                 </div>
             </div>
             <div>
                 <label className="block text-xs font-bold text-gray-700 uppercase mb-1">Linha do Cabeçalho (Opcional):</label>
                 <div className="flex items-center gap-2">
                     <input 
                       type="number" 
                       min="1" 
                       placeholder="Auto" 
                       value={manualHeaderRow}
                       onChange={e => setManualHeaderRow(e.target.value)}
                       className="w-24 p-2 text-sm border rounded focus:ring-2 focus:ring-blue-500 outline-none"
                     />
                     <span className="text-[10px] text-gray-500 leading-tight max-w-[200px]">Se a busca automática falhar, informe o número da linha (ex: 7).</span>
                 </div>
             </div>
         </div>
      </Modal>

      {/* Export Modal */}
      <Modal
         isOpen={exportModalOpen}
         onClose={() => setExportModalOpen(false)}
         title={`Exportar ${exportType?.toUpperCase()}`}
         footer={
             <div className="flex gap-2">
                 <button onClick={() => setExportModalOpen(false)} className="px-4 py-2 text-sm text-gray-600 hover:bg-gray-100 rounded transition">Cancelar</button>
                 <button onClick={confirmExport} className="px-4 py-2 text-sm bg-green-600 text-white rounded font-bold hover:bg-green-700 transition">Baixar Arquivo</button>
             </div>
         }
      >
         <div className="flex flex-col gap-4">
             <div className="flex gap-2">
                 <button onClick={() => setExportCols(allExportOptions)} className="text-xs px-3 py-1 bg-gray-100 hover:bg-gray-200 rounded border">Marcar Todos</button>
                 <button onClick={() => setExportCols([])} className="text-xs px-3 py-1 bg-gray-100 hover:bg-gray-200 rounded border">Desmarcar Todos</button>
             </div>
             <div className="grid grid-cols-2 gap-2">
                 {allExportOptions.map(col => (
                     <label key={col} className="flex items-center gap-2 p-2 border rounded hover:bg-gray-50 cursor-pointer">
                         <input 
                            type="checkbox"
                            checked={exportCols.includes(col)}
                            onChange={(e) => {
                                if(e.target.checked) setExportCols([...exportCols, col]);
                                else setExportCols(exportCols.filter(c => c !== col));
                            }}
                            className="rounded border-gray-300 text-green-600 w-4 h-4"
                         />
                         <span className="text-xs text-gray-700">{col}</span>
                     </label>
                 ))}
             </div>
         </div>
      </Modal>
    </div>
  );
};

export default App;