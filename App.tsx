import React, { useState, useEffect, useMemo, useRef, useCallback } from 'react';
import * as XLSX from 'xlsx';
import { jsPDF } from 'jspdf';
import autoTable from 'jspdf-autotable';
import { ReportItem, SdsInfo, NddInfo, MapInfo, CorporateInfo, CepInvalidEntry, FilterState, MapColumnConfig, ColumnDef } from './types';
import { readExcelFile, readNddCsv, processReportData, parseMapWorkbook, processMapSheet, mapColumnsConfig, processCorporateFile } from './services/excelService';
import { bulkLookupCeps } from './services/viacepService';
import { parseDateRobust } from './utils/dateUtils';
import { Modal } from './components/Modal';
import { ProgressBar } from './components/ProgressBar';
import { useDebounce } from './hooks/useDebounce';
import { Dashboard, DashboardStats, LocationBreakdown } from './components/Dashboard';

// Initial Column Definitions
const INITIAL_COLUMNS: ColumnDef[] = [
  // SDS — hidden until SDS base is loaded
  { id: 'mon', label: 'Monitoramento', visible: false, type: 'sds', key: 'status' },
  { id: 'lastUpdate', label: 'Ult. Atualização', visible: false, type: 'sds', key: 'lastUpdate' },
  { id: 'detection', label: 'Data Detecção', visible: false, type: 'sds', key: 'detection' },
  // NDD MPS — hidden until NDD base is loaded
  { id: 'nddMon', label: 'Monitoramento MPS', visible: false, type: 'ndd', key: 'status' },
  { id: 'nddLastUpdate', label: 'Ult. Leitura MPS', visible: false, type: 'ndd', key: 'lastUpdate' },
  { id: 'nddDays', label: 'Dias s/ Contador MPS', visible: false, type: 'ndd', key: 'daysWithoutMeters' },
  { id: 'nddAccounting', label: 'Contabilização MPS', visible: false, type: 'ndd', key: 'accountingStatus' },
  { id: 'nddConnectionType', label: 'Conexão MPS', visible: false, type: 'ndd', key: 'connectionType' },
  { id: 'nddMpsIp', label: 'IP MPS', visible: false, type: 'ndd', key: 'mpsIp' },
  // Map — hidden until Mapa is loaded
  ...mapColumnsConfig.map(c => ({ id: `map_${c.key}`, label: c.label, visible: false, type: 'map', key: c.key } as ColumnDef)),
  // Corporate — hidden until Corporate is loaded
  { id: 'corp_status',      label: 'Status Contrato',       visible: false, type: 'corporate', key: 'status' },
  { id: 'corp_modelo',      label: 'Modelo (Contrato)',      visible: false, type: 'corporate', key: 'modelo' },
  { id: 'corp_cliente',     label: 'Cliente (Contrato)',     visible: false, type: 'corporate', key: 'clienteInstalacao' },
  { id: 'corp_logradouro',  label: 'Logradouro (Ctto)',      visible: false, type: 'corporate', key: 'logradouro' },
  { id: 'corp_complemento', label: 'Complemento (Ctto)',     visible: false, type: 'corporate', key: 'complemento' },
  { id: 'corp_bairro',      label: 'Bairro (Ctto)',          visible: false, type: 'corporate', key: 'bairro' },
  { id: 'corp_cidade',      label: 'Cidade (Ctto)',          visible: false, type: 'corporate', key: 'cidade' },
  { id: 'corp_uf',          label: 'UF (Ctto)',              visible: false, type: 'corporate', key: 'uf' },
  { id: 'corp_cep',         label: 'CEP (Ctto)',             visible: false, type: 'corporate', key: 'cep' },
  { id: 'corp_dtInstal',    label: 'Dt. Instalação (Ctto)',  visible: false, type: 'corporate', key: 'dataInstalacao' },
  { id: 'corp_endereco',    label: 'Endereço Completo (Ctto)',visible: false, type: 'corporate', key: 'enderecoInstalacao' },
  // Standard — always visible
  { id: 'serie',         label: 'Série',           visible: true, type: 'standard', key: 'serie' },
  { id: 'contrato',      label: 'Contrato',         visible: true, type: 'standard', key: 'contrato' },
  { id: 'statusOs',      label: 'Status OS',        visible: true, type: 'standard', key: 'statusOs' },
  { id: 'tipo',          label: 'Tipo',             visible: true, type: 'standard', key: 'tipo' },
  { id: 'dataCriacao',   label: 'Data Criação',     visible: true, type: 'standard', key: 'dataCriacao' },
  { id: 'dataConclusao', label: 'Data Conclusão',   visible: true, type: 'standard', key: 'dataConclusao' },
  { id: 'os',            label: 'OS',               visible: true, type: 'standard', key: 'os' },
  { id: 'idOsCorp',      label: 'ID OS Corp',       visible: true, type: 'standard', key: 'idOsCorp' },
  { id: 'situacaoEquip', label: 'Situação Equip.',  visible: true, type: 'standard', key: 'situacaoEquip' },
  { id: 'equipProduzindo',label:'Produzindo',        visible: true, type: 'standard', key: 'equipProduzindo' },
  { id: 'tipoConexao',   label: 'Tipo Conexão',     visible: true, type: 'standard', key: 'tipoConexao' },
  { id: 'ip',            label: 'IP',               visible: true, type: 'standard', key: 'ip' },
  { id: 'hostname',      label: 'Hostname',         visible: true, type: 'standard', key: 'hostname' },
  { id: 'bairro',        label: 'Bairro',           visible: true, type: 'standard', key: 'bairro' },
  { id: 'cidade',        label: 'Cidade',           visible: true, type: 'standard', key: 'cidade' },
  { id: 'filial',        label: 'Filial',           visible: true, type: 'standard', key: 'filial' },
  { id: 'origem',        label: 'Origem',           visible: true, type: 'standard', key: 'origem' },
];

const App: React.FC = () => {
  // State
  const [allData, setAllData] = useState<ReportItem[]>([]);
  const [sdsData, setSdsData] = useState<Map<string, { rawLastUpdate: Date | null, rawDetection: Date | null }>>(new Map());
  const [nddData, setNddData] = useState<Map<string, { status: string, lastUpdate: string, daysWithoutMeters: string, rawLastUpdate: Date | null, accountingStatus: string, connectionType: string, mpsIp: string }>>(new Map());
  const [mapData, setMapData] = useState<Map<string, MapInfo>>(new Map());
  const [corporateData, setCorporateData] = useState<Map<string, CorporateInfo>>(new Map());
  const [useCepValidation, setUseCepValidation] = useState(true);
  const [cepValidationProgress, setCepValidationProgress] = useState<{ done: number; total: number } | null>(null);
  const cepAbortRef = useRef<AbortController | null>(null);

  const [isProcessing, setIsProcessing] = useState(false);
  const [progress, setProgress] = useState(0);
  const [progressText, setProgressText] = useState('');

  // UI State
  const [activeTab, setActiveTab] = useState<'table' | 'dashboard'>('table');

  // Visibility toggles — when unchecked the data is treated as empty for display/dashboard
  const [showSds, setShowSds] = useState(true);
  const [showNdd, setShowNdd] = useState(true);
  const [showMap, setShowMap] = useState(true);
  const [showCorp, setShowCorp] = useState(true);
  const [sheetModalOpen, setSheetModalOpen] = useState(false);
  const [exportModalOpen, setExportModalOpen] = useState(false);
  const [columnModalOpen, setColumnModalOpen] = useState(false);
  
  const [mapWorkbook, setMapWorkbook] = useState<XLSX.WorkBook | null>(null);
  const [mapSheetNames, setMapSheetNames] = useState<string[]>([]);
  const [selectedSheet, setSelectedSheet] = useState<string | null>(null);
  const [manualHeaderRow, setManualHeaderRow] = useState<string>('');
  
  // Export State
  const [exportType, setExportType] = useState<'csv' | 'xlsx' | 'pdf' | null>(null);
  const [exportCols, setExportCols] = useState<string[]>([]);
  
  // PDF Customization State
  const [pdfTitle, setPdfTitle] = useState("Just Report");
  const [pdfObservation, setPdfObservation] = useState("");
  const [pdfLogo, setPdfLogo] = useState<string | null>(null);
  
  // Table Columns State
  const [columns, setColumns] = useState<ColumnDef[]>(INITIAL_COLUMNS);

  // Filter State (Immediate UI)
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
    selectedMon: [],
    selectedNddMon: []
  });

  // Debounce the filters to avoid heavy calculation on every keystroke
  const debouncedFilters = useDebounce(filters, 400);

  const [currentPage, setCurrentPage] = useState(1);
  const itemsPerPage = 50;

  // File Inputs Refs
  const folderInputRef = useRef<HTMLInputElement>(null);
  const sdsInputRef = useRef<HTMLInputElement>(null);
  const mapInputRef = useRef<HTMLInputElement>(null);
  const nddInputRef = useRef<HTMLInputElement>(null);
  const corpInputRef = useRef<HTMLInputElement>(null);

  // --- Effects ---
  useEffect(() => {
    document.title = "Just Report";
  }, []);

  // Auto-show columns the first time each base is loaded
  const sdsColumnsRevealed  = useRef(false);
  const nddColumnsRevealed  = useRef(false);
  const mapColumnsRevealed  = useRef(false);
  const corpColumnsRevealed = useRef(false);

  useEffect(() => {
    if (sdsData.size > 0 && !sdsColumnsRevealed.current) {
      sdsColumnsRevealed.current = true;
      setColumns(prev => prev.map(c => c.type === 'sds' ? { ...c, visible: true } : c));
    }
  }, [sdsData.size]);

  useEffect(() => {
    if (nddData.size > 0 && !nddColumnsRevealed.current) {
      nddColumnsRevealed.current = true;
      setColumns(prev => prev.map(c => c.type === 'ndd' ? { ...c, visible: true } : c));
    }
  }, [nddData.size]);

  useEffect(() => {
    if (mapData.size > 0 && !mapColumnsRevealed.current) {
      mapColumnsRevealed.current = true;
      setColumns(prev => prev.map(c => c.type === 'map' ? { ...c, visible: true } : c));
    }
  }, [mapData.size]);

  useEffect(() => {
    if (corporateData.size > 0 && !corpColumnsRevealed.current) {
      corpColumnsRevealed.current = true;
      setColumns(prev => prev.map(c =>
        c.type === 'corporate' || c.id === 'serie' ? { ...c, visible: true } : c
      ));
    }
  }, [corporateData.size]);

  // Synthetic rows from Corporate when no reports are loaded (corporate-only mode)
  const corporateRows = useMemo((): ReportItem[] => {
    if (allData.length > 0 || corporateData.size === 0) return [];
    return Array.from(corporateData.values() as Iterable<CorporateInfo>).map((c): ReportItem => ({
      id: `corp-${c.serial}`,
      dataCriacao: '-', dataConclusao: '-',
      os: '-', idOsCorp: '-', tipo: '-',
      statusOs: c.status,
      contrato: c.clienteInstalacao || '-',
      serie: c.serial,
      situacaoEquip: '-', equipProduzindo: '-',
      tipoConexao: '-', ip: '-', hostname: '-',
      bairro: c.bairro || '-', cidade: c.cidade || '-', filial: c.uf || '-',
      origem: 'Corporate',
      _rawCriacao: null, _rawConclusao: null,
    }));
  }, [allData.length, corporateData]);

  // The data that drives the table — reports when loaded, corporate rows otherwise
  const effectiveData = allData.length > 0 ? allData : corporateRows;

  // --- Helpers (Memoized to avoid recreation) ---

  const getSdsInfo = useCallback((serial: string): SdsInfo => {
    const empty = { status: '-', colorClass: '', lastUpdate: '-', detection: '-', rawLastUpdate: null, rawDetection: null } as SdsInfo;
    if (!showSds || sdsData.size === 0 || !serial || serial === '-') return empty;
    
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
  }, [sdsData, showSds, filters.offlineDays, filters.alertDays]);

  const getNddInfo = useCallback((serial: string): NddInfo => {
    const empty = { status: '-', colorClass: '', lastUpdate: '-', daysWithoutMeters: '-', rawLastUpdate: null, accountingStatus: '-', connectionType: '-', mpsIp: '-' } as NddInfo;
    if (!showNdd || nddData.size === 0 || !serial || serial === '-') return empty;

    const key = String(serial).trim().toUpperCase();
    const record = nddData.get(key);

    if (!record) {
      return { ...empty, status: 'Não Monitorado', colorClass: 'bg-red-100 text-red-800 font-semibold' };
    }

    const { status, lastUpdate, daysWithoutMeters, rawLastUpdate, accountingStatus, connectionType, mpsIp } = record;
    const days = parseInt(daysWithoutMeters) || 0;
    const extras = { accountingStatus, connectionType, mpsIp };

    if (days > filters.offlineDays || status === 'NoMonitoringData') {
        return { status: 'Não Monitorado', colorClass: 'bg-red-100 text-red-800 font-semibold', lastUpdate, daysWithoutMeters, rawLastUpdate, ...extras };
    }
    if (days > filters.alertDays || status === 'RedEvent') {
        return { status: 'Alerta', colorClass: 'bg-yellow-100 text-yellow-800 font-semibold', lastUpdate, daysWithoutMeters, rawLastUpdate, ...extras };
    }

    return { status: 'Monitorado', colorClass: 'bg-green-100 text-green-800 font-semibold', lastUpdate, daysWithoutMeters, rawLastUpdate, ...extras };
  }, [nddData, showNdd, filters.offlineDays, filters.alertDays]);

  const getMapInfo = useCallback((serial: string): MapInfo => {
    if (!showMap || mapData.size === 0 || !serial || serial === '-') {
      const empty: any = {};
      mapColumnsConfig.forEach(c => empty[c.key] = '-');
      return empty;
    }
    const key = String(serial).trim().toUpperCase();
    return mapData.get(key) || mapColumnsConfig.reduce((acc, col) => ({...acc, [col.key]: 'N/A'}), {} as MapInfo);
  }, [mapData, showMap]);

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

  const handleNddSelect = async (e: React.ChangeEvent<HTMLInputElement>) => {
    if (!e.target.files || !e.target.files[0]) return;
    setIsProcessing(true);
    setProgressText("Lendo base NDD MPS...");
    
    try {
        const raw = await readNddCsv(e.target.files[0]);
        const newNddData = new Map();
        
        raw.forEach((row: any) => {
             const findVal = (keys: string[]) => {
                for (const k of keys) {
                    if (row[k] !== undefined && row[k] !== null && row[k] !== "") return row[k];
                    const upperK = k.toUpperCase();
                    const match = Object.keys(row).find(rk => rk.trim().toUpperCase() === upperK);
                    if (match && row[match] !== undefined && row[match] !== null && row[match] !== "") return row[match];
                }
                return null;
             };

             const serial = findVal(['Serial', 'Número de Série', 'Nº Série', 'Série', 'Numero de serie']);
             const lastUpdate = findVal(['Last meter', 'Última leitura', 'Ultima leitura', 'Data Leitura', 'Ultimo medidor']);
             const alertsStatus = findVal(['Alerts status', 'Status de Alerta', 'Status Alerta', 'Status de alertas']);
             const daysWithoutMeters = findVal(['Days without meters', 'Dias sem medidores', 'Dias sem leitura', 'Dias sem contadores']);
             const accountingStatusRaw = findVal(['Accounting status', 'Status Contabilização', 'Billing Status', 'Accounting Status']);
             const connectionTypeRaw = findVal(['Connection type', 'Tipo Conexão', 'Tipo de Conexão', 'Connection Type']);
             const mpsIpRaw = findVal(["Printer's address", "Printers address", 'IP Impressora', 'Endereco Impressora']);

             const accountingStatusMap: Record<string, string> = {
                 'BillingEnabled': 'Bilhetagem Ativa',
                 'NoBillingRecently': 'Sem Bilhetagem Recente',
                 'NoBillingData': 'Nunca Bilhetado',
             };
             const connectionTypeMap: Record<string, string> = {
                 'Network': 'Rede',
                 'Local': 'USB/Local',
             };
             const rawAccounting = String(accountingStatusRaw || '');
             const rawConnection = String(connectionTypeRaw || '');

             if (serial) {
                 const key = String(serial).trim().toUpperCase();
                 newNddData.set(key, {
                     status: String(alertsStatus || 'Unknown'),
                     lastUpdate: String(lastUpdate || '-'),
                     daysWithoutMeters: String(daysWithoutMeters || '0'),
                     rawLastUpdate: parseDateRobust(lastUpdate),
                     accountingStatus: accountingStatusMap[rawAccounting] || rawAccounting || '-',
                     connectionType: connectionTypeMap[rawConnection] || rawConnection || '-',
                     mpsIp: String(mpsIpRaw || '-'),
                 });
             }
        });
        setNddData(newNddData);
        alert(`Base NDD carregada: ${newNddData.size} registros.`);
    } catch (err) {
        alert("Erro ao ler arquivo NDD");
        console.error(err);
    } finally {
        setIsProcessing(false);
        if(nddInputRef.current) nddInputRef.current.value = '';
    }
  };

  const handleCorpSelect = async (e: React.ChangeEvent<HTMLInputElement>) => {
    if (!e.target.files || !e.target.files[0]) return;

    // Cancel any previous CEP validation still running
    cepAbortRef.current?.abort();
    cepAbortRef.current = null;
    setCepValidationProgress(null);

    setIsProcessing(true);
    setProgressText("Lendo base Corporate...");
    try {
      const data = await processCorporateFile(e.target.files[0]);
      setCorporateData(data);
      setIsProcessing(false);
      setProgressText('');
      if (corpInputRef.current) corpInputRef.current.value = '';

      if (!useCepValidation) return;

      // --- ViaCEP validation phase ---
      const allCeps = Array.from(data.values() as Iterable<CorporateInfo>)
        .map(c => c.cep)
        .filter(c => c.length === 8);

      if (allCeps.length === 0) return;

      const controller = new AbortController();
      cepAbortRef.current = controller;
      setCepValidationProgress({ done: 0, total: allCeps.length });
      setProgressText(`Validando CEPs via ViaCEP (0 / ${allCeps.length})...`);
      setIsProcessing(true);

      try {
        const results = await bulkLookupCeps(
          allCeps,
          (done, total) => {
            setCepValidationProgress({ done, total });
            setProgressText(`Validando CEPs via ViaCEP (${done} / ${total})...`);
          },
          controller.signal
        );

        // Merge results back into corporateData (mutable update on a new Map)
        setCorporateData(prev => {
          const next = new Map(prev);
          (next as Map<string, CorporateInfo>).forEach((info, key) => {
            const hit = results.get(info.cep);
            if (hit === undefined) return; // CEP had no digits or wasn't queried
            const updated: CorporateInfo = {
              ...info,
              cepStatus: hit ? 'valid' : 'invalid',
              // Overwrite address fields only if ViaCEP returned data
              ...(hit ? {
                logradouro: hit.logradouro || info.logradouro,
                complemento: hit.complemento || info.complemento,
                bairro: hit.bairro || info.bairro,
                cidade: hit.localidade || info.cidade,
                uf: hit.uf || info.uf,
              } : {}),
            };
            next.set(key, updated);
          });
          return next;
        });
      } catch {
        // aborted or network error — leave cepStatus as 'unchecked'
      } finally {
        setIsProcessing(false);
        setProgressText('');
        setCepValidationProgress(null);
        cepAbortRef.current = null;
      }

    } catch (err: any) {
      alert(`Erro ao ler arquivo Corporate:\n${err?.message || err}`);
      console.error(err);
      setIsProcessing(false);
      setProgressText('');
      if (corpInputRef.current) corpInputRef.current.value = '';
    }
  };

  const getCorporateInfo = useCallback((serial: string): CorporateInfo | null => {
    if (!showCorp || corporateData.size === 0 || !serial || serial === '-') return null;
    const key = String(serial).replace(/[^a-zA-Z0-9]/g, '').toUpperCase();
    return corporateData.get(key) || null;
  }, [corporateData, showCorp]);

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
      setProgressText("Mapeando dados...");

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
          mon: new Set<string>(),
          nddMon: new Set<string>()
      };
      
      const now = new Date();

      effectiveData.forEach(item => {
          if (item.tipo) sets.tipo.add(item.tipo);
          if (item.equipProduzindo) sets.prod.add(item.equipProduzindo);
          if (item.statusOs) sets.status.add(item.statusOs);
          if (item.situacaoEquip) sets.situacao.add(item.situacaoEquip);
          if (item.tipoConexao) sets.conexao.add(item.tipoConexao);
          
          let status = '-';
          const key = String(item.serie).trim().toUpperCase();
          const record = sdsData.get(key);
          if (!record) status = 'Não Monitorado';
          else if (!record.rawLastUpdate) status = 'Dados Incompletos';
          else {
              const diffDays = Math.ceil(Math.abs(now.getTime() - record.rawLastUpdate.getTime()) / (1000 * 60 * 60 * 24));
              if (diffDays > filters.offlineDays) status = 'Não Monitorado';
              else if (diffDays > filters.alertDays) status = 'Alerta';
              else status = 'Monitorado';
          }
          sets.mon.add(status);

          let nddStatus = '-';
          const nddRecord = nddData.get(key);
          if (!nddRecord) nddStatus = 'Não Monitorado';
          else {
              const days = parseInt(nddRecord.daysWithoutMeters) || 0;
              if (days > filters.offlineDays || nddRecord.status === 'NoMonitoringData') nddStatus = 'Não Monitorado';
              else if (days > filters.alertDays || nddRecord.status === 'RedEvent') nddStatus = 'Alerta';
              else nddStatus = 'Monitorado';
          }
          sets.nddMon.add(nddStatus);
      });
      return {
          tipo: Array.from(sets.tipo).sort(),
          prod: Array.from(sets.prod).sort(),
          status: Array.from(sets.status).sort(),
          situacao: Array.from(sets.situacao).sort(),
          conexao: Array.from(sets.conexao).sort(),
          mon: Array.from(sets.mon).sort(),
          nddMon: Array.from(sets.nddMon).sort()
      };
  }, [effectiveData, sdsData, nddData, filters.alertDays, filters.offlineDays]);

  const filteredData = useMemo(() => {
      let data = [...effectiveData];
      const activeFilters = debouncedFilters;

      const stripTime = (d: string) => d ? new Date(d + "T00:00:00").getTime() : null;
      const startC = stripTime(activeFilters.startCreation);
      const endC = stripTime(activeFilters.endCreation);
      const startF = stripTime(activeFilters.startConclusion);
      const endF = stripTime(activeFilters.endConclusion);
      
      const hasSearch = !!activeFilters.search;
      const term = activeFilters.search.toLowerCase();
      const isGlobal = activeFilters.searchField === 'all';
      
      const now = new Date();

      data = data.filter(item => {
          if (startC || endC) {
             const t = item._rawCriacao ? new Date(item._rawCriacao).setHours(0,0,0,0) : null;
             if (t) {
                if (startC && t < startC) return false;
                if (endC && t > endC) return false;
             } else if (startC || endC) return false; 
          }
          if (startF || endF) {
             const t = item._rawConclusao ? new Date(item._rawConclusao).setHours(0,0,0,0) : null;
             if (t) {
                if (startF && t < startF) return false;
                if (endF && t > endF) return false;
             } else if (startF || endF) return false;
          }

          if (activeFilters.selectedTypes.length && !activeFilters.selectedTypes.includes(item.tipo)) return false;
          if (activeFilters.selectedProds.length && !activeFilters.selectedProds.includes(item.equipProduzindo)) return false;
          if (activeFilters.selectedStatus.length && !activeFilters.selectedStatus.includes(item.statusOs)) return false;
          if (activeFilters.selectedSituacao.length && !activeFilters.selectedSituacao.includes(item.situacaoEquip)) return false;
          if (activeFilters.selectedConexao.length && !activeFilters.selectedConexao.includes(item.tipoConexao)) return false;
          
          if (activeFilters.selectedMon.length) {
              const key = String(item.serie).trim().toUpperCase();
              const record = sdsData.get(key);
              let status = 'Não Monitorado';
              
              if (!record) status = 'Não Monitorado';
              else if (!record.rawLastUpdate) status = 'Dados Incompletos';
              else {
                  const diffDays = Math.ceil(Math.abs(now.getTime() - record.rawLastUpdate.getTime()) / (1000 * 60 * 60 * 24));
                  if (diffDays > activeFilters.offlineDays) status = 'Não Monitorado';
                  else if (diffDays > activeFilters.alertDays) status = 'Alerta';
                  else status = 'Monitorado';
              }
              
              if (!activeFilters.selectedMon.includes(status)) return false;
          }

          if (activeFilters.selectedNddMon.length) {
              const key = String(item.serie).trim().toUpperCase();
              const nddRecord = nddData.get(key);
              let nddStatus = 'Não Monitorado';
              
              if (!nddRecord) nddStatus = 'Não Monitorado';
              else {
                  const days = parseInt(nddRecord.daysWithoutMeters) || 0;
                  if (days > activeFilters.offlineDays || nddRecord.status === 'NoMonitoringData') nddStatus = 'Não Monitorado';
                  else if (days > activeFilters.alertDays || nddRecord.status === 'RedEvent') nddStatus = 'Alerta';
                  else nddStatus = 'Monitorado';
              }
              
              if (!activeFilters.selectedNddMon.includes(nddStatus)) return false;
          }

          if (hasSearch) {
              if (isGlobal) {
                  if (
                      (item.os && item.os.toLowerCase().includes(term)) ||
                      (item.serie && item.serie.toLowerCase().includes(term)) ||
                      (item.contrato && item.contrato.toLowerCase().includes(term)) ||
                      (item.ip && item.ip.toLowerCase().includes(term)) ||
                      (item.tipo && item.tipo.toLowerCase().includes(term)) ||
                      (item.bairro && item.bairro.toLowerCase().includes(term))
                  ) return true;

                  const serialKey = String(item.serie).trim().toUpperCase();

                  if (mapData.size > 0) {
                      const mapInfo = mapData.get(serialKey);
                      if (mapInfo) {
                           for (const k in mapInfo) {
                               if (String(mapInfo[k]).toLowerCase().includes(term)) return true;
                           }
                      }
                  }

                  // Direct map lookup instead of calling getSdsInfo inside filter loop
                  if (sdsData.size > 0) {
                      const sdsRecord = sdsData.get(serialKey);
                      const sdsStatus = !sdsRecord ? 'não monitorado' : !sdsRecord.rawLastUpdate ? 'dados incompletos' : 'monitorado';
                      if (sdsStatus.includes(term)) return true;
                  }

                  if (nddData.size > 0) {
                      const nddRecord = nddData.get(serialKey);
                      if (nddRecord) {
                          if (nddRecord.accountingStatus.toLowerCase().includes(term)) return true;
                          if (nddRecord.connectionType.toLowerCase().includes(term)) return true;
                          if (nddRecord.mpsIp.toLowerCase().includes(term)) return true;
                      }
                  }

                  for (const k in item) {
                      if (k.startsWith('_')) continue;
                      if (String(item[k]).toLowerCase().includes(term)) return true;
                  }

                  return false;
              } else {
                  const val = String(item[activeFilters.searchField] || '').toLowerCase();
                  if (!val.includes(term)) return false;
              }
          }

          return true;
      });

      return data;
  }, [effectiveData, debouncedFilters, sdsData, nddData, mapData]);

  useEffect(() => setCurrentPage(1), [filteredData.length]);

  const startIndex = (currentPage - 1) * itemsPerPage;
  const endIndex = startIndex + itemsPerPage;
  const totalPages = Math.ceil(filteredData.length / itemsPerPage);

  // Pre-calculate sds/ndd/map info for the visible page only — avoids N*columns redundant calls
  const pagedData = useMemo(() => filteredData.slice(startIndex, endIndex), [filteredData, startIndex, endIndex]);

  const rowInfoCache = useMemo(() => {
    const cache = new Map<string, { sds: SdsInfo; ndd: NddInfo; map: MapInfo; corp: CorporateInfo | null }>();
    pagedData.forEach(row => {
      const key = String(row.serie).trim().toUpperCase();
      if (!cache.has(key)) {
        cache.set(key, {
          sds: getSdsInfo(row.serie),
          ndd: getNddInfo(row.serie),
          map: getMapInfo(row.serie),
          corp: getCorporateInfo(row.serie),
        });
      }
    });
    return cache;
  }, [pagedData, getSdsInfo, getNddInfo, getMapInfo, getCorporateInfo]);

  // --- Dashboard Stats ---

  const dashboardStats = useMemo((): DashboardStats => {
    const now = new Date();
    const sdsLoaded = sdsData.size > 0;
    const nddLoaded = nddData.size > 0;
    const corpLoaded = corporateData.size > 0 && showCorp;

    const sds = { monitored: 0, alert: 0, notMonitored: 0, incomplete: 0 };
    const ndd = { monitored: 0, alert: 0, notMonitored: 0 };
    const billing = { active: 0, noRecent: 0, never: 0 };
    const corp = { inContract: 0, outOfContract: 0, ativo: 0, inativo: 0 };
    const producingMap: Record<string, number> = {};
    const situacaoMap: Record<string, number> = {};
    const tipoMap: Record<string, number> = {};
    const cidadeMap: Record<string, number> = {};
    const contratoMap: Record<string, number> = {};
    const modeloMap: Record<string, number> = {};
    const ufMap: Record<string, number> = {};
    const connMap: Record<string, number> = {};

    // Location breakdown maps (keyed by contrato and cidade)
    const locContratoMap = new Map<string, LocationBreakdown>();
    const locCityMap     = new Map<string, LocationBreakdown>();

    const getOrCreateLoc = (map: Map<string, LocationBreakdown>, name: string): LocationBreakdown => {
      if (!map.has(name)) {
        map.set(name, {
          name,
          total: 0,
          sds: { monitored: 0, alert: 0, notMonitored: 0, noData: 0 },
          ndd: { monitored: 0, alert: 0, notMonitored: 0, noData: 0 },
          billing: { active: 0, noRecent: 0, never: 0 },
          situacao: [],
          serials: [],
        });
      }
      return map.get(name)!;
    };

    // Temporary situacao maps per location (finalized after loop)
    const locContratoSit = new Map<string, Record<string, number>>();
    const locCitySit     = new Map<string, Record<string, number>>();

    filteredData.forEach(item => {
      const key = String(item.serie).trim().toUpperCase();

      // Classify SDS status for this item
      let sdsStatus: 'monitored' | 'alert' | 'notMonitored' | 'noData' = 'noData';
      if (sdsLoaded) {
        const rec = sdsData.get(key);
        if (!rec) {
          sds.notMonitored++; sdsStatus = 'notMonitored';
        } else if (!rec.rawLastUpdate) {
          sds.incomplete++; sdsStatus = 'notMonitored';
        } else {
          const diff = Math.ceil(Math.abs(now.getTime() - rec.rawLastUpdate.getTime()) / 86400000);
          if (diff > filters.offlineDays)      { sds.notMonitored++; sdsStatus = 'notMonitored'; }
          else if (diff > filters.alertDays)   { sds.alert++;        sdsStatus = 'alert'; }
          else                                  { sds.monitored++;    sdsStatus = 'monitored'; }
        }
      }

      // Classify NDD status for this item
      let nddStatus: 'monitored' | 'alert' | 'notMonitored' | 'noData' = 'noData';
      let itemBilling: 'active' | 'noRecent' | 'never' | null = null;
      if (nddLoaded) {
        const rec = nddData.get(key);
        if (!rec) {
          ndd.notMonitored++; nddStatus = 'notMonitored';
        } else {
          const days = parseInt(rec.daysWithoutMeters) || 0;
          if (days > filters.offlineDays || rec.status === 'NoMonitoringData') { ndd.notMonitored++; nddStatus = 'notMonitored'; }
          else if (days > filters.alertDays || rec.status === 'RedEvent')      { ndd.alert++;        nddStatus = 'alert'; }
          else                                                                   { ndd.monitored++;    nddStatus = 'monitored'; }

          if (rec.accountingStatus === 'Bilhetagem Ativa')       { billing.active++;  itemBilling = 'active'; }
          else if (rec.accountingStatus === 'Sem Bilhetagem Recente') { billing.noRecent++; itemBilling = 'noRecent'; }
          else if (rec.accountingStatus === 'Nunca Bilhetado')    { billing.never++;   itemBilling = 'never'; }

          const conn = rec.connectionType || '-';
          connMap[conn] = (connMap[conn] || 0) + 1;
        }
      }

      if (corpLoaded) {
        const corpKey = String(item.serie || '').replace(/[^a-zA-Z0-9]/g, '').toUpperCase();
        const corpRec = corpKey ? corporateData.get(corpKey) : null;
        if (corpRec) {
          corp.inContract++;
          if (corpRec.status.toLowerCase().includes('ativo')) corp.ativo++;
          else corp.inativo++;
          if (corpRec.modelo) modeloMap[corpRec.modelo] = (modeloMap[corpRec.modelo] || 0) + 1;
          if (corpRec.uf) ufMap[corpRec.uf] = (ufMap[corpRec.uf] || 0) + 1;
        } else {
          corp.outOfContract++;
        }
      }

      const prod = item.equipProduzindo || '-';
      if (prod !== '-') producingMap[prod] = (producingMap[prod] || 0) + 1;

      const sit = item.situacaoEquip || '-';
      if (sit !== '-') situacaoMap[sit] = (situacaoMap[sit] || 0) + 1;

      const t = item.tipo || '-';
      if (t !== '-') tipoMap[t] = (tipoMap[t] || 0) + 1;

      const city = item.cidade || '-';
      if (city !== '-') cidadeMap[city] = (cidadeMap[city] || 0) + 1;

      const contrato = item.contrato || '-';
      if (contrato !== '-') contratoMap[contrato] = (contratoMap[contrato] || 0) + 1;

      // --- Accumulate location breakdowns ---
      const applyToLoc = (loc: LocationBreakdown, sitMap: Map<string, Record<string, number>>, locName: string) => {
        loc.total++;
        if (item.serie && item.serie !== '-') loc.serials.push(item.serie);

        if (sdsLoaded)  loc.sds[sdsStatus]++;
        else            loc.sds.noData++;

        if (nddLoaded)  loc.ndd[nddStatus]++;
        else            loc.ndd.noData++;

        if (itemBilling) loc.billing[itemBilling]++;

        if (sit !== '-') {
          if (!sitMap.has(locName)) sitMap.set(locName, {});
          const sm = sitMap.get(locName)!;
          sm[sit] = (sm[sit] || 0) + 1;
        }
      };

      if (contrato !== '-') {
        applyToLoc(getOrCreateLoc(locContratoMap, contrato), locContratoSit, contrato);
      }
      if (city !== '-') {
        applyToLoc(getOrCreateLoc(locCityMap, city), locCitySit, city);
      }
    });

    // Finalize situacao arrays for each location
    const finalizeLocs = (
      map: Map<string, LocationBreakdown>,
      sitMap: Map<string, Record<string, number>>
    ): LocationBreakdown[] =>
      Array.from(map.values())
        .map(loc => ({
          ...loc,
          situacao: Object.entries(sitMap.get(loc.name) || {})
            .sort((a, b) => b[1] - a[1])
            .slice(0, 6)
            .map(([name, count]) => ({ name, count })),
        }))
        .sort((a, b) => b.total - a.total);

    const sortedTop = (map: Record<string, number>, limit = 10) =>
      Object.entries(map).sort((a, b) => b[1] - a[1]).slice(0, limit).map(([name, count]) => ({ name, count }));

    return {
      total: filteredData.length,
      sdsLoaded,
      nddLoaded,
      corpLoaded,
      sds,
      ndd,
      billing,
      corp,
      producing: sortedTop(producingMap, 6),
      situacao: sortedTop(situacaoMap, 8),
      tipo: sortedTop(tipoMap, 8),
      byCidade: sortedTop(cidadeMap, 10),
      byContrato: sortedTop(contratoMap, 10),
      byModelo: sortedTop(modeloMap, 15),
      byUf: sortedTop(ufMap, 30),
      connectionType: sortedTop(connMap, 5),
      locationsByContrato: finalizeLocs(locContratoMap, locContratoSit),
      locationsByCity:     finalizeLocs(locCityMap, locCitySit),
      cepStats: corpLoaded ? (() => {
        const allCorp = Array.from(corporateData.values() as Iterable<CorporateInfo>);
        let valid = 0, invalid = 0, unchecked = 0;
        const invalidList: CepInvalidEntry[] = [];
        allCorp.forEach(c => {
          if (!c.cep || c.cep.length !== 8) { unchecked++; return; }
          if (c.cepStatus === 'valid')   { valid++; }
          else if (c.cepStatus === 'invalid') {
            invalid++;
            invalidList.push({ serial: c.serial, cep: c.cep, enderecoRaw: c.enderecoInstalacao, cidade: c.cidade, uf: c.uf, modelo: c.modelo });
          }
          else { unchecked++; }
        });
        return { total: valid + invalid, valid, invalid, unchecked, invalidList };
      })() : null,
    };
  }, [filteredData, sdsData, nddData, corporateData, showCorp, filters.alertDays, filters.offlineDays]);

  // --- Column Management ---

  const moveColumn = (index: number, direction: 'up' | 'down') => {
      const newCols = [...columns];
      const targetIndex = direction === 'up' ? index - 1 : index + 1;
      if (targetIndex < 0 || targetIndex >= newCols.length) return;
      
      const temp = newCols[index];
      newCols[index] = newCols[targetIndex];
      newCols[targetIndex] = temp;
      setColumns(newCols);
  };

  const toggleColumnVisibility = (id: string) => {
      setColumns(columns.map(c => c.id === id ? { ...c, visible: !c.visible } : c));
  };

  const visibleColumns = useMemo(() => columns.filter(c => c.visible), [columns]);

  // --- Export ---
  
  const handleExportClick = (type: 'csv' | 'xlsx' | 'pdf') => {
      setExportType(type);
      setExportCols(visibleColumns.map(c => c.label)); // Default to currently visible columns
      setExportModalOpen(true);
  };

  const handleLogoSelect = (e: React.ChangeEvent<HTMLInputElement>) => {
      if (e.target.files && e.target.files[0]) {
          const reader = new FileReader();
          reader.onload = (ev) => {
              if(ev.target?.result) {
                  setPdfLogo(ev.target.result as string);
              }
          };
          reader.readAsDataURL(e.target.files[0]);
      }
  };

  const confirmExport = () => {
      setExportModalOpen(false);
      if (!exportType) return;

      const exportData = filteredData.map(row => {
          const sds = getSdsInfo(row.serie);
          const ndd = getNddInfo(row.serie);
          const map = getMapInfo(row.serie);
          const corp = getCorporateInfo(row.serie);

          const rowData: any = {};

          columns.forEach(col => {
              if (!exportCols.includes(col.label)) return;

              let val = '';
              if (col.type === 'sds') {
                   if (col.key === 'status') val = sds.status;
                   else if (col.key === 'lastUpdate') val = sds.lastUpdate;
                   else if (col.key === 'detection') val = sds.detection;
              } else if (col.type === 'ndd') {
                   if (col.key === 'status') val = ndd.status;
                   else if (col.key === 'lastUpdate') val = ndd.lastUpdate;
                   else if (col.key === 'daysWithoutMeters') val = ndd.daysWithoutMeters;
                   else if (col.key === 'accountingStatus') val = ndd.accountingStatus;
                   else if (col.key === 'connectionType') val = ndd.connectionType;
                   else if (col.key === 'mpsIp') val = ndd.mpsIp;
              } else if (col.type === 'map') {
                   val = map[col.key] || '';
              } else if (col.type === 'corporate') {
                   val = corp ? (corp as any)[col.key] || '-' : (corporateData.size > 0 ? 'Fora do contrato' : '-');
              } else {
                   val = row[col.key] || '';
              }
              rowData[col.label] = val;
          });
          return rowData;
      });

      const fname = `JustReport_${new Date().toLocaleDateString('pt-BR').replace(/\//g,'-')}`;

      if (exportType === 'pdf') {
          const doc = new jsPDF({ orientation: 'landscape', unit: 'mm', format: 'a4' });
          const pageWidth = doc.internal.pageSize.width;
          
          // Header Config
          const title = pdfTitle || "Just Report";
          const margin = 14;
          let currentY = 15;

          // Logo Rendering (Right aligned)
          let logoHeight = 0;
          if (pdfLogo) {
               const logoW = 30; // 30mm width
               const logoRatio = 0.5; // Aspect ratio assumption or calculate from img
               // Since we don't have natural dimensions easily in jsPDF without loading Image object, we assume standard aspect or square.
               // Better approach: Fit in box 30x15
               const logoH = 15;
               logoHeight = logoH;
               
               doc.addImage(pdfLogo, 'PNG', pageWidth - margin - logoW, 10, logoW, logoH);
          }

          // Title Rendering
          doc.setFontSize(14);
          doc.setTextColor(40);
          doc.text(title, margin, currentY);
          currentY += 6;

          // Observations Rendering
          if (pdfObservation) {
              doc.setFontSize(9);
              doc.setTextColor(100);
              // Split text to fit page width minus margins and potential logo space
              const maxWidth = pageWidth - (margin * 2);
              const lines = doc.splitTextToSize(pdfObservation, maxWidth);
              doc.text(lines, margin, currentY);
              currentY += (lines.length * 4) + 2;
          }

          // Ensure table starts below logo if logo is taller than text
          if (pdfLogo) {
              const minTableY = 10 + logoHeight + 5; 
              if (currentY < minTableY) currentY = minTableY;
          } else {
              currentY += 2;
          }

          if (exportData.length > 0) {
              const head = [Object.keys(exportData[0])];
              const body = exportData.map(Object.values);
              autoTable(doc, {
                  head,
                  body,
                  startY: currentY,
                  styles: { fontSize: 6, cellPadding: 1 },
                  headStyles: { fillColor: [37, 99, 235] }, // Blue-600
                  theme: 'grid',
                  margin: { left: margin, right: margin }
              });
          }
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
          <div className="relative inline-block ml-1">
              <button 
                 onClick={() => setOpen(!open)}
                 className={`p-0.5 rounded hover:bg-${color}-200 transition-colors ${selected.length ? 'text-blue-600 bg-blue-50' : 'text-gray-400'}`}
              >
                  <svg className="w-3 h-3" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path strokeLinecap="round" strokeLinejoin="round" strokeWidth="2" d="M19 9l-7 7-7-7"/></svg>
              </button>
              {open && (
                  <>
                  <div className="fixed inset-0 z-30" onClick={() => setOpen(false)}></div>
                  <div className="absolute z-40 mt-1 w-48 bg-white border border-gray-200 shadow-lg rounded p-2 max-h-60 overflow-y-auto left-0">
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

  const renderHeaderCell = useCallback((col: ColumnDef, index: number) => {
      let content = <span>{col.label}</span>;
      let bgColor = 'bg-gray-50';
      let textColor = 'text-gray-600';
      let borderColor = 'border-gray-200';

      const addFilter = (options: string[], selected: string[], onChange: (v: string[]) => void, color?: string) => (
          <div className="flex items-center justify-between gap-1">
              <span>{col.label}</span>
              <FilterDropdown options={options} selected={selected} onChange={onChange} color={color} />
          </div>
      );

      if (col.type === 'sds') {
          bgColor = 'bg-blue-50/80';
          textColor = 'text-gray-800';
          if (col.key === 'status') {
              content = addFilter(uniqueValues.mon, filters.selectedMon, (v) => setFilters(prev => ({...prev, selectedMon: v})), 'blue');
          }
      } else if (col.type === 'ndd') {
          bgColor = 'bg-green-50/80';
          textColor = 'text-gray-800';
          if (col.key === 'status') {
              content = addFilter(uniqueValues.nddMon, filters.selectedNddMon, (v) => setFilters(prev => ({...prev, selectedNddMon: v})), 'green');
          }
      } else if (col.type === 'map') {
          bgColor = 'bg-purple-50/80';
          borderColor = 'border-purple-100';
          textColor = 'text-purple-900';
      } else if (col.type === 'corporate') {
          bgColor = 'bg-amber-50/80';
          borderColor = 'border-amber-100';
          textColor = 'text-amber-900';
      } else {
          if (col.key === 'tipo') content = addFilter(uniqueValues.tipo, filters.selectedTypes, v => setFilters(prev => ({...prev, selectedTypes: v})));
          else if (col.key === 'statusOs') content = addFilter(uniqueValues.status, filters.selectedStatus, v => setFilters(prev => ({...prev, selectedStatus: v})));
          else if (col.key === 'situacaoEquip') content = addFilter(uniqueValues.situacao, filters.selectedSituacao, v => setFilters(prev => ({...prev, selectedSituacao: v})));
          else if (col.key === 'equipProduzindo') content = addFilter(uniqueValues.prod, filters.selectedProds, v => setFilters(prev => ({...prev, selectedProds: v})));
          else if (col.key === 'tipoConexao') content = addFilter(uniqueValues.conexao, filters.selectedConexao, v => setFilters(prev => ({...prev, selectedConexao: v})));
      }

      const stickyClass = index === 0 ? 'sticky left-0 z-20 shadow-[2px_0_5px_-2px_rgba(0,0,0,0.1)]' : '';

      return (
          <th 
            key={col.id} 
            className={`px-3 py-2 text-left font-bold border-b whitespace-nowrap ${bgColor} ${textColor} ${borderColor} ${stickyClass}`}
          >
              {content}
          </th>
      );
  }, [uniqueValues, filters]);

  const renderRowCell = useCallback((row: ReportItem, col: ColumnDef, index: number) => {
      const cacheKey = String(row.serie).trim().toUpperCase();
      const cached = rowInfoCache.get(cacheKey);
      const sds = cached?.sds ?? getSdsInfo(row.serie);
      const ndd = cached?.ndd ?? getNddInfo(row.serie);
      const map = cached?.map ?? getMapInfo(row.serie);
      const corp = cached !== undefined ? cached.corp : getCorporateInfo(row.serie);
      let val: React.ReactNode = '';
      let cellClass = '';
      let title = '';

      if (col.type === 'sds') {
          cellClass = 'border-r border-gray-100';
          if (col.key === 'status') {
               val = sds.status;
               cellClass += ` ${sds.colorClass}`;
          } else if (col.key === 'lastUpdate') val = sds.lastUpdate;
          else if (col.key === 'detection') val = sds.detection;
      } else if (col.type === 'ndd') {
          cellClass = 'border-r border-gray-100';
          if (col.key === 'status') {
               val = ndd.status;
               cellClass += ` ${ndd.colorClass}`;
          } else if (col.key === 'lastUpdate') val = ndd.lastUpdate;
          else if (col.key === 'daysWithoutMeters') val = ndd.daysWithoutMeters;
          else if (col.key === 'accountingStatus') {
               val = ndd.accountingStatus;
               if (ndd.accountingStatus === 'Bilhetagem Ativa') cellClass += ' bg-green-50 text-green-700 font-semibold';
               else if (ndd.accountingStatus === 'Sem Bilhetagem Recente') cellClass += ' bg-yellow-50 text-yellow-700 font-semibold';
               else if (ndd.accountingStatus === 'Nunca Bilhetado') cellClass += ' bg-red-50 text-red-700 font-semibold';
          } else if (col.key === 'connectionType') val = ndd.connectionType;
          else if (col.key === 'mpsIp') val = ndd.mpsIp;
      } else if (col.type === 'map') {
          val = map[col.key];
          title = String(val);
          cellClass = 'border-r border-purple-100 bg-purple-50/30 text-purple-900 group-hover:bg-purple-100/50';
          val = <div className="max-w-[200px] truncate">{val}</div>;
      } else if (col.type === 'corporate') {
          cellClass = 'border-r border-amber-100 bg-amber-50/30 text-amber-900 group-hover:bg-amber-100/50';
          if (!corp) {
              val = corporateData.size > 0 ? <span className="text-xs text-gray-400 italic">Fora do contrato</span> : '-';
          } else if (col.key === 'status') {
              const isAtivo = corp.status.toLowerCase().includes('ativo');
              val = <span className={`text-xs font-semibold px-1.5 py-0.5 rounded ${isAtivo ? 'bg-green-100 text-green-700' : 'bg-red-100 text-red-700'}`}>{corp.status || '-'}</span>;
          } else {
              val = (corp as any)[col.key] || '-';
              title = String(val);
              val = <div className="max-w-[200px] truncate">{val}</div>;
          }
      } else {
          val = row[col.key];
          if (col.key === 'serie') cellClass = "font-mono";
          cellClass += " border-r border-gray-100 text-gray-700";
      }

      const stickyClass = index === 0 ? 'sticky left-0 group-hover:bg-blue-100 bg-white transition-colors z-10 shadow-[2px_0_5px_-2px_rgba(0,0,0,0.1)]' : '';
      
      if (index === 0) {
           cellClass += " bg-white"; 
      }

      return (
          <td key={col.id} className={`px-3 py-2 whitespace-nowrap ${cellClass} ${stickyClass}`} title={title}>
              {val}
          </td>
      );
  }, [getSdsInfo, getNddInfo, getMapInfo, getCorporateInfo, rowInfoCache, corporateData]);

  return (
    <div className="h-full flex flex-col bg-white">
      {/* Header */}
      <div className="bg-white shadow-sm p-3 z-30 flex flex-col gap-2 border-b border-gray-200">
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
                  onClick={() => setColumnModalOpen(true)}
                  className="flex items-center gap-1 px-3 py-1 text-xs font-bold text-gray-700 bg-white border border-gray-300 rounded hover:bg-gray-50 transition"
                  title="Configurar Colunas"
               >
                   <svg className="w-4 h-4" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path strokeLinecap="round" strokeLinejoin="round" strokeWidth="2" d="M10.325 4.317c.426-1.756 2.924-1.756 3.35 0a1.724 1.724 0 002.573 1.066c1.543-.94 3.31.826 2.37 2.37a1.724 1.724 0 001.065 2.572c1.756.426 1.756 2.924 0 3.35a1.724 1.724 0 00-1.066 2.573c.94 1.543-.826 3.31-2.37 2.37a1.724 1.724 0 00-2.572 1.065c-.426 1.756-2.924 1.756-3.35 0a1.724 1.724 0 00-2.573-1.066c-1.543.94-3.31-.826-2.37-2.37a1.724 1.724 0 00-1.065-2.572c-1.756-.426-1.756-2.924 0-3.35a1.724 1.724 0 001.066-2.573c-.94-1.543.826-3.31 2.37-2.37.996.608 2.296.07 2.572-1.065z"/><path strokeLinecap="round" strokeLinejoin="round" strokeWidth="2" d="M15 12a3 3 0 11-6 0 3 3 0 016 0z"/></svg>
                   Colunas
               </button>

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
                 {...({ webkitdirectory: "", directory: "" } as any)}
                 multiple 
                 onChange={handleFolderSelect}
                 className="text-xs text-gray-600 file:bg-blue-600 file:text-white file:border-0 file:rounded file:px-2 file:py-0.5 file:text-[10px] file:font-bold hover:file:bg-blue-700 cursor-pointer"
               />
           </div>
        </div>

        {/* Secondary Inputs */}
        <div className="flex flex-wrap items-center gap-3 bg-gray-50 p-2 rounded-lg border border-gray-200 shadow-inner">
           <div className="flex items-center gap-2 pr-3 border-r border-gray-300">
               <label className="flex items-center gap-1 cursor-pointer" title={showSds ? 'Ocultar dados SDS' : 'Exibir dados SDS'}>
                 <input type="checkbox" checked={showSds} onChange={e => setShowSds(e.target.checked)}
                   className="w-3.5 h-3.5 rounded border-gray-400 text-blue-600 focus:ring-blue-400 focus:ring-1 cursor-pointer" />
                 <span className="text-[10px] font-bold text-gray-500 uppercase">2. Base SDS</span>
                 {sdsData.size > 0 && (
                   <span className={`text-[9px] font-bold text-white rounded-full px-1.5 py-0.5 transition-colors ${showSds ? 'bg-blue-500' : 'bg-gray-400'}`} title={`${sdsData.size} registros SDS`}>{sdsData.size}</span>
                 )}
               </label>
               <input ref={sdsInputRef} type="file" accept=".xlsx,.xls,.csv" onChange={handleSdsSelect} className="text-xs text-gray-500 w-40 file:bg-gray-200 file:text-gray-700 file:border-0 file:rounded file:px-2 file:py-0.5 file:text-[10px] file:font-bold hover:file:bg-gray-300"/>
           </div>
           <div className="flex items-center gap-2 pr-3 border-r border-gray-300">
               <label className="flex items-center gap-1 cursor-pointer" title={showNdd ? 'Ocultar dados MPS' : 'Exibir dados MPS'}>
                 <input type="checkbox" checked={showNdd} onChange={e => setShowNdd(e.target.checked)}
                   className="w-3.5 h-3.5 rounded border-gray-400 text-green-600 focus:ring-green-400 focus:ring-1 cursor-pointer" />
                 <span className="text-[10px] font-bold text-green-700 uppercase">3. Base NDD MPS</span>
                 {nddData.size > 0 && (
                   <span className={`text-[9px] font-bold text-white rounded-full px-1.5 py-0.5 transition-colors ${showNdd ? 'bg-green-500' : 'bg-gray-400'}`} title={`${nddData.size} registros MPS`}>{nddData.size}</span>
                 )}
               </label>
               <input ref={nddInputRef} type="file" accept=".csv" onChange={handleNddSelect} className="text-xs text-gray-500 w-40 file:bg-green-100 file:text-green-700 file:border-0 file:rounded file:px-2 file:py-0.5 file:text-[10px] file:font-bold hover:file:bg-green-200"/>
           </div>
           <div className="flex items-center gap-2 pr-3 border-r border-gray-300">
               <label className="flex items-center gap-1 cursor-pointer" title={showMap ? 'Ocultar dados Mapa' : 'Exibir dados Mapa'}>
                 <input type="checkbox" checked={showMap} onChange={e => setShowMap(e.target.checked)}
                   className="w-3.5 h-3.5 rounded border-gray-400 text-purple-600 focus:ring-purple-400 focus:ring-1 cursor-pointer" />
                 <span className="text-[10px] font-bold text-purple-700 uppercase">4. Mapa</span>
                 {mapData.size > 0 && (
                   <span className={`text-[9px] font-bold text-white rounded-full px-1.5 py-0.5 transition-colors ${showMap ? 'bg-purple-500' : 'bg-gray-400'}`} title={`${mapData.size} registros Mapa`}>{mapData.size}</span>
                 )}
               </label>
               <input ref={mapInputRef} type="file" accept=".xlsx,.xls,.xlsb" onChange={handleMapSelect} className="text-xs text-gray-500 w-40 file:bg-purple-100 file:text-purple-700 file:border-0 file:rounded file:px-2 file:py-0.5 file:text-[10px] file:font-bold hover:file:bg-purple-200"/>
           </div>
           <div className="flex flex-col gap-1 pr-3 border-r border-gray-300">
             <div className="flex items-center gap-2">
               <label className="flex items-center gap-1 cursor-pointer" title={showCorp ? 'Ocultar dados Contrato' : 'Exibir dados Contrato'}>
                 <input type="checkbox" checked={showCorp} onChange={e => setShowCorp(e.target.checked)}
                   className="w-3.5 h-3.5 rounded border-gray-400 text-amber-600 focus:ring-amber-400 focus:ring-1 cursor-pointer" />
                 <span className="text-[10px] font-bold text-amber-700 uppercase">5. Contrato</span>
                 {corporateData.size > 0 && (
                   <span className={`text-[9px] font-bold text-white rounded-full px-1.5 py-0.5 transition-colors ${showCorp ? 'bg-amber-500' : 'bg-gray-400'}`} title={`${corporateData.size} equipamentos no contrato`}>{corporateData.size}</span>
                 )}
               </label>
               <input ref={corpInputRef} type="file" accept=".xlsx,.xls,.xlsb,.xls" onChange={handleCorpSelect} className="text-xs text-gray-500 w-40 file:bg-amber-100 file:text-amber-700 file:border-0 file:rounded file:px-2 file:py-0.5 file:text-[10px] file:font-bold hover:file:bg-amber-200"/>
             </div>
             <div className="flex items-center gap-2">
               <label className="flex items-center gap-1 cursor-pointer select-none" title="Quando ativo, valida os CEPs do contrato via ViaCEP e preenche os campos de endereço">
                 <input type="checkbox" checked={useCepValidation} onChange={e => setUseCepValidation(e.target.checked)}
                   className="w-3 h-3 rounded border-gray-400 text-amber-600 focus:ring-amber-400 focus:ring-1 cursor-pointer" />
                 <span className="text-[9px] text-amber-700">Validar CEPs via ViaCEP</span>
               </label>
               {cepValidationProgress && (
                 <span className="text-[9px] text-amber-600 font-semibold animate-pulse">
                   {cepValidationProgress.done}/{cepValidationProgress.total} CEPs…
                 </span>
               )}
             </div>
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
                    <option value="contrato">Contrato</option>
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

      {/* Tab Bar */}
      <div className="bg-white border-b border-gray-200 px-4 flex items-center gap-1 flex-shrink-0 z-20">
        {([
          {
            id: 'table' as const,
            label: 'Relatórios',
            activeColor: 'border-blue-600 text-blue-600',
            badge: null as React.ReactNode,
            icon: <svg className="w-3.5 h-3.5" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path strokeLinecap="round" strokeLinejoin="round" strokeWidth="2" d="M3 10h18M3 14h18M3 6h18M3 18h18"/></svg>,
          },
          {
            id: 'dashboard' as const,
            label: 'Dashboard',
            activeColor: 'border-purple-600 text-purple-600',
            badge: filteredData.length > 0
              ? <span className={`ml-1 text-[9px] font-bold px-1.5 py-0.5 rounded-full ${activeTab === 'dashboard' ? 'bg-purple-100 text-purple-700' : 'bg-gray-100 text-gray-500'}`}>{filteredData.length.toLocaleString('pt-BR')}</span>
              : null,
            icon: <svg className="w-3.5 h-3.5" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path strokeLinecap="round" strokeLinejoin="round" strokeWidth="2" d="M9 19v-6a2 2 0 00-2-2H5a2 2 0 00-2 2v6a2 2 0 002 2h2a2 2 0 002-2zm0 0V9a2 2 0 012-2h2a2 2 0 012 2v10m-6 0a2 2 0 002 2h2a2 2 0 002-2m0 0V5a2 2 0 012-2h2a2 2 0 012 2v14a2 2 0 01-2 2h-2a2 2 0 01-2-2z"/></svg>,
          },
        ]).map(tab => (
          <button
            key={tab.id}
            onClick={() => setActiveTab(tab.id)}
            className={`flex items-center gap-1.5 px-3 py-2 text-xs font-bold border-b-2 -mb-px transition-colors ${
              activeTab === tab.id
                ? tab.activeColor
                : 'border-transparent text-gray-500 hover:text-gray-700'
            }`}
          >
            {tab.icon}
            {tab.label}
            {tab.badge}
          </button>
        ))}
      </div>

      {activeTab === 'table' && (
        <>
          {/* Corporate-only mode banner */}
          {allData.length === 0 && corporateData.size > 0 && (
            <div className="flex items-center gap-2 bg-amber-50 border-b border-amber-200 px-4 py-2 text-xs text-amber-800 flex-shrink-0">
              <svg className="w-3.5 h-3.5 flex-shrink-0 text-amber-600" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path strokeLinecap="round" strokeLinejoin="round" strokeWidth="2" d="M13 16h-1v-4h-1m1-4h.01M21 12a9 9 0 11-18 0 9 9 0 0118 0z"/></svg>
              <span>
                <strong>Modo Contrato</strong> — exibindo {corporateData.size.toLocaleString('pt-BR')} equipamentos do arquivo de contrato.
                Carregue a <strong>Pasta Relatórios</strong> para cruzar com OS abertas.
              </span>
            </div>
          )}
          {/* Table */}
          <div className="flex-grow overflow-auto custom-scrollbar relative">
              <table className="w-full border-collapse min-w-max">
                  <thead className="bg-gray-50 sticky top-0 z-20 shadow-sm text-xs">
                      <tr>
                          {visibleColumns.map((col, index) => renderHeaderCell(col, index))}
                      </tr>
                  </thead>
                  <tbody className="text-xs text-gray-700 bg-white divide-y divide-gray-100">
                      {filteredData.length === 0 ? (
                          <tr><td colSpan={visibleColumns.length} className="px-6 py-8 text-center text-gray-400 text-sm">
                            {corporateData.size > 0 ? 'Nenhum equipamento corresponde ao filtro atual.' : 'Nenhum dado para exibir. Carregue a Pasta Relatórios ou o arquivo de Contrato para começar.'}
                          </td></tr>
                      ) : (
                          pagedData.map((row) => (
                              <tr key={row.id} className="hover:bg-blue-50 transition-colors group">
                                  {visibleColumns.map((col, index) => renderRowCell(row, col, index))}
                              </tr>
                          ))
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
        </>
      )}
      {activeTab === 'dashboard' && <Dashboard stats={dashboardStats} />}

      {/* Column Config Modal */}
      <Modal 
          isOpen={columnModalOpen}
          onClose={() => setColumnModalOpen(false)}
          title="Configurar Colunas"
          footer={
             <div className="flex justify-end">
                <button onClick={() => setColumnModalOpen(false)} className="px-4 py-2 text-sm bg-blue-600 text-white rounded font-bold hover:bg-blue-700 transition">Fechar</button>
             </div>
          }
      >
         <div className="flex flex-col gap-1">
             <div className="text-xs text-gray-500 mb-2">Marque para exibir. Use as setas para reordenar.</div>
             {columns.map((col, idx) => (
                 <div key={col.id} className={`flex items-center justify-between p-2 rounded border border-gray-100 ${col.visible ? 'bg-white' : 'bg-gray-50 opacity-70'}`}>
                     <label className="flex items-center gap-2 cursor-pointer flex-grow">
                         <input 
                            type="checkbox" 
                            checked={col.visible}
                            onChange={() => toggleColumnVisibility(col.id)}
                            className="rounded border-gray-300 text-blue-600 w-4 h-4 focus:ring-blue-500"
                         />
                         <span className={`text-sm ${col.visible ? 'text-gray-800' : 'text-gray-500'}`}>{col.label}</span>
                     </label>
                     <div className="flex gap-1">
                         <button 
                            onClick={() => moveColumn(idx, 'up')}
                            disabled={idx === 0}
                            className="p-1 text-gray-400 hover:text-blue-600 hover:bg-blue-50 rounded disabled:opacity-30 disabled:hover:bg-transparent"
                         >
                            <svg className="w-4 h-4" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path strokeLinecap="round" strokeLinejoin="round" strokeWidth="2" d="M5 15l7-7 7 7"/></svg>
                         </button>
                         <button 
                            onClick={() => moveColumn(idx, 'down')}
                            disabled={idx === columns.length - 1}
                            className="p-1 text-gray-400 hover:text-blue-600 hover:bg-blue-50 rounded disabled:opacity-30 disabled:hover:bg-transparent"
                         >
                            <svg className="w-4 h-4" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path strokeLinecap="round" strokeLinejoin="round" strokeWidth="2" d="M19 9l-7 7-7-7"/></svg>
                         </button>
                     </div>
                 </div>
             ))}
         </div>
      </Modal>

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
            {exportType === 'pdf' && (
                <div className="p-3 bg-blue-50 border border-blue-100 rounded-lg space-y-3">
                    <h4 className="text-xs font-bold text-blue-800 uppercase">Customização PDF</h4>
                    
                    {/* Title & Logo Row */}
                    <div className="flex gap-3">
                        <div className="flex-grow">
                             <label className="block text-[10px] font-bold text-gray-600 uppercase mb-1">Título do Relatório</label>
                             <input 
                               type="text" 
                               value={pdfTitle} 
                               onChange={e => setPdfTitle(e.target.value)} 
                               className="w-full text-sm p-1.5 border rounded focus:ring-1 focus:ring-blue-400 outline-none"
                             />
                        </div>
                        <div className="flex-shrink-0">
                            <label className="block text-[10px] font-bold text-gray-600 uppercase mb-1">Logo (Direita)</label>
                            <div className="relative overflow-hidden w-24">
                                <button className="w-full text-[10px] py-1.5 bg-white border border-gray-300 rounded hover:bg-gray-50 text-gray-700">Escolher Imagem</button>
                                <input 
                                   type="file" 
                                   accept="image/*" 
                                   onChange={handleLogoSelect} 
                                   className="absolute inset-0 opacity-0 cursor-pointer"
                                />
                            </div>
                            {pdfLogo && <div className="text-[9px] text-green-600 mt-0.5 text-center">Imagem carregada</div>}
                        </div>
                    </div>

                    {/* Observations */}
                    <div>
                        <label className="block text-[10px] font-bold text-gray-600 uppercase mb-1">Observações (Opcional)</label>
                        <textarea 
                           value={pdfObservation}
                           onChange={e => setPdfObservation(e.target.value)}
                           className="w-full text-xs p-1.5 border rounded focus:ring-1 focus:ring-blue-400 outline-none h-16 resize-none"
                           placeholder="Digite observações adicionais para o cabeçalho..."
                        />
                    </div>
                </div>
            )}

             <div className="flex gap-2">
                 <button onClick={() => setExportCols(columns.map(c => c.label))} className="text-xs px-3 py-1 bg-gray-100 hover:bg-gray-200 rounded border">Marcar Todos</button>
                 <button onClick={() => setExportCols([])} className="text-xs px-3 py-1 bg-gray-100 hover:bg-gray-200 rounded border">Desmarcar Todos</button>
             </div>
             <div className="grid grid-cols-2 gap-2 max-h-60 overflow-y-auto custom-scrollbar">
                 {columns.map(col => (
                     <label key={col.id} className="flex items-center gap-2 p-2 border rounded hover:bg-gray-50 cursor-pointer">
                         <input 
                            type="checkbox"
                            checked={exportCols.includes(col.label)}
                            onChange={(e) => {
                                if(e.target.checked) setExportCols([...exportCols, col.label]);
                                else setExportCols(exportCols.filter(c => c !== col.label));
                            }}
                            className="rounded border-gray-300 text-green-600 w-4 h-4"
                         />
                         <span className="text-xs text-gray-700">{col.label}</span>
                     </label>
                 ))}
             </div>
         </div>
      </Modal>
    </div>
  );
};

export default App;