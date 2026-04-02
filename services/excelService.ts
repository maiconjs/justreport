import * as XLSX from 'xlsx';
import { ReportItem, MapColumnConfig, MapInfo } from '../types';
import { parseDateRobust, formatDate, excelDateToJSDate } from '../utils/dateUtils';

const columnMap = {
  dataCriacao: ['Data de Criação', 'Data Criação', 'Data Abertura'],
  dataConclusao: ['Data de Conclusão', 'Data Conclusão'],
  os: ['OS', 'Ordem de Serviço'],
  idOsCorp: ['ID OS Corporate', 'ID OS', 'ID Corporate'],
  tipo: ['Tipo', 'Tipo OS'],
  statusOs: ['Status da OS', 'Status'],
  contrato: ['Contrato/Item', 'Contrato', 'Item'],
  serie: ['Número de Série', 'Nº de Série', 'Serial'],
  situacaoEquip: ['Situação do Equipamento', 'Situação'],
  equipProduzindo: ['Equipamento Produzindo', 'Produzindo'],
  tipoConexao: ['Tipo de Conexão', 'Conexão'],
  ip: ['Endereço IP', 'IP'],
  hostname: ['Hostname', 'Nome do Host'],
  bairro: ['Bairro de Atendimento', 'Bairro'],
  cidade: ['Cidade', 'City'],
  filial: ['Filial', 'Unidade']
};

export const mapColumnsConfig: MapColumnConfig[] = [
  { key: 'statusGeral', label: 'Status Geral (Mapa)', search: ['Status Geral', 'Status Gera'], strict: false },
  { key: 'statusItem', label: 'Status Item (Mapa)', search: ['Status Item'], strict: false },
  { key: 'modelo', label: 'Modelo (Mapa)', search: ['Modelo'], strict: false },
  { key: 'bairro', label: 'Bairro (Mapa)', search: ['Bairro'], strict: false },
  { key: 'cidade', label: 'Cidade (Mapa)', search: ['Cidade'], strict: false },
  { key: 'uf', label: 'UF (Mapa)', search: ['UF', 'Estado'], strict: true },
  { key: 'cep', label: 'CEP (Mapa)', search: ['CEP'], strict: true },
  { key: 'cnpj', label: 'CNPJ (Mapa)', search: ['CNPJ INSTALAÇÃO', 'CNPJ'], strict: false },
  { key: 'dtInstalacao', label: 'Dt. Instalação (Mapa)', search: ['Dt. Instalação', 'Data Instalação'], strict: false },
  { key: 'obs', label: 'OBS (Mapa)', search: ['OBS / comentário', 'OBS / comentários', 'OBS/comentários', 'OBS', 'Observações', 'Comentários'], strict: true },
  { key: 'ip', label: 'IP (Mapa)', search: ['Endereço do Sistema', 'Endereço Sistema', 'IP'], strict: true },
  { key: 'mascara', label: 'Máscara (Mapa)', search: ['Máscara', 'Mascara'], strict: false },
  { key: 'gateway', label: 'Gateway (Mapa)', search: ['Gateway'], strict: false },
  { key: 'dns', label: 'DNS (Mapa)', search: ['DNS'], strict: true }
];

export const readExcelFile = async (file: File): Promise<any[]> => {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const data = new Uint8Array(e.target?.result as ArrayBuffer);
        const workbook = XLSX.read(data, { type: 'array', cellDates: true });
        const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
        const jsonData = XLSX.utils.sheet_to_json(firstSheet, { defval: "" });
        resolve(jsonData);
      } catch (error) {
        reject(error);
      }
    };
    reader.onerror = reject;
    reader.readAsArrayBuffer(file);
  });
};

export const readNddCsv = async (file: File): Promise<any[]> => {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const arrayBuffer = e.target?.result as ArrayBuffer;
        const data = new Uint8Array(arrayBuffer);
        
        // Use XLSX to read the CSV, it's generally better at detecting delimiters and encodings
        const workbook = XLSX.read(data, { type: 'array', cellDates: true });
        const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
        let jsonData = XLSX.utils.sheet_to_json(firstSheet, { defval: "" });

        // Clean up keys in jsonData if they have BOM or quotes
        if (jsonData.length > 0) {
            jsonData = jsonData.map((row: any) => {
                const newRow: any = {};
                for (const key in row) {
                    const cleanKey = key.trim().replace(/^\uFEFF/, '').replace(/^"|"$/g, '').trim();
                    newRow[cleanKey] = row[key];
                }
                return newRow;
            });
        }

        // Check if it parsed correctly (if it only has one column and it contains semicolons, it failed to detect delimiter)
        if (jsonData.length > 0) {
           const firstRow = jsonData[0] as any;
           const keys = Object.keys(firstRow);
           if (keys.length === 1 && String(firstRow[keys[0]]).includes(';')) {
              // Fallback to manual semicolon parsing if XLSX failed
              const decoder = new TextDecoder('iso-8859-1');
              const text = decoder.decode(data);
              const lines = text.split(/\r?\n/);
              const headers = lines[0].split(';').map(h => h.trim().replace(/^"|"$/g, ''));
              const result = [];
              for (let i = 1; i < lines.length; i++) {
                if (!lines[i].trim()) continue;
                const currentLine = lines[i].split(';').map(v => v.trim().replace(/^"|"$/g, ''));
                const obj: any = {};
                for (let j = 0; j < headers.length; j++) {
                  obj[headers[j]] = currentLine[j] || "";
                }
                result.push(obj);
              }
              return resolve(result);
           }
        }
        
        resolve(jsonData);
      } catch (error) {
        reject(error);
      }
    };
    reader.onerror = reject;
    reader.readAsArrayBuffer(file);
  });
};

export const processReportData = (rawData: any[], fileName: string): ReportItem[] => {
  const getRaw = (row: any, keys: string[]) => {
    for (const key of keys) {
      if (row[key] !== undefined) return row[key];
      const upperKey = Object.keys(row).find(k => k.toUpperCase() === key.toUpperCase());
      if (upperKey) return row[upperKey];
    }
    return null;
  };

  return rawData.map((row, index) => {
    const rawCriacao = getRaw(row, columnMap.dataCriacao);
    const rawConclusao = getRaw(row, columnMap.dataConclusao);
    let rawContrato = getRaw(row, columnMap.contrato);
    
    if (rawContrato && String(rawContrato).includes('/')) {
      rawContrato = String(rawContrato).split('/')[0].trim();
    }

    const rawOs = getRaw(row, columnMap.os);
    const rawSerie = getRaw(row, columnMap.serie);

    // Generate a synthetic ID if needed
    const syntheticId = `${fileName}-${index}`;

    return {
      id: syntheticId,
      dataCriacao: formatDate(parseDateRobust(rawCriacao) || rawCriacao),
      dataConclusao: formatDate(parseDateRobust(rawConclusao) || rawConclusao),
      os: formatDate(rawOs),
      idOsCorp: formatDate(getRaw(row, columnMap.idOsCorp)),
      tipo: formatDate(getRaw(row, columnMap.tipo)),
      statusOs: formatDate(getRaw(row, columnMap.statusOs)),
      contrato: formatDate(rawContrato),
      serie: formatDate(rawSerie),
      situacaoEquip: formatDate(getRaw(row, columnMap.situacaoEquip)),
      equipProduzindo: formatDate(getRaw(row, columnMap.equipProduzindo)),
      tipoConexao: formatDate(getRaw(row, columnMap.tipoConexao)),
      ip: formatDate(getRaw(row, columnMap.ip)),
      hostname: formatDate(getRaw(row, columnMap.hostname)),
      bairro: formatDate(getRaw(row, columnMap.bairro)),
      cidade: formatDate(getRaw(row, columnMap.cidade)),
      filial: formatDate(getRaw(row, columnMap.filial)),
      origem: fileName,
      _rawCriacao: parseDateRobust(rawCriacao),
      _rawConclusao: parseDateRobust(rawConclusao)
    };
  });
};

export const parseMapWorkbook = async (file: File): Promise<XLSX.WorkBook> => {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const data = new Uint8Array(e.target?.result as ArrayBuffer);
        const wb = XLSX.read(data, { type: 'array', bookDeps: true });
        resolve(wb);
      } catch (err) {
        reject(err);
      }
    };
    reader.readAsArrayBuffer(file);
  });
};

export const processMapSheet = (ws: XLSX.WorkSheet, manualHeaderRow: number | null): Map<string, MapInfo> => {
  const mapData = new Map<string, MapInfo>();
  
  if (!ws['!ref']) return mapData;
  const range = XLSX.utils.decode_range(ws['!ref']);
  
  let headerRowIndex = -1;
  
  if (manualHeaderRow && manualHeaderRow > 0) {
    headerRowIndex = manualHeaderRow - 1;
  } else {
    const mustHave = ['MODELO', 'STATUS', 'CONTRATO', 'SÉRIE', 'SERIAL', 'ENDEREÇO', 'BAIRRO'];
    const scanLimit = Math.min(range.e.r, 100);
    for (let r = range.s.r; r <= scanLimit; r++) {
      let matches = 0;
      for (let c = range.s.c; c <= range.e.c; c++) {
        const cell = ws[XLSX.utils.encode_cell({ r, c })];
        if (cell && cell.v) {
          const val = String(cell.v).toUpperCase();
          if (mustHave.some(k => val.includes(k))) matches++;
        }
      }
      if (matches >= 2) {
        headerRowIndex = r;
        break;
      }
    }
  }

  if (headerRowIndex === -1) throw new Error("Cabeçalho não encontrado");

  const colIndices: { [key: string]: number } = {};
  let serialIdx = -1;

  // Map columns logic
  for (let c = range.s.c; c <= range.e.c; c++) {
    const cell = ws[XLSX.utils.encode_cell({ r: headerRowIndex, c })];
    if (cell && cell.v) {
      const val = String(cell.v).trim().toUpperCase();
      
      // Identify Serial Column
      const serialSearch = ['SÉRIE', 'SERIE', 'SERIAL', 'Nº SÉRIE', 'NUMERO DE SERIE', 'CONTRATO/ITEM'];
      if (serialSearch.some(s => val.includes(s))) serialIdx = c;

      mapColumnsConfig.forEach(mc => {
        for (const term of mc.search) {
          const upperTerm = term.toUpperCase();
          const exactMatch = val === upperTerm;
          const startsWithMatch = val.startsWith(upperTerm);
          const includesMatch = val.includes(upperTerm);

          if (exactMatch) {
            colIndices[mc.key] = c;
            break;
          } else if (!mc.strict && startsWithMatch) {
            if (colIndices[mc.key] === undefined) colIndices[mc.key] = c;
          } else if (!mc.strict && includesMatch) {
            if (colIndices[mc.key] === undefined) colIndices[mc.key] = c;
          }
        }
      });
    }
  }

  // Fallback for serial
  if (serialIdx === -1) {
     // Try generic ID
     for (let c = range.s.c; c <= range.e.c; c++) {
        const cell = ws[XLSX.utils.encode_cell({ r: headerRowIndex, c })];
        if(cell && cell.v && String(cell.v).toUpperCase() === 'ID') {
            serialIdx = c;
            break;
        }
     }
  }

  if (serialIdx === -1) throw new Error(`Coluna chave (Série/Serial/ID) não encontrada na linha ${headerRowIndex + 1}`);

  const startRow = headerRowIndex + 1;
  for (let r = startRow; r <= range.e.r; r++) {
    const cellSerial = ws[XLSX.utils.encode_cell({ r, c: serialIdx })];
    if (cellSerial && cellSerial.v) {
      const key = String(cellSerial.v).trim().toUpperCase();
      const entry: any = {};
      mapColumnsConfig.forEach(mc => {
        const idx = colIndices[mc.key];
        if (idx !== undefined) {
          const cellVal = ws[XLSX.utils.encode_cell({ r, c: idx })];
          let val = (cellVal && cellVal.v !== undefined) ? cellVal.v : "N/A";
          if (mc.key === 'dtInstalacao') {
             val = excelDateToJSDate(val);
          } else {
             val = String(val).trim();
          }
          entry[mc.key] = val;
        } else {
          entry[mc.key] = "N/A";
        }
      });
      mapData.set(key, entry as MapInfo);
    }
  }
  return mapData;
};