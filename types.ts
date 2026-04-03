export interface ReportItem {
  id: string; // Unique ID (often based on OS or generated)
  dataCriacao: string;
  dataConclusao: string;
  os: string;
  idOsCorp: string;
  tipo: string;
  statusOs: string;
  contrato: string;
  serie: string;
  situacaoEquip: string;
  equipProduzindo: string;
  tipoConexao: string;
  ip: string;
  hostname: string;
  bairro: string;
  cidade: string;
  filial: string;
  origem: string;
  _rawCriacao: Date | null;
  _rawConclusao: Date | null;
  [key: string]: any;
}

export interface SdsInfo {
  status: 'Monitorado' | 'Alerta' | 'Não Monitorado' | 'Dados Incompletos' | '-';
  colorClass: string;
  lastUpdate: string;
  detection: string;
  rawLastUpdate: Date | null;
  rawDetection: Date | null;
}

export interface NddInfo {
  status: string;
  colorClass: string;
  lastUpdate: string;
  daysWithoutMeters: string;
  rawLastUpdate: Date | null;
  accountingStatus: string;
  connectionType: string;
  mpsIp: string;
}

export interface MapInfo {
  statusGeral: string;
  statusItem: string;
  modelo: string;
  bairro: string;
  cidade: string;
  uf: string;
  cep: string;
  cnpj: string;
  dtInstalacao: string;
  obs: string;
  ip: string;
  mascara: string;
  gateway: string;
  dns: string;
  [key: string]: string;
}

export interface FilterState {
  search: string;
  searchField: string;
  startCreation: string;
  endCreation: string;
  startConclusion: string;
  endConclusion: string;
  alertDays: number;
  offlineDays: number;
  selectedTypes: string[];
  selectedProds: string[];
  selectedStatus: string[];
  selectedSituacao: string[];
  selectedConexao: string[];
  selectedMon: string[];
  selectedNddMon: string[];
}

export interface MapColumnConfig {
  key: keyof MapInfo;
  label: string;
  search: string[];
  strict: boolean;
}

export type CepStatus = 'unchecked' | 'valid' | 'invalid';

export interface CorporateInfo {
  serial: string;
  status: string;         // 'Ativo' | 'Inativo'
  modelo: string;
  enderecoInstalacao: string;  // raw original
  // Parsed/validated address fields
  logradouro: string;
  complemento: string;
  bairro: string;
  cidade: string;
  uf: string;
  cep: string;
  cepStatus: CepStatus;
  dataInstalacao: string;
  clienteInstalacao: string;
  inContract: boolean;
}

export interface CepInvalidEntry {
  serial: string;
  cep: string;
  enderecoRaw: string;
  cidade: string;
  uf: string;
  modelo: string;
}

export interface ColumnDef {
  id: string;
  label: string;
  visible: boolean;
  type: 'sds' | 'map' | 'standard' | 'ndd' | 'corporate';
  key: string; // Key in the respective data object
}