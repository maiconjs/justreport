import { HTMLAttributes } from 'react';

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
}

export interface MapColumnConfig {
  key: keyof MapInfo;
  label: string;
  search: string[];
  strict: boolean;
}

// Add support for webkitdirectory to InputHTMLAttributes
declare module 'react' {
  interface InputHTMLAttributes<T> extends HTMLAttributes<T> {
    webkitdirectory?: string;
    directory?: string;
  }
}