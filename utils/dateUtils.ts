export const parseDateRobust = (value: any): Date | null => {
  if (!value) return null;
  if (value instanceof Date && !isNaN(value.getTime())) return value;

  if (typeof value === 'string') {
    const trimmed = value.trim();
    // Match DD/MM/YYYY or DD-MM-YYYY
    const ptBrMatch = trimmed.match(/^(\d{1,2})[\/.-](\d{1,2})[\/.-](\d{4})/);
    if (ptBrMatch) {
      return new Date(parseInt(ptBrMatch[3]), parseInt(ptBrMatch[2]) - 1, parseInt(ptBrMatch[1]));
    }
    // Try ISO
    const isoDate = new Date(trimmed);
    if (!isNaN(isoDate.getTime())) return isoDate;
  }
  return null;
};

export const formatDate = (val: any): string => {
  if (val instanceof Date && !isNaN(val.getTime())) return val.toLocaleDateString('pt-BR');
  return (val !== null && val !== undefined && val !== '') ? String(val).trim() : "-";
};

export const excelDateToJSDate = (serial: any): string => {
  if (serial == null) return "N/A";
  if (typeof serial === 'string' && serial.includes('/')) return serial.trim();
  if (typeof serial === 'number') {
    const utc_days = Math.floor(serial - 25569);
    const utc_value = utc_days * 86400;
    const date_info = new Date(utc_value * 1000);
    const day = String(date_info.getUTCDate()).padStart(2, '0');
    const month = String(date_info.getUTCMonth() + 1).padStart(2, '0');
    const year = date_info.getUTCFullYear();
    return `${day}/${month}/${year}`;
  }
  return String(serial).trim();
};