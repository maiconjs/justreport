export const parseDateRobust = (value: any): Date | null => {
  if (!value) return null;
  if (value instanceof Date && !isNaN(value.getTime())) return value;

  if (typeof value === 'number') {
    // Excel serial date
    const utc_days = Math.floor(value - 25569);
    const d = new Date(utc_days * 86400 * 1000);
    if (!isNaN(d.getTime())) return d;
  }

  if (typeof value === 'string') {
    const trimmed = value.trim();
    // Match one or two digit segments separated by / . -  (handles D/M/YYYY, DD/MM/YYYY, M/D/YYYY, MM/DD/YYYY)
    const parts = trimmed.match(/^(\d{1,2})[\/.-](\d{1,2})[\/.-](\d{4})/);
    if (parts) {
      const first  = parseInt(parts[1]);
      const second = parseInt(parts[2]);
      const year   = parseInt(parts[3]);

      // If first part > 12, it cannot be a month → it's DD/MM/YYYY (BR)
      if (first > 12) return new Date(year, second - 1, first);
      // If second part > 12, it cannot be a month → it's MM/DD/YYYY (US)
      if (second > 12) return new Date(year, first - 1, second);
      // Ambiguous (both <= 12): assume BR format DD/MM/YYYY
      return new Date(year, second - 1, first);
    }
    // Try ISO 8601 (YYYY-MM-DD or full datetime)
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