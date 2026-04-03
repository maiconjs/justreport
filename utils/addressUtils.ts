export interface ParsedAddress {
  logradouro: string;
  complemento: string;
  bairro: string;
  cidade: string;
  uf: string;
  cep: string;
}

/**
 * Returns true if the character is a Unicode uppercase letter (including accented, like Á É Ã).
 */
const isUpper = (c: string) => /\p{Lu}/u.test(c);
const isLower = (c: string) => /\p{Ll}/u.test(c);
const isAlpha = (c: string) => /\p{L}/u.test(c);

/**
 * Splits a concatenated "BairroCidade" string into its two parts.
 *
 * The corporate XLS concatenates them without a separator, e.g.:
 *   "Ponto CertoCamaçari"       → bairro="Ponto Certo",    cidade="Camaçari"
 *   "PirajáSalvador"            → bairro="Pirajá",         cidade="Salvador"
 *   "Dt IndustrialCandeias"     → bairro="Dt Industrial",  cidade="Candeias"
 *   "CENTROJequié"              → bairro="CENTRO",         cidade="Jequié"
 *   "TANQUEFeira de Santana"    → bairro="TANQUE",         cidade="Feira de Santana"
 *   "SSA PIRAJÁPirajá"          → bairro="SSA PIRAJÁ",     cidade="Pirajá"
 *   "B S JOÃOItaberaba"         → bairro="B S JOÃO",       cidade="Itaberaba"
 *   "Alphaville ISalvador"      → bairro="Alphaville I",   cidade="Salvador"
 *   "CENTROMata de São João"    → bairro="CENTRO",         cidade="Mata de São João"
 *
 * Algorithm:
 *  Scan left-to-right collecting all candidate split points (positions where a
 *  boundary between bairro and cidade could be). A boundary exists at position i
 *  when the transition at i looks like the START of a new word that could be a city:
 *    - lowercase letter immediately followed by an uppercase letter  (PirajáSalvador)
 *    - digit immediately followed by an uppercase letter             (rare)
 *    - uppercase run ending, followed by an uppercase letter whose NEXT char is lowercase
 *      (handles ALL-CAPS bairro + Title-case city:  CENTROJequié, TANQUEFeira)
 *
 *  We take the LAST qualifying boundary that leaves a non-empty bairro.
 */
export function splitBairroCidade(bc: string): { bairro: string; cidade: string } {
  if (!bc) return { bairro: '', cidade: '' };

  const candidates: number[] = [];

  for (let i = 1; i < bc.length; i++) {
    const prev = bc[i - 1];
    const cur  = bc[i];
    const next = i + 1 < bc.length ? bc[i + 1] : '';

    // Rule 1: lowercase/digit → uppercase  (most common: "PirajáSalvador", "CertoCamaçari")
    if ((isLower(prev) || /\d/.test(prev)) && isUpper(cur)) {
      candidates.push(i);
      continue;
    }

    // Rule 2: uppercase run → Title-case word start
    // cur is uppercase AND next is lowercase AND prev is also uppercase (we're mid-caps-run)
    // This catches "CENTROJequié" (O→J, e is next) and "TANQUEFeira" (E→F, e is next)
    if (isUpper(prev) && isUpper(cur) && next && isLower(next)) {
      candidates.push(i);
      continue;
    }

    // Rule 3: space + uppercase after uppercase run  e.g. "CENTRO Mata" — space handled naturally
    // (spaces between words don't create split; only the word-start transition matters)
  }

  // Take the LAST candidate that leaves a meaningful bairro (at least 2 chars)
  for (let k = candidates.length - 1; k >= 0; k--) {
    const idx = candidates[k];
    const bairro = bc.slice(0, idx).trim();
    const cidade = bc.slice(idx).trim();
    if (bairro.length >= 2 && cidade.length >= 2 && isAlpha(cidade[0])) {
      return { bairro, cidade };
    }
  }

  // No valid split found
  return { bairro: bc, cidade: '' };
}

/**
 * Parses the concatenated address format used in the Corporate XLS.
 *
 * Format: "<street>, <number> [- <complement>...] - <BairroCity> - <UF> - <CEP>"
 */
export function parseEnderecoInstalacao(raw: string): ParsedAddress {
  const empty: ParsedAddress = { logradouro: raw || '', complemento: '', bairro: '', cidade: '', uf: '', cep: '' };
  if (!raw || raw === '-') return empty;

  const parts = raw.split(' - ');
  if (parts.length < 3) return empty;

  // Tail: CEP — last 8 consecutive digits
  const cepRaw = parts[parts.length - 1].trim();
  const cep = cepRaw.replace(/\D/g, '');
  if (cep.length !== 8) return empty;

  // Second-to-last: UF — exactly 2 uppercase ASCII letters
  const uf = parts[parts.length - 2].trim();
  if (!/^[A-Z]{2}$/.test(uf)) return empty;

  // Third-to-last: BairroCidade concatenated
  const bairroCidade = parts[parts.length - 3].trim();

  // Everything before BairroCidade
  const streetParts = parts.slice(0, parts.length - 3);

  const { bairro, cidade } = splitBairroCidade(bairroCidade);

  const logradouro = (streetParts[0] ?? '').trim();
  const complemento = streetParts.slice(1).join(' - ').trim();

  return { logradouro, complemento, bairro, cidade, uf, cep };
}
