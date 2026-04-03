export interface ViaCepResult {
  cep: string;        // "41295-360"
  logradouro: string;
  complemento: string;
  bairro: string;
  localidade: string; // cidade
  uf: string;
}

// ---------------------------------------------------------------------------
// 1. In-memory cache (instant, survives only the current session)
// ---------------------------------------------------------------------------
const memoryCache = new Map<string, { result: ViaCepResult | null; ts: number }>();

// ---------------------------------------------------------------------------
// 2. IndexedDB persistent cache (survives page reloads, 30-day TTL)
// ---------------------------------------------------------------------------
const DB_NAME = 'viacep-cache';
const STORE_NAME = 'ceps';
const DB_VERSION = 1;
const TTL_MS = 30 * 24 * 60 * 60 * 1000; // 30 days

const openDB = (): Promise<IDBDatabase> =>
  new Promise((resolve, reject) => {
    const req = indexedDB.open(DB_NAME, DB_VERSION);
    req.onupgradeneeded = () => {
      const db = req.result;
      if (!db.objectStoreNames.contains(STORE_NAME)) {
        db.createObjectStore(STORE_NAME, { keyPath: 'cep' });
      }
    };
    req.onsuccess = () => resolve(req.result);
    req.onerror = () => reject(req.error);
  });

interface CacheEntry {
  cep: string;
  result: ViaCepResult | null;
  ts: number;
}

/** Bulk-read from IndexedDB in a single transaction (very fast). */
const bulkGetFromIDB = async (ceps: string[]): Promise<Map<string, CacheEntry>> => {
  const out = new Map<string, CacheEntry>();
  if (ceps.length === 0) return out;
  try {
    const db = await openDB();
    return new Promise((resolve) => {
      const tx = db.transaction(STORE_NAME, 'readonly');
      const store = tx.objectStore(STORE_NAME);
      let pending = ceps.length;
      for (const cep of ceps) {
        const req = store.get(cep);
        req.onsuccess = () => {
          if (req.result) out.set(cep, req.result as CacheEntry);
          if (--pending === 0) resolve(out);
        };
        req.onerror = () => {
          if (--pending === 0) resolve(out);
        };
      }
      if (pending === 0) resolve(out);
    });
  } catch {
    return out;
  }
};

/** Bulk-write to IndexedDB in a single transaction (fire-and-forget). */
const bulkSaveToIDB = async (entries: CacheEntry[]): Promise<void> => {
  if (entries.length === 0) return;
  try {
    const db = await openDB();
    const tx = db.transaction(STORE_NAME, 'readwrite');
    const store = tx.objectStore(STORE_NAME);
    for (const entry of entries) {
      store.put(entry);
    }
  } catch {
    // IndexedDB unavailable — silently ignore
  }
};

// ---------------------------------------------------------------------------
// 3. Multi-provider CEP lookup (race for speed)
// ---------------------------------------------------------------------------

/**
 * Normalizes BrasilAPI v2 response to the ViaCepResult format.
 */
function normalizeBrasilApi(data: any, cleanCep: string): ViaCepResult | null {
  if (!data || data.errors) return null;
  return {
    cep: data.cep || cleanCep,
    logradouro: data.street || '',
    complemento: '',
    bairro: data.neighborhood || '',
    localidade: data.city || '',
    uf: data.state || '',
  };
}

/**
 * Normalizes OpenCEP / ViaCEP response (same format) to ViaCepResult.
 */
function normalizeViaCep(data: any, cleanCep: string): ViaCepResult | null {
  if (!data || data.erro) return null;
  return {
    cep: data.cep || cleanCep,
    logradouro: data.logradouro || '',
    complemento: data.complemento || '',
    bairro: data.bairro || '',
    localidade: data.localidade || '',
    uf: data.uf || '',
  };
}

/**
 * Fetches a single CEP from a specific provider with timeout.
 * Returns null on any error (timeout, network, parse).
 */
async function fetchFromProvider(
  url: string,
  normalize: (data: any, cep: string) => ViaCepResult | null,
  cep: string,
  timeoutMs: number,
  parentSignal?: AbortSignal
): Promise<ViaCepResult | null> {
  const controller = new AbortController();
  const timer = setTimeout(() => controller.abort(), timeoutMs);

  // Link parent abort to this child
  const onParentAbort = () => controller.abort();
  parentSignal?.addEventListener('abort', onParentAbort, { once: true });

  try {
    const res = await fetch(url, { signal: controller.signal });
    if (res.status === 429) throw Object.assign(new Error('rate-limit'), { status: 429 });
    if (!res.ok) return null;
    const data = await res.json();
    return normalize(data, cep);
  } catch (err: any) {
    if (err?.status === 429) throw err;
    if (parentSignal?.aborted) throw new DOMException('Aborted', 'AbortError');
    return null;
  } finally {
    clearTimeout(timer);
    parentSignal?.removeEventListener('abort', onParentAbort);
  }
}

/**
 * Races multiple CEP providers for a single CEP.
 * Returns the first successful non-null result.
 * 
 * Provider priority (all launched simultaneously, first wins):
 *   1. OpenCEP    — CloudFlare CDN, fastest, same format as ViaCEP
 *   2. BrasilAPI  — fallback aggregator, high reliability
 *   3. ViaCEP     — original provider, can be slow
 */
const PER_REQUEST_TIMEOUT = 4000; // 4s max per individual provider

export const lookupCep = async (
  cep: string,
  signal?: AbortSignal
): Promise<ViaCepResult | null> => {
  const clean = cep.replace(/\D/g, '');
  if (clean.length !== 8) return null;

  // Race all providers — first successful response wins
  const providers = [
    fetchFromProvider(
      `https://opencep.com/v1/${clean}`,
      normalizeViaCep, clean, PER_REQUEST_TIMEOUT, signal
    ),
    fetchFromProvider(
      `https://brasilapi.com.br/api/cep/v2/${clean}`,
      normalizeBrasilApi, clean, PER_REQUEST_TIMEOUT, signal
    ),
    fetchFromProvider(
      `https://viacep.com.br/ws/${clean}/json/`,
      normalizeViaCep, clean, PER_REQUEST_TIMEOUT, signal
    ),
  ];

  // Promise.any() resolves with the first fulfilled (non-rejected) promise.
  // We wrap nulls as rejections so only actual data wins the race.
  try {
    const result = await Promise.any(
      providers.map(p =>
        p.then(r => {
          if (r === null) throw new Error('no-data');
          return r;
        })
      )
    );
    return result;
  } catch {
    // All providers failed or returned null
    return null;
  }
};

// ---------------------------------------------------------------------------
// 4. Bulk lookup with multi-layer cache + multi-provider race
// ---------------------------------------------------------------------------

/**
 * Bulk-validates a list of CEPs using a three-tier caching strategy:
 *
 *  1. In-memory cache (instant, current session)
 *  2. IndexedDB persistent cache (fast, survives reloads, 30-day TTL)
 *  3. Multi-provider network race (OpenCEP + BrasilAPI + ViaCEP)
 *
 * Key design choices:
 *  - Deduplication: only unique 8-digit CEPs are queried.
 *  - Cache-first: most CEPs are served from cache after the first load.
 *  - Continuous pool: as soon as one slot finishes, the next CEP starts.
 *  - Multi-provider race: 3 APIs raced simultaneously per CEP.
 *  - Concurrency 30: balanced for browser connection limits (6 per host × 3 hosts ≈ 18 actual).
 */
export const bulkLookupCeps = async (
  ceps: string[],
  onProgress: (done: number, total: number) => void,
  signal?: AbortSignal
): Promise<Map<string, ViaCepResult | null>> => {
  const unique = [...new Set(
    ceps.map(c => c.replace(/\D/g, '')).filter(c => c.length === 8)
  )];

  const results = new Map<string, ViaCepResult | null>();
  const total = unique.length;
  let done = 0;

  if (total === 0) return results;

  const now = Date.now();

  // --- Layer 1: In-memory cache ---
  const needIDB: string[] = [];
  for (const cep of unique) {
    const cached = memoryCache.get(cep);
    if (cached && (now - cached.ts) < TTL_MS) {
      results.set(cep, cached.result);
      done++;
    } else {
      needIDB.push(cep);
    }
  }

  // --- Layer 2: IndexedDB cache ---
  const needNetwork: string[] = [];
  if (needIDB.length > 0) {
    const idbEntries = await bulkGetFromIDB(needIDB);
    for (const cep of needIDB) {
      const entry = idbEntries.get(cep);
      if (entry && (now - entry.ts) < TTL_MS) {
        results.set(cep, entry.result);
        memoryCache.set(cep, { result: entry.result, ts: entry.ts });
        done++;
      } else {
        needNetwork.push(cep);
      }
    }
  }

  // Report progress after cache resolution
  onProgress(done, total);

  // If everything was cached, we're done!
  if (needNetwork.length === 0) return results;

  // --- Layer 3: Network (multi-provider race, only for cache misses) ---
  // Browser limits: ~6 connections per hostname. With 3 hosts and concurrency 30,
  // we get ~10 CEPs active × 3 providers = 30 connections spread across 3 hosts.
  // This keeps all connection pools busy without overwhelming any single host.
  const CONCURRENCY = 30;
  const BACKOFF_MS  = 500;

  let idx = 0;
  const newEntries: CacheEntry[] = [];

  const sleep = (ms: number) => new Promise<void>(r => setTimeout(r, ms));

  const worker = async () => {
    while (true) {
      if (signal?.aborted) break;
      const i = idx++;
      if (i >= needNetwork.length) break;

      const cep = needNetwork[i];
      let result: ViaCepResult | null = null;

      for (let attempt = 0; attempt < 2; attempt++) {
        try {
          result = await lookupCep(cep, signal);
          break;
        } catch (err: any) {
          if (err?.name === 'AbortError') return;
          if (err?.status === 429 && attempt < 1) {
            await sleep(BACKOFF_MS * (attempt + 1));
            continue;
          }
          break;
        }
      }

      const ts = Date.now();
      memoryCache.set(cep, { result, ts });
      newEntries.push({ cep, result, ts });

      results.set(cep, result);
      done++;
      onProgress(done, total);
    }
  };

  await Promise.all(
    Array.from({ length: Math.min(CONCURRENCY, needNetwork.length) }, () => worker())
  );

  // Persist to IndexedDB (non-blocking)
  bulkSaveToIDB(newEntries).catch(() => {});

  return results;
};

/**
 * Returns cache statistics for UI display.
 */
export const getCacheStats = (): { memorySize: number } => ({
  memorySize: memoryCache.size,
});

/**
 * Clears all cached CEP data (both memory and IndexedDB).
 */
export const clearCepCache = async (): Promise<void> => {
  memoryCache.clear();
  try {
    const db = await openDB();
    const tx = db.transaction(STORE_NAME, 'readwrite');
    tx.objectStore(STORE_NAME).clear();
  } catch {
    // IndexedDB unavailable
  }
};
