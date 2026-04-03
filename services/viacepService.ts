export interface ViaCepResult {
  cep: string;        // "41295-360"
  logradouro: string;
  complemento: string;
  bairro: string;
  localidade: string; // cidade
  uf: string;
}

export const lookupCep = async (
  cep: string,
  signal?: AbortSignal
): Promise<ViaCepResult | null> => {
  const clean = cep.replace(/\D/g, '');
  if (clean.length !== 8) return null;
  try {
    const res = await fetch(`https://viacep.com.br/ws/${clean}/json/`, { signal });
    // 429 = rate limited — caller should retry with backoff
    if (res.status === 429) throw Object.assign(new Error('rate-limit'), { status: 429 });
    if (!res.ok) return null;
    const data = await res.json();
    if (data.erro) return null;
    return {
      cep: data.cep || clean,
      logradouro: data.logradouro || '',
      complemento: data.complemento || '',
      bairro: data.bairro || '',
      localidade: data.localidade || '',
      uf: data.uf || '',
    };
  } catch (err: any) {
    if (err?.name === 'AbortError') throw err; // propagate cancellation
    if (err?.status === 429) throw err;        // propagate rate-limit for retry
    return null;
  }
};

/**
 * Bulk-validates a list of CEPs using a continuous concurrency pool.
 *
 * Key design choices:
 *  - Deduplication: only unique 8-digit CEPs are queried (huge saving for repeated serials).
 *  - Continuous pool (not batch): as soon as one slot finishes, the next CEP starts.
 *    This keeps CONCURRENCY requests in-flight at all times instead of waiting for
 *    the slowest in each batch.
 *  - Automatic backoff on 429: pauses the whole pool for BACKOFF_MS then retries.
 *  - No artificial delay between requests — ViaCEP has no documented strict rate limit
 *    for reasonable concurrency levels.
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
  let idx = 0;

  const CONCURRENCY = 20;  // simultaneous in-flight requests
  const BACKOFF_MS  = 800; // pause on 429 before resuming

  const sleep = (ms: number) => new Promise<void>(r => setTimeout(r, ms));

  // Worker: continuously picks the next CEP until the queue is empty
  const worker = async () => {
    while (true) {
      if (signal?.aborted) break;
      const i = idx++;
      if (i >= unique.length) break;

      const cep = unique[i];
      let result: ViaCepResult | null = null;

      // Retry loop for transient rate-limits only
      for (let attempt = 0; attempt < 3; attempt++) {
        try {
          result = await lookupCep(cep, signal);
          break;
        } catch (err: any) {
          if (err?.name === 'AbortError') return;
          if (err?.status === 429 && attempt < 2) {
            await sleep(BACKOFF_MS * (attempt + 1));
            continue;
          }
          break; // other errors → treat as null
        }
      }

      results.set(cep, result);
      done++;
      onProgress(done, total);
    }
  };

  // Launch CONCURRENCY workers in parallel; they all share the idx counter
  await Promise.all(
    Array.from({ length: Math.min(CONCURRENCY, total) }, () => worker())
  );

  return results;
};
