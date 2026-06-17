// AES-256-GCM Format-Preserving Encryption — Browser Simulation
// Spec: AES256_GCM_ENCRYPTION.md (xorshift128+ keystream + FPE layer)

// ── §9 — xorshift128+ PRNG ────────────────────────────────────────────────────
function makeKeystream(seed: number) {
  let a = ((seed ^ 0x9E3779B9) >>> 0) || 1;
  let b = ((seed ^ 0x6C62272E) >>> 0) || 2;
  return () => {
    a ^= a << 13; a = a >>> 0;
    a ^= a >> 17;
    a ^= a << 5;  a = a >>> 0;
    b ^= b >> 7;  b = b >>> 0;
    b ^= b << 9;  b = b >>> 0;
    b ^= b >> 8;  b = b >>> 0;
    return (((a + b) >>> 0) / 0x100000000);
  };
}

// ── §8.1 — Random key from seed ───────────────────────────────────────────────
function generateRandomKey(seed: number): string {
  const rng = makeKeystream((seed ^ 0xDEADBEEF) >>> 0);
  const bytes: string[] = [];
  for (let i = 0; i < 32; i++)
    bytes.push(Math.floor(rng() * 256).toString(16).padStart(2, "0"));
  return bytes.join("");
}

// ── §8.2 — PBKDF2-like passphrase key ────────────────────────────────────────
function deriveKeyFromPassphrase(passphrase: string, iterations: number): string {
  let h = 0x5A827999;
  for (let i = 0; i < passphrase.length; i++)
    h = (Math.imul(h, 31) + passphrase.charCodeAt(i)) >>> 0;
  const rng = makeKeystream(h);
  for (let i = 0; i < Math.min(iterations, 200); i++) rng();
  const bytes: string[] = [];
  for (let i = 0; i < 32; i++)
    bytes.push(Math.floor(rng() * 256).toString(16).padStart(2, "0"));
  return bytes.join("");
}

// ── §11 — Column IV hash (deterministic per key+col) ─────────────────────────
function hashColIV(keyHex: string, colName: string): number {
  let h = parseInt(keyHex.slice(0, 8), 16) ^ 0xA5A5A5A5;
  const s = "COL\x00" + colName;
  for (let i = 0; i < s.length; i++)
    h = (Math.imul(h, 1664525) + s.charCodeAt(i) + 1013904223) >>> 0;
  return h;
}

// ── §12 — Per-cell keystream bytes (AES-GCM simulation) ──────────────────────
function makeCellKsBytes(size: number, keyHex: string, ivSeed: number): Uint8Array {
  const combined = (parseInt(keyHex.slice(0, 8), 16) ^ ivSeed) >>> 0;
  const ksRng = makeKeystream(combined);
  const ksBytes = new Uint8Array(size);
  for (let i = 0; i < size; i++)
    ksBytes[i] = Math.floor(ksRng() * 256);
  return ksBytes;
}

// ── §10 — Format-preserving encryption ───────────────────────────────────────
function encryptFPECell(ksBytes: Uint8Array, value: string): string {
  const isAllNumeric = /^\d+$/.test(value) && value.length > 1;
  let ki = 0;
  return [...value].map((ch, idx) => {
    const code = ch.charCodeAt(0);
    const k = ksBytes[ki++ % ksBytes.length];
    if (code >= 48 && code <= 57) {
      if (isAllNumeric && idx === 0) {
        // Leading digit: keep in 1–9 (no leading zeros)
        const d = code - 49; // map 1–9 → 0–8
        return String.fromCharCode(49 + ((d + (k % 9) + 9) % 9));
      }
      return String.fromCharCode(48 + ((code - 48 + k) % 10));
    } else if (code >= 65 && code <= 90) {
      return String.fromCharCode(65 + ((code - 65 + k) % 26));
    } else if (code >= 97 && code <= 122) {
      return String.fromCharCode(97 + ((code - 97 + k) % 26));
    }
    return ch;
  }).join("");
}

// ── Public API ────────────────────────────────────────────────────────────────

export interface AnonymizeOptions {
  keyMode: "random" | "pbkdf2";
  seed: number;
  passphrase: string;
  pbkdf2Iterations: number;
  deterministic: boolean;
}

export interface AnonymizeResult {
  rows: Record<string, string>[];
  keyHex: string;
}

export async function anonymizeRows(
  rows: Record<string, string>[],
  columns: string[],
  options: AnonymizeOptions,
  onProgress: (pct: number) => void
): Promise<AnonymizeResult> {
  const keyHex =
    options.keyMode === "pbkdf2" && options.passphrase.trim().length > 0
      ? deriveKeyFromPassphrase(options.passphrase, options.pbkdf2Iterations)
      : generateRandomKey(options.seed);

  // Pre-compute per-column keystreams (deterministic mode — §6 Step 4)
  const colKsBytes: Record<string, Uint8Array> = {};
  if (options.deterministic) {
    for (const col of columns) {
      colKsBytes[col] = makeCellKsBytes(256, keyHex, hashColIV(keyHex, col));
    }
  }

  const result: Record<string, string>[] = [];
  const total = rows.length;
  const cache = new Map<string, string>(); // deterministic cache
  let ivCounter = 0;
  const CHUNK = 5000;

  for (let ri = 0; ri < total; ri++) {
    const row = { ...rows[ri] };

    for (const col of columns) {
      const val = row[col] ?? "";
      if (!val.trim()) continue;

      if (options.deterministic) {
        const cacheKey = col + "\x00" + val;
        if (cache.has(cacheKey)) {
          row[col] = cache.get(cacheKey)!;
        } else {
          const enc = encryptFPECell(colKsBytes[col], val);
          cache.set(cacheKey, enc);
          row[col] = enc;
        }
      } else {
        ivCounter = (ivCounter + 1) >>> 0;
        const ksBytes = makeCellKsBytes(val.length + 32, keyHex, ivCounter);
        row[col] = encryptFPECell(ksBytes, val);
      }
    }

    result.push(row);

    if (ri % CHUNK === 0) {
      onProgress(Math.min(99, Math.round((ri / total) * 100)));
      await new Promise((r) => setTimeout(r, 0));
    }
  }

  onProgress(100);
  return { rows: result, keyHex };
}
