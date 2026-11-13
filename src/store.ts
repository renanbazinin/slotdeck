/**
 * Where plans, decks, pictures and rendered files are kept between calls. Keys look like paths
 * (`deck/ab12.json`, `asset/gen-1f2e`), so a folder on a volume, a bucket or a table all fit.
 *
 * Pictures are stored on their own, apart from the deck JSON and the pptx. The JSON stays a
 * few kilobytes of text, which is all a language model ever sees; a pptx can be rebuilt from
 * the JSON and the pictures whenever it is wanted.
 */
export interface BlobStore {
  get(key: string): Promise<Uint8Array | undefined>;
  set(key: string, data: Uint8Array): Promise<void>;
}

/** A store that lives in memory, for tests and short-lived processes. */
export function memoryStore(): BlobStore & { keys(): string[] } {
  const blobs = new Map<string, Uint8Array>();
  return {
    async get(key) {
      return blobs.get(key);
    },
    async set(key, data) {
      blobs.set(key, data);
    },
    keys: () => [...blobs.keys()],
  };
}

const encoder = new TextEncoder();
const decoder = new TextDecoder();

export async function readJson<T>(store: BlobStore, key: string): Promise<T | undefined> {
  const data = await store.get(key);
  return data ? (JSON.parse(decoder.decode(data)) as T) : undefined;
}

export async function writeJson(store: BlobStore, key: string, value: unknown): Promise<void> {
  await store.set(key, encoder.encode(JSON.stringify(value)));
}
