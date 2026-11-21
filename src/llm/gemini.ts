// Gemini through the Interactions API, over plain fetch (no SDK dependency).
// Contract: https://ai.google.dev/static/api/interactions-v1.openapi.json
import type { ImageProvider, ImageRequest, JsonRequest, LlmProvider } from './provider.ts';

export const GEMINI_DEVELOPER_ENDPOINT = 'https://generativelanguage.googleapis.com/v1/interactions';

/** Endpoint of the Interactions API on Vertex AI. Authenticate with a bearer token via `headers`. */
export function vertexEndpoint(project: string, location = 'global'): string {
  const host = location === 'global' ? 'aiplatform.googleapis.com' : `${location}-aiplatform.googleapis.com`;
  return `https://${host}/v1beta1/projects/${project}/locations/${location}/interactions`;
}

type Headers = Record<string, string>;

interface GeminiConnection {
  model: string;
  /** Gemini Developer API key, sent as `x-goog-api-key`. */
  apiKey?: string;
  /** Interactions endpoint. Defaults to the Developer API; use {@link vertexEndpoint} for Vertex AI. */
  endpoint?: string;
  /** Extra headers, or a function returning them, e.g. a fresh `Authorization: Bearer` token. */
  headers?: Headers | (() => Headers | Promise<Headers>);
  /** Custom fetch, for proxies and tests. */
  fetch?: typeof fetch;
  /** Gives up on a request that has not answered after this long. Defaults to 3 minutes; 0 waits for ever. */
  timeoutMs?: number;
  /**
   * How many times a request that failed for a passing reason is sent again: a rate limit (429)
   * that is no spending cap or daily quota, a request the service timed out (408), a server error
   * (500, 502, 503, 504), the recitation filter tripping on an ordinary answer, a network that
   * failed, or no answer within `timeoutMs`. Each try waits longer than the last, by a doubling
   * step with some chance in it, or as long as the service asks. Defaults to 2: three tries in all.
   * 0 sends every request once. From 0.7.0; before, a text request was sent once more after a
   * server error, without waiting, and a picture was never sent again.
   */
  retries?: number;
  /**
   * Gives up on a call, its tries and its waits together, after this long, in milliseconds: no try
   * runs past it, and no wait is begun that would end after it. Left out, only `timeoutMs` limits
   * each try. Set it where a person waits for the answer. From 0.7.0.
   */
  deadlineMs?: number;
}

export interface GeminiOptions extends GeminiConnection {
  thinkingLevel?: 'minimal' | 'low' | 'medium' | 'high';
  /**
   * Output limit of a request in tokens, reasoning included. By default it is the caller's own
   * estimate of the answer plus 8192 for reasoning, or 24576 when the caller gives no estimate.
   */
  maxOutputTokens?: number;
}

/** Room for an answer when the caller does not say how large it may be. */
const DEFAULT_ANSWER_TOKENS = 16_384;
/** Room for reasoning, which Gemini counts against the same limit as the answer. */
const REASONING_TOKENS = 8_192;

export interface GeminiImageOptions extends GeminiConnection {
  /** Output size class. Defaults to `1K`. */
  imageSize?: '512' | '1K' | '2K' | '4K';
  /** Appended to every prompt, e.g. a house photo style. */
  style?: string;
}

export class GeminiError extends Error {
  readonly status: number | undefined;
  /** How long the service asked to wait before the next try, in milliseconds, where it said. */
  readonly retryAfterMs: number | undefined;
  /** No answer came within `timeoutMs`. */
  readonly timedOut: boolean;

  constructor(message: string, status?: number, more: { retryAfterMs?: number | undefined; timedOut?: boolean } = {}) {
    super(message);
    this.name = 'GeminiError';
    this.status = status;
    this.retryAfterMs = more.retryAfterMs;
    this.timedOut = more.timedOut === true;
  }
}

/** The first wait before another try, and the longest: a longer wait the service asks for is not waited. */
const FIRST_WAIT = 1_000;
const LONGEST_WAIT = 60_000;

/**
 * Whether a request that failed so is worth another try: a passing failure, which the same request
 * usually does not meet again. A bad key, a bad request, a spending cap or a daily quota is met
 * again however often it is sent; so is a stop the caller asked for.
 */
function passing(error: unknown): boolean {
  if (error instanceof GeminiError) {
    if (error.timedOut) return true;
    if (error.status === 429) return !/spending|billing|per day|perday|daily|limit: 0\b/i.test(error.message);
    return error.status === 408 || error.status === 500 || error.status === 502 || error.status === 503 || error.status === 504 || /recitation/i.test(error.message);
  }
  // fetch fails with a TypeError when the network does ("fetch failed"); a stop is an AbortError.
  return error instanceof TypeError;
}

/**
 * How long to wait, in milliseconds, before trying a request again that failed with `error` on try
 * number `attempt` (from 1); `undefined` where another try is not worth it. The wait doubles from
 * one second with each try, taken at random between half and all of it so that requests that failed
 * together do not come back together, up to a minute; where the service says how long to wait, that
 * long, and no more tries where it asks for more than a minute.
 */
export function retryDelay(error: unknown, attempt: number, random: () => number = Math.random): number | undefined {
  if (!passing(error)) return undefined;
  const asked = error instanceof GeminiError ? error.retryAfterMs : undefined;
  if (asked !== undefined) return asked <= LONGEST_WAIT ? Math.max(0, asked) : undefined;
  const step = Math.min(LONGEST_WAIT, FIRST_WAIT * 2 ** Math.max(0, attempt - 1));
  return Math.round(step * (0.5 + random() / 2));
}

/** The wait a failed answer asks for: its Retry-After header (seconds, or a date), or the service's RetryInfo ("37s"). */
function askedWait(response: Response, parsed: unknown): number | undefined {
  const header = response.headers.get('retry-after');
  if (header !== null && header.trim() !== '') {
    const seconds = Number(header);
    if (Number.isFinite(seconds)) return Math.max(0, seconds * 1000);
    const date = Date.parse(header);
    if (Number.isFinite(date)) return Math.max(0, date - Date.now());
  }
  const details = (parsed as { error?: { details?: unknown } } | undefined)?.error?.details;
  for (const detail of Array.isArray(details) ? details : []) {
    const delay = (detail as { retryDelay?: unknown } | null)?.retryDelay;
    const seconds = typeof delay === 'string' ? /^(\d+(?:\.\d+)?)s$/.exec(delay.trim())?.[1] : undefined;
    if (seconds !== undefined) return Number(seconds) * 1000;
  }
  return undefined;
}

/** Waits so long, or until the caller stops, which ends the wait with the caller's reason. */
function pause(ms: number, signal: AbortSignal | undefined): Promise<void> {
  if (signal?.aborted) return Promise.reject(signal.reason);
  return new Promise((resolve, reject) => {
    const stop = () => {
      clearTimeout(timer);
      reject(signal?.reason);
    };
    const timer = setTimeout(() => {
      signal?.removeEventListener('abort', stop);
      resolve();
    }, ms);
    signal?.addEventListener('abort', stop, { once: true });
  });
}

/**
 * Runs a request, and runs it again after a passing failure, as `retries` and `deadlineMs` allow.
 * `run` is given the most time its try may take, in milliseconds (`Infinity` for no limit).
 */
async function withRetries<T>(connection: GeminiConnection, signal: AbortSignal | undefined, run: (most: number) => Promise<T>): Promise<T> {
  const tries = 1 + Math.max(0, Math.floor(connection.retries ?? 2));
  const end = connection.deadlineMs !== undefined && connection.deadlineMs > 0 ? Date.now() + connection.deadlineMs : Infinity;
  for (let attempt = 1; ; attempt++) {
    try {
      return await run(end - Date.now());
    } catch (error) {
      if (signal?.aborted) throw error;
      const wait = attempt < tries ? retryDelay(error, attempt) : undefined;
      if (wait === undefined || Date.now() + wait >= end) throw error;
      await pause(wait, signal);
    }
  }
}

interface ContentBlock {
  type?: string;
  text?: string;
  data?: string;
  mime_type?: string;
}

interface Interaction {
  status?: string;
  steps?: Array<{ type?: string; content?: ContentBlock[] }>;
  errors?: Array<{ message?: string }>;
  usage?: { total_output_tokens?: number; total_thought_tokens?: number };
}

async function createInteraction(
  connection: GeminiConnection,
  body: Record<string, unknown>,
  signal: AbortSignal | undefined,
  acceptCutOff = false,
  most = Infinity,
): Promise<Interaction> {
  const extra = typeof connection.headers === 'function' ? await connection.headers() : connection.headers;
  const headers: Headers = { 'Content-Type': 'application/json', ...extra };
  if (connection.apiKey) headers['x-goog-api-key'] = connection.apiKey;

  const init: RequestInit = { method: 'POST', headers, body: JSON.stringify(body) };
  // A request can hang without failing, so every one has a limit of its own next to the caller's
  // signal: its own, or the time left before the call's deadline where that is less.
  const own = connection.timeoutMs ?? 180_000;
  if (!(most > 0)) throw new GeminiError('Gemini did not answer before the deadline of the call');
  const limit = Math.min(own > 0 ? own : Infinity, most);
  const timeout = Number.isFinite(limit) ? AbortSignal.timeout(Math.max(1, Math.round(limit))) : undefined;
  const combined = signal && timeout ? AbortSignal.any([signal, timeout]) : (signal ?? timeout);
  if (combined) init.signal = combined;

  let response: Response;
  let text: string;
  try {
    response = await (connection.fetch ?? fetch)(connection.endpoint ?? GEMINI_DEVELOPER_ENDPOINT, init);
    text = await response.text();
  } catch (error) {
    if (timeout?.aborted && !signal?.aborted) throw new GeminiError(`Gemini did not answer within ${Math.round(limit / 1000)} seconds`, undefined, { timedOut: true });
    throw error;
  }
  let parsed: unknown;
  try {
    parsed = JSON.parse(text);
  } catch {
    parsed = undefined;
  }
  if (!response.ok) {
    const message = (parsed as { error?: { message?: string } } | undefined)?.error?.message ?? text.slice(0, 300);
    throw new GeminiError(`Gemini request failed (${response.status}): ${message}`, response.status, { retryAfterMs: askedWait(response, parsed) });
  }
  const interaction = (parsed ?? {}) as Interaction;
  // An answer cut off at the output limit is for the caller to deal with, when it says it can.
  if (interaction.status === 'incomplete' && acceptCutOff) return interaction;
  if (interaction.status !== 'completed') {
    // A gateway may answer in another shape than the service: what is not a list of errors with messages says nothing.
    const errors: unknown[] = Array.isArray(interaction.errors) ? interaction.errors : [];
    const reason = errors
      .map((error) => (typeof error === 'object' && error !== null ? (error as { message?: unknown }).message : undefined))
      .filter((message) => typeof message === 'string' && message !== '')
      .join('; ');
    throw new GeminiError(`Gemini interaction ended with status "${interaction.status ?? 'unknown'}"${reason ? `: ${reason}` : ''}`);
  }
  return interaction;
}

/** The blocks the model wrote. Steps and blocks of another shape than the service's (from a gateway, say) are passed over. */
function outputBlocks(interaction: Interaction): ContentBlock[] {
  const isObject = (value: unknown): boolean => typeof value === 'object' && value !== null && !Array.isArray(value);
  const steps: unknown[] = Array.isArray(interaction.steps) ? interaction.steps : [];
  return steps
    .filter((step): step is { type: string; content?: unknown } => isObject(step) && (step as { type?: unknown }).type === 'model_output')
    .flatMap((step) => (Array.isArray(step.content) ? step.content : []).filter((block): block is ContentBlock => isObject(block)));
}

/**
 * The tokens an answer cut off at `limit` spent reasoning and writing, when the service says so and they fill
 * the limit. A gateway may say nothing, say it in another shape, or give one count only; what it gives is
 * used only when it accounts for the cut-off. The service counts both against the limit: an answer seen cut
 * off at 11334 had spent 11319.
 */
function tokensSpent(interaction: Interaction, limit: number): { thought: number | undefined; output: number | undefined } | undefined {
  const usage: unknown = interaction.usage;
  if (typeof usage !== 'object' || usage === null) return undefined;
  const count = (value: unknown): number | undefined => (typeof value === 'number' && Number.isFinite(value) && value >= 0 ? value : undefined);
  const thought = count((usage as { total_thought_tokens?: unknown }).total_thought_tokens);
  const output = count((usage as { total_output_tokens?: unknown }).total_output_tokens);
  const total = (thought ?? 0) + (output ?? 0);
  return total > 0 && total >= 0.9 * limit ? { thought, output } : undefined;
}

/** An answer cut off twice, with what filled the limit the second time when the service says. */
function cutOffTwice(limit: number, last: Interaction, thinkingLevel: string): GeminiError {
  // The adapter's limit replaces that of every request, so one set for this answer must not lower another's.
  const larger = `set maxOutputTokens on the adapter above ${limit}, and above any other request's limit, since it replaces them all`;
  const spent = tokensSpent(last, limit);
  let cause: string;
  if (!spent) cause = `Try again, or for a large answer ${larger}`;
  else if ((spent.thought ?? 0) > (spent.output ?? 0)) {
    // The retry reasons as lightly as the adapter lets it; a level set on the adapter is the caller's to lower.
    const lower = thinkingLevel === 'medium' || thinkingLevel === 'high' ? `set a lower thinkingLevel than "${thinkingLevel}", or ` : '';
    const answer = spent.output === undefined ? '' : ` and the answer ${spent.output}`;
    cause = `The second time reasoning took ${spent.thought} of them${answer}: try again, or ${lower}${larger}`;
  } else {
    const reasoning = spent.thought === undefined ? '' : ` and reasoning ${spent.thought}`;
    cause =
      `The second time the answer took ${spent.output} of them${reasoning}: the model ran away inside a text field, ` +
      `or the answer is larger than its estimate. Try again, or for a large answer ${larger}`;
  }
  return new GeminiError(`Gemini's answer was cut off at the limit of ${limit} tokens, twice. ${cause}`);
}

/** A Gemini text model as an {@link LlmProvider}. */
export function gemini(options: GeminiOptions): LlmProvider {
  return {
    async generateJson(request: JsonRequest): Promise<unknown> {
      // Every request has an output limit. Now and then a model runs away inside one text field and
      // writes until it is stopped; without a limit that takes minutes and tens of thousands of tokens.
      const limit = options.maxOutputTokens ?? Math.ceil(request.maxOutputTokens ?? DEFAULT_ANSWER_TOKENS) + REASONING_TOKENS;
      const generationConfig: Record<string, unknown> = { max_output_tokens: limit };
      // A level set on the adapter is the caller's decision; otherwise a request that needs little reasoning gets little.
      if (options.thinkingLevel) generationConfig.thinking_level = options.thinkingLevel;
      else if (request.effort === 'low') generationConfig.thinking_level = 'low';

      // With pictures, the input becomes one user step holding the text and each picture.
      const input = request.images?.length
        ? [
            {
              type: 'user_input',
              content: [
                { type: 'text', text: request.prompt },
                ...request.images.map((image) => ({ type: 'image', mime_type: image.mimeType, data: toBase64(image.data) })),
              ],
            },
          ]
        : request.prompt;
      const body: Record<string, unknown> = {
        model: options.model,
        input,
        system_instruction: request.system,
        response_format: { type: 'text', mime_type: 'application/json', schema: request.schema },
        // One-shot calls: nothing to chain to, so nothing needs to be kept server-side.
        store: false,
      };
      body.generation_config = generationConfig;

      // A failure that comes and goes by chance is worth another try; any other is the caller's to see.
      const send = (sent: Record<string, unknown>) => withRetries(options, request.signal, (most) => createInteraction(options, sent, request.signal, true, most));
      let interaction = await send(body);
      if (interaction.status === 'incomplete') {
        // Cut off at the limit: by a model running away inside a text field, or by a long reasoning before a short
        // answer, which was seen to fill the limit by itself. Both are a matter of chance, so the request is sent
        // once more, with light thinking unless the adapter sets a level. The limit stays.
        const thinkingLevel = options.thinkingLevel || 'low';
        interaction = await send({ ...body, generation_config: { ...generationConfig, thinking_level: thinkingLevel } });
        if (interaction.status === 'incomplete') throw cutOffTwice(limit, interaction, thinkingLevel);
      }
      const text = outputBlocks(interaction)
        .filter((block) => block.type === 'text')
        .map((block) => block.text ?? '')
        .join('');
      try {
        return JSON.parse(text);
      } catch {
        throw new GeminiError(`Gemini did not return valid JSON: ${text.slice(0, 200)}`);
      }
    },
  };
}

function toBase64(data: Uint8Array): string {
  let binary = '';
  for (let offset = 0; offset < data.length; offset += 0x8000) binary += String.fromCharCode(...data.subarray(offset, offset + 0x8000));
  return btoa(binary);
}

const ASPECT_RATIOS = ['1:1', '2:3', '3:2', '3:4', '4:3', '4:5', '5:4', '9:16', '16:9', '21:9', '1:8', '8:1', '1:4', '4:1'];

/** The supported output ratio closest to the frame (compared on a log scale, so 2:1 and 1:2 are equally far from 1:1). */
export function nearestAspectRatio(ratio: number): string {
  let best = '1:1';
  let bestDistance = Infinity;
  for (const candidate of ASPECT_RATIOS) {
    const [w, h] = candidate.split(':').map(Number) as [number, number];
    const distance = Math.abs(Math.log(ratio) - Math.log(w / h));
    if (distance < bestDistance) {
      best = candidate;
      bestDistance = distance;
    }
  }
  return best;
}

/** A Gemini image model as an {@link ImageProvider}. */
export function geminiImages(options: GeminiImageOptions): ImageProvider {
  return {
    async generateImage(request: ImageRequest): Promise<Uint8Array> {
      const prompt = options.style ? `${request.prompt}\n\n${options.style}` : request.prompt;
      const body = {
        model: options.model,
        input: [{ type: 'user_input', content: [{ type: 'text', text: prompt }] }],
        response_format: [
          { type: 'text' },
          { type: 'image', aspect_ratio: nearestAspectRatio(request.aspectRatio), image_size: options.imageSize ?? '1K' },
        ],
        store: false,
      };
      const interaction = await withRetries(options, request.signal, (most) => createInteraction(options, body, request.signal, false, most));
      const image = outputBlocks(interaction).find((block) => block.type === 'image' && block.data);
      if (!image?.data) throw new GeminiError('Gemini returned no image');
      // The picture comes as base64 text. Anything else, or text that is not base64, is no picture.
      let binary: string;
      try {
        if (typeof image.data !== 'string') throw new TypeError('not text');
        binary = atob(image.data);
      } catch {
        throw new GeminiError('Gemini returned an image that cannot be read');
      }
      return Uint8Array.from(binary, (char) => char.charCodeAt(0));
    },
  };
}
