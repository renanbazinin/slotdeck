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

  constructor(message: string, status?: number) {
    super(message);
    this.name = 'GeminiError';
    this.status = status;
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
}

async function createInteraction(
  connection: GeminiConnection,
  body: Record<string, unknown>,
  signal: AbortSignal | undefined,
  acceptCutOff = false,
): Promise<Interaction> {
  const extra = typeof connection.headers === 'function' ? await connection.headers() : connection.headers;
  const headers: Headers = { 'Content-Type': 'application/json', ...extra };
  if (connection.apiKey) headers['x-goog-api-key'] = connection.apiKey;

  const init: RequestInit = { method: 'POST', headers, body: JSON.stringify(body) };
  // A request can hang without failing, so every one has a limit of its own next to the caller's signal.
  const limit = connection.timeoutMs ?? 180_000;
  const timeout = limit > 0 ? AbortSignal.timeout(limit) : undefined;
  const combined = signal && timeout ? AbortSignal.any([signal, timeout]) : (signal ?? timeout);
  if (combined) init.signal = combined;

  let response: Response;
  let text: string;
  try {
    response = await (connection.fetch ?? fetch)(connection.endpoint ?? GEMINI_DEVELOPER_ENDPOINT, init);
    text = await response.text();
  } catch (error) {
    if (timeout?.aborted && !signal?.aborted) throw new GeminiError(`Gemini did not answer within ${Math.round(limit / 1000)} seconds`);
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
    throw new GeminiError(`Gemini request failed (${response.status}): ${message}`, response.status);
  }
  const interaction = (parsed ?? {}) as Interaction;
  // An answer cut off at the output limit is for the caller to deal with, when it says it can.
  if (interaction.status === 'incomplete' && acceptCutOff) return interaction;
  if (interaction.status !== 'completed') {
    const reason = interaction.errors?.map((error) => error.message).filter(Boolean).join('; ');
    throw new GeminiError(`Gemini interaction ended with status "${interaction.status ?? 'unknown'}"${reason ? `: ${reason}` : ''}`);
  }
  return interaction;
}

/**
 * Failures that the same request usually does not repeat: a server error, or the service's
 * recitation filter tripping on an ordinary answer. Bad keys, quotas and bad requests are not among them.
 */
function passesByItself(error: unknown): boolean {
  if (!(error instanceof GeminiError)) return false;
  return (error.status !== undefined && error.status >= 500 && error.status <= 504) || /recitation/i.test(error.message);
}

function outputBlocks(interaction: Interaction): ContentBlock[] {
  return (interaction.steps ?? []).filter((step) => step.type === 'model_output').flatMap((step) => step.content ?? []);
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

      const send = () => createInteraction(options, body, request.signal, true);
      let interaction: Interaction;
      try {
        interaction = await send();
      } catch (error) {
        // A failure that comes and goes by chance is worth one more try; any other is the caller's to see.
        if (!passesByItself(error)) throw error;
        interaction = await send();
      }
      if (interaction.status === 'incomplete') {
        // Running away is a matter of chance too, so the same request is sent once more.
        interaction = await send();
        if (interaction.status === 'incomplete') {
          throw new GeminiError(
            `Gemini's answer was cut off at the limit of ${limit} tokens, twice. The model is probably running away inside a text field; try again, or set thinkingLevel to "low"`,
          );
        }
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
      const interaction = await createInteraction(options, body, request.signal);
      const image = outputBlocks(interaction).find((block) => block.type === 'image' && block.data);
      if (!image?.data) throw new GeminiError('Gemini returned no image');
      return Uint8Array.from(atob(image.data), (char) => char.charCodeAt(0));
    },
  };
}
