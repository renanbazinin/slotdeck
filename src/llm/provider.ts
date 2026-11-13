import type { ImageKind } from '../template/manifest.ts';

/** A JSON Schema object. The planner only uses a small, widely supported subset. */
export type JsonSchema = { [key: string]: unknown };

export interface JsonRequest {
  system: string;
  prompt: string;
  /** Schema the answer must follow. Providers without schema support may treat it as guidance. */
  schema: JsonSchema;
  /** Pictures to look at alongside the prompt, e.g. a rendered slide. A provider that cannot see may ignore them. */
  images?: Array<{ mimeType: string; data: Uint8Array }>;
  signal?: AbortSignal;
}

/**
 * The one thing slotdeck needs from a language model: answer a prompt with JSON.
 * Implement this to plug in any model or gateway.
 */
export interface LlmProvider {
  /** Resolves with the parsed JSON value. */
  generateJson(request: JsonRequest): Promise<unknown>;
}

export interface ImageRequest {
  /** Description of the picture, written by the planning model. */
  prompt: string;
  kind: ImageKind;
  /** Width divided by height of the frame the picture goes into. */
  aspectRatio: number;
  signal?: AbortSignal;
}

/** Produces a picture for a slot. Implement this for any image model, stock library or search. */
export interface ImageProvider {
  /** Resolves with PNG, JPEG or GIF bytes. */
  generateImage(request: ImageRequest): Promise<Uint8Array>;
}
