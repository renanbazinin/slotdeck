# Model providers

slotdeck does not know any model service. It defines two small interfaces, and you pass objects that implement
them. This is how it works with a public API, a cloud account, or a gateway inside a network with no internet
access.

```ts
import type { LlmProvider, JsonRequest, JsonSchema, ImageProvider, ImageRequest } from 'slotdeck';
```

## `LlmProvider`: the text model

```ts
interface LlmProvider {
  /** Resolves with the parsed JSON value. */
  generateJson(request: JsonRequest): Promise<unknown>;
}

interface JsonRequest {
  system: string;       // the system prompt: the role and the rules
  prompt: string;       // the user message: the brief, the manifest's patterns, what to answer
  schema: JsonSchema;   // the JSON Schema the answer must follow
  images?: Array<{ mimeType: string; data: Uint8Array }>; // pictures to look at; only draftManifest sends them
  signal?: AbortSignal; // pass it to your HTTP call so work can be cancelled
}
```

The contract:

- Resolve with the **parsed** value (an object), not a string of JSON.
- The value should follow `schema`. slotdeck checks what it gets back against the template's manifest and asks
  again when content breaks a limit, but it expects the overall shape the schema describes.
- Reject (throw) when the call fails. slotdeck does not retry failed calls; add retries in your provider if you
  want them.
- One call is one request and one answer. There is no conversation state: every request carries everything it
  needs.

### What the schemas contain

Only a small, widely supported subset of JSON Schema:

- `type`: `object`, `array`, `string`, `integer`, `number`, `boolean`
- `properties`, `required`, `items`, `enum`, `description`
- in a few places `minItems`, `maxItems`, `minimum`, `maximum`

There is no `anyOf`, `oneOf`, `$ref`, `pattern` or `format`. If your backend rejects the length or range
keywords, your provider can delete them from the schema before sending it: the result is checked by code in any
case.

### A provider for a backend with structured output

Most hosted models and gateways accept a JSON Schema and return JSON that follows it. The adapter is then a
single HTTP call. The request and response fields below are placeholders: use the ones your gateway documents.

```ts
import type { LlmProvider } from 'slotdeck';

export function gatewayModel(options: { url: string; model: string; token: () => Promise<string> }): LlmProvider {
  return {
    async generateJson({ system, prompt, schema, signal }) {
      const response = await fetch(options.url, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json', Authorization: `Bearer ${await options.token()}` },
        body: JSON.stringify({
          model: options.model,
          messages: [
            { role: 'system', content: system },
            { role: 'user', content: prompt },
          ],
          // However your gateway asks for JSON that follows a schema:
          response_format: { type: 'json_schema', json_schema: { name: 'answer', schema } },
        }),
        signal,
      });
      if (!response.ok) throw new Error(`Model gateway answered ${response.status}: ${(await response.text()).slice(0, 300)}`);
      const body = await response.json();
      const text = body.choices[0].message.content; // wherever your gateway puts the answer text
      return JSON.parse(text);
    },
  };
}
```

### A provider for a model without structured output

Put the schema into the prompt and parse what comes back. Models often wrap JSON in a code fence or add a
sentence around it, so extract the JSON before parsing.

```ts
export function plainTextModel(complete: (system: string, prompt: string, signal?: AbortSignal) => Promise<string>): LlmProvider {
  return {
    async generateJson({ system, prompt, schema, signal }) {
      const instruction =
        `${prompt}\n\nAnswer with one JSON value that follows this JSON Schema, and nothing else:\n${JSON.stringify(schema)}`;
      const text = await complete(system, instruction, signal);
      const start = text.indexOf('{');
      const end = text.lastIndexOf('}');
      if (start < 0 || end <= start) throw new Error(`The model did not answer with JSON: ${text.slice(0, 200)}`);
      return JSON.parse(text.slice(start, end + 1));
    },
  };
}
```

Every top-level answer slotdeck asks for is a JSON object, so looking for the outermost braces is safe.

### Choosing a model

- The requests are small. Measured on the demo template (15 patterns): the outline request is a system prompt
  of about 600 characters, a prompt of about 3,600 characters and a schema of about 700 characters. The writing
  request for a 15-slide deck is a prompt of about 1,500 characters and a schema of about 9,800 characters. Any
  current general-purpose model handles this.
- Writing a deck is a single call that returns all slides, so the model needs enough output room for the whole
  deck. The JSON of a nine-slide deck is about 6 KB.
- `editManifest` sends the whole manifest with the measured shapes of every pattern's slide: about 58,000
  characters for a template with 24 patterns.
- `draftManifest` can send a picture of each slide (`images`). A provider for a model that cannot see may ignore
  them; the draft then relies on measurements alone.

### Timeouts and cancellation

slotdeck passes `signal` through when the caller gave one. Your provider should pass it to its HTTP call, and
should have a time limit of its own: a model call can hang without failing.

```ts
const limit = AbortSignal.timeout(180_000);
const combined = signal ? AbortSignal.any([signal, limit]) : limit;
await fetch(url, { ...init, signal: combined });
```

### Logging and auditing

A provider is the one place every model call passes through, so it is where to count calls, log prompts or
enforce a budget:

```ts
const counted: LlmProvider = {
  async generateJson(request) {
    const started = Date.now();
    try {
      return await inner.generateJson(request);
    } finally {
      log({ promptChars: request.prompt.length, ms: Date.now() - started });
    }
  },
};
```

## `ImageProvider`: the picture source

```ts
interface ImageProvider {
  /** Resolves with PNG, JPEG or GIF bytes. */
  generateImage(request: ImageRequest): Promise<Uint8Array>;
}

interface ImageRequest {
  prompt: string;       // a description of the picture, written by the planning model
  kind: 'photo' | 'icon' | 'graphic'; // what the slot holds; generated pictures are photos
  aspectRatio: number;  // width divided by height of the frame the picture goes into
  signal?: AbortSignal;
}
```

It does not have to be an image model. A stock library, an internal picture search or a fixed set of brand
photos all fit: return the bytes of the picture that best matches `prompt`.

```ts
const brandPhotos: ImageProvider = {
  async generateImage({ prompt, aspectRatio }) {
    const hit = await photoLibrary.search(prompt, { orientation: aspectRatio >= 1 ? 'landscape' : 'portrait' });
    return new Uint8Array(await hit.download());
  },
};
```

The picture is cropped to the frame when the slot's fit is `cover`, so any proportions work; matching
`aspectRatio` just loses less of the picture.

Without an `ImageProvider`, leave `images` out. `generateDeck` and the chat tools then plan decks with no
generated photos (patterns that need a photo are not offered), unless you supply photos as `assets`.

## The included Gemini adapters

```ts
import { gemini, geminiImages, vertexEndpoint, GeminiError, GEMINI_DEVELOPER_ENDPOINT } from 'slotdeck';

const llm = gemini({ model: 'gemini-3.8-flash', apiKey });
const images = geminiImages({ model: 'gemini-3.1-flash-image', apiKey });
```

They call Google's Interactions API with plain `fetch`, so they need network access to that API. Inside a
network without it, write a provider for your own gateway as shown above.

Options of both:

| Option | Meaning |
| --- | --- |
| `model` | Model id. Required. |
| `apiKey` | Gemini Developer API key, sent as the `x-goog-api-key` header. |
| `endpoint` | Interactions endpoint. Defaults to `GEMINI_DEVELOPER_ENDPOINT`. Use `vertexEndpoint(project, location?)` for Vertex AI, or the address of a proxy. |
| `headers` | Extra headers, or a function returning them (may be async), for example a fresh `Authorization: Bearer` token. |
| `fetch` | A custom `fetch`, for proxies and tests. |
| `timeoutMs` | Gives up on a request after this long. Default 180000 (3 minutes). `0` waits for ever. |

Only for `gemini`: `thinkingLevel` (`minimal`, `low`, `medium`, `high`) and `maxOutputTokens`.

Only for `geminiImages`: `imageSize` (`512`, `1K`, `2K`, `4K`; default `1K`) and `style`, a sentence appended to
every prompt, such as a house photo style.

Failures throw `GeminiError`, which has the HTTP `status` when there was one.

Vertex AI:

```ts
const llm = gemini({
  model: 'gemini-3.8-flash',
  endpoint: vertexEndpoint('my-project'), // location defaults to 'global'
  headers: async () => ({ Authorization: `Bearer ${await getAccessToken()}` }),
});
```

The adapters were run live against the Gemini Developer API. The Vertex endpoint is supported by the adapter
but has not been run live.

## Testing without a model

A scripted provider makes tests deterministic and free:

```ts
const scripted: LlmProvider = {
  async generateJson({ prompt, schema }) {
    return cannedAnswers.shift(); // objects shaped like the schema
  },
};
```
