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
  maxOutputTokens?: number; // a ceiling on the answer's size in tokens, reasoning not counted; see "Runaway answers"
  effort?: 'low';       // set on short structural requests: use light reasoning if your model has a setting for it
  images?: Array<{ mimeType: string; data: Uint8Array }>; // pictures to look at; only draftManifest sends them
  signal?: AbortSignal; // pass it to your HTTP call so work can be cancelled
}
```

The contract:

- Resolve with the **parsed** value (an object), not a string of JSON.
- The value should follow `schema`. slotdeck checks what it gets back against the template's manifest and asks
  again when content breaks a limit, but it expects the overall shape the schema describes.
- Reject (throw) when the call fails. The planner does not retry failed calls; retries belong in the provider,
  which knows which failures pass by themselves (see "Runaway answers" and the Gemini adapter below).
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

- The requests are small. Measured on the demo template when it had 15 patterns: the outline request is a system prompt
  of about 600 characters, a prompt of about 3,600 characters and a schema of about 700 characters. The writing
  request for a 15-slide deck is a prompt of about 1,500 characters and a schema of about 9,800 characters. Any
  current general-purpose model handles this.
- Writing a deck is a single call that returns all slides, so the model needs enough output room for the whole
  deck. The JSON of a nine-slide deck is about 6 KB.
- `editManifest` sends the whole manifest with the measured shapes of every pattern's slide: about 58,000
  characters for a template with 24 patterns.
- `draftManifest` can send a picture of each slide (`images`). A provider for a model that cannot see may ignore
  them; the draft then relies on measurements alone.

### Runaway answers: always set an output limit

Now and then a model runs away inside one text field of its answer: it keeps adding sentence fragments to a
single string until something stops it. This was seen with a planning request whose proper answer is about 200
tokens: the model wrote more than 27,000 tokens into one field. Without a limit such a call takes minutes, costs
tens of thousands of output tokens, and ends in a timeout. It is a matter of chance: the same request usually
works the next time.

It was tied to reasoning. With the model's default reasoning, a request for a text-only plan ran away in 6 of
24 calls. With light reasoning it did not run away in 18 calls, and each took 4 seconds instead of about 35.
That is why planning requests carry `effort: 'low'`.

A long reasoning before a short answer can also fill the limit by itself, where the model counts reasoning
against it. A writing request whose answer was about 550 tokens took 5,300 to 10,900 reasoning tokens, and one
send in five was cut off at its limit of 11,334: a normal answer, cut off part of the way through. Asked again
with light reasoning it took none, and every answer was complete. It is no setting for the first send of every
request, though: where the writer copied a passage of 55 words into a layout's two columns, it kept 41 to 44 of
them with light reasoning and 46 to 50 with the default.

So a provider should do three things:

1. **Use light reasoning when a request says `effort: 'low'`**, if your model has such a setting.
2. **Limit the output of every call.** The planner passes `maxOutputTokens` with its requests: a generous
   ceiling for that answer (for an outline of six slides, 1,360 tokens). Pass it to your model as its output
   limit. If your model counts reasoning tokens against the same limit, add room for them. When a request
   carries no estimate, use a default of your own rather than no limit.
3. **Ask once more, with light reasoning, when an answer is cut off at the limit.** A cut-off answer is not
   valid JSON. Send the request once more with the same limit and your model's light reasoning setting, and
   throw if it is cut off again. The first send keeps the reasoning of point 1.

```ts
async generateJson({ system, prompt, schema, maxOutputTokens, effort, signal }) {
  const limit = (maxOutputTokens ?? 16_000) + 8_000; // room for reasoning, if your model counts it here
  const reasoning = effort === 'low' ? 'low' : undefined; // whatever your model calls its light setting
  const answer = await callModel({ system, prompt, schema, limit, reasoning, signal });
  if (!answer.cutOffAtLimit) return JSON.parse(answer.text);
  // Cut off: once more with light reasoning, which leaves the limit to the answer.
  const again = await callModel({ system, prompt, schema, limit, reasoning: 'low', signal });
  if (!again.cutOffAtLimit) return JSON.parse(again.text);
  throw new Error(`The model's answer was cut off at ${limit} tokens twice`);
}
```

The included Gemini adapter does all three. Its limit is the request's estimate plus 8,192 tokens for reasoning
(24,576 when there is no estimate, or its own `maxOutputTokens` option when set). It uses thinking level `low`
for requests with `effort: 'low'`, unless `thinkingLevel` is set on the adapter, which then applies to every
request. It sends a cut-off request once more with the same limit and thinking level `low`, or with the
adapter's `thinkingLevel` when one is set; the first send is as before. When the answer is cut off again, the
error says what filled the limit the second time, reasoning or the answer itself, where the service's counts
show it. It also sends a request once more after a server error (500 to 504) or when the service's recitation
filter blocks an ordinary answer, which happens at random; a bad key, a quota or a bad request is not repeated.

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
| `retries` | How many times a request that failed for a passing reason is sent again (from 0.7.0): a rate limit (429) that is no spending cap or daily quota, 408, 500, 502, 503, 504, the recitation filter, a network that failed, or, within a `deadlineMs`, no answer within `timeoutMs`. A 429 is tried again unless it is a spending cap or a quota of the day (by its message, or by the quota it names): Gemini's limit per minute also says to check billing details. Each try waits longer: one second doubled with each try, taken at random between half and all of it, up to a minute, or as long as the service asks (`Retry-After`, or its `RetryInfo`); a wait of more than a minute is not waited. Default 2, three tries in all. `0` sends once. Before 0.7.0 a text request was sent once more after a server error, and a picture never. |
| `deadlineMs` | Gives up on a call after this long, its tries and its waits together (from 0.7.0): no try runs past it, and no wait is begun that would end after it. One deadline covers the whole call, an answer sent again after a cut-off too. Left out, only `timeoutMs` limits each try, and a try that times out is not sent again. Set it where a person waits, such as `120000`. |

Only for `gemini`: `thinkingLevel` (`minimal`, `low`, `medium`, `high`) and `maxOutputTokens`, the output limit
of a request with reasoning included. By default the limit is the request's own estimate plus 8,192 tokens, or
24,576 when the request has none. Set on the adapter, it replaces the limit of every request, larger or smaller.

Only for `geminiImages`: `imageSize` (`512`, `1K`, `2K`, `4K`; default `1K`) and `style`, a sentence appended to
every prompt, such as a house photo style.

Failures throw `GeminiError`, which has the HTTP `status` when there was one, the wait the service asked for
(`retryAfterMs`), and `timedOut` when no answer came in time. `retryDelay(error, attempt)` is the rule the
adapters wait by, for a provider of your own: the milliseconds to wait before another try after `error` on try
number `attempt`, or `undefined` where another try is not worth it.

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
