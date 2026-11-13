# Chat tools

`presentationTools` packages the planner as four tools for a chat model ("function calling"). The chat model
talks to the person; the tools do the work; the person approves a plan before anything is written.

```ts
import { presentationTools, TOOL_NAMES, storeKeys, memoryStore } from 'slotdeck';
import { directoryStore } from 'slotdeck/node';
```

## The four tools

| Tool name | What it does | Model calls inside |
| --- | --- | --- |
| `plan_presentation` | Decides the slides, the layout of each and the point each makes. Writes nothing yet. | 1 |
| `edit_presentation_plan` | Changes a plan by instruction. Returns a new plan with a new id. | 1 |
| `create_presentation` | Writes the approved plan's content, makes its pictures, renders the file. | 1, plus repairs, plus one per photo |
| `edit_presentation` | Changes a created presentation by instruction. Returns a new presentation with a new id. | 1 or 2, plus repairs |

The intended conversation:

1. The person asks for a deck. The chat model calls `plan_presentation`.
2. The chat model shows the plan. The person asks for changes; the chat model calls `edit_presentation_plan`
   as often as needed.
3. The person approves. The chat model calls `create_presentation`.
4. The person asks for changes to the result. The chat model calls `edit_presentation`.

## Setting up

```ts
const tools = presentationTools({
  template,                                     // from loadTemplate
  llm,                                          // LlmProvider used inside the tools
  store: directoryStore('/mnt/presentations'),  // where plans, decks, pictures and files are kept
  images,                                       // optional ImageProvider; without it decks have no generated photos
  icons,                                        // optional IconDef[]
  assets,                                       // optional AssetDef[]
  layout: 'adaptive',                           // optional; default 'strict'
  fields: { organizationDepartment: 'Acme  |  Operations' }, // optional deck-level values every deck gets
  numberPatterns: 'from-brief',                 // optional
  thumbnail: (patternId) => `/static/thumbnails/${patternId}.png`, // optional address of a picture of each layout
  newId: (kind) => `${kind}_${crypto.randomUUID()}`, // optional; default is a short random id
  onProgress: (event, tool) => console.log(tool, event.stage), // optional
});
```

The `llm` here is the model the tools use to plan and write. It can be the same model as the chat model or a
different one.

## Giving the tools to the chat model

`tools.definitions` is an array of `{ name, description, parameters }`, where `parameters` is a JSON Schema
object. Wrap it in whatever shape your chat model's API expects:

```ts
// One common shape:
const declarations = tools.definitions.map((tool) => ({ type: 'function', ...tool }));

// Another:
const declarations = tools.definitions.map((tool) => ({ type: 'function', function: tool }));
```

When the chat model calls a tool, pass the name and the parsed arguments to `tools.call`:

```ts
const result = await tools.call(call.name, call.arguments, { signal });
// result.output  -> send this back to the chat model as the tool's result (JSON, small, no binary data)
// result.plan    -> the DeckPlan, when the tool made or changed one (for your own UI)
// result.deck    -> the DeckSpec, when the tool made or changed one
// result.pptx    -> the file bytes, when one was rendered (offer it as a download)
```

`call` does not throw for expected failures. A missing id, an unknown tool or a model failure comes back as
`{ output: { error: "..." } }`, in words the chat model can act on. It throws only when `signal` was aborted.

## Arguments and results of each tool

### `plan_presentation`

Arguments:

| Name | Type | Required | Meaning |
| --- | --- | --- | --- |
| `topic` | string | yes | What the presentation is about, with every fact, figure and name the person gave. Numbers not stated here will not appear on slides. |
| `slide_count` | integer 1 to 30 | no | Number of slides, when the person asked for one. |
| `language` | string | no | BCP 47 tag, such as `he` or `en`. Defaults to the template's language. |
| `image_density` | `none`, `low`, `medium`, `high` | no | How many slides carry a photo. Default `medium`. Forced to `none` when no photos can be made. |
| `text_size` | `normal`, `large`, `x-large` | no | Size of body text. Default `normal`. |
| `instructions` | string | no | Audience, tone and anything else to respect. |

Output:

```json
{
  "plan_id": "plan_dc7cfffab9",
  "title": "Adopting a Rescue Dog",
  "language": "en",
  "slides": [
    {
      "number": 1,
      "layout": "cover",
      "layout_description": "Opening slide with the presentation's title and a subtitle.",
      "purpose": "Introduce the guide.",
      "thumbnail": "/static/thumbnails/cover.png"
    }
  ],
  "next": "Show this plan to the user. Use edit_presentation_plan for changes and create_presentation once they approve."
}
```

`thumbnail` is present only when the `thumbnail` option returns an address.

### `edit_presentation_plan`

Arguments: `plan_id` (string, required), `instruction` (string, required: what to change, in the person's words,
with any new facts they gave).

Output: the same shape as `plan_presentation` with a new `plan_id`, plus:

```json
{ "summary": "Added a table slide comparing three options before the closing slide.", "previous_plan_id": "plan_dc7cfffab9" }
```

### `create_presentation`

Arguments: `plan_id` (string, required: the approved plan).

Output:

```json
{
  "presentation_id": "deck_851fbb9677",
  "title": "Adopting a Rescue Dog",
  "file": "pptx/deck_851fbb9677.pptx",
  "slides": [{ "number": 1, "layout": "cover", "title": "Welcome Home Guide" }],
  "plan_id": "plan_7e890671a8",
  "warnings": ["slide 4 body: text is 230 characters; the slot fits about 211"]
}
```

`file` is the store key of the rendered pptx. `warnings` is present only when there are any.

### `edit_presentation`

Arguments: `presentation_id` (string, required), `instruction` (string, required).

Output: the same shape as `create_presentation` with a new `presentation_id`, plus:

```json
{ "summary": "Shortened the body text on slide 2 and made the title punchier.", "rewritten_slides": [2], "previous_presentation_id": "deck_851fbb9677" }
```

### Errors

```json
{ "error": "There is no plan with id \"plan_x\"" }
```

Other messages: `"topic" is required`, `There is no presentation with id "..."`, `Unknown tool "..."`, and the
messages of `PlanError` and `DeckValidationError` (see [generating.md](generating.md)).

## The store

The tools keep everything in a `BlobStore`, so the chat model only ever handles ids and short summaries.

```ts
interface BlobStore {
  get(key: string): Promise<Uint8Array | undefined>;
  set(key: string, data: Uint8Array): Promise<void>;
}
```

| Key | Content | Typical size |
| --- | --- | --- |
| `plan/<id>.json` | A plan: the brief and the slide list | 1 KB |
| `deck/<id>.json` | A deck: all text, and for each picture its id and description | 4 KB |
| `asset/<id>` | One generated picture | tens of KB to 1 MB |
| `pptx/<id>.pptx` | The rendered file | hundreds of KB and up |

`storeKeys.plan(id)`, `storeKeys.deck(id)`, `storeKeys.asset(id)` and `storeKeys.pptx(id)` build these keys, so
your app can read what the tools saved:

```ts
const pptx = await store.get(storeKeys.pptx(presentationId)); // to serve a download later
```

Why pictures are kept apart: the deck JSON stays a few kilobytes of text, which is all a model ever sees, and a
pptx can be rebuilt from the JSON and the pictures at any time. Icons and supplied assets are not stored, since
your app already has them.

Every change makes a new plan or deck under a new id. Earlier versions stay, so "undo" is just using the previous
id.

### Stores that come with the package

```ts
const store = memoryStore();                       // in memory: tests and short-lived processes; store.keys() lists keys
const store = directoryStore('/mnt/presentations'); // a folder, such as a mounted volume; keys become file paths under it
```

### Writing your own

Any object with `get` and `set` works: a bucket, a database table, a key-value service.

```ts
const store: BlobStore = {
  async get(key) {
    const row = await db.query('select data from blobs where key = $1', [key]);
    return row ? new Uint8Array(row.data) : undefined;
  },
  async set(key, data) {
    await db.query('insert into blobs (key, data) values ($1, $2) on conflict (key) do update set data = $2', [key, Buffer.from(data)]);
  },
};
```

`get` must resolve with `undefined` for a key that does not exist, not throw.

## Showing a plan visually

A plan lists a layout per slide. To show layouts as pictures, render the template's slides to images once (with
any tool you have), name each picture after the pattern made from that slide, and return its address from the
`thumbnail` option. The address is passed through to the chat model's output and is not fetched by slotdeck.

## A complete loop

```ts
async function runTurn(messages) {
  let reply = await chatModel.send({ messages, tools: declarations });
  while (reply.toolCalls?.length) {
    for (const call of reply.toolCalls) {
      const result = await tools.call(call.name, call.arguments);
      if (result.pptx) offerDownload(result.output.presentation_id, result.pptx);
      messages.push({ role: 'tool', toolCallId: call.id, content: JSON.stringify(result.output) });
    }
    reply = await chatModel.send({ messages, tools: declarations });
  }
  return reply.text;
}
```

`chatModel`, `offerDownload` and the message shape are your application's; only `tools.definitions` and
`tools.call` come from slotdeck.

## Guidance for the chat model's system prompt

The tool descriptions already tell the model the order to work in. It helps to add, in your own system prompt:

- Put every fact, figure and name the person gave into `topic`. Numbers that are not in `topic` will not be used.
- Show the plan and wait for approval before calling `create_presentation`.
- For a change to a finished presentation, call `edit_presentation` with the person's words; do not plan again
  from scratch.
- Refer to presentations and plans by id only. Never ask for or produce deck JSON.
