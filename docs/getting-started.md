# Getting started

## Install

```bash
npm install slotdeck
```

- Node 22.18 or later.
- ESM only. Your `package.json` needs `"type": "module"`, or your files need the `.mjs` extension (or you use
  TypeScript with ESM output).
- Two runtime dependencies come with it: `fflate` (zip) and `@xmldom/xmldom` (XML). There is no native code
  and no office program involved.

## What you need

| Thing | What it is | Where it comes from |
| --- | --- | --- |
| Template | A `.pptx` file with sample slides | Your designer, or your organization's standard deck |
| Manifest | JSON describing the template | Written once per template: by `draftManifest` (see [template-studio.md](template-studio.md)) or by hand (see [manifest-format.md](manifest-format.md)) |
| Deck | JSON describing one new presentation | Written by your code, or by a model through `generateDeck` |
| Assets | Picture bytes by id | Your icon set, your own pictures, or an image source you plug in |

## 1. Render a deck from JSON (no model)

```ts
import { readFileSync, writeFileSync } from 'node:fs';
import { loadTemplate, renderDeck } from 'slotdeck';
import type { DeckSpec, TemplateManifest } from 'slotdeck';

const manifest = JSON.parse(readFileSync('template.manifest.json', 'utf8')) as TemplateManifest;
const template = loadTemplate(readFileSync('template.pptx'), manifest); // do this once and keep it

const deck: DeckSpec = {
  version: 1,
  title: 'Quarterly review',
  fields: { organizationDepartment: 'Acme  |  Operations' },
  slides: [
    { pattern: 'cover', slots: { title: 'Quarterly review', subtitle: 'What changed since March' } },
    {
      pattern: 'text-and-photo',
      slots: {
        title: 'Deliveries are faster',
        body: ['Median delivery time fell from four days to three.', 'The new depot explains most of it.'],
        photo: { asset: 'depot', alt: 'The new depot' },
      },
    },
    { pattern: 'closing', slots: { title: 'Thank you' } },
  ],
};

const { pptx, warnings } = renderDeck(template, deck, {
  assets: { depot: readFileSync('depot.jpg') },
  layout: 'adaptive', // optional slots that the deck leaves out are removed from the slide
});
for (const warning of warnings) console.log(`slide ${warning.slide} ${warning.slot ?? ''}: ${warning.message}`);
writeFileSync('review.pptx', pptx);
```

The pattern ids (`cover`, `text-and-photo`, `closing`) and slot names (`title`, `body`, `photo`) are whatever
the manifest defines. To list them:

```ts
for (const pattern of manifest.patterns) {
  console.log(pattern.id, '-', pattern.description);
  for (const slot of pattern.slots) console.log('   ', slot.name, slot.kind, slot.empty ?? 'required');
}
```

If the deck does not fit the manifest (unknown pattern, a required slot without a value, a picture whose
asset was not passed), `renderDeck` throws a `DeckValidationError` whose `issues` array says what is wrong. Call
`validateDeck` first if you want the list without an exception. See [rendering.md](rendering.md).

## 2. Generate a deck with a language model

```ts
import { generateDeck, loadTemplate } from 'slotdeck';
import type { LlmProvider } from 'slotdeck';

// Any model works: the only requirement is this one method. See providers.md.
const llm: LlmProvider = {
  async generateJson({ system, prompt, schema }) {
    return await myModel.completeAsJson({ system, prompt, schema }); // must resolve with a parsed JSON value
  },
};

const { pptx, deck, warnings, outline } = await generateDeck(
  template,
  {
    prompt: 'Quarterly review for the operations team: deliveries got faster, returns fell, two depots opened.',
    slideCount: 6,
    imageDensity: 'none', // no image source was given, so plan without photos
    fields: { organizationDepartment: 'Acme  |  Operations' },
  },
  { llm, layout: 'adaptive' },
);
```

Keep the returned `deck`: it is what you store, show, edit and render again. See [generating.md](generating.md).

With the included Gemini adapter instead of your own provider:

```ts
import { gemini, geminiImages } from 'slotdeck';

const llm = gemini({ model: 'gemini-3.8-flash', apiKey: process.env.GEMINI_API_KEY });
const images = geminiImages({ model: 'gemini-3.1-flash-image', apiKey: process.env.GEMINI_API_KEY });
```

## 3. Change a deck by instruction

```ts
import { editDeck } from 'slotdeck';

const edited = await editDeck(template, deck, 'Shorten slide 2 and add a slide about returns before the end', { llm });
// edited.deck, edited.pptx, edited.summary, edited.rewritten
```

## 4. Describe a new template

```ts
import { checkManifest, draftManifest, loadTemplate, renderDeck, sampleDeck } from 'slotdeck';

const { manifest, skipped, issues } = await draftManifest(readFileSync('new-template.pptx'), { llm, name: 'new-template' });
const draft = loadTemplate(readFileSync('new-template.pptx'), manifest);
console.log(checkManifest(draft)); // [] when the manifest matches the file

// A deck that uses every pattern once with every slot full: render it and look at it.
const { deck: sample, assets } = sampleDeck(draft);
writeFileSync('sample.pptx', renderDeck(draft, sample, { assets }).pptx);
```

A drafted manifest is a starting point that a person reviews. See [template-studio.md](template-studio.md).

## Where to go next

- The two JSON formats: [deck-format.md](deck-format.md), [manifest-format.md](manifest-format.md)
- Using it inside a chat assistant: [chat-tools.md](chat-tools.md)
- Connecting an internal model gateway: [providers.md](providers.md)
