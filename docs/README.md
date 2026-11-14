# slotdeck documentation

This folder is the complete documentation of slotdeck. It is plain Markdown with no links to the internet, and
it ships inside the npm package, so it is available wherever the package is installed:

```
node_modules/slotdeck/docs/
```

It is written to be read by a person or by a coding assistant that has no network access. Every page stands on
its own, names real exports, and shows complete examples.

## What slotdeck is, in one minute

slotdeck makes new PowerPoint files from an existing PowerPoint template.

1. A **template** is an ordinary `.pptx` with sample slides.
2. A **manifest** (JSON) describes the template: which slides are reusable layouts (**patterns**) and which
   shapes on them hold content (**slots**).
3. A **deck** (JSON) lists the slides of a new presentation: for each slide, a pattern and a value for each of
   its slots.
4. `renderDeck(template, deck)` clones the pattern's slide for every deck slide, puts the values into the
   slots, and returns the bytes of a new `.pptx`. This step uses no model and no network.
5. Optionally, a language model writes the deck JSON from a brief (`generateDeck`), changes it by instruction
   (`editDeck`), or drafts the manifest for a new template (`draftManifest`). The model is something you plug
   in; slotdeck never calls anything you did not give it.

## Pages

| Page | Read it when you need to |
| --- | --- |
| [getting-started.md](getting-started.md) | Install the package and render a first deck, with and without a model |
| [concepts.md](concepts.md) | Understand templates, patterns, slots, manifests, decks, assets and layout modes |
| [deck-format.md](deck-format.md) | Write or check deck JSON: every field and every kind of slot value |
| [manifest-format.md](manifest-format.md) | Write or read a manifest: every property of patterns and slots |
| [rendering.md](rendering.md) | Call `loadTemplate`, `renderDeck`, `validateDeck`, `extractAssets`; understand warnings and errors |
| [generating.md](generating.md) | Make a deck from a brief, or change one by instruction, with a language model |
| [chat-tools.md](chat-tools.md) | Give a chat model the four presentation tools and store what they produce |
| [providers.md](providers.md) | Connect your own language model or image source, including an internal gateway |
| [layout-freedom.md](layout-freedom.md) | Vary the number of cards, enlarge text, derive several patterns from one slide, and find other layouts for a slide |
| [languages.md](languages.md) | Make right-to-left decks from a left-to-right template, or the reverse |
| [text-fit.md](text-fit.md) | Measure text with the template's fonts: overflow in lines, shrinking to fit, budgets from measurement |
| [preview.md](preview.md) | Get a picture of each slide without PowerPoint, and know how close it is |
| [design-checks.md](design-checks.md) | Find empty boxes, uneven cards and orphan words by measurement, and have drafts reviewed while a deck is written |
| [template-studio.md](template-studio.md) | Draft a manifest with a model, check it, change it by hand or by instruction |
| [api-reference.md](api-reference.md) | Look up any export: its signature and what it does |
| [troubleshooting.md](troubleshooting.md) | Find the meaning of an error or warning, and the known limits |
| [changelog.md](changelog.md) | See what changed in each version, and what to do when upgrading |

## Which function do I need?

| Goal | Function | Model calls |
| --- | --- | --- |
| Turn deck JSON into a pptx | `renderDeck` | none |
| Check deck JSON before rendering | `validateDeck` | none |
| See which other layouts a slide's content fits | `slideAlternatives` | none |
| Brief in, pptx out | `generateDeck` | text: 2 plus repairs; images: one per photo |
| Brief in, deck JSON out | `planDeck` | 2 plus repairs |
| Show a plan first, write later | `planOutline`, `editPlan`, `writeDeck` | 1 each, `writeDeck` plus repairs |
| Change an existing deck by instruction | `editDeck` (with pictures and pptx) or `planEdit` (JSON only) | 1 or 2 plus repairs |
| Let a chat model do all of this | `presentationTools` | as above |
| Describe a new template | `draftManifest` | 1 per slide, plus 1 |
| Verify a manifest against its pptx | `checkManifest` | none |
| See whether a manifest's budgets hold | `sampleDeck` then `renderDeck` | none |
| Know whether a text fits its box, in lines | `loadTemplate` with `fonts`, or `textMeasurer` | none |
| Set budgets from what the boxes hold | `measuredBudgets` | none |
| Show a deck or a template's layouts as pictures | `previewDeck`, `patternPreviews` with a `PreviewProvider` | none |
| List design problems of a deck | `checkDesign` | none |
| Have drafts checked for design while a deck is written | `review: designReview(template)` on the planner | none of its own; each repair is one writing call |
| Change a manifest by operations | `changeManifest` | none |
| Add a mirrored or text-only variant of a pattern | `changeManifest` with `derivePattern`; `patternVariants` lists what is possible | none |
| Change a manifest by instruction | `editManifest` | 1 or 2 |
| List the shapes of a template | `inspectTemplate`, `studioView` | none |

## Facts worth knowing before writing code

- The package is ESM only (`import`, not `require`) and needs Node 22.18 or later.
- Entry points: `slotdeck` (everything that runs anywhere) and `slotdeck/node` (`directoryStore` and
  `fontFiles`, which use the file system).
- All binary data is `Uint8Array`. In Node a `Buffer` is a `Uint8Array`, so `fs.readFileSync(path)` can be
  passed directly.
- Functions that call a model are `async`. `renderDeck`, `validateDeck`, `slideAlternatives`, `checkManifest`,
  `sampleDeck`, `changeManifest`, `patternVariants`, `inspectTemplate` and `studioView` are synchronous.
- A deck never contains binary data. Pictures are referenced by an **asset id** and passed to the renderer as
  a map from id to bytes.
- Nothing is written to disk and nothing is sent anywhere by slotdeck itself, except what your own provider
  objects do and what `directoryStore` writes to the folder you give it.
