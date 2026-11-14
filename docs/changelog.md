# Changelog

What changed in each version, newest first, with what to do when upgrading. The version you have installed is
in `node_modules/slotdeck/package.json`.

## Unreleased

### Added

- **Variants the studio derives by itself.** A new `changeManifest` operation,
  `{ op: 'derivePattern', pattern, variant }`, makes a second pattern from a pattern's slide and works out every
  position and budget from the shapes of the slide. Variants: `mirror` (the picture and what sits beside it
  change sides, including a photo that runs off the edge of the slide), `text-only` (the picture is left out
  and the text widens over the room it leaves) and `without` (an optional text named in `slot` is left out and
  what sits under it moves up). A variant the slide does not allow is refused with the reason. See "Variants
  the code works out" in [layout-freedom.md](layout-freedom.md) and `derivePattern` in
  [template-studio.md](template-studio.md).
- **`editManifest` understands requests for a variant**, such as "add a mirrored version of the photo slides".
  The model names the variant; it still never sets a position.
- **`patternVariants(template, pattern)`** lists the variants a pattern allows, with the reason for each one it
  does not, for a studio screen that offers them as buttons. `VARIANTS` is the list of variant names.
- **`variants` on `draftManifest`** adds the mirror image and the text-only version of every drafted pattern
  that allows one. Off by default: see "Variants in a draft" in [template-studio.md](template-studio.md).
- **A slot can show a part of the design in another place.** A slot with `"empty": "keep"` and a `box` keeps
  its shape's content and puts the shape at the box, and may point at any shape that is not inside a group,
  a group as a whole included. `derivePattern` adds such slots (`moved1`, `moved2`, ...) for the shapes of a
  variant that are not slots. See [manifest-format.md](manifest-format.md).
- **The demo template has mirrored photo patterns**: `text-and-photo-mirrored`, `photo-stat-mirrored` and
  `list-with-photo-mirrored`, made with the new operation.

### Changed

- **A slot with `"empty": "keep"` and a `box` now moves its shape.** Before, the `box` of such a slot was
  ignored unless a deck gave the slot a value. `checkManifest` no longer reports that such a slot points at a
  shape of another kind, or that it has no character budget.

### Upgrading

- Nothing has to change. Decks and manifests made before render as they did, with one exception: a manifest
  that already had a slot with both `"empty": "keep"` and a `box` now shows that shape at the box. Remove the
  `box` to keep the shape where the template drew it.
- To get mirrored layouts from your own template, run
  `changeManifest(template, [{ op: 'derivePattern', pattern: '<id>', variant: 'mirror' }])` for each pattern with
  a picture beside its text, render `sampleDeck` to look at the result, and store the new manifest.
  `patternVariants(template, '<id>')` says which patterns allow it.
- A planner that is offered a pattern and its mirror image chooses between them by their descriptions. Give
  each a description that says where the picture is.

## 0.1.2

### Added

- **Text-only decks.** `textOnly: true` on a brief, or `text_only: true` in the `plan_presentation` tool, plans
  every slide as plain text: a title with paragraphs or a list, and no pictures, icons, cards, figures, tables
  or charts. The planner is then offered only the template's plain text patterns, so no other kind can be
  chosen. See "Asking for a kind of slide" in [generating.md](generating.md).
- **Wishes about the kind of slides are honoured.** A wish stated in words in the brief or in an edit
  instruction ("text only", "no tables", "only cards") now decides which patterns may be used and comes before
  variety. It works when planning, when changing a plan, and when changing a finished deck. The pattern list
  the model reads marks each plain text pattern with "Plain text.".
- **`isPlainText(pattern)`** tells whether a pattern is plain text, for an app that wants to show or check
  which layouts a text-only deck can use.
- **Larger body text for a whole template.** `textScale` on the manifest enlarges the body text of every deck
  made from the template (`1.3` turns 9 pt into about 12 pt). Budgets shrink to match in validation, in what
  the planner shows the model, and in `sampleDeck`. Set it by hand or with
  `changeManifest(template, [{ op: 'setManifest', textScale: 1.3 }])`; `editManifest` understands requests such
  as "make body text 20 percent larger". See "Larger text for a whole template" in
  [layout-freedom.md](layout-freedom.md).
- **`maxOutputTokens` on requests to the model.** The planner passes a ceiling for each answer with every
  request it makes. See "Runaway answers" in [providers.md](providers.md).
- **`effort: 'low'` on planning requests.** Calls that choose and arrange rather than write (the outline, a
  change to a plan, the first step of an edit) ask for light reasoning.

### Changed

- **Table text and chart labels follow the text scale**, where the manifest states their size
  (`bodyText.size`, `headerText.size`, `text.size`). This applies to a deck's own text size (`textScale`,
  `textSize`, `text_size`) as well as to the new template scale, and a table cell then holds fewer characters.
  Before, tables and charts kept their size when a deck asked for larger text.
- **The Gemini adapter limits the output of every request.** The limit is the request's estimate plus 8,192
  tokens for reasoning, or 24,576 when the request has none, or the adapter's `maxOutputTokens` when set.
  Before, a request had no limit unless `maxOutputTokens` was set.
- **The Gemini adapter sends a request once more** when the answer was cut off at the limit, after a server
  error (500 to 504), and when the service's recitation filter blocks an ordinary answer. A bad key, a quota or
  a bad request is not repeated.
- **The Gemini adapter uses thinking level `low` for requests with `effort: 'low'`**, unless `thinkingLevel`
  is set on the adapter. Planning calls became several times faster: a plan takes about 5 seconds.
- The per-slide sentence of a plan is asked for as "a single sentence of at most 25 words".

### Fixed

- **A request for a text-only presentation gave cards and tables.** The planner was told to vary the layouts
  and nothing said that a wish about the kind of slides comes first; and "no pictures" was the most that
  reached it. Fixed by the two additions above.
- **A model call could hang for minutes.** With its default reasoning the model sometimes ran away inside one
  text field of a planning answer and wrote tens of thousands of tokens. Measured on a text-only plan: 6 of 24
  calls with default reasoning, none of 18 with light reasoning. Planning calls now use light reasoning, every
  call has an output limit, and a cut-off answer is asked for once more.

### Upgrading from 0.1.1

No code has to change. To benefit:

- **Chat assistants:** `plan_presentation` has a new optional argument, `text_only`. It is in
  `tools.definitions`, so it reaches the chat model when the definitions are passed as before. If your system
  prompt lists the tool's arguments, add it, and tell the model that "no pictures" (`image_density: "none"`)
  still allows cards, tables and charts.
- **Small body text:** if decks from your template have body text that is too small, add `"textScale": 1.3`
  (or another factor) to the manifest, render `sampleDeck` once to see that everything still fits, and keep the
  manifest. A manifest with `textScale` is ignored by 0.1.1 and earlier, which show the designed size.
- **Your own `LlmProvider`:** requests now carry two optional fields, `maxOutputTokens` and `effort`. A provider
  that ignores them keeps working. To get the protection and the speed, pass `maxOutputTokens` on as your
  model's output limit (adding room for reasoning tokens if your model counts them in the same limit), use
  light reasoning when `effort` is `'low'`, and retry once when an answer is cut off. See "Runaway answers" in
  [providers.md](providers.md).
- **Tables with larger text:** a deck that asks for larger text now enlarges table text too, where the
  manifest states its size, and gives cells a smaller budget. A deck stored with `textScale` renders with
  larger table text than before.

## 0.1.1

### Added

- **Documentation inside the package.** The `docs/` folder ships with the npm package
  (`node_modules/slotdeck/docs/`), in plain Markdown with no links to the internet, so it can be read where
  there is no network access. Start at [README.md](README.md).

### Changed

- **Unknown chart format.** `validateDeck` warns about a chart `format` other than `"number"` or `"percent"`
  (`chart format "x" is not known and is ignored`). Before, it was ignored silently.

### Upgrading from 0.1.0

Nothing to change.

## 0.1.0

First release.

- **Renderer:** `loadTemplate`, `renderDeck`, `validateDeck`, `extractAssets`. Text, pictures, tables and
  native charts; rows and grids with a variable number of items; larger text on request; derived patterns.
- **Planner and editor:** `generateDeck`, `planDeck`, `planOutline`, `editPlan`, `writeDeck`, `editDeck`,
  `planEdit`, `resolveAssets`, with any model through `LlmProvider` and `ImageProvider`.
- **Chat tools:** `presentationTools` with `plan_presentation`, `edit_presentation_plan`,
  `create_presentation` and `edit_presentation`, and stores (`memoryStore`, `directoryStore`).
- **Gemini adapters:** `gemini`, `geminiImages`, `vertexEndpoint`.
- **Reading direction:** right-to-left decks on left-to-right templates and the reverse.
- **Template studio:** `draftManifest`, `checkManifest`, `sampleDeck`, `changeManifest`, `editManifest`,
  `studioView`, `describeManifest`, `inspectTemplate`.
