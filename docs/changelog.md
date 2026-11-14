# Changelog

What changed in each version, newest first, with what to do when upgrading. The version you have installed is
in `node_modules/slotdeck/package.json`.

## 0.2.0

Two groups of additions. The first is about seeing and checking a deck: text measured with the template's
fonts, pictures of slides without PowerPoint, and design checks. The second is about more layouts from the
same template: variants the studio derives by itself, rows that fill their line, and other layouts for a
slide. Everything new is off until it is asked for.

### Added

- **Text is measured with the template's fonts.** Load a template with font metrics
  (`loadTemplate(pptx, manifest, { fonts })`) and text is broken into lines at each slot's typeface and size,
  the way PowerPoint sets it. Validation and `renderDeck` then warn by the lines a text runs over
  (`text takes 7 lines at 18 pt; the box holds 5 (2 lines too many). Text of about 290 characters fits.`), and
  the planner sends a draft back to the model with that sentence. On the demo template every line broke at
  the same word as in PowerPoint in all 1,153 text slots checked at the designed sizes, in English and in
  Hebrew, and 699 of 700 slots had PowerPoint's line count with every font size halved. See
  [text-fit.md](text-fit.md).
- **`fontkitMetrics(fontkit, files)`** builds the metrics from font files with the `fontkit` library, which
  you install yourself: slotdeck does not depend on it. `FontMetrics` is the small interface behind it, for
  using another library.
- **`templateFonts(template)`** lists the typefaces, weights and slants a template's text is set in, and
  **`fontFiles(fontkit, folder, typefaces)`** from `slotdeck/node` finds their files in a font folder.
- **Shrinking text to fit.** `shrinkText: { to: 0.85, minSize: 10 }` on the manifest lets text that does not
  fit be set smaller, in half points, down to the stated limit. `renderDeck` returns what it shrank in `shrunk`.
- **`textMeasurer(template)`** measures any text in any text slot: lines, capacity, overflow, the lines
  themselves. `validateDeck` takes it as the `measure` option.
- **`measuredBudgets(template)`** says how many characters each text slot's box really holds and suggests a
  budget, to set with `changeManifest`.
- **`checkManifest`** tries every text budget out when the template has fonts, and warns about a budget that
  lets text run over its box. It also names typefaces that were not supplied.
- **`sampleDeck(template, { fit: 'box' })`** fills every text slot to what its box holds by measurement.
- **Slide previews.** `PreviewProvider` is a new interface for a renderer that turns a pptx into one picture
  for each slide. slotdeck draws nothing itself and depends on no renderer. `pagusPreviews(core, renderer)` is
  an adapter for the Pagus library, which you install yourself; it gives SVG for a browser to show, with
  Hebrew and Arabic paragraphs marked right to left. `previewDeck` renders a deck and draws it,
  `patternPreviews` makes a thumbnail for every pattern of a template, and `presentationTools` takes a
  `preview` option: the pictures of every presentation it makes go to the store under `preview/<id>/` and to
  the app in `ToolResult.previews`. No renderer tried reproduces PowerPoint exactly; what was found, for
  English, Hebrew, tables and native charts, is in [preview.md](preview.md).
- **Design checks.** `checkDesign(template, deck)` lists problems of a deck that fits and still looks wrong,
  from measurement alone: a box left almost empty, a card far shorter than its neighbours, a heading that ends
  with one word on a line of its own. Every message says what length of text would put it right. See
  [design-checks.md](design-checks.md).
- **Reviews.** The planner takes a `review` option: a second look at every draft, whose findings go back to
  the writing model in the repair rounds. `designReview(template)` is the design checks as a review.
  `visionReview({ template, llm, preview })` shows a model pictures of the slides; it needs a preview provider
  that gives bitmap pictures, and it has been tested with a scripted model only, not run against a live one.
  Nothing reviews a draft unless `review` is set. `presentationTools` takes the same option.
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
- **Rows that fill their line.** `widen` on a `repeat` lets the items of a row or grid grow wider when a line
  holds fewer of them than the template draws, up to the factor given (`"widen": 1.5`), so that two or three
  cards fill a row drawn for four. Backgrounds and text boxes widen; icons, pictures and circles keep their
  size and their side of the item, also inside a group. The character budgets of the row's texts grow with
  the items, in validation, in what the planner shows the model, and in `sampleDeck`. Set it by hand or with
  `changeManifest(template, [{ op: 'setRepeat', pattern: '<id>', widen: 1.5 }])`; `editManifest` understands
  requests such as "let two cards fill the row". See "Fewer items that fill the row" in
  [layout-freedom.md](layout-freedom.md).
- **The demo template's `icon-cards` and `stat-cards` use `widen`**, so the example decks with three cards
  or four figures fill their rows.
- **`slideAlternatives(manifest, deck, slide)`** lists the other patterns that can hold what one slide of a deck
  says and returns the slide written for each: a row of cards as a list or as columns, a slide as its mirror
  image, a list as paragraphs. Code only: no model call, and no word is changed. A photo or an icon the new
  pattern has no place for is reported, and so is a text that is over its new budget. The chat tools do not
  offer it yet. See "Another layout for the same slide" in [layout-freedom.md](layout-freedom.md).

- **Measuring follows the layout.** With fonts, the text of a row that widens its items is measured in the
  box the renderer gives it at the slide's number of items, so validation, shrinking and the design checks
  agree with the slide: text that fits two wide cards is accepted there and reported in four narrow ones. On
  the demo template all 22 texts in widened cards and figures broke where PowerPoint broke them, in English
  and in Hebrew. `checkManifest` tries the larger budgets of such a row in the wider boxes, and
  `sampleDeck(template, { fit: 'box', items: 'min' })` fills them. A mirrored pattern, whose boxes only move
  their shapes, is checked for empty boxes against the template's own text like the pattern it mirrors.

### Changed

- `renderDeck` returns `shrunk` next to `pptx` and `warnings`. It is an empty array unless the manifest has
  `shrinkText` and the template has fonts.
- `loadTemplate` takes a third, optional argument. `Template` has an optional `fonts` field.
- `storeKeys` has a `preview` key builder. With a `preview` provider the outputs of `create_presentation`
  and `edit_presentation` gain `previews`, or `preview_error` when the pictures could not be made.
- **A slot with `"empty": "keep"` and a `box` now moves its shape.** Before, the `box` of such a slot was
  ignored unless a deck gave the slot a value. `checkManifest` no longer reports that such a slot points at a
  shape of another kind, or that it has no character budget.

### Upgrading from 0.1.2

- A dependency range written for 0.1, such as `"slotdeck": "^0.1.2"`, does not take 0.2.0. Install it by
  version: `npm install slotdeck@0.2.0`. The package itself has no new dependency; `fontkit` and the Pagus
  packages are installed only by an app that wants measuring or previews.
- Nothing has to change. With a manifest made before and a template loaded without fonts, every result is the
  same as in 0.1.2: character budgets, the same warnings, the same files. There is one exception: a manifest
  that already had a slot with both `"empty": "keep"` and a `box` now shows that shape at the box. Remove the
  `box` to keep the shape where the template drew it.
- To turn measuring on: `npm install fontkit`, build metrics with `fontkitMetrics`, and pass them to
  `loadTemplate`. Code that builds a `Template` object by hand (`{ pkg, manifest }`) keeps working; add
  `fonts` to it to measure.
- Once a template has fonts, a text longer than its `maxChars` that fits its box is no longer reported, and a
  text within `maxChars` that does not fit is. Code that looks for the words `characters; the slot fits` in a
  warning should also expect `lines`.
- Previews are off unless you pass a `preview` provider. To turn them on with Pagus:
  `npm install @pagus-kit/core @pagus-kit/renderer` and `pagusPreviews(pagusCore, pagusRenderer)`.
- Design checks and reviews are off unless you call `checkDesign` or pass `review`. A review can make the
  planner use more writing calls, since a draft within its limits may now be sent back.
- Budgets drafted without fonts are cautious. `measuredBudgets` usually suggests larger ones; setting them
  lets the planning model write more per slide.
- To get mirrored layouts from your own template, run
  `changeManifest(template, [{ op: 'derivePattern', pattern: '<id>', variant: 'mirror' }])` for each pattern with
  a picture beside its text, render `sampleDeck` to look at the result, and store the new manifest.
  `patternVariants(template, '<id>')` says which patterns allow it.
- A planner that is offered a pattern and its mirror image chooses between them by their descriptions. Give
  each a description that says where the picture is.
- Rows and grids keep the size of their items unless the manifest opts in with `widen`, and decks made from
  a manifest without it render byte for byte as before. After opting in, a slide with fewer items than drawn
  looks different (its items are wider) and its text budgets are larger; text that fitted before still fits.
  Render `sampleDeck(template, { items: 'min' })` to look at it before using the manifest.

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
