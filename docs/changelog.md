# Changelog

What changed in each version, newest first, with what to do when upgrading. The version you have installed is
in `node_modules/slotdeck/package.json`.

## 0.3.1

One step from a draft to a manifest that can be used. A studio built on 0.3.0 could show a person every
suggestion and every issue, thirty and more for one template, and nobody got to the end of the list.

### Added

- **`prepareManifest(template, { icons })`** prepares a manifest in one call and without a model. It takes
  every suggestion that is a fact of the slide (`icon`, `row`, `row-max`), lets rows widen, adds the mirrored
  and text-only variants and settles their rows, keeps the template's own pictograms where the app has no
  icon set (`kept` lists those slots, to give back later), sets the measured budget where a budget lets text
  run over its box, and takes out a pattern that cannot be used. It returns the manifest, whether it is fit to
  use (`sound`), what was done (`done`), the few sentences that are left for a person (`left`), and a value
  for each deck field taken from the template's own text (`fields`). It renames nothing, and a second run
  changes nothing. `variants: []` and `widen: false` make it take less. With `remove: false` it takes nothing
  away that a stored deck depends on, which is how to run it on a manifest that is in use. See
  "prepareManifest" in [template-studio.md](template-studio.md).
- **[prompts/one-step-preparation.md](prompts/one-step-preparation.md)**: a short prompt for a coding agent, to
  change what an upload leads to in a studio: prepare, a summary and a sample to look at, publish.

### Changed

- **[prompts/template-studio.md](prompts/template-studio.md)** now starts with that path and asks for it to
  be built first. Its prepare operation is `draftManifest` followed by `prepareManifest`, and answers with a
  summary instead of lists; going through suggestions one by one has become "looking closer".

### Fixed

- **A row is no longer made of items that hold a slot the row does not have.** An item takes everything in it
  along: with fewer items its shapes are gone, and a slide with fewer items than drawn could not be rendered
  when a card held, say, a pictogram that was labelled a decoration, or an icon named `icon1` beside a text
  named `step1`. A draft, `setRepeat` without numbers and the `row` suggestion now make such a slot part of
  the row when every item has one under its own number. A slot only one item has holds the row to no fewer
  items than that one, and in the last item it means there is no row. A text of the whole deck inside an
  item counts the same way. `prepareManifest` repairs a manifest that already has such a row.
- **`row-max` is no longer suggested for a row whose last item holds a slot that is not the row's** and that
  a deck fills. A further item is a copy of the last, and would show the template's own sample in that place.
- **A shape that takes its place from the slide's layout is not moved.** The usual title of a real template is
  a placeholder with no position of its own on the slide. `derivePattern` gave it a box when mirroring, and
  the variant could not be rendered. It now refuses with the reason, `suggestChanges` no longer lists such a
  variant, and the text-only variant leaves that shape where the layout puts it.
- **[template-studio.md](template-studio.md) no longer says to draft one slide again** with `slides: [n]`: the
  pattern it returns names the deck's repeated texts anew. Use `addPattern` and `addSlot`.

### Measured

On one 36-slide template, a draft left 30 suggestions for a person. `prepareManifest` took 21 of them in about
three seconds and left three sentences: 24 layouts became 41, with no name changed. A brief for five slides with
2, 3, 4, 5 and 6 columns then got the asked number on one line on 5 of 5 slides, as before, and with two or
three columns the cards now fill the row.

### Upgrading from 0.3.0

- Nothing changes for code that does not call `prepareManifest`, except what is described under "Fixed": a
  draft and `setRepeat` find rows by the new rule, and `derivePattern` refuses a mirror that could not be
  rendered. A stored manifest is not touched by the upgrade.
- A studio that shows the result of `suggestChanges` and `checkManifest` to a person can call
  `prepareManifest` after the draft and show its `left` first. An app whose chat side has no icon set passes
  `icons: false`.
- For a manifest that is in use, call it with `remove: false`. A manifest drafted with 0.3.0 can have a row
  that renders only with as many items as the template draws; this repairs it, and every deck that rendered
  still does.

## 0.3.0

A draft that can be used as it comes, and the studio's review step as a function. A row of cards now takes as
many as its slide has room for, a request for a number of columns gets them side by side, and a manifest that
is already in use can be repaired without a model and without renaming anything.

### Added

- **`suggestChanges(template, options?)`** reads a manifest against its pptx and lists what the slides allow
  that the manifest does not use yet, each with the changes for `changeManifest`: photo slots whose pictures
  are pictograms (`icon`), items named alike that are not a row (`row`), a row with room for more items
  (`row-max`), a row or grid whose items could widen (`row-widen`), and variants the slide allows (`variant`).
  No model. It works on any manifest, and every suggestion is tried against the pptx before it is listed.
  `SUGGESTION_KINDS` lists the kinds. See "suggestChanges" in [template-studio.md](template-studio.md).
- **`rows` on `draftManifest`**: `'fit'`, the default, lets a drafted row take as many items as fit side by
  side at their drawn width while the space between them stays at a third of the drawn space or more, and sets
  the row's `span` for it. `'drawn'` keeps rows at what the template draws. Any other value is refused.
- **`widen` on `draftManifest`** gives drafted rows and grids a `widen` factor, wherever some number of items
  they take leaves a line short of what the template draws. Off by default. A value that is not a number above
  1 and at most 4 is refused.
- **`notes` and `suggestions` on the result of `draftManifest`**: what the code settled by measurement, in
  sentences, and what the slides allow beyond the draft.
- **A ceiling on each answer of a draft.** Every request of `draftManifest` now carries `maxOutputTokens`: for
  a slide, 400 tokens plus 60 for each shape listed; for the deck, 300 tokens plus 40 and the size in bytes of
  each repeated text. It ends a model that runs away inside a text field.
- **`docs/prompts/`**: prompts to hand to a coding agent, shipped in the package like the rest of the
  documentation. The first, [prompts/template-studio.md](prompts/template-studio.md), is about the logic of a
  template studio built on these functions: what it stores, the order of its steps, what it must never do, and
  how to prove it works.

### Changed

- **A drafted row takes more than it draws.** With 0.2.0 a row drawn with four cards took two to four, so a
  deck that wanted five went to a table or a grid. It now takes what the slide has room for. Some rows still
  take what they draw: a grid; a row whose items touch, or differ in width by more than its spacing can absorb;
  and a row with a line, an arrow or a highlight among its items. Manifests that exist are not touched.
- **A row with something among its items holds exactly what it draws.** Where an arrow, a line or a highlight
  lies among the items of a row, fewer items would be laid out across it as well. Such a row is now drafted,
  and found by `setRepeat`, with its fewest and most both at the number drawn; 0.2.0 gave it two to that
  number. A rule under the whole row, or a panel behind all of it, is not among the items.
- **A small picture is an icon.** A picture with no side over 1.1 inch that is not a JPEG and not cropped into
  a frame is marked `small: a pictogram` for the drafting model, and becomes an icon slot even when the model
  answers `photo`. For a row this is decided for all its items together, and the slot keeps the name the model
  gave it. A model that does not see the slide was observed to call the pictograms of a row of cards
  photographs; the layout then asked for a photograph per card and was not offered when pictures were off.
- **A draft names layouts by their structure.** The drafting model is told to name and describe a layout by
  what it holds, never by the subject of the filler and without a number of items, and to name repeated items
  `card`, `column`, `step` or `stat`. A layout described as "four team member profiles" was passed over by the
  planner for every other subject.
- **The planner is told how a pattern sets its items.** The line it reads for a row or grid now says
  `items: 2 to 6 cards side by side in one row, numbered N = 1, 2, ...` or
  `items: 2 to 6 cards in lines of up to 3, numbered N = 1, 2, ...`; a row that takes one number of items gets
  a line too, `items: always 3 cards side by side in one row, numbered N = 1, ..., 3`; and a grid that never
  takes more than one line holds reads as a row. A new rule, in `planOutline`, `editPlan` and `planEdit`, tells
  it to answer a wish for a number of columns, or for items side by side, with a pattern that takes that number
  and sets it on one line, and to use the closest when there is none.
- **A row with no number of items given** is planned with the number it draws or, when it does not take that
  number, the nearest it takes. Before, a row that always takes two and draws three was planned with three and
  then refused. `sampleDeck` does the same by default, so its sample is a deck the manifest allows. An edit
  still copies a stored slide that states no number as it is.
- **A row that takes one number says so**: `pattern "x" takes exactly 2 cards, not 3`, from `validateDeck` and
  the planner, where it said `takes 2 to 2 cards`; `describeManifest` says `takes exactly 2` too.
- **`setRepeat` is stricter about what cannot render.** The room it counts for a row stops at the edge of a
  shape the row is drawn on, such as a panel. More items than drawn are refused for a row whose items are
  drawn overlapping, or differ in width so much that copies of the last, set out by the step from the first to
  the second, would lie on top of each other. A
  `span` is refused for a row whose items sit inside a group, where the renderer never used one, and such a
  span already in a manifest is no longer counted. The fit is counted when a change gives `max` or a `span`,
  or finds the row, and a refusal names fewer than the row draws only for a span narrower than the row. With
  a `span`, the message says how far the row may spread; `span: null` on a row that only fitted its most with
  the span lowers the most to what fits, and says so. A row whose item slots carry a `box` is laid out from
  where the slide draws them, so it is not found as a row and takes no more items or wider ones; `setSlot` with
  `box: null` on such a slot is now allowed, and takes the box away in every item.
- **A text-only variant leaves a row alone.** `derivePattern` with `text-only` widened each text beside the
  picture by itself, also one item of a row, whose copies then lay on top of each other. It now widens only
  text outside the row, and is refused where the row is all the text there is.

### Fixed

- `setRepeat` refused the number of items a row draws (`4 cards do not fit side by side`) when the items of a
  row inside a group sat closer together than 3% of their width, for example when only `widen` was being set.
  What the template draws fits in the width it is drawn in.

### Measured

One 36-slide template with two rows of four cards, drafted with `gemini-3.8-flash` and no slide pictures,
twice with 0.2.0 and three times with 0.3.0. Requests then asked for 2, 3, 4, 5 and 6 columns side by side,
with pictures off and a set of icons.

| | 0.2.0 | 0.3.0 |
| --- | --- | --- |
| Model calls for a draft | 37 | 37 |
| Patterns drafted, of 36 slides | 23 and 23 | 23, 23 and 24 |
| The two rows of cards take more than the four they draw | in neither draft | in every draft (up to 5 and up to 6) |
| The pictograms of those cards | photo slots, in both drafts | icon slots, in every draft |
| Those rows are described as | team or speaker profiles with photos, in both drafts | a row of columns, each with an icon, a heading and a text, in every draft |
| One brief for five slides, 2 to 6 columns: slides with the asked number on one line | 2 of 5, both on layouts for figures | 5 of 5, all on the rows of cards |
| The same brief on a 0.2.0 draft after `suggestChanges` (`icon`, `row`, `row-max`) and no new draft | | 5 of 5 |
| Each number asked for on its own, on a new draft and on two repaired 0.2.0 drafts | | 5 of 5 on each (two columns as one line of a two-column grid) |
| Sample decks at the fewest and at the most items | no validation problem | no validation problem |

What code cannot repair in a 0.2.0 draft is its descriptions: a row still described as team profiles is passed
over more easily for other subjects. Give it a description of its structure with `setPattern`.

### Upgrading from 0.2.0

- **Version range.** A range written for 0.2, such as `^0.2.0`, does not take 0.3.0. Change it to `^0.3.0`.
  A dependency pinned to `0.2.0` exactly needs the number changed.
- **Manifests you have keep working and do not change.** To let their rows take more, and to turn pictogram
  photo slots into icons, without a model and with every id unchanged:

  ```ts
  const template = loadTemplate(pptxBytes, savedManifest);
  const facts = suggestChanges(template, { kinds: ['icon', 'row', 'row-max'] });
  const { manifest } = changeManifest(template, facts.flatMap((s) => s.changes));
  ```

  Then give any pattern whose description names the filler's subject a description of its structure, with
  `{ op: 'setPattern', pattern, description }`.
- **Icon slots need an icon set.** An icon slot is filled from the `icons` given to `planDeck`, `generateDeck`
  or `presentationTools`. If your app gives none, follow the repair with
  `{ op: 'setSlot', pattern, slot, empty: 'keep' }` for each such row, which keeps the template's own
  pictograms: a pattern with a required icon slot and no icon set is not offered to the planner. The same holds
  for a new draft, which makes such slots icons.
- **Do not draft again to get this.** A new draft renames every pattern, and stored decks then point at ids
  that do not exist.
- **A new draft differs from a 0.2.0 draft**: rows take more, small pictures are icons. Pass `rows: 'drawn'`
  for rows as before. Render `sampleDeck(template, { items: 'max' })` to look at the rows at their most.
- **Code that builds a `DraftResult` by hand**, such as a test double, needs the two new fields, `notes` and
  `suggestions`.
- **Tests with a scripted model** that match the planner's prompt: the line
  `items: 1 to 5 cards, numbered` is now `items: 1 to 5 cards side by side in one row, numbered` (for a grid
  that can break a line, `in lines of up to 2`); a row that takes one number of items now has an
  `items: always ...` line; and the heading `Patterns for new slides` of `planEdit` is now
  `Patterns, for new slides and for a slide that changes its layout`.
- **Code that matches refusals**: `takes 2 to 2 cards, not 3` is now `takes exactly 2 cards, not 3`.
- **`setRepeat` may now refuse** a `span` on a row inside a group; more items than drawn on a row whose items
  are drawn overlapping or differ much in width; more items than fit inside a panel the row is drawn on, or in
  the drawn width of a row with an arrow, a line or a highlight among its items; and more or wider items for a
  row whose slots carry a `box`. Show `rejected[i].reason` as it is; `span: null` clears a span that no longer
  applies. `derivePattern` with `text-only` may now be refused for a layout whose text is a row.
- **`draftManifest` now rejects** a `variants` entry other than `mirror` and `text-only`, which 0.2.0 passed
  over. `without` needs a slot to be named and is made with `derivePattern`.
- **Your own `LlmProvider`**: draft requests now carry `maxOutputTokens`. Pass it on as the output limit, with
  room for reasoning on top if your model counts reasoning against it, as the Gemini adapter does.

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
