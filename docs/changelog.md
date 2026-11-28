# Changelog

What changed in each version, newest first, with what to do when upgrading. The version you have installed is
in `node_modules/slotdeck/package.json`.

## 0.10.2

Two things that failed when a deck was rendered, with errors that did not say where. No part of a manifest or
a deck changes, and a deck that rendered before renders as it did.

### Fixed

- **A sample table with merged cells is a sample like any other.** A template whose sample table had a heading
  across two columns, or a label down two rows, passed every check and then failed when a deck used it:
  `sample tables with merged cells are not supported`, with no slide named. Its cells are now read for their
  styles, and the table that is built has none merged: one cell for each column of each row, as a deck's table
  has. Every cell keeps the fill, the borders and the room PowerPoint kept for it. A cell that a neighbour's span
  covers states no text style of its own, so it takes its text from the nearest cell of its row that does (the
  one that covers it, where it is merged along the row), and from the cell above it where its whole row is
  covered. Nothing has to be unmerged in the template. Free slides still draw a table of their own on a template
  whose only sample table has merged cells, as they did.
- **A `box` on a placeholder that states no place of its own.** A placeholder takes its place from its slide
  layout, and most titles of real templates are such placeholders. `changeManifest` (`setSlot` with a `box`)
  and a manifest written by hand could give one a `box`, and the render then failed with `a box can only be set
  for a shape that is not inside a group`, of a shape that was in no group. Such a shape is now given a place of
  its own at the box, as PowerPoint gives it one when a person moves it: it stays the placeholder it was, with
  its typeface, size and bullets from its layout. A shape inside a group is refused as before, and so is a
  group that states no place (its children are placed by it).
- **What a table or a box cannot be made of names its slide and its slot.** `slide 3 / table: table slots must
  point at a table`, where the message stood alone before. The messages themselves are the same.

### Known

- Budgets, the studio's view and the drafts of a template whose text is in such placeholders still read a
  placeholder without a place as a shape at the corner of the slide with no size (`inspectTemplate`): a budget
  drafted for it is the length of its sample. That is the next release.

## 0.10.1

Style-only mode on the themes and templates that ship with an office suite: eleven themes with an ordinary deck
on each, and six designed templates. Thirteen of the seventeen carry a deck; the fixes below are what they
showed. A template that carried free slides before 0.10.0 renders as it did.

### Fixed

- **The content of a slide of placeholders stands where its layout puts it.** A slide whose every text is a
  placeholder states no place of its own, and the room of free slides was the usual margins (6% to 94% of the
  slide) whatever the template drew there. It is now where the layout, or the master, puts the body: on a theme
  with a decoration down one side, or a paper card in the middle of the page, the blocks stand where the
  theme's own text does. The band the blocks are kept in starts and ends where that body does, where it was
  a guess from the title's place.
- **The room is not widened into what the slide layout or the master draws.** A free slide's room is widened
  at the side reading ends to match the margin where reading starts. It is no longer widened where the layout
  or the master draws something in that strip under the title. Shapes inside a group are read one by one: a
  group that spans the slide draws where its children are.
- **The larger room of a crowded slide starts with the band where something is drawn in between.** It started
  just under the title. Where the layout or the master draws something that starts or ends between that line
  and the band (a banner behind the title, the top of a panel the content stands on), it starts with the band:
  text set in that gap stood on the edge of the drawing.
- **Where placeholders are copied, running text is the body of the content slides.** The line under a cover's
  title is plain, small and has no bullet, and was taken for running text before the bulleted body was: every
  slide then had the look of a subtitle (capitals, the accent colour). Running text is now the text most of
  the template's content is set in, without its bullet, and a subtitle is the last choice.
- **A placeholder its layout draws a line round is no text style.** A boxed caption was copied with its box
  round every text.
- **`styleManifest`: a cover and a closing slide are title slides.** A first or last slide of content (a table,
  a list, headed columns) is no longer taken for one: a deck had to fill its table to end. A subtitle
  placeholder is the line under the title whatever its sample holds, such as a presenter's name over a date.

### Known

- A theme whose title stands under its content (a title at the foot of the slide) cannot carry free slides:
  `styleManifest` gives the reason. So does a photo album, which has no slide of content.

## 0.10.0

Style-only mode: a deck in the style of a file nobody prepared. `styleManifest` makes a manifest from a `.pptx`
by rules, with no model, and a deck takes the file's cover, its closing slide and, between them, slides the
software lays out in the file's typefaces, sizes and colours. And a file whose text is all in placeholders, as
most decks people make by hand are, can carry such slides.

### Added

- **`styleManifest(pptx, { name? })`**: a manifest from a file alone, in a fraction of a second. Every slide
  becomes a layout whose texts are slots named by what can be measured (a title, the line under it, a heading
  over its text, a big number, a list); footers and texts that repeat in the margins become deck fields; the
  first and the last slide become `cover` and `closing`; the frame of free slides is the slide whose title sits
  where most titles sit, with the least drawn under it. It returns the manifest, `patterns` (what a deck is
  planned from: the cover, `free`, the closing slide), `fields`, `notes` in sentences, and a `reason` where the
  file cannot carry slides the software lays out. See [style-only.md](style-only.md).
- **A placeholder can be the source of a free slide's text style.** A free slide copies the template's shapes,
  and a placeholder states almost nothing itself: its place, typeface, size and bullets come from the slide
  layout, the master and the master's text styles. It is now copied as a shape that states its own look, with
  the list style of every level, the insets and the anchor written into it, the nearest statement of each
  property kept. `manifest.free.styles` may name a placeholder, and `changeManifest` no longer refuses one.

### Changed

- **A template whose running text is all in placeholders can carry free slides.** Before, `validateDeck` said
  `the template's running text is in placeholders, which a free slide cannot copy`, and a planner was offered
  no free slide. Now the placeholders are copied. Running text taken from a bulleted body has no bullet, and a
  body set larger than seven tenths of the title's size stands in at under half that size. A template with a
  text box of its own to copy keeps the styles it always had: placeholders are taken only where there is
  nothing else, or where the manifest names one.

### Upgrading from 0.9.0

- Nothing changes for a template that carried free slides before: its styles, its stored decks and its
  rendered files are the same.
- A template that was refused for placeholders now offers free slides to a planner, and a deck on it may hold
  them. To keep it as it was, set `"free": false` in its manifest (`changeManifest` with
  `{ op: 'setManifest', free: false }`).
- An app that shows the sentence of that refusal, or tests for it: the sentence is no longer said.
- A prompt for a coding agent that adds style-only mode to an app is in
  [prompts/style-only.md](prompts/style-only.md).

## 0.9.0

Free slides use their room. A row of three cards, a few figures or a timeline under a title used to stand
at the template's own sizes over an empty lower half. Such a slide now sets its text larger, a sentence alone
is set as a statement, and a planner is told to give a low row its bottom line.

### Changed

- **A free slide whose blocks leave much of its room free sets its text larger.** Where the blocks take
  less than two thirds of the room's height, every text style is tried at up to one and a half times its
  size, the largest first, and the first that suits is taken: the blocks are set out as they were, no text is
  smaller than before, the texts grew together (none more than 1.3 times what another did), a heading of one
  line is still one line, and nothing that fitted is now over its box. The size of a text still follows the
  width of its block, so a row of many narrow blocks grows little or not at all, and the title stays the
  largest text. Cards, steps, figures, phases, milestones, options and notes grow so. A slide with a chart, a
  table or a picture does not, and neither does a slide of running text, lists or quotations alone. A slide is
  planned for a writer at the template's own sizes, so what a writer is told always fits. See "A slide that
  leaves room" in [free-slides.md](free-slides.md).
- **One note alone is a statement.** A free slide of one note and nothing else sets it at up to seven tenths
  of the title's size, and no more than twice the size of running text, in up to four lines, where it was a
  thin band of small text. A slot of such a note is told what those lines hold (163 characters on the demo
  template, where a bottom line is told 153).
- **Notes alone sit down their room**, as other blocks that are lower than it do: two fifths of what they
  leave is above them. They stood under the title.
- **Where a template's own room ends well short of the page, a low row sits down the room to the page's
  foot.** Blocks that are lower than their room are settled in the room the template's own layouts use. On a
  template whose layouts use the upper part of the page, that room ends a tenth of the slide's height or
  more above the foot of the larger room a full slide may take, and a low row stood over an empty lower
  third. Such a row is now settled in the room down to that foot. The blocks keep their size, and a template
  whose own room ends near the foot of the page (the demo template does) sets its slides as before.
- **A page that is one colour but for a glow in a corner is read by that colour.** The text of a free slide
  is chosen to read on the page of the slide it stands on. A page with a gradient was read by the gradient's
  first colour, so a white page with a green glow in one corner was a green page: the dates of a timeline
  were set in black, and a band of colour was mixed from the green. A gradient in which one colour stands for
  three fifths of its length or more is now read by that colour. Any other gradient is read by its first
  colour, as before.
- **A planner is told what a low row needs.** Where notes and figures, phases or milestones are on offer, the
  line that offers free slides ends: `A slide of one row of figures, phases or milestones alone leaves its
  lower half empty: give it a note under the row, with the point the row makes.`
- **A deck that asks for larger text (`textScale`) never gets smaller text than one that does not**, also
  now that a slide grows without the wish: such a slide grows from the size the wish gave it, and where the
  slide without the wish still has some text larger, that slide is made.

### Added

- **`arrange.grow: false`** on a free slide (and `grow` on `freeSlide`): the slide keeps the template's own
  sizes whatever room its blocks leave, as before 0.9.0.

### Upgrading from 0.8.0

- A stored deck renders again with larger text on its free slides of cards, steps, figures, phases,
  milestones, options and notes that leave room, and with a lone note as a statement. Nothing that fitted
  stops fitting. To keep a slide as it was, set `"arrange": { "grow": false }` on it.
- An app that pins the writer's or the planner's prompts in tests: the line that offers free slides has one
  more sentence where notes and figures, phases or milestones are on offer.
- A template whose own layouts keep their content in the upper part of the page: a stored deck renders again
  with its low rows lower on the slide. Nothing changes its size.
- A template whose free slides stand on a page with a gradient that is mostly one colour and starts in
  another: a stored deck renders again with the texts, the band of colour and the strong fills that read on
  the colour most of the page shows. `manifest.free.styles` still names a style where another is wanted.
- A prompt for a coding agent that does the upgrade in an app is in
  [prompts/upgrade-0.9.0.md](prompts/upgrade-0.9.0.md).

## 0.8.0

Slides that say more at a glance. A text can set the words that carry its point in bold. A free slide can
draw a process, a timeline and a comparison, where it had cards and lists for all three. And a slide can
stand on a band of the template's colour, so that a deck is not one white page after another. Nothing changes
for a deck that uses none of it: such a deck gives the file it gave.

### Added

- **Strong phrases.** A phrase written between two pairs of asterisks, `**like this**`, is set in bold, in a
  run of its own that is otherwise the paragraph's: in every text slot of every slide, a layout's and a free
  slide's. On a free slide a strong phrase with a number in it (`**34%**`) also takes the template's accent,
  the colour of its big numbers, where that reads on the slide and on its cards. The marks are no part of the
  text: every length, budget and check counts what the slide shows. In a cell of a table, a label of a chart
  and a picture's alternative text the marks are taken off and the words stay. See "Strong phrases" in
  [deck-format.md](deck-format.md).
- **`emphasis`** (on `PlanOptions` and the chat tools, `true` unless set): whether the writer is told of that
  mark. It is told to use it sparingly: the two or three words or the one figure that carry a point, at most
  one phrase in a paragraph, on some slides and not on all. `false` tells it plain text only, as before.
- **Three kinds of block that are drawn** on a free slide, in the template's colours and its own text styles,
  as native shapes a person can edit:
  - `phase`: a link of a process. Phases stand as a chain of arrows, each with its name on it and a short
    text under it.
  - `milestone`: a point on a timeline. A line with a dot for each, its date over the line, what happened
    under it.
  - `option`: one of the things a comparison sets side by side: a name and a few points. The option with a
    `badge` (a word or two, such as `Recommended`) is picked out: its ground is the template's main colour
    and its text white.

  Up to six phases or milestones stand on one line, and up to four options. In a deck read from the right the
  arrows point left and the first stands at the right. The planner offers them with cards, in every deck
  that is not text only. See "A process, a timeline and a comparison" in [free-slides.md](free-slides.md).
- **`arrange.tone`** on a free slide: `accent` sets its blocks on a band in the template's main colour, their
  text in a colour that reads on it (white, on most templates); `tint` on a pale band of that colour. For the
  slide or two of a deck that carry its point: key figures, a quotation, a conclusion. `freeSlide` takes it as
  `tone`. A plan may ask for it (`shows_on` in an outline's answer), and code keeps it rare: never two such
  slides in a row, and no more than one slide in four. See "A band of colour" in
  [free-slides.md](free-slides.md).

- **`iconFiles` and `assetFiles`** from `slotdeck/node`: an app's own icons and pictures read from a folder.
  `iconFiles('/app/assets/icons')` gives the list that `icons` takes, and
  `assetFiles(folder, 'photo' | 'graphic')` the list that `assets` takes. Every PNG, JPEG and GIF in the folder
  is a picture, known by its file name without the ending. What each one shows is read from the folder's
  `icons.json` or `assets.json` (an object of id and description, or a list of `{ id, description, file? }`),
  from a file named in `descriptions`, or from an object given there; a picture without a line is described
  by the words of its file name. `recursive: true` reads the folders inside. `onSkipped` is told of every file
  that is passed over: an SVG or a WebP, which are not read, a file named as a picture that is none, a line
  that describes no picture. Two pictures of one id are an `Error`. Until now an app read its folder with code
  of its own, and the guides did not say how.
- **A guide to an app's own pictures**, "Your own icons and pictures" in [generating.md](generating.md): the
  folder, the formats, how an icon's description decides whether it is chosen, and that the same lists are
  given again wherever a stored deck is rendered.

### Changed

- **`FREE_BLOCK_KINDS` has three more kinds**, after `step`: `phase`, `milestone`, `option`. Code that lists
  every kind (a `Record<FreeBlockKind, ...>`) needs an entry for each.
- **The line that offers free slides to a planner names what they can draw**: "a process as a chain of arrows,
  a timeline or a comparison of options" stands among the content no layout shows. A template that has a layout
  for one of these is still planned with its own layout.
- **A deck that holds `**` around a phrase shows that phrase in bold**, where it showed the asterisks. Asterisks
  that are no such pair (a footnote mark, `2 ** 3`, marks with a space inside) stay as written.

### Upgrading from 0.7.1

- Nothing to do for decks, manifests and stored plans.
- An app that does not want bold phrases in its decks sets `emphasis: false` on the chat tools or on the
  planning calls. Decks written before have none.
- An app that shows a slide's text itself (a preview, a search index) reads it with the marks: take pairs of
  `**` off, or show them as bold.
- An app that reads its icon folder with code of its own may keep it, or use `iconFiles` in its place: a
  folder with an `icons.json` of `{ id, description }` and a file `<id>.png` for each gives the same list.
- A prompt for a coding agent that does the upgrade in an app is in
  [prompts/upgrade-0.8.0.md](prompts/upgrade-0.8.0.md).

## 0.7.1

A text too long for one presentation is no longer a dead end. The plan tool used to tell the chat model to ask
the person to paste the part the presentation is about, though the person had given the whole text already.
Now it lists the text's parts, by heading or by first words, and the chat model names the ones wanted in a new
argument, `document_parts`. An app can also set how much text a presentation is made from.

### Added

- **`maxDocumentChars`** (on `PlanOptions` and the chat tools): the most characters of a new brief's document,
  a whole number from 1,000 to 30,000, and 12,000 (`MOST_DOCUMENT_CHARS`, about 1,800 words) unless set.
  `undefined` and `null` leave it out; any other value is a `RangeError`, from `presentationTools` when the
  tools are built, from every planning stage, and from `planEdit` and `editDeck`. Only `planOutline` holds a new brief to it, and so
  `planDeck`, `generateDeck` and the plan tool. The messages say the words it holds ("about 3,000 words" for
  20,000). More text makes a longer planning call and a longer writing call, and a deck made from a text has at
  most 30 slides, so more of it is set small or cut. See "Texts the person gave" in
  [chat-tools.md](chat-tools.md).
- **`document_parts`** on `plan_presentation`, with `documents: true`: the parts of a text too long to take
  whole that the presentation is made from, by the number or heading the tool listed (`"2"`, `"4-6"`,
  `"Risks"`), or by a heading the person named. Each text numbers its own parts; with several texts named, a
  number goes after the text's name (`"report.pdf: 2"`). A named text none of whose parts is named is taken
  whole. A text is cut at its headings, a long section in pieces of about a quarter of the limit, into about 40
  parts at most, and every part is listed. Each listed count holds 2 for the join, so parts whose counts add up
  to the limit fit. See "A text too long to take whole" in [chat-tools.md](chat-tools.md).
- **`document.parts_used`** in the plan tool's output, where parts were chosen: each part the plan is made from,
  as `"bike-share.md: 2. Riders"`. Where only some parts of a text were chosen, `next` begins
  `Only some parts of "bike-share.md" are in this plan, as document.parts_used lists: tell the user which.`
- **A second argument for `describeDocuments`**: the tools' options, or any object with their
  `maxDocumentChars`, so that a text is marked too long by the same limit as the tools. Where a listed text is
  too long, the line ends with one more sentence, which tells the chat model to name it all the same and use
  `document_parts`, and never to ask the person to paste a part. Left out, it judges by 12,000 as before.

### Changed

- **The error for a text too long lists its parts** instead of sending the person to paste. It was
  `The text is 15230 characters, and a presentation is made from at most 12000 (about 1,800 words). Ask the
  user to paste the part the presentation is about, and name that.` It now names the text, tells the chat model
  to call again with the parts in `document_parts`, and ends with the parts of every named text, one to a line:
  `1. Background (2175)`. New errors say when the parts chosen are still too long, when a part is not in the
  text, when a heading or a number could be more than one part, and when `document_parts` comes without
  `documents`. The error for a name the list does not have adds a hint where the name has a part after it
  (`"bike-share.md, Risks"`), or more listed names after it (`"report.pdf, notes.docx"`). The messages are in "Errors" in [chat-tools.md](chat-tools.md).
- **The later stages read a stored plan's document up to 30,000 characters**, whatever their own
  `maxDocumentChars`: `editPlan`, `writeDeck`, `composeLayouts`, `planTextFit` and `freeSlidesOff`, and so the
  tools that change a plan and create a presentation. A plan made under a larger limit is written by tools built
  with a smaller one.
- **No second planning call for chosen parts, or for a text of more than 12,000 characters**, that the planner
  put on no slide. A named text the planner puts on no slide is made again as instructions only where it is
  12,000 characters or fewer and no parts were chosen: chosen parts are what the presentation is about, and no
  request is that long. Otherwise `next` begins `The planner put none of the parts named on a slide. Ask the
  user what the presentation should carry, and plan again.`, with the text's name in place of
  `the parts named` where no parts were chosen.
- **A document's sentences are found faster on a long line with many stops**: the word before a stop is read
  back from the stop to the space before it, not from the start of the line. The passages are the same.

### Upgrading from 0.7.0

Nothing breaks. With no `maxDocumentChars` and no `document_parts`, every tool output, message and system-prompt
line for texts of 12,000 characters or fewer is as in 0.7.0. To get what 0.7.1 brings:

- Give `describeDocuments` the tools' options: `describeDocuments(list, toolOptions)`. Without them it marks a
  text too long by 12,000, whatever the tools take.
- If the app filters the arguments of tool calls or what the tools answer, let `document_parts` and
  `document.parts_used` through.
- Set `maxDocumentChars` where people give long texts and the app's deadline allows the longer calls.

Tests that pinned the old text of the too-long error see the new one, which names the text and lists its parts.

### Known limits

- A limit above 12,000 is not measured with a live model. A translated document is written whole in one call:
  try a text of the limit's length into another language within the app's deadline before raising it.
- A table, or a line with no sentence end and no space, longer than the limit is one part that cannot be cut: it
  can only be left out, or narrowed with the person.
- A part number holds for one text and one limit. A number from an earlier turn means another part if the app
  changes the limit or reads the same file into other text; a heading does not.

## 0.7.0

Hebrew decks read right, and mailed files are clean. Every change here is small, none needs the template's
fonts, and each one changes what a reader of a mailed deck sees today: charts that can read from the right,
a highlighted card that stays on the right option in an English deck, file properties that describe the deck
and not its template, no links, comments or broken animations left over from the template's samples, and text
a screen reader can use. The chat tools now carry their own system-prompt lines, and model calls ride out a
short rate limit instead of failing.

### Added

- **Charts that read right to left** (`chartDirection` on the manifest, for decks that read right to left, or on
  the deck). The first category of columns and lines stands on the right, with the value axis; horizontal bars
  grow from the right with their names on the right; the room around a free slide's chart is mirrored, and a
  legend of its own stands upright at the left, first series at the top (PowerPoint sets a legend under a chart
  from the left whatever the chart reads); the data behind it (Edit Data) is a sheet that reads right to left. A pie is drawn as it is. Off
  unless set: many Hebrew charts of time read from the left. See "Charts that read right to left" in
  [languages.md](languages.md). `changeManifest` sets it with `{ op: 'setManifest', chartDirection: 'rtl' }`.
- **An item drawn to stand out keeps its place in a deck that reads the other way** (`repeat.emphasis`). In an
  English deck from a Hebrew template, the green "recommended" card used to stay on the right while the first
  option's text moved left: the emphasis landed on the last option. With `emphasis`, the items change places
  with their looks. `suggestChanges` finds such a row from the slide's fills (kind `row-emphasis`) and
  `prepareManifest` takes it; on the right-to-left template it finds the one row that has it.
- **Fewer lines spread down their area** (`repeat.spread`). An agenda of three in a list drawn for six bunched in
  the middle beside a full-height divider; with `spread: 1.6` the space between lines grows up to 1.6 times,
  within the area they are drawn in.
- **Tables on a layout fill their frame** (`fill: true` on a table slot). A table with fewer rows than its sample
  keeps the sample's height, its rows in the sample's proportions, and its columns of numbers end on the right
  in a deck read from the left.
- **File properties** (`RenderOptions.properties`, set by `generateDeck`, `editDeck` and the chat tools unless
  their `properties` is `false`). The file has the deck's title, language, real number of slides and the
  version of slotdeck that made it (custom property `slotdeck`), and no longer shows the template's sample
  cover as its thumbnail in Outlook and Explorer. The template's other custom properties, such as a
  confidentiality label, stay. A title that is no text is not written. `embedDeck: true` keeps the deck (without its brief and document) in the file;
  `readEmbeddedDeck(pptx)` takes it out with its pictures, also after PowerPoint saved the file again. See
  "File properties" in [rendering.md](rendering.md).
- **`tools.instructions`**: the lines the chat model's system prompt needs for the tools, made from the same
  options as the tool definitions. Add them to the app's own prompt in place of lines written by hand; they
  change with an update as the tools do. See "The chat model's system prompt" in [chat-tools.md](chat-tools.md).
- **`delivery: 'read'`** (on `PlanOptions` and the chat tools) for decks that are mailed and read: full
  sentences, no closing slide of thanks or questions, and an agenda only in decks of more than ten slides.
- **Retries** on `gemini()` and `geminiImages()`: `retries` (default 2) and `deadlineMs`. A rate limit that is no
  spending cap or daily quota, 408, 500, 502, 503, 504, the recitation filter, a network that failed and a
  timeout within a `deadlineMs` are tried again after a doubling wait with chance in it, or the wait the service
  asks for; one deadline covers a whole call. Pictures are retried too. `retryDelay(error, attempt)` is the rule, for a provider of your own.
- **Warnings a person should check** (planner only, with codes and Hebrew sentences, never sent back to the
  model): contact details and figures the brief does not give (`text-details-unstated`: an email or web
  address, a phone number, "24/7", a percentage), stand-ins such as "XX%" or "[name]" (`text-stand-in`), and
  text about the slide itself such as "this slide" or "בשקף זה" (`text-about-slide`). On the 14 stored decks of
  the right-to-left template they tell only the two invented "24/7" hotlines.
- **A word no line of its box holds** (`word-wider-than-line`): without fonts, `validateDeck` and `renderDeck`
  with a loaded template tell an address or a long number that PowerPoint can only break inside it.
- **`null` in a chart's values** is a value that is not known: a gap in the chart and an empty cell in its data,
  where it was drawn as 0. A chart written again in an edit keeps its gaps, unless the request gives the numbers
  (a writer is given numbers only, and code puts a gap it filled back).

### Changed

- **What a screen reader is told.** A chart's frame says its kind and its values ("Column chart of Revenue: Q1
  12, Q2 15"), a table's its headings and number of rows, both in Hebrew for a Hebrew deck; an icon that no one
  described is marked as decoration (Office's "Mark as decorative"). Before, a chart said "column chart:" and
  its series, a table nothing, and an icon was an unnamed picture.
- **A chart's text is marked with the language of its names.** In a deck that reads right to left, a chart whose
  categories or series have a Hebrew letter is marked `he-IL`, so PowerPoint places its brackets and punctuation
  as Hebrew; it was `en-US`. A chart in English is written as before.
- **A figure in a table cell** ("+12%") reads left to right in a deck that reads left to right from a template
  that reads right to left, as text boxes have since 0.6.5.
- **Every writer is told** never to state the organisation's own procedures, terms, prices, contacts, results or
  plans that the brief does not give, and never to write about the slide itself as slide text. In a live check of
  three thin briefs (six calls), no deck invented a contact, a hotline or a price, and decks set to be read were
  written in sentences with no thank-you slide; a table still gave its plans terms of its own ("Unlimited 45m").
  Review a deck made from a thin brief.

### Fixed

- **Cloned slides leave the template's sample leftovers behind.** A link on a sample run made the person's whole
  paragraph a link to the template's page; it is taken off the filled run. A slide keeps only its layout and
  what its own shapes point at, so the sample's comments and add-in tags go with their parts, and add-in data
  on the slide and its shapes is dropped. An animation of a shape or a paragraph the slide no longer has is
  taken out, and one with nothing left to animate goes whole: PowerPoint refused to open 0.6.9's file of a
  template whose animated optional text a deck left out. A link of the design, such as a logo that links to a
  website, stays.
- **With fonts, a text in columns is measured at its column's width less the space between columns**
  (`spcCol`), and the space before a first paragraph and after a last one is counted where the shape says
  PowerPoint draws it (`spcFirstLastPara`). A free slide's copy of a shape no longer keeps `spcFirstLastPara`,
  which its plan does not count. (These were planned as 0.6.10.)
- The package no longer ships source maps that point at sources it does not include: 2.8 MB unpacked, from 4.2.

### Upgrading from 0.6.9

Nothing breaks. To get what 0.7.0 brings, the app does a few things once:

- Add `tools.instructions` to the chat model's system prompt, and take out the lines about the tools that were
  written there by hand.
- In the studio, for a template whose Hebrew charts should read from the right:
  `changeManifest(template, [{ op: 'setManifest', chartDirection: 'rtl' }])`. Prepare the template again, or
  apply `suggestChanges(template, { kinds: ['row-emphasis'] })`, to mark rows with a card that stands out. Set
  `spread` on lists and `fill` on tables where wanted. [prompts/upgrade-0.7.0.md](prompts/upgrade-0.7.0.md)
  says all of it to the app's coding agent.
- For an app whose decks are mailed and read, set `delivery: 'read'` on the chat tools.
- Set `deadlineMs` on the Gemini adapters where a person waits for the answer.

A stored deck changes where it renders again: the alternative text of its charts, tables and icons, and the
language mark of a Hebrew chart. Its slides are otherwise the same file: on 14 stored decks of the right-to-left
template nothing else changed, and the demo template's decks change only in those words (the golden test lists
each). A deck renders with the template's file properties unless `properties` is given; the chat tools and
`generateDeck` give it. Tests that pinned a chart's `descr`, retried a 503 exactly once, or expected a picture
call to fail on the first 503, see the new behaviour. Decks are now told the new warnings above.

### Known limits

- The highlighted item is found only where exactly one item of three or more looks unlike the rest; two unlike
  items do not say which stands out, and a pair of columns (`sequences`) is not moved.
- Counting capitals and wide letters against character budgets without fonts was tried and left out: an
  ordinary sentence near its budget came out a few hundredths over, and the count is not checked against
  PowerPoint for that use.

## 0.6.9

On a free slide, spacing gives way before size. A text a little too long for its box was set half a point
smaller, while the space between its paragraphs and its line spacing never changed. Now that space closes up
first, and on a template that sets its lines wider than single, so do its lines: the text keeps more of its
size. On a right-to-left template whose lines are 112%, a document of ten parts under headings went from 14.5
to 16 pt.

### Changed

- **Running text and lists on a free slide are set larger where closer spacing lets them.** Only a text that fits
  at the template's own spacing is set so, at the largest size at which it fits with:
  - the space between its paragraphs closed up to no less than 0.7 of what it is (half the size, in whole
    points);
  - its lines, where the template sets them wider than single, closed up by up to a tenth, never closer than
    single, in whole percents (or whole points, where the template states points).

  The loosest spacing that fits is used, and like texts of a slide share it. A long text a person asked to
  have whole may take two columns where 0.6.8 took three, at a size no smaller; never more columns. In a deck
  that asks for larger text, how much of the wish a slide takes is chosen as 0.6.8 chose it; spacing then gives
  way within that, and no text is smaller than 0.6.8 set it. On 2,216 texts of free slides
  on the demo template (English and Hebrew), a right-to-left fixture and a right-to-left template:
  - 395 are set larger. On the right-to-left template that is 203 of 554, by 1 pt on a typical one and up to
    3.5. On the demo template a text gains half a point on a typical one, up to 2.
  - None is set smaller, and none is newly reported.
  - What a writer is told before writing is the same in every one of 372 plans.
  - In PowerPoint's own rendering of the 356 of them it could open, it drew exactly the spacing planned. Every
    text ended inside its box, or within the part of a line its last line may hang, except three in two
    columns not set dense, whose box is an estimate (see Known limits). They stayed inside their room, and two
    of them ran further past in 0.6.8.

  What is cut, reported or set whole is decided at the template's own spacing, as in 0.6.8 (the one judge of
  0.6.3 is not loosened). The renderer draws the spacing the plan counted: the space after each paragraph, and
  where lines close up, an explicit line spacing. A check with the template's fonts measures them as drawn.

### Fixed

- **A text of many short paragraphs is no longer set down the slide past its box.** A short text is set a
  little down its room; counted by its length, twelve one-line paragraphs were set so far down that their last
  two ran 38.5 pt under the box in PowerPoint, and the deck's check reported them. A written text is set no
  further down than its paragraphs, each in whole lines, leave free, with the part of a line its last line may
  hang under its box. A text the check did not report stays where it was.

### Upgrading from 0.6.8

Nothing to change in code. Decks without free slides render byte for byte as in 0.6.8, and so do 29 stored
decks of two right-to-left templates. A stored deck with free slides changes where:
- a running text or a list that fits is set larger with closer spacing;
- a text of many short paragraphs was set down past its box: it is set higher, and where it fits there it is
  no longer reported (one that does not is reported with the larger budget of its higher box).

A test that pins the size of a written free slide's text, or its `paragraphGap`, sees the larger size and the
closer gap. The budgets writers are told, and plans made before there is content, do not change, except that a
person's document that code sets as it is is planned at the size it is now set at.

### Known limits

- Text set closer reads tighter. The floors (0.7 of the space between paragraphs, single line spacing) were
  chosen where PowerPoint's rendering still kept paragraphs apart and letters clear; a template whose lines are
  single gains only from the space between paragraphs.
- The box of running text in two columns that is not set dense is an estimate, made at the template's own
  spacing. Its second column can run a little under that box, inside its room, as in 0.6.8: set larger, one
  Hebrew text on the demo template now runs 4.5 pt under it.
- In a deck that asks for larger text, the share of the wish is chosen at the template's own spacing. A running
  text there can come out a little smaller than the same text without the wish, where closer spacing set that
  one larger: never smaller than 0.6.8 set it with the wish.

## 0.6.8

A layout's text no longer runs past its box when it is written to the length the writer is told. A layout's
budget (`maxChars`, `maxParagraphs`) is counted once, for the template's own text, without the space the
template keeps around paragraphs. A text that fills it in many paragraphs takes that space many times: on a
right-to-left template, six points beside a photo, 480 characters as told, ran 19.5 pt past their box in
PowerPoint, and three paragraphs in each of two columns 17 pt. The check counts characters and paragraphs, so
it said nothing.

What a writer is told a layout's slot of paragraphs holds now counts that space, without fonts, from the
template itself. The check is unchanged. On a right-to-left template whose manifest states `characterWidth`
0.45, a writer is told, at the deck's usual text size:

| Slot | 0.6.7 told | 0.6.8 tells |
| --- | --- | --- |
| Text beside a photo | 715 characters in 3 paragraphs | 518 in 3 |
| Points beside a photo | 480 in 6 | 357 in 5 |
| Two columns, each | 550 in 3 | 428 in 3 |
| A page of text | 1,299 in 7 | 1,071 in 6 |
| Points on cards, each | 86 in 4 | 79 in 3 |

In all, 19% fewer characters on that template; 28.5% fewer where the manifest does not state its width. Each
paragraph is counted in whole lines, each line 85% full, so the pair holds a text of any paragraph lengths. On
17,109 texts PowerPoint set on three templates, none inside what it tells ran past its box. Stored decks render
as in 0.6.7.

### Changed

- **A writer is told what a layout's slot of paragraphs holds with its template's spacing counted.** For each
  slot of two paragraphs or more whose template keeps space before or after its paragraphs, or sets its lines
  further apart than 1.2 times their size, the count reads from the slide:
  - the box;
  - the lines it holds at the slot's size;
  - the space PowerPoint draws between paragraphs. With `spcFirstLastPara` it also draws the space before the
    first and after the last.

  A text of p paragraphs is held where it fits in the lines left, each paragraph ending on a line of its own.
  The writer is told one pair of characters and paragraphs:
  - Where the stored budget fits, nothing changes.
  - Where keeping the paragraphs costs few characters, they are kept with fewer characters: two columns, 550 in 3
    become 428 in 3.
  - Else it is told fewer, longer paragraphs: points beside a photo, 480 in 6 become 357 in 5.

  The pair is in the slot's limits, the outline's list of layouts, and the writer's `maxItems`. A slot that
  keeps no such space, a free slide's texts, a deck's fields and slots of one paragraph are told their stored
  budgets, as before.
- **A slot's description loses a length it states, where the pair is less.** "About 420 characters in all"
  would tell a writer more than the pair; the rest of the description stays.
- **An asked length is held by what a layout holds in the paragraphs it takes.** A slide asked for 200 words
  was held by nine tenths of a layout's limits, a margin for the spacing that was not counted. A slot the count
  speaks for now holds what it holds in some number of paragraphs. Its writer is told that number and what it
  holds: on the demo template, 45 words in a body of three short paragraphs are told 331 characters in two,
  where three hold 263. It is told no fewer than half the paragraphs of the pair, so a list stays a list: the
  demo's list beside a photo holds at most 306 characters of a length, in two items. A longer text is told
  that most, in those paragraphs, and is cut or moved as before. Other slots keep the nine tenths.
- **A document's text on a layout is counted the same way.** A text of known paragraphs is told what that many
  hold, or fewer that hold more: the writer joins them, as before, and the slide says so. A text code sets as it
  is, in a layout's one running slot, is held where its own paragraphs fit, each counted in its own whole
  lines. On the demo template, a lead line and three short items now stay in the list beside a photo, set as
  they are; 0.6.7 moved them to a free slide. A change that keeps such a slide's passages keeps its text as
  code set it: a new title does not hand it to a writer. A text that a writer writes, not code, is held to what
  the writer is told, also where a home is found for it: a slide asked shorter, a table of the text that goes
  into the layout's table, and in `planEdit` a text asked in other words, which goes to a free slide where its
  layout held it only as code set it. A layout of four paragraphs or more no longer tells a text of three or
  more one paragraph fewer where the count speaks for it: the count says how many.
- **A document is shared over slides by what a slide holds of a document's text.** On a template whose layout
  of plain text the count speaks for, the room the outline is told a slide holds of a document is the pair
  that layout is told: 263 characters on the demo template, where 0.6.7 said 297.

### Added

- **`PlanOptions.countSpacing`, and `countSpacing` of `presentationTools` and `planTextFit`.** On unless
  `false`. With `false`, every writer is told the slot's stored budget, as in 0.6.7.

### Upgrading from 0.6.7

Nothing to change in code. Stored decks render byte for byte as in 0.6.7, and their check is the same: a deck
written to 0.6.7's budgets gets no new warning.

Plans change on templates that keep space between paragraphs, for a layout's text, an asked length on a
layout, and a document's text on a layout. The numbers move both ways:
- A writer is told fewer characters, or fewer paragraphs, where the spacing costs room: most of a layout's
  slots of many paragraphs.
- An asked length and a document's text lose 0.6.7's shares of nine tenths and three quarters, so they may be
  told more characters, often in fewer paragraphs: on the demo template, 45 words are told up to 310
  characters in at most two paragraphs, where 0.6.7 told up to 290 in three.

A slide may then move to a free slide where 0.6.7 kept its layout, or keep a layout it left before. A test that
compares a plan's budgets, a writer's schema or `planTextFit`'s `holds` with stored copies sees the new numbers.
`countSpacing: false` gives 0.6.7's.

A template that states `characterWidth` loses less: on the right-to-left template above, 19% where it would
lose 28.5%. To set it, see "How wide a template's text is set" in [text-fit.md](text-fit.md) and
[prompts/character-width.md](prompts/character-width.md).

### Known limits

- The count holds any paragraph lengths, so a text of even paragraphs often has room left: in PowerPoint, the
  writer could have been told more. It counts lines 85% full, as free slides do.
- A slot of one paragraph is not counted: its stored budget stands. On a template whose cards hold one
  paragraph each, a card filled to its budget can still run past its box, as in 0.6.7.
- Measured with the template's fonts (`textMeasurer`, `shrinkText`), the space before the first paragraph and
  after the last that `spcFirstLastPara` draws is still not counted.

## 0.6.7

Running text on free slides is set larger where it fits. On a right-to-left template, a document's text under
six headings was set at 9.5 pt in three columns, the last a third full, where PowerPoint fits it at 12. Three
things held such text small:
- the count that sizes running text takes a character to be as wide as any face may set it;
- it counted each heading of a document on a line of its own, though 0.6.6 sets a heading at the head of its
  paragraph;
- a long text went down half a point at a time in two columns before it was allowed three.

A template can now say how wide its text is set, headings are counted as they are set, and three columns come
sooner. The count still keeps a margin: that text is now set at 10 pt, not at 12 (see Known limits). On 990
test slides on three templates, compared with 0.6.6:
- Nearly half the long texts with a document's headings are set larger: by 1 pt on a typical one of them, and
  up to 5.
- On a template that states its width, running text that is set larger is set 0.5 to 1.5 pt larger on a
  typical slide, up to 4, and up to 6.5 with headings. Most short texts not set dense do not change.
- Nothing is set smaller, unless a template states its width: wider than the estimate, it counts more
  cautiously; or, on a small template, a long text stays whole where 0.6.6 cut it (see Known limits).
- No text is reported over its budget that was not before.

In PowerPoint's own rendering of 134 texts on two templates, capitals, figures and mixed scripts among them,
nothing ran past its box that did not before.

### Added

- **`TemplateManifest.characterWidth`: how wide the template's text is set.** It is the width PowerPoint gives
  a character of the template's running text, as a share of its size, measured on the template's own face: the
  width that 19 full lines in 20 stay under, rounded up. Where it is stated, the running text of free slides
  (texts and lists, set dense or not) is counted with it, in place of 0.55, and 0.48 for text set dense. On the
  demo template at 0.44, six paragraphs of 1,400 characters went from 14.5 to 16 pt, and a text of 2,000 set
  dense from 13.5 to 14.5. On a right-to-left template at 0.45, three texts side by side went from 13.5 to 15.
  The budgets a writer is told grow with it. Other things are counted as before:
  - cards, steps, quotations, notes, headings, labels, tables and charts;
  - whether a long text takes one column or two;
  - whether a long text set dense needs the larger room, where it is set in columns. Counted narrower, a text
    would otherwise stay in the first room, smaller than in the larger one. Where the larger room holds such a
    text at no size and the first room holds it, it stays in the first: in columns it would be cut;
  - how far down the slide a short text not set dense starts. Counted narrower, a text of many short paragraphs
    was set so far down that it ran under its box;
  - how small texts beside each other may be set: four fifths of the top size. Counted narrower, the top is
    higher in a deck that asks for larger text, and three texts that need smaller were set over their budget.

  The value is between 0.3 and 0.8. A value under the face's own lets text run past its box: measure it, don't
  guess it. A value over the estimate (0.48 for a long text set dense, 0.55 for other running text) counts that
  text more cautiously, and sets it smaller: right for a wide face. To manage it:
  - `changeManifest` sets it (`{ op: 'setManifest', characterWidth: 0.45 }`), and `null` clears it.
  - `editManifest` takes it in words.
  - `checkManifest` warns of a value out of range, which is then not used.
- **`scripts/measure-character-width.ts` measures it.** It renders free slides of the template's kind of text,
  reads how PowerPoint breaks their lines (`scripts/measure-pptx.ps1`), and prints the width 19 lines in 20 stay
  under, rounded up. It is a development tool for a PC with PowerPoint, in the repository, not in the package.
  See "How wide a template's text is set" in [text-fit.md](text-fit.md).
- **`TemplateManifest.leastTextSize`: the least size, in points, of a long text set dense.**
  - Such a text is set no smaller, also where the first room would hold it at four fifths of the template's
    size. Where it needs smaller, it is cut, or the planner gives it more slides.
  - What a slide of one text holds (`freeTextRoom`, and with it the room the planner counts for a document) is
    counted at it: a slide set dense is said to hold no more than it holds at that size, and, where that is
    less than the shortest length the room is asked about (650 characters), what it holds.
  - Where it is over the size that reads (9 pt on a slide 13.33 inches wide), a writer of such a text is told
    what its box holds at the least size it is set at there, in the paragraphs it is told. Before there is
    content the text may be set larger, but what is written can hardly be set smaller: a writer told what the
    box holds in a few long paragraphs wrote it in many, and was cut.
  - A value over the template's own size of running text is taken as that size, and a value between half
    points as the next half point up (12.2 as 12.5): text is set in half points.
  - A value under 9 pt (on a slide 13.33 inches wide) also lowers the size under which the checks call a
    slide's text small.

  Left out, it is 9 pt on a slide 13.33 inches wide and less on a narrower one, by its width, as before. The
  value is between 5 and 24. `changeManifest` and `editManifest` set it, as above.

### Changed

- **A document's headings are counted as they are set.** On a free slide, a heading set run in takes no line of
  its own in the count that sizes the text:
  - its bold letters and the em space go to the head of its paragraph;
  - the space after it goes before that paragraph.

  The count asks the renderer's own question of each heading (`runsIn`), for a line a character shorter, so
  that a heading counted run in is always set run in. On the demo template, a text of twelve sections went from
  10.5 to 11.5 pt. Whether a long text set dense needs the larger room is still decided as 0.6.6 counted it,
  with its headings on lines of their own. Counted run in, a text of short parts fits the first room more often,
  and kept there it came out smaller: in one column at 14.5 pt where the larger room set it at 22 in two. While
  a deck is planned, a document's text that code sets as it is carries its headings into either room, so the
  plan counts it as the slide will be set.
- **With the template's fonts, a heading is measured as it is set.** The deck's check and the renderer's
  `shrinkText` measured each heading of a document on a line of its own (`FitOptions.bold`). A text the plan now
  sets larger, its headings counted run in, was then reported over its box, and set smaller with `shrinkText`:
  twelve short sections planned at 20.5 pt were measured as 24 lines in a box of 21, and set at 18. A heading
  over a paragraph of text is now measured run in at its head, in bold, an em space after it, where the renderer
  sets it so, asked for a line a character shorter as the plans ask. As the renderer, it counts a line of a text
  in columns without the space between them, and sets a heading on a line of its own where the slide states no
  size for its run or no place for its shape. A word that does not fit what the heading leaves of the line
  starts the next, and a heading larger than its paragraph makes its line as high as it is.
  `TextFit.paragraphs` and `TextFit.widths` have an empty entry for it: its words open the first line of the
  paragraph under it.
- **A deck that asks for larger text gets what 0.6.6 gave it.** The size each text of a free slide may not go
  under is counted as 0.6.6 counted it, its headings on lines of their own and a character no narrower than the
  estimate. Counted as now, a text with headings, or on a template that states its width, came out larger
  without the wish, and the wish then cost the rest of its slide half of it: a note set at 19.5 pt went to
  18.5, with a warning. Where the slide takes none of the wish, it is set as without it, counted as now.
- **A long text set dense takes three columns where two do not hold it**, at the same size, as long as a line
  in three columns still holds 45 characters as PowerPoint sets them. Before, it took three only where two
  would make lines longer than 95 characters. It went down half a point at a time in two columns, then jumped
  to three with the last column part empty. On a right-to-left template, a document's text of 2,000 characters
  under ten headings went from 10.5 to 11 pt, and to 11.5 with `characterWidth` 0.45. A text that fits at no
  size is set as before: at the least size, in the columns its lines take there.
- `setHeadings` counts a line with the template's `characterWidth` where it is wider than its own 0.55. The box
  of running text set in two columns (not dense) is as high as that width needs, where it is wider than 0.47.

### Upgrading from 0.6.6

Nothing to change in code. Stored decks render as in 0.6.6, with two exceptions, on every template:
- A slide set dense whose text two columns did not hold at some size, and three columns do, is now set larger,
  in three columns.
- A slide whose text has a document's headings (`SlideSpec.headings`, from 0.6.6) is counted with its headings
  run in, and is set at the same size or larger. With the template's fonts it is measured so too, and
  `shrinkText` sets it smaller less often.

Plans of such texts change the same way: a test that compares a plan's budgets or a rendered deck with a stored
copy sees larger sizes and budgets there. Code that reads `TextFit.paragraphs` or `TextFit.widths` of a text
measured with `bold` finds an empty entry for a heading set run in.

To have a template's text set larger, measure its width and set it once:

1. On a PC with PowerPoint, in a clone of the repository, run:
   `node scripts/measure-character-width.ts template.pptx manifest.json --text sample.txt`
   `sample.txt` holds paragraphs of the kind the decks hold, one per line, some 3,000 characters.
2. In the app, set the printed value: `changeManifest(template, [{ op: 'setManifest', characterWidth: 0.45 }])`.
3. Store the manifest it returns.

[prompts/character-width.md](prompts/character-width.md) is the prompt for a coding agent. A template without
the setting is counted as before.

### Known limits

- `characterWidth` is one width for all of a template's running text. A text much wider than the template's
  usual (capitals, long figures, words in a wider script) is counted as the usual one, and can run past its box
  where it is set larger. Measure it on text of the kind the decks hold.
- The count keeps its margins with the setting too: it counts lines 85% full, and keeps a line free in each
  column. In PowerPoint, a text set by it still has room for about 2 to 3 pt more: the document's text under
  six headings above is set at 10 pt where PowerPoint fits it at 12, with or without the setting. Templates
  loaded with fonts measure the layouts' text as before, and the setting does not change that.
- On a template that states its width, a long text set dense that the larger room holds at no size stays in
  the first room where that holds it whole. There it is set down to four fifths of the template's size, which
  on a small template is under the least size of the larger room: on a slide 10 inches wide, at a text size
  of 0.8, 6.5 pt where 0.6.6 set it at 7 pt and cut it. A text with headings can stay so where the same text
  without them, counted longer, goes to the larger room at 7 pt and is cut. On the demo template and a
  13.33-inch right-to-left template four fifths of the size is over the least size, and this does not arise.
- The plans and the measured check ask of a heading for a line a character shorter than the renderer does. A
  heading at that edge is counted on a line of its own where it is set run in: a line more, never fewer. On
  1,080 texts measured with fonts, 11 had one such heading, all on a template wider than the estimate.
- A text with headings that are not set run in, such as headings over list items or table rows, is counted
  with each heading on a line of its own, as before.
- A `leastTextSize` over 9 pt does not raise the size under which the checks call a slide's text small. That
  check looks at all of a slide's text, its captions and labels among it.

## 0.6.6

A person's document reads as a document on its slides. On a slide set dense, its headings stood in the text as
plain lines, indistinguishable from it ("Riders", "תוכנית"); one could end a column with its text in the next;
and the heading the slide opened under stood nowhere when the writer gave the slide no title. Stored decks, and
decks made without a document, render as in 0.6.5.

### Fixed

- **Headings in a document's text are set as headings.** Each is set in bold at the head of the paragraph
  under it, an em space between them: `Riders  Students and commuters make up most members…`. The space after
  the heading goes before the paragraph, so sections stay apart. A heading shares its line with its text, so
  none ends a column with its text in the next. The em space opens the paragraph's own text, in its language,
  so an English heading over Hebrew text (`KPI`) is not glued to it. A heading keeps a line of its own, in
  bold and without a bullet, over a list item (by its marker or by the bullet its paragraph draws), a table
  row or another heading, over a paragraph it would turn the other way (a Hebrew heading over an English
  paragraph of a Hebrew deck), and where its bold letters would take more than a line of the text, counted
  wider than any plan counts one.
- **The heading a slide opens under is its title.** Where code sets a slide's text from its passages, code
  sets its title too: the heading of its first passage, word for word, where the title holds it, on a free
  slide in one line. Slides that go on from a section get the section's heading. Before, the writer chose the
  title, and a slide without one lost the heading.
- **A text does not open with its title.** A first paragraph that says the title again, as a heading line
  with a colon (`רקע:` under `רקע`), is taken off; its words are the title.
- A writer that keeps a heading in brackets (`[Riders]`) in a text it writes has the brackets taken off.

### Added

- **`SlideSpec.headings`:** the headings of the document that stand in the slide's text, word for word, in
  order. The planner and `planEdit` set it; the renderer sets the paragraphs it names. A paragraph no entry
  names is plain text, and a value that is not a list of strings names none.

### Changed

- **The writer is asked for less.** A slide whose every slot code sets is left out of the writer's request,
  and a deck with nothing left to write asks no model at all. Before, such a slide was sent as an empty object
  (a closing slide whose only text is a deck field was one).
- The writer is told: a title is the heading in brackets at the top of the slide's text, never one further
  down, which the slide shows in its text; a heading in brackets inside a text it writes is a paragraph of its
  own, without the brackets.
- `planEdit` sets a slide's title to the heading its text opens under where it sets the text again from its
  passages, the change is to its passages alone, and the slide was titled with its heading or had no title. A
  change in words may speak of the title: the writer then titles the slide, offered the heading where the
  slide had it, and a title the person gave stays. A slide the change writes is settled as in a new deck; a
  slide it leaves alone keeps its text and its headings as they were.
- **Text is measured in bold where it is drawn bold.** With fonts, the headings of a slide are measured in the
  bold face (`FitOptions.bold`), so a heading that takes a second line in bold is counted so. Without the bold
  face supplied, such a text is not measured, and the character budgets hold.
- **Previews keep the space after a heading.** Pagus trims the text of each run; the adapter keeps a space at
  either end of a run (a word joiner beside it), so a heading set run in is not shown glued to its text.
- In a deck written in another language than its document, a heading the writer translated and kept in
  brackets loses the brackets and is set as a heading.
- `slideAlternatives` keeps a slide's `headings` on the layouts it offers.

### Upgrading from 0.6.5

Nothing to change in code. A stored deck renders as in 0.6.5, byte for byte: it has no `headings`, and its
headings stay plain lines until the deck is made again or changed. An app that stores decks whole keeps
`headings` with them; one that rebuilds slides field by field must copy it. A test that counts the writer's
requests, or compares its schema, sees no request for a slide whose every slot code sets.

### Known limits

- A heading too long for the slide's title is titled by the writer, which may leave it out of a free slide's
  title. A heading over a list item or a table row is on a line of its own and can end a column.
- On a layout that draws no bullets, the items of a list from a document are set without their markers (as
  before); a heading over such a list stays a plain line.
- A text the writer writes (translated, asked to be shorter, or cut) has its headings marked only where it
  keeps them word for word.
- A heading set run in uses U+2003 (an em space). Fonts without it fall back to another font for that space.

## 0.6.5

English and mixed text on a right-to-left template reads as it was written. An English deck made from a Hebrew
template showed `The digital transaction share reached landing just ,74% 75% shy of our ambitioustarget` for
`The digital transaction share reached 74%, landing just shy of our ambitious 75% target`, `a minute login-40
outage` for `a 40-minute login outage`, and `12%+` for `+12%` in a table: the English sat in runs the template
marks Hebrew, and PowerPoint orders the words of a run by its language. Plans, model requests and Hebrew decks
are as in 0.6.4.

### Fixed

- **English text in a run marked Hebrew.** A text with letters that is not right to left, written into a run
  the template marks with a right-to-left language (`he-IL`), is marked with a left-to-right one: the deck's
  language when it reads left to right (`en` gives `en-US`, `en-GB` stays `en-GB`), and `en-US` otherwise. The
  end of its paragraph is marked the same when it states a right-to-left language. PowerPoint then shows the
  words and figures in their order, and checks their spelling in English.
- **Figures in an English deck on a right-to-left template.** A figure without letters (`+12%`, `-3.5%`,
  `4.6 / 5`, `2027 – 2029`) in a deck that reads left to right is set left to right, in a left-to-right
  language. In a right-to-left paragraph they showed as `12%+`, `3.5%-`, `5 / 4.6` and `2029 – 2027`. In a deck
  that reads right to left, figures keep what the template gives them.
- **A Hebrew line that starts with a Latin word, in a right-to-left deck.** `AI בשירות הלקוחות` was set left
  to right, from its first letter: its bullet went to the left and `AI` to the end of the line. In a deck that
  reads right to left, a paragraph with any right-to-left letter now reads right to left and is marked with a
  right-to-left language. In a deck that reads left to right the rules are as before.
- The same rules hold for every text slotdeck writes: slots, deck fields, table cells, the labels of a free
  slide (a shared legend) and the sample text of a composed layout.
- `directionOf` knows `iw` (Hebrew's old tag), `ckb` (Central Kurdish) and `syr` (Syriac) as right to left, as
  the renderer already did. One list of right-to-left languages serves the planner, the renderer and the
  measurer.

### Changed

- **The measurer follows the renderer.** With fonts, text is measured in the language and direction it is
  written in. `FitOptions.direction` says which way the deck reads, and defaults to the template's direction;
  slotdeck's own checks pass the deck's.
- **Previews follow the pptx.** `PreviewRequest.direction` tells a renderer which way the deck reads.
  `previewDeck`, `patternPreviews`, the chat tools and `visionReview` fill it, and the Pagus adapter marks a
  paragraph right to left by the renderer's rule.

### Upgrading from 0.6.4

Nothing to change in code. A stored deck renders as before, byte for byte, unless it has, in a run the template
marks right to left, text with letters and no right-to-left letter, or a figure in a deck that reads left to
right; or, in a deck that reads right to left, a paragraph that starts with a Latin word and holds a
right-to-left letter. Those paragraphs change as above. On a right-to-left test template, 12 Hebrew decks
rendered byte for byte as in 0.6.4, and 2 English ones changed in 85 English paragraphs, one language tag
each. Re-render stored English decks made from a right-to-left template to correct them. A test that compares
the rendered bytes of such decks sees the new tags. A `PreviewProvider` of your own may read
`request.direction`, or leave it.

### Known limits

- A run that states no language of its own and inherits a right-to-left one from the template's styles keeps
  it: English written there is still ordered as Hebrew. Templates saved by PowerPoint state the language on
  every run.
- An English-only line in a Hebrew deck ("KPI" in a list) reads left to right, as before, so its bullet stands
  on its left.
- Chart text keeps its language and direction as before.
- Spelling underlines of the changed runs were checked on a computer without Hebrew proofing only.

## 0.6.4

Warnings and notes can reach a person in their own language. A chat app showed people who write in Hebrew
slotdeck's English: a box of warnings such as `the text from the document for this slide (p1-p14) is 3727
characters, and the slide holds about 2480 even with its text set small, at 7 pt: it is cut to fit`, and a
chat model that answered a request in Hebrew with the English of a note. Plans, decks, files and model
requests are as in 0.6.3.

### Added

- **A code on every warning.** `Issue.code` says what it is (`text-cut-dense`), and `Issue.values` holds the
  numbers and labels its sentence states (`{ label: "p1-p14", chars: 3727, holds: 2480, size: 7 }`). A code is
  never renamed. The English `message` is as it was. The 56 codes are listed in "Warnings in a person's
  language" in [api-reference.md](api-reference.md), and `ISSUE_CODES` lists them.
- **`formatIssue(issue, language, { deck })`** gives the line in English or in Hebrew. With the deck it is the
  whole line, as the chat tools write it: in English exactly their line, in Hebrew `שקף 1: הטקסט מהמסמך לשקף
  הזה (קטעים p1-p14) באורך 3727 תווים, והשקף מכיל כ-2480 תווים גם בגופן קטן של 7 נק': הטקסט נחתך כדי שייכנס. ...`
  (the label `p1-p14` stands between the invisible marks U+2068 and U+2069, left out here). Any other
  language gets the English.
- **`ToolResult.issues`:** every line of `output.warnings`, in its order, and a note for each slide the `note`
  of a plan or of a change that stops before a cut speaks of (`slide-holds-less`, `slide-set-small`,
  `change-would-cut`), and for each text of `not_on_slides` (`text-retold`). It is not in `output`, so the chat
  model's tokens do not grow. `TextCutError` has the same notes in `issues`, and `freeOff`.

### Changed

- **The chat model is told to speak the user's language.** A result with `note`, `warnings` or
  `not_on_slides` ends `next` with `Tell the user what this result says in the language they write in, in your
  own words and with the same slide numbers and figures; never paste its sentences as they are.`
  `create_presentation` and `edit_presentation` then have a `next`. A change that stops before a cut asks the
  chat model to ask the person, in their language, whether they accept it. The plan tool's description says
  to show the plan in the language the user writes in: 6 words more in the declarations, about 7 tokens.

### Upgrading from 0.6.3

Nothing to change in code. To show warnings and notes in the person's language, show
`(result.issues ?? []).map((issue) => formatIssue(issue, uiLanguage, { deck: result.deck }))` in place of
`output.warnings` and the `note`, and add the language line to the chat's system prompt
([own-text.md](prompts/own-text.md)). [warnings-in-language.md](prompts/warnings-in-language.md) is the prompt
for a coding agent. A test that compares whole `Issue` objects now sees `code` and `values`; one that matches
`output.next` sees the language line where a result has something to tell.

### Known limits

- English and Hebrew only.
- What is not slotdeck's own text stays as it was inside a Hebrew line: a review model's sentence
  (`seen-in-picture`), values from the deck, and the names of slots other than the title, the subtitle, the
  body and the headings and texts of columns, cards and points. The reason a layout cannot be a free slide's
  frame (`free-base-refused`) is said in Hebrew where it is about the layout itself, and stays English
  otherwise.
- Errors (`output.error`) and the `summary` of a change of colour or layouts are still English, and get no
  language line. The chat model retells them as it did.
- The template check (`checkManifest`) gives no codes: it speaks to template authors.

## 0.6.3

A slide that carries a person's text holds what one rule says it holds. Planned, set by code, or written as
told, such text ended inside its box in PowerPoint in the template's own face. Without a document or an amount
of text, a brief is planned, written and rendered as before.

### Fixed

- **A slide held about 1 character of its text.** On a right-to-left template, a document of 1,359 words on 2
  slides, each text beside a card, a figure or a step, was cut to `the slide holds about 1`. Each slide is now
  its text alone, set dense, planned at 7 pt, and holds 474 and 469 of its 685 and 674 words.
- **A text set dense was planned one way and checked another.** 0.6.2 left a document's headings out: 52 lines
  with no full stop were planned as held, set word for word, and ran 18 pt past their box. One rule now counts
  such text, headings and all: paragraphs in whole lines by length, filled into columns as PowerPoint fills
  them, each column keeping a line free. Code sets a document's text as it is only where it fits. In 3 to 100
  paragraphs alike, the rule accepts no longer text than 0.6.2 did.
- **A writer inside what it was told was sent back:** one in six, in random plans of 0.6.2. Text set dense and
  small is now told what any text of its paragraphs holds; in 3,500 random plans no writer of dense text was
  sent back. A text that does not fit is always told so.
- In a fixed deck of three slides or fewer, an opening or a closing slide keeps passages its neighbour cannot
  hold. A table layout given only the document's tables keeps them at an end of any deck, as in its middle.
  With free slides off, no slide "holds about 0".

### Changed

- **The text comes first.** In a plan, a chart, a table, a list or cards that would cut a document's text, or
  leave under half of the words asked, give way where the text alone keeps more of it. A part a change asks
  for stays, then and in later changes (`editPlan` marks the slide, `OutlineEntry.partsAsked`, internal): the
  parts of a free slide, and a layout's chart or table. A layout of cards, figures, a list or a quote gives
  way to an amount of text asked for it, as in 0.6.2.
- **A slide set dense holds less:** about 4,350 characters, some 720 English words, at 9 pt on a slide 13.33
  inches wide (0.6.2: 4,850); about 2,900, some 440 Hebrew words, at 7 pt on one 10 inches wide (0.6.2:
  3,400). A cut text is told twelve paragraphs at most.
- A short dense list that one column does not hold goes into two. A layout's slot of four or more paragraphs
  takes one fewer of a document's text of three or more. A cut with free slides off says so, and a slide that
  holds a document's text only with its paragraphs joined says `its paragraphs are joined`.

### Upgrading from 0.6.2

Nothing to change in code. A 0.6.2 deck renders its dense slides (`arrange.dense`) by the new rule: the same
text can be set half a point smaller or up to 2 pt larger, in a box of another height. One whose dense text
the rule does not hold now warns `text is N characters; the slot fits about M`, such as 500 Hebrew words in
twelve paragraphs on a slide 10 inches wide; in fewer paragraphs the same words can still fit. Such a slide is
cut when it is written again. Plan again a 0.6.2 plan whose `holds_words` is far below its `words` or `text_words`. Apps that
followed [own-text.md](prompts/own-text.md) should upgrade; its first test now expects about 440 Hebrew words.
[text-first.md](prompts/text-first.md) is the prompt for a coding agent: the upgrade, and a check that
`holds_words`, the `note` and the warnings reach the person before a presentation is made.

### Known limits

- The rule knows no face. Capitals and long unbroken words, such as web addresses, can run past the box
  unsaid: up to 12 pt in the templates' own faces, 223 pt in a wide face. In a wide face, text at the top of
  its told range can too: 500 German words ran 24 pt past.
- The last column can come out short, at times almost empty.
- A list of short items, each a little shorter than a line, is counted two lines an item: a slide can cut it,
  or join its items, where it would fit whole. 0.6.2 set such lists whole, and most ran past their box.
- A table layout at an end of a fixed deck, given text and the table, becomes text holding the table, cut.
- `editPlan` can take a slide inserted before a reworded one of the same parts for the plan's slide.

## 0.6.2

A slide keeps the text a person asks for, on the slides they asked for, and a deck can carry a document a
person gives, in its own words. A deck, a plan and a manifest from 0.6.1 work unchanged, and a deck renders to
the same bytes. A brief without a document and without an amount of text is planned and written as before,
with the same requests to the model as 0.6.1; only a request the Gemini adapter sends again, after an answer
cut off at its limit, asks for light thinking, unless the adapter sets `thinkingLevel`.

### Fixed

- **An amount of text was spread over more slides, or cut.** Asked for 2 slides of 500 words each on the
  right-to-left template, 0.6.1 wrote 271 words on each: the most a slide held at four fifths of the template's
  text size. With no count given it spread them over 4 slides of 250 words; with free slides off it wrote 76.
  A slide now keeps the amount asked for. Where the template's own size cannot hold it, the text is set
  smaller on a free slide of that text alone (`FreeArrange.dense`), in two or three columns in the larger room
  under the title, down to the least size that reads: 9 pt on a slide 13.33 inches wide, 7 pt on one 10 inches
  wide. 500 Hebrew words go on one slide of the right-to-left template at 7 to 7.5 pt, and 500 English words on
  one slide of the demo template at 10.5 to 12 pt. A slide asked for words is never spread over more slides, by
  a plan, a change to a plan or a change to a deck. A text too long even at that size is cut to what the slide
  holds, and that is said before anything is written.
- **A text of an asked length could overflow its free slide.** The budget of a free slide's text counted two
  paragraphs. A writer that kept to the length it was told, in six to twelve paragraphs, went over the box: 250
  Hebrew words at every such count, 225 English words at six, ten and twelve. The text was then sent back to be
  cut. Every free slide of one text with an asked length is now marked dense, and is set smaller where its
  paragraphs need it.
- **A document came out as a summary.** Given a document of 491 words as its brief, the planner of 0.6.1 wrote
  8 slides on the right-to-left template that sum it up: 34 to 55 words on each slide of content, 271 words in
  all, titles included. Nothing gave a slide its own part of the document, the writer was told to keep slide
  text short, and in the chat tools the chat model retold the document in `topic` before the planner saw it.
- **The chat model dropped the person's numbers.** "2 slides" went into `instructions`, where the planner does
  not read a count, and the planner planned 6 to 10 slides. The tool descriptions now say to set `slide_count`
  whenever a number of slides is named, to pass an amount of text on in `instructions` as it was said, and not
  to split, shorten or sum up anything; the planner sets long text smaller.
- **A Gemini answer cut off at its limit was asked for again the same way.** A writing request whose answer is
  about 550 tokens was seen to fill its limit of 11,334 with 10,900 tokens of reasoning before the answer was
  done: one send in five, and twice in a row in one run, which failed the deck. The adapter now sends a cut-off
  request once more with thinking level `low` and the same limit, unless the adapter sets `thinkingLevel`; in 3
  such sends it took no reasoning and every answer was complete. The first request is unchanged. When the
  answer is cut off again, the error says what filled the limit, reasoning or the answer itself, where the
  service counts the tokens. See "Runaway answers" in [providers.md](providers.md).

### Added

- **Dense text.** `FreeArrange.dense`, set by the planner on a free slide of one text that carries an amount of
  text a person asked for, or their own text. Where that text is too long for four fifths of the template's
  size it is set as large as fits, down to the least size that reads, in the larger room under the title, in
  two columns, or three where two would make lines longer than about 95 characters. It is counted as PowerPoint
  sets such text: a character at 0.48 of its size, where PowerPoint was measured at 0.41 on a middling line and
  never over 0.47. It is planned for the most the writer is told, in any number of paragraphs up to twelve: the
  range a writer is told ends where the text fits however many paragraphs of about one length it writes, so a
  writer inside it is not sent back to shorten its text. A slide with the flag whose text fits as before is set
  out as before. `freeContent` and `freeSlide` keep the flag (`FreeSlideContent.dense`). See "A long text a
  person asked for" in [free-slides.md](free-slides.md).
- **`planTextFit(template, plan)`**: before anything is written, the size each slide's text is set at and
  what it holds, for every slide given words or text from a document. See "Before writing: `planTextFit`" in
  [generating.md](generating.md).
- **Making a deck from a document.** `Brief.document` takes a document word for word. Code splits it into
  numbered passages: paragraphs, lists and tables, under the document's headings. The outline gives each slide
  the passages it carries (`OutlineEntry.passages`), every passage on one slide and in order unless the brief
  asks otherwise, and code puts right what the answer missed. A slide's length comes from its passages. With
  no count, a text longer than a slide holds at the template's own size goes over more slides, each with its
  own passages. With a count, from `slideCount`, from a number of slides the brief names or from "each part on
  its own slide", the passages are shared over those slides and set smaller where they need it, and none is
  left out to make room. A slide of several texts that would be cut is one text, set dense, and its passages
  are set into it by code. Where a slide is still cut, the passages are shared again, in order, over the
  slides that carry them, one a change keeps as it was too, and only where one is cut even so, over the slides
  of text that carry none; a chart slide keeps its chart. A slide asked to be shorter keeps its passages. A deck
  of three slides or fewer has no opening or closing slide unless asked: where the number is fixed, an answer
  with a slide that carries no passage is asked again once, and one that keeps an opening slide then stands. A
  slide that holds its text gets it word for word, set by code, where the document is in the deck's language,
  and the writer writes its title. A slide the writer writes (a translation, a summary, a cut) is told to keep the
  text's words and order; one in the deck's language with room for nearly all of its text is sent back once
  when it keeps less than 85% of them. The outline also says the language the document is written in
  (`DeckPlan.documentLanguage`, `DeckSpec.documentLanguage`); one the text's letters belie, with under a tenth
  of them in its script, is not kept. A Hebrew text full of English terms, said to be Hebrew, is Hebrew.
  See "Making a deck from a document" in [generating.md](generating.md).
- **Text pasted from a word processor reads as it looks.** A tab at the start of a paragraph sets it in, and a
  tab after an item's marker is the item's: neither makes a table, which is lines with a tab between filled
  cells, one after another. The `o` and `§` bullets of a list within a list are bullets, shown as `◦` and `▪`.
  A table pasted from a spreadsheet keeps its empty corner cell. A run of three short lines or more with no
  stops (a list typed without bullets) is text, all of it, not every other line a heading. A table of the
  document on a slide with no table is part of its running text, and the slide has room for it. Where
  passages are shared out again, over the parts of a slide spread over several or over a fixed number of
  slides, a slide's table (a free slide's or a layout's) goes to the slide that carries the document's table.
  In a layout's list, an item set by code has the list's bullet in place of its own, and a number one space
  after it. With no count and an answer that places nothing, code adds no more slides than the document has
  passages.
- **Kept with the plan and the deck.** `DeckSpec.document` and `SlideSpec.passages`. A change to a plan or a
  deck shows the document, numbered, with the passages of each slide and those on no slide. The plan of a
  change to a deck says which slides the request asks for in other words (`reword`): such a slide is written by
  the writer, and not set by code, also where its passages move.
- **Texts in the chat tools.** With the new `documents: true` option, `plan_presentation` takes `documents`:
  the names of texts the person gave in the conversation, which the app passes with each call
  (`call(name, args, { documents })`). `describeDocuments` writes the line that lists them for the chat's
  system prompt. A `topic` or `instructions` that copies a listed text is an error that names the argument
  and says to name the text instead; one that retells it gets a note (`not_on_slides`). A named text that only
  asks for a presentation is planned again as instructions. A shorter paste, which is not listed, goes in
  `instructions` whole, in quotation marks, as the argument and the error for a name not listed say. Any
  number of texts may be named, a text under two names (a file and the message it came in) is carried once,
  and names are compared in one Unicode form, without direction marks. See "Texts the person gave" in
  [chat-tools.md](chat-tools.md).
- **What a plan says.** For slides given words or text from a document, the plan tool's output says the size
  of the text where it is below four fifths of the template's (`text_pt`), what a slide holds where it is cut
  (`holds_words`), and a `note` for the chat model to pass on. Slides of a document have `passages` and
  `text_words`; `create_presentation` says how many of the document's words are on the slides
  (`document.kept_words`), or that the deck is in another language than the text (`document.translated`).
- **A cut is said before a change is written.** `planEdit` with `stopBeforeCut` stops after its plan call when
  a slide would hold less text than it carries, with a `TextCutError` that names the slides. `edit_presentation`
  does so unless it is called with `accept_cut: true`, and then answers with a `note` for the person. What a
  slide holds is counted as it will be written, with the heading or caption it keeps, so the error gives what
  the slide would hold. A part added beside a dense text, or a text from the document, leaves that text less
  room: a text that still holds there stays as it is, and one that does not is written again for that room,
  and a cut there is said like any other, where 0.6.1 kept the text as it was and let it run past its box. A
  photo that leaves a dense text too little room gives way to it: `the photo of this slide is left out to make
  room for its text of about 300 words`. A slide asked for no length whose parts and passages stay ("simpler
  words", a new title, another language) is never stopped: it keeps the length of its characters, or is
  written to what it holds.
- **The writer is held to an asked length.** A text written under 85% of the least it was told is sent back
  once; still short, the deck is told: `500 words were asked for this slide, and it has about 380: ask for it
  again to have them all`.
- **Names**: `FreeArrange.dense`, `FreeSlideContent.dense`, `planTextFit`, `SlideTextFit`, `TextCutError`,
  `EditOptions.stopBeforeCut`, `Brief.document`, `OutlineEntry.passages`, `DeckPlan.documentLanguage`,
  `DeckSpec.document`, `DeckSpec.documentLanguage`, `SlideSpec.passages`, `documentPassages`, `DocumentPassage`,
  `MOST_DOCUMENT_CHARS`, `describeDocuments`, `ChatDocument`, `ToolCallContext` and
  `PresentationToolsOptions.documents`.

### Changed

- **Asked lengths, since 0.6.1.** A slide asked for words no longer goes over more slides in a plan
  (`planOutline`, `editPlan`) or a change to a deck (`planEdit`): it is set dense where it needs it, or cut
  where even that cannot hold it. The purposes numbered "(1/3)" for such a text, and the warning
  `the 200 words asked for slide 1 go over slides 1 to 5: ...`, are gone. A slide set dense that is still cut
  is told `1000 words were asked for this slide, and it holds about 527 even with its text set small, at 7 pt:
  the rest is cut. Ask for the text over more slides, or for a shorter one`. A text-and-photo slide asked for
  more than it holds becomes a text alone, and so does a plan's slide of several texts that would be cut: its
  texts become one. A slide with a chart, a table, a list or cards, or a deck's slide with a second text
  (`planEdit`), keeps them and holds what it can. With free slides off a slide holds what its layout holds,
  and no slide is added. Every free slide of one text with words now carries `arrange.dense`; where its text
  fits at the template's size the writer is told the same range as before. A change to a plan with words shows
  each slide's size and most (`text at 7 pt`, `at most about 520 words`, `cut to fit`) and has one more rule on
  words. A change to a deck that writes a dense slide's text again keeps its length unless the request gives
  another; one that keeps the text (a new title) writes the title alone.
- **The chat tools' declarations change for every app**: the `slide_count`, `instructions` and `instruction`
  texts, one sentence of `plan_presentation`, and `accept_cut` on `edit_presentation`: about 1,100 characters,
  some 270 tokens, on each chat turn, and about 1,020 characters more with `documents: true`. These are what the
  chat model reads; what the planner sends is unchanged.
- Only for a brief or a deck with a document: the prompts that plan show it in numbered passages, with rules on
  passages in place of "vary the patterns across the deck" and of the opening and closing line; the writer's
  rules say to keep the document's words; and fewer slides have a photo unless the brief says how many
  (`imageDensity` defaults to `low`).

### Measured

- **PowerPoint.** In the design, 161 dense slides on the right-to-left and the demo template (texts of 1,000 to
  5,500 characters in 3 to 12 paragraphs, with and without a heading, at text sizes normal and large): the text
  ends inside its drawn box on every one, at worst right at its bottom, and PowerPoint's own line count agrees.
  With this release's first budget, of what 12 equal paragraphs hold (for 500 words, 3,700 Hebrew characters
  and 3,810 English ones), 96 more on both templates, at text scale 1 and 1.25, with and without a heading: 500
  words in 5 to 12 paragraphs, and the top of the told range and that budget in 12 equal paragraphs and in one
  long paragraph with eleven short ones. 3,250 Hebrew characters are set at 7 to 7.5 pt in 3 columns, 3,000
  English characters at 10.5 to 12 pt. The running text ends inside its box on every one, at worst right at its
  bottom on the demo template and 0.6 pt above it on the right-to-left one. With the budget that holds in any
  number of paragraphs, 36 more: the top of the told range for 500 Hebrew words on the right-to-left template
  (3,430 characters) and for 800 and 500 English words on the demo template (4,890 and 3,440), each in 1 to 12
  equal paragraphs, at 7 to 11.5 pt. The running text ends inside its box on every one, at worst right at its
  bottom. A budget of what 12 paragraphs hold would send a writer inside its range back to shorten its text:
  3,450 Hebrew characters in 11 paragraphs, 5,220 English ones in 10.
- **Live**, the production text model, each call capped as the adapter caps it, on the right-to-left template,
  with no pictures. A Hebrew document of 491 words (474 of them in its 8 passages): given as the prompt to
  0.6.1, it became 8 slides of 271 words that kept 29% of its words. Given as `document` to 0.6.2, it became 8
  slides that kept 467 of the 474, with no warning (the 7 others on a two-column layout the writer filled to its
  columns' limits); with `slideCount: 3`, all 474 on two dense slides at 9.5 and 10 pt, and a third slide of
  steps whose passages moved to the second because its layout could not hold them. Two such texts together,
  907 words in 16 passages, asked for on 2 slides "about 500 words on each, as it is": 2 dense slides at 7 and
  8 pt that kept 885 of 885. "2 slides, 500 words on each" with `slideCount: 2` and no document: 2 dense slides
  of 519 and 513 words at 7.5 pt. Each deck took two text calls, or three, and in PowerPoint's own rendering
  every text ends inside its box. Before the fix of the cut-off retry above, two of these runs ended in a
  writing call cut off twice, which is how it was found.
- **Offline**, a model that answers as told, on the right-to-left template: 2 slides of 500 words with
  `slideCount: 2` gave 2 slides of 500 words at 7 pt (0.6.1: 271 each), each told
  `Write 500 words, as was asked: 3250 to 3430 characters in all`, and a writer that wrote 3,280 characters in 7
  paragraphs on each was taken with no repair and no warning. Without a count the two slides kept 500 words
  each (0.6.1: spread over 4 of 250); a change "keep it to 2 slides" to 0.6.1's 4-slide plan gave 2 (0.6.1: 4
  again). A free text of 200 to 270 Hebrew words, written in 6 to 12 paragraphs, overflowed in 13 of 16 cases
  on 0.6.1 and in none on 0.6.2. A Hebrew document of 474 words in 8 passages went on 2 slides at 10 and 9.5 pt
  with `slideCount: 2`, and on 1 slide at 7.5 pt with `slideCount: 1`; every passage was placed, every slide was
  filled by code, and all 474 words were kept.

### Upgrading from 0.6.1

- Nothing to do for decks without a document and without an amount of text. A plan or a deck from 0.6.1 with
  `words` works; a change to it no longer spreads its text.
- In a chat, pass the person's numbers and text on: see [own-text.md](prompts/own-text.md). It asks for
  `documents: true`, the list of the conversation's texts on each call, `describeDocuments` in the system
  prompt, and the lines of that prompt that make the chat model set `slide_count`, pass amounts as they were
  said, name a text instead of retelling it, pass a short paste in quotation marks, and set `accept_cut` only
  when the person accepts a cut. Take out any line that sends attached files into `topic`.
- Keep free slides on: a long text is set smaller only on a free slide. With `free: false`, or a list of
  patterns without `free`, a slide holds what its layout holds.
- Show `text_pt`, `holds_words` and `note` where the app shows a plan, and let the new warnings and the `note`
  of `edit_presentation` through.
- An app that calls the planner itself sets `brief.document`, keeps `prompt` to what the deck is for, and can
  call `planTextFit` between `planOutline` and `writeDeck` to say what is set small or cut. For changes it can
  pass `stopBeforeCut` to `planEdit` and catch `TextCutError`.
- A store that keeps plans and decks keeps the person's document too, once for every version. Delete them with
  the document when the app deletes it.
- An app that sets `thinkingLevel` on the Gemini adapter keeps its level when a request is sent again after an
  answer cut off at its limit. An app with a provider of its own can ask once more with light reasoning when
  an answer is cut off, as "Runaway answers" in [providers.md](providers.md) shows. The cut-off error still
  starts `Gemini's answer was cut off at the limit of N tokens, twice`; what follows it is new.
- Known: a document of up to 12,000 characters (about 1,800 Hebrew or 2,000 English words). Several texts
  named together are one document. A text for a plan or deck that has none means a new plan. `planEdit`
  without `stopBeforeCut` says a cut after writing, not before. Titles, headings and table cells are cut to
  their limits. A heading of 60 characters or more, with no "#" and no blank line under it, is read as text,
  and so is one right under a line of text with no stop at its end, or one over two short lines with no stops.
  A short line over one other short line is read as a heading over it. A fixed number of slides shares the
  passages again by the size of their text, not by each slide's room: a slide with a table beside its text
  may be cut while another has room. A plan or deck whose document has no language recorded (the outline gave
  none) fills slides by code where the document is in the script of the deck's language. A slide of a deck made
  by 0.6.1 carries no dense flag, so a change that writes its text again plans it for the usual length unless
  the request names one.

## 0.6.1

A slide gets the amount of text a person asks for. A deck, a plan and a manifest from 0.6.0 work unchanged, and
a deck renders to the same bytes.

### Fixed

- **200 words asked, 76 written.** The planner did not read an amount of text from a brief or a request. The
  writer was told what each slot says of length ("About 420 characters in all") and that limits are hard, so
  it wrote what the slot said, whatever was asked. On the right-to-left template one slide of "200 words of
  running text" came out as 76 words on its text layout, and as 53 on a free slide.
- **A free slide of one text was told far less than it holds.** A free slide is planned for texts of the usual
  length, 300 characters for running text. One such text alone is set large on a short line, so its slot said
  `Write 260 to 500 characters` with a limit of 528 on the right-to-left template, and 584 on the demo template.
  The same slide holds 1,194 characters in two columns at the running-text size, with no warning.

### Added

- **Asking for an amount of text.** The outline gives a slide `words` (`OutlineEntry.words`, 1 to 3000) when
  the brief asks for an amount of text on it: "200 words", "a paragraph of about 80 words". Code then finds
  the text a slide that holds it, counting about 6 characters to a word, 6.5 in Hebrew: the layout planned
  when nine tenths of its limits hold it, else the smallest layout of plain text that does, else a free slide
  set out for that length, which keeps a chart, a table or a photo the layout had. Where none holds it and the
  slide shows text alone, the text goes over as many slides as it needs while the deck may have more, each
  with its share and its purpose numbered "(1/3)"; otherwise the slide holds what it can and the deck gets a
  warning. The writer is told the length before the rest of the
  slot's description, as a range that starts at it: `Write 200 words, as was asked: 1200 to 1370 characters in
  all`. See "Asking for an amount of text" in [generating.md](generating.md).
- **The same in a change.** A change to a plan keeps the words a slide has, and lists them; a slide whose
  layout and words it leaves as they were is not fitted again. A change to a deck ("make slide 2 about 200
  words") gives the slide the room as a plan does, and a slide that becomes a free slide for it keeps its
  chart, table or photo as it is; where nothing holds the text, it goes over new slides after it, and a
  warning says where.
- **The plan tool shows them.** A slide of `plan_presentation` and `edit_presentation_plan` that was asked an
  amount of text has `"words"` in its line.

### Changed

- Every prompt that plans slides (a new plan, a change to a plan, a change to a deck) has one more rule, on
  amounts of text, and an optional `words` property for each slide. A brief without such a wish is planned as
  before.
- A free slide's running text is planned at the length asked for, where a plan has `words` for its slide. Its
  budget then says what its box holds at that length: 1,377 characters for 200 words on the demo template,
  1,450 for 200 Hebrew words on the right-to-left template.

### Measured

- **Live**, the production text model, each call capped as the adapter caps it: on the right-to-left template,
  a deck of one slide asked for "200 words of running text". 0.6.0: 53 words. 0.6.1: 212 words in three
  paragraphs, on a free slide set out for them in two columns, inside their box in PowerPoint's own rendering.
  Two text calls. Told "about 200 words" alone the model wrote 164, and told a range from nine tenths of the
  length it wrote 168: the range now starts at the length.
- **Offline** on the right-to-left template, with a model that answers as told: 200 words planned on its text
  layout go to a free slide that holds 1,450 characters; 80 words (520 characters) go to a free slide too,
  since nine tenths of its page layout's 556 are 500; with free slides off, 200 words go over three slides of
  67 when the deck may have six, and a deck of one slide is told what the slide holds (76 words) with a
  warning; 600 words go over three free slides of 200; 200 words on its chart layout go to a free slide of the
  text and the chart, which holds 96, with a warning, and are not spread; a change to a deck that asks for 200
  words on a slide makes it a free slide, or, with free slides off, three slides with a warning. A brief
  without such a wish gives the same plan as before.
- **PowerPoint**: free slides of one text filled to their budgets (1,281 characters in 9 paragraphs and in 4,
  1,200 in 6, 1,450 in 10 and in 3) lie inside their boxes with room to spare.

### Upgrading from 0.6.0

- Nothing to do. A plan an app keeps may now have `words` on a slide: keep it with the plan, as the plan's
  other keys. An app that shows a plan's slides may show it ("about 200 words").
- An app that passes `patterns` to the planner or the tools without `free` gets no free slides: a long text
  then goes over several slides where the deck may have them, and otherwise holds what the slide's layout
  holds, with a warning. Add `free` to the list.
- An app whose own prompts tell the chat model, or the planner through `instructions`, to keep slides short
  works against a person who asks for an amount: pass such a wish on in the person's words.
- Known: a layout's limits can be more than its box holds when its text has many paragraphs and the template
  keeps space between them in points: 532 characters in five paragraphs run over the 556 of the right-to-left
  template's page layout, whose paragraphs have 12 pt before them. An asked length fills a layout to nine
  tenths of its limits for that reason, and a free slide counts that space. Text written to a layout's own
  description is as it was in 0.6.0.

## 0.6.0

More freedom: slides the software lays out itself, a deck's own colours, and a planner that does what was
asked about the form of a deck. A deck, a plan and a manifest from 0.5 work unchanged, except a deck whose
`textScale` is refused (see Upgrading), and a deck that uses none of the new things renders to the same bytes
as with 0.5.2, unless 0.5.2 rendered it to a file PowerPoint did not open (see Fixed).

### Added

- **Free slides.** A slide may list **blocks** instead of naming a pattern: texts, lists, cards, figures,
  steps, a quotation, a note, charts, tables and pictures, any number of them. The software lays them out on
  the frame of one of the template's layouts, with texts, panels and tables copied from shapes the template
  has, so the slide looks like a slide of that template. One text alone works; so do twelve charts, a chart
  beside three cards, and two hundred charts. `freeSlide` writes such a slide and `freeContent` reads it back.
  Nothing has to be set up; a manifest may name the frame and the styles (`free`). See
  [free-slides.md](free-slides.md).
- **The planner uses free slides** for content that no layout of the template shows: several charts, tables
  or pictures on one slide, more items than a layout takes, parts that no layout combines. It names the parts
  in the plan, and the writer fills them. On by default; `free: false` on the planner or the chat tools turns
  it off, and so does `"free": false` in a manifest. See "Free slides when a deck is planned" in
  [free-slides.md](free-slides.md).
- **Colours.** `colors` on a deck (`{ "accent": "1565C0" }`) gives every slide another colour in place of the
  template's main colour; its tints and shades follow, greys, photos and the logo stay, text stays readable,
  icons are recoloured. `colors` on a slide does it for that slide alone. `recolorDeck` sets them with no model
  call, `templatePalette` shows what the software takes for the template's colours, and a manifest may name
  them or forbid changes (`colors`). See [colors.md](colors.md).
- **Colours in the chat tools.** `accent_color` on `plan_presentation`, `edit_presentation_plan` and
  `edit_presentation`, and `color_slides` on the last for single slides. A colour alone recolours a
  presentation without a model call. No model reads a colour out of a brief: the chat model passes it.
- **A list of layouts.** `Brief.patterns` (the `layouts` argument of the chat tools) limits a deck to the
  layouts a person names: "only the text layout". The list is kept with the plan and the deck
  (`DeckSpec.patterns`), and a later change keeps to it unless it asks for another layout.
- **Wishes come first.** The first rule of every prompt that chooses layouts now says that what the brief
  asks about the form of the deck (how many slides, which kinds, how many things on a slide, how many side by
  side) comes before variety and before the planner's habits. A deck of one slide is a deck of one slide; the
  opening and closing slides are for decks of three or more.
- **`validateDeck` takes the loaded template** as well as the manifest. With the template it builds a deck's
  own layouts in, lays free slides out and measures text with the template's fonts, so no option is needed.
- **Chart detail.** A chart slot may say `detail: "compact"` or `"spark"` for a chart drawn small; free slides
  choose it by the size of each chart.
- **Studio settings.** `setManifest` takes `colors` and `free`; `editManifest` hears them in a person's words;
  `checkManifest` reports a setting that points at nothing.
- **A change to a free slide, part by part.** A change from the chat, `planEdit` or `editDeck` names the parts
  of a free slide it writes again ("add two charts", "shorten the third card"); the other parts are kept as
  they are, and so is the slide's arrangement unless the change is about it. Where the plan of a change named
  no part to write, the slide is left as it was with the warning `slide N: nothing on it was changed: the
  plan of the change named no part to write again`.
- **Names.** New exports: `freeSlide`, `freeContent`, `isFreeSlide`, `freeBaseRefusal`, `FREE_PATTERN`,
  `FREE_BLOCK_KINDS`, `recolorDeck`, `templatePalette`, and the types `DeckColors`, `FreeArrange`,
  `FreeBlockKind`, `FreeBlockContent`, `FreeSlideContent`, `PlanSettings`, `TemplatePalette`, `ColorFamily`.
  New fields: `SlideSpec.blocks`, `arrange`, `base`, `colors`; `DeckSpec.colors`, `patterns`; `Brief.colors`,
  `patterns`; `OutlineEntry.blocks`, `arrange`; `free` in the options of the planner; `colors` and
  `deckPatterns` in the options of `editPlan`, with `colorSlides` in those of `planEdit` and `editDeck`;
  `recolored` in their results; `free`, `patterns` and `colors` in the options of `presentationTools`;
  `TemplateManifest.colors`, `free`; `ChartSlotDef.detail`; `accent_color` (of the deck, and of a slide in its
  own colours), `layouts`, `free_layout` and a slide's `parts` in the tools' outputs.

### Changed

- The planner's prompts were reworded for the points above. Decks planned from the same brief can come out
  differently than with 0.5.2, as with any change of prompt.
- With free slides on offer, `composeLayouts` makes no layouts: a free slide is what a deck gets for content
  no layout shows. Pass `free: false` to keep composing layouts as in 0.5.
- `edit_presentation_plan` and `edit_presentation` no longer require `instruction`: a colour alone is a
  complete request. Called with neither, they answer `"instruction" is required unless a colour is given`.
  With `colors: false` both are as in 0.5.
- `tools.definitions` are made for the template and the options the tools were built with: the `layouts`
  argument lists the layouts those tools can plan with, and the colour arguments are left out where colours
  are off. They were one constant before.
- A warning that has no slide is printed by the chat tools as its message alone.
- The chat tools' `warnings` also carry what the renderer says of the deck as a whole, once, also after a
  change of colour alone.
- The chat tools hold a `slide_count` above 30, which their declaration already forbids, to 30.
- Where the software itself fails on a chat tool's request (a stored plan or presentation that is not what the
  tools put there), the chat model is told so in a sentence it can pass on (`The stored plan could not be
  read. Make a new plan with plan_presentation.`), not `Cannot read properties of undefined`. What the app's
  own model, image provider, store or review throws is passed on as before.
- Three warnings say `1 line`, `1 character` and `1 row` where they said `1 lines`, `1 characters` and
  `1 rows`.

### Fixed

- **A file PowerPoint could not open.** A control character in a text, a table cell, a chart label or an
  alternative text (a vertical tab, which comes with text copied out of a text box, a form feed and the like)
  was written into the file as it was. Such characters are now left out; a vertical tab breaks the line in a
  text slot and is a space elsewhere. A deck that has none renders to the same bytes as before.
- **A number where text belongs.** A text slot given a number, an object or a list holding one made
  `validateDeck` and `renderDeck` throw a `TypeError`. They answer with the issue `expected text` on the slot.
- **Layouts built for a deck** (`DeckSpec.layouts`) are built again for a template object that was given
  another manifest, as a template's other answers are. A manifest object edited in place is still not seen.
- **The shape a generated picture is asked for.** Where a picture slot states a `box` (a mirrored layout, a
  wide diagram), the image provider is asked for the aspect ratio of that box, where the picture is placed,
  not of the sample picture's own frame. Pictures already stored are used as they are.
- **More files PowerPoint could not open.** Each of these rendered without a word to a file PowerPoint did
  not open; each is now an error that says what is wrong:
  - a deck's `textScale` that is no number (`NaN`, an object, `"large"`, `""`) or is outside 0.1 to 300, on a
    deck with a chart or other text whose size the factor sets (text of 12 pt or less: most layouts of the
    right-to-left template, only the native chart of the demo template). `null`, which 0.5.2 took as a factor
    of 0, is now read as not set; a number written as text (`"1.25"`) is read as that number, as before;
  - a manifest's own `textScale` that is no number (`NaN`, `"abc"`) or is 0 or less: `validateDeck` and
    `renderDeck` stop the deck, and `checkManifest` reports it;
  - a chart of more than 255 series;
  - a size or colour that `changeManifest` took for a chart's or a table's text (`"red"`, `0`, `1e9`), a
    line colour, a list of chart colours that is no list, or a text slot's `paragraphGap` that is no number:
    `changeManifest` now refuses each. A manifest that already holds such a value (taken by 0.5.2, or written
    by hand) is not checked, and a deck with such a slot still renders to a file PowerPoint does not open,
    except a text size, which is now written inside 1 to 4000 pt (below).
- **Text larger than PowerPoint takes.** The size written is the template's size times the manifest's
  `textScale` times the deck's. A product past 4000 pt or under 1 pt (a deck's 300 on a template whose
  manifest says 1.3, or 0.1 on one that says 0.5) gave a file PowerPoint did not open. Such a size is now
  written as 4000 pt or 1 pt; every other size is written as before.
- **Errors of the library where a `TypeError` was thrown.** A deck, slide or `layouts` of another shape
  (`null`, a string) gets `the deck is no object with a list of slides` and the like from `validateDeck` and
  `renderDeck`; a brief without a prompt and a plan of another shape get a `PlanError`; `checkManifest`
  reports patterns or slots that are no lists; an asset whose bytes are no picture is a `DeckValidationError`
  naming the slot (before: a plain `Error`, `Unsupported image format: expected PNG, JPEG or GIF`, or a
  `TypeError` for a value that is no `Uint8Array`); the Gemini adapter rejects an answer of the right status
  and the wrong shape with a `GeminiError`. The messages are in [troubleshooting.md](troubleshooting.md).
- **Words the planner did not know.** A model's `text_size` or a brief's `textSize` that is not `normal`,
  `large` or `x-large` keeps the size there is, also names every object has (`"constructor"`); a brief's
  `imageDensity` outside its four values is planned as `medium`. Before, a `textSize` such as `"constructor"`
  gave the deck a text size that was no number, and a model's `text_size` dropped the deck's size; an
  `imageDensity` outside the four values put `undefined` into the model's instructions; an object such as
  `{ "toString": 1 }` in either threw a `TypeError`.
- **Time and memory.** A word of 20,000 letters with no space was measured in quadratic time; it takes half a
  second now. `checkDesign` on a slide whose `items` was `Infinity`, `sampleDeck` with budgets of a billion,
  `composeLayouts` with thousands of parts and `writeDeck` with item counts of a billion ran for minutes or ran
  out of memory; they now answer at once, the last with a `PlanError`. A table of more than about 125,000 cells
  (the number depends on Node's stack) no longer makes `validateDeck` throw a `RangeError`; one of a million
  cells is reported like any other.

### Measured

- **The same bytes.** 24 decks of 0.5.2 in 31 renders (7 on the demo template, each in both layout modes, and
  17 on a right-to-left template of 35 layouts), with every kind of slot and larger text, render to the same
  bytes as with 0.5.2; so do the eleven decks of `test/golden.test.ts`, one of them with a layout of its own.
- **Free slides in PowerPoint.** 13 decks of free slides on the two templates: cards, steps, figures, tables,
  charts and quotations, and on each template twenty slides from one sentence to two hundred charts.
  PowerPoint, asked how it sets the text of every block, finds none outside its box, card or slide. The
  twenty slides render in 0.7 s on the demo template and 1.3 s on the right-to-left one.
- **The planner with a stub model.** A deck with twelve charts on one slide, planned, written and changed to
  add two charts: 5 model calls, each with an output ceiling; the change wrote only the two new charts and
  kept the twelve as they were. Sixty charts on one slide: 4 calls.
- **The chat tools with a production model**, every call capped, pictures as placeholders, on both templates:
  plan with a colour asked for, write, add two charts, change an unrelated title, add a table to the slide of
  key figures, take a part off a slide, rename one chart. 7 steps, 12 model calls on each template. Every
  change left the other four slides as they were, and the chart slide kept every chart it was not asked to
  change. The colour asked for in the first request is on every slide. PowerPoint finds no text of a free
  slide's blocks outside its box in any of the decks. One step on the right-to-left template failed once: the
  model ran on to the tight output ceiling of the test, twice. It passed when run again, and in two earlier
  runs.
- **Text sizes.** A deck with charts, a table, text and a free slide opens in PowerPoint at every `textScale`
  from 0.1 to 300 on the demo template. On the right-to-left template, whose manifest says 1.3, 0.5.2 gave
  files PowerPoint did not open from 219.8 up; with the sizes held, they open.

### Upgrading from 0.5.2

- Stored decks, plans and manifests need nothing, except a deck whose `textScale` is refused (below). Render as
  before.
- Warnings without a slide number: free slides (on by default when a deck is planned or changed) and colours
  give warnings whose `Issue.slide` is not set; the slide is named in the message (`slide 3 holds 9 blocks, as
  asked; ...`). An app that prints `slide ${issue.slide}: ${issue.message}`, as the chat tools of 0.5.2 did,
  prints `slide undefined`; print the message alone when `slide` is not set.
- An app that shows a plan's slides by `layout`: a free slide has the layout `free`, `free_layout: true` and no
  thumbnail. Show its `layout_description`, which says what the slide holds.
- An app that walks `SlideSpec.slots` by its manifest's slot names: a free slide's slots are named `b1Text`,
  `b2Chart`, and so on. `isFreeSlide(slide)` tells; `freeContent(slide)` gives the blocks.
- An app that passes `patterns` to the planner gets no free slides unless the list names `free`; the tools take
  the same option from 0.6.0.
- An app that sets `composeLayouts` gets no new layouts while free slides are on offer; pass `free: false` to
  have them made as in 0.5. `free: false` offers what 0.5 offered and nothing else; the reworded prompts stay.
- The chat tools are declared with reworded descriptions, and `plan_presentation`, `edit_presentation_plan`
  and `edit_presentation` take a `layouts` argument wherever a layout can be named. Both are so whatever
  `free` and `colors` say; `colors: false` leaves the colour arguments out. Calls made as with 0.5.2 are
  accepted as before; their answers may hold free slides unless the tools are built with `free: false`.
- An app that keeps one set of tool declarations for all templates: take `definitions` from the tools built
  for the template a conversation uses.
- `presentationTools` now looks at the template's slides when it is called, to see whether the template can
  carry free slides: about 0.2 s for the demo template and half a second for one of 35 layouts, once for a
  loaded template. Load the template once and build the tools from that object; an app that loads it anew for every
  chat turn pays this on every turn. See "Build the tools once for a loaded template" in
  [chat-tools.md](chat-tools.md).
- A stored deck whose `textScale` is no number or is outside 0.1 to 300 is now refused with a
  `DeckValidationError`. Such a deck rendered to a file PowerPoint did not open when it had a chart or other
  text whose size the factor sets (text of 12 pt or less: most layouts of the right-to-left template, only
  the native chart of the demo template), for values such as `NaN`, `"large"`, `""`, `false`, `0`, `0.05` and
  `1000`. A deck without such text opened at any value, and so did every deck with a value 0.5.2 read as a
  number, such as `true` or `[1.5]`. All of them are refused now: render it without the value, or with a
  number.
- A manifest whose chart or table text, line colour, chart colours or `paragraphGap` 0.5.2's `changeManifest`
  took with a value that is no colour or number still gives files PowerPoint does not open: set those values
  again with `changeManifest`, which now refuses what does not work. A manifest whose own `textScale` is no
  number stops every deck with an error: set it again, or remove it.
- Code that caught a `TypeError` for input of the wrong shape gets the library's own error now:
  `DeckValidationError` from `renderDeck`, an error issue from `validateDeck`, `PlanError` from the planner,
  `GeminiError` from the adapter. Code that caught `Unsupported image format` from `renderDeck` gets a
  `DeckValidationError`.
- Where none of the template's patterns can be filled with what was supplied, a deck is now planned from free
  slides when they are on offer. `No pattern of template "x" can be used` comes only when they are not: the
  template cannot carry them, the deck fields their frame shows were not supplied, `free: false` (on the
  planner or the tools, or `"free": false` in the manifest), a list of patterns that does not name `free`, or a
  template that has a pattern of its own called `free`.

## 0.5.2

Hebrew text that PowerPoint underlined as misspelled.

### Fixed

- **Red underlines under new text.** PowerPoint saves the result of its spelling check on each run of text:
  `err="1"` says a mistake was found in it, `dirty="0"` says it needs no new check. A filled text took its
  formatting from the template's sample run, and with it that saved result. Where the template's sample was
  Hebrew typed while the text was marked English, PowerPoint had saved it as misspelled, so it underlined the
  whole new text in red without checking it, until someone right-clicked a word. A filled text now drops both,
  and PowerPoint checks it in its own language when the deck is opened.
- **Hebrew that starts with an English word was marked English.** A text such as "AI בשירות הלקוחות" kept the
  template's English language on its run, so PowerPoint checked its Hebrew words as English. A run is now marked
  with a right-to-left language when most of its letters are right to left, not only when the first one is. The
  paragraph's direction still comes from the first letter, and text fitting measures such a run as the
  renderer marks it.

### Measured

On the right-to-left template, a sample deck of its 35 layouts plus a layout made from the template slide whose
sample PowerPoint had saved as misspelled, on two slides: with 0.5.1, 4 filled texts carried `err="1"` and 14
carried `dirty="0"`; with 0.5.2, none. The `err="1"` runs left are in that slide's texts that are not slots,
as the template has them.

### Upgrading from 0.5.1

- Nothing to do. Render a deck again to drop the saved results from it.

## 0.5.1

Fixes for templates that repeat a sample text under their titles, and for lines on composed layouts.

### Fixed

- **A sample under every title was a deck field.** A draft made every text that repeats across a template's
  slides one value for the whole deck. A sample such as "Subtitle" under every title repeats too, so every
  slide of every deck showed it. Now only a text in the header or footer may become a deck field, and the
  draft's deck call asks the model whether each is one value for the deck or a sample each slide replaces,
  such as "Bottom line: text here". The rest are listed with each slide's texts, for the model to name.
- **`prepareManifest` repairs a manifest drafted before.** A deck field whose text the template repeats in the
  content of its slides becomes a text of each slide that a slide may leave out (`empty: 'remove'`), and
  `prepared.fields` no longer holds the sample. A field in the header or footer stays. `done` says which.
- **A line through a composed layout.** A layout composed on a slide whose frame has a line through the
  content, such as a divider beside a column of titles, kept the line, and the new parts lay across it. Such
  a line is not kept now. Neither is a deck field that lies in the content, such as a drafted "Subtitle".
- **A rule through a card.** Where everything is drawn, a rule of a design that runs through a text or cuts
  through a panel is left out, with a warning. `composePattern` first asks the model once more. This was seen
  as a green rule down the middle one of three cards.
- **A rule within a card stayed behind.** When cards were made equally wide, a rule across a card did not move
  with it and could stick out of the card. It moves with its card now.

### Measured

On the right-to-left template, with a manifest drafted before this version (its "Subtitle" sample a deck field
on 28 layouts, its department name on 40):

- `prepareManifest` made the sample a text of each slide on the 28 layouts and left the department name a deck
  field.
- The grid of each of the 27 layouts a new layout can stand on is the same as in 0.5.0, before and after that
  repair, so a layout made for a deck with 0.5.0 is built in the same place.
- The 15 decks with layouts of their own from the 0.5.0 measurements render as before, except in 5, where a
  rule within a card now moves with the card: by up to a third of a percent of the slide's width, and in one
  from 4 percent outside the card to inside it.

### Upgrading from 0.5.0

- Decks need nothing: they render as before, with the changes above.
- For a template drafted before 0.5.1 whose slides show a sample such as "Subtitle" in every deck, run
  `prepareManifest` on it again, with `remove: false` for a template in use. Its `done` names each field it
  changed. Stop passing that field in `fields`: it is no longer in `prepared.fields`.
- After that repair, a stored deck that does not fill that text renders without it with `layout: 'adaptive'`.
  With `'strict'`, a slide must fill a text it may leave out, and such a deck does not render. A studio that
  compares stored decks before and after a repair should call `validateDeck` with the layout mode the chat
  uses: with the default, `'strict'`, every such deck is listed as worse.
- The value an app stored for that field is no longer read. It can stay or be deleted.
- A sample in the header or footer, such as a bottom line, stays a deck field in that repair. Where it is one,
  `{ op: 'setSlot', pattern, slot, deckField: null, empty: 'remove' }` on each layout that has it makes it a
  text of each slide. A new draft asks the model about it.
- The draft's deck call has a ceiling of 300 tokens plus 50 for each repeated text, 10 more than before.

## 0.5.0

New layouts made while a deck is planned. A text-only deck from a template with one plain text layout used that
layout for every slide, and three goals or a plan in four steps became paragraphs. With `composeLayouts`, those
slides get cards, steps or a timeline, in the template's own look. Off by default.

### Added

- **`composeLayouts`** in the planner's options (`planOutline`, `planDeck`, `generateDeck`) and in
  `presentationTools`: up to that many new layouts per deck. After the outline, one light call names the slides
  of content that none of the offered layouts shows well, and what each needs. The layouts are composed side by
  side with `composePattern`, on the frame of the layout each slide had, so the title stays where the deck's
  other titles are. The slides then use them, and the deck is written once. A layout that cannot be made leaves
  its slides as planned. See "New layouts for a deck" in [generating.md](generating.md).
- **`composeLayouts(template, plan, options)`**: the same as a step of its own, with what was made and what could
  not be (`made`, `failed`).
- **`layouts` in `DeckSpec` and `DeckPlan`**: the designs of the layouts made for the deck. The template file is
  not changed; the stored deck JSON is all a later render needs, and it needs no model.
- **`withLayouts(template, layouts)`**: the template with a deck's layouts built onto it, once per template and
  list. `renderDeck`, `writeDeck`, `editPlan`, `planEdit`, `editDeck`, `checkDesign`, `designReview` and
  `visionReview` call it themselves. When the template gains a layout of the same id, the template's own is used.
- **`base`** in `ComposeOptions`, in `PatternDesign` and as the third argument of `designKit`: the layout a new
  one stands on. `composePattern` records the base it used, so that a design is built on the same frame again.
- In the chat tools, a planned slide with a new layout has `new_layout: true` and the new layout's description.
- A progress event `{ stage: 'layouts' }`. `validateDeck` against a manifest without the deck's layouts says to
  use `withLayouts`.
- **[prompts/deck-layouts.md](prompts/deck-layouts.md)**: a prompt for a coding agent, to turn this on in the
  path an app's chat already takes: one setting, a mark in the plan, the deck kept whole. No new tool.
- [prompts/compose-layout.md](prompts/compose-layout.md) gains a fourth operation: keeping in the template a
  layout the chat made for a presentation, from its design with `buildPattern`.

### Changed

- A composed text or list whose box holds much more than its sample is described as "Drawn for about N
  characters", N being the sample's length, so that a deck fills it as the design meant. This holds for layouts
  composed in a studio too.

### Measured

On Gemini 3.8 Flash, 10 briefs on two templates (a right-to-left template in Hebrew and the demo template; 8
text-only, 2 with varied layouts), each written from one outline as before and after composing:

- Every deck got one or two new layouts: 17 made, none failed. The text-only decks went from one layout for four
  slides to rows of cards, numbered steps, a timeline, a comparison, an action list with owners, a quotation
  and key figures. The decks with varied layouts got one each, for a slide of running text.
- Composing took 9 to 24 seconds, in two to five light calls. Writing took about as long as before.
- Rendering the stored deck again, with the template loaded afresh, gave the same file every time.
- One slide was wrong: a section divider was given cards on its dark frame. Only slides of content may get a
  new layout now.
- The writer filled the cards of the Hebrew template to about a third of the design's sample length, which left
  them mostly empty; with the length the box was drawn for, about three quarters.

Five of the briefs again after these two fixes: 9 layouts made, none failed, in 12 to 25 seconds; the section
dividers kept their layout. Lists in tall cards are still shorter than the card.

### Upgrading from 0.4.0

- Nothing changes unless `composeLayouts` is set. Decks without `layouts` render as before.
- Store the deck JSON as before; it may now carry `layouts`. Keep the template as it is.
- Code that checks decks itself against `template.manifest` (`validateDeck`, `slideAlternatives`) should use
  `withLayouts(template, deck.layouts).manifest` for a deck with layouts.
- A slide with `new_layout: true` has no thumbnail in the app.
- Designs stored from 0.4.0 have no `base` and are built on the same frame as before.

## 0.4.0

A layout the template does not have, made from its own parts. A text-only variant of a slide with a photo keeps
the text where it was and leaves half the slide empty; a new layout fills it.

### Added

- **`composePattern(template, instruction, { llm, effort: 'low' })`** asks the model for a new layout in the
  template's look and builds it: a new slide after the last one, on the frame of one of the template's own
  slides (background, header, logo, page number, its title where every title is), made of copies of the
  template's own text styles and decorations, and a pattern for it with measured budgets. The model designs on
  a grid of 12 columns and 12 rows and can only name styles the template has; the code checks the design,
  sends one with problems back once, makes numbered items equal, and proves the result renders. It returns a
  new pptx and manifest to store together. See "composePattern" in [template-studio.md](template-studio.md).
- **`buildPattern(template, design)`**: the code half alone, for a design the app has. **`designKit(template)`**
  lists the grid, the base slide and the styles. **`DesignError`** carries every reason a design cannot be built.
- **[prompts/compose-layout.md](prompts/compose-layout.md)**: a prompt for a coding agent, to add "a new layout
  from words" to a template studio: compose, look at the sample, keep or discard.
- A composed slide is named `composed <id>` in the pptx and left out of the kit, so composing a second layout
  does not change the keys a first design was made with.
- `keepCorners` keeps a rounded panel's corners when it is resized (used inside; not exported).

### Measured

On Gemini 3.8 Flash with `effort: 'low'`, 10 requests with this release's code on two templates (a right-to-left
template of 36 slides and the demo template), such as "create a text only layout", "an agenda with 6 items", "a
slide with three key figures and a caption under each" and a request in Hebrew: all 10 were built, in 6 to 25
seconds, 7 of them after the design was sent back once. Twenty earlier requests, made while the code was
written, were all built as well. With the model's default reasoning one request took 210 seconds.

### Upgrading from 0.3.2

- Nothing changes for code that does not call the new functions.
- A composed layout changes the template's pptx: store `made.pptx` and `made.manifest` together. Every other
  pattern points at the same slide as before, so stored decks still render.

## 0.3.2

A prompt for a studio that cannot yet add a variant of one layout. No code changed.

### Added

- **[prompts/layout-variants.md](prompts/layout-variants.md)**: a prompt for a coding agent, to add two
  operations to a template studio: list the variants a layout allows, with the reason for each one it does not
  (`patternVariants`), and make one (`changeManifest` with `derivePattern`): the mirror image, a text-only
  version, or a version without an optional text. It covers a template that the chat already uses: the check
  before storing, taking a variant away again, and prepare adding variants back. Server logic only; a screen
  only where the studio already shows one layout with actions beside it.

### Upgrading from 0.3.1

- Nothing to do. The library is the same as in 0.3.1.

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
