# Troubleshooting

Every message slotdeck produces, what it means, and what to do. Then the known limits.

## Setup problems

| Symptom | Cause | Fix |
| --- | --- | --- |
| `ERR_REQUIRE_ESM`, or `require() of ES Module` | The package is ESM only. | Use `import`. Set `"type": "module"` in your `package.json`, or use `.mjs` files, or `await import('slotdeck')` from CommonJS. |
| `Cannot find module 'slotdeck/node'` in a bundler | `slotdeck/node` uses `node:fs` and is for Node only. | Import it only in server code. Everything else is in `slotdeck`. |
| `Not an Office Open XML package: [Content_Types].xml is missing` | The bytes passed to `loadTemplate` are not a `.pptx` (an old `.ppt`, a PDF, an HTML error page saved as a file). | Pass the bytes of a real `.pptx`. Save a `.ppt` as `.pptx` in PowerPoint first. |
| Types are not found | TypeScript module resolution is too old for package `exports`. | Use `"moduleResolution": "bundler"`, `"node16"` or `"nodenext"`. |

## Deck errors

These make `renderDeck` throw `DeckValidationError` and are returned by `validateDeck` with `severity: 'error'`.
Each has the slide number and usually the slot name.

| Message | Meaning | Fix |
| --- | --- | --- |
| `the deck has no slides` | `slides` is missing or empty. | Give at least one slide. |
| `the deck is no object with a list of slides, such as { "version": 1, "slides": [...] }` (also `the slide is no object with a pattern and slots, ...` on a slide, and `"layouts" is an object; it is a list of the deck's own layouts, ...`) | What was passed as a deck, a slide or `layouts` has another shape: `null`, a string, a number. Before 0.6.0 this was a `TypeError`. | Pass the parsed deck. An application that stores decks: check what was read back. |
| `the deck's textScale is NaN; it is a factor on body text between 0.1 and 300, such as 1.25` | `textScale` is no number, or is outside 0.1 to 300. `null` is read as not set, and a number written as text (`"1.25"`) as that number. Before 0.6.0 such a deck rendered to a file PowerPoint did not open when it had a chart or text of 12 pt or less, whose size the factor sets. | Leave `textScale` out, or give a factor such as `1.25`. |
| `the template's textScale is NaN; it is a factor on body text such as 1.25, which changeManifest sets with "textScale"` | The manifest's own `textScale` is no number, or is 0 or less. Every deck on the template is stopped: its sizes would be written as `NaN`, which PowerPoint does not open. `checkManifest` reports it too. | Set it with `changeManifest` (`{ op: 'setManifest', textScale: 1.25 }`), or remove it from the manifest. |
| `the chart has 256 series; a chart takes at most 255` | PowerPoint does not open a chart of more series. | Fewer series: a chart reads well with up to about eight. |
| `asset "x" is not a PNG, JPEG or GIF picture` | (From `renderDeck`.) The bytes passed for the asset are no picture, or empty. | Pass the bytes of the picture file. |
| `a free slide holds at most 2000 blocks, not 5000` | Past that a slide is no slide, and laying it out would take minutes. | Fewer blocks; spread them over slides. |
| `unknown pattern "x" (available: a, b, c)` | The slide names a pattern the manifest does not have. | Use one of the listed ids. |
| `pattern "x" takes 2 to 4 cards, not 5` | `items` is outside the pattern's range. | Use a count in the range, or raise `max` in the manifest if the template has room. |
| `this slot is required` | A required slot has no value (or an empty string), or a required deck field was not given. | Give the slot a value, or add the field to `deck.fields`. |
| `leaving this slot empty needs layout: 'adaptive'` | An optional slot was left out in `strict` layout. | Render with `layout: 'adaptive'`, or fill the slot. |
| `expected text` | A text slot got something that is no text: a picture, table or chart value, a number such as `87`, an object, or a list that holds one of them. | Give a string or an array of strings. A number goes in quotation marks: `"87"`. |
| `expected an image reference such as { "asset": "id" }` | An image slot got something else. | Give `{ "asset": "<id>" }`. |
| `expected a table such as { "header": [...], "rows": [[...]] }` | A table slot got something else. | Give a `header` array and a `rows` array of arrays. |
| `the table has no columns` | `header` is empty. | Give at least one column heading. |
| `row 2 has 3 cells; the table has 4 columns` | A row's length differs from the header's. | Give every row one cell per column. Use `""` for an empty cell. |
| `expected a chart such as { "type": "column", ... }` | A chart slot got something else. | Give `type`, `categories` and `series`. |
| `unknown chart type "x"` | `type` is not `column`, `bar`, `line` or `pie`. | Use one of the four. |
| `the chart has no categories` / `the chart has no series` | An empty array. | Give at least one of each. |
| `series "x" has 3 values; the chart has 4 categories` | A series' length differs from the categories'. | One value per category. |
| `series "x" has a value that is not a number` | A value is `null`, a string that is not a number, or missing. | Give numbers. |
| `a pie chart takes exactly one series` | A pie with several series. | Give one series, or use another type. |
| `asset "x" was not provided` | The deck refers to a picture whose bytes were not passed. | Add it to `options.assets` under that id. With the planner, pass the same `icons` and `assets` to rendering as to planning, or use `resolveAssets`. |
| `a free slide needs at least one block, such as "blocks": ["text"]` | A slide with an empty `blocks` list. | Name at least one block, or use a pattern. See [free-slides.md](free-slides.md). |
| `"blocks" names "x"; a free slide is made of text, list, card, ...` | A kind of block that does not exist. | Use one of the kinds listed. |
| `a free slide cannot be made from this template: ...` | The template has no layout a free slide can stand on, or no running text to copy. The sentence says which: no layout with a title, a title whose place neither its slide layout nor the master states, running text that does not read on the frame's background, or running text that is all in placeholders. | Add a pattern with a title at the top and running text in text boxes to the manifest, or name a frame with `free.base`. |

## Deck warnings

These do not stop rendering. `renderDeck` and the planner return them.

| Message | Meaning | What to do |
| --- | --- | --- |
| `text is 230 characters; the slot fits about 211` | Text over the slot's `maxChars`. It is rendered and may overflow its box. | Shorten the text, or raise the budget in the manifest if the sample deck shows room. |
| `text is 38 characters; the slot fits about 1` on a part of a free slide (`b3Text`) | The box of that part is lower than one line of its text, even at the smallest size: a card among too many, one under the other, or a template that sets its lines far apart. The text would be drawn across the card's edge. | Fewer blocks on the slide, or more of them side by side (`arrange.columns`). |
| `text takes 7 lines at 18 pt; the box holds 5 (2 lines too many). Text of about 290 characters fits.` | (Template loaded with fonts.) The text was measured and runs over its box by that many lines. | Shorten the text to about the length given, or allow shrinking with `shrinkText`. See [text-fit.md](text-fit.md). |
| `text takes 6 lines even at 15.5 pt; the box holds 5 ...` | The text was set as small as `shrinkText` allows and still does not fit. | Shorten the text. |
| `text is 30% wider than its box at 12 pt, and the box does not wrap. ...` | (With fonts.) A one-line box whose text is too long. | Shorten the text. |
| `text takes 1 line where the template's own text takes 5, so its box will look empty. ...` | (From `checkDesign` or a `review`.) The text fits but leaves its box almost empty. Also: `...so the row will look uneven`, and `the last of its 2 lines holds one word, ...`. | Lengthen or shorten the text as the message says. See [design-checks.md](design-checks.md). |
| `seen in the picture of the slide: ...` | (From `visionReview`.) A model looked at a picture of the slide and reported this. It is only as reliable as the picture. | Check the slide; see the cautions in [design-checks.md](design-checks.md). |
| `text has 5 paragraphs; the slot fits 4` | Over `maxParagraphs`. | Fewer paragraphs. |
| `the table has 6 columns; the slot fits 5` (also rows) | Over the table's limits. The table may leave the slide. On a free slide (slot `b1Table`) the limit is what the table's box holds at the smallest text size, so the table will be drawn past its box, over what is under it. | A smaller table. On a free slide also: fewer blocks beside it, or a slide of its own. |
| `a cell is 40 characters; cells fit about 22` | A long cell makes its row taller. On a free slide a row is made as high as its cells need, and the text is set smaller until the table fits its box. This is said when that takes the text under four fifths of the size the table would have with a word in every cell: the table is still drawn inside its box, in small text, down to the smallest size. | Shorter cells: they make the table's text larger. |
| `the chart has 12 categories; the slot fits 8` (also series) | Over the chart's limits. Labels may crowd. On a free slide (slot `b1Chart`) the limits are what reads well at the size the chart is drawn at: 8 categories and 3 series for a large chart, fewer for a small one, 12 categories and one series for a chart drawn without axes. The chart is drawn all the same. | Fewer categories or series, or fewer blocks on the slide so that the chart is larger. |
| `pattern "x" has no such slot; ignored` | The slide gives a slot name the pattern does not have. | Check the name against the manifest. With `items`, slots of items beyond the count are reported this way. |
| `pattern "x" has a fixed number of items; "items" is ignored` | `items` on a pattern without a `repeat`. | Remove `items`. |
| `chart format "x" is not known and is ignored; use "number" or "percent"` | An unknown `format`. | Use one of the two, or leave it out. |
| `chart values not stated in the brief: 89` | (Planner only.) The model charted a number the brief does not contain. It may be invented, or worked out from the brief. | Check the number. Put the figures you want charted into the brief. |
| `200 words were asked for this slide, and it holds about 92: ask for the text over more slides, or for a shorter one` (from 0.6.3, with `; free slides are off, so its text is not set smaller` at the end where they are) | (Planner only.) An amount of text was asked for a slide, and the slide the plan could give it holds less: free slides are off, or the slide keeps a chart, a table, a list or cards beside which it holds half of the text or more, or a deck's second text beside its text, or, in a change to a deck (`planEdit`) or to parts a change to a plan (`editPlan`) asked for, then or in an earlier change, a free slide's parts or a layout's chart or table, whatever they hold, or it stays on a layout. The slide holds what it can, and the writer was told so. The plan said so before writing (`planTextFit`, `holds_words`). | Ask for the text over more slides, or for less. Allow free slides (`free`, and `free` in an app's `patterns`), which set a long text smaller. See "Asking for an amount of text" in [generating.md](generating.md). |
| `1000 words were asked for this slide, and it holds about 446 even with its text set small, at 7 pt: the rest is cut. Ask for the text over more slides, or for a shorter one` | (Planner only.) The slide was set dense for the text asked, down to the least size that reads, and still holds less. The rest is cut. | Ask for the text over more slides, or for less. A slide holds about 440 Hebrew words on a slide 10 inches wide, and about 720 English words on one 13.33 inches wide, in twelve paragraphs or fewer; a text of more paragraphs holds less. |
| `500 words were asked for this slide, and it has about 380: ask for it again to have them all` | (Planner only.) The writer wrote the slide well short of the length asked, under 85% of the least it was told, and again when it was sent back. | Ask for the slide again, or change it with the length in the words of the request. |
| `the photo of this slide is left out to make room for the 500 words asked for it` | (Edits only.) A change asked for more text on a slide of a text and a photo than it holds, and the slide became a text alone, set dense. | Nothing, or ask for the text over two slides and put the photo on one of them. |
| `the photo of this slide is left out to make room for its text of about 300 words` | (Edits only.) A change added a photo beside a dense text, or a text from the document, that does not hold beside it. The text stays as it is, alone, and the photo is not added. | Ask for the text over two slides and put the photo on one of them, or for a shorter text. |
| `the text from the document for this slide (p3-p5) is 4100 characters, and the slide holds about 3200: it is cut to fit. Ask for more slides to keep it whole` (also `... even with its text set small, at 7 pt: ...`, and from 0.6.3 `...; free slides are off, so its text is not set smaller`) | (Planner only.) A deck made from a document: the slide's passages are more than it holds, where no slide could be added for them (the number of slides was fixed, or the deck had its most). With free slides on, its text was set as small as reads, alone on the slide; with them off, it is on the layout of plain text that holds most. A part a change to a plan (`editPlan`) asked for stays, then and in later changes, and a change to a deck (`planEdit`) keeps the parts a slide has: the text holds what it can beside them. A text code would set as it is, headings and all, that the slide does not hold so is written by the writer, told what the slide holds (from 0.6.3). In 0.6.2 a slide that kept cards, figures or steps beside its text could hold about 1 character of it, and a plan made then is written as it was planned. The writer shortened the text. The plan said so before writing. | Ask for more slides, or for fewer passages on that slide. Allow free slides (`free`, and `free` in an app's `patterns`), which set a long text smaller. Plan a deck from 0.6.2 again. |
| `the text from the document for this slide (p1-p27) has 53 paragraphs, and the slide holds it only as 25: its paragraphs are joined. Ask for more slides to keep them apart` (from 0.6.3) | (Planner only.) A deck made from a document: the slide holds all of its text's characters, but not as so many paragraphs, such as the lines of a long list. Its writer was told the most paragraphs that hold all of it, and joined the rest; no word is cut. Text set dense counts each paragraph in whole lines, so a list of short items, each a little shorter than a line, is counted two lines an item. | Ask for more slides, or for fewer passages on that slide. |
| `the slide keeps 72% of the words of its text from the document (p3-p5), though it has room for nearly all of it` | (Planner only.) A deck made from a document: the writer wrote the slide's text in other words, and kept under 85% of its words when it was sent back. | Ask for that slide again, with its text "word for word". |
| `slide 4 holds 60 blocks, as asked; its smallest text is set at 5.5 pt` | A free slide was made with as many blocks as asked for, and some of its text is under 9 pt (on a slide 13.33 inches wide; in proportion on a narrower one), which nobody reads on a projected slide. It has no slide number in `Issue.slide`: it is a note on what was asked for, not a request to write the slide again. Where the small text is a table's and its long cells are the cause, the table's slot is told as well. | Nothing, when that many on one slide is what was wanted. Otherwise spread the blocks over more slides. |
| `slide 4: 12 side by side were asked for; the slide holds 5` | A free slide's `arrange.columns` would have left some block under its smallest size: narrower than its kind sits well in, or under half the least height. The slide was set out with the nearest number of blocks side by side at which every block has that size. No slide number in `Issue.slide`. | Nothing, or state the number the slide holds. See "How the blocks are set out" in [free-slides.md](free-slides.md). |
| `slide 4 holds 9 blocks, as asked; the deck asks for larger text, and this slide has no room for it` (also `... has room for half of that`) | The deck's `textScale` is above 1. On this free slide larger text would leave some part less room than it needs, and its text would have to be set smaller than without the wish. The slide is made with half of the larger text, or with none. No slide number in `Issue.slide`. | Nothing, or fewer blocks on the slide. |
| `slide 4 holds 200 blocks, as asked; 200 charts are drawn small, without axes, legend or labels` | The charts of a free slide are too small for anything but their bars or lines. | The same. |
| `slide 4 holds 9 blocks, as asked; one of them reaches below the room the slide has, and may be cut off at its edge` | A block of a free slide was set out below the room for content. The software lays nothing out there, and says so should it happen. | Fewer blocks on the slide. Please report the deck: it is a fault in the layout. |
| `slide 4 was planned on the frame of the layout "x", which the template no longer offers; it stands on "y"` | A free slide's `base` names a pattern that is gone. | Remove `base` from the slide, or restore the pattern. |
| `slide 4 was planned on the frame of the layout "x", which cannot carry a free slide (the layout "x" has its title low on the slide: blocks would get 25% of the slide's height, and a frame leaves them 45% or more); it stands on "y"` | A free slide's `base` names a pattern that is no frame: a cover, a section divider or a closing slide, whose title leaves no room under it; a pattern with a row of items; one without a title. On a template none of whose layouts leaves 45%, the sentence compares with the layout the slide stands on: `... and "y" leaves them 40%`. | Remove `base` from the slide, or name a content layout. |
| `the colour "accent" is "blue", which is not six hex digits such as "1565C0"; it is ignored` | A colour in `colors` that is no colour. The deck renders without it. | Give six hex digits. See [colors.md](colors.md). |
| `the template keeps its own colours in every deck; the colours the deck asks for are left out` | The manifest has `"colors": false`. | Remove `colors` from the deck, or allow colours in the manifest. |
| `the colour FFFFFF is close to the page's own; fills, chart bars and icons in it will hardly show` | (From `renderDeck` and the chat tools.) The deck's colour is about the colour of a page that stays as it is. `validateDeck` does not say this: it has no page to compare with. | A darker or stronger colour. |
| `the template has no second colour; "accent2" is left out` | The deck asks for a second colour, and the template has only one main colour. | Nothing, or leave `accent2` out. |

## Planner errors (`PlanError`)

| Message | Meaning | Fix |
| --- | --- | --- |
| `No pattern of template "x" can be used: ...` | Nothing in the manifest can be filled with what was supplied, and free slides are not on offer: the template cannot carry them, the deck fields their frame shows were not supplied, `free: false` (on the planner or the tools, or `"free": false` in the manifest), a list of patterns without `free`, or a template with a pattern of its own called `free`. The reasons follow. | See the reasons below. |
| `The model did not produce a usable outline: ...` | The model's answer did not have the expected shape. | Retry. If it repeats, check that your provider returns parsed JSON that follows the schema. |
| `The brief has no prompt: "prompt" is the text of what the deck is about` | `brief.prompt` is missing or no string. Before 0.6.0 this was a `TypeError`. | Give the text of what the deck is about. |
| `The plan is no object with a list of slides` (also `The plan has slides that are no object with a pattern and a purpose: 2, 5`) | A plan passed to `writeDeck`, `editPlan` or `composeLayouts` has another shape: often a stored plan read back wrongly. | Pass the plan as `planOutline` or `editPlan` returned it. |
| `Slide 3: a free slide takes at most 400 parts, not 1000` | A plan whose free slide names more parts than a deck is written with. | Fewer parts on the slide. |
| `Template "x" has no plain text pattern, so a text-only deck cannot be planned from it` | `textOnly` was asked for, and no pattern of the manifest is plain text. | Add a plain text pattern to the manifest (see [layout-freedom.md](layout-freedom.md)), or plan without `textOnly`. |
| `The plan uses patterns that cannot be filled with what was supplied: a, b` | A plan was made with icons, assets or fields that are not passed now. | Pass the same `icons`, `assets`, `fields` and `layout` to every step. |
| `Slide 3: pattern "x" needs icons, and no icon set was supplied` | The same, for an edit. | As above. |
| `The document is 15230 characters, and a deck is made from at most 12000 (about 1,800 words): pass the part the deck is about` | `brief.document` is longer than `MOST_DOCUMENT_CHARS`. One writing call carries at most that much. | Pass the part the deck is about. |
| `"document" must be text` | `brief.document` is no string: a stored plan read back wrongly, or a file's bytes. | Pass the text of the document. A blank one is read as none. |
| `Nothing was changed: slide 2 would hold about 446 of its 1000 words (text at 7 pt). Ask for more slides or a shorter text, or accept that the rest is cut.` | A `TextCutError`, from `planEdit` with `stopBeforeCut`: a slide the change writes would hold less than it carries. Nothing was written. `error.fits` names the slides. | Tell the person. Ask for more slides or a shorter text, or call again without `stopBeforeCut` to cut the text. |

Reasons a pattern is not offered:

| Reason | Fix |
| --- | --- |
| `is built around numbers, and the brief gives none` | Put figures in the brief, or set `numberPatterns: 'always'`. |
| `needs a photo, and images are turned off` | Give an `ImageProvider` or photo assets, and an `imageDensity` other than `none`. |
| `needs icons, and no icon set was supplied` | Pass `icons`. |
| `needs a graphic asset, and none was supplied` | Pass an asset of kind `graphic`. |
| `needs the deck field "x"` | Add it to `brief.fields`. |

## Chat tool errors

The tools answer these to the chat model as `{ "error": "..." }`, which it can act on. The rest are in "Errors"
in [chat-tools.md](chat-tools.md). These come from `plan_presentation` built with `documents: true`.

| Message | Meaning | Fix |
| --- | --- | --- |
| `"documents" names a text the app does not have: "budget.xlsx". The texts the app lists for this conversation are "pasted text 1" and "notes.docx": use those names. A text the user pasted in a short message is not listed: ...` | The chat model named a text that is not in the list passed with the call: a short paste, which is not listed, or a file the app has no text for. | Pass the conversation's list with every call (`call(name, args, { documents })`), and put `describeDocuments(list)` in the chat's system prompt. The chat model then passes a short paste in `instructions`, in quotation marks, and tells the person that a file not listed could not be read. |
| `The text "notes.txt" is empty: tell the user it could not be read.` | The named text is blank: a file whose text the app could not read. | Make the text of such files, or leave them out of the list. |
| `The text is 15230 characters, and a presentation is made from at most 12000 (about 1,800 words). Ask the user to paste the part the presentation is about, and name that.` | The named texts together are longer than a deck is made from. | Nothing: the person pastes the part, which becomes a new text of the list. |
| `"topic" copies the text "pasted text 1". Name it in documents instead ...` (or `"instructions" copies ...`) | No text was named, and that argument copies a listed text of 1,500 characters or more. | Nothing: the chat model plans again with the name. A named text that only asks for a presentation is planned as instructions. If it keeps happening, take out the lines of the system prompt that send attached files or all the person's facts into `topic`. |

## Manifest problems (`checkManifest`)

| Message | Fix |
| --- | --- |
| `patterns is no list of patterns, such as [...]` (also `slots is no list of slots, ...` and `pattern 2 is no object with an id, a slide and slots`) | A manifest of another shape, often one written by hand. Nothing else is checked until it is mended. Other functions given such a manifest may still throw a `TypeError`: check a manifest you did not make with `checkManifest` first. |
| `points at slide 9, but the template has 5 slides` | Correct `sourceSlide`, or the pptx lost slides. |
| `points at shape #99, which is not on slide 1` | The shape id is wrong, or the slide was edited. List ids with `inspectTemplate` or `studioView`. |
| `is an image slot, but shape #6 is a text box` (and similar) | The slot's kind does not match the shape. Change the kind or the shape id. |
| `another pattern has the same id` / `another slot of the pattern has the same name` | Rename one. |
| `has a box, but its shape is inside a group` | A `box` only works for a shape placed directly on the slide. Remove the `box`, or ungroup in the template. |
| `is a chart slot, but its shape is inside a group` | A native chart cannot replace a picture inside a group. Ungroup in the template. |
| `repeat expects a slot named "card1Text"` | A row's name template has no matching slot for one of the drawn items. Add the slot or correct the template name. |
| `repeat item 2 points at shape #77, which is not on slide 4` | Correct the shape ids in `repeat.shapes`. |
| `repeat items must all sit directly on the slide, or all in the same group` | Items in different groups cannot be laid out together. |
| `repeat range 3 to 2 is not valid` | `min` must be at least 1 and not more than `max`. |
| warning: `has no character budget, so nothing stops text that does not fit` | Add `maxChars`. |
| warning: `points at the same shape as "x"` | Two slots on one shape: usually a mistake. |
| warning: `a grid cannot hold more items than it draws` | Lower `max` to the number drawn. |
| warning: `has a budget of 900 characters, but sample text of that length takes 14 lines where 9 fit; about 560 characters fit` | (Template loaded with fonts.) The budget lets text run over the box. Lower `maxChars` to about the length given, or use `measuredBudgets`. See [text-fit.md](text-fit.md). |
| warning: `has a budget of 520 characters on a slide with 2 of its 4 items, but sample text of that length takes 10 lines where 7 fit; about 376 characters fit` | (Template loaded with fonts, row with `widen`.) With that number of items the row's items are wider and their texts have a larger budget, which still lets text run over the wider box. Lower the slot's `maxChars`: the budget at each number of items follows from it. See "Fewer items that fill the row" in [layout-freedom.md](layout-freedom.md). |
| warning: `no font was supplied for Cambria bold; text set in it is checked by its character budget` | Add the font file to `fontkitMetrics`, or accept character budgets for those slots. |
| warning: `free.base names the pattern "x", which cannot be the frame of free slides: ...; they stand on "y"` | The pattern is there and is no frame: its title sits low on the slide, it has a row of items, or it has no title. Free slides stand on the layout named last. Name a content layout, or remove `free.base`. On a template that has no content layout, the layout that leaves most room is the frame and can be named. See [free-slides.md](free-slides.md). |
| warning: `free.styles.heading points at shape #6 of slide 11, which no layout of the manifest uses as a text that can be copied; the software's own choice is used` | The shape is in the template, and no pattern of the manifest fills it as a text any more (its pattern was removed), or it is a placeholder. Name another shape, or remove the entry. |
| `free.styles.heading is null; it names a shape of the template, such as { "slide": 7, "shape": "7" }` | A manifest written by hand. Give the slide number and the shape id, or remove the entry. |

## Refusals from `changeManifest`

`rejected[i].reason` is one of the manifest messages above, or one of these:

| Reason | Meaning |
| --- | --- |
| `there is no pattern "x"` / `"p" has no slot "x"` / `slide 4 has no shape #99` | The change names something that does not exist. |
| `5 cards do not fit side by side: each is 21% of the slide wide and the row has 93%. The most is 4` | The row cannot hold that many. Use the stated maximum. |
| `a grid holds at most the 6 cards it draws; 8 would need a new line of them` | A grid's `max` is the number drawn. |
| `"title" is a text slot; maxColumns does not apply to it` | A setting for another kind of slot. |
| `maxChars must be a whole number of 1 or more` | A budget of zero, a fraction or a negative number. |
| `the size of bodyText is between 4 and 200 points` (also `the colour of text is six hex digits such as "0E7C7B", not "red"`, `lineColor is ...`, `colors is a list of colours, ...`, `paragraphGap is a number of points, 0 or more`) | A size or colour of a chart's or a table's text, a line colour, a list of chart colours, or a text slot's `paragraphGap` that PowerPoint could not open the file with. Before 0.6.0 these were taken, and every deck with such a slot rendered to a file PowerPoint did not open. (A text slot's `fontSize` outside 4 to 200, `a font size is between 4 and 200 points`, was refused before as well.) |
| `"card1Text" is part of the row of cards, so it cannot be renamed, moved or tied to a deck field; remove the row first` | Row slots must stay alike. Use `removeRepeat` first if you really need to. |
| `shape #4 is a picture; it can be photo, icon, graphic, chart, decoration` | The role does not suit the shape. |
| `"x" has no row of like items. Name the slots of each item alike with a number ...` | `setRepeat` found no numbered slots to make a row from. |
| `another pattern is already called "x"` | Choose another id. |
| `a template needs at least one pattern` | The last pattern cannot be removed. |
| `a free slide cannot stand on "cover": the layout "cover" has its title low on the slide: blocks would get 22% of the slide's height, and a frame leaves them 45% or more; name a layout with a title that leaves room under it and has no row of items, such as "list-with-photo"` | `free.base` names a pattern that is no frame. The reason is in the middle: a low title, a row of items, no title. The layout named at the end is one that can be used. Where no layout of the template leaves 45%, the reason compares with that layout: `... and "section" leaves them 40%`. |
| `shape #4 of slide 11 cannot be the quote of a free slide: it is a placeholder, which takes its look from the slide layout, and a free slide copies a text box that states its own` | `free.styles` names a placeholder. Also: `it is not a text of one of the template's layouts`, `it is not a box that the template sets text on`. Name a text box that a pattern fills. |

## Rendering errors that are not about the deck

Plain `Error`s from `renderDeck`. They mean the manifest and the pptx disagree, or a picture is not usable.
`checkManifest` catches the first group ahead of time.

| Message | Fix |
| --- | --- |
| `Pattern "x": slot "y" points at shape #12, which is not on template slide 3` | Fix the manifest (see above). |
| `Pattern "x" points at slide 9, but the template has 5 slides` | Fix `sourceSlide`. |
| `Unsupported image format: expected PNG, JPEG or GIF` | Convert the picture (WebP, SVG, HEIC and others are not supported). |
| `sample tables with merged cells are not supported` | The template's sample table has merged cells. Unmerge them in the template. |
| `a chart slot must point at a shape that is not inside a group` | Ungroup the sample picture in the template. |

## Gemini adapter errors (`GeminiError`)

| Message | Meaning |
| --- | --- |
| `Gemini request failed (401): ...` | The key or token is wrong or missing. `error.status` has the HTTP status. |
| `Gemini request failed (429): ...` | Rate limit or quota. Retry later. |
| `Gemini request failed (400): The specified schema produces a constraint that has too many states for serving` | The service rejected a schema as too complex. This was seen during development with length limits on nested arrays, which the shipped schemas do not use. If it appears with a long deck, asking for fewer slides may help. |
| `Gemini did not answer within 180 seconds` | The request hung. Retry, or raise `timeoutMs`. |
| `Gemini did not return valid JSON: ...` | The model's text was not JSON. Retry. |
| `Gemini interaction ended with status "..."` | The service did not complete the request; the reason follows when it gives one. |
| `Gemini's answer was cut off at the limit of 9552 tokens, twice. ...` | The output limit was reached twice in a row, the second time with light thinking, or with the adapter's `thinkingLevel` when one is set. The rest of the message says what filled the limit the second time, where the service's counts of tokens show it (see "Runaway answers" in [providers.md](providers.md)). `reasoning took ... of them`: a long reasoning before the answer. Try again, or lower the adapter's `thinkingLevel` if it sets one above `low`. `the answer took ... of them`: the model ran away inside one text field, or the answer is larger than the request's estimate. Try again. With neither, the service gave no counts, or none that fill the limit: try again. If it keeps happening, set `maxOutputTokens` on the adapter. It replaces the limit of every request, raising some and lowering others, so set it above the largest limit the app's requests get, not only above the one in the message: a request's limit is its estimate plus 8,192, or 24,576 when it has none. |
| `Gemini request failed (400): Request blocked due to copyright/recitation content` | The service's recitation filter tripped on an ordinary answer. It happens at random. The adapter already sent the request twice; try again. |
| `Gemini request failed (400): Thinking level ... is not supported for this model` | The `thinkingLevel` set on the adapter is not one the model accepts (`minimal` is refused by some models). Use `low`, or leave it unset. |
| `Gemini request failed (429): Your project has exceeded its monthly spending cap` | The key's project reached the spending cap set for it. Nothing works until the cap is raised in the provider's console. |
| `Gemini returned no image` | The image model answered without a picture, often because the prompt was refused. |
| `Gemini returned an image that cannot be read` | The answer held image data that is not base64 text, as a gateway may give. Before 0.6.0 this was a `DOMException`, or, for image data that is no text (a number), a picture of a byte or two that rendering then refused; an answer whose steps, content or errors had another shape was a `TypeError`, and is now one of the errors above. Check what the gateway returns. |

## The result looks wrong

| Observation | Likely cause | Fix |
| --- | --- | --- |
| Text runs out of its box or over a neighbour | The slot's budget is too generous, or the deck ignored a budget warning. | Render `sampleDeck` to see every slot at its budget, then lower `maxChars` with `changeManifest`. |
| A leftover sample element appears (a sticker, a second page number, a sample caption) | The shape is not a slot, so it is copied as it is. | Mark it as a decoration: `{ op: 'addSlot', pattern, shape, role: 'decoration' }`. |
| A picture is cropped badly | `cover` crops to the frame. | Supply a picture closer to the frame's proportions, or set `fit: 'contain'` on the slot. |
| Cards are in reverse order in a right-to-left deck | The pattern's items are not declared as a row. | Add a `repeat` (`setRepeat`) or `sequences`. |
| Hebrew or Arabic text is left-aligned | The manifest keeps alignment. | Set `alignOnDirectionChange: 'flip'` on the manifest or the slot. |
| A font changes for another script | The template's font lacks that script. | Set a complex-script font in the template. |
| PowerPoint underlines Hebrew or Arabic text in red, and the line goes away on a right-click | Before 0.5.2 the text kept the template run's saved spelling result (`err="1"`, `dirty="0"`), so PowerPoint showed the template's mistakes on the new text without checking it. | Upgrade to 0.5.2 and render the deck again. A text that starts with an English word and is mostly Hebrew is also marked Hebrew from 0.5.2. |
| PowerPoint underlines red text that the template left as it was (not a slot) | The template saved that text as misspelled, often Hebrew typed while the text was marked English. | Mark the text Hebrew in the template (Review > Language) and save it, or make the shape a slot or a decoration. |
| Body text looks small | The engine keeps the template's own font sizes, and the template was designed with small body text (9 pt on a 10-inch slide is common in decks exported from Google Slides). | For every deck: set `textScale` on the manifest, such as `1.3`. For one deck: `textSize: 'large'` on the brief, `text_size` in the chat tool, or `textScale` on the deck. See "Larger text" in [layout-freedom.md](layout-freedom.md). |
| Asked for 5 columns (or 2, or 3), the deck uses another layout: a table, a grid, figures | The row that draws the columns does not take that number, or the planner passes it over. In a manifest drafted before 0.3.0 a row takes at most what the template draws; its small pictures may be photo slots, which need photographs and drop the pattern when pictures are off; its description may name the filler's subject ("team profiles"). | `changeManifest(template, suggestChanges(template, { kinds: ['icon', 'row', 'row-max'] }).flatMap((s) => s.changes))`, then give the pattern a description of its structure with `setPattern`. Icon slots are filled from the `icons` given to the planner: an app that gives none follows the repair with `setSlot` and `empty: 'keep'`, which keeps the template's pictograms. `describeManifest(template)` shows each row with its range. Also check that the number the user asked for is in the brief's `prompt` (`topic` or `instructions` of `plan_presentation`). See "Repairing a manifest that is in use" in [template-studio.md](template-studio.md). |
| Asked for 3 columns, the cards come as two on one line and one below | The planner chose a grid. Before 0.3.0 it was not told which patterns set their items in one row. | From 0.3.0 the planner reads `side by side in one row` or `in lines of up to 2` with a pattern's number of items, and is told to answer a wish for columns with a pattern that takes that number and sets it on one line. If the template has no row that takes the number, give one the room with `setRepeat`, or see the row above. |
| A pattern with icons is never used | It has a required icon slot and the app gives the planner no icon set, so it is not offered (`needs icons, and no icon set was supplied`). | Pass `icons` to `planDeck`, `generateDeck` or `presentationTools`, or keep the template's own pictograms: `{ op: 'setSlot', pattern, slot, empty: 'keep' }`. |
| The planner never uses a pattern | It is not offered. | See "Reasons a pattern is not offered". Also check the pattern's `description`: the model chooses by it. |
| A deck that was asked to be "text only" has cards, tables or charts | The wish did not reach the planner as a wish about slides: only "no pictures" was passed (`imageDensity: 'none'`, or `image_density` in the chat tool), which still allows cards, tables and charts. | Set `textOnly: true` on the brief, or `text_only: true` in `plan_presentation`. See "Asking for a kind of slide" in [generating.md](generating.md). |
| A text-only deck uses one layout for every content slide | The template has one plain text pattern besides its opening and closing. | Add more plain text patterns to the manifest as derived patterns: a list without its photo, a quotation, two paragraphs side by side. |
| Asked in a chat for "2 slides of 500 words", the plan has 6 to 10 slides | The chat model put the number in `instructions` or `topic`, where the planner reads no number of slides. The plan's `next` then says so, and asks it to plan again with `slide_count`. | Add the lines of "Guidance for the chat model's system prompt" in [chat-tools.md](chat-tools.md). See [own-text.md](prompts/own-text.md). |
| A presentation of a text the person pasted or attached sums it up | The chat model retold the text in `topic`, and the planner never saw it. | Build the tools with `documents: true` and pass the conversation's texts on each call. See "Texts the person gave" in [chat-tools.md](chat-tools.md). |
| A slide states a number or contact detail that nobody gave | The model invented it. | Figures are guarded only for patterns marked `figures` and for charts. State details in the brief and review the deck. |

## Known limits

- **Shapes need their own position.** The engine measures the shapes on a slide. A shape that takes its position
  and font size from a layout placeholder, without stating them on the slide, is filled correctly but measured
  as empty: a drafted budget for it is only as long as its sample text, and `box` cannot be applied to it. Decks
  exported from Google Slides and decks built from free-standing text boxes work as they are. For a free
  slide, the title of the frame may be such a placeholder: its place is read from the layout or the master.
  The texts a free slide copies its styles from may not be.
- **Control characters are left out.** A character that no pptx part can hold (a vertical tab, a form feed and
  the like, which come with text pasted from other programs) is taken out of every text, cell, label and
  alternative text before it is written. A vertical tab, the line break of a text box, breaks the line in a
  text slot and is a space elsewhere.
- **Budgets are estimates.** A character budget cannot see lines lost to wrapping. Leave room in budgets for
  slots with several paragraphs. A template loaded with its fonts measures text in lines and does not have this
  limit for text slots: see [text-fit.md](text-fit.md). Tables and charts are not measured.
- **Text boxes do not resize.** Text longer than its box overflows.
- **A slide holds so much.** A long text a person asked for is set smaller on its slide, down to the least
  size that reads, and then cut: at most about 440 Hebrew words on a slide 10 inches wide, and about 720
  English words on one 13.33 inches wide, in twelve paragraphs or fewer. A text of many paragraphs holds less:
  each may end a character into a line, and that line is kept for it. Only a free slide of that text alone is
  set smaller: not with free slides off, nor beside a chart, a table, a list, cards or a deck's second text. A
  slide of a deck made by 0.6.1 carries no mark that its text was asked for, so a change that writes its text
  again plans it for the usual length unless the request names one.
- **Text set dense is counted by its length.** The count knows no face: capitals and long unbroken words, such
  as web addresses, can run past the box with nothing said, and further in a wide face, where a text written at
  the top of what its writer was told can too. A list of short items, each a little shorter than a line, is
  counted two lines an item: a slide can cut it, or hold it only with its items joined, where it would fit
  whole. A slide that holds a document's text only with its paragraphs joined says so (`its paragraphs are
  joined`). See "A long text a person asked for" in [free-slides.md](free-slides.md).
- **`planEdit` says a cut after writing** unless it is given `stopBeforeCut`: then it stops before writing with
  a `TextCutError`. The chat tool `edit_presentation` stops so unless it is called with `accept_cut: true`.
- **Documents.** A deck is made from at most 12,000 characters of a document, about 1,800 Hebrew or 2,000
  English words. Several texts named together are one document, and a new document needs a new plan. Titles,
  headings and table cells are cut to their limits. A heading of 60 characters or more, with no `#` and no
  blank line under it, is read as text, and so is one right under a line of text with no stop at its end, or
  one over two short lines with no stops. A short line over one other short line is read as a heading over it.
- **Charts:** `column`, `bar`, `line` and `pie` only. No stacked or combined charts.
- **Tables:** the sample table must not have merged cells.
- **Pictures:** PNG, JPEG and GIF only.
- **Not carried over from the template:** speaker notes, named sections and revision history.
- **Animations** on template slides were not tested. A slide whose animated shape is removed may make PowerPoint
  offer to repair the file; use templates without animations, or keep animated shapes out of optional slots.
- **Not read back:** changes made to a rendered file in PowerPoint. The deck JSON is the source of truth.
- **Not possible:** positioning a shape from a deck; mirroring a slide because of a deck's reading direction
  (a mirrored layout is a separate pattern: see "Variants the code works out" in
  [layout-freedom.md](layout-freedom.md)).
- **What was tested.** Output was checked by rendering in PowerPoint on Windows. Other programs that open pptx
  files were not tested. The Gemini adapters were run against the Gemini Developer API; the Vertex AI endpoint
  is supported but was not run live.
