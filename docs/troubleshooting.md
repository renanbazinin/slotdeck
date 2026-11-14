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
| `unknown pattern "x" (available: a, b, c)` | The slide names a pattern the manifest does not have. | Use one of the listed ids. |
| `pattern "x" takes 2 to 4 cards, not 5` | `items` is outside the pattern's range. | Use a count in the range, or raise `max` in the manifest if the template has room. |
| `this slot is required` | A required slot has no value (or an empty string), or a required deck field was not given. | Give the slot a value, or add the field to `deck.fields`. |
| `leaving this slot empty needs layout: 'adaptive'` | An optional slot was left out in `strict` layout. | Render with `layout: 'adaptive'`, or fill the slot. |
| `expected text` | A text slot got a picture, table or chart value. | Give a string or an array of strings. |
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

## Deck warnings

These do not stop rendering. `renderDeck` and the planner return them.

| Message | Meaning | What to do |
| --- | --- | --- |
| `text is 230 characters; the slot fits about 211` | Text over the slot's `maxChars`. It is rendered and may overflow its box. | Shorten the text, or raise the budget in the manifest if the sample deck shows room. |
| `text takes 7 lines at 18 pt; the box holds 5 (2 lines too many). Text of about 290 characters fits.` | (Template loaded with fonts.) The text was measured and runs over its box by that many lines. | Shorten the text to about the length given, or allow shrinking with `shrinkText`. See [text-fit.md](text-fit.md). |
| `text takes 6 lines even at 15.5 pt; the box holds 5 ...` | The text was set as small as `shrinkText` allows and still does not fit. | Shorten the text. |
| `text is 30% wider than its box at 12 pt, and the box does not wrap. ...` | (With fonts.) A one-line box whose text is too long. | Shorten the text. |
| `text has 5 paragraphs; the slot fits 4` | Over `maxParagraphs`. | Fewer paragraphs. |
| `the table has 6 columns; the slot fits 5` (also rows) | Over the table's limits. The table may leave the slide. | A smaller table. |
| `a cell is 40 characters; cells fit about 22` | A long cell makes its row taller. | Shorter cells. |
| `the chart has 12 categories; the slot fits 8` (also series) | Over the chart's limits. Labels may crowd. | Fewer categories. |
| `pattern "x" has no such slot; ignored` | The slide gives a slot name the pattern does not have. | Check the name against the manifest. With `items`, slots of items beyond the count are reported this way. |
| `pattern "x" has a fixed number of items; "items" is ignored` | `items` on a pattern without a `repeat`. | Remove `items`. |
| `chart format "x" is not known and is ignored; use "number" or "percent"` | An unknown `format`. | Use one of the two, or leave it out. |
| `chart values not stated in the brief: 89` | (Planner only.) The model charted a number the brief does not contain. It may be invented, or worked out from the brief. | Check the number. Put the figures you want charted into the brief. |

## Planner errors (`PlanError`)

| Message | Meaning | Fix |
| --- | --- | --- |
| `No pattern of template "x" can be used: ...` | Nothing in the manifest can be filled with what was supplied. The reasons follow. | See the reasons below. |
| `The model did not produce a usable outline: ...` | The model's answer did not have the expected shape. | Retry. If it repeats, check that your provider returns parsed JSON that follows the schema. |
| `Template "x" has no plain text pattern, so a text-only deck cannot be planned from it` | `textOnly` was asked for, and no pattern of the manifest is plain text. | Add a plain text pattern to the manifest (see [layout-freedom.md](layout-freedom.md)), or plan without `textOnly`. |
| `The plan uses patterns that cannot be filled with what was supplied: a, b` | A plan was made with icons, assets or fields that are not passed now. | Pass the same `icons`, `assets`, `fields` and `layout` to every step. |
| `Slide 3: pattern "x" needs icons, and no icon set was supplied` | The same, for an edit. | As above. |

Reasons a pattern is not offered:

| Reason | Fix |
| --- | --- |
| `is built around numbers, and the brief gives none` | Put figures in the brief, or set `numberPatterns: 'always'`. |
| `needs a photo, and images are turned off` | Give an `ImageProvider` or photo assets, and an `imageDensity` other than `none`. |
| `needs icons, and no icon set was supplied` | Pass `icons`. |
| `needs a graphic asset, and none was supplied` | Pass an asset of kind `graphic`. |
| `needs the deck field "x"` | Add it to `brief.fields`. |

## Manifest problems (`checkManifest`)

| Message | Fix |
| --- | --- |
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
| warning: `no font was supplied for Cambria bold; text set in it is checked by its character budget` | Add the font file to `fontkitMetrics`, or accept character budgets for those slots. |

## Refusals from `changeManifest`

`rejected[i].reason` is one of the manifest messages above, or one of these:

| Reason | Meaning |
| --- | --- |
| `there is no pattern "x"` / `"p" has no slot "x"` / `slide 4 has no shape #99` | The change names something that does not exist. |
| `5 cards do not fit side by side: each is 21% of the slide wide and the row has 93%. The most is 4` | The row cannot hold that many. Use the stated maximum. |
| `a grid holds at most the 6 cards it draws; 8 would need a new line of them` | A grid's `max` is the number drawn. |
| `"title" is a text slot; maxColumns does not apply to it` | A setting for another kind of slot. |
| `maxChars must be a whole number of 1 or more` | A budget of zero, a fraction or a negative number. |
| `"card1Text" is part of the row of cards, so it cannot be renamed, moved or tied to a deck field; remove the row first` | Row slots must stay alike. Use `removeRepeat` first if you really need to. |
| `shape #4 is a picture; it can be photo, icon, graphic, chart, decoration` | The role does not suit the shape. |
| `"x" has no row of like items. Name the slots of each item alike with a number ...` | `setRepeat` found no numbered slots to make a row from. |
| `another pattern is already called "x"` | Choose another id. |
| `a template needs at least one pattern` | The last pattern cannot be removed. |

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
| `Gemini's answer was cut off at the limit of 9552 tokens, twice. ...` | The model wrote until the output limit, twice in a row. It is nearly always a model running away inside one text field (see "Runaway answers" in [providers.md](providers.md)). Try again. If it keeps happening, set `thinkingLevel: 'low'`; if the limit is too small for a legitimately large answer, set `maxOutputTokens`. |
| `Gemini request failed (400): Request blocked due to copyright/recitation content` | The service's recitation filter tripped on an ordinary answer. It happens at random. The adapter already sent the request twice; try again. |
| `Gemini request failed (400): Thinking level ... is not supported for this model` | The `thinkingLevel` set on the adapter is not one the model accepts (`minimal` is refused by some models). Use `low`, or leave it unset. |
| `Gemini request failed (429): Your project has exceeded its monthly spending cap` | The key's project reached the spending cap set for it. Nothing works until the cap is raised in the provider's console. |
| `Gemini returned no image` | The image model answered without a picture, often because the prompt was refused. |

## The result looks wrong

| Observation | Likely cause | Fix |
| --- | --- | --- |
| Text runs out of its box or over a neighbour | The slot's budget is too generous, or the deck ignored a budget warning. | Render `sampleDeck` to see every slot at its budget, then lower `maxChars` with `changeManifest`. |
| A leftover sample element appears (a sticker, a second page number, a sample caption) | The shape is not a slot, so it is copied as it is. | Mark it as a decoration: `{ op: 'addSlot', pattern, shape, role: 'decoration' }`. |
| A picture is cropped badly | `cover` crops to the frame. | Supply a picture closer to the frame's proportions, or set `fit: 'contain'` on the slot. |
| Cards are in reverse order in a right-to-left deck | The pattern's items are not declared as a row. | Add a `repeat` (`setRepeat`) or `sequences`. |
| Hebrew or Arabic text is left-aligned | The manifest keeps alignment. | Set `alignOnDirectionChange: 'flip'` on the manifest or the slot. |
| A font changes for another script | The template's font lacks that script. | Set a complex-script font in the template. |
| Body text looks small | The engine keeps the template's own font sizes, and the template was designed with small body text (9 pt on a 10-inch slide is common in decks exported from Google Slides). | For every deck: set `textScale` on the manifest, such as `1.3`. For one deck: `textSize: 'large'` on the brief, `text_size` in the chat tool, or `textScale` on the deck. See "Larger text" in [layout-freedom.md](layout-freedom.md). |
| The planner never uses a pattern | It is not offered. | See "Reasons a pattern is not offered". Also check the pattern's `description`: the model chooses by it. |
| A deck that was asked to be "text only" has cards, tables or charts | The wish did not reach the planner as a wish about slides: only "no pictures" was passed (`imageDensity: 'none'`, or `image_density` in the chat tool), which still allows cards, tables and charts. | Set `textOnly: true` on the brief, or `text_only: true` in `plan_presentation`. See "Asking for a kind of slide" in [generating.md](generating.md). |
| A text-only deck uses one layout for every content slide | The template has one plain text pattern besides its opening and closing. | Add more plain text patterns to the manifest as derived patterns: a list without its photo, a quotation, two paragraphs side by side. |
| A slide states a number or contact detail that nobody gave | The model invented it. | Figures are guarded only for patterns marked `figures` and for charts. State details in the brief and review the deck. |

## Known limits

- **Shapes need their own position.** The engine measures the shapes on a slide. A shape that takes its position
  and font size from a layout placeholder, without stating them on the slide, is filled correctly but measured
  as empty: a drafted budget for it is only as long as its sample text, and `box` cannot be applied to it. Decks
  exported from Google Slides and decks built from free-standing text boxes work as they are.
- **Budgets are estimates.** A character budget cannot see lines lost to wrapping. Leave room in budgets for
  slots with several paragraphs. A template loaded with its fonts measures text in lines and does not have this
  limit for text slots: see [text-fit.md](text-fit.md). Tables and charts are not measured.
- **Text boxes do not resize.** Text longer than its box overflows.
- **Charts:** `column`, `bar`, `line` and `pie` only. No stacked or combined charts.
- **Tables:** the sample table must not have merged cells.
- **Pictures:** PNG, JPEG and GIF only.
- **Not carried over from the template:** speaker notes, named sections and revision history.
- **Animations** on template slides were not tested. A slide whose animated shape is removed may make PowerPoint
  offer to repair the file; use templates without animations, or keep animated shapes out of optional slots.
- **Not read back:** changes made to a rendered file in PowerPoint. The deck JSON is the source of truth.
- **Not possible:** swapping the side of a photo; positioning a shape from a deck.
- **What was tested.** Output was checked by rendering in PowerPoint on Windows. Other programs that open pptx
  files were not tested. The Gemini adapters were run against the Gemini Developer API; the Vertex AI endpoint
  is supported but was not run live.
