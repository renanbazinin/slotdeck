# Chat tools

`presentationTools` packages the planner as four tools for a chat model ("function calling"). The chat model
talks to the person; the tools do the work; the person approves a plan before anything is written.

```ts
import { presentationTools, describeDocuments, TOOL_NAMES, storeKeys, memoryStore } from 'slotdeck';
import { directoryStore } from 'slotdeck/node';
```

## The four tools

| Tool name | What it does | Model calls inside |
| --- | --- | --- |
| `plan_presentation` | Decides the slides, the layout of each and the point each makes. Writes nothing yet. | 1; 2 when a named text only asks for a presentation |
| `edit_presentation_plan` | Changes a plan by instruction. Returns a new plan with a new id. | 1; none for a colour alone |
| `create_presentation` | Writes the approved plan's content, makes its pictures, renders the file. | 1, plus repairs, plus one per photo; more for a free slide of over nine parts |
| `edit_presentation` | Changes a created presentation by instruction. Returns a new presentation with a new id. | 1 or 2, plus repairs; none for a colour alone; more when many parts of a free slide of over nine parts change |

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
  review,                                       // optional DeckReview: a second look at every draft (design-checks.md)
  free: true,                                   // optional: free slides, for content no layout shows; on unless false (see below)
  countSpacing: true,                           // optional: a layout's budget told with its template's spacing counted (generating.md); on unless false
  composeLayouts: 2,                            // optional: new layouts for such slides in place of free slides; needs free: false (generating.md)
  patterns: ['cover', 'text-only', 'free'],     // optional: the only layouts the tools may use
  colors: true,                                 // optional: whether a person may ask for another colour; false leaves accent_color and color_slides out
  thumbnail: (patternId) => `/static/thumbnails/${patternId}.png`, // optional address of a picture of each layout
  preview,                                      // optional PreviewProvider: draws every presentation the tools make
  newId: (kind) => `${kind}_${crypto.randomUUID()}`, // optional; default is a short random id
  onProgress: (event, tool) => console.log(tool, event.stage), // optional
  documents: true,                              // optional: the plan tool takes the names of texts the person gave (see below); off unless true
});
```

The `llm` here is the model the tools use to plan and write. It can be the same model as the chat model or a
different one.

### Free slides

A slide can hold what no layout of the template has: six charts and a note, twelve cards, a table beside two
figures. The software lays such a **free slide** out itself ([free-slides.md](free-slides.md)). The tools plan
them unless the `free` option is `false`, when the template can carry them, its manifest does not say
`"free": false`, the deck fields their frame shows are supplied, and the `patterns` option, if you pass one,
names `free`. Then:

- `plan_presentation` says in its description that a slide can hold what no layout has (several charts or
  tables on one slide, any number of cards or figures), and that such a wish goes in `instructions` in the
  person's words. It names pictures among them only when the tools can put one on a slide: with an `images`
  provider, or with a photo or a graphic among `assets`.
- The `layouts` argument takes `free` beside the template's layouts, so a person can ask for free slides only,
  or leave them out.
- A plan and a presentation show such a slide with `"layout": "free"` and `"free_layout": true` (see below).
- `edit_presentation` says in its description that several charts or tables can be put on one slide, and
  pictures under the same condition.
  It changes a free slide part by part: a part the person does not speak of stays exactly as it is, also when
  parts are added, taken off or put in another order beside it (see "Free slides in an edit" in
  [generating.md](generating.md)).

While free slides are on offer the `composeLayouts` option makes nothing. A plan or a presentation that has a
free slide is written, changed and shown whatever the option says.

**Build the tools once for a loaded template.** Whether the template can carry free slides is worked out when
`presentationTools` is called, from the template's slides: about 0.2 s for the demo template and half a
second for one of 35 layouts. It is remembered for the loaded template, so a second call with
the same `template` object costs nothing. An app that loads the template anew and builds the tools for every
chat turn pays it on every turn, also on turns that call no tool: load the template once and keep it. With
`free: false`, or a `patterns` list that does not name `free`, nothing is worked out.

### Texts the person gave

A person pastes a report or attaches a file and wants it on the slides. With `documents: true` the plan tool
takes `documents`: the names of such texts. The slides then carry every part of them, in their own words and
order, and the planner sets a long text smaller where it must. Without the option the chat model can only
retell the text in `topic`, and the slides sum it up.

The app builds the list of the conversation's texts on every turn, passes it with every tool call, and puts
the line `describeDocuments` makes into the chat's system prompt:

```ts
const list: ChatDocument[] = [
  { name: 'risk-review.docx', text: textOfTheFile },   // an attached file whose text the app has, under its name
  { name: 'pasted text 1', text: longMessage },         // a message of 1,000 characters or more, whole
];
const system = [appPrompt, describeDocuments(list)].filter(Boolean).join('\n\n');
const result = await tools.call(call.name, call.arguments, { signal, documents: list });
```

Build the list from the whole conversation, oldest first: each file whose text the app has, under its file
name, and each message of the person's of 1,000 characters or more as "pasted text 1", "pasted text 2", so
that a message keeps its name on every turn. A message's length and text leave out the text of a file that
came with it, which is listed under the file's name. Take a text as the person gave it, with its line breaks,
which mark its paragraphs, lists and headings. No two texts have the same name. `describeDocuments` lists the
last ten, each with its words and first words, and returns `''` for an empty list.

A shorter paste is not listed. With the option, the `instructions` argument says that a text of the person's
that the app does not list goes there whole and word for word, in quotation marks, and the error for a name
the list does not have says so too.

The plan tool reads a name exactly, then without regard to case, then as the beginning of one listed name
("plan 2027" finds "Plan 2027.docx"), each time in one Unicode form and without direction marks: a file name
in decomposed letters is found by the same name composed. It reads only the list on the call: the tools are
built once, and the list belongs to the conversation. Any number of texts may be named. A text named twice, or
under two names (a file and the message it came in), is carried once: a named text whose words all come in a
row in another named text is left out. A name it cannot find, an empty text and more than 12,000
characters in all are errors the chat model can act on (see "Errors"). A new text for an existing plan or
presentation needs a new plan: the tools that change one take no `documents`.

When no text is named, the plan tool checks `topic` and `instructions`, each on its own, against each listed
text of 1,500 characters or more. A shorter message is most often the request itself. An argument that copies
a listed text (120 or more runs of four words from it, at least half of the argument's runs) is an error that
names the argument and tells the chat model to name the text instead. A topic and instructions that retell
one (40 or more words, three quarters of them words of the text) plan, and the output says so in
`not_on_slides`.

A named text of which the planner puts no part on a slide only asks for the presentation: a long request,
listed and named. The plan is then made again with that text as `instructions` and no document, which is a
second planning call, and `next` begins `"pasted text 1" only asks for a presentation, so the planner took it
as instructions, and no slide carries it word for word.`

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
const result = await tools.call(call.name, call.arguments, { signal }); // and documents, with that option
// result.output  -> send this back to the chat model as the tool's result (JSON, small, no binary data)
// result.plan    -> the DeckPlan, when the tool made or changed one (for your own UI)
// result.deck    -> the DeckSpec, when the tool made or changed one
// result.pptx    -> the file bytes, when one was rendered (offer it as a download)
```

`call` does not throw for expected failures. A missing id, an unknown tool or a model failure comes back as
`{ output: { error: "..." } }`, in words the chat model can act on. It throws only when `signal` was aborted.
Only the four declared names are tools: any other name, and anything that is not a string, is an unknown tool.

## Arguments and results of each tool

### `plan_presentation`

Arguments:

| Name | Type | Required | Meaning |
| --- | --- | --- | --- |
| `topic` | string | yes | What the presentation is about, with every fact, figure and name the person gave. Numbers not stated here will not appear on slides. |
| `slide_count` | integer 1 to 30 | no | Number of slides in all. The chat model must set it whenever the person names a number of slides: the planner does not read one out of the topic or the instructions, and chooses the number without it. With it the presentation has exactly that many slides: long text is set smaller, and no slide is added. Where no free slide of text is on offer, the description says instead that a slide holds what its layout holds, and that the plan says what is cut. A number over 30 is taken as 30. |
| `language` | string | no | BCP 47 tag, such as `he` or `en`. Defaults to the template's language. |
| `image_density` | `none`, `low`, `medium`, `high` | no | How many slides carry a photo. Default `medium`. Forced to `none` when no photos can be made. |
| `text_size` | `normal`, `large`, `x-large` | no | Size of body text. Default `normal`. |
| `text_only` | boolean | no | `true` when the person asks for a presentation of text only. Every slide is then plain text: a title with paragraphs or a list, with no pictures, icons, cards, figures, tables or charts. Only the template's plain text layouts are used, and free slides of text, lists, quotations and notes where free slides are on offer. Not together with `layouts`: the layouts named then decide, and the plan is not marked text only. |
| `instructions` | string | no | Everything else the person asked for, in their own words: audience, tone, every amount of text they ask for ("500 words on each slide"), and every wish about how slides look or are arranged that has no argument of its own ("three columns", "no cover slide"). With the `documents` option, also a text of the person's for the slides that the app does not list (a short paste), whole and word for word, in quotation marks. |
| `documents` | array of strings | no | Only with the `documents` option: the names of the texts of the conversation the presentation is made from, as the app lists them. The slides carry every part of them in their own words. Named also for a summary or some parts of a text, which `instructions` then says. `topic` is then a sentence or two on what the presentation is for, and `language` the language of the text unless the person asks for another. |
| `accent_color` | string | no | Only when the person asks for a colour: the main colour as six hex digits, such as `1E6B3A`, in place of the template's. |
| `layouts` | array of layout ids | no | Only when the person limits the layouts: the only layouts the presentation may use, the first and last slides too. The ids on offer are the layouts these tools can plan with, and `free` where free slides are on offer: a layout the `patterns` option leaves out, or one that needs icons, photos, a graphic or a deck field the app did not supply, is not among them. Naming one all the same is an error that says which and why. A number layout named for a topic without a figure is an error too. |

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

`thumbnail` is present only when the `thumbnail` option returns an address. With the `composeLayouts` option, a
slide that uses a layout made for this presentation has `"new_layout": true` and no thumbnail, since the app has
no picture of it; its description is the new layout's.

A free slide has no layout to show a picture of either. Its line says what it holds:

```json
{ "number": 3, "layout": "free", "layout_description": "A slide laid out for this presentation: 6 charts and a note.", "purpose": "Every store at a glance.", "free_layout": true }
```

A plan and a presentation that have a colour carry `"accent_color"`, and one limited to some layouts
`"layouts"`; in a presentation, a slide with a colour of its own has `"accent_color"` in its line (`"template"`
where it keeps the template's inside a coloured presentation).

A slide the person asked an amount of text for ("200 words on the expansion", passed on in `instructions` or
`topic`) has `"words"` in its line: the words its text will have. The plan gives such a slide the room for the
text on that slide: a layout that holds it, or a free slide set out for it, set smaller where it must. A slide
asked for words is never spread over more slides. See "Asking for an amount of text" in
[generating.md](generating.md).

```json
{ "number": 2, "layout": "free", "layout_description": "A slide laid out for this presentation: a text.", "purpose": "Our expansion north.", "words": 200, "free_layout": true }
```

A plan with words or a text the person gave says, before anything is written, what each such slide will carry.
Each of these is present only when it applies:

| Field | Where | Meaning |
| --- | --- | --- |
| `passages` | slide | The parts of the person's text the slide carries, as `"p3-p5"`. |
| `text_words` | slide | The words of those parts. |
| `text_pt` | slide | The size its running text is set at, in points, where that is under four fifths of the template's size. |
| `holds_words` | slide | The words it keeps, where even the least size that reads cannot hold its text: the rest is cut. |
| `document` | plan | `{ "passages": 9, "words": 491, "left_out": "p1" }`: the person's text, and the parts no slide carries. |
| `note` | plan | What the chat model is to tell the person before the presentation is made: the slides that are cut, then those set in much smaller type. |
| `not_on_slides` | plan | The listed texts the topic retells, which no slide will carry word for word. `next` then says so first. |

```json
{
  "slides": [{ "number": 1, "layout": "free", "purpose": "Our plan.", "words": 1000, "text_pt": 9, "holds_words": 725, "free_layout": true }],
  "note": "Not all the text fits: slide 1 holds about 725 of its 1000 words (text at 9 pt). The rest will be cut. Tell the user before creating the presentation, and offer more slides or a shorter text."
}
```

Where free slides are off (`free: false`, `"free": false` in the manifest, or a `patterns` list without
`free`), a cut slide's text is not set smaller, and the `note` ends: `Free slides are off, so the text is not
set smaller.` So does the `note` of a change that stops before writing (see `edit_presentation`).

Where a slide has words and the call had no `slide_count`, `next` begins: `No slide_count was given, so the
planner chose 6 slides. If the user named a number of slides, plan again with slide_count.` A plan with a
`note` or `not_on_slides` has `next` end with the language line (from 0.6.4; see "Warnings and notes in the
person's language"):

```
Tell the user what this result says in the language they write in, in your own words and with the same slide numbers and figures; never paste its sentences as they are.
```

### `edit_presentation_plan`

Arguments:

| Name | Type | Required | Meaning |
| --- | --- | --- | --- |
| `plan_id` | string | yes | The plan to change. |
| `instruction` | string | no | What to change, in the person's words, with any new facts they gave. Text the person gives for a slide goes in whole and word for word, in quotation marks. It may be left out or empty when `accent_color` is given: the tool then writes the colour to the plan and calls no model. With neither, the call is an error. |
| `accent_color` | string | no | Only when the person asks for a colour: six hex digits, or `template` to return to the template's colours. |
| `layouts` | array of layout ids | no | Only when the person changes which layouts may be used. Slides that are planned already keep their layout unless the instruction says otherwise. An empty array lifts the limit; an argument that names no layout sets nothing. |

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
  "warnings": ["slide 4 body: text is 230 characters; the slot fits about 211"],
  "next": "Tell the user what this result says in the language they write in, in your own words and with the same slide numbers and figures; never paste its sentences as they are."
}
```

`file` is the store key of the rendered pptx. `warnings` is present only when there are any, and so is `next`,
which then tells the chat model to say them in the user's language (from 0.6.4; see "Warnings and notes in the
person's language").

A presentation made from a text the person gave has `"document": { "words": 491, "kept_words": 488 }`, with
`"left_out"` naming the parts no slide carries. A slide that holds its part of the text has it set word for word
by the software. A presentation in another language than the text has `"translated": true` in place of
`kept_words`: the words of a translation are not the text's, so none are counted.

A slide that holds less than was asked for it, or less than its part of the person's text, is cut, and
`warnings` says so, as the plan's `note` did before: `slide 1: 1000 words were asked for this slide, and it holds
about 446 even with its text set small, at 7 pt: the rest is cut. Ask for the text over more slides, or for a
shorter one`. So does a slide written well short of the words asked (`... and it has about 380: ask for it again
to have them all`), and one that kept too few of the words of the person's text. The chat model passes these
on. See "Deck warnings" in [troubleshooting.md](troubleshooting.md).

A free slide is listed with what it is made of, and a warning about one of its parts names the part by its
number on the slide and its kind:

```json
{
  "slides": [{ "number": 3, "layout": "free", "title": "Sales by store", "free_layout": true, "parts": "6 charts and a note" }],
  "warnings": ["slide 3, part 7 (note): text is 230 characters; the slot fits about 150", "slide 3: 2 parts had no content and were left out"]
}
```

With the `preview` option set, the output also has `"previews"`: the store keys of the pictures of the slides,
such as `["preview/deck_851fbb9677/slide-01.svg", ...]`. If the pictures could not be made, it has
`"preview_error"` with the reason instead, and the presentation is created all the same.

### `edit_presentation`

Arguments:

| Name | Type | Required | Meaning |
| --- | --- | --- | --- |
| `presentation_id` | string | yes | The presentation to change. |
| `instruction` | string | no | What to change, in the person's words. Text the person gives for a slide goes in whole and word for word, in quotation marks. It may be left out when `accent_color` is given: the tool then recolours and renders, and calls no model. |
| `accent_color` | string | no | Only when the person asks for a colour: six hex digits, or `template` to return to the template's colours. With `color_slides`, `template` returns those slides to the template's colours, also inside a presentation that has a colour of its own. |
| `color_slides` | array of integers | no | The slides to colour, by number, when the colour is for some slides only. Left out, the whole presentation takes it. It needs `accent_color`: without one the call is an error. The numbers are those of the presentation before the change; a slide the same change removes is named in `warnings`. |
| `layouts` | array of layout ids | no | Only when the person changes which layouts may be used. Slides that are there already keep their layout unless the instruction says otherwise. An empty array lifts the limit; an argument that names no layout sets nothing. |
| `accept_cut` | boolean | no | Only after the person accepted that a slide's text is cut, when an earlier call with the same instruction answered with a `note` saying so. Left out, a change that would cut a slide's text is not made. |

Output: the same shape as `create_presentation` with a new `presentation_id`, plus:

```json
{ "summary": "Shortened the body text on slide 2 and made the title punchier.", "rewritten_slides": [2], "previous_presentation_id": "deck_851fbb9677" }
```

`rewritten_slides` are the slides that were written again, and the free slides whose parts or arrangement were
changed with nothing to write: a part taken off, three side by side.

An instruction that asks for an amount of text on a slide ("make slide 2 about 200 words") gives that slide
the room for it: a layout that holds the text, or a free slide set out for it, set smaller where it must. No
slide is added for it. Where even the least size that reads cannot hold the text, nothing is changed and
nothing is stored. The output is then:

```json
{
  "note": "Nothing was changed: slide 2 would hold about 725 of its 1500 words (text at 9 pt). Ask for more slides or a shorter text, or accept that the rest is cut.",
  "next": "Nothing was written. Tell the user what the note says, and ask them, in their language, whether they accept that the rest is cut. If they accept, call edit_presentation again with the same instruction and accept_cut: true; otherwise change the request (more slides, a shorter text). Tell the user what this result says in the language they write in, in your own words and with the same slide numbers and figures; never paste its sentences as they are."
}
```

So does any other change that would leave a slide less of its text than it carries, such as a part added
beside a long text. A change that asks no length of a slide and keeps its parts and passages (other words, a
new title, another language) is never stopped: its text is written to what it holds.

With `accept_cut: true` the change is made, the slide holds what it can, and `warnings` says the rest is cut.

A change that was planned for a part of a free slide and did not come back from the writer is in `warnings`,
such as `slide 3: nothing was written for 1 part, and it stays as it was`. So is a change the plan spoke of
without marking any part of the slide to be written:
`slide 3: nothing on it was changed: the plan of the change named no part to write again` (the slide is then
not among `rewritten_slides`). The `summary` is
written before the slides are, so it may say that the change was made: have the chat model pass such a warning
on, and ask again naming the part.

Two kinds of warning say what is so on a free slide, and not that a change failed. A number side by side that
the slide cannot hold is set to the nearest number it can, and the result says
`slide 3: 12 side by side were asked for; the slide holds 5`. And a chart that holds more than reads well at
its size is named, as in `slide 3, part 2 (chart): the chart has 3 series; the slot fits 2`: that is said of
the presentation before a change as after it, and an edit that is not about the chart leaves it as it is.

Two things an edit does not do on a free slide. It does not hand a slide that is set out some way (three side
by side, the main part at the left) back to the software's own choice unless the same instruction changes
something else on that slide: a number of columns or a side that the person names is always set. And it changes
the title of a slide of ten parts or more but does not remove it. See "Free slides in an edit" in
[generating.md](generating.md).

### Errors

```json
{ "error": "There is no plan with id \"plan_x\"" }
```

Other messages: `"topic" is required`, `"instruction" is required unless a colour is given`,
`There is no presentation with id "..."`, `Unknown tool "..."`, and the messages of `PlanError` and
`DeckValidationError` (see [generating.md](generating.md)).

With the `documents` option, `plan_presentation` answers these before any model is asked:

| Case | Message |
| --- | --- |
| A name the list does not have | `"documents" names a text the app does not have: "budget.xlsx". The texts the app lists for this conversation are "pasted text 1" and "Garden notes.docx": use those names. A text the user pasted in a short message is not listed: pass it in instructions, whole and word for word, in quotation marks. A file that is not listed could not be read: tell the user.` With an empty list the second sentence is `The app lists no text for this conversation.` |
| An empty text | `The text "notes.txt" is empty: tell the user it could not be read.` |
| Too much text | `The text is 15230 characters, and a presentation is made from at most 12000 (about 1,800 words). Ask the user to paste the part the presentation is about, and name that.` |
| A topic or instructions that copies a listed text | `"topic" copies the text "pasted text 1". Name it in documents instead and keep topic to what the presentation is for: the planner reads the text word for word. If the user wants a presentation about it or a summary of it, name it all the same and say so in instructions.` It begins `"instructions" copies` where that argument is the copy. |

A named text that only asks for a presentation is no error: the plan is made again with it as instructions
(see "Texts the person gave").

What your own parts throw comes back as its message, unchanged: the `llm`, the `images` provider, the `store`
and the `review`. A network that fails (`fetch failed`) is told to the chat model in those words.

What the store holds may not be what the tools put there: a file cut short, or JSON that is no plan. The
software then fails on it with a `TypeError`, a `RangeError` or a `SyntaxError`, whose text
(`Cannot read properties of undefined`) says nothing a chat model can act on. The chat model is told one of
these instead:

| Tool | Message |
| --- | --- |
| `create_presentation`, `edit_presentation_plan` | `The stored plan could not be read. Make a new plan with plan_presentation.` |
| `edit_presentation` | `The stored presentation could not be read. Make it again: a new plan with plan_presentation, then create_presentation.` |
| `plan_presentation` | `The presentation could not be planned: the software failed on this request. Tell the user that it did not work.` |

The same is said when the software fails this way for another reason, such as a template or an option that is
not what `presentationTools` takes. The error itself is not in the result: to see it, run the step outside the
tools (`writeDeck` with the stored plan, `planEdit` with the stored deck).

## Warnings and notes in the person's language

What the tools say is English: the `note`, the `warnings` and `not_on_slides`. A person who writes in another
language gets it in theirs in two ways, from 0.6.4.

**The app shows them.** A result that tells a person something has `issues` beside `output`: every line of
`output.warnings`, in its order, and a note for each slide the `note` speaks of (a cut slide
`slide-holds-less`, a slide in small type `slide-set-small`, a slide a change would cut `change-would-cut`)
and for each text of `not_on_slides` (`text-retold`). Each has a `code` and its `values`. `issues` is not in
`output`, so it costs the chat model nothing. Show them with `formatIssue`:

```ts
const result = await tools.call(call.name, call.arguments);
const lines = (result.issues ?? []).map((issue) => formatIssue(issue, uiLanguage, { deck: result.deck }));
```

In English a line is exactly the line of `output.warnings`. In Hebrew the line of a slide cut to fit is:

```
שקף 1: הטקסט מהמסמך לשקף הזה (קטעים p1-p14) באורך 3727 תווים, והשקף מכיל כ-2480 תווים גם בגופן קטן של 7 נק': הטקסט נחתך כדי שייכנס. אפשר לבקש יותר שקפים כדי שהטקסט יישאר שלם
```

The label `p1-p14` stands between the invisible marks U+2068 and U+2069 (see
[api-reference.md](api-reference.md)); they are left out here.

Show a Hebrew line right to left (`dir="rtl"`, or `dir="auto"`: it begins with a Hebrew word). Any other
language gets the English. The codes and their values are in "Warnings in a person's language" in
[api-reference.md](api-reference.md).

**The chat model says them.** A result with `note`, `warnings` or `not_on_slides` ends its `next` with one
sentence, and has a `next` of that sentence alone where it had none (`create_presentation`,
`edit_presentation`):

```
Tell the user what this result says in the language they write in, in your own words and with the same slide numbers and figures; never paste its sentences as they are.
```

The `next` of a change that stops before a cut asks the chat model to ask the person, in their language,
whether they accept it. The plan tool's description says to show the plan in the language the user writes in.
A result with nothing to tell has `next` as before. The English `note`, `warnings` and their order are as in
0.6.3.

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
| `plan/<id>.json` | A plan: the brief and the slide list, and the person's text when one was named | 1 KB, plus the text |
| `deck/<id>.json` | A deck: all text, and for each picture its id and description, and the person's text when one was named | 4 KB, plus the text |
| `asset/<id>` | One generated picture | tens of KB to 1 MB |
| `pptx/<id>.pptx` | The rendered file | hundreds of KB and up |
| `preview/<id>/slide-01.svg` | A picture of one slide, when the tools have a `preview` provider. The extension follows the provider: `svg`, `png` | tens of KB each |

`storeKeys.plan(id)`, `storeKeys.deck(id)`, `storeKeys.asset(id)`, `storeKeys.pptx(id)` and
`storeKeys.preview(id, slide, extension)` build these keys, so your app can read what the tools saved:

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

A plan lists a layout per slide. To show layouts as pictures, make one picture for each pattern once, name it
after the pattern, and return its address from the `thumbnail` option. The address is passed through to the
chat model's output and is not fetched by slotdeck.

`patternPreviews(template, preview)` makes those pictures without PowerPoint, through a preview provider; see
[preview.md](preview.md). Pictures of the template's own slides from any other tool work as well.

## Showing a presentation

Set the `preview` option to a `PreviewProvider` and every presentation the tools create or change is drawn.
The pictures go to the store under `preview/<presentation id>/`, their keys are in the tool's output, and the
`ToolResult` carries the pictures themselves in `previews` for your app to show. How to get a provider, and how
close its pictures are to PowerPoint, is in [preview.md](preview.md).

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

- When the person wants their own text on the slides, have the chat model name it in `documents` and keep
  `topic` to what the presentation is for. A shorter paste, which is not listed, goes in `instructions`, whole
  and word for word, in quotation marks. Otherwise put every fact, figure and name the person gave into
  `topic`. A number that is in none of these is not used.
- Set `slide_count` whenever the person names a number of slides, and pass an amount of text on in
  `instructions` as they said it. The chat model should not split slides or shorten text itself: the planner
  sets long text smaller.
- When a tool's result has `note`, `not_on_slides` or `warnings`, the chat model tells the person before it goes
  on. It sets `accept_cut` only after the person accepted the cut.
- The chat model speaks the language the person writes in. What the tools say is English: it retells it in its
  own words, in the person's language, with the same slide numbers and figures, and never pastes it as it is.
  The results say so in `next` (from 0.6.4); a line in the system prompt makes it hold for the whole
  conversation.
- Pass wishes about the kind of slides on, in the person's words ("no tables", "only cards"), in `topic` or
  `instructions`. For a presentation of text only, set `text_only: true`: "no pictures" alone
  (`image_density: "none"`) still allows cards, tables and charts.
- A slide can hold several charts or tables, or any number of cards or figures, and several pictures where your
  app has an image provider or pictures of its own. When the person asks for that ("all six charts on one
  slide", "twelve logos on one page"), pass the wish on in `instructions` in their words, with the number they
  said. Do not answer that the template has no such layout.
- Every number a chart is to show must appear in `topic`. A chart the person gave no numbers for is left off the
  slide, and the result says so in `warnings`: pass that on. The check looks for each number of a chart in what
  the person gave; it does not tell whether the number stands in the right chart.
- Show the plan and wait for approval before calling `create_presentation`.
- For a change to a finished presentation, call `edit_presentation` with the person's words; do not plan again
  from scratch. That goes for a slide of many parts too: "add a chart for the new store to slide 3" or "take the
  second picture off" changes those parts and leaves the others as they are. Give the numbers of a new chart in
  the instruction.
- Refer to presentations and plans by id only. Never ask for or produce deck JSON.
