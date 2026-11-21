# Upgrade to 0.7.0: a prompt for a coding agent

This is a prompt for a coding agent working in an app whose chat makes presentations with slotdeck 0.6.9 or
earlier, often in Hebrew on a right-to-left template, and whose decks are mailed and read. 0.7.0 changes what a
reader of such a deck sees: charts can read from the right, a highlighted card stays on the right option in an
English deck, the file shows the deck's title and slide count instead of the template's, nothing of the
template's sample (a link, a comment, a broken animation) is left in it, and a screen reader can use it. The
chat tools carry their own system-prompt lines, and model calls ride out a short rate limit.

Paste everything below the line to the agent. Fill in the table first: which templates' charts should read from
the right is a choice, not a fact of the template. Server logic only, for slotdeck 0.7.0 or later.

---

You are working in an application whose chat makes presentations with slotdeck. Upgrade to slotdeck 0.7.0 and do
the few things below once. Change nothing else in how the app plans, writes or renders, and add no screen.

The templates to change:

| Template | Charts read from the right (`chartDirection: 'rtl'`) | Lists that spread (`spread`) | Tables that fill their frame (`fill`) |
| --- | --- | --- | --- |
| (the name of the template in the app) | (yes or no) | (the patterns, such as agenda, icon-list, numbered-points, with 1.6) | (the table patterns, such as table, table-text) |

## The short path

1. Upgrade to slotdeck 0.7.0 or later.
2. In the chat's system prompt, replace the lines about the presentation tools that the app wrote by hand with
   `tools.instructions`, which the tools now carry.
3. For each template above, one change through the studio's own path for changing a manifest (below), and
   prepare the template again so that a row with a card that stands out is marked.
4. Prove it with the tests below.

The references are "Upgrading from 0.6.9" under 0.7.0 at the top of `node_modules/slotdeck/docs/changelog.md`,
"The chat model's system prompt" in `docs/chat-tools.md`, "Charts that read right to left" in `docs/languages.md`
and "File properties" in `docs/rendering.md`, all under `node_modules/slotdeck/`, not the internet.

## Find out first

Change nothing yet. Report in a few lines, then say what you will change, file by file:

1. The installed version (`node_modules/slotdeck/package.json`).
2. Where the chat's system prompt is built, and which of its lines are about the presentation tools: a number of
   slides, the person's own text, cuts, the language, the plan before the presentation, ids. Quote them.
3. Where the app builds `presentationTools`, `gemini` and `geminiImages`, with which options.
4. Where the app keeps each template's manifest, the server logic it changes one with (`changeManifest`,
   `editManifest`, or the studio's own operation), and the operation that prepares a template
   (`prepareManifest`). Name the file and the function.
5. Every test or snapshot of the app that compares a rendered deck, a chart's or a picture's `descr`, a
   system prompt, or a model call that fails or is retried.

## What to change

**1. Upgrade.** Install 0.7.0 or later the way the app installs its packages. Nothing breaks. A stored deck that
renders again changes only in what a screen reader is told (its charts, tables and icons) and in the language
mark of a chart with Hebrew names, unless its template is changed below.

**2. The system prompt.** `presentationTools(...)` returns `instructions` beside `definitions`: the lines for the
chat model's system prompt, made from the same options. Join them to the app's own prompt
(`[appPrompt, tools.instructions].join('\n\n')`) and remove the app's own lines that say the same thing, so that
nothing is said twice or differently. Keep the app's lines that are about the app itself, such as its name, its
tone, and `describeDocuments`.

**3. Decks that are read.** The app's decks are mailed and read, not presented: build the tools with
`delivery: 'read'`. Slides are then written in full sentences, with no closing slide of thanks or questions, and
an agenda only in a deck of more than ten slides.

**4. File properties.** The tools now write what each file says of itself: its title, language, slide count and
the version of slotdeck, without the template's cover as its thumbnail. Nothing to do. Where the app knows who
asked for the deck, pass `properties: { author: <their name> }` to `presentationTools`. Do not set `embedDeck`
unless asked: it keeps the deck's JSON in the file, which then goes wherever the file is mailed.

**5. Retries.** `gemini()` and `geminiImages()` now try a request again after a passing failure (a rate limit, a
server error, a network that failed), two more times by default, and pictures too. Where a person waits for the
answer, give both adapters `deadlineMs: 120000`, so that a call gives up after two minutes in all. Remove any
retry the app wrapped around slotdeck's calls itself, or the tries multiply.

**6. The templates.** For each template in the table, through the studio's existing path for a change to a
manifest, so that it is stored, checked and published as any other change:
- `{ op: 'setManifest', chartDirection: 'rtl' }` where its charts should read from the right;
- `{ op: 'setRepeat', pattern: <id>, spread: 1.6 }` for each list named;
- `{ op: 'setSlot', pattern: <id>, slot: <the table slot>, fill: true }` for each table named.
Then run the studio's preparation again (`prepareManifest`), or apply
`suggestChanges(template, { kinds: ['row-emphasis'] })`: it marks a row whose slide draws one card unlike the
others (`repeat.emphasis`), so that in an English deck that card is still the first. Load the template again
after the change, where the app keeps it loaded.

**7. Warnings.** The planner now gives more warnings, with codes and Hebrew sentences: details the person did not
give (`text-details-unstated`: an email, a phone number, "24/7", a percentage), stand-ins such as "XX%"
(`text-stand-in`), text about the slide itself (`text-about-slide`), and a word too long for its line
(`word-wider-than-line`). The app shows warnings with `formatIssue` already; check that the new codes reach the
person like the others.

## Tests

- The chat's system prompt contains `tools.instructions` once, and none of the lines it replaced.
- A Hebrew deck with a column chart, rendered with the changed template: in `ppt/charts/chart1.xml` the category
  axis has `<c:orientation val="maxMin"/>`, and the chart's names are marked `lang="he-IL"`. An English deck from
  the same template has `minMax`.
- An English deck of the layout whose card stands out: the card in the other colour holds the first card's text.
- A deck's file: `docProps/app.xml` says the deck's number of slides, `docProps/core.xml` its title, and there is
  no `docProps/thumbnail.jpeg`.
- A fake fetch that answers 503 with `Retry-After: 0` and then succeeds: the call succeeds, with two requests.
- A test that pinned a chart's `descr`, a system prompt, or a call that fails on the first 503 sees the new
  behaviour: update it, and say so in your report.
