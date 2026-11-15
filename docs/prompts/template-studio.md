# Template studio: a prompt for a coding agent

This is a prompt for a coding agent working in an app that embeds slotdeck and has, or is getting, a template
studio: the part where a person turns an uploaded `.pptx` into a template and keeps it good.

Paste everything below the line to the agent. It covers server logic only, for slotdeck 0.3.1 or later. The
agent first reports what it found; read that before it changes anything.

The path it asks for is short: a person uploads a file, starts one operation that prepares the template by
itself, looks at one summary and a sample, and publishes. Everything else is for looking closer.

---

You are working in an application that embeds slotdeck. Your task is the logic of its template studio: what
the server does, in what order, what it stores, what it never does, and the tests that prove it. Do not design
screens. Where a person must decide, this says what they are shown and what they decide.

The reference is `node_modules/slotdeck/docs/template-studio.md` and the files beside it, not the internet.
Where a document and the types (the `.d.ts` files under `node_modules/slotdeck/dist`) disagree, the types are
right.

## The path a person takes

A person uploads a file, starts one operation, reads one short summary, looks at a sample, and publishes. That
is four operations of your server: upload (4.1), prepare and publish (4.2), and the sample (4.8).

Build this path first, and make it what an upload leads to: if the app has a draft operation today, it becomes
the prepare operation, and is not left where it is with a new one beside it. A studio that answers an upload
with thirty suggestions and a list of issues is correct and unusable: nobody reaches the end of the list.
`prepareManifest` takes those suggestions and solves those issues in one call, without a model, and returns
the few sentences that are left for a person.

Everything else in this document (taking one suggestion, changing one slot, history, replacing the file) is
for the person who wants to look closer. It stays behind the main path, never in front of it.

## 1. Find out first

Change nothing yet. Report:

1. The installed version (`node_modules/slotdeck/package.json`). Below 0.3.1, report that first and stop
   there: 0.3.0 has no `prepareManifest`, and 0.2.0 has no `suggestChanges` either, no `rows` or `widen` on
   `draftManifest`, and no `notes` or `suggestions` in its result. On agreement upgrade, or ask the person to
   if the install fails; then follow "Upgrading from" in the new `node_modules/slotdeck/docs/changelog.md`,
   run the app's tests, and resume here.
2. Where each template's `.pptx`, manifest, fonts and icons are stored; where plans and decks are, and what
   ties them to a template; whether a write can be read half done; what storage server instances share.
3. Every `loadTemplate` call, hand-built `{ pkg, manifest }` and `presentationTools` call: from which
   manifest, with which `icons` and `fields`, when it runs (build, start-up, request), and what is cached.
4. Which studio operations exist (upload, draft, change, instruct, import, export, replace the pptx, delete,
   undo), and what guards each on the server. And what a person is given after an upload and a draft: how many
   suggestions, issues and other items they must go through before a template can be used.
5. For each model call: its ceiling, its stop, what a failure leaves behind, and whether the app's
   `LlmProvider` passes `maxOutputTokens` and `signal` on, limits a call's time, and retries by itself.
6. Which rules of section 3 already hold, and which have a test.

Send the report with the rules the app breaks, most serious first, and your proposed changes. Wait for
agreement before editing.

## 2. The model

A template is a `.pptx` and a manifest, which is JSON.

- The manifest refers to the pptx: a pattern names a slide by number (`sourceSlide`), a slot a shape by id
  (`shapeId`). Editing the file in PowerPoint can change both, and rendering then fails.
- Stored plans and decks refer to the manifest by pattern id and slot name. Lose a pattern id and a deck fails
  with `unknown pattern "..."`.
- The planner chooses a pattern by its `description`. With a wrong one, a layout is never chosen.
- The chat tools hold the template they were built with. `changeManifest` and `editManifest` return a new
  manifest and leave that template as it was.

Store for each template:

- revisions, append-only: a number, the manifest, its pptx and fonts, what made it, and the `applied`
  messages;
- which revision is published, meaning the chat side uses it;
- a value for each `deckField`, passed as `fields` to `presentationTools`: a draft makes every text that
  repeats across slides a required slot with one (R6). Its first value is the template's own text (4.2), until
  a person sets theirs;
- the ids of its plans and decks. The tools' `BlobStore` has only `get` and `set` and cannot be listed: record
  `plan_id` and `presentation_id` from each tool result's `output`, and read with `storeKeys.plan(id)` and
  `storeKeys.deck(id)`.

Revisions and "published" are yours to build; slotdeck has no such fields.

## 3. Rules that always hold

**R1. A stored manifest is whole, or it is an error.** Write a revision in one step: a transaction, or a
temporary file renamed into place. An unreadable manifest is an error, never "no manifest yet"; an upload
without one is a recorded state, not an empty file. *Why:* a half-written file read as empty invites a draft
over what was there. *Check:* with a truncated manifest, operations fail clearly and no draft runs.

**R2. A write names the revision it was made from** and is refused when that is not the newest. *Why:*
otherwise the later of two writes silently erases the other. *Check:* of two changes made from one revision,
the second is refused.

**R3. A write adds a revision and removes none.** Undo restores an earlier one as a new revision. *Why:* a
removed revision cannot be brought back. *Check:* after a change and a restore there are three revisions, and
the last equals the first.

**R4. A template is drafted once.** The server refuses `draftManifest` for a manifest that has patterns,
before any model call, by whatever operation it is asked; a confirmation in the browser is not a guard. A
second start while a draft runs is refused too (section 5). To start over, upload the file as a new template.
*Why:* a model names, describes and even labels a slide differently each time, so a second draft replaces
every pattern id and description, drops every correction, and strands stored decks. *Check:* a draft request
on a drafted template makes zero model calls.

**R5. Nothing that breaks stored plans or decks is stored unseen.** If a new manifest loses a name, or a
stored deck validates worse against it, store it only after the person saw the list and confirmed. *Why:* such
decks can no longer be edited, and a removed slot's content is dropped with only a warning. *Check:*
`setPattern` with `id`, `removeSlot`, and a `max` below the `items` of a stored deck's slide are each refused
without confirmation.

```ts
const names = (m: TemplateManifest) => m.patterns.flatMap((p) => [p.id, ...p.slots.map((s) => `${p.id}.${s.name}`)]);
const lostNames = (before: TemplateManifest, after: TemplateManifest) => names(before).filter((name) => !names(after).includes(name));

// decks: this template's stored decks; layout: the chat side's
const harm = (m: TemplateManifest, deck: DeckSpec) =>
  validateDeck(m, deck, { layout }).filter((i) => i.severity === 'error' || i.message.endsWith('has no such slot; ignored')).length;
const hitDecks = decks.filter((deck) => harm(after, deck) > harm(before, deck));
```

`hitDecks` also sees a `max` below a deck's `items`, `removeRepeat` and a slot made required, which keep every
name. `lostNames` also covers stored plans, which name patterns; check a plan's `items` against the new `min`
and `max` too. No suggestion changes a name.

**R6. Only a sound manifest is published, and no other template stands in.** A revision is sound when
`checkManifest` reports no `error`, its sample renders at the fewest and at the most items, and the planner
would be offered a pattern. One is not offered while a required slot (an optional one too, unless `layout` is
`'adaptive'`) lacks its `fields` value, `icons`, photos or graphic asset on the chat side. slotdeck exports no
function for this: work it out on the revision as it is stored, and return the patterns left out with every
revision. A template that was never published is published only when a person says so, after the summary of
4.2: until then no commit publishes a revision of it, whatever made the commit. Once it has been published,
each later revision is published as soon as it is sound; one that is not is stored, the earlier one stays
published, and the person is told. A template adopted in 4.5 counts as published. With nothing published the
chat gets an error, never another template. *Why:* errors otherwise surface while a deck is made. *Check:*
import a manifest whose slot names a missing shape; the chat still gets the earlier revision. And a change
committed to a template nobody published yet does not publish it.

**R7. The chat side builds from the published revision**, never from a manifest loaded at start-up or bundled
at build time: a template shipped with the app's code only seeds its first revision (4.5). *Why:* a studio
change would never arrive. *Check:* after a changed description, the prompt of the next `plan_presentation`
call holds it.

**R8. A pptx and its manifest change together.** A revision records its pptx, and a restore brings back both.
*Why:* another file has other shape ids. *Check:* the same pptx again is accepted on confirmation; one lacking
a slide a pattern uses is not published, and the old pair still serves.

**R9. A model call has a ceiling, a time limit and a stop, and a failed one stores nothing** (section 5).
*Why:* only a limit ends a model that runs away. *Check:* tests 1 and 2 of section 7.

## 4. Operations

Every way state changes is an operation of the app's own server, and ends in one commit:

1. Refuse a stale revision (R2).
2. Refuse what is not a manifest. `loadTemplate` does not look at it, and `checkManifest` throws on JSON
   without a `patterns` array: check `version`, `name`, `direction` and each pattern's `id`, `sourceSlide`,
   `description` and `slots` yourself, and treat a throw as a refusal.
3. Run the checks of R5.
4. Store in one step (R1, R3), then decide publication (R6).
5. Return the revision, the `issues`, what was applied and refused, and whether it is published. The prepare
   operation alone returns its summary (4.2) in their place.

### 4.1 Upload

`PptxPackage.load(bytes)` throws for bytes that are not a pptx. Refuse a file with no slides
(`inspectTemplate(pkg).slides.length`). Never reuse an id, also after deletion, or a new template inherits the
old plans and decks. Tell the person the slide count: a draft makes up to one model call per slide, plus one.

### 4.2 Prepare: draft and settle, in one operation

This is the one operation a person starts after an upload. It drafts the template once (R4), then settles the
draft by itself. Asked for a template that has patterns, it does not draft: it is then the repair of 4.5.

```ts
const draft = await draftManifest(pptx, { llm: capped(llm, CEILING), name, slideImage, onProgress, signal });
if (draft.manifest.patterns.length === 0) throw new Error('no layout was found in this file'); // store nothing

// hasIcons: whether the chat side gives presentationTools an icon set for this template
const prepared = prepareManifest(loadTemplate(pptx, draft.manifest, fonts ? { fonts } : {}), { icons: hasIcons });
```

`prepareManifest` calls no model. It takes every suggestion that is a fact of the slide (`icon`, `row`,
`row-max`), lets rows widen, adds the mirrored and text-only variants and settles their rows, sets the
measured budget where a budget lets text run over its box, and takes out a pattern that cannot be used: one
for which `checkManifest` reports an error, or whose own sample does not render. With `icons: false` it keeps
the template's own pictograms in required icon slots, so that layouts with icons are still offered. It renames
nothing, and a second run with the same options changes nothing. It leaves `without` variants alone: there is
one for every optional text.

Commit `prepared.manifest` as revision 1, not yet published. Keep with it `draft.skipped`, the sentences of
`draft.notes` and `prepared.done`, `prepared.left` and `prepared.kept`. Store `prepared.fields` as the first
value of each deck field, where the chat side reads its `fields`: it is the text the template itself shows
there, and without a value the planner is offered no layout that has that field.

**Answer with a summary, not with lists.** The response to this operation holds, in this order:

1. numbers: layouts made, how many of them are added variants (`prepared.added.length`), slides skipped
   (`draft.skipped.length`);
2. `prepared.left`: the few sentences only a person can settle, such as the deck's own texts and the
   descriptions to read. When a layout was taken out, its first sentence says which and why;
3. the deck fields with their first values (`prepared.fields`), for the person to replace with their own;
4. where to get the sample at the fewest and at the most items (4.8);
5. whether it may be published: `prepared.sound`, and the patterns the planner would not be offered (R6).

`prepared.done`, `draft.skipped`, the `issues` and the suggestions that remain are returned only when asked
for. A person who has read `left` and looked at the sample can publish.

**Publish** is an operation of its own, started by a person. It works out R6 on the newest revision as it is
stored at that moment, not from a `sound` kept at prepare, which says nothing about a manifest changed since.
It refuses a revision that is not sound and says why.

When `prepared.sound` is false the revision stays unpublished, and `left` says why. After a draft that is
rare, because a pattern that cannot be used was taken out. The way on is 4.4: the sentence names the pattern,
the person corrects it or removes it with `{ op: 'removePattern', pattern }`, and publish is tried again.

Keep `rows` at its default, `'fit'`. Leave `slideImage` out unless the app can make a PNG or JPEG of a slide.
`onProgress` fires only for slides that get a call, so finish on the promise, not on `done === total`. To take
fewer things by itself, pass `variants: []` or `widen: false` to `prepareManifest`.

A slide added later becomes a pattern with `addPattern` and `addSlot`, not with a draft of `slides: [n]`,
which would name the deck fields anew.

### 4.3 Looking closer

For the person who wants more than the summary. None of it is needed before publishing a revision that is
sound.

- **Each `description`**, with its picture (`patternPreviews`). It must describe structure ("a row of cards,
  each with an icon and a short text"), not the filler's subject ("four team member profiles with photos"), or
  the planner passes the layout over for every other subject. A draft made with this version is told so; a
  manifest drafted earlier may not be. Only a person can repair it, with `{ op: 'setPattern', pattern,
  description }`. A variant made earlier has its own copy.
- **What is still suggested.** `suggestChanges(template)` calls no model and changes nothing. After
  `prepareManifest` with its default options it lists the `without` variants and nothing else; with `widen:
  false` or fewer `variants`, also what those left out. `changeManifest` accepts each suggestion's `changes`
  whole; if any change of an accepted one is rejected, store none. Remember declined ones by `kind`,
  `pattern`, `variant` and `slot`.
- **What was done.** The sentences of `prepared.done`, for a person who asks why a layout is as it is.
- **Preparing again.** `prepareManifest` does not know what a person declined or undid: it adds a variant that
  was removed again, and widens a row that was set back to its drawn width. On a template a person has worked
  on, call it with `variants: []` and `widen: false`, or with the variants and the factor they chose.

An icon slot is filled from the `icons` passed to `presentationTools`. `prepareManifest` is told with `icons`
whether there are any. Without, it sets the required icon slots to keep the template's pictogram and lists
them in `prepared.kept`; an icon slot that may be left out stays as it is. Keep that list with the template.
When an icon set is added later, those are the slots to give back to it, each with `{ op: 'setSlot', pattern,
slot, empty: 'error' }`. A slot a person set to keep is not in the list and stays kept.

### 4.4 Change

`changeManifest(template, changes)` returns `{ manifest, applied, rejected, issues }`. It skips a bad change,
gives its `reason` in `rejected`, and applies the rest. A malformed change from a client, such as `addPattern`
without `id`, can still throw: catch and refuse. Commit when `applied` is not empty, and return each `message`
and `reason`.

`editManifest(template, instruction, { llm: capped(llm, CEILING), signal })` returns the same plus `note`. It
is a proposal: show it, store nothing. On acceptance, send `result.applied.map((a) => a.change)` through
`changeManifest` on the current revision.

### 4.5 Repair a manifest in use

A manifest drafted with slotdeck 0.2.0 lets a row take no more than the number drawn: asked for five columns
on a layout drawn with four, the planner picks a table. It may also have photo slots on pictograms, and such a
layout is not offered when pictures are off.

First adopt what is stored, byte for byte, as the first revision, published as it is even where R6 would not,
so the repair can be undone. Adoption runs once per template under the lock of section 5 (propose when: for
example on the first request that finds no revision) and keeps the original until the copy reads back equal.
Record the ids of the plans and decks already in the storage behind the tools' store, or R5 sees none. A
stored manifest without patterns is an undrafted template and gets no revision.

Then it is the call of 4.2, told that the manifest is in use: `prepareManifest(template, { icons: hasIcons,
remove: false, fields: storedFields })`. With `remove: false` no pattern is taken out, so it loses no pattern
id and no slot name, and stored decks keep working. A pattern that cannot be used stays in: `sound` is then
false and `left` says what is wrong with it, for a person to correct through 4.4. `fields` are the values the
app has for the deck fields; only a missing one is asked about. Show its `left` and the number of things in
`done` as a dry run, with the lists of R5, which are empty unless something is wrong; on agreement commit its
manifest once, and R6 decides whether it is published. To repair the facts only and add no layouts, pass
`variants: [], widen: false` as well. Never run it unasked at start-up. Then a person reads the descriptions
(4.3).

### 4.6 Replace the pptx

Load the new bytes with the current manifest and compare `checkManifest`'s errors with the current pair's.

- **No new error, and the sample renders.** Show the person the new sample before they confirm: the check
  proves that slides and shapes exist, not that they are the same ones. Then commit both as one revision and
  drop every cache made from the old file.
- **New errors.** Publish nothing and return each one's `pattern`, `slot` and `message`. The person corrects
  the file, or the new pair is stored as an unpublished revision and repaired through 4.4 until R6 publishes
  it: `changeManifest` refuses only new errors. Use `removeSlot`, then `addSlot` on the new shape under the
  old `name`, then `setSlot` with the old slot's settings, which `addSlot` measures anew; a row slot needs
  `removeRepeat` first, and afterwards `setRepeat` with the old `min`, `max`, `span`, `widen` and `item`.
- **Slides reordered or redrawn.** A new template, since no operation changes `sourceSlide`. Keep the old one
  for its decks.

### 4.7 Import, export, history, delete

Export returns a revision's manifest and number; import names that revision (R2), commits against the stored
pptx, and can replace every name like a second draft, so R5 guards it. History lists the revisions; restore
commits an earlier one's manifest, pptx and fonts as the newest. Delete retires a template: out of the chat's
choice, but still loadable for edits of its decks, or such an edit is refused clearly.

### 4.8 Samples

`sampleDeck(template, { items })` fills every text slot close to its budget with the template's filler words.
Unlabelled, it is read as a broken real deck. Mark it as a sample, filled close to every limit, in the file
name and in every response that carries it, and keep it out of the store the chat uses. Render it at `items:
'min'` and `'max'` with `renderDeck(template, deck, { assets, layout: 'adaptive' })`.

### 4.9 Reaching the chat

```ts
function toolsFor(id: string) { // readPublished, cache, storeFor, iconsFor and layout are yours
  const record = readPublished(id); // throws when nothing is published (R6)
  const cached = cache.get(id);
  if (cached?.revision === record.revision) return cached.tools;
  const template = loadTemplate(record.pptx, record.manifest, { fonts: record.fonts });
  const tools = presentationTools({ template, llm, store: storeFor(id), icons: iconsFor(id), fields: record.fields, layout });
  cache.set(id, { revision: record.revision, tools });
  return tools;
}
```

Read the published revision number from shared storage on every chat request: another instance may have
written it. A running turn keeps its tools. Key icons and previews by revision too.

## 5. Model calls

- **Ceiling.** `draftManifest` puts `maxOutputTokens` on every request (a slide call asks 400 plus 60 per
  listed shape, the deck call 300 plus 40 and the bytes of each text that repeats across slides);
  `editManifest` sends none. Wrap the provider so one number, at least 4,096 and above what the largest draft
  request asks, bounds every studio call, and `LIMIT_MS` its time. A request that asks more is refused, not
  cut: a cut answer is broken JSON.

```ts
const capped = (llm: LlmProvider, most: number): LlmProvider => ({
  async generateJson(request) {
    if ((request.maxOutputTokens ?? 0) > most) throw new Error(`${request.maxOutputTokens} tokens asked, ${most} allowed`);
    const signal = AbortSignal.any([AbortSignal.timeout(LIMIT_MS), ...(request.signal ? [request.signal] : [])]);
    return llm.generateJson({ ...request, maxOutputTokens: request.maxOutputTokens ?? most, signal });
  },
});
```

- **Provider.** It must pass the number on as its output limit, with room on top for reasoning if its model
  counts that against it, and may retry a failure once, but never after an abort or a time limit. slotdeck's
  `gemini()` does so, unless `maxOutputTokens` is set in its own options, which replaces every request's
  ceiling.
- **Stop.** slotdeck only forwards `signal`; the provider must reject when it fires. Abort the run's
  `AbortController` when the person stops, the connection closes or the run fails: the other workers of a
  failed draft go on calling the model. Check `signal.aborted` before storing. A provider error rejects the
  whole run: store nothing, and say what failed.
- **One at a time.** One model operation per template, under a lock in shared storage that expires.

## 6. Fonts and measuring

Text is measured as it will be set only when the template is loaded with the metrics of its own fonts. Without
them budgets are character counts that `checkManifest` cannot try out, and `textMeasurer`, `measuredBudgets`
and `checkDesign` throw.

```ts
import * as fontkit from 'fontkit'; // installed by the app, not by slotdeck
import { fontFiles } from 'slotdeck/node';

const typefaces = templateFonts(loadTemplate(pptx, manifest)).map((font) => font.typeface);
const fonts = fontkitMetrics(fontkit, fontFiles(fontkit, fontFolder, typefaces));
const template = loadTemplate(pptx, manifest, { fonts });
const missing = textMeasurer(template).missingFonts();
```

`fontFiles` reads one folder, without subfolders, and throws when it does not exist. Load with fonts wherever
a template is built, studio and chat alike. Font files added through the app make a new revision;
`fontkitMetrics(fontkit, { [typeface]: [bytes] })` takes their stored bytes. Whether a font may be copied to a
server is a licence question for a person. Without `fontkit`, build the unmeasured path and report it. Always
state which is true: measured, measured except the `missing` typefaces, or not measured.

With fonts, `checkManifest` warns of budgets that overflow, and `measuredBudgets(template)` gives a
`suggested` budget per text slot to offer as `maxChars`.

## 7. Tests

No test calls a live model. Script one that honours `signal`:

```ts
const requests: JsonRequest[] = [];
const llm: LlmProvider = {
  async generateJson(request) {
    request.signal?.throwIfAborted();
    requests.push(request);
    return answer(request, requests.length); // your prepared answer
  },
};
```

Tell requests apart by `request.schema.properties`. The draft's deck call has `fields`: answer `{ language:
'en', fields: [] }`. A slide call has `shapes`: `{ usable: false, reason: 'scripted' }` skips the slide; a
pattern needs `{ usable, reason, id, description, figures, shapes: [{ shape, role, name, optional }] }`, with
ids from the `enum` of `shape`. `editManifest` has `changes`: answer `{ changes: [{ op, pattern, shape, slot,
values: [{ property, value }] }], note }`.

Use a template from the app's test data, or stop and ask for one. `sampleDeck(template, { items: 'max'
}).deck` serves as a stored deck, `memoryStore()` as the tools' `store`, and a scripted draft with `rows:
'drawn'` and numbered names (`card1Text`, `card2Text`) gives rows as 0.2.0 drafted them; `{ op: 'setSlot',
pattern, slot, imageKind: 'photo' }` on a pictogram's slot then gives its photo slots.

The first test to write is the main path, end to end, started by the same call an upload makes in the app:
upload, prepare with the scripted model, publish, and then the tools built the way the chat side builds them,
with the `fields` and `icons` it would pass. The prompt of the next `plan_presentation` call lists the
prepared layouts, among them one with a deck field and one with an icon slot: a template that is published and
offers the planner next to nothing fails this test. In it, check also that the response to prepare holds the
numbers, `left`, the deck fields, the sample and `sound`, and no list of suggestions or issues; that the
revision is not published before the person says so, also after a change was committed to it; that publish is
refused for a revision whose only pattern has a slot on a shape its slide does not have; and that
`prepareManifest` on the stored manifest returns an empty `done`.

Every *Check* of section 3 is a test. Add these:

1. A draft of N slides makes at most N + 1 calls. Every request, also of `editManifest`, has `maxOutputTokens`
   at or under the ceiling, and the `signal`.
2. A draft aborted midway, or whose model throws on its third call or never answers before the time limit,
   stores nothing, and the request count stops growing.
3. After repair, `max` has risen, `lostNames` is empty, `suggestChanges` lists nothing but `without` variants,
   and a deck stored before still passes `validateDeck`. With a pattern added whose slot is on a shape that
   does not exist, the repair still loses no name, the pattern is reported, and the revision is not published.
4. An instruction answered with one valid and one invalid operation makes two requests, stores nothing before
   acceptance, and returns the reason.
5. Every sample response carries the sample mark.
6. An id is not reused after a delete. The repair of 4.6 loses no name, and a restore serves its revision's
   pptx. A held lock stops a second model operation before any call, and prepare on a template that has
   patterns makes no model call. Without an icon set a repair excludes no pattern that was offered, and a
   missing `fields` value is reported.

## 8. Order of work, and done

With no studio yet, build the main path first (4.1, 4.2, 4.8, publishing, 4.9), then the rest of section 4.
With one that exists: the upgrade; then the prepare operation and its summary, in the place of today's draft
operation and in front of whatever lists the studio shows today, because that is what makes it usable. Every
template the chat uses today stays in use: adoption (4.5) publishes it as it is; then the draft guard (R4) and
storage (R1 to R3); then the repair of 4.5, which a person runs for each template; then the rest.

You are done when:

- the report of section 1 was given and agreed;
- the main path has a passing test from upload to a published template that the planner is offered layouts of,
  and the response to prepare holds a summary and no list of suggestions;
- every rule has a passing test, no test calls a live model, and no operation writes around the commit;
- the repair of 4.5 is a tested operation, and your final message says which templates a person must still
  repair and whose descriptions to read;
- your final message lists what you changed, what you left alone, and what a person must still decide.
