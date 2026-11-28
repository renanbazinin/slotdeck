# Style-only templates: a prompt for a coding agent

This is a prompt for a coding agent working in an app whose chat makes presentations with slotdeck, and where a
file becomes a template only after its manifest is prepared. People upload a deck and want more decks "that look
like this one", with no preparation. slotdeck 0.10.0 has that: `styleManifest` makes a manifest from the file
alone, by rules and with no model, and a deck takes the file's cover, its closing slide and, between them,
slides the software lays out in the file's typefaces, sizes and colours.

Paste everything below the line to the agent. Server logic only, for slotdeck 0.10.0 or later.

---

You are working in an application whose chat makes presentations with slotdeck. Add one more way for an uploaded
file to become a template: by its style alone. Change nothing else in how the app prepares templates, plans,
writes or renders, and add no screen.

## The short path

1. Upgrade to slotdeck 0.10.0 or later.
2. Where a file is uploaded as a template, add a second path beside the one the app has: call
   `styleManifest(bytes, { name })`. With a `reason`, store nothing and tell the person that sentence. Without
   one, store the manifest, `patterns` and `fields` with the template, marked as a template by style.
3. Where the presentation tools are built for a template so marked, give them `patterns` from what was stored
   and `layout: 'adaptive'`.
4. Prove it with the tests below.

The reference is `node_modules/slotdeck/docs/style-only.md`, and "Upgrading from 0.9.0" under 0.10.0 at the top
of `node_modules/slotdeck/docs/changelog.md`, not the internet. An app on a version before 0.9.0 does the
upgrade prompts of the versions between first.

## Find out first

Change nothing yet. Report in a few lines, then say what you will change, file by file:

1. The installed version (`node_modules/slotdeck/package.json`).
2. The operation an upload leads to today: what it calls (`draftManifest`, `prepareManifest`), what it stores for
   a template (the file, the manifest, the deck fields, whether it is published), and what a person is shown.
3. Where the app builds `presentationTools`, and with which options: above all `patterns`, `layout`, `fields`,
   `icons` and `assets`.
4. How a template's deck fields reach a deck: where their values are stored and how a person gives them.
5. Every test of the app that pins the sentence `the template's running text is in placeholders`, or a list of
   the templates that offer slides the software lays out.

## What to change

**1. Upgrade.** Install 0.10.0 or later the way the app installs its packages. A template that carried slides
the software lays out before has the same styles, stored decks and rendered files. One thing changes without a
line of code: a template whose running text is all in placeholders was refused such slides, and now carries
them, so a planner is offered them on it. To keep such a template as it was, set `"free": false` in its
manifest.

**2. The upload.** Which path a file takes is the person's choice, or the app owner's: add it as an option of
the upload operation the app has, and leave the default as it is today.

```ts
import { styleManifest, loadTemplate, checkManifest } from 'slotdeck';

const style = styleManifest(bytes, { name });
if (style.reason) {
  // No template by style from this file. Say the sentence, and offer the usual preparation.
} else {
  // Store with the template: style.manifest, style.patterns, style.fields, and that it is a template by style.
}
```

- It calls no model and takes a fraction of a second. There is no draft to wait for and no list to review: the
  template is ready when it is stored, and is published as the app's own rules say.
- `style.fields` are the texts a deck gives once for all its slides, such as a unit's name in the header. Each
  has a name (`field1`, `field2`) and the sample the file shows there. Store them the way the app stores a
  template's deck fields, with the sample as what a person is asked to replace. A field with no value shows
  nothing on the slides.
- `style.notes` are sentences about what was decided. Log them; show them where the app shows a template's
  details.
- Do not run `draftManifest`, `prepareManifest`, `suggestChanges` or the studio's review on this manifest, and do
  not offer its layouts for editing. Its layouts are the file's slides as they are drawn, kept for their text
  styles. A person who wants the file's own layouts takes the usual preparation instead.
- The same file gives the same manifest. It may be made again when the file is loaded, or after an upgrade.

**3. Making presentations.** For a template by style:

```ts
presentationTools({ template: loadTemplate(bytes, manifest), patterns, layout: 'adaptive', llm, store, /* icons, assets, fields as for any template */ });
```

`patterns` is the stored list: `cover`, `free`, `closing`, without the cover or the closing slide where the file
has none. Where the app already gives the tools a list of layouts a person limited a deck to, give the layouts
that are in both. Everything else is as for a prepared template.

**4. The person's language.** `reason` and `notes` are English sentences. Where the app shows the reason, put
its own sentence in the person's language before it ("This file cannot be used by its style alone"), and keep
the reason as the detail.

## Tests

1. An ordinary deck, a title slide and a few slides of a title and bullets, taken from the app's test files or
   made with PowerPoint: `styleManifest` gives no `reason`, `patterns` is `['cover', 'free', 'closing']`, and
   `checkManifest(loadTemplate(bytes, style.manifest))` has no error.
2. A presentation made on it through the app's own path, from a brief of a few sentences: the deck passes
   `validateDeck` with no error, its first slide is on `cover`, its last on `closing`, and every slide between
   them has `blocks`. Open the file once and look at it.
3. A file with a title slide and nothing else: `styleManifest` gives a `reason`, nothing is stored as a template
   by style, and the person is told.
4. A prepared template the app already has: a stored deck renders to the same bytes as before the upgrade.
5. The tests found in "Find out first" pass again, changed only where they pin the refusal for placeholders.

## Report

Say where the option was added, what is stored for a template by style, and how its `patterns` reach the tools.
Say which of the tests above ran and what they showed, and what the deck of test 2 looked like. Do not say a
thing was checked that was not run.
