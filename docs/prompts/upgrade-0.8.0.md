# Upgrade to 0.8.0: a prompt for a coding agent

This is a prompt for a coding agent working in an app whose chat makes presentations with slotdeck 0.7.1 or
earlier. 0.8.0 changes what a reader of a deck sees: the words that carry a point are set in bold, a slide the
software lays out itself can draw a process as a chain of arrows, a timeline and a comparison that picks one
option out, and such a slide can stand on a band of the template's colour. All of it is on once the package is
upgraded. The prompt asks for the upgrade and the few places where the app's own code can stand in its way. It
also covers an app that has a folder of its own icons and pictures and gives the presentations none of them, or
reads the folder with code of its own: 0.8.0 reads such a folder.

Paste everything below the line to the agent. Fill in the two lines first, or leave them empty where the app
has no such folder. Server logic only, for slotdeck 0.8.0 or later.

---

You are working in an application whose chat makes presentations with slotdeck. Upgrade to slotdeck 0.8.0 and do
the few things below once. Change nothing else in how the app plans, writes or renders, and add no screen.

The app's own pictures:

| What | Folder on the server |
| --- | --- |
| Icons | (the folder, or nothing) |
| Photographs, and pictures to show whole (diagrams, logos) | (the folders, each with its kind, or nothing) |

## The short path

1. Upgrade to slotdeck 0.8.0 or later. Nothing breaks, and no template, manifest or stored deck is changed.
2. Wherever the app's own code shows or reads the text of a slide, take the marks of a strong phrase into
   account (below). Where it does not, there is nothing to do.
3. Where a folder is named above: read it once when the server starts, with `iconFiles` and `assetFiles`, and
   give the lists to the presentation tools as `icons` and `assets`.
4. Prove it with the tests below.

The references are "Upgrading from 0.7.1" under 0.8.0 at the top of `node_modules/slotdeck/docs/changelog.md`,
"Strong phrases" in `docs/deck-format.md`, and "A process, a timeline and a comparison" and "A band of colour"
in `docs/free-slides.md`, and "Your own icons and pictures" in `docs/generating.md`, all under
`node_modules/slotdeck/`, not the internet.

## Find out first

Change nothing yet. Report in a few lines, then say what you will change, file by file:

1. The installed version (`node_modules/slotdeck/package.json`).
2. Every place where the app's own code reads the text of a deck's slots: to show a slide's text outside the
   pptx (a list of slides, a preview in HTML, a search index), to count it, or to compare it. Name the file
   and the function.
3. Every place that lists the kinds of part of a free slide, or names them for a person: a table with one
   entry for each kind, a translation of their names, a filter on a plan's or a deck's fields.
4. Where the app builds `presentationTools`, and with which options: above all whether it is given `icons` and
   `assets`, and from where.
5. For each folder named above: how many pictures it holds, in which formats (count the PNG, JPEG, GIF, SVG and
   other files), whether it has folders inside, and whether a file in it says what each picture shows.
6. Every test or snapshot of the app that compares a writer's system prompt, the schema of a planning call, or
   the list of kinds a planner is offered.

## What to change

**1. Upgrade.** Install 0.8.0 or later the way the app installs its packages. A stored deck renders to the file
it rendered to. A deck written from now on may hold strong phrases, the three new kinds of part and a band of
colour: all of them are in the deck's JSON, and render with no option.

**2. Strong phrases in the app's own code.** A text of a deck may now hold a phrase between two pairs of
asterisks, `**like this**`, which the pptx shows in bold. The marks are in the deck's JSON. Where the app shows
such a text itself, show the phrase in bold or take the pairs off; where it counts or compares a text, count it
without them, as slotdeck does. Asterisks that are no such pair are text. Do not strip them before the deck is
stored or rendered: the renderer needs them.

**3. The kinds of part.** `FREE_BLOCK_KINDS` has three more: `phase`, `milestone` and `option`. A table of the
app with an entry for each kind needs theirs, and a list that filters the kinds of a plan must let them
through. Their parts are `b1Heading` and `b1Text` (a phase), `b1When`, `b1Heading` and `b1Text` (a milestone),
`b1Badge`, `b1Heading` and `b1Items` (an option). A deck's `arrange` may hold `tone` (`"accent"` or `"tint"`):
a filter on that object must keep it. Warnings and notes in the app's language (`formatIssue`) name the new
kinds and parts already.

**4. Only if the app's owner asks for plain text.** `emphasis: false` on `presentationTools` (or on the
planning calls the app makes itself) tells the writer to use no bold. Left out, the writer is told to use it
sparingly. Do not set it unless you were asked to.

**5. The app's own icons and pictures, where a folder is named above.** Read each folder once, when the server
starts or when the template is loaded, and keep the lists in memory:

```ts
import { iconFiles, assetFiles } from 'slotdeck/node';

const skipped: string[] = [];
const icons = iconFiles(ICONS_FOLDER, { onSkipped: (file, reason) => skipped.push(`${file}: ${reason}`) });
const assets = [...assetFiles(PHOTOS_FOLDER, 'photo'), ...assetFiles(DIAGRAMS_FOLDER, 'graphic')];
// presentationTools({ template, llm, store, icons, assets, ... })
```

- Give the same two lists wherever the app makes a file from a stored deck by itself (`resolveAssets`, then
  `renderDeck`): a deck stores the ids of its pictures, not the pictures.
- Every PNG, JPEG and GIF of a folder is a picture, known by its file name without the ending. That id is kept
  in every deck: do not rename files that decks already use. Folders inside the folder are read with
  `recursive: true`.
- An SVG or a WebP is not read. Log `skipped` once at start, and report how many files were passed over and
  why. Do not convert pictures in the app: say in your report that those files have to be saved as PNG, and
  let the app's owner decide.
- A model chooses an icon by its description alone. Where the folder has no file that says what each icon
  shows, write one beside the pictures, `icons.json` (`assets.json` in a folder of photographs or diagrams): an
  object of id and one line, what the icon shows and what it stands for, such as
  `"clock": "a clock: time, schedule, duration"`. Take the lines from what the app already knows of its icons
  (a list, their names); where all you have is a file name, write a plain line from it and list those icons
  in your report as needing a person's look. Never invent what a picture shows without having something that
  says so.
- Where the app read the folder with code of its own before, keep its ids exactly as they were, so that
  stored decks still find their pictures, and replace the code only if the lists come out the same.
- A template whose layouts have icon slots was prepared with or without an icon set
  (`prepareManifest(template, { icons })`). If the app gets its first icon set here, say so in your report:
  the templates prepared without one keep their own pictograms until they are prepared again.

## Tests

1. A deck with `"body": "Passes grew **34%** this year"` on a layout of the app's template renders, its
   check (`validateDeck`) counts 25 characters, and the slide's XML holds a run with `b="1"` whose text is
   `34%`.
2. A free slide of three phases, one of three milestones and one of three options (one with a badge) on the
   app's template each pass `validateDeck` with no error and render. Where the template cannot carry free
   slides, `validateDeck` says why, as it did before: that is no failure of the upgrade.
3. The app's own display of a slide's text, where it has one, shows no asterisks for a strong phrase.
4. Where a folder is named above: the lists hold as many pictures as the folder has PNG, JPEG and GIF files,
   every description is a line of words, and a presentation made from a brief that calls for cards shows icons
   of the folder on a layout with icon slots. A stored deck that names an icon renders again after a restart.
5. The tests and snapshots found in "Find out first" pass again, changed only where they pin the writer's
   system prompt (one more rule), the schema of an outline (one more property, `shows_on`) or the list of kinds.

## Report

Say what the app's own code does with a strong phrase now, which lists of kinds were extended, and whether
`emphasis` was set. For each folder of pictures: how many were read, how many were passed over and why, and
which descriptions a person should look at. Say which of the tests above ran and what they showed. Do not say a thing was checked that
was not run.
