# Larger text on free slides: a prompt for a coding agent

This is a prompt for a coding agent working in an app whose chat makes presentations with slotdeck 0.6.6 or
earlier. People see slides of running text set small, at 9 or 10 pt, where 12 would fit. One case is a long text
of their own, in three columns with the last part empty. Up to 0.6.6 slotdeck counted a character as wide as
any face may set it. From 0.6.7 a template can say how wide its text is set (`characterWidth` in its manifest),
and such text is set larger: by 0.5 to 1.5 pt on a typical slide, and up to 4. The count still keeps a margin
of 2 to 3 pt. Headings of a document are counted as they are set, and a long text takes three columns sooner.

Paste everything below the line to the agent, with the value you measured for each template. The value comes
from a PC with PowerPoint (`scripts/measure-character-width.ts` in slotdeck's repository, see "How wide a
template's text is set" in `docs/text-fit.md`), never from the agent. The prompt asks for the upgrade and one
setting stored in each template's manifest. No new tool, option or screen. Server logic only, for slotdeck 0.6.7
or later.

---

You are working in an application whose chat makes presentations with slotdeck. Upgrade to slotdeck 0.6.7, and
store a measured `characterWidth` in the manifest of the templates named below. Measure nothing yourself, guess
no value, and change nothing else in how the app plans, writes or renders.

The values to set, as measured on a PC with PowerPoint:

| Template | `characterWidth` |
| --- | --- |
| (the name of the template in the app) | (the value measured, such as 0.45) |

## The short path

1. Upgrade to slotdeck 0.6.7 or later.
2. For each template above, run the app's own server logic that changes a manifest, with
   `changeManifest(template, [{ op: 'setManifest', characterWidth: <value> }])`, and store the manifest it
   returns where the app keeps that template's manifest, as the studio stores any change.
3. Prove it with the tests below.

The references are "How wide a template's text is set" in `node_modules/slotdeck/docs/text-fit.md`, the
`characterWidth` row in `node_modules/slotdeck/docs/manifest-format.md`, and "Upgrading from 0.6.6" at the top
of `node_modules/slotdeck/docs/changelog.md`, not the internet.

## Find out first

Change nothing yet. Report in a few lines, then say what you will change, file by file:

1. The installed version (`node_modules/slotdeck/package.json`).
2. Where the app keeps each template's manifest, and the server logic it changes one with: `changeManifest`,
   `editManifest`, or the studio's own operation. Name the file and the function.
3. Whether the app keeps a cached, loaded template (`loadTemplate`) that has to be loaded again once its
   manifest changes, and where.
4. Every test or snapshot of the app that compares a rendered deck, a plan's budgets or the size of a text with
   a stored copy.

## What to change

**1. Upgrade.** Install 0.6.7 or later the way the app installs its packages. A deck stored before renders as it
did, with two exceptions, on every template:
- a slide whose long text set dense two columns did not hold at some size is now set larger, in three columns;
- a slide whose text has a document's headings (`SlideSpec.headings`, from 0.6.6) is counted with its headings
  set run in, and is set at the same size or larger.

Plans of a document's text with headings, and of long texts set dense, change the same way. Everything else is
as in 0.6.6 until a template states its width.

**2. Store the width with each template.** Use the studio's existing path for a change to a manifest, so that the
change is stored, checked and published as any other: one more change of kind `setManifest` with
`characterWidth`. Do not write the field into the JSON by hand, and do not add a screen for it. Where the studio
changes a manifest in words, `editManifest` takes it too ("set the character width to 0.45").
`checkManifest` warns of a value out of range (0.3 to 0.8), which slotdeck then does not use. Load the template
again after the change, where the app keeps it loaded.

**3. Tests.**

- A deck of one free slide with a long text, rendered with the template before and after the setting: after,
  its text is set larger, and `validateDeck` reports nothing new.
- The stored manifest has `characterWidth`, and a manifest without it renders as before.
- A test that compares a plan's budgets or a rendered deck with a stored copy sees larger budgets and sizes
  for a template that now has the setting, and for document decks with headings and long texts set dense on any
  template: update the copy, and say so in your report.
