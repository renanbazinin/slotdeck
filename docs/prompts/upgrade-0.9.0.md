# Upgrade to 0.9.0: a prompt for a coding agent

This is a prompt for a coding agent working in an app whose chat makes presentations with slotdeck 0.8.0 or
earlier. 0.9.0 changes how a slide the software lays out itself uses its room: a low row of cards, figures or
steps sets its text larger, one note alone on a slide is set as a statement, and a planner gives a low row its
bottom line. All of it is on once the package is upgraded. The prompt asks for the upgrade and the one place
where the app's own code can stand in its way.

Paste everything below the line to the agent. Server logic only, for slotdeck 0.9.0 or later.

---

You are working in an application whose chat makes presentations with slotdeck. Upgrade to slotdeck 0.9.0 and do
the few things below once. Change nothing else in how the app plans, writes or renders, and add no screen.

## The short path

1. Upgrade to slotdeck 0.9.0 or later. No template, manifest or stored deck is changed.
2. Where the app's own code copies or filters the `arrange` of a slide, let `grow` through.
3. Prove it with the tests below.

The references are "Upgrading from 0.8.0" under 0.9.0 at the top of `node_modules/slotdeck/docs/changelog.md`
and "A slide that leaves room" in `node_modules/slotdeck/docs/free-slides.md`, not the internet. An app that is
on a version before 0.8.0 does [upgrade-0.8.0.md](upgrade-0.8.0.md) first.

## Find out first

Change nothing yet. Report in a few lines, then say what you will change, file by file:

1. The installed version (`node_modules/slotdeck/package.json`).
2. Every place where the app's own code reads, copies or filters the `arrange` of a slide (a schema of the
   app's own, a list of allowed keys, an editor of a deck's JSON). Name the file and the function.
3. Every test or snapshot of the app that compares the planner's prompt, or the sizes and places of the shapes
   on a slide the software lays out (a pinned XML, a picture of a slide).

## What to change

**1. Upgrade.** Install 0.9.0 or later the way the app installs its packages. A stored deck is not changed.
Rendered again, its slides of cards, steps, figures, phases, milestones, options and notes that leave much of
their room free have larger text, and a note alone on a slide is large. On a template whose own layouts keep to
the upper part of the page, a low row sits lower on the slide. Nothing that fitted stops fitting.

**2. `arrange.grow`.** A slide's `arrange` may now hold `grow: false`: the slide keeps the template's own text
sizes whatever room it leaves. A filter on that object must keep the key. Do not set it yourself: it is there
for a person who wants one slide as it was.

**3. Nothing for the templates.** No field of a manifest is new, and no template is prepared again. On a
template whose page is a gradient that is mostly one colour (a white page with a glow in a corner), the text of
such a slide is now chosen for that colour.

## Tests

1. A slide of three short cards (`"pattern": "free"`, a heading and one line each) on the app's template passes
   `validateDeck` with no error and renders, once as it is and once with `"arrange": { "grow": false }`. No text
   of the first is smaller than in the second (compare the `sz` of the runs in the two slides' XML). Equal sizes
   are no failure: a template whose blocks are narrow, or whose text already fills the room, keeps its sizes.
   Where the template cannot carry such slides, `validateDeck` says why, as before.
2. A slide of one note alone passes `validateDeck` with no error and renders.
3. The tests and snapshots found in "Find out first" pass again, changed only where they pin the planner's line
   that offers such slides (one more sentence) or the sizes and places of their shapes.

## Report

Say which version was installed, where `grow` had to be let through, and which tests or snapshots were changed
and why. Say which of the tests above ran and what they showed. Do not say a thing was checked that was not run.
