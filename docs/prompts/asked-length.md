# An amount of text: a prompt for a coding agent

This is a prompt for a coding agent working in an app whose chat makes presentations with slotdeck. Up to
0.6.0 a person who asked for "200 words" on a slide got what the layout said it holds: 76 words, on a template
in Hebrew. From 0.6.1 the planner reads the amount, gives the text a slide that holds it (a layout, a free slide
set out for that length, or several slides that share it), and tells the writer the length.

Paste everything below the line to the agent. It is on by default, so the prompt asks for little: the upgrade,
and the few places where the app's own code or prompts can stand between the person's wish and the planner. No
new tool and no new setting. Server logic only, for slotdeck 0.6.1 or later.

---

You are working in an application whose chat makes presentations with slotdeck. slotdeck 0.6.1 gives a slide
the amount of text a person asks for ("200 words on our expansion", "make slide 2 about 150 words"). The person
does nothing new: they ask in the chat, in their own words, and the planner reads the amount. Your task is to
upgrade, and to make sure nothing in the app shortens that wish or takes away the room the planner needs for
it. Do not add a tool, an argument, a setting or a screen for it.

The short path: upgrade. With free slides on, which they are by default, a long text then works in the chat with
nothing else changed. The rest of this prompt is the few places where the app's own code or prompts can work
against it.

The references are "Asking for an amount of text" in `node_modules/slotdeck/docs/generating.md`,
`node_modules/slotdeck/docs/chat-tools.md` and "Upgrading from 0.6.0" in
`node_modules/slotdeck/docs/changelog.md`, not the internet.

## Find out first

Change nothing yet. Report in a few lines, then say what you plan to change, file by file:

1. The installed version (`node_modules/slotdeck/package.json`). Below 0.6.1 the upgrade is step 1 below; the
   references above come with it, so read them once it is installed.
2. The `presentationTools({ ... })` call, or the app's own calls of the planner: whether it passes `patterns`,
   and whether it passes `free` (and from where).
3. Every text of the app's own that the chat model or the planner reads: the chat's system prompt, and anything
   the app adds to a request's `instructions` or `topic`. Quote every line that limits the text of a slide:
   "keep slides short", "at most 40 words a slide", "bullets only", "summarise the user's request".
4. How a plan and a presentation are shown to the person: which keys of the tools' outputs the app reads, and
   whether the warnings of `create_presentation` and `edit_presentation` reach the chat model.

## What to change

**1. Upgrade.** Install slotdeck 0.6.1 or later the way the app installs its packages, and read "Upgrading
from 0.6.0" in its changelog. Stored plans and presentations need nothing.

**2. Let free slides hold a long text.** A slide of one long text is a free slide: one text set at the
template's running-text size, in two columns where it needs them. If the app passes `patterns`, free slides are
offered only when that list names `free`: add `'free'` unless the list exists to forbid them. If the app turns
free slides off (`free: false`), a text longer than a layout holds goes over several slides where the
presentation may have them, and is otherwise cut to what the layout holds, with a warning. Leave free slides on
unless there is a reason in the app's configuration to turn them off.

**3. Let the wish through.** The chat model passes a wish for an amount of text on in `instructions` (or in
`topic`), in the person's words; the planner reads it from there. A line of the app's own prompts that caps the
text of a slide makes the chat model shorten or drop that wish. Change such a line so that it gives way to the
person, for example: "Keep slide text short unless the user asks for an amount of text; then pass their words on
in instructions as they said them." Do not add a number of your own.

**4. Show the words of a plan.** A slide of a plan that was asked an amount of text has `"words"` in its line of
the `plan_presentation` and `edit_presentation_plan` outputs: the number of words its text will have. If the
app shows plans itself, show it beside the slide ("about 200 words"). A text that goes over several slides has
its share on each, and their purposes end in "(1/3)", "(2/3)", "(3/3)".

**5. Pass the warnings on.** Two warnings are new, and the chat model should tell the person both:
`200 words were asked for this slide, and it holds about 76: ask for the text over more slides, or for a shorter one`,
and, after a change, `the 200 words asked for slide 2 go over slides 2 to 4: one slide holds about 76`. If the
app filters warnings before the chat model sees them, let these two through.

## How to prove it

In the language of the app's templates, with the production model:

- Ask the chat for a presentation of one slide with 200 words about a subject of your choosing. The plan has a
  slide with `"words": 200`; the presentation's slide has between 180 and 230 words, inside its box when it is
  opened in PowerPoint, with no warning about its words.
- On that presentation, ask to make the slide about 100 words. The slide is rewritten with about 100.
- Ask for a slide of 600 words. The plan has three slides of 200, whose purposes end in "(1/3)" to "(3/3)".

Report the words counted on each slide, the warnings, and anything that came out shorter than asked.
