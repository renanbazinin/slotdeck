# The person's text comes first: a prompt for a coding agent

This is a prompt for a coding agent working in an app whose chat makes presentations with slotdeck 0.6.2, and
follows [own-text.md](own-text.md). In 0.6.2 a person who gave a document of about 1,300 words and asked for 2
slides could get both slides cut to almost nothing: the planner put a card, a figure or a step beside the text,
the slide kept it, and the deck said `the slide holds about 1: it is cut to fit`. The plan had said so before,
in `holds_words: 0`, but an app that shows plans itself or filters the tools' outputs did not show it. From
0.6.3 the person's text comes first: such a slide is its text alone, set small, and one rule says what a slide
of small text holds, from the plan to the finished deck.

Paste everything below the line to the agent. It asks for the upgrade and a check that what the tools say
reaches the person. No new tool, option or screen. Server logic only, for slotdeck 0.6.3 or later. An app on a
version before 0.6.2 follows own-text.md first, which asks for 0.6.3 too.

---

You are working in an application whose chat makes presentations with slotdeck 0.6.2. Upgrade it to 0.6.3.
In 0.6.2 a slide that carries a person's text could keep a card, a chart or a figure beside it and hold almost
none of the text; 0.6.3 gives the text the slide. Your task is the upgrade, and to make sure that what the
tools say of a cut text reaches the person before a presentation is made. Split, shorten, join or count
nothing in code of your own, and change nothing else in how the chat plans or edits.

## The short path

1. Upgrade to slotdeck 0.6.3 or later, and keep free slides on.
2. Let `holds_words`, `text_pt`, the `note` and the `warnings` of the tools' outputs through to the chat model,
   and, where the app shows a plan itself, show them beside the plan.
3. Remove any code or test that expects a 0.6.2 figure, or that works around a cut text.
4. Prove it with the tests below.

The references are "Upgrading from 0.6.2" at the top of `node_modules/slotdeck/docs/changelog.md`, "Making a
deck from a document" in `node_modules/slotdeck/docs/generating.md`, "A long text a person asked for" in
`node_modules/slotdeck/docs/free-slides.md`, and "Deck warnings" in
`node_modules/slotdeck/docs/troubleshooting.md`, not the internet.

## Find out first

Change nothing yet. Report in a few lines, then say what you will change, file by file:

1. The installed version (`node_modules/slotdeck/package.json`), and whether the app followed own-text.md
   (`documents: true`, the list of texts, the lines of the chat's system prompt).
2. The `presentationTools({ ... })` call: whether it passes `free: false` or a `patterns` list without `free`,
   and whether the template's manifest says `"free": false`.
3. Whether the app shows plans itself, and whether it filters what of the tools' outputs reaches the chat
   model. Name the fields it keeps and the ones it drops.
4. Every place in the app's code or tests that matches the text of a warning or a note, or expects a number of
   words a slide holds (such as 500 or 527 Hebrew words on one slide, or 815 English words).
5. Any code that splits, shortens, joins or resizes a person's text, or sets a font size of its own.
6. How long the app keeps plans: whether a plan made before the upgrade can be turned into a presentation
   after it.

## What to change

**1. Upgrade, and keep free slides on.** Install 0.6.3 or later the way the app installs its packages. A long
text is set smaller only on a slide the software lays out itself: if the app passes `free: false`, a
`patterns` list without `free`, or a manifest with `"free": false`, say so in your report and ask before
changing it. With free slides off, a cut slide's warning ends `; free slides are off, so its text is not set
smaller`, and the plan's `note` ends `Free slides are off, so the text is not set smaller.`

**2. What the tools say reaches the person.** Keep these fields wherever the app passes the tools' outputs on
or shows them: on a slide of a plan `words`, `passages`, `text_words`, `text_pt` and `holds_words`; on the
plan `note` and `not_on_slides`; on a presentation `document.kept_words` and `warnings`; the `note` and `next`
that `edit_presentation` answers with in place of a presentation. Where the app shows a plan, show "holds 474
of 685 words, 7 pt" beside a cut slide and the note above the plan. The chat model's system prompt already
says to pass on `note` and `warnings` if the app took own-text.md's lines; if it did not, add these two:

```
- When a plan has a slide with "holds_words", tell the user what "note" says before you create the presentation.
- When the result of a tool has "note", "not_on_slides" or "warnings", tell the user what they say before you go on.
```

**3. What changed in the numbers and the warnings.** Remove expectations of 0.6.2's figures; do not replace
them with 0.6.3's in code:

- A slide of small text holds a little less: about 2,900 characters, some 440 Hebrew words, at 7 pt on a slide
  10 inches wide (0.6.2: about 3,400 characters), and about 720 English words at 9 pt on one 13.33 inches wide
  (0.6.2: about 800 words). A text that code sets as it is, in few paragraphs, can hold more. Text set at 7
  to 9 pt is how a long text fits: it is not a bug.
- The start of every 0.6.2 warning is unchanged; some end with a new clause. Match a warning by its start, if
  at all.
- A new warning: `the text from the document for this slide (p1-p27) has 53 paragraphs, and the slide holds it
  only as 25: its paragraphs are joined. Ask for more slides to keep them apart`. No word is cut; the lines of
  a long list are joined into fewer paragraphs.
- The last column of small text can come out short.

**4. Plans and presentations made by 0.6.2.** Nothing to migrate. A plan made by 0.6.2 is written as it was
planned: one whose slides show `holds_words` far below their `words` or `text_words` still gives `holds about 1`
when it is made into a presentation. If the app keeps plans across the upgrade, say so in your report: the
person plans again (the chat model does so when asked). A presentation made by 0.6.2 whose text is longer than
0.6.3 holds may now warn `text is N characters; the slot fits about M`, and a change that writes that slide
again cuts it to what 0.6.3 holds.

## How to prove it

Offline, with the chat model and the planner scripted:

- A scripted `plan_presentation` call with `documents` naming a text longer than 2 slides hold (about 1,400
  Hebrew words on a template 10 inches wide, 1,800 English words on one 13.33 inches wide) and
  `slide_count: 2`, whose scripted outline gives each slide a text and a card (`"shows": [{ "kind": "text",
  "count": 1 }, { "kind": "card", "count": 1 }]`): the plan has 2 slides of the text alone, each with
  `text_pt`, `holds_words` well over 0 on each slide whose text is cut, and a `note`, and the chat model's next
  turn gets all three. `create_presentation` gives each cut slide a warning that says `holds about N` with N
  in the thousands, not 1.
- A call that names no text and asks no amount makes the same planner request as before the upgrade.

Then with the production model, in the language of the app's templates:

1. Paste a text longer than 2 slides hold (about 1,400 Hebrew words on a template 10 inches wide, 1,800
   English words on one 13.33 inches wide) and ask for 2 slides with all of it. The plan has 2 slides, each
   with `text_pt`, `holds_words` on each slide whose text is cut, and a `note`, and the chat model tells the
   person before it creates the presentation. Each slide of the presentation is the text alone, at 7 to 9 pt,
   inside its box in PowerPoint, and `kept_words` is close to what the slides hold added up: `holds_words`
   where a slide has it, else its `text_words`.
2. Ask for the same text on 3 slides. No slide has `holds_words`, and `kept_words` is at least 95% of the
   words.
3. Paste a list of 50 to 60 lines of a few words each (30 to 40 characters) and ask for it on one slide. The
   presentation keeps its words (count them on the slide: `document.words` leaves out lines it reads as
   headings); if its lines are joined, the chat model passes on the warning that says so. A longer list, or one
   of longer lines, can be cut: the plan then has `holds_words` and a `note` first.
4. Ask for a presentation about a subject, with no text and no numbers. The plan is as before the upgrade.

Report the arguments of every tool call, the slides of each plan with `words`, `text_pt` and `holds_words`,
`kept_words`, the notes and warnings, and what the person was told before each presentation was made.
