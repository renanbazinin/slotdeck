# The person's own text and numbers: a prompt for a coding agent

This is a prompt for a coding agent working in an app whose chat makes presentations with slotdeck. Up to
0.6.1 a person who asked for "2 slides of 500 words each" could get 8 slides of 250 words, and a person who
pasted a text or attached a file got slides that sum it up: the chat model left the number out of
`slide_count` and retold the text in `topic`. From 0.6.2 a slide keeps the amount of text asked for, set
smaller where it must, on the number of slides asked for, and the chat model can name the person's text so
that all of it reaches the slides word for word.

Paste everything below the line to the agent. It asks for the upgrade, one option, a list the app builds on
every turn, and a few lines of the chat's system prompt. No new tool and no screen. Server logic only, for
slotdeck 0.6.2 or later.

---

You are working in an application whose chat makes presentations with slotdeck. People ask for numbers ("2
slides", "500 words on each"), and they paste a text or attach a file and want it on the slides. slotdeck 0.6.2
keeps both when the chat model passes them on: with `slide_count` the presentation has exactly that many
slides, an amount of text passed in `instructions` stays on its slide, set smaller where it must, and the
texts named in `documents` go on the slides word for word. Your task is to make the chat model pass these on,
and to give the tools the person's texts. Add no tool, screen or setting, split, shorten or count nothing in
code of your own, and change nothing else in how the chat plans or edits.

## The short path

1. Upgrade to slotdeck 0.6.2 or later, and keep free slides on.
2. Build the tools with `documents: true`.
3. On every turn, build the list of the conversation's texts, pass it with every tool call, and add what
   `describeDocuments(list)` returns to the chat's system prompt.
4. Put the system prompt lines below in place of the app's lines that send attached files into `topic` or keep
   slides short.
5. Let the new fields of the tools' outputs through.

The references are "Asking for an amount of text" and "Making a deck from a document" in
`node_modules/slotdeck/docs/generating.md`, "Texts the person gave" in
`node_modules/slotdeck/docs/chat-tools.md`, and "Upgrading from 0.6.1" in
`node_modules/slotdeck/docs/changelog.md`, not the internet.

## Find out first

Change nothing yet. Report in a few lines, then say what you will change, file by file:

1. The installed version (`node_modules/slotdeck/package.json`). The references come with 0.6.2: read them
   once it is installed.
2. The `presentationTools({ ... })` call, whether it passes `free: false` or a `patterns` list without `free`,
   and where the tools' `call` is made, with what of the conversation is in scope there.
3. How each kind of text a person gives reaches the server: a file whose text comes with the message (and the
   wrapper around it), a file sent to the chat model as it is (a PDF), a file kept in the app's store, text
   pasted into a message. Say for each whether the server can read its text when a tool runs.
4. Every line of the system prompt, or of the app's code, that tells the chat model what to put in `topic`,
   `instructions` or `slide_count`. Quote those about attached files, every fact the user gave, summaries,
   short slides or a number of slides.
5. One real conversation, from the logs, in which the person asks for "2 slides, 500 words each": the
   arguments of the `plan_presentation` call, and the slides and words of the plan.
6. Whether the app shows plans itself, and whether it filters what of the tools' outputs reaches the chat model.
7. What calls the model in the app's `LlmProvider`: slotdeck's `gemini()`, wrapped or not, or the app's own code.

## What to change

**1. Upgrade, and keep free slides on.** Install 0.6.2 or later the way the app installs its packages. Stored
plans and presentations need nothing. A long text is set smaller only on a slide the software lays out itself:
if the app passes `free: false`, or a `patterns` list without `free`, say so in your report and ask before
changing it. If the app's own code calls the model, not `gemini()`, which does so already, make it ask once more
with light reasoning when an answer is cut off: see "Runaway answers" in `node_modules/slotdeck/docs/providers.md`.

**2 and 3. The option and the list.**

```ts
import { describeDocuments, presentationTools } from 'slotdeck';
import type { ChatDocument } from 'slotdeck';

const tools = presentationTools({ template, llm, store, /* as before */ documents: true });

// On every turn, from the whole conversation, oldest first:
const list: ChatDocument[] = [
  { name: 'risk-review.docx', text: textOfTheFile },
  { name: 'pasted text 1', text: longMessage },
];
// describeDocuments returns '' for an empty list. The lines of step 4 are part of the app's prompt.
const system = [appPrompt, describeDocuments(list)].filter(Boolean).join('\n\n');
// On every tool call the chat model makes:
const result = await tools.call(call.name, call.arguments, { signal, documents: list });
```

What goes in the list:

- Each attached file whose text the server has, under its file name: the text that came with the message, the
  text the app keeps in its store, or the text it already makes of a PDF. Leave out a file it has no text for.
- Each message of the person's of 1,000 characters or more, as `pasted text 1`, `pasted text 2`, numbered from
  the oldest so that a message keeps its name on every turn. Take the whole message, the request in it too,
  but not the text of a file that came with it, which is listed under the file's name: count the 1,000
  characters without it. A shorter paste is not listed; the chat model puts it in `instructions` (step 4).

A text goes in as the person gave it, with its line breaks, which mark its paragraphs, lists and headings. No
two texts have the same name: number the second file of a name. Do not shorten, clean or convert a text, and
do not leave a long one out: the tools say when one is too long.

**4. The system prompt.** Take out every line that tells the chat model to put attached files, or every fact
the user gave, into `topic`, to keep slides short, or to keep presentations to a number of slides of its own.
Put these lines in their place, as they are:

```
- When the user wants their own text on the slides (a text they pasted or a file they attached), name it in "documents" of plan_presentation by the name listed for it, and say in "topic" in a sentence or two what the presentation is for. Never copy, shorten or retell it in "topic": the planner puts all of it on the slides, word for word. Name it also when the user wants a summary of it or some parts of it, and say which in "instructions".
- A shorter text the user pasted is not listed: put it in "instructions", whole and word for word, in quotation marks.
- Otherwise put every fact, figure and name the user gave into "topic".
- When the user names a number of slides, set "slide_count" to it, and when they ask for an amount of text ("500 words on each slide"), pass it on in "instructions" as they said it. Never split a slide, add slides or shorten the text yourself because it seems long: the planner sets long text smaller.
- If a plan has another number of slides than the user asked for, or a slide whose "words" is not the amount they asked for, plan again with their numbers before you show it. A slide with "holds_words" is not such a case: tell the user what "note" says before you create the presentation.
- When the result of a tool has "note", "not_on_slides" or "warnings", tell the user what they say before you go on.
- When edit_presentation answers with a "note" that a slide's text would be cut, nothing was changed: tell the user, and call it again with the same instruction and "accept_cut": true only if they accept. Otherwise change the request as they say (more slides, a shorter text).
```

**5. The outputs.** If the app filters the tools' outputs, or shows plans itself, keep these: on a slide of a
plan `words`, `passages`, `text_words`, `text_pt` (its text is set well below the template's size, at that
size) and `holds_words` (the words it keeps where its text is cut); on the plan `note`, `not_on_slides` and
`document`; on a presentation `document.kept_words`, or `document.translated` for one in another language than
the text; the `note` and `next` that `edit_presentation` answers with in place of a presentation; and the
warnings. Where the app shows a plan, show "500 words, 7 pt" beside a slide and the note above the plan.

## How to prove it

Offline, with the chat model and the planner scripted:

- A message of 1,200 characters gives a system prompt with the `describeDocuments` text and `"pasted text 1"`
  in it. Messages that are all shorter give no such text, and so does a message that is the text of an
  attached file and a short line: the file is listed, the message is not.
- A scripted `plan_presentation` call with `documents: ["pasted text 1"]` reaches the planner: its first
  request has the sentences of that message.
- A call that names no text, in a conversation with no listed text, makes the same planner request as before
  the upgrade.

Then with the production model, in the language of the app's templates:

1. "A presentation of 2 slides, 500 words each", on a subject of your choosing. The call has `slide_count: 2`
   and the 500 words in `instructions`; the plan has 2 slides with `"words": 500` and a note the chat model
   passes on; each slide of the presentation has 450 to 560 words, inside its box in PowerPoint.
2. Paste a text of about 1,000 words in two headed parts and ask for 2 slides, one for each part, with all the
   text. The call names `"pasted text 1"`, with `slide_count: 2` and a topic of a sentence or two.
   `kept_words` is at least 95% of the words, and in PowerPoint each part is on its slide, in order, inside its
   box.
3. Attach a file of about 500 words, in each format people attach (a Word file, a PDF), and ask for a
   presentation of it. The call names the file, and `kept_words` is at least 95% of its words.
4. Paste a paragraph of about 80 words and ask for it on a slide as it is. The call has it in `instructions`,
   in quotation marks, and the slide has it word for word.
5. On a plan of 8 slides, say "only 2 slides, 500 words each". The next plan has 2 slides of 500 words.
6. Ask for "one slide of 1,000 words". The plan has `holds_words` and a `note`, and the chat model tells the
   person before it creates the presentation, without planning again.
7. On the presentation of test 1, ask for "1,000 words on slide 1". The answer has a `note` and no new
   presentation; with the person's yes the next call has `accept_cut: true`, and the presentation comes with
   the warning that the rest is cut.
8. Ask for a presentation about a subject, with no text and no numbers. The call has neither `documents` nor
   `slide_count`, and the plan is as before the upgrade.

Report the arguments of every tool call, the slides and words of each plan and presentation, `text_pt`,
`kept_words`, the notes and warnings, and anything that came out shorter or longer than asked.
