# Warnings and notes in the person's language: a prompt for a coding agent

This is a prompt for a coding agent working in an app whose chat makes presentations with slotdeck 0.6.3, for
people who write in a language other than English. Up to 0.6.3 everything slotdeck says is English: the app
showed the tools' warnings as they are, such as `the text from the document for this slide (p1-p14) is 3727
characters, and the slide holds about 2480 even with its text set small, at 7 pt: it is cut to fit`, and the
chat model answered a person who wrote in Hebrew with the English of a note and an English question. From
0.6.4 every warning and note has a code and values, slotdeck gives the line in English and in Hebrew
(`formatIssue`), and the tools tell the chat model to speak the user's language.

Paste everything below the line to the agent. It asks for the upgrade, the app's warnings shown with
`formatIssue`, and one line of the chat's system prompt. No new tool, option or screen. Server logic and the
place the app shows warnings, for slotdeck 0.6.4 or later.

---

You are working in an application whose chat makes presentations with slotdeck 0.6.3. Its people write in
Hebrew, and maybe other languages; slotdeck's warnings and notes are English. Upgrade to slotdeck 0.6.4, show
the warnings and notes in the language of the app's interface, and keep the chat model speaking the user's
language. Translate nothing yourself, match no warning by its text, and change nothing else in how the chat
plans or edits.

## The short path

1. Upgrade to slotdeck 0.6.4 or later.
2. Where the app shows the tools' warnings and notes, show
   `(result.issues ?? []).map((issue) => formatIssue(issue, uiLanguage, { deck: result.deck }))` in place of
   `output.warnings` and the `note`. `uiLanguage` is the language of the app's interface, such as `"he"`.
3. Keep the chat's language line: add it to the system prompt if it is not there (below).
4. Prove it with the tests below.

The references are "Warnings and notes in the person's language" in
`node_modules/slotdeck/docs/chat-tools.md`, "Warnings in a person's language" in
`node_modules/slotdeck/docs/api-reference.md`, and "Upgrading from 0.6.3" at the top of
`node_modules/slotdeck/docs/changelog.md`, not the internet.

## Find out first

Change nothing yet. Report in a few lines, then say what you will change, file by file:

1. The installed version (`node_modules/slotdeck/package.json`).
2. Every place the app shows a person what the tools said: `output.warnings`, `output.note`,
   `output.not_on_slides`, and anything that reads `Issue.message` (from `generateDeck`, `editDeck`,
   `renderDeck`, `validateDeck`). Name the file and what it shows.
3. Where the app knows the language of its interface, and whether it stores tool results (the whole
   `ToolResult`, or only `output`) to show them later.
4. Every place in the app's code or tests that matches the text of a warning or a note, compares whole `Issue`
   objects, or expects an exact `output.next`.
5. The chat's system prompt: whether it has a line on the user's language, and the lines on `note` and
   `warnings`.

## What to change

**1. Upgrade.** Install 0.6.4 or later the way the app installs its packages. Plans, decks, files and model
requests are as in 0.6.3.

**2. Show the warnings and notes with `formatIssue`.** A result that tells a person something has `issues`
beside `output`: every line of `output.warnings`, in its order, and a note for each slide the `note` speaks
of, and for each text of `not_on_slides`. Show each as
`formatIssue(issue, uiLanguage, { deck: result.deck })`. In English that is exactly the line of
`output.warnings`; in Hebrew a line about a slide begins `שקף 1: `, or `שקף 1, <the slot>: `, and names the part of a free slide in
Hebrew, and a line about the whole deck is its sentence alone. Any language slotdeck has no sentences for gets
the English. Show a Hebrew line right to left (`dir="rtl"`, or
`dir="auto"`); keep the invisible marks around labels such as `p1-p14`, which keep them reading as written.
Where the app stores results to show later, store `issues` and `deck` with them. Where it shows
`Issue.message` of `generateDeck` or `renderDeck`, use `formatIssue(issue, uiLanguage)` there too. Keep
`output` as it goes to the chat model: do not add `issues` to it.

**3. The chat model speaks the user's language.** From 0.6.4 a result with `note`, `warnings` or
`not_on_slides` ends its `next` with a sentence that says so, and the plan tool's description says to show the
plan in the user's language. Pass `next` on as it is. Add this line to the chat's system prompt, next to the
lines on `note` and `warnings`:

```
- Speak to the user in the language they write in. What the tools say in "note", "warnings" and "not_on_slides" is English: tell it in the user's language, in your own words, with the same slide numbers and figures, and never paste it as it is.
```

**4. Tests.** A test that compares whole `Issue` objects now sees `code` and `values`: compare the fields it
means. A test that expects an exact `output.next` sees the language line where the result has something to
tell. Match a warning by its `code`, never by its text.

## How to prove it

Offline, with the chat model and the planner scripted:

- A scripted `plan_presentation` call with `slide_count: 1` and 1,500 words asked for the slide (in
  `instructions`), whose scripted outline gives the slide those words: the plan has a `note`, `result.issues`
  has one issue with code `slide-holds-less`, and `formatIssue(issue, "he")` is Hebrew with the slide's number
  and words. The plan's `next` ends with `never paste its sentences as they are.`
- `create_presentation` of that plan: `(result.issues ?? []).map((issue) => formatIssue(issue, "en", { deck:
  result.deck }))` equals `output.warnings`, and in Hebrew each line is Hebrew, and begins `שקף N` where the
  English line begins `slide N`.
- A scripted `edit_presentation` that asks 1,500 words of a slide, without `accept_cut`: `output.note` is the
  English note, and `result.issues` has a `change-would-cut` issue for that slide.
- A call that has nothing to tell gives no `issues`, and the same `next` as before the upgrade.

Then with the production model, with the interface in Hebrew:

1. Paste a long Hebrew text and ask, in Hebrew, for 2 slides with all of it. The app shows the plan's note and
   the presentation's warnings in Hebrew, with the slide numbers and figures of the English. The chat model
   answers in Hebrew, in its own words, and asks in Hebrew before it creates a presentation that cuts text.
2. Ask, in Hebrew, for a larger font on a slide whose text already fills it. The chat model answers in Hebrew;
   no English sentence of slotdeck's reaches the person.
3. Switch the interface to English: the app shows the warnings exactly as `output.warnings` has them.

Report the files you changed, the lines the app showed in Hebrew and in English, and what the chat model said.
