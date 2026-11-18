# Prompts

Prompts to hand to a coding agent that works on an application built with slotdeck. Each file is a complete
prompt: paste its text to the agent, or tell the agent to read it at
`node_modules/slotdeck/docs/prompts/<file>`. They are written for an agent that works without internet access
and reads the rest of this documentation from the package.

| Prompt | Hand it over when |
| --- | --- |
| [one-step-preparation.md](one-step-preparation.md) | The app has a template studio, and a person who uploads a template is met by lists of suggestions and issues. The prompt asks for one operation that prepares a template by itself, a short summary to look at, and publishing. It is short, and the place to start. |
| [layout-variants.md](layout-variants.md) | The studio prepares and publishes templates, but a person cannot add a variant of one layout by themselves: its mirror image, a text-only version without its picture, or a version without an optional line. The prompt asks for an operation that lists the variants a layout allows and one that makes one, also for a template the chat already uses. Server logic only. |
| [compose-layout.md](compose-layout.md) | A person needs a layout the template does not have (a text-only slide, three points, an agenda), and a variant of an existing layout cannot give it. The prompt asks for one operation that makes such a layout from words with `composePattern`, a look at its sample, and keeping it or not; from 0.5.0 also from a layout the chat made for a presentation. Server logic only. |
| [deck-layouts.md](deck-layouts.md) | The chat makes presentations, and a text-only one repeats the same layout because the template has no other. The prompt asks for one setting on the path the chat already takes (`composeLayouts`), so that the planner makes layouts for those slides in the template's look, a mark for them in the plan, and keeping the deck whole. No new tool. Server logic only, 0.5.0 or later; on 0.6.0 only for an app that keeps free slides off, since free slides take that place. |
| [free-slides-and-colors.md](free-slides-and-colors.md) | The app is to be upgraded to 0.6.0, where a presentation can hold slides the software lays out itself (twelve charts, a chart beside three cards, one paragraph alone), other colours than the template's, and a list of layouts a person limits it to. Almost all of it is on by default. The prompt asks for the upgrade, four places where the app must not trip over the new things, and two switches in the studio. Server logic only, 0.6.0 or later. |
| [asked-length.md](asked-length.md) | The app is to be upgraded to 0.6.1, where a slide gets the amount of text a person asks for ("200 words on our expansion"), where 0.6.0 wrote what the layout said it holds. It is on by default. The prompt asks for the upgrade and the few places where the app's own code or prompts can stand between that wish and the planner: a list of layouts without `free`, a prompt line that caps slide text, the warnings. Server logic only, 0.6.1 or later. On 0.6.2, use own-text.md. |
| [own-text.md](own-text.md) | People ask for "2 slides of 500 words each" and get eight slides of 250 or two of 270, or paste a text or attach a file and get slides that sum it up. The prompt asks for the upgrade to 0.6.2, one option, the list of the conversation's texts passed with every tool call, and the lines of the chat's system prompt that make the chat model pass the number of slides and the amount of text on as said, name the text instead of retelling it, and ask the person before a cut. No new tool. Server logic only, 0.6.2 or later. |
| [template-studio.md](template-studio.md) | The app has, or is getting, a template studio: the part where a person turns a `.pptx` into a template and keeps its manifest good. The prompt is about the logic of that part as a whole: the short path from an upload to a published template, what the studio stores, what it must never do, and how to prove it works. It says nothing about the user interface. |

A prompt names functions, options and messages as they are in the version it ships with. Use the one inside the
version of slotdeck the app has installed, not a copy kept from an earlier one.

A prompt cannot see the app. Each one therefore starts by telling the agent what to find out first and to
report what it found before it changes anything. Read that report: it is where a wrong assumption shows.
