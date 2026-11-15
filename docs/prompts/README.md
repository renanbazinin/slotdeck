# Prompts

Prompts to hand to a coding agent that works on an application built with slotdeck. Each file is a complete
prompt: paste its text to the agent, or tell the agent to read it at
`node_modules/slotdeck/docs/prompts/<file>`. They are written for an agent that works without internet access
and reads the rest of this documentation from the package.

| Prompt | Hand it over when |
| --- | --- |
| [one-step-preparation.md](one-step-preparation.md) | The app has a template studio, and a person who uploads a template is met by lists of suggestions and issues. The prompt asks for one operation that prepares a template by itself, a short summary to look at, and publishing. It is short, and the place to start. |
| [template-studio.md](template-studio.md) | The app has, or is getting, a template studio: the part where a person turns a `.pptx` into a template and keeps its manifest good. The prompt is about the logic of that part as a whole: the short path from an upload to a published template, what the studio stores, what it must never do, and how to prove it works. It says nothing about the user interface. |

A prompt names functions, options and messages as they are in the version it ships with. Use the one inside the
version of slotdeck the app has installed, not a copy kept from an earlier one.

A prompt cannot see the app. Each one therefore starts by telling the agent what to find out first and to
report what it found before it changes anything. Read that report: it is where a wrong assumption shows.
