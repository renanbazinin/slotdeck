# Layout variants: a prompt for a coding agent

This is a prompt for a coding agent working in an app that has a template studio built on slotdeck, where a
person cannot yet add a variant of one layout by themselves: its mirror image, a text-only version without
its picture, or a version without an optional line. Typically the template is prepared and published, and
later someone needs, say, a text-only slide the template does not have.

Paste everything below the line to the agent. It asks for two operations, one that lists the variants a
layout allows and one that makes one, and for what happens to a template that the chat already uses. Server
logic only, for slotdeck 0.3.1 or later. The full set of rules for a studio is in
[template-studio.md](template-studio.md); the prepare and publish operations this one builds on are in
[one-step-preparation.md](one-step-preparation.md).

---

You are working in an application that has a template studio built on slotdeck. A person can upload a
template, prepare it and publish it, but cannot add a variant of one of its layouts: the mirror image (the
picture and its text change sides), a text-only version (the picture is left out and the text widens over the
room it leaves), or a version without one optional text (what sits under it moves up). Your task is to add
that: a person picks a layout, sees which variants it allows and why the others are not possible, and makes
one. The code works out every position and text budget. Nobody types a box, and no model is called.

This is about what the server does and returns. Do not design screens. The reference is
`node_modules/slotdeck/docs/template-studio.md`, sections "`derivePattern`" and "`patternVariants(template,
pattern)`", and "Variants the code works out" in `node_modules/slotdeck/docs/layout-freedom.md`, not the
internet.

## Find out first

Change nothing yet. Report in a few lines, then say what you plan to change, file by file:

1. The installed version (`node_modules/slotdeck/package.json`). Below 0.3.1, say so and stop: the upgrade
   comes first, with "Upgrading from" in `node_modules/slotdeck/docs/changelog.md`. Before 0.3.1 a mirror
   could move a title that takes its place from the slide's layout, and that variant could not be rendered.
2. How the studio changes a stored manifest today: whether it calls `changeManifest` with operations, and
   which `op` values it lets through. If it filters them, whether `derivePattern` is among them.
3. Whether the studio has a "change in words" operation (`editManifest`). That one can already answer with
   `derivePattern`; say whether a person could reach a variant that way today.
4. For a published template: where the manifest the chat uses is stored, whether a change to it reaches the
   chat at once or only after publishing again, and what the publish operation checks.
5. Whether prepare runs on a template that has layouts, and with which `variants` option. Prepare with the
   default options adds the mirrored and text-only variants itself, so a person who removed one gets it back.

## What to build

**1. List the variants of a layout.** An operation that takes a template and a layout id and answers with
`patternVariants`:

```ts
const options = patternVariants(template, patternId); // throws when the template has no such layout: answer "not found"
// [{ variant: 'mirror', available: true, note: 'the mirror image of "text-and-photo", ...' },
//  { variant: 'text-only', available: false, note: '"cover" has no photo, graphic or chart of its own, so there is nothing to leave out' },
//  { variant: 'without', slot: 'subtitle', available: true, note: '"text-and-photo" without "subtitle": body moves up ...' }]
```

Return each option as it comes, with the id the new layout would get: `<layout>-mirrored`,
`<layout>-text-only` or `<layout>-no-<slot>` (the slot name in lower case words joined by `-`, such as
`no-sub-title` for `subTitle`). Where the manifest already has a layout of that id, add that to the option, so
that a person is not offered what is there. The mirror image of a layout that is itself a mirror (its id ends
in `-mirrored`, or the studio stored that it was made with `mirror`) is its source again: mark it the same way.
`note` is the sentence to show: what the variant would do, or why the slide does not allow it. Do not write
your own reasons.

`patternVariants` reads the template's shapes on each call and changes nothing. Call it on the template as it
is stored now, `loadTemplate(pptx, storedManifest)` or however the studio loads it, not on one kept from
start-up: a variant made a moment ago changes which ids are taken.

**2. Make one.** An operation that takes the layout id, the variant, `slot` for `without`, and optionally an
`id` and a `description`, and goes through `changeManifest`, the same as every other change to a manifest:

```ts
const result = changeManifest(template, [{ op: 'derivePattern', pattern, variant, ...(slot ? { slot } : {}), ...(id ? { id } : {}), ...(description ? { description } : {}) }]);
if (result.applied.length === 0) {
  // result.rejected[0].reason says why: the slide does not allow it, the id is taken, the slot is not an optional text
  // store nothing, answer with that reason
}
// store result.manifest; answer with result.applied[0].message and the new layout
```

Check the input before calling it: `variant` is one of `VARIANTS`, `slot` is given for `without` and only for
it. The new layout is placed right after its source, and the source is not changed. A variant can be made from
a variant: the mirror image, then that without its subtitle.

When the source layout's description names a side ("a photo on the right"), the default description of a
mirror adds a sentence that contradicts it. Let the person pass their own `description`, and say so in the
answer when the source's description has the words left or right in it.

**3. A template the chat already uses.** Adding a layout takes nothing away: every stored presentation still
names layouts and slots that exist. Make it reach the chat the way other changes to a published template reach
it in this app; do not unpublish the template for it. Before storing, check the result the way publish checks
it, on the manifest as it would be stored:

```ts
function fitToUse(template: Template): boolean {
  if (checkManifest(template).some((issue) => issue.severity === 'error')) return false;
  try {
    for (const items of ['min', 'max'] as const) {
      const { deck, assets } = sampleDeck(template, { items });
      renderDeck(template, deck, { assets, layout: 'adaptive' });
    }
    return true;
  } catch {
    return false;
  }
}
```

If the template is published and the result is not fit to use, store nothing and answer with the reason.
`changeManifest` already refuses a change that adds an error, so this catches a sample that does not render.

**4. Taking a variant away again.** A variant is a layout like the others: `{ op: 'removePattern', pattern }`
removes it, if the studio offers removing layouts. Before removing a layout of a published template, check
each stored presentation of the template with `validateDeck(manifest, deck)` against the manifest without it,
and refuse, naming them, when one uses it. If prepare runs with its default options on a template a person has
changed by hand, it adds the mirrored and text-only variants again: where the studio lets people remove them,
run prepare on such a template with `variants: []`, as said in part 4 of `one-step-preparation.md`.

**5. A screen, only if there is a place for it.** If the studio already shows one layout with actions beside
it, the list from part 1 can be shown there: one button for each available option that starts part 2, and the
`note` of each one that is not available. Nothing more is needed. If there is no such place, leave the screen
and say so in your final message.

## Rules

- No model is called in any of this. `patternVariants` and `derivePattern` are code only.
- Every position, size and budget comes from `derivePattern`. Do not compute a box, and do not offer one to
  type.
- A refused variant stores nothing and changes nothing; its reason is passed on as it is.
- No change renames or removes a layout that exists, except part 4, which a person asks for.
- Do not add a variant by any other path, such as copying a pattern and editing its slots by hand: that is what
  `derivePattern` does, with the checks.

## Tests

No test calls a live model. Use a template the app's tests already have with a layout that has one picture
beside its text, and one without a picture. Each test starts from the same call the studio makes, not from a
function only the test calls.

1. The list for the picture-and-text layout has `mirror` and `text-only` available, each with the id it would
   get. The list for the layout without a picture has both unavailable, each with its `note`. An unknown
   layout id is answered "not found".
2. Making `text-only` adds `<layout>-text-only` right after its source, leaves the source as it was, and
   stores the manifest. The sample deck renders at `min` and at `max`. Asking for it again is refused because
   the id is taken, and stores nothing; the list now marks that option as there.
3. Making `without` with a slot that is required, or with no slot, is refused with the reason and stores
   nothing. Making `mirror` on the layout without a picture is refused with the reason `note` gave.
4. On a published template: after making a variant, the tools built the way the chat side builds them offer
   the new layout to the planner, and every stored presentation of the template still passes `validateDeck`.
   The template is still published.
5. No request reaches a model in any of the tests above.

## Done when

- the report and the plan were given;
- the five tests pass;
- your final message says which operations a person now has, how a new variant reaches the chat for a
  published template, what you did about prepare adding variants back, and whether a screen shows the list.
