// What the code settles in a draft by measurement, and the suggestions it makes for a manifest
// that already exists. No test here asks a model anything but scripted labels.
import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { test } from 'node:test';
import { fileURLToPath } from 'node:url';
import { SUGGESTION_KINDS, changeManifest, checkManifest, draftManifest, loadTemplate, renderDeck, sampleDeck, suggestChanges, validateDeck } from '../src/index.ts';
import type { JsonRequest, LlmProvider, ManifestSuggestion, PatternDef, Template, TemplateManifest } from '../src/index.ts';
import { buildFixturePptx, cardsManifest, fixtureManifest, gridManifest } from './fixture.ts';

const root = join(dirname(fileURLToPath(import.meta.url)), '..');
const pptx = buildFixturePptx();

/**
 * Labels as a model gave them for a slide it could not see: the small pictograms of the cards as
 * photographs, and the cards as people because the filler is about a team.
 */
function labelling(): { llm: LlmProvider; requests: JsonRequest[] } {
  const requests: JsonRequest[] = [];
  const llm: LlmProvider = {
    async generateJson(request) {
      requests.push(request);
      if (request.prompt.startsWith('Sample texts')) return { language: 'en', fields: [] };
      if (request.prompt.startsWith('Slide 4')) {
        return {
          usable: true,
          reason: '',
          id: 'card-row',
          description: 'A row of cards.',
          figures: false,
          shapes: [
            { shape: '2', role: 'title', name: 'title', optional: false },
            ...[1, 2, 3].flatMap((n) => [
              { shape: `${n}1`, role: 'text', name: `card${n}Text`, optional: false },
              { shape: `${n}2`, role: 'photo', name: `card${n}Photo`, optional: false },
            ]),
          ],
        };
      }
      if (request.prompt.startsWith('Slide 5')) {
        return {
          usable: true,
          reason: '',
          id: 'stat-grid',
          description: 'A grid of figures.',
          figures: true,
          shapes: [{ shape: '2', role: 'title', name: 'title', optional: false }, ...[11, 21, 31, 41].map((id, index) => ({ shape: String(id), role: 'number', name: `stat${index + 1}`, optional: false }))],
        };
      }
      return { usable: false, reason: 'not a layout', id: '', description: '', figures: false, shapes: [] };
    },
  };
  return { llm, requests };
}

const patternOf = (manifest: TemplateManifest, id: string) => manifest.patterns.find((pattern) => pattern.id === id) as PatternDef;
const errorsOf = (template: Template) => checkManifest(template).filter((issue) => issue.severity === 'error');
const kindsOf = (suggestions: ManifestSuggestion[]) => suggestions.map((suggestion) => `${suggestion.kind} ${suggestion.pattern}${suggestion.variant ? ` ${suggestion.variant}` : ''}`);

test('a draft settles by measurement how many items a row takes and what a small picture is', async () => {
  const { llm, requests } = labelling();
  const { manifest, issues, notes, suggestions } = await draftManifest(pptx, { llm, slides: [4, 5] });
  assert.deepEqual(issues.filter((issue) => issue.severity === 'error'), []);

  // The model was told which pictures are pictograms, and each call carries a ceiling on its answer.
  const slide4 = requests.find((request) => request.prompt.startsWith('Slide 4'));
  assert.match(slide4?.prompt ?? '', /#12 picture at x69 y41, 4 wide 8 high, in group #10 \| 8x8 px png, same picture on 3 slides, small: a pictogram/);
  assert.match(slide4?.prompt ?? '', /never the subject of the filler/);
  assert.equal(slide4?.maxOutputTokens, 400 + 60 * 8, 'eight shapes are listed for slide 4');
  assert.ok(requests.every((request) => typeof request.maxOutputTokens === 'number'));

  // It answered "photo" all the same. A picture that small is an icon, and its name says so.
  const row = patternOf(manifest, 'card-row');
  assert.deepEqual(row.slots.filter((slot) => slot.kind === 'image'), [
    { name: 'card1Icon', kind: 'image', imageKind: 'icon', shapeId: '12' },
    { name: 'card2Icon', kind: 'image', imageKind: 'icon', shapeId: '22' },
    { name: 'card3Icon', kind: 'image', imageKind: 'icon', shapeId: '32' },
  ]);

  // Three cards are drawn. Six fit with a third of the drawn space between them; eight would touch.
  assert.deepEqual(row.repeat, { item: 'card', slots: ['card{n}Text', 'card{n}Icon'], shapes: [['10'], ['20'], ['30']], min: 2, max: 6, span: { left: 0.03, right: 0.97 } });
  // A grid takes what it draws: one more would need a line the slide does not have.
  assert.deepEqual([patternOf(manifest, 'stat-grid').repeat?.max, patternOf(manifest, 'stat-grid').repeat?.span], [4, undefined]);

  assert.equal(notes.length, 2);
  assert.match(notes[0] ?? '', /^Slide 4: 3 pictures labelled photo are no more than 1\.1 inch on a side, so they were made icons\.$/);
  assert.match(notes[1] ?? '', /^"card-row": the row of cards takes 2 to 6\. The row may now spread from 3% to 97% of the slide's width\.$/);

  // What is left to a person is a matter of taste, and comes as suggestions.
  assert.deepEqual(kindsOf(suggestions), ['row-widen card-row', 'row-widen stat-grid']);

  // The draft is usable as it is: a deck with six cards stays in every limit and renders.
  const template = loadTemplate(pptx, manifest);
  const { deck, assets } = sampleDeck(template, { items: 'max' });
  assert.equal(deck.slides.find((slide) => slide.pattern === 'card-row')?.items, 6);
  assert.deepEqual(validateDeck(manifest, deck, { assets: Object.keys(assets) }), []);
  assert.ok(renderDeck(template, deck, { assets }).pptx.length > 1000);
});

test('a draft can keep rows as drawn, or let their items widen', async () => {
  const drawn = await draftManifest(pptx, { llm: labelling().llm, slides: [4, 5], rows: 'drawn' });
  assert.deepEqual(patternOf(drawn.manifest, 'card-row').repeat, { item: 'card', slots: ['card{n}Text', 'card{n}Icon'], shapes: [['10'], ['20'], ['30']], min: 2, max: 3 });
  assert.equal(drawn.notes.length, 1, 'only the icons are noted');
  // What `fit` would have done is then the first suggestion.
  const more = drawn.suggestions.find((suggestion) => suggestion.kind === 'row-max');
  // Six fit only once the row spreads into the room beside it, so the span is part of the change.
  assert.deepEqual(more?.changes, [{ op: 'setRepeat', pattern: 'card-row', max: 6, span: { left: 0.03, right: 0.97 } }]);
  assert.match(more?.message ?? '', /takes up to 3, and the slide has room for 6 side by side .* For that the row spreads from 3% to 97% of the slide's width\. Up to 8 fit when they sit closer than that\.$/);
  // Taken, it gives the row that `fit` drafts.
  const taken = changeManifest(loadTemplate(pptx, drawn.manifest), more?.changes ?? []);
  assert.deepEqual(taken.rejected, []);
  assert.deepEqual(patternOf(taken.manifest, 'card-row').repeat, patternOf((await draftManifest(pptx, { llm: labelling().llm, slides: [4, 5] })).manifest, 'card-row').repeat);

  const wide = await draftManifest(pptx, { llm: labelling().llm, slides: [4, 5], widen: 1.5 });
  assert.deepEqual([patternOf(wide.manifest, 'card-row').repeat?.widen, patternOf(wide.manifest, 'stat-grid').repeat?.widen], [1.5, 1.5]);
  assert.deepEqual(wide.suggestions.filter((suggestion) => suggestion.kind === 'row-widen'), []);
  assert.deepEqual(errorsOf(loadTemplate(pptx, wide.manifest)), []);

  for (const widen of [1, 5, Number.NaN]) {
    await assert.rejects(draftManifest(pptx, { llm: labelling().llm, slides: [4], widen }), /widen is a factor above 1 and at most 4/);
  }
});

test('suggestChanges lists what the slides allow that a manifest does not use, and each applies as it stands', () => {
  // The manifest a draft made before any of this was settled: photographs, and a row of three.
  const cards = patternOf(cardsManifest, 'cards');
  const before: TemplateManifest = {
    ...cardsManifest,
    patterns: [
      ...fixtureManifest.patterns,
      {
        ...cards,
        slots: cards.slots.map((slot) => (slot.kind === 'image' ? { name: slot.name, kind: 'image', imageKind: 'photo', shapeId: slot.shapeId } : slot)),
        repeat: { item: 'card', slots: ['card{n}Text', 'card{n}Icon'], shapes: [['10', '13'], ['20'], ['30']], min: 2, max: 3 },
      },
    ],
  };
  const template = loadTemplate(pptx, before);
  const suggestions = suggestChanges(template);
  assert.deepEqual(kindsOf(suggestions), ['icon cards', 'row-max cards', 'row-widen cards']);
  assert.deepEqual(suggestions.map((suggestion) => suggestion.slide), [4, 4, 4]);
  // One change covers the row: a change to one item's slot is made to every item.
  assert.deepEqual(suggestions[0]?.changes, [{ op: 'setSlot', pattern: 'cards', slot: 'card1Icon', imageKind: 'icon' }]);
  assert.match(suggestions[0]?.message ?? '', /^"card1Icon", "card2Icon", "card3Icon" are photo slots, but their pictures are no more than 4% of the slide wide: pictograms\./);
  assert.deepEqual(before, template.manifest, 'nothing is changed by asking');
  assert.equal(patternOf(template.manifest, 'cards').repeat?.max, 3);

  assert.deepEqual(kindsOf(suggestChanges(template, { kinds: ['row-max'] })), ['row-max cards']);
  assert.deepEqual(suggestChanges(template, { kinds: ['row-widen'], widen: 1.5 })[0]?.changes, [{ op: 'setRepeat', pattern: 'cards', widen: 1.5 }]);
  assert.deepEqual([...SUGGESTION_KINDS], ['icon', 'row', 'row-max', 'row-widen', 'variant']);

  // Taking them all leaves nothing more to suggest.
  const result = changeManifest(template, suggestions.flatMap((suggestion) => suggestion.changes));
  assert.deepEqual(result.rejected, []);
  const after = loadTemplate(pptx, result.manifest);
  assert.deepEqual(errorsOf(after), []);
  assert.deepEqual(patternOf(result.manifest, 'cards').repeat, { item: 'card', slots: ['card{n}Text', 'card{n}Icon'], shapes: [['10', '13'], ['20'], ['30']], min: 2, max: 6, span: { left: 0.03, right: 0.97 }, widen: 2 });
  assert.deepEqual(patternOf(result.manifest, 'cards').slots.filter((slot) => slot.kind === 'image').map((slot) => slot.imageKind), ['icon', 'icon', 'icon']);
  assert.deepEqual(suggestChanges(after), []);

  // A large picture stays a photograph, and a hand-made manifest that uses its slides fully gets no advice.
  assert.deepEqual(suggestChanges(loadTemplate(pptx, fixtureManifest)), []);
  assert.deepEqual(kindsOf(suggestChanges(loadTemplate(pptx, gridManifest))), ['row-widen grid']);
});

test('suggestChanges finds a row that numbered slot names describe but the pattern does not have', () => {
  const cards = patternOf(cardsManifest, 'cards');
  const { repeat: _repeat, ...loose } = cards;
  const template = loadTemplate(pptx, { ...cardsManifest, patterns: [...fixtureManifest.patterns, loose] });
  const suggestions = suggestChanges(template);
  assert.deepEqual(kindsOf(suggestions), ['row cards']);
  assert.match(suggestions[0]?.message ?? '', /name 3 cards alike, but the pattern has no row, so a deck must fill all 3\. As a row, a slide can use 2 to 3 of them\./);

  const found = changeManifest(template, suggestions[0]?.changes ?? []);
  assert.deepEqual(found.rejected, []);
  assert.equal(patternOf(found.manifest, 'cards').repeat?.max, 3);
  // With the row found, the next question is how many it takes.
  assert.deepEqual(kindsOf(suggestChanges(loadTemplate(pptx, found.manifest))), ['row-max cards', 'row-widen cards']);
});

test('suggestChanges offers each variant once, and none of a variant', () => {
  const demoPptx = readFileSync(join(root, 'templates', 'demo.pptx'));
  const demoManifest = JSON.parse(readFileSync(join(root, 'templates', 'demo.manifest.json'), 'utf8')) as TemplateManifest;
  // The demo manifest as it was before it had derived variants of its own.
  const plain: TemplateManifest = { ...demoManifest, patterns: demoManifest.patterns.filter((pattern) => !pattern.id.endsWith('-mirrored')) };
  const template = loadTemplate(demoPptx, plain);

  const variants = suggestChanges(template, { kinds: ['variant'] });
  const mirrors = variants.filter((suggestion) => suggestion.variant === 'mirror');
  assert.ok(mirrors.length >= 2 && variants.some((suggestion) => suggestion.variant === 'text-only'), kindsOf(variants).join(', '));
  assert.ok(variants.filter((suggestion) => suggestion.variant === 'without').every((suggestion) => typeof suggestion.slot === 'string'));

  // Every one of them is made as offered.
  const all = changeManifest(template, variants.flatMap((suggestion) => suggestion.changes));
  assert.deepEqual(all.rejected, []);
  assert.equal(all.manifest.patterns.length, plain.patterns.length + variants.length);
  assert.deepEqual(errorsOf(loadTemplate(demoPptx, all.manifest)), []);
  // Made once, a variant is not offered again, and a variant gets no variants of its own.
  assert.deepEqual(suggestChanges(loadTemplate(demoPptx, all.manifest), { kinds: ['variant'] }), []);

  // The demo manifest ships with its mirrored variants: only what it lacks is offered.
  const shipped = kindsOf(suggestChanges(loadTemplate(demoPptx, demoManifest), { kinds: ['variant'] }));
  const made = demoManifest.patterns.filter((pattern) => pattern.id.endsWith('-mirrored')).map((pattern) => pattern.id.replace(/-mirrored$/, ''));
  assert.ok(made.length > 0);
  for (const id of made) assert.ok(!shipped.includes(`variant ${id} mirror`), `${id} is already mirrored`);
});
