// What the code settles in a draft by measurement, and the suggestions it makes for a manifest
// that already exists. No test here asks a model anything but scripted labels.
import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { test } from 'node:test';
import { fileURLToPath } from 'node:url';
import { SUGGESTION_KINDS, changeManifest, checkManifest, draftManifest, loadTemplate, renderDeck, sampleDeck, suggestChanges, validateDeck } from '../src/index.ts';
import type { JsonRequest, LlmProvider, ManifestSuggestion, PatternDef, Template, TemplateManifest } from '../src/index.ts';
import { buildFixturePptx, cardsManifest, fixtureManifest, fixtureWithSlide, fixtureWithTextCards, gridManifest } from './fixture.ts';

const root = join(dirname(fileURLToPath(import.meta.url)), '..');
const pptx = buildFixturePptx();

/**
 * Labels as a model gave them for a slide it could not see: the small pictograms of the cards as
 * photographs, and the cards as people because the filler is about a team.
 */
function labelling(pictures: [string, string, string] = ['photo', 'photo', 'photo']): { llm: LlmProvider; requests: JsonRequest[] } {
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
              { shape: `${n}2`, role: pictures[n - 1], name: `card${n}Photo`, optional: false },
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

  // It answered "photo" all the same. A picture that small is an icon. The name stays the model's,
  // so the row its names describe is the row that is found.
  const row = patternOf(manifest, 'card-row');
  assert.deepEqual(row.slots.filter((slot) => slot.kind === 'image'), [
    { name: 'card1Photo', kind: 'image', imageKind: 'icon', shapeId: '12' },
    { name: 'card2Photo', kind: 'image', imageKind: 'icon', shapeId: '22' },
    { name: 'card3Photo', kind: 'image', imageKind: 'icon', shapeId: '32' },
  ]);

  // Three cards are drawn. Six fit with a third of the drawn space between them; eight would touch.
  assert.deepEqual(row.repeat, { item: 'card', slots: ['card{n}Text', 'card{n}Photo'], shapes: [['10'], ['20'], ['30']], min: 2, max: 6, span: { left: 0.03, right: 0.97 } });
  // A grid takes what it draws: one more would need a line the slide does not have.
  assert.deepEqual([patternOf(manifest, 'stat-grid').repeat?.max, patternOf(manifest, 'stat-grid').repeat?.span], [4, undefined]);

  assert.equal(notes.length, 2);
  assert.match(notes[0] ?? '', /^Slide 4: "card1Photo", "card2Photo", "card3Photo" were labelled photo, but their pictures are no more than 1\.1 inch on a side: icon slots\.$/);
  assert.match(notes[1] ?? '', /^"card-row": the row of cards takes 2 to 6\. The row may now spread from 3% to 97% of the slide's width\.$/);

  // What is left to a person is a matter of taste, and comes as suggestions. The two-by-two grid
  // gets none: it takes two to four figures, and with any of those its fullest line is full.
  assert.deepEqual(kindsOf(suggestions), ['row-widen card-row']);

  // The draft is usable as it is: a deck with six cards stays in every limit and renders.
  const template = loadTemplate(pptx, manifest);
  const { deck, assets } = sampleDeck(template, { items: 'max' });
  assert.equal(deck.slides.find((slide) => slide.pattern === 'card-row')?.items, 6);
  assert.deepEqual(validateDeck(manifest, deck, { assets: Object.keys(assets) }), []);
  assert.ok(renderDeck(template, deck, { assets }).pptx.length > 1000);
});

test('a draft can keep rows as drawn, or let their items widen', async () => {
  const drawn = await draftManifest(pptx, { llm: labelling().llm, slides: [4, 5], rows: 'drawn' });
  assert.deepEqual(patternOf(drawn.manifest, 'card-row').repeat, { item: 'card', slots: ['card{n}Text', 'card{n}Photo'], shapes: [['10'], ['20'], ['30']], min: 2, max: 3 });
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
  // Only where it can change a slide: the grid's fullest line is full with every number it takes.
  assert.deepEqual([patternOf(wide.manifest, 'card-row').repeat?.widen, patternOf(wide.manifest, 'stat-grid').repeat?.widen], [1.5, undefined]);
  assert.deepEqual(wide.suggestions.filter((suggestion) => suggestion.kind === 'row-widen'), []);
  assert.deepEqual(errorsOf(loadTemplate(pptx, wide.manifest)), []);
  // One sentence for the row, though two things were settled for it.
  const rowNotes = wide.notes.filter((note) => note.startsWith('"card-row"'));
  assert.equal(rowNotes.length, 1);
  assert.match(rowNotes[0] ?? '', /takes 2 to 6\. The row may now spread from 3% to 97% of the slide's width\. With fewer cards on a line than drawn, each grows up to 1\.5 times as wide/);

  for (const widen of [1, 5, Number.NaN, '2' as unknown as number]) {
    await assert.rejects(draftManifest(pptx, { llm: labelling().llm, slides: [4], widen }), /widen is a factor above 1 and at most 4/);
    assert.throws(() => suggestChanges(loadTemplate(pptx, cardsManifest), { widen }), /widen is a factor above 1 and at most 4/);
  }
  // A value of `rows` that is neither of the two is a mistake to say, not a draft of the other kind.
  await assert.rejects(draftManifest(pptx, { llm: labelling().llm, slides: [4], rows: 'Fit' as unknown as 'fit' }), /rows is "fit" or "drawn", not "Fit"/);
  await assert.rejects(draftManifest(pptx, { llm: labelling().llm, slides: [4], variants: ['without' as unknown as 'mirror'] }), /a draft adds the variants "mirror" and "text-only", not "without"/);
  // The same for a kind of suggestion that does not exist. A factor is looked at only when widening is asked about.
  assert.throws(() => suggestChanges(loadTemplate(pptx, cardsManifest), { kinds: ['rows' as unknown as 'row'] }), /"rows" is not a kind of suggestion; the kinds are icon, row, row-max, row-widen, variant/);
  assert.doesNotThrow(() => suggestChanges(loadTemplate(pptx, cardsManifest), { kinds: ['row-max'], widen: 9 }));
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
  assert.match(suggestions[0]?.message ?? '', /^"card1Icon", "card2Icon", "card3Icon" are photo slots, but their pictures are no more than 4% of the slide wide: pictograms\. As icon slots they are filled from the icon set given to the planner, .* Without an icon set, a required icon slot keeps the layout from being offered\.$/);
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
  assert.match(suggestions[0]?.message ?? '', /name 3 cards alike, but the pattern has no row, so every slide draws all 3\. As a row, a slide can use 2 to 3 of them\./);

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
  // Its text-only layout of the same slide was made by hand, under an id of its own, with its own
  // boxes. It is known by what it does: it takes the picture away.
  assert.ok(demoManifest.patterns.some((pattern) => pattern.id === 'text-only' && pattern.sourceSlide === patternOf(demoManifest, 'text-and-photo').sourceSlide));
  assert.ok(!shipped.includes('variant text-and-photo text-only'), shipped.join(', '));

  // A variant given its own id stays known after the pattern it came from is tuned.
  const named = changeManifest(template, [
    { op: 'derivePattern', pattern: 'text-and-photo', variant: 'mirror', id: 'photo-left' },
    { op: 'setSlot', pattern: 'text-and-photo', slot: 'body', maxChars: 300 },
  ]);
  assert.deepEqual(named.rejected, []);
  const again = kindsOf(suggestChanges(loadTemplate(demoPptx, named.manifest), { kinds: ['variant'] }));
  assert.ok(!again.includes('variant text-and-photo mirror'), again.join(', '));
  assert.ok(again.includes('variant text-and-photo without'), 'what it does not have is still offered');

  // Nor is a variant offered again because the pattern it came from was given a decoration since.
  const decorated = changeManifest(loadTemplate(demoPptx, demoManifest), [{ op: 'addSlot', pattern: 'text-and-photo', shape: '3', role: 'decoration' }]);
  assert.deepEqual(decorated.rejected, []);
  const after = kindsOf(suggestChanges(loadTemplate(demoPptx, decorated.manifest), { kinds: ['variant'] }));
  assert.ok(!after.includes('variant text-and-photo text-only') && !after.includes('variant text-and-photo mirror'), after.join(', '));
});

test('a row with arrows among its items is drafted to hold what it draws', async () => {
  // An arrow in each gap between the three cards. With two cards, or four, they would be drawn across the arrows.
  const arrow = (id: number, x: number) => `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="Arrow"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="${x}" y="2600000"/><a:ext cx="600000" cy="200000"/></a:xfrm><a:prstGeom prst="rightArrow"><a:avLst/></a:prstGeom></p:spPr></p:sp>`;
  const steps = fixtureWithSlide(4, (xml) => xml.replace('</p:spTree>', `${arrow(91, 3200000)}${arrow(92, 5200000)}</p:spTree>`));
  const { manifest, notes, suggestions } = await draftManifest(steps, { llm: labelling(['icon', 'icon', 'icon']).llm, slides: [4] });
  assert.deepEqual([patternOf(manifest, 'card-row').repeat?.min, patternOf(manifest, 'card-row').repeat?.max], [3, 3]);
  assert.deepEqual(notes, []);
  assert.deepEqual(kindsOf(suggestions), [], 'no more items, and no widening across the arrows');

  // A manifest that takes the arrows off every slide has nothing in the way: the row is free again.
  const without: TemplateManifest = {
    ...manifest,
    patterns: manifest.patterns.map((pattern) => ({
      ...pattern,
      slots: [...pattern.slots, { name: 'arrow1', kind: 'text', shapeId: '91', empty: 'drop' }, { name: 'arrow2', kind: 'text', shapeId: '92', empty: 'drop' }],
      repeat: { ...(pattern.repeat as NonNullable<typeof pattern.repeat>), min: 2 },
    })),
  } as TemplateManifest;
  assert.deepEqual(kindsOf(suggestChanges(loadTemplate(steps, without), { kinds: ['row-max', 'row-widen'] })), ['row-max card-row', 'row-widen card-row']);

  // A manifest that 0.2.0 drafted takes two to three of such a row. It is offered neither more nor wider.
  const asBefore: TemplateManifest = { ...manifest, patterns: manifest.patterns.map((pattern) => ({ ...pattern, repeat: { ...(pattern.repeat as NonNullable<typeof pattern.repeat>), min: 2 } })) };
  assert.deepEqual(suggestChanges(loadTemplate(steps, asBefore), { kinds: ['row-max', 'row-widen'] }), []);
});

test('a text-only variant never gives one item of a row a place of its own', () => {
  // Three cards on the left, as text boxes on the slide, and a photograph on the right.
  const photo = `<p:pic><p:nvPicPr><p:cNvPr id="60" name="Photo"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr><p:blipFill><a:blip r:embed="rId2"/><a:stretch><a:fillRect/></a:stretch></p:blipFill><p:spPr><a:xfrm><a:off x="5700000" y="1300000"/><a:ext cx="3000000" cy="3000000"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic>`;
  const body = `<p:sp><p:nvSpPr><p:cNvPr id="40" name="Body"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="500000" y="1300000"/><a:ext cx="3800000" cy="500000"/></a:xfrm></p:spPr><p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr sz="1200"/><a:t>SAMPLE BODY TEXT</a:t></a:r></a:p></p:txBody></p:sp>`;
  // The slide's title is left as designed here, so the cards are the only text the deck writes.
  const slots: TemplateManifest['patterns'][number]['slots'] = [
    ...[1, 2, 3].map((n): TemplateManifest['patterns'][number]['slots'][number] => ({ name: `card${n}Text`, kind: 'text', shapeId: `${n}1`, maxChars: 20 })),
    { name: 'photo', kind: 'image', imageKind: 'photo', shapeId: '60' },
  ];
  const repeat = { item: 'card', slots: ['card{n}Text'], shapes: [['11'], ['21'], ['31']], min: 2, max: 3 };
  const rowOnly: TemplateManifest = { ...fixtureManifest, patterns: [{ id: 'cards', sourceSlide: 4, description: 'Cards beside a photo.', slots, repeat }] };

  // The row is all the text beside the picture: the software lays a row out as one, so there is nothing to widen.
  const alone = loadTemplate(fixtureWithTextCards([500000, 1900000, 3300000], photo), rowOnly);
  assert.deepEqual(errorsOf(alone), []);
  const refused = changeManifest(alone, [{ op: 'derivePattern', pattern: 'cards', variant: 'text-only' }]);
  assert.match(refused.rejected[0]?.reason ?? '', /the text beside the picture is the row of cards, which the software lays out itself and cannot widen/);
  assert.ok(!kindsOf(suggestChanges(alone)).includes('variant cards text-only'));

  // With a body text above the row, that text widens and the row is left alone.
  const withBody: TemplateManifest = { ...rowOnly, patterns: [{ ...(rowOnly.patterns[0] as TemplateManifest['patterns'][number]), slots: [...slots, { name: 'body', kind: 'text', shapeId: '40', maxChars: 60 }] }] };
  const made = changeManifest(loadTemplate(fixtureWithTextCards([500000, 1900000, 3300000], body + photo), withBody), [{ op: 'derivePattern', pattern: 'cards', variant: 'text-only' }]);
  assert.deepEqual(made.rejected, []);
  const variant = patternOf(made.manifest, 'cards-text-only');
  assert.ok((variant.slots.find((slot) => slot.name === 'body') as { box?: unknown }).box, 'the body takes the room the picture leaves');
  for (const slot of variant.slots) if (/^card\d/.test(slot.name)) assert.equal((slot as { box?: unknown }).box, undefined, `${slot.name} stays where the row puts it`);
});

test('an icon suggestion is one change for a row, true of every item, and gone once taken', () => {
  const cards = patternOf(cardsManifest, 'cards');
  const withKinds = (kinds: string[], more: Record<string, unknown> = {}): TemplateManifest => ({
    ...cardsManifest,
    patterns: [
      {
        ...cards,
        slots: cards.slots.map((slot) => (slot.kind === 'image' ? ({ name: slot.name, kind: 'image', imageKind: kinds[Number(slot.name[4]) - 1], shapeId: slot.shapeId, ...more } as typeof slot) : slot)),
      },
    ],
  });

  // A draft labels each shape on its own, so a row can come out icon, photo, photo.
  const mixed = loadTemplate(pptx, withKinds(['icon', 'photo', 'photo']));
  const [icon, ...rest] = suggestChanges(mixed, { kinds: ['icon'] });
  assert.deepEqual(rest, []);
  assert.deepEqual(icon?.changes, [{ op: 'setSlot', pattern: 'cards', slot: 'card2Icon', imageKind: 'icon' }]);
  assert.match(icon?.message ?? '', /^"card2Icon", "card3Icon" are photo slots/);
  const taken = changeManifest(mixed, icon?.changes ?? []);
  assert.deepEqual(patternOf(taken.manifest, 'cards').slots.filter((slot) => slot.kind === 'image').map((slot) => slot.imageKind), ['icon', 'icon', 'icon']);
  assert.deepEqual(suggestChanges(loadTemplate(pptx, taken.manifest), { kinds: ['icon'] }), []);

  // One card's picture is cropped into its frame, as a portrait is: not a pictogram. A change to
  // one item is made to every item, so the row is left as it is.
  const cropped = fixtureWithSlide(4, (xml) => xml.replace(/(<p:cNvPr id="32" name="Card icon"\/><p:cNvPicPr\/><p:nvPr\/><\/p:nvPicPr><p:blipFill><a:blip r:embed="rId2"\/>)/, '$1<a:srcRect l="10000" r="10000"/>'));
  assert.deepEqual(suggestChanges(loadTemplate(cropped, withKinds(['photo', 'photo', 'photo'])), { kinds: ['icon'] }), []);

  // A picture the manifest keeps as designed, or always removes, is filled by nothing: no advice.
  assert.deepEqual(suggestChanges(loadTemplate(pptx, withKinds(['photo', 'photo', 'photo'], { empty: 'keep' })), { kinds: ['icon'] }), []);
  // One picture outside a row reads in the singular.
  const single: TemplateManifest = { ...fixtureManifest, patterns: fixtureManifest.patterns.map((pattern) => (pattern.id === 'content' ? { ...pattern, slots: pattern.slots.map((slot) => (slot.name === 'icon' ? { name: 'icon', kind: 'image', imageKind: 'photo', shapeId: '8' } : slot)) } : pattern)) } as TemplateManifest;
  const lone = suggestChanges(loadTemplate(pptx, single), { kinds: ['icon'] });
  assert.match(lone[0]?.message ?? '', /^"icon" is a photo slot, but its picture is no more than \d+% of the slide wide: a pictogram\. As an icon slot it is filled/);
});

test('icon and row taken together leave a row whose items are filled alike', () => {
  // The manifest names three cards alike but has no row, and their pictures are photo slots.
  const cards = patternOf(cardsManifest, 'cards');
  const { repeat: _repeat, ...loose } = cards;
  const noRow: TemplateManifest = {
    ...cardsManifest,
    patterns: [{ ...loose, slots: loose.slots.map((slot) => (slot.kind === 'image' ? { name: slot.name, kind: 'image', imageKind: 'photo', shapeId: slot.shapeId } : slot)) }],
  };
  const facts = suggestChanges(loadTemplate(pptx, noRow), { kinds: ['icon', 'row', 'row-max'] });
  assert.deepEqual(kindsOf(facts), ['icon cards', 'row cards']);
  // The row is not made yet, so one change does not reach every item: each slot has its own.
  assert.deepEqual(facts[0]?.changes.map((change) => (change as { slot: string }).slot), ['card1Icon', 'card2Icon', 'card3Icon']);
  const both = changeManifest(loadTemplate(pptx, noRow), facts.flatMap((suggestion) => suggestion.changes));
  assert.deepEqual(both.rejected, []);
  assert.deepEqual(patternOf(both.manifest, 'cards').slots.filter((slot) => slot.kind === 'image').map((slot) => slot.imageKind), ['icon', 'icon', 'icon']);
  assert.equal(patternOf(both.manifest, 'cards').repeat?.shapes.length, 3);

  // With one picture that is not a pictogram, the row that is about to be made decides: no icon at all.
  const cropped = fixtureWithSlide(4, (xml) => xml.replace(/(<p:cNvPr id="32" name="Card icon"\/><p:cNvPicPr\/><p:nvPr\/><\/p:nvPicPr><p:blipFill><a:blip r:embed="rId2"\/>)/, '$1<a:srcRect l="10000" r="10000"/>'));
  assert.deepEqual(kindsOf(suggestChanges(loadTemplate(cropped, noRow), { kinds: ['icon', 'row'] })), ['row cards']);
});

test('a draft settles photo or icon for a row as a whole, and keeps the row', async () => {
  const pictures = (manifest: TemplateManifest) => patternOf(manifest, 'card-row').slots.filter((slot) => slot.kind === 'image').map((slot) => slot.imageKind);
  // One of three pictures is cropped into a frame. Labelled photo: all stay photo, and the row stands.
  const cropped = fixtureWithSlide(4, (xml) => xml.replace(/(<p:cNvPr id="32" name="Card icon"\/><p:cNvPicPr\/><p:nvPr\/><\/p:nvPicPr><p:blipFill><a:blip r:embed="rId2"\/>)/, '$1<a:srcRect l="10000" r="10000"/>'));
  const kept = await draftManifest(cropped, { llm: labelling().llm, slides: [4] });
  assert.deepEqual(pictures(kept.manifest), ['photo', 'photo', 'photo']);
  assert.equal(patternOf(kept.manifest, 'card-row').repeat?.max, 6);
  assert.equal(kept.notes.length, 1, 'only the row is noted');

  // The model follows the marker shape by shape: icon, icon, photo. Items of a row are filled alike.
  const evened = await draftManifest(cropped, { llm: labelling(['icon', 'icon', 'photo']).llm, slides: [4] });
  assert.deepEqual(pictures(evened.manifest), ['photo', 'photo', 'photo']);
  assert.match(evened.notes[0] ?? '', /^Slide 4: "card1Photo", "card2Photo" were labelled icon, but not every picture of the row measures as a pictogram: all of them are photo slots/);
  assert.equal(patternOf(evened.manifest, 'card-row').repeat?.slots.length, 2);

  // On the plain fixture, where all three are pictograms, a mixed answer becomes icons.
  const icons = await draftManifest(pptx, { llm: labelling(['icon', 'photo', 'photo']).llm, slides: [4] });
  assert.deepEqual(pictures(icons.manifest), ['icon', 'icon', 'icon']);
  assert.match(icons.notes[0] ?? '', /^Slide 4: "card2Photo", "card3Photo" were labelled photo/);
});

test('the ceiling on the answer about the deck grows with the texts it must repeat', async () => {
  const demoPptx = readFileSync(join(root, 'templates', 'demo.pptx'));
  const { llm, requests } = labelling();
  await draftManifest(demoPptx, { llm, slides: [1] });
  const deckCall = requests.find((request) => request.prompt.startsWith('Sample texts'));
  const repeated = [...(deckCall?.prompt.split('Give each a short camelCase field name:')[1]?.split('Also give')[0] ?? '').matchAll(/^- (".*")$/gm)].map((match) => JSON.parse(match[1] as string) as string);
  assert.ok(repeated.length > 0, 'the demo template has a text that repeats across its slides');
  assert.equal(deckCall?.maxOutputTokens, 300 + repeated.reduce((sum, text) => sum + 40 + new TextEncoder().encode(text).length, 0));
});
