// Variants the studio derives by itself: the mirror image of a slide with a picture, the same
// slide without its picture, and a slide without one of its optional lines. A person or a model
// names the variant; every position and budget is worked out by the code.
import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { test } from 'node:test';
import { fileURLToPath } from 'node:url';
import {
  PptxPackage,
  VARIANTS,
  changeManifest,
  checkManifest,
  draftManifest,
  inspectTemplate,
  editManifest,
  isPlainText,
  loadTemplate,
  patternVariants,
  renderDeck,
  sampleDeck,
  validateDeck,
} from '../src/index.ts';
import type { DeckSpec, LlmProvider, ManifestChange, PatternDef, SlideBox, SlotDef, Template, TemplateManifest } from '../src/index.ts';
import { NS, childElement, descendants, parseXml, path, rootElement } from '../src/pptx/xml.ts';
import type { Element } from '../src/pptx/xml.ts';
import { buildFixturePptx, cardsManifest, fixtureManifest } from './fixture.ts';

const root = join(dirname(fileURLToPath(import.meta.url)), '..');
const demoPptx = readFileSync(join(root, 'templates', 'demo.pptx'));
const demoManifest = JSON.parse(readFileSync(join(root, 'templates', 'demo.manifest.json'), 'utf8')) as TemplateManifest;
// The demo manifest as it was before it had derived variants of its own, so the tests make them.
const plain: TemplateManifest = { ...demoManifest, patterns: demoManifest.patterns.filter((pattern) => !pattern.id.endsWith('-mirrored')) };
const demo = loadTemplate(demoPptx, plain);
const picture = new Uint8Array(readFileSync(join(root, 'examples', 'icons', 'star.png')));

const { slideWidth: WIDTH, slideHeight: HEIGHT } = inspectTemplate(demo.pkg);

function derive(template: Template, changes: ManifestChange[]): { manifest: TemplateManifest; template: Template; messages: string[] } {
  const result = changeManifest(template, changes);
  assert.deepEqual(result.rejected, [], 'every change is made');
  assert.deepEqual(result.issues, checkManifest(template), 'a derived pattern passes the checker: it adds no error and no warning');
  return { manifest: result.manifest, template: { pkg: template.pkg, manifest: result.manifest }, messages: result.applied.map((entry) => entry.message) };
}

const patternOf = (manifest: TemplateManifest, id: string) => manifest.patterns.find((pattern) => pattern.id === id) as PatternDef;
const slotOf = (pattern: PatternDef, name: string) => pattern.slots.find((slot) => slot.name === name) as SlotDef;
const boxOf = (pattern: PatternDef, name: string) => (slotOf(pattern, name) as { box?: SlideBox }).box;
const budgetOf = (pattern: PatternDef, name: string) => (slotOf(pattern, name) as { maxChars?: number }).maxChars;

function slideOf(template: Template, deck: DeckSpec, assets: Record<string, Uint8Array> = { photo: picture }): Element {
  const rendered = renderDeck(template, deck, { assets, layout: 'adaptive' });
  assert.deepEqual(rendered.warnings, []);
  return rootElement(parseXml(PptxPackage.load(rendered.pptx).readText('ppt/slides/slide1.xml')));
}

function shapeById(slide: Element, id: string): Element | undefined {
  const props = descendants(slide, NS.p, 'cNvPr').find((el) => el.getAttribute('id') === id);
  return props?.parentNode?.parentNode as Element | undefined;
}

/** x, y, width and height of a shape on a rendered slide, in EMU. */
function frameOf(slide: Element, id: string): number[] | undefined {
  const shape = shapeById(slide, id);
  if (!shape) return undefined;
  const xfrm = (shape.localName === 'grpSp' ? path(shape, [NS.p, 'grpSpPr'], [NS.a, 'xfrm']) : path(shape, [NS.p, 'spPr'], [NS.a, 'xfrm'])) as Element;
  const off = childElement(xfrm, NS.a, 'off');
  const ext = childElement(xfrm, NS.a, 'ext');
  return [off?.getAttribute('x'), off?.getAttribute('y'), ext?.getAttribute('cx'), ext?.getAttribute('cy')].map(Number);
}

const textOf = (slide: Element, id: string) =>
  descendants(shapeById(slide, id) as Element, NS.a, 't')
    .map((node) => node.textContent ?? '')
    .join('');

const fields = { organizationDepartment: 'Org' };

// --- mirror -----------------------------------------------------------------------

test('mirror: a picture that runs off the slide changes sides with its text, and the header keeps its place', () => {
  const { manifest, template, messages } = derive(demo, [{ op: 'derivePattern', pattern: 'text-and-photo', variant: 'mirror' }]);
  assert.deepEqual(manifest.patterns.map((pattern) => pattern.id).slice(1, 3), ['text-and-photo', 'text-and-photo-mirrored'], 'the variant follows its source');
  const source = patternOf(manifest, 'text-and-photo');
  const mirrored = patternOf(manifest, 'text-and-photo-mirrored');
  assert.deepEqual(source, patternOf(plain, 'text-and-photo'), 'the source pattern is not touched');

  // The photo ran off the right edge at full height. Mirrored it runs off the left edge, and starts
  // under the logo and the organisation line, which no pattern may move.
  assert.deepEqual(boxOf(mirrored, 'photo'), { x: 0, y: 0.1367, w: 0.4277, h: 0.8633 });
  // The three texts kept one left edge at 3.75%; they keep one left edge on the other side, and their sizes.
  assert.deepEqual(boxOf(mirrored, 'title'), { x: 0.4825, y: 0.2, w: 0.48, h: 0.1867 });
  assert.deepEqual([boxOf(mirrored, 'subtitle')?.x, boxOf(mirrored, 'body')?.x], [0.4825, 0.4825]);
  // Nothing changed size, so no budget changed, and the deck's own line is left alone.
  assert.deepEqual(mirrored.slots.map((slot) => budgetOf(mirrored, slot.name)), source.slots.map((slot) => budgetOf(source, slot.name)));
  assert.equal(boxOf(mirrored, 'organizationDepartment'), undefined);
  assert.match(mirrored.description, /^A title, subtitle, and body text .* Mirrored: here the picture is on the left and the text on the right\. Alternate it with "text-and-photo" for variety\.$/);
  assert.match(messages[0] ?? '', /^"text-and-photo-mirrored" is a new pattern from slide 2: the mirror image of "text-and-photo", with the picture on the left .* The picture is cut back to clear shape #3 and shape #4 \("Organization name \| Depa"\), which keep their place\.$/);

  const deck: DeckSpec = { version: 1, fields, slides: [{ pattern: 'text-and-photo-mirrored', slots: { title: 'T', subtitle: 'S', body: 'B', photo: { asset: 'photo' } } }] };
  assert.deepEqual(validateDeck(manifest, deck, { assets: ['photo'] }), []);
  const slide = slideOf(template, deck);
  assert.deepEqual(frameOf(slide, '2'), [0, Math.round(0.1367 * HEIGHT), Math.round(0.4277 * WIDTH), Math.round(0.8633 * HEIGHT)]);
  assert.equal(frameOf(slide, '5')?.[0], Math.round(0.4825 * WIDTH));
  assert.equal(textOf(slide, '5'), 'T');
  // The square picture is cropped at its sides to the new, taller frame, not stretched into it.
  const crop = descendants(shapeById(slide, '2') as Element, NS.a, 'srcRect')[0];
  assert.ok(Number(crop?.getAttribute('l')) > 0 && crop?.getAttribute('l') === crop?.getAttribute('r'));
  // A deck that reads the other way gets the same places: a direction never moves a shape.
  const hebrew = slideOf(template, { ...deck, direction: 'rtl' });
  assert.deepEqual(frameOf(hebrew, '5'), frameOf(slide, '5'));
  assert.deepEqual(frameOf(hebrew, '2'), frameOf(slide, '2'));
});

test('mirror: a picture inside the margins changes places with its text inside the same margins', () => {
  const { manifest } = derive(demo, [
    { op: 'derivePattern', pattern: 'photo-stat', variant: 'mirror', id: 'Stat Left', description: 'A figure on the left, a photo on the right.' },
    { op: 'derivePattern', pattern: 'chart-native', variant: 'mirror' },
  ]);
  const mirrored = patternOf(manifest, 'stat-left');
  assert.equal(mirrored.description, 'A figure on the left, a photo on the right.');
  assert.equal(mirrored.figures, true, 'the pattern is still offered only for briefs with figures');
  // Photo 3.75%..45.75% and text 50%..96% become text 3.75%..49.75% and photo 54%..96%.
  assert.deepEqual(boxOf(mirrored, 'photo'), { x: 0.54, y: 0.1733, w: 0.42, h: 0.7333 });
  assert.deepEqual([boxOf(mirrored, 'title'), boxOf(mirrored, 'statDescription')?.x], [{ x: 0.0375, y: 0.1867, w: 0.4575, h: 0.16 }, 0.0375]);
  // A native chart is placed by its box like a picture. The title above both stays where it is.
  const chart = patternOf(manifest, 'chart-native-mirrored');
  assert.deepEqual([boxOf(chart, 'chart')?.x, boxOf(chart, 'chartDescription')?.x, boxOf(chart, 'title')], [0.4373, 0.0375, undefined]);
});

test('mirror: a shape that is not a slot, and a group, move with their side', () => {
  // Without its slot, the subtitle of the photo slide is part of the design. It still has to follow the title.
  const fixed = changeManifest(demo, [{ op: 'removeSlot', pattern: 'text-and-photo', slot: 'subtitle' }]).manifest;
  const { manifest, template } = derive({ pkg: demo.pkg, manifest: fixed }, [{ op: 'derivePattern', pattern: 'text-and-photo', variant: 'mirror' }]);
  const mirrored = patternOf(manifest, 'text-and-photo-mirrored');
  assert.deepEqual(slotOf(mirrored, 'moved1'), { name: 'moved1', kind: 'text', shapeId: '6', empty: 'keep', box: { x: 0.4825, y: 0.4067, w: 0.4725, h: 0.0533 } });
  const deck: DeckSpec = { version: 1, fields, slides: [{ pattern: 'text-and-photo-mirrored', slots: { title: 'T', body: 'B', photo: { asset: 'photo' } } }] };
  const slide = slideOf(template, deck);
  assert.equal(frameOf(slide, '6')?.[0], Math.round(0.4825 * WIDTH));
  assert.equal(textOf(slide, '6'), 'A supporting line in one sentence', 'it keeps the template’s own words');

  // On the fixture, the picture and the text are each inside a group. A shape inside a group cannot
  // take a box, so the groups themselves are moved, and what is inside them is filled where it is.
  const pptx = buildFixturePptx();
  const grouped: TemplateManifest = {
    ...fixtureManifest,
    patterns: [
      ...fixtureManifest.patterns,
      {
        id: 'grouped',
        sourceSlide: 4,
        description: 'A picture beside text.',
        slots: [
          { name: 'photo', kind: 'image', imageKind: 'photo', shapeId: '12' },
          { name: 'text', kind: 'text', shapeId: '21', maxChars: 20 },
        ],
      },
    ],
  };
  const made = derive(loadTemplate(pptx, grouped), [{ op: 'derivePattern', pattern: 'grouped', variant: 'mirror' }]);
  const pattern = patternOf(made.manifest, 'grouped-mirrored');
  assert.deepEqual(slotOf(pattern, 'moved1'), { name: 'moved1', kind: 'text', shapeId: '10', empty: 'keep', box: { x: 0.0109, y: 0.3888, w: 0.1094, h: 0.2916 } });
  assert.deepEqual([boxOf(pattern, 'photo'), boxOf(pattern, 'text')], [undefined, undefined]);
  const rendered = slideOf(made.template, { version: 1, slides: [{ pattern: 'grouped-mirrored', slots: { photo: { asset: 'photo' }, text: 'Inside' } }] });
  // The group with the picture went from the right end to the left; its size is exactly as drawn, so nothing inside is rescaled.
  assert.deepEqual(frameOf(rendered, '10'), [Math.round(0.0109 * 9144000), Math.round(0.3888 * 5143500), 1000000, 1500000]);
  assert.equal(frameOf(rendered, '20')?.[0], Math.round(0.6562 * 9144000));
  assert.equal(textOf(rendered, '21'), 'Inside');
});

// --- text only ----------------------------------------------------------------------

test('text-only: the picture is left out, the text widens to one edge and its budgets are measured again', () => {
  const { manifest, template, messages } = derive(demo, [
    { op: 'derivePattern', pattern: 'text-and-photo', variant: 'text-only', id: 'words' },
    { op: 'derivePattern', pattern: 'photo-stat', variant: 'text-only' },
  ]);
  const words = patternOf(manifest, 'words');
  assert.deepEqual(slotOf(words, 'photo'), { name: 'photo', kind: 'image', imageKind: 'graphic', shapeId: '2', empty: 'drop' });
  // The photo ran off the right edge, so the text may reach the margin it keeps on its own side.
  // It stops earlier, at 90 characters a line of the body's 18 pt, and every text stops on that edge.
  assert.deepEqual(['title', 'subtitle', 'body'].map((name) => boxOf(words, name)), [
    { x: 0.0375, y: 0.2, w: 0.8438, h: 0.1867 },
    { x: 0.0375, y: 0.4067, w: 0.8438, h: 0.0533 },
    { x: 0.0375, y: 0.4933, w: 0.8438, h: 0.3867 },
  ]);
  assert.deepEqual(['title', 'subtitle', 'body'].map((name) => budgetOf(words, name)), [75, 75, 639]);
  assert.match(messages[0] ?? '', /^"words" is a new pattern from slide 2: "text-and-photo" without its picture "photo"\. Widened over the freed area, with budgets to match: title \(75 characters\), subtitle \(75 characters\) and body \(639 characters\)\.$/);
  assert.match(words.description, /Without the picture: the text takes the room it leaves\.$/);
  assert.ok(isPlainText(words), 'a text-only deck may use it');

  // Text on the right of its picture grows to the left, and starts where the slide's own margin is.
  const stat = patternOf(manifest, 'photo-stat-text-only');
  assert.deepEqual(['title', 'statValue', 'statDescription'].map((name) => boxOf(stat, name)?.x), [0.0375, 0.0375, 0.0375]);
  assert.deepEqual([boxOf(stat, 'title')?.w, budgetOf(stat, 'title'), budgetOf(stat, 'statDescription')], [0.8906, 90, 425]);

  const slide = slideOf(template, { version: 1, fields, slides: [{ pattern: 'words', slots: { title: 'T', body: 'B' } }] }, {});
  assert.equal(shapeById(slide, '2'), undefined, 'the photo is gone');
  assert.deepEqual(frameOf(slide, '7'), [0.0375, 0.4933, 0.8438, 0.3867].map((value, index) => Math.round(value * (index % 2 === 0 ? WIDTH : HEIGHT))));
});

// --- without an optional line ----------------------------------------------------------

test('without: an optional line is left out, what sits under it moves up and gets the room', () => {
  const { manifest, template, messages } = derive(demo, [
    { op: 'derivePattern', pattern: 'text-and-photo', variant: 'without', slot: 'subtitle' },
    { op: 'derivePattern', pattern: 'chart-with-takeaway', variant: 'without', slot: 'takeawayText' },
    { op: 'derivePattern', pattern: 'cover', variant: 'without', slot: 'subtitle', id: 'cover-plain' },
  ]);
  const pattern = patternOf(manifest, 'text-and-photo-no-subtitle');
  assert.equal(slotOf(pattern, 'subtitle').empty, 'drop');
  // The body takes the subtitle's place and keeps its own lower edge: 8.7% of the slide's height more to write in.
  assert.deepEqual(boxOf(pattern, 'body'), { x: 0.0375, y: 0.4067, w: 0.465, h: 0.4733 });
  assert.deepEqual([budgetOf(patternOf(plain, 'text-and-photo'), 'body'), budgetOf(pattern, 'body')], [352, 431]);
  assert.deepEqual([boxOf(pattern, 'title'), boxOf(pattern, 'photo')], [undefined, undefined], 'what is above it or beside it stays');
  assert.match(messages[0] ?? '', /"text-and-photo" without "subtitle": body moves up and has more room \(431 characters\)\.$/);
  assert.match(pattern.description, /Without the subtitle, which leaves more room for the rest\.$/);

  // Nothing sits under the takeaway, so the text above it gets its room.
  const chart = patternOf(manifest, 'chart-with-takeaway-no-takeaway-text');
  assert.deepEqual([boxOf(chart, 'chartDescription'), budgetOf(chart, 'chartDescription')], [{ x: 0.6, y: 0.3, w: 0.36, h: 0.58 }, 407]);
  // On the cover the title's text sits at the bottom of its box, so a taller box would move it: only the line is left out.
  const cover = patternOf(manifest, 'cover-plain');
  assert.deepEqual([slotOf(cover, 'subtitle').empty, boxOf(cover, 'title')], ['drop', undefined]);
  assert.match(messages[2] ?? '', /"cover" without "subtitle"\. Nothing else needed to move\.$/);

  // In strict layout an optional line has to be filled; the variant is the same slide for a deck that has none.
  const deck: DeckSpec = { version: 1, fields, slides: [{ pattern: 'text-and-photo-no-subtitle', slots: { title: 'T', body: 'B', photo: { asset: 'photo' } } }] };
  assert.deepEqual(validateDeck(manifest, deck, { layout: 'strict', assets: ['photo'] }), []);
  const slide = slideOf(template, deck);
  assert.equal(shapeById(slide, '6'), undefined);
  assert.deepEqual(frameOf(slide, '7')?.slice(1), [Math.round(0.4067 * HEIGHT), Math.round(0.465 * WIDTH), Math.round(0.4733 * HEIGHT)]);
});

test('variants build on each other, and the sample deck shows every one', () => {
  const { manifest, template } = derive(demo, [
    { op: 'derivePattern', pattern: 'text-and-photo', variant: 'mirror' },
    { op: 'derivePattern', pattern: 'text-and-photo-mirrored', variant: 'without', slot: 'subtitle' },
    { op: 'derivePattern', pattern: 'list-with-photo', variant: 'text-only' },
  ]);
  const both = patternOf(manifest, 'text-and-photo-mirrored-no-subtitle');
  // The body is where the mirror put it, and moved up from there.
  assert.deepEqual(boxOf(both, 'body'), { x: 0.4825, y: 0.4067, w: 0.465, h: 0.4733 });

  const { deck, assets } = sampleDeck(template);
  assert.deepEqual(deck.slides.map((slide) => slide.pattern), manifest.patterns.map((pattern) => pattern.id));
  assert.deepEqual(validateDeck(manifest, deck, { assets: Object.keys(assets), layout: 'adaptive' }), []);
  assert.deepEqual(renderDeck(template, deck, { assets, layout: 'adaptive' }).warnings, []);
  // The sample fills the widened list to the budget that was measured for it.
  const list = deck.slides.find((slide) => slide.pattern === 'list-with-photo-text-only')?.slots.bulletList as string[];
  assert.ok(list.join('').length > 369 && list.join('').length <= (budgetOf(patternOf(manifest, 'list-with-photo-text-only'), 'bulletList') ?? 0));
});

// --- refusals ---------------------------------------------------------------------------

test('a variant the slide does not allow is refused with the reason', () => {
  const reasons = (template: Template, changes: ManifestChange[]) => {
    const result = changeManifest(template, changes);
    assert.deepEqual(result.applied, []);
    assert.deepEqual(result.manifest, template.manifest);
    return result.rejected.map((entry) => entry.reason);
  };
  assert.deepEqual(
    reasons(loadTemplate(demoPptx, demoManifest), [
      { op: 'derivePattern', pattern: 'cover', variant: 'mirror' },
      { op: 'derivePattern', pattern: 'wide-diagram', variant: 'mirror' },
      { op: 'derivePattern', pattern: 'wide-diagram', variant: 'text-only' },
      { op: 'derivePattern', pattern: 'text-and-photo', variant: 'without' },
      { op: 'derivePattern', pattern: 'text-and-photo', variant: 'without', slot: 'title' },
      { op: 'derivePattern', pattern: 'text-and-photo', variant: 'without', slot: 'photo' },
      { op: 'derivePattern', pattern: 'icon-cards', variant: 'without', slot: 'card1Text' },
      { op: 'derivePattern', pattern: 'text-and-photo', variant: 'flip' as 'mirror' },
      { op: 'derivePattern', pattern: 'text-and-photo', variant: 'mirror' },
      { op: 'derivePattern', pattern: 'nothing', variant: 'mirror' },
    ]),
    [
      '"cover" has no photo, graphic or chart of its own, so there is nothing to change sides with',
      'nothing sits beside the picture "diagramGraphic", so there is nothing for it to change sides with',
      'nothing sits beside the picture "diagramGraphic", so there is nothing for it to make room for',
      'name the slot to leave out, an optional text; "text-and-photo" has "subtitle"',
      '"title" is required; only an optional text can be left out; "text-and-photo" has "subtitle"',
      '"photo" is an image; only an optional text can be left out; "text-and-photo" has "subtitle"',
      '"card1Text" is required; only an optional text can be left out; "icon-cards" has none',
      '"flip" is not a variant; the variants are mirror, text-only, without',
      'another pattern is already called "text-and-photo-mirrored"',
      'there is no pattern "nothing"',
    ],
  );

  const pptx = buildFixturePptx();
  const grouped: TemplateManifest = {
    ...cardsManifest,
    patterns: [
      ...cardsManifest.patterns.map((pattern) => (pattern.id === 'cards' ? { ...pattern, slots: pattern.slots.map((slot) => (slot.name === 'title' ? { ...slot, empty: 'remove' as const } : slot)) } : pattern)),
      {
        id: 'grouped',
        sourceSlide: 4,
        description: 'A picture beside text.',
        slots: [
          { name: 'photo', kind: 'image', imageKind: 'photo', shapeId: '12' },
          { name: 'text', kind: 'text', shapeId: '21', maxChars: 20 },
        ],
      },
    ],
  };
  assert.deepEqual(
    reasons(loadTemplate(pptx, grouped), [
      // The only text beside the picture is inside a group, and a shape inside a group cannot take a box.
      { op: 'derivePattern', pattern: 'grouped', variant: 'text-only' },
      // The note lies on top of the title, so what is under it has nowhere to go.
      { op: 'derivePattern', pattern: 'content', variant: 'without', slot: 'note' },
      // A row of items is laid out by the renderer for each item count; a box would pin it.
      { op: 'derivePattern', pattern: 'cards', variant: 'without', slot: 'title' },
    ]),
    [
      'the text beside the picture is inside a group, and a shape inside a group cannot take a box',
      'shape #2 ("SAMPLE LABEL SAMPLE TITL"), which keeps its place, is where shape #4 would go',
      'shape #30 is part of the row of cards, which the software lays out itself and cannot move',
    ],
  );
});

// --- for a studio screen, a model and a draft ------------------------------------------------

test('patternVariants says which variants a pattern allows, for a screen that offers them', () => {
  assert.deepEqual(VARIANTS, ['mirror', 'text-only', 'without']);
  const options = patternVariants(demo, 'text-and-photo');
  assert.deepEqual(options.map((option) => [option.variant, option.slot, option.available]), [
    ['mirror', undefined, true],
    ['text-only', undefined, true],
    ['without', 'subtitle', true],
  ]);
  assert.match(options[2]?.note ?? '', /body moves up and has more room/);
  assert.deepEqual(patternVariants(demo, 'icon-cards').map((option) => `${option.variant}: ${option.available ? 'yes' : option.note}`), [
    'mirror: "icon-cards" has no photo, graphic or chart of its own, so there is nothing to change sides with',
    'text-only: "icon-cards" has no photo, graphic or chart of its own, so there is nothing to leave out',
  ]);
  assert.throws(() => patternVariants(demo, 'nothing'), /there is no pattern "nothing"/);
});

test('a slot that keeps its shape may show it in another place, whatever kind of shape it is', () => {
  const pptx = buildFixturePptx();
  const manifest: TemplateManifest = {
    ...fixtureManifest,
    patterns: [
      {
        id: 'content',
        sourceSlide: 1,
        description: 'The icon group at the top.',
        slots: [
          { name: 'title', kind: 'text', shapeId: '2', maxChars: 20, styleRun: 'last' },
          { name: 'body', kind: 'text', shapeId: '3', maxChars: 40 },
          { name: 'photo', kind: 'image', imageKind: 'photo', shapeId: '4', empty: 'keep', box: { x: 0.5, y: 0.5, w: 0.2, h: 0.2 } },
          { name: 'note', kind: 'text', shapeId: '6', maxChars: 20, empty: 'drop' },
          { name: 'iconGroup', kind: 'text', shapeId: '7', empty: 'keep', box: { x: 0.5, y: 0.1, w: 0.2187, h: 0.1944 } },
        ],
      },
    ],
  };
  const template = loadTemplate(pptx, manifest);
  assert.deepEqual(checkManifest(template), []);
  const slide = slideOf(template, { version: 1, slides: [{ pattern: 'content', slots: { title: 'T', body: 'B' } }] }, {});
  // The group's box is its own size to within rounding, so only its position is set and its 2:1 scale is untouched.
  assert.deepEqual(frameOf(slide, '7'), [4572000, 514350, 2000000, 1000000]);
  // The kept picture has another size in its new place, so it is given that size.
  assert.deepEqual(frameOf(slide, '4'), [4572000, 2571750, 1828800, 1028700]);
});

test('editManifest lets a model ask for a variant by name; the positions come from the software', async () => {
  const prompts: string[] = [];
  const llm: LlmProvider = {
    async generateJson(request) {
      prompts.push(request.prompt);
      return {
        changes: [
          { op: 'derivePattern', pattern: 'text-and-photo', shape: '', slot: '', values: [{ property: 'variant', value: 'mirror' }, { property: 'id', value: 'photo-left' }, { property: 'description', value: 'A tall photo on the left, the text on the right.' }] },
          { op: 'derivePattern', pattern: 'text-and-photo', shape: '', slot: 'subtitle', values: [{ property: 'variant', value: 'without' }] },
          { op: 'derivePattern', pattern: 'cover', shape: '', slot: '', values: [{ property: 'variant', value: 'mirror' }] },
        ],
        note: '',
      };
    },
  };
  const result = await editManifest(demo, 'Add a mirrored version of the photo slide and one without the subtitle. Mirror the cover too.', { llm });
  assert.match(prompts[0] ?? '', /- derivePattern \(pattern\): a second pattern from the same slide that the software lays out itself\./);
  assert.deepEqual(result.applied.map((entry) => entry.change), [
    { op: 'derivePattern', pattern: 'text-and-photo', variant: 'mirror', id: 'photo-left', description: 'A tall photo on the left, the text on the right.' },
    { op: 'derivePattern', pattern: 'text-and-photo', variant: 'without', slot: 'subtitle' },
  ]);
  // The refusal went back to the model once, with its reason.
  assert.equal(prompts.length, 2);
  assert.match(prompts[1] ?? '', /It refused these, and nothing of them was done:\n3\. "cover" has no photo, graphic or chart of its own/);
  const made = result.manifest.patterns.map((pattern) => pattern.id);
  assert.ok(made.includes('photo-left') && made.includes('text-and-photo-no-subtitle'));
  assert.deepEqual(boxOf(patternOf(result.manifest, 'photo-left'), 'photo'), { x: 0, y: 0.1367, w: 0.4277, h: 0.8633 });
});

test('draftManifest adds the variants a slide allows when asked to, and none by default', async () => {
  const llm: LlmProvider = {
    async generateJson(request) {
      if (/Sample texts from a presentation template/.test(request.prompt)) {
        return { language: 'en', fields: [{ text: 'Organization name | Department', name: 'organizationDepartment' }] };
      }
      const cover = /^Slide 1 of/.test(request.prompt);
      return {
        usable: true,
        reason: 'a content layout',
        id: cover ? 'cover' : 'text-and-photo',
        description: cover ? 'Opening slide.' : 'Text on the left, a photo on the right.',
        figures: false,
        shapes: cover
          ? [{ shape: '5', role: 'title', name: 'title', optional: false }]
          : [
              { shape: '2', role: 'photo', name: 'photo', optional: false },
              { shape: '5', role: 'title', name: 'title', optional: false },
              { shape: '6', role: 'subtitle', name: 'subtitle', optional: true },
              { shape: '7', role: 'text', name: 'body', optional: false },
            ],
      };
    },
  };
  const bare = await draftManifest(demoPptx, { llm, slides: [1, 2] });
  assert.deepEqual(bare.manifest.patterns.map((pattern) => pattern.id), ['cover', 'text-and-photo']);

  const drafted = await draftManifest(demoPptx, { llm, slides: [1, 2], variants: ['mirror', 'text-only'] });
  // The cover has no picture, so it gets no variant; nothing is reported, because nothing was asked of it in particular.
  assert.deepEqual(drafted.manifest.patterns.map((pattern) => pattern.id), ['cover', 'text-and-photo', 'text-and-photo-mirrored', 'text-and-photo-text-only']);
  assert.deepEqual(drafted.issues, []);
  const mirrored = patternOf(drafted.manifest, 'text-and-photo-mirrored');
  assert.deepEqual(boxOf(mirrored, 'photo'), { x: 0, y: 0.1367, w: 0.4277, h: 0.8633 });
  assert.match(mirrored.description, /^Text on the left, a photo on the right\. Mirrored: here the picture is on the left/);
});
