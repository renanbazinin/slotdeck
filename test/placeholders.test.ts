// A sample text the template repeats under its titles, such as "Subtitle", is content: each slide
// has its own. Only a text in the header or footer is one value for the whole deck. And a line
// that runs through the content area belongs to how a slide lays out its content, so a layout
// composed on that slide's frame does not keep it. No test calls a live model.
import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { test } from 'node:test';
import { fileURLToPath } from 'node:url';
import { PptxPackage, buildPattern, designKit, draftManifest, inspectTemplate, loadTemplate, prepareManifest, renderDeck, sampleDeck } from '../src/index.ts';
import type { DesignElement, JsonRequest, LlmProvider, PatternDesign, Template, TemplateManifest } from '../src/index.ts';
import { readPresentation } from '../src/pptx/presentation.ts';
import { frameFacts, shapeText } from '../src/template/measure.ts';

const root = join(dirname(fileURLToPath(import.meta.url)), '..');
const demoPptx = readFileSync(join(root, 'templates', 'demo.pptx'));
const demoManifest = JSON.parse(readFileSync(join(root, 'templates', 'demo.manifest.json'), 'utf8')) as TemplateManifest;

const SAMPLE = 'Subtitle goes here';
const KICKER = '901';

/**
 * The demo template with two things many house templates have: the sample "Subtitle goes here"
 * under the title of every content slide (2 to 11), and a green divider through the content of
 * slides 4 to 7, in the same place on each, so that it counts as part of the frame. Slide 7 has
 * one more divider of its own, between its columns, which a layout may use as a rule.
 */
function houseStylePptx(): Uint8Array {
  const pkg = PptxPackage.load(demoPptx);
  const info = inspectTemplate(pkg);
  const emu = (x: number, y: number, w: number, h: number) =>
    `<a:xfrm><a:off x="${Math.round(x * info.slideWidth)}" y="${Math.round(y * info.slideHeight)}"/><a:ext cx="${Math.round(w * info.slideWidth)}" cy="${Math.round(h * info.slideHeight)}"/></a:xfrm>`;
  const kicker =
    `<p:sp><p:nvSpPr><p:cNvPr id="${KICKER}" name="Kicker"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>` +
    `<p:spPr>${emu(0.6, 0.245, 0.35, 0.04)}<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr>` +
    `<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr lang="en-US" sz="1400"/><a:t>${SAMPLE}</a:t></a:r></a:p></p:txBody></p:sp>`;
  const line = (id: string, x: number, y: number, h: number) =>
    `<p:cxnSp><p:nvCxnSpPr><p:cNvPr id="${id}" name="Divider"/><p:cNvCxnSpPr/><p:nvPr/></p:nvCxnSpPr>` +
    `<p:spPr>${emu(x, y, 0, h)}<a:prstGeom prst="line"><a:avLst/></a:prstGeom><a:ln w="9525"><a:solidFill><a:srgbClr val="00A252"/></a:solidFill></a:ln></p:spPr></p:cxnSp>`;
  readPresentation(pkg).slides.forEach((ref, index) => {
    const number = index + 1;
    let added = '';
    if (number >= 2 && number <= 11) added += kicker;
    if (number >= 4 && number <= 7) added += line('902', 0.55, 0.15, 0.75);
    if (number === 7) added += line('903', 0.5, 0.3, 0.58);
    if (added) pkg.write(ref.part, pkg.readText(ref.part).replace('</p:spTree>', `${added}</p:spTree>`));
  });
  return pkg.save();
}

/** The demo's manifest, with the sample under each title made a deck field, as a draft before 0.5.1 made it. */
function draftedAsField(): TemplateManifest {
  const manifest = structuredClone(demoManifest);
  for (const pattern of manifest.patterns) {
    if (pattern.sourceSlide >= 2 && pattern.sourceSlide <= 11) pattern.slots.push({ name: 'kicker', kind: 'text', shapeId: KICKER, maxChars: 40, deckField: 'kicker' });
  }
  return manifest;
}

const pptx = houseStylePptx();

test('a text repeated in the header or footer is one value for the deck; one repeated under the titles is not', () => {
  const facts = frameFacts(inspectTemplate(PptxPackage.load(pptx)));
  assert.ok(facts.repeatedTexts.includes(SAMPLE), 'the sample repeats on ten slides');
  assert.ok(!facts.fieldTexts.includes(SAMPLE), 'it lies under the title, in the content');
  assert.ok(facts.fieldTexts.includes('Organization name | Department'), 'the name in the header is the deck\'s');
});

test('a draft asks about the header and footer texts only, and makes a field of none the model calls a sample', async () => {
  const requests: JsonRequest[] = [];
  const llm: LlmProvider = {
    async generateJson(request) {
      requests.push(request);
      // The model says the name in the header is a sample too: then it is not a field either.
      if (request.prompt.startsWith('Sample texts')) return { language: 'en', fields: [{ text: 'Organization name | Department', name: 'organization', each: 'slide' }] };
      return { usable: true, reason: '', id: 'list-with-photo', description: 'A list beside a photo.', figures: false, shapes: [{ shape: '5', role: 'title', name: 'title', optional: false }] };
    },
  };
  const { manifest } = await draftManifest(pptx, { llm, slides: [4] });
  const asked = requests.find((request) => request.prompt.startsWith('Sample texts'))?.prompt.split('Give each a short camelCase field name:')[1] ?? '';
  assert.match(asked, /- "Organization name \| Department"/);
  assert.doesNotMatch(asked, new RegExp(`- "${SAMPLE}"`), 'the sample under the titles is not offered as a field');
  const slideCall = requests.find((request) => request.prompt.startsWith('Slide 4'));
  assert.match(slideCall?.prompt ?? '', new RegExp(`#${KICKER} text .*"${SAMPLE}"`), 'it is listed with the slide\'s texts, for the model to name');
  assert.match(slideCall?.prompt ?? '', /"Organization name \| Department"/, 'a text the model called a sample is listed with the slide\'s texts');
  assert.ok(!manifest.patterns.some((pattern) => pattern.slots.some((slot) => slot.kind === 'text' && slot.deckField)), 'no deck field was made');
});

test('prepareManifest makes a field drafted from a sample under the titles a text of each slide, and leaves the header\'s', () => {
  const template = loadTemplate(pptx, draftedAsField());
  const prepared = prepareManifest(template, { icons: true, variants: [], widen: false, remove: false, layout: 'adaptive' });
  const kickers = prepared.manifest.patterns.flatMap((pattern) => pattern.slots.filter((slot) => slot.name === 'kicker'));
  assert.ok(kickers.length >= 10);
  for (const slot of kickers) {
    assert.equal(slot.kind === 'text' ? slot.deckField : 'not text', undefined, 'no longer a deck field');
    assert.equal(slot.empty, 'remove', 'a slide may leave it out');
  }
  assert.equal(prepared.fields.kicker, undefined, 'the app is not given the sample as a value for the deck');
  assert.equal(prepared.fields.organizationDepartment, 'Organization name | Department');
  assert.ok(prepared.done.some((sentence) => sentence.includes('"kicker" was one value for the whole deck')), prepared.done.join('\n'));
  assert.ok(!prepared.left.some((sentence) => sentence.includes(SAMPLE)), 'nothing asks a person to set the sample');

  // A deck that says nothing about it renders without the sample: the slide leaves it out.
  const ready = loadTemplate(pptx, prepared.manifest);
  const { deck, assets } = sampleDeck(ready);
  for (const slide of deck.slides) delete slide.slots.kicker;
  const rendered = inspectTemplate(PptxPackage.load(renderDeck(ready, deck, { assets, layout: 'adaptive' }).pptx));
  assert.ok(!rendered.slides.some((slide) => slide.shapes.some((shape) => shapeText(shape) === SAMPLE)), 'no slide shows the sample');

  // Once is enough: a second run finds nothing more to do about it.
  const again = prepareManifest(loadTemplate(pptx, prepared.manifest), { icons: true, variants: [], widen: false, remove: false, layout: 'adaptive' });
  assert.deepEqual(again.manifest, prepared.manifest);
});

// --- composing on such a frame ------------------------------------------------------------

const text = (name: string, style: string, column: number, row: number, width: number, height: number, sample: string): DesignElement => ({ name, style, role: 'text', optional: false, sample, column, row, width, height });
const decoration = (style: string, column: number, row: number, width: number, height: number): DesignElement => ({ name: '', style, role: 'decoration', optional: false, sample: '', column, row, width, height });

function styles(template: Template) {
  const kit = designKit(template, undefined, BASE);
  assert.ok(kit);
  const all = kit.styles;
  const find = (test: (entry: (typeof all)[number]) => boolean, what: string) => {
    const style = all.find(test);
    assert.ok(style, `the kit has ${what}`);
    return style.key;
  };
  return {
    heading: find((style) => style.text.includes('used for card1Heading'), 'a card heading'),
    body: find((style) => style.text.includes('used for card1Text'), 'a card text'),
    panel: find((style) => style.kind === 'decoration' && style.shape === 'area' && style.text.includes('panel'), 'a panel'),
    rule: find((style) => style.kind === 'decoration' && style.shape === 'down', 'a rule down'),
  };
}

/**
 * Three cards of three columns each, with a gap of one column between them, and rules where asked.
 * The texts on a card take its first `wide` columns.
 */
function cards(template: Template, rules: Array<[column: number, row: number, height: number]>, wide = 3): PatternDesign {
  const style = styles(template);
  const elements: DesignElement[] = [];
  for (const [n, column] of [[1, 1], [2, 5], [3, 9]] as const) {
    elements.push(decoration(style.panel, column, 4, 3, 8));
    elements.push(text(`point${n}Heading`, style.heading, column, 4, wide, 2, `Point ${n}`));
    elements.push(text(`point${n}Text`, style.body, column, 6, wide, 6, 'A sentence or two about this point.'));
  }
  for (const [column, row, height] of rules) elements.push(decoration(style.rule, column, row, 1, height));
  return { id: 'three-cards', description: 'Three points side by side on cards.', keep: ['title'], elements, base: BASE };
}

/** The layout on slide 4, whose frame has the divider. */
const BASE = 'list-with-photo';

const lastSlide = (made: { pptx: Uint8Array }) => inspectTemplate(PptxPackage.load(made.pptx)).slides.at(-1)?.shapes ?? [];

test('a layout composed on a slide with a divider and a drafted sample keeps neither', () => {
  const template = loadTemplate(pptx, draftedAsField());
  const facts = frameFacts(inspectTemplate(template.pkg));
  const base = inspectTemplate(template.pkg).slides[3];
  assert.ok(base?.shapes.some((shape) => shape.id === '902' && facts.isFrame(shape)), 'the divider is part of the frame of slide 4, which the layout stands on');
  assert.equal(designKit(template, undefined, BASE)?.baseSlide, 4);

  const made = buildPattern(template, cards(template, []));
  const shapes = lastSlide(made);
  assert.ok(!shapes.some((shape) => shape.id === '902'), 'the divider through the content is not kept: the cards would lie across it');
  assert.ok(!shapes.some((shape) => shapeText(shape) === SAMPLE), 'the sample under the title is not kept');
  assert.ok(!made.manifest.patterns.at(-1)?.slots.some((slot) => slot.name === 'kicker'), 'nor is it a slot of the new layout');
  assert.ok(shapes.some((shape) => shapeText(shape) === 'Organization name | Department'), 'the deck\'s name in the header stays');
  assert.ok(shapes.some((shape) => shape.kind === 'picture' && shape.box.y < 0.15), 'the logo stays');
});

test('a rule between two cards is kept; one that cuts through a card is left out, with a warning', () => {
  const template = loadTemplate(pptx, demoManifest);
  const rules = (made: { pptx: Uint8Array }) => lastSlide(made).filter((shape) => shape.kind === 'line' && shape.id !== '902');

  // In the gaps, columns 4 and 8: each stays between its two cards.
  const between = buildPattern(template, cards(template, [[4, 4, 8], [8, 4, 8]]));
  assert.equal(rules(between).length, 2);
  assert.deepEqual(between.warnings.filter((warning) => warning.includes('rule')), []);

  // Down the last column of the middle card, from above it to the bottom: the texts on the cards
  // leave that column free, so the design passes, but drawn the rule runs through the card.
  const through = buildPattern(template, cards(template, [[7, 3, 10]], 2));
  assert.equal(rules(through).length, 0);
  assert.ok(through.warnings.some((warning) => /The rule D\d+ \(columns 7-7, rows 3-12\) runs through the panel/.test(warning)), through.warnings.join('\n'));
  // The rest of the layout is as designed.
  assert.equal(through.manifest.patterns.at(-1)?.repeat?.max, 3);
});
