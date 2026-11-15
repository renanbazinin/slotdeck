// A new layout composed from the template's own parts: the model designs it on a grid with the
// template's text styles and decorations, the code builds a real slide from the frame of one the
// template has, and the slide becomes a pattern. No test calls a live model.
import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { test } from 'node:test';
import { fileURLToPath } from 'node:url';
import {
  DESIGN_GRID,
  DesignError,
  PptxPackage,
  buildPattern,
  checkManifest,
  composePattern,
  designKit,
  inspectTemplate,
  loadTemplate,
  renderDeck,
  sampleDeck,
  validateDeck,
} from '../src/index.ts';
import type { DeckSpec, DesignElement, JsonRequest, LlmProvider, PatternDesign, Template, TemplateManifest } from '../src/index.ts';
import { NS, descendants, parseXml, rootElement } from '../src/pptx/xml.ts';

const root = join(dirname(fileURLToPath(import.meta.url)), '..');
const pptx = readFileSync(join(root, 'templates', 'demo.pptx'));
const manifest = JSON.parse(readFileSync(join(root, 'templates', 'demo.manifest.json'), 'utf8')) as TemplateManifest;
const demo = (): Template => loadTemplate(pptx, manifest);

/** The kit's style for a slot of the demo template, found by what the kit says it is used for. */
function styleFor(template: Template, use: string): string {
  const style = designKit(template)?.styles.find((candidate) => candidate.text.includes(`used for ${use}`));
  assert.ok(style, `the kit has a style used for ${use}`);
  return style.key;
}
const panel = (template: Template) => (designKit(template)?.styles.find((style) => style.kind === 'decoration' && style.shape === 'area' && style.text.includes('panel')) as { key: string }).key;

const text = (name: string, style: string, column: number, row: number, width: number, height: number, sample: string, role: DesignElement['role'] = 'text'): DesignElement => ({
  name,
  style,
  role,
  optional: false,
  sample,
  column,
  row,
  width,
  height,
});
const decoration = (style: string, column: number, row: number, width: number, height: number): DesignElement => ({ name: '', style, role: 'decoration', optional: false, sample: '', column, row, width, height });

/** Three cards across the width, each a panel with a heading and a sentence on it. */
function threeCards(template: Template, id = 'three-cards'): PatternDesign {
  const heading = styleFor(template, 'card1Heading');
  const body = styleFor(template, 'card1Text');
  const card = panel(template);
  const elements: DesignElement[] = [];
  for (const [n, column] of [[1, 1], [2, 5], [3, 9]] as const) {
    elements.push(decoration(card, column, 4, 4, 8));
    elements.push(text(`point${n}Heading`, heading, column, 4, 4, 2, `Point number ${n}`, 'subtitle'));
    elements.push(text(`point${n}Text`, body, column, 6, 4, 6, 'A sentence or two that says what this point means for the reader.'));
  }
  return { id, description: 'Three points side by side on cards, each with a heading and a sentence.', keep: ['title'], elements };
}

test('designKit reads the grid, the slide to stand on, and the template\'s own styles', () => {
  const template = demo();
  const kit = designKit(template);
  assert.ok(kit);
  assert.equal(kit.base, 'list-with-photo', 'the base is a layout whose title sits where most titles do');
  assert.deepEqual(kit.keepable.map((entry) => entry.slot), ['title']);
  assert.ok(kit.area.x > 0 && kit.area.y > 0 && kit.area.x + kit.area.w < 1 && kit.area.y + kit.area.h < 1, 'the grid lies inside the slide');
  assert.equal(kit.taken.length, DESIGN_GRID);
  const keys = kit.styles.map((style) => style.key);
  assert.equal(new Set(keys).size, keys.length, 'every style has its own key');
  assert.ok(kit.styles.some((style) => style.kind === 'text' && style.bullet), 'a style with bullets is offered for lists');
  assert.ok(kit.styles.some((style) => style.kind === 'decoration' && style.shape === 'area'), 'the cards\' panel is offered');
  assert.ok(!kit.styles.some((style) => /organizationDepartment/.test(style.text)), 'the deck\'s own texts are not offered as styles');
  assert.ok(!kit.styles.some((style) => style.kind === 'decoration' && style.text.includes('filled FFFFFF')), 'the white disc behind each card\'s icon is not offered: without its icon it is an empty circle');
});

test('buildPattern adds a slide made of the frame and copies of the template\'s shapes, and a pattern for it', () => {
  const template = demo();
  const made = buildPattern(template, threeCards(template));
  const before = PptxPackage.load(pptx);
  const after = PptxPackage.load(made.pptx);
  const slides = inspectTemplate(after).slides;
  assert.equal(slides.length, inspectTemplate(before).slides.length + 1);
  assert.equal(made.slide, slides.length);
  for (const slide of inspectTemplate(before).slides) assert.equal(after.readText(slide.part), before.readText(slide.part), `slide ${slide.number} is untouched`);

  const pattern = made.manifest.patterns.at(-1);
  assert.equal(pattern?.id, 'three-cards');
  assert.equal(pattern?.sourceSlide, made.slide);
  assert.deepEqual(made.manifest.patterns.slice(0, -1), manifest.patterns, 'the other patterns are as they were');
  const names = pattern?.slots.map((slot) => slot.name) ?? [];
  assert.ok(names.includes('title') && names.includes('organizationDepartment'), 'the kept title and the deck\'s own text are slots, as on the base slide');
  for (const n of [1, 2, 3]) assert.ok(names.includes(`point${n}Heading`) && names.includes(`point${n}Text`));
  assert.deepEqual({ min: pattern?.repeat?.min, max: pattern?.repeat?.max }, { min: 2, max: 3 }, 'numbered parts make a row');
  assert.ok((pattern?.repeat?.widen ?? 1) > 1, 'with fewer items the row widens');

  // The new slide: the base slide's title in its place, its photo and list gone, ids all different.
  const base = slides[3];
  const fresh = slides.at(-1);
  const baseTitle = manifest.patterns.find((candidate) => candidate.id === 'list-with-photo')?.slots.find((slot) => slot.name === 'title')?.shapeId;
  assert.deepEqual(fresh?.shapes.find((shape) => shape.id === baseTitle)?.box, base?.shapes.find((shape) => shape.id === baseTitle)?.box);
  assert.ok(!fresh?.shapes.some((shape) => shape.kind === 'picture' && !base?.shapes.some((other) => other.id === shape.id && other.kind === 'picture' && !/photo/i.test(other.name))), 'no photo is carried over');
  const ids = fresh?.shapes.map((shape) => shape.id) ?? [];
  assert.equal(new Set(ids).size, ids.length);
  const list = rootElement(parseXml(after.readText('ppt/presentation.xml')));
  const slideIds = descendants(list, NS.p, 'sldId').map((entry) => entry.getAttribute('id'));
  assert.equal(new Set(slideIds).size, slideIds.length);

  // It is a pattern like any other: checked, sampled at fewest and most, and filled by a deck.
  const built = loadTemplate(made.pptx, made.manifest);
  assert.deepEqual(checkManifest(built).filter((issue) => issue.pattern === 'three-cards' && issue.severity === 'error'), []);
  for (const items of ['min', 'max'] as const) {
    const { deck, assets } = sampleDeck(built, { items });
    renderDeck(built, deck, { assets, layout: 'adaptive' });
  }
  const deck: DeckSpec = {
    version: 1,
    slides: [{ pattern: 'three-cards', items: 2, slots: { title: 'Two things to know', point1Heading: 'Speed', point1Text: 'It is fast.', point2Heading: 'Cost', point2Text: 'It costs less.' } }],
    fields: { organizationDepartment: 'Team' },
  };
  assert.deepEqual(validateDeck(made.manifest, deck, { layout: 'adaptive' }).filter((issue) => issue.severity === 'error'), []);
  const rendered = PptxPackage.load(renderDeck(built, deck, { layout: 'adaptive' }).pptx);
  const slide = rendered.readText('ppt/slides/slide1.xml');
  for (const words of ['Two things to know', 'Speed', 'It costs less.']) assert.ok(slide.includes(words));
  assert.ok(!slide.includes('Point number 3'), 'the third card is gone with two items');
});

test('a layout composed earlier does not change the kit, so a design is built the same after another was added', () => {
  const template = demo();
  const first = buildPattern(template, threeCards(template, 'first-cards'));
  const extended = loadTemplate(first.pptx, first.manifest);
  assert.deepEqual(designKit(extended), designKit(template), 'the same base, grid and style keys');
  const second = buildPattern(extended, threeCards(template, 'second-cards'));
  assert.deepEqual(second.manifest.patterns.map((pattern) => pattern.id).slice(-2), ['first-cards', 'second-cards']);
  assert.equal(second.slide, first.slide + 1);
});

test('a composed budget is what the box holds, not what the model\'s sample was', () => {
  const template = demo();
  const design = threeCards(template);
  const long = design.elements.map((element) => (element.name === 'point1Heading' ? { ...element, sample: 'A heading far longer than the box it was given can ever hold on one line' } : element));
  const budget = (made: ReturnType<typeof buildPattern>, name: string) => (made.manifest.patterns.at(-1)?.slots.find((slot) => slot.name === name) as { maxChars?: number }).maxChars ?? 0;
  const plain = buildPattern(template, design);
  const longer = buildPattern(template, { ...design, elements: long });
  assert.equal(budget(longer, 'point1Heading'), budget(plain, 'point1Heading'), 'a long sample does not raise the budget');
  assert.ok(budget(longer, 'point1Heading') < 40);
});

test('buildPattern refuses a design it cannot build, with every reason', () => {
  const template = demo();
  const heading = styleFor(template, 'card1Heading');
  const body = styleFor(template, 'card1Text');
  const card = panel(template);
  const design: PatternDesign = {
    id: 'cover',
    description: 'A layout with every mistake.',
    keep: ['title', 'subtitle'],
    elements: [
      text('overTitle', body, 1, 1, 4, 2, 'Sits on the kept title.'),
      text('first', body, 1, 5, 6, 3, 'One text.'),
      text('second', body, 4, 6, 6, 3, 'Another text on it.'),
      text('across', heading, 9, 5, 4, 2, 'Half on a panel.'),
      decoration(card, 9, 6, 4, 4),
      text('onPanel', body, 9, 8, 4, 2, 'On the panel, listed before it.'),
      decoration(card, 9, 7, 4, 4),
      text('outside', body, 10, 11, 4, 3, 'Off the grid.'),
      text('styled', 'T99', 1, 10, 2, 2, 'No such style.'),
      decoration(body, 1, 9, 2, 1),
    ],
  };
  assert.throws(
    () => buildPattern(template, design),
    (error: unknown) => {
      assert.ok(error instanceof DesignError);
      const all = error.problems.join('\n');
      for (const words of [
        'already has a layout called "cover"',
        '"subtitle" cannot be kept',
        'covers the kept title',
        'lie on each other',
        'reaches over the edge of the panel',
        'would hide it',
        'outside the grid',
        '"T99"',
        'text roles take T styles and decorations take D styles',
      ]) {
        assert.ok(all.includes(words), `a reason says: ${words}\n${all}`);
      }
      return true;
    },
  );
});

/** A model that answers from a list, and keeps what it was asked. */
function scripted(answers: unknown[]): LlmProvider & { requests: JsonRequest[] } {
  const requests: JsonRequest[] = [];
  return {
    requests,
    async generateJson(request) {
      requests.push(request);
      const answer = answers[requests.length - 1];
      if (answer === undefined) throw new Error('no more answers');
      return answer;
    },
  };
}

test('composePattern sends a design that leaves the slide half empty back once, and builds the second', async () => {
  const template = demo();
  const good = threeCards(template, 'three-points');
  const half: PatternDesign = { ...good, elements: good.elements.filter((element) => element.column === 1) };
  const llm = scripted([half, good]);
  const made = await composePattern(template, 'A text-only layout with three points', { llm, effort: 'low' });
  assert.equal(llm.requests.length, 2);
  const [first, second] = llm.requests as [JsonRequest, JsonRequest];
  assert.equal(first.effort, 'low');
  assert.ok(first.maxOutputTokens !== undefined, 'the answer has a limit');
  assert.ok(first.prompt.includes('A text-only layout with three points'));
  assert.ok(first.prompt.includes('Text styles') && first.prompt.includes('Decorations'));
  const items = (first.schema.properties as { elements: { items: { properties: { style: { enum: string[] } } } } }).elements.items.properties.style.enum;
  assert.deepEqual(items, designKit(template)?.styles.map((style) => style.key), 'the model can only name the template\'s styles');
  assert.match(second.prompt, /columns 5-12[^\n]*stays empty/, 'the second request says what was wrong');
  assert.equal(made.pattern, 'three-points');
  assert.deepEqual(made.warnings, []);
  assert.match(made.summary, /new layout on slide 14/);
});

test('composePattern keeps a first answer that is better than the second, and fails when neither can be built', async () => {
  const template = demo();
  const good = threeCards(template);
  const half: PatternDesign = { ...good, elements: good.elements.filter((element) => element.column === 1) };
  const broken: PatternDesign = { ...good, elements: [...good.elements, text('point1Text', 'T1', 1, 1, 2, 2, 'Again.')] };

  const kept = await composePattern(template, 'three points', { llm: scripted([half, broken]) });
  assert.equal(kept.design.elements.length, 3, 'the half-empty first answer is built, as it can be');
  assert.ok(kept.warnings.some((warning) => warning.includes('stays empty')), 'and the warning says what is wrong with it');

  await assert.rejects(composePattern(template, 'three points', { llm: scripted([broken, broken]) }), (error: unknown) => error instanceof DesignError && error.problems.some((problem) => problem.includes('is taken')));
  await assert.rejects(composePattern(template, '  ', { llm: scripted([]) }), /Say what the new layout should be/);
});

test('numbered items of different widths are made equal, so that a deck lays them out without overlap', () => {
  const template = demo();
  const heading = styleFor(template, 'card1Heading');
  const body = styleFor(template, 'card1Text');
  const widths = [3, 2, 2, 2, 3];
  const elements: DesignElement[] = [];
  let column = 1;
  widths.forEach((width, index) => {
    elements.push(text(`step${index + 1}Heading`, heading, column, 4, width, 2, `Step ${index + 1}`, 'subtitle'));
    elements.push(text(`step${index + 1}Text`, body, column, 6, width, 5, 'What happens in this step.'));
    column += width;
  });
  const made = buildPattern(template, { id: 'five-steps', description: 'Five steps in a row.', keep: ['title'], elements });
  const slide = inspectTemplate(PptxPackage.load(made.pptx)).slides.at(-1);
  const pattern = made.manifest.patterns.at(-1);
  const boxes = [1, 2, 3, 4, 5].map((n) => slide?.shapes.find((shape) => shape.id === pattern?.slots.find((slot) => slot.name === `step${n}Text`)?.shapeId)?.box as { x: number; w: number });
  const sizes = boxes.map((box) => box.w);
  assert.ok(Math.max(...sizes) - Math.min(...sizes) < 0.002, `the steps are equally wide: ${sizes.join(', ')}`);
  const gaps = boxes.slice(1).map((box, index) => box.x - ((boxes[index] as { x: number; w: number }).x + (boxes[index] as { x: number; w: number }).w));
  assert.ok(Math.min(...gaps) > 0 && Math.max(...gaps) - Math.min(...gaps) < 0.002, `and equally far apart: ${gaps.join(', ')}`);
  assert.equal(pattern?.repeat?.max, 5);
  const built = loadTemplate(made.pptx, made.manifest);
  const { deck, assets } = sampleDeck(built, { items: 'max' });
  renderDeck(built, deck, { assets, layout: 'adaptive' });
});
