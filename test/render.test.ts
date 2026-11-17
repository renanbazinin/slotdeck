import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
import { test } from 'node:test';
import { placeholderPhoto } from '../examples/placeholder.ts';
import { DeckValidationError, PptxPackage, loadTemplate, renderDeck, validateDeck } from '../src/index.ts';
import type { DeckSpec, RenderOptions, SlideSpec, TemplateManifest } from '../src/index.ts';
import { ContentTypes, readRels, relsOwner, resolveTarget } from '../src/pptx/package.ts';
import { readPresentation } from '../src/pptx/presentation.ts';
import { NS, childElement, childElements, descendants, parseXml, path, rootElement } from '../src/pptx/xml.ts';
import type { Element } from '../src/pptx/xml.ts';
import { writable } from '../src/util/chars.ts';
import { buildFixturePptx, fixtureManifest } from './fixture.ts';

const fixture = buildFixturePptx();
const square = placeholderPhoto(100, 100, [10, 20, 30], [200, 210, 220]);
const wide = placeholderPhoto(400, 100, [10, 20, 30], [200, 210, 220]);
const assets = { square, wide };

const contentSlide = (slots: SlideSpec['slots'] = {}): SlideSpec => ({
  pattern: 'content',
  slots: { title: 'שלום', body: ['כותרת', 'גוף'], photo: { asset: 'square' }, note: 'הערה', ...slots },
});

function render(deck: Partial<DeckSpec> & { slides: SlideSpec[] }, options: RenderOptions = {}, manifest: TemplateManifest = fixtureManifest) {
  const result = renderDeck(loadTemplate(fixture, manifest), { version: 1, ...deck }, { assets, ...options });
  const pkg = PptxPackage.load(result.pptx);
  const slide = (number: number) => rootElement(parseXml(pkg.readText(`ppt/slides/slide${number}.xml`)));
  return { ...result, pkg, slide };
}

function shape(root: Element, id: string): Element | undefined {
  const props = descendants(root, NS.p, 'cNvPr').find((el) => el.getAttribute('id') === id);
  return props?.parentNode?.parentNode as Element | undefined;
}

function paragraphs(root: Element, id: string): Element[] {
  const body = childElement(shape(root, id) as Element, NS.p, 'txBody') as Element;
  return childElements(body, NS.a, 'p');
}

const textOf = (el: Element) =>
  descendants(el, NS.a, 't')
    .map((t) => t.textContent ?? '')
    .join('');

/** Every relationship resolves, every part has a content type, and nothing is orphaned. */
function assertIntact(pkg: PptxPackage): void {
  const files = pkg.paths();
  const types = rootElement(parseXml(pkg.readText('[Content_Types].xml')));
  const defaults = new Set(childElements(types, NS.contentTypes, 'Default').map((el) => el.getAttribute('Extension')));
  const overrides = new Set(childElements(types, NS.contentTypes, 'Override').map((el) => el.getAttribute('PartName')));
  for (const override of overrides) assert.ok(pkg.has((override as string).slice(1)), `override for missing part ${override}`);

  const reachable = new Set<string>();
  for (const file of files) {
    const owner = relsOwner(file);
    if (owner === undefined) continue;
    assert.ok(owner === '' || pkg.has(owner), `relationships without an owner: ${file}`);
    for (const rel of readRels(pkg, owner)) {
      const target = resolveTarget(owner, rel.target);
      assert.ok(pkg.has(target), `${file}: ${rel.id} points at missing part ${target}`);
      reachable.add(target);
    }
  }
  for (const file of files) {
    if (file === '[Content_Types].xml' || relsOwner(file) !== undefined) continue;
    assert.ok(reachable.has(file), `orphaned part ${file}`);
    const extension = file.slice(file.lastIndexOf('.') + 1);
    assert.ok(overrides.has(`/${file}`) || defaults.has(extension), `no content type for ${file}`);
  }
}

test('renders deck slides in order and reuses a pattern', () => {
  const { pkg, slide } = render({
    fields: { closing: 'תודה' },
    slides: [contentSlide({ title: 'ראשון' }), contentSlide({ title: 'שני' }), { pattern: 'closing', slots: {} }],
  });

  const pres = readPresentation(pkg);
  assert.deepEqual(
    pres.slides.map((s) => s.part),
    ['ppt/slides/slide1.xml', 'ppt/slides/slide2.xml', 'ppt/slides/slide3.xml'],
  );
  assert.equal(new Set(pres.slides.map((s) => s.id)).size, 3);
  assert.equal(textOf(shape(slide(1), '2') as Element), 'ראשון');
  assert.equal(textOf(shape(slide(2), '2') as Element), 'שני');
  assert.equal(textOf(shape(slide(3), '2') as Element), 'תודה', 'deck field fills the slot');
  assertIntact(pkg);
});

test('leaves no template sample text or stale parts behind', () => {
  const { pkg } = render({ slides: [contentSlide()] });
  for (const file of pkg.paths().filter((name) => name.endsWith('.xml'))) {
    assert.ok(!pkg.readText(file).includes('SAMPLE'), `sample text left in ${file}`);
  }
  assert.ok(!pkg.has('ppt/slides/slide2.xml'), 'unused template slide removed');
  assert.ok(!pkg.has('ppt/notesSlides/notesSlide1.xml'), 'template notes removed');
  assert.ok(!pkg.has('ppt/changesInfos/changesInfo1.xml'), 'revision history removed');
  assert.ok(!pkg.readText('ppt/presentation.xml').includes('sectionLst'), 'section list removed');
  assert.ok(pkg.has('ppt/media/image1.png'), 'media still used by the kept icon stays');
  assertIntact(pkg);
});

test('drops template media once nothing uses it', () => {
  const { pkg } = render({ slides: [contentSlide({ icon: { asset: 'square' } })] });
  assert.ok(!pkg.has('ppt/media/image1.png'));
  assertIntact(pkg);
});

test('stores an asset once however often it is used', () => {
  const { pkg } = render({ slides: [contentSlide(), contentSlide({ icon: { asset: 'square' } })] });
  const media = pkg.paths().filter((name) => name.startsWith('ppt/media/slotdeck-'));
  assert.equal(media.length, 1);
  assertIntact(pkg);
});

test('new paragraphs take the formatting of the template paragraph in the same position', () => {
  const { slide } = render({ slides: [contentSlide({ body: ['א', 'ב', 'ג'] })] });
  const sizes = paragraphs(slide(1), '3').map((p) => {
    const runs = childElements(p, NS.a, 'r');
    assert.equal(runs.length, 1, 'one run per paragraph');
    assert.equal(childElements(p, NS.a, 'br').length, 0, 'template line breaks are gone');
    return childElement(runs[0] as Element, NS.a, 'rPr')?.getAttribute('sz');
  });
  assert.deepEqual(sizes, ['2000', '1200', '1200']);
});

test('styleRun picks which run of a mixed paragraph supplies the formatting', () => {
  const { slide } = render({ slides: [contentSlide()] });
  const run = childElement(paragraphs(slide(1), '2')[0] as Element, NS.a, 'r') as Element;
  assert.equal(childElement(run, NS.a, 'rPr')?.getAttribute('b'), null, 'the plain run, not the bold label');
});

test('paragraph direction follows the text; numbers keep the template direction', () => {
  const { slide } = render({ slides: [contentSlide({ body: ['עברית', 'English', '78%'] })] });
  const direction = paragraphs(slide(1), '3').map((p) => childElement(p, NS.a, 'pPr')?.getAttribute('rtl'));
  assert.deepEqual(direction, ['1', '0', '1']);
});

test('alignment flips for a deck in the other direction unless the manifest says keep', () => {
  const english = { direction: 'ltr' as const, slides: [contentSlide({ title: 'Hello', body: 'World' })] };
  const align = (root: Element, id: string) => childElement(paragraphs(root, id)[0] as Element, NS.a, 'pPr')?.getAttribute('algn');

  assert.equal(align(render(english).slide(1), '2'), 'l');
  assert.equal(align(render({ slides: [contentSlide()] }).slide(1), '2'), 'r', 'same direction: untouched');

  const keep: TemplateManifest = { ...fixtureManifest, alignOnDirectionChange: 'keep' };
  assert.equal(align(render(english, {}, keep).slide(1), '2'), 'r');

  const keepButFlipBody: TemplateManifest = {
    ...keep,
    patterns: keep.patterns.map((pattern) => ({
      ...pattern,
      slots: pattern.slots.map((slot) => (slot.name === 'body' && slot.kind === 'text' ? { ...slot, alignOnDirectionChange: 'flip' as const } : slot)),
    })),
  };
  const mixed = render(english, {}, keepButFlipBody).slide(1);
  assert.equal(align(mixed, '2'), 'r');
  assert.equal(align(mixed, '3'), 'l', 'slot override wins');
});

test('cover fit crops the image to the frame and keeps the frame', () => {
  const crop = (asset: string) => {
    const pic = shape(render({ slides: [contentSlide({ photo: { asset } })] }).slide(1), '4') as Element;
    const rect = path(pic, [NS.p, 'blipFill'], [NS.a, 'srcRect']);
    const ext = path(pic, [NS.p, 'spPr'], [NS.a, 'xfrm'], [NS.a, 'ext']);
    assert.equal(ext?.getAttribute('cx'), '4000000');
    assert.equal(ext?.getAttribute('cy'), '2000000');
    assert.equal(path(pic, [NS.p, 'spPr'], [NS.a, 'prstGeom'])?.getAttribute('prst'), 'roundRect', 'frame geometry kept');
    return ['l', 't', 'r', 'b'].map((side) => rect?.getAttribute(side) ?? null);
  };
  // The frame is 2:1. A square image loses a quarter top and bottom; a 4:1 image a quarter left and right.
  assert.deepEqual(crop('square'), [null, '25000', null, '25000']);
  assert.deepEqual(crop('wide'), ['25000', null, '25000', null]);
});

test('contain fit shrinks the frame around its centre, allowing for group scaling', () => {
  const pic = shape(render({ slides: [contentSlide({ icon: { asset: 'square' } })] }).slide(1), '8') as Element;
  const xfrm = path(pic, [NS.p, 'spPr'], [NS.a, 'xfrm']) as Element;
  // The group doubles widths, so a square image needs a frame half as wide as it is tall.
  assert.equal(childElement(xfrm, NS.a, 'ext')?.getAttribute('cx'), '250000');
  assert.equal(childElement(xfrm, NS.a, 'ext')?.getAttribute('cy'), '500000');
  assert.equal(childElement(xfrm, NS.a, 'off')?.getAttribute('x'), '325000');
  assert.equal(path(pic, [NS.p, 'blipFill'], [NS.a, 'srcRect']), undefined);
});

test('replaced pictures point at the new image and carry its alt text', () => {
  const { pkg, slide } = render({ slides: [contentSlide({ photo: { asset: 'wide', alt: 'A wide photo' } })] });
  const pic = shape(slide(1), '4') as Element;
  const relId = path(pic, [NS.p, 'blipFill'], [NS.a, 'blip'])?.getAttributeNS(NS.r, 'embed');
  const rel = readRels(pkg, 'ppt/slides/slide1.xml').find((r) => r.id === relId);
  assert.ok(rel, 'relationship exists');
  assert.deepEqual(pkg.read(resolveTarget('ppt/slides/slide1.xml', (rel as { target: string }).target)), wide);
  const props = path(pic, [NS.p, 'nvPicPr'], [NS.p, 'cNvPr']);
  assert.equal(props?.getAttribute('descr'), 'A wide photo');
  assert.equal(props?.getAttribute('title'), null);
});

test('slide-number fields show the new position', () => {
  const { slide } = render({ slides: [contentSlide(), contentSlide()] });
  assert.equal(textOf(shape(slide(2), '5') as Element), '2');
});

test('strict layout requires every slot; adaptive removes empty optional ones', () => {
  const withoutNote = contentSlide();
  delete withoutNote.slots.note;

  assert.throws(() => render({ slides: [withoutNote] }), DeckValidationError);

  const { slide } = render({ slides: [withoutNote] }, { layout: 'adaptive' });
  assert.equal(shape(slide(1), '6'), undefined, 'optional shape removed');
  assert.ok(shape(slide(1), '8'), "a 'keep' slot stays without a value");

  const dropIcon = render({ slides: [contentSlide({ icon: null })] }, { layout: 'adaptive' });
  assert.equal(shape(dropIcon.slide(1), '8'), undefined, 'explicit null removes a kept decoration');
});

test('validation reports problems without rendering', () => {
  const check = (slides: SlideSpec[], options = {}) =>
    validateDeck(fixtureManifest, { version: 1, slides }, { assets: Object.keys(assets), ...options });

  assert.deepEqual(check([contentSlide()]), []);
  assert.match(check([{ pattern: 'nope', slots: {} }])[0]?.message ?? '', /unknown pattern/);
  assert.match(check([contentSlide({ title: null })])[0]?.message ?? '', /required/);
  assert.match(check([contentSlide({ photo: { asset: 'missing' } })])[0]?.message ?? '', /was not provided/);
  assert.match(check([contentSlide({ photo: 'text' })])[0]?.message ?? '', /image reference/);
  assert.match(check([{ pattern: 'closing', slots: {} }])[0]?.message ?? '', /required/, 'deck field not set');

  const warnings = check([contentSlide({ title: 'x'.repeat(21), body: ['a', 'b', 'c', 'd'], extra: 'y' })]);
  assert.deepEqual(
    warnings.map((issue) => issue.severity),
    ['warning', 'warning', 'warning'],
  );
  assert.deepEqual(warnings.map((issue) => issue.slot).sort(), ['body', 'extra', 'title']);

  const result = render({ slides: [contentSlide({ title: 'x'.repeat(21) })] });
  assert.equal(result.warnings.length, 1, 'warnings are returned, not thrown');
});

test('a value that is no text is reported as one, and nothing throws', () => {
  const check = (slots: Record<string, unknown>) =>
    validateDeck(fixtureManifest, { version: 1, slides: [contentSlide(slots as SlideSpec['slots'])] }, { assets: Object.keys(assets) }).map((issue) => `${issue.slot}: ${issue.message}`);
  assert.deepEqual(check({ title: 2025 }), ['title: expected text']);
  assert.deepEqual(check({ title: true }), ['title: expected text']);
  assert.deepEqual(check({ body: ['One', 24] }), ['body: expected text']);
  assert.deepEqual(check({ body: ['One', null] }), ['body: expected text']);
  assert.deepEqual(check({ body: [{ text: 'One' }] }), ['body: expected text']);
  assert.deepEqual(check({ note: { asset: 'square' } }), ['note: expected text']);
  // Rendering answers with the same issue, as an error of the deck.
  assert.throws(() => render({ slides: [contentSlide({ title: 2025 } as unknown as SlideSpec['slots'])] }), (error: unknown) => error instanceof DeckValidationError && /title: expected text/.test(error.message));
});

test('characters that no file can hold never reach a part, and a vertical tab breaks the line', () => {
  // A vertical tab is the line break of text copied out of a text box; the others come with pasted text too.
  const { pkg, slide } = render({ slides: [contentSlide({ title: 'אחד\u000Bשניים\u0001', body: ['ראשון\u000C', 'שני\uFFFE\uD83D'], photo: { asset: 'wide', alt: 'A wide\u000Bphoto\u0008' } })] });
  const xml = pkg.readText('ppt/slides/slide1.xml');
  assert.doesNotMatch(xml, /[\u0000-\u0008\u000B\u000C\u000E-\u001F\uFFFE\uFFFF]|[\uD800-\uDBFF](?![\uDC00-\uDFFF])/);
  assert.deepEqual(paragraphs(slide(1), '2').map(textOf), ['אחד', 'שניים']);
  assert.deepEqual(paragraphs(slide(1), '3').map(textOf), ['ראשון', 'שני']);
  assert.equal(path(shape(slide(1), '4') as Element, [NS.p, 'nvPicPr'], [NS.p, 'cNvPr'])?.getAttribute('descr'), 'A wide photo');
  // The count that is checked is the count that is drawn.
  const issues = validateDeck(fixtureManifest, { version: 1, slides: [contentSlide({ title: `${'x'.repeat(20)}\u0001\u0002` })] }, { assets: Object.keys(assets) });
  assert.deepEqual(issues, []);
  // The two characters at the end of the plane go, and a pair of surrogates stays whole.
  assert.equal(writable('a\uFFFEb\uFFFFc\u{1F6B2}d'), 'abc\u{1F6B2}d');
  // The code names these characters by their numbers. Written as themselves they are invisible, and
  // a tool that drops such characters from a file would take them out of the check without a trace.
  assert.doesNotMatch(readFileSync(new URL('../src/util/chars.ts', import.meta.url), 'utf8'), /[^\t\n\r\x20-\x7E]/);
  // The tests that give such characters to the code name them by their numbers too: one written as
  // itself here would be dropped by the same tools, and the assertion over it would test nothing.
  for (const file of ['render.test.ts', 'free.test.ts']) {
    assert.doesNotMatch(readFileSync(new URL(`./${file}`, import.meta.url), 'utf8'), /[\u0000-\u0008\u000B\u000C\u000E-\u001F\uFFFE\uFFFF]/, file);
  }
});

test('content types list exactly the parts that exist', () => {
  const { pkg } = render({ slides: [contentSlide()] });
  const text = pkg.readText('[Content_Types].xml');
  assert.ok(text.includes('/ppt/slides/slide1.xml'));
  assert.ok(!text.includes('/ppt/slides/slide2.xml'));
  assert.ok(!text.includes('notesSlide'));
  ContentTypes.read(pkg); // still parses
});
