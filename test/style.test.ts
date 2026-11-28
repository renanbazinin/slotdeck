// Style-only mode (0.10.0): a manifest made from a pptx alone (`styleManifest`), enough of one for
// the template to carry a cover, a closing slide and the slides the software lays out itself; and
// what makes it work on an ordinary file, whose text is all in placeholders.
import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { test } from 'node:test';
import { fileURLToPath } from 'node:url';
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate';
import { pageColor } from '../src/color/recolor.ts';
import { freeKitOrReason, widened } from '../src/free/kit.ts';
import { drawnBoxes } from '../src/template/drawn.ts';
import { PptxPackage } from '../src/pptx/package.ts';
import { parseXml, rootElement } from '../src/pptx/xml.ts';
import { checkManifest, freeSlide, loadTemplate, renderDeck, styleManifest, validateDeck } from '../src/index.ts';
import type { DeckSpec, StyleManifest } from '../src/index.ts';
import { serializeXml } from '../src/pptx/xml.ts';
import { TextStyles } from '../src/text/styles.ts';

const root = join(dirname(fileURLToPath(import.meta.url)), '..');
const demoPptx = new Uint8Array(readFileSync(join(root, 'templates', 'demo.pptx')));

test('a manifest is made from the pptx alone, the same every time, and passes the check', () => {
  const style = styleManifest(demoPptx, { name: 'demo' });
  assert.deepEqual(styleManifest(demoPptx, { name: 'demo' }), style, 'the same file gives the same manifest');
  const { manifest } = style;
  assert.equal(manifest.name, 'demo');
  assert.equal(manifest.direction, 'ltr');
  assert.equal(manifest.language, 'en');
  // Every slide of the demo is read, and its first and last are the cover and the closing slide.
  assert.equal(manifest.patterns.length, 13);
  assert.equal(style.cover, 'cover');
  assert.equal(style.closing, 'closing');
  // What a deck in the template's style is planned from, and no reason why it could not be.
  assert.deepEqual(style.patterns, ['cover', 'free', 'closing']);
  assert.equal(style.reason, undefined);
  assert.equal(manifest.patterns[0]?.id, 'cover');
  assert.equal(manifest.patterns.at(-1)?.id, 'closing');
  for (const id of ['cover', 'closing']) assert.ok(manifest.patterns.find((pattern) => pattern.id === id)?.slots.some((slot) => slot.name === 'title'), `${id} has a title`);
  // The line of the organisation that every slide of the demo shows is the deck's own text: a field.
  assert.deepEqual(Object.values(style.fields), ['Organization name | Department']);
  assert.ok(manifest.patterns.some((pattern) => pattern.slots.some((slot) => slot.kind === 'text' && slot.deckField === 'field1')));
  // No slot is named twice on a slide, and none points at a shape the slide does not have.
  assert.deepEqual(checkManifest(loadTemplate(demoPptx, manifest)).filter((issue) => issue.severity === 'error'), []);
  for (const pattern of manifest.patterns) assert.equal(new Set(pattern.slots.map((slot) => slot.name)).size, pattern.slots.length, pattern.id);
});

test('the template\'s style is found without anyone naming a slot: its frame, its text styles, its colours', () => {
  const style = styleManifest(demoPptx);
  const template = loadTemplate(demoPptx, style.manifest);
  const kit = freeKitOrReason(template);
  assert.ok(typeof kit !== 'string', String(kit));
  if (typeof kit === 'string') return;
  // The frame is a slide whose title sits where most titles sit, with room under it and nothing drawn there.
  assert.equal(kit.base, (style.manifest.free as { base?: string }).base);
  assert.ok(kit.band.bottom - kit.band.top >= 0.45);
  assert.match(style.notes.join('\n'), /free slides stand on the frame of slide \d+: its title sits where the titles of \d+ slides sit, and nothing is drawn under it/);
  // The styles are the demo's own: running text at 18 pt, a bold heading over it, a big number in the main colour.
  assert.equal(kit.styles.body.size, 18);
  assert.ok(kit.styles.heading.bold && kit.styles.heading.size >= kit.styles.body.size && kit.styles.heading.set === undefined, 'a heading of the template, not one made up');
  assert.ok(kit.styles.figure.size >= 36 && kit.styles.figure.set === undefined, 'a big number of the template');
  assert.equal(kit.styles.figure.color, '0E7C7B');
  assert.equal(kit.palette.accent, '0E7C7B');
  assert.ok(kit.panel, 'the box its cards stand on');
});

test('a deck of the cover, slides the software lays out and the closing slide renders on it', () => {
  const style = styleManifest(demoPptx);
  const template = loadTemplate(demoPptx, style.manifest);
  const deck: DeckSpec = {
    version: 1,
    slides: [
      { pattern: 'cover', slots: { title: 'Riverbend 2027', subtitle: 'The plan for the city council' } },
      freeSlide({ title: 'A year in three numbers', tone: 'accent', blocks: [{ figure: { value: '34%', label: 'more passes' } }, { figure: { value: '11', label: 'new stations' } }] }),
      freeSlide({ title: 'How a station opens', blocks: [{ phase: { heading: 'Survey', text: 'Two weeks.' } }, { phase: { heading: 'Build', text: 'Three days.' } }, { phase: { heading: 'Launch' } }] }),
      { pattern: 'closing', slots: { title: 'Thank you' } },
    ],
  };
  assert.deepEqual(validateDeck(template, deck, { layout: 'adaptive' }).filter((issue) => issue.severity === 'error'), []);
  const files = unzipSync(renderDeck(template, deck, { layout: 'adaptive' }).pptx);
  const slide = (number: number) => strFromU8(files[`ppt/slides/slide${number}.xml`] as Uint8Array);
  assert.match(slide(1), /Riverbend 2027/);
  assert.match(slide(2), /name="Ground"/);
  assert.match(slide(3), /prst="homePlate"/);
  assert.match(slide(4), /Thank you/);
  // The deck gave the frame's own line no text, so no slide shows the template's sample of it.
  for (const number of [1, 2, 3, 4]) assert.doesNotMatch(slide(number), /Organization name/);
});

// A page of the template's colour in one corner and white over the rest: the look of many brand templates.
const page = (stops: Array<[number, string]>): string =>
  `<p:bg><p:bgPr><a:gradFill><a:gsLst>${stops.map(([pos, hex]) => `<a:gs pos="${pos}"><a:srgbClr val="${hex}"/></a:gs>`).join('')}</a:gsLst><a:path path="circle"><a:fillToRect l="100000" t="100000"/></a:path></a:gradFill><a:effectLst/></p:bgPr></p:bg>`;
const GLOW = page([[0, '7CF29C'], [30000, 'FFFFFF'], [100000, 'FFFFFF']]);
const BLEND = page([[0, '0E7C7B'], [100000, 'FFFFFF']]);
const withPage = (part: string, xml: string): Uint8Array => {
  const zip = unzipSync(demoPptx);
  zip[part] = strToU8(strFromU8(zip[part] as Uint8Array).replace('<p:spTree>', `${xml}<p:spTree>`));
  return zipSync(zip);
};
const pageOf = (pptx: Uint8Array, slide: number) => {
  const pkg = PptxPackage.load(pptx);
  const part = `ppt/slides/slide${slide}.xml`;
  return pageColor(pkg, part, rootElement(parseXml(pkg.readText(part))));
};

test('a page that is one colour but for a glow in a corner is read by that colour', () => {
  // Green at the corner, white from three tenths of the way on: white stands for 85% of the page.
  const glow = pageOf(withPage('ppt/slides/slide7.xml', GLOW), 7);
  assert.equal(glow.color, 'FFFFFF');
  assert.ok(Math.abs((glow.glow ?? 0) - 0.85) < 1e-9, String(glow.glow));
  // A blend of two colours is no page of one colour: it is read by its first, as it always was.
  assert.deepEqual(pageOf(withPage('ppt/slides/slide7.xml', BLEND), 7), { color: '0E7C7B' });
  // A page of one colour has no glow, and neither has a slide that states no page of its own.
  assert.deepEqual(pageOf(withPage('ppt/slides/slide7.xml', '<p:bg><p:bgPr><a:solidFill><a:srgbClr val="0E7C7B"/></a:solidFill><a:effectLst/></p:bgPr></p:bg>'), 7), { color: '0E7C7B' });
  assert.equal(pageOf(demoPptx, 7).glow, undefined);
  // The page a slide takes from its layout is read the same way.
  const layout = /slideLayouts\/(slideLayout\d+\.xml)/.exec(strFromU8(unzipSync(demoPptx)['ppt/slides/_rels/slide7.xml.rels'] as Uint8Array))?.[1];
  const inherited = pageOf(withPage(`ppt/slideLayouts/${layout}`, GLOW), 7);
  assert.equal(inherited.color, 'FFFFFF');
  assert.ok(Math.abs((inherited.glow ?? 0) - 0.85) < 1e-9);
});

test('free slides on a page with a glow choose their text for the page, and such a frame is taken before a plain one', () => {
  const plain = styleManifest(demoPptx);
  assert.equal((plain.manifest.free as { base?: string }).base, 'slide4');
  // Slide 7 is a frame as clear as slide 4: with the template's glow on its page, it is the one taken.
  const pptx = withPage('ppt/slides/slide7.xml', GLOW);
  const style = styleManifest(pptx);
  assert.equal((style.manifest.free as { base?: string }).base, 'slide7');
  assert.match(style.notes.join('\n'), /free slides stand on the frame of slide 7: .*its page is one colour over 85% of it, with the template's glow in the rest/);
  // A page that is a blend is not such a page: the first clear frame is taken, as before.
  assert.equal((styleManifest(withPage('ppt/slides/slide7.xml', BLEND)).manifest.free as { base?: string }).base, 'slide4');
  // The text of a free slide is chosen for the white the content stands on: the same styles as on the plain slide.
  const template = loadTemplate(pptx, style.manifest);
  const [kit, same] = [freeKitOrReason(template), freeKitOrReason(loadTemplate(demoPptx, { ...plain.manifest, free: { base: 'slide7' } }))];
  assert.ok(typeof kit !== 'string' && typeof same !== 'string', `${String(kit)} ${String(same)}`);
  if (typeof kit === 'string' || typeof same === 'string') return;
  assert.equal(kit.background, 'FFFFFF');
  assert.deepEqual(kit.styles, same.styles);
  assert.equal(kit.palette.accent, same.palette.accent);
  // And the slide that is made keeps the glow: it is the frame's own page.
  const deck: DeckSpec = { version: 1, slides: [freeSlide({ title: 'The road to 2027', blocks: [{ milestone: { when: 'Q1 2026', heading: 'North Bridge' } }, { milestone: { when: 'Q3 2026', heading: 'E-bikes arrive' } }] })] };
  assert.deepEqual(validateDeck(template, deck, { layout: 'adaptive' }).filter((issue) => issue.severity === 'error'), []);
  const made = strFromU8(unzipSync(renderDeck(template, deck, { layout: 'adaptive' }).pptx)['ppt/slides/slide1.xml'] as Uint8Array);
  assert.match(made, /<p:bg><p:bgPr><a:gradFill>/);
  assert.match(made, /7CF29C/);
});

// An ordinary file: a title slide, three slides of a title and a bulleted body, a closing title
// slide, on the default template of python-pptx. Every text is a placeholder that states neither
// its place nor its look: both come from the slide layout, the master and the master's text styles.
const plainPptx = new Uint8Array(readFileSync(join(root, 'test', 'fixtures', 'placeholders.pptx')));
const errors = (issues: Array<{ severity: string }>) => issues.filter((issue) => issue.severity === 'error');

test('a file whose text is all in placeholders is read where they stand and at the sizes they inherit', () => {
  const style: StyleManifest = styleManifest(plainPptx, { name: 'plain' });
  assert.equal(style.reason, undefined);
  assert.deepEqual(style.patterns, ['cover', 'free', 'closing']);
  // The cover's title is the 44 pt of the master's title style, with the room of the layout's box; the line under it is a subtitle.
  const cover = style.manifest.patterns.find((pattern) => pattern.id === 'cover');
  const title = cover?.slots.find((slot) => slot.name === 'title');
  const subtitle = cover?.slots.find((slot) => slot.name === 'subtitle');
  assert.ok(title?.kind === 'text' && title.size === 44 && (title.maxChars ?? 0) >= 30, JSON.stringify(title));
  assert.ok(subtitle?.kind === 'text' && subtitle.size === 32 && subtitle.empty === 'remove', JSON.stringify(subtitle));
  // A bulleted body is a list, with the room of its box and not of its sample.
  const body = style.manifest.patterns.find((pattern) => pattern.id === 'slide2')?.slots.find((slot) => slot.name === 'list1');
  assert.ok(body?.kind === 'text' && body.size === 32 && (body.maxChars ?? 0) >= 200, JSON.stringify(body));
  // No slide draws anything in a colour: the colours are the theme's first two accents.
  assert.deepEqual(style.manifest.colors, { accent: '4F81BD', accent2: 'C0504D' });
  assert.match(style.notes.join('\n'), /10 texts are placeholders that take their place or their size from the slide layout/);
  assert.match(style.notes.join('\n'), /the main colour is the first accent of the theme, 4F81BD/);
  assert.deepEqual(errors(checkManifest(loadTemplate(plainPptx, style.manifest))), []);
  assert.deepEqual(styleManifest(plainPptx, { name: 'plain' }), style);
});

test('a placeholder as a shape that states its own look reads as the placeholder does', () => {
  const styles = new TextStyles(PptxPackage.load(plainPptx));
  const stated = styles.standalone(2, '3');
  assert.ok(stated);
  if (!stated) return;
  const xml = String(stated).replace(/ xmlns:\w+="[^"]*"/g, '');
  // No placeholder any more: its place, its outline and its list style are its own.
  assert.doesNotMatch(xml, /<p:ph\b/);
  assert.match(xml, /<a:xfrm><a:off x="457200" y="1600200"\/><a:ext cx="8229600" cy="4525963"\/><\/a:xfrm><a:prstGeom prst="rect">/);
  assert.match(xml, /<a:lstStyle><a:lvl1pPr marL="342900" indent="-342900"[^>]*>.*?<a:buChar char="•"\/><a:defRPr sz="3200"[^>]*>/s);
  assert.match(xml, /<a:lvl2pPr marL="742950"/);
  // The same copy every time, and none for a shape that is no placeholder or is not there.
  assert.equal(styles.standalone(2, '3'), stated);
  assert.equal(styles.standalone(2, '99'), undefined);
  const demo = new TextStyles(PptxPackage.load(demoPptx));
  const box = styleManifest(demoPptx).manifest.patterns.find((pattern) => pattern.id === 'slide4')?.slots.find((slot) => slot.kind === 'text' && slot.name !== 'title');
  assert.equal(demo.standalone(4, box?.shapeId ?? ''), undefined);

  // Put in the placeholder's place, the copy is measured as the placeholder was: typeface, size, margins, bullet, spacing, insets.
  const zip = unzipSync(plainPptx);
  const doc = parseXml(strFromU8(zip['ppt/slides/slide2.xml'] as Uint8Array));
  const old = [...doc.getElementsByTagName('p:sp')].find((shape) => shape.getElementsByTagName('p:cNvPr')[0]?.getAttribute('id') === '3');
  assert.ok(old?.parentNode);
  old?.parentNode?.replaceChild(doc.importNode(stated, true), old);
  zip['ppt/slides/slide2.xml'] = strToU8(serializeXml(doc));
  const before = styles.frame(2, '3');
  const after = new TextStyles(PptxPackage.load(zipSync(zip))).frame(2, '3');
  assert.ok(before && after);
  // One thing differs, and is the point: the copy states its own place, where the placeholder took its layout's.
  assert.deepEqual([before?.placed, after?.placed], [false, true]);
  assert.deepEqual(JSON.parse(JSON.stringify({ ...after, placed: false })), JSON.parse(JSON.stringify(before)));
  assert.equal(before?.paragraphs[0]?.run.size, 32);
  assert.equal(before?.paragraphs[0]?.bullet, true);
});

test('a file whose text is all in placeholders carries the slides the software lays out', () => {
  const style = styleManifest(plainPptx);
  const template = loadTemplate(plainPptx, style.manifest);
  const kit = freeKitOrReason(template);
  assert.ok(typeof kit !== 'string', String(kit));
  if (typeof kit === 'string') return;
  // Running text is copied from the body placeholder. Its 32 pt under a 44 pt title is a size for
  // five bullets on a slide: it stands in at under half the title's size, without its bullet.
  assert.deepEqual([kit.styles.body.slide, kit.styles.body.shapeId, kit.styles.body.size], [2, '3', 20]);
  assert.equal(kit.styles.body.set?.bullets, false);
  assert.equal(kit.styles.list.bullet, true);
  assert.equal(kit.palette.accent, '4F81BD');

  const cards = [
    { card: { heading: 'Longer hours', text: 'Bikes unlock from five.' } },
    { card: { heading: 'Safer docks', text: 'Every station gets a light.' } },
    { card: { heading: 'Group passes', text: 'One pass for five riders.' } },
  ];
  const deck: DeckSpec = {
    version: 1,
    slides: [
      { pattern: 'cover', slots: { title: 'Riverbend Bike Share', subtitle: 'The plan for 2027' } },
      freeSlide({ title: 'Three ways in', blocks: cards }),
      freeSlide({ title: 'What riders asked for', blocks: [{ list: ['Longer hours on weekends', 'More docks near the tram', 'A pass for families'] }] }),
      { pattern: 'closing', slots: { title: 'Thank you' } },
    ],
  };
  assert.deepEqual(errors(validateDeck(template, deck, { layout: 'adaptive' })), []);
  const files = unzipSync(renderDeck(template, deck, { layout: 'adaptive' }).pptx);
  const slide = (number: number) => strFromU8(files[`ppt/slides/slide${number}.xml`] as Uint8Array);
  // The cover is the template's own slide, placeholders and all. A slide that is laid out keeps one
  // placeholder, the title of its frame: every text on it is a copy that states its own look.
  assert.equal((slide(1).match(/<p:ph\b/g) ?? []).length, 2);
  for (const number of [2, 3]) assert.equal((slide(number).match(/<p:ph\b/g) ?? []).length, 1, `slide ${number}`);
  assert.match(slide(2), /Longer hours/);
  assert.match(slide(2), /<a:lstStyle><a:lvl1pPr/);
  // Running text has no bullet, whatever the body placeholder it is copied from gives; a list has the template's.
  assert.match(slide(2), /<a:buNone\/>/);
  assert.doesNotMatch(slide(3), /<a:buNone\/>/);
  assert.match(slide(3), /<a:buChar char="•"\/>/);
  // The closing slide's line is left out, and its placeholder with it: no sample text stays.
  assert.doesNotMatch(slide(4), /Questions and discussion/);
});

test('a file that cannot carry such slides says why, and offers none', () => {
  // The fixture with its three slides of a title and a body taken out: a cover and a closing slide, and no frame.
  const zip = unzipSync(plainPptx);
  const presentation = strFromU8(zip['ppt/presentation.xml'] as Uint8Array);
  const ids = [...presentation.matchAll(/<p:sldId [^>]*\/>/g)].map((match) => match[0]);
  assert.equal(ids.length, 5);
  zip['ppt/presentation.xml'] = strToU8(ids.slice(1, 4).reduce((xml, id) => xml.replace(id, ''), presentation));
  const style = styleManifest(zipSync(zip));
  assert.match(style.reason ?? '', /^no slide has a title in its upper part with room under it/);
  assert.deepEqual(style.patterns, ['cover', 'closing']);
  // The manifest itself offers none: left to itself the engine would stand them on the cover.
  assert.equal(style.manifest.free, false);
  assert.match(style.notes.join('\n'), /the template cannot carry slides the software lays out: no slide has a title/);
  assert.deepEqual(errors(checkManifest(loadTemplate(zipSync(zip), style.manifest))), []);
});

// 0.10.1: what decks on the themes of an office suite showed. Each case is the fixture with one part changed.
const plainWith = (edits: Record<string, (xml: string) => string>): Uint8Array => {
  const zip = unzipSync(plainPptx);
  for (const [part, edit] of Object.entries(edits)) {
    const before = strFromU8(zip[part] as Uint8Array);
    const after = edit(before);
    assert.notEqual(after, before, `${part} was not changed`);
    zip[part] = strToU8(after);
  }
  return zipSync(zip);
};
const kitOf = (pptx: Uint8Array) => {
  const kit = freeKitOrReason(loadTemplate(pptx, styleManifest(pptx).manifest));
  assert.ok(typeof kit !== 'string', String(kit));
  return kit as Exclude<typeof kit, string>;
};
const LAYOUT = 'ppt/slideLayouts/slideLayout2.xml';
const MASTER = 'ppt/slideMasters/slideMaster1.xml';
// The body of the content layout drawn over the left seven tenths of the slide, as a theme with a decoration down its right side has it.
const narrowBody = (xml: string) => xml.replace('<p:ph idx="1"/></p:nvPr></p:nvSpPr><p:spPr/>', '<p:ph idx="1"/></p:nvPr></p:nvSpPr><p:spPr><a:xfrm><a:off x="457200" y="1600200"/><a:ext cx="5943600" cy="4525963"/></a:xfrm></p:spPr>');
// A decoration of the master down the right fifth of the slide, inside a group that spans the slide and draws its children at half size.
const SIDE =
  '<p:grpSp><p:nvGrpSpPr><p:cNvPr id="90" name="Decoration"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="9144000" cy="6858000"/><a:chOff x="0" y="0"/><a:chExt cx="18288000" cy="13716000"/></a:xfrm></p:grpSpPr>' +
  '<p:sp><p:nvSpPr><p:cNvPr id="91" name="Side"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="14630400" y="0"/><a:ext cx="3657600" cy="13716000"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:solidFill><a:srgbClr val="4F81BD"/></a:solidFill></p:spPr></p:sp></p:grpSp>';
const drawnOnMaster = (shape: string) => (xml: string) => xml.replace('</p:spTree>', `${shape}</p:spTree>`);

test('the content of a slide of placeholders stands where its layout puts it, and the room is not widened into what the master draws', () => {
  // What a master draws is read shape by shape: a group that spans the slide draws where its children are.
  const pkg = PptxPackage.load(plainWith({ [MASTER]: drawnOnMaster(SIDE) }));
  const boxes = drawnBoxes(rootElement(parseXml(pkg.readText(MASTER))), 9144000, 6858000);
  assert.deepEqual(boxes, [{ x: 0.8, y: 0, w: 0.2, h: 1 }]);

  // With nothing drawn beside it, a body that leaves the right of the slide free has its room widened to match the left margin.
  const open = kitOf(plainWith({ [LAYOUT]: narrowBody }));
  assert.ok(Math.abs(open.area.x - 0.05) < 0.005, String(open.area.x));
  assert.ok(open.area.x + open.area.w > 0.9, `to ${open.area.x + open.area.w}`);
  // With the master's decoration there, the room ends where the template's own body ends: nothing is set on the decoration.
  const beside = kitOf(plainWith({ [LAYOUT]: narrowBody, [MASTER]: drawnOnMaster(SIDE) }));
  assert.ok(Math.abs(beside.area.x - 0.05) < 0.005);
  assert.ok(Math.abs(beside.area.x + beside.area.w - 0.7) < 0.01, `to ${beside.area.x + beside.area.w}`);

  // The rule by itself: a strip is added at the side reading ends, unless something is drawn in it under the title.
  const area = { x: 0.05, y: 0.04, w: 0.65, h: 0.9 };
  const title = { x: 0.05, y: 0.04, w: 0.9, h: 0.17 };
  assert.ok(widened(area, 'ltr', [], title).w > 0.85);
  assert.deepEqual(widened(area, 'ltr', [], title, [{ x: 0.8, y: 0, w: 0.2, h: 1 }]), area);
  // What is drawn beside the title alone, a logo in the corner, takes nothing from the room under it.
  assert.ok(widened(area, 'ltr', [], title, [{ x: 0.85, y: 0.02, w: 0.1, h: 0.12 }]).w > 0.85);
  assert.deepEqual(widened({ ...area, x: 0.3 }, 'rtl', [], title, [{ x: 0, y: 0, w: 0.2, h: 1 }]), { ...area, x: 0.3 });
});

test('the band starts where the template\'s own body starts, and the larger room does not reach over a banner behind the title', () => {
  // The band starts where the master puts the body: 0.233 of the way down.
  const plain = kitOf(plainPptx);
  assert.ok(Math.abs(plain.band.top - 1600200 / 6858000) < 0.002, String(plain.band.top));
  assert.ok(plain.dense.top <= plain.band.top);
  // A layout that puts its body lower, three tenths of the way down: the band starts there, and a
  // crowded slide may start just under the title, in the larger room.
  const lowBody = (xml: string) => xml.replace('<p:ph idx="1"/></p:nvPr></p:nvSpPr><p:spPr/>', '<p:ph idx="1"/></p:nvPr></p:nvSpPr><p:spPr><a:xfrm><a:off x="457200" y="2057400"/><a:ext cx="8229600" cy="4069080"/></a:xfrm></p:spPr>');
  const low = kitOf(plainWith({ [LAYOUT]: lowBody }));
  assert.ok(Math.abs(low.band.top - 0.3) < 0.01, String(low.band.top));
  assert.ok(low.dense.top < low.band.top - 0.05, `${low.dense.top} and ${low.band.top}`);
  // With a banner of the master from the top of the slide to between the title and that body, the
  // larger room starts with the band: text set in the gap stood on the banner's edge.
  const banner = '<p:sp><p:nvSpPr><p:cNvPr id="92" name="Banner"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="9144000" cy="1783080"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:solidFill><a:srgbClr val="4F81BD"/></a:solidFill></p:spPr></p:sp>';
  const under = kitOf(plainWith({ [LAYOUT]: lowBody, [MASTER]: drawnOnMaster(banner) }));
  assert.equal(under.dense.top, under.band.top);
  assert.equal(under.band.top, low.band.top);
});

test('running text is the body of the content slides, not the line under the cover\'s title, and not a boxed caption', () => {
  // The subtitle of the cover at 18 pt: plain, small and with no bullet, it was taken for running text before the body was.
  const small = kitOf(plainWith({ 'ppt/slides/slide1.xml': (xml) => xml.replace('<a:r><a:t>The plan for 2027</a:t></a:r>', '<a:r><a:rPr sz="1800"/><a:t>The plan for 2027</a:t></a:r>') }));
  assert.deepEqual([small.styles.body.slide, small.styles.body.shapeId], [2, '3']);
  // A body placeholder its layout draws a line round is a box with a text in it: its look is no text style.
  const boxed = (xml: string) => xml.replace('<p:ph idx="1"/></p:nvPr></p:nvSpPr><p:spPr/>', '<p:ph idx="1"/></p:nvPr></p:nvSpPr><p:spPr><a:ln w="38100"><a:solidFill><a:srgbClr val="C0504D"/></a:solidFill></a:ln></p:spPr>');
  const outlined = freeKitOrReason(loadTemplate(plainWith({ [LAYOUT]: boxed }), styleManifest(plainWith({ [LAYOUT]: boxed })).manifest));
  if (typeof outlined !== 'string') assert.notDeepEqual([outlined.styles.body.slide, outlined.styles.body.shapeId], [2, '3']);
});

test('a cover and a closing slide are title slides: a slide of content that comes last is no closing slide', () => {
  // The fixture without its last slide: the deck ends on a title and a bulleted body.
  const zip = unzipSync(plainPptx);
  const presentation = strFromU8(zip['ppt/presentation.xml'] as Uint8Array);
  const ids = [...presentation.matchAll(/<p:sldId [^>]*\/>/g)].map((match) => match[0]);
  zip['ppt/presentation.xml'] = strToU8(presentation.replace(ids[4] as string, ''));
  const style = styleManifest(zipSync(zip));
  assert.equal(style.closing, undefined);
  assert.deepEqual(style.patterns, ['cover', 'free']);
  assert.match(style.notes.join('\n'), /the last slide is no title slide/);
  // A subtitle placeholder is the line under the title whatever its sample holds: a presenter's name over a date.
  const twoLines = styleManifest(plainWith({ 'ppt/slides/slide1.xml': (xml) => xml.replace('<a:p><a:r><a:t>The plan for 2027</a:t></a:r></a:p>', '<a:p><a:r><a:t>Dana Levi</a:t></a:r></a:p><a:p><a:r><a:t>November 2026</a:t></a:r></a:p>') }));
  assert.equal(twoLines.cover, 'cover');
  assert.ok(twoLines.manifest.patterns[0]?.slots.some((slot) => slot.name === 'subtitle'));
});
