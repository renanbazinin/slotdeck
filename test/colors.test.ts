// A deck in other colours than its template's: what follows the new colour, and what stays.
import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { test } from 'node:test';
import { fileURLToPath } from 'node:url';
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate';
import { paletteOf, recolorTree, schemeOf } from '../src/color/recolor.ts';
import { contrast, hexToOklch, hueDistance } from '../src/color/space.ts';
import { extractAssets, loadTemplate, renderDeck, validateDeck } from '../src/index.ts';
import type { DeckSpec, TemplateManifest } from '../src/index.ts';
import { PptxPackage } from '../src/pptx/package.ts';
import { parseXml, rootElement, serializeXml } from '../src/pptx/xml.ts';
import { readPngPixels } from '../src/util/png.ts';
import { placeholderPhoto } from '../examples/placeholder.ts';

const root = join(dirname(fileURLToPath(import.meta.url)), '..');
const manifest = JSON.parse(readFileSync(join(root, 'templates', 'demo.manifest.json'), 'utf8')) as TemplateManifest;
const demoPptx = new Uint8Array(readFileSync(join(root, 'templates', 'demo.pptx')));
const template = loadTemplate(demoPptx, manifest);
const icon = (id: string) => new Uint8Array(readFileSync(join(root, 'examples', 'icons', `${id}.png`)));
const photo = placeholderPhoto(400, 300, [14, 124, 123], [244, 162, 89]);
const assets = { clock: icon('clock'), shield: icon('shield'), photo };

const TEAL = '0E7C7B';
const TINT = 'E8F1F0';
const INK = '1B2A41';
const AMBER = 'F4A259';
const BLUE = '1565C0';

const deck: DeckSpec = {
  version: 1,
  fields: { organizationDepartment: 'Riverbend' },
  slides: [
    { pattern: 'cover', slots: { title: 'A year of rides', subtitle: 'What the season taught us' } },
    { pattern: 'icon-cards', items: 2, slots: { title: 'Two changes', card1Icon: { asset: 'clock' }, card1Heading: 'Longer hours', card1Text: 'Bikes unlock from five.', card2Icon: { asset: 'shield' }, card2Heading: 'Safer docks', card2Text: 'Every station gets a light.' } },
    { pattern: 'chart-native', slots: { title: 'Rides grew', chart: { type: 'column', categories: ['Spring', 'Summer'], series: [{ name: '2024', values: [5, 9] }, { name: '2025', values: [7, 19] }] }, chartDescription: 'Rides in thousands.' } },
    { pattern: 'photo-stat', slots: { title: 'Most rides replace a car trip', photo: { asset: 'photo' }, statValue: '62%', statDescription: 'of riders said so.' } },
  ],
};

const files = (pptx: Uint8Array) => unzipSync(pptx);
const text = (pptx: Uint8Array, name: string) => strFromU8(files(pptx)[name] as Uint8Array);
const colours = (xml: string) => new Set([...xml.matchAll(/<a:srgbClr val="([0-9A-F]{6})"/g)].map((match) => match[1] as string));
const render = (given: DeckSpec) => renderDeck(template, given, { assets, layout: 'adaptive' }).pptx;

test('the palette of the demo template is read from its slides', () => {
  const palette = paletteOf(template.pkg, manifest);
  assert.equal(palette.accent, TEAL);
  assert.equal(palette.accent2, AMBER);
  assert.equal(palette.ink, INK);
  assert.deepEqual(palette.families.find((family) => family.base === TEAL)?.members, [TINT, TEAL]);
  // The blue of the stock theme, which shapes only name as a fallback behind their own fill, is not counted.
  assert.ok(!palette.colors.some((entry) => entry.hex === '4F81BD'));
  // Read once for a template; a manifest may state the main colour itself.
  assert.equal(paletteOf(template.pkg, manifest), palette);
  assert.equal(paletteOf(template.pkg, { ...manifest, colors: { accent: AMBER } }).accent, AMBER);
});

test('a deck that asks for no colour is the file it always was', () => {
  const plain = render(deck);
  assert.deepEqual(render({ ...deck, colors: {} }), plain);
  assert.deepEqual(render({ ...deck, colors: { accent: TEAL } }), plain);
  // A colour that is no colour is left out, with a warning that names no slide.
  const bad = { ...deck, colors: { accent: 'blue' } };
  assert.deepEqual(render(bad), plain);
  const issues = validateDeck(manifest, bad, { layout: 'adaptive', assets: Object.keys(assets) });
  assert.equal(issues.length, 1);
  assert.equal(issues[0]?.severity, 'warning');
  assert.equal(issues[0]?.slide, undefined);
  assert.match(issues[0]?.message ?? '', /"accent" is "blue", which is not six hex digits/);
});

test('the main colour of every slide follows the deck, its tints with it; the rest stays', () => {
  const plain = render(deck);
  const blue = render({ ...deck, colors: { accent: BLUE } });
  for (const number of [1, 2, 3, 4]) {
    const before = colours(text(plain, `ppt/slides/slide${number}.xml`));
    const after = colours(text(blue, `ppt/slides/slide${number}.xml`));
    assert.ok(!after.has(TEAL) && !after.has(TINT), `slide ${number} keeps no teal`);
    assert.equal(after.has(BLUE), before.has(TEAL), `slide ${number}: teal became the blue asked for`);
    // The ink and the second colour are where they were.
    assert.equal(after.has(INK), before.has(INK));
    assert.equal(after.has(AMBER), before.has(AMBER));
    assert.equal(after.size, before.size, `slide ${number} has as many colours as before`);
  }
  // The pale tint of the cards is now a pale tint of the blue.
  const tint = [...colours(text(blue, 'ppt/slides/slide2.xml'))].map((hex) => ({ hex, ...hexToOklch(hex) })).find((entry) => entry.l > 0.9 && entry.c > 0.005);
  assert.ok(tint && hueDistance(tint.h, hexToOklch(BLUE).h) < 15, JSON.stringify(tint));
});

test('the chart, the icons and nothing else of the pictures change', () => {
  const plain = files(render(deck));
  const blue = files(render({ ...deck, colors: { accent: BLUE } }));
  const chart = strFromU8(blue['ppt/charts/chart1.xml'] as Uint8Array);
  assert.ok(chart.includes(`<a:srgbClr val="${BLUE}"/>`) && !chart.includes(TEAL));
  assert.ok(strFromU8(plain['ppt/charts/chart1.xml'] as Uint8Array).includes(TEAL));

  const media = (zip: Record<string, Uint8Array>) => Object.keys(zip).filter((name) => /media\/slotdeck-/.test(name)).sort();
  assert.deepEqual(media(blue), media(plain));
  const changed = media(blue).filter((name) => Buffer.compare(blue[name] as Uint8Array, plain[name] as Uint8Array) !== 0);
  assert.equal(changed.length, 2, 'the two icons, and not the photo');
  for (const name of changed) {
    const { pixels } = readPngPixels(blue[name] as Uint8Array) as { pixels: Uint8Array };
    let teal = 0;
    let blueish = 0;
    for (let i = 0; i < pixels.length; i += 4) {
      if ((pixels[i + 3] as number) < 200) continue;
      if (pixels[i] === 0x0e && pixels[i + 1] === 0x7c && pixels[i + 2] === 0x7b) teal++;
      if (pixels[i] === 0x15 && pixels[i + 1] === 0x65 && pixels[i + 2] === 0xc0) blueish++;
    }
    assert.equal(teal, 0);
    assert.ok(blueish > 50, `${name} is drawn in the new colour`);
  }
  // The template's own pictures, the logo among them, are the same bytes.
  for (const name of Object.keys(plain).filter((entry) => /media\/image/.test(entry))) assert.deepEqual(blue[name], plain[name], name);
});

test('one slide in its own colours leaves the others alone', () => {
  const plain = render(deck);
  const one = structuredClone(deck);
  (one.slides[3] as DeckSpec['slides'][number]).colors = { accent: 'C8102E' };
  const mixed = render(one);
  for (const number of [1, 2, 3]) assert.equal(text(mixed, `ppt/slides/slide${number}.xml`), text(plain, `ppt/slides/slide${number}.xml`));
  assert.ok(colours(text(mixed, 'ppt/slides/slide4.xml')).has('C8102E'));
  // The theme and the masters belong to every slide, so one slide's colours do not reach them.
  for (const name of Object.keys(files(plain)).filter((entry) => /theme|slideMaster|slideLayout/.test(entry))) assert.deepEqual(files(mixed)[name], files(plain)[name], name);
  // A slide's colours take the place of the deck's on that slide.
  const both = render({ ...one, colors: { accent: BLUE } });
  assert.ok(colours(text(both, 'ppt/slides/slide4.xml')).has('C8102E') && !colours(text(both, 'ppt/slides/slide4.xml')).has(BLUE));
  assert.ok(colours(text(both, 'ppt/slides/slide1.xml')).has(BLUE));
  // The same icon in two colours is two pictures; in one colour it is stored once.
  const twice = structuredClone(deck);
  twice.slides.push({ ...structuredClone(deck.slides[1] as DeckSpec['slides'][number]), colors: { accent: 'C8102E' } });
  const media = (pptx: Uint8Array) => Object.keys(files(pptx)).filter((name) => /media\/slotdeck-/.test(name)).length;
  assert.equal(media(render({ ...twice, colors: { accent: BLUE } })), media(plain) + 2);
});

test('a colour named through the theme changes too, and what tints it stays on it', () => {
  // The demo template with its main colour moved into the theme: the cover's circle names it as
  // accent1, darkened a little, and the theme says what accent1 is.
  const zip = unzipSync(demoPptx);
  const edit = (name: string, change: (xml: string) => string) => {
    zip[name] = strToU8(change(strFromU8(zip[name] as Uint8Array)));
  };
  edit('ppt/theme/theme1.xml', (xml) => xml.replace(/<a:accent1>.*?<\/a:accent1>/s, `<a:accent1><a:srgbClr val="${TEAL}"/></a:accent1>`));
  let swapped = 0;
  edit('ppt/slides/slide1.xml', (xml) =>
    xml.replace(new RegExp(`<a:solidFill><a:srgbClr val="${TEAL}"/></a:solidFill>`, 'g'), () => {
      swapped++;
      return '<a:solidFill><a:schemeClr val="accent1"><a:lumMod val="90000"/></a:schemeClr></a:solidFill>';
    }),
  );
  assert.ok(swapped > 0, 'the cover has a teal fill to move into the theme');
  const themed = loadTemplate(zipSync(zip), manifest);
  assert.equal(schemeOf(themed.pkg, 'ppt/slides/slide1.xml').resolve('accent1'), TEAL);
  assert.equal(schemeOf(themed.pkg, 'ppt/slides/slide1.xml').resolve('tx1'), '000000');
  assert.equal(paletteOf(themed.pkg, manifest).accent, TEAL);

  const blue = renderDeck(themed, { ...deck, colors: { accent: BLUE } }, { assets, layout: 'adaptive' }).pptx;
  // The theme of the deck says the new colour, so the slide may go on naming it; the template's own package still says the old.
  assert.ok(text(blue, 'ppt/theme/theme1.xml').includes(`<a:accent1><a:srgbClr val="${BLUE}"/></a:accent1>`));
  assert.ok(text(blue, 'ppt/slides/slide1.xml').includes('<a:schemeClr val="accent1"><a:lumMod val="90000"/></a:schemeClr>'));
  assert.ok(themed.pkg.readText('ppt/theme/theme1.xml').includes(`<a:accent1><a:srgbClr val="${TEAL}"/></a:accent1>`));
  // One slide's colours change the slide and leave the theme: there the named colour becomes a literal, with its shade kept.
  const one = structuredClone(deck);
  (one.slides[0] as DeckSpec['slides'][number]).colors = { accent: BLUE };
  const single = renderDeck(themed, one, { assets, layout: 'adaptive' }).pptx;
  assert.ok(text(single, 'ppt/slides/slide1.xml').includes(`<a:srgbClr val="${BLUE}"><a:lumMod val="90000"/></a:srgbClr>`));
  assert.ok(text(single, 'ppt/theme/theme1.xml').includes(`<a:accent1><a:srgbClr val="${TEAL}"/></a:accent1>`));
  // A slide in its own colours, in a deck in others, shows its own: the deck's theme would give it the deck's.
  const red = renderDeck(themed, { ...one, colors: { accent: 'C8102E' } }, { assets, layout: 'adaptive' }).pptx;
  assert.ok(text(red, 'ppt/theme/theme1.xml').includes('<a:accent1><a:srgbClr val="C8102E"/></a:accent1>'));
  assert.ok(text(red, 'ppt/slides/slide1.xml').includes(`<a:srgbClr val="${BLUE}"><a:lumMod val="90000"/></a:srgbClr>`));
  // And one that asks for the template's own colours there keeps them.
  (one.slides[0] as DeckSpec['slides'][number]).colors = {};
  const kept = renderDeck(themed, { ...one, colors: { accent: 'C8102E' } }, { assets, layout: 'adaptive' }).pptx;
  assert.ok(text(kept, 'ppt/slides/slide1.xml').includes(`<a:srgbClr val="${TEAL}"><a:lumMod val="90000"/></a:srgbClr>`));
  assert.ok(!colours(text(kept, 'ppt/slides/slide1.xml')).has('C8102E'));
});

test('every colour is changed once: a table, a layout that two slides share, an icon read back', () => {
  const tables: DeckSpec = {
    version: 1,
    fields: { organizationDepartment: 'Riverbend' },
    colors: { accent: BLUE },
    slides: [1, 2].map(() => ({ pattern: 'table', slots: { title: 'Stations', table: { header: ['Station', 'Docks'], rows: [['Tram Square', '24'], ['River Gate', '18']] } } })),
  };
  // The demo manifest states the colour of the table's text; a second mapping would move it again.
  const tableSlot = manifest.patterns.find((pattern) => pattern.id === 'table')?.slots.find((slot) => slot.kind === 'table');
  const stated = [tableSlot?.kind === 'table' ? tableSlot.headerText?.color : undefined, tableSlot?.kind === 'table' ? tableSlot.bodyText?.color : undefined].filter((color): color is string => !!color);
  const pptx = renderDeck(template, tables, { layout: 'adaptive' }).pptx;
  const palette = paletteOf(template.pkg, manifest);
  const once = (hex: string) => (hex === TEAL ? BLUE : hex);
  const slide = colours(text(pptx, 'ppt/slides/slide1.xml'));
  for (const color of stated) assert.ok(slide.has(once(color.toUpperCase())), `the table text stated as ${color} is mapped once`);
  assert.equal(text(pptx, 'ppt/slides/slide1.xml'), text(pptx, 'ppt/slides/slide2.xml').replace(/<a:t>2<\/a:t>/, '<a:t>1</a:t>'));
  assert.equal(palette.accent, TEAL);

  // A layout with a teal background, shared by both slides, is teal once more in the template and blue once in the deck.
  const zip = unzipSync(demoPptx);
  const layoutOf = (strFromU8(zip['ppt/slides/_rels/slide8.xml.rels'] as Uint8Array).match(/slideLayouts\/(slideLayout\d+\.xml)/) as RegExpMatchArray)[1] as string;
  const layoutPart = `ppt/slideLayouts/${layoutOf}`;
  zip[layoutPart] = strToU8(strFromU8(zip[layoutPart] as Uint8Array).replace('<p:spTree>', `<p:bg><p:bgPr><a:solidFill><a:srgbClr val="${TEAL}"/></a:solidFill><a:effectLst/></p:bgPr></p:bg><p:spTree>`));
  const withBackground = loadTemplate(zipSync(zip), manifest);
  const coloured = renderDeck(withBackground, tables, { layout: 'adaptive' }).pptx;
  assert.ok(text(coloured, layoutPart).includes(`<p:bg><p:bgPr><a:solidFill><a:srgbClr val="${BLUE}"/>`));
  assert.ok(withBackground.pkg.readText(layoutPart).includes(`<a:srgbClr val="${TEAL}"/>`));
  // One slide alone in other colours takes the background along as its own, and the layout stays.
  const alone = structuredClone(tables);
  delete alone.colors;
  (alone.slides[1] as DeckSpec['slides'][number]).colors = { accent: 'C8102E' };
  const mixed = renderDeck(withBackground, alone, { layout: 'adaptive' }).pptx;
  assert.ok(text(mixed, 'ppt/slides/slide2.xml').includes('<p:bg><p:bgPr><a:solidFill><a:srgbClr val="C8102E"/>'));
  assert.ok(!text(mixed, 'ppt/slides/slide1.xml').includes('<p:bg>'));
  assert.ok(text(mixed, layoutPart).includes(`<a:srgbClr val="${TEAL}"/>`));

  // A recoloured icon read back from the file is not handed on as the original.
  const blue = renderDeck(template, { ...deck, colors: { accent: BLUE } }, { assets, layout: 'adaptive' }).pptx;
  assert.deepEqual(Object.keys(extractAssets(template, { ...deck, colors: { accent: BLUE } }, blue)).sort(), ['photo']);
  assert.deepEqual(Object.keys(extractAssets(template, deck, render(deck))).sort(), ['clock', 'photo', 'shield']);
});

test('text stays readable when a colour changes under it or on it', () => {
  // The cover's subtitle is pale on the dark ink; the big figure of the photo slide is teal on white.
  const yellow = renderDeck(template, { ...deck, colors: { accent: 'FFD400' } }, { assets, layout: 'adaptive' }).pptx;
  const figure = text(yellow, 'ppt/slides/slide4.xml');
  const number = /<a:solidFill><a:srgbClr val="([0-9A-F]{6})"\/><\/a:solidFill>[^<]*(?:<a:latin[^>]*>)?(?:<[^>]*>)*?<a:t>62%<\/a:t>/.exec(figure)?.[1] ?? [...figure.matchAll(/<a:r><a:rPr[^>]*>.*?<a:srgbClr val="([0-9A-F]{6})"\/>.*?<a:t>62%<\/a:t>/gs)].map((match) => match[1])[0];
  assert.ok(number, 'the figure states its colour');
  assert.notEqual(number, 'FFD400', 'a yellow figure on white would not read');
  assert.ok(contrast(number as string, 'FFFFFF') >= 4.5, `${number} reads on white`);
  assert.ok(hueDistance(hexToOklch(number as string).h, hexToOklch('FFD400').h) < 25, 'and is still a yellow');
  // A fill in the new colour is exactly the colour asked for.
  assert.ok(colours(text(yellow, 'ppt/slides/slide1.xml')).has('FFD400'));
  // A dark blue needs no help: the figure is the colour asked for.
  const navy = renderDeck(template, { ...deck, colors: { accent: '0B3D91' } }, { assets, layout: 'adaptive' }).pptx;
  assert.ok(/<a:srgbClr val="0B3D91"\/>(?:(?!<a:t>).)*<a:t>62%<\/a:t>/s.test(text(navy, 'ppt/slides/slide4.xml')));
});

test('the colours of a picture and of a placeholder are not touched', () => {
  const xml =
    '<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:cSld><p:spTree>' +
    '<p:pic><p:blipFill><a:blip><a:duotone><a:srgbClr val="0E7C7B"/><a:prstClr val="white"/></a:duotone></a:blip></p:blipFill><p:spPr><a:ln><a:solidFill><a:srgbClr val="0E7C7B"/></a:solidFill></a:ln></p:spPr></p:pic>' +
    '<p:sp><p:spPr><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></p:spPr><p:txBody><a:p><a:r><a:rPr><a:solidFill><a:prstClr val="teal"><a:alpha val="50000"/></a:prstClr></a:solidFill></a:rPr><a:t>x</a:t></a:r></a:p></p:txBody></p:sp>' +
    '</p:spTree></p:cSld></p:sld>';
  const doc = parseXml(xml);
  const map = (hex: string) => (hex === '0E7C7B' || hex === '008080' ? BLUE : hex);
  assert.equal(recolorTree(rootElement(doc), map, { resolve: () => undefined }), 2);
  const out = serializeXml(doc);
  assert.ok(out.includes('<a:duotone><a:srgbClr val="0E7C7B"/><a:prstClr val="white"/></a:duotone>'), 'the picture keeps its effect');
  assert.ok(out.includes(`<a:ln><a:solidFill><a:srgbClr val="${BLUE}"/></a:solidFill></a:ln>`), 'its outline follows');
  assert.ok(out.includes('<a:schemeClr val="phClr"/>'));
  assert.ok(out.includes(`<a:srgbClr val="${BLUE}"><a:alpha val="50000"/></a:srgbClr>`), 'a named colour becomes a literal and keeps what is on it');
  // A package without a theme, or a part that is not there, has nothing to look up and does not fail.
  assert.equal(schemeOf(PptxPackage.load(demoPptx), 'ppt/slides/none.xml').resolve('accent1'), undefined);
});

/** The demo template's package with one part changed. */
const demoWith = (name: string, change: (xml: string) => string) => {
  const zip = unzipSync(demoPptx);
  zip[name] = strToU8(change(strFromU8(zip[name] as Uint8Array)));
  return zip;
};
/** A slide's XML with every colour that shape #5, the title, states set to black. */
const blackTitle = (xml: string): string => {
  const at = xml.indexOf('<p:cNvPr id="5"');
  const end = xml.indexOf('</p:sp>', at);
  return xml.slice(0, at) + xml.slice(at, end).replace(/<a:srgbClr val="[0-9A-F]{6}"\/>/g, '<a:srgbClr val="000000"/>') + xml.slice(end);
};
const GRADIENT = `<p:bg><p:bgPr><a:gradFill><a:gsLst><a:gs pos="0"><a:srgbClr val="${TEAL}"/></a:gs><a:gs pos="23000"><a:srgbClr val="${TINT}"/></a:gs><a:gs pos="100000"><a:srgbClr val="FFFFFF"/></a:gs></a:gsLst><a:lin ang="0"/></a:gradFill><a:effectLst/></p:bgPr></p:bg>`;

test('a text on a gradient is judged against the whole gradient, not its first stop', () => {
  // The slide of the text layouts on a gradient that is teal at one edge and white where the
  // title stands, with the title in black: as a template with pale gradient slides has it.
  const zip = demoWith('ppt/slides/slide2.xml', (xml) => blackTitle(xml.replace('<p:spTree>', `${GRADIENT}<p:spTree>`)));
  const gradient = loadTemplate(zipSync(zip), manifest);
  const one: DeckSpec = { version: 1, fields: { organizationDepartment: 'Riverbend' }, slides: [{ pattern: 'text-only', slots: { title: 'What the season taught us', body: 'Rides grew in every month.' } }] };
  const titleColour = (pptx: Uint8Array): string | undefined => {
    const slide = text(pptx, 'ppt/slides/slide1.xml');
    return /<a:solidFill><a:srgbClr val="([0-9A-F]{6})"/.exec(slide.slice(slide.indexOf('<p:cNvPr id="5"')))?.[1];
  };
  assert.equal(titleColour(renderDeck(gradient, one, { layout: 'adaptive' }).pptx), '000000');
  // A dark colour makes the first stop dark. The title stands on the light part and stays black:
  // judged against the first stop alone it was turned into a pale grey, which nobody reads there.
  const navy = renderDeck(gradient, { ...one, colors: { accent: '0B3D91' } }, { layout: 'adaptive' }).pptx;
  assert.ok(text(navy, 'ppt/slides/slide1.xml').includes('<a:gs pos="0"><a:srgbClr val="0B3D91"/>'), 'the gradient follows the colour');
  assert.equal(titleColour(navy), '000000');

  // The same for a slide that takes the gradient from its layout.
  const inherited = demoWith('ppt/slides/slide2.xml', blackTitle);
  const layoutPart = `ppt/slideLayouts/${(strFromU8(inherited['ppt/slides/_rels/slide2.xml.rels'] as Uint8Array).match(/slideLayouts\/(slideLayout\d+\.xml)/) as RegExpMatchArray)[1]}`;
  inherited[layoutPart] = strToU8(strFromU8(inherited[layoutPart] as Uint8Array).replace('<p:spTree>', `${GRADIENT}<p:spTree>`));
  const fromLayout = renderDeck(loadTemplate(zipSync(inherited), manifest), { ...one, colors: { accent: '0B3D91' } }, { layout: 'adaptive' }).pptx;
  assert.ok(text(fromLayout, layoutPart).includes('<a:gs pos="0"><a:srgbClr val="0B3D91"/>'));
  assert.equal(titleColour(fromLayout), '000000');

  // A gradient that is dark all over is another matter: black on it never read, and is left alone as before.
  // A solid fill is judged as it always was: the teal figure on white is made to read in yellow.
  const yellow = renderDeck(template, { ...deck, colors: { accent: 'FFD400' } }, { assets, layout: 'adaptive' }).pptx;
  assert.ok(!/<a:srgbClr val="FFD400"\/>(?:(?!<a:t>).)*<a:t>62%<\/a:t>/s.test(text(yellow, 'ppt/slides/slide4.xml')));
});

test('a slide number is text too: a field keeps reading when its colour changes', () => {
  // The slide number of the photo slide set in the main colour.
  const zip = demoWith('ppt/slides/slide3.xml', (xml) => {
    const field = /<a:fld\b[^>]*type="slidenum"[^>]*>.*?<\/a:fld>/s.exec(xml);
    assert.ok(field, 'the slide has a slide number');
    const fill = `<a:solidFill><a:srgbClr val="${TEAL}"/></a:solidFill>`;
    let coloured = field[0];
    if (/<a:rPr[^>]*>.*?<a:solidFill>.*?<\/a:solidFill>/s.test(coloured)) coloured = coloured.replace(/<a:solidFill>.*?<\/a:solidFill>/s, fill);
    else if (/<a:rPr[^>]*\/>/.test(coloured)) coloured = coloured.replace(/<a:rPr([^>]*)\/>/, `<a:rPr$1>${fill}</a:rPr>`);
    else coloured = coloured.replace(/(<a:rPr[^>]*>)/, `$1${fill}`);
    assert.ok(coloured.includes(fill), 'the field states the main colour');
    return xml.replace(field[0], coloured);
  });
  const numbered = loadTemplate(zipSync(zip), manifest);
  const one: DeckSpec = { version: 1, fields: { organizationDepartment: 'Riverbend' }, slides: [deck.slides[3] as DeckSpec['slides'][number]] };
  const fieldOf = (given: DeckSpec): string => /<a:fld\b[^>]*type="slidenum"[^>]*>(.*?)<\/a:fld>/s.exec(text(renderDeck(numbered, given, { assets, layout: 'adaptive' }).pptx, 'ppt/slides/slide1.xml'))?.[1] ?? '';
  const colourOf = (field: string) => /<a:srgbClr val="([0-9A-F]{6})"/.exec(field)?.[1];
  assert.equal(colourOf(fieldOf(one)), TEAL);
  const field = fieldOf({ ...one, colors: { accent: 'FFD400' } });
  const colour = colourOf(field);
  assert.ok(colour, 'the field states a colour');
  assert.notEqual(colour, 'FFD400', 'a yellow slide number on white would not read');
  assert.ok(contrast(colour as string, 'FFFFFF') >= 3, `${colour} reads on white`);
  // Its run properties stay the first thing in it, as the file format wants them.
  assert.ok(/^<a:rPr\b/.test(field));
  // A colour that reads is the colour asked for.
  assert.equal(colourOf(fieldOf({ ...one, colors: { accent: BLUE } })), BLUE);
});

test('every layout of the master follows the deck, also one that no slide of the deck uses', () => {
  // The layout of the closing slide gets a teal background; the deck has no closing slide.
  const zip = unzipSync(demoPptx);
  const layoutOf = (slide: number) => `ppt/slideLayouts/${(strFromU8(zip[`ppt/slides/_rels/slide${slide}.xml.rels`] as Uint8Array).match(/slideLayouts\/(slideLayout\d+\.xml)/) as RegExpMatchArray)[1]}`;
  // The deck's slides stand on template slides 1, 5, 9 and 3.
  const used = [1, 5, 9, 3].map(layoutOf);
  const unused = Object.keys(zip).filter((name) => /^ppt\/slideLayouts\/slideLayout\d+\.xml$/.test(name)).sort().find((part) => !used.includes(part) && strFromU8(zip[part] as Uint8Array).includes('<p:spTree>') && !strFromU8(zip[part] as Uint8Array).includes('<p:bg>'));
  assert.ok(unused, 'the template has a layout the deck does not use');
  zip[unused] = strToU8(strFromU8(zip[unused] as Uint8Array).replace('<p:spTree>', `<p:bg><p:bgPr><a:solidFill><a:srgbClr val="${TEAL}"/></a:solidFill><a:effectLst/></p:bgPr></p:bg><p:spTree>`));
  const withBackground = loadTemplate(zipSync(zip), manifest);
  const blue = renderDeck(withBackground, { ...deck, colors: { accent: BLUE } }, { assets, layout: 'adaptive' }).pptx;
  // A person who adds a slide in PowerPoint and picks that layout gets the deck's colour.
  assert.ok(text(blue, unused).includes(`<p:bg><p:bgPr><a:solidFill><a:srgbClr val="${BLUE}"/>`));
  assert.ok(withBackground.pkg.readText(unused).includes(`<a:srgbClr val="${TEAL}"/>`), 'the template keeps its own');
  // A deck without colours, and one slide in its own, leave every layout as it is.
  const plain = renderDeck(withBackground, deck, { assets, layout: 'adaptive' }).pptx;
  const one = structuredClone(deck);
  (one.slides[3] as DeckSpec['slides'][number]).colors = { accent: 'C8102E' };
  const mixed = renderDeck(withBackground, one, { assets, layout: 'adaptive' }).pptx;
  assert.deepEqual(files(mixed)[unused], files(plain)[unused]);
  assert.ok(text(plain, unused).includes(`<a:srgbClr val="${TEAL}"/>`));
});

test('a colour close to the page, and a second colour the template has not, are said and not refused', () => {
  const warnings = (given: DeckSpec) => renderDeck(template, given, { assets, layout: 'adaptive' }).warnings.map((issue) => issue.message);
  const near = (hex: string) => `the colour ${hex} is close to the page's own; fills, chart bars and icons in it will hardly show`;
  assert.deepEqual(warnings({ ...deck, colors: { accent: 'FFFFFF' } }), [near('FFFFFF')]);
  assert.deepEqual(warnings({ ...deck, colors: { accent: 'F7F5EE' } }), [near('F7F5EE')]);
  for (const accent of ['FFD400', BLUE, '0B0C10']) assert.deepEqual(warnings({ ...deck, colors: { accent } }), []);
  // Told once for the deck, without a slide number, so that no slide is written again for it.
  assert.deepEqual(renderDeck(template, { ...deck, colors: { accent: 'FFFFFF' } }, { assets, layout: 'adaptive' }).warnings.map((issue) => issue.slide), [undefined]);
  // For one slide in such a colour too.
  const one = structuredClone(deck);
  (one.slides[2] as DeckSpec['slides'][number]).colors = { accent: 'FFFFFF' };
  assert.deepEqual(warnings(one), [near('FFFFFF')]);

  // A template whose layouts use one colour only has no second colour to change.
  const single = { ...manifest, patterns: manifest.patterns.filter((pattern) => ['text-only', 'stat-cards', 'icon-cards', 'two-columns'].includes(pattern.id)) };
  const few = loadTemplate(demoPptx, single);
  assert.equal(paletteOf(few.pkg, single).accent2, undefined);
  const texts: DeckSpec = { version: 1, fields: { organizationDepartment: 'Riverbend' }, slides: [{ pattern: 'text-only', slots: { title: 'A title', body: 'A text.' } }] };
  const plain = renderDeck(few, texts, { layout: 'adaptive' });
  const second = renderDeck(few, { ...texts, colors: { accent2: 'C8102E' } }, { layout: 'adaptive' });
  assert.deepEqual(second.pptx, plain.pptx);
  assert.deepEqual(second.warnings.map((issue) => [issue.slide, issue.message]), [[undefined, 'the template has no second colour; "accent2" is left out']]);
  const onSlide = renderDeck(few, { ...texts, slides: [{ ...(texts.slides[0] as DeckSpec['slides'][number]), colors: { accent2: 'C8102E' } }] }, { layout: 'adaptive' });
  assert.deepEqual(onSlide.warnings.map((issue) => [issue.slide, issue.message]), [[undefined, 'slide 1: the template has no second colour; "accent2" is left out']]);
  // With a second colour to change, nothing is said.
  assert.deepEqual(warnings({ ...deck, colors: { accent2: 'C8102E' } }), []);
});

test('a colour close to a page that changes with the colours is not called close to it', () => {
  // A page in the main colour takes the new colour too: what is drawn on it in that colour is
  // drawn on a page of another colour, and shows.
  const zip = unzipSync(demoPptx);
  zip['ppt/slides/slide2.xml'] = strToU8(strFromU8(zip['ppt/slides/slide2.xml'] as Uint8Array).replace('<p:spTree>', `<p:bg><p:bgPr><a:solidFill><a:srgbClr val="${TEAL}"/></a:solidFill><a:effectLst/></p:bgPr></p:bg><p:spTree>`));
  const teal = loadTemplate(zipSync(zip), manifest);
  const texts: DeckSpec = { version: 1, fields: { organizationDepartment: 'Riverbend' }, slides: [{ pattern: 'text-only', slots: { title: 'A title', body: 'A text.' } }] };
  assert.ok(teal.pkg.readText('ppt/slides/slide2.xml').includes(`<p:bg><p:bgPr><a:solidFill><a:srgbClr val="${TEAL}"/>`), 'the text slide of this template has a teal page');
  assert.deepEqual(renderDeck(teal, { ...texts, colors: { accent: '117F7E' } }, { layout: 'adaptive' }).warnings, []);
  // On the white page of the template as it ships, the same colour is far from the page, and white is close.
  assert.deepEqual(renderDeck(template, { ...texts, colors: { accent: '117F7E' } }, { layout: 'adaptive' }).warnings, []);
  assert.equal(renderDeck(template, { ...texts, colors: { accent: 'FFFFFF' } }, { layout: 'adaptive' }).warnings.length, 1);
});

test('colours of a slide that are no colours follow the deck; only {} asks for the template\'s own', () => {
  const blueDeck = (colors: unknown): DeckSpec => {
    const given = structuredClone(deck);
    (given.slides[3] as unknown as { colors: unknown }).colors = colors;
    return { ...given, colors: { accent: BLUE } };
  };
  const figureSlide = (colors: unknown) => colours(text(render(blueDeck(colors)), 'ppt/slides/slide4.xml'));
  // Nothing usable is left of these, and the deck's check says so where there is something to say.
  for (const none of [{ accent: 'red' }, { accent: '#15' }, { primary: 'C8102E' }, 'C8102E', null]) {
    const slide = figureSlide(none);
    assert.ok(slide.has(BLUE) && !slide.has(TEAL), `${JSON.stringify(none)} follows the deck`);
  }
  assert.ok(validateDeck(template, blueDeck({ accent: 'red' }), { layout: 'adaptive', assets: Object.keys(assets) }).some((issue) => /^slide 4: .*"red".*ignored/.test(issue.message)));
  // Colours that ask for nothing keep the template's own, as `{}` is documented to.
  for (const own of [{}, { accent: '' }]) {
    const slide = figureSlide(own);
    assert.ok(slide.has(TEAL) && !slide.has(BLUE), `${JSON.stringify(own)} keeps the template's colours`);
  }
  // A slide that follows the deck is the slide it would be with no colours of its own.
  assert.equal(text(render(blueDeck(null)), 'ppt/slides/slide4.xml'), text(render({ ...deck, colors: { accent: BLUE } }), 'ppt/slides/slide4.xml'));
});

test('the text of a native chart keeps reading when the manifest sets it in the main colour', () => {
  const inTeal = structuredClone(manifest);
  const slot = inTeal.patterns.find((pattern) => pattern.id === 'chart-native')?.slots.find((entry) => entry.kind === 'chart');
  assert.ok(slot?.kind === 'chart' && slot.text);
  slot.text = { ...slot.text, color: TEAL };
  const charted = loadTemplate(demoPptx, inTeal);
  const one: DeckSpec = { version: 1, fields: { organizationDepartment: 'Riverbend' }, slides: [deck.slides[2] as DeckSpec['slides'][number]] };
  // The colour of the chart's own text: the default it states for everything written in it.
  const chartText = (given: DeckSpec) => /<a:defRPr sz="\d+"[^>]*><a:solidFill><a:srgbClr val="([0-9A-F]{6})"/.exec(text(renderDeck(charted, given, { layout: 'adaptive' }).pptx, 'ppt/charts/chart1.xml'))?.[1];
  assert.equal(chartText(one), TEAL);
  const yellow = chartText({ ...one, colors: { accent: 'FFD400' } });
  assert.ok(yellow && yellow !== 'FFD400' && contrast(yellow, 'FFFFFF') >= 3, `${yellow} reads on the white page`);
  assert.ok(hueDistance(hexToOklch(yellow as string).h, hexToOklch('FFD400').h) < 25, 'and is still a yellow');
  // A colour that reads is the colour asked for.
  assert.equal(chartText({ ...one, colors: { accent: BLUE } }), BLUE);
});
