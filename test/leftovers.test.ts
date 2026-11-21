// What a slide cloned from the template leaves behind of the template's sample (0.7.0): a link on
// the sample's run, its comments, its add-ins' tags, and animations of shapes or paragraphs that
// are gone. No test calls a model.
import assert from 'node:assert/strict';
import { test } from 'node:test';
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate';
import { placeholderPhoto } from '../examples/placeholder.ts';
import { PptxPackage, loadTemplate, renderDeck } from '../src/index.ts';
import type { DeckSpec, SlideSpec } from '../src/index.ts';
import { readRels } from '../src/pptx/package.ts';
import { parseXml, rootElement, serializeXml } from '../src/pptx/xml.ts';
import { dropLeftovers, keptUnnamed } from '../src/render/leftovers.ts';
import { buildFixturePptx, fixtureManifest } from './fixture.ts';

const P = 'http://schemas.openxmlformats.org/presentationml/2006/main';
const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';

/** One click that shows a shape, or some of its paragraphs: an entrance effect, as PowerPoint writes it. */
const click = (id: number, spid: string, paragraphs = '') =>
  `<p:par><p:cTn id="${id}" fill="hold"><p:stCondLst><p:cond delay="indefinite"/></p:stCondLst><p:childTnLst>` +
  `<p:par><p:cTn id="${id + 1}" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst>` +
  `<p:par><p:cTn id="${id + 2}" presetID="1" presetClass="entr" presetSubtype="0" fill="hold" nodeType="clickEffect"><p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst>` +
  `<p:set><p:cBhvr><p:cTn id="${id + 3}" dur="1" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst></p:cTn><p:tgtEl><p:spTgt spid="${spid}">${paragraphs}</p:spTgt></p:tgtEl>` +
  `<p:attrNameLst><p:attrName>style.visibility</p:attrName></p:attrNameLst></p:cBhvr><p:to><p:strVal val="visible"/></p:to></p:set>` +
  `</p:childTnLst></p:cTn></p:par></p:childTnLst></p:cTn></p:par></p:childTnLst></p:cTn></p:par>`;

/** Animations of the title (#2), the optional text (#6) and the second paragraph of the body (#3). */
const timing =
  `<p:timing><p:tnLst><p:par><p:cTn id="1" dur="indefinite" restart="never" nodeType="tmRoot"><p:childTnLst>` +
  `<p:seq concurrent="1" nextAc="seek"><p:cTn id="2" dur="indefinite" nodeType="mainSeq"><p:childTnLst>` +
  `${click(3, '2')}${click(7, '6')}${click(11, '3', '<p:txEl><p:pRg st="1" end="1"/></p:txEl>')}` +
  `</p:childTnLst></p:cTn><p:prevCondLst><p:cond evt="onPrev" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:prevCondLst>` +
  `<p:nextCondLst><p:cond evt="onNext" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:nextCondLst></p:seq>` +
  `</p:childTnLst></p:cTn></p:par></p:tnLst><p:bldLst><p:bldP spid="3" grpId="0" build="p"/><p:bldP spid="6" grpId="0"/><p:bldP spid="2" grpId="0"/></p:bldLst></p:timing>`;

/**
 * The fixture whose first slide carries what a template's sample may: a link on the title's sample
 * run, a link on its icon (part of the design), its add-ins' tags, a comment and animations.
 */
function leftoverFixture(withTiming = true): Uint8Array {
  const files = unzipSync(buildFixturePptx());
  const slide = strFromU8(files['ppt/slides/slide1.xml'] as Uint8Array)
    .replace('<a:rPr lang="he-IL" sz="2600"/><a:t>SAMPLE TITLE</a:t>', '<a:rPr lang="he-IL" sz="2600"><a:hlinkClick r:id="rId10"/></a:rPr><a:t>SAMPLE TITLE</a:t>')
    .replace('<a:endParaRPr sz="2600"/>', '<a:endParaRPr sz="2600"><a:hlinkClick r:id="rId10"/></a:endParaRPr>')
    .replace('<p:cNvPr id="8" name="Icon"/>', '<p:cNvPr id="8" name="Icon"><a:hlinkClick r:id="rId11"/></p:cNvPr>')
    .replace('<p:cNvPr id="6" name="Optional"/><p:cNvSpPr txBox="1"/><p:nvPr/>', '<p:cNvPr id="6" name="Optional"/><p:cNvSpPr txBox="1"/><p:nvPr><p:custDataLst><p:tags r:id="rId12"/></p:custDataLst></p:nvPr>')
    .replace('</p:spTree>', '</p:spTree><p:custDataLst><p:tags r:id="rId12"/></p:custDataLst>')
    .replace('</p:cSld>', `</p:cSld>${withTiming ? timing : ''}`);
  assert.equal((slide.match(/hlinkClick/g) ?? []).length, 3);
  files['ppt/slides/slide1.xml'] = strToU8(slide);
  files['ppt/slides/_rels/slide1.xml.rels'] = strToU8(
    strFromU8(files['ppt/slides/_rels/slide1.xml.rels'] as Uint8Array).replace(
      '</Relationships>',
      `<Relationship Id="rId10" Type="${REL}/hyperlink" Target="https://template.example/sample" TargetMode="External"/>` +
        `<Relationship Id="rId11" Type="${REL}/hyperlink" Target="https://brand.example/" TargetMode="External"/>` +
        `<Relationship Id="rId12" Type="${REL}/tags" Target="../tags/tag1.xml"/>` +
        `<Relationship Id="rId13" Type="${REL}/comments" Target="../comments/comment1.xml"/></Relationships>`,
    ),
  );
  files['ppt/tags/tag1.xml'] = strToU8(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<p:tagLst xmlns:p="${P}"><p:tag name="SAMPLE_ADDIN" val="1"/></p:tagLst>`);
  files['ppt/comments/comment1.xml'] = strToU8(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<p:cmLst xmlns:p="${P}"><p:cm authorId="0" dt="2026-01-01T00:00:00.000" idx="1"><p:pos x="10" y="10"/><p:text>Fix the sample before it ships</p:text></p:cm></p:cmLst>`);
  files['[Content_Types].xml'] = strToU8(
    strFromU8(files['[Content_Types].xml'] as Uint8Array).replace(
      '</Types>',
      '<Override PartName="/ppt/tags/tag1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.tags+xml"/>' +
        '<Override PartName="/ppt/comments/comment1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.comments+xml"/></Types>',
    ),
  );
  return zipSync(files);
}

const photo = placeholderPhoto(100, 100, [10, 20, 30], [200, 210, 220]);
const content = (slots: SlideSpec['slots']): DeckSpec => ({ version: 1, slides: [{ pattern: 'content', slots: { title: 'שלום', body: 'גוף', photo: { asset: 'p' }, ...slots } }] });
const render = (file: Uint8Array, deck: DeckSpec) => {
  const out = renderDeck(loadTemplate(file, fixtureManifest), deck, { assets: { p: photo }, layout: 'adaptive' }).pptx;
  return { pkg: PptxPackage.load(out), slide: PptxPackage.load(out).readText('ppt/slides/slide1.xml') };
};

test('the link of a sample run, the sample\'s comments and its add-ins\' tags are left behind; a link of the design stays', () => {
  const { pkg, slide } = render(leftoverFixture(), content({}));
  // The person's title is not a link to the template's page; the icon still links to the brand.
  assert.equal((slide.match(/hlinkClick/g) ?? []).length, 1);
  assert.match(slide, /<p:cNvPr id="8" name="Icon"><a:hlinkClick r:id="rId11"\/><\/p:cNvPr>/);
  assert.ok(!slide.includes('custDataLst'));
  const rels = readRels(pkg, 'ppt/slides/slide1.xml');
  // The photo, the template's icon that stays, the icon's link and the layout.
  assert.deepEqual(rels.map((rel) => rel.type.split('/').at(-1)).sort(), ['hyperlink', 'image', 'image', 'slideLayout']);
  assert.equal(rels.find((rel) => rel.type.endsWith('/hyperlink'))?.target, 'https://brand.example/');
  for (const part of ['ppt/tags/tag1.xml', 'ppt/comments/comment1.xml']) assert.ok(!pkg.has(part), part);
  assert.ok(!pkg.readText('[Content_Types].xml').includes('tag1.xml') && !pkg.readText('[Content_Types].xml').includes('comment1.xml'));
});

test('animations of shapes or paragraphs that are gone are taken out, and the rest stays', () => {
  const targets = (slide: string) => [...slide.matchAll(/<p:spTgt spid="(\d+)"(?:\/>|>(<p:txEl>.*?<\/p:txEl>)?)/g)].map((found) => `${found[1]}${found[2] ? ' paragraph' : ''}`);
  const builds = (slide: string) => [...slide.matchAll(/<p:bldP spid="(\d+)"/g)].map((found) => found[1]);
  // The optional text is left out and the body has one paragraph: only the title's click is left.
  const short = render(leftoverFixture(), content({})).slide;
  assert.deepEqual(targets(short), ['2']);
  assert.deepEqual(builds(short), ['2']);
  assert.equal((short.match(/nodeType="clickEffect"/g) ?? []).length, 1);
  // With both, all three stay as the template has them.
  const full = render(leftoverFixture(), content({ note: 'Note', body: ['One', 'Two'] })).slide;
  assert.deepEqual(targets(full), ['2', '6', '3 paragraph']);
  assert.deepEqual(builds(full), ['3', '6', '2']);
  // A slide with nothing left to animate has no animations at all.
  const files = unzipSync(leftoverFixture());
  files['ppt/slides/slide1.xml'] = strToU8(strFromU8(files['ppt/slides/slide1.xml'] as Uint8Array).replace(click(3, '2'), ''));
  assert.ok(!render(zipSync(files), content({})).slide.includes('p:timing'));
  // A template without animations is drawn as before.
  assert.ok(!render(leftoverFixture(false), content({})).slide.includes('p:timing'));
});

test('a sequence left empty goes, and a missing trigger takes only its own sequence', () => {
  const P = 'http://schemas.openxmlformats.org/presentationml/2006/main';
  const shape = (id: number) => `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="S${id}"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr/></p:sp>`;
  const interactive = (trigger: string, id: number, target: string) =>
    `<p:seq concurrent="1" nextAc="seek"><p:cTn id="${id}" restart="whenNotActive" fill="hold" evtFilter="cancelBubble" nodeType="interactiveSeq"><p:stCondLst><p:cond evt="onClick" delay="0"><p:tgtEl><p:spTgt spid="${trigger}"/></p:tgtEl></p:cond></p:stCondLst><p:childTnLst>${click(id + 1, target)}</p:childTnLst></p:cTn><p:nextCondLst><p:cond evt="onClick" delay="0"><p:tgtEl><p:spTgt spid="${trigger}"/></p:tgtEl></p:cond></p:nextCondLst></p:seq>`;
  const slideWith = (sequences: string) =>
    `<p:sld xmlns:p="${P}" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>${shape(2)}${shape(3)}${shape(4)}</p:spTree></p:cSld>` +
    `<p:timing><p:tnLst><p:par><p:cTn id="1" dur="indefinite" restart="never" nodeType="tmRoot"><p:childTnLst>${sequences}</p:childTnLst></p:cTn></p:par></p:tnLst></p:timing></p:sld>`;
  const main = (target: string) => `<p:seq concurrent="1" nextAc="seek"><p:cTn id="2" dur="indefinite" nodeType="mainSeq"><p:childTnLst>${click(3, target)}</p:childTnLst></p:cTn></p:seq>`;
  const clean = (xml: string) => {
    const doc = parseXml(xml);
    dropLeftovers(rootElement(doc));
    return serializeXml(doc);
  };
  // The main sequence animates a shape that is gone; a click on #4 shows #3: only the click is left, and no list is empty.
  const kept = clean(slideWith(main('9') + interactive('4', 20, '3')));
  assert.ok(kept.includes('nodeType="interactiveSeq"') && !kept.includes('nodeType="mainSeq"') && !kept.includes('<p:childTnLst/>'), kept);
  // The shape that is clicked is gone: its sequence goes, and the main one stays.
  const trigger = clean(slideWith(main('2') + interactive('9', 20, '3')));
  assert.ok(trigger.includes('nodeType="mainSeq"') && !trigger.includes('interactiveSeq'), trigger);
});

test('a relationship the slide does not name stays where what it shows needs it, and a leftover of the sample goes', () => {
  const D = 'http://schemas.openxmlformats.org/drawingml/2006/diagram';
  const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
  const slide = (body: string) => rootElement(parseXml(`<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:r="${R}"><p:cSld><p:spTree>${body}</p:spTree></p:cSld></p:sld>`));
  const diagram = slide(`<p:graphicFrame><a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:graphicData><dgm:relIds xmlns:dgm="${D}" r:dm="rId2" r:lo="rId3" r:qs="rId4" r:cs="rId5"/></a:graphicData></a:graphic></p:graphicFrame>`);
  const plain = slide('');
  const drawing = 'http://schemas.microsoft.com/office/2007/relationships/diagramDrawing';
  // A diagram's drawing is named in the diagram's data: it stays while the slide shows a diagram.
  assert.deepEqual([keptUnnamed(drawing, diagram, false), keptUnnamed(drawing, plain, false)], [true, false]);
  const vml = `${R}/vmlDrawing`;
  assert.deepEqual([keptUnnamed(vml, slide('<p:graphicFrame><p:oleObj/></p:graphicFrame>'), false), keptUnnamed(vml, plain, false)], [true, false]);
  // The sample's link, comments and tags go; a kind not known to be a leftover stays, except on a free slide.
  for (const kind of ['hyperlink', 'comments', 'tags', 'image', 'chart']) assert.equal(keptUnnamed(`${R}/${kind}`, plain, false), false, kind);
  assert.equal(keptUnnamed('http://schemas.microsoft.com/office/2018/10/relationships/comments', plain, false), false);
  assert.deepEqual([keptUnnamed(`${R}/someAddIn`, plain, false), keptUnnamed(`${R}/someAddIn`, plain, true)], [true, false]);
});
