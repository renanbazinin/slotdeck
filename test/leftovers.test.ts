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
