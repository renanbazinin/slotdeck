import assert from 'node:assert/strict';
import { test } from 'node:test';
import { placeholderPhoto } from '../examples/placeholder.ts';
import { PptxPackage, inspectTemplate } from '../src/index.ts';
import { relativeTarget, relsOwner, relsPath, resolveTarget } from '../src/pptx/package.ts';
import { NS, descendants, parseXml, rootElement } from '../src/pptx/xml.ts';
import { detectDirection, marksRightToLeft, setBodyText } from '../src/render/text.ts';
import { readImageInfo } from '../src/util/image.ts';
import { buildFixturePptx } from './fixture.ts';

test('relationship paths', () => {
  assert.equal(relsPath('ppt/slides/slide1.xml'), 'ppt/slides/_rels/slide1.xml.rels');
  assert.equal(relsPath(''), '_rels/.rels');
  assert.equal(relsOwner('ppt/slides/_rels/slide1.xml.rels'), 'ppt/slides/slide1.xml');
  assert.equal(relsOwner('_rels/.rels'), '');
  assert.equal(relsOwner('ppt/slides/slide1.xml'), undefined);

  assert.equal(resolveTarget('ppt/slides/slide1.xml', '../media/image1.png'), 'ppt/media/image1.png');
  assert.equal(resolveTarget('ppt/presentation.xml', 'slides/slide1.xml'), 'ppt/slides/slide1.xml');
  assert.equal(resolveTarget('ppt/presentation.xml', '/docProps/app.xml'), 'docProps/app.xml');
  assert.equal(resolveTarget('', 'ppt/presentation.xml'), 'ppt/presentation.xml');

  assert.equal(relativeTarget('ppt/slides/slide1.xml', 'ppt/media/image1.png'), '../media/image1.png');
  assert.equal(relativeTarget('ppt/presentation.xml', 'ppt/slides/slide1.xml'), 'slides/slide1.xml');
  assert.equal(relativeTarget('ppt/slides/slide1.xml', 'ppt/slides/slide2.xml'), 'slide2.xml');
});

test('image headers', () => {
  assert.deepEqual(readImageInfo(placeholderPhoto(40, 30, [0, 0, 0], [9, 9, 9])), {
    format: 'png',
    extension: 'png',
    contentType: 'image/png',
    width: 40,
    height: 30,
  });

  const gif = new Uint8Array([0x47, 0x49, 0x46, 0x38, 0x39, 0x61, 0x20, 0x03, 0x58, 0x02]);
  assert.deepEqual([readImageInfo(gif).width, readImageInfo(gif).height], [800, 600]);

  // SOI, an APP0 segment to skip, then a baseline frame header for 640 x 480.
  const jpeg = new Uint8Array([
    0xff, 0xd8, 0xff, 0xe0, 0x00, 0x04, 0x00, 0x00, 0xff, 0xc0, 0x00, 0x11, 0x08, 0x01, 0xe0, 0x02, 0x80, 0x03, 0x00, 0x00, 0x00,
  ]);
  const info = readImageInfo(jpeg);
  assert.deepEqual([info.format, info.width, info.height], ['jpeg', 640, 480]);

  assert.throws(() => readImageInfo(new Uint8Array([1, 2, 3, 4, 5, 6, 7, 8])), /Unsupported image format/);
});

test('text direction detection', () => {
  assert.equal(detectDirection('שלום world'), 'rtl');
  assert.equal(detectDirection('Hello שלום'), 'ltr');
  assert.equal(detectDirection('12 כלבים'), 'rtl');
  assert.equal(detectDirection('مرحبا'), 'rtl');
  assert.equal(detectDirection('78%'), undefined);
  assert.equal(detectDirection(''), undefined);

  assert.equal(marksRightToLeft('שלום world'), true);
  assert.equal(marksRightToLeft('AI בשירות הלקוחות'), true);
  assert.equal(marksRightToLeft('Meeting with דני'), false);
  assert.equal(marksRightToLeft('78%'), false);
});

test('filled text: no saved spelling check, and a Hebrew language on mostly Hebrew text', () => {
  const a = 'http://schemas.openxmlformats.org/drawingml/2006/main';
  // The run as a template saved it after PowerPoint checked Hebrew words as English.
  const run = '<a:r><a:rPr lang="en-GB" sz="1038" b="1" dirty="0" err="1"/><a:t>משימת</a:t></a:r>';
  const fill = (text: string) => {
    const doc = parseXml(`<a:txBody xmlns:a="${a}"><a:bodyPr/><a:p>${run}<a:endParaRPr lang="en-GB" dirty="0"/></a:p></a:txBody>`);
    setBodyText(rootElement(doc), [text], { flipAlign: false });
    const props = descendants(doc, NS.a, 'rPr')[0];
    const end = descendants(doc, NS.a, 'endParaRPr')[0];
    return { lang: props?.getAttribute('lang'), err: props?.hasAttribute('err'), dirty: props?.hasAttribute('dirty') || end?.hasAttribute('dirty') };
  };

  assert.deepEqual(fill('תוכנית העבודה לרבעון הבא'), { lang: 'he-IL', err: false, dirty: false });
  assert.deepEqual(fill('AI בשירות הלקוחות'), { lang: 'he-IL', err: false, dirty: false });
  assert.deepEqual(fill('Meeting with דני'), { lang: 'en-GB', err: false, dirty: false });
});

test('inspector lists shapes with slide-relative boxes through groups', () => {
  const info = inspectTemplate(PptxPackage.load(buildFixturePptx()));
  assert.equal(info.slides.length, 5);
  assert.equal(info.slides[2]?.shapes.find((shape) => shape.id === '3')?.kind, 'table');
  const first = info.slides[0];
  assert.ok(first);
  const byId = new Map(first.shapes.map((shape) => [shape.id, shape]));

  assert.equal(byId.get('2')?.kind, 'text');
  assert.deepEqual(byId.get('3')?.paragraphs, ['SAMPLE HEAD', 'SAMPLE BODYSAMPLE MORE']);
  assert.deepEqual(byId.get('3')?.fontSizes, [20, 12]);
  assert.equal(byId.get('5')?.placeholder, 'sldNum');
  assert.equal(byId.get('4')?.image?.cropped, true);
  assert.equal(byId.get('4')?.image?.geometry, 'roundRect');
  assert.equal(byId.get('4')?.image?.width, 8);

  // Child at (200000, 200000) sized 500000 in a group that doubles x and sits at y = 3000000.
  const icon = byId.get('8');
  assert.equal(icon?.groupId, '7');
  assert.deepEqual(icon?.box, {
    x: Math.round((400000 / 9144000) * 10000) / 10000,
    y: Math.round((3200000 / 5143500) * 10000) / 10000,
    w: Math.round((1000000 / 9144000) * 10000) / 10000,
    h: Math.round((500000 / 5143500) * 10000) / 10000,
  });
});
