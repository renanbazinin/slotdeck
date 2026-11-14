// A tiny synthetic pptx with the features the renderer has to cope with: mixed-format
// paragraphs, a cropped picture, a picture inside a scaled group, a slide-number field,
// speaker notes, revision history and named sections. It is only read by slotdeck itself;
// the master and layout are stubs, so PowerPoint would not open it.
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate';
import { placeholderPhoto } from '../examples/placeholder.ts';
import type { TemplateManifest } from '../src/index.ts';

const A = 'http://schemas.openxmlformats.org/drawingml/2006/main';
const P = 'http://schemas.openxmlformats.org/presentationml/2006/main';
const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const HEAD = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';
const NS = `xmlns:a="${A}" xmlns:r="${R}" xmlns:p="${P}"`;

const rels = (items: Array<[id: string, type: string, target: string]>) =>
  `${HEAD}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${items
    .map(([id, type, target]) => `<Relationship Id="${id}" Type="${type}" Target="${target}"/>`)
    .join('')}</Relationships>`;

const textBox = (id: number, name: string, x: number, paragraphs: string) =>
  `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="${name}"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>` +
  `<p:spPr><a:xfrm><a:off x="${x}" y="100000"/><a:ext cx="3000000" cy="500000"/></a:xfrm></p:spPr>` +
  `<p:txBody><a:bodyPr/><a:lstStyle/>${paragraphs}</p:txBody></p:sp>`;

const slide1 =
  `${HEAD}<p:sld ${NS}><p:cSld><p:spTree>` +
  `<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>` +
  textBox(
    2,
    'Title',
    100000,
    `<a:p><a:pPr algn="r" rtl="1"/><a:r><a:rPr lang="he-IL" sz="2600" b="1"/><a:t>SAMPLE LABEL </a:t></a:r>` +
      `<a:r><a:rPr lang="he-IL" sz="2600"/><a:t>SAMPLE TITLE</a:t></a:r><a:endParaRPr sz="2600"/></a:p>`,
  ) +
  textBox(
    3,
    'Body',
    4000000,
    `<a:p><a:pPr algn="r" rtl="1"/><a:r><a:rPr sz="2000" b="1"/><a:t>SAMPLE HEAD</a:t></a:r></a:p>` +
      `<a:p><a:pPr algn="r" rtl="1"/><a:r><a:rPr sz="1200"/><a:t>SAMPLE BODY</a:t></a:r><a:br/>` +
      `<a:r><a:rPr sz="1200" i="1"/><a:t>SAMPLE MORE</a:t></a:r></a:p>`,
  ) +
  `<p:pic><p:nvPicPr><p:cNvPr id="4" name="Photo" title="old.png" descr="old"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr>` +
  `<p:blipFill><a:blip r:embed="rId2"/><a:srcRect l="10000" r="10000"/><a:stretch><a:fillRect/></a:stretch></p:blipFill>` +
  `<p:spPr><a:xfrm><a:off x="1000000" y="1000000"/><a:ext cx="4000000" cy="2000000"/></a:xfrm>` +
  `<a:prstGeom prst="roundRect"><a:avLst/></a:prstGeom></p:spPr></p:pic>` +
  `<p:sp><p:nvSpPr><p:cNvPr id="5" name="Number"/><p:cNvSpPr/><p:nvPr><p:ph type="sldNum" idx="12"/></p:nvPr></p:nvSpPr><p:spPr/>` +
  `<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:fld id="{B6F15528-21DE-4FAA-801E-634DDDAF4B2B}" type="slidenum">` +
  `<a:rPr lang="en-US"/><a:t>1</a:t></a:fld></a:p></p:txBody></p:sp>` +
  textBox(6, 'Optional', 100000, `<a:p><a:r><a:rPr sz="1000"/><a:t>SAMPLE OPTIONAL</a:t></a:r></a:p>`) +
  // Group that stretches its children 2x horizontally: a 500000 x 500000 child renders 2:1.
  `<p:grpSp><p:nvGrpSpPr><p:cNvPr id="7" name="Group"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>` +
  `<p:grpSpPr><a:xfrm><a:off x="0" y="3000000"/><a:ext cx="2000000" cy="1000000"/><a:chOff x="0" y="0"/><a:chExt cx="1000000" cy="1000000"/></a:xfrm></p:grpSpPr>` +
  `<p:pic><p:nvPicPr><p:cNvPr id="8" name="Icon"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr>` +
  `<p:blipFill><a:blip r:embed="rId2"/><a:stretch><a:fillRect/></a:stretch></p:blipFill>` +
  `<p:spPr><a:xfrm><a:off x="200000" y="200000"/><a:ext cx="500000" cy="500000"/></a:xfrm></p:spPr></p:pic></p:grpSp>` +
  `</p:spTree></p:cSld></p:sld>`;

const slide2 =
  `${HEAD}<p:sld ${NS}><p:cSld><p:spTree>` +
  `<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>` +
  textBox(2, 'Closing', 100000, `<a:p><a:pPr algn="r" rtl="1"/><a:r><a:rPr sz="5200"/><a:t>SAMPLE CLOSING</a:t></a:r></a:p>`) +
  `</p:spTree></p:cSld></p:sld>`;

const tableCell = (text: string, fill: string) =>
  `<a:tc><a:txBody><a:bodyPr/><a:lstStyle/><a:p><a:pPr algn="l" rtl="0"/>` +
  (text ? `<a:r><a:rPr lang="en-US"/><a:t>${text}</a:t></a:r>` : '') +
  `<a:endParaRPr/></a:p></a:txBody><a:tcPr>${fill ? `<a:solidFill><a:srgbClr val="${fill}"/></a:solidFill>` : ''}</a:tcPr></a:tc>`;

const tableRow = (id: number, text: string, fill: string) =>
  `<a:tr h="300000">${tableCell(text, fill).repeat(3)}` +
  `<a:extLst><a:ext uri="{0D108BD9-81ED-4DB2-BD59-A6C34878D82A}"><a16:rowId xmlns:a16="http://schemas.microsoft.com/office/drawing/2014/main" val="${id}"/></a:ext></a:extLst></a:tr>`;

const tableColumn = (id: number) =>
  `<a:gridCol w="2000000"><a:extLst><a:ext uri="{9D8B030D-6E8A-4147-A177-3AD203B41FA5}">` +
  `<a16:colId xmlns:a16="http://schemas.microsoft.com/office/drawing/2014/main" val="${id}"/></a:ext></a:extLst></a:gridCol>`;

// A three-column table: a filled header row, then a plain and a shaded body row.
const slide3 =
  `${HEAD}<p:sld ${NS}><p:cSld><p:spTree>` +
  `<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>` +
  textBox(2, 'Title', 100000, `<a:p><a:pPr algn="r" rtl="1"/><a:r><a:rPr sz="2600"/><a:t>SAMPLE TABLE TITLE</a:t></a:r></a:p>`) +
  `<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="3" name="Table"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr>` +
  `<p:xfrm><a:off x="100000" y="1000000"/><a:ext cx="6000000" cy="900000"/></p:xfrm>` +
  `<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table"><a:tbl><a:tblPr/>` +
  `<a:tblGrid>${tableColumn(1)}${tableColumn(2)}${tableColumn(3)}</a:tblGrid>` +
  tableRow(10, 'SAMPLE H', '00683A') +
  tableRow(11, '', '') +
  tableRow(12, '', 'DCDDDE') +
  `</a:tbl></a:graphicData></a:graphic></p:graphicFrame>` +
  `</p:spTree></p:cSld></p:sld>`;

// A row of three cards, first in reading order at the right. Each card is a group holding a text
// box and a picture; the first card also has a text box that sits outside its group.
const card = (id: number, x: number, y = 2000000) =>
  `<p:grpSp><p:nvGrpSpPr><p:cNvPr id="${id}" name="Card"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>` +
  `<p:grpSpPr><a:xfrm><a:off x="${x}" y="${y}"/><a:ext cx="1000000" cy="1500000"/><a:chOff x="${x}" y="${y}"/><a:chExt cx="1000000" cy="1500000"/></a:xfrm></p:grpSpPr>` +
  `<p:sp><p:nvSpPr><p:cNvPr id="${id + 1}" name="Card text"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>` +
  `<p:spPr><a:xfrm><a:off x="${x}" y="${y + 800000}"/><a:ext cx="1000000" cy="400000"/></a:xfrm></p:spPr>` +
  `<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:pPr algn="ctr" rtl="1"/><a:r><a:rPr sz="900"/><a:t>SAMPLE CARD</a:t></a:r></a:p></p:txBody></p:sp>` +
  `<p:pic><p:nvPicPr><p:cNvPr id="${id + 2}" name="Card icon"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr>` +
  `<p:blipFill><a:blip r:embed="rId2"/><a:stretch><a:fillRect/></a:stretch></p:blipFill>` +
  `<p:spPr><a:xfrm><a:off x="${x + 300000}" y="${y + 100000}"/><a:ext cx="400000" cy="400000"/></a:xfrm></p:spPr></p:pic></p:grpSp>`;

const slide4 =
  `${HEAD}<p:sld ${NS}><p:cSld><p:spTree>` +
  `<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>` +
  textBox(2, 'Title', 100000, `<a:p><a:pPr algn="r" rtl="1"/><a:r><a:rPr sz="2600"/><a:t>SAMPLE CARDS TITLE</a:t></a:r></a:p>`) +
  card(10, 6000000) +
  `<p:sp><p:nvSpPr><p:cNvPr id="13" name="Stray label"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>` +
  `<p:spPr><a:xfrm><a:off x="6100000" y="3600000"/><a:ext cx="800000" cy="200000"/></a:xfrm></p:spPr>` +
  `<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr sz="800"/><a:t>first</a:t></a:r></a:p></p:txBody></p:sp>` +
  card(20, 4000000) +
  card(30, 2000000) +
  `</p:spTree></p:cSld></p:sld>`;

// A two-by-two grid of the same cards: first row at the top, first card of each row at the right.
const slide5 =
  `${HEAD}<p:sld ${NS}><p:cSld><p:spTree>` +
  `<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>` +
  textBox(2, 'Title', 100000, `<a:p><a:pPr algn="r" rtl="1"/><a:r><a:rPr sz="2600"/><a:t>SAMPLE GRID TITLE</a:t></a:r></a:p>`) +
  card(10, 6000000, 1000000) +
  card(20, 4000000, 1000000) +
  card(30, 6000000, 3000000) +
  card(40, 4000000, 3000000) +
  `</p:spTree></p:cSld></p:sld>`;

const presentation =
  `${HEAD}<p:presentation ${NS}>` +
  `<p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rId1"/></p:sldMasterIdLst>` +
  `<p:sldIdLst><p:sldId id="256" r:id="rId2"/><p:sldId id="257" r:id="rId3"/><p:sldId id="258" r:id="rId6"/><p:sldId id="259" r:id="rId7"/><p:sldId id="260" r:id="rId8"/></p:sldIdLst>` +
  `<p:sldSz cx="9144000" cy="5143500"/><p:notesSz cx="6858000" cy="9144000"/>` +
  `<p:extLst><p:ext uri="{521415D9-36F7-43E2-AB2F-B90AF26B5E84}">` +
  `<p14:sectionLst xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main">` +
  `<p14:section name="Intro" id="{6B2B4B2C-6F4E-4C8B-9B0A-0D1E2F3A4B5C}"><p14:sldIdLst><p14:sldId id="256"/></p14:sldIdLst></p14:section>` +
  `</p14:sectionLst></p:ext></p:extLst></p:presentation>`;

const contentTypes =
  `${HEAD}<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">` +
  `<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>` +
  `<Default Extension="xml" ContentType="application/xml"/>` +
  `<Default Extension="png" ContentType="image/png"/>` +
  `<Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>` +
  `<Override PartName="/ppt/slideMasters/slideMaster1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"/>` +
  `<Override PartName="/ppt/slideLayouts/slideLayout1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/>` +
  `<Override PartName="/ppt/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>` +
  `<Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>` +
  `<Override PartName="/ppt/slides/slide2.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>` +
  `<Override PartName="/ppt/slides/slide3.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>` +
  `<Override PartName="/ppt/slides/slide4.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>` +
  `<Override PartName="/ppt/slides/slide5.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>` +
  `<Override PartName="/ppt/notesSlides/notesSlide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.notesSlide+xml"/>` +
  `<Override PartName="/ppt/changesInfos/changesInfo1.xml" ContentType="application/vnd.ms-powerpoint.changesinfo+xml"/>` +
  `</Types>`;

export function buildFixturePptx(): Uint8Array {
  const files: Record<string, string | Uint8Array> = {
    '[Content_Types].xml': contentTypes,
    '_rels/.rels': rels([['rId1', `${REL}/officeDocument`, 'ppt/presentation.xml']]),
    'ppt/presentation.xml': presentation,
    'ppt/_rels/presentation.xml.rels': rels([
      ['rId1', `${REL}/slideMaster`, 'slideMasters/slideMaster1.xml'],
      ['rId2', `${REL}/slide`, 'slides/slide1.xml'],
      ['rId3', `${REL}/slide`, 'slides/slide2.xml'],
      ['rId4', 'http://schemas.microsoft.com/office/2016/11/relationships/changesInfo', 'changesInfos/changesInfo1.xml'],
      ['rId5', `${REL}/theme`, 'theme/theme1.xml'],
      ['rId6', `${REL}/slide`, 'slides/slide3.xml'],
      ['rId7', `${REL}/slide`, 'slides/slide4.xml'],
      ['rId8', `${REL}/slide`, 'slides/slide5.xml'],
    ]),
    'ppt/slideMasters/slideMaster1.xml': `${HEAD}<p:sldMaster ${NS}><p:cSld><p:spTree/></p:cSld></p:sldMaster>`,
    'ppt/slideMasters/_rels/slideMaster1.xml.rels': rels([
      ['rId1', `${REL}/slideLayout`, '../slideLayouts/slideLayout1.xml'],
      ['rId2', `${REL}/theme`, '../theme/theme1.xml'],
    ]),
    'ppt/slideLayouts/slideLayout1.xml': `${HEAD}<p:sldLayout ${NS}><p:cSld><p:spTree/></p:cSld></p:sldLayout>`,
    'ppt/slideLayouts/_rels/slideLayout1.xml.rels': rels([['rId1', `${REL}/slideMaster`, '../slideMasters/slideMaster1.xml']]),
    'ppt/theme/theme1.xml': `${HEAD}<a:theme xmlns:a="${A}" name="Fixture"/>`,
    'ppt/slides/slide1.xml': slide1,
    'ppt/slides/_rels/slide1.xml.rels': rels([
      ['rId1', `${REL}/slideLayout`, '../slideLayouts/slideLayout1.xml'],
      ['rId2', `${REL}/image`, '../media/image1.png'],
      ['rId3', `${REL}/notesSlide`, '../notesSlides/notesSlide1.xml'],
    ]),
    'ppt/slides/slide2.xml': slide2,
    'ppt/slides/_rels/slide2.xml.rels': rels([['rId1', `${REL}/slideLayout`, '../slideLayouts/slideLayout1.xml']]),
    'ppt/slides/slide3.xml': slide3,
    'ppt/slides/_rels/slide3.xml.rels': rels([['rId1', `${REL}/slideLayout`, '../slideLayouts/slideLayout1.xml']]),
    'ppt/slides/slide5.xml': slide5,
    'ppt/slides/_rels/slide5.xml.rels': rels([
      ['rId1', `${REL}/slideLayout`, '../slideLayouts/slideLayout1.xml'],
      ['rId2', `${REL}/image`, '../media/image1.png'],
    ]),
    'ppt/slides/slide4.xml': slide4,
    'ppt/slides/_rels/slide4.xml.rels': rels([
      ['rId1', `${REL}/slideLayout`, '../slideLayouts/slideLayout1.xml'],
      ['rId2', `${REL}/image`, '../media/image1.png'],
    ]),
    'ppt/notesSlides/notesSlide1.xml': `${HEAD}<p:notes ${NS}><p:cSld><p:spTree/></p:cSld></p:notes>`,
    'ppt/notesSlides/_rels/notesSlide1.xml.rels': rels([['rId1', `${REL}/slide`, '../slides/slide1.xml']]),
    'ppt/changesInfos/changesInfo1.xml': `${HEAD}<pc:chgInfo xmlns:pc="http://schemas.microsoft.com/office/powerpoint/2013/main/command"/>`,
    'ppt/media/image1.png': placeholderPhoto(8, 8, [0, 0, 0], [255, 255, 255]),
  };
  return zipSync(Object.fromEntries(Object.entries(files).map(([name, data]) => [name, typeof data === 'string' ? strToU8(data) : data])));
}

/** The fixture with the XML of one of its slides changed, for a case the fixture does not draw. */
export function fixtureWithSlide(number: number, change: (xml: string) => string): Uint8Array {
  const files = unzipSync(buildFixturePptx());
  const name = `ppt/slides/slide${number}.xml`;
  files[name] = strToU8(change(strFromU8(files[name] as Uint8Array)));
  return zipSync(files);
}

export const fixtureManifest: TemplateManifest = {
  version: 1,
  name: 'fixture',
  language: 'he',
  direction: 'rtl',
  patterns: [
    {
      id: 'content',
      sourceSlide: 1,
      description: 'Title, body, photo and icon.',
      slots: [
        { name: 'title', kind: 'text', shapeId: '2', maxChars: 20, styleRun: 'last' },
        { name: 'body', kind: 'text', shapeId: '3', maxParagraphs: 3 },
        { name: 'photo', kind: 'image', imageKind: 'photo', shapeId: '4' },
        { name: 'note', kind: 'text', shapeId: '6', empty: 'remove' },
        { name: 'icon', kind: 'image', imageKind: 'icon', shapeId: '8', empty: 'keep' },
      ],
    },
    {
      id: 'closing',
      sourceSlide: 2,
      description: 'Closing words.',
      slots: [{ name: 'title', kind: 'text', shapeId: '2', deckField: 'closing' }],
    },
  ],
};

/** The fixture manifest plus a pattern whose row of cards can hold one to five. */
export const cardsManifest: TemplateManifest = {
  ...fixtureManifest,
  patterns: [
    ...fixtureManifest.patterns,
    {
      id: 'cards',
      sourceSlide: 4,
      description: 'A row of cards.',
      slots: [
        { name: 'title', kind: 'text', shapeId: '2', size: 26 },
        ...[11, 21, 31].flatMap((textId, index): TemplateManifest['patterns'][number]['slots'] => [
          { name: `card${index + 1}Text`, kind: 'text', shapeId: String(textId), size: 9, maxChars: 20 },
          { name: `card${index + 1}Icon`, kind: 'image', imageKind: 'icon', shapeId: String(textId + 1), empty: 'keep' },
        ]),
      ],
      repeat: {
        item: 'card',
        slots: ['card{n}Text', 'card{n}Icon'],
        shapes: [['10', '13'], ['20'], ['30']],
        min: 1,
        max: 5,
        span: { left: 0.05, right: 0.95 },
      },
    },
  ],
};

/** The fixture manifest plus a pattern whose two-by-two grid of cards can hold one to four. */
export const gridManifest: TemplateManifest = {
  ...fixtureManifest,
  patterns: [
    ...fixtureManifest.patterns,
    {
      id: 'grid',
      sourceSlide: 5,
      description: 'A grid of cards.',
      slots: [
        { name: 'title', kind: 'text', shapeId: '2', size: 26 },
        ...[11, 21, 31, 41].map((textId, index): TemplateManifest['patterns'][number]['slots'][number] => ({
          name: `card${index + 1}Text`,
          kind: 'text',
          shapeId: String(textId),
          size: 9,
        })),
      ],
      repeat: { item: 'card', slots: ['card{n}Text'], shapes: [['10'], ['20'], ['30'], ['40']], columns: 2, min: 1, max: 4 },
    },
  ],
};

/** The fixture manifest plus a pattern for the table slide, kept apart so tests of the two basic patterns stay simple. */
export const tableManifest: TemplateManifest = {
  ...fixtureManifest,
  patterns: [
    ...fixtureManifest.patterns,
    {
      id: 'table',
      sourceSlide: 3,
      description: 'A table.',
      slots: [
        { name: 'title', kind: 'text', shapeId: '2' },
        {
          name: 'table',
          kind: 'table',
          shapeId: '3',
          maxColumns: 3,
          maxRows: 3,
          maxCellChars: 10,
          headerText: { size: 10, bold: true, color: 'FFFFFF', font: 'Assistant' },
          bodyText: { size: 9 },
        },
      ],
    },
  ],
};
