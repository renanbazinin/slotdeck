// Builds a native PowerPoint chart: the chart part, and the small Excel workbook that holds its
// data so the chart can be edited with "Edit Data".
import { strToU8, zipSync } from 'fflate';
import type { ChartValue } from '../deck/spec.ts';
import { XML_HEADER, escapeXmlAttr } from '../pptx/xml.ts';
import type { ChartSlotDef } from '../template/manifest.ts';

const NS_C = 'http://schemas.openxmlformats.org/drawingml/2006/chart';
const NS_A = 'http://schemas.openxmlformats.org/drawingml/2006/main';
const NS_R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';

export const CHART_CONTENT_TYPE = 'application/vnd.openxmlformats-officedocument.drawingml.chart+xml';
export const CHART_REL_TYPE = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart';
export const WORKBOOK_REL_TYPE = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/package';
export const WORKBOOK_CONTENT_TYPE = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
export const CHART_GRAPHIC_URI = NS_C;

/** The earliest date a zip entry can carry. */
const ZIP_EPOCH = new Date(1980, 0, 1);

const DEFAULT_COLORS = ['4472C4', 'ED7D31', 'A5A5A5', 'FFC000', '5B9BD5', '70AD47'];
const SHEET = 'Sheet1';

const text = (value: string) => value.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
const hex = (value: string) => value.replace(/^#/, '').toUpperCase();

/** Spreadsheet column name for a zero-based index: 0 is A, 26 is AA. */
function columnName(index: number): string {
  let name = '';
  for (let n = index + 1; n > 0; n = Math.floor((n - 1) / 26)) name = String.fromCharCode(65 + ((n - 1) % 26)) + name;
  return name;
}

const fill = (color: string) => `<a:solidFill><a:srgbClr val="${hex(color)}"/></a:solidFill>`;
const line = (color: string, width = 9525) => `<a:ln w="${width}" cap="flat">${fill(color)}<a:round/></a:ln>`;
const noLine = '<a:ln><a:noFill/></a:ln>';

/** The workbook cell range of a series' values, of its name, and of the categories. */
const ranges = (chart: ChartValue) => {
  const last = chart.categories.length + 1;
  return {
    categories: `${SHEET}!$A$2:$A$${last}`,
    name: (series: number) => `${SHEET}!$${columnName(series + 1)}$1`,
    values: (series: number) => `${SHEET}!$${columnName(series + 1)}$2:$${columnName(series + 1)}$${last}`,
  };
};

function stringRef(formula: string, values: string[]): string {
  const points = values.map((value, index) => `<c:pt idx="${index}"><c:v>${text(value)}</c:v></c:pt>`).join('');
  return `<c:strRef><c:f>${text(formula)}</c:f><c:strCache><c:ptCount val="${values.length}"/>${points}</c:strCache></c:strRef>`;
}

function numberRef(formula: string, values: number[], format: string): string {
  const points = values.map((value, index) => `<c:pt idx="${index}"><c:v>${value}</c:v></c:pt>`).join('');
  return `<c:numRef><c:f>${text(formula)}</c:f><c:numCache><c:formatCode>${text(format)}</c:formatCode><c:ptCount val="${values.length}"/>${points}</c:numCache></c:numRef>`;
}

function dataLabels(format: string, show: 'value' | 'none', position?: string): string {
  if (show === 'none') return '<c:dLbls><c:delete val="1"/></c:dLbls>';
  return (
    `<c:dLbls><c:numFmt formatCode="${escapeXmlAttr(format)}" sourceLinked="0"/><c:spPr><a:noFill/>${noLine}</c:spPr>` +
    (position ? `<c:dLblPos val="${position}"/>` : '') +
    `<c:showLegendKey val="0"/><c:showVal val="1"/><c:showCatName val="0"/><c:showSerName val="0"/><c:showPercent val="0"/><c:showBubbleSize val="0"/></c:dLbls>`
  );
}

/** The chart part. `workbookRelId` is the relationship from the chart to its workbook. */
export function buildChartXml(chart: ChartValue, slot: ChartSlotDef, workbookRelId: string, direction: 'rtl' | 'ltr' = 'ltr'): string {
  const colors = slot.colors && slot.colors.length > 0 ? slot.colors : DEFAULT_COLORS;
  const color = (index: number) => colors[index % colors.length] as string;
  const lineColor = slot.lineColor ?? 'D9D9D9';
  const format = chart.format === 'percent' ? '0"%"' : 'General';
  const range = ranges(chart);
  const several = chart.series.length > 1;

  const seriesHead = (index: number) =>
    `<c:idx val="${index}"/><c:order val="${index}"/><c:tx>${stringRef(range.name(index), [chart.series[index]?.name ?? ''])}</c:tx>`;
  const seriesData = (index: number) =>
    `<c:cat>${stringRef(range.categories, chart.categories)}</c:cat>` +
    `<c:val>${numberRef(range.values(index), chart.series[index]?.values ?? [], format)}</c:val>`;

  const axisText = '<c:majorTickMark val="none"/><c:minorTickMark val="none"/><c:tickLblPos val="nextTo"/>';
  const axisLine = `<c:spPr>${line(lineColor)}</c:spPr>`;
  const gridlines = `<c:majorGridlines><c:spPr>${line(lineColor)}</c:spPr></c:majorGridlines>`;
  const [categoryAxisId, valueAxisId] = [50010001, 50010002];
  // Bars are read by their length, so their axis starts at zero; left to itself PowerPoint
  // starts it near the smallest value, which exaggerates differences.
  const zeroBased = chart.type !== 'line' && chart.series.every((series) => series.values.every((value) => value >= 0));
  /** Category and value axes. Horizontal bars list their categories top to bottom, as a table would. */
  const axes = (horizontal: boolean) =>
    `<c:catAx><c:axId val="${categoryAxisId}"/><c:scaling><c:orientation val="${horizontal ? 'maxMin' : 'minMax'}"/></c:scaling>` +
    `<c:delete val="0"/><c:axPos val="${horizontal ? 'l' : 'b'}"/><c:numFmt formatCode="General" sourceLinked="0"/>${axisText}${axisLine}` +
    `<c:crossAx val="${valueAxisId}"/><c:crosses val="autoZero"/><c:auto val="1"/><c:lblAlgn val="ctr"/><c:lblOffset val="100"/><c:noMultiLvlLbl val="0"/></c:catAx>` +
    `<c:valAx><c:axId val="${valueAxisId}"/><c:scaling><c:orientation val="minMax"/>${zeroBased ? '<c:min val="0"/>' : ''}</c:scaling>` +
    `<c:delete val="0"/><c:axPos val="${horizontal ? 'b' : 'l'}"/>${gridlines}<c:numFmt formatCode="${escapeXmlAttr(format)}" sourceLinked="0"/>${axisText}` +
    `<c:spPr>${noLine}</c:spPr><c:crossAx val="${categoryAxisId}"/><c:crosses val="${horizontal ? 'max' : 'autoZero'}"/><c:crossBetween val="between"/></c:valAx>`;
  const axisIds = `<c:axId val="${categoryAxisId}"/><c:axId val="${valueAxisId}"/>`;

  let plot: string;
  if (chart.type === 'pie') {
    const slices = chart.categories
      .map((_, index) => `<c:dPt><c:idx val="${index}"/><c:bubble3D val="0"/><c:spPr>${fill(color(index))}<a:ln w="19050">${fill('FFFFFF')}</a:ln></c:spPr></c:dPt>`)
      .join('');
    plot =
      `<c:pieChart><c:varyColors val="1"/><c:ser>${seriesHead(0)}${slices}${dataLabels(format, 'value', 'outEnd')}${seriesData(0)}</c:ser>` +
      `<c:firstSliceAng val="0"/></c:pieChart>`;
  } else if (chart.type === 'line') {
    const series = chart.series
      .map(
        (_, index) =>
          `<c:ser>${seriesHead(index)}<c:spPr><a:ln w="28575" cap="rnd">${fill(color(index))}<a:round/></a:ln></c:spPr>` +
          `<c:marker><c:symbol val="circle"/><c:size val="6"/><c:spPr>${fill(color(index))}${noLine}</c:spPr></c:marker>` +
          `${seriesData(index)}<c:smooth val="0"/></c:ser>`,
      )
      .join('');
    plot = `<c:lineChart><c:grouping val="standard"/><c:varyColors val="0"/>${series}${dataLabels(format, 'none')}<c:marker val="1"/>${axisIds}</c:lineChart>${axes(false)}`;
  } else {
    const horizontal = chart.type === 'bar';
    const series = chart.series
      .map((_, index) => `<c:ser>${seriesHead(index)}<c:spPr>${fill(color(index))}</c:spPr><c:invertIfNegative val="0"/>${seriesData(index)}</c:ser>`)
      .join('');
    plot =
      `<c:barChart><c:barDir val="${horizontal ? 'bar' : 'col'}"/><c:grouping val="clustered"/><c:varyColors val="0"/>${series}` +
      `${dataLabels(format, 'value', 'outEnd')}<c:gapWidth val="${several ? 80 : 60}"/>${axisIds}</c:barChart>${axes(horizontal)}`;
  }

  const legend = several || chart.type === 'pie' ? '<c:legend><c:legendPos val="b"/><c:overlay val="0"/></c:legend>' : '';
  const style = slot.text ?? {};
  const font = style.font ? `<a:latin typeface="${escapeXmlAttr(style.font)}"/><a:ea typeface="${escapeXmlAttr(style.font)}"/><a:cs typeface="${escapeXmlAttr(style.font)}"/>` : '';
  const defaults =
    // The paragraph direction makes mixed labels such as "רבעון 1" read in the right order.
    `<c:txPr><a:bodyPr/><a:lstStyle/><a:p><a:pPr${direction === 'rtl' ? ' rtl="1"' : ''}><a:defRPr sz="${Math.round((style.size ?? 10) * 100)}"${style.bold ? ' b="1"' : ''}>` +
    `${style.color ? fill(style.color) : ''}${font}</a:defRPr></a:pPr><a:endParaRPr lang="en-US"/></a:p></c:txPr>`;

  return (
    `${XML_HEADER}<c:chartSpace xmlns:c="${NS_C}" xmlns:a="${NS_A}" xmlns:r="${NS_R}">` +
    `<c:date1904 val="0"/><c:roundedCorners val="0"/>` +
    `<c:chart><c:autoTitleDeleted val="1"/><c:plotArea><c:layout/>${plot}<c:spPr><a:noFill/>${noLine}</c:spPr></c:plotArea>${legend}` +
    `<c:plotVisOnly val="1"/><c:dispBlanksAs val="gap"/></c:chart>` +
    `<c:spPr><a:noFill/>${noLine}</c:spPr>${defaults}` +
    `<c:externalData r:id="${escapeXmlAttr(workbookRelId)}"><c:autoUpdate val="0"/></c:externalData>` +
    `</c:chartSpace>`
  );
}

/** The data as an Excel workbook: categories down column A, one series per column from B, names in row 1. */
export function buildChartWorkbook(chart: ChartValue): Uint8Array {
  const stringCell = (reference: string, value: string) => `<c r="${reference}" t="inlineStr"><is><t xml:space="preserve">${text(value)}</t></is></c>`;
  const header = `<row r="1">${chart.series.map((series, index) => stringCell(`${columnName(index + 1)}1`, series.name)).join('')}</row>`;
  const rows = chart.categories
    .map((category, row) => {
      const cells = chart.series.map((series, index) => `<c r="${columnName(index + 1)}${row + 2}"><v>${series.values[row] ?? 0}</v></c>`).join('');
      return `<row r="${row + 2}">${stringCell(`A${row + 2}`, category)}${cells}</row>`;
    })
    .join('');

  const main = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
  const packageRels = 'http://schemas.openxmlformats.org/package/2006/relationships';
  const files: Record<string, string> = {
    '[Content_Types].xml':
      `${XML_HEADER}<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">` +
      `<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>` +
      `<Default Extension="xml" ContentType="application/xml"/>` +
      `<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>` +
      `<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>` +
      `</Types>`,
    '_rels/.rels':
      `${XML_HEADER}<Relationships xmlns="${packageRels}">` +
      `<Relationship Id="rId1" Type="${NS_R}/officeDocument" Target="xl/workbook.xml"/></Relationships>`,
    'xl/workbook.xml':
      `${XML_HEADER}<workbook xmlns="${main}" xmlns:r="${NS_R}"><sheets><sheet name="${SHEET}" sheetId="1" r:id="rId1"/></sheets></workbook>`,
    'xl/_rels/workbook.xml.rels':
      `${XML_HEADER}<Relationships xmlns="${packageRels}">` +
      `<Relationship Id="rId1" Type="${NS_R}/worksheet" Target="worksheets/sheet1.xml"/></Relationships>`,
    'xl/worksheets/sheet1.xml': `${XML_HEADER}<worksheet xmlns="${main}"><sheetData>${header}${rows}</sheetData></worksheet>`,
  };
  // A fixed timestamp, so the same chart always gives the same bytes.
  return zipSync(Object.fromEntries(Object.entries(files).map(([name, content]) => [name, strToU8(content)])), { mtime: ZIP_EPOCH });
}

/** The frame that puts a chart on a slide, in place of the sample picture. Sizes are in EMU. */
export function chartFrameXml(shapeId: string, relId: string, area: { x: number; y: number; cx: number; cy: number }, alt: string): string {
  return (
    `<p:graphicFrame xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="${NS_A}" xmlns:r="${NS_R}">` +
    `<p:nvGraphicFramePr><p:cNvPr id="${escapeXmlAttr(shapeId)}" name="Chart" descr="${escapeXmlAttr(alt)}"/>` +
    `<p:cNvGraphicFramePr><a:graphicFrameLocks noGrp="1"/></p:cNvGraphicFramePr><p:nvPr/></p:nvGraphicFramePr>` +
    `<p:xfrm><a:off x="${Math.round(area.x)}" y="${Math.round(area.y)}"/><a:ext cx="${Math.round(area.cx)}" cy="${Math.round(area.cy)}"/></p:xfrm>` +
    `<a:graphic><a:graphicData uri="${CHART_GRAPHIC_URI}"><c:chart xmlns:c="${NS_C}" r:id="${escapeXmlAttr(relId)}"/></a:graphicData></a:graphic>` +
    `</p:graphicFrame>`
  );
}
