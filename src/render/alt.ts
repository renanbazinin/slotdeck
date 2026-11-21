// The alternative text of what the renderer draws itself: what a screen reader says of a chart
// or a table, in the deck's language.
import type { ChartType, ChartValue, TableValue } from '../deck/spec.ts';
import type { Direction } from '../template/manifest.ts';

/** The words of an alternative text, in English and in Hebrew. */
const WORDS = {
  en: {
    chart: { column: 'Column chart', bar: 'Bar chart', line: 'Line chart', pie: 'Pie chart' } satisfies Record<ChartType, string>,
    of: (kind: string, name: string) => `${kind} of ${name}`,
    missing: 'no value',
    table: 'Table',
    rows: (count: number) => (count === 1 ? '1 row' : `${count} rows`),
  },
  he: {
    chart: { column: 'תרשים עמודות', bar: 'תרשים פסים', line: 'תרשים קו', pie: 'תרשים עוגה' } satisfies Record<ChartType, string>,
    of: (kind: string, name: string) => `${kind} של ${name}`,
    missing: 'אין נתון',
    table: 'טבלה',
    rows: (count: number) => (count === 1 ? 'שורה אחת' : `${count} שורות`),
  },
};

/** The words for a deck: Hebrew for a deck in Hebrew that reads right to left, and English for any other. */
const wordsFor = (language: string | undefined, direction: Direction | undefined) => (direction === 'rtl' && /^(he|iw)([-_]|$)/i.test((language ?? '').trim()) ? WORDS.he : WORDS.en);

/**
 * What a chart shows, for a screen reader: its kind, and every series with its value at each
 * category, as "Column chart of Adoptions: Q1 62, Q2 74, Q3 95". A chart of several series names
 * each before its values. A value that is not known is said so.
 */
export function chartAlt(chart: ChartValue, language?: string, direction?: Direction): string {
  const words = wordsFor(language, direction);
  const kind = words.chart[chart.type] ?? words.chart.column;
  const unit = chart.format === 'percent' ? '%' : '';
  const valuesOf = (values: ReadonlyArray<number | null>) =>
    chart.categories.map((category, index) => {
      const value = values[index];
      return `${category} ${value === null || value === undefined ? words.missing : `${value}${unit}`}`;
    });
  const [only] = chart.series;
  if (chart.series.length === 1 && only) return `${only.name ? words.of(kind, only.name) : kind}: ${valuesOf(only.values).join(', ')}`;
  return [kind, ...chart.series.map((series) => `${series.name ? `${series.name}: ` : ''}${valuesOf(series.values).join(', ')}`)].join('. ');
}

/** What a table holds, for a screen reader: its headings and how many rows it has, as "Table: Name, Region, Sales; 5 rows". */
export function tableAlt(table: TableValue, language?: string, direction?: Direction): string {
  const words = wordsFor(language, direction);
  const header = table.header.filter((cell) => cell.trim() !== '').join(', ');
  return `${words.table}${header ? `: ${header}` : ''}; ${words.rows(table.rows.length)}`;
}
