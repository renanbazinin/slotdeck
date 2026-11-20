// The headings of a person's document that stand in a slide's text as paragraphs of their own
// (`SlideSpec.headings`): which paragraphs of a text they are, and how each is set.

/** A line that is an item of a list: a bullet or a number, then a space. */
export const ITEM = /^\s*(?:[-*•‣◦▪–·]|[o§](?=\t)|\(?\d{1,3}[.)]|\(?[a-z][.)]|\(?[א-ת][.)])\s+/u;

/**
 * Which paragraphs of a text are the slide's headings (`SlideSpec.headings`), by position. The
 * entries are matched in order, each to the first paragraph after the one before it that says
 * exactly that, spaces at the ends aside; an entry that matches nothing is passed over. A text of
 * one paragraph has none, and neither has the last paragraph of a text: a heading stands over
 * text. Anything but a list of strings names none.
 */
export function headingPlaces(paragraphs: readonly string[], headings: unknown): number[] {
  if (!Array.isArray(headings) || !headings.every((entry) => typeof entry === 'string') || paragraphs.length < 2) return [];
  const places: number[] = [];
  let from = 0;
  for (const entry of headings as string[]) {
    const wanted = entry.trim();
    if (!wanted) continue;
    for (let at = from; at < paragraphs.length - 1; at++) {
      if ((paragraphs[at] as string).trim() !== wanted) continue;
      places.push(at);
      from = at + 1;
      break;
    }
  }
  return places;
}

/**
 * Whether a heading is set at the head of the paragraph under it: not over an item of a list or a
 * row of a table, which keep the line they start with, the heading then on a line of its own.
 */
export const leadsInto = (next: string): boolean => !ITEM.test(next) && !next.includes(' | ');
