// Slide pictures from the Pagus library (`@pagus-kit/core` and `@pagus-kit/renderer`). slotdeck
// does not depend on it: the developer installs it and passes the two modules in.
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate';
import { paragraphDirection } from '../render/text.ts';
import type { Direction } from '../template/manifest.ts';
import type { PreviewProvider, SlidePreview } from './provider.ts';

/** The part of `@pagus-kit/core` this adapter uses: `import * as core from '@pagus-kit/core'`. */
export interface PagusCore {
  parse(data: ArrayBuffer): Promise<{ slides: readonly unknown[]; slideSize: unknown }>;
}

/** The part of `@pagus-kit/renderer` this adapter uses: `import * as renderer from '@pagus-kit/renderer'`. */
export interface PagusRenderer {
  renderSlide(slide: never, slideSize: never, options?: never): { svg: string; width: number; height: number };
}

export interface PagusOptions {
  /** Scale of the drawing: 1 gives 96 pixels to the inch, so a 13.33 inch slide is 1280 wide. Defaults to 1. */
  scale?: number;
  /** A typeface to fall back on for each of the template's, for a viewer that lacks it: `{ Calibri: 'Carlito' }`. */
  fontSubstitutes?: Record<string, string>;
  /** `@font-face` or `@import` rules to embed, so the picture loads its own fonts. */
  fontCss?: string;
}

const encoder = new TextEncoder();

/**
 * Pagus writes a paragraph's text as HTML and leaves its direction to the browser, which takes
 * a paragraph to run left to right unless told otherwise. A Hebrew or Arabic sentence then gets
 * its full stop on the wrong side, and a line that mixes Hebrew with an address or a number is
 * put in the wrong order. A paragraph is therefore marked as running right to left by the rule
 * slotdeck's renderer writes it by (see `paragraphDirection`): one that starts with a
 * right-to-left letter, and in a deck that reads right to left, one with any right-to-left letter.
 */
function markDirection(svg: string, direction: Direction | undefined): string {
  return svg.replace(/<p ([^>]*)>([\s\S]*?)<\/p>/g, (whole, attributes: string, inner: string) => {
    if (/(^|\s)dir=/.test(attributes)) return whole;
    return paragraphDirection(inner.replace(/<[^>]+>/g, ''), direction) === 'rtl' ? `<p dir="rtl" ${attributes}>${inner}</p>` : whole;
  });
}

/**
 * The file with the spaces at the ends of its runs kept. Pagus trims the text of each run, and a
 * space a run starts or ends with is then lost: a heading set run in (`SlideSpec.headings`) is
 * followed by an em space that opens the run after it, and without it the preview shows the heading
 * glued to its text. A word joiner (U+2060), which draws nothing and no trimming takes, stands
 * beside such a space. A file whose runs have none comes back as it is.
 */
function keepSpaces(pptx: Uint8Array): Uint8Array {
  const parts = unzipSync(pptx);
  let changed = false;
  for (const [name, data] of Object.entries(parts)) {
    if (!/^ppt\/slides\/slide\d+\.xml$/.test(name)) continue;
    const xml = strFromU8(data);
    const kept = xml.replace(/<a:t>(\s+)/g, '<a:t>\u2060$1').replace(/(\s+)<\/a:t>/g, '$1\u2060</a:t>');
    if (kept === xml) continue;
    parts[name] = strToU8(kept);
    changed = true;
  }
  return changed ? zipSync(parts) : pptx;
}

/**
 * A {@link PreviewProvider} that draws slides with Pagus: one SVG for each slide, with the text
 * as real text, so a browser draws it with its own fonts and line breaking. The pictures are for
 * showing in a web page. They are close to PowerPoint and not identical: see docs/preview.md.
 */
export function pagusPreviews(core: PagusCore, renderer: PagusRenderer, options: PagusOptions = {}): PreviewProvider {
  const render = renderer.renderSlide as (slide: unknown, slideSize: unknown, options?: unknown) => { svg: string; width: number; height: number };
  return {
    async renderSlides({ pptx, slides, direction, signal }) {
      signal?.throwIfAborted();
      // The parser takes an ArrayBuffer of exactly the file, which a view into a larger buffer is not.
      const file = keepSpaces(pptx);
      const buffer = file.buffer.slice(file.byteOffset, file.byteOffset + file.byteLength) as ArrayBuffer;
      const presentation = await core.parse(buffer);
      signal?.throwIfAborted();
      const wanted = slides ? new Set(slides) : undefined;
      const pictures: SlidePreview[] = [];
      presentation.slides.forEach((slide, index) => {
        if (wanted && !wanted.has(index + 1)) return;
        const drawn = render(slide, presentation.slideSize, options);
        pictures.push({ slide: index + 1, mimeType: 'image/svg+xml', data: encoder.encode(markDirection(drawn.svg, direction)), width: drawn.width, height: drawn.height });
      });
      return pictures;
    },
  };
}
