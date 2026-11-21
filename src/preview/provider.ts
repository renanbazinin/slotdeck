// Pictures of slides, made by a renderer the developer plugs in. slotdeck draws nothing itself.
import type { Direction } from '../template/manifest.ts';

/** A picture of one slide of a pptx. */
export interface SlidePreview {
  /** 1-based slide number. */
  slide: number;
  /** `image/svg+xml` for a drawing a browser shows, `image/png` (or another bitmap type) for pixels. */
  mimeType: string;
  /** The picture: SVG as UTF-8 text, or the bitmap's bytes. */
  data: Uint8Array;
  /** Size in pixels, when the renderer states it. */
  width?: number;
  height?: number;
}

export interface PreviewRequest {
  /** The presentation to draw, e.g. the result of `renderDeck`. */
  pptx: Uint8Array;
  /** 1-based numbers of the slides wanted. All of them when left out. */
  slides?: number[];
  /**
   * The direction the deck reads, for a renderer that has to give a paragraph its direction
   * itself: in a right-to-left deck a paragraph with any right-to-left letter reads right to
   * left, as slotdeck writes it. slotdeck's own calls fill it.
   */
  direction?: Direction;
  signal?: AbortSignal;
}

/**
 * The one thing slotdeck needs from a slide renderer: pictures of the slides of a pptx.
 * Implement this over any renderer: a library, a service that runs PowerPoint or LibreOffice,
 * a headless browser. {@link pagusPreviews} does it for the Pagus library.
 */
export interface PreviewProvider {
  /** Resolves with one picture for each slide asked for, in slide order. */
  renderSlides(request: PreviewRequest): Promise<SlidePreview[]>;
}

/** The file extension for a preview's type, for naming it in a store or on disk. */
export function previewExtension(preview: Pick<SlidePreview, 'mimeType'>): string {
  const type = preview.mimeType.toLowerCase().split(';')[0]?.trim() ?? '';
  if (type === 'image/svg+xml') return 'svg';
  if (type === 'image/jpeg') return 'jpg';
  return /^image\/([a-z0-9]+)$/.exec(type)?.[1] ?? 'bin';
}
