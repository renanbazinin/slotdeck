// What a slide cloned from the template brings along that is the template's own sample content and
// not the deck's: tags that add-ins keep on it, and animations of shapes or paragraphs that are gone.
import { NS, childElement, childElements, descendants, path, removeNode } from '../pptx/xml.ts';
import type { Element } from '../pptx/xml.ts';

/**
 * Takes the template's leftovers off a filled slide: the data add-ins keep on the slide and its
 * shapes (`p:custDataLst`), and the animations of shapes that are no longer on it, or of paragraphs
 * their shape no longer has. PowerPoint asks to repair a file whose animation names a shape that
 * is not there. An animation left with nothing to animate is taken out whole.
 */
export function dropLeftovers(root: Element): void {
  for (const data of descendants(root, NS.p, 'custDataLst')) removeNode(data);
  const timing = childElement(root, NS.p, 'timing');
  if (!timing) return;
  const shapes = new Map<string, Element>();
  for (const props of descendants(root, NS.p, 'cNvPr')) {
    const id = props.getAttribute('id');
    const shape = props.parentNode?.parentNode;
    if (id && shape && shape.nodeType === 1) shapes.set(id, shape as Element);
  }
  /** Whether an animation's target is on the slide: its shape, and the paragraphs it names. */
  const there = (target: Element): boolean => {
    const shape = shapes.get(target.getAttribute('spid') ?? '');
    if (!shape) return false;
    const range = path(target, [NS.p, 'txEl'], [NS.p, 'pRg']);
    if (!range) return true;
    const paragraphs = descendants(shape, NS.a, 'p').length;
    return Number(range.getAttribute('end') ?? range.getAttribute('st')) < paragraphs;
  };
  // The effect a missing target is in: the nearest step of the timeline that names a preset effect, or else the nearest step.
  for (const target of descendants(timing, NS.p, 'spTgt').filter((entry) => !there(entry))) {
    let step: Element | undefined;
    let nearest: Element | undefined;
    for (let node = target.parentNode as Element | null; node && node !== timing; node = node.parentNode as Element | null) {
      if (node.namespaceURI !== NS.p) continue;
      // A shape that starts a sequence when clicked, or that a video or a sound plays in: that sequence or that node goes.
      if (node.localName === 'seq' || node.localName === 'video' || node.localName === 'audio') {
        step = node;
        break;
      }
      if (node.localName !== 'par') continue;
      nearest ??= node;
      if (childElement(node, NS.p, 'cTn')?.hasAttribute('presetClass')) {
        step = node;
        break;
      }
    }
    const gone = step ?? nearest;
    if (gone?.parentNode) removeNode(gone);
  }
  // Steps and sequences left with nothing in them, from the inside out: a list of steps may not be
  // empty, and a sequence of clicks whose shapes are all gone is one.
  for (let removed = true; removed; ) {
    removed = false;
    for (const step of [...descendants(timing, NS.p, 'par'), ...descendants(timing, NS.p, 'seq')]) {
      const children = path(step, [NS.p, 'cTn'], [NS.p, 'childTnLst']);
      if (children && childElements(children).length === 0 && step.parentNode) {
        removeNode(step);
        removed = true;
      }
    }
  }
  // How each animated shape is built (whole, or paragraph by paragraph): for the shapes still animated.
  const animated = new Set(descendants(timing, NS.p, 'spTgt').map((target) => target.getAttribute('spid') ?? ''));
  const list = childElement(timing, NS.p, 'bldLst');
  if (list) {
    for (const build of childElements(list)) if (!animated.has(build.getAttribute('spid') ?? '')) removeNode(build);
    if (childElements(list).length === 0) removeNode(list);
  }
  if (descendants(timing, NS.p, 'spTgt').length === 0) removeNode(timing);
}

const NS_DIAGRAM = 'http://schemas.openxmlformats.org/drawingml/2006/diagram';
/**
 * Kinds of relationship of a slide that the template's sample may leave unused, by the last part of
 * their type: pictures, links, comments, add-in tags, charts, embedded files, sound and film, the
 * parts of a diagram, another slide, notes. Unused, they go with their parts; a kind not named here
 * is kept, as it was before 0.7.0.
 */
const LEFT_BEHIND = new Set(['image', 'hyperlink', 'comments', 'tags', 'chart', 'oleObject', 'package', 'audio', 'video', 'media', 'diagramData', 'diagramLayout', 'diagramQuickStyle', 'diagramColors', 'slide', 'notesSlide']);

/**
 * Whether a slide keeps a relationship that its XML does not name with an `r:` attribute: one a
 * part it shows needs all the same. A diagram's drawing is named in the diagram's data, not in the
 * slide; the drawing of an embedded object or a control is found by its kind. A relationship of a
 * kind that a template's sample may leave behind (`LEFT_BEHIND`) is not kept by this: unused, it goes.
 * `free`: the slide is a free slide, which keeps nothing of its base slide but what it shows.
 */
export function keptUnnamed(type: string, root: Element, free: boolean): boolean {
  const kind = type.slice(type.lastIndexOf('/') + 1);
  if (kind === 'diagramDrawing') return descendants(root, NS_DIAGRAM, 'relIds').length > 0;
  if (kind === 'vmlDrawing') return descendants(root, NS.p, 'oleObj').length > 0 || descendants(root, NS.p, 'control').length > 0;
  return !free && !LEFT_BEHIND.has(kind);
}
