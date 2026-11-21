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
      if (node.localName !== 'par' || node.namespaceURI !== NS.p) continue;
      nearest ??= node;
      if (childElement(node, NS.p, 'cTn')?.hasAttribute('presetClass')) {
        step = node;
        break;
      }
    }
    const gone = step ?? nearest;
    if (gone?.parentNode) removeNode(gone);
  }
  // Steps left with nothing in them, from the inside out.
  for (let removed = true; removed; ) {
    removed = false;
    for (const step of descendants(timing, NS.p, 'par')) {
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
