// Variants of a pattern that the code lays out by itself. A person or a model names the variant
// (the mirror image, the slide without its picture, the slide without an optional line); where
// each shape then goes is worked out here from the shapes of the slide, so nobody types a box.
import type { PptxPackage } from '../pptx/package.ts';
import { NS, childElement, descendants, parseXml, rootElement } from '../pptx/xml.ts';
import type { Template } from '../render/render.ts';
import { inspectTemplate } from './inspect.ts';
import type { Box, ShapeInfo, SlideInfo, TemplateInfo } from './inspect.ts';
import { scalesWithText } from './manifest.ts';
import type { PatternDef, SlideBox, SlotDef, TemplateManifest, TextSlotDef } from './manifest.ts';
import { boxCapacity, frameFacts, shapeText, slotFromShape } from './measure.ts';

/**
 * The variants `derivePattern` can make:
 * - `mirror`: the picture and what sits beside it change sides
 * - `text-only`: the picture is left out and the text beside it widens over the freed area
 * - `without`: one optional text is left out and what sits under it moves up
 */
export const VARIANTS = ['mirror', 'text-only', 'without'] as const;
export type PatternVariant = (typeof VARIANTS)[number];

export interface VariantRequest {
  variant: PatternVariant;
  /** For `without`: the optional text slot to leave out. */
  slot?: string;
}

/** One variant of a pattern, and whether the template's slide allows it. */
export interface VariantOption {
  variant: PatternVariant;
  /** For `without`: the slot that would be left out. */
  slot?: string;
  available: boolean;
  /** What the variant does, or why it cannot be made. */
  note: string;
}

export interface DerivedVariant {
  /** A copy of the source pattern with the variant's slots. Its id and description are still the source's. */
  pattern: PatternDef;
  /** Ending for the new pattern's id, such as `mirrored`. */
  suffix: string;
  /** A description the planner can choose the variant by. */
  description: string;
  /** What was done, in a sentence or two. */
  summary: string;
}

export interface VariantContext {
  pkg: PptxPackage;
  info: TemplateInfo;
  manifest: TemplateManifest;
}

class Refused extends Error {}
const refuse = (reason: string): never => {
  throw new Refused(reason);
};

const EPS = 0.005;
/** Room left between a picture and a header or footer it is cut back from, as a share of the slide height. */
const GAP = 0.03;
/** The most characters a widened text box should hold on one line. Longer lines are hard to read. */
const LINE = 90;

const round = (value: number) => Math.round(value * 10000) / 10000;
const rounded = (box: Box): SlideBox => ({ x: round(box.x), y: round(box.y), w: round(box.w), h: round(box.h) });
const rightOf = (box: Box) => box.x + box.w;
const bottomOf = (box: Box) => box.y + box.h;
const middleOf = (box: Box) => box.x + box.w / 2;
const shared = (a0: number, a1: number, b0: number, b1: number) => Math.min(a1, b1) - Math.max(a0, b0);
const sharedAcross = (a: Box, b: Box) => shared(a.x, rightOf(a), b.x, rightOf(b));
const sharedDown = (a: Box, b: Box) => shared(a.y, bottomOf(a), b.y, bottomOf(b));
const touches = (a: Box, b: Box) => sharedAcross(a, b) > EPS && sharedDown(a, b) > EPS;
const words = (name: string) => name.replace(/([a-z0-9])([A-Z])/g, '$1 $2').toLowerCase();
const list = (names: string[]) => (names.length <= 1 ? names.join('') : `${names.slice(0, -1).join(', ')} and ${names.at(-1)}`);

/** A shape that sits directly on the slide, with everything the pattern says about it. */
interface Unit {
  shape: ShapeInfo;
  /** Where the pattern shows it: a slot's own box, or the shape as drawn. */
  box: Box;
  /** The slot that points at the shape itself. */
  own?: SlotDef;
  /** Slots that point at shapes inside it, when it is a group. */
  inner: SlotDef[];
  /** Header, logo, page number or a value of the whole deck: never moved. */
  locked: boolean;
  /** Always removed by the pattern, so it is not on the slide. */
  gone: boolean;
  /** An item of the pattern's row, which the renderer lays out itself. */
  inRow: boolean;
}

function unitsOf(pattern: PatternDef, slide: SlideInfo, info: TemplateInfo): Unit[] {
  const frame = frameFacts(info);
  const byId = new Map(slide.shapes.map((shape) => [shape.id, shape]));
  const topOf = (id: string): string => {
    const parent = byId.get(id)?.groupId;
    return parent === undefined ? id : topOf(parent);
  };
  const row = new Set((pattern.repeat?.shapes ?? []).flat().map(topOf));
  return slide.shapes
    .filter((shape) => shape.groupId === undefined)
    .map((shape): Unit => {
      const own = pattern.slots.find((slot) => slot.shapeId === shape.id);
      const inner = pattern.slots.filter((slot) => slot.shapeId !== shape.id && byId.has(slot.shapeId) && topOf(slot.shapeId) === shape.id);
      const placed = own && 'box' in own ? own.box : undefined;
      const unit: Unit = {
        shape,
        box: placed ?? shape.box,
        inner,
        locked:
          frame.isFrame(shape) ||
          (shape.kind === 'text' && frame.repeatedTexts.includes(shapeText(shape))) ||
          [own, ...inner].some((slot) => slot?.kind === 'text' && slot.deckField !== undefined),
        gone: own?.empty === 'drop',
        inRow: row.has(shape.id),
      };
      if (own) unit.own = own;
      return unit;
    });
}

/** Ids of the text shapes whose text starts at the top of its box, so that a taller box leaves the text where it is. */
function topAnchored(pkg: PptxPackage, slide: SlideInfo): Set<string> {
  const ids = new Set<string>();
  for (const shape of descendants(rootElement(parseXml(pkg.readText(slide.part))), NS.p, 'sp')) {
    const id = descendants(shape, NS.p, 'cNvPr')[0]?.getAttribute('id');
    const body = childElement(shape, NS.p, 'txBody');
    const props = body && childElement(body, NS.a, 'bodyPr');
    // A placeholder takes its anchor from the layout, which is not read here.
    const inherited = descendants(shape, NS.p, 'ph').length > 0;
    if (id && (props?.getAttribute('anchor') ?? (inherited ? undefined : 't')) === 't') ids.add(id);
  }
  return ids;
}

const label = (unit: Unit) => {
  const text = unit.shape.kind === 'text' ? shapeText(unit.shape) : '';
  return `shape #${unit.shape.id}${text ? ` ("${text.slice(0, 24)}")` : ''}`;
};

function freeName(pattern: PatternDef, base: string): string {
  const names = new Set(pattern.slots.map((slot) => slot.name));
  for (let n = 1; ; n++) if (!names.has(`${base}${n}`)) return `${base}${n}`;
}

/** Shows a unit in another place: through its own slot, or through a slot that keeps it as designed. */
function place(pattern: PatternDef, unit: Unit, box: Box, what: string): void {
  if (unit.inRow) refuse(`${what} is part of the row of ${pattern.repeat?.item ?? 'item'}s, which the software lays out itself and cannot move`);
  const own = unit.own;
  if (own?.kind === 'table') refuse(`the table "${own.name}" cannot be moved`);
  if (own) {
    (own as { box?: SlideBox }).box = rounded(box);
    return;
  }
  const name = freeName(pattern, 'moved');
  pattern.slots.push(
    unit.shape.kind === 'picture'
      ? { name, kind: 'image', imageKind: 'graphic', shapeId: unit.shape.id, empty: 'keep', box: rounded(box) }
      : { name, kind: 'text', shapeId: unit.shape.id, empty: 'keep', box: rounded(box) },
  );
}

/** Works out a text slot's budget again after its box grew. A budget never falls when the box only grows. */
function rebudget(slot: TextSlotDef, before: Box, after: Box, shape: ShapeInfo, info: TemplateInfo): void {
  if (slot.maxChars === undefined || !(before.w * before.h > 0)) return;
  const size = slot.fontSize ?? slot.size ?? shape.fontSizes?.[0] ?? 14;
  const scaled = Math.round((slot.maxChars * (after.w * after.h)) / (before.w * before.h));
  slot.maxChars = Math.max(slot.maxChars, Math.min(scaled, boxCapacity(after, info, size, slot.maxParagraphs ?? 1, slot.paragraphGap ?? 0)));
}

interface Arrangement {
  picture: Unit;
  slot: SlotDef;
  /** The picture's slot points at a shape inside a group, so the picture itself cannot take a box. */
  grouped: boolean;
  /** Shapes in the picture's own column: a sticker on it, a caption under it. */
  riders: Unit[];
  /** What sits beside the picture. */
  block: Unit[];
  /** Everything that keeps its place: the header, a title above both, a band behind both. */
  stays: Unit[];
  /** The side of the picture the block is on. */
  side: 'left' | 'right';
  blockLeft: number;
  blockRight: number;
}

/** Finds the pattern's one picture and sorts the other shapes into those that ride with it, sit beside it, or stay. */
function arrange(pattern: PatternDef, units: Unit[], purpose: string): Arrangement {
  const pictures = pattern.slots.filter(
    (slot) => ((slot.kind === 'image' && slot.imageKind !== 'icon') || slot.kind === 'chart') && slot.empty !== 'drop' && slot.empty !== 'keep',
  );
  const found = pictures
    .map((slot) => ({ slot, unit: units.find((unit) => unit.own === slot || unit.inner.includes(slot)) }))
    .filter((entry): entry is { slot: SlotDef; unit: Unit } => entry.unit !== undefined && !entry.unit.inRow);
  if (found.length === 0) refuse(`"${pattern.id}" has no photo, graphic or chart of its own, so there is nothing to ${purpose}`);
  if (found.length > 1) refuse(`"${pattern.id}" has ${found.length} pictures (${found.map((entry) => entry.slot.name).join(', ')}); this needs a slide with one picture beside its text`);
  const { slot, unit: picture } = found[0] as { slot: SlotDef; unit: Unit };
  const area = picture.box;

  const riders: Unit[] = [];
  const block: Unit[] = [];
  const stays: Unit[] = [];
  const bands: Unit[] = [];
  for (const unit of units) {
    if (unit === picture || unit.gone) continue;
    if (unit.locked) stays.push(unit);
    else if (sharedAcross(unit.box, area) <= EPS) block.push(unit);
    else if (unit.box.x >= area.x - 0.02 && rightOf(unit.box) <= rightOf(area) + 0.02) riders.push(unit);
    // A title above the picture and its text, or a caption under both, belongs to neither side.
    else if (sharedDown(unit.box, area) <= EPS) stays.push(unit);
    else if (unit.box.x <= area.x + EPS && rightOf(unit.box) >= rightOf(area) - EPS) bands.push(unit);
    // A panel that reaches over the picture's edge belongs to the side it comes from.
    else block.push(unit);
  }
  if (block.length === 0) refuse(`nothing sits beside the picture "${slot.name}", so there is nothing for it to ${purpose === 'leave out' ? 'make room for' : 'change sides with'}`);
  const lefts = block.filter((unit) => middleOf(unit.box) < middleOf(area)).length;
  if (lefts > 0 && lefts < block.length) refuse(`the picture "${slot.name}" has content on both of its sides, so there is no one side to ${purpose === 'leave out' ? 'widen' : 'change'}`);
  const blockLeft = Math.min(...block.map((unit) => unit.box.x));
  const blockRight = Math.max(...block.map((unit) => rightOf(unit.box)));
  // A shape behind the picture that also lies behind the text is a background; a smaller one is the picture's own frame.
  for (const band of bands) (band.box.x <= blockLeft + EPS && rightOf(band.box) >= blockRight - EPS ? stays : riders).push(band);
  return { picture, slot, grouped: picture.own !== slot, riders, block, stays, side: lefts > 0 ? 'left' : 'right', blockLeft, blockRight };
}

function mirror(pattern: PatternDef, units: Unit[]): Omit<DerivedVariant, 'pattern'> {
  const found = arrange(pattern, units, 'change sides with');
  const { picture, block, riders, stays } = found;
  const area = picture.box;
  const bleeds = area.x <= EPS || rightOf(area) >= 1 - EPS;
  // A picture that runs off one side of the slide runs off the other; otherwise the two change places within the room they share.
  const left = bleeds ? 0 : Math.min(area.x, found.blockLeft);
  const right = bleeds ? 1 : Math.max(rightOf(area), found.blockRight);
  const pictureShift = left + right - rightOf(area) - area.x;
  const blockShift = left + right - found.blockRight - found.blockLeft;

  const moved = new Map<Unit, Box>();
  for (const unit of riders) moved.set(unit, { ...unit.box, x: unit.box.x + pictureShift });
  for (const unit of block) moved.set(unit, { ...unit.box, x: unit.box.x + blockShift });
  for (const [unit, box] of moved) {
    const hit = stays.find((stay) => touches(box, stay.box) && !touches(unit.box, stay.box));
    if (hit) refuse(`${label(hit)}, which keeps its place, is where ${label(unit)} would go`);
  }

  // The header and footer stay where they are, so a picture that would cover them is cut back to start under or end above them.
  let target: Box = { ...area, x: area.x + pictureShift };
  let top = target.y;
  let bottom = bottomOf(target);
  const covered = stays.filter((stay) => touches(target, stay.box) && !touches(area, stay.box));
  for (const stay of covered) {
    if (found.grouped) refuse(`${label(stay)}, which keeps its place, is where the picture would go, and a picture inside a group cannot be cut back`);
    if (bottomOf(stay.box) <= 0.3) top = Math.max(top, bottomOf(stay.box) + GAP);
    else if (stay.box.y >= 0.7) bottom = Math.min(bottom, stay.box.y - GAP);
    else refuse(`${label(stay)}, which keeps its place, is where the picture would go`);
  }
  if (bottom - top < area.h * 0.5) refuse('the header and footer leave the picture less than half its height on the other side');
  target = { ...target, y: top, h: bottom - top };

  place(pattern, picture, target, `the picture "${found.slot.name}"`);
  for (const [unit, box] of moved) place(pattern, unit, box, label(unit));

  const now = found.side;
  const other = now === 'left' ? 'right' : 'left';
  const cut = covered.length > 0 ? ` The picture is cut back to clear ${list(covered.map(label))}, which ${covered.length === 1 ? 'keeps its' : 'keep their'} place.` : '';
  return {
    suffix: 'mirrored',
    description: `${pattern.description} Mirrored: here the picture is on the ${now} and the text on the ${other}. Alternate it with "${pattern.id}" for variety.`,
    summary: `the mirror image of "${pattern.id}", with the picture on the ${now} and what sat beside it on the ${other}.${cut}`,
  };
}

function textOnly(pattern: PatternDef, units: Unit[], context: VariantContext): Omit<DerivedVariant, 'pattern'> {
  const found = arrange(pattern, units, 'leave out');
  const { picture, block, riders, side } = found;
  const area = picture.box;
  const shapes = new Map((context.info.slides[pattern.sourceSlide - 1]?.shapes ?? []).map((shape) => [shape.id, shape]));
  const removed = new Set<Unit>();
  const drop = (unit: Unit) => {
    removed.add(unit);
    if (unit.own) unit.own.empty = 'drop';
    else {
      const slot = slotFromShape(unit.shape, context.info, 'decoration', freeName(pattern, 'decoration'));
      if (slot) pattern.slots.push(slot);
    }
  };

  // The picture's slot stays, under its name, as a slot that always removes the shape.
  const at = pattern.slots.indexOf(found.slot);
  const shape = shapes.get(found.slot.shapeId);
  const gone = shape && slotFromShape(shape, context.info, 'decoration', found.slot.name);
  if (gone) pattern.slots.splice(at, 1, gone);
  if (!found.grouped) removed.add(picture);
  else if (picture.inner.length === 1 && !picture.own) drop(picture);
  // What was drawn on the picture goes with it: a sticker, a frame. A text the deck writes stays.
  for (const rider of riders) {
    const on = Math.max(0, sharedAcross(rider.box, area)) * Math.max(0, sharedDown(rider.box, area)) >= 0.5 * rider.box.w * rider.box.h;
    const written = rider.inner.length > 0 || (rider.own !== undefined && rider.own.empty !== 'keep' && rider.own.kind !== 'image');
    if (on && !written) drop(rider);
  }

  // The text grows towards the side the picture was on: up to the picture's far edge, or, for a
  // picture that ran off the slide, up to the margin the text keeps on its own side.
  const far = side === 'left' ? (rightOf(area) >= 1 - EPS ? 1 - found.blockLeft : rightOf(area)) : area.x <= EPS ? 1 - found.blockRight : area.x;
  let columnLeft = side === 'left' ? found.blockLeft : far;
  let columnRight = side === 'left' ? far : found.blockRight;
  const texts = block.filter((unit): unit is Unit & { own: TextSlotDef } => unit.own?.kind === 'text' && unit.own.empty !== 'keep' && unit.own.empty !== 'drop');
  // Lines that run the width of a slide are hard to read, so the column is no wider than a
  // comfortable line of the slot with the most text. It keeps the edge where reading starts,
  // which lines up with the other slides, and gives up room at the far edge.
  const main = texts.slice().sort((a, b) => (b.own.maxChars ?? 0) - (a.own.maxChars ?? 0))[0];
  if (main) {
    const size = (main.own.fontSize ?? main.own.size ?? main.shape.fontSizes?.[0] ?? 14) * (scalesWithText(main.own) ? (context.manifest.textScale ?? 1) : 1);
    const readable = (LINE * size * 0.5 * 12700) / context.info.slideWidth;
    if (columnRight - columnLeft > readable) {
      if (context.manifest.direction === 'rtl') columnLeft = columnRight - readable;
      else columnRight = columnLeft + readable;
    }
  }
  const widened: string[] = [];
  for (const unit of texts) {
    const slot = unit.own;
    const from = unit.box;
    // Every text reaches the column's edge on the side it grows to, so the boxes end on one line.
    let left = side === 'left' ? Math.max(from.x, columnLeft) : columnLeft;
    let right = side === 'left' ? columnRight : Math.min(rightOf(from), columnRight);
    for (const other of units) {
      if (other === unit || other.gone || removed.has(other) || sharedDown(other.box, from) <= EPS) continue;
      if (side === 'left' && other.box.x >= rightOf(from) - EPS) right = Math.min(right, other.box.x - 0.02);
      if (side === 'right' && rightOf(other.box) <= from.x + EPS) left = Math.max(left, rightOf(other.box) + 0.02);
    }
    if (right - left <= from.w + EPS) continue;
    const to = { ...from, x: left, w: right - left };
    rebudget(slot, from, to, unit.shape, context.info);
    slot.box = rounded(to);
    widened.push(`${slot.name}${slot.maxChars !== undefined ? ` (${slot.maxChars} characters)` : ''}`);
  }
  if (widened.length === 0) {
    const inGroup = block.some((unit) => unit.inner.some((slot) => slot.kind === 'text'));
    refuse(inGroup ? 'the text beside the picture is inside a group, and a shape inside a group cannot take a box' : 'no text beside the picture has room to widen');
  }
  return {
    suffix: 'text-only',
    description: `${pattern.description} Without the picture: the text takes the room it leaves.`,
    summary: `"${pattern.id}" without its picture "${found.slot.name}". Widened over the freed area, with budgets to match: ${list(widened)}.`,
  };
}

function without(pattern: PatternDef, units: Unit[], name: string | undefined, context: VariantContext): Omit<DerivedVariant, 'pattern'> {
  const optional = pattern.slots.filter((slot) => slot.kind === 'text' && slot.empty === 'remove' && !slot.deckField).map((slot) => slot.name);
  const hint = optional.length > 0 ? `; "${pattern.id}" has ${list(optional.map((entry) => `"${entry}"`))}` : `; "${pattern.id}" has none`;
  if (!name) refuse(`name the slot to leave out, an optional text${hint}`);
  const slot = pattern.slots.find((candidate) => candidate.name === name) ?? refuse(`"${pattern.id}" has no slot "${name}"`);
  if (slot.kind !== 'text' || slot.deckField || slot.empty !== 'remove') {
    const state = slot.kind !== 'text' ? `is ${slot.kind === 'image' ? 'an image' : `a ${slot.kind}`}` : slot.deckField ? "is one of the deck's own fields" : slot.empty === 'drop' ? 'is always removed already' : slot.empty === 'keep' ? 'is kept as designed' : 'is required';
    refuse(`"${name}" ${state}; only an optional text can be left out${hint}`);
  }
  const unit = units.find((candidate) => candidate.own === slot) ?? refuse(`"${name}" is inside a group, and what sits under a shape inside a group cannot be moved`);
  if (unit.inRow) refuse(`"${name}" is part of the row of ${pattern.repeat?.item ?? 'item'}s`);
  const hole = unit.box;
  const anchored = topAnchored(context.pkg, context.info.slides[pattern.sourceSlide - 1] as SlideInfo);
  slot.empty = 'drop';
  unit.gone = true;

  const sameColumn = (other: Unit) => sharedAcross(other.box, hole) >= 0.5 * Math.min(other.box.w, hole.w);
  const candidates = units.filter((other) => other !== unit && !other.gone && !other.locked && sameColumn(other));
  const under = candidates.filter((other) => other.box.y >= bottomOf(hole) - 0.01);
  const done: string[] = [];
  const grow = (target: Unit, to: Box): boolean => {
    const text = target.own;
    if (!text || text.kind !== 'text' || text.empty === 'keep' || !anchored.has(target.shape.id)) return false;
    rebudget(text, target.box, to, target.shape, context.info);
    return true;
  };

  if (under.length > 0) {
    const shift = Math.min(...under.map((other) => other.box.y)) - hole.y;
    for (const other of under) {
      let to: Box = { ...other.box, y: other.box.y - shift };
      // The lowest text of a column keeps its lower edge, so the room that was gained is room to write in.
      const lowest = !under.some((below) => below !== other && below.box.y >= bottomOf(other.box) - 0.01 && sharedAcross(below.box, other.box) > EPS);
      const taller = { ...to, h: to.h + shift };
      const grew = lowest && grow(other, taller);
      if (grew) to = taller;
      const hit = units.find((stay) => stay !== other && stay !== unit && !stay.gone && !under.includes(stay) && touches(to, stay.box) && !touches(other.box, stay.box));
      if (hit) refuse(`${label(hit)}, which keeps its place, is where ${label(other)} would go`);
      place(pattern, other, to, other.own ? `"${other.own.name}"` : label(other));
      const text = other.own?.kind === 'text' ? other.own : undefined;
      done.push(`${other.own?.name ?? label(other)} moves up${grew ? ` and has more room${text?.maxChars !== undefined ? ` (${text.maxChars} characters)` : ''}` : ''}`);
    }
  } else {
    // Nothing sits under it: the text above it gets the room.
    const above = candidates.filter((other) => bottomOf(other.box) <= hole.y + 0.01).sort((a, b) => bottomOf(b.box) - bottomOf(a.box))[0];
    const to = above && { ...above.box, h: bottomOf(hole) - above.box.y };
    if (above && to && grow(above, to)) {
      place(pattern, above, to, `"${above.own?.name}"`);
      const text = above.own?.kind === 'text' ? above.own : undefined;
      done.push(`${above.own?.name} has the room it leaves${text?.maxChars !== undefined ? ` (${text.maxChars} characters)` : ''}`);
    }
  }
  return {
    suffix: `no-${words(name as string).replace(/[^a-z0-9]+/g, '-')}`,
    description: `${pattern.description} Without the ${words(name as string)}${done.length > 0 ? ', which leaves more room for the rest' : ''}.`,
    summary: `"${pattern.id}" without "${name}"${done.length > 0 ? `: ${list(done)}.` : '. Nothing else needed to move.'}`,
  };
}

/**
 * Works out a variant of a pattern from the shapes of its slide. The source is not touched. A
 * variant that the slide does not allow comes back as `refused`, with the reason.
 */
export function deriveVariant(source: PatternDef, request: VariantRequest, context: VariantContext): DerivedVariant | { refused: string } {
  const slide = context.info.slides[source.sourceSlide - 1];
  if (!slide) return { refused: `"${source.id}" points at slide ${source.sourceSlide}, which the template does not have` };
  const pattern = structuredClone(source);
  const units = unitsOf(pattern, slide, context.info);
  try {
    const made =
      request.variant === 'mirror'
        ? mirror(pattern, units)
        : request.variant === 'text-only'
          ? textOnly(pattern, units, context)
          : request.variant === 'without'
            ? without(pattern, units, request.slot, context)
            : refuse(`"${String(request.variant)}" is not a variant; the variants are ${VARIANTS.join(', ')}`);
    return { pattern, ...made };
  } catch (error) {
    if (!(error instanceof Refused)) throw error;
    return { refused: error.message };
  }
}

/**
 * Lists the variants of a pattern with whether each can be made: the mirror image, the text-only
 * version, and the pattern without each of its optional texts. For a studio screen that offers
 * them as buttons; `changeManifest` makes one with the `derivePattern` operation.
 */
export function patternVariants(template: Template, pattern: string): VariantOption[] {
  const source = template.manifest.patterns.find((candidate) => candidate.id === pattern);
  if (!source) throw new Error(`there is no pattern "${pattern}"`);
  const context: VariantContext = { pkg: template.pkg, info: inspectTemplate(template.pkg), manifest: template.manifest };
  const requests: VariantRequest[] = [
    { variant: 'mirror' },
    { variant: 'text-only' },
    ...source.slots.filter((slot) => slot.kind === 'text' && slot.empty === 'remove' && !slot.deckField).map((slot): VariantRequest => ({ variant: 'without', slot: slot.name })),
  ];
  return requests.map((request): VariantOption => {
    const made = deriveVariant(source, request, context);
    const option: VariantOption = 'refused' in made ? { variant: request.variant, available: false, note: made.refused } : { variant: request.variant, available: true, note: made.summary };
    if (request.slot !== undefined) option.slot = request.slot;
    return option;
  });
}
