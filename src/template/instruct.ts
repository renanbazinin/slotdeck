// Changes a manifest as a person asks in words. The model reads the manifest next to the
// measured shapes of each slide and answers with operations for `changeManifest`; it writes no
// manifest itself, so whatever it answers goes through the same checks as a change made by hand.
import type { JsonSchema, LlmProvider } from '../llm/provider.ts';
import type { Template } from '../render/render.ts';
import { changeManifest, describeSlot } from './change.ts';
import type { ChangeResult, ManifestChange } from './change.ts';
import { VARIANTS } from './derive.ts';
import { inspectTemplate } from './inspect.ts';
import { templateTextSetting } from './manifest.ts';
import type { TableTextStyle } from './manifest.ts';
import { ROLES, describeShape, frameFacts } from './measure.ts';
import type { Role } from './measure.ts';

export interface ManifestEditOptions {
  llm: LlmProvider;
  signal?: AbortSignal;
}

export interface ManifestEditResult extends ChangeResult {
  /** What the model has to say to the person: what it could not do, or what it needs to know. Often empty. */
  note: string;
}

const OPS = ['setSlot', 'addSlot', 'removeSlot', 'setRepeat', 'removeRepeat', 'setPattern', 'copyPattern', 'derivePattern', 'removePattern', 'addPattern', 'setManifest'] as const;

/** What an operation's values may set. A table's text styles are spelled out so that every value is one word or number. */
const PROPERTIES = [
  'name',
  'id',
  'description',
  'figures',
  'role',
  'optional',
  'empty',
  'maxChars',
  'maxParagraphs',
  'deckField',
  'fontSize',
  'imageKind',
  'maxColumns',
  'maxRows',
  'maxCellChars',
  'headerSize',
  'headerBold',
  'headerColor',
  'headerFont',
  'bodySize',
  'bodyBold',
  'bodyColor',
  'bodyFont',
  'maxCategories',
  'maxSeries',
  'min',
  'max',
  'slide',
  'language',
  'direction',
  'alignOnDirectionChange',
  'textScale',
  'characterWidth',
  'leastTextSize',
  'variant',
  'widen',
  'accent',
  'accent2',
  'keepColors',
  'freeSlides',
  'freeBase',
] as const;

const SYSTEM = [
  'You maintain the manifest of a presentation template. The manifest says which slides of a pptx are reusable layouts ("patterns") and which shapes on them are slots that new content replaces.',
  'A person asks for a change in their own words. You answer with the smallest list of operations that does what they asked, and nothing they did not ask for.',
  'You never write positions, sizes or budgets the person did not state: the software measures the shapes and works those out.',
].join(' ');

const GUIDE = [
  'Every operation has "op", "pattern", "shape", "slot" and "values". Give "" for a pattern, shape or slot the operation does not use. List in "values" only the properties that change, each as { property, value } with the value written as text ("6", "true", "drop").',
  'Operations:',
  '- setSlot (pattern, slot): change a slot. Values: empty ("error" makes it required, "remove" optional, "keep" always leaves the template\'s own content, "drop" always removes the shape); maxChars, maxParagraphs (limits of a text); fontSize (a text size in points); deckField (ties a text to one value for the whole deck); name (renames the slot); imageKind ("photo", "icon" or "graphic"). For a table: maxColumns, maxRows, maxCellChars, headerSize, headerBold, headerColor (hex), headerFont, bodySize, bodyBold, bodyColor, bodyFont. For a chart: maxCategories, maxSeries.',
  '- addSlot (pattern, shape): make a shape of the slide a slot, or give the shape of an existing slot another role. Values: role, and optionally name and optional ("true"). Roles: title, subtitle, text, list (one paragraph per item), number (a large figure), caption, photo, icon, graphic (a supplied chart or diagram picture), chart (a native editable chart in place of a picture), table, decoration (removed from every slide made with the pattern), fixed (stays exactly as designed).',
  '- removeSlot (pattern, slot): the shape stays as the template shows it.',
  '- setRepeat (pattern): how many items the row or grid of a pattern takes. Values: min, max, and widen (a factor such as "1.5": with fewer items than the slide draws, each item may grow up to that many times as wide, so the row is filled and does not float in the middle; "1" keeps the drawn width).',
  '- removeRepeat (pattern): the items stop being a row whose number can change.',
  '- setPattern (pattern): Values: id (renames the pattern), description, figures ("true" offers it only when a brief has numbers).',
  '- copyPattern (pattern): a second pattern from the same slide, to turn into a variant by hand with further operations. Values: id of the new pattern, and optionally description.',
  '- derivePattern (pattern): a second pattern from the same slide that the software lays out itself. Values: variant, and optionally id and description. Variants: "mirror" (the picture and the text beside it change sides), "text-only" (the picture is left out and the text widens over the room it leaves), "without" (leaves out the optional text named in "slot"; what sits under it moves up). The software works out every position and budget, and refuses a variant the slide does not allow. In a description, say what the layout holds and where the picture is.',
  '- removePattern (pattern).',
  '- addPattern: make a slide without a pattern into one. Values: slide, id, and optionally description.',
  '- setManifest: settings of the whole template. Values: name, language, direction ("rtl" or "ltr"), alignOnDirectionChange ("flip" or "keep"), textScale (a factor on the body text of every deck, such as "1.3" when the template\'s body text is too small; "1" for the designed size), characterWidth (how wide PowerPoint sets a character of the template\'s text, as a share of its size, such as "0.43"; only a number the person gives, which they measured; "" has the software estimate it), leastTextSize (the least size in points of a long text a person asked to have whole, set dense on a slide the software lays out, such as "10"; other text is not held to it; "" for the size the slide\'s width allows), accent and accent2 (the template\'s main and second colour as six hex digits such as "0E7C7B", which a deck that asks for other colours replaces; "" has the software find them), keepColors ("true" keeps every deck in the template\'s own colours whatever it asks for; "false" lets decks ask again), freeSlides ("false" stops slides that the software lays out itself being planned; "true" allows them), freeBase (the id of the pattern whose frame such slides stand on; "" has the software choose).',
  'A change to one item of a row (card1Icon) is made to that slot of every item, so give it once.',
  "When a size, a table's columns or a row count changes, do not also set a character budget: the software measures the new one.",
  'Operations run in order, so a pattern made by copyPattern can be named by the operations after it.',
  'A person may speak of a slide by its number. The list shows which pattern each slide is.',
  'If something asked for cannot be done with these operations, leave it out and say so in "note", in the language the person wrote in. Otherwise "note" is "".',
].join('\n');

// Models that fill a schema write its properties in alphabetical order. The names are chosen so
// that this order is also the order of thought: the operation, what it is about, then its values.
// Every property is required, so the model never has to decide whether to skip one.
// The arrays carry no length limits: Gemini refuses a schema whose nested limits multiply into too many states.
const SCHEMA: JsonSchema = {
  type: 'object',
  properties: {
    changes: {
      type: 'array',
      items: {
        type: 'object',
        properties: {
          op: { type: 'string', enum: [...OPS] },
          pattern: { type: 'string', description: 'id of the pattern, or ""' },
          shape: { type: 'string', description: 'shape id without the #, or ""' },
          slot: { type: 'string', description: 'name of the slot, or ""' },
          values: {
            type: 'array',
            items: {
              type: 'object',
              properties: { property: { type: 'string', enum: [...PROPERTIES] }, value: { type: 'string' } },
              required: ['property', 'value'],
            },
          },
        },
        required: ['op', 'pattern', 'shape', 'slot', 'values'],
      },
    },
    note: { type: 'string' },
  },
  required: ['changes', 'note'],
};

/**
 * The manifest as text: every pattern with its slots and their limits, next to the measured
 * shapes of its slide that are not slots. This is what the model reads; an app can show it to
 * its own chat model too.
 */
export function describeManifest(template: Template): string {
  const { manifest } = template;
  const info = inspectTemplate(template.pkg);
  const frame = frameFacts(info);
  const used = new Set(manifest.patterns.map((pattern) => pattern.sourceSlide));
  const unused = info.slides.map((slide) => slide.number).filter((number) => !used.has(number));

  const blocks = manifest.patterns.map((pattern) => {
    const slide = info.slides[pattern.sourceSlide - 1];
    const shapes = new Map((slide?.shapes ?? []).map((shape) => [shape.id, shape]));
    const slotted = new Set(pattern.slots.map((slot) => slot.shapeId));
    const lines = [`Pattern "${pattern.id}", slide ${pattern.sourceSlide}${pattern.figures ? ', offered only for briefs with figures' : ''}. ${pattern.description}`];
    const repeat = pattern.repeat;
    if (repeat) {
      const drawn = repeat.shapes.length;
      lines.push(`  ${repeat.columns ? `grid of ${drawn} ${repeat.item}s, ${repeat.columns} to a line` : `row of ${drawn} ${repeat.item}s`} (${repeat.slots.join(', ')}); takes ${(repeat.min ?? drawn) === (repeat.max ?? drawn) ? `exactly ${repeat.min ?? drawn}` : `${repeat.min ?? drawn} to ${repeat.max ?? drawn}`}${repeat.widen && repeat.widen > 1 ? `; fewer on a line grow up to ${repeat.widen} times as wide` : ''}`);
    }
    lines.push('  slots:');
    for (const slot of pattern.slots) {
      const shape = shapes.get(slot.shapeId);
      lines.push(`    ${slot.name}: ${describeSlot(slot)} | ${shape ? describeShape(shape, info) : `#${slot.shapeId}, not on the slide`}`);
    }
    const others = (slide?.shapes ?? []).filter((shape) => !slotted.has(shape.id) && shape.kind !== 'group');
    if (others.length > 0) {
      lines.push('  other shapes on the slide, which stay as designed:');
      for (const shape of others) lines.push(`    ${describeShape(shape, info)}${frame.isFrame(shape) ? ' (repeats across slides)' : ''}`);
    }
    return lines.join('\n');
  });

  return [
    `Template "${manifest.name}"${manifest.language ? `, written in ${manifest.language}` : ''}, reads ${manifest.direction === 'rtl' ? 'right to left' : 'left to right'}; ${(manifest.alignOnDirectionChange ?? 'flip') === 'flip' ? 'mirrors' : 'keeps'} alignment in decks that read the other way${manifest.textScale && manifest.textScale !== 1 ? `; body text is shown at ${manifest.textScale} times its designed size` : ''}${templateTextSetting(manifest, 'characterWidth') !== undefined ? `; a character of its text is counted ${manifest.characterWidth} of its size wide` : ''}${templateTextSetting(manifest, 'leastTextSize') !== undefined ? `; a long text a person asked to have whole, set dense on a slide the software lays out, is set no smaller than ${templateTextSetting(manifest, 'leastTextSize')} pt` : ''}${manifest.colors === false ? '; every deck keeps its colours' : manifest.colors ? `; its ${[manifest.colors.accent ? `main colour is ${manifest.colors.accent}` : '', manifest.colors.accent2 ? `second colour is ${manifest.colors.accent2}` : ''].filter(Boolean).join(' and its ')}` : ''}${manifest.free === false ? '; no slides are laid out by the software' : manifest.free && manifest.free.base ? `; slides the software lays out stand on the frame of "${manifest.free.base}"` : ''}. Positions and sizes are percentages of the slide.`,
    ...blocks,
    unused.length > 0 ? `Slides without a pattern: ${unused.join(', ')}.` : '',
  ]
    .filter(Boolean)
    .join('\n\n');
}

const text = (value: unknown): string | undefined => (typeof value === 'string' && value.trim() !== '' ? value.trim() : undefined);
const oneOf = <T extends string>(value: unknown, options: readonly T[]): T | undefined => options.find((option) => option === value);

/** Turns one answered operation into a change, reading only the values that operation uses. */
function toChange(raw: Record<string, unknown>): ManifestChange | undefined {
  const op = oneOf(raw.op, OPS);
  const pattern = text(raw.pattern) ?? '';
  const slot = text(raw.slot) ?? '';
  const shape = (text(raw.shape) ?? '').replace(/^#/, '');
  const values = new Map<string, string>();
  for (const entry of Array.isArray(raw.values) ? raw.values : []) {
    const property = text((entry as { property?: unknown } | null)?.property);
    const value = (entry as { value?: unknown } | null)?.value;
    if (property && (typeof value === 'string' || typeof value === 'number' || typeof value === 'boolean')) values.set(property, String(value).trim());
  }
  const str = (property: string) => text(values.get(property));
  const num = (property: string) => {
    const value = Number(values.get(property));
    return values.has(property) && Number.isFinite(value) && value > 0 ? value : undefined;
  };
  const count = (property: string) => {
    const value = num(property);
    return value === undefined ? undefined : Math.round(value);
  };
  const bool = (property: string) => {
    const value = values.get(property)?.toLowerCase();
    return value === 'true' ? true : value === 'false' ? false : undefined;
  };
  const choice = <T extends string>(property: string, options: readonly T[]) => oneOf(values.get(property), options);
  const style = (prefix: 'header' | 'body'): TableTextStyle | undefined => {
    const result: TableTextStyle = {};
    const size = num(`${prefix}Size`);
    if (size !== undefined) result.size = size;
    const bold = bool(`${prefix}Bold`);
    if (bold !== undefined) result.bold = bold;
    const color = str(`${prefix}Color`)?.replace(/^#/, '').toUpperCase();
    if (color && /^[0-9A-F]{6}$/.test(color)) result.color = color;
    const font = str(`${prefix}Font`);
    if (font) result.font = font;
    return Object.keys(result).length > 0 ? result : undefined;
  };
  // Colours and free slides are several properties that make one setting each.
  /** A number, or null where the value is empty: the setting goes back to the software's own. */
  const cleared = (property: string): number | null | undefined => (!values.has(property) ? undefined : (values.get(property) ?? '').trim() === '' ? null : Number(values.get(property)));
  const colorsAsked = (): { accent?: string | null; accent2?: string | null } | false | null | undefined => {
    const keep = bool('keepColors');
    if (keep === true) return false;
    const named: { accent?: string | null; accent2?: string | null } = {};
    for (const name of ['accent', 'accent2'] as const) {
      if (values.has(name)) named[name] = (values.get(name) ?? '').replace(/^#/, '').toUpperCase() || null;
    }
    return Object.keys(named).length > 0 ? named : keep === false ? null : undefined;
  };
  const freeAsked = (): boolean | { base: string | null } | undefined => {
    const offered = bool('freeSlides');
    if (offered === false) return false;
    if (values.has('freeBase')) return { base: values.get('freeBase') || null };
    return offered;
  };
  const put = <T extends object>(target: T, entries: Record<string, unknown>): T => {
    for (const [key, value] of Object.entries(entries)) if (value !== undefined) (target as Record<string, unknown>)[key] = value;
    return target;
  };

  switch (op) {
    case 'setSlot':
      return put({ op, pattern, slot } as ManifestChange, {
        name: str('name'),
        description: str('description'),
        empty: choice('empty', ['error', 'remove', 'keep', 'drop'] as const),
        maxChars: count('maxChars'),
        maxParagraphs: count('maxParagraphs'),
        deckField: str('deckField'),
        fontSize: num('fontSize'),
        imageKind: choice('imageKind', ['photo', 'icon', 'graphic'] as const),
        maxColumns: count('maxColumns'),
        maxRows: count('maxRows'),
        maxCellChars: count('maxCellChars'),
        headerText: style('header'),
        bodyText: style('body'),
        maxCategories: count('maxCategories'),
        maxSeries: count('maxSeries'),
      });
    case 'addSlot':
      return put({ op, pattern, shape, role: (choice('role', ROLES) ?? values.get('role') ?? '') as Role } as ManifestChange, {
        name: str('name'),
        optional: bool('optional') === true ? true : undefined,
      });
    case 'removeSlot':
      return { op, pattern, slot };
    case 'setRepeat':
      return put({ op, pattern } as ManifestChange, { min: count('min'), max: count('max'), widen: num('widen') });
    case 'removeRepeat':
    case 'removePattern':
      return { op, pattern };
    case 'setPattern':
      return put({ op, pattern } as ManifestChange, { id: str('id'), description: str('description'), figures: bool('figures') });
    case 'copyPattern':
      return put({ op, pattern, id: str('id') ?? '' } as ManifestChange, { description: str('description') });
    case 'derivePattern':
      return put({ op, pattern, variant: choice('variant', VARIANTS) ?? values.get('variant') ?? '' } as ManifestChange, { slot: slot || undefined, id: str('id'), description: str('description') });
    case 'addPattern':
      return put({ op, slide: count('slide') ?? 0, id: str('id') ?? '' } as ManifestChange, { description: str('description') });
    case 'setManifest':
      return put({ op } as ManifestChange, {
        name: str('name'),
        language: str('language'),
        direction: choice('direction', ['rtl', 'ltr'] as const),
        alignOnDirectionChange: choice('alignOnDirectionChange', ['flip', 'keep'] as const),
        textScale: num('textScale'),
        characterWidth: cleared('characterWidth'),
        leastTextSize: cleared('leastTextSize'),
        colors: colorsAsked(),
        free: freeAsked(),
      });
    default:
      return undefined;
  }
}

/**
 * Changes a template's manifest as an instruction asks: "drop the second page number on slide
 * 13", "let the cards take up to six", "set the table text to 9 pt". Returns the new manifest
 * with what was changed and what was refused; the template itself is not touched.
 *
 * One model call, and a second when the software refused some of the first answer, so the model
 * can correct it (six cards do not fit; the most is five).
 */
export async function editManifest(template: Template, instruction: string, options: ManifestEditOptions): Promise<ManifestEditResult> {
  const ask = async (prompt: string): Promise<{ note: string; changes: ManifestChange[] }> => {
    const answer = (await options.llm.generateJson({ system: SYSTEM, prompt, schema: SCHEMA, ...(options.signal ? { signal: options.signal } : {}) })) as {
      note?: unknown;
      changes?: unknown;
    } | null;
    const changes = (Array.isArray(answer?.changes) ? answer.changes : [])
      .map((raw) => (raw && typeof raw === 'object' ? toChange(raw as Record<string, unknown>) : undefined))
      .filter((change): change is ManifestChange => change !== undefined);
    return { note: typeof answer?.note === 'string' ? answer.note.trim() : '', changes };
  };

  const prompt = [describeManifest(template), GUIDE, `The person asks:\n${instruction.trim()}`].join('\n\n');
  const first = await ask(prompt);
  let result = changeManifest(template, first.changes);
  let note = first.note;

  if (result.rejected.length > 0) {
    const second = await ask(
      [
        prompt,
        `You answered with these operations:\n${first.changes.map((change, index) => `${index + 1}. ${JSON.stringify(change)}`).join('\n')}`,
        `The software made these:\n${result.applied.map((entry) => `${entry.index + 1}. ${entry.message}`).join('\n') || '(none)'}`,
        `It refused these, and nothing of them was done:\n${result.rejected.map((entry) => `${entry.index + 1}. ${entry.reason}`).join('\n')}`,
        'Answer with operations that replace the refused ones only, so that what the person asked is done as far as it can be. Do not repeat operations that were made. Where the reason shows it cannot be done, give no operation and tell the person in "note" what the limit is.',
      ].join('\n\n'),
    );
    if (second.changes.length > 0) {
      const again = changeManifest({ pkg: template.pkg, manifest: result.manifest }, second.changes);
      const offset = first.changes.length;
      result = {
        manifest: again.manifest,
        applied: [...result.applied, ...again.applied.map((entry) => ({ ...entry, index: entry.index + offset }))],
        rejected: again.rejected.map((entry) => ({ ...entry, index: entry.index + offset })),
        issues: again.issues,
      };
    }
    note = second.note || note;
  }
  return { ...result, note };
}
