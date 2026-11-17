import assert from 'node:assert/strict';
import { test } from 'node:test';
import { buildPalette, colorMap, isGrey, readColors } from '../src/color/palette.ts';
import { hexToOklch, hueDistance } from '../src/color/space.ts';

// The colours of the demo template, weighted as its slides use them.
const demo = [
  { hex: '1B2A41', weight: 60 },
  { hex: '5C6B73', weight: 35 },
  { hex: 'E8F1F0', weight: 55 },
  { hex: '0E7C7B', weight: 24 },
  { hex: 'FFFFFF', weight: 27 },
  { hex: 'F4A259', weight: 9 },
  { hex: 'FCE9D2', weight: 3 },
];

// A template in two greens and a violet, with black text and a grey.
const greens = [
  { hex: '000000', weight: 170 },
  { hex: '00683A', weight: 453 },
  { hex: '00A252', weight: 154 },
  { hex: '6924E8', weight: 122 },
  { hex: '0C0C0C', weight: 109 },
  { hex: 'DCDDDE', weight: 53 },
  { hex: '61FF74', weight: 8 },
];

test('greys, black and white are no family; a pale tint of a colour is', () => {
  for (const hex of ['000000', 'FFFFFF', '808080', '0C0C0C', 'DCDDDE', '5C6B73']) assert.equal(isGrey(hex), true, hex);
  for (const hex of ['0E7C7B', 'E8F1F0', 'FCE9D2', '1B2A41', '6924E8']) assert.equal(isGrey(hex), false, hex);
});

test('the main colour, the second colour and the ink are found, each with its tints', () => {
  const palette = buildPalette(demo);
  assert.equal(palette.accent, '0E7C7B');
  assert.equal(palette.accent2, 'F4A259');
  assert.equal(palette.ink, '1B2A41');
  assert.deepEqual(palette.families.find((family) => family.base === '0E7C7B')?.members, ['E8F1F0', '0E7C7B']);
  assert.deepEqual(palette.families.find((family) => family.base === 'F4A259')?.members, ['F4A259', 'FCE9D2']);
  assert.deepEqual(palette.families.find((family) => family.base === '1B2A41')?.members, ['1B2A41']);
  assert.deepEqual(palette.colors.map((entry) => entry.hex).slice(0, 2), ['1B2A41', 'E8F1F0']);
  // The same colours in another order give the same palette.
  assert.deepEqual(buildPalette(demo.slice().reverse()), palette);
});

test('shades of one hue are one family, named by its most used strong member', () => {
  const palette = buildPalette(greens);
  assert.equal(palette.accent, '00683A');
  assert.equal(palette.accent2, '6924E8');
  assert.equal(palette.ink, undefined);
  assert.deepEqual(palette.families[0]?.members, ['00683A', '00A252', '61FF74']);
  // A manifest may say which colour is the main one.
  const pinned = buildPalette(greens, { accent: '6924E8' });
  assert.equal(pinned.accent, '6924E8');
  assert.equal(pinned.accent2, '00683A');
  // A template of greys has no colour to change.
  assert.deepEqual(buildPalette([{ hex: '000000', weight: 3 }, { hex: 'FFFFFF', weight: 9 }]), { families: [], colors: [{ hex: 'FFFFFF', share: 0.75 }, { hex: '000000', share: 0.25 }] });
});

test('a new main colour takes its tints and shades along and leaves the rest', () => {
  const map = colorMap(buildPalette(demo), { accent: '1565C0' });
  assert.ok(map);
  assert.equal(map('0E7C7B'), '1565C0');
  assert.equal(map('#0e7c7b'), '1565C0');
  const blue = hexToOklch('1565C0');
  // The pale tint becomes a pale tint of the new colour.
  const tint = hexToOklch(map('E8F1F0'));
  assert.ok(tint.l > 0.93 && tint.c > 0.005 && tint.c < 0.06, JSON.stringify(tint));
  assert.ok(hueDistance(tint.h, blue.h) < 15);
  // A soft edge between the colour and white follows too, though the palette never saw it.
  const edge = hexToOklch(map('87BDBD'));
  assert.ok(hueDistance(edge.h, blue.h) < 12 && edge.l > 0.7);
  // Everything else stays: the second colour and its tint, the ink, the greys.
  for (const hex of ['F4A259', 'FCE9D2', '1B2A41', '5C6B73', 'FFFFFF', '000000']) assert.equal(map(hex), hex);
  assert.equal(map('not a colour'), 'not a colour');
});

test('a darker or a lighter target moves the whole family, and keeps their order', () => {
  const palette = buildPalette(greens);
  for (const target of ['0B3D91', '7FB3FF', 'C8102E']) {
    const map = colorMap(palette, { accent: target });
    assert.ok(map);
    assert.equal(map('00683A'), target);
    const [dark, middle, light] = ['00683A', '00A252', '61FF74'].map((hex) => hexToOklch(map(hex)).l) as [number, number, number];
    assert.ok(dark < middle && middle < light, `${target}: ${dark} ${middle} ${light}`);
    assert.equal(map('6924E8'), '6924E8');
    assert.equal(map('000000'), '000000');
  }
  const both = colorMap(palette, { accent: '0B3D91', accent2: 'E07A00' });
  assert.equal(both?.('6924E8'), 'E07A00');
});

test('exact replacements win, and a deck that changes nothing has no map', () => {
  const palette = buildPalette(demo);
  const map = colorMap(palette, { accent: '1565C0', replace: { E8F1F0: 'F5F5F5', '1b2a41': '#000000' } });
  assert.equal(map?.('E8F1F0'), 'F5F5F5');
  assert.equal(map?.('1B2A41'), '000000');
  assert.equal(map?.('0E7C7B'), '1565C0');
  assert.equal(colorMap(palette, undefined), undefined);
  assert.equal(colorMap(palette, {}), undefined);
  assert.equal(colorMap(palette, { accent: '0E7C7B' }), undefined);
  assert.equal(colorMap(palette, { accent: 'blue' }), undefined);
  // A second colour for a template that has none changes nothing.
  assert.equal(colorMap(buildPalette([{ hex: '0E7C7B', weight: 1 }]), { accent2: '1565C0' }), undefined);
});

test('colours that are not colours are named and left out', () => {
  assert.deepEqual(readColors(undefined), { problems: [] });
  assert.deepEqual(readColors({ accent: '#1565c0', accent2: '' }), { colors: { accent: '1565C0' }, problems: [] });
  const bad = readColors({ accent: 'blue', replace: { '0E7C7B': 'teal', FFFFFF: '000000' }, background: '000000' });
  assert.deepEqual(bad.colors, { replace: { FFFFFF: '000000' } });
  assert.equal(bad.problems.length, 3);
  assert.match(bad.problems[0] as string, /"accent" is "blue", which is not six hex digits/);
  assert.match(bad.problems[2] as string, /no setting called "background"/);
  assert.match(readColors('blue').problems[0] as string, /must be an object/);
});
