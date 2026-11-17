import assert from 'node:assert/strict';
import { test } from 'node:test';
import { hexToOklch, hexToRgb, hueDistance, hueOffset, oklchToHex, parseHex, rgbToHex } from '../src/color/space.ts';

test('a colour is six hex digits, with or without #, and nothing else', () => {
  assert.equal(parseHex('#0e7c7b'), '0E7C7B');
  assert.equal(parseHex(' F4A259 '), 'F4A259');
  for (const bad of ['blue', '#fff', '0E7C7B00', '', 'GGGGGG', 12, null, undefined, '0E7C7B"/><x']) assert.equal(parseHex(bad), undefined);
});

test('every colour survives the way to OKLCH and back', () => {
  const colours = ['000000', 'FFFFFF', '808080', 'FF0000', '00FF00', '0000FF', '0E7C7B', 'F4A259', '1B2A41', 'E8F1F0', '6924E8', 'FCE9D2', '5C6B73'];
  for (const hex of colours) assert.equal(oklchToHex(hexToOklch(hex)), hex);
  for (let value = 0; value <= 0xffffff; value += 0x010305) {
    const hex = value.toString(16).padStart(6, '0').toUpperCase();
    assert.equal(oklchToHex(hexToOklch(hex)), hex);
  }
});

test('lightness, chroma and hue are those of OKLCH', () => {
  const white = hexToOklch('FFFFFF');
  assert.ok(Math.abs(white.l - 1) < 1e-3 && white.c < 1e-3);
  assert.ok(hexToOklch('000000').l < 1e-6);
  const grey = hexToOklch('808080');
  assert.ok(grey.c < 1e-3 && Math.abs(grey.l - 0.6) < 0.01);
  // Reference values of pure red and blue.
  const red = hexToOklch('FF0000');
  assert.ok(Math.abs(red.l - 0.628) < 0.002 && Math.abs(red.c - 0.2577) < 0.002 && Math.abs(red.h - 29.23) < 0.1);
  const blue = hexToOklch('0000FF');
  assert.ok(Math.abs(blue.l - 0.452) < 0.002 && Math.abs(blue.c - 0.3132) < 0.002 && Math.abs(blue.h - 264.05) < 0.1);
  // A tint keeps the hue of its colour.
  assert.ok(hueDistance(hexToOklch('0E7C7B').h, hexToOklch('E8F1F0').h) < 12);
});

test('a colour no screen can show keeps its hue and lightness and gives up chroma', () => {
  const teal = hexToOklch('0E7C7B');
  const vivid = hexToOklch(oklchToHex({ l: teal.l, c: 0.4, h: teal.h }));
  assert.ok(vivid.c > teal.c && vivid.c < 0.4);
  assert.ok(Math.abs(vivid.l - teal.l) < 0.02);
  assert.ok(hueDistance(vivid.h, teal.h) < 3);
  assert.equal(oklchToHex({ l: 1.4, c: 0.3, h: 120 }), 'FFFFFF');
  assert.equal(oklchToHex({ l: -1, c: 0.3, h: 120 }), '000000');
});

test('hues are compared around the circle', () => {
  assert.equal(hueDistance(350, 10), 20);
  assert.equal(hueDistance(10, 350), 20);
  assert.equal(hueOffset(350, 10), 20);
  assert.equal(hueOffset(10, 350), -20);
  assert.equal(hueOffset(0, 180), -180);
  assert.deepEqual(hexToRgb('#0E7C7B'), [14, 124, 123]);
  assert.equal(rgbToHex(14, 124, 123), '0E7C7B');
  assert.equal(rgbToHex(300, -4, 127.6), 'FF0080');
});
