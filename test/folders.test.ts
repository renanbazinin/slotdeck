// An app's own icons and pictures read from a folder (`iconFiles`, `assetFiles` of slotdeck/node).
import assert from 'node:assert/strict';
import { mkdirSync, mkdtempSync, readFileSync, rmSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { dirname, join } from 'node:path';
import { after, test } from 'node:test';
import { fileURLToPath } from 'node:url';
import { placeholderPhoto } from '../examples/placeholder.ts';
import { freeSlide, loadTemplate, renderDeck, resolveAssets, validateDeck } from '../src/index.ts';
import type { DeckSpec, TemplateManifest } from '../src/index.ts';
import { assetFiles, iconFiles } from '../src/node.ts';

const root = join(dirname(fileURLToPath(import.meta.url)), '..');
const png = placeholderPhoto(40, 40, [14, 124, 123], [244, 162, 89]);
const made: string[] = [];
/** A folder with these files, taken away when the tests are done. */
function folder(files: Record<string, Uint8Array | string>): string {
  const dir = mkdtempSync(join(tmpdir(), 'slotdeck-folder-'));
  made.push(dir);
  for (const [name, content] of Object.entries(files)) {
    mkdirSync(dirname(join(dir, name)), { recursive: true });
    writeFileSync(join(dir, name), content);
  }
  return dir;
}
after(() => {
  for (const dir of made) rmSync(dir, { recursive: true, force: true });
});

test('a folder of icons is an icon set: each picture by its file name, described by the folder\'s list', () => {
  // The demo's own icon folder is such a folder: pictures, and icons.json with a line for each.
  const icons = iconFiles(join(root, 'examples', 'icons'));
  const list = JSON.parse(readFileSync(join(root, 'examples', 'icons', 'icons.json'), 'utf8')) as Array<{ id: string; description: string }>;
  assert.equal(icons.length, list.length);
  assert.deepEqual(icons.map((icon) => icon.id), list.map((entry) => entry.id).sort(), 'in the order of their ids');
  for (const icon of icons) {
    assert.equal(icon.description, list.find((entry) => entry.id === icon.id)?.description);
    assert.deepEqual([...icon.data.subarray(1, 4)], [0x50, 0x4e, 0x47], `${icon.id} is the PNG of its file`);
  }

  // It is what the calls that take an icon set take: a deck that names an icon gets its picture.
  const template = loadTemplate(new Uint8Array(readFileSync(join(root, 'templates', 'demo.pptx'))), JSON.parse(readFileSync(join(root, 'templates', 'demo.manifest.json'), 'utf8')) as TemplateManifest);
  const deck: DeckSpec = { version: 1, fields: { organizationDepartment: 'Riverbend' }, slides: [freeSlide({ title: 'Why', blocks: [{ card: { heading: 'Fast', text: 'Two seconds.', icon: 'bolt' } }, { card: { heading: 'On time', text: 'Every day.', icon: 'clock' } }] })] };
  assert.deepEqual(validateDeck(template, deck, { layout: 'adaptive', assets: icons.map((icon) => icon.id) }), []);
  return resolveAssets(template, deck, { icons }).then((assets) => {
    assert.deepEqual(Object.keys(assets).sort(), ['bolt', 'clock']);
    assert.ok(renderDeck(template, deck, { assets, layout: 'adaptive' }).pptx.length > 1000);
  });
});

test('a picture nobody described is described by the words of its file name, and what is no picture is passed over and said', () => {
  const dir = folder({ 'money-bag_2.png': png, 'ClockFace.PNG': png, 'photo.JPG': new Uint8Array([0xff, 0xd8, 0xff, 0xc0, 0, 11, 8, 0, 4, 0, 6, 1, 1, 0x11, 0]), 'logo.svg': '<svg/>', 'shot.webp': 'RIFF', 'broken.png': 'not a picture', 'notes.txt': 'hello', 'inside/deep.png': png });
  const skipped: string[] = [];
  const icons = iconFiles(dir, { onSkipped: (file, reason) => skipped.push(`${file}: ${reason}`) });
  assert.deepEqual(icons.map((icon) => [icon.id, icon.description]), [['ClockFace', 'Clock Face'], ['money-bag_2', 'money bag 2'], ['photo', 'photo']]);
  assert.deepEqual(skipped, [
    'broken.png: it is named as a picture, and its content is no PNG, JPEG or GIF',
    'logo.svg: pictures are read as PNG, JPEG or GIF: save it as a PNG',
    'shot.webp: pictures are read as PNG, JPEG or GIF: save it as a PNG',
  ]);
  // Without anyone to tell, the same icons come back, and nothing is thrown for a file that is no picture.
  assert.deepEqual(iconFiles(dir).map((icon) => icon.id), ['ClockFace', 'money-bag_2', 'photo']);
  // The folders inside are read only when asked, and a picture there is still known by its own name.
  assert.deepEqual(iconFiles(dir, { recursive: true }).map((icon) => icon.id), ['ClockFace', 'deep', 'money-bag_2', 'photo']);
  assert.deepEqual(iconFiles(folder({})), [], 'an empty folder is an empty set');
});

test('descriptions come from an object, from a named file, or from the folder\'s usual list', () => {
  const pictures = { 'clock.png': png, 'coin.png': png };
  // An object in code.
  assert.deepEqual(iconFiles(folder(pictures), { descriptions: { clock: 'a clock: time, schedule' } }).map((icon) => icon.description), ['a clock: time, schedule', 'coin']);
  // The usual file, as an object or as a list; a line with nothing in it leaves the file's words.
  assert.deepEqual(iconFiles(folder({ ...pictures, 'icons.json': JSON.stringify({ clock: 'a clock: time', coin: '  ' }) })).map((icon) => icon.description), ['a clock: time', 'coin']);
  assert.deepEqual(iconFiles(folder({ ...pictures, 'icons.json': JSON.stringify([{ id: 'coin', description: 'a coin: money, price' }]) })).map((icon) => icon.description), ['clock', 'a coin: money, price']);
  // A file of another name is read when it is named, and an object given in code leaves the folder's file unread.
  const named = folder({ ...pictures, 'he.json': JSON.stringify({ clock: 'שעון: זמן, לוח זמנים' }), 'icons.json': JSON.stringify({ clock: 'a clock' }) });
  assert.equal(iconFiles(named, { descriptions: 'he.json' })[0]?.description, 'שעון: זמן, לוח זמנים');
  assert.equal(iconFiles(named, { descriptions: {} })[0]?.description, 'clock');
  // A named file that is not there, or is no JSON, is an error that names it; the usual one may be missing.
  assert.throws(() => iconFiles(folder(pictures), { descriptions: 'missing.json' }), /missing\.json cannot be read as the descriptions/);
  assert.throws(() => iconFiles(folder({ ...pictures, 'icons.json': '{ not json' })), /icons\.json cannot be read as the descriptions/);
  // A line for a picture the folder does not have is said, and stops nothing.
  const skipped: string[] = [];
  iconFiles(folder({ ...pictures, 'icons.json': JSON.stringify({ clock: 'a clock', gone: 'a ghost' }) }), { onSkipped: (file, reason) => skipped.push(`${file}: ${reason}`) });
  assert.deepEqual(skipped, ['icons.json: it describes "gone", and the folder has no such picture']);
});

test('a list may give a picture an id that is not its file name, and two pictures of one id are an error', () => {
  const dir = folder({ 'IMG_0042.png': png, 'team.png': png, 'assets.json': JSON.stringify([{ id: 'handshake', file: 'IMG_0042.png', description: 'Two hands shaking: agreement, a deal closed' }]) });
  assert.deepEqual(assetFiles(dir, 'photo').map((asset) => [asset.id, asset.kind, asset.description]), [['handshake', 'photo', 'Two hands shaking: agreement, a deal closed'], ['team', 'photo', 'team']]);
  assert.deepEqual(assetFiles(dir, 'graphic').map((asset) => asset.kind), ['graphic', 'graphic']);
  // The same name in two formats, or in two folders, or given to two files by a list.
  assert.throws(() => iconFiles(folder({ 'a.png': png, 'a.gif': new Uint8Array([0x47, 0x49, 0x46, 0x38, 0x39, 0x61, 1, 0, 1, 0]) })), /two pictures with the id "a": a\.gif and a\.png/);
  assert.throws(() => iconFiles(folder({ 'a.png': png, 'more/a.png': png }), { recursive: true }), /two pictures with the id "a": a\.png and more\/a\.png/);
  assert.throws(() => iconFiles(folder({ 'a.png': png, 'b.png': png, 'icons.json': JSON.stringify([{ id: 'a', file: 'b.png', description: 'x' }]) })), /two pictures with the id "a"/);
  assert.throws(() => iconFiles(join(root, 'no-such-folder')), /ENOENT/);
});
