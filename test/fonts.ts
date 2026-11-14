// Font files for the tests. The demo template is set in Calibri and Cambria, which are licensed
// and not in this repository. Two things stand in for them:
// - open fonts from dev dependencies, which every machine and CI has. Carlito and Caladea were
//   drawn to the metrics of Calibri and Cambria, so Latin text measures the same; Heebo covers Hebrew.
// - the real fonts, where they are installed (Windows or macOS with Office), for the checks
//   against PowerPoint's own line breaks.
import { existsSync, readFileSync } from 'node:fs';
import { createRequire } from 'node:module';
import { dirname, join } from 'node:path';
import type { FontFiles } from '../src/index.ts';

const require = createRequire(import.meta.url);

function packaged(name: string, files: string[]): Uint8Array[] {
  const folder = dirname(require.resolve(`@expo-google-fonts/${name}/package.json`));
  return files.map((file) => new Uint8Array(readFileSync(join(folder, file))));
}

const FOUR = (family: string) => [`400Regular/${family}_400Regular.ttf`, `700Bold/${family}_700Bold.ttf`, `400Regular_Italic/${family}_400Regular_Italic.ttf`, `700Bold_Italic/${family}_700Bold_Italic.ttf`];

/** Open fonts with the metrics of the demo template's typefaces, and Heebo for Hebrew. */
export function openFonts(): FontFiles {
  return {
    Calibri: packaged('carlito', FOUR('Carlito')),
    Cambria: packaged('caladea', FOUR('Caladea')),
    Heebo: packaged('heebo', ['400Regular/Heebo_400Regular.ttf', '700Bold/Heebo_700Bold.ttf']),
  };
}

const OFFICE_FOLDERS = [process.env.SLOTDECK_FONT_DIR, 'C:/Windows/Fonts', '/Library/Fonts/Microsoft', '/Applications/Microsoft PowerPoint.app/Contents/Resources/DFonts'];

/** Calibri and Cambria themselves, or `undefined` where they are not installed. */
export function officeFonts(): FontFiles | undefined {
  const names = {
    Calibri: ['calibri.ttf', 'calibrib.ttf', 'calibrii.ttf', 'calibriz.ttf'],
    Cambria: ['cambria.ttc', 'cambriab.ttf', 'cambriai.ttf', 'cambriaz.ttf'],
  };
  for (const folder of OFFICE_FOLDERS) {
    if (!folder || !Object.values(names).flat().every((file) => existsSync(join(folder, file)))) continue;
    return Object.fromEntries(Object.entries(names).map(([family, files]) => [family, files.map((file) => new Uint8Array(readFileSync(join(folder, file))))]));
  }
  return undefined;
}
