// Breaking a paragraph into lines the way PowerPoint does, given a way to measure a piece of text.
import { breaksAnywhere } from './fonts.ts';

/** A piece of a paragraph that stays on one line, with the space that follows it. */
interface Unit {
  word: string;
  gap: string;
}

/** A line may break after a hyphen or a dash, as well as at a space. Not at a slash, a full stop or a colon. */
const BREAK_AFTER = /[-‐–—]/;
/** Chinese and Japanese closing punctuation, which never starts a line. */
const CLOSING = /[、。，．！？」』）】ー]/;
const SPACE = /[ \t  - 　]/;

function units(text: string): Unit[] {
  const out: Unit[] = [];
  let word = '';
  const close = (gap = '') => {
    if (word !== '' || gap !== '') out.push({ word, gap });
    word = '';
  };
  const chars = [...text];
  for (let index = 0; index < chars.length; index++) {
    const char = chars[index] as string;
    if (SPACE.test(char)) {
      let gap = char;
      while (index + 1 < chars.length && SPACE.test(chars[index + 1] as string)) gap += chars[++index];
      close(gap);
    } else if (char === '​') {
      close();
    } else if (breaksAnywhere(char)) {
      if (CLOSING.test(char)) {
        word += char;
        continue;
      }
      close();
      word = char;
      const next = chars[index + 1];
      if (next === undefined || !CLOSING.test(next)) close();
    } else {
      word += char;
      if (BREAK_AFTER.test(char) && index + 1 < chars.length && !SPACE.test(chars[index + 1] as string)) close();
    }
  }
  close();
  return out;
}

/**
 * Breaks one paragraph into lines. `first` and `rest` are the widths the first and the later
 * lines may take; `measure` gives the width of a piece of text set on one line, in the same
 * unit. Words go on a line while they fit; a word wider than a whole line is cut where the
 * line ends. Spaces at the end of a line hang outside it and do not count.
 *
 * `inside` names the exception to that: it is asked whether the space after a line that would
 * end with the given text, with the given word to follow, stays inside the line and so needs
 * room. PowerPoint keeps it inside when the text on both sides of that space runs against the
 * paragraph's direction.
 */
export function breakLines(text: string, first: number, rest: number, measure: (text: string) => number, inside?: (line: string, next: string) => boolean): string[] {
  const lines: string[] = [];
  let line = '';
  let width = 0;
  let gap = '';
  const limit = () => (lines.length === 0 ? first : rest);

  const all = units(text);
  for (let index = 0; index < all.length; index++) {
    const unit = all[index] as Unit;
    let word = unit.word;
    let wordWidth = word === '' ? 0 : measure(word);
    const gapWidth = line !== '' && gap !== '' ? measure(gap) : 0;
    const next = all[index + 1]?.word;
    const tail = inside && line !== '' && unit.gap !== '' && next !== undefined && inside(line + gap + word, next) ? measure(unit.gap) : 0;
    if (line !== '' && width + gapWidth + wordWidth + tail <= limit()) {
      line += gap + word;
      width += gapWidth + wordWidth;
    } else {
      if (line !== '') {
        lines.push(line + gap);
        line = '';
        width = 0;
      }
      // A word wider than a whole line is cut by characters. The word is spread once, and what is
      // left of it is not measured whole after every line, which took a time that grew with the
      // square of its length: a piece of it that is wider than a line shows that the rest is, and
      // only a rest that may fit a line is measured whole, for its width. The pieces start at
      // twice the length of the line before and double.
      if (wordWidth > limit()) {
        const chars = [...word];
        let from = 0;
        let last = 0;
        for (;;) {
          const left = chars.length - from;
          if (left === 0) {
            word = '';
            wordWidth = 0;
            break;
          }
          let wider = false;
          for (let piece = Math.max(8, 2 * last); !wider; piece *= 2) {
            if (piece >= left) {
              wordWidth = measure(chars.slice(from).join(''));
              wider = wordWidth > limit();
              break;
            }
            wider = measure(chars.slice(from, from + piece).join('')) > limit();
          }
          if (!wider) {
            word = chars.slice(from).join('');
            break;
          }
          let count = 1;
          while (count < left && measure(chars.slice(from, from + count + 1).join('')) <= limit()) count++;
          lines.push(chars.slice(from, from + count).join(''));
          from += count;
          last = count;
        }
      }
      line = word;
      width = wordWidth;
    }
    gap = unit.gap;
  }
  if (line !== '' || lines.length === 0) lines.push(line + gap);
  else if (gap !== '') lines[lines.length - 1] += gap;
  return lines;
}
