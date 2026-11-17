// Characters that no part of a pptx can hold.

/**
 * What XML 1.0 has no way to write: the control characters other than tab, line feed and
 * carriage return, U+FFFE and U+FFFF, and one half of a surrogate pair standing alone.
 */
const UNWRITABLE = /[\u0000-\u0008\u000B\u000C\u000E-\u001F\uFFFE\uFFFF]|[\uD800-\uDBFF](?![\uDC00-\uDFFF])|(?<![\uD800-\uDBFF])[\uDC00-\uDFFF]/g;

/**
 * A text without the characters XML cannot hold. PowerPoint does not open a file that has one
 * of them in a part, and they come with text pasted from other programs: a vertical tab is the
 * line break of a text box. A text that has none comes back as it is.
 */
export function writable(value: string): string {
  return value.replace(UNWRITABLE, '');
}
