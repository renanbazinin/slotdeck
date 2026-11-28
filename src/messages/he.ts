// The Hebrew sentence of every code, for an app whose people read Hebrew. Each says what the
// English one says, with the same numbers, addressed to no one in particular (no imperative of
// either gender). Numbers are digits, but for a count of one, which Hebrew says with the word
// after the noun ("שורה אחת"); a label, a name or a value given in the deck stands in an isolate
// of its own, so that it reads in its own direction inside a right-to-left line.
import type { FreeBlockKind } from '../deck/spec.ts';
import type { IssueCatalogue } from './codes.ts';

/**
 * A text inside a right-to-left line, kept in its own direction: "p1-p14", "3-cards", "-5, 12.5"
 * and a sentence of the review model's come out as they were written. U+2068 FIRST STRONG ISOLATE
 * and U+2069 POP DIRECTIONAL ISOLATE: the text takes the direction of its first letter.
 */
const FIRST_STRONG_ISOLATE = '\u2068';
const POP_DIRECTIONAL_ISOLATE = '\u2069';
export const isolate = (text: string): string => `${FIRST_STRONG_ISOLATE}${text}${POP_DIRECTIONAL_ISOLATE}`;

const at = (slide: number | undefined): string => (slide !== undefined ? `שקף ${slide}: ` : '');
/** So many of a thing: "5 שורות", and one of it with the word, "שורה אחת". */
const count = (amount: number, one: string, many: string, oneWord: 'אחד' | 'אחת'): string => (amount === 1 ? `${one} ${oneWord}` : `${amount} ${many}`);
/** About so many of a thing, always in digits: "כ-41 תווים", "כ-1 תו". */
const about = (amount: number, one: string, many: string): string => `כ-${amount} ${amount === 1 ? one : many}`;
/** A one-letter preposition joined to what follows: "ל-5 שורות", "לשורה אחת", "לכ-22 תווים". */
const joined = (preposition: 'ל' | 'ב', phrase: string): string => (/^\d/.test(phrase) ? `${preposition}-${phrase}` : `${preposition}${phrase}`);
const characters = (amount: number): string => count(amount, 'תו', 'תווים', 'אחד');
const aboutCharacters = (amount: number): string => about(amount, 'תו', 'תווים');
const lines = (amount: number): string => count(amount, 'שורה', 'שורות', 'אחת');
const words = (amount: number): string => count(amount, 'מילה', 'מילים', 'אחת');
const parts = (amount: number): string => count(amount, 'חלק', 'חלקים', 'אחד');
const paragraphs = (amount: number): string => count(amount, 'פסקה', 'פסקאות', 'אחת');
const columns = (amount: number): string => count(amount, 'עמודה', 'עמודות', 'אחת');
const points = (size: number): string => `${size} נק'`;
/** The passages of a document a slide carries: "קטעים p3-p5", "קטע p3". */
const passages = (label: string): string => `${/^p\d+$/.test(label) ? 'קטע' : 'קטעים'} ${isolate(label)}`;
const holds = (slide: number, blocks: number): string => `שקף ${slide} מכיל ${parts(blocks)}, כפי שהתבקש`;
const fitsHint = (fitsChars: number): string => `נכנס בה טקסט של ${aboutCharacters(fitsChars)}.`;
const atSize = (shrunk: boolean, size: number): string => `${shrunk ? 'גם ב-' : 'ב-'}${points(size)}`;
const inSize = (size: number | undefined): string => (size !== undefined ? ` (בגופן ${points(size)})` : '');
const ignored = 'ולכן לא נלקחה בחשבון';
/** The slot beside a text in a row, named as a Hebrew line names slots: "הטקסט של כרטיס 2", or its own name in quotes. */
const hebrewBeside = (slot: string): string => {
  const named = hebrewSlot(slot);
  return named.startsWith(FIRST_STRONG_ISOLATE) ? `"${named}"` : `ה${named}`;
};

/** Said at the end of a cut where free slides are off. */
export const HE_FREE_OFF = '; שקפים חופשיים אינם פעילים, כך שהטקסט לא מוצג בגופן קטן יותר';

/**
 * Why a layout cannot carry a free slide, in Hebrew, for the reasons `freeBaseRefusal` gives about
 * the layout itself; a reason it does not know stays as it was, in English, isolated.
 */
function refusal(why: string, base: string): string {
  const its = (name: string, own: string, other: string): string => (name === base ? own : other.replace('%', `"${isolate(name)}"`));
  let found = /^the layout "(.+)" has its title low on the slide: blocks would get (\d+)% of the slide's height, and (?:a frame leaves them (\d+)% or more|"(.+)" leaves them (\d+)%)$/.exec(why);
  if (found) {
    const [, name = '', room, least, instead, insteadRoom] = found;
    const than = least !== undefined ? `ומסגרת משאירה להם ${least}% או יותר` : `והפריסה "${isolate(instead ?? '')}" משאירה להם ${insteadRoom}%`;
    return `${its(name, 'הכותרת שלה', 'הכותרת של הפריסה %')} נמוכה בשקף: לחלקים היו נשארים ${room}% מגובה השקף, ${than}`;
  }
  found = /^the layout "(.+)" has no title$/.exec(why);
  if (found) return its(found[1] ?? '', 'אין בה כותרת', 'אין כותרת בפריסה %');
  found = /^the layout "(.+)" gives no frame of its own$/.exec(why);
  if (found) return its(found[1] ?? '', 'אין לה מסגרת משלה', 'לפריסה % אין מסגרת משלה');
  found = /^the layout "(.+)" has a row of (.+)s, which the software lays out itself$/.exec(why);
  if (found) return `${its(found[1] ?? '', 'יש בה', 'יש בפריסה %')} שורה של פריטים (${isolate(found[2] ?? '')}), שהתוכנה מסדרת בעצמה`;
  return isolate(why);
}

export const HE: IssueCatalogue = {
  'text-too-wide': (v) => `הטקסט רחב ${v.approximate ? 'בכ-' : 'ב-'}${v.percent}% מהתיבה שלו ${atSize(v.shrunk, v.size)}, והתיבה אינה מעבירה טקסט לשורה הבאה. ${fitsHint(v.fitsChars)}`,
  'text-overflow': (v) =>
    `הטקסט תופס ${v.approximate ? about(v.lines, 'שורה', 'שורות') : lines(v.lines)} ${atSize(v.shrunk, v.size)}; בתיבה יש מקום ${joined('ל', lines(v.capacity))} (${lines(v.over)} יותר מדי). ${fitsHint(v.fitsChars)}`,
  'colors-not-object': (v) => `${at(v.slide)}ההגדרה "colors" צריכה להיות אובייקט כמו ${isolate('{ "accent": "1565C0" }')}; היא לא נלקחה בחשבון`,
  'color-not-hex': (v) => `${at(v.slide)}הצבע "${isolate(v.name)}" הוא ${isolate(v.given)}, שאינו קוד של שש ספרות הקסדצימליות כמו "1565C0", ולכן לא נלקח בחשבון`,
  'colors-replace-not-map': (v) => `${at(v.slide)}ההגדרה "replace" צריכה למפות צבעים לצבעים, כמו ${isolate('{ "0E7C7B": "1565C0" }')}; היא לא נלקחה בחשבון`,
  'colors-replace-not-pair': (v) => `${at(v.slide)}ההחלפה ${isolate(`${v.from}: ${v.to}`)} אינה זוג צבעים של שש ספרות הקסדצימליות, ${ignored}`,
  'colors-unknown-setting': (v) => `${at(v.slide)}אין ב-"colors" הגדרה בשם "${isolate(v.name)}", ${ignored}`,
  'colors-kept-by-template': () => 'התבנית שומרת על הצבעים שלה בכל מצגת, ולכן הצבעים שהמצגת מבקשת לא הוחלו',
  'columns-not-whole': (v) => `שקף ${v.slide}: מספר החלקים בשורה, ${isolate(v.given)}, אינו מספר שלם (1 ומעלה), ולכן לא נלקח בחשבון`,
  'columns-relaxed': (v) => `שקף ${v.slide}: ${v.asked === 1 ? 'התבקש חלק אחד' : `התבקשו ${v.asked} חלקים`} בשורה, והשקף מכיל ${parts(v.columns)} בשורה`,
  'free-text-small': (v) => `${holds(v.slide, v.blocks)}; הטקסט הקטן ביותר בו מוצג בגופן ${points(v.size)}`,
  'free-larger-text': (v) => `${holds(v.slide, v.blocks)}; המצגת מבקשת טקסט גדול יותר, ${v.half ? 'ובשקף הזה יש מקום רק לחצי מההגדלה' : 'ובשקף הזה אין לכך מקום'}`,
  'free-charts-bare': (v) => `${holds(v.slide, v.blocks)}; ${v.charts === 1 ? 'תרשים אחד מוצג' : `${v.charts} תרשימים מוצגים`} בגודל קטן, בלי צירים, מקרא או תוויות`,
  'free-blocks-low': (v) =>
    `${holds(v.slide, v.blocks)}; ${v.count === 1 ? 'אחד מהם יורד' : `${v.count} מהם יורדים`} אל מתחת לשטח הפנוי בשקף, ${v.count === 1 ? 'ועלול' : 'ועלולים'} להיחתך בקצה השקף`,
  'free-base-refused': (v) =>
    `שקף ${v.slide} תוכנן על המסגרת של הפריסה "${isolate(v.base)}", שאי אפשר לבנות עליה שקף חופשי (${refusal(v.why, v.base)}); לכן הוא בנוי על המסגרת של הפריסה "${isolate(v.stands)}"`,
  'free-base-gone': (v) => `שקף ${v.slide} תוכנן על המסגרת של הפריסה "${isolate(v.base)}", שכבר אינה בתבנית; לכן הוא בנוי על המסגרת של הפריסה "${isolate(v.stands)}"`,
  'items-ignored': (v) => `לפריסה "${isolate(v.pattern)}" יש מספר פריטים קבוע, ולכן "items" לא נלקח בחשבון`,
  'slot-unknown': (v) => `אין בפריסה "${isolate(v.pattern)}" תיבה כזו, ולכן התוכן שלה לא נלקח בחשבון`,
  'slot-over-chars': (v) => `הטקסט באורך ${characters(v.chars)}, ובתיבה יש מקום ${joined('ל', aboutCharacters(v.most))}`,
  'slot-over-paragraphs': (v) => `בטקסט יש ${paragraphs(v.paragraphs)}, ובתיבה יש מקום ${joined('ל', paragraphs(v.most))}`,
  'table-over-columns': (v) => `בטבלה יש ${columns(v.columns)}, ובתיבה יש מקום ${joined('ל', columns(v.most))}`,
  'table-over-rows': (v) => `בטבלה יש ${lines(v.rows)}, ובתיבה יש מקום ${joined('ל', lines(v.most))}`,
  'table-cell-long': (v) => `תא בטבלה באורך ${characters(v.chars)}, ובתא יש מקום ${joined('ל', aboutCharacters(v.most))}`,
  'chart-over-categories': (v) =>
    `בתרשים יש ${count(v.categories, 'קטגוריה', 'קטגוריות', 'אחת')}, ובתיבה יש מקום ${joined('ל', count(v.most, 'קטגוריה', 'קטגוריות', 'אחת'))}`,
  'chart-over-series': (v) => `בתרשים יש ${count(v.series, 'סדרה', 'סדרות', 'אחת')}, ובתיבה יש מקום ${joined('ל', count(v.most, 'סדרה', 'סדרות', 'אחת'))}`,
  'chart-format-unknown': (v) => `פורמט התרשים ${isolate(v.format)} אינו מוכר ולא נלקח בחשבון; אפשר להשתמש ב-"number" או ב-"percent"`,
  'chart-values-unstated': (v) => `ערכים בתרשים שלא צוינו בבקשה למצגת: ${isolate(v.numbers)}. כדאי לבדוק אותם לפני השימוש במצגת.`,
  'text-details-unstated': (v) => `${v.count === 1 ? 'פרט שלא צוין' : 'פרטים שלא צוינו'} בבקשה למצגת: ${isolate(v.details)}. כדאי לבדוק ${v.count === 1 ? 'אותו' : 'אותם'} לפני השימוש במצגת.`,
  'word-wider-than-line': (v) => `המילה ${isolate(v.word)} (${characters(v.chars)}) רחבה משורה של התיבה, ולכן היא נשברת באמצע`,
  'text-stand-in': (v) => `במקום ערך מופיע ממלא מקום: ${isolate(v.found)}. כדאי לכתוב במקומו את הערך האמיתי, או להשמיט אותו.`,
  'text-about-slide': (v) => `הטקסט מדבר על השקף (${isolate(v.found)}) במקום לומר את התוכן עצמו.`,
  'text-orphan': (v) =>
    `הטקסט תופס ${lines(v.lines)}, ובאחרונה שבהן יש מילה אחת בלבד, ${isolate(v.word)}. טקסט של ${aboutCharacters(v.chars)} יסתיים שורה אחת קודם; עוד כמה מילים ימלאו את השורה.`,
  'text-empty': (v) =>
    `הטקסט תופס ${lines(v.lines)}, ${v.sampleLines !== undefined ? `ואילו הטקסט של התבנית תופס ${lines(v.sampleLines)}` : `ובתיבה יש מקום ${joined('ל', lines(v.capacity ?? 0))}`}, ולכן התיבה תיראה ריקה. טקסט של ${aboutCharacters(v.chars)} ימלא אותה כפי שתוכננה.`,
  'row-uneven': (v) =>
    `הטקסט תופס ${lines(v.lines)}, ואילו ${hebrewBeside(v.beside)} שלידו תופס ${lines(v.besideLines)}, ולכן התיבות שבשורה ייראו לא אחידות. טקסט של ${aboutCharacters(v.chars)} או יותר יאזן ביניהן.`,
  'seen-in-picture': (v) => `נראה בתמונת השקף: ${isolate(v.problem)}`,
  'accent2-missing': (v) => `${at(v.slide)}לתבנית אין צבע שני, ולכן "accent2" לא הוחל`,
  'color-near-page': (v) => `הצבע ${isolate(v.hex)} קרוב לצבע הרקע של השקף; מילויים, עמודות של תרשימים וסמלים בצבע הזה כמעט לא ייראו`,
  'part-unwritten': () => 'לא נכתב עבורו דבר, והוא ייכתב שוב',
  'photos-over': (v) => `אפשר ליצור לכל היותר ${v.most} תמונות לשקף אחד; לשאר יש להשתמש בתמונה שסופקה`,
  'parts-left-out-empty': (v) => `שקף ${v.slide}: ${v.count === 1 ? 'חלק אחד לא הכיל תוכן והושמט' : `${v.count} חלקים לא הכילו תוכן והושמטו`}`,
  'parts-left-out-unusable': (v) => `שקף ${v.slide}: ${v.count === 1 ? 'אי אפשר היה להציג חלק אחד, והוא הושמט' : `אי אפשר היה להציג ${v.count} חלקים, והם הושמטו`}`,
  'photos-left-out': (v) => `שקף ${v.slide}: ${v.count === 1 ? 'תמונה אחת הושמטה' : `${v.count} תמונות הושמטו`}, כי לשקף אחד נוצרות לכל היותר ${v.most} תמונות`,
  'parts-unchanged': (v) => `שקף ${v.slide}: אי אפשר היה להציג את מה שנכתב עבור ${parts(v.count)}, ולכן ${v.count === 1 ? 'הוא נשאר כפי שהיה' : 'הם נשארים כפי שהיו'}`,
  'parts-unwritten': (v) => `שקף ${v.slide}: לא נכתב דבר עבור ${parts(v.count)}, ולכן ${v.count === 1 ? 'הוא נשאר כפי שהיה' : 'הם נשארים כפי שהיו'}`,
  'written-under-retry': (v) => `הטקסט באורך ${characters(v.chars)}, ולשקף הזה התבקשו ${words(v.words)}: יש לכתוב ${v.least} עד ${v.most} תווים, כפי שהתבקש`,
  'written-under': (v) => `לשקף הזה התבקשו ${words(v.words)}, ויש בו ${about(v.hasWords, 'מילה', 'מילים')}; אפשר לבקש אותו שוב כדי לקבל את כולן`,
  'words-cut': (v) => `לשקף הזה התבקשו ${words(v.words)}, והוא מכיל כ-${v.holdsWords} מהן; אפשר לבקש את הטקסט על פני יותר שקפים, או טקסט קצר יותר${v.freeOff ? HE_FREE_OFF : ''}`,
  'words-cut-dense': (v) =>
    `לשקף הזה התבקשו ${words(v.words)}, והוא מכיל כ-${v.holdsWords} מהן גם בגופן קטן של ${points(v.size)}: יתר הטקסט נחתך. אפשר לבקש את הטקסט על פני יותר שקפים, או טקסט קצר יותר`,
  'photo-left-out': (v) =>
    `התמונה של השקף הזה הושמטה כדי לפנות מקום ${v.asked ? `ל-${v.words} המילים שהתבקשו עבורו` : `לטקסט שלו, של ${about(v.words, 'מילה', 'מילים')}`}`,
  'words-kept-low-retry': (v) =>
    `נשמרו ${v.percent}% מהמילים של הטקסט מהמסמך (${passages(v.label)}), ובשקף יש מקום כמעט לכל הטקסט: יש לכתוב את הטקסט במילים ובסדר שלו, ולקצר רק היכן שמגבלה מחייבת`,
  'words-kept-low': (v) => `בשקף נשארו ${v.percent}% מהמילים של הטקסט שלו מהמסמך (${passages(v.label)}), אף שיש בו מקום כמעט לכל הטקסט`,
  'text-cut': (v) =>
    `הטקסט מהמסמך לשקף הזה (${passages(v.label)}) באורך ${characters(v.chars)}, והשקף מכיל כ-${v.holds} תווים: הטקסט נחתך כדי שייכנס. אפשר לבקש יותר שקפים כדי שהטקסט יישאר שלם${v.freeOff ? HE_FREE_OFF : ''}`,
  'paragraphs-joined': (v) =>
    `בטקסט מהמסמך לשקף הזה (${passages(v.label)}) יש ${paragraphs(v.paragraphs)}, והשקף מציג אותו ${joined('ב', paragraphs(v.holdsParagraphs))} בלבד: חלק מהפסקאות חוברו זו לזו. אפשר לבקש יותר שקפים כדי שהפסקאות יישארו נפרדות`,
  'text-cut-dense': (v) =>
    `הטקסט מהמסמך לשקף הזה (${passages(v.label)}) באורך ${characters(v.chars)}, והשקף מכיל כ-${v.holds} תווים גם בגופן קטן של ${points(v.size)}: הטקסט נחתך כדי שייכנס. אפשר לבקש יותר שקפים כדי שהטקסט יישאר שלם`,
  'nothing-changed': (v) => `שקף ${v.slide}: לא שונה בו דבר, כי תוכנית השינוי לא ציינה אף חלק לכתוב מחדש`,
  'recolor-slide-removed': (v) => `השינוי הסיר את שקף ${v.slide}, שהיה אמור לקבל צבע חדש`,
  'change-would-cut': (v) => `שקף ${v.slide} היה מכיל רק כ-${v.holdsWords} מתוך ${v.words} המילים שלו${inSize(v.size)}, ולכן השינוי לא בוצע${v.freeOff ? HE_FREE_OFF : ''}`,
  'slide-holds-less': (v) => `שקף ${v.slide} מכיל כ-${v.holdsWords} מתוך ${v.words} המילים שלו${inSize(v.size)}: יתר הטקסט ייחתך${v.freeOff ? HE_FREE_OFF : ''}`,
  'slide-set-small': (v) => `שקף ${v.slide} מכיל את כל הטקסט שלו בגופן קטן מזה של התבנית: ${points(v.size)}, לעומת ${points(v.templateSize)} בתבנית`,
  'text-retold': (v) => `תיאור הנושא מנסח מחדש את "${isolate(v.name)}", ולכן השקפים לא יכילו את הטקסט הזה מילה במילה`,
};

/** The kinds of part of a free slide, as a line names them: "חלק 3 (תרשים)". */
export const HE_KINDS: Record<FreeBlockKind, string> = {
  text: 'טקסט',
  list: 'רשימה',
  card: 'כרטיס',
  figure: 'נתון',
  step: 'שלב',
  phase: 'שלב בתהליך',
  milestone: 'אבן דרך',
  option: 'חלופה',
  quote: 'ציטוט',
  note: 'הערה',
  chart: 'תרשים',
  table: 'טבלה',
  photo: 'תמונה',
  graphic: 'איור',
};

/** The parts of a part, as a line names them: "חלק 7 (כרטיס), כותרת". */
export const HE_PARTS: Record<string, string> = {
  heading: 'כותרת',
  text: 'טקסט',
  items: 'פריטים',
  icon: 'סמל',
  value: 'ערך',
  label: 'תווית',
  quote: 'ציטוט',
  source: 'מקור',
  caption: 'כיתוב',
  chart: 'תרשים',
  table: 'טבלה',
  photo: 'תמונה',
  graphic: 'איור',
  when: 'מועד',
  badge: 'תג',
};

/** Slots that most templates name alike, and the groups of a row of them ("card2Heading"). */
const HE_SLOTS: Record<string, string> = { title: 'כותרת', subtitle: 'כותרת משנה', body: 'גוף הטקסט' };
const HE_GROUPS: Record<string, string> = { column: 'עמודה', col: 'עמודה', card: 'כרטיס', point: 'נקודה' };
const HE_GROUP_PARTS: Record<string, string> = { Heading: 'כותרת', Text: 'טקסט' };

/**
 * A slot of a layout as a Hebrew line names it: "כותרת", "גוף הטקסט", "כותרת של כרטיס 2". A slot
 * named otherwise keeps its name, isolated.
 */
export function hebrewSlot(slot: string): string {
  // Only the tables' own names: "constructor" names no slot.
  if (Object.hasOwn(HE_SLOTS, slot)) return HE_SLOTS[slot] as string;
  const found = /^(column|col|card|point)(\d+)(Heading|Text)$/.exec(slot);
  if (found) return `${HE_GROUP_PARTS[found[3] as string]} של ${HE_GROUPS[found[1] as string]} ${found[2]}`;
  return isolate(slot);
}
