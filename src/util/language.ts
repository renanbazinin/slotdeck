// Language tags: which ones are written right to left, and which tag left-to-right text gets.

/**
 * The languages written right to left, by their primary tag: Hebrew (also under its old tag
 * `iw`), Yiddish, Arabic, Persian, Urdu, Pashto, Uyghur, Sindhi, Central Kurdish, Syriac and
 * Dhivehi. A tag may carry a region after `-` (or `_`, as some systems write it).
 */
const RTL_LANGUAGE = /^(he|iw|yi|ar|fa|ur|ps|ug|sd|ckb|syr|dv)([-_]|$)/i;

/** Whether a language tag (`he`, `he-IL`, `ar-SA`) names a language written right to left. */
export function isRtlLanguage(tag: string | undefined): boolean {
  return RTL_LANGUAGE.test((tag ?? '').trim());
}

/**
 * The region PowerPoint's own tags pair with a bare language tag. PowerPoint checks spelling by
 * the run's language and names languages with a region (`en-US`, never `en`), so a bare tag gets
 * the one it most often has.
 */
const REGIONS: Record<string, string> = {
  en: 'en-US', fr: 'fr-FR', de: 'de-DE', es: 'es-ES', it: 'it-IT', pt: 'pt-BR', nl: 'nl-NL', ru: 'ru-RU', uk: 'uk-UA',
  pl: 'pl-PL', cs: 'cs-CZ', sk: 'sk-SK', hu: 'hu-HU', ro: 'ro-RO', bg: 'bg-BG', el: 'el-GR', tr: 'tr-TR', sv: 'sv-SE',
  da: 'da-DK', nb: 'nb-NO', no: 'nb-NO', fi: 'fi-FI', ja: 'ja-JP', zh: 'zh-CN', ko: 'ko-KR', hi: 'hi-IN', th: 'th-TH',
};

/**
 * The language tag for left-to-right text in a deck written in `language`: the deck's own
 * language when it is written left to right, with a region added to a bare tag PowerPoint knows
 * (`en` gives `en-US`), and English (`en-US`) when the deck's language is written right to left
 * or not known. Latin text in a Hebrew deck is most often English: a brand, an acronym, a term.
 */
export function leftToRightTag(language: string | undefined): string {
  const tag = (language ?? '').trim();
  if (!tag || isRtlLanguage(tag)) return 'en-US';
  if (/[-_]/.test(tag)) return tag.replace('_', '-');
  return REGIONS[tag.toLowerCase()] ?? tag.toLowerCase();
}
