// What the slides state that a person should check (0.7.0): contact details and figures the brief
// does not give, stand-ins where a value belongs, and text about the slide itself. No model.
import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
import { test } from 'node:test';
import { formatIssue } from '../src/index.ts';
import type { DeckSpec, Issue } from '../src/index.ts';
import { unstatedDetails } from '../src/deck/numbers.ts';

const brief = 'A deck for new adopters. Call us on 03-555-1234 or write to hello@shelter.example. 9 in 10 adoptions last; 40% of our dogs are seniors. Visit www.shelter.example/adopt.';
const deckOf = (...texts: Array<string | string[]>): DeckSpec => ({ version: 1, brief, fields: { contact: 'office@shelter.example' }, slides: texts.map((text) => ({ pattern: 'text', slots: { body: text } })) });
const found = (issues: Issue[]) => issues.map((issue) => `${issue.slide}:${issue.code}:${JSON.stringify(issue.values)}`);

test('the showcase deck says an email address its brief never gave', () => {
  const showcase = JSON.parse(readFileSync(new URL('../examples/showcase/rescue-dog/deck.json', import.meta.url), 'utf8')) as DeckSpec;
  const issues = unstatedDetails(showcase).filter((issue) => issue.code === 'text-details-unstated');
  assert.deepEqual(issues.map((issue) => [issue.slot, issue.values]), [['subtitle', { details: 'adopt@example.org', count: 1 }]]);
  assert.equal(issues[0]?.message, 'a detail not stated in the brief: adopt@example.org. Check it before using the deck.');
  assert.equal(formatIssue(issues[0] as Issue, 'he'), 'פרט שלא צוין בבקשה למצגת: ⁨adopt@example.org⁩. כדאי לבדוק אותו לפני השימוש במצגת.');
});

test('details the brief, the document or the fields give are not told; made-up ones are', () => {
  // Given: the phone number in another form, the address, a field's address, the percentage, the web address.
  assert.deepEqual(found(unstatedDetails(deckOf('Call 035551234 or write to hello@shelter.example; office@shelter.example answers too. 40% are seniors. See www.shelter.example/adopt.'))), []);
  // Made up: a hotline open day and night, an address, a percentage, a mobile number, a web page.
  assert.deepEqual(found(unstatedDetails(deckOf(['Our hotline is open 24/7 on +972 50 123 4567.', 'Write to adopt@shelter.example; 85% return within a week.', 'More at https://shelter.example/faq.']))), [
    '1:text-details-unstated:{"details":"24/7, +972 50 123 4567, adopt@shelter.example, 85%, https://shelter.example/faq","count":5}',
  ]);
  // A span of years, a large number and a date are no phone numbers.
  assert.deepEqual(found(unstatedDetails(deckOf('From 2023-2025 we rehomed 1,200,000 meals on 12.03.2026.'))), []);
  // A document is given text too, and a check can be limited to some slides.
  assert.deepEqual(found(unstatedDetails({ ...deckOf('85% stay'), document: 'Of our dogs, 85% stay.' })), []);
  assert.deepEqual(found(unstatedDetails(deckOf('fine', '85% stay'), new Set([1]))), []);
});

test('stand-ins and text about the slide are told, in English and in Hebrew', () => {
  assert.deepEqual(found(unstatedDetails(deckOf('Growth reached XX% under [name], TBD.'))), ['1:text-stand-in:{"found":"XX, [name], TBD"}']);
  assert.deepEqual(found(unstatedDetails(deckOf('This slide shows our results.', 'הנתונים בשקף זה מראים צמיחה', 'כפי שמוצג בתרשים שלהלן', 'See the chart below for details.'))), [
    '1:text-about-slide:{"found":"This slide"}',
    '2:text-about-slide:{"found":"בשקף זה"}',
    '3:text-about-slide:{"found":"כפי שמוצג בתרשים"}',
    '4:text-about-slide:{"found":"See the chart below"}',
  ]);
  // Words that only look like them are left alone: an XXL size, a footnote [1], a slide in a playground.
  assert.deepEqual(found(unstatedDetails(deckOf('Shirts in XXL, as noted [1]. The new slide in the park is open.'))), []);
  const stand = unstatedDetails(deckOf('[שם הלקוח] יקבל הנחה של XX%'))[0] as Issue;
  // In the order the text has them; the list is isolated, so that its brackets and Latin letters keep their places.
  assert.equal(formatIssue(stand, 'he'), 'במקום ערך מופיע ממלא מקום: ⁨[שם הלקוח], XX⁩. כדאי לכתוב במקומו את הערך האמיתי, או להשמיט אותו.');
});

test('a date is no phone number, and a bracket with a number in it is a reference, not a stand-in', () => {
  assert.deepEqual(found(unstatedDetails(deckOf('The fair opens on 05.11.2026 and closes 01-12-2026 [Smith 2020].'))), []);
  // A number to dial is told where the brief does not give it, in any of its forms.
  assert.deepEqual(found(unstatedDetails(deckOf('Call 03-555-9876.'))), ['1:text-details-unstated:{"details":"03-555-9876","count":1}']);
});
