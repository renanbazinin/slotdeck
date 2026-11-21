// The rules every writer is given about what it may not invent (0.7.0), and the rules for a deck that
// is read on its own rather than presented (`delivery: 'read'`). No test calls a model.
import assert from 'node:assert/strict';
import { test } from 'node:test';
import { loadTemplate, memoryStore, presentationTools } from '../src/index.ts';
import type { JsonRequest, LlmProvider } from '../src/index.ts';
import { systemPrompt } from '../src/plan/write.ts';
import { buildFixturePptx, fixtureManifest } from './fixture.ts';

test('every writer is told not to invent the organisation\'s procedures, terms or contacts, nor to write about the slide', () => {
  for (const [source, document, given] of [
    ['brief', false, 'the brief do not give'],
    ['brief', true, 'the brief and the document do not give'],
    ['deck', false, 'the request and the deck as it is do not give'],
    ['deck', true, 'the request, the document and the deck as it is do not give'],
  ] as const) {
    const prompt = systemPrompt('he', source, document);
    assert.ok(prompt.includes(`- Never state the organisation's own procedures, terms, prices, contacts (people, phone numbers, email or web addresses), results or plans that ${given}.`), prompt);
    assert.ok(prompt.includes('- Write the content itself. Never write about the slide or its layout as slide text'));
    // Presented, as before: short phrases, and nothing about a closing slide.
    assert.ok(prompt.includes('a phrase beats a sentence') && !prompt.includes('is read, not presented'));
  }
});

test('a deck that is read is written in full sentences, with no slide of thanks and an agenda only when it is long', () => {
  const read = systemPrompt('en', 'brief', true, 'read');
  assert.ok(read.includes('elsewhere write full sentences: the deck is read on its own, with no one to present it.') && !read.includes('a phrase beats a sentence'));
  assert.ok(read.includes('- The deck is read, not presented: no closing slide of thanks or questions, and an agenda only in a deck of more than ten slides.'));
  assert.equal(systemPrompt('en', 'brief', false, 'present'), systemPrompt('en', 'brief'));
});

test('tools set once to decks that are read tell every planning call so', async () => {
  const requests: JsonRequest[] = [];
  const llm: LlmProvider = {
    async generateJson(request) {
      requests.push(request);
      throw new Error('stop after the first call');
    },
  };
  const template = loadTemplate(buildFixturePptx(), fixtureManifest);
  for (const delivery of ['read', undefined] as const) {
    const tools = presentationTools({ template, llm, store: memoryStore(), ...(delivery ? { delivery } : {}) });
    await tools.call('plan_presentation', { topic: 'Our branch hours' });
  }
  assert.equal(requests.length, 2);
  assert.match(requests[0]?.system ?? '', /The deck is read, not presented/);
  assert.doesNotMatch(requests[1]?.system ?? '', /The deck is read, not presented/);
});
