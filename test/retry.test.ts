// Trying a model call again after a failure that passes by itself (0.7.0): which failures, how long
// to wait, and the deadline of a call. No test calls a model; a fake fetch answers.
import assert from 'node:assert/strict';
import { test } from 'node:test';
import { GeminiError, gemini, geminiImages, retryDelay } from '../src/index.ts';

const fixed = () => 0.5;

test('retryDelay: passing failures wait a doubling step with chance in it, or what the service asks; others are not tried again', () => {
  const overloaded = new GeminiError('Gemini request failed (503): overloaded', 503);
  // From one second, doubled with each try, between half and all of it, up to a minute.
  assert.deepEqual([1, 2, 3].map((attempt) => retryDelay(overloaded, attempt, () => 0)), [500, 1000, 2000]);
  assert.deepEqual([1, 2, 3].map((attempt) => retryDelay(overloaded, attempt, () => 1)), [1000, 2000, 4000]);
  assert.equal(retryDelay(overloaded, 20, () => 1), 60000);
  for (const status of [408, 429, 500, 502, 503, 504]) assert.equal(retryDelay(new GeminiError('x', status), 1, fixed), 750, String(status));
  assert.equal(retryDelay(new GeminiError('Request blocked due to copyright/recitation content.', 400), 1, fixed), 750);
  assert.equal(retryDelay(new GeminiError('Gemini did not answer within 180 seconds', undefined, { timedOut: true }), 1, fixed), 750);
  assert.equal(retryDelay(new TypeError('fetch failed'), 1, fixed), 750, 'a network that failed');
  // What fails again however often it is sent: a bad request or key, a spending cap or a daily quota, a stop.
  for (const [status, message] of [
    [400, 'Schema is too complex'],
    [401, 'API key not valid'],
    [501, 'Not implemented'],
    [429, 'Your project has exceeded its monthly spending cap'],
    [429, 'You exceeded your current quota, please check your plan and billing details.'],
    [429, 'Quota exceeded for metric: generate_content_requests_per_day, limit: 0'],
    [429, 'Quota exceeded for quota metric GenerateRequestsPerDayPerProjectPerModel'],
  ] as const) {
    assert.equal(retryDelay(new GeminiError(message, status), 1, fixed), undefined, message);
  }
  assert.equal(retryDelay(new DOMException('stopped', 'AbortError'), 1, fixed), undefined);
  assert.equal(retryDelay(new Error('anything else'), 1, fixed), undefined);
  // The service's own wait: as asked, and no more tries where it asks for more than a minute.
  assert.equal(retryDelay(new GeminiError('slow down', 429, { retryAfterMs: 2500 }), 1, fixed), 2500);
  assert.equal(retryDelay(new GeminiError('slow down', 429, { retryAfterMs: 3_600_000 }), 1, fixed), undefined);
});

/** A fetch that answers from a list, the last answer again when the list ends, and counts the calls. */
function answers(...list: Array<[status: number, body: unknown, headers?: Record<string, string>]>) {
  let calls = 0;
  const fetch = (async () => {
    const [status, body, headers] = list[Math.min(calls, list.length - 1)] as [number, unknown, Record<string, string> | undefined];
    calls++;
    return new Response(JSON.stringify(body), { status, headers: headers ?? {} });
  }) as typeof globalThis.fetch;
  return { fetch, calls: () => calls };
}
const text = { status: 'completed', steps: [{ type: 'model_output', content: [{ type: 'text', text: '{"ok":true}' }] }] };
const picture = { status: 'completed', steps: [{ type: 'model_output', content: [{ type: 'image', data: btoa('png'), mime_type: 'image/png' }] }] };
const request = { system: 's', prompt: 'p', schema: {} };

test('a rate limit waits as long as the service asks, in its header or its RetryInfo, and pictures are tried again too', async () => {
  const limited = answers([429, { error: { message: 'Resource has been exhausted (e.g. check quota).' } }, { 'retry-after': '0' }], [200, text]);
  assert.deepEqual(await gemini({ model: 'm', fetch: limited.fetch }).generateJson(request), { ok: true });
  assert.equal(limited.calls(), 2);
  const told = answers([429, { error: { message: 'Resource exhausted', details: [{ '@type': 'type.googleapis.com/google.rpc.RetryInfo', retryDelay: '0s' }] } }], [200, text]);
  assert.deepEqual(await gemini({ model: 'm', fetch: told.fetch }).generateJson(request), { ok: true });
  assert.equal(told.calls(), 2);
  // A picture that failed once is made on the second try; before 0.7.0 the call failed.
  const images = answers([503, { error: { message: 'overloaded' } }, { 'retry-after': '0' }], [200, picture]);
  assert.deepEqual(await geminiImages({ model: 'i', fetch: images.fetch }).generateImage({ prompt: 'a dog', kind: 'photo', aspectRatio: 1 }), new TextEncoder().encode('png'));
  assert.equal(images.calls(), 2);
  // A spending cap is not tried again, whatever wait it names.
  const capped = answers([429, { error: { message: 'Your project has exceeded its monthly spending cap' } }, { 'retry-after': '0' }], [200, text]);
  await assert.rejects(gemini({ model: 'm', fetch: capped.fetch }).generateJson(request), /spending cap/);
  assert.equal(capped.calls(), 1);
});

test('a deadline ends a call, its tries and its waits together, and a stop ends a wait at once', async () => {
  // The service asks for two seconds; a deadline of one second leaves no room for the wait, so the first failure is the answer.
  const slow = answers([503, { error: { message: 'overloaded' } }, { 'retry-after': '2' }], [200, text]);
  const started = Date.now();
  await assert.rejects(gemini({ model: 'm', fetch: slow.fetch, deadlineMs: 1000 }).generateJson(request), /overloaded/);
  assert.equal(slow.calls(), 1);
  assert.ok(Date.now() - started < 900, 'no wait was begun');
  // Without the deadline the wait is waited; a stop from the caller ends it.
  const stop = new AbortController();
  const waiting = gemini({ model: 'm', fetch: answers([503, { error: { message: 'overloaded' } }, { 'retry-after': '30' }]).fetch }).generateJson({ ...request, signal: stop.signal });
  setTimeout(() => stop.abort(new Error('the person left')), 50);
  await assert.rejects(waiting, /the person left/);
  assert.ok(Date.now() - started < 5000);
  // A try is given no more time than the deadline leaves: a request that hangs ends with it.
  const hanging = (async (_url: string, init: RequestInit) =>
    new Promise<Response>((_resolve, reject) => init.signal?.addEventListener('abort', () => reject(init.signal?.reason)))) as unknown as typeof globalThis.fetch;
  const before = Date.now();
  await assert.rejects(gemini({ model: 'm', fetch: hanging, deadlineMs: 300 }).generateJson(request), /did not answer/);
  assert.ok(Date.now() - before < 2000, String(Date.now() - before));
});
