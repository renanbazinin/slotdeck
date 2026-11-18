// Issues as the tests written before 0.6.4 compare them: their severity, slide, slot and message.
// Every warning they are given is first checked to carry what 0.6.4 adds: a code, its values, an
// English sentence that is its message, and a Hebrew one. So each scenario of those tests is also
// a check that its warnings have codes (messages.test.ts checks the sentences themselves).
import assert from 'node:assert/strict';
import { ISSUE_CODES, formatIssue } from '../src/index.ts';
import type { Issue } from '../src/index.ts';

/** Fails on a warning without a known code and values, or whose sentences do not come out of them. */
export function assertCoded(issue: Issue): void {
  if (issue.severity !== 'warning') return;
  const said = JSON.stringify(issue);
  assert.ok(issue.code !== undefined && ISSUE_CODES.includes(issue.code), `a warning without a code: ${said}`);
  assert.ok(typeof issue.values === 'object' && issue.values !== null, `a warning without values: ${said}`);
  assert.equal(formatIssue(issue, 'en'), issue.message, `the English sentence of ${issue.code} is not its message`);
  const hebrew = formatIssue(issue, 'he');
  assert.match(hebrew, /^[א-ת]/, `the Hebrew sentence of ${issue.code} does not begin with a Hebrew word: ${hebrew}`);
  assert.doesNotMatch(hebrew, /undefined|NaN|\[object|\$\{/, hebrew);
}

export const withoutCodes = <T extends Issue>(issues: readonly T[] | undefined): Array<Omit<T, 'code' | 'values'>> =>
  (issues ?? []).map((issue) => {
    assertCoded(issue);
    const { code: _code, values: _values, ...rest } = issue;
    return rest;
  });
