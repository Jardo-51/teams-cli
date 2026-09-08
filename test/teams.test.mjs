import { test } from 'node:test';
import assert from 'node:assert/strict';

import { emojiArgumentError, parseMessageIds } from '../teams.mjs';

test('parseMessageIds reads a single id', () => {
  assert.deepEqual(parseMessageIds('1785922526738'), { ids: ['1785922526738'] });
});

test('parseMessageIds reads a comma-separated list', () => {
  assert.deepEqual(
    parseMessageIds('1785922526738, 1785922526739 ,1785922526740'),
    { ids: ['1785922526738', '1785922526739', '1785922526740'] },
  );
});

test('parseMessageIds drops blank entries and repeats', () => {
  // A trailing or doubled comma says nothing about which messages are meant,
  // and a repeated id would only find what its first turn left behind.
  assert.deepEqual(
    parseMessageIds('1785922526738,,1785922526739,1785922526738,'),
    { ids: ['1785922526738', '1785922526739'] },
  );
});

test('parseMessageIds refuses a list with no id in it', () => {
  for (const list of ['', '   ', ',', ' , ']) {
    assert.match(parseMessageIds(list).error, /No message id/);
  }
});

test('parseMessageIds refuses an id that could break out of a selector', () => {
  // The ids end up inside CSS attribute selectors, so a quote or a bracket in
  // one is refused rather than escaped.
  for (const id of ['17859"2526738', "17859'2526738", '1785922526738]', '178592 2526738', '#message']) {
    const { ids, error } = parseMessageIds(id);
    assert.equal(ids, undefined, `expected "${id}" to be refused`);
    assert.match(error, /Invalid message id/);
  }
});

test('parseMessageIds accepts the punctuation Teams ids can carry', () => {
  assert.deepEqual(
    parseMessageIds('1785922526738:1.2-3_4'),
    { ids: ['1785922526738:1.2-3_4'] },
  );
});

test('emojiArgumentError accepts an emoji character', () => {
  for (const emoji of ['👍', '👀', '🎉', '❤️']) {
    assert.equal(emojiArgumentError(emoji), null, `expected "${emoji}" to be accepted`);
  }
});

test('emojiArgumentError refuses an emoji name rather than the character', () => {
  // Plain text would otherwise only be refused minutes later, once the browser
  // is open and the picker has been walked.
  for (const name of ['thumbsup', ':+1:', 'eyes']) {
    assert.match(emojiArgumentError(name), /the emoji character itself/);
  }
});

test('emojiArgumentError refuses what could break out of a selector', () => {
  for (const emoji of ['👍"', "👍'", '👍\\']) {
    assert.match(emojiArgumentError(emoji), /expected a single emoji character/);
  }
});
