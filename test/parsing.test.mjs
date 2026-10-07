import { test } from 'node:test';
import assert from 'node:assert/strict';

import {
  AUTHOR_GROUP_WINDOW_MS, carryForwardAuthors, collectRenderedMessages, emojiArgumentError,
  findSentMessage, isPendingDelivery, isSessionEvidence, parseMessageIds, parsePeriod,
  parsePostMessageArgs, parseReadChatMessagesArgs, renderedMessageIds,
} from '../parsing.mjs';
import { el, fakeDocument } from './fake-dom.mjs';

test('parsePeriod reads each unit', () => {
  assert.equal(parsePeriod('10m'), 10 * 60_000);
  assert.equal(parsePeriod('6h'), 6 * 3_600_000);
  assert.equal(parsePeriod('2d'), 2 * 86_400_000);
});

test('parsePeriod ignores case and the space before the unit', () => {
  assert.equal(parsePeriod('6H'), 6 * 3_600_000);
  assert.equal(parsePeriod('  30 m '), 30 * 60_000);
});

test('parsePeriod refuses a span of zero', () => {
  // "0m" names no messages at all, so it is a typo rather than a request.
  assert.equal(parsePeriod('0m'), null);
  assert.equal(parsePeriod('00h'), null);
});

test('parsePeriod refuses what is not a period', () => {
  for (const value of ['', 'm', '10', '10w', '10 minutes', '-5m', '1.5h', '5m5m', 'now']) {
    assert.equal(parsePeriod(value), null, `expected "${value}" to be refused`);
  }
});

test('parseReadChatMessagesArgs takes the positional arguments in order', () => {
  const args = parseReadChatMessagesArgs(['Developers', '6h', 'out/messages.json']);
  assert.deepEqual(args, {
    chatName: 'Developers',
    period: '6h',
    outputFile: 'out/messages.json',
    withoutReactionsOnly: false,
    unknownFlag: undefined,
  });
});

test('parseReadChatMessagesArgs takes the flag wherever it is given', () => {
  for (const args of [
    ['--without-reactions-only', 'Developers', '6h', 'out.json'],
    ['Developers', '--without-reactions-only', '6h', 'out.json'],
    ['Developers', '6h', 'out.json', '--without-reactions-only'],
  ]) {
    const parsed = parseReadChatMessagesArgs(args);
    assert.equal(parsed.withoutReactionsOnly, true);
    assert.equal(parsed.chatName, 'Developers');
    assert.equal(parsed.period, '6h');
    assert.equal(parsed.outputFile, 'out.json');
    assert.equal(parsed.unknownFlag, undefined);
  }
});

test('parseReadChatMessagesArgs reports a mistyped flag rather than reading it as an argument', () => {
  // The near miss is the case that matters: it must not end up as a positional
  // argument, where it would pass the usage check and read the wrong chat.
  for (const flag of ['--without-reactions', '-without-reactions-only', '--without_reactions_only', '--dry-run']) {
    const parsed = parseReadChatMessagesArgs(['Developers', '6h', 'out.json', flag]);
    assert.equal(parsed.unknownFlag, flag);
    assert.equal(parsed.withoutReactionsOnly, false);
    assert.equal(parsed.outputFile, 'out.json');
  }
});

test('parseReadChatMessagesArgs leaves missing arguments undefined', () => {
  const parsed = parseReadChatMessagesArgs(['Developers']);
  assert.equal(parsed.chatName, 'Developers');
  assert.equal(parsed.period, undefined);
  assert.equal(parsed.outputFile, undefined);
});

test('parsePostMessageArgs takes the chat name and the message', () => {
  assert.deepEqual(parsePostMessageArgs(['Developers', 'ship it?']), {
    chatName: 'Developers',
    message: 'ship it?',
    dryRun: false,
  });
});

// A near miss of --dry-run is issue #3 and is deliberately not pinned here, so
// that fixing it does not have to argue with a test.
test('parsePostMessageArgs keeps --dry-run out of the positional arguments', () => {
  for (const args of [
    ['--dry-run', 'Developers', 'ship it?'],
    ['Developers', '--dry-run', 'ship it?'],
    ['Developers', 'ship it?', '--dry-run'],
  ]) {
    assert.deepEqual(parsePostMessageArgs(args), {
      chatName: 'Developers',
      message: 'ship it?',
      dryRun: true,
    });
  }
});

test('parsePostMessageArgs keeps a message that only looks like a flag', () => {
  // The message is whatever the caller passed; only the exact flag is removed.
  assert.deepEqual(parsePostMessageArgs(['Developers', '--dry-run --dry-run']), {
    chatName: 'Developers',
    message: '--dry-run --dry-run',
    dryRun: false,
  });
});

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

test('carryForwardAuthors attributes a message to the group above it', () => {
  const messages = [
    { ts: 0, author: 'Ada Lovelace' },
    { ts: AUTHOR_GROUP_WINDOW_MS - 1, author: '' },
  ];
  carryForwardAuthors(messages);
  assert.deepEqual(messages.map(m => m.author), ['Ada Lovelace', 'Ada Lovelace']);
});

test('carryForwardAuthors still attributes a message exactly one window away', () => {
  const messages = [
    { ts: 0, author: 'Ada Lovelace' },
    { ts: AUTHOR_GROUP_WINDOW_MS, author: '' },
  ];
  carryForwardAuthors(messages);
  assert.deepEqual(messages.map(m => m.author), ['Ada Lovelace', 'Ada Lovelace']);
});

test('carryForwardAuthors leaves a message past the window unattributed', () => {
  const messages = [
    { ts: 0, author: 'Ada Lovelace' },
    { ts: AUTHOR_GROUP_WINDOW_MS + 1, author: '' },
  ];
  carryForwardAuthors(messages);
  assert.deepEqual(messages.map(m => m.author), ['Ada Lovelace', '']);
});

test('carryForwardAuthors leaves the whole group after a gap unattributed', () => {
  // The header of that group was never collected, so nothing after the gap can
  // be attributed either — naming the author from before it would name the
  // wrong person.
  const messages = [
    { ts: 0, author: 'Ada Lovelace' },
    { ts: AUTHOR_GROUP_WINDOW_MS + 1, author: '' },
    { ts: AUTHOR_GROUP_WINDOW_MS + 2, author: '' },
  ];
  carryForwardAuthors(messages);
  assert.deepEqual(messages.map(m => m.author), ['Ada Lovelace', '', '']);
});

test('carryForwardAuthors starts a new group from its own header', () => {
  const messages = [
    { ts: 0, author: 'Ada Lovelace' },
    { ts: 1000, author: '' },
    { ts: 2000, author: 'Grace Hopper' },
    { ts: 3000, author: '' },
  ];
  carryForwardAuthors(messages);
  assert.deepEqual(
    messages.map(m => m.author),
    ['Ada Lovelace', 'Ada Lovelace', 'Grace Hopper', 'Grace Hopper'],
  );
});

test('carryForwardAuthors leaves a first message with no header unattributed', () => {
  const messages = [{ ts: 1_700_000_000_000, author: '' }];
  carryForwardAuthors(messages);
  assert.equal(messages[0].author, '');
});

// The markup below is what the collector's selectors assume Teams renders: the
// message carries data-mid, its item wraps it, and the timestamp, author and
// content carry ids derived from the message id.
function chatItem(children) {
  return el('div', { 'data-tid': 'chat-pane-item' }, children);
}

function chatMessage(mid, children) {
  return el('div', { 'data-tid': 'chat-pane-message', 'data-mid': mid }, children);
}

function reactionPill() {
  return el('button', { 'data-tid': 'diverse-reaction-pill-button' }, ['👍 1']);
}

test('collectRenderedMessages reads a message from the elements ids point at', () => {
  const doc = fakeDocument([
    chatItem([
      el('div', { id: 'author-1785922526738', 'data-tid': 'message-author-name' }, ['  Ada Lovelace  ']),
      chatMessage('1785922526738', [
        el('time', { id: 'timestamp-1785922526738', datetime: '2026-08-01T09:15:26.738Z' }, ['09:15']),
        el('div', { id: 'content-1785922526738' }, [' ship it? ']),
      ]),
    ]),
  ]);

  assert.deepEqual(collectRenderedMessages(doc), [{
    id: '1785922526738',
    time: '2026-08-01T09:15:26.738Z',
    ts: Date.parse('2026-08-01T09:15:26.738Z'),
    author: 'Ada Lovelace',
    body: 'ship it?',
    links: [],
    hasReactions: false,
  }]);
});

test('collectRenderedMessages falls back to the elements inside the item', () => {
  // Same message without the ids: the timestamp and the author are then found
  // by their own markers within the item, and the content within the message.
  const doc = fakeDocument([
    chatItem([
      el('span', { 'data-tid': 'message-author-name' }, ['Grace Hopper']),
      chatMessage('1785922526738', [
        el('time', { datetime: '2026-08-01T09:15:26.738Z' }, ['09:15']),
        el('div', { 'data-message-content': '' }, ['ship it?']),
      ]),
    ]),
  ]);

  assert.deepEqual(collectRenderedMessages(doc), [{
    id: '1785922526738',
    time: '2026-08-01T09:15:26.738Z',
    ts: Date.parse('2026-08-01T09:15:26.738Z'),
    author: 'Grace Hopper',
    body: 'ship it?',
    links: [],
    hasReactions: false,
  }]);
});

test('collectRenderedMessages falls back to the id for the time', () => {
  // Teams message ids are the send time in epoch milliseconds.
  const doc = fakeDocument([
    chatItem([chatMessage('1785922526738', [el('div', { id: 'content-1785922526738' }, ['no timestamp here'])])]),
  ]);

  const [message] = collectRenderedMessages(doc);
  assert.equal(message.time, new Date(1785922526738).toISOString());
  assert.equal(message.ts, 1785922526738);
});

test('collectRenderedMessages survives an id no date can hold', () => {
  // Date only covers ±8.64e15 ms; a longer numeric id would make toISOString()
  // throw and take the whole run down with it.
  const doc = fakeDocument([
    chatItem([chatMessage('99999999999999999', [el('div', { id: 'content-99999999999999999' }, ['from the far future'])])]),
  ]);

  assert.deepEqual(collectRenderedMessages(doc), [{
    id: '99999999999999999',
    time: '',
    ts: null,
    author: '',
    body: 'from the far future',
    links: [],
    hasReactions: false,
  }]);
});

test('collectRenderedMessages leaves an unparseable time null', () => {
  const doc = fakeDocument([
    chatItem([chatMessage('not-an-epoch', [el('div', { id: 'content-not-an-epoch' }, ['hello'])])]),
  ]);

  const [message] = collectRenderedMessages(doc);
  assert.equal(message.time, '');
  assert.equal(message.ts, null);
});

test('collectRenderedMessages skips a message with no id', () => {
  const doc = fakeDocument([
    chatItem([el('div', { 'data-tid': 'chat-pane-message' }, ['nothing identifies this one'])]),
    chatItem([chatMessage('1785922526738', [el('div', { id: 'content-1785922526738' }, ['this one is read'])])]),
  ]);

  assert.deepEqual(collectRenderedMessages(doc).map(m => m.id), ['1785922526738']);
});

test('collectRenderedMessages keeps the full href next to the shortened link text', () => {
  const doc = fakeDocument([
    chatItem([chatMessage('1785922526738', [
      el('div', { id: 'content-1785922526738' }, [
        'see ',
        el('a', { href: 'https://example.com/a/very/long/path?with=query' }, [' https://example.com/a/… ']),
      ]),
    ])]),
  ]);

  const [message] = collectRenderedMessages(doc);
  assert.deepEqual(message.links, [{
    text: 'https://example.com/a/…',
    href: 'https://example.com/a/very/long/path?with=query',
  }]);
});

test('collectRenderedMessages reports a reaction pill on the message that carries it', () => {
  const doc = fakeDocument([
    chatItem([chatMessage('1785922526738', [
      el('div', { id: 'content-1785922526738' }, ['reacted to']),
      reactionPill(),
    ])]),
    chatItem([chatMessage('1785922526739', [
      el('div', { id: 'content-1785922526739' }, ['not reacted to']),
    ])]),
  ]);

  assert.deepEqual(
    collectRenderedMessages(doc).map(m => [m.id, m.hasReactions]),
    [['1785922526738', true], ['1785922526739', false]],
  );
});

test('collectRenderedMessages reads a message whose content is missing', () => {
  const doc = fakeDocument([chatItem([chatMessage('1785922526738', [])])]);

  const [message] = collectRenderedMessages(doc);
  assert.equal(message.body, '');
  assert.deepEqual(message.links, []);
});

test('collectRenderedMessages returns nothing for a pane with no messages', () => {
  assert.deepEqual(collectRenderedMessages(fakeDocument([])), []);
});

test('collectRenderedMessages survives being serialised into the page', () => {
  // page.evaluate() sends this function's source, not the module around it, so
  // a reference to anything outside its own body — AUTHOR_GROUP_WINDOW_MS one
  // declaration above it is the natural reach — compiles, imports and passes
  // every test that calls it here, then throws in the page at runtime.
  // new Function gives it exactly the bare global scope the page does, so such
  // a reference fails here instead.
  const inPage = new Function(`return (${collectRenderedMessages});`)();
  const doc = fakeDocument([
    chatItem([
      el('div', { id: 'author-1785922526738' }, ['Ada Lovelace']),
      chatMessage('1785922526738', [
        el('time', { id: 'timestamp-1785922526738', datetime: '2026-08-01T09:15:26.738Z' }, ['09:15']),
        el('div', { id: 'content-1785922526738' }, [
          'see ',
          el('a', { href: 'https://example.com/a/very/long/path' }, ['https://example.com/a/…']),
        ]),
        reactionPill(),
      ]),
    ]),
  ]);

  assert.deepEqual(inPage(doc), collectRenderedMessages(doc));
});

test('collectRenderedMessages reads the global document when given none', () => {
  // The default argument is the only form the live command uses: page.evaluate()
  // calls the function with nothing, and it reads the document of the page it
  // landed in. Standing one up as the global is that same call.
  const doc = fakeDocument([
    chatItem([chatMessage('1785922526738', [el('div', { id: 'content-1785922526738' }, ['ship it?'])])]),
  ]);

  globalThis.document = doc;
  try {
    assert.deepEqual(collectRenderedMessages(), collectRenderedMessages(doc));
  } finally {
    delete globalThis.document;
  }
});

test('isSessionEvidence accepts a successful answer from the chat service', () => {
  // The requests a signed-in client makes while it boots, as a live run made them.
  for (const url of [
    'https://eu-api.asm.skype.com/v1/skypetokenauth',
    'https://eu-prod.asyncgw.teams.microsoft.com/v1/skypetokenauth',
    'https://teams.cloud.microsoft/api/chatsvc/emea/v1/threads/19%3Aabc%40thread.v2?view=msnp24Equivalent',
    'https://teams.cloud.microsoft/api/chatsvc/emea/v1/users/ME/conversations/48%3Anotes/messages',
  ]) {
    assert.equal(isSessionEvidence(url, 200), true, url);
    assert.equal(isSessionEvidence(url, 204), true, url);
  }
});

test('isSessionEvidence refuses a failed answer from the chat service', () => {
  // A 404 for a conversation that does not exist is still the service answering
  // a signed-in client, but nothing here needs to lean on that; and a 401 is the
  // very thing being looked for.
  for (const status of [301, 401, 403, 404, 500]) {
    assert.equal(isSessionEvidence('https://teams.cloud.microsoft/api/chatsvc/emea/v1/threads/x', status), false);
    assert.equal(isSessionEvidence('https://eu-api.asm.skype.com/v1/skypetokenauth', status), false);
  }
});

test('isSessionEvidence refuses what an expired session still gets answered', () => {
  // An expired session still loads the client, its config and its telemetry,
  // and the profile pictures cached behind a cookie — none of which needs a
  // working token.
  for (const url of [
    'https://teams.microsoft.com/v2/',
    'https://config.teams.microsoft.com/config/v1/MicrosoftTeams/1415_1.0.0.0',
    'https://eu-teams.events.data.microsoft.com/OneCollector/1.0/',
    'https://teams.cloud.microsoft/api/mt/part/emea-02/beta/users/x/profilepicturev2/8:orgid:x',
    'https://login.microsoftonline.com/tenant/oauth2/v2.0/token',
    'https://example.com/?next=/api/chatsvc/',
    'not a url',
  ]) {
    assert.equal(isSessionEvidence(url, 200), false, url);
  }
});

test('isPendingDelivery recognises the statuses of a message still on its way', () => {
  for (const status of ['Sending...', 'Retrying...', 'sending', '  Retrying…']) {
    assert.equal(isPendingDelivery(status), true, status);
  }
  for (const status of ['Sent', 'Seen', 'Read by Ada Lovelace', '', null, undefined]) {
    assert.equal(isPendingDelivery(status), false, String(status));
  }
});

// A message of your own as Teams renders it, with the delivery status it puts on
// the newest one of them.
function ownMessage(mid, { status = null, failed = false } = {}) {
  return chatItem([
    chatMessage(mid, [el('div', { id: `content-${mid}` }, ['ship it?'])]),
    ...(status === null ? [] : [el('span', { id: `read-status-icon-${mid}`, 'aria-label': status, role: 'img' })]),
    ...(failed ? [el('div', { id: `failure-status-${mid}` }, ['Failed to send'])] : []),
  ]);
}

test('renderedMessageIds lists the rendered messages that have an id', () => {
  const doc = fakeDocument([
    ownMessage('1785922526738'),
    chatItem([el('div', { 'data-tid': 'chat-pane-message' }, ['no id'])]),
    ownMessage('1785922526739', { status: 'Sent' }),
  ]);
  assert.deepEqual(renderedMessageIds(doc), ['1785922526738', '1785922526739']);
});

test('findSentMessage finds the new message and reads its status', () => {
  const doc = fakeDocument([
    ownMessage('1785922526738'),
    ownMessage('2730311426311008073', { status: 'Retrying...' }),
  ]);
  assert.deepEqual(findSentMessage(['1785922526738'], doc), {
    mid: '2730311426311008073', status: 'Retrying...', failed: false,
  });
});

test('findSentMessage ignores the messages that were there before', () => {
  // The previous message of your own carries the status until the new one
  // takes it over, and must not be read as the new one delivered.
  const doc = fakeDocument([ownMessage('1785922526738', { status: 'Sent' })]);
  assert.equal(findSentMessage(['1785922526738'], doc), null);
});

test('findSentMessage ignores a new message without a delivery status', () => {
  // Someone else's message arriving meanwhile carries none.
  const doc = fakeDocument([
    ownMessage('1785922526738', { status: 'Sending...' }),
    ownMessage('1785922526740'),
  ]);
  assert.deepEqual(findSentMessage([], doc), { mid: '1785922526738', status: 'Sending...', failed: false });
});

test('findSentMessage reports a message marked as not sent', () => {
  const doc = fakeDocument([ownMessage('2730311426311008073', { failed: true })]);
  assert.deepEqual(findSentMessage([], doc), { mid: '2730311426311008073', status: null, failed: true });
});

test('findSentMessage returns nothing before the message renders', () => {
  assert.equal(findSentMessage(['1785922526738'], fakeDocument([ownMessage('1785922526738')])), null);
});

test('findSentMessage and renderedMessageIds survive being serialised into the page', () => {
  // Both are sent to the page by page.evaluate(), as collectRenderedMessages is.
  const doc = fakeDocument([
    ownMessage('1785922526738'),
    ownMessage('1785922526739', { status: 'Sent' }),
  ]);
  const ids = new Function(`return (${renderedMessageIds});`)();
  const sent = new Function(`return (${findSentMessage});`)();
  assert.deepEqual(ids(doc), renderedMessageIds(doc));
  assert.deepEqual(sent(['1785922526738'], doc), findSentMessage(['1785922526738'], doc));
});
