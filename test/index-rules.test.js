const test = require('node:test');
const assert = require('node:assert/strict');

const {
  extractStyleId,
  imageRecordMatchesSource,
  retryableDeltaSkips,
  shouldAdvanceDeltaState
} = require('../server');

test('extractStyleId preserves distinct compact and hyphenated variants', () => {
  const cases = new Map([
    ['MY30089-4 black (1).jpg', 'MY30089-4'],
    ['MY30230C 粉色.jpg', 'MY30230C'],
    ['MB40482B-blue.png', 'MB40482B'],
    ['MY30216-2B (12).JPG', 'MY30216-2B'],
    ['MY30216-2L detail.webp', 'MY30216-2L'],
    ['MC20181-D gold.jpg', 'MC20181-D'],
    ['MY30230.jpg', 'MY30230']
  ]);

  for (const [fileName, expected] of cases) {
    assert.equal(extractStyleId(fileName), expected, fileName);
  }
});

test('image records dedupe by stable drive id or normalized file name', () => {
  assert.equal(imageRecordMatchesSource(
    { driveId: '01ABC', driveName: 'old-name.jpg' },
    { driveId: '01abc', name: 'renamed.jpg' }
  ), true);
  assert.equal(imageRecordMatchesSource(
    { driveName: 'MY30089-4 BLACK (1).JPG' },
    { name: 'my30089-4 black (1).jpg' }
  ), true);
  assert.equal(imageRecordMatchesSource(
    { driveId: 'first', driveName: 'a.jpg' },
    { driveId: 'second', name: 'b.jpg' }
  ), false);
});

test('transient and safety-limit skips remain retryable', () => {
  const deferred = retryableDeltaSkips([
    { reason: 'add_limit' },
    { reason: 'no_download_url' },
    { reason: 'download_503' },
    { reason: 'already_exists' }
  ]);
  assert.deepEqual(deferred.map(item => item.reason), [
    'add_limit',
    'no_download_url',
    'download_503'
  ]);
});

test('delta cursor advances only after every discovered image is handled', () => {
  assert.equal(shouldAdvanceDeltaState({
    finalDeltaLink: 'delta:next',
    nextLink: '',
    errors: [],
    skipped: []
  }), true);
  assert.equal(shouldAdvanceDeltaState({
    finalDeltaLink: 'delta:next',
    nextLink: '',
    errors: [],
    skipped: [{ reason: 'add_limit' }]
  }), false);
  assert.equal(shouldAdvanceDeltaState({
    finalDeltaLink: 'delta:next',
    nextLink: 'page:2',
    errors: [],
    skipped: []
  }), false);
  assert.equal(shouldAdvanceDeltaState({
    finalDeltaLink: 'delta:next',
    nextLink: '',
    errors: [{ error: 'Vertex timeout' }],
    skipped: []
  }), false);
});
