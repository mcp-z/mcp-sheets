import assert from 'assert';
import { isInvalidRequest } from '../../../src/lib/google-errors.ts';

describe('isInvalidRequest', () => {
  it('is true for a Google 400 however the client reports the status', () => {
    assert.strictEqual(isInvalidRequest({ response: { status: 400 } }), true);
    assert.strictEqual(isInvalidRequest({ status: 400 }), true);
    assert.strictEqual(isInvalidRequest({ code: 400 }), true);
    assert.strictEqual(isInvalidRequest({ code: '400' }), true);
  });

  it('is false for provider failures and unknown shapes', () => {
    for (const status of [401, 403, 404, 429, 500, 502, 503]) {
      assert.strictEqual(isInvalidRequest({ response: { status } }), false, `status ${status}`);
    }
    assert.strictEqual(isInvalidRequest({ code: 'ECONNRESET' }), false);
    assert.strictEqual(isInvalidRequest(new Error('every write failed')), false);
    assert.strictEqual(isInvalidRequest('400'), false);
    assert.strictEqual(isInvalidRequest(undefined), false);
  });
});
