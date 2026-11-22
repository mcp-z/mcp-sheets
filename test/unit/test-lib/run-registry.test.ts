import assert from 'assert';
import { RunRegistry } from '../../lib/run-registry.ts';

// A real delete function that fails for chosen ids and records every attempt.
function deleter(failing: string[]) {
  const attempts: string[] = [];
  const del = async (id: string) => {
    attempts.push(id);
    if (failing.includes(id)) throw new Error(`503 for ${id}`);
  };
  return { del, attempts };
}

describe('run registry', () => {
  it('cleanup deletes every id no delete was attempted on, and passes when all succeed', async () => {
    const registry = new RunRegistry();
    registry.add('a', 'title-a');
    registry.add('b', 'title-b');
    const { del, attempts } = deleter([]);
    await registry.cleanup(del);
    assert.deepStrictEqual(attempts.sort(), ['a', 'b']);
    assert.strictEqual(registry.size, 0);
  });

  it('a successful delete removes the id, so cleanup does not touch it', async () => {
    const registry = new RunRegistry();
    registry.add('a', 'title-a');
    const { del, attempts } = deleter([]);
    await registry.remove('a', del);
    await registry.cleanup(del);
    assert.deepStrictEqual(attempts, ['a']);
  });

  it('a failed delete is rethrown, kept, never tried again, and named at the end with its id and title', async () => {
    const registry = new RunRegistry();
    registry.add('bad', 'title-bad');
    const { del, attempts } = deleter(['bad']);
    await assert.rejects(() => registry.remove('bad', del), /503 for bad/);
    await assert.rejects(
      () => registry.cleanup(del),
      (error: unknown) => {
        assert.ok(error instanceof Error);
        assert.ok(error.message.includes('bad "title-bad"'), error.message);
        assert.ok(error.message.includes('503 for bad'), error.message);
        return true;
      }
    );
    assert.deepStrictEqual(attempts, ['bad'], 'the failed id must not be attempted a second time');
    assert.strictEqual(registry.size, 1);
  });

  it('cleanup deletes the untried ids, skips the failed one, and fails naming every id left', async () => {
    const registry = new RunRegistry();
    registry.add('failed-earlier', 'first');
    registry.add('untried-ok', 'second');
    registry.add('untried-bad', 'third');
    const { del, attempts } = deleter(['failed-earlier', 'untried-bad']);
    await assert.rejects(() => registry.remove('failed-earlier', del));
    await assert.rejects(
      () => registry.cleanup(del),
      (error: unknown) => {
        assert.ok(error instanceof Error);
        assert.ok(error.message.includes('failed-earlier "first"'), error.message);
        assert.ok(error.message.includes('untried-bad "third"'), error.message);
        assert.ok(!error.message.includes('untried-ok'), error.message);
        return true;
      }
    );
    assert.deepStrictEqual(attempts.sort(), ['failed-earlier', 'untried-bad', 'untried-ok']);
  });
});
