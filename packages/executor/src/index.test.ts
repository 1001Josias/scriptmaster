import { describe, expect, it } from 'vitest';

import { createLogger } from '@scriptmaster/runtime';

import { executeEntry } from './index.js';

describe('executeEntry', () => {
  it('executes an entry function and captures logs and return value', async () => {
    const Logger = createLogger();
    const result = await executeEntry({
      module: {
        Logger,
        main(name: unknown) {
          Logger.log('hello %s', name);
          return { ok: true };
        },
      },
      entryFunction: 'main',
      args: ['Josias'],
    });

    expect(result.status).toBe('succeeded');
    expect(result.returnValue).toEqual({ ok: true });
    expect(result.logs.map((entry) => entry.message)).toEqual(['hello Josias']);
  });

  it('returns a structured diagnostic when the entry does not exist', async () => {
    const result = await executeEntry({ module: {}, entryFunction: 'missing' });

    expect(result).toEqual(
      expect.objectContaining({
        status: 'failed',
        failure: expect.objectContaining({ type: 'entry_not_found' }),
      }),
    );
  });

  it('captures runtime errors without throwing from the executor', async () => {
    const result = await executeEntry({
      module: {
        main() {
          throw new TypeError('broken');
        },
      },
      entryFunction: 'main',
    });

    expect(result.status).toBe('failed');
    expect(result.failure).toEqual(
      expect.objectContaining({ type: 'runtime_error', name: 'TypeError', message: 'broken' }),
    );
  });

  it('returns a timed_out result when execution exceeds the limit', async () => {
    const result = await executeEntry({
      module: {
        async main() {
          await new Promise((resolve) => setTimeout(resolve, 25));
        },
      },
      entryFunction: 'main',
      timeoutMs: 1,
    });

    expect(result.status).toBe('timed_out');
    expect(result.failure).toEqual(expect.objectContaining({ type: 'timeout' }));
  });
});
