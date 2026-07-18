import { describe, expect, it, vi } from 'vitest';

import { createLogger } from './index.js';

describe('createLogger', () => {
  it('captures string logs', () => {
    const logger = createLogger();

    logger.log('hello');

    expect(logger.getEntries()).toEqual([
      {
        level: 'info',
        message: 'hello',
        values: ['hello'],
      },
    ]);
  });

  it('preserves useful formatting for objects and placeholders', () => {
    const logger = createLogger();

    logger.log('user: %s', 'Josias');
    logger.log({ status: 'ok', count: 2 });

    expect(logger.getEntries().map((entry) => entry.message)).toEqual([
      'user: Josias',
      "{ status: 'ok', count: 2 }",
    ]);
  });

  it('captures multiple calls, exposes entries and supports an execution sink', () => {
    const onEntry = vi.fn();
    const logger = createLogger(onEntry);

    logger.log('first');
    logger.log('second', 2);

    expect(logger.getEntries()).toHaveLength(2);
    expect(onEntry).toHaveBeenCalledTimes(2);

    logger.clear();
    expect(logger.getEntries()).toEqual([]);
  });

  it('returns snapshots that cannot mutate the captured state', () => {
    const logger = createLogger();
    const value = { nested: true };

    logger.log(value);
    const snapshot = logger.getEntries();
    (snapshot[0]?.values as unknown[]).push('changed');

    expect(logger.getEntries()[0]?.values).toEqual([value]);
  });
});
