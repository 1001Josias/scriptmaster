import { describe, expect, it, vi } from 'vitest';

import { createRuntimeBindings } from './backend.js';
import { createGasFakesBackend } from './gas-fakes.js';

describe('gas-fakes backend', () => {
  it('loads gas-fakes once and exposes only requested globals', async () => {
    const globalObject: Record<string, unknown> = {};
    const load = vi.fn(async () => {
      globalObject.SpreadsheetApp = { openById: vi.fn() };
      globalObject.DriveApp = { getFiles: vi.fn() };
    });
    const backend = createGasFakesBackend({ load, globalObject });

    const first = await createRuntimeBindings({
      backend,
      services: ['SpreadsheetApp', 'SpreadsheetApp'],
    });
    const second = await createRuntimeBindings({
      backend,
      services: ['DriveApp'],
    });

    expect(load).toHaveBeenCalledTimes(1);
    expect(first.SpreadsheetApp).toBe(globalObject.SpreadsheetApp);
    expect(first.DriveApp).toBeUndefined();
    expect(second.DriveApp).toBe(globalObject.DriveApp);
  });

  it('fails when the backend does not expose a requested service', async () => {
    const backend = createGasFakesBackend({
      load: async () => undefined,
      globalObject: {},
    });

    await expect(
      createRuntimeBindings({ backend, services: ['SpreadsheetApp'] }),
    ).rejects.toThrow('did not provide the requested SpreadsheetApp binding');
  });
});
