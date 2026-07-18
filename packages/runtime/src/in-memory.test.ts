import { describe, expect, it } from 'vitest';

import { createRuntimeBindings } from './backend.js';
import { createInMemoryRuntimeBackend } from './in-memory.js';

describe('in-memory runtime backend', () => {
  it('reads deterministic spreadsheet fixtures through synchronous Apps Script APIs', async () => {
    const backend = createInMemoryRuntimeBackend()
      .spreadsheet('orders')
      .sheet('Orders')
      .values([
        ['id', 'amount'],
        ['A-100', 120],
        ['A-101', 80],
      ]);

    const bindings = await createRuntimeBindings({
      backend,
      services: ['SpreadsheetApp'],
    });

    const values = bindings.SpreadsheetApp!
      .openById('orders')
      .getSheetByName('Orders')!
      .getRange('A1:B3')
      .getValues();

    expect(values).toEqual([
      ['id', 'amount'],
      ['A-100', 120],
      ['A-101', 80],
    ]);
  });

  it('supports setValues and appendRow while keeping fixture reads isolated', async () => {
    const backend = createInMemoryRuntimeBackend()
      .spreadsheet('orders')
      .sheet('Orders')
      .values([['id', 'amount']]);

    const bindings = await createRuntimeBindings({
      backend,
      services: ['SpreadsheetApp'],
    });
    const sheet = bindings.SpreadsheetApp!.openById('orders').getSheetByName('Orders')!;

    sheet.getRange('A2:B2').setValues([['A-100', 120]]);
    sheet.appendRow(['A-101', 80]);

    const snapshot = backend.getValues('orders', 'Orders');
    expect(snapshot).toEqual([
      ['id', 'amount'],
      ['A-100', 120],
      ['A-101', 80],
    ]);

    snapshot[1][1] = 999;
    expect(backend.getValues('orders', 'Orders')[1][1]).toBe(120);
  });

  it('rejects invalid ranges and setValues dimensions', async () => {
    const backend = createInMemoryRuntimeBackend()
      .spreadsheet('orders')
      .sheet('Orders')
      .values([]);
    const bindings = await createRuntimeBindings({
      backend,
      services: ['SpreadsheetApp'],
    });
    const sheet = bindings.SpreadsheetApp!.openById('orders').getSheetByName('Orders')!;

    expect(() => sheet.getRange('not-a-range')).toThrow('Unsupported A1 notation');
    expect(() => sheet.getRange('A1:B2').setValues([[1, 2]])).toThrow(
      'setValues expected 2x2 values',
    );
  });
});
