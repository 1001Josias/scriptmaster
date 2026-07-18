import { describe, expect, it, vi } from 'vitest';

import {
  SpreadsheetCompatibilityError,
  UnsupportedSpreadsheetMethodError,
  createSpreadsheetApp,
  type SheetsApiClient,
} from './spreadsheet.js';

function createClient(): SheetsApiClient {
  return {
    spreadsheets: {
      get: vi.fn().mockResolvedValue({
        data: { sheets: [{ properties: { title: 'Data' } }] },
      }),
      values: {
        get: vi.fn().mockResolvedValue({ data: { values: [['name'], ['Ada']] } }),
        update: vi.fn().mockResolvedValue({}),
        append: vi.fn().mockResolvedValue({}),
      },
    },
  };
}

describe('SpreadsheetApp compatibility', () => {
  it('reads and writes values using GAS-style calls', async () => {
    const client = createClient();
    const SpreadsheetApp = createSpreadsheetApp(client);
    const spreadsheet = SpreadsheetApp.openById('sheet-123');
    const sheet = await spreadsheet.getSheetByName('Data');

    expect(sheet).not.toBeNull();
    const range = sheet!.getRange('A1:A2');

    await expect(range.getValues()).resolves.toEqual([['name'], ['Ada']]);
    await range.setValues([['name'], ['Grace']]);
    await sheet!.appendRow(['Linus', 42, true]);

    expect(client.spreadsheets.values.get).toHaveBeenCalledWith({
      spreadsheetId: 'sheet-123',
      range: 'Data!A1:A2',
    });
    expect(client.spreadsheets.values.update).toHaveBeenCalledWith({
      spreadsheetId: 'sheet-123',
      range: 'Data!A1:A2',
      valueInputOption: 'USER_ENTERED',
      requestBody: { values: [['name'], ['Grace']] },
    });
    expect(client.spreadsheets.values.append).toHaveBeenCalledWith({
      spreadsheetId: 'sheet-123',
      range: 'Data',
      valueInputOption: 'USER_ENTERED',
      insertDataOption: 'INSERT_ROWS',
      requestBody: { values: [['Linus', 42, true]] },
    });
  });

  it('returns null when the requested sheet does not exist', async () => {
    const client = createClient();
    const SpreadsheetApp = createSpreadsheetApp(client);

    await expect(
      SpreadsheetApp.openById('sheet-123').getSheetByName('Missing'),
    ).resolves.toBeNull();
  });

  it('normalizes Google API errors', async () => {
    const client = createClient();
    vi.mocked(client.spreadsheets.values.get).mockRejectedValueOnce(new Error('permission denied'));
    const SpreadsheetApp = createSpreadsheetApp(client);
    const sheet = await SpreadsheetApp.openById('sheet-123').getSheetByName('Data');

    await expect(sheet!.getRange('A1').getValues()).rejects.toEqual(
      expect.objectContaining({
        name: 'SpreadsheetCompatibilityError',
        message: expect.stringContaining('permission denied'),
      }),
    );
  });

  it('rejects invalid arguments and unsupported methods explicitly', () => {
    const SpreadsheetApp = createSpreadsheetApp(createClient());

    expect(() => SpreadsheetApp.openById('')).toThrow(SpreadsheetCompatibilityError);
    expect(() =>
      (SpreadsheetApp as unknown as { create(): void }).create(),
    ).toThrow(UnsupportedSpreadsheetMethodError);
  });
});
