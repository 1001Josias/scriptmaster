import { describe, expect, it } from 'vitest';

import { analyzeAppsScript, generateCompatibilityReport } from './index.js';

describe('analyzeAppsScript', () => {
  it('detects supported services, methods, instance methods and trigger functions', () => {
    const result = analyzeAppsScript(`function onOpen() {
  const sheet = SpreadsheetApp.openById('sheet-id');
  Logger.log(sheet.getName());
}`);

    expect(result.diagnostics).toEqual([]);
    expect(result.symbols).toEqual([
      {
        kind: 'trigger',
        name: 'onOpen',
        supported: true,
        location: { line: 1, column: 10 },
      },
      {
        kind: 'service',
        name: 'SpreadsheetApp',
        supported: true,
        location: { line: 2, column: 17 },
      },
      {
        kind: 'method',
        name: 'openById',
        service: 'SpreadsheetApp',
        supported: true,
        location: { line: 2, column: 32 },
      },
      {
        kind: 'service',
        name: 'Logger',
        supported: true,
        location: { line: 3, column: 3 },
      },
      {
        kind: 'method',
        name: 'log',
        service: 'Logger',
        supported: true,
        location: { line: 3, column: 10 },
      },
      {
        kind: 'method',
        name: 'getName',
        service: 'SpreadsheetApp',
        receiverType: 'Spreadsheet',
        supported: true,
        location: { line: 3, column: 20 },
      },
    ]);
  });

  it('resolves chained and assigned SpreadsheetApp receiver types', () => {
    const result = analyzeAppsScript(`function main() {
  const spreadsheet = SpreadsheetApp.openById('sheet-id');
  const sheet = spreadsheet.getSheetByName('Data');
  const values = sheet.getRange('A1:B2').getValues();
  sheet.appendRow(['x', 1]);
  return values;
}`);

    expect(result.symbols).toEqual(
      expect.arrayContaining([
        expect.objectContaining({
          kind: 'method',
          name: 'getSheetByName',
          service: 'SpreadsheetApp',
          receiverType: 'Spreadsheet',
        }),
        expect.objectContaining({
          kind: 'method',
          name: 'getRange',
          service: 'SpreadsheetApp',
          receiverType: 'Sheet',
        }),
        expect.objectContaining({
          kind: 'method',
          name: 'getValues',
          service: 'SpreadsheetApp',
          receiverType: 'Range',
        }),
        expect.objectContaining({
          kind: 'method',
          name: 'appendRow',
          service: 'SpreadsheetApp',
          receiverType: 'Sheet',
        }),
      ]),
    );

    const report = generateCompatibilityReport(result);
    for (const method of ['getSheetByName', 'getRange', 'getValues', 'appendRow']) {
      expect(report.items).toContainEqual(
        expect.objectContaining({ name: method, status: 'supported' }),
      );
    }
  });

  it('represents duplicate usages as separate consistently ordered occurrences', () => {
    const result = analyzeAppsScript(`DriveApp.getFiles();
DriveApp.getFolders();`);

    expect(result.symbols.filter((symbol) => symbol.kind === 'service')).toEqual([
      expect.objectContaining({ name: 'DriveApp', location: { line: 1, column: 1 } }),
      expect.objectContaining({ name: 'DriveApp', location: { line: 2, column: 1 } }),
    ]);
  });

  it('detects unknown Apps Script-like services as unsupported', () => {
    const result = analyzeAppsScript('UnknownApp.execute();');

    expect(result.symbols).toEqual([
      expect.objectContaining({
        kind: 'service',
        name: 'UnknownApp',
        supported: false,
      }),
      expect.objectContaining({
        kind: 'method',
        name: 'execute',
        service: 'UnknownApp',
        supported: false,
      }),
    ]);
  });

  it('returns structured syntax diagnostics without throwing', () => {
    const result = analyzeAppsScript('function broken( {');

    expect(result.diagnostics.length).toBeGreaterThan(0);
    expect(result.diagnostics[0]).toEqual(
      expect.objectContaining({
        category: 'error',
        code: expect.any(Number),
        message: expect.any(String),
      }),
    );
  });
});
