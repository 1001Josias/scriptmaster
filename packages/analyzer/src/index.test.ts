import { describe, expect, it } from 'vitest';

import { analyzeAppsScript } from './index.js';

describe('analyzeAppsScript', () => {
  it('detects supported services, methods and trigger functions with locations', () => {
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
    ]);
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
