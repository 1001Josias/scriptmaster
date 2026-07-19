import { describe, expect, it } from 'vitest';

import { analyzeAppsScript, generateCompatibilityReport } from './index.js';

describe('generateCompatibilityReport', () => {
  it('classifies mixed compatibility scenarios deterministically', () => {
    const analysis = analyzeAppsScript(`function onOpen() {
  SpreadsheetApp.openById('sheet-id');
  DriveApp.getFiles();
  UnknownApp.execute();
}`);

    const report = generateCompatibilityReport(analysis);

    expect(report.schemaVersion).toBe(1);
    expect(
      report.items.map(({ kind, name, service, status }) => ({
        kind,
        name,
        ...(service ? { service } : {}),
        status,
      })),
    ).toEqual([
      { kind: 'trigger', name: 'onOpen', status: 'partially_supported' },
      { kind: 'service', name: 'SpreadsheetApp', status: 'supported' },
      {
        kind: 'method',
        name: 'openById',
        service: 'SpreadsheetApp',
        status: 'supported',
      },
      { kind: 'service', name: 'DriveApp', status: 'unsupported' },
      { kind: 'method', name: 'getFiles', service: 'DriveApp', status: 'unsupported' },
      { kind: 'service', name: 'UnknownApp', status: 'unknown' },
      { kind: 'method', name: 'execute', service: 'UnknownApp', status: 'unknown' },
    ]);
    expect(report.summary).toEqual({
      total: 7,
      supported: 2,
      partiallySupported: 1,
      unsupported: 2,
      unknown: 2,
      score: 36,
    });
  });

  it('emits one service summary while preserving every method occurrence', () => {
    const report = generateCompatibilityReport(
      analyzeAppsScript(`Logger.log('first');
Logger.log('second');`),
    );

    expect(report.items.filter((item) => item.kind === 'service')).toEqual([
      expect.objectContaining({
        name: 'Logger',
        location: { line: 1, column: 1 },
      }),
    ]);
    expect(report.items.filter((item) => item.kind === 'method')).toEqual([
      expect.objectContaining({ name: 'log', location: { line: 1, column: 8 } }),
      expect.objectContaining({ name: 'log', location: { line: 2, column: 8 } }),
    ]);
    expect(report.summary).toEqual({
      total: 3,
      supported: 3,
      partiallySupported: 0,
      unsupported: 0,
      unknown: 0,
      score: 100,
    });
  });

  it('preserves source locations and migration notes', () => {
    const report = generateCompatibilityReport(analyzeAppsScript('Logger.log("hello");'));

    expect(report.items).toEqual([
      expect.objectContaining({
        kind: 'service',
        name: 'Logger',
        status: 'supported',
        location: { line: 1, column: 1 },
        note: expect.any(String),
      }),
      expect.objectContaining({
        kind: 'method',
        name: 'log',
        service: 'Logger',
        status: 'supported',
        location: { line: 1, column: 8 },
        note: expect.any(String),
      }),
    ]);
    expect(report.summary).toEqual({
      total: 2,
      supported: 2,
      partiallySupported: 0,
      unsupported: 0,
      unknown: 0,
      score: 100,
    });
  });

  it('returns a stable empty summary and carries parser diagnostics', () => {
    const report = generateCompatibilityReport(analyzeAppsScript('function broken( {'));

    expect(report.summary).toEqual({
      total: 0,
      supported: 0,
      partiallySupported: 0,
      unsupported: 0,
      unknown: 0,
      score: 100,
    });
    expect(report.diagnostics.length).toBeGreaterThan(0);
  });
});
