import { describe, expect, it } from 'vitest';

import { analyzeAppsScript, generateCompatibilityReport } from '@scriptmaster/analyzer';

import { CompilationBlockedError, generateNodeProject } from './index.js';

function reportFor(source: string) {
  return generateCompatibilityReport(analyzeAppsScript(source));
}

describe('generateNodeProject', () => {
  it('generates a deterministic Logger project while preserving user source', () => {
    const source = `function main() {
  Logger.log('hello');
}`;
    const input = { name: 'My GAS Script', source, report: reportFor(source) };

    const first = generateNodeProject(input);
    const second = generateNodeProject(input);

    expect(first).toEqual(second);
    expect(first.name).toBe('my-gas-script');
    expect(first.files.map((file) => file.path)).toEqual([
      'package.json',
      'scriptmaster-report.json',
      'src/index.ts',
      'tsconfig.json',
    ]);
    expect(first.files.find((file) => file.path === 'src/index.ts')?.content).toContain(source);
    expect(first.files.find((file) => file.path === 'src/index.ts')?.content).toContain(
      'const Logger = runtime.Logger;',
    );
  });

  it('injects SpreadsheetApp configuration without changing GAS calls', () => {
    const source = `async function main() {
  const spreadsheet = SpreadsheetApp.openById('sheet-id');
  const sheet = await spreadsheet.getSheetByName('Data');
  await sheet?.appendRow(['value']);
}`;
    const project = generateNodeProject({ source, report: reportFor(source) });
    const entry = project.files.find((file) => file.path === 'src/index.ts')?.content;

    expect(entry).toContain('sheetsClient: runtime.SheetsApiClient;');
    expect(entry).toContain(
      'const SpreadsheetApp = runtime.createSpreadsheetApp(configuration.sheetsClient);',
    );
    expect(entry).toContain("SpreadsheetApp.openById('sheet-id')");
  });

  it('blocks unsupported and unknown APIs with actionable locations', () => {
    const source = 'DriveApp.getFiles();\nUnknownApp.execute();';

    expect(() =>
      generateNodeProject({ source, report: reportFor(source) }),
    ).toThrowError(CompilationBlockedError);

    try {
      generateNodeProject({ source, report: reportFor(source) });
    } catch (error) {
      expect(error).toBeInstanceOf(CompilationBlockedError);
      expect((error as CompilationBlockedError).diagnostics).toEqual(
        expect.arrayContaining([
          expect.objectContaining({ symbol: 'DriveApp', status: 'unsupported', line: 1 }),
          expect.objectContaining({ symbol: 'UnknownApp', status: 'unknown', line: 2 }),
        ]),
      );
    }
  });
});
