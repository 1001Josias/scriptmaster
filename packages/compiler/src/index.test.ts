import { describe, expect, it } from 'vitest';

import { analyzeAppsScript, generateCompatibilityReport } from '@scriptmaster/analyzer';

import { CompilationBlockedError, generateNodeProject } from './index.js';

function reportFor(source: string) {
  return generateCompatibilityReport(analyzeAppsScript(source));
}

describe('generateNodeProject', () => {
  it('generates a deterministic Logger project while preserving and exporting user source', () => {
    const source = `function main() {
  Logger.log('hello');
}`;
    const input = { name: 'My GAS Script', source, report: reportFor(source) };

    const first = generateNodeProject(input);
    const second = generateNodeProject(input);

    expect(first).toEqual(second);
    expect(first.name).toBe('my-gas-script');
    expect(first.entryFunctions).toEqual(['main']);
    expect(first.files.map((file) => file.path)).toEqual([
      'package.json',
      'scriptmaster-report.json',
      'src/index.ts',
      'tsconfig.json',
    ]);
    const entry = first.files.find((file) => file.path === 'src/index.ts')?.content;
    expect(entry).toContain(source);
    expect(entry).toContain('const Logger = runtime.Logger;');
    expect(entry).toContain('export { Logger, main };');
  });

  it('binds SpreadsheetApp through gas-fakes without rewriting synchronous GAS calls', () => {
    const source = `function main() {
  const spreadsheet = SpreadsheetApp.openById('sheet-id');
  const sheet = spreadsheet.getSheetByName('Data');
  sheet?.appendRow(['value']);
}`;
    const project = generateNodeProject({ source, report: reportFor(source) });
    const entry = project.files.find((file) => file.path === 'src/index.ts')?.content;
    const tsconfig = project.files.find((file) => file.path === 'tsconfig.json')?.content;

    expect(project.entryFunctions).toEqual(['main']);
    expect(entry).toContain(
      'await runtime.createRuntimeBindings({ backend: runtime.gasFakesBackend, services: ["SpreadsheetApp"] })',
    );
    expect(entry).toContain('const SpreadsheetApp = gas.SpreadsheetApp!;');
    expect(entry).toContain("spreadsheet.getSheetByName('Data')");
    expect(entry).not.toContain('sheetsClient');
    expect(entry).toContain('export { SpreadsheetApp, main };');
    expect(tsconfig).toContain('google-apps-script');
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
