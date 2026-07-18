import { describe, expect, it, vi } from 'vitest';

import { analyzeAppsScript, generateCompatibilityReport } from '@scriptmaster/analyzer';

import {
  buildMigrationSuggestionPrompt,
  generateMigrationSuggestions,
  type MigrationSuggestionProvider,
} from './index.js';

function reportFor(source: string) {
  return generateCompatibilityReport(analyzeAppsScript(source, 'script.js'));
}

describe('migration suggestions', () => {
  it('does not require AI configuration', async () => {
    const source = 'DriveApp.getFiles();';

    await expect(
      generateMigrationSuggestions({ source, report: reportFor(source) }),
    ).resolves.toEqual({ status: 'not_configured', suggestions: [] });
  });

  it('builds a prompt only from non-supported findings', () => {
    const source = "Logger.log('ok');\nDriveApp.getFiles();";
    const prompt = buildMigrationSuggestionPrompt(source, reportFor(source));

    expect(prompt.input.source).toBe(source);
    expect(prompt.input.findings.some((finding) => finding.symbol === 'Logger')).toBe(false);
    expect(prompt.input.findings).toEqual(
      expect.arrayContaining([
        expect.objectContaining({ symbol: 'DriveApp', status: 'unsupported', line: 2 }),
        expect.objectContaining({ symbol: 'DriveApp.getFiles', status: 'unsupported', line: 2 }),
      ]),
    );
  });

  it('returns a source-linked AI suggestion without changing deterministic status', async () => {
    const source = 'DriveApp.getFiles();';
    const provider: MigrationSuggestionProvider = {
      generate: vi.fn().mockResolvedValue({
        suggestions: [
          {
            symbol: 'DriveApp.getFiles',
            summary: 'List files through the Google Drive API v3.',
            replacement: 'Use google.drive({ version: "v3", auth }).files.list(...).',
            behaviorChanges: ['Results are paginated and require explicit field selection.'],
            risks: ['OAuth scopes and shared-drive flags must be configured.'],
            confidence: 'high',
          },
        ],
      }),
    };

    const result = await generateMigrationSuggestions({
      source,
      report: reportFor(source),
      provider,
    });

    expect(result.status).toBe('generated');
    expect(result.suggestions).toEqual([
      expect.objectContaining({
        symbol: 'DriveApp.getFiles',
        location: { line: 1, column: 10 },
        originalStatus: 'unsupported',
        source: 'ai',
        confidence: 'high',
      }),
    ]);
  });

  it('ignores hallucinated symbols and normalizes provider failures', async () => {
    const source = 'DriveApp.getFiles();';
    const hallucinatingProvider: MigrationSuggestionProvider = {
      generate: async () => ({
        suggestions: [
          {
            symbol: 'ImaginaryApp.fly',
            summary: 'Invalid',
            behaviorChanges: [],
            risks: [],
            confidence: 'high',
          },
        ],
      }),
    };

    await expect(
      generateMigrationSuggestions({ source, report: reportFor(source), provider: hallucinatingProvider }),
    ).resolves.toEqual({ status: 'generated', suggestions: [] });

    const failingProvider: MigrationSuggestionProvider = {
      generate: async () => {
        throw new Error('provider unavailable');
      },
    };

    await expect(
      generateMigrationSuggestions({ source, report: reportFor(source), provider: failingProvider }),
    ).resolves.toEqual({
      status: 'failed',
      suggestions: [],
      error: 'provider unavailable',
    });
  });
});
