import type { CompatibilityItem, CompatibilityReport } from '@scriptmaster/analyzer';

export type SuggestionConfidence = 'low' | 'medium' | 'high';

export interface MigrationSuggestion {
  symbol: string;
  location: CompatibilityItem['location'];
  originalStatus: CompatibilityItem['status'];
  summary: string;
  replacement?: string;
  behaviorChanges: string[];
  risks: string[];
  confidence: SuggestionConfidence;
  source: 'ai';
}

export interface MigrationSuggestionResult {
  status: 'generated' | 'not_configured' | 'no_findings' | 'failed';
  suggestions: MigrationSuggestion[];
  error?: string;
}

export interface MigrationSuggestionPrompt {
  system: string;
  input: {
    source: string;
    findings: Array<{
      symbol: string;
      status: CompatibilityItem['status'];
      line: number;
      column: number;
      analyzerNote: string;
    }>;
  };
}

export interface MigrationSuggestionProvider {
  generate(prompt: MigrationSuggestionPrompt): Promise<unknown>;
}

export interface GenerateMigrationSuggestionsInput {
  source: string;
  report: CompatibilityReport;
  provider?: MigrationSuggestionProvider;
}

function symbolName(item: CompatibilityItem): string {
  return item.service ? `${item.service}.${item.name}` : item.name;
}

function actionableFindings(report: CompatibilityReport): CompatibilityItem[] {
  return report.items.filter((item) => item.status !== 'supported');
}

export function buildMigrationSuggestionPrompt(
  source: string,
  report: CompatibilityReport,
): MigrationSuggestionPrompt {
  return {
    system: [
      'You are a Google Apps Script to Node.js migration assistant.',
      'Return only JSON matching the requested suggestion schema.',
      'Treat the analyzer findings as authoritative and never change their compatibility status.',
      'Provide practical Node.js or Google API replacements, behavior changes, risks, and confidence.',
      'Do not claim that suggested code has been executed or verified.',
    ].join(' '),
    input: {
      source,
      findings: actionableFindings(report).map((item) => ({
        symbol: symbolName(item),
        status: item.status,
        line: item.location.line,
        column: item.location.column,
        analyzerNote: item.note,
      })),
    },
  };
}

function isStringArray(value: unknown): value is string[] {
  return Array.isArray(value) && value.every((item) => typeof item === 'string');
}

function parseSuggestion(value: unknown, findings: CompatibilityItem[]): MigrationSuggestion | null {
  if (!value || typeof value !== 'object') return null;
  const candidate = value as Record<string, unknown>;
  const symbol = candidate.symbol;
  const finding = findings.find((item) => symbolName(item) === symbol);

  if (
    typeof symbol !== 'string' ||
    !finding ||
    typeof candidate.summary !== 'string' ||
    !isStringArray(candidate.behaviorChanges) ||
    !isStringArray(candidate.risks) ||
    !['low', 'medium', 'high'].includes(String(candidate.confidence))
  ) {
    return null;
  }

  return {
    symbol,
    location: finding.location,
    originalStatus: finding.status,
    summary: candidate.summary,
    ...(typeof candidate.replacement === 'string' ? { replacement: candidate.replacement } : {}),
    behaviorChanges: candidate.behaviorChanges,
    risks: candidate.risks,
    confidence: candidate.confidence as SuggestionConfidence,
    source: 'ai',
  };
}

function parseProviderResponse(value: unknown, findings: CompatibilityItem[]): MigrationSuggestion[] {
  const rawSuggestions =
    value && typeof value === 'object' && Array.isArray((value as { suggestions?: unknown }).suggestions)
      ? (value as { suggestions: unknown[] }).suggestions
      : [];

  return rawSuggestions
    .map((suggestion) => parseSuggestion(suggestion, findings))
    .filter((suggestion): suggestion is MigrationSuggestion => suggestion !== null);
}

export async function generateMigrationSuggestions(
  input: GenerateMigrationSuggestionsInput,
): Promise<MigrationSuggestionResult> {
  const findings = actionableFindings(input.report);

  if (findings.length === 0) {
    return { status: 'no_findings', suggestions: [] };
  }

  if (!input.provider) {
    return { status: 'not_configured', suggestions: [] };
  }

  try {
    const response = await input.provider.generate(buildMigrationSuggestionPrompt(input.source, input.report));
    return {
      status: 'generated',
      suggestions: parseProviderResponse(response, findings),
    };
  } catch (error) {
    return {
      status: 'failed',
      suggestions: [],
      error: error instanceof Error ? error.message : String(error),
    };
  }
}
