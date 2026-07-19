import type {
  AnalysisDiagnostic,
  AnalysisResult,
  DetectedSymbol,
  SourceLocation,
  SymbolKind,
} from './index.js';

export type CompatibilityStatus =
  | 'supported'
  | 'partially_supported'
  | 'unsupported'
  | 'unknown';

export interface CompatibilityItem {
  kind: SymbolKind;
  name: string;
  service?: string;
  receiverType?: string;
  status: CompatibilityStatus;
  location: SourceLocation;
  note: string;
}

export interface CompatibilitySummary {
  total: number;
  supported: number;
  partiallySupported: number;
  unsupported: number;
  unknown: number;
  score: number;
}

export interface CompatibilityReport {
  schemaVersion: 1;
  summary: CompatibilitySummary;
  items: CompatibilityItem[];
  diagnostics: AnalysisDiagnostic[];
}

interface CatalogEntry {
  status: Exclude<CompatibilityStatus, 'unknown'>;
  note: string;
}

const SERVICE_CATALOG: Readonly<Record<string, CatalogEntry>> = {
  Logger: {
    status: 'supported',
    note: 'Supported by the @scriptmaster/runtime Logger compatibility adapter.',
  },
  SpreadsheetApp: {
    status: 'supported',
    note: 'Supported by the configured SpreadsheetApp runtime backend for the MVP surface.',
  },
  DriveApp: {
    status: 'unsupported',
    note: 'Recognized, but no Drive compatibility adapter is available yet.',
  },
  GmailApp: {
    status: 'unsupported',
    note: 'Recognized, but no Gmail compatibility adapter is available yet.',
  },
  UrlFetchApp: {
    status: 'unsupported',
    note: 'Recognized, but automatic migration to a Node.js HTTP client is not available yet.',
  },
};

const METHOD_CATALOG: Readonly<Record<string, CatalogEntry>> = {
  'Logger.log': {
    status: 'supported',
    note: 'Supported with structured capture and Node.js-compatible value formatting.',
  },
  'SpreadsheetApp.openById': {
    status: 'supported',
    note: 'Supported by the configured runtime backend; authentication remains executor-owned.',
  },
  'SpreadsheetApp.Spreadsheet.getSheetByName': {
    status: 'supported',
    note: 'Supported by the SpreadsheetApp MVP runtime surface.',
  },
  'SpreadsheetApp.Sheet.getRange': {
    status: 'supported',
    note: 'Supported for A1 notation by the SpreadsheetApp MVP runtime surface.',
  },
  'SpreadsheetApp.Range.getValues': {
    status: 'supported',
    note: 'Supported by the SpreadsheetApp MVP runtime surface.',
  },
  'SpreadsheetApp.Range.setValues': {
    status: 'supported',
    note: 'Supported by the SpreadsheetApp MVP runtime surface.',
  },
  'SpreadsheetApp.Sheet.appendRow': {
    status: 'supported',
    note: 'Supported by the SpreadsheetApp MVP runtime surface.',
  },
};

const TRIGGER_ENTRY: CatalogEntry = {
  status: 'partially_supported',
  note: 'Trigger detected. Scheduling or HTTP exposure must be configured in the target runtime.',
};

function catalogEntryFor(symbol: DetectedSymbol): CatalogEntry | undefined {
  if (symbol.kind === 'trigger') return TRIGGER_ENTRY;
  if (symbol.kind === 'service') return SERVICE_CATALOG[symbol.name];

  if (symbol.service) {
    const typedKey = symbol.receiverType
      ? `${symbol.service}.${symbol.receiverType}.${symbol.name}`
      : undefined;
    const methodEntry =
      (typedKey ? METHOD_CATALOG[typedKey] : undefined) ??
      METHOD_CATALOG[`${symbol.service}.${symbol.name}`];

    if (methodEntry) return methodEntry;

    const serviceEntry = SERVICE_CATALOG[symbol.service];
    return serviceEntry?.status === 'unsupported' ? serviceEntry : undefined;
  }

  return undefined;
}

function toCompatibilityItem(symbol: DetectedSymbol): CompatibilityItem {
  const entry = catalogEntryFor(symbol);
  const context = {
    ...(symbol.service ? { service: symbol.service } : {}),
    ...(symbol.receiverType ? { receiverType: symbol.receiverType } : {}),
  };

  if (!entry) {
    return {
      kind: symbol.kind,
      name: symbol.name,
      ...context,
      status: 'unknown',
      location: symbol.location,
      note: 'No compatibility information is available for this symbol.',
    };
  }

  return {
    kind: symbol.kind,
    name: symbol.name,
    ...context,
    status: entry.status,
    location: symbol.location,
    note: entry.note,
  };
}

function normalizeSymbols(symbols: readonly DetectedSymbol[]): DetectedSymbol[] {
  const seenServices = new Set<string>();

  return symbols.filter((symbol) => {
    if (symbol.kind !== 'service') return true;
    if (seenServices.has(symbol.name)) return false;
    seenServices.add(symbol.name);
    return true;
  });
}

function summarize(items: CompatibilityItem[]): CompatibilitySummary {
  const summary: CompatibilitySummary = {
    total: items.length,
    supported: 0,
    partiallySupported: 0,
    unsupported: 0,
    unknown: 0,
    score: 100,
  };

  for (const item of items) {
    if (item.status === 'supported') summary.supported += 1;
    else if (item.status === 'partially_supported') summary.partiallySupported += 1;
    else if (item.status === 'unsupported') summary.unsupported += 1;
    else summary.unknown += 1;
  }

  if (summary.total > 0) {
    const compatibleUnits = summary.supported + summary.partiallySupported * 0.5;
    summary.score = Math.round((compatibleUnits / summary.total) * 100);
  }

  return summary;
}

export function generateCompatibilityReport(analysis: AnalysisResult): CompatibilityReport {
  const items = normalizeSymbols(analysis.symbols).map(toCompatibilityItem);

  return {
    schemaVersion: 1,
    summary: summarize(items),
    items,
    diagnostics: [...analysis.diagnostics],
  };
}
