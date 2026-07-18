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
    status: 'partially_supported',
    note: 'Recognized. The Node.js compatibility adapter is planned for the MVP.',
  },
  SpreadsheetApp: {
    status: 'partially_supported',
    note: 'Recognized. Initial migration support will use the Google Sheets API.',
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
    status: 'partially_supported',
    note: 'Will be migrated to the ScriptMaster logging compatibility layer.',
  },
  'SpreadsheetApp.openById': {
    status: 'partially_supported',
    note: 'Will be migrated to the Google Sheets API using configured credentials.',
  },
};

const TRIGGER_ENTRY: CatalogEntry = {
  status: 'partially_supported',
  note: 'Trigger detected. Scheduling or HTTP exposure must be configured in the target runtime.',
};

function catalogEntryFor(symbol: DetectedSymbol): CatalogEntry | undefined {
  if (symbol.kind === 'trigger') {
    return TRIGGER_ENTRY;
  }

  if (symbol.kind === 'service') {
    return SERVICE_CATALOG[symbol.name];
  }

  if (symbol.service) {
    return METHOD_CATALOG[`${symbol.service}.${symbol.name}`] ?? SERVICE_CATALOG[symbol.service];
  }

  return undefined;
}

function toCompatibilityItem(symbol: DetectedSymbol): CompatibilityItem {
  const entry = catalogEntryFor(symbol);

  if (!entry) {
    return {
      kind: symbol.kind,
      name: symbol.name,
      ...(symbol.service ? { service: symbol.service } : {}),
      status: 'unknown',
      location: symbol.location,
      note: 'No compatibility information is available for this symbol.',
    };
  }

  return {
    kind: symbol.kind,
    name: symbol.name,
    ...(symbol.service ? { service: symbol.service } : {}),
    status: entry.status,
    location: symbol.location,
    note: entry.note,
  };
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
    if (item.status === 'supported') {
      summary.supported += 1;
    } else if (item.status === 'partially_supported') {
      summary.partiallySupported += 1;
    } else if (item.status === 'unsupported') {
      summary.unsupported += 1;
    } else {
      summary.unknown += 1;
    }
  }

  if (summary.total > 0) {
    const compatibleUnits = summary.supported + summary.partiallySupported * 0.5;
    summary.score = Math.round((compatibleUnits / summary.total) * 100);
  }

  return summary;
}

export function generateCompatibilityReport(analysis: AnalysisResult): CompatibilityReport {
  const items = analysis.symbols.map(toCompatibilityItem);

  return {
    schemaVersion: 1,
    summary: summarize(items),
    items,
    diagnostics: [...analysis.diagnostics],
  };
}
