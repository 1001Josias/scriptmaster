import type { CompatibilityItem, CompatibilityReport } from '@scriptmaster/analyzer';

export interface GeneratedFile {
  path: string;
  content: string;
}

export interface GeneratedProject {
  name: string;
  files: GeneratedFile[];
  report: CompatibilityReport;
  entryFunctions: string[];
}

export interface GenerateNodeProjectInput {
  name?: string;
  source: string;
  report: CompatibilityReport;
}

export interface CompilationDiagnostic {
  kind: CompatibilityItem['kind'];
  symbol: string;
  status: CompatibilityItem['status'];
  line: number;
  column: number;
  message: string;
}

export class CompilationBlockedError extends Error {
  constructor(readonly diagnostics: CompilationDiagnostic[]) {
    super(
      `Node.js project generation blocked by ${diagnostics.length} incompatible symbol${diagnostics.length === 1 ? '' : 's'}.`,
    );
    this.name = 'CompilationBlockedError';
  }
}

interface RuntimeBinding {
  name: string;
  declaration: string;
}

function sanitizePackageName(value: string): string {
  const sanitized = value
    .trim()
    .toLowerCase()
    .replace(/[^a-z0-9._-]+/g, '-')
    .replace(/^-+|-+$/g, '');

  return sanitized || 'scriptmaster-project';
}

function blockingDiagnostics(report: CompatibilityReport): CompilationDiagnostic[] {
  return report.items
    .filter((item) => item.status === 'unsupported' || item.status === 'unknown')
    .map((item) => ({
      kind: item.kind,
      symbol: item.service ? `${item.service}.${item.name}` : item.name,
      status: item.status,
      line: item.location.line,
      column: item.location.column,
      message: item.note,
    }));
}

function runtimeBindings(report: CompatibilityReport): RuntimeBinding[] {
  const services = new Set(
    report.items.filter((item) => item.kind === 'service').map((item) => item.name),
  );
  const bindings: RuntimeBinding[] = [];

  if (services.has('Logger')) {
    bindings.push({ name: 'Logger', declaration: 'const Logger = runtime.Logger;' });
  }

  if (services.has('SpreadsheetApp')) {
    bindings.push({
      name: 'SpreadsheetApp',
      declaration:
        'const SpreadsheetApp = runtime.createSpreadsheetApp(configuration.sheetsClient);',
    });
  }

  return bindings;
}

function discoverEntryFunctions(source: string): string[] {
  const names = new Set<string>();
  const pattern = /^(?:export\s+)?(?:async\s+)?function\s+([A-Za-z_$][\w$]*)\s*\(/gm;
  for (const match of source.matchAll(pattern)) {
    if (match[1]) names.add(match[1]);
  }
  return [...names].sort();
}

function generateEntrySource(
  source: string,
  report: CompatibilityReport,
  entryFunctions: string[],
): string {
  const bindings = runtimeBindings(report);
  const usesSpreadsheet = bindings.some((binding) => binding.name === 'SpreadsheetApp');
  const imports = "import * as runtime from '@scriptmaster/runtime';";

  if (usesSpreadsheet) {
    const returnedNames = [...bindings.map((binding) => binding.name), ...entryFunctions];

    return [
      imports,
      '',
      'export interface ScriptMasterConfiguration {',
      '  sheetsClient: runtime.SheetsApiClient;',
      '}',
      '',
      'export function createScript(configuration: ScriptMasterConfiguration) {',
      ...bindings.map((binding) => `  ${binding.declaration}`),
      '',
      source
        .trim()
        .split('\n')
        .map((line) => `  ${line}`)
        .join('\n'),
      '',
      '  return {',
      ...returnedNames.map((name) => `    ${name},`),
      '  };',
      '}',
      '',
    ].join('\n');
  }

  const exportedNames = [...bindings.map((binding) => binding.name), ...entryFunctions];
  const exportLine = exportedNames.length > 0 ? `export { ${exportedNames.join(', ')} };` : 'export {};';

  return [
    imports,
    '',
    ...bindings.map((binding) => binding.declaration),
    '',
    source.trim(),
    '',
    exportLine,
    '',
  ].join('\n');
}

function packageJson(name: string): string {
  return `${JSON.stringify(
    {
      name,
      version: '0.0.0',
      private: true,
      type: 'module',
      scripts: {
        build: 'tsc',
        start: 'node dist/index.js',
      },
      dependencies: {
        '@scriptmaster/runtime': '^0.0.0',
      },
      devDependencies: {
        typescript: '^5.8.3',
      },
    },
    null,
    2,
  )}\n`;
}

const TSCONFIG = `${JSON.stringify(
  {
    compilerOptions: {
      target: 'ES2022',
      module: 'NodeNext',
      moduleResolution: 'NodeNext',
      strict: true,
      outDir: 'dist',
      rootDir: 'src',
      skipLibCheck: true,
    },
    include: ['src/**/*.ts'],
  },
  null,
  2,
)}\n`;

export function generateNodeProject(input: GenerateNodeProjectInput): GeneratedProject {
  const diagnostics = blockingDiagnostics(input.report);
  if (diagnostics.length > 0) {
    throw new CompilationBlockedError(diagnostics);
  }

  const name = sanitizePackageName(input.name ?? 'scriptmaster-project');
  const entryFunctions = discoverEntryFunctions(input.source);
  const files: GeneratedFile[] = [
    { path: 'package.json', content: packageJson(name) },
    { path: 'tsconfig.json', content: TSCONFIG },
    {
      path: 'src/index.ts',
      content: generateEntrySource(input.source, input.report, entryFunctions),
    },
    { path: 'scriptmaster-report.json', content: `${JSON.stringify(input.report, null, 2)}\n` },
  ];

  return {
    name,
    files: files.sort((left, right) => left.path.localeCompare(right.path)),
    report: input.report,
    entryFunctions,
  };
}
