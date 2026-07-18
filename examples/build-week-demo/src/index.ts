import { generateMigrationSuggestions, type MigrationSuggestionProvider } from '@scriptmaster/ai';
import { analyzeAppsScript, generateCompatibilityReport } from '@scriptmaster/analyzer';
import { generateNodeProject } from '@scriptmaster/compiler';
import { executeEntry, type ExecutableScriptModule } from '@scriptmaster/executor';
import {
  Logger,
  createRuntimeBindings,
  type AppsScriptRuntimeBindings,
  type RuntimeBackend,
} from '@scriptmaster/runtime';

const SOURCE = `function onOpen() {
  Logger.log('Configure this trigger in the target scheduler');
}

function main() {
  const spreadsheet = SpreadsheetApp.openById('demo-spreadsheet');
  const sheet = spreadsheet.getSheetByName('Orders');
  if (!sheet) throw new Error('Orders sheet was not found');
  const values = sheet.getRange('A1:B3').getValues();
  const total = values.slice(1).reduce((sum, row) => sum + Number(row[1]), 0);
  Logger.log('Processed %s orders with total %s', values.length - 1, total);
  return { orderCount: values.length - 1, total };
}`;

const suggestionProvider: MigrationSuggestionProvider = {
  async generate(prompt) {
    return {
      suggestions: prompt.input.findings.map((finding) => ({
        symbol: finding.symbol,
        summary: 'Configure the Apps Script trigger as an explicit schedule or event in the target platform.',
        replacement: 'Create a BotMaster schedule that invokes the generated onOpen entry function.',
        behaviorChanges: ['The trigger lifecycle becomes explicit infrastructure configuration.'],
        risks: ['The schedule timezone and execution identity must match the original Apps Script project.'],
        confidence: 'high',
      })),
    };
  },
};

function createDemoBackend(): RuntimeBackend {
  const rows = [
    ['order_id', 'amount'],
    ['A-100', 120],
    ['A-101', 80],
  ];

  const range = {
    getValues: () => rows.map((row) => [...row]),
  } as unknown as GoogleAppsScript.Spreadsheet.Range;

  const sheet = {
    getRange: (a1Notation: string) => {
      if (a1Notation !== 'A1:B3') throw new Error(`Unexpected demo range: ${a1Notation}`);
      return range;
    },
  } as unknown as GoogleAppsScript.Spreadsheet.Sheet;

  const spreadsheet = {
    getSheetByName: (name: string) => (name === 'Orders' ? sheet : null),
  } as unknown as GoogleAppsScript.Spreadsheet.Spreadsheet;

  const SpreadsheetApp = {
    openById: (id: string) => {
      if (id !== 'demo-spreadsheet') throw new Error(`Unexpected demo spreadsheet: ${id}`);
      return spreadsheet;
    },
  } as unknown as GoogleAppsScript.Spreadsheet.SpreadsheetApp;

  return {
    name: 'demo-memory',
    async initialize(): Promise<AppsScriptRuntimeBindings> {
      return { SpreadsheetApp };
    },
  };
}

function instantiateSource(bindings: AppsScriptRuntimeBindings): ExecutableScriptModule {
  const factory = new Function(
    'Logger',
    'SpreadsheetApp',
    `'use strict';\n${SOURCE}\nreturn { main, onOpen };`,
  ) as (
    logger: typeof Logger,
    spreadsheetApp: GoogleAppsScript.Spreadsheet.SpreadsheetApp,
  ) => ExecutableScriptModule;

  return {
    ...factory(Logger, bindings.SpreadsheetApp!),
    Logger,
  };
}

async function run(): Promise<void> {
  console.log('\n1. SOURCE\n', SOURCE);

  const analysis = analyzeAppsScript(SOURCE, 'demo.js');
  const report = generateCompatibilityReport(analysis);
  console.log('\n2. COMPATIBILITY REPORT\n', JSON.stringify(report, null, 2));

  const suggestions = await generateMigrationSuggestions({
    source: SOURCE,
    report,
    provider: suggestionProvider,
  });
  console.log('\n3. AI GUIDANCE (NON-DETERMINISTIC)\n', JSON.stringify(suggestions, null, 2));

  const project = generateNodeProject({
    name: 'scriptmaster-build-week-demo',
    source: SOURCE,
    report,
  });
  console.log('\n4. GENERATED NODE PROJECT\n', project.files.map((file) => file.path).join('\n'));

  const bindings = await createRuntimeBindings({
    backend: createDemoBackend(),
    services: ['SpreadsheetApp'],
  });
  const module = instantiateSource(bindings);
  const result = await executeEntry({ module, entryFunction: 'main' });
  console.log('\n5. EXECUTION RESULT\n', JSON.stringify(result, null, 2));

  console.log('\n6. NEXT STEP\nThe generated entry function can be packaged as a BotMaster Worker.');

  if (result.status !== 'succeeded') process.exitCode = 1;
}

run().catch((error: unknown) => {
  console.error(error);
  process.exitCode = 1;
});
