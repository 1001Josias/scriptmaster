import ts from 'typescript';

const KNOWN_SERVICES = new Set([
  'CalendarApp',
  'ContactsApp',
  'DocumentApp',
  'DriveApp',
  'FormApp',
  'GmailApp',
  'HtmlService',
  'LanguageApp',
  'Logger',
  'MailApp',
  'Maps',
  'PropertiesService',
  'ScriptApp',
  'Session',
  'SlidesApp',
  'SpreadsheetApp',
  'UrlFetchApp',
  'Utilities',
]);

const TRIGGER_FUNCTIONS = new Set([
  'doGet',
  'doPost',
  'onEdit',
  'onFormSubmit',
  'onInstall',
  'onOpen',
]);

interface ReceiverContext {
  service: string;
  type: string;
}

const METHOD_RETURN_TYPES: Readonly<Record<string, string>> = {
  'SpreadsheetApp.service.openById': 'Spreadsheet',
  'SpreadsheetApp.Spreadsheet.getSheetByName': 'Sheet',
  'SpreadsheetApp.Sheet.getRange': 'Range',
  'SpreadsheetApp.Sheet.appendRow': 'Sheet',
};

export type SymbolKind = 'service' | 'method' | 'trigger';

export interface SourceLocation {
  line: number;
  column: number;
}

export interface DetectedSymbol {
  kind: SymbolKind;
  name: string;
  service?: string;
  receiverType?: string;
  supported: boolean;
  location: SourceLocation;
}

export interface AnalysisDiagnostic {
  category: 'error' | 'warning';
  code: number;
  message: string;
  location?: SourceLocation;
}

export interface AnalysisResult {
  symbols: DetectedSymbol[];
  diagnostics: AnalysisDiagnostic[];
}

function toLocation(sourceFile: ts.SourceFile, position: number): SourceLocation {
  const location = sourceFile.getLineAndCharacterOfPosition(position);

  return {
    line: location.line + 1,
    column: location.character + 1,
  };
}

function flattenDiagnosticMessage(message: string | ts.DiagnosticMessageChain): string {
  return ts.flattenDiagnosticMessageText(message, '\n');
}

function collectDiagnostics(source: string, fileName: string): AnalysisDiagnostic[] {
  const output = ts.transpileModule(source, {
    fileName,
    reportDiagnostics: true,
    compilerOptions: {
      allowJs: true,
      target: ts.ScriptTarget.ES2022,
      module: ts.ModuleKind.ESNext,
    },
  });

  return (output.diagnostics ?? []).map((diagnostic) => {
    const location =
      diagnostic.start === undefined
        ? undefined
        : toLocation(
            ts.createSourceFile(fileName, source, ts.ScriptTarget.ES2022, true),
            diagnostic.start,
          );

    return {
      category: diagnostic.category === ts.DiagnosticCategory.Error ? 'error' : 'warning',
      code: diagnostic.code,
      message: flattenDiagnosticMessage(diagnostic.messageText),
      ...(location ? { location } : {}),
    };
  });
}

function isAppsScriptServiceName(name: string): boolean {
  return KNOWN_SERVICES.has(name) || name.endsWith('App');
}

function returnTypeFor(context: ReceiverContext, method: string): string | undefined {
  return METHOD_RETURN_TYPES[`${context.service}.${context.type}.${method}`];
}

export function analyzeAppsScript(source: string, fileName = 'script.ts'): AnalysisResult {
  const sourceFile = ts.createSourceFile(
    fileName,
    source,
    ts.ScriptTarget.ES2022,
    true,
    fileName.endsWith('.ts') ? ts.ScriptKind.TS : ts.ScriptKind.JS,
  );
  const diagnostics = collectDiagnostics(source, fileName);
  const symbols: DetectedSymbol[] = [];
  const variables = new Map<string, ReceiverContext>();

  function addSymbol(symbol: DetectedSymbol): void {
    symbols.push(symbol);
  }

  function resolveExpression(expression: ts.Expression): ReceiverContext | undefined {
    if (ts.isIdentifier(expression)) {
      if (isAppsScriptServiceName(expression.text)) {
        return { service: expression.text, type: 'service' };
      }
      return variables.get(expression.text);
    }

    if (ts.isParenthesizedExpression(expression)) {
      return resolveExpression(expression.expression);
    }

    if (ts.isCallExpression(expression) && ts.isPropertyAccessExpression(expression.expression)) {
      const receiver = resolveExpression(expression.expression.expression);
      if (!receiver) return undefined;
      const returnType = returnTypeFor(receiver, expression.expression.name.text);
      return returnType ? { service: receiver.service, type: returnType } : undefined;
    }

    return undefined;
  }

  function visit(node: ts.Node): void {
    if (ts.isFunctionDeclaration(node) && node.name && TRIGGER_FUNCTIONS.has(node.name.text)) {
      addSymbol({
        kind: 'trigger',
        name: node.name.text,
        supported: true,
        location: toLocation(sourceFile, node.name.getStart(sourceFile)),
      });
    }

    if (
      ts.isVariableDeclaration(node) &&
      ts.isIdentifier(node.name) &&
      node.initializer
    ) {
      const context = resolveExpression(node.initializer);
      if (context) variables.set(node.name.text, context);
    }

    if (ts.isCallExpression(node) && ts.isPropertyAccessExpression(node.expression)) {
      const receiver = resolveExpression(node.expression.expression);
      if (receiver) {
        if (receiver.type === 'service') {
          addSymbol({
            kind: 'service',
            name: receiver.service,
            supported: KNOWN_SERVICES.has(receiver.service),
            location: toLocation(sourceFile, node.expression.expression.getStart(sourceFile)),
          });
        }

        addSymbol({
          kind: 'method',
          name: node.expression.name.text,
          service: receiver.service,
          ...(receiver.type === 'service' ? {} : { receiverType: receiver.type }),
          supported: KNOWN_SERVICES.has(receiver.service),
          location: toLocation(sourceFile, node.expression.name.getStart(sourceFile)),
        });
      }
    }

    ts.forEachChild(node, visit);
  }

  visit(sourceFile);

  symbols.sort((left, right) => {
    return (
      left.location.line - right.location.line ||
      left.location.column - right.location.column ||
      left.kind.localeCompare(right.kind) ||
      left.name.localeCompare(right.name)
    );
  });

  return { symbols, diagnostics };
}

export {
  generateCompatibilityReport,
  type CompatibilityItem,
  type CompatibilityReport,
  type CompatibilityStatus,
  type CompatibilitySummary,
} from './compatibility.js';
