import type {
  AppsScriptRuntimeBindings,
  RuntimeBackend,
  RuntimeBackendContext,
} from './backend.js';

export type InMemoryCellValue = unknown;
export type InMemoryMatrix = InMemoryCellValue[][];

interface ParsedRange {
  startRow: number;
  startColumn: number;
  rowCount: number;
  columnCount: number;
}

interface SheetState {
  values: InMemoryMatrix;
}

interface SpreadsheetState {
  sheets: Map<string, SheetState>;
}

function cloneMatrix(values: readonly (readonly InMemoryCellValue[])[]): InMemoryMatrix {
  return values.map((row) => [...row]);
}

function columnToIndex(column: string): number {
  let value = 0;
  for (const character of column.toUpperCase()) {
    value = value * 26 + character.charCodeAt(0) - 64;
  }
  return value - 1;
}

function parseA1Notation(notation: string): ParsedRange {
  const match = /^([A-Za-z]+)(\d+)(?::([A-Za-z]+)(\d+))?$/.exec(notation.trim());
  if (!match) throw new Error(`Unsupported A1 notation: ${notation}`);

  const startColumn = columnToIndex(match[1]);
  const startRow = Number(match[2]) - 1;
  const endColumn = match[3] ? columnToIndex(match[3]) : startColumn;
  const endRow = match[4] ? Number(match[4]) - 1 : startRow;

  if (startRow < 0 || endRow < startRow || endColumn < startColumn) {
    throw new Error(`Invalid A1 notation: ${notation}`);
  }

  return {
    startRow,
    startColumn,
    rowCount: endRow - startRow + 1,
    columnCount: endColumn - startColumn + 1,
  };
}

function ensureSize(state: SheetState, rowCount: number, columnCount: number): void {
  while (state.values.length < rowCount) state.values.push([]);
  for (const row of state.values) {
    while (row.length < columnCount) row.push('');
  }
}

function createRange(state: SheetState, notation: string): GoogleAppsScript.Spreadsheet.Range {
  const parsed = parseA1Notation(notation);

  return {
    getValues(): Object[][] {
      ensureSize(
        state,
        parsed.startRow + parsed.rowCount,
        parsed.startColumn + parsed.columnCount,
      );

      return Array.from({ length: parsed.rowCount }, (_, rowOffset) =>
        Array.from(
          { length: parsed.columnCount },
          (_, columnOffset) =>
            state.values[parsed.startRow + rowOffset][parsed.startColumn + columnOffset],
        ),
      );
    },

    setValues(values: Object[][]): GoogleAppsScript.Spreadsheet.Range {
      if (
        values.length !== parsed.rowCount ||
        values.some((row) => row.length !== parsed.columnCount)
      ) {
        throw new Error(
          `setValues expected ${parsed.rowCount}x${parsed.columnCount} values for ${notation}.`,
        );
      }

      ensureSize(
        state,
        parsed.startRow + parsed.rowCount,
        parsed.startColumn + parsed.columnCount,
      );

      values.forEach((row, rowOffset) => {
        row.forEach((value, columnOffset) => {
          state.values[parsed.startRow + rowOffset][parsed.startColumn + columnOffset] = value;
        });
      });

      return this;
    },
  } as unknown as GoogleAppsScript.Spreadsheet.Range;
}

function createSheet(state: SheetState): GoogleAppsScript.Spreadsheet.Sheet {
  return {
    getRange(a1Notation: string): GoogleAppsScript.Spreadsheet.Range {
      return createRange(state, a1Notation);
    },

    appendRow(rowContents: Object[]): GoogleAppsScript.Spreadsheet.Sheet {
      state.values.push([...rowContents]);
      return this;
    },
  } as unknown as GoogleAppsScript.Spreadsheet.Sheet;
}

export class InMemoryRuntimeBackend implements RuntimeBackend {
  readonly name = 'in-memory';

  private readonly spreadsheets = new Map<string, SpreadsheetState>();
  private selectedSpreadsheetId?: string;
  private selectedSheetName?: string;

  spreadsheet(id: string): this {
    if (!id.trim()) throw new Error('Spreadsheet id must not be empty.');
    if (!this.spreadsheets.has(id)) this.spreadsheets.set(id, { sheets: new Map() });
    this.selectedSpreadsheetId = id;
    this.selectedSheetName = undefined;
    return this;
  }

  sheet(name: string): this {
    const spreadsheet = this.selectedSpreadsheet();
    if (!name.trim()) throw new Error('Sheet name must not be empty.');
    if (!spreadsheet.sheets.has(name)) spreadsheet.sheets.set(name, { values: [] });
    this.selectedSheetName = name;
    return this;
  }

  values(values: readonly (readonly InMemoryCellValue[])[]): this {
    const sheet = this.selectedSheet();
    sheet.values = cloneMatrix(values);
    return this;
  }

  getValues(spreadsheetId: string, sheetName: string): InMemoryMatrix {
    const spreadsheet = this.spreadsheets.get(spreadsheetId);
    const sheet = spreadsheet?.sheets.get(sheetName);
    if (!sheet) throw new Error(`Unknown in-memory sheet: ${spreadsheetId}/${sheetName}`);
    return cloneMatrix(sheet.values);
  }

  async initialize(context: RuntimeBackendContext): Promise<AppsScriptRuntimeBindings> {
    const requested = new Set(context.services);
    const bindings: AppsScriptRuntimeBindings = {};

    if (requested.has('SpreadsheetApp')) {
      bindings.SpreadsheetApp = {
        openById: (id: string): GoogleAppsScript.Spreadsheet.Spreadsheet => {
          const spreadsheet = this.spreadsheets.get(id);
          if (!spreadsheet) throw new Error(`Unknown in-memory spreadsheet: ${id}`);

          return {
            getSheetByName: (name: string): GoogleAppsScript.Spreadsheet.Sheet | null => {
              const sheet = spreadsheet.sheets.get(name);
              return sheet ? createSheet(sheet) : null;
            },
          } as unknown as GoogleAppsScript.Spreadsheet.Spreadsheet;
        },
      } as unknown as GoogleAppsScript.Spreadsheet.SpreadsheetApp;
    }

    return bindings;
  }

  private selectedSpreadsheet(): SpreadsheetState {
    if (!this.selectedSpreadsheetId) {
      throw new Error('Call spreadsheet(id) before configuring a sheet.');
    }
    return this.spreadsheets.get(this.selectedSpreadsheetId)!;
  }

  private selectedSheet(): SheetState {
    const spreadsheet = this.selectedSpreadsheet();
    if (!this.selectedSheetName) {
      throw new Error('Call sheet(name) before configuring values.');
    }
    return spreadsheet.sheets.get(this.selectedSheetName)!;
  }
}

export function createInMemoryRuntimeBackend(): InMemoryRuntimeBackend {
  return new InMemoryRuntimeBackend();
}
