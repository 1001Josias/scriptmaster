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

interface InMemoryRangeApi {
  getValues(): InMemoryMatrix;
  setValues(values: InMemoryMatrix): InMemoryRangeApi;
}

interface InMemorySheetApi {
  getRange(a1Notation: string): GoogleAppsScript.Spreadsheet.Range;
  appendRow(rowContents: InMemoryCellValue[]): InMemorySheetApi;
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

  const [, startColumnName, startRowNumber, endColumnName, endRowNumber] = match;
  if (!startColumnName || !startRowNumber) {
    throw new Error(`Unsupported A1 notation: ${notation}`);
  }

  const startColumn = columnToIndex(startColumnName);
  const startRow = Number(startRowNumber) - 1;
  const endColumn = endColumnName ? columnToIndex(endColumnName) : startColumn;
  const endRow = endRowNumber ? Number(endRowNumber) - 1 : startRow;

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

function cellAt(state: SheetState, row: number, column: number): InMemoryCellValue {
  return state.values[row]?.[column] ?? '';
}

function setCell(
  state: SheetState,
  row: number,
  column: number,
  value: InMemoryCellValue,
): void {
  const targetRow = state.values[row];
  if (!targetRow) throw new Error(`In-memory row ${row + 1} was not allocated.`);
  targetRow[column] = value;
}

function createRange(state: SheetState, notation: string): GoogleAppsScript.Spreadsheet.Range {
  const parsed = parseA1Notation(notation);
  const range: InMemoryRangeApi = {
    getValues(): InMemoryMatrix {
      ensureSize(
        state,
        parsed.startRow + parsed.rowCount,
        parsed.startColumn + parsed.columnCount,
      );

      return Array.from({ length: parsed.rowCount }, (_, rowOffset) =>
        Array.from({ length: parsed.columnCount }, (_, columnOffset) =>
          cellAt(state, parsed.startRow + rowOffset, parsed.startColumn + columnOffset),
        ),
      );
    },

    setValues(values: InMemoryMatrix): InMemoryRangeApi {
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
          setCell(
            state,
            parsed.startRow + rowOffset,
            parsed.startColumn + columnOffset,
            value,
          );
        });
      });

      return range;
    },
  };

  return range as unknown as GoogleAppsScript.Spreadsheet.Range;
}

function createSheet(state: SheetState): GoogleAppsScript.Spreadsheet.Sheet {
  const sheet: InMemorySheetApi = {
    getRange(a1Notation: string): GoogleAppsScript.Spreadsheet.Range {
      return createRange(state, a1Notation);
    },

    appendRow(rowContents: InMemoryCellValue[]): InMemorySheetApi {
      state.values.push([...rowContents]);
      return sheet;
    },
  };

  return sheet as unknown as GoogleAppsScript.Spreadsheet.Sheet;
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
    this.selectedSheet().values = cloneMatrix(values);
    return this;
  }

  getValues(spreadsheetId: string, sheetName: string): InMemoryMatrix {
    const sheet = this.spreadsheets.get(spreadsheetId)?.sheets.get(sheetName);
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
    const spreadsheet = this.spreadsheets.get(this.selectedSpreadsheetId);
    if (!spreadsheet) throw new Error(`Unknown in-memory spreadsheet: ${this.selectedSpreadsheetId}`);
    return spreadsheet;
  }

  private selectedSheet(): SheetState {
    const spreadsheet = this.selectedSpreadsheet();
    if (!this.selectedSheetName) {
      throw new Error('Call sheet(name) before configuring values.');
    }
    const sheet = spreadsheet.sheets.get(this.selectedSheetName);
    if (!sheet) throw new Error(`Unknown in-memory sheet: ${this.selectedSheetName}`);
    return sheet;
  }
}

export function createInMemoryRuntimeBackend(): InMemoryRuntimeBackend {
  return new InMemoryRuntimeBackend();
}
