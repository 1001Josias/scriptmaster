export type CellValue = string | number | boolean | null;
export type RowValues = CellValue[];
export type MatrixValues = RowValues[];

export interface SheetsApiClient {
  spreadsheets: {
    get(params: {
      spreadsheetId: string;
      fields?: string;
    }): Promise<{
      data?: {
        sheets?: Array<{ properties?: { title?: string } }>;
      };
    }>;
    values: {
      get(params: {
        spreadsheetId: string;
        range: string;
      }): Promise<{ data?: { values?: unknown[][] } }>;
      update(params: {
        spreadsheetId: string;
        range: string;
        valueInputOption: 'USER_ENTERED';
        requestBody: { values: MatrixValues };
      }): Promise<unknown>;
      append(params: {
        spreadsheetId: string;
        range: string;
        valueInputOption: 'USER_ENTERED';
        insertDataOption: 'INSERT_ROWS';
        requestBody: { values: MatrixValues };
      }): Promise<unknown>;
    };
  };
}

export class SpreadsheetCompatibilityError extends Error {
  constructor(
    message: string,
    readonly cause?: unknown,
  ) {
    super(message);
    this.name = 'SpreadsheetCompatibilityError';
  }
}

export class UnsupportedSpreadsheetMethodError extends Error {
  constructor(method: string) {
    super(`Spreadsheet compatibility method is not supported: ${method}`);
    this.name = 'UnsupportedSpreadsheetMethodError';
  }
}

export interface RangeLike {
  getValues(): Promise<MatrixValues>;
  setValues(values: MatrixValues): Promise<void>;
}

export interface SheetLike {
  getRange(a1Notation: string): RangeLike;
  appendRow(values: RowValues): Promise<void>;
}

export interface SpreadsheetLike {
  getSheetByName(name: string): Promise<SheetLike | null>;
}

export interface SpreadsheetAppLike {
  openById(spreadsheetId: string): SpreadsheetLike;
}

function normalizeValues(values: unknown[][] | undefined): MatrixValues {
  return (values ?? []).map((row) =>
    row.map((value) => {
      if (
        value === null ||
        typeof value === 'string' ||
        typeof value === 'number' ||
        typeof value === 'boolean'
      ) {
        return value;
      }

      return String(value);
    }),
  );
}

function normalizeError(operation: string, error: unknown): SpreadsheetCompatibilityError {
  const detail = error instanceof Error ? error.message : String(error);
  return new SpreadsheetCompatibilityError(`Google Sheets ${operation} failed: ${detail}`, error);
}

function withUnsupportedMethods<T extends object>(target: T, typeName: string): T {
  return new Proxy(target, {
    get(current, property, receiver) {
      if (typeof property === 'string' && !(property in current)) {
        return () => {
          throw new UnsupportedSpreadsheetMethodError(`${typeName}.${property}`);
        };
      }

      return Reflect.get(current, property, receiver);
    },
  });
}

class RangeAdapter implements RangeLike {
  constructor(
    private readonly client: SheetsApiClient,
    private readonly spreadsheetId: string,
    private readonly range: string,
  ) {}

  async getValues(): Promise<MatrixValues> {
    try {
      const response = await this.client.spreadsheets.values.get({
        spreadsheetId: this.spreadsheetId,
        range: this.range,
      });
      return normalizeValues(response.data?.values);
    } catch (error) {
      throw normalizeError(`read for range ${this.range}`, error);
    }
  }

  async setValues(values: MatrixValues): Promise<void> {
    try {
      await this.client.spreadsheets.values.update({
        spreadsheetId: this.spreadsheetId,
        range: this.range,
        valueInputOption: 'USER_ENTERED',
        requestBody: { values },
      });
    } catch (error) {
      throw normalizeError(`write for range ${this.range}`, error);
    }
  }
}

class SheetAdapter implements SheetLike {
  constructor(
    private readonly client: SheetsApiClient,
    private readonly spreadsheetId: string,
    private readonly title: string,
  ) {}

  getRange(a1Notation: string): RangeLike {
    if (!a1Notation.trim()) {
      throw new SpreadsheetCompatibilityError('Range notation must not be empty.');
    }

    return withUnsupportedMethods(
      new RangeAdapter(this.client, this.spreadsheetId, `${this.title}!${a1Notation}`),
      'Range',
    );
  }

  async appendRow(values: RowValues): Promise<void> {
    try {
      await this.client.spreadsheets.values.append({
        spreadsheetId: this.spreadsheetId,
        range: this.title,
        valueInputOption: 'USER_ENTERED',
        insertDataOption: 'INSERT_ROWS',
        requestBody: { values: [values] },
      });
    } catch (error) {
      throw normalizeError(`append for sheet ${this.title}`, error);
    }
  }
}

class SpreadsheetAdapter implements SpreadsheetLike {
  constructor(
    private readonly client: SheetsApiClient,
    private readonly spreadsheetId: string,
  ) {}

  async getSheetByName(name: string): Promise<SheetLike | null> {
    try {
      const response = await this.client.spreadsheets.get({
        spreadsheetId: this.spreadsheetId,
        fields: 'sheets.properties.title',
      });
      const exists = response.data?.sheets?.some((sheet) => sheet.properties?.title === name) ?? false;

      return exists
        ? withUnsupportedMethods(new SheetAdapter(this.client, this.spreadsheetId, name), 'Sheet')
        : null;
    } catch (error) {
      throw normalizeError(`metadata lookup for spreadsheet ${this.spreadsheetId}`, error);
    }
  }
}

export function createSpreadsheetApp(client: SheetsApiClient): SpreadsheetAppLike {
  return withUnsupportedMethods(
    {
      openById(spreadsheetId: string): SpreadsheetLike {
        if (!spreadsheetId.trim()) {
          throw new SpreadsheetCompatibilityError('Spreadsheet id must not be empty.');
        }

        return withUnsupportedMethods(
          new SpreadsheetAdapter(client, spreadsheetId),
          'Spreadsheet',
        );
      },
    },
    'SpreadsheetApp',
  );
}
