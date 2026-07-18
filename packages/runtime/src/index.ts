import { format } from 'node:util';

export interface LogEntry {
  level: 'info';
  message: string;
  values: readonly unknown[];
}

export interface LoggerLike {
  log(...values: unknown[]): void;
  getEntries(): readonly LogEntry[];
  clear(): void;
}

export function createLogger(onEntry?: (entry: LogEntry) => void): LoggerLike {
  const entries: LogEntry[] = [];

  return {
    log(...values: unknown[]): void {
      const entry: LogEntry = {
        level: 'info',
        message: format(...values),
        values: [...values],
      };

      entries.push(entry);
      onEntry?.(entry);
    },

    getEntries(): readonly LogEntry[] {
      return entries.map((entry) => ({ ...entry, values: [...entry.values] }));
    },

    clear(): void {
      entries.length = 0;
    },
  };
}

export const Logger = createLogger();
