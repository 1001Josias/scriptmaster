import type { LogEntry, LoggerLike } from '@scriptmaster/runtime';

export type ScriptFunction = (...args: unknown[]) => unknown | Promise<unknown>;

export interface ExecutableScriptModule {
  [name: string]: unknown;
  Logger?: LoggerLike;
}

export interface ExecuteEntryInput {
  module: ExecutableScriptModule;
  entryFunction: string;
  args?: unknown[];
  timeoutMs?: number;
  now?: () => number;
}

export interface ExecutionFailure {
  type: 'entry_not_found' | 'timeout' | 'runtime_error';
  message: string;
  name?: string;
  stack?: string;
}

export interface ExecutionResult {
  status: 'succeeded' | 'failed' | 'timed_out';
  entryFunction: string;
  startedAt: string;
  finishedAt: string;
  durationMs: number;
  returnValue?: unknown;
  logs: readonly LogEntry[];
  failure?: ExecutionFailure;
}

const DEFAULT_TIMEOUT_MS = 30_000;

function getLogs(module: ExecutableScriptModule): readonly LogEntry[] {
  return module.Logger?.getEntries() ?? [];
}

function serializeFailure(error: unknown): ExecutionFailure {
  if (error instanceof Error) {
    return {
      type: 'runtime_error',
      message: error.message,
      name: error.name,
      stack: error.stack,
    };
  }

  return {
    type: 'runtime_error',
    message: String(error),
  };
}

export async function executeEntry(input: ExecuteEntryInput): Promise<ExecutionResult> {
  const timeoutMs = input.timeoutMs ?? DEFAULT_TIMEOUT_MS;
  if (!Number.isFinite(timeoutMs) || timeoutMs <= 0) {
    throw new RangeError('timeoutMs must be a positive finite number.');
  }

  const now = input.now ?? Date.now;
  const started = now();
  const startedAt = new Date(started).toISOString();
  const candidate = input.module[input.entryFunction];

  if (typeof candidate !== 'function') {
    const finished = now();
    return {
      status: 'failed',
      entryFunction: input.entryFunction,
      startedAt,
      finishedAt: new Date(finished).toISOString(),
      durationMs: Math.max(0, finished - started),
      logs: getLogs(input.module),
      failure: {
        type: 'entry_not_found',
        message: `Entry function was not found: ${input.entryFunction}`,
      },
    };
  }

  input.module.Logger?.clear();

  let timeoutHandle: ReturnType<typeof setTimeout> | undefined;
  const timeout = new Promise<never>((_, reject) => {
    timeoutHandle = setTimeout(() => {
      const error = new Error(`Execution exceeded timeout of ${timeoutMs}ms.`);
      error.name = 'ExecutionTimeoutError';
      reject(error);
    }, timeoutMs);
  });

  try {
    const returnValue = await Promise.race([
      Promise.resolve((candidate as ScriptFunction)(...(input.args ?? []))),
      timeout,
    ]);
    const finished = now();

    return {
      status: 'succeeded',
      entryFunction: input.entryFunction,
      startedAt,
      finishedAt: new Date(finished).toISOString(),
      durationMs: Math.max(0, finished - started),
      returnValue,
      logs: getLogs(input.module),
    };
  } catch (error) {
    const finished = now();
    const timedOut = error instanceof Error && error.name === 'ExecutionTimeoutError';

    return {
      status: timedOut ? 'timed_out' : 'failed',
      entryFunction: input.entryFunction,
      startedAt,
      finishedAt: new Date(finished).toISOString(),
      durationMs: Math.max(0, finished - started),
      logs: getLogs(input.module),
      failure: timedOut
        ? {
            type: 'timeout',
            message: error.message,
            name: error.name,
          }
        : serializeFailure(error),
    };
  } finally {
    if (timeoutHandle) clearTimeout(timeoutHandle);
  }
}
