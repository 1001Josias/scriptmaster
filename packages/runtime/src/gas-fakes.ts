import type { AppsScriptRuntimeBindings, RuntimeBackend } from './backend.js';

const SUPPORTED_GLOBALS = [
  'Logger',
  'SpreadsheetApp',
  'DriveApp',
  'GmailApp',
  'Utilities',
  'ScriptApp',
] as const;

type GasFakesLoader = () => Promise<unknown>;

export interface CreateGasFakesBackendOptions {
  load?: GasFakesLoader;
  globalObject?: Record<string, unknown>;
}

function defaultLoader(): Promise<unknown> {
  return import('@mcpher/gas-fakes');
}

export function createGasFakesBackend(
  options: CreateGasFakesBackendOptions = {},
): RuntimeBackend {
  const load = options.load ?? defaultLoader;
  const globalObject = options.globalObject ?? (globalThis as Record<string, unknown>);
  let initialized: Promise<void> | null = null;

  return {
    name: 'gas-fakes',

    async initialize(context): Promise<AppsScriptRuntimeBindings> {
      initialized ??= load().then(() => undefined);
      await initialized;

      const requested = new Set(context.services);
      const bindings: AppsScriptRuntimeBindings = {};

      for (const service of SUPPORTED_GLOBALS) {
        if (requested.has(service) && globalObject[service] !== undefined) {
          (bindings as Record<string, unknown>)[service] = globalObject[service];
        }
      }

      return bindings;
    },
  };
}

export const gasFakesBackend = createGasFakesBackend();
