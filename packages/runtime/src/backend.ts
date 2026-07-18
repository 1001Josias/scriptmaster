export interface AppsScriptRuntimeBindings {
  Logger?: GoogleAppsScript.Base.Logger;
  SpreadsheetApp?: GoogleAppsScript.Spreadsheet.SpreadsheetApp;
  DriveApp?: GoogleAppsScript.Drive.DriveApp;
  GmailApp?: GoogleAppsScript.Gmail.GmailApp;
  Utilities?: GoogleAppsScript.Utilities.Utilities;
  ScriptApp?: GoogleAppsScript.Script.ScriptApp;
  [service: string]: unknown;
}

export interface RuntimeBackendContext {
  services: readonly string[];
}

export interface RuntimeBackend {
  readonly name: string;
  initialize(context: RuntimeBackendContext): Promise<AppsScriptRuntimeBindings>;
}

export interface CreateRuntimeBindingsOptions {
  backend: RuntimeBackend;
  services: readonly string[];
}

export async function createRuntimeBindings(
  options: CreateRuntimeBindingsOptions,
): Promise<AppsScriptRuntimeBindings> {
  const services = [...new Set(options.services)].sort();
  const bindings = await options.backend.initialize({ services });

  for (const service of services) {
    if (!(service in bindings) || bindings[service] === undefined) {
      throw new Error(
        `Runtime backend "${options.backend.name}" did not provide the requested ${service} binding.`,
      );
    }
  }

  return bindings;
}
