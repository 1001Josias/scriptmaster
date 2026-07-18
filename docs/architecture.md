# Architecture

ScriptMaster is a development experience for migrating and creating Node.js automations with Google Apps Script compatibility.

## MVP flow

1. Import or paste a Google Apps Script project.
2. Analyze supported and unsupported APIs.
3. Generate a standard Node.js project.
4. Execute the generated automation with logs and execution history.

## Boundaries

- `apps/web`: user interface and API composition.
- `packages/analyzer`: static analysis and compatibility reports.
- `packages/gas-compat`: runtime adapters backed by official Google APIs.
- `packages/compiler`: generation of standard Node.js projects.
- `examples`: migration and execution examples.

ScriptMaster does not emulate the Google Apps Script runtime. Compatibility is implemented through explicit libraries and adapters.

In the long-term product vision, generated projects can be published as BotMaster Workers. The Build Week MVP remains standalone.
