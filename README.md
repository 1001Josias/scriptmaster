# ScriptMaster

ScriptMaster turns Google Apps Script projects into analyzable, executable Node.js projects without requiring source-level rewrites of supported APIs.

## End-to-end demo

The Build Week scenario covers source analysis, deterministic compatibility findings, AI migration guidance, Node project generation, execution, logs and result output.

```bash
pnpm install
pnpm --filter @scriptmaster/build-week-demo demo
```

The demo uses an in-memory `SpreadsheetApp` backend, so no Google credentials are required. See [`examples/build-week-demo`](examples/build-week-demo/README.md) for expected output, failure behavior and the production runtime path through `gasFakesBackend`.

## Packages

- `@scriptmaster/analyzer` — Apps Script symbol detection and compatibility reports.
- `@scriptmaster/ai` — optional provider-neutral migration suggestions.
- `@scriptmaster/compiler` — deterministic Node.js project generation.
- `@scriptmaster/runtime` — Apps Script-compatible runtime bindings and backends.
- `@scriptmaster/executor` — structured entry-function execution.
- `apps/web` — browser-based source editor and compatibility report.
