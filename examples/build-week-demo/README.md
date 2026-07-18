# ScriptMaster Build Week demo

This scenario demonstrates the product flow without Google credentials or manual source edits:

1. analyze a Google Apps Script;
2. generate the deterministic compatibility report;
3. display clearly separated AI migration guidance;
4. generate the Node.js project files;
5. execute `main` with an in-memory `SpreadsheetApp` backend;
6. display the return value and structured Logger entries.

## Run

From the repository root:

```bash
pnpm install
pnpm --filter @scriptmaster/build-week-demo demo
```

Expected business result:

```json
{
  "orderCount": 2,
  "total": 200
}
```

The demo intentionally uses an in-memory backend. Production execution uses the runtime backend abstraction and can select `gasFakesBackend` with Google credentials inside an isolated process or container.

## Failure behavior

- Analyzer findings remain deterministic even when AI is unavailable.
- AI output is labeled as non-deterministic guidance.
- Unexpected spreadsheet IDs, sheet names, or ranges fail with explicit diagnostics.
- The executor returns structured `failed` or `timed_out` results rather than hiding runtime errors.

The generated entry function is the future boundary for packaging the output as a BotMaster Worker.
