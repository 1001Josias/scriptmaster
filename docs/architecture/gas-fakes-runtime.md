# gas-fakes runtime backend

ScriptMaster uses `@mcpher/gas-fakes` behind the `RuntimeBackend` contract to preserve synchronous Google Apps Script service calls while running generated projects on Node.js.

## Boundary

- The compiler depends only on `@scriptmaster/runtime`.
- `@scriptmaster/runtime` owns backend initialization and validates requested globals.
- `gas-fakes` is loaded once and exposes its emulated Apps Script services through `globalThis`.
- `Logger` remains ScriptMaster-owned so execution logs keep the existing structured history contract.
- Generated projects use `@types/google-apps-script` as the compile-time API contract.

## Current limitations

`gas-fakes` reads authentication and runtime configuration from process-level environment and installs process-global service objects. Therefore, the in-process executor must not run concurrent executions with different credentials or tenant identities.

Production execution must use the isolated execution backend tracked in issue #18. Each worker process or container should receive one execution identity, a restricted environment, explicit scopes, and hard timeout/resource limits.

The legacy native `createSpreadsheetApp` adapter remains temporarily available for compatibility, but new generated projects use the gas-fakes backend.
