import { useMemo, useRef, useState } from 'react';

import {
  analyzeAppsScript,
  generateCompatibilityReport,
  type CompatibilityItem,
  type CompatibilityReport,
} from '@scriptmaster/analyzer';

const SAMPLE_SOURCE = `function main() {
  Logger.log('Starting migration');
  const spreadsheet = SpreadsheetApp.openById('spreadsheet-id');
  DriveApp.getFiles();
}`;

const STATUS_LABELS = {
  supported: 'Supported',
  partially_supported: 'Partial',
  unsupported: 'Blocking',
  unknown: 'Unknown',
} as const;

function symbolName(item: CompatibilityItem): string {
  return item.service ? `${item.service}.${item.name}` : item.name;
}

export function App() {
  const [source, setSource] = useState(SAMPLE_SOURCE);
  const [report, setReport] = useState<CompatibilityReport | null>(null);
  const [error, setError] = useState<string | null>(null);
  const editorRef = useRef<HTMLTextAreaElement>(null);

  const blockingCount = useMemo(
    () => report?.items.filter((item) => item.status === 'unsupported' || item.status === 'unknown').length ?? 0,
    [report],
  );

  function analyze(): void {
    if (!source.trim()) {
      setReport(null);
      setError('Paste a Google Apps Script before running the analysis.');
      return;
    }

    try {
      setError(null);
      setReport(generateCompatibilityReport(analyzeAppsScript(source, 'script.js')));
    } catch (cause) {
      setReport(null);
      setError(cause instanceof Error ? cause.message : 'The script could not be analyzed.');
    }
  }

  function focusLocation(item: CompatibilityItem): void {
    const editor = editorRef.current;
    if (!editor) return;

    const lines = source.split('\n');
    const offset = lines.slice(0, item.location.line - 1).reduce((total, line) => total + line.length + 1, 0);
    const position = offset + item.location.column - 1;
    editor.focus();
    editor.setSelectionRange(position, position + item.name.length);
  }

  const reportBadge = report
    ? report.assessment.scoreReliable
      ? `${report.summary.score}%`
      : report.assessment.status === 'blocked'
        ? 'Blocked'
        : 'Incomplete'
    : null;

  return (
    <main className="app-shell">
      <header className="hero">
        <div>
          <p className="eyebrow">Google Apps Script migration</p>
          <h1>ScriptMaster</h1>
          <p className="hero-copy">Paste a script, identify migration blockers, and understand what can move to Node.js today.</p>
        </div>
        <div className="hero-badge">Analysis runs locally</div>
      </header>

      <section className="workspace" aria-label="Script compatibility workspace">
        <article className="panel editor-panel">
          <div className="panel-heading">
            <div>
              <p className="step">01</p>
              <h2>Source script</h2>
            </div>
            <button className="secondary" type="button" onClick={() => setSource(SAMPLE_SOURCE)}>Load example</button>
          </div>
          <textarea
            ref={editorRef}
            aria-label="Google Apps Script source code"
            spellCheck={false}
            value={source}
            onChange={(event) => setSource(event.target.value)}
          />
          <div className="editor-footer">
            <span>{source.split('\n').length} lines</span>
            <button className="primary" type="button" onClick={analyze}>Analyze compatibility</button>
          </div>
          {error && <p className="error-state" role="alert">{error}</p>}
        </article>

        <article className="panel report-panel">
          <div className="panel-heading">
            <div>
              <p className="step">02</p>
              <h2>Migration report</h2>
            </div>
            {report && (
              <div
                className={`score ${blockingCount || !report.assessment.scoreReliable ? 'score-blocked' : ''}`}
                title={report.assessment.scoreReliable ? 'Compatibility score' : 'Compatibility score is not reliable because parsing did not complete cleanly'}
              >
                {reportBadge}
              </div>
            )}
          </div>

          {!report ? (
            <div className="empty-state">
              <strong>No analysis yet</strong>
              <span>Your compatibility summary and source-linked findings will appear here.</span>
            </div>
          ) : (
            <>
              <div className="summary-grid">
                <Summary label="Supported" value={report.summary.supported} />
                <Summary label="Partial" value={report.summary.partiallySupported} />
                <Summary label="Blocking" value={report.summary.unsupported + report.summary.unknown} />
              </div>

              {report.diagnostics.length > 0 && (
                <div className="diagnostics" role={report.assessment.status === 'blocked' ? 'alert' : undefined}>
                  <strong>
                    {report.assessment.status === 'blocked'
                      ? 'Parser errors block a reliable compatibility result'
                      : 'Parser diagnostics make this report incomplete'}
                  </strong>
                  {report.diagnostics.map((diagnostic, index) => (
                    <p key={`${diagnostic.code}-${index}`}>{diagnostic.message}</p>
                  ))}
                </div>
              )}

              <div className="findings">
                {report.items.length === 0 ? (
                  <div className="empty-state compact"><span>No Apps Script services were detected.</span></div>
                ) : report.items.map((item, index) => (
                  <button
                    className="finding"
                    type="button"
                    key={`${symbolName(item)}-${item.location.line}-${item.location.column}-${index}`}
                    onClick={() => focusLocation(item)}
                  >
                    <span className={`status status-${item.status}`}>{STATUS_LABELS[item.status]}</span>
                    <span className="finding-copy">
                      <strong>{symbolName(item)}</strong>
                      <small>Line {item.location.line}, column {item.location.column}</small>
                      <span>{item.note}</span>
                    </span>
                  </button>
                ))}
              </div>
            </>
          )}
        </article>
      </section>
    </main>
  );
}

function Summary({ label, value }: { label: string; value: number }) {
  return <div className="summary-card"><span>{label}</span><strong>{value}</strong></div>;
}
