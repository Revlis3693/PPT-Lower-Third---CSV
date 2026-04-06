import React, { useCallback, useEffect, useMemo, useState } from "react";
import type { CsvRow, ShapeMapping, StatusMessage } from "../types";
import { getCharThresholdForColumn } from "../config/columnCharThresholds";
import { parseCsv } from "../services/csv";
import { cleanRow, dedupeRows, dropRowsWithNoCellData } from "../services/dataTransforms";
import {
  applyRowToMappingsOnSlide,
  duplicateTemplateSlideAndPopulate,
  getTemplateSlideIdFromMappings,
  getTemplateSlideIdFromSelection,
  isPowerPointApiSupported,
  listMappableShapesOnTemplateSlide
} from "../services/pptService";

function uuid(): string {
  return Math.random().toString(16).slice(2) + "-" + Date.now().toString(16);
}

function StatusBox({ messages }: { messages: StatusMessage[] }) {
  if (!messages.length) return <div className="muted">No messages.</div>;
  return (
    <div className="status">
      {messages.map((m, idx) => {
        const prefix = m.kind.toUpperCase();
        return (
          <div key={idx}>
            [{prefix}] {m.message}
          </div>
        );
      })}
    </div>
  );
}

export default function App() {
  const [headers, setHeaders] = useState<string[]>([]);
  /** Raw CSV rows; never mutated for cleaning/dedupe toggles. */
  const [rawRows, setRawRows] = useState<CsvRow[]>([]);

  const [mappings, setMappings] = useState<ShapeMapping[]>([]);
  const [selectedColumn, setSelectedColumn] = useState<string>("");
  const [rowIndex, setRowIndex] = useState<number>(0);
  const [status, setStatus] = useState<StatusMessage[]>([]);
  const [busy, setBusy] = useState<boolean>(false);
  const [progressText, setProgressText] = useState<string>("");
  /** True only while Generate slides is running (for spinner + progress row, not other busy work). */
  const [isGeneratingSlides, setIsGeneratingSlides] = useState(false);

  const [skipBlankValues, setSkipBlankValues] = useState<boolean>(false);
  const [removeDuplicates, setRemoveDuplicates] = useState<boolean>(false);
  const [cleanData, setCleanData] = useState<boolean>(false);

  /** Per effective-row index: include when generating slides (default all true when data loads). */
  const [selectedForGenerate, setSelectedForGenerate] = useState<boolean[]>([]);

  const [shapeChoices, setShapeChoices] = useState<Array<{ shapeId: string; shapeName: string }>>([]);
  const [selectedShapeId, setSelectedShapeId] = useState<string>("");
  const [templateSlideIdForMapping, setTemplateSlideIdForMapping] = useState<string | null>(null);

  const [fileInputKey, setFileInputKey] = useState(0);

  /** CSV parse + shape scan: show banner; avoid locking the whole UI during slide scan. */
  const [importPhase, setImportPhase] = useState<"idle" | "parsing" | "scanningShapes">("idle");

  const effectiveRows = useMemo(() => {
    if (!headers.length || !rawRows.length) return [];
    let r = rawRows.map((row) => (cleanData ? cleanRow(row, headers) : { ...row }));
    // Rows can become all-empty after cleaning (e.g. whitespace-only cells).
    r = dropRowsWithNoCellData(r, headers);
    if (removeDuplicates) r = dedupeRows(r, headers);
    return r;
  }, [rawRows, headers, cleanData, removeDuplicates]);

  useEffect(() => {
    const n = effectiveRows.length;
    setRowIndex((i) => (n === 0 ? 0 : Math.min(Math.max(0, i), n - 1)));
    setSelectedForGenerate((prev) => {
      if (prev.length === n) return prev;
      return new Array(n).fill(true);
    });
  }, [effectiveRows.length]);

  const dataLoaded = rawRows.length > 0;
  const canMap = dataLoaded && headers.length > 0;
  const canPreview = effectiveRows.length > 0 && mappings.length > 0;
  const canGenerate = effectiveRows.length > 0 && mappings.length > 0;

  const safeRowIndex = effectiveRows.length === 0 ? 0 : Math.min(Math.max(rowIndex, 0), effectiveRows.length - 1);
  const selectedRow = effectiveRows[safeRowIndex];

  /** True once `selectedForGenerate` length matches `effectiveRows` (avoids unchecked flash before sync). */
  const generateSelectionSynced =
    effectiveRows.length > 0 && selectedForGenerate.length === effectiveRows.length;

  function pushStatus(msg: StatusMessage) {
    setStatus((s) => [msg, ...s].slice(0, 50));
  }

  const previewRowAtIndex = useCallback(
    async (idx: number, silent?: boolean) => {
      const row = effectiveRows[idx];
      if (!row || !mappings.length) return;
      try {
        let slideId = getTemplateSlideIdFromMappings(mappings);
        if (!slideId) slideId = await getTemplateSlideIdFromSelection();
        await applyRowToMappingsOnSlide({
          slideId,
          row,
          mappings,
          skipBlankValues
        });
        if (!silent) pushStatus({ kind: "success", message: `Preview: row ${idx + 1} of ${effectiveRows.length}` });
      } catch (e: any) {
        pushStatus({ kind: "error", message: e?.message ?? String(e) });
      }
    },
    [effectiveRows, mappings, skipBlankValues]
  );

  /**
   * Reads shapes from the currently selected slide (same as Refresh shapes).
   * Errors are shown in status; does not throw (safe to call after CSV load).
   */
  async function runRefreshShapesFlow(): Promise<void> {
    await new Promise<void>((r) => window.setTimeout(() => r(), 0));
    try {
      const { templateSlideId, shapes } = await listMappableShapesOnTemplateSlide();
      setTemplateSlideIdForMapping(templateSlideId);
      setShapeChoices(shapes);
      setSelectedShapeId(shapes[0]?.shapeId ?? "");

      if (shapes.length === 0) {
        pushStatus({
          kind: "warning",
          message:
            "No text boxes, placeholders, or shapes found on the selected slide. Use Insert → Text Box, or name shapes in the Selection Pane (Home → Arrange → Selection Pane), then click Refresh shapes."
        });
      } else {
        pushStatus({ kind: "success", message: `Found ${shapes.length} shape(s) on this slide. Pick a CSV column and a shape, then Add mapping.` });
      }
    } catch (e: any) {
      pushStatus({ kind: "error", message: e?.message ?? String(e) });
    }
  }

  async function onCsvFileSelected(file: File | null) {
    if (!file) return;
    setImportPhase("parsing");
    try {
      setStatus([]);

      const parsed = await parseCsv(file);
      setHeaders(parsed.headers);
      setRawRows(parsed.rows);
      setRowIndex(0);
      setSelectedColumn(parsed.headers[0] ?? "");
      pushStatus({ kind: "success", message: `Loaded ${parsed.rows.length} rows with ${parsed.headers.length} columns.` });

      if (parsed.headers.length > 0 && parsed.rows.length > 0) {
        setImportPhase("scanningShapes");
        await runRefreshShapesFlow();
      }
    } catch (e: any) {
      pushStatus({ kind: "error", message: e?.message ?? String(e) });
    } finally {
      setImportPhase("idle");
    }
  }

  async function onRefreshShapes() {
    if (!canMap) {
      pushStatus({ kind: "warning", message: "Load a CSV first." });
      return;
    }
    setImportPhase("scanningShapes");
    try {
      await runRefreshShapesFlow();
    } finally {
      setImportPhase("idle");
    }
  }

  function onAddMapping() {
    if (!canMap) {
      pushStatus({ kind: "warning", message: "Load a CSV first." });
      return;
    }
    if (!templateSlideIdForMapping) {
      pushStatus({
        kind: "warning",
        message:
          "No shapes loaded yet. Select your template slide in the thumbnails, then click Refresh shapes (shapes also scan automatically when you load a CSV)."
      });
      return;
    }
    if (!selectedColumn) {
      pushStatus({ kind: "warning", message: "Choose a CSV column." });
      return;
    }
    if (!selectedShapeId) {
      pushStatus({ kind: "warning", message: "Choose a shape from the list, or refresh if the list is empty." });
      return;
    }

    const shape = shapeChoices.find((s) => s.shapeId === selectedShapeId);
    if (!shape) {
      pushStatus({ kind: "warning", message: "Selected shape not found. Click Refresh shapes again." });
      return;
    }

    const mapping: ShapeMapping = {
      id: uuid(),
      templateSlideId: templateSlideIdForMapping,
      shapeId: shape.shapeId,
      shapeName: shape.shapeName,
      columnName: selectedColumn,
      label: shape.shapeName || selectedColumn
    };

    setMappings((prev) => {
      const existingIdx = prev.findIndex((m) => m.templateSlideId === mapping.templateSlideId && m.shapeId === mapping.shapeId);
      if (existingIdx >= 0) {
        const clone = [...prev];
        clone[existingIdx] = { ...clone[existingIdx], ...mapping };
        return clone;
      }
      return [mapping, ...prev];
    });

    pushStatus({
      kind: "success",
      message: `Mapped shape "${shape.shapeName}" to CSV column "${selectedColumn}".`
    });
  }

  function handleTableRowClick(idx: number) {
    setRowIndex(idx);
    if (canPreview) void previewRowAtIndex(idx, true);
  }

  async function goPrev() {
    if (safeRowIndex <= 0) return;
    const n = safeRowIndex - 1;
    setRowIndex(n);
    if (canPreview) await previewRowAtIndex(n, true);
  }

  async function goNext() {
    if (safeRowIndex >= effectiveRows.length - 1) return;
    const n = safeRowIndex + 1;
    setRowIndex(n);
    if (canPreview) await previewRowAtIndex(n, true);
  }

  async function onPreviewButton() {
    if (!selectedRow) {
      pushStatus({ kind: "warning", message: "No row selected." });
      return;
    }
    try {
      setBusy(true);
      await previewRowAtIndex(safeRowIndex, true);
      pushStatus({ kind: "success", message: `Preview applied. Row ${safeRowIndex + 1} of ${effectiveRows.length}.` });
    } catch (e: any) {
      pushStatus({ kind: "error", message: e?.message ?? String(e) });
    } finally {
      setBusy(false);
    }
  }

  function onToggleGenerateRow(idx: number) {
    setSelectedForGenerate((prev) => {
      const next = [...prev];
      next[idx] = !next[idx];
      return next;
    });
  }

  function selectAllForGenerate() {
    const n = effectiveRows.length;
    if (n === 0) return;
    setSelectedForGenerate(new Array(n).fill(true));
  }

  function selectNoneForGenerate() {
    const n = effectiveRows.length;
    if (n === 0) return;
    setSelectedForGenerate(new Array(n).fill(false));
  }

  /** Rows checked for generation; at least one must be checked to run. */
  function rowsCheckedForGenerate(): CsvRow[] {
    if (!generateSelectionSynced) return effectiveRows;
    return effectiveRows.filter((_, i) => selectedForGenerate[i]);
  }

  async function onGenerateSlides() {
    if (!effectiveRows.length) {
      pushStatus({ kind: "warning", message: "Load a CSV file first." });
      return;
    }
    const rows = rowsCheckedForGenerate();
    if (!rows.length) {
      pushStatus({ kind: "warning", message: "No rows selected. Use the checkboxes or Select all, then try again." });
      return;
    }
    try {
      setBusy(true);
      setIsGeneratingSlides(true);
      setStatus([]);
      setProgressText("Working…");

      const templateSlideId = await getTemplateSlideIdFromSelection();
      const res = await duplicateTemplateSlideAndPopulate({
        templateSlideId,
        rows,
        mappings,
        skipBlankValues,
        onProgress: (msg) => setProgressText(msg)
      });

      const total = effectiveRows.length;
      const picked = rows.length;
      const scope =
        picked === total ? `all ${picked} row(s)` : `${picked} of ${total} row(s) (unchecked rows skipped)`;
      pushStatus({ kind: "success", message: `Generated ${res.created} slide(s) from ${scope}.` });
      for (const w of res.warnings) pushStatus({ kind: "warning", message: w });
    } catch (e: any) {
      pushStatus({ kind: "error", message: e?.message ?? String(e) });
    } finally {
      setBusy(false);
      setIsGeneratingSlides(false);
      setProgressText("");
    }
  }

  function onReset() {
    setBusy(false);
    setIsGeneratingSlides(false);
    setProgressText("");
    setHeaders([]);
    setRawRows([]);
    setMappings([]);
    setSelectedColumn("");
    setRowIndex(0);
    setStatus([]);
    setSkipBlankValues(false);
    setRemoveDuplicates(false);
    setCleanData(false);
    setSelectedForGenerate([]);
    setShapeChoices([]);
    setSelectedShapeId("");
    setTemplateSlideIdForMapping(null);
    setImportPhase("idle");
    setFileInputKey((k) => k + 1);
    pushStatus({ kind: "info", message: "Reset complete. Load a CSV again if needed." });
  }

  const envOk = isPowerPointApiSupported();

  const displayRowCount = effectiveRows.length;
  const rawRowCount = rawRows.length;

  return (
    <div className="app">
      <div className="header">
        <h1>Lower Third Builder</h1>
        <div className="row">
          <button
            type="button"
            onClick={onReset}
            disabled={busy || importPhase !== "idle"}
            title="Clears data and unlocks the pane if an operation got stuck"
          >
            Reset
          </button>
        </div>
      </div>

      {importPhase !== "idle" ? (
        <div className="importBanner" role="status" aria-live="polite">
          <span className="importBanner__spinner" aria-hidden="true" />
          <div className="importBanner__body">
            <div className="importBanner__title">
              {importPhase === "parsing" ? "Reading CSV file…" : "Scanning slide for shapes…"}
            </div>
            <div className="importBanner__detail">
              {importPhase === "parsing"
                ? "Parsing rows and columns."
                : "This can take a few seconds. The selected slide in the thumbnail pane is scanned for text boxes and placeholders."}
            </div>
          </div>
        </div>
      ) : null}

      {!envOk && (
        <div className="section">
          <h2>Host support</h2>
          <div className="muted">
            This add-in requires PowerPoint JavaScript APIs (PowerPointApi 1.5+) for slide and shape selection. Update Office if this message appears.
          </div>
        </div>
      )}

      <div className="section">
        <h2>1) Data Source</h2>
        <div className="row">
          <input
            key={fileInputKey}
            type="file"
            accept=".csv,text/csv"
            disabled={busy || importPhase !== "idle"}
            onChange={(e) => onCsvFileSelected(e.target.files?.[0] ?? null)}
          />
            <label className="row" title="When mapping to shapes: do not overwrite with empty CSV cells. Empty table rows are always omitted.">
              <input
                type="checkbox"
                checked={skipBlankValues}
                disabled={busy}
                onChange={(e) => setSkipBlankValues(e.target.checked)}
              />
              <span>Skip blank values (when filling shapes)</span>
            </label>
          <label className="row">
            <input
              type="checkbox"
              checked={removeDuplicates}
              disabled={busy}
              onChange={(e) => setRemoveDuplicates(e.target.checked)}
            />
            <span>Remove duplicate rows</span>
          </label>
          <label className="row">
            <input type="checkbox" checked={cleanData} disabled={busy} onChange={(e) => setCleanData(e.target.checked)} />
            <span>Clean data</span>
          </label>
        </div>
        {!dataLoaded ? (
          <div className="muted" style={{ marginTop: 8 }}>
            Load a CSV with headers. Example headers: <code>Name</code>, <code>Title</code>, <code>Company</code>.
          </div>
        ) : (
          <div style={{ marginTop: 8 }}>
            <div className="muted">
              Columns: {headers.length} · Raw rows: {rawRowCount}
              {removeDuplicates || cleanData ? (
                <>
                  {" "}
                  · Effective rows: {displayRowCount}
                  {cleanData ? " (cleaned)" : ""}
                  {removeDuplicates ? " (deduped)" : ""}
                </>
              ) : null}
            </div>

            <div className="muted" style={{ marginTop: 8, marginBottom: 4 }}>
              Preview data — click a row to select and preview on the slide. Checkboxes choose which rows are included when you
              generate slides (all are selected by default). ⚠️ = over character limit (see <code>columnCharThresholds.ts</code>).
            </div>
            <div className="previewToolbar row">
              <button type="button" onClick={() => void goPrev()} disabled={busy || safeRowIndex <= 0 || !canPreview}>
                Prev
              </button>
              <button type="button" onClick={() => void goNext()} disabled={busy || safeRowIndex >= effectiveRows.length - 1 || !canPreview}>
                Next
              </button>
              <span className="muted">
                Row {effectiveRows.length ? safeRowIndex + 1 : 0} of {effectiveRows.length}
              </span>
              <span className="previewToolbar__sep muted" aria-hidden="true">
                |
              </span>
              <button type="button" onClick={selectAllForGenerate} disabled={busy || effectiveRows.length === 0}>
                Select all
              </button>
              <button type="button" onClick={selectNoneForGenerate} disabled={busy || effectiveRows.length === 0}>
                Select none
              </button>
            </div>
            <div className="previewTableWrap">
              <table className="previewTable">
                <thead>
                  <tr>
                    <th className="previewTable__check" title="Include this row when generating slides">
                      Include
                    </th>
                    <th className="previewTable__idx">#</th>
                    {headers.map((h) => (
                      <th key={h}>{h}</th>
                    ))}
                  </tr>
                </thead>
                <tbody>
                  {effectiveRows.map((row, idx) => (
                    <tr
                      key={idx}
                      className={idx === safeRowIndex ? "previewRow previewRow--selected" : "previewRow"}
                      onClick={() => handleTableRowClick(idx)}
                    >
                      <td className="previewTable__check" onClick={(e) => e.stopPropagation()}>
                        <input
                          type="checkbox"
                          checked={generateSelectionSynced ? !!selectedForGenerate[idx] : true}
                          disabled={busy}
                          onChange={() => onToggleGenerateRow(idx)}
                          aria-label={`Include row ${idx + 1} when generating slides`}
                        />
                      </td>
                      <td className="previewTable__idx muted">{idx + 1}</td>
                      {headers.map((h) => {
                        const cell = row[h] ?? "";
                        const len = cell.length;
                        const max = getCharThresholdForColumn(h);
                        const warn = len > max;
                        return (
                          <td key={h}>
                            {warn ? <span title={`${len} chars (limit ${max})`}>⚠️ </span> : null}
                            {cell}
                          </td>
                        );
                      })}
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          </div>
        )}
      </div>

      <div className="section">
        <h2>2) Column Mapping</h2>
        {!dataLoaded ? (
          <div className="muted">Load a CSV first, then map columns to shapes on your template slide.</div>
        ) : (
          <>
            <div className="muted" style={{ marginBottom: 10 }}>
              Map each CSV column to a text box on your template slide. With your template slide selected in the thumbnails, shapes are scanned automatically when you load a CSV; use <b>Refresh shapes</b> again if you switch slides. Then pick a column and a shape, then <b>Add mapping</b>.
            </div>
            <div className="row">
              <button className="primary" onClick={onRefreshShapes} disabled={busy || importPhase !== "idle" || !canMap}>
                Refresh shapes
              </button>
              <span className="muted">Uses the slide selected in the thumbnail pane (one slide only).</span>
            </div>
            <div className="row" style={{ marginTop: 10 }}>
              <div className="muted">CSV column</div>
              <select value={selectedColumn} disabled={busy || !headers.length} onChange={(e) => setSelectedColumn(e.target.value)}>
                {headers.map((h) => (
                  <option key={h} value={h}>
                    {h}
                  </option>
                ))}
              </select>
              <div className="muted">Shape on slide</div>
              <select
                value={selectedShapeId}
                disabled={busy || shapeChoices.length === 0}
                onChange={(e) => setSelectedShapeId(e.target.value)}
              >
                {shapeChoices.length === 0 ? (
                  <option value="">— Load CSV or Refresh shapes —</option>
                ) : (
                  shapeChoices.map((s) => (
                    <option key={s.shapeId} value={s.shapeId}>
                      {s.shapeName}
                    </option>
                  ))
                )}
              </select>
              <button className="primary" onClick={onAddMapping} disabled={busy || !canMap || !templateSlideIdForMapping || shapeChoices.length === 0}>
                Add mapping
              </button>
            </div>

            <div style={{ marginTop: 10 }}>
              {mappings.length === 0 ? (
                <div className="muted">
                  <b>No mappings yet.</b> Load a CSV (shapes scan automatically) or click <b>Refresh shapes</b>, then add mappings for each field.
                </div>
              ) : (
                <div className="mappingList">
                  {mappings.map((m) => (
                    <div key={m.id} className="mappingItem">
                      <div className="mappingMeta">
                        <div>
                          <b>{m.label ?? m.columnName}</b> <span className="muted">→</span> <code>{m.columnName}</code>
                        </div>
                        <button
                          type="button"
                          onClick={() => setMappings((prev) => prev.filter((x) => x.id !== m.id))}
                          disabled={busy}
                        >
                          Remove
                        </button>
                      </div>
                      <div className="row">
                        <span className="muted">Label</span>
                        <input
                          type="text"
                          value={m.label ?? ""}
                          disabled={busy}
                          onChange={(e) =>
                            setMappings((prev) => prev.map((x) => (x.id === m.id ? { ...x, label: e.target.value } : x)))
                          }
                        />
                      </div>
                      <div className="muted">
                        Shape: <code>{m.shapeId}</code>
                        {m.shapeName ? (
                          <>
                            {" "}
                            · Name: <code>{m.shapeName}</code>
                          </>
                        ) : null}
                      </div>
                    </div>
                  ))}
                </div>
              )}
            </div>
          </>
        )}
      </div>

      <div className="section">
        <h2>3) Preview</h2>
        {!dataLoaded ? (
          <div className="muted">Load a CSV first.</div>
        ) : mappings.length === 0 ? (
          <div className="muted">Create at least one mapping first.</div>
        ) : (
          <>
            <div className="muted" style={{ marginBottom: 8 }}>
              The table in section 1 updates the slide when you click a row (or use Prev/Next). Use the button below if you need an explicit refresh.
            </div>
            <div className="row">
              <button className="primary" type="button" onClick={() => void onPreviewButton()} disabled={busy || !canPreview}>
                Preview current row on slide
              </button>
            </div>
          </>
        )}
      </div>

      <div className="section">
        <h2>4) Generate</h2>
        <div className="muted" style={{ marginBottom: 8 }}>
          Uses the currently selected slide as the template. Only rows checked under <b>Include</b> in the table above are turned
          into slides. Use <b>Select all</b> / <b>Select none</b> to adjust quickly.
        </div>
        <div className="row">
          <button className="primary" type="button" onClick={() => void onGenerateSlides()} disabled={busy || !canGenerate}>
            Generate slides
          </button>
        </div>
        {isGeneratingSlides ? (
          <div className="generateProgress" role="status" aria-live="polite">
            <span className="generateProgress__spinner" aria-hidden="true" />
            <span className="generateProgress__text">{progressText || "Working…"}</span>
          </div>
        ) : null}
      </div>

      <div className="section">
        <h2>Status</h2>
        <StatusBox messages={status} />
      </div>
    </div>
  );
}
