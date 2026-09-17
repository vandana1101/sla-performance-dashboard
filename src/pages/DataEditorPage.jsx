import React, { useMemo, useState } from "react";
import {
  ArrowLeft,
  ChevronDown,
  Download,
  Eye,
  EyeOff,
  FileSpreadsheet,
  RefreshCcw,
  Save,
  Search,
  Undo2,
  Redo2,
  X,
} from "lucide-react";
import * as XLSX from "xlsx";

function columnLetter(index) {
  let result = "";
  let n = index + 1;

  while (n > 0) {
    const rem = (n - 1) % 26;
    result = String.fromCharCode(65 + rem) + result;
    n = Math.floor((n - 1) / 26);
  }

  return result;
}

function normalizeWorkbook(raw) {
  if (!raw) return [];

  // Preferred structure:
  // raw.workbookSheets = [{ name, rows, headers }]
  if (Array.isArray(raw.workbookSheets)) {
    return raw.workbookSheets.map((sheet) => ({
      name: sheet.name,
      headers:
        sheet.headers ||
        (sheet.rows?.length ? Object.keys(sheet.rows[0]) : []),
      rows: sheet.rows || [],
    }));
  }

  // Common structure used by the current reader:
  // raw.sheets = { fg: [], lr: [], ... }
  if (raw.sheets && typeof raw.sheets === "object") {
    return Object.entries(raw.sheets).map(([name, rows]) => ({
      name,
      headers: rows?.length ? Object.keys(rows[0]) : [],
      rows: rows || [],
    }));
  }

  return [];
}

export default function DataEditorPage({
  raw,
  onBack,
  onSave,
  onRecalculate,
}) {
  const sheets = useMemo(() => normalizeWorkbook(raw), [raw]);

  const [activeSheet, setActiveSheet] = useState(0);
  const [data, setData] = useState(sheets);
  const [selected, setSelected] = useState({
    row: 0,
    col: 0,
  });
  const [editing, setEditing] = useState(false);
  const [editValue, setEditValue] = useState("");
  const [search, setSearch] = useState("");
  const [hiddenColumns, setHiddenColumns] = useState({});
  const [history, setHistory] = useState([]);
  const [future, setFuture] = useState([]);

  const sheet = data[activeSheet];

  const visibleHeaders = useMemo(() => {
    if (!sheet) return [];

    return sheet.headers
      .map((header, index) => ({
        header,
        index,
      }))
      .filter(
        ({ index }) =>
          !hiddenColumns[`${activeSheet}-${index}`]
      );
  }, [sheet, hiddenColumns, activeSheet]);

  const selectedHeader =
    sheet?.headers?.[selected.col] || "";

  const selectedValue =
    sheet?.rows?.[selected.row]?.[selectedHeader] ?? "";

  function pushHistory(nextData) {
    setHistory((prev) => [...prev.slice(-30), data]);
    setFuture([]);
    setData(nextData);
  }

  function updateCell(rowIndex, colIndex, value) {
    if (!sheet) return;

    const header = sheet.headers[colIndex];

    const next = data.map((s, si) => {
      if (si !== activeSheet) return s;

      return {
        ...s,
        rows: s.rows.map((row, ri) =>
          ri === rowIndex
            ? {
                ...row,
                [header]: value,
              }
            : row
        ),
      };
    });

    pushHistory(next);
  }

  function commitEdit() {
    updateCell(selected.row, selected.col, editValue);
    setEditing(false);
  }

  function selectCell(row, col) {
    setSelected({ row, col });
    setEditing(false);
  }

  function startEdit() {
    setEditValue(String(selectedValue ?? ""));
    setEditing(true);
  }

  function undo() {
    if (!history.length) return;

    const previous = history[history.length - 1];

    setFuture((prev) => [...prev, data]);
    setData(previous);
    setHistory((prev) => prev.slice(0, -1));
  }

  function redo() {
    if (!future.length) return;

    const next = future[future.length - 1];

    setHistory((prev) => [...prev, data]);
    setData(next);
    setFuture((prev) => prev.slice(0, -1));
  }

  function toggleColumn(colIndex) {
    const key = `${activeSheet}-${colIndex}`;

    setHiddenColumns((prev) => ({
      ...prev,
      [key]: !prev[key],
    }));
  }

  function exportWorkbook() {
    const workbook = XLSX.utils.book_new();

    data.forEach((s) => {
      const rows = s.rows.map((row) => {
        const ordered = {};

        s.headers.forEach((header) => {
          ordered[header] = row[header] ?? "";
        });

        return ordered;
      });

      const worksheet = XLSX.utils.json_to_sheet(rows, {
        header: s.headers,
      });

      XLSX.utils.book_append_sheet(
        workbook,
        worksheet,
        String(s.name).slice(0, 31)
      );
    });

    XLSX.writeFile(
      workbook,
      `${raw?.fileName || "SLA_Edited_Workbook"}.xlsx`
    );
  }

  function saveChanges() {
    if (onSave) {
      onSave({
        ...raw,
        workbookSheets: data,
        sheets: Object.fromEntries(
          data.map((s) => [s.name, s.rows])
        ),
      });
    }

    alert("Changes saved.");
  }

  function recalculate() {
    if (onRecalculate) {
      onRecalculate({
        ...raw,
        workbookSheets: data,
        sheets: Object.fromEntries(
          data.map((s) => [s.name, s.rows])
        ),
      });
    }

    alert("Dashboard recalculated from the edited data.");
  }

  if (!sheet) {
    return (
      <div className="data-editor-page">
        <div className="data-editor-empty">
          <FileSpreadsheet size={44} />
          <h2>No workbook data available</h2>
          <button onClick={onBack}>Back to Dashboard</button>
        </div>
      </div>
    );
  }

  return (
    <div className="data-editor-page">

      {/* =====================================================
          TOP BAR
      ===================================================== */}
      <header className="data-editor-topbar">

        <div className="editor-brand">
          <button
            className="editor-back"
            onClick={onBack}
            title="Back to Dashboard"
          >
            <ArrowLeft size={18} />
          </button>

          <div className="editor-file-icon">
            <FileSpreadsheet size={19} />
          </div>

          <div>
            <div className="editor-title">
              Workbook Data Editor
            </div>

            <div className="editor-subtitle">
              {raw?.fileName || "Uploaded Workbook"}
            </div>
          </div>
        </div>

        <div className="editor-actions">

          <button
            className="editor-tool-btn"
            onClick={undo}
            disabled={!history.length}
            title="Undo"
          >
            <Undo2 size={16} />
          </button>

          <button
            className="editor-tool-btn"
            onClick={redo}
            disabled={!future.length}
            title="Redo"
          >
            <Redo2 size={16} />
          </button>

          <div className="editor-divider" />

          <button
            className="editor-secondary-btn"
            onClick={recalculate}
          >
            <RefreshCcw size={15} />
            Recalculate Dashboard
          </button>

          <button
            className="editor-secondary-btn"
            onClick={exportWorkbook}
          >
            <Download size={15} />
            Export
          </button>

          <button
            className="editor-primary-btn"
            onClick={saveChanges}
          >
            <Save size={15} />
            Save Changes
          </button>
        </div>
      </header>

      {/* =====================================================
          SHEET TABS
      ===================================================== */}
      <div className="editor-sheet-tabs">

        {data.map((s, index) => (
          <button
            key={`${s.name}-${index}`}
            className={
              index === activeSheet
                ? "editor-sheet-tab active"
                : "editor-sheet-tab"
            }
            onClick={() => {
              setActiveSheet(index);
              setSelected({ row: 0, col: 0 });
              setEditing(false);
            }}
          >
            {s.name}
          </button>
        ))}

      </div>

      {/* =====================================================
          TOOLBAR
      ===================================================== */}
      <div className="editor-toolbar">

        <div className="editor-search">
          <Search size={15} />

          <input
            value={search}
            onChange={(e) => setSearch(e.target.value)}
            placeholder="Search this sheet..."
          />

          {search && (
            <button onClick={() => setSearch("")}>
              <X size={13} />
            </button>
          )}
        </div>

        <div className="editor-toolbar-spacer" />

        <div className="editor-sheet-info">
          <span>{sheet.rows.length.toLocaleString()} rows</span>
          <span>•</span>
          <span>{sheet.headers.length} columns</span>
        </div>

        <ColumnMenu
          headers={sheet.headers}
          hiddenColumns={hiddenColumns}
          activeSheet={activeSheet}
          onToggle={toggleColumn}
        />

      </div>

      {/* =====================================================
          FORMULA BAR
      ===================================================== */}
      <div className="editor-formula-bar">

        <div className="editor-name-box">
          {columnLetter(selected.col)}
          {selected.row + 2}
        </div>

        <div className="formula-symbol">
          fx
        </div>

        <input
          className="formula-input"
          value={
            editing
              ? editValue
              : String(selectedValue ?? "")
          }
          onChange={(e) => {
            if (!editing) {
              setEditing(true);
            }

            setEditValue(e.target.value);
          }}
          onKeyDown={(e) => {
            if (e.key === "Enter") {
              commitEdit();
            }

            if (e.key === "Escape") {
              setEditing(false);
            }
          }}
          onFocus={() => {
            if (!editing) {
              setEditValue(String(selectedValue ?? ""));
              setEditing(true);
            }
          }}
        />

        <div className="formula-column-name">
          {selectedHeader}
        </div>

      </div>

      {/* =====================================================
          GRID
      ===================================================== */}
      <div className="editor-grid-wrapper">

        <table className="editor-grid">

          <thead>
            <tr>
              <th className="row-number-header">
                #
              </th>

              {visibleHeaders.map(({ header, index }) => (
                <th
                  key={`${header}-${index}`}
                  className={
                    selected.col === index
                      ? "selected-column"
                      : ""
                  }
                >
                  <div className="column-header-content">

                    <span>
                      {columnLetter(index)}
                    </span>

                    <strong>
                      {header}
                    </strong>

                  </div>
                </th>
              ))}
            </tr>
          </thead>

          <tbody>
            {sheet.rows.map((row, rowIndex) => {

              const rowText = Object.values(row)
                .join(" ")
                .toLowerCase();

              if (
                search &&
                !rowText.includes(search.toLowerCase())
              ) {
                return null;
              }

              return (
                <tr key={rowIndex}>

                  <td
                    className={
                      selected.row === rowIndex
                        ? "row-number selected-row"
                        : "row-number"
                    }
                  >
                    {rowIndex + 2}
                  </td>

                  {visibleHeaders.map(
                    ({ header, index }) => {

                      const value =
                        row[header] ?? "";

                      const isSelected =
                        selected.row === rowIndex &&
                        selected.col === index;

                      const isFormula =
                        typeof value === "string" &&
                        value.startsWith("=");

                      return (
                        <td
                          key={`${rowIndex}-${index}`}
                          className={[
                            isSelected
                              ? "cell selected-cell"
                              : "cell",
                            isFormula
                              ? "formula-cell"
                              : "",
                          ].join(" ")}
                          onClick={() =>
                            selectCell(rowIndex, index)
                          }
                          onDoubleClick={() => {
                            selectCell(rowIndex, index);
                            setEditValue(
                              String(value)
                            );
                            setEditing(true);
                          }}
                        >

                          {isSelected && editing ? (
                            <input
                              autoFocus
                              className="cell-editor"
                              value={editValue}
                              onChange={(e) =>
                                setEditValue(
                                  e.target.value
                                )
                              }
                              onBlur={commitEdit}
                              onKeyDown={(e) => {
                                if (e.key === "Enter") {
                                  commitEdit();
                                }

                                if (
                                  e.key === "Escape"
                                ) {
                                  setEditing(false);
                                }
                              }}
                            />
                          ) : (
                            <span>
                              {String(value)}
                            </span>
                          )}

                        </td>
                      );
                    }
                  )}

                </tr>
              );
            })}
          </tbody>

        </table>

      </div>

      {/* =====================================================
          FOOTER
      ===================================================== */}
      <footer className="editor-footer">

        <div>
          Sheet: <strong>{sheet.name}</strong>
        </div>

        <div>
          Selected:{" "}
          <strong>
            {selectedHeader || "—"}
          </strong>
        </div>

        <div className="editor-footer-status">
          <span className="status-dot" />
          Workbook ready
        </div>

      </footer>

    </div>
  );
}


function ColumnMenu({
  headers,
  hiddenColumns,
  activeSheet,
  onToggle,
}) {
  const [open, setOpen] = useState(false);

  return (
    <div className="column-menu-wrapper">

      <button
        className="editor-secondary-btn"
        onClick={() => setOpen((v) => !v)}
      >
        <EyeOff size={15} />
        Columns
        <ChevronDown size={14} />
      </button>

      {open && (
        <div className="column-menu">

          <div className="column-menu-title">
            Show / Hide Columns
          </div>

          {headers.map((header, index) => {

            const hidden =
              hiddenColumns[
                `${activeSheet}-${index}`
              ];

            return (
              <button
                key={`${header}-${index}`}
                onClick={() => onToggle(index)}
              >
                {hidden ? (
                  <EyeOff size={14} />
                ) : (
                  <Eye size={14} />
                )}

                <span>
                  {columnLetter(index)} — {header}
                </span>
              </button>
            );
          })}

        </div>
      )}

    </div>
  );
}