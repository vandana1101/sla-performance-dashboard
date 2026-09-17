import React, { useEffect, useMemo, useState } from "react";
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

/* =========================================================
   COLUMN LETTER
========================================================= */

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

/* =========================================================
   READ EXCEL FILE

   Converts an uploaded browser File into editable
   workbook sheet objects.
========================================================= */

async function readExcelFile(file, fileIndex) {
  const buffer = await file.arrayBuffer();

  const workbook = XLSX.read(buffer, {
    type: "array",
    cellFormula: true,
    cellNF: true,
    cellStyles: true,
    cellDates: true,
  });

  const sheets = workbook.SheetNames.map((sheetName) => {
    const ws = workbook.Sheets[sheetName];

    const range = XLSX.utils.decode_range(
      ws["!ref"] || "A1"
    );

    const headers = [];

    for (
      let c = range.s.c;
      c <= range.e.c;
      c++
    ) {
      const address = XLSX.utils.encode_cell({
        r: range.s.r,
        c,
      });

      const cell = ws[address];

      headers.push(
        cell?.v !== undefined &&
        cell?.v !== null &&
        cell?.v !== ""
          ? String(cell.v)
          : `Column ${c + 1}`
      );
    }

    const rows = [];

    for (
      let r = range.s.r + 1;
      r <= range.e.r;
      r++
    ) {
      const row = {};

      for (
        let c = range.s.c;
        c <= range.e.c;
        c++
      ) {
        const address = XLSX.utils.encode_cell({
          r,
          c,
        });

        const cell = ws[address];

        const header =
          headers[c - range.s.c];

        if (cell?.f) {
          row[header] = `=${cell.f}`;
        } else {
          row[header] =
            cell?.v !== undefined &&
            cell?.v !== null
              ? cell.v
              : "";
        }
      }

      rows.push(row);
    }

    return {
      name: sheetName,
      headers,
      rows,
      fileIndex,
    };
  });

  return {
    file,
    fileIndex,
    fileName: file.name,
    sheets,
  };
}

/* =========================================================
   MAIN DATA EDITOR
========================================================= */

export default function DataEditorPage({
  raw,
  onBack,
  onSave,
  onRecalculate,
}) {

  /* =======================================================
     WORKBOOK DATA
  ======================================================= */

  const [workbooks, setWorkbooks] = useState([]);

  const [loading, setLoading] = useState(true);

  const [loadError, setLoadError] = useState("");

  /* =======================================================
     EDITOR STATE
  ======================================================= */

  const [activeSheet, setActiveSheet] = useState(0);

  const [data, setData] = useState([]);

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

  /* =======================================================
     READ UPLOADED FILES

     IMPORTANT:
     raw is File[]
========================================================= */

  useEffect(() => {

    let cancelled = false;

    async function loadFiles() {

      setLoading(true);
      setLoadError("");

      try {

        if (!raw || !Array.isArray(raw) || !raw.length) {
          throw new Error(
            "No uploaded workbook files were found."
          );
        }

        const loaded = [];

        for (
          let i = 0;
          i < raw.length;
          i++
        ) {

          const file = raw[i];

          if (!(file instanceof File)) {
            continue;
          }

          const workbook =
            await readExcelFile(
              file,
              i
            );

          loaded.push(workbook);
        }

        if (!loaded.length) {
          throw new Error(
            "The uploaded files could not be read as Excel workbooks."
          );
        }

        if (cancelled) return;

        setWorkbooks(loaded);

        const allSheets =
          loaded.flatMap(
            (workbook) =>
              workbook.sheets
          );

        setData(allSheets);

        setActiveSheet(0);

        setSelected({
          row: 0,
          col: 0,
        });

        setHistory([]);

        setFuture([]);

      } catch (error) {

        if (cancelled) return;

        console.error(
          "Workbook loading error:",
          error
        );

        setLoadError(
          error?.message ||
          "Unable to read the workbook."
        );

      } finally {

        if (!cancelled) {
          setLoading(false);
        }

      }
    }

    loadFiles();

    return () => {
      cancelled = true;
    };

  }, [raw]);

  /* =======================================================
     ACTIVE SHEET
========================================================= */

  const sheet = data[activeSheet];

  /* =======================================================
     VISIBLE HEADERS
========================================================= */

  const visibleHeaders = useMemo(() => {

    if (!sheet) return [];

    return sheet.headers
      .map((header, index) => ({
        header,
        index,
      }))
      .filter(
        ({ index }) =>
          !hiddenColumns[
            `${activeSheet}-${index}`
          ]
      );

  }, [
    sheet,
    hiddenColumns,
    activeSheet,
  ]);

  /* =======================================================
     SELECTED CELL
========================================================= */

  const selectedHeader =
    sheet?.headers?.[
      selected.col
    ] || "";

  const selectedValue =
    sheet?.rows?.[
      selected.row
    ]?.[selectedHeader] ?? "";

  /* =======================================================
     HISTORY
========================================================= */

  function pushHistory(nextData) {

    setHistory((prev) => [
      ...prev.slice(-30),
      data,
    ]);

    setFuture([]);

    setData(nextData);
  }

  /* =======================================================
     UPDATE CELL
========================================================= */

  function updateCell(
    rowIndex,
    colIndex,
    value
  ) {

    if (!sheet) return;

    const header =
      sheet.headers[colIndex];

    const next =
      data.map(
        (currentSheet, sheetIndex) => {

          if (
            sheetIndex !==
            activeSheet
          ) {
            return currentSheet;
          }

          return {
            ...currentSheet,

            rows:
              currentSheet.rows.map(
                (row, currentRowIndex) => {

                  if (
                    currentRowIndex !==
                    rowIndex
                  ) {
                    return row;
                  }

                  return {
                    ...row,
                    [header]: value,
                  };
                }
              ),
          };
        }
      );

    pushHistory(next);
  }

  /* =======================================================
     COMMIT EDIT
========================================================= */

  function commitEdit() {

    updateCell(
      selected.row,
      selected.col,
      editValue
    );

    setEditing(false);
  }

  /* =======================================================
     SELECT CELL
========================================================= */

  function selectCell(
    row,
    col
  ) {

    setSelected({
      row,
      col,
    });

    setEditing(false);
  }

  /* =======================================================
     START EDIT
========================================================= */

  function startEdit() {

    setEditValue(
      String(
        selectedValue ?? ""
      )
    );

    setEditing(true);
  }

  /* =======================================================
     UNDO
========================================================= */

  function undo() {

    if (!history.length) return;

    const previous =
      history[
        history.length - 1
      ];

    setFuture((prev) => [
      ...prev,
      data,
    ]);

    setData(previous);

    setHistory((prev) =>
      prev.slice(0, -1)
    );
  }

  /* =======================================================
     REDO
========================================================= */

  function redo() {

    if (!future.length) return;

    const next =
      future[
        future.length - 1
      ];

    setHistory((prev) => [
      ...prev,
      data,
    ]);

    setData(next);

    setFuture((prev) =>
      prev.slice(0, -1)
    );
  }

  /* =======================================================
     TOGGLE COLUMN
========================================================= */

  function toggleColumn(
    colIndex
  ) {

    const key =
      `${activeSheet}-${colIndex}`;

    setHiddenColumns((prev) => ({
      ...prev,
      [key]: !prev[key],
    }));
  }

  /* =======================================================
     BUILD WORKBOOK FILES

     Converts edited data back into actual .xlsx Files.

     This is the critical bridge between:
       Editor → Dashboard
========================================================= */

  function buildEditedFiles() {

    const result = [];

    workbooks.forEach(
      (originalWorkbook) => {

        const workbook =
          XLSX.utils.book_new();

        const workbookSheets =
          data.filter(
            (sheetData) =>
              sheetData.fileIndex ===
              originalWorkbook.fileIndex
          );

        workbookSheets.forEach(
          (sheetData) => {

            const rows =
              sheetData.rows.map(
                (row) => {

                  const ordered = {};

                  sheetData.headers.forEach(
                    (header) => {

                      ordered[header] =
                        row[header] ??
                        "";

                    }
                  );

                  return ordered;
                }
              );

            const worksheet =
              XLSX.utils.json_to_sheet(
                rows,
                {
                  header:
                    sheetData.headers,
                }
              );

            XLSX.utils.book_append_sheet(
              workbook,
              worksheet,
              String(
                sheetData.name
              ).slice(0, 31)
            );

          }
        );

        const arrayBuffer =
          XLSX.write(
            workbook,
            {
              bookType: "xlsx",
              type: "array",
            }
          );

        const editedFile =
          new File(
            [
              arrayBuffer,
            ],
            originalWorkbook.fileName,
            {
              type:
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            }
          );

        result.push(
          editedFile
        );
      }
    );

    return result;
  }

  /* =======================================================
     SAVE CHANGES

     Sends File[] back to App.
========================================================= */

  function saveChanges() {

    try {

      const editedFiles =
        buildEditedFiles();

      if (!editedFiles.length) {
        alert(
          "No workbook files available to save."
        );

        return;
      }

      if (onSave) {
        onSave(
          editedFiles
        );
      }

      alert(
        "Changes saved successfully."
      );

    } catch (error) {

      console.error(
        "Save error:",
        error
      );

      alert(
        "Unable to save the edited workbook."
      );
    }
  }

  /* =======================================================
     RECALCULATE DASHBOARD

     Creates edited File[] and sends them
     back through the dashboard calculation
     pipeline.
========================================================= */

  function recalculate() {

    try {

      const editedFiles =
        buildEditedFiles();

      if (!editedFiles.length) {
        alert(
          "No workbook files available."
        );

        return;
      }

      if (onRecalculate) {
        onRecalculate(
          editedFiles
        );
      } else if (onSave) {
        onSave(
          editedFiles
        );
      }

      alert(
        "Dashboard recalculated from the edited data."
      );

    } catch (error) {

      console.error(
        "Recalculation error:",
        error
      );

      alert(
        "Unable to recalculate the dashboard."
      );
    }
  }

  /* =======================================================
     EXPORT WORKBOOK

     Downloads the currently selected workbook
     after edits.
========================================================= */

  function exportWorkbook() {

    try {

      const editedFiles =
        buildEditedFiles();

      if (!editedFiles.length) {
        alert(
          "No workbook available to export."
        );

        return;
      }

      editedFiles.forEach(
        (file) => {

          const url =
            URL.createObjectURL(
              file
            );

          const anchor =
            document.createElement(
              "a"
            );

          anchor.href = url;

          anchor.download =
            file.name;

          document.body.appendChild(
            anchor
          );

          anchor.click();

          document.body.removeChild(
            anchor
          );

          URL.revokeObjectURL(
            url
          );
        }
      );

    } catch (error) {

      console.error(
        "Export error:",
        error
      );

      alert(
        "Unable to export workbook."
      );
    }
  }

  /* =======================================================
     LOADING STATE
========================================================= */

  if (loading) {

    return (
      <div className="data-editor-page">

        <div className="data-editor-empty">

          <FileSpreadsheet
            size={44}
          />

          <h2>
            Loading workbook...
          </h2>

          <p>
            Reading your uploaded Excel data.
          </p>

        </div>

      </div>
    );
  }

  /* =======================================================
     ERROR STATE
========================================================= */

  if (loadError) {

    return (
      <div className="data-editor-page">

        <div className="data-editor-empty">

          <FileSpreadsheet
            size={44}
          />

          <h2>
            Unable to load workbook
          </h2>

          <p>
            {loadError}
          </p>

          <button
            onClick={onBack}
          >
            Back to Dashboard
          </button>

        </div>

      </div>
    );
  }

  /* =======================================================
     EMPTY STATE
========================================================= */

  if (!sheet) {

    return (
      <div className="data-editor-page">

        <div className="data-editor-empty">

          <FileSpreadsheet
            size={44}
          />

          <h2>
            No workbook data available
          </h2>

          <button
            onClick={onBack}
          >
            Back to Dashboard
          </button>

        </div>

      </div>
    );
  }

  /* =======================================================
     EDITOR
========================================================= */

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
            <ArrowLeft
              size={18}
            />
          </button>

          <div className="editor-file-icon">

            <FileSpreadsheet
              size={19}
            />

          </div>

          <div>

            <div className="editor-title">
              Workbook Data Editor
            </div>

            <div className="editor-subtitle">

              {workbooks.length === 1
                ? workbooks[0].fileName
                : `${workbooks.length} uploaded workbooks`}

            </div>

          </div>

        </div>

        <div className="editor-actions">

          {/* UNDO */}

          <button
            className="editor-tool-btn"
            onClick={undo}
            disabled={!history.length}
            title="Undo"
          >
            <Undo2
              size={16}
            />
          </button>

          {/* REDO */}

          <button
            className="editor-tool-btn"
            onClick={redo}
            disabled={!future.length}
            title="Redo"
          >
            <Redo2
              size={16}
            />
          </button>

          <div className="editor-divider" />

          {/* RECALCULATE */}

          <button
            className="editor-secondary-btn"
            onClick={recalculate}
          >
            <RefreshCcw
              size={15}
            />

            Recalculate Dashboard
          </button>

          {/* EXPORT */}

          <button
            className="editor-secondary-btn"
            onClick={exportWorkbook}
          >
            <Download
              size={15}
            />

            Export
          </button>

          {/* SAVE */}

          <button
            className="editor-primary-btn"
            onClick={saveChanges}
          >
            <Save
              size={15}
            />

            Save Changes
          </button>

        </div>

      </header>

      {/* =====================================================
          SHEET TABS
      ===================================================== */}

      <div className="editor-sheet-tabs">

        {data.map(
          (s, index) => (

            <button
              key={`${s.fileIndex}-${s.name}-${index}`}
              className={
                index === activeSheet
                  ? "editor-sheet-tab active"
                  : "editor-sheet-tab"
              }
              onClick={() => {

                setActiveSheet(
                  index
                );

                setSelected({
                  row: 0,
                  col: 0,
                });

                setEditing(false);

              }}
            >
              {s.name}
            </button>

          )
        )}

      </div>

      {/* =====================================================
          TOOLBAR
      ===================================================== */}

      <div className="editor-toolbar">

        <div className="editor-search">

          <Search
            size={15}
          />

          <input
            value={search}
            onChange={(e) =>
              setSearch(
                e.target.value
              )
            }
            placeholder="Search this sheet..."
          />

          {search && (

            <button
              onClick={() =>
                setSearch("")
              }
            >
              <X
                size={13}
              />
            </button>

          )}

        </div>

        <div className="editor-toolbar-spacer" />

        <div className="editor-sheet-info">

          <span>
            {sheet.rows.length.toLocaleString()}
            {" "}rows
          </span>

          <span>
            •
          </span>

          <span>
            {sheet.headers.length}
            {" "}columns
          </span>

        </div>

        <ColumnMenu
          headers={
            sheet.headers
          }
          hiddenColumns={
            hiddenColumns
          }
          activeSheet={
            activeSheet
          }
          onToggle={
            toggleColumn
          }
        />

      </div>

      {/* =====================================================
          FORMULA BAR
      ===================================================== */}

      <div className="editor-formula-bar">

        <div className="editor-name-box">

          {columnLetter(
            selected.col
          )}

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
              : String(
                  selectedValue ?? ""
                )
          }
          onChange={(e) => {

            if (!editing) {
              setEditing(true);
            }

            setEditValue(
              e.target.value
            );

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

              setEditValue(
                String(
                  selectedValue ?? ""
                )
              );

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

              {visibleHeaders.map(
                ({
                  header,
                  index,
                }) => (

                  <th
                    key={`${header}-${index}`}
                    className={
                      selected.col ===
                      index
                        ? "selected-column"
                        : ""
                    }
                  >

                    <div className="column-header-content">

                      <span>
                        {columnLetter(
                          index
                        )}
                      </span>

                      <strong>
                        {header}
                      </strong>

                    </div>

                  </th>

                )
              )}

            </tr>

          </thead>

          <tbody>

            {sheet.rows.map(
              (
                row,
                rowIndex
              ) => {

                const rowText =
                  Object.values(row)
                    .join(" ")
                    .toLowerCase();

                if (
                  search &&
                  !rowText.includes(
                    search.toLowerCase()
                  )
                ) {
                  return null;
                }

                return (

                  <tr
                    key={rowIndex}
                  >

                    {/* ROW NUMBER */}

                    <td
                      className={
                        selected.row ===
                        rowIndex
                          ? "row-number selected-row"
                          : "row-number"
                      }
                    >
                      {rowIndex + 2}
                    </td>

                    {/* CELLS */}

                    {visibleHeaders.map(
                      ({
                        header,
                        index,
                      }) => {

                        const value =
                          row[
                            header
                          ] ?? "";

                        const isSelected =
                          selected.row ===
                            rowIndex &&
                          selected.col ===
                            index;

                        const isFormula =
                          typeof value ===
                            "string" &&
                          value.startsWith(
                            "="
                          );

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
                              selectCell(
                                rowIndex,
                                index
                              )
                            }
                            onDoubleClick={() => {

                              selectCell(
                                rowIndex,
                                index
                              );

                              setEditValue(
                                String(
                                  value
                                )
                              );

                              setEditing(
                                true
                              );

                            }}
                          >

                            {isSelected &&
                            editing ? (

                              <input
                                autoFocus
                                className="cell-editor"
                                value={
                                  editValue
                                }
                                onChange={(e) =>
                                  setEditValue(
                                    e.target.value
                                  )
                                }
                                onBlur={
                                  commitEdit
                                }
                                onKeyDown={(e) => {

                                  if (
                                    e.key ===
                                    "Enter"
                                  ) {
                                    commitEdit();
                                  }

                                  if (
                                    e.key ===
                                    "Escape"
                                  ) {
                                    setEditing(
                                      false
                                    );
                                  }

                                }}
                              />

                            ) : (

                              <span>
                                {String(
                                  value
                                )}
                              </span>

                            )}

                          </td>

                        );

                      }
                    )}

                  </tr>

                );

              }
            )}

          </tbody>

        </table>

      </div>

      {/* =====================================================
          FOOTER
      ===================================================== */}

      <footer className="editor-footer">

        <div>

          Sheet:
          {" "}
          <strong>
            {sheet.name}
          </strong>

        </div>

        <div>

          Selected:
          {" "}
          <strong>
            {selectedHeader ||
              "—"}
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

/* =========================================================
   COLUMN MENU
========================================================= */

function ColumnMenu({
  headers,
  hiddenColumns,
  activeSheet,
  onToggle,
}) {

  const [open, setOpen] =
    useState(false);

  return (

    <div className="column-menu-wrapper">

      <button
        className="editor-secondary-btn"
        onClick={() =>
          setOpen(
            (value) => !value
          )
        }
      >

        <EyeOff
          size={15}
        />

        Columns

        <ChevronDown
          size={14}
        />

      </button>

      {open && (

        <div className="column-menu">

          <div className="column-menu-title">
            Show / Hide Columns
          </div>

          {headers.map(
            (
              header,
              index
            ) => {

              const hidden =
                hiddenColumns[
                  `${activeSheet}-${index}`
                ];

              return (

                <button
                  key={`${header}-${index}`}
                  onClick={() =>
                    onToggle(
                      index
                    )
                  }
                >

                  {hidden ? (

                    <EyeOff
                      size={14}
                    />

                  ) : (

                    <Eye
                      size={14}
                    />

                  )}

                  <span>
                    {columnLetter(
                      index
                    )}
                    {" — "}
                    {header}
                  </span>

                </button>

              );

            }
          )}

        </div>

      )}

    </div>
  );
}