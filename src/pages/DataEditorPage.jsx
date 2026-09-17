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
   NORMALIZE ONE WORKBOOK OBJECT
========================================================= */

function normalizeSingleWorkbook(raw, fileIndex = 0) {
  if (!raw) return [];

  /*
     CASE 1:
     Parsed workbook object:

     {
       fileName,
       workbookSheets: [
         {
           name,
           headers,
           rows
         }
       ],
       sheets: {
         Sheet1: [...]
       }
     }
  */

  if (Array.isArray(raw.workbookSheets)) {
    return raw.workbookSheets.map((sheet) => ({
      name: sheet.name,
      headers:
        sheet.headers ||
        (sheet.rows?.length
          ? Object.keys(sheet.rows[0])
          : []),
      rows: sheet.rows || [],
      fileIndex,
    }));
  }

  /*
     CASE 2:
     Parsed workbook using sheets object
  */

  if (
    raw.sheets &&
    typeof raw.sheets === "object"
  ) {
    return Object.entries(raw.sheets).map(
      ([name, rows]) => ({
        name,
        headers:
          rows?.length
            ? Object.keys(rows[0])
            : [],
        rows: rows || [],
        fileIndex,
      })
    );
  }

  return [];
}

/* =========================================================
   NORMALIZE RAW INPUT

   Supports:

   1. Array of parsed workbook objects
   2. Single parsed workbook object
   3. Array of actual File objects

========================================================= */

async function normalizeRaw(raw) {
  /*
     IMPORTANT:

     Your current UploadPage passes parsed workbook
     objects, NOT browser File objects.

     Therefore this function first handles those objects.
  */

  if (Array.isArray(raw)) {
    const allSheets = [];

    for (
      let fileIndex = 0;
      fileIndex < raw.length;
      fileIndex++
    ) {
      const item = raw[fileIndex];

      /*
         Already parsed workbook
      */
      if (
        item?.workbookSheets ||
        item?.sheets
      ) {
        const sheets =
          normalizeSingleWorkbook(
            item,
            fileIndex
          );

        allSheets.push(...sheets);
        continue;
      }

      /*
         Actual browser File object
      */
      if (
        item &&
        typeof item.arrayBuffer === "function"
      ) {
        const buffer =
          await item.arrayBuffer();

        const workbook = XLSX.read(
          buffer,
          {
            type: "array",
            cellFormula: true,
            cellNF: true,
            cellStyles: true,
            cellDates: true,
          }
        );

        workbook.SheetNames.forEach(
          (sheetName) => {
            const ws =
              workbook.Sheets[sheetName];

            const range =
              XLSX.utils.decode_range(
                ws["!ref"] || "A1"
              );

            const headers = [];

            for (
              let c = range.s.c;
              c <= range.e.c;
              c++
            ) {
              const address =
                XLSX.utils.encode_cell({
                  r: range.s.r,
                  c,
                });

              const cell =
                ws[address];

              headers.push(
                cell?.v ??
                  `Column ${c + 1}`
              );
            }

            const rows = [];

            for (
              let r =
                range.s.r + 1;
              r <= range.e.r;
              r++
            ) {
              const row = {};

              for (
                let c = range.s.c;
                c <= range.e.c;
                c++
              ) {
                const address =
                  XLSX.utils.encode_cell({
                    r,
                    c,
                  });

                const cell =
                  ws[address];

                const header =
                  headers[
                    c - range.s.c
                  ];

                if (cell?.f) {
                  row[header] =
                    `=${cell.f}`;
                } else {
                  row[header] =
                    cell?.v ?? "";
                }
              }

              rows.push(row);
            }

            allSheets.push({
              name: sheetName,
              headers,
              rows,
              fileIndex,
            });
          }
        );
      }
    }

    return allSheets;
  }

  /*
     Single parsed workbook
  */

  if (
    raw?.workbookSheets ||
    raw?.sheets
  ) {
    return normalizeSingleWorkbook(
      raw,
      0
    );
  }

  return [];
}

/* =========================================================
   BUILD EDITED PARSED WORKBOOKS

   This is the critical part.

   We return the SAME structure that your existing
   Dashboard / calculations already expect.

========================================================= */

function buildEditedRaw(
  originalRaw,
  editorData
) {
  /*
     If the application stores multiple uploaded
     workbook objects, preserve that structure.
  */

  if (Array.isArray(originalRaw)) {
    return originalRaw.map(
      (originalWorkbook, fileIndex) => {
        const fileSheets =
          editorData.filter(
            (sheet) =>
              sheet.fileIndex ===
              fileIndex
          );

        const workbookSheets =
          fileSheets.map((sheet) => ({
            name: sheet.name,
            headers: sheet.headers,
            rows: sheet.rows,
          }));

        return {
          ...originalWorkbook,
          workbookSheets,
          sheets:
            Object.fromEntries(
              workbookSheets.map(
                (sheet) => [
                  sheet.name,
                  sheet.rows,
                ]
              )
            ),
        };
      }
    );
  }

  /*
     Single workbook object
  */

  const workbookSheets =
    editorData.map((sheet) => ({
      name: sheet.name,
      headers: sheet.headers,
      rows: sheet.rows,
    }));

  return {
    ...originalRaw,
    workbookSheets,
    sheets:
      Object.fromEntries(
        workbookSheets.map(
          (sheet) => [
            sheet.name,
            sheet.rows,
          ]
        )
      ),
  };
}

/* =========================================================
   EXPORT EDITOR DATA TO XLSX
========================================================= */

function downloadEditorData(
  data,
  fileName
) {
  /*
     Group sheets by original file.
  */

  const grouped = {};

  data.forEach((sheet) => {
    const index =
      sheet.fileIndex ?? 0;

    if (!grouped[index]) {
      grouped[index] = [];
    }

    grouped[index].push(sheet);
  });

  Object.entries(grouped).forEach(
    ([fileIndex, sheets]) => {
      const workbook =
        XLSX.utils.book_new();

      sheets.forEach((sheet) => {
        const rows =
          sheet.rows.map((row) => {
            const ordered = {};

            sheet.headers.forEach(
              (header) => {
                ordered[header] =
                  row[header] ?? "";
              }
            );

            return ordered;
          });

        const worksheet =
          XLSX.utils.json_to_sheet(
            rows,
            {
              header:
                sheet.headers,
            }
          );

        XLSX.utils.book_append_sheet(
          workbook,
          worksheet,
          String(
            sheet.name
          ).slice(0, 31)
        );
      });

      const baseName =
        Array.isArray(fileName)
          ? fileName[
              Number(fileIndex)
            ]
          : fileName;

      XLSX.writeFile(
        workbook,
        baseName
          ? String(baseName)
              .replace(
                /\.xlsx$/i,
                ""
              ) + ".xlsx"
          : `SLA_Edited_Workbook_${
              Number(fileIndex) + 1
            }.xlsx`
      );
    }
  );
}

/* =========================================================
   MAIN COMPONENT
========================================================= */

export default function DataEditorPage({
  raw,
  onBack,
  onSave,
  onRecalculate,
}) {
  const initialSheets = useMemo(() => {
    /*
       Only synchronous normalization here.

       Parsed workbook objects are what your current
       application uses.
    */

    if (Array.isArray(raw)) {
      return raw.flatMap(
        (item, fileIndex) =>
          normalizeSingleWorkbook(
            item,
            fileIndex
          )
      );
    }

    return normalizeSingleWorkbook(
      raw,
      0
    );
  }, [raw]);

  const [data, setData] =
    useState(initialSheets);

  const [activeSheet, setActiveSheet] =
    useState(0);

  const [selected, setSelected] =
    useState({
      row: 0,
      col: 0,
    });

  const [editing, setEditing] =
    useState(false);

  const [editValue, setEditValue] =
    useState("");

  const [search, setSearch] =
    useState("");

  const [hiddenColumns, setHiddenColumns] =
    useState({});

  const [history, setHistory] =
    useState([]);

  const [future, setFuture] =
    useState([]);

  const sheet =
    data[activeSheet];

  /* =======================================================
     VISIBLE HEADERS
  ======================================================= */

  const visibleHeaders =
    useMemo(() => {
      if (!sheet) return [];

      return sheet.headers
        .map(
          (header, index) => ({
            header,
            index,
          })
        )
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
  ======================================================= */

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
  ======================================================= */

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
  ======================================================= */

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
                (row, rowIndex2) =>
                  rowIndex2 ===
                  rowIndex
                    ? {
                        ...row,
                        [header]:
                          value,
                      }
                    : row
              ),
          };
        }
      );

    pushHistory(next);
  }

  /* =======================================================
     COMMIT EDIT
  ======================================================= */

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
  ======================================================= */

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
     UNDO
  ======================================================= */

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
  ======================================================= */

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
  ======================================================= */

  function toggleColumn(
    colIndex
  ) {
    const key =
      `${activeSheet}-${colIndex}`;

    setHiddenColumns(
      (prev) => ({
        ...prev,
        [key]:
          !prev[key],
      })
    );
  }

  /* =======================================================
     EXPORT
  ======================================================= */

  function exportWorkbook() {
    const fileNames =
      Array.isArray(raw)
        ? raw.map(
            (item, index) =>
              item?.fileName ||
              `SLA_Edited_Workbook_${
                index + 1
              }`
          )
        : raw?.fileName ||
          "SLA_Edited_Workbook";

    downloadEditorData(
      data,
      fileNames
    );
  }

  /* =======================================================
     SAVE CHANGES
  ======================================================= */

  function saveChanges() {
    const editedRaw =
      buildEditedRaw(
        raw,
        data
      );

    if (onSave) {
      onSave(
        editedRaw
      );
    }

    alert(
      "Changes saved."
    );
  }

  /* =======================================================
     RECALCULATE DASHBOARD
  ======================================================= */

  function recalculate() {
    const editedRaw =
      buildEditedRaw(
        raw,
        data
      );

    if (onRecalculate) {
      onRecalculate(
        editedRaw
      );
    } else if (onSave) {
      onSave(
        editedRaw
      );
    }

    alert(
      "Dashboard recalculated from the edited data."
    );
  }

  /* =======================================================
     EMPTY STATE
  ======================================================= */

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

          <p>
            The uploaded workbook data
            was not passed to the editor.
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
     MAIN EDITOR
  ======================================================= */

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
              {Array.isArray(raw)
                ? raw
                    .map(
                      (item) =>
                        item?.fileName
                    )
                    .filter(Boolean)
                    .join(" · ") ||
                  "Uploaded Workbooks"
                : raw?.fileName ||
                  "Uploaded Workbook"}
            </div>
          </div>

        </div>

        <div className="editor-actions">

          <button
            className="editor-tool-btn"
            onClick={undo}
            disabled={
              !history.length
            }
            title="Undo"
          >
            <Undo2
              size={16}
            />
          </button>

          <button
            className="editor-tool-btn"
            onClick={redo}
            disabled={
              !future.length
            }
            title="Redo"
          >
            <Redo2
              size={16}
            />
          </button>

          <div className="editor-divider" />

          <button
            className="editor-secondary-btn"
            onClick={
              recalculate
            }
          >
            <RefreshCcw
              size={15}
            />

            Recalculate Dashboard
          </button>

          <button
            className="editor-secondary-btn"
            onClick={
              exportWorkbook
            }
          >
            <Download
              size={15}
            />

            Export
          </button>

          <button
            className="editor-primary-btn"
            onClick={
              saveChanges
            }
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
          (currentSheet, index) => (
            <button
              key={`${currentSheet.fileIndex}-${currentSheet.name}-${index}`}
              className={
                index ===
                activeSheet
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

                setEditing(
                  false
                );

                setSearch("");
              }}
            >
              {currentSheet.name}
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
            {sheet.rows.length.toLocaleString()} rows
          </span>

          <span>
            •
          </span>

          <span>
            {sheet.headers.length} columns
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
                  selectedValue ??
                    ""
                )
          }
          onChange={(e) => {
            if (!editing) {
              setEditing(
                true
              );
            }

            setEditValue(
              e.target.value
            );
          }}
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
          onFocus={() => {
            if (!editing) {
              setEditValue(
                String(
                  selectedValue ??
                    ""
                )
              );

              setEditing(
                true
              );
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
                  Object.values(
                    row
                  )
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
                    key={
                      rowIndex
                    }
                  >

                    <td
                      className={
                        selected.row ===
                        rowIndex
                          ? "row-number selected-row"
                          : "row-number"
                      }
                    >
                      {rowIndex +
                        2}
                    </td>

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
                            ].join(
                              " "
                            )}
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
                                onChange={(
                                  e
                                ) =>
                                  setEditValue(
                                    e.target.value
                                  )
                                }
                                onBlur={
                                  commitEdit
                                }
                                onKeyDown={(
                                  e
                                ) => {

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
          Sheet:{" "}
          <strong>
            {sheet.name}
          </strong>
        </div>

        <div>
          Selected:{" "}
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
            (value) =>
              !value
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
                    )}{" "}
                    — {header}
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