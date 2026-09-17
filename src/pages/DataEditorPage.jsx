import React, {
  useEffect,
  useMemo,
  useState,
} from "react";

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

    result =
      String.fromCharCode(65 + rem) +
      result;

    n = Math.floor((n - 1) / 26);
  }

  return result;
}


/* =========================================================
   READ AN ACTUAL EXCEL FILE
========================================================= */

async function readExcelFile(file, fileIndex) {
  if (
    !file ||
    typeof file.arrayBuffer !== "function"
  ) {
    throw new Error(
      `Uploaded item ${
        fileIndex + 1
      } is not a readable Excel file object.`
    );
  }

  const buffer =
    await file.arrayBuffer();

  const workbook =
    XLSX.read(buffer, {
      type: "array",
      cellFormula: true,
      cellNF: true,
      cellStyles: true,
      cellDates: true,
    });

  const sheets = [];

  workbook.SheetNames.forEach(
    (sheetName) => {
      const ws =
        workbook.Sheets[
          sheetName
        ];

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

      sheets.push({
        name: sheetName,
        headers,
        rows,
        fileIndex,
      });
    }
  );

  return sheets;
}


/* =========================================================
   READ PARSED WORKBOOK OBJECT
========================================================= */

function readParsedWorkbook(
  workbook,
  fileIndex
) {
  if (!workbook) {
    return [];
  }

  /*
     Structure:

     {
       fileName,
       workbookSheets: [
         {
           name,
           headers,
           rows
         }
       ]
     }
  */

  if (
    Array.isArray(
      workbook.workbookSheets
    )
  ) {
    return workbook.workbookSheets.map(
      (sheet) => ({
        name: sheet.name,

        headers:
          sheet.headers ||
          (
            sheet.rows?.length
              ? Object.keys(
                  sheet.rows[0]
                )
              : []
          ),

        rows:
          sheet.rows || [],

        fileIndex,
      })
    );
  }

  /*
     Structure:

     {
       sheets: {
         Sheet1: [],
         Sheet2: []
       }
     }
  */

  if (
    workbook.sheets &&
    typeof workbook.sheets ===
      "object"
  ) {
    return Object.entries(
      workbook.sheets
    ).map(
      ([name, rows]) => ({
        name,

        headers:
          rows?.length
            ? Object.keys(
                rows[0]
              )
            : [],

        rows:
          rows || [],

        fileIndex,
      })
    );
  }

  return [];
}


/* =========================================================
   LOAD ANY RAW INPUT

   Supports BOTH:

   1. File[]
   2. Parsed workbook objects[]

========================================================= */

async function loadRawFiles(raw) {
  if (!raw) {
    return [];
  }

  const items =
    Array.isArray(raw)
      ? raw
      : [raw];

  const allSheets = [];

  for (
    let fileIndex = 0;
    fileIndex < items.length;
    fileIndex++
  ) {
    const item =
      items[fileIndex];

    /*
       ACTUAL EXCEL FILE
    */

    if (
      item &&
      typeof item.arrayBuffer ===
        "function"
    ) {
      const sheets =
        await readExcelFile(
          item,
          fileIndex
        );

      allSheets.push(
        ...sheets
      );

      continue;
    }

    /*
       ALREADY PARSED WORKBOOK
    */

    const sheets =
      readParsedWorkbook(
        item,
        fileIndex
      );

    allSheets.push(
      ...sheets
    );
  }

  return allSheets;
}


/* =========================================================
   CONVERT EDITOR DATA BACK TO PARSED WORKBOOKS
========================================================= */

function buildEditedRaw(
  raw,
  data
) {
  const original =
    Array.isArray(raw)
      ? raw
      : [raw];

  return original.map(
    (
      originalWorkbook,
      fileIndex
    ) => {
      const fileSheets =
        data.filter(
          (sheet) =>
            sheet.fileIndex ===
            fileIndex
        );

      const workbookSheets =
        fileSheets.map(
          (sheet) => ({
            name:
              sheet.name,

            headers:
              sheet.headers,

            rows:
              sheet.rows,
          })
        );

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


/* =========================================================
   EXPORT
========================================================= */

function exportData(
  data,
  raw
) {
  const groups = {};

  data.forEach(
    (sheet) => {
      const index =
        sheet.fileIndex || 0;

      if (!groups[index]) {
        groups[index] = [];
      }

      groups[index].push(
        sheet
      );
    }
  );

  Object.entries(groups).forEach(
    ([fileIndex, sheets]) => {
      const workbook =
        XLSX.utils.book_new();

      sheets.forEach(
        (sheet) => {
          const rows =
            sheet.rows.map(
              (row) => {
                const ordered =
                  {};

                sheet.headers.forEach(
                  (header) => {
                    ordered[
                      header
                    ] =
                      row[
                        header
                      ] ?? "";
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
        }
      );

      let fileName =
        `SLA_Edited_Workbook_${
          Number(fileIndex) + 1
        }.xlsx`;

      if (
        Array.isArray(raw) &&
        raw[fileIndex]?.fileName
      ) {
        fileName =
          String(
            raw[fileIndex]
              .fileName
          ).replace(
            /\.xlsx$/i,
            "_Edited.xlsx"
          );
      }

      XLSX.writeFile(
        workbook,
        fileName
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
  const [data, setData] =
    useState([]);

  const [loading, setLoading] =
    useState(true);

  const [error, setError] =
    useState("");

  const [
    activeSheet,
    setActiveSheet,
  ] = useState(0);

  const [selected, setSelected] =
    useState({
      row: 0,
      col: 0,
    });

  const [editing, setEditing] =
    useState(false);

  const [
    editValue,
    setEditValue,
  ] = useState("");

  const [search, setSearch] =
    useState("");

  const [
    hiddenColumns,
    setHiddenColumns,
  ] = useState({});

  const [history, setHistory] =
    useState([]);

  const [future, setFuture] =
    useState([]);

  /* =======================================================
     LOAD DATA
  ======================================================= */

  useEffect(() => {
    let cancelled = false;

    async function load() {
      try {
        setLoading(true);
        setError("");

        const sheets =
          await loadRawFiles(
            raw
          );

        if (cancelled) {
          return;
        }

        setData(sheets);
        setActiveSheet(0);

        setSelected({
          row: 0,
          col: 0,
        });

        setHistory([]);
        setFuture([]);
      } catch (err) {
        if (!cancelled) {
          console.error(
            "DataEditor load error:",
            err
          );

          setError(
            err?.message ||
              "Unable to load workbook."
          );
        }
      } finally {
        if (!cancelled) {
          setLoading(false);
        }
      }
    }

    load();

    return () => {
      cancelled = true;
    };
  }, [raw]);


  /* =======================================================
     CURRENT SHEET
  ======================================================= */

  const sheet =
    data[activeSheet];


  /* =======================================================
     VISIBLE COLUMNS
  ======================================================= */

  const visibleHeaders =
    useMemo(() => {
      if (!sheet) {
        return [];
      }

      return sheet.headers
        .map(
          (
            header,
            index
          ) => ({
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
    ]?.[selectedHeader] ??
    "";


  /* =======================================================
     UPDATE CELL
  ======================================================= */

  function updateCell(
    rowIndex,
    colIndex,
    value
  ) {
    if (!sheet) {
      return;
    }

    const header =
      sheet.headers[
        colIndex
      ];

    const previous =
      data;

    const next =
      data.map(
        (
          current,
          sheetIndex
        ) => {
          if (
            sheetIndex !==
            activeSheet
          ) {
            return current;
          }

          return {
            ...current,

            rows:
              current.rows.map(
                (
                  row,
                  index
                ) => {
                  if (
                    index !==
                    rowIndex
                  ) {
                    return row;
                  }

                  return {
                    ...row,
                    [header]:
                      value,
                  };
                }
              ),
          };
        }
      );

    setHistory(
      (prev) => [
        ...prev.slice(-30),
        previous,
      ]
    );

    setFuture([]);

    setData(next);
  }


  /* =======================================================
     COMMIT
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
     SELECT
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
    if (!history.length) {
      return;
    }

    const previous =
      history[
        history.length - 1
      ];

    setFuture(
      (prev) => [
        ...prev,
        data,
      ]
    );

    setData(previous);

    setHistory(
      (prev) =>
        prev.slice(
          0,
          -1
        )
    );
  }


  /* =======================================================
     REDO
  ======================================================= */

  function redo() {
    if (!future.length) {
      return;
    }

    const next =
      future[
        future.length - 1
      ];

    setHistory(
      (prev) => [
        ...prev,
        data,
      ]
    );

    setData(next);

    setFuture(
      (prev) =>
        prev.slice(
          0,
          -1
        )
    );
  }


  /* =======================================================
     HIDE / SHOW COLUMN
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
     SAVE
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
     RECALCULATE
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
     EXPORT
  ======================================================= */

  function exportWorkbook() {
    exportData(
      data,
      raw
    );
  }


  /* =======================================================
     LOADING SCREEN
  ======================================================= */

  if (loading) {
    return (
      <div className="data-editor-page">

        <div className="data-editor-empty">

          <RefreshCcw
            size={44}
            className="spin"
          />

          <h2>
            Loading workbook...
          </h2>

          <p>
            Reading the uploaded
            workbook data.
          </p>

        </div>

      </div>
    );
  }


  /* =======================================================
     ERROR SCREEN
  ======================================================= */

  if (error) {
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
            {error}
          </p>

          <button
            onClick={onBack}
          >
            <ArrowLeft
              size={16}
            />

            Back to Dashboard
          </button>

        </div>

      </div>
    );
  }


  /* =======================================================
     NO DATA
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
            No readable sheets were
            found in the uploaded
            workbook.
          </p>

          <button
            onClick={onBack}
          >
            <ArrowLeft
              size={16}
            />

            Back to Dashboard
          </button>

        </div>

      </div>
    );
  }


  /* =======================================================
     EDITOR
  ======================================================= */

  return (
    <div className="data-editor-page">

      {/* =====================================================
          HEADER
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
                  "Uploaded Workbook"
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
          (
            currentSheet,
            index
          ) => (
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