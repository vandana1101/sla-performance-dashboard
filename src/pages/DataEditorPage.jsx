import React, { useEffect, useMemo, useState } from "react";
import {
  ArrowLeft,
  Download,
  Eye,
  EyeOff,
  FileSpreadsheet,
  RotateCcw,
  Save,
  Search,
  RefreshCw,
  ChevronLeft,
  ChevronRight,
} from "lucide-react";
import * as XLSX from "xlsx";

function isReadableFile(file) {
  return (
    file &&
    typeof file.arrayBuffer === "function"
  );
}

/* -------------------------------------------------------
   READ ONE EXCEL FILE
------------------------------------------------------- */

async function readExcelFile(file, fileIndex = 0) {
  if (!isReadableFile(file)) {
    throw new Error(
      `Uploaded item ${fileIndex + 1} is not a readable Excel file object.`
    );
  }

  const buffer = await file.arrayBuffer();

  if (!buffer || buffer.byteLength === 0) {
    throw new Error(
      `${file.name || `File ${fileIndex + 1}`} is empty.`
    );
  }

  let workbook;

  try {
    workbook = XLSX.read(buffer, {
      type: "array",
      cellFormula: true,
      cellNF: true,
      cellStyles: true,
      cellDates: true,
    });
  } catch (error) {
    throw new Error(
      `Could not read ${file.name || `File ${fileIndex + 1}`}: ${
        error?.message || "Invalid Excel workbook."
      }`
    );
  }

  if (!workbook?.SheetNames?.length) {
    throw new Error(
      `${file.name || `File ${fileIndex + 1}`} does not contain any worksheets.`
    );
  }

  const sheets = [];

  workbook.SheetNames.forEach((sheetName) => {
    const ws = workbook.Sheets[sheetName];

    if (!ws) return;

    const range = XLSX.utils.decode_range(ws["!ref"] || "A1");

    const headers = [];

    for (let c = range.s.c; c <= range.e.c; c++) {
      const address = XLSX.utils.encode_cell({
        r: range.s.r,
        c,
      });

      const value = ws[address]?.v;

      headers.push(
        value !== undefined &&
          value !== null &&
          String(value).trim() !== ""
          ? String(value)
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
        const header = headers[c - range.s.c];

        if (cell?.f) {
          row[header] = `=${cell.f}`;
        } else if (cell?.v !== undefined && cell?.v !== null) {
          row[header] = cell.v;
        } else {
          row[header] = "";
        }
      }

      rows.push(row);
    }

    sheets.push({
      name: sheetName,
      headers,
      rows,
      fileIndex,
      fileName: file.name || `Workbook ${fileIndex + 1}`,
    });
  });

  return sheets;
}

/* -------------------------------------------------------
   CONVERT EDITED DATA BACK TO ORIGINAL FILE STRUCTURE
------------------------------------------------------- */

function buildEditedFiles(originalFiles, editorSheets) {
  const grouped = {};

  editorSheets.forEach((sheet) => {
    if (!grouped[sheet.fileIndex]) {
      grouped[sheet.fileIndex] = [];
    }

    grouped[sheet.fileIndex].push(sheet);
  });

  return originalFiles.map((originalFile, fileIndex) => {
    const sheets = grouped[fileIndex] || [];

    const workbook = XLSX.utils.book_new();

    sheets.forEach((sheet) => {
      const rows = sheet.rows.map((row) => {
        const orderedRow = {};

        sheet.headers.forEach((header) => {
          orderedRow[header] =
            row[header] === undefined ||
            row[header] === null
              ? ""
              : row[header];
        });

        return orderedRow;
      });

      const worksheet = XLSX.utils.json_to_sheet(
        rows,
        {
          header: sheet.headers,
        }
      );

      XLSX.utils.book_append_sheet(
        workbook,
        worksheet,
        String(sheet.name).slice(0, 31)
      );
    });

    const output = XLSX.write(workbook, {
      bookType: "xlsx",
      type: "array",
    });

    return new File(
      [output],
      originalFile.name || `Edited_${fileIndex + 1}.xlsx`,
      {
        type:
          "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
      }
    );
  });
}

/* -------------------------------------------------------
   MAIN COMPONENT
------------------------------------------------------- */

export default function DataEditorPage({
  raw,
  onBack,
  onSave,
  onRecalculate,
}) {
  const [data, setData] = useState([]);
  const [activeSheet, setActiveSheet] = useState(0);

  const [loading, setLoading] = useState(true);
  const [error, setError] = useState("");

  const [search, setSearch] = useState("");
  const [hiddenColumns, setHiddenColumns] = useState([]);

  const [history, setHistory] = useState([]);
  const [historyIndex, setHistoryIndex] = useState(-1);

  const [saving, setSaving] = useState(false);

  /* -------------------------------------------------------
     NORMALIZE INPUT
  ------------------------------------------------------- */

  const inputFiles = useMemo(() => {
    if (Array.isArray(raw)) {
      return raw;
    }

    if (raw?.files && Array.isArray(raw.files)) {
      return raw.files;
    }

    if (raw?.file) {
      return [raw.file];
    }

    return [];
  }, [raw]);

  /* -------------------------------------------------------
     LOAD WORKBOOKS
  ------------------------------------------------------- */

  useEffect(() => {
    let cancelled = false;

    async function loadFiles() {
      setLoading(true);
      setError("");
      setData([]);
      setActiveSheet(0);

      try {
        if (!inputFiles.length) {
          throw new Error(
            "No uploaded files were passed to the data editor."
          );
        }

        const allSheets = [];

        for (
          let fileIndex = 0;
          fileIndex < inputFiles.length;
          fileIndex++
        ) {
          const file = inputFiles[fileIndex];

          const sheets = await readExcelFile(
            file,
            fileIndex
          );

          allSheets.push(...sheets);
        }

        if (!allSheets.length) {
          throw new Error(
            "The uploaded workbooks contain no readable worksheets."
          );
        }

        if (!cancelled) {
          setData(allSheets);

          setHistory([
            JSON.parse(JSON.stringify(allSheets)),
          ]);

          setHistoryIndex(0);
        }
      } catch (err) {
        console.error(
          "DATA EDITOR WORKBOOK LOAD ERROR:",
          err
        );

        if (!cancelled) {
          setError(
            err?.message ||
              "The uploaded files could not be read as Excel workbooks."
          );
        }
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
  }, [inputFiles]);

  /* -------------------------------------------------------
     ACTIVE SHEET
  ------------------------------------------------------- */

  const sheet = data[activeSheet];

  /* -------------------------------------------------------
     FILTER ROWS
  ------------------------------------------------------- */

  const filteredRows = useMemo(() => {
    if (!sheet) return [];

    if (!search.trim()) {
      return sheet.rows;
    }

    const term = search.toLowerCase();

    return sheet.rows.filter((row) =>
      sheet.headers.some((header) =>
        String(row[header] ?? "")
          .toLowerCase()
          .includes(term)
      )
    );
  }, [sheet, search]);

  /* -------------------------------------------------------
     UPDATE CELL
  ------------------------------------------------------- */

  const updateCell = (
    rowIndex,
    column,
    value
  ) => {
    if (!sheet) return;

    const actualRowIndex = sheet.rows.indexOf(
      filteredRows[rowIndex]
    );

    if (actualRowIndex === -1) return;

    const newData = data.map((currentSheet, index) => {
      if (index !== activeSheet) {
        return currentSheet;
      }

      return {
        ...currentSheet,
        rows: currentSheet.rows.map(
          (row, index2) => {
            if (index2 !== actualRowIndex) {
              return row;
            }

            return {
              ...row,
              [column]: value,
            };
          }
        ),
      };
    });

    setData(newData);

    const snapshot = JSON.parse(
      JSON.stringify(newData)
    );

    const newHistory = history.slice(
      0,
      historyIndex + 1
    );

    newHistory.push(snapshot);

    setHistory(newHistory);
    setHistoryIndex(newHistory.length - 1);
  };

  /* -------------------------------------------------------
     UNDO
  ------------------------------------------------------- */

  const undo = () => {
    if (historyIndex <= 0) return;

    const previousIndex =
      historyIndex - 1;

    const previousData =
      JSON.parse(
        JSON.stringify(
          history[previousIndex]
        )
      );

    setData(previousData);
    setHistoryIndex(previousIndex);
  };

  /* -------------------------------------------------------
     REDO
  ------------------------------------------------------- */

  const redo = () => {
    if (
      historyIndex >=
      history.length - 1
    ) {
      return;
    }

    const nextIndex =
      historyIndex + 1;

    const nextData =
      JSON.parse(
        JSON.stringify(
          history[nextIndex]
        )
      );

    setData(nextData);
    setHistoryIndex(nextIndex);
  };

  /* -------------------------------------------------------
     HIDE / SHOW COLUMN
  ------------------------------------------------------- */

  const toggleColumn = (column) => {
    setHiddenColumns((current) => {
      if (current.includes(column)) {
        return current.filter(
          (item) => item !== column
        );
      }

      return [...current, column];
    });
  };

  /* -------------------------------------------------------
     SAVE
  ------------------------------------------------------- */

  const saveChanges = async () => {
    if (!inputFiles.length || !data.length) {
      return;
    }

    setSaving(true);

    try {
      const editedFiles =
        buildEditedFiles(
          inputFiles,
          data
        );

      if (onSave) {
        await onSave(editedFiles);
      }
    } catch (err) {
      console.error(
        "SAVE ERROR:",
        err
      );

      setError(
        err?.message ||
          "Unable to save the edited workbook."
      );
    } finally {
      setSaving(false);
    }
  };

  /* -------------------------------------------------------
     SAVE + RECALCULATE
  ------------------------------------------------------- */

  const recalculate = async () => {
    if (!inputFiles.length || !data.length) {
      return;
    }

    setSaving(true);
    setError("");

    try {
      const editedFiles =
        buildEditedFiles(
          inputFiles,
          data
        );

      if (onRecalculate) {
        await onRecalculate(
          editedFiles
        );
      } else if (onSave) {
        await onSave(editedFiles);
      }
    } catch (err) {
      console.error(
        "RECALCULATE ERROR:",
        err
      );

      setError(
        err?.message ||
          "Unable to recalculate the dashboard."
      );
    } finally {
      setSaving(false);
    }
  };

  /* -------------------------------------------------------
     EXPORT
  ------------------------------------------------------- */

  const exportWorkbook = () => {
    if (!data.length) return;

    const files =
      buildEditedFiles(
        inputFiles,
        data
      );

    files.forEach((file) => {
      const url =
        URL.createObjectURL(file);

      const link =
        document.createElement("a");

      link.href = url;
      link.download =
        file.name ||
        "Edited_Workbook.xlsx";

      document.body.appendChild(link);
      link.click();
      document.body.removeChild(link);

      URL.revokeObjectURL(url);
    });
  };

  /* -------------------------------------------------------
     LOADING
  ------------------------------------------------------- */

  if (loading) {
    return (
      <div className="data-editor-page">
        <div className="editor-loading">
          <RefreshCw
            size={42}
            className="spin"
          />

          <h2>
            Loading workbook...
          </h2>

          <p>
            Reading the uploaded Excel
            files.
          </p>
        </div>
      </div>
    );
  }

  /* -------------------------------------------------------
     ERROR
  ------------------------------------------------------- */

  if (error) {
    return (
      <div className="data-editor-page">
        <div className="editor-error">
          <FileSpreadsheet
            size={56}
          />

          <h1>
            Unable to load workbook
          </h1>

          <p>
            {error}
          </p>

          <button
            type="button"
            className="editor-primary-btn"
            onClick={onBack}
          >
            <ArrowLeft size={16} />
            Back to Dashboard
          </button>
        </div>
      </div>
    );
  }

  /* -------------------------------------------------------
     EMPTY
  ------------------------------------------------------- */

  if (!data.length || !sheet) {
    return (
      <div className="data-editor-page">
        <div className="editor-error">
          <FileSpreadsheet
            size={56}
          />

          <h1>
            No workbook data available
          </h1>

          <p>
            No readable worksheets were
            found in the uploaded files.
          </p>

          <button
            type="button"
            className="editor-primary-btn"
            onClick={onBack}
          >
            <ArrowLeft size={16} />
            Back to Dashboard
          </button>
        </div>
      </div>
    );
  }

  /* -------------------------------------------------------
     MAIN EDITOR
  ------------------------------------------------------- */

  return (
    <div className="data-editor-page">

      {/* HEADER */}

      <div className="data-editor-header">

        <div className="editor-title-area">

          <button
            type="button"
            className="editor-back-btn"
            onClick={onBack}
          >
            <ArrowLeft size={17} />
            Dashboard
          </button>

          <div className="editor-title">
            <FileSpreadsheet
              size={25}
            />

            <div>
              <h1>
                Data Editor
              </h1>

              <p>
                Edit uploaded workbook
                data before recalculating
                the dashboard.
              </p>
            </div>
          </div>

        </div>

        <div className="editor-actions">

          <button
            type="button"
            className="editor-secondary-btn"
            onClick={undo}
            disabled={
              historyIndex <= 0
            }
          >
            <RotateCcw size={15} />
            Undo
          </button>

          <button
            type="button"
            className="editor-secondary-btn"
            onClick={redo}
            disabled={
              historyIndex >=
              history.length - 1
            }
          >
            <RefreshCw size={15} />
            Redo
          </button>

          <button
            type="button"
            className="editor-secondary-btn"
            onClick={exportWorkbook}
          >
            <Download size={15} />
            Export
          </button>

          <button
            type="button"
            className="editor-secondary-btn"
            onClick={saveChanges}
            disabled={saving}
          >
            <Save size={15} />
            {saving
              ? "Saving..."
              : "Save Changes"}
          </button>

          <button
            type="button"
            className="editor-primary-btn"
            onClick={recalculate}
            disabled={saving}
          >
            <RefreshCw size={15} />
            {saving
              ? "Recalculating..."
              : "Save & Recalculate"}
          </button>

        </div>
      </div>

      {/* TOOLBAR */}

      <div className="data-editor-toolbar">

        <div className="sheet-tabs">

          {data.map(
            (currentSheet, index) => (
              <button
                key={`${currentSheet.fileIndex}-${currentSheet.name}-${index}`}
                type="button"
                className={
                  index === activeSheet
                    ? "sheet-tab active"
                    : "sheet-tab"
                }
                onClick={() => {
                  setActiveSheet(index);
                  setSearch("");
                  setHiddenColumns([]);
                }}
              >
                <FileSpreadsheet
                  size={14}
                />

                {currentSheet.name}

                <span>
                  {currentSheet.rows.length}
                </span>
              </button>
            )
          )}

        </div>

        <div className="editor-search">

          <Search size={16} />

          <input
            type="text"
            placeholder="Search this sheet..."
            value={search}
            onChange={(e) =>
              setSearch(
                e.target.value
              )
            }
          />

          {search && (
            <span>
              {filteredRows.length} rows
            </span>
          )}

        </div>

      </div>

      {/* COLUMN CONTROLS */}

      <div className="column-controls">

        <span className="column-control-label">
          Columns:
        </span>

        {sheet.headers.map(
          (header) => {
            const hidden =
              hiddenColumns.includes(
                header
              );

            return (
              <button
                key={header}
                type="button"
                className={
                  hidden
                    ? "column-toggle hidden"
                    : "column-toggle"
                }
                onClick={() =>
                  toggleColumn(
                    header
                  )
                }
                title={
                  hidden
                    ? `Show ${header}`
                    : `Hide ${header}`
                }
              >
                {hidden ? (
                  <EyeOff
                    size={13}
                  />
                ) : (
                  <Eye
                    size={13}
                  />
                )}

                {header}
              </button>
            );
          }
        )}

      </div>

      {/* TABLE */}

      <div className="data-editor-table-wrapper">

        <table className="data-editor-table">

          <thead>
            <tr>

              <th className="row-number-header">
                #
              </th>

              {sheet.headers
                .filter(
                  (header) =>
                    !hiddenColumns.includes(
                      header
                    )
                )
                .map(
                  (header) => (
                    <th
                      key={header}
                      title={header}
                    >
                      {header}
                    </th>
                  )
                )}

            </tr>
          </thead>

          <tbody>

            {filteredRows.map(
              (row, visibleRowIndex) => {

                const actualRowIndex =
                  sheet.rows.indexOf(
                    row
                  );

                return (
                  <tr
                    key={
                      `${actualRowIndex}-${visibleRowIndex}`
                    }
                  >

                    <td className="row-number">
                      {actualRowIndex +
                        2}
                    </td>

                    {sheet.headers
                      .filter(
                        (header) =>
                          !hiddenColumns.includes(
                            header
                          )
                      )
                      .map(
                        (header) => (
                          <td
                            key={header}
                          >
                            <input
                              className="cell-input"
                              value={
                                row[
                                  header
                                ] ?? ""
                              }
                              onChange={(
                                e
                              ) =>
                                updateCell(
                                  visibleRowIndex,
                                  header,
                                  e.target.value
                                )
                              }
                            />
                          </td>
                        )
                      )}

                  </tr>
                );
              }
            )}

          </tbody>

        </table>

        {!filteredRows.length && (
          <div className="editor-no-results">
            No matching rows found.
          </div>
        )}

      </div>

      {/* FOOTER */}

      <div className="data-editor-footer">

        <div>
          <strong>
            {sheet.name}
          </strong>

          <span>
            {" "}
            · {sheet.rows.length} rows
            {" "}
            · {sheet.headers.length} columns
          </span>
        </div>

        <div>
          Showing{" "}
          <strong>
            {filteredRows.length}
          </strong>{" "}
          of{" "}
          <strong>
            {sheet.rows.length}
          </strong>{" "}
          rows
        </div>

      </div>

    </div>
  );
}