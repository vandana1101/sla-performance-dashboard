import * as XLSX from "xlsx";

/* =========================================================
   READ COMPLETE WORKBOOK
========================================================= */

export async function readCompleteWorkbook(file) {
  const buffer = await file.arrayBuffer();

  const workbook = XLSX.read(buffer, {
    type: "array",
    cellFormula: true,
    cellNF: true,
    cellStyles: true,
    cellDates: true,
  });

  const workbookSheets = workbook.SheetNames.map(
    (sheetName) => {
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

        headers.push(
          ws[address]?.v ??
            `Column ${c + 1}`
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
          const address =
            XLSX.utils.encode_cell({
              r,
              c,
            });

          const cell = ws[address];

          const header =
            headers[c - range.s.c];

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

      return {
        name: sheetName,
        headers,
        rows,
      };
    }
  );

  return {
    fileName: file.name,
    workbookSheets,
    sheets: Object.fromEntries(
      workbookSheets.map((s) => [
        s.name,
        s.rows,
      ])
    ),
  };
}


/* =========================================================
   CONVERT EDITOR STATE BACK TO XLSX
========================================================= */

export function workbookStateToXlsx(raw) {
  const workbook =
    XLSX.utils.book_new();

  const sheets =
    raw?.workbookSheets || [];

  sheets.forEach((sheet) => {

    const data = sheet.rows.map(
      (row) => {
        const ordered = {};

        sheet.headers.forEach(
          (header) => {
            ordered[header] =
              row[header] ?? "";
          }
        );

        return ordered;
      }
    );

    const worksheet =
      XLSX.utils.json_to_sheet(
        data,
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

  return workbook;
}


/* =========================================================
   DOWNLOAD
========================================================= */

export function downloadWorkbook(
  raw,
  fileName = "Edited_Workbook.xlsx"
) {
  const workbook =
    workbookStateToXlsx(raw);

  XLSX.writeFile(
    workbook,
    fileName
  );
}


/* =========================================================
   UPDATE ONE CELL
========================================================= */

export function updateWorkbookCell(
  raw,
  sheetName,
  rowIndex,
  columnName,
  value
) {
  const workbookSheets =
    (raw.workbookSheets || []).map(
      (sheet) => {

        if (sheet.name !== sheetName) {
          return sheet;
        }

        return {
          ...sheet,
          rows: sheet.rows.map(
            (row, index) =>
              index === rowIndex
                ? {
                    ...row,
                    [columnName]: value,
                  }
                : row
          ),
        };
      }
    );

  return {
    ...raw,
    workbookSheets,
    sheets: Object.fromEntries(
      workbookSheets.map((s) => [
        s.name,
        s.rows,
      ])
    ),
  };
}