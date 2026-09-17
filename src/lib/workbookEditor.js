import * as XLSX from "xlsx";

/* =========================================================
   READ COMPLETE WORKBOOK

   Converts an uploaded Excel File into an editor-friendly
   workbook object.

   This is useful anywhere outside DataEditorPage where
   the complete workbook structure is required.
========================================================= */

export async function readCompleteWorkbook(file) {
  if (!file) {
    throw new Error("No workbook file provided.");
  }

  const buffer = await file.arrayBuffer();

  const workbook = XLSX.read(buffer, {
    type: "array",
    cellFormula: true,
    cellNF: true,
    cellStyles: true,
    cellDates: true,
  });

  const workbookSheets =
    workbook.SheetNames.map(
      (sheetName) => {

        const ws =
          workbook.Sheets[
            sheetName
          ];

        const range =
          XLSX.utils.decode_range(
            ws["!ref"] || "A1"
          );

        /* -----------------------------------------
           HEADERS
        ----------------------------------------- */

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
            cell?.v !== undefined &&
            cell?.v !== null &&
            cell?.v !== ""
              ? String(cell.v)
              : `Column ${c + 1}`
          );

        }

        /* -----------------------------------------
           ROWS
        ----------------------------------------- */

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

            const cell =
              ws[address];

            const header =
              headers[
                c - range.s.c
              ];

            /*
             * Preserve formulas.
             *
             * Example:
             * Excel formula:
             * =IF(A2>10,"Yes","No")
             *
             * Stored as:
             * "=IF(A2>10,"Yes","No")"
             */

            if (cell?.f) {

              row[header] =
                `=${cell.f}`;

            } else {

              row[header] =
                cell?.v ??
                "";

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
   CONVERT EDITOR STATE → XLSX WORKBOOK

   Input:

   {
     workbookSheets: [
       {
         name,
         headers,
         rows
       }
     ]
   }

========================================================= */

export function workbookStateToXlsx(
  raw
) {

  const workbook =
    XLSX.utils.book_new();

  const sheets =
    raw?.workbookSheets || [];

  if (!sheets.length) {
    throw new Error(
      "No workbook sheets available."
    );
  }

  sheets.forEach(
    (sheet) => {

      const data =
        (sheet.rows || []).map(
          (row) => {

            const ordered = {};

            (sheet.headers || [])
              .forEach(
                (header) => {

                  ordered[header] =
                    row?.[header] ??
                    "";

                }
              );

            return ordered;
          }
        );

      const worksheet =
        XLSX.utils.json_to_sheet(
          data,
          {
            header:
              sheet.headers || [],
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

  return workbook;
}


/* =========================================================
   WORKBOOK STATE → FILE

   This is the important function for the new editor.

   It creates a real browser File object so App.jsx
   can put the edited workbook back into its `files`
   state and run buildDashboardModel() again.
========================================================= */

export function workbookStateToFile(
  raw,
  fileName
) {

  const workbook =
    workbookStateToXlsx(
      raw
    );

  const arrayBuffer =
    XLSX.write(
      workbook,
      {
        bookType: "xlsx",
        type: "array",
      }
    );

  return new File(
    [
      arrayBuffer,
    ],
    fileName ||
      raw?.fileName ||
      "Edited_Workbook.xlsx",
    {
      type:
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    }
  );
}


/* =========================================================
   DOWNLOAD WORKBOOK
========================================================= */

export function downloadWorkbook(
  raw,
  fileName = "Edited_Workbook.xlsx"
) {

  const workbook =
    workbookStateToXlsx(
      raw
    );

  XLSX.writeFile(
    workbook,
    fileName
  );
}


/* =========================================================
   UPDATE ONE CELL

   Works on the editor-style workbook state.

   Does NOT mutate the original object.
========================================================= */

export function updateWorkbookCell(
  raw,
  sheetName,
  rowIndex,
  columnName,
  value
) {

  if (!raw) {
    throw new Error(
      "No workbook data provided."
    );
  }

  const workbookSheets =
    (
      raw.workbookSheets ||
      []
    ).map(
      (sheet) => {

        if (
          sheet.name !==
          sheetName
        ) {
          return sheet;
        }

        return {

          ...sheet,

          rows:
            (
              sheet.rows ||
              []
            ).map(
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
                  [columnName]:
                    value,
                };

              }
            ),

        };

      }
    );

  return {

    ...raw,

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
   CREATE EDITOR STATE FROM A FILE

   Convenience helper.

   Example:

   const workbook =
     await fileToEditorState(file);

========================================================= */

export async function fileToEditorState(
  file
) {

  return await readCompleteWorkbook(
    file
  );
}