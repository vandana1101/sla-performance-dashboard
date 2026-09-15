import * as XLSX from "xlsx";

export async function readInputFile(file) {
  const ext = file.name.split(".").pop().toLowerCase();
  if (ext === "csv") {
    const text = await file.text();
    const wb = XLSX.read(text, { type: "string", cellDates: true });
    return {
      name: file.name,
      sheets: wb.SheetNames.map(name => ({
        name: name || file.name.replace(/\.[^.]+$/, ""),
        rows: XLSX.utils.sheet_to_json(wb.Sheets[name], { defval: null, raw: false })
      }))
    };
  }
  const buffer = await file.arrayBuffer();
  const wb = XLSX.read(buffer, { type: "array", cellDates: true, raw: false, dateNF: "dd-mm-yyyy" });
  return {
    name: file.name,
    sheets: wb.SheetNames.map(name => ({
      name,
      rows: XLSX.utils.sheet_to_json(wb.Sheets[name], { defval: null, raw: false })
    }))
  };
}