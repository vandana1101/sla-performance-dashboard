import React, { useRef, useState } from "react";
import { FileSpreadsheet, UploadCloud, X, ArrowRight, CheckCircle2, FileText } from "lucide-react";
import { readInputFile } from "../lib/fileReader";

export default function UploadPage({ files, setFiles, format, setFormat, onBuild }) {
  const inputRef = useRef(null);
  const [dragging, setDragging] = useState(false);
  const [error, setError] = useState("");

  const addFiles = async (incoming) => {
    const selected = Array.from(incoming || []);
    if (!selected.length) return;
    setError("");
    try {
      const parsed = [];
      for (const file of selected) {
        const ext = file.name.split(".").pop().toLowerCase();
        if (!["xlsx", "xls", "csv"].includes(ext)) {
          throw new Error(`${file.name}: upload Excel or CSV only.`);
        }
        parsed.push(await readInputFile(file));
      }
      setFiles((prev) => [...prev, ...parsed]);
    } catch (e) {
      setError(e.message);
    }
  };

  const remove = (idx) => setFiles(files.filter((_, i) => i !== idx));

  return (
    <main className="upload-page">
      <section className="hero-copy">
        <div className="eyebrow">DELIVERY INTELLIGENCE</div>
        <h1>Build your <span>SLA dashboard.</span></h1>
        <p>
          Upload your operational files. The dashboard calculates TAT, SLA,
          state and LR performance automatically — without requiring pre-calculated columns.
        </p>
      </section>

      <section className="upload-card">
        <div className="upload-card-head">
          <div>
            <h2>1. Upload source data</h2>
            <p>Excel workbook or CSV files. You can upload multiple files.</p>
          </div>
          <div className="supported"><FileSpreadsheet size={17}/> XLSX · XLS · CSV</div>
        </div>

        <div
          className={`dropzone ${dragging ? "dragging" : ""}`}
          onDragOver={(e) => { e.preventDefault(); setDragging(true); }}
          onDragLeave={() => setDragging(false)}
          onDrop={(e) => { e.preventDefault(); setDragging(false); addFiles(e.dataTransfer.files); }}
          onClick={() => inputRef.current?.click()}
        >
          <input ref={inputRef} hidden type="file" multiple accept=".xlsx,.xls,.csv" onChange={(e) => addFiles(e.target.files)} />
          <div className="upload-icon"><UploadCloud size={28}/></div>
          <h3>Drop your files here</h3>
          <p>or click to browse from your computer</p>
          <span className="drop-hint">Raw data is enough. Calculated columns are created at runtime.</span>
        </div>

        {error && <div className="error-box">{error}</div>}

        {files.length > 0 && (
          <div className="file-list">
            {files.map((f, i) => (
              <div className="file-row" key={`${f.name}-${i}`}>
                <div className="file-icon"><FileText size={18}/></div>
                <div className="file-info">
                  <strong>{f.name}</strong>
                  <span>{f.sheets?.length || 1} source sheet{(f.sheets?.length || 1) > 1 ? "s" : ""}</span>
                </div>
                <CheckCircle2 className="ok" size={19}/>
                <button className="remove" onClick={(e) => { e.stopPropagation(); remove(i); }}><X size={17}/></button>
              </div>
            ))}
          </div>
        )}

        <div className="format-section">
          <div>
            <h2>2. Choose dashboard format</h2>
            <p>This lets the same data engine support multiple customer-facing layouts.</p>
          </div>
          <select value={format} onChange={(e) => setFormat(e.target.value)} className="format-select">
            <option value="format1">Format 1 — SLA Delivery Performance</option>
          </select>
        </div>

        <button className="build-button" disabled={!files.length} onClick={onBuild}>
          Build my dashboard <ArrowRight size={19}/>
        </button>
      </section>

      <div className="upload-note">
        <span>Privacy by design</span>
        <p>Your calculations happen in the browser. No calculated workbook is stored by this application.</p>
      </div>
    </main>
  );
}