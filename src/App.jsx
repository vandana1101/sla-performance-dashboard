import React, { useMemo, useState } from "react";
import { LayoutDashboard, UploadCloud, ChevronRight, RotateCcw } from "lucide-react";
import UploadPage from "./pages/UploadPage";
import DashboardPage from "./pages/DashboardPage";
import { buildDashboardModel } from "./lib/calculations";

export default function App() {
  const [page, setPage] = useState("upload");
  const [files, setFiles] = useState([]);
  const [format, setFormat] = useState("format1");

  const model = useMemo(() => {
    if (!files.length) return null;
    try {
      return buildDashboardModel(files);
    } catch (e) {
      return { error: e?.message || "Unable to calculate dashboard." };
    }
  }, [files]);

  const build = () => {
    if (!files.length) return;
    setPage("dashboard");
  };

  const reset = () => {
    setFiles([]);
    setPage("upload");
  };

  return (
    <div className="app-shell">
      <header className="topbar">
        <div className="brand">
          <div className="brand-mark"><LayoutDashboard size={20} /></div>
          <div>
            <div className="brand-name">SLA Performance</div>
            <div className="brand-sub">Customer Analytics Portal</div>
          </div>
        </div>
        <div className="topbar-right">
          <span className="format-pill">FORMAT 1</span>
          {page === "dashboard" && (
            <button className="ghost-button" onClick={reset}>
              <RotateCcw size={16} /> New upload
            </button>
          )}
        </div>
      </header>

      <div className="page-wrap">
        {page === "upload" ? (
          <UploadPage
            files={files}
            setFiles={setFiles}
            format={format}
            setFormat={setFormat}
            onBuild={build}
          />
        ) : (
          <DashboardPage model={model} />
        )}
      </div>
    </div>
  );
}