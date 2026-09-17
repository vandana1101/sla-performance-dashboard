import React, { useMemo, useState } from "react";
import {
  LayoutDashboard,
  RotateCcw,
} from "lucide-react";

import UploadPage from "./pages/UploadPage";
import DashboardPage from "./pages/DashboardPage";
import DataEditorPage from "./pages/DataEditorPage";

import { buildDashboardModel } from "./lib/calculations";

export default function App() {
  const [page, setPage] = useState("upload");

  const [files, setFiles] = useState([]);

  const [format, setFormat] =
    useState("format1");


  /* =========================================================
     DASHBOARD MODEL
  ========================================================= */

  const model = useMemo(() => {
    if (!files.length) return null;

    try {
      return buildDashboardModel(files);
    } catch (e) {
      console.error(
        "Dashboard calculation error:",
        e
      );

      return {
        error:
          e?.message ||
          "Unable to calculate dashboard.",
      };
    }
  }, [files]);


  /* =========================================================
     BUILD DASHBOARD
  ========================================================= */

  const build = () => {
    if (!files.length) return;

    setPage("dashboard");
  };


  /* =========================================================
     OPEN DATA EDITOR
  ========================================================= */

  const openEditor = () => {
    if (!files.length) return;

    setPage("editor");
  };


  /* =========================================================
     SAVE EDITED DATA
  ========================================================= */

  const handleEditorSave = (
    editedFiles
  ) => {
    if (!editedFiles) return;

    setFiles(editedFiles);
  };


  /* =========================================================
     RECALCULATE DASHBOARD
  ========================================================= */

  const handleEditorRecalculate = (
    editedFiles
  ) => {
    if (!editedFiles) return;

    /*
      Updating files forces buildDashboardModel()
      to run again because `files` is a dependency
      of the useMemo above.
    */

    setFiles(editedFiles);

    setPage("dashboard");
  };


  /* =========================================================
     RESET / NEW UPLOAD
  ========================================================= */

  const reset = () => {
    setFiles([]);

    setPage("upload");
  };


  /* =========================================================
     HEADER
  ========================================================= */

  return (
    <div className="app-shell">

      <header className="topbar">

        {/* ===============================================
            BRAND
        =============================================== */}

        <div className="brand">

          <div className="brand-mark">
            <LayoutDashboard size={20} />
          </div>

          <div>

            <div className="brand-name">
              SLA Performance
            </div>

            <div className="brand-sub">
              Customer Analytics Portal
            </div>

          </div>

        </div>


        {/* ===============================================
            TOP RIGHT
        =============================================== */}

        <div className="topbar-right">

          <span className="format-pill">
            {format === "format1"
              ? "FORMAT 1"
              : format.toUpperCase()}
          </span>


          {/* ============================================
              DASHBOARD ACTIONS
          ============================================ */}

          {page === "dashboard" && (
            <>

              <button
                className="ghost-button"
                onClick={openEditor}
              >
                <LayoutDashboard size={16} />

                Edit Data
              </button>


              <button
                className="ghost-button"
                onClick={reset}
              >
                <RotateCcw size={16} />

                New upload
              </button>

            </>
          )}


          {/* ============================================
              EDITOR ACTION
          ============================================ */}

          {page === "editor" && (

            <button
              className="ghost-button"
              onClick={() =>
                setPage("dashboard")
              }
            >
              <LayoutDashboard size={16} />

              Dashboard
            </button>

          )}

        </div>

      </header>


      {/* =================================================
          MAIN PAGE AREA
      ================================================= */}

      <div className="page-wrap">


        {/* =================================================
            UPLOAD
        ================================================= */}

        {page === "upload" && (

          <UploadPage
            files={files}
            setFiles={setFiles}

            format={format}
            setFormat={setFormat}

            onBuild={build}
          />

        )}


        {/* =================================================
            DASHBOARD
        ================================================= */}

        {page === "dashboard" && (

          <DashboardPage
            model={model}

            onEditData={openEditor}

          />

        )}


        {/* =================================================
            DATA EDITOR
        ================================================= */}

        {page === "editor" && (

          <DataEditorPage
            raw={files}

            onBack={() =>
              setPage("dashboard")
            }

            onSave={
              handleEditorSave
            }

            onRecalculate={
              handleEditorRecalculate
            }

          />

        )}

      </div>

    </div>
  );
}