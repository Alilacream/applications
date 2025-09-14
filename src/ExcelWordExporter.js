import React, { useState, useRef } from "react";
import * as XLSX from "xlsx-js-style";
import { saveAs } from "file-saver";

import PizZip from "pizzip";
import Docxtemplater from "docxtemplater";

const ExcelWordExporter = () => {
  const [data, setData] = useState([]);
  const [selectedRows, setSelectedRows] = useState(new Set());
  const [fileName, setFileName] = useState("");
  const [showError, setShowError] = useState(false);
  const [errorMessage, setErrorMessage] = useState("");
  const fileInputRef = useRef(null);

  // Handle file import
  const handleFileImport = () => {
    fileInputRef.current?.click();
  };

  const handleFileChange = (event) => {
    const file = event.target.files[0];
    if (!file) return;

    const validExtensions = [".xlsx", ".xls", ".csv"];
    const fileExtension = file.name
      .toLowerCase()
      .slice(file.name.lastIndexOf("."));

    if (!validExtensions.includes(fileExtension)) {
      setErrorMessage(
        `Invalid file type. Please select an Excel file (.xlsx, .xls, or .csv). You selected: ${fileExtension}`
      );
      setShowError(true);
      event.target.value = "";
      return;
    }

    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, { type: "array" });
        const firstSheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[firstSheetName];
        const jsonData = XLSX.utils.sheet_to_json(worksheet);

        if (jsonData.length === 0) {
          setErrorMessage("No data found in the file.");
          setShowError(true);
          return;
        }

        setData(jsonData);
        setFileName(file.name);
        setShowError(false);
      } catch (error) {
        setErrorMessage(
          "Error reading Excel file. Please make sure the file is not corrupted."
        );
        setShowError(true);
      }
    };
    reader.readAsArrayBuffer(file);
  };

  const closeError = () => {
    setShowError(false);
    setErrorMessage("");
  };

  // Toggle row selection
  const toggleRowSelection = (index) => {
    const newSelected = new Set(selectedRows);
    if (newSelected.has(index)) {
      newSelected.delete(index);
    } else {
      newSelected.add(index);
    }
    setSelectedRows(newSelected);
  };

  // Get selected data
  const getSelectedData = () => {
    return data.filter((_, index) => selectedRows.has(index));
  };


  // Export Excel (without description) — with sequential ID reset
  const exportExcelFile = () => {
    const selected = getSelectedData();
    if (selected.length === 0) {
      setErrorMessage("No rows selected. Please select at least one row.");
      setShowError(true);
      return;
    }

    try {
      const columns = Object.keys(data[0]);
      const excelData = selected.map((row, exportIndex) => {
        const newRow = {};
        columns.forEach((col) => {
          if (col.trim().toLowerCase() === "id") {
            // ✅ Override original ID with sequential number (1, 2, 3...)
            newRow[col] = exportIndex + 1;
          } else if (col.trim().toLowerCase() !== "description") {
            newRow[col] = row[col];
          }
        });
        return newRow;
      });

      const wb = XLSX.utils.book_new();
      const ws = XLSX.utils.json_to_sheet(excelData);
      XLSX.utils.book_append_sheet(wb, ws, "Selected Projects");

      const exportFileName = fileName
        ? `export_${fileName.replace(/\.[^/.]+$/, "")}.xlsx`
        : "project_export.xlsx";

      XLSX.writeFile(wb, exportFileName);
    } catch (error) {
      console.error("Export Excel Error:", error);
      setErrorMessage("Error exporting Excel file.");
      setShowError(true);
    }
  };






  // Export Excel (without description)
  const exportWordFile = async () => {
    const selected = getSelectedData();
    if (selected.length === 0) {
      setErrorMessage("No rows selected. Please select at least one row.");
      setShowError(true);
      return;
    }

    try {
      // Load the official template
      const templateResponse = await fetch("/templates/ROYAUME_DU_MAROC.docx");
      if (!templateResponse.ok) {
        setErrorMessage("Could not load the royaume template");
        setShowError(true);
        return;
      }
      const templateArrayBuffer = await templateResponse.arrayBuffer();

      // Initialize docxtemplater
      const zip = new PizZip(templateArrayBuffer);
      const doc = new Docxtemplater(zip, {
        paragraphLoop: true,
        linebreaks: true,
      });

      // Find column names dynamically
      const descriptionColumn =
        Object.keys(data[0]).find((col) =>
          col.toLowerCase().includes("description")
        ) || "description";

      const priceColumn =
        Object.keys(data[0]).find(
          (col) =>
            col.toLowerCase().includes("prix") ||
            col.toLowerCase().includes("price")
        ) || null;

      // Prepare data with title, description, and optional price
      const projects = selected.map((row,indexId) => ({
        id: indexId + 1,
        title: row["titre de project"] || "Sans Titre",
        description: row[descriptionColumn] || "",
        price: priceColumn ? row[priceColumn] : null, // null = won't render price block
      }));

      // Inject data
      doc.setData({ projects });

      // Render
      doc.render();

      // Generate and save
      const blob = doc.getZip().generate({
        type: "blob",
        mimeType:
          "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
      });

      saveAs(
        blob,
        fileName
          ? `descriptions_${fileName.replace(/\.[^/.]+$/, "")}.docx`
          : "project_descriptions.docx"
      );
    } catch (error) {
      console.error("🚨 docxtemplater Error:", error);
      setErrorMessage(`Export failed: ${error.message || "Unknown error"}`);
      setShowError(true);
    }
  };
  // where the app begins
  return (
    <div className="app-container">
      {/* Header Section */}
      <div className="header-section">
        <h1 className="header-title">📊 Project Data Selector & Exporter</h1>
        <p className="header-subtitle">
          Import your Excel file, select the projects you want, and export clean
          data or rich descriptions — all in a few clicks.
        </p>
      </div>

      {/* Import Button */}
      <div className="import-button-wrapper">
        <button onClick={handleFileImport} className="import-button">
          📁 Import Excel File
        </button>
      </div>

      <input
        className="hidden"
        ref={fileInputRef}
        type="file"
        accept=".xlsx,.xls,.csv"
        onChange={handleFileChange}
      />

      {/* Error Modal */}
      {showError && (
        <div className="error-modal-overlay">
          <div className="error-modal">
            <div className="error-modal-header">
              <div className="error-icon">
                <svg
                  className="error-svg"
                  fill="none"
                  stroke="currentColor"
                  viewBox="0 0 24 24"
                >
                  <path
                    strokeLinecap="round"
                    strokeLinejoin="round"
                    strokeWidth="2"
                    d="M12 9v2m0 4h.01m-6.938 4h13.856c1.54 0 2.502-1.667 1.732-2.5L13.732 4c-.77-.833-1.732-.833-2.5 0L4.268 13.5c-.77.833.192 2.5 1.732 2.5z"
                  />
                </svg>
              </div>
              <div>
                <h3 className="error-title">Oops!</h3>
                <p className="error-message">{errorMessage}</p>
              </div>
            </div>
            <button onClick={closeError} className="error-close-button">
              Close
            </button>
          </div>
        </div>
      )}

      {/* File Info & Data Table */}
      {fileName && data.length > 0 && (
        <div className="file-info-section">
          {/* File Summary Card */}
          <div className="file-summary-card">
            <h3 className="file-summary-title">
              ✅ File Imported Successfully
            </h3>
            <div className="file-stats-grid">
              <div className="stat-card">
                <span className="stat-label">File Name</span>
                <span className="stat-value">{fileName}</span>
              </div>
              <div className="stat-card">
                <span className="stat-label">Total Rows</span>
                <span className="stat-value">{data.length}</span>
              </div>
              <div className="stat-card">
                <span className="stat-label">Columns</span>
                <span className="stat-value">
                  {Object.keys(data[0]).length}
                </span>
              </div>
            </div>
          </div>

          {/* Full Projects Table */}
          <div className="projects-table-container">
            <div className="table-header">
              <h4 className="table-title">📋 All Projects ({data.length})</h4>
            </div>
            <div className="table-wrapper">
              <table className="projects-table">
                <thead>
                  <tr>
                    <th className="table-head-cell">Select</th>
                    {Object.keys(data[0]).map((key) => (
                      <th
                        key={key}
                        className={`table-head-cell ${
                          key.toLowerCase().includes("description")
                            ? "description-column"
                            : ""
                        }`}
                      >
                        {key}
                      </th>
                    ))}
                  </tr>
                </thead>
                <tbody>
                  {data.map((row, index) => {
                    const keys = Object.keys(row); // ← Get column names to detect "description"
                    return (
                      <tr key={index} className="table-row">
                        <td className="table-cell">
                          <input
                            type="checkbox"
                            checked={selectedRows.has(index)}
                            onChange={() => toggleRowSelection(index)}
                            className="select-checkbox"
                          />
                        </td>
                        {Object.values(row).map((value, cellIndex) => {
                          const key = keys[cellIndex];
                          let displayValue = String(value);

                          // 👇 ONLY truncate if column name includes "description"
                          if (
                            key.toLowerCase().includes("description") &&
                            displayValue.length > 50
                          ) {
                            displayValue =
                              displayValue.substring(0, 50) + "...";
                          }

                          return (
                            <td
                              key={cellIndex}
                              className="table-cell"
                              title={String(value)} // ← Full text still visible on hover
                            >
                              {displayValue}
                            </td>
                          );
                        })}
                      </tr>
                    );
                  })}
                </tbody>
              </table>
            </div>
          </div>

          {/* Selected Projects Preview */}
          {selectedRows.size > 0 && (
            <div className="selected-projects-container">
              <div className="table-header">
                <h4 className="table-title">
                  ✅ Selected Projects ({selectedRows.size})
                </h4>
              </div>
              <div className="table-wrapper">
                <table className="projects-table">
                  <thead>
                    <tr>
                      <th className="table-head-cell">Id</th>
                      {Object.keys(data[0])
                        .filter(
                          (col) =>
                            !col.toLowerCase().includes("description") &&
                            col.toLowerCase() !== "id"
                        )
                        .map((key) => (
                          <th key={key} className="table-head-cell">
                            {key}
                          </th>
                        ))}
                    </tr>
                  </thead>
                  <tbody>
                    {getSelectedData().map((row, index) => (
                      <tr key={index} className="table-row">
                        <td className="table-cell">{index + 1}</td>
                        {Object.entries(row).map(([key, value]) => {
                          if (
                            key.toLowerCase() === "id" ||
                            key.toLowerCase().includes("description")
                          )
                            return null;
                          return (
                            <td key={key} className="table-cell">
                              {String(value)}
                            </td>
                          );
                        })}
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            </div>
          )}

          {/* Export Options */}
          <div className="export-section">
            <h4 className="export-title">📤 Export Options</h4>
            <div className="export-cards-grid">
              {/* Excel Export Card */}
              <div className="export-card excel-card">
                <div className="export-card-header">
                  <div className="export-icon">📊</div>
                  <h5 className="export-card-title">Export Clean Excel</h5>
                </div>
                <p className="export-card-description">
                  Exports selected rows without the description column — perfect
                  for data processing or sharing.
                </p>
                <button
                  onClick={exportExcelFile}
                  className="export-button excel-button"
                >
                  📊 Export Excel
                </button>
              </div>

              {/* Word Export Card */}
              <div className="export-card word-card">
                <div className="export-card-header">
                  <div className="export-icon">📝</div>
                  <h5 className="export-card-title">
                    Export Descriptions to Word
                  </h5>
                </div>
                <p className="export-card-description">
                  Exports project IDs, titles, and full descriptions in a
                  beautifully formatted Word document.
                </p>
                <button
                  onClick={exportWordFile}
                  className="export-button word-button"
                >
                  📝 Export Word
                </button>
              </div>
            </div>
          </div>
        </div>
      )}
    </div>
  );
};

export default ExcelWordExporter;
