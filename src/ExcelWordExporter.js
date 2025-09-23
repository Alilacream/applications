import React, { useState, useRef } from "react";
import * as XLSX from "xlsx-js-style";
import { saveAs } from "file-saver";

import PizZip from "pizzip";
import Docxtemplater from "docxtemplater";
// dectects tauri app folder
const isTauri = () => {
  return (
    typeof window !== "undefined" &&
    "__TAURI__" in window &&
    typeof window.__TAURI__ !== "undefined"
  );
};
// 👇 Universal file saver — works in browser AND Tauri
const saveFile = async (
  data,
  fileName,
  mimeType = "application/octet-stream"
) => {
  if (isTauri()) {
    console.log("🚀 Tauri detected — using native dialog...");

    try {
      const { save } = await import("@tauri-apps/api/dialog");
      const { writeBinaryFile } = await import("@tauri-apps/api/fs");
      const { open } = await import("@tauri-apps/api/shell");
      const { dirname } = await import("@tauri-apps/api/path");

      // Show that dialog is about to open
      console.log("📂 Opening save dialog...");

      const filePath = await save({
        filters: [
          {
            name: fileName.includes(".xlsx")
              ? "Excel File"
              : fileName.includes(".docx")
              ? "Word Document"
              : "File",
            extensions: [fileName.split(".").pop()],
          },
        ],
        defaultPath: fileName,
      });

      if (!filePath) {
        console.log("❌ User canceled save dialog");
        alert("Export canceled.");
        return false;
      }

      console.log(`✅ Saving file to: ${filePath}`);
      await writeBinaryFile(filePath, data);

      const folderPath = await dirname(filePath);
      console.log(`📂 Opening folder: ${folderPath}`);
      await open(folderPath);

      return true;
    } catch (err) {
      console.error("🚨 Tauri save dialog ERROR:", err);
      alert(`Export failed: ${err.message}`);
      return false;
    }
  } else {
    console.log("🌐 Browser detected — using file-saver download...");
    const { saveAs } = await import("file-saver");
    const blob = new Blob([data], { type: mimeType });
    saveAs(blob, fileName);
    return true;
  }
};

// 👇 Helper function to normalize strings (remove accents, lowercase, clean spaces)
const normalizeString = (str) => {
  if (typeof str !== "string") return "";
  return str
    .normalize("NFD") // Split accented characters
    .replace(/[\u0300-\u036f]/g, "") // Remove diacritics
    .toLowerCase()
    .trim()
    .replace(/\s+/g, " "); // Collapse spaces
};

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

  // Handles any given extra column at the beginning so we can find the first column index:
  // N° PRIX
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

        // ✅ Read sheet as 2D array — with empty string fallback
        const range = XLSX.utils.decode_range(worksheet["!ref"]);
        const raw_data = [];
        for (let R = range.s.r; R <= range.e.r; ++R) {
          const row = [];
          for (let C = range.s.c; C <= range.e.c; ++C) {
            const cell_address = { c: C, r: R };
            const cell_ref = XLSX.utils.encode_cell(cell_address);
            const cell = worksheet[cell_ref];
            row.push(cell ? cell.w || cell.v : ""); // ← Use formatted (w) or raw (v) value
          }
          raw_data.push(row);
        }

        // ✅ Find header row — look for "N° PRIX"
        let headerRowIndex = -1;
        for (let i = 0; i < raw_data.length; i++) {
          const row = raw_data[i];
          if (
            Array.isArray(row) &&
            row.some(
              (cell) =>
                typeof cell === "string" &&
                normalizeString(cell).includes("n° prix")
            )
          ) {
            headerRowIndex = i;
            break;
          }
        }

        if (headerRowIndex === -1) {
          throw new Error(
            "Could not detect header row. Please make sure 'N° PRIX' column exists."
          );
        }

        // ✅ Extract headers and data rows
        const headers = raw_data[headerRowIndex];
        const dataRows = raw_data.slice(headerRowIndex + 1);

        // ✅ Convert to array of objects — with normalized keys
        const jsonData = dataRows
          .filter((row) => row && row.length > 0)
          .map((row) => {
            const obj = {};
            headers.forEach((header, index) => {
              const key = normalizeString(header || `col_${index}`);
              obj[key] = index < row.length ? row[index] : "";
            });
            return obj;
          })
          .filter((obj) =>
            Object.values(obj).some((val) => String(val).trim() !== "")
          );

        if (jsonData.length === 0) {
          throw new Error("No valid data found after header row.");
        }

        setData(jsonData);
        setFileName(file.name);
        setShowError(false);
      } catch (error) {
        console.error("File Read Error:", error);
        setErrorMessage(
          error.message || "Error reading Excel file. Please check format."
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

  // 🎨 Excel Styles
  const headerStyle = {
    fill: { fgColor: { rgb: "43A047" } }, // Dark blue
    font: { color: { rgb: "000000" }, bold: true, sz: 15 },
    alignment: { horizontal: "center", vertical: "center", wrapText: true },
    border: {
      top: { style: "thin", color: { rgb: "000000" } },
      bottom: { style: "thin", color: { rgb: "000000" } },
      left: { style: "thin", color: { rgb: "000000" } },
      right: { style: "thin", color: { rgb: "000000" } },
    },
  };

  const cellStyle = {
    alignment: { horizontal: "left", vertical: "top", wrapText: true },
    border: {
      top: { style: "thin", color: { rgb: "D9D9D9" } },
      bottom: { style: "thin", color: { rgb: "D9D9D9" } },
      left: { style: "thin", color: { rgb: "D9D9D9" } },
      right: { style: "thin", color: { rgb: "D9D9D9" } },
    },
    font: { sz: 11 },
  };

  const numberStyle = {
    ...cellStyle,
    alignment: { horizontal: "right", vertical: "top" },
    numFmt: "#,##0.00",
  };

  // Export Excel (without description) — with sequential ID reset + STYLING
  const exportExcelFile = async () => {
    const selected = getSelectedData();
    if (selected.length === 0) {
      setErrorMessage("No rows selected. Please select at least one row.");
      setShowError(true);
      return;
    }

    try {
      // Detect columns dynamically — using normalized search
      const parseNumber = (value) => {
        const num = parseFloat(value);
        return !isNaN(num) ? num : 0;
      };
      const nPrixCol =
        Object.keys(data[0]).find((col) =>
          normalizeString(col).includes("n° prix")
        ) || normalizeString("N° Prix");
      const titleCol =
        Object.keys(data[0]).find((col) =>
          normalizeString(col).includes("designation")
        ) || normalizeString("designation des ouvrages");
      const unitCol =
        Object.keys(data[0]).find((col) =>
          normalizeString(col).includes("unite")
        ) || normalizeString("unite");
      const qtyCol =
        Object.keys(data[0]).find((col) =>
          normalizeString(col).includes("quantite")
        ) || normalizeString("quantites");
      const priceCol =
        Object.keys(data[0]).find((col) =>
          normalizeString(col).includes("p.u")
        ) || normalizeString("p.u dh.ht");
      const totalCol =
        Object.keys(data[0]).find((col) =>
          normalizeString(col).includes("montant total h.t")
        ) || normalizeString("Montant total H.T");

      // Prepare data
      const excelData = selected.map((row) => {
        const qty = parseNumber(row[qtyCol]);
        const price = parseNumber(row[priceCol]);
        const totalFromSource = parseNumber(row[totalCol]);

        return {
          "N°Prix": row[nPrixCol] || "", // ← PRESERVE ORIGINAL VALUE (A, a, 1, 2, B, etc.)
          Désignation: row[titleCol] || "Sans Titre",
          Unité: row[unitCol] || "",
          Quantité: qty,
          "P.U DH.HT": price,
          "Montant Total HT": totalFromSource || qty * price,
        };
      });

      // Create worksheet
      const ws = XLSX.utils.json_to_sheet(excelData, { skipHeader: true });

      // Define columns
      const headers = [
        "N°Prix",
        "Désignation",
        "Unité",
        "Quantité",
        "P.U DH.HT",
        "Montant Total HT",
      ];

      const columns = [
        { wch: 8 }, // N°Prix
        { wch: 40 }, // Désignation
        { wch: 10 }, // Unité
        { wch: 12 }, // Quantité
        { wch: 15 }, // P.U DH.HT
        { wch: 18 }, // Montant Total HT
      ];
      ws["!cols"] = columns;

      // Add styled headers
      headers.forEach((header, colIndex) => {
        const cellAddress = XLSX.utils.encode_cell({ r: 0, c: colIndex });
        ws[cellAddress] = { v: header, s: headerStyle };
      });

      // Style data cells
      for (let rowIndex = 1; rowIndex <= excelData.length; rowIndex++) {
        for (let colIndex = 0; colIndex < headers.length; colIndex++) {
          const cellAddress = XLSX.utils.encode_cell({
            r: rowIndex,
            c: colIndex,
          });
          if (!ws[cellAddress]) continue;

          if ([3, 4, 5].includes(colIndex)) {
            // Numeric columns
            ws[cellAddress].s = numberStyle;
            ws[cellAddress].t = "n";
            ws[cellAddress].v = Number(ws[cellAddress].v);
          } else {
            ws[cellAddress].s = cellStyle;
          }
        }
      }

      // Finalize and export
      const wb = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(wb, ws, "Selected Projects");

      const exportFileName = fileName
        ? `export_${fileName.replace(/\.[^/.]+$/, "")}.xlsx`
        : "project_export.xlsx";

      // ✅ START REPLACEMENT — Replace XLSX.writeFile with this block
      // ✅ Generate buffer (same as before)
      const excelBuffer = XLSX.write(wb, { bookType: "xlsx", type: "array" });
      const uint8Array = new Uint8Array(excelBuffer);

      // ✅ Use universal saver
      const success = await saveFile(
        uint8Array,
        exportFileName,
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
      );

      if (!success) {
        setErrorMessage("Export canceled by user.");
        setShowError(true);
        return;
      }

      alert(`✅ Excel exported successfully!`);

      // ✅ END REPLACEMENT
    } catch (error) {
      console.error("Export Excel Error:", error);
      setErrorMessage("Error exporting Excel file.");
      setShowError(true);
    }
  };

  // Export Word (with description) — FIXED
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

      // 🔍 Detect columns dynamically — with normalization
      const titleCol =
        Object.keys(data[0]).find((col) =>
          normalizeString(col).includes("designation")
        ) || "designation des ouvrages";

      const descCol =
        Object.keys(data[0]).find((col) =>
          normalizeString(col).includes("descriptif")
        ) || "descriptif";

      // Prepare data
      const projects = selected.map((row, indexId) => ({
        id: indexId + 1,
        title: row[titleCol] || "Sans Titre",
        descriptif: row[descCol] || "",
      }));
      //.filter((project) => project.descriptif.trim() !== ""); // ← Skip if empty or only whitespace
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
      const wordFileName = fileName
        ? `descriptions_${fileName.replace(/\.[^/.]+$/, "")}.docx`
        : "project_descriptions.docx";

      const success = await saveFile(
        await blob.arrayBuffer(),
        wordFileName,
        "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
      );

      if (!success) {
        setErrorMessage("Export canceled by user.");
        setShowError(true);
        return;
      }

      alert(`✅ Word document exported successfully!`);
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
                    {Object.keys(data[0]).map((key) => {
                      let displayName = key;
                      if (normalizeString(key).includes("n° prix"))
                        displayName = "Code";
                      else if (normalizeString(key).includes("designation"))
                        displayName = "Désignation";
                      else if (normalizeString(key).includes("unite"))
                        displayName = "Unité";
                      else if (normalizeString(key).includes("quantit"))
                        displayName = "Qté";
                      else if (normalizeString(key).includes("p.u"))
                        displayName = "P.U HT";
                      else if (normalizeString(key).includes("motant"))
                        displayName = "Total HT";
                      else if (normalizeString(key).includes("description"))
                        displayName = "Description (longue)";

                      return (
                        <th
                          key={key}
                          className={`table-head-cell ${
                            normalizeString(key).includes("description")
                              ? "description-column"
                              : ""
                          }`}
                        >
                          {displayName}
                        </th>
                      );
                    })}
                  </tr>
                </thead>

                <tbody>
                  {data.map((row, index) => {
                    const keys = Object.keys(row);
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

                          if (
                            normalizeString(key).includes("description") &&
                            displayValue.length > 50
                          ) {
                            displayValue =
                              displayValue.substring(0, 50) + "...";
                          }

                          return (
                            <td
                              key={cellIndex}
                              className="table-cell"
                              title={String(value)}
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
                            !normalizeString(col).includes("description") &&
                            normalizeString(col) !== "id"
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
                            normalizeString(key) === "id" ||
                            normalizeString(key).includes("description")
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
