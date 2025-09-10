import React, { useState, useRef } from "react";
import * as XLSX from "xlsx";
import { HotTable } from "@handsontable/react";
import HyperFormula from "hyperformula";
import "handsontable/dist/handsontable.full.min.css";

export default function App() {
  const [workbookData, setWorkbookData] = useState({});
  const [activeSheet, setActiveSheet] = useState(null);
  const hfInstanceRef = useRef(null);

  // Extract sheet data with formulas
  const extractSheetData = (ws) => {
    const range = XLSX.utils.decode_range(ws["!ref"]);
    const rows = [];
    for (let R = range.s.r; R <= range.e.r; ++R) {
      const row = [];
      for (let C = range.s.c; C <= range.e.c; ++C) {
        const cellAddr = XLSX.utils.encode_cell({ r: R, c: C });
        const cell = ws[cellAddr];
        if (cell) {
          if (cell.f) {
            row.push("=" + cell.f); // keep formula
          } else {
            row.push(cell.v ?? "");
          }
        } else {
          row.push("");
        }
      }
      rows.push(row);
    }
    return rows;
  };

  // Upload Excel
  const handleFileUpload = (e) => {
    const file = e.target.files?.[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = (evt) => {
      const bstr = evt.target.result;
      const wb = XLSX.read(bstr, { type: "binary", cellFormula: true });

      // Collect all sheets first
      const sheetsData = {};
      wb.SheetNames.forEach((sheetName) => {
        const ws = wb.Sheets[sheetName];
        sheetsData[sheetName] = extractSheetData(ws);
      });

      // Init HyperFormula with all sheets at once
      const config = {
        licenseKey: "gpl-v3",
        sheetNamePrefix: "Sheet",
      };
      hfInstanceRef.current = HyperFormula.buildFromSheets(sheetsData, config);

      // Build workbook state
      const newWorkbook = {};
      wb.SheetNames.forEach((sheetName) => {
        const sheetId = hfInstanceRef.current.getSheetId(sheetName);
        const values = hfInstanceRef.current.getSheetValues(sheetId);
        newWorkbook[sheetName] = {
          data: values,
          raw: sheetsData[sheetName],
          colWidths: new Array(sheetsData[sheetName][0]?.length || 10).fill(120),
          rowHeights: new Array(sheetsData[sheetName].length).fill(26),
        };
      });

      setWorkbookData(newWorkbook);
      setActiveSheet(wb.SheetNames[0]);
    };
    reader.readAsBinaryString(file);
  };

  // Manual refresh
  const handleRefresh = () => {
    if (!hfInstanceRef.current) return;
    const updatedWorkbook = { ...workbookData };

    Object.keys(updatedWorkbook).forEach((sheetName) => {
      const sheetId = hfInstanceRef.current.getSheetId(sheetName);
      updatedWorkbook[sheetName].data =
        hfInstanceRef.current.getSheetValues(sheetId);
    });

    setWorkbookData(updatedWorkbook);
  };

  // Export with formulas
  const handleExport = () => {
    const wb = XLSX.utils.book_new();
    Object.entries(workbookData).forEach(([sheetName, sheetObj]) => {
      const ws = XLSX.utils.aoa_to_sheet(sheetObj.raw);
      XLSX.utils.book_append_sheet(wb, ws, sheetName);
    });
    XLSX.writeFile(wb, "updated.xlsx");
  };

  // Auto-refresh after edit
  const handleAfterChange = (changes, source) => {
    if (source === "loadData" || !changes) return;

    const updatedWorkbook = { ...workbookData };

    changes.forEach(([row, col, oldValue, newValue]) => {
      if (activeSheet) {
        const input =
          typeof newValue === "string" && newValue.startsWith("=")
            ? newValue
            : newValue || "";

        updatedWorkbook[activeSheet].raw[row][col] = input;

        const sheetId = hfInstanceRef.current.getSheetId(activeSheet);
        hfInstanceRef.current.setCellContents(
          { sheet: sheetId, col, row },
          [[input]]
        );
      }
    });

    // Recalc all sheets
    Object.keys(updatedWorkbook).forEach((sheetName) => {
      const sheetId = hfInstanceRef.current.getSheetId(sheetName);
      updatedWorkbook[sheetName].data =
        hfInstanceRef.current.getSheetValues(sheetId);
    });

    setWorkbookData(updatedWorkbook);
  };

  const activeSheetData = activeSheet ? workbookData[activeSheet] : null;
  const hasData = activeSheetData && activeSheetData.data?.length > 0;

  return (
    <div style={{ padding: 16 }}>
      <h2>Mexico UI (Multi-sheet + Cross-Sheet Formulas)</h2>

      <div style={{ marginTop: 8, marginBottom: 12 }}>
        <input type="file" accept=".xlsx,.xls" onChange={handleFileUpload} />
        {hasData && (
          <>
            <button
              onClick={handleExport}
              style={{ marginLeft: 8, padding: "6px 12px" }}
            >
              Download Updated Excel
            </button>
            <button
              onClick={handleRefresh}
              style={{ marginLeft: 8, padding: "6px 12px" }}
            >
              Refresh Workbook
            </button>
          </>
        )}
      </div>

      {/* Tabs */}
      {Object.keys(workbookData).length > 0 && (
        <div style={{ display: "flex", borderBottom: "1px solid #ddd" }}>
          {Object.keys(workbookData).map((name) => (
            <button
              key={name}
              onClick={() => setActiveSheet(name)}
              style={{
                padding: "6px 12px",
                border: "none",
                borderBottom:
                  activeSheet === name
                    ? "3px solid blue"
                    : "3px solid transparent",
                background: "transparent",
                fontWeight: activeSheet === name ? "bold" : "normal",
                cursor: "pointer",
              }}
            >
              {name}
            </button>
          ))}
        </div>
      )}

      {/* Grid */}
      {activeSheet && (
        <div style={{ border: "1px solid #ddd", overflow: "auto" }}>
          <HotTable
            data={activeSheetData.data}
            colHeaders={true}
            rowHeaders={true}
            licenseKey="non-commercial-and-evaluation"
            width="100%"
            height={600}
            manualColumnResize={true}
            manualRowResize={true}
            colWidths={activeSheetData.colWidths || undefined}
            rowHeights={activeSheetData.rowHeights || undefined}
            afterChange={handleAfterChange}
          />
        </div>
      )}
    </div>
  );
}