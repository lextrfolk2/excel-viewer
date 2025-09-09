import React, { useState } from "react";
import * as XLSX from "xlsx";
import { HotTable } from "@handsontable/react";
import HyperFormula from "hyperformula";
import "handsontable/dist/handsontable.full.min.css";

export default function App() {
  const [workbookData, setWorkbookData] = useState({});
  const [activeSheet, setActiveSheet] = useState(null);

  const normalizeSheetData = (ws) => {
    const sheetData = XLSX.utils.sheet_to_json(ws, { header: 1, raw: false });
    const maxCols = sheetData.reduce((max, row) => Math.max(max, row.length), 0);
    const normalized = sheetData.map((row) => {
      const newRow = [...row];
      while (newRow.length < maxCols) newRow.push("");
      return newRow;
    });
    return { normalized, maxCols };
  };

  const handleFileUpload = (e) => {
    const file = e.target.files?.[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = (evt) => {
      const bstr = evt.target.result;
      const wb = XLSX.read(bstr, { type: "binary", cellFormula: true }); // keep formulas

      const newWorkbook = {};
      wb.SheetNames.forEach((sheetName) => {
        const ws = wb.Sheets[sheetName];
        const { normalized, maxCols } = normalizeSheetData(ws);

        newWorkbook[sheetName] = {
          data: normalized,
          colWidths: new Array(maxCols).fill(120),
          rowHeights: new Array(normalized.length).fill(26),
        };
      });

      setWorkbookData(newWorkbook);
      setActiveSheet(wb.SheetNames[0]);
    };
    reader.readAsBinaryString(file);
  };

  const handleExport = () => {
    const wb = XLSX.utils.book_new();
    Object.entries(workbookData).forEach(([sheetName, sheetObj]) => {
      const ws = XLSX.utils.aoa_to_sheet(sheetObj.data);
      XLSX.utils.book_append_sheet(wb, ws, sheetName);
    });
    XLSX.writeFile(wb, "updated.xlsx");
  };

  const activeSheetData = activeSheet ? workbookData[activeSheet] : null;
  const hasData = activeSheetData && activeSheetData.data?.length > 0;

  return (
    <div style={{ padding: 16 }}>
      <h2>Mexico UI (Multi-sheet + Formulas)</h2>

      <div style={{ marginTop: 8, marginBottom: 12 }}>
        <input type="file" accept=".xlsx,.xls" onChange={handleFileUpload} />
        {hasData && (
          <button
            onClick={handleExport}
            style={{ marginLeft: 8, padding: "6px 12px" }}
          >
            Download Updated Excel
          </button>
        )}
      </div>

      {/* Tab Navigation */}
      {Object.keys(workbookData).length > 0 && (
        <div style={{ display: "flex", borderBottom: "1px solid #ddd" }}>
          {Object.keys(workbookData).map((name) => (
            <button
              key={name}
              onClick={() => setActiveSheet(name)}
              style={{
                padding: "6px 12px",
                border: "none",
                borderBottom: activeSheet === name ? "3px solid blue" : "3px solid transparent",
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

      {/* Handsontable */}
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
            formulas={{
              engine: HyperFormula.buildEmpty({
                licenseKey: "gpl-v3",
              }),
            }}
          />
        </div>
      )}
    </div>
  );
}