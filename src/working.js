import React, { useState, useRef } from "react";
import * as XLSX from "xlsx";
import { HotTable } from "@handsontable/react";
import Handsontable from "handsontable";
import { HyperFormula } from "hyperformula";
import "handsontable/dist/handsontable.full.min.css";

export default function App() {
  const [workbookData, setWorkbookData] = useState({});
  const [activeSheet, setActiveSheet] = useState(null);
  const hfInstanceRef = useRef(null);
  const hotRef = useRef(null);
  const stylesRef = useRef({}); // store per-sheet per-cell style (background colors)

  // helper: convert SheetJS style RGB to CSS hex (SheetJS often gives 'FFRRGGBB' or 'RRGGBB')
  const rgbToHex = (rgb) => {
    if (!rgb) return null;
    const s = String(rgb);
    // drop leading alpha if present (e.g. 'FFRRGGBB')
    const hex = s.length === 8 ? s.slice(2) : s.length === 6 ? s : null;
    return hex ? `#${hex}` : null;
  };

  // Extract sheet data with formulas + capture styles (works if SheetJS populated cell.s and ws['!cols']/ws['!rows'])
  const extractSheetData = (ws) => {
    // safe range detection
    const ref = ws["!ref"];
    if (!ref) return { aoa: [[]], styles: [[]], colWidths: [], rowHeights: [] };

    const range = XLSX.utils.decode_range(ref);
    const rows = [];
    const styles = [];

    for (let R = range.s.r; R <= range.e.r; ++R) {
      const row = [];
      const styleRow = [];
      for (let C = range.s.c; C <= range.e.c; ++C) {
        const cellAddr = XLSX.utils.encode_cell({ r: R, c: C });
        const cell = ws[cellAddr];

        if (cell) {
          // preserve formula strings (ensure leading '=')
          // SheetJS exposes formula as cell.f (no leading '='), and cell.v as value
          if (cell.f) {
            row.push("=" + String(cell.f));
          } else if (typeof cell.v !== "undefined") {
            row.push(cell.v);
          } else {
            row.push("");
          }

          // style extraction (best-effort)
          let bg = null;
          if (cell.s && cell.s.fill) {
            // try common style paths used by SheetJS
            const fg = cell.s.fill.fgColor || cell.s.fill.FgColor || cell.s.fill; // tolerant
            if (fg && (fg.rgb || fg.RGB || fg.rgba)) {
              bg = rgbToHex(fg.rgb || fg.RGB || fg.rgba);
            } else if (cell.s.fill.patternType && cell.s.fill.bgColor) {
              // fallback pattern
              bg = rgbToHex(cell.s.fill.bgColor.rgb || cell.s.fill.bgColor.RGB);
            }
          }
          styleRow.push(bg);
        } else {
          row.push("");
          styleRow.push(null);
        }
      }
      rows.push(row);
      styles.push(styleRow);
    }

    // column widths (SheetJS uses ws['!cols'] with objects that may contain wpx or wch)
    const colWidths = (ws["!cols"] || []).map((c) => {
      if (c && c.wpx) return c.wpx;
      if (c && c.width) return c.width; // some builds
      if (c && c.wch) return Math.round(c.wch * 7); // approximate char -> px
      return 120; // fallback
    });
    // normalize to at least the number of columns
    const colsCount = rows[0]?.length || 0;
    while (colWidths.length < colsCount) colWidths.push(120);

    // row heights (SheetJS uses ws['!rows'] with objects that may contain hpx or h)
    const rowHeights = (ws["!rows"] || []).map((r) => {
      if (r && r.hpx) return r.hpx;
      if (r && r.h) return Math.round(r.h * 1.333); // fallback conversion
      return 26;
    });
    while (rowHeights.length < rows.length) rowHeights.push(26);

    return { aoa: rows, styles, colWidths, rowHeights };
  };

  // Upload Excel
  const handleFileUpload = (e) => {
    const file = e.target.files?.[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = (evt) => {
      const bstr = evt.target.result;

      // request styles and formulas from SheetJS (cellStyles may require a SheetJS build that supports styles)
      const wb = XLSX.read(bstr, {
        type: "binary",
        cellFormula: true,
        cellStyles: true, // best-effort: include styles if SheetJS build supports it
      });

      // Collect all sheets first (raw formulas + styles + dims)
      const sheetsData = {};
      const sheetStyles = {};
      const sheetColWidths = {};
      const sheetRowHeights = {};

      wb.SheetNames.forEach((sheetName) => {
        const ws = wb.Sheets[sheetName];
        const extracted = extractSheetData(ws);
        sheetsData[sheetName] = extracted.aoa;
        sheetStyles[sheetName] = extracted.styles;
        sheetColWidths[sheetName] = extracted.colWidths;
        sheetRowHeights[sheetName] = extracted.rowHeights;
      });

      // Build a HyperFormula external instance (recommended for Handsontable formulas plugin)
      const hf = HyperFormula.buildEmpty({
        licenseKey: "internal-use-in-handsontable",
      });

      // Add sheets & set their content
      wb.SheetNames.forEach((sheetName) => {
        hf.addSheet(sheetName); // create sheet with this name
        const sheetId = hf.getSheetId(sheetName);
        // setSheetContent expects a 2D array
        hf.setSheetContent(sheetId, sheetsData[sheetName]);
      });

      // attach to ref
      hfInstanceRef.current = hf;

      // Build workbook state (preserve raw formulas and computed values)
      const newWorkbook = {};
      wb.SheetNames.forEach((sheetName) => {
        const sheetId = hfInstanceRef.current.getSheetId(sheetName);
        const values = hfInstanceRef.current.getSheetValues(sheetId);
        newWorkbook[sheetName] = {
          data: values, // computed values (for refresh needs)
          raw: sheetsData[sheetName], // raw AOA with formulas (HotTable needs formula strings)
          colWidths: sheetColWidths[sheetName] || new Array(sheetsData[sheetName][0]?.length || 10).fill(120),
          rowHeights: sheetRowHeights[sheetName] || new Array(sheetsData[sheetName].length).fill(26),
        };
      });

      // store styles for render-time
      stylesRef.current = sheetStyles;

      setWorkbookData(newWorkbook);
      // set first sheet as active
      setActiveSheet(wb.SheetNames[0]);

      // Force-hottable update (if table already mounted)
      setTimeout(() => {
        if (hotRef.current?.hotInstance) {
          hotRef.current.hotInstance.updateSettings({
            formulas: { engine: hfInstanceRef.current, sheetName: wb.SheetNames[0] },
          });
          hotRef.current.hotInstance.render();
        }
      }, 0);
    };
    reader.readAsBinaryString(file);
  };

  // Manual refresh (forces recalculation)
  const handleRefresh = () => {
    if (!hfInstanceRef.current) return;
    const updatedWorkbook = { ...workbookData };

    Object.keys(updatedWorkbook).forEach((sheetName) => {
      const sheetId = hfInstanceRef.current.getSheetId(sheetName);
      updatedWorkbook[sheetName].data = hfInstanceRef.current.getSheetValues(sheetId);
    });

    setWorkbookData(updatedWorkbook);
    if (hotRef.current?.hotInstance) hotRef.current.hotInstance.render();
  };

  // Export with formulas (keeps formulas intact)
  const handleExport = () => {
    const wb = XLSX.utils.book_new();
    Object.entries(workbookData).forEach(([sheetName, sheetObj]) => {
      const ws = XLSX.utils.aoa_to_sheet(sheetObj.raw);
      XLSX.utils.book_append_sheet(wb, ws, sheetName);
    });
    XLSX.writeFile(wb, "updated.xlsx");
  };

  // Auto-refresh after edit - batch updates into HyperFormula, update cached computed values, re-render
  const handleAfterChange = (changes, source) => {
    if (source === "loadData" || !changes || !hfInstanceRef.current) return;

    const updatedWorkbook = { ...workbookData };

    // Use HyperFormula batch for better performance
    hfInstanceRef.current.batch(() => {
      changes.forEach(([row, col, oldValue, newValue]) => {
        if (!activeSheet) return;
        const input = typeof newValue === "string" && newValue.startsWith("=") ? newValue : (newValue ?? "");
        // ensure raw array has the row
        if (!updatedWorkbook[activeSheet].raw[row]) updatedWorkbook[activeSheet].raw[row] = [];
        updatedWorkbook[activeSheet].raw[row][col] = input;

        const sheetId = hfInstanceRef.current.getSheetId(activeSheet);
        const cellVal = input === "" ? "" : String(input);
        hfInstanceRef.current.setCellContents({ sheet: sheetId, col, row }, [[cellVal]]);
      });
    });

    // After batch, refresh computed data cache
    Object.keys(updatedWorkbook).forEach((sheetName) => {
      const sheetId = hfInstanceRef.current.getSheetId(sheetName);
      updatedWorkbook[sheetName].data = hfInstanceRef.current.getSheetValues(sheetId);
    });

    setWorkbookData(updatedWorkbook);

    if (hotRef.current?.hotInstance) hotRef.current.hotInstance.render();
  };

  // renderer factory that applies background color then falls back to default text renderer
  const makeBgRenderer = (bg) => {
    return function (hotInstance, td, row, col, prop, value, cellProperties) {
      Handsontable.renderers.TextRenderer.apply(this, arguments);
      if (bg) td.style.background = bg;
    };
  };

  const activeSheetData = activeSheet ? workbookData[activeSheet] : null;
  const hasData = activeSheetData && (activeSheetData.data?.length > 0 || activeSheetData.raw?.length > 0);

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
              onClick={() => {
                setActiveSheet(name);
                // when switching sheet, tell formulas plugin to use that sheetName
                setTimeout(() => {
                  if (hotRef.current?.hotInstance && hfInstanceRef.current) {
                    hotRef.current.hotInstance.updateSettings({
                      formulas: { engine: hfInstanceRef.current, sheetName: name },
                    });
                    hotRef.current.hotInstance.render();
                  }
                }, 0);
              }}
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
            ref={hotRef}
            data={activeSheetData.raw} // pass raw formulas (Handsontable will evaluate via HyperFormula)
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
            formulas={{
              engine: hfInstanceRef.current,
              sheetName: activeSheet,
            }}
            // per-cell renderer for Excel background color (best-effort)
            cells={(row, col) => {
              const sheetStyles = stylesRef.current[activeSheet];
              const bg = sheetStyles?.[row]?.[col] ?? null;
              if (bg) {
                return { renderer: makeBgRenderer(bg) };
              }
              return {};
            }}
          />
        </div>
      )}
    </div>
  );
}