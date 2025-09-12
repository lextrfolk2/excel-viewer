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
  const stylesRef = useRef({}); // per-sheet per-cell background colors (hex)

  // helper: convert Excel 'char width' or wch/wpx to px (approx)
  const excelWidthToPx = (w) => {
    if (typeof w === "number") {
      // For Excel: px ≈ width * 7 + 5 (approx for default fonts)
      return Math.round(w * 7 + 5);
    }
    return undefined;
  };

  // safe color normalizer (sheetcolor like '00305496' -> '#305496')
  const colorToCss = (rgb) => {
    if (!rgb) return null;
    const s = String(rgb);
    const body = s.length === 8 ? s.slice(2) : s.length === 6 ? s : s;
    return body ? `#${body}` : null;
  };

  // Extract sheet data with formulas + styles + column/row dims.
  // Returns: { aoa, styles, colWidthsPx (array or undefined), rowHeightsPx (array or undefined), hasExplicitCols, hasExplicitRows }
  const extractSheetData = (ws) => {
    const ref = ws["!ref"];
    if (!ref) return { aoa: [[]], styles: [[]], colWidthsPx: undefined, rowHeightsPx: undefined, hasExplicitCols: false, hasExplicitRows: false };

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
          if (cell.f) {
            row.push("=" + String(cell.f));
          } else if (typeof cell.v !== "undefined") {
            row.push(cell.v);
          } else {
            row.push("");
          }

          // try best-effort style extraction from SheetJS cell.s (if present)
          let bg = null;
          try {
            const s = cell.s;
            if (s && s.fill) {
              // possible paths (tolerant)
              const fg = s.fill.fgColor || s.fill.FgColor || s.fill;
              const bgc = fg && (fg.rgb || fg.RGB || fg.rgba) ? (fg.rgb || fg.RGB || fg.rgba) : null;
              if (bgc) bg = colorToCss(bgc);
              else if (s.fill.bgColor && (s.fill.bgColor.rgb || s.fill.bgColor.RGB)) {
                bg = colorToCss(s.fill.bgColor.rgb || s.fill.bgColor.RGB);
              }
            }
          } catch (e) {
            bg = null;
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

    // Column widths - only if spreadsheet returned '!cols' (SheetJS must support styles)
    let colWidthsPx, rowHeightsPx;
    const cols = ws["!cols"];
    if (Array.isArray(cols) && cols.length > 0) {
      colWidthsPx = cols.map((c) => {
        if (!c) return excelWidthToPx(10);
        if (c.wpx) return c.wpx;
        if (c.width) return excelWidthToPx(c.width);
        if (c.wch) return Math.round(c.wch * 7);
        // fallback numeric width
        if (typeof c === "number") return excelWidthToPx(c);
        return excelWidthToPx(10);
      });
    } else {
      colWidthsPx = undefined; // allow autoColumnSize
    }

    const rowsMeta = ws["!rows"];
    if (Array.isArray(rowsMeta) && rowsMeta.length > 0) {
      rowHeightsPx = rowsMeta.map((r) => {
        if (!r) return 26;
        if (r.hpx) return r.hpx;
        if (r.h) return Math.round(r.h * 1.333);
        return 26;
      });
    } else {
      rowHeightsPx = undefined; // allow autoRowSize
    }

    return {
      aoa: rows,
      styles,
      colWidthsPx,
      rowHeightsPx,
      hasExplicitCols: !!colWidthsPx,
      hasExplicitRows: !!rowHeightsPx,
    };
  };

  // Upload Excel
  const handleFileUpload = (e) => {
    const file = e.target.files?.[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = (evt) => {
      const bstr = evt.target.result;
      const wb = XLSX.read(bstr, {
        type: "binary",
        cellFormula: true,
        cellStyles: true, // best-effort -- many SheetJS builds ignore this
      });

      const sheetsData = {};
      const sheetStyles = {};
      const sheetColWidths = {};
      const sheetRowHeights = {};
      const sheetHasCols = {};
      const sheetHasRows = {};

      wb.SheetNames.forEach((sheetName) => {
        const ws = wb.Sheets[sheetName];
        const extracted = extractSheetData(ws);
        sheetsData[sheetName] = extracted.aoa;
        sheetStyles[sheetName] = extracted.styles;
        sheetColWidths[sheetName] = extracted.colWidthsPx;
        sheetRowHeights[sheetName] = extracted.rowHeightsPx;
        sheetHasCols[sheetName] = extracted.hasExplicitCols;
        sheetHasRows[sheetName] = extracted.hasExplicitRows;
      });

      // Build HyperFormula external instance recommended for Handsontable formulas plugin
      const hf = HyperFormula.buildEmpty({ licenseKey: "internal-use-in-handsontable" });
      wb.SheetNames.forEach((sheetName) => {
        hf.addSheet(sheetName);
        const sheetId = hf.getSheetId(sheetName);
        hf.setSheetContent(sheetId, sheetsData[sheetName]);
      });
      hfInstanceRef.current = hf;

      // Build workbook state (preserve raw formulas and computed values)
      const newWorkbook = {};
      wb.SheetNames.forEach((sheetName) => {
        const sheetId = hfInstanceRef.current.getSheetId(sheetName);
        const values = hfInstanceRef.current.getSheetValues(sheetId);

        newWorkbook[sheetName] = {
          data: values,
          raw: sheetsData[sheetName],
          // only set widths/heights if the sheet actually contained them
          colWidths: sheetHasCols[sheetName] ? sheetColWidths[sheetName] : undefined,
          rowHeights: sheetHasRows[sheetName] ? sheetRowHeights[sheetName] : undefined,
        };
      });

      stylesRef.current = sheetStyles;
      setWorkbookData(newWorkbook);
      setActiveSheet(wb.SheetNames[0]);

      // update HT formulas plugin + force autosize recalcs after mount
      setTimeout(() => {
        if (hotRef.current?.hotInstance && hfInstanceRef.current) {
          hotRef.current.hotInstance.updateSettings({
            formulas: { engine: hfInstanceRef.current, sheetName: wb.SheetNames[0] },
            autoColumnSize: true,
            autoRowSize: true,
          });

          // recalc auto column sizes (if plugin present)
          try {
            const autoCol = hotRef.current.hotInstance.getPlugin("autoColumnSize");
            if (autoCol) {
              if (typeof autoCol.recalculateAllColumnsWidth === "function") {
                autoCol.recalculateAllColumnsWidth();
              } else if (typeof autoCol.calculateColumnsWidth === "function") {
                // fallback: calculate each column
                const colsCount = hotRef.current.hotInstance.countCols();
                for (let i = 0; i < colsCount; i++) {
                  try {
                    autoCol.calculateColumnsWidth(i, 0, true);
                  } catch (e) {}
                }
              }
            }
            const autoRow = hotRef.current.hotInstance.getPlugin("autoRowSize");
            if (autoRow && typeof autoRow.recalculateAllRowsHeight === "function") {
              autoRow.recalculateAllRowsHeight();
            }
          } catch (err) {
            /* ignore plugin API differences across versions */
          }

          hotRef.current.hotInstance.render();
        }
      }, 50);
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
    if (hotRef.current?.hotInstance) {
      // make sure auto-size recalculates after refresh
      try {
        const autoCol = hotRef.current.hotInstance.getPlugin("autoColumnSize");
        if (autoCol && typeof autoCol.recalculateAllColumnsWidth === "function") autoCol.recalculateAllColumnsWidth();
        const autoRow = hotRef.current.hotInstance.getPlugin("autoRowSize");
        if (autoRow && typeof autoRow.recalculateAllRowsHeight === "function") autoRow.recalculateAllRowsHeight();
      } catch (e) {}
      hotRef.current.hotInstance.render();
    }
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

  // Auto-refresh after edit - batch updates into HyperFormula, update cached computed values, re-render & recalc autosize
  const handleAfterChange = (changes, source) => {
    if (source === "loadData" || !changes || !hfInstanceRef.current) return;
    const updatedWorkbook = { ...workbookData };

    hfInstanceRef.current.batch(() => {
      changes.forEach(([row, col, oldValue, newValue]) => {
        if (!activeSheet) return;
        const input = typeof newValue === "string" && newValue.startsWith("=") ? newValue : (newValue ?? "");
        if (!updatedWorkbook[activeSheet].raw[row]) updatedWorkbook[activeSheet].raw[row] = [];
        updatedWorkbook[activeSheet].raw[row][col] = input;
        const sheetId = hfInstanceRef.current.getSheetId(activeSheet);
        const cellVal = input === "" ? "" : String(input);
        hfInstanceRef.current.setCellContents({ sheet: sheetId, col, row }, [[cellVal]]);
      });
    });

    // refresh computed data snapshot
    Object.keys(updatedWorkbook).forEach((sheetName) => {
      const sheetId = hfInstanceRef.current.getSheetId(sheetName);
      updatedWorkbook[sheetName].data = hfInstanceRef.current.getSheetValues(sheetId);
    });

    setWorkbookData(updatedWorkbook);

    // force re-evaluate autosize and re-render
    if (hotRef.current?.hotInstance) {
      try {
        const autoCol = hotRef.current.hotInstance.getPlugin("autoColumnSize");
        if (autoCol && typeof autoCol.recalculateAllColumnsWidth === "function") autoCol.recalculateAllColumnsWidth();
        const autoRow = hotRef.current.hotInstance.getPlugin("autoRowSize");
        if (autoRow && typeof autoRow.recalculateAllRowsHeight === "function") autoRow.recalculateAllRowsHeight();
      } catch (e) {}
      hotRef.current.hotInstance.render();
    }
  };

  // renderer that applies background color then default text renderer
  const makeBgRenderer = (bg) => {
    return function (hotInstance, td, row, col, prop, value, cellProperties) {
      Handsontable.renderers.TextRenderer.apply(this, arguments);
      if (bg) {
        td.style.backgroundColor = bg;
        td.style.backgroundImage = "none";
      }
    };
  };

  // optional helper: apply external JSON of styles (for SheetJS builds that don't include styles)
  // JSON format:
  // { "<sheetName>": { "colWidths": [px,...] | null, "rowHeights": [px,...] | null, "styles": { "r{row}c{col}": "#RRGGBB", ... } } }
  const applyExternalStyles = (json) => {
    if (!json) return;
    // merge into our stylesRef and workbookData if shapes match
    const nextWB = { ...workbookData };
    Object.keys(json).forEach((sheetName) => {
      const entry = json[sheetName];
      if (entry.colWidths) {
        if (!nextWB[sheetName]) nextWB[sheetName] = {};
        nextWB[sheetName].colWidths = entry.colWidths;
      }
      if (entry.rowHeights) {
        if (!nextWB[sheetName]) nextWB[sheetName] = {};
        nextWB[sheetName].rowHeights = entry.rowHeights;
      }
      // convert styles object into 2D array if possible
      if (entry.styles) {
        // build 2D array same dims as raw
        const raw = nextWB[sheetName]?.raw || [[]];
        const styles2d = raw.map((r, ri) => r.map((c, ci) => null));
        Object.entries(entry.styles).forEach(([coord, hex]) => {
          // coord like 'r10c3'
          const m = coord.match(/^r(\d+)c(\d+)$/);
          if (m) {
            const rr = parseInt(m[1], 10);
            const cc = parseInt(m[2], 10);
            if (styles2d[rr]) styles2d[rr][cc] = hex;
          }
        });
        stylesRef.current = { ...stylesRef.current, [sheetName]: styles2d };
      }
    });
    setWorkbookData(nextWB);
    // force plugin recalcs
    setTimeout(() => {
      if (hotRef.current?.hotInstance) {
        try {
          const autoCol = hotRef.current.hotInstance.getPlugin("autoColumnSize");
          if (autoCol && typeof autoCol.recalculateAllColumnsWidth === "function") autoCol.recalculateAllColumnsWidth();
          const autoRow = hotRef.current.hotInstance.getPlugin("autoRowSize");
          if (autoRow && typeof autoRow.recalculateAllRowsHeight === "function") autoRow.recalculateAllRowsHeight();
        } catch (e) {}
        hotRef.current.hotInstance.render();
      }
    }, 50);
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
            <button onClick={handleExport} style={{ marginLeft: 8, padding: "6px 12px" }}>
              Download Updated Excel
            </button>
            <button onClick={handleRefresh} style={{ marginLeft: 8, padding: "6px 12px" }}>
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
                setTimeout(() => {
                  if (hotRef.current?.hotInstance && hfInstanceRef.current) {
                    hotRef.current.hotInstance.updateSettings({
                      formulas: { engine: hfInstanceRef.current, sheetName: name },
                    });
                    // recalc autosize for new sheet
                    try {
                      const autoCol = hotRef.current.hotInstance.getPlugin("autoColumnSize");
                      if (autoCol && typeof autoCol.recalculateAllColumnsWidth === "function") autoCol.recalculateAllColumnsWidth();
                      const autoRow = hotRef.current.hotInstance.getPlugin("autoRowSize");
                      if (autoRow && typeof autoRow.recalculateAllRowsHeight === "function") autoRow.recalculateAllRowsHeight();
                    } catch (e) {}
                    hotRef.current.hotInstance.render();
                  }
                }, 0);
              }}
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

      {/* Grid */}
      {activeSheet && (
        <div style={{ border: "1px solid #ddd", overflow: "auto" }}>
          <HotTable
            ref={hotRef}
            data={activeSheetData.raw}
            colHeaders={true}
            rowHeaders={true}
            licenseKey="non-commercial-and-evaluation"
            width="100%"
            height={600}
            manualColumnResize={true}
            manualRowResize={true}
            autoColumnSize={true}
            autoRowSize={true}
            colWidths={activeSheetData.colWidths || undefined}
            rowHeights={activeSheetData.rowHeights || undefined}
            afterChange={handleAfterChange}
            formulas={{
              engine: hfInstanceRef.current,
              sheetName: activeSheet,
            }}
            cells={(row, col) => {
              const sheetStyles = stylesRef.current[activeSheet];
              const bg = sheetStyles?.[row]?.[col] ?? null;
              if (bg) return { renderer: makeBgRenderer(bg) };
              return {};
            }}
          />
        </div>
      )}
    </div>
  );
}