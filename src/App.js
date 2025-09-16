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

  const excelWidthToPx = (w) => {
    if (typeof w === "number") {
      return Math.round(w * 7 + 5);
    }
    return undefined;
  };

  const colorToCss = (rgb) => {
    if (!rgb) return null;
    const s = String(rgb);
    const body = s.length === 8 ? s.slice(2) : s.length === 6 ? s : s;
    return body ? `#${body}` : null;
  };

  const extractSheetData = (ws) => {
    const ref = ws["!ref"];
    if (!ref)
      return {
        aoa: [[]],
        styles: [[]],
        colWidthsPx: undefined,
        rowHeightsPx: undefined,
        hasExplicitCols: false,
        hasExplicitRows: false,
      };

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
          if (cell.f) {
            row.push("=" + String(cell.f));
          } else if (typeof cell.v !== "undefined") {
            row.push(cell.v);
          } else {
            row.push("");
          }

          let bg = null;
          try {
            const s = cell.s;
            if (s && s.fill) {
              const fg = s.fill.fgColor || s.fill.FgColor || s.fill;
              const bgc =
                fg && (fg.rgb || fg.RGB || fg.rgba)
                  ? fg.rgb || fg.RGB || fg.rgba
                  : null;
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

    let colWidthsPx, rowHeightsPx;
    const cols = ws["!cols"];
    if (Array.isArray(cols) && cols.length > 0) {
      colWidthsPx = cols.map((c) => {
        if (!c) return excelWidthToPx(10);
        if (c.wpx) return c.wpx;
        if (c.width) return excelWidthToPx(c.width);
        if (c.wch) return Math.round(c.wch * 7);
        if (typeof c === "number") return excelWidthToPx(c);
        return excelWidthToPx(10);
      });
    } else {
      colWidthsPx = undefined;
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
      rowHeightsPx = undefined;
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

  const handleFileUpload = (e) => {
    const file = e.target.files?.[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = (evt) => {
      const bstr = evt.target.result;
      const wb = XLSX.read(bstr, {
        type: "binary",
        cellFormula: true,
        cellStyles: true,
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

      const hf = HyperFormula.buildEmpty({
        licenseKey: "internal-use-in-handsontable",
      });
      wb.SheetNames.forEach((sheetName) => {
        hf.addSheet(sheetName);
        const sheetId = hf.getSheetId(sheetName);
        hf.setSheetContent(sheetId, sheetsData[sheetName]);
      });
      hfInstanceRef.current = hf;

      const newWorkbook = {};
      wb.SheetNames.forEach((sheetName) => {
        const sheetId = hfInstanceRef.current.getSheetId(sheetName);
        const values = hfInstanceRef.current.getSheetValues(sheetId);

        newWorkbook[sheetName] = {
          data: values,
          raw: sheetsData[sheetName],
          colWidths: sheetHasCols[sheetName] ? sheetColWidths[sheetName] : undefined,
          rowHeights: sheetHasRows[sheetName]
            ? sheetRowHeights[sheetName]
            : undefined,
        };
      });

      stylesRef.current = sheetStyles;
      setWorkbookData(newWorkbook);
      setActiveSheet(wb.SheetNames[0]);

      setTimeout(() => {
        if (hotRef.current?.hotInstance && hfInstanceRef.current) {
          hotRef.current.hotInstance.updateSettings({
            formulas: { engine: hfInstanceRef.current, sheetName: wb.SheetNames[0] },
            autoColumnSize: true,
            autoRowSize: true,
          });
          hotRef.current.hotInstance.render();
        }
      }, 50);
    };
    reader.readAsBinaryString(file);
  };

  const handleRefresh = () => {
    if (!hfInstanceRef.current) return;
    const updatedWorkbook = { ...workbookData };

    Object.keys(updatedWorkbook).forEach((sheetName) => {
      const sheetId = hfInstanceRef.current.getSheetId(sheetName);
      updatedWorkbook[sheetName].data = hfInstanceRef.current.getSheetValues(sheetId);
    });

    setWorkbookData(updatedWorkbook);
    if (hotRef.current?.hotInstance) {
      hotRef.current.hotInstance.render();
    }
  };

  const handleExport = () => {
    const wb = XLSX.utils.book_new();
    Object.entries(workbookData).forEach(([sheetName, sheetObj]) => {
      const ws = XLSX.utils.aoa_to_sheet(sheetObj.data);
      XLSX.utils.book_append_sheet(wb, ws, sheetName);
    });
    XLSX.writeFile(wb, "updated.xlsx");
  };

  const handleAfterChange = (changes, source) => {
    if (source === "loadData" || !changes || !hfInstanceRef.current) return;
    const updatedWorkbook = { ...workbookData };

    hfInstanceRef.current.batch(() => {
      changes.forEach(([row, col, oldValue, newValue]) => {
        if (!activeSheet) return;
        const input =
          typeof newValue === "string" && newValue.startsWith("=")
            ? newValue
            : newValue ?? "";
        if (!updatedWorkbook[activeSheet].raw[row])
          updatedWorkbook[activeSheet].raw[row] = [];
        updatedWorkbook[activeSheet].raw[row][col] = input;
        const sheetId = hfInstanceRef.current.getSheetId(activeSheet);
        const cellVal = input === "" ? "" : String(input);
        hfInstanceRef.current.setCellContents(
          { sheet: sheetId, col, row },
          [[cellVal]]
        );
      });
    });

    Object.keys(updatedWorkbook).forEach((sheetName) => {
      const sheetId = hfInstanceRef.current.getSheetId(sheetName);
      updatedWorkbook[sheetName].data = hfInstanceRef.current.getSheetValues(sheetId);
    });

    setWorkbookData(updatedWorkbook);
    if (hotRef.current?.hotInstance) {
      hotRef.current.hotInstance.render();
    }
  };

  const makeBgRenderer = (bg) => {
    return function (hotInstance, td, row, col, prop, value, cellProperties) {
      Handsontable.renderers.TextRenderer.apply(this, arguments);
      if (bg) {
        td.style.backgroundColor = bg;
        td.style.backgroundImage = "none";
      }
    };
  };

  // 🔹 Replace placeholders like {{9999}} using API /data
  const handleReplacePlaceholders = async () => {
    if (!hfInstanceRef.current) return;

    try {
      // 🔹 Call your Java API
      const res = await fetch("http://localhost:9090/excel-viewer/data", {
        method: "GET",
        headers: {
          "Content-Type": "application/json",
        },
      });
  
      if (!res.ok) {
        throw new Error(`API request failed with status ${res.status}`);
      }

      const replacements = await res.json(); // e.g. { "9999": "Hi" }

      const updatedWorkbook = { ...workbookData };

      Object.keys(updatedWorkbook).forEach((sheetName) => {
        const sheetId = hfInstanceRef.current.getSheetId(sheetName);
        const rawSheet = updatedWorkbook[sheetName].raw;

        for (let r = 0; r < rawSheet.length; r++) {
          for (let c = 0; c < rawSheet[r].length; c++) {
            const val = rawSheet[r][c];
            if (typeof val === "string") {
              const match = val.match(/^{{(\w+)}}$/);
              if (match) {
                const key = match[1];
                if (replacements[key] !== undefined) {
                  const newVal = replacements[key];
                  rawSheet[r][c] = newVal;

                  hfInstanceRef.current.setCellContents(
                    { sheet: sheetId, row: r, col: c },
                    [[newVal]]
                  );
                }
              }
            }
          }
        }
        updatedWorkbook[sheetName].data =
          hfInstanceRef.current.getSheetValues(sheetId);
      });

      setWorkbookData(updatedWorkbook);
      if (hotRef.current?.hotInstance) {
        hotRef.current.hotInstance.render();
      }
    } catch (err) {
      console.error("Error replacing placeholders:", err);
    }
  };

  const activeSheetData = activeSheet ? workbookData[activeSheet] : null;
  const hasData =
    activeSheetData &&
    (activeSheetData.data?.length > 0 || activeSheetData.raw?.length > 0);

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
            <button
              onClick={handleReplacePlaceholders}
              style={{ marginLeft: 8, padding: "6px 12px" }}
            >
              Replace Placeholders
            </button>
          </>
        )}
      </div>

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