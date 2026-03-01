/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global Office */

Office.onReady(() => {
  // If needed, Office.js is ready to be called.
});

async function fetchMhr(desc, timeoutMs = 4000) {
  const ctrl = new AbortController();
  const t = setTimeout(() => ctrl.abort(), timeoutMs);
  try {
    const resp = await fetch("/api/mhr", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ description: desc }),
      signal: ctrl.signal,
    });
    clearTimeout(t);
    if (!resp.ok) return "";
    const data = await resp.json();
    return data && data.mhr ? data.mhr : "";
  } catch (e) {
    clearTimeout(t);
    return "";
  }
}

async function calculateMhrForCell(event) {
  try {
    if (typeof Excel === "undefined") {
      event.completed();
      return;
    }

    const info = await Excel.run(async (context) => {
      const ws = context.workbook.worksheets.getActiveWorksheet();
      const used = ws.getUsedRange(false);
      const selection = context.workbook.getSelectedRange();
      used.load(["values", "rowCount", "columnCount", "rowIndex", "columnIndex"]);
      selection.load(["values", "rowIndex", "columnIndex"]);
      await context.sync();

      const values = used.values || [];
      const rowCount = used.rowCount || (values ? values.length : 0);
      const colCount = used.columnCount || (values && values[0] ? values[0].length : 0);
      if (!values.length) return { ok: false, reason: "No data" };

      let headerRow = -1;
      let descCol = -1;
      const searchRows = Math.min(30, rowCount);
      for (let r = 0; r < searchRows; r++) {
        const row = values[r] || [];
        for (let c = 0; c < row.length; c++) {
          const val = row[c];
          if (typeof val === "string" && val.trim().toLowerCase() === "description") {
            headerRow = r;
            descCol = c;
            break;
          }
        }
        if (headerRow !== -1) break;
      }
      if (headerRow === -1 || descCol === -1)
        return { ok: false, reason: "Missing 'description' header" };

      const headers = (values[headerRow] || []).map((v) => (v == null ? "" : String(v).trim()));
      const findCol = (name) =>
        headers.findIndex((h) => h.toLowerCase() === String(name).toLowerCase());
      let mhrCol = findCol("Mhr");
      let setMhrHeader = false;
      if (mhrCol === -1) {
        const desired = descCol + 1;
        const cap = Math.max(headers.length, colCount) + 20;
        if (!headers[desired] || String(headers[desired]).trim() === "") {
          mhrCol = desired;
        } else {
          let c = desired + 1;
          while (c < cap && headers[c] && String(headers[c]).trim() !== "") c++;
          mhrCol = c;
        }
        setMhrHeader = true;
      }

      const absDescCol = used.columnIndex + descCol;
      const absMhrCol = used.columnIndex + mhrCol;
      const absHeaderRow = used.rowIndex + headerRow;
      const absRow = selection.rowIndex;
      const absCol = selection.columnIndex;
      if (absCol !== absDescCol || absRow <= absHeaderRow) {
        return { ok: false, reason: "Selection not in Description column" };
      }

      const desc = String(selection.values?.[0]?.[0] ?? "").trim();
      if (!desc) return { ok: false, reason: "Empty description" };

      return {
        ok: true,
        desc,
        absRow,
        absMhrCol,
        absHeaderRow,
        setMhrHeader,
      };
    });

    if (!info || !info.ok) {
      event.completed();
      return;
    }

    const mhr = await fetchMhr(info.desc);

    await Excel.run(async (context) => {
      const ws = context.workbook.worksheets.getActiveWorksheet();
      if (info.setMhrHeader) {
        const hdrCell = ws.getRangeByIndexes(info.absHeaderRow, info.absMhrCol, 1, 1);
        hdrCell.values = [["Mhr"]];
      }
      const cell = ws.getRangeByIndexes(info.absRow, info.absMhrCol, 1, 1);
      cell.values = [[mhr || ""]];
      await context.sync();
    });
  } finally {
    event.completed();
  }
}

Office.actions.associate("calculateMhrForCell", calculateMhrForCell);
