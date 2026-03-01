/* global Excel, fetch */

// Price calculator feature (my2.db pipeline)
// - Find header row within first 30 rows
// - Ensure Description column exists
// - Ensure Price column exists (create to the right of Description if missing)
// - For each data row, fetch TOP 5 candidates from /api/price
// - Write candidates to hidden sheet `_PriceData`
// - Set Data Validation on Price cell to point to `_PriceData`

export async function runPriceOnActiveSheet(updateStatus, addMessage) {
  if (typeof Excel === "undefined") {
    updateStatus && updateStatus("Excel APIs not available", "error");
    return;
  }
  updateStatus && updateStatus("Processing Price...", "processing");
  try {
    // Phase 1: Read sheet and decide target indices
    const phase1 = await Excel.run(async (context) => {
      const ws = context.workbook.worksheets.getActiveWorksheet();
      const used = ws.getUsedRange(false);
      const selection = context.workbook.getSelectedRange();
      used.load(["values", "rowCount", "columnCount", "rowIndex", "columnIndex"]);
      selection.load(["rowIndex", "rowCount", "columnIndex", "columnCount"]);
      await context.sync();

      const values = used.values || [];
      const rowCount = used.rowCount || (values ? values.length : 0);
      const colCount = used.columnCount || (values && values[0] ? values[0].length : 0);
      if (!values.length) return { ok: false, reason: "No data" };

      // Find header row with 'description'
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
      let priceCol = findCol("Price");
      let setPriceHeader = false;
      if (priceCol === -1) {
        const desired = descCol + 1;
        const cap = Math.max(headers.length, colCount) + 20;
        if (!headers[desired] || String(headers[desired]).trim() === "") {
          priceCol = desired;
        } else {
          let c = desired + 1;
          while (c < cap && headers[c] && String(headers[c]).trim() !== "") c++;
          priceCol = c;
        }
        setPriceHeader = true;
      }

      const selStart = selection.rowIndex - used.rowIndex;
      const selEnd = selStart + selection.rowCount - 1;
      const dataStart = headerRow + 1;
      const dataEnd = rowCount - 1;
      const targetStart = Math.max(dataStart, selStart);
      const targetEnd = Math.min(dataEnd, selEnd);
      if (targetStart > targetEnd)
        return { ok: false, reason: "Select one or more data rows to process" };

      const selectedRows = [];
      for (let r = targetStart; r <= targetEnd; r++) {
        selectedRows.push(r);
      }

      return {
        ok: true,
        values,
        rowCount,
        rowIndex: used.rowIndex,
        colIndex: used.columnIndex,
        headerRow,
        descCol,
        priceCol,
        setPriceHeader,
        selectedRows,
      };
    });

    if (!phase1.ok) {
      updateStatus && updateStatus(phase1.reason || "Processing failed", "error");
      return;
    }

    const {
      values,
      rowIndex,
      colIndex,
      headerRow,
      descCol,
      priceCol,
      setPriceHeader,
      selectedRows,
    } = phase1;

    // Phase 2: fetch candidates
    async function fetchCandidates(desc, timeoutMs = 8000) {
      const topMatchesInput = document.getElementById("priceTopMatches");
      let limit = 5;
      if (topMatchesInput) {
        const val = parseInt(topMatchesInput.value, 10);
        if (!isNaN(val) && val > 0) limit = val;
      }

      const ctrl = new AbortController();
      const t = setTimeout(() => ctrl.abort(), timeoutMs);
      try {
        const resp = await fetch("/api/price", {
          method: "POST",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify({ description: desc, limit: limit }),
          signal: ctrl.signal,
        });
        clearTimeout(t);
        if (!resp.ok) return [];
        const data = await resp.json();
        return data && data.candidates ? data.candidates : [];
      } catch (e) {
        clearTimeout(t);
        return [];
      }
    }

    const rowResults = [];
    const chunkSize = 5;
    for (let i = 0; i < selectedRows.length; i += chunkSize) {
      const chunk = selectedRows.slice(i, i + chunkSize);
      await Promise.all(
        chunk.map(async (r) => {
          const desc = String(values[r]?.[descCol] ?? "").trim();
          if (!desc) {
            rowResults.push({ row: r, candidates: [] });
            return;
          }
          const candidates = await fetchCandidates(desc);
          rowResults.push({ row: r, candidates });
        })
      );
    }

    // Phase 3: write back only target cells
    await Excel.run(async (context) => {
      const ws = context.workbook.worksheets.getActiveWorksheet();
      const absHeaderRow = rowIndex + headerRow;
      const absPriceCol = colIndex + priceCol;

      // Ensure hidden sheet exists
      let refSheet = context.workbook.worksheets.getItemOrNullObject("_PriceData");
      refSheet.load("name");
      await context.sync();

      if (refSheet.isNullObject) {
        refSheet = context.workbook.worksheets.add("_PriceData");
        refSheet.visibility = Excel.SheetVisibility.hidden;
      }

      if (setPriceHeader) {
        ws.getRangeByIndexes(absHeaderRow, absPriceCol, 1, 1).values = [["Price"]];
      }

      for (const item of rowResults) {
        const r = item.row;
        const candidates = item.candidates || [];
        if (candidates.length === 0) continue;

        const absRow = rowIndex + r;

        // Prepare validation strings: "Value || Description"
        const validationValues = candidates.map((c) => {
          let v = c.value || "0";
          let d = c.description || "";
          if (d.length > 200) d = d.substring(0, 200) + "...";
          return `${v} || ${d}`;
        });

        if (validationValues.length > 0) {
          // Write to _PriceData at row `absRow`
          const refRange = refSheet.getRangeByIndexes(absRow, 0, 1, validationValues.length);
          refRange.values = [validationValues];

          // Target Price Cell
          const priceCell = ws.getRangeByIndexes(absRow, absPriceCol, 1, 1);

          const excelRow = absRow + 1;
          const endChar = String.fromCharCode(65 + validationValues.length - 1);
          const address = `='_PriceData'!$A$${excelRow}:$${endChar}$${excelRow}`;

          priceCell.dataValidation.rule = {
            list: {
              inCellDropDown: true,
              source: address,
            },
          };
          priceCell.dataValidation.errorAlert = { showAlert: false };

          // Set default value (clean)
          const firstVal = validationValues[0].split(" || ")[0];
          priceCell.values = [[firstVal]];
          priceCell.numberFormat = [["$#,##0.00"]];
        }
      }

      await context.sync();
    });
    updateStatus && updateStatus("Price processed (Dropdowns created)", "success");
    addMessage && addMessage("I've added dropdowns with top candidates to the Price column.", "ai");
  } catch (e) {
    console.error("Price processing failed", e);
    updateStatus && updateStatus("Processing failed: " + e.message, "error");
  }
}
