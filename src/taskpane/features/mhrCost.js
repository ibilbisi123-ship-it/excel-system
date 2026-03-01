/* global Excel, fetch */

// Mhr & Cost calculator feature
// - Find header row (first row with non-empty cells)
// - Ensure Description column exists
// - Ensure Mhr column exists (create next to Description if missing)
// - For each data row, fetch TOP 5 candidates from DB.
// - Write candidates to a hidden sheet `_CalcData`.
// - Set Data Validation on Mhr cell to point to the hidden candidates.
// - Default Mhr cell key to the first candidate.
// - If Rate exists, set Cost column to formula: Rate * Parse(Mhr)

export async function runMhrCostOnActiveSheet(updateStatus, addMessage) {
  if (typeof Excel === "undefined") {
    updateStatus && updateStatus("Excel APIs not available", "error");
    return;
  }
  updateStatus && updateStatus("Processing Mhr & Cost...", "processing");
  try {
    // Phase 1: read sheet values and decide target columns without modifying layout
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

      // Find header row within first 30 rows
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
      const rateCol = findCol("Rate");
      let costCol = findCol("Cost");
      const wantCost = rateCol !== -1;

      let setMhrHeader = false;
      let setCostHeader = false;
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
      if (wantCost && costCol === -1) {
        let c = Math.max(headers.length, colCount);
        const cap = c + 20;
        while (c < cap && headers[c] && String(headers[c]).trim() !== "") c++;
        costCol = c;
        setCostHeader = true;
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
        colIndex: used.columnIndex,
        rowIndex: used.rowIndex,
        headerRow,
        descCol,
        mhrCol,
        rateCol,
        costCol,
        wantCost,
        setMhrHeader,
        setCostHeader,
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
      mhrCol,
      rateCol,
      costCol,
      wantCost,
      setMhrHeader,
      setCostHeader,
      selectedRows,
    } = phase1;

    // Phase 2: fetch candidates
    async function fetchCandidates(desc, timeoutMs = 8000) {
      const topMatchesInput = document.getElementById("mhrTopMatches");
      let limit = 5;
      if (topMatchesInput) {
        const val = parseInt(topMatchesInput.value, 10);
        if (!isNaN(val) && val > 0) limit = val;
      }

      const ctrl = new AbortController();
      const t = setTimeout(() => ctrl.abort(), timeoutMs);
      try {
        const resp = await fetch("/api/mhr", {
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

    await Excel.run(async (context) => {
      const ws = context.workbook.worksheets.getActiveWorksheet();
      const absHeaderRow = rowIndex + headerRow;
      const absMhrCol = colIndex + mhrCol;
      const absCostCol = wantCost ? colIndex + costCol : null;

      // Ensure hidden sheet exists
      let refSheet = context.workbook.worksheets.getItemOrNullObject("_CalcData");
      refSheet.load("name");
      await context.sync();

      if (refSheet.isNullObject) {
        refSheet = context.workbook.worksheets.add("_CalcData");
        refSheet.visibility = Excel.SheetVisibility.hidden;
      }

      // Headers
      if (setMhrHeader) {
        ws.getRangeByIndexes(absHeaderRow, absMhrCol, 1, 1).values = [["Mhr"]];
      }
      if (wantCost && setCostHeader) {
        ws.getRangeByIndexes(absHeaderRow, absCostCol, 1, 1).values = [["Cost"]];
      }

      // We iterate through results and update each row
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
          // Write to _CalcData at row `absRow`
          const refRange = refSheet.getRangeByIndexes(absRow, 0, 1, validationValues.length);
          refRange.values = [validationValues];

          // Target Mhr Cell
          const mhrCell = ws.getRangeByIndexes(absRow, absMhrCol, 1, 1);

          // Construct address manually for DataValidation: ='_CalcData'!$A$Row:$E$Row
          const excelRow = absRow + 1;
          const endChar = String.fromCharCode(65 + validationValues.length - 1);
          const address = `='_CalcData'!$A$${excelRow}:$${endChar}$${excelRow}`;

          mhrCell.dataValidation.rule = {
            list: {
              inCellDropDown: true,
              source: address,
            },
          };
          mhrCell.dataValidation.errorAlert = { showAlert: false };

          // Set default value (first choice) - clean value immediately
          const firstVal = validationValues[0].split(" || ")[0];
          mhrCell.values = [[firstVal]];

          // Set Cost Formula
          if (wantCost && absCostCol !== null) {
            const costCell = ws.getRangeByIndexes(absRow, absCostCol, 1, 1);
            const rateOffset = rateCol - costCol;
            const mhrOffset = mhrCol - costCol;

            // Note: R1C1 notation
            const mhrRef = `RC[${mhrOffset}]`;
            const rateRef = `RC[${rateOffset}]`;

            // Formula: =Rate * ValueFromMhr
            // Handle both the temporary "Value || Description" and the cleaned "Value"
            const formula = `=${rateRef} * IFERROR(VALUE(LEFT(${mhrRef}, FIND(" ||", ${mhrRef}) - 1)), IFERROR(VALUE(${mhrRef}), 0))`;
            costCell.formulasR1C1 = [[formula]];
            costCell.numberFormat = [["$#,##0.00"]];
          }
        }
      }

      await context.sync();
    });
    updateStatus && updateStatus("Mhr & Cost processed (Dropdowns created)", "success");
    addMessage && addMessage("I've added dropdowns with top candidates to the Mhr column.", "ai");
  } catch (e) {
    console.error("Mhr & Cost processing failed", e);
    updateStatus && updateStatus("Processing failed: " + e.message, "error");
  }
}
