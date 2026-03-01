/* global Excel */

export async function runCableFillerOnActiveSheet(updateStatus, addMessage) {
  try {
    if (typeof Excel === "undefined") {
      updateStatus("Excel APIs not available", "error");
      return;
    }
    updateStatus("Running cable filler...", "processing");
    await Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getActiveWorksheet();

      const used = sheet.getUsedRange();
      used.load(["rowIndex", "rowCount"]);
      await context.sync();
      let lastRow = Math.max(11, (used.rowIndex || 0) + (used.rowCount || 0));

      const fillDownAndFixValues = async (addressTop, formula, endRow) => {
        const top = sheet.getRange(addressTop);
        if (formula !== null && formula !== undefined) {
          top.formulas = [[formula]];
        }
        const rng = sheet.getRange(
          `${addressTop.replace(/\d+$/, "11")}:${addressTop.replace(/\d+$/, String(endRow))}`
        );
        sheet.getRange(addressTop).autoFill(rng, Excel.AutoFillType.fillCopy);
        await context.sync();
        rng.load("values");
        await context.sync();
        const vals = rng.values;
        rng.values = vals;
        await context.sync();
      };

      await fillDownAndFixValues(
        "U11",
        '=IF(P11=0,"",CONCATENATE(P11,",","  ",R11,",","  ",Q11,",","  ",T11))',
        lastRow
      );
      await fillDownAndFixValues("V11", "=A11", lastRow);
      await fillDownAndFixValues(
        "W11",
        '=IFERROR( IF( ISNUMBER(SEARCH("mm2", LOWER(P11))), SUBSTITUTE(SUBSTITUTE(MID(LOWER(P11), SEARCH("mm2", LOWER(P11)) - 6, 6), "-", ""), " ", ""), IF(ISNUMBER(SEARCH("18", SUBSTITUTE(P11," ",""))), 1, IF(ISNUMBER(SEARCH("16", SUBSTITUTE(P11," ",""))), 1.5, IF(ISNUMBER(SEARCH("14", SUBSTITUTE(P11," ",""))), 2.5, "")))), "")',
        lastRow
      );
      const rngX = sheet.getRange(`X11:X${lastRow}`);
      rngX.values = Array.from({ length: Math.max(0, lastRow - 10) }, () => ["mm2"]);
      await context.sync();
      await fillDownAndFixValues("Y11", '=IF(W11=0,"",W11&X11)', lastRow);
      await fillDownAndFixValues("AA11", '=IFERROR(LEFT(P11, FIND("-", P11)-1), "")', lastRow);
      await fillDownAndFixValues(
        "AB11",
        '=IFERROR( IF(ISNUMBER(FIND("PR", P11)), MID(P11, FIND("-", P11)+1, FIND("PR", P11)-FIND("-", P11)-1)*2, IF(ISNUMBER(FIND("CR", P11)), MID(P11, FIND("-", P11)+1, FIND("CR", P11)-FIND("-", P11)-1)*1, "")), "")',
        lastRow
      );
      await fillDownAndFixValues("AC11", '=IFERROR(LEFT(P11, FIND("-", P11)-1) * 2, "")', lastRow);
      await fillDownAndFixValues("Z11", "=AA11*AB11*AC11", lastRow);
      const lastRowM = lastRow;
      await fillDownAndFixValues(
        "AF11",
        "=IFERROR(INDEX('[Parameters & Ref. (E&I).xlsx]CABLE GLAND (SIZING)'!$D$4:$Q$27, " +
          "MATCH(W11,'[Parameters & Ref. (E&I).xlsx]CABLE GLAND (SIZING)'!$B$4:$B$27,0), " +
          "MATCH(MAX(FILTER('[Parameters & Ref. (E&I).xlsx]CABLE GLAND (SIZING)'!$D$3:$Q$3, " +
          "'[Parameters & Ref. (E&I).xlsx]CABLE GLAND (SIZING)'!$D$3:$Q$3<=AB11)), " +
          "'[Parameters & Ref. (E&I).xlsx]CABLE GLAND (SIZING)'!$D$3:$Q$3,0)),\"Not Found\")",
        lastRowM
      );
      await fillDownAndFixValues("AD11", '=IF(AF11<>"", MID(AF11,2,LEN(AF11)-1), "")', lastRow);
      await fillDownAndFixValues("AE11", '=IF(AF11<>"", LEFT(AF11,1), "")', lastRow);
      await fillDownAndFixValues("AH11", "=AA11", lastRow);
      await fillDownAndFixValues("AI11", "=AC11", lastRow);
      await fillDownAndFixValues("AG11", '=IF(AD11=0,"",AH11*AI11)', lastRow);
      await fillDownAndFixValues("AJ11", "=V11", lastRow);
      await fillDownAndFixValues("AK11", "=L11", lastRow);
      await fillDownAndFixValues("AL11", "=AC11", lastRow);
      await fillDownAndFixValues(
        "AM11",
        '=IF(AND($K11<>"",$M11<>""), IF($K11="F","FIBER CONDUIT, BEND, " & $L11, IF($K11="S","STEEL CONDUIT, BEND, " & $L11, IF($K11="SP","RGS CONDUIT, PVC COATED, BEND, " & $L11, IF($K11="P","PVC CONDUIT, BEND, " & $L11,"")))), "")',
        lastRow
      );
      await fillDownAndFixValues("AN11", "=M11", lastRow);
      await fillDownAndFixValues("AO11", "=L11", lastRow);
      await fillDownAndFixValues(
        "AQ11",
        '=IF(AND($K11<>"",$M11<>""), IF($K11="F","FIBER CONDUIT, BEND, " & $L11, IF($K11="S","STEEL CONDUIT, BEND, " & $L11, IF($K11="SP","RGS CONDUIT, PVC COATED, BEND, " & $L11, IF($K11="P","PVC CONDUIT, BEND, " & $L11,"")))), "")',
        lastRow
      );
      await fillDownAndFixValues("AR11", "=N11", lastRow);
    });
    updateStatus("Cable filler completed", "success");
    addMessage("Cable filler completed on the active sheet.", "ai");
  } catch (e) {
    console.error("Cable filler failed", e);
    updateStatus("Cable filler failed", "error");
  }
}
