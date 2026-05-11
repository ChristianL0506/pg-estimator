/**
 * Excel exports for estimating methods and the Compare-All view.
 *
 * - generateMethodFactorsWorkbook(methodKey, methodData):
 *     Returns an ExcelJS workbook containing every factor cell in a method,
 *     organized into one sheet per category (Pipe, Welds, Valves, Bolts,
 *     Threads, Other, Cost Params). Used by GET /api/methods/:key/export.
 *
 * - generateCompareWorkbook(compareResult):
 *     Returns an ExcelJS workbook with a Summary sheet (one row per method)
 *     and a Line Items sheet (every BOM item × every method). Used by
 *     POST /api/estimates/:id/compare-methods/export.
 *
 * Both functions share the navy/light styling used elsewhere in excelExport.ts
 * so the deliverables look consistent across the app.
 */

import ExcelJS from "exceljs";

const NAVY = { argb: "FF1A3650" };
const HEADER_FILL: ExcelJS.Fill = { type: "pattern", pattern: "solid", fgColor: NAVY };
const HEADER_FONT: Partial<ExcelJS.Font> = { color: { argb: "FFFFFFFF" }, bold: true, size: 10 };
const SECTION_FILL: ExcelJS.Fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFE2E8F0" } };
const ZEBRA_FILL: ExcelJS.Fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFF8FAFC" } };
const BORDER_THIN: Partial<ExcelJS.Borders> = {
  top: { style: "thin", color: { argb: "FFB4BEC8" } },
  left: { style: "thin", color: { argb: "FFB4BEC8" } },
  bottom: { style: "thin", color: { argb: "FFB4BEC8" } },
  right: { style: "thin", color: { argb: "FFB4BEC8" } },
};

/**
 * Render a key/value table of factor cells onto a worksheet. Used for any
 * simple "size -> { col1, col2, ... }" map (welds, pipe, valves, bolts, ...).
 *
 * The table is sorted by NPS size where possible. Non-numeric keys fall to
 * the bottom in original order.
 */
function renderFactorTable(
  ws: ExcelJS.Worksheet,
  startRow: number,
  title: string,
  table: Record<string, any> | undefined | null,
): number {
  if (!table || typeof table !== "object" || Object.keys(table).length === 0) {
    ws.getCell(startRow, 1).value = `${title} — no data`;
    ws.getCell(startRow, 1).font = { italic: true, color: { argb: "FF777777" } };
    return startRow + 2;
  }

  // Section header
  ws.mergeCells(startRow, 1, startRow, 6);
  const headerCell = ws.getCell(startRow, 1);
  headerCell.value = title;
  headerCell.fill = HEADER_FILL;
  headerCell.font = HEADER_FONT;
  headerCell.alignment = { vertical: "middle", horizontal: "left", indent: 1 };

  // Detect column keys from the first row
  const sizeKeys = Object.keys(table);
  const firstVal = table[sizeKeys[0]];
  if (typeof firstVal !== "object" || firstVal === null) {
    // Simple flat table (e.g. Other: { Hydro: { factor: 16 }, ... } when value is just a number)
    // Render as 2-column: key | value
    let r = startRow + 1;
    ws.getCell(r, 1).value = "Item";
    ws.getCell(r, 2).value = "Value";
    for (const c of [1, 2]) {
      const cell = ws.getCell(r, c);
      cell.fill = SECTION_FILL;
      cell.font = { bold: true, size: 10 };
      cell.border = BORDER_THIN;
    }
    r++;
    let rowNum = 0;
    for (const k of sizeKeys) {
      const cell = ws.getCell(r, 1); cell.value = k; cell.border = BORDER_THIN;
      const vc = ws.getCell(r, 2); vc.value = table[k]; vc.border = BORDER_THIN;
      if (rowNum % 2 === 1) { cell.fill = ZEBRA_FILL; vc.fill = ZEBRA_FILL; }
      rowNum++;
      r++;
    }
    return r + 1;
  }

  // Nested table — discover sub-keys
  const subKeys: string[] = [];
  for (const k of sizeKeys) {
    const v = table[k];
    if (v && typeof v === "object") {
      for (const sk of Object.keys(v)) {
        if (sk.startsWith("_")) continue; // hide private notes (e.g. _note)
        if (!subKeys.includes(sk)) subKeys.push(sk);
      }
    }
  }

  // Column headers: Size | sub-key1 | sub-key2 | ...
  let r = startRow + 1;
  ws.getCell(r, 1).value = "Size";
  for (let i = 0; i < subKeys.length; i++) {
    ws.getCell(r, 2 + i).value = subKeys[i];
  }
  for (let c = 1; c <= 1 + subKeys.length; c++) {
    const cell = ws.getCell(r, c);
    cell.fill = SECTION_FILL;
    cell.font = { bold: true, size: 10 };
    cell.alignment = { horizontal: c === 1 ? "left" : "right", vertical: "middle" };
    cell.border = BORDER_THIN;
  }
  r++;

  // Stable sort: numeric NPS first ascending, then anything else in original order
  const parseNpsSafe = (s: string): number => {
    const cleaned = s.replace(/["'″]/g, "").trim();
    const compound = cleaned.match(/^(\d+)\s*[-]?\s*(\d+)\/(\d+)/);
    if (compound) return parseInt(compound[1]) + parseInt(compound[2]) / parseInt(compound[3]);
    const frac = cleaned.match(/^(\d+)\/(\d+)/);
    if (frac) return parseInt(frac[1]) / parseInt(frac[2]);
    const m = cleaned.match(/^(\d+(?:\.\d+)?)/);
    if (m) return parseFloat(m[1]);
    return Number.POSITIVE_INFINITY;
  };
  const sortedKeys = [...sizeKeys].sort((a, b) => parseNpsSafe(a) - parseNpsSafe(b));

  let rowNum = 0;
  for (const k of sortedKeys) {
    const sizeCell = ws.getCell(r, 1);
    sizeCell.value = k;
    sizeCell.border = BORDER_THIN;
    sizeCell.alignment = { horizontal: "left", vertical: "middle" };
    for (let i = 0; i < subKeys.length; i++) {
      const valCell = ws.getCell(r, 2 + i);
      const v = (table[k] || {})[subKeys[i]];
      valCell.value = (typeof v === "number" || typeof v === "string") ? v : null;
      valCell.border = BORDER_THIN;
      valCell.alignment = { horizontal: "right", vertical: "middle" };
      if (typeof v === "number") valCell.numFmt = "0.0000";
    }
    if (rowNum % 2 === 1) {
      for (let c = 1; c <= 1 + subKeys.length; c++) ws.getCell(r, c).fill = ZEBRA_FILL;
    }
    rowNum++;
    r++;
  }

  return r + 1; // leave a blank row before the next section
}

export function generateMethodFactorsWorkbook(methodKey: string, methodName: string, methodData: any): ExcelJS.Workbook {
  const wb = new ExcelJS.Workbook();
  wb.creator = "Picou Group Estimator";
  wb.lastModifiedBy = "Picou Group Estimator";
  wb.created = new Date();

  // Cover sheet with method metadata.
  const cover = wb.addWorksheet("Method", { properties: { defaultColWidth: 18 } });
  cover.getColumn(1).width = 22;
  cover.getColumn(2).width = 80;
  let cr = 1;
  cover.mergeCells(cr, 1, cr, 2);
  cover.getCell(cr, 1).value = methodName;
  cover.getCell(cr, 1).font = { bold: true, size: 16, color: { argb: "FF1A3650" } };
  cr += 2;
  const meta: Array<[string, string]> = [
    ["Key", methodKey],
    ["Description", methodData?.description || ""],
    ["Source", methodData?.source || ""],
    ["Exported at", new Date().toISOString()],
  ];
  for (const [k, v] of meta) {
    cover.getCell(cr, 1).value = k;
    cover.getCell(cr, 1).font = { bold: true };
    cover.getCell(cr, 2).value = v;
    cover.getCell(cr, 2).alignment = { wrapText: true, vertical: "top" };
    cr++;
  }

  // For Bill's EI method the shape is different (labor_rates, material_factor_groups, ...).
  // For Justin / Industry / custom-based the shape is labor_factors { pipe, welds, ... }.
  if (methodData?.labor_factors) {
    const factors = methodData.labor_factors;
    const sections: Array<[string, any]> = [
      ["Pipe Handling (MH per LF)", factors.pipe],
      ["Welds (MH per joint)", factors.welds],
      ["Valves (MH per valve)", factors.valves],
      ["Bolts (MH per joint)", factors.bolts],
      ["Threaded Connections", factors.threads],
      ["Other Factors", factors.other],
    ];
    for (const [title, table] of sections) {
      const safeTitle = title.split(" ")[0];
      const ws = wb.addWorksheet(safeTitle, { properties: { defaultColWidth: 14 } });
      ws.getColumn(1).width = 22;
      for (let c = 2; c <= 8; c++) ws.getColumn(c).width = 18;
      renderFactorTable(ws, 1, title, table);
    }
  } else if (methodData?.labor_rates) {
    // Bill's EI tables: nested as { butt_welds_ei: { wall: { schedule: mh } } } etc.
    const lr = methodData.labor_rates;
    for (const [tableName, table] of Object.entries(lr)) {
      const safeTitle = tableName.replace(/_/g, " ");
      const ws = wb.addWorksheet(safeTitle.substring(0, 31), { properties: { defaultColWidth: 14 } });
      ws.getColumn(1).width = 22;
      for (let c = 2; c <= 10; c++) ws.getColumn(c).width = 14;
      renderFactorTable(ws, 1, safeTitle, table as any);
    }
    if (methodData.material_factor_groups) {
      const ws = wb.addWorksheet("Material Groups");
      ws.getColumn(1).width = 8;
      ws.getColumn(2).width = 80;
      ws.getCell(1, 1).value = "Group";
      ws.getCell(1, 2).value = "Description";
      ws.getCell(1, 1).fill = HEADER_FILL; ws.getCell(1, 1).font = HEADER_FONT;
      ws.getCell(1, 2).fill = HEADER_FILL; ws.getCell(1, 2).font = HEADER_FONT;
      let r = 2;
      for (const [g, desc] of Object.entries(methodData.material_factor_groups)) {
        ws.getCell(r, 1).value = g;
        ws.getCell(r, 2).value = String(desc || "");
        ws.getCell(r, 2).alignment = { wrapText: true };
        r++;
      }
    }
  }

  // Cost params (common to all methods)
  if (methodData?.cost_params) {
    const ws = wb.addWorksheet("Cost Params", { properties: { defaultColWidth: 22 } });
    ws.getColumn(1).width = 32;
    ws.getColumn(2).width = 22;
    ws.getCell(1, 1).value = "Parameter";
    ws.getCell(1, 2).value = "Value";
    for (const c of [1, 2]) {
      const cell = ws.getCell(1, c);
      cell.fill = HEADER_FILL; cell.font = HEADER_FONT;
      cell.alignment = { horizontal: c === 1 ? "left" : "right" };
    }
    let r = 2;
    for (const [k, v] of Object.entries(methodData.cost_params)) {
      ws.getCell(r, 1).value = k.replace(/_/g, " ");
      ws.getCell(r, 2).value = typeof v === "number" || typeof v === "string" ? v : JSON.stringify(v);
      ws.getCell(r, 2).alignment = { horizontal: "right" };
      r++;
    }
  }

  return wb;
}

export function generateCompareWorkbook(compareResult: any): ExcelJS.Workbook {
  const wb = new ExcelJS.Workbook();
  wb.creator = "Picou Group Estimator";
  wb.lastModifiedBy = "Picou Group Estimator";
  wb.created = new Date();

  const summary = compareResult.summary || [];
  const lineItems = compareResult.lineItems || [];

  // --- Summary sheet
  const sum = wb.addWorksheet("Summary", { properties: { defaultColWidth: 16 } });
  sum.getColumn(1).width = 28;
  for (let c = 2; c <= 8; c++) sum.getColumn(c).width = 18;
  sum.getCell(1, 1).value = `Estimate Comparison — ${compareResult.estimateName || ""}`;
  sum.getCell(1, 1).font = { bold: true, size: 14, color: { argb: "FF1A3650" } };
  sum.mergeCells(1, 1, 1, 5);
  sum.getCell(2, 1).value = `Items: ${compareResult.itemCount}  ·  Effective labor rate: $${(compareResult.effectiveLaborRate || 0).toFixed(2)}/hr`;
  sum.getCell(2, 1).font = { italic: true, color: { argb: "FF555555" } };
  sum.mergeCells(2, 1, 2, 5);

  const sumHeader = ["Method", "Total MH", "Labor Cost", "Material Cost", "Grand Total"];
  for (let i = 0; i < sumHeader.length; i++) {
    const cell = sum.getCell(4, 1 + i);
    cell.value = sumHeader[i];
    cell.fill = HEADER_FILL; cell.font = HEADER_FONT;
    cell.alignment = { horizontal: i === 0 ? "left" : "right" };
    cell.border = BORDER_THIN;
  }

  let r = 5;
  for (let i = 0; i < summary.length; i++) {
    const s = summary[i];
    const cells = [
      { v: s.label, fmt: undefined as string | undefined },
      { v: s.totalMH, fmt: "#,##0.00" },
      { v: s.totalLaborCost, fmt: '"$"#,##0.00' },
      { v: s.totalMaterialCost, fmt: '"$"#,##0.00' },
      { v: s.totalCost, fmt: '"$"#,##0.00' },
    ];
    for (let c = 0; c < cells.length; c++) {
      const cell = sum.getCell(r, 1 + c);
      cell.value = cells[c].v;
      if (cells[c].fmt) cell.numFmt = cells[c].fmt!;
      cell.alignment = { horizontal: c === 0 ? "left" : "right" };
      cell.border = BORDER_THIN;
      if (i % 2 === 1) cell.fill = ZEBRA_FILL;
    }
    r++;
  }

  // --- Line Items sheet (long form: one row per item per method)
  // Wide form is hard to read in Excel when there are many methods; long form
  // works better with pivot tables / filters.
  const li = wb.addWorksheet("Line Items", { properties: { defaultColWidth: 14 } });
  li.getColumn(1).width = 6;
  li.getColumn(2).width = 14;
  li.getColumn(3).width = 50;
  li.getColumn(4).width = 12;
  li.getColumn(5).width = 8;
  li.getColumn(6).width = 6;
  li.getColumn(7).width = 22;
  li.getColumn(8).width = 12;
  li.getColumn(9).width = 14;
  li.getColumn(10).width = 14;
  li.getColumn(11).width = 14;

  const liHeader = ["#", "Category", "Description", "Size", "Qty", "Unit", "Method", "MH/Unit", "Total MH", "Labor $", "Total $"];
  for (let i = 0; i < liHeader.length; i++) {
    const cell = li.getCell(1, 1 + i);
    cell.value = liHeader[i];
    cell.fill = HEADER_FILL; cell.font = HEADER_FONT;
    cell.alignment = { horizontal: "left" };
    cell.border = BORDER_THIN;
  }

  let lr = 2;
  for (const item of lineItems) {
    for (const s of summary) {
      const bm = item.byMethod?.[s.key];
      if (!bm) continue;
      const row = [
        item.lineNumber,
        item.category,
        item.description,
        item.size,
        item.quantity,
        item.unit,
        s.label,
        bm.mhPerUnit,
        bm.totalMH,
        bm.laborCost,
        bm.totalCost,
      ];
      for (let c = 0; c < row.length; c++) {
        const cell = li.getCell(lr, 1 + c);
        cell.value = row[c] as any;
        cell.border = BORDER_THIN;
        if (c >= 7) {
          cell.alignment = { horizontal: "right" };
          if (c === 7) cell.numFmt = "0.0000";
          else if (c === 8) cell.numFmt = "#,##0.00";
          else cell.numFmt = '"$"#,##0.00';
        }
        if (lr % 2 === 0) cell.fill = ZEBRA_FILL;
      }
      lr++;
    }
  }
  // Freeze the header
  li.views = [{ state: "frozen", ySplit: 1 }];

  return wb;
}
