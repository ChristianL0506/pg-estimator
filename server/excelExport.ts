import ExcelJS from "exceljs";
import type { EstimateProject, EstimateItem } from "@shared/schema";

// ============================================================
// Shared styles
// ============================================================
const NAVY = { argb: "FF1A3650" };
const WHITE_FONT: Partial<ExcelJS.Font> = { color: { argb: "FFFFFFFF" }, bold: true, size: 10 };
const HEADER_FILL: ExcelJS.Fill = { type: "pattern", pattern: "solid", fgColor: NAVY };
const LIGHT_FILL: ExcelJS.Fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFF0F4F8" } };
const BORDER_THIN: Partial<ExcelJS.Borders> = {
  top: { style: "thin", color: { argb: "FFB4BEC8" } },
  left: { style: "thin", color: { argb: "FFB4BEC8" } },
  bottom: { style: "thin", color: { argb: "FFB4BEC8" } },
  right: { style: "thin", color: { argb: "FFB4BEC8" } },
};

function parseSizeNum(size: string): number {
  if (!size) return 0;
  const s = size.replace(/[""''″]/g, "").trim();
  const m = s.match(/^(\d+(?:\.\d+)?)/);
  if (m) return parseFloat(m[1]);
  if (s.includes("/")) {
    const parts = s.split("-");
    if (parts.length === 2) {
      const whole = parseInt(parts[0]) || 0;
      const [n, d] = parts[1].split("/").map(Number);
      return whole + (n / (d || 1));
    }
    const [n, d] = s.split("/").map(Number);
    return n / (d || 1);
  }
  return 0;
}

function categorizeItem(item: EstimateItem): string {
  const d = item.description.toUpperCase();
  const c = item.category.toLowerCase();
  if (c === "pipe" || d.includes("PIPE")) return "pipe";
  if (d.includes("ELBOW") || d.includes("TEE") || d.includes("REDUCER") || d.includes("CAP") ||
      d.includes("COUPLING") || d.includes("NIPPLE") || d.includes("UNION") || c === "fitting") return "fitting";
  if (d.includes("VALVE") || c === "valve") return "valve";
  if (d.includes("GASKET") || c === "gasket") return "gasket";
  if (d.includes("BOLT") || d.includes("STUD") || c === "bolt") return "bolt";
  if (d.includes("FLANGE") || c === "flange") return "flange";
  if (d.includes("WELD") || d.includes("BUTT") || c === "weld") return "weld";
  if (d.includes("SUPPORT") || d.includes("SHOE") || d.includes("HANGER") || c === "support") return "support";
  if (d.includes("STEEL") || d.includes("BEAM") || d.includes("COLUMN") || c === "steel") return "steel";
  if (d.includes("CONCRETE") || d.includes("FOOTING") || c === "concrete") return "concrete";
  return "misc";
}

function addProjectHeader(ws: ExcelJS.Worksheet, project: EstimateProject) {
  ws.getCell("A1").value = "Client:";
  ws.getCell("B1").value = project.client || "";
  ws.getCell("A2").value = "Project:";
  ws.getCell("B2").value = project.name;
  ws.getCell("A3").value = "Location:";
  ws.getCell("B3").value = project.location || "";
  ws.getCell("A4").value = "Estimate No:";
  ws.getCell("B4").value = project.projectNumber || "";
  ws.getCell("A5").value = "Date:";
  ws.getCell("B5").value = new Date().toLocaleDateString();
  for (let r = 1; r <= 5; r++) {
    ws.getCell(`A${r}`).font = { bold: true, size: 9 };
    ws.getCell(`B${r}`).font = { size: 9 };
  }
}

function styleHeaderRow(ws: ExcelJS.Worksheet, row: number, colCount: number) {
  const r = ws.getRow(row);
  for (let c = 1; c <= colCount; c++) {
    const cell = r.getCell(c);
    cell.fill = HEADER_FILL;
    cell.font = WHITE_FONT;
    cell.border = BORDER_THIN;
    cell.alignment = { horizontal: "center", vertical: "middle", wrapText: true };
  }
  r.height = 24;
}

function styleDataRow(ws: ExcelJS.Worksheet, row: number, colCount: number, alt: boolean) {
  const r = ws.getRow(row);
  for (let c = 1; c <= colCount; c++) {
    const cell = r.getCell(c);
    if (alt) cell.fill = LIGHT_FILL;
    cell.border = BORDER_THIN;
    cell.font = { size: 9 };
    cell.alignment = { vertical: "middle" };
  }
}

// ============================================================
// BILL'S BID SHEET FORMAT
// ============================================================
export async function generateBillsWorkbook(project: EstimateProject): Promise<ExcelJS.Workbook> {
  const wb = new ExcelJS.Workbook();
  wb.creator = "Picou Group Contractors — Takeoff & Estimating Tool";
  wb.created = new Date();

  const items = project.items;
  const cats = {
    pipe: items.filter(i => categorizeItem(i) === "pipe"),
    fitting: items.filter(i => categorizeItem(i) === "fitting"),
    valve: items.filter(i => categorizeItem(i) === "valve"),
    gasket: items.filter(i => categorizeItem(i) === "gasket"),
    bolt: items.filter(i => categorizeItem(i) === "bolt"),
    flange: items.filter(i => categorizeItem(i) === "flange"),
    weld: items.filter(i => categorizeItem(i) === "weld"),
    support: items.filter(i => categorizeItem(i) === "support"),
    steel: items.filter(i => categorizeItem(i) === "steel"),
    concrete: items.filter(i => categorizeItem(i) === "concrete"),
    misc: items.filter(i => categorizeItem(i) === "misc"),
  };

  // --- SUMMARY SHEET ---
  const summary = wb.addWorksheet("Summary");
  addProjectHeader(summary, project);
  summary.getColumn(1).width = 16;
  summary.getColumn(2).width = 35;
  summary.getColumn(3).width = 8;
  summary.getColumn(4).width = 12;
  summary.getColumn(5).width = 14;
  summary.getColumn(6).width = 14;
  summary.getColumn(7).width = 14;

  summary.getCell("A7").value = "MECHANICAL TAKE OFF SUMMARY";
  summary.getCell("A7").font = { bold: true, size: 11, color: NAVY };

  // Labor summary
  const laborRows = [
    ["Pipe Handling / Erection", "LF", cats.pipe.reduce((s, i) => s + i.quantity, 0), cats.pipe.reduce((s, i) => s + i.quantity * (i.laborHoursPerUnit || 0), 0)],
    ["Field Welds", "EA", cats.weld.reduce((s, i) => s + i.quantity, 0), cats.weld.reduce((s, i) => s + i.quantity * (i.laborHoursPerUnit || 0), 0)],
    ["Flange Bolt Ups", "EA", cats.flange.reduce((s, i) => s + i.quantity, 0) + cats.bolt.reduce((s, i) => s + i.quantity, 0), (cats.flange.reduce((s, i) => s + i.quantity * (i.laborHoursPerUnit || 0), 0) + cats.bolt.reduce((s, i) => s + i.quantity * (i.laborHoursPerUnit || 0), 0))],
    ["Tagged Items (Valves)", "EA", cats.valve.reduce((s, i) => s + i.quantity, 0), cats.valve.reduce((s, i) => s + i.quantity * (i.laborHoursPerUnit || 0), 0)],
    ["Pipe Supports / Shoes", "EA", cats.support.reduce((s, i) => s + i.quantity, 0), cats.support.reduce((s, i) => s + i.quantity * (i.laborHoursPerUnit || 0), 0)],
    ["Fittings", "EA", cats.fitting.reduce((s, i) => s + i.quantity, 0), cats.fitting.reduce((s, i) => s + i.quantity * (i.laborHoursPerUnit || 0), 0)],
    ["Structural Steel", "EA", cats.steel.reduce((s, i) => s + i.quantity, 0), cats.steel.reduce((s, i) => s + i.quantity * (i.laborHoursPerUnit || 0), 0)],
    ["Civil / Concrete", "EA", cats.concrete.reduce((s, i) => s + i.quantity, 0), cats.concrete.reduce((s, i) => s + i.quantity * (i.laborHoursPerUnit || 0), 0)],
    ["Misc Items", "EA", cats.misc.reduce((s, i) => s + i.quantity, 0), cats.misc.reduce((s, i) => s + i.quantity * (i.laborHoursPerUnit || 0), 0)],
  ];

  summary.getRow(9).values = ["DESCRIPTION", "UOM", "QUANTITY", "MANHOURS", "LABOR $", "MATERIAL $", "TOTAL $"];
  styleHeaderRow(summary, 9, 7);

  let row = 10;
  const totalMH = items.reduce((s, i) => s + i.quantity * (i.laborHoursPerUnit || 0), 0);
  const totalMat = items.reduce((s, i) => s + (i.materialExtension || 0), 0);
  const totalLab = items.reduce((s, i) => s + (i.laborExtension || 0), 0);

  for (const [desc, uom, qty, mh] of laborRows) {
    const matchItems = items.filter(i => {
      const cat = categorizeItem(i);
      if (desc === "Pipe Handling / Erection") return cat === "pipe";
      if (desc === "Field Welds") return cat === "weld";
      if (desc === "Flange Bolt Ups") return cat === "flange" || cat === "bolt";
      if (desc === "Tagged Items (Valves)") return cat === "valve";
      if (desc === "Pipe Supports / Shoes") return cat === "support";
      if (desc === "Fittings") return cat === "fitting";
      if (desc === "Structural Steel") return cat === "steel";
      if (desc === "Civil / Concrete") return cat === "concrete";
      return cat === "misc";
    });
    const matCost = matchItems.reduce((s, i) => s + (i.materialExtension || 0), 0);
    const labCost = matchItems.reduce((s, i) => s + (i.laborExtension || 0), 0);
    summary.getRow(row).values = [desc, uom, qty, Math.round((mh as number) * 100) / 100, labCost, matCost, labCost + matCost];
    summary.getCell(`E${row}`).numFmt = '$#,##0.00';
    summary.getCell(`F${row}`).numFmt = '$#,##0.00';
    summary.getCell(`G${row}`).numFmt = '$#,##0.00';
    summary.getCell(`D${row}`).numFmt = '#,##0.00';
    summary.getCell(`C${row}`).numFmt = '#,##0';
    styleDataRow(summary, row, 7, (row - 10) % 2 === 1);
    row++;
  }

  // Totals row
  const totalQty = items.reduce((s, i) => s + i.quantity, 0);
  summary.getRow(row).values = ["TOTALS", `${items.length} items`, Math.round(totalQty * 100) / 100, totalMH, totalLab, totalMat, totalLab + totalMat];
  const totRow = summary.getRow(row);
  for (let c = 1; c <= 7; c++) {
    totRow.getCell(c).font = { bold: true, size: 10 };
    totRow.getCell(c).border = { top: { style: "double", color: NAVY }, bottom: { style: "double", color: NAVY } };
  }
  summary.getCell(`D${row}`).numFmt = '#,##0.00';
  summary.getCell(`E${row}`).numFmt = '$#,##0.00';
  summary.getCell(`F${row}`).numFmt = '$#,##0.00';
  summary.getCell(`G${row}`).numFmt = '$#,##0.00';
  row += 2;

  // Markups section
  summary.getCell(`A${row}`).value = "ESTIMATE TOTALS";
  summary.getCell(`A${row}`).font = { bold: true, size: 11, color: NAVY };
  row++;
  const sub = totalLab + totalMat;
  const oAmt = sub * (project.markups.overhead / 100);
  const pAmt = (sub + oAmt) * (project.markups.profit / 100);
  const tAmt = totalMat * (project.markups.tax / 100);
  const bAmt = (sub + oAmt + pAmt + tAmt) * (project.markups.bond / 100);
  const grand = sub + oAmt + pAmt + tAmt + bAmt;

  const totals = [
    ["Direct Material", totalMat],
    ["Direct Labor", totalLab],
    ["Subtotal", sub],
    [`Overhead (${project.markups.overhead}%)`, oAmt],
    [`Profit (${project.markups.profit}%)`, pAmt],
    [`Tax (${project.markups.tax}%)`, tAmt],
    [`Bond (${project.markups.bond}%)`, bAmt],
    ["GRAND TOTAL", grand],
  ];
  for (const [label, val] of totals) {
    summary.getCell(`A${row}`).value = label;
    summary.getCell(`B${row}`).value = val;
    summary.getCell(`B${row}`).numFmt = '$#,##0.00';
    if (label === "GRAND TOTAL" || label === "Subtotal") {
      summary.getCell(`A${row}`).font = { bold: true, size: 10 };
      summary.getCell(`B${row}`).font = { bold: true, size: 10 };
    }
    row++;
  }

  row += 2;
  const stRate = project.laborRate || 56;
  const otRate = (project as any).overtimeRate || 79;
  const dtRate = (project as any).doubleTimeRate || 100;
  const otPct = (project as any).overtimePercent || 15;
  const dtPct = (project as any).doubleTimePercent || 2;
  const pd = project.perDiem || 75;
  const stPct = 100 - otPct - dtPct;
  const blended = (stRate * stPct / 100) + (otRate * otPct / 100) + (dtRate * dtPct / 100);
  const pdHr = pd / 10;
  const effective = blended + pdHr;
  summary.getCell(`A${row}`).value = `ST Rate: $${stRate}/hr | OT Rate: $${otRate}/hr | DT Rate: $${dtRate}/hr`;
  summary.getCell(`A${row}`).font = { size: 8, italic: true };
  row++;
  summary.getCell(`A${row}`).value = `OT%: ${otPct}% | DT%: ${dtPct}% | Per Diem: $${pd}/day ($${pdHr.toFixed(2)}/hr)`;
  summary.getCell(`A${row}`).font = { size: 8, italic: true };
  row++;
  summary.getCell(`A${row}`).value = `Blended Rate: $${blended.toFixed(2)}/hr | Effective Rate (incl per diem): $${effective.toFixed(2)}/hr`;
  summary.getCell(`A${row}`).font = { size: 8, italic: true };
  row++;
  summary.getCell(`A${row}`).value = `Total Labor Hours: ${totalMH.toFixed(1)}`;
  summary.getCell(`A${row}`).font = { size: 8, italic: true };
  row++;
  summary.getCell(`A${row}`).value = `Method: ${project.estimateMethod === "bill" ? "Bill's EI Method" : project.estimateMethod === "justin" ? "Justin's Factor Method" : "Manual"}`;
  summary.getCell(`A${row}`).font = { size: 8, italic: true };

  // --- PIPE LABOR SHEET ---
  const pipeLab = wb.addWorksheet("Pipe Labor");
  addProjectHeader(pipeLab, project);
  pipeLab.getCell("E6").value = "PIPING FAB / INSTALL LABOR";
  pipeLab.getCell("E6").font = { bold: true, size: 11, color: NAVY };

  const pipeHeaders = ["#", "Description", "Size", "Qty (LF)", "Schedule", "Material", "MH/LF", "Total MH", "Labor $/LF", "Labor Total", "Mat $/LF", "Mat Total", "Sheet Ref"];
  pipeLab.getRow(9).values = pipeHeaders;
  styleHeaderRow(pipeLab, 9, pipeHeaders.length);
  [5, 30, 8, 10, 10, 8, 8, 10, 10, 12, 10, 12, 18].forEach((w, i) => { pipeLab.getColumn(i + 1).width = w; });

  row = 10;
  for (const item of cats.pipe) {
    pipeLab.getRow(row).values = [
      item.lineNumber, item.description, item.size,
      Math.round(item.quantity * 100) / 100,
      (item as any).itemSchedule || "", (item as any).itemMaterial || "",
      Math.round((item.laborHoursPerUnit || 0) * 10000) / 10000,
      Math.round(item.quantity * (item.laborHoursPerUnit || 0) * 100) / 100,
      item.laborUnitCost || 0,
      item.laborExtension || 0,
      item.materialUnitCost || 0,
      item.materialExtension || 0,
      item.notes || "",
    ];
    ["I", "J", "K", "L"].forEach(col => { pipeLab.getCell(`${col}${row}`).numFmt = '$#,##0.00'; });
    pipeLab.getCell(`D${row}`).numFmt = '#,##0.00';
    pipeLab.getCell(`G${row}`).numFmt = '0.0000';
    pipeLab.getCell(`H${row}`).numFmt = '#,##0.00';
    styleDataRow(pipeLab, row, pipeHeaders.length, (row - 10) % 2 === 1);
    row++;
  }

  // --- MATERIAL SHEETS (Fittings, Valves, Gaskets, Bolts) ---
  const materialSets: [string, string, EstimateItem[]][] = [
    ["Fittings", "FITTINGS MATERIAL PRICING", cats.fitting],
    ["Valves", "VALVE MATERIAL PRICING", cats.valve],
    ["Gaskets", "GASKETS MATERIAL PRICING", cats.gasket],
    ["Studs-Bolts", "STUDS / BOLTS MATERIAL PRICING", cats.bolt],
    ["Welds", "WELD LABOR", cats.weld],
    ["Supports", "PIPE SUPPORTS", cats.support],
    ["Misc Material", "MISCELLANEOUS MATERIALS", cats.misc],
  ];

  for (const [sheetName, title, sheetItems] of materialSets) {
    const ws = wb.addWorksheet(sheetName);
    addProjectHeader(ws, project);
    ws.getCell("A7").value = title;
    ws.getCell("A7").font = { bold: true, size: 11, color: NAVY };

    const headers = ["#", "Qty", "Unit", "Size", "Description", "Mat $/Unit", "Mat Total", "Labor Hrs/Unit", "Labor $/Unit", "Labor Total", "Total Cost", "Sheet Ref"];
    ws.getRow(9).values = headers;
    styleHeaderRow(ws, 9, headers.length);
    [5, 8, 6, 10, 38, 12, 12, 10, 12, 12, 12, 18].forEach((w, i) => { ws.getColumn(i + 1).width = w; });

    let r = 10;
    for (const item of sheetItems) {
      ws.getRow(r).values = [
        item.lineNumber,
        item.unit === "LF" ? Math.round(item.quantity * 100) / 100 : item.quantity,
        item.unit,
        item.size,
        item.description,
        item.materialUnitCost || 0,
        item.materialExtension || 0,
        item.laborHoursPerUnit || 0,
        item.laborUnitCost || 0,
        item.laborExtension || 0,
        item.totalCost || 0,
        item.notes || "",
      ];
      ["F", "G", "I", "J", "K"].forEach(col => { ws.getCell(`${col}${r}`).numFmt = '$#,##0.00'; });
      ws.getCell(`H${r}`).numFmt = '#,##0.0000';
      ws.getCell(`B${r}`).numFmt = item.unit === "LF" ? '#,##0.00' : '#,##0';
      styleDataRow(ws, r, headers.length, (r - 10) % 2 === 1);
      r++;
    }

    // Totals
    if (sheetItems.length > 0) {
      r++;
      ws.getCell(`E${r}`).value = "TOTALS";
      ws.getCell(`E${r}`).font = { bold: true };
      ws.getCell(`G${r}`).value = sheetItems.reduce((s, i) => s + (i.materialExtension || 0), 0);
      ws.getCell(`G${r}`).numFmt = '$#,##0.00';
      ws.getCell(`G${r}`).font = { bold: true };
      ws.getCell(`J${r}`).value = sheetItems.reduce((s, i) => s + (i.laborExtension || 0), 0);
      ws.getCell(`J${r}`).numFmt = '$#,##0.00';
      ws.getCell(`J${r}`).font = { bold: true };
      ws.getCell(`K${r}`).value = sheetItems.reduce((s, i) => s + (i.totalCost || 0), 0);
      ws.getCell(`K${r}`).numFmt = '$#,##0.00';
      ws.getCell(`K${r}`).font = { bold: true };
    }
  }

  return wb;
}

// ============================================================
// JUSTIN'S BID PROGRAM FORMAT
// ============================================================
export async function generateJustinsWorkbook(project: EstimateProject): Promise<ExcelJS.Workbook> {
  const wb = new ExcelJS.Workbook();
  wb.creator = "Picou Group Contractors — Takeoff & Estimating Tool";
  wb.created = new Date();

  const ws = wb.addWorksheet("Estimate");
  ws.getColumn(1).width = 22;
  ws.getColumn(2).width = 14;
  ws.getColumn(3).width = 14;
  ws.getColumn(4).width = 14;
  ws.getColumn(5).width = 10;
  ws.getColumn(6).width = 10;
  ws.getColumn(7).width = 10;
  ws.getColumn(8).width = 14;
  ws.getColumn(9).width = 14;
  ws.getColumn(10).width = 14;
  ws.getColumn(11).width = 16;

  // Title
  ws.getCell("A1").value = "ESTIMATE WORK SHEET";
  ws.getCell("A1").font = { bold: true, size: 14, color: NAVY };
  ws.getCell("A2").value = "   CUSTOMER:";
  ws.getCell("B2").value = project.client || "";
  ws.getCell("E2").value = "JOB:";
  ws.getCell("F2").value = project.name;
  ws.getCell("A2").font = { bold: true, size: 9 };
  ws.getCell("E2").font = { bold: true, size: 9 };

  // Group items by Justin's categories
  const items = project.items;
  const pipeItems: { label: string; qty: number; factor: number; mh: number }[] = [];
  const weldItems: { label: string; qty: number; factor: number; mh: number }[] = [];
  const boltItems: { label: string; qty: number; factor: number; mh: number }[] = [];
  const valveItems: { label: string; qty: number; factor: number; mh: number }[] = [];
  const otherItems: { label: string; qty: number; factor: number; mh: number }[] = [];

  // Aggregate pipe by size
  const pipeBySz = new Map<number, { qty: number; mh: number }>();
  const weldBySz = new Map<number, { qty: number; mh: number }>();
  const boltBySz = new Map<number, { qty: number; mh: number }>();
  const valveBySz = new Map<number, { qty: number; mh: number }>();

  for (const item of items) {
    const cat = categorizeItem(item);
    const sz = parseSizeNum(item.size);
    const mhTotal = item.quantity * (item.laborHoursPerUnit || 0);

    if (cat === "pipe") {
      const agg = pipeBySz.get(sz) || { qty: 0, mh: 0 };
      agg.qty += item.quantity;
      agg.mh += mhTotal;
      pipeBySz.set(sz, agg);
    } else if (cat === "weld") {
      const agg = weldBySz.get(sz) || { qty: 0, mh: 0 };
      agg.qty += item.quantity;
      agg.mh += mhTotal;
      weldBySz.set(sz, agg);
    } else if (cat === "bolt" || cat === "flange") {
      const agg = boltBySz.get(sz) || { qty: 0, mh: 0 };
      agg.qty += item.quantity;
      agg.mh += mhTotal;
      boltBySz.set(sz, agg);
    } else if (cat === "valve") {
      const agg = valveBySz.get(sz) || { qty: 0, mh: 0 };
      agg.qty += item.quantity;
      agg.mh += mhTotal;
      valveBySz.set(sz, agg);
    } else {
      otherItems.push({
        label: `${item.description}${item.size ? " " + item.size : ""}`,
        qty: item.quantity,
        factor: item.laborHoursPerUnit || 0,
        mh: mhTotal,
      });
    }
  }

  for (const [sz, { qty, mh }] of [...pipeBySz.entries()].sort((a, b) => a[0] - b[0])) {
    const label = sz < 2 ? `1/2"-1" Pipe` : `${sz}"Pipe`;
    pipeItems.push({ label, qty: Math.round(qty * 100) / 100, factor: qty > 0 ? Math.round((mh / qty) * 10000) / 10000 : 0, mh: Math.round(mh * 100) / 100 });
  }
  for (const [sz, { qty, mh }] of [...weldBySz.entries()].sort((a, b) => a[0] - b[0])) {
    const label = sz < 2 ? `1/2"-1"Welds` : `${sz}"Welds`;
    weldItems.push({ label, qty, factor: qty > 0 ? Math.round((mh / qty) * 100) / 100 : 0, mh: Math.round(mh * 100) / 100 });
  }
  for (const [sz, { qty, mh }] of [...boltBySz.entries()].sort((a, b) => a[0] - b[0])) {
    const label = `${sz}"Bolts`;
    boltItems.push({ label, qty, factor: qty > 0 ? Math.round((mh / qty) * 100) / 100 : 0, mh: Math.round(mh * 100) / 100 });
  }
  for (const [sz, { qty, mh }] of [...valveBySz.entries()].sort((a, b) => a[0] - b[0])) {
    const label = `${sz}"Valve`;
    valveItems.push({ label, qty, factor: qty > 0 ? Math.round((mh / qty) * 100) / 100 : 0, mh: Math.round(mh * 100) / 100 });
  }

  // Page 1 headers
  ws.getRow(3).values = ["Description", "Qty", "", "", "Factor", "", "", "Sub Total", "", "", "Total"];
  styleHeaderRow(ws, 3, 11);

  let r = 4;
  // Pipe section
  ws.getCell(`A${r}`).value = "Pipe";
  ws.getCell(`A${r}`).font = { bold: true, underline: true, size: 10 };
  r++;
  for (const p of pipeItems) {
    ws.getRow(r).values = [p.label, p.qty, "", "", p.factor, "", "", "", "", "", p.mh];
    ws.getCell(`B${r}`).numFmt = '#,##0.00';
    ws.getCell(`E${r}`).numFmt = '0.0000';
    ws.getCell(`K${r}`).numFmt = '#,##0.00';
    styleDataRow(ws, r, 11, (r - 5) % 2 === 1);
    r++;
  }

  r++;
  // Welds section
  ws.getCell(`A${r}`).value = "Welds";
  ws.getCell(`A${r}`).font = { bold: true, underline: true, size: 10 };
  r++;
  for (const w of weldItems) {
    ws.getRow(r).values = [w.label, w.qty, "", "", w.factor, "", "", "", "", "", w.mh];
    ws.getCell(`B${r}`).numFmt = '#,##0';
    ws.getCell(`E${r}`).numFmt = '0.00';
    ws.getCell(`K${r}`).numFmt = '#,##0.00';
    styleDataRow(ws, r, 11, (r % 2 === 0));
    r++;
  }

  r++;
  // Bolts section
  ws.getCell(`A${r}`).value = "Bolts";
  ws.getCell(`A${r}`).font = { bold: true, underline: true, size: 10 };
  r++;
  for (const b of boltItems) {
    ws.getRow(r).values = [b.label, b.qty, "", "", b.factor, "", "", "", "", "", b.mh];
    ws.getCell(`B${r}`).numFmt = '#,##0';
    ws.getCell(`E${r}`).numFmt = '0.00';
    ws.getCell(`K${r}`).numFmt = '#,##0.00';
    styleDataRow(ws, r, 11, (r % 2 === 0));
    r++;
  }

  r++;
  // Valves section
  ws.getCell(`A${r}`).value = "Valves";
  ws.getCell(`A${r}`).font = { bold: true, underline: true, size: 10 };
  r++;
  for (const v of valveItems) {
    ws.getRow(r).values = [v.label, v.qty, "", "", v.factor, "", "", "", "", "", v.mh];
    ws.getCell(`B${r}`).numFmt = '#,##0';
    ws.getCell(`E${r}`).numFmt = '0.00';
    ws.getCell(`K${r}`).numFmt = '#,##0.00';
    styleDataRow(ws, r, 11, (r % 2 === 0));
    r++;
  }

  r++;
  // Other section
  ws.getCell(`A${r}`).value = "Other";
  ws.getCell(`A${r}`).font = { bold: true, underline: true, size: 10 };
  r++;
  for (const o of otherItems) {
    ws.getRow(r).values = [o.label, o.qty, "", "", o.factor, "", "", "", "", "", o.mh];
    ws.getCell(`B${r}`).numFmt = '#,##0';
    ws.getCell(`E${r}`).numFmt = '0.00';
    ws.getCell(`K${r}`).numFmt = '#,##0.00';
    styleDataRow(ws, r, 11, (r % 2 === 0));
    r++;
  }

  // Totals section
  r += 2;
  const totalHours = items.reduce((s, i) => s + i.quantity * (i.laborHoursPerUnit || 0), 0);
  const laborRate = project.laborRate || 56;
  const jOtRate = (project as any).overtimeRate || 79;
  const jDtRate = (project as any).doubleTimeRate || 100;
  const jOtPct = (project as any).overtimePercent || 15;
  const jDtPct = (project as any).doubleTimePercent || 2;
  const perDiem = project.perDiem || 75;
  const jStPct = 100 - jOtPct - jDtPct;
  const jBlended = (laborRate * jStPct / 100) + (jOtRate * jOtPct / 100) + (jDtRate * jDtPct / 100);
  const jPdHr = perDiem / 10;
  const jEffective = jBlended + jPdHr;
  const totalMat = items.reduce((s, i) => s + (i.materialExtension || 0), 0);
  const totalLabor = items.reduce((s, i) => s + (i.laborExtension || 0), 0);
  const sub = totalMat + totalLabor;
  const oAmt = sub * (project.markups.overhead / 100);
  const pAmt = (sub + oAmt) * (project.markups.profit / 100);
  const tAmt = totalMat * (project.markups.tax / 100);
  const bAmt = (sub + oAmt + pAmt + tAmt) * (project.markups.bond / 100);
  const grand = sub + oAmt + pAmt + tAmt + bAmt;

  ws.getCell(`I${r}`).value = "HOURS";
  ws.getCell(`K${r}`).value = totalHours;
  ws.getCell(`K${r}`).numFmt = '#,##0.00';
  ws.getCell(`I${r}`).font = { bold: true };
  ws.getCell(`K${r}`).font = { bold: true };
  r++;
  ws.getCell(`A${r}`).value = "Blended Rate";
  ws.getCell(`K${r}`).value = jBlended;
  ws.getCell(`K${r}`).numFmt = '$#,##0.00';
  r++;
  ws.getCell(`A${r}`).value = `ST: $${laborRate} (${jStPct}%) | OT: $${jOtRate} (${jOtPct}%) | DT: $${jDtRate} (${jDtPct}%)`;
  ws.getCell(`A${r}`).font = { italic: true, size: 9, color: { argb: "FF888888" } };
  r++;
  ws.getCell(`I${r}`).value = "LABOR";
  ws.getCell(`K${r}`).value = totalLabor;
  ws.getCell(`K${r}`).numFmt = '$#,##0.00';
  ws.getCell(`I${r}`).font = { bold: true };
  ws.getCell(`K${r}`).font = { bold: true };
  r++;
  ws.getCell(`I${r}`).value = "Per Diem";
  ws.getCell(`K${r}`).value = `$${perDiem}/day ($${jPdHr.toFixed(2)}/hr)`;
  r++;
  ws.getCell(`I${r}`).value = "Per Diem Total (included in labor above)";
  ws.getCell(`K${r}`).value = totalHours * jPdHr;
  ws.getCell(`K${r}`).numFmt = '$#,##0.00';
  ws.getCell(`I${r}`).font = { italic: true, size: 9, color: { argb: "FF888888" } };
  r++;
  ws.getCell(`I${r}`).value = "Effective Rate (incl per diem)";
  ws.getCell(`K${r}`).value = jEffective;
  ws.getCell(`K${r}`).numFmt = '$#,##0.00';
  ws.getCell(`I${r}`).font = { italic: true, size: 9, color: { argb: "FF888888" } };
  r++;
  ws.getCell(`I${r}`).value = "Material";
  ws.getCell(`K${r}`).value = totalMat;
  ws.getCell(`K${r}`).numFmt = '$#,##0.00';
  r++;
  ws.getCell(`I${r}`).value = `Overhead (${project.markups.overhead}%)`;
  ws.getCell(`K${r}`).value = oAmt;
  ws.getCell(`K${r}`).numFmt = '$#,##0.00';
  r++;
  ws.getCell(`I${r}`).value = `Profit (${project.markups.profit}%)`;
  ws.getCell(`K${r}`).value = pAmt;
  ws.getCell(`K${r}`).numFmt = '$#,##0.00';
  r++;
  ws.getCell(`I${r}`).value = `Tax (${project.markups.tax}%)`;
  ws.getCell(`K${r}`).value = tAmt;
  ws.getCell(`K${r}`).numFmt = '$#,##0.00';
  r++;
  ws.getCell(`I${r}`).value = `Bond (${project.markups.bond}%)`;
  ws.getCell(`K${r}`).value = bAmt;
  ws.getCell(`K${r}`).numFmt = '$#,##0.00';
  r += 2;
  ws.getCell(`I${r}`).value = "TOTAL:";
  ws.getCell(`K${r}`).value = grand;
  ws.getCell(`I${r}`).font = { bold: true, size: 12, color: NAVY };
  ws.getCell(`K${r}`).font = { bold: true, size: 12, color: NAVY };
  ws.getCell(`K${r}`).numFmt = '$#,##0.00';

  // Page 2 - Line-by-line detail
  const detail = wb.addWorksheet("Detail");
  addProjectHeader(detail, project);
  detail.getCell("A7").value = "LINE-BY-LINE DETAIL";
  detail.getCell("A7").font = { bold: true, size: 11, color: NAVY };

  const dHeaders = ["#", "Category", "Description", "Size", "Qty", "Unit", "Mat $/Unit", "Mat Total", "Hrs/Unit", "Labor $/Unit", "Labor Total", "Total", "Sheet"];
  detail.getRow(9).values = dHeaders;
  styleHeaderRow(detail, 9, dHeaders.length);
  [4, 10, 35, 8, 8, 6, 10, 12, 8, 10, 12, 12, 16].forEach((w, i) => { detail.getColumn(i + 1).width = w; });

  r = 10;
  for (const item of items) {
    detail.getRow(r).values = [
      item.lineNumber, item.category, item.description, item.size,
      item.unit === "LF" ? Math.round(item.quantity * 100) / 100 : item.quantity,
      item.unit,
      item.materialUnitCost || 0, item.materialExtension || 0,
      item.laborHoursPerUnit || 0, item.laborUnitCost || 0,
      item.laborExtension || 0, item.totalCost || 0,
      item.notes || "",
    ];
    ["G", "H", "J", "K", "L"].forEach(col => { detail.getCell(`${col}${r}`).numFmt = '$#,##0.00'; });
    detail.getCell(`I${r}`).numFmt = '0.0000';
    detail.getCell(`E${r}`).numFmt = item.unit === "LF" ? '#,##0.00' : '#,##0';
    styleDataRow(detail, r, dHeaders.length, (r - 10) % 2 === 1);
    r++;
  }

  return wb;
}
