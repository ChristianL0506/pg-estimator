import jsPDF from "jspdf";
import autoTable from "jspdf-autotable";
import type { TakeoffProject, TakeoffItem, EstimateProject, EstimateItem } from "@shared/schema";

function formatDollar(n: number) {
  return `$${n.toFixed(2)}`;
}

function addPGHeader(doc: jsPDF, title: string) {
  doc.setFontSize(18);
  doc.setTextColor(30, 80, 140);
  doc.text("PICOU GROUP CONTRACTORS", 14, 20);
  doc.setFontSize(12);
  doc.setTextColor(60, 60, 60);
  doc.text(title, 14, 30);
  doc.setDrawColor(30, 80, 140);
  doc.setLineWidth(0.5);
  doc.line(14, 34, 200, 34);
}

export function exportMechanicalPdf(project: TakeoffProject) {
  const doc = new jsPDF({ orientation: "landscape" });
  addPGHeader(doc, `Mechanical BOM — ${project.name}`);

  const rows = project.items.map(item => [
    String(item.lineNumber),
    item.category.toUpperCase(),
    item.size,
    item.description,
    item.unit === "LF" ? item.quantity.toFixed(2) : String(Math.round(item.quantity)),
    item.unit,
    item.schedule || "",
    item.rating || "",
    item.notes || "",
  ]);

  autoTable(doc, {
    startY: 38,
    head: [["#", "Category", "Size", "Description", "Qty", "Unit", "Schedule", "Rating", "Notes"]],
    body: rows,
    styles: { fontSize: 7, cellPadding: 2 },
    headStyles: { fillColor: [30, 80, 140] },
    columnStyles: { 3: { cellWidth: 80 } },
  });

  doc.save(`${project.name}_mechanical_bom.pdf`);
}

export function exportStructuralPdf(project: TakeoffProject) {
  const doc = new jsPDF({ orientation: "landscape" });
  addPGHeader(doc, `Structural BOM — ${project.name}`);

  const rows = project.items.map(item => [
    String(item.lineNumber),
    item.mark || "",
    item.category.replace("_", " ").toUpperCase(),
    item.size,
    item.description,
    item.unit === "LF" || item.unit === "CY" ? item.quantity.toFixed(2) : String(Math.round(item.quantity)),
    item.unit,
    item.grade || "",
    item.weight ? String(item.weight) + " lbs" : "",
    item.notes || "",
  ]);

  autoTable(doc, {
    startY: 38,
    head: [["#", "Mark", "Category", "Size", "Description", "Qty", "Unit", "Grade", "Weight", "Notes"]],
    body: rows,
    styles: { fontSize: 7, cellPadding: 2 },
    headStyles: { fillColor: [30, 80, 140] },
    columnStyles: { 4: { cellWidth: 70 } },
  });

  doc.save(`${project.name}_structural_bom.pdf`);
}

export function exportCivilPdf(project: TakeoffProject) {
  const doc = new jsPDF({ orientation: "landscape" });
  addPGHeader(doc, `Civil Takeoff — ${project.name}`);

  const rows = project.items.map(item => [
    String(item.lineNumber),
    item.category.replace("_", " ").toUpperCase(),
    item.size,
    item.description,
    item.material || "",
    item.depth || "",
    ["LF", "SF", "CY", "SY", "TON"].includes(item.unit) ? item.quantity.toFixed(2) : String(Math.round(item.quantity)),
    item.unit,
    item.notes || "",
  ]);

  autoTable(doc, {
    startY: 38,
    head: [["#", "Category", "Size", "Description", "Material", "Depth", "Qty", "Unit", "Notes"]],
    body: rows,
    styles: { fontSize: 7, cellPadding: 2 },
    headStyles: { fillColor: [30, 80, 140] },
    columnStyles: { 3: { cellWidth: 75 } },
  });

  doc.save(`${project.name}_civil_bom.pdf`);
}

export function exportEstimatePdf(project: EstimateProject) {
  const doc = new jsPDF({ orientation: "landscape" });
  
  // Cover page
  doc.setFontSize(22);
  doc.setTextColor(30, 80, 140);
  doc.text("PICOU GROUP CONTRACTORS", 14, 30);
  doc.setFontSize(16);
  doc.setTextColor(60, 60, 60);
  doc.text("ESTIMATE PACKAGE", 14, 42);
  doc.setFontSize(12);
  doc.text(project.name, 14, 54);
  if (project.projectNumber) doc.text(`Project No: ${project.projectNumber}`, 14, 64);
  if (project.client) doc.text(`Client: ${project.client}`, 14, 74);
  if (project.location) doc.text(`Location: ${project.location}`, 14, 84);
  doc.text(`Date: ${new Date(project.createdAt).toLocaleDateString()}`, 14, 94);

  // Compute totals
  const totalMaterial = project.items.reduce((sum, i) => sum + (i.materialExtension || 0), 0);
  const totalLabor = project.items.reduce((sum, i) => sum + (i.laborExtension || 0), 0);
  const totalHours = project.items.reduce((sum, i) => sum + (i.quantity || 0) * (i.laborHoursPerUnit || 0), 0);
  const subtotal = totalMaterial + totalLabor;
  const overheadAmt = subtotal * (project.markups.overhead / 100);
  const profitAmt = (subtotal + overheadAmt) * (project.markups.profit / 100);
  const taxAmt = (totalMaterial) * (project.markups.tax / 100);
  const bondAmt = (subtotal + overheadAmt + profitAmt + taxAmt) * (project.markups.bond / 100);
  const grandTotal = subtotal + overheadAmt + profitAmt + taxAmt + bondAmt;

  // Summary box
  doc.setDrawColor(30, 80, 140);
  doc.setLineWidth(0.5);
  doc.rect(14, 104, 120, 80);
  doc.setFontSize(10);
  doc.setTextColor(30, 80, 140);
  doc.text("ESTIMATE SUMMARY", 18, 114);
  const summaryRows = [
    ["Material Cost", formatDollar(totalMaterial)],
    ["Labor Cost", formatDollar(totalLabor)],
    ["Total Labor Hours", totalHours.toFixed(1) + " hrs"],
    ["Subtotal", formatDollar(subtotal)],
    [`Overhead (${project.markups.overhead}%)`, formatDollar(overheadAmt)],
    [`Profit (${project.markups.profit}%)`, formatDollar(profitAmt)],
    [`Tax (${project.markups.tax}%)`, formatDollar(taxAmt)],
    [`Bond (${project.markups.bond}%)`, formatDollar(bondAmt)],
  ];
  let y = 120;
  doc.setTextColor(60, 60, 60);
  for (const [label, val] of summaryRows) {
    doc.text(label, 18, y);
    doc.text(val, 120, y, { align: "right" });
    y += 8;
  }
  doc.setFontSize(11);
  doc.setTextColor(30, 80, 140);
  doc.text("GRAND TOTAL", 18, y + 4);
  doc.text(formatDollar(grandTotal), 120, y + 4, { align: "right" });

  // Detail page
  doc.addPage();
  addPGHeader(doc, `Estimate Detail — ${project.name}`);

  const rows = project.items.map(item => [
    String(item.lineNumber),
    item.category.toUpperCase(),
    item.size,
    item.description,
    item.unit === "LF" ? item.quantity.toFixed(2) : String(Math.round(item.quantity)),
    item.unit,
    formatDollar(item.materialUnitCost),
    formatDollar(item.laborUnitCost),
    item.laborHoursPerUnit.toFixed(2),
    formatDollar(item.materialExtension || 0),
    formatDollar(item.laborExtension || 0),
    formatDollar(item.totalCost || 0),
  ]);

  autoTable(doc, {
    startY: 38,
    head: [["#", "Cat", "Size", "Description", "Qty", "Unit", "Mat $", "Labor $", "Hrs/Unit", "Mat Ext", "Labor Ext", "Total"]],
    body: rows,
    foot: [["", "", "", "TOTALS", "", "", "", "", totalHours.toFixed(1), formatDollar(totalMaterial), formatDollar(totalLabor), formatDollar(subtotal)]],
    styles: { fontSize: 7, cellPadding: 2 },
    headStyles: { fillColor: [30, 80, 140] },
    footStyles: { fillColor: [30, 80, 140] },
    columnStyles: { 3: { cellWidth: 65 } },
  });

  doc.save(`${project.name}_estimate.pdf`);
}
