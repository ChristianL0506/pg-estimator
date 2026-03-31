import jsPDF from "jspdf";
import autoTable from "jspdf-autotable";
import type { EstimateProject } from "@shared/schema";

function fmtDollar(n: number) {
  return "$" + n.toLocaleString("en-US", { minimumFractionDigits: 2, maximumFractionDigits: 2 });
}

export function generateBidReport(project: EstimateProject) {
  const doc = new jsPDF({ orientation: "portrait" });

  // 1. Header
  doc.setFontSize(20);
  doc.setTextColor(30, 80, 140);
  doc.text("PICOU GROUP CONTRACTORS", 14, 22);
  doc.setFontSize(10);
  doc.setTextColor(100, 100, 100);
  doc.text("Industrial Piping & Mechanical Contractors", 14, 29);
  doc.setDrawColor(30, 80, 140);
  doc.setLineWidth(0.8);
  doc.line(14, 33, 196, 33);

  doc.setFontSize(14);
  doc.setTextColor(30, 30, 30);
  doc.text("BID PROPOSAL", 14, 44);

  doc.setFontSize(10);
  doc.setTextColor(60, 60, 60);
  let y = 54;
  doc.text(`Project: ${project.name}`, 14, y); y += 7;
  if (project.projectNumber) { doc.text(`Project No: ${project.projectNumber}`, 14, y); y += 7; }
  if (project.client) { doc.text(`Client: ${project.client}`, 14, y); y += 7; }
  if (project.location) { doc.text(`Location: ${project.location}`, 14, y); y += 7; }
  doc.text(`Date: ${new Date().toLocaleDateString()}`, 14, y); y += 12;

  // 2. Scope summary
  doc.setFontSize(11);
  doc.setTextColor(30, 80, 140);
  doc.text("SCOPE SUMMARY", 14, y); y += 7;
  doc.setFontSize(9);
  doc.setTextColor(60, 60, 60);

  const catCounts: Record<string, { count: number; qty: number }> = {};
  for (const item of project.items) {
    const cat = item.category.toUpperCase();
    if (!catCounts[cat]) catCounts[cat] = { count: 0, qty: 0 };
    catCounts[cat].count++;
    catCounts[cat].qty += item.quantity;
  }

  doc.text(`Total Line Items: ${project.items.length}`, 14, y); y += 5;
  const catList = Object.entries(catCounts).sort((a, b) => b[1].count - a[1].count);
  for (const [cat, info] of catList.slice(0, 8)) {
    doc.text(`  ${cat}: ${info.count} items`, 14, y); y += 5;
  }
  y += 5;

  // 3. Cost summary table
  const totalMaterial = project.items.reduce((sum, i) => sum + (i.materialExtension || 0), 0);
  const totalLabor = project.items.reduce((sum, i) => sum + (i.laborExtension || 0), 0);
  const totalHours = project.items.reduce((sum, i) => sum + (i.quantity || 0) * (i.laborHoursPerUnit || 0), 0);
  const subtotal = totalMaterial + totalLabor;
  const overheadAmt = subtotal * (project.markups.overhead / 100);
  const profitAmt = (subtotal + overheadAmt) * (project.markups.profit / 100);
  const taxAmt = totalMaterial * (project.markups.tax / 100);
  const bondAmt = (subtotal + overheadAmt + profitAmt + taxAmt) * (project.markups.bond / 100);
  const grandTotal = subtotal + overheadAmt + profitAmt + taxAmt + bondAmt;

  doc.setFontSize(11);
  doc.setTextColor(30, 80, 140);
  doc.text("COST SUMMARY", 14, y); y += 3;

  autoTable(doc, {
    startY: y,
    head: [["Item", "Amount"]],
    body: [
      ["Material Cost", fmtDollar(totalMaterial)],
      ["Labor Cost", fmtDollar(totalLabor)],
      ["Total Labor Hours", totalHours.toFixed(1) + " hrs"],
      ["Subtotal", fmtDollar(subtotal)],
      [`Overhead (${project.markups.overhead}%)`, fmtDollar(overheadAmt)],
      [`Profit (${project.markups.profit}%)`, fmtDollar(profitAmt)],
      [`Sales Tax (${project.markups.tax}%)`, fmtDollar(taxAmt)],
      [`Bond (${project.markups.bond}%)`, fmtDollar(bondAmt)],
    ],
    foot: [["GRAND TOTAL", fmtDollar(grandTotal)]],
    styles: { fontSize: 9, cellPadding: 3 },
    headStyles: { fillColor: [30, 80, 140] },
    footStyles: { fillColor: [30, 80, 140], fontStyle: "bold" },
    columnStyles: { 1: { halign: "right" } },
    margin: { left: 14, right: 14 },
  });

  y = (doc as any).lastAutoTable.finalY + 10;

  // 4. BOM summary by category (pivot-style)
  if (y > 240) { doc.addPage(); y = 20; }
  doc.setFontSize(11);
  doc.setTextColor(30, 80, 140);
  doc.text("BOM SUMMARY BY CATEGORY", 14, y); y += 3;

  const catSummary: Record<string, { qty: number; matCost: number; labCost: number }> = {};
  for (const item of project.items) {
    const cat = item.category.toUpperCase();
    if (!catSummary[cat]) catSummary[cat] = { qty: 0, matCost: 0, labCost: 0 };
    catSummary[cat].qty += item.quantity;
    catSummary[cat].matCost += item.materialExtension || 0;
    catSummary[cat].labCost += item.laborExtension || 0;
  }

  const bomRows = Object.entries(catSummary)
    .sort((a, b) => (b[1].matCost + b[1].labCost) - (a[1].matCost + a[1].labCost))
    .map(([cat, info]) => [cat, String(Math.round(info.qty)), fmtDollar(info.matCost), fmtDollar(info.labCost), fmtDollar(info.matCost + info.labCost)]);

  autoTable(doc, {
    startY: y,
    head: [["Category", "Qty", "Material Cost", "Labor Cost", "Total"]],
    body: bomRows,
    styles: { fontSize: 8, cellPadding: 2 },
    headStyles: { fillColor: [30, 80, 140] },
    columnStyles: { 1: { halign: "right" }, 2: { halign: "right" }, 3: { halign: "right" }, 4: { halign: "right" } },
    margin: { left: 14, right: 14 },
  });

  y = (doc as any).lastAutoTable.finalY + 10;

  // 5. Notes/assumptions
  if (y > 250) { doc.addPage(); y = 20; }
  doc.setFontSize(11);
  doc.setTextColor(30, 80, 140);
  doc.text("NOTES & ASSUMPTIONS", 14, y); y += 7;
  doc.setFontSize(8);
  doc.setTextColor(80, 80, 80);
  const assumptions = [
    "All pricing based on current material costs and labor rates.",
    `Base labor rate: $${project.laborRate}/hr.`,
    "Excludes scaffolding, insulation, and painting unless specifically listed.",
    "Assumes normal working conditions (no confined space, no hazmat).",
    "Valid for 30 days from date of proposal.",
  ];
  for (const note of assumptions) {
    doc.text(`• ${note}`, 14, y); y += 5;
  }

  // Footer on all pages
  const pageCount = doc.getNumberOfPages();
  for (let i = 1; i <= pageCount; i++) {
    doc.setPage(i);
    doc.setFontSize(7);
    doc.setTextColor(150, 150, 150);
    doc.text(`Picou Group Contractors — ${project.name} — Page ${i} of ${pageCount}`, 14, 287);
  }

  doc.save(`${project.name} - Bid Report.pdf`);
}
