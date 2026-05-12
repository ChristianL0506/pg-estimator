import jsPDF from "jspdf";
import autoTable from "jspdf-autotable";
import type { EstimateProject } from "@shared/schema";
import logoPic from "@assets/logo-pic.jpg";

/* ============================================================
   BID REPORT (customer-facing)
   ============================================================
   Goal: a clean proposal the customer sees. They get only what
   they need to make a decision:

     - Picou Group branding (logo)
     - Project info (name / no. / client / location / date)
     - High-level scope description (categories present, NOT
       item counts or line-level qty)
     - One bottom-line price (Grand Total Bid Amount)
     - Notes & assumptions
     - 30-day validity

   We deliberately DO NOT expose:
     - Material vs labor split
     - Labor hours
     - Per-category cost breakdown
     - Markup percentages
     - Per-item quantities
     - Internal cost detail

   Anything the client doesn't need to make the buy decision
   stays internal.
   ============================================================ */

function fmtDollar(n: number) {
  return "$" + n.toLocaleString("en-US", { minimumFractionDigits: 2, maximumFractionDigits: 2 });
}

// Friendly category labels for the scope description.
const CATEGORY_LABELS: Record<string, string> = {
  pipe: "Pipe",
  fitting: "Fittings",
  flange: "Flanges",
  valve: "Valves",
  gasket: "Gaskets",
  bolt: "Bolts & Studs",
  hanger: "Pipe Hangers & Supports",
  instrument: "Instrumentation",
  weld: "Field Welds",
  steel: "Structural Steel",
  rebar: "Reinforcing Steel",
  concrete: "Concrete",
  earthwork: "Earthwork",
  paving: "Paving",
  utility: "Underground Utilities",
};

function prettyCategory(raw: string): string {
  const key = raw.toLowerCase().trim();
  if (CATEGORY_LABELS[key]) return CATEGORY_LABELS[key];
  // Fall back to title-casing whatever the raw category is.
  return raw
    .toLowerCase()
    .split(/\s+/)
    .map(w => w.charAt(0).toUpperCase() + w.slice(1))
    .join(" ");
}

// Convert an imported image asset (from Vite, which gives us a URL)
// into a data-URI so jsPDF can embed it synchronously.
async function imageUrlToDataUri(url: string): Promise<string> {
  const res = await fetch(url);
  const blob = await res.blob();
  return await new Promise<string>((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve(reader.result as string);
    reader.onerror = reject;
    reader.readAsDataURL(blob);
  });
}

export async function generateBidReport(project: EstimateProject) {
  const doc = new jsPDF({ orientation: "portrait", unit: "mm", format: "letter" });
  const pageWidth = doc.internal.pageSize.getWidth();
  const pageHeight = doc.internal.pageSize.getHeight();
  const marginX = 18;
  const contentRight = pageWidth - marginX;

  // ---- 1. Header band with logo + branding ----
  let logoDataUri = "";
  try {
    logoDataUri = await imageUrlToDataUri(logoPic);
  } catch {
    // Logo fetch failed; we'll just skip it. The rest of the report
    // still works.
  }

  // Brand color band across the top
  doc.setFillColor(30, 80, 140);
  doc.rect(0, 0, pageWidth, 24, "F");

  // Logo on the left side of the band
  if (logoDataUri) {
    try {
      // Aspect ratio is preserved if width and height are both passed.
      // Source image is 363x219 ~ 1.66:1, so 18mm wide -> ~10.85mm tall.
      doc.addImage(logoDataUri, "JPEG", marginX, 4, 17, 16);
    } catch {
      // If addImage fails for any reason, fall through silently.
    }
  }

  // Branding text on the right of the band
  doc.setFont("helvetica", "bold");
  doc.setFontSize(16);
  doc.setTextColor(255, 255, 255);
  doc.text("PICOU GROUP CONTRACTORS", marginX + 22, 13);
  doc.setFont("helvetica", "normal");
  doc.setFontSize(9);
  doc.text("Industrial Piping & Mechanical Contractors", marginX + 22, 19);

  // ---- 2. Title ----
  doc.setFont("helvetica", "bold");
  doc.setFontSize(18);
  doc.setTextColor(30, 80, 140);
  doc.text("BID PROPOSAL", marginX, 36);

  // Underline accent
  doc.setDrawColor(30, 80, 140);
  doc.setLineWidth(0.6);
  doc.line(marginX, 38.5, marginX + 35, 38.5);

  // ---- 3. Project information block ----
  doc.setFont("helvetica", "bold");
  doc.setFontSize(10);
  doc.setTextColor(60, 60, 60);
  let y = 48;
  const label = (s: string) => { doc.setFont("helvetica", "bold"); doc.text(s, marginX, y); };
  const value = (s: string) => { doc.setFont("helvetica", "normal"); doc.text(s, marginX + 30, y); };

  label("Project:"); value(project.name); y += 6;
  if (project.projectNumber) { label("Project No:"); value(project.projectNumber); y += 6; }
  if (project.client) { label("Prepared For:"); value(project.client); y += 6; }
  if (project.location) { label("Location:"); value(project.location); y += 6; }
  label("Proposal Date:"); value(new Date().toLocaleDateString("en-US", { year: "numeric", month: "long", day: "numeric" })); y += 6;
  label("Valid Through:"); value(new Date(Date.now() + 30 * 24 * 60 * 60 * 1000).toLocaleDateString("en-US", { year: "numeric", month: "long", day: "numeric" })); y += 12;

  // ---- 4. Scope of Work (high level, NO quantities) ----
  doc.setFont("helvetica", "bold");
  doc.setFontSize(12);
  doc.setTextColor(30, 80, 140);
  doc.text("SCOPE OF WORK", marginX, y); y += 2;
  doc.setDrawColor(30, 80, 140);
  doc.setLineWidth(0.3);
  doc.line(marginX, y, contentRight, y); y += 6;

  doc.setFont("helvetica", "normal");
  doc.setFontSize(10);
  doc.setTextColor(50, 50, 50);

  // Gather unique categories present in the BOM. We expose ONLY
  // the categories — never counts or quantities.
  const catSet = new Set<string>();
  for (const item of project.items) {
    if (item.category) catSet.add(item.category.toLowerCase().trim());
  }
  const categories = Array.from(catSet)
    .map(prettyCategory)
    .sort();

  const scopeIntro = "Picou Group Contractors is pleased to submit this bid proposal for the above-referenced project. Our scope includes furnishing all labor, supervision, tools, equipment, and materials required to complete the following work:";
  const wrappedIntro = doc.splitTextToSize(scopeIntro, contentRight - marginX);
  doc.text(wrappedIntro, marginX, y);
  y += wrappedIntro.length * 5 + 4;

  // Render categories as a bulleted list (no qty, no $ exposure)
  if (categories.length > 0) {
    for (const cat of categories) {
      doc.text(`\u2022 ${cat}`, marginX + 4, y);
      y += 5.5;
    }
  } else {
    doc.text("\u2022 Scope per attached drawings and specifications.", marginX + 4, y);
    y += 5.5;
  }
  y += 6;

  // ---- 5. Total Bid Amount (the headline number, single value) ----
  // We compute the same grand total as the internal estimate, but
  // we present ONLY the total \u2014 not the breakdown.
  const totalMaterial = project.items.reduce((sum, i) => sum + (i.materialExtension || 0), 0);
  const totalLabor = project.items.reduce((sum, i) => sum + (i.laborExtension || 0), 0);
  const subtotal = totalMaterial + totalLabor;
  const overheadAmt = subtotal * ((project.markups?.overhead || 0) / 100);
  const profitAmt = (subtotal + overheadAmt) * ((project.markups?.profit || 0) / 100);
  const taxAmt = totalMaterial * ((project.markups?.tax || 0) / 100);
  const bondAmt = (subtotal + overheadAmt + profitAmt + taxAmt) * ((project.markups?.bond || 0) / 100);
  const grandTotal = subtotal + overheadAmt + profitAmt + taxAmt + bondAmt;

  if (y > pageHeight - 80) { doc.addPage(); y = 24; }

  // Headline price box
  doc.setFillColor(30, 80, 140);
  doc.rect(marginX, y, contentRight - marginX, 22, "F");
  doc.setFont("helvetica", "bold");
  doc.setFontSize(11);
  doc.setTextColor(220, 230, 245);
  doc.text("TOTAL BID AMOUNT", marginX + 6, y + 9);
  doc.setFontSize(20);
  doc.setTextColor(255, 255, 255);
  const priceText = fmtDollar(grandTotal);
  const priceWidth = doc.getTextWidth(priceText);
  doc.text(priceText, contentRight - 6 - priceWidth, y + 15);
  y += 30;

  // Sub-note: lump-sum / inclusive language
  doc.setFont("helvetica", "italic");
  doc.setFontSize(9);
  doc.setTextColor(80, 80, 80);
  const lumpNote = "Lump-sum price, inclusive of all labor, material, overhead, applicable sales tax, and bonding. No additional charges apply unless the scope changes.";
  const wrappedLump = doc.splitTextToSize(lumpNote, contentRight - marginX);
  doc.text(wrappedLump, marginX, y);
  y += wrappedLump.length * 4.5 + 6;

  // ---- 6. Notes & Assumptions ----
  if (y > pageHeight - 70) { doc.addPage(); y = 24; }

  doc.setFont("helvetica", "bold");
  doc.setFontSize(12);
  doc.setTextColor(30, 80, 140);
  doc.text("NOTES & ASSUMPTIONS", marginX, y); y += 2;
  doc.setDrawColor(30, 80, 140);
  doc.setLineWidth(0.3);
  doc.line(marginX, y, contentRight, y); y += 6;

  doc.setFont("helvetica", "normal");
  doc.setFontSize(9);
  doc.setTextColor(60, 60, 60);

  const assumptions = [
    "Pricing based on drawings and specifications provided at time of bid.",
    "Work performed during normal business hours (Monday\u2013Friday, daytime shift).",
    "Excludes scaffolding, insulation, painting, fireproofing, and X-ray inspection unless specifically listed in scope.",
    "Excludes confined-space entry, hot-work permits, hazardous material handling, and asbestos abatement.",
    "Owner to provide site access, laydown area, water, power, and sanitary facilities at no cost to contractor.",
    "Change orders required for scope additions, revisions, or field-discovered conditions.",
    "Pricing valid for 30 days from proposal date.",
    "Standard payment terms: net 30, monthly progress billing.",
  ];
  for (const note of assumptions) {
    const wrapped = doc.splitTextToSize(`\u2022 ${note}`, contentRight - marginX - 4);
    doc.text(wrapped, marginX + 4, y);
    y += wrapped.length * 4.5 + 1.5;
  }
  y += 8;

  // ---- 7. Signature block ----
  if (y > pageHeight - 50) { doc.addPage(); y = 24; }

  doc.setFont("helvetica", "bold");
  doc.setFontSize(10);
  doc.setTextColor(30, 80, 140);
  doc.text("ACCEPTANCE", marginX, y); y += 8;

  doc.setFont("helvetica", "normal");
  doc.setFontSize(9);
  doc.setTextColor(60, 60, 60);
  doc.text("Accepted and authorized to proceed on the terms above:", marginX, y); y += 14;

  // Two signature lines side-by-side
  const col2X = marginX + (contentRight - marginX) / 2 + 4;
  doc.setDrawColor(120, 120, 120);
  doc.setLineWidth(0.3);
  doc.line(marginX, y, marginX + (contentRight - marginX) / 2 - 4, y);
  doc.line(col2X, y, contentRight, y);
  doc.setFontSize(8);
  doc.setTextColor(100, 100, 100);
  doc.text("Customer Signature", marginX, y + 4);
  doc.text("Date", col2X, y + 4);
  y += 14;

  doc.line(marginX, y, marginX + (contentRight - marginX) / 2 - 4, y);
  doc.line(col2X, y, contentRight, y);
  doc.text("Printed Name", marginX, y + 4);
  doc.text("Title", col2X, y + 4);

  // ---- 8. Footer on every page ----
  const pageCount = doc.getNumberOfPages();
  for (let i = 1; i <= pageCount; i++) {
    doc.setPage(i);
    doc.setFont("helvetica", "normal");
    doc.setFontSize(7);
    doc.setTextColor(150, 150, 150);
    doc.text(
      `Picou Group Contractors  \u2014  ${project.name}  \u2014  Page ${i} of ${pageCount}`,
      marginX,
      pageHeight - 8,
    );
    doc.text(
      "Confidential. Pricing valid for 30 days.",
      contentRight,
      pageHeight - 8,
      { align: "right" },
    );
  }

  doc.save(`${project.name} - Bid Proposal.pdf`);
}
