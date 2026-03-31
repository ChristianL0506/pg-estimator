// Quick Entry Parser for natural language estimating input
// Examples:
//   "3 4\" butt welds" → { qty: 3, size: "4\"", description: "butt welds", unit: "EA" }
//   "100 LF 6\" pipe" → { qty: 100, size: "6\"", description: "pipe", unit: "LF" }
//   "50 CY concrete" → { qty: 50, size: "", description: "concrete", unit: "CY" }

export interface ParsedEntry {
  qty: number;
  size: string;
  description: string;
  unit: string;
  category: "pipe" | "fitting" | "valve" | "steel" | "concrete" | "rebar" | "earthwork" | "paving" | "electrical" | "other";
}

const UNIT_KEYWORDS = /\b(LF|SF|CY|SY|EA|TON|TONS|LBS|LB|AC|FT|GAL)\b/i;

const SIZE_PATTERNS = [
  /\b(\d+(?:-\d+\/\d+)?")\s*(?:x\s*\d+(?:-\d+\/\d+)?")?/,  // 4", 1-1/2", 6"x4"
  /\b(\d+(?:\/\d+)?)\s*IN(?:CH)?/i,                           // 4 INCH
  /\b(\d+(?:-\d+\/\d+)?\s*INCH)/i,                            // 4-inch
  /\b(#\d+)\b/,                                                // #5 rebar
  /\b(W\d+x\d+)\b/i,                                          // W14x30 steel
  /\b(HSS\s*\d+x\d+)/i,                                       // HSS 6x6
];

function detectCategory(desc: string): ParsedEntry["category"] {
  const d = desc.toUpperCase();
  if (/\bPIPE\b/.test(d)) return "pipe";
  if (/\bELBOW|TEE|REDUCER|COUPLING|FLANGE|GASKET|FITTING|NIPPLE|CAP\b/.test(d)) return "fitting";
  if (/\bVALVE\b/.test(d)) return "valve";
  if (/\bW\d+X|HSS|BEAM|COLUMN|BRACE|ANGLE|CHANNEL|PLATE\b/.test(d)) return "steel";
  if (/\bCONCRETE|FOOTING|SLAB|WALL|GRADE BEAM\b/.test(d)) return "concrete";
  if (/\bREBAR|#\d+\b/.test(d)) return "rebar";
  if (/\bCUT|FILL|EARTHWORK|EXCAVAT|BACKFILL\b/.test(d)) return "earthwork";
  if (/\bASPHALT|PAVING|CURB|SIDEWALK|BASE COURSE\b/.test(d)) return "paving";
  if (/\bELECTRICAL|CONDUIT|WIRE\b/.test(d)) return "electrical";
  return "other";
}

export function parseQuickEntry(input: string): ParsedEntry | null {
  const trimmed = input.trim();
  if (!trimmed) return null;

  // Try to extract leading quantity
  const qtyMatch = trimmed.match(/^(\d+(?:\.\d+)?)\s+(.+)/);
  if (!qtyMatch) return null;

  const qty = parseFloat(qtyMatch[1]);
  let remainder = qtyMatch[2].trim();

  // Extract unit if present (right after qty or before description)
  let unit = "EA";
  const unitMatch = remainder.match(UNIT_KEYWORDS);
  if (unitMatch) {
    const matchedUnit = unitMatch[1].toUpperCase();
    // Normalize
    if (matchedUnit === "FT") unit = "LF";
    else if (matchedUnit === "TONS") unit = "TON";
    else if (matchedUnit === "LB") unit = "LBS";
    else unit = matchedUnit;
    remainder = remainder.replace(UNIT_KEYWORDS, "").trim();
  }

  // Extract size
  let size = "";
  for (const pattern of SIZE_PATTERNS) {
    const sizeMatch = remainder.match(pattern);
    if (sizeMatch) {
      size = sizeMatch[0].trim();
      remainder = remainder.replace(sizeMatch[0], "").trim();
      break;
    }
  }

  // Clean up remaining text as description
  const description = remainder.replace(/\s+/g, " ").trim();
  if (!description) return null;

  // Infer unit from description keywords if still EA
  if (unit === "EA") {
    const desc = description.toUpperCase();
    if (/\bPIPE\b/.test(desc)) unit = "LF";
    else if (/\bCONCRETE\b|\bEXCAVAT\b|\bFILL\b/.test(desc)) unit = "CY";
    else if (/\bASPHALT\b|\bPAVING\b|\bSLAB\b/.test(desc)) unit = "SF";
  }

  return {
    qty,
    size,
    description,
    unit,
    category: detectCategory(description + " " + size),
  };
}
