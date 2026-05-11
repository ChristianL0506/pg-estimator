import type { Express, Request, Response, NextFunction } from "express";
import { createServer, type Server } from "http";
import { storage } from "./storage";
import multer from "multer";
import { randomUUID } from "crypto";
import fs from "fs";
import path from "path";
import { execSync, execFileSync, execFile, exec } from "child_process";
import Anthropic from "@anthropic-ai/sdk";
import type { JobProgress } from "@shared/schema";
import { insertEstimateProjectSchema, insertCostDatabaseEntrySchema, estimateItemSchema, markupsSchema, insertPurchaseRecordSchema, insertBidSchema, insertVendorQuoteSchema, insertCustomEstimatorMethodSchema } from "@shared/schema";
import { z } from "zod";
import { parse as csvParse } from "csv-parse/sync";
import { PDFDocument } from "pdf-lib";

// Login rate limiting: 5 failed attempts per IP in 15 min window
const loginAttempts = new Map<string, { count: number; firstAttempt: number; blocked: boolean }>();
const LOGIN_MAX_ATTEMPTS = 5;
const LOGIN_WINDOW_MS = 15 * 60 * 1000; // 15 minutes

// Factory function to get a fresh Anthropic client
// Re-reads env var each time in case token was refreshed
// User-configured API key (stored in SQLite, overrides env var)
let userApiKey: string | null = null;
let userGeminiKey: string | null = null;

function setUserApiKey(key: string | null) {
  userApiKey = key;
}

function getUserApiKey(): string | null {
  return userApiKey;
}

function setUserGeminiKey(key: string | null) {
  userGeminiKey = key;
}

function getUserGeminiKey(): string | null {
  return userGeminiKey;
}

function getAnthropicClient(): Anthropic {
  const key = userApiKey || process.env.ANTHROPIC_API_KEY || process.env.LLM_API_KEY || "";
  // If user has their own key, use Anthropic DIRECTLY (bypass any proxy)
  if (userApiKey) {
    return new Anthropic({ apiKey: key, baseURL: "https://api.anthropic.com" });
  }
  // Otherwise use platform-provided key (may go through proxy)
  return new Anthropic({
    apiKey: key,
  });
}

/** Check if an error is an authentication/token-expiry error */
function isAuthError(err: any): boolean {
  const msg = String(err?.message || err?.error?.message || "").toLowerCase();
  const status = err?.status || err?.statusCode || 0;
  return (
    status === 401 ||
    msg.includes("authentication_error") ||
    msg.includes("invalid or expired session token") ||
    msg.includes("invalid x-api-key") ||
    msg.includes("invalid api key") ||
    msg.includes("401")
  );
}

const autoCalculateBodySchema = z.object({
  method: z.enum(["bill", "justin", "industry"]).default("justin"),
  // Optional: ID of a saved CustomEstimatorMethod. When provided, the base
  // method (above) is used with this profile's overrides layered on top.
  customMethodId: z.string().optional(),
  laborRate: z.number().min(0).max(500).default(56),
  overtimeRate: z.number().min(0).max(500).default(79),
  doubleTimeRate: z.number().min(0).max(500).default(100),
  perDiem: z.number().min(0).max(500).default(75),
  overtimePercent: z.number().min(0).max(100).default(15),
  doubleTimePercent: z.number().min(0).max(100).default(2),
  material: z.enum(["CS", "SS"]).default("CS"),
  schedule: z.string().default("STD"),
  installType: z.enum(["standard", "rack"]).default("standard"),
  pipeLocation: z.string().default("Open Rack"),
  elevation: z.string().default("0-20ft"),
  alloyGroup: z.string().default("4"),
  rackFactor: z.number().min(1).max(5).default(1.3),
  // How to treat fitting labor vs separate weld rows.
  // "bundled" (default): fitting MH includes its weld labor via the weld-end multiplier.
  // "separate": fitting MH is handling only; weld rows in the BOM carry the weld labor.
  fittingWeldMode: z.enum(["bundled", "separate"]).default("bundled"),
}).refine(data => data.overtimePercent + data.doubleTimePercent <= 100, {
  message: "overtimePercent + doubleTimePercent cannot exceed 100",
  path: ["overtimePercent"],
});

const patchEstimateSchema = z.object({
  name: z.string().optional(),
  projectNumber: z.string().optional(),
  client: z.string().optional(),
  location: z.string().optional(),
  markups: markupsSchema.optional(),
  items: z.array(estimateItemSchema).optional(),
  laborRate: z.number().optional(),
  overtimeRate: z.number().optional(),
  doubleTimeRate: z.number().optional(),
  perDiem: z.number().optional(),
  overtimePercent: z.number().optional(),
  doubleTimePercent: z.number().optional(),
  estimateMethod: z.enum(["bill", "justin", "industry", "manual"]).optional(),
  customMethodId: z.string().optional(),
  fittingWeldMode: z.enum(["bundled", "separate"]).optional(),
  // Pass a number to override, null to clear, undefined to leave unchanged.
  // Percentage value (e.g. 15 = 15%). Bill method ignores this.
  contingencyOverride: z.number().nullable().optional(),
  scopeAdders: z.array(z.object({
    id: z.string(),
    label: z.string(),
    mode: z.enum(["hours", "cost"]).default("hours"),
    hours: z.number().default(0),
    ratePerHour: z.number().optional(),
    flatCost: z.number().default(0),
    note: z.string().optional(),
  })).optional(),
});
import { generateBillsWorkbook, generateJustinsWorkbook } from "./excelExport";
import { generateMethodFactorsWorkbook, generateCompareWorkbook } from "./methodExports";

// ============================================================
// AUTH MIDDLEWARE
// ============================================================

function authMiddleware(req: Request, res: Response, next: NextFunction): void {
  // Skip auth for login endpoint
  if (req.path === "/api/login") return next();
  // Skip non-API routes
  if (!req.path.startsWith("/api/")) return next();

  const authHeader = req.headers.authorization;
  if (!authHeader || !authHeader.startsWith("Bearer ")) {
    res.status(401).json({ message: "Authentication required" });
    return;
  }
  const token = authHeader.slice(7);
  const session = storage.validateToken(token);
  if (!session) {
    res.status(401).json({ message: "Invalid or expired token" });
    return;
  }
  (req as any).user = session;
  next();
}

const MAX_FILE_SIZE = 300 * 1024 * 1024; // 300MB — supports large drawing packages (500+ pages)
const UPLOAD_DIR = "/tmp/pg-unified-uploads";
const RENDER_DIR = "/tmp/pg-unified-renders";
const CHUNK_SIZE = 40;
const LARGE_PACKAGE_THRESHOLD = 100; // pages — switch to streaming mode above this
const LARGE_CHUNK_SIZE = 10; // much smaller chunks for large packages — only 10 images in memory at a time

for (const dir of [UPLOAD_DIR, RENDER_DIR]) {
  if (!fs.existsSync(dir)) fs.mkdirSync(dir, { recursive: true });
}

const upload = multer({
  storage: multer.diskStorage({
    destination: UPLOAD_DIR,
    filename: (_req, file, cb) => {
      const safeName = file.originalname.replace(/[^a-zA-Z0-9._-]/g, "_");
      cb(null, `${Date.now()}_${safeName}`);
    },
  }),
  limits: { fileSize: MAX_FILE_SIZE },
  fileFilter: (_req, file, cb) => {
    if (file.mimetype === "application/pdf") cb(null, true);
    else cb(new Error("Only PDF files are allowed."));
  },
});

// ============================================================
// HELPER FUNCTIONS
// ============================================================

function getPdfPageCount(pdfPath: string): number {
  try {
    const info = execFileSync("pdfinfo", [pdfPath], { timeout: 10000 }).toString();
    const m = info.match(/Pages:\s*(\d+)/);
    if (m) return parseInt(m[1], 10);
  } catch (e) { console.warn("Suppressed error:", e); }
  return 1;
}

function extractText(pdfPath: string): string {
  try {
    return execFileSync("pdftotext", ["-layout", pdfPath, "-"], {
      maxBuffer: 50 * 1024 * 1024, timeout: 60000,
    }).toString("utf-8");
  } catch (e) {
    console.warn("Suppressed error:", e);
    return "";
  }
}

// Per-page PDF text extraction. Returns the embedded vector text for one page,
// or empty string if the PDF is image-only (scanned) or pdftotext is missing.
// This is the lowest-risk implementation of the council's PDF text suggestion:
// when CAD-exported ISOs are loaded, every BOM cell is in the PDF as text and
// pdftotext gives us the rows for free — we then pass it to Claude as evidence.
// For scanned PDFs this returns "" and behavior is unchanged.
function extractPageText(pdfPath: string, pageNum: number): string {
  try {
    return execFileSync("pdftotext", [
      "-layout",
      "-f", String(pageNum),
      "-l", String(pageNum),
      pdfPath, "-",
    ], { maxBuffer: 8 * 1024 * 1024, timeout: 15000 }).toString("utf-8");
  } catch (e) {
    return "";
  }
}

// Probes a PDF to see whether it has embedded vector text. Returns true if
// the first 3 pages average >100 chars of text (typical CAD-exported ISO).
// Used to decide whether to pay the per-page extractPageText cost on every
// page or skip it entirely.
function pdfHasVectorText(pdfPath: string, pageCount: number): boolean {
  try {
    const samplePages = Math.min(3, pageCount);
    let totalChars = 0;
    for (let p = 1; p <= samplePages; p++) {
      totalChars += extractPageText(pdfPath, p).length;
    }
    return (totalChars / samplePages) > 100;
  } catch {
    return false;
  }
}

function cleanupJobDir(jobDir: string) {
  try {
    const files = fs.readdirSync(jobDir);
    for (const f of files) {
      try { fs.unlinkSync(path.join(jobDir, f)); } catch (e) { console.warn("Suppressed error:", e); }
    }
    fs.rmSync(jobDir, { recursive: true, force: true });
  } catch (e) { console.warn("Suppressed error:", e); }
}

const THUMBNAIL_DIR = path.join(process.cwd(), "data", "page-thumbnails");
if (!fs.existsSync(THUMBNAIL_DIR)) fs.mkdirSync(THUMBNAIL_DIR, { recursive: true });

function savePageThumbnails(
  projectId: string,
  rendered: any,
  startPage: number,
) {
  try {
    const projDir = path.join(THUMBNAIL_DIR, projectId);
    if (!fs.existsSync(projDir)) fs.mkdirSync(projDir, { recursive: true });

    // Mechanical with revisions: has cloudPageImages with fullImagePath
    const cloudImages = rendered.cloudPageImages as { pageNum: number; fullImagePath: string }[] | undefined;
    if (cloudImages && cloudImages.length > 0) {
      for (const pi of cloudImages) {
        const globalPage = pi.pageNum + startPage - 1;
        const dest = path.join(projDir, `page-${globalPage}.png`);
        if (fs.existsSync(pi.fullImagePath)) {
          fs.copyFileSync(pi.fullImagePath, dest);
        }
      }
      return;
    }

    // Mechanical without revisions: has bomPageImages with imagePath
    const bomImages = rendered.bomPageImages as { pageNum: number; imagePath: string }[] | undefined;
    if (bomImages && bomImages.length > 0) {
      for (const pi of bomImages) {
        const globalPage = pi.pageNum + startPage - 1;
        const dest = path.join(projDir, `page-${globalPage}.png`);
        if (fs.existsSync(pi.imagePath)) {
          fs.copyFileSync(pi.imagePath, dest);
        }
      }
      return;
    }

    // Structural/Civil: has pageImages with imagePath
    const pageImages = rendered.pageImages as { pageNum: number; imagePath: string }[] | undefined;
    if (pageImages && pageImages.length > 0) {
      for (const pi of pageImages) {
        const globalPage = pi.pageNum + startPage - 1;
        const dest = path.join(projDir, `page-${globalPage}.png`);
        if (fs.existsSync(pi.imagePath)) {
          fs.copyFileSync(pi.imagePath, dest);
        }
      }
    }
  } catch (err: any) {
    console.warn(`  Failed to save thumbnails for project ${projectId}:`, err.message);
  }
}

// ============================================================
// DUAL-MODEL EXTRACTION — Gemini (Feature 1)
// ============================================================

async function extractWithGemini(imageBase64: string, prompt: string): Promise<any[]> {
  const geminiKey = getUserGeminiKey();
  if (!geminiKey) return [];
  try {
    const resp = await fetch(
      `https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent`,
      {
        method: "POST",
        headers: { "Content-Type": "application/json", "x-goog-api-key": geminiKey },
        body: JSON.stringify({
          contents: [{
            parts: [
              { text: prompt },
              { inlineData: { mimeType: "image/png", data: imageBase64 } },
            ],
          }],
          generationConfig: { temperature: 0, maxOutputTokens: 8192 },
        }),
      }
    );
    if (!resp.ok) {
      console.warn(`  Gemini API error: ${resp.status} ${resp.statusText}`);
      return [];
    }
    const data = await resp.json();
    const text = data?.candidates?.[0]?.content?.parts?.[0]?.text || "";
    // Try to parse JSON from the response
    const jsonMatch = text.match(/\[[\s\S]*\]/);
    if (jsonMatch) {
      return JSON.parse(jsonMatch[0]);
    }
  } catch (err: any) {
    console.warn(`  Gemini extraction failed: ${err.message}`);
  }
  return [];
}

// ============================================================
// DRAWING TEMPLATE DETECTION (Feature 2)
// ============================================================

function detectDrawingTemplate(ocrText: string): any | null {
  if (!ocrText || ocrText.trim().length < 50) return null;
  const templates = storage.getDrawingTemplates();
  if (templates.length === 0) return null;

  const textLower = ocrText.toLowerCase();
  let bestMatch: any = null;
  let bestScore = 0;

  for (const tmpl of templates) {
    let score = 0;
    // Check match patterns (JSON array of regex strings)
    if (tmpl.matchPatterns) {
      try {
        const patterns: string[] = JSON.parse(tmpl.matchPatterns);
        for (const pat of patterns) {
          try {
            if (new RegExp(pat, "i").test(ocrText)) score += 10;
          } catch (e) { console.warn("Suppressed error:", e); }
        }
      } catch (e) { console.warn("Suppressed error:", e); }
    }
    // Check engineering firm name
    if (tmpl.engineeringFirm && textLower.includes(tmpl.engineeringFirm.toLowerCase())) {
      score += 20;
    }
    // Check sample OCR text similarity (simple word overlap)
    if (tmpl.sampleOcrText) {
      const sampleWords = new Set(tmpl.sampleOcrText.toLowerCase().split(/\s+/).filter((w: string) => w.length > 3));
      const textWords = textLower.split(/\s+/).filter(w => w.length > 3);
      let overlap = 0;
      for (const w of textWords) {
        if (sampleWords.has(w)) overlap++;
      }
      if (sampleWords.size > 0) score += (overlap / sampleWords.size) * 15;
    }
    if (score > bestScore) {
      bestScore = score;
      bestMatch = tmpl;
    }
  }

  return bestScore >= 10 ? bestMatch : null;
}

async function splitPdfIntoChunks(pdfPath: string, pageCount: number): Promise<{ chunkPath: string; startPage: number; endPage: number }[]> {
  const effectiveChunkSize = pageCount > LARGE_PACKAGE_THRESHOLD ? LARGE_CHUNK_SIZE : CHUNK_SIZE;
  if (pageCount <= effectiveChunkSize) {
    return [{ chunkPath: pdfPath, startPage: 1, endPage: pageCount }];
  }
  const chunks: { chunkPath: string; startPage: number; endPage: number }[] = [];
  const chunkDir = path.join(RENDER_DIR, `chunks_${Date.now()}`);
  fs.mkdirSync(chunkDir, { recursive: true });

  // Use pdf-lib (pure Node.js) instead of qpdf system dependency
  const pdfBytes = fs.readFileSync(pdfPath);
  const srcDoc = await PDFDocument.load(pdfBytes);

  for (let start = 1; start <= pageCount; start += effectiveChunkSize) {
    const end = Math.min(start + effectiveChunkSize - 1, pageCount);
    const chunkPath = path.join(chunkDir, `chunk_${start}_${end}.pdf`);
    try {
      const chunkDoc = await PDFDocument.create();
      const pageIndices = Array.from({ length: end - start + 1 }, (_, i) => start - 1 + i);
      const copiedPages = await chunkDoc.copyPages(srcDoc, pageIndices);
      for (const page of copiedPages) {
        chunkDoc.addPage(page);
      }
      const chunkBytes = await chunkDoc.save();
      fs.writeFileSync(chunkPath, chunkBytes);
      chunks.push({ chunkPath, startPage: start, endPage: end });
    } catch (err: any) {
      console.error(`  Failed to split pages ${start}-${end}:`, err.message);
    }
  }
  return chunks;
}

function parsePipeLength(raw: string): { feet: number; original: string; wasInches?: boolean } {
  const s = String(raw).trim();
  // Match feet-inches: 16'-8", 3'-0", 0'-11", etc.
  const m = s.match(/^(\d+)['''\u2019]\s*[-–]?\s*(\d+)?(?:\s*[-–]?\s*(\d+)\/(\d+))?\s*[""\u201D]?$/);
  if (m) {
    const ft = parseInt(m[1], 10) || 0;
    const inches = parseInt(m[2] || "0", 10);
    const fracNum = parseInt(m[3] || "0", 10);
    const fracDen = parseInt(m[4] || "1", 10) || 1;
    const totalInches = ft * 12 + inches + fracNum / fracDen;
    return { feet: Math.round((totalInches / 12) * 100) / 100, original: s };
  }
  // Match inches-only: 11", 6", 8-1/2", etc. (double-quote means inches)
  const inchMatch = s.match(/^(\d+)(?:\s*[-–]?\s*(\d+)\/(\d+))?\s*[""\u201D]$/);
  if (inchMatch) {
    const inches = parseInt(inchMatch[1], 10) || 0;
    const fracNum = parseInt(inchMatch[2] || "0", 10);
    const fracDen = parseInt(inchMatch[3] || "1", 10) || 1;
    const totalInches = inches + fracNum / fracDen;
    return { feet: Math.round((totalInches / 12) * 100) / 100, original: s, wasInches: true };
  }
  const num = parseFloat(s);
  return { feet: isNaN(num) ? 0 : num, original: s };
}

/**
 * Detect and correct pipe quantities where inches were misread as feet.
 * On ISOs, short pipe runs (spool pieces, drops, branch connections) are common.
 * When the AI extracts just "11" without a unit marker, it's ambiguous.
 *
 * Heuristic based on Stolthaven Phase 6 calibration data:
 * - Values 1-11 without unit marker: auto-correct to inches (very common spool lengths)
 * - Values 12-18: auto-correct for small bore (<=2"), flag for larger pipe
 * - Values 19+: assume feet (longer runs are typically in feet)
 */
function correctPipeLengthIfInches(qty: number, rawQty: string, description: string, options?: { installLocation?: string; size?: string }): { correctedQty: number; wasCorrection: boolean; note: string } {
  // If the raw string already had feet/inch markers, parsePipeLength handled it
  if (/[\u2018\u2019\u201C\u201D''""\u0027]/.test(rawQty)) {
    return { correctedQty: qty, wasCorrection: false, note: "" };
  }

  // Parse the pipe size for context
  const pipeSize = parseFloat(options?.size || "0") || 0;
  const isSmallBore = pipeSize > 0 && pipeSize <= 4; // Expanded: <=4" is small-bore for inch correction

  // *** PRIMARY GUARD: any bare integer 1-11 with no unit marker is ALMOST CERTAINLY
  // a misread — the AI grabbed the wrong column (size, item NO., or some other
  // small integer cell) instead of the QTY field. We previously auto-converted
  // these to inches/12 = LF, which silently produced fake 0.08, 0.17, 0.25, etc.
  // pipe lengths across many sizes whenever extraction failed.
  //
  // The right behavior is to flag the row for review (qty=0) and let the
  // estimator either (a) correct it from the drawing, or (b) re-run extraction.
  // True 1"–11" inch spool pieces in real BOMs are written WITH a unit marker
  // ("3\"", "0'-8\""), and those are handled by parsePipeLength upstream.
  // True 1'–11' foot pipe lengths in real BOMs are also written WITH the foot
  // marker ("5'", "3'-0\""). A bare integer with no marker means the AI failed
  // to read the qty cell.
  if (Number.isInteger(qty) && qty >= 1 && qty <= 11) {
    const matchesSize = pipeSize > 0 && qty === Math.round(pipeSize);
    const reason = matchesSize
      ? `QTY ${qty} matches pipe SIZE ${pipeSize}\" \u2014 AI likely read the size column.`
      : `QTY ${qty} has no feet/inch marker \u2014 AI likely missed the qty cell.`;
    return {
      correctedQty: 0,
      wasCorrection: true,
      note: `\u26a0\ufe0f Pipe quantity flagged for review. ${reason} Set to 0 LF \u2014 please verify against drawing and edit the row.`
    };
  }

  // Values 12-18 without unit marker: still ambiguous. Could be 12-18 feet of pipe
  // (a real but uncommon length on a single BOM line) or 12-18 inches (1-1.5 feet).
  // We previously auto-converted small-bore to inches; that risks the same
  // silent-corruption issue. Flag instead.
  if (Number.isInteger(qty) && qty >= 12 && qty <= 18) {
    return {
      correctedQty: qty,
      wasCorrection: false,
      note: `\u26a0\ufe0f Ambiguous: ${qty} with no unit marker. Could be ${qty}' or ${qty}\" (${Math.round((qty / 12) * 100) / 100} LF). Verify against drawing.`
    };
  }

  // Values 19-36: assume feet but note for verification
  if (Number.isInteger(qty) && qty >= 19 && qty <= 36) {
    return {
      correctedQty: qty,
      wasCorrection: false,
      note: `Pipe qty ${qty}: assumed feet. Verify if this should be inches.`
    };
  }

  return { correctedQty: qty, wasCorrection: false, note: "" };
}


function isTitlePage(text: string): boolean {
  // If no OCR text available (tesseract not installed), assume it's NOT a title page
  // — Claude will handle filtering during extraction
  if (text.trim().length === 0) return false;
  if (text.trim().length < 50) return true;
  const hasBom = /\b(PIPE|ELBOW|TEE|VALVE|FLANGE|GASKET|BOLT|STUD|REDUCER|W\d+X|HSS|FOOTING|REBAR|STORM|SEWER|ASPHALT)\b/i.test(text);
  if (!hasBom && /\bAREA\s+\d+\b/i.test(text)) return true;
  if (/FOR REFERENCE ONLY/i.test(text) && !hasBom) return true;
  return false;
}


// ============================================================
// CALIBRATION DATA — Known-good factors from completed projects
// ============================================================

const CALIBRATION_DATA: Record<string, {
  project: string;
  benchmark: { actual_ipmh: number; target_ipmh: number; variance_pct: number; total_mh: number };
  ss_weld_factors: Record<string, { mh_per_weld: number; schedule: string; material: string }>;
  bolt_methodology: string;
  small_bore_threshold: number;
  pipe_scope_note: string;
}> = {
  "stolthaven_phase6": {
    project: "Stolthaven Phase 6 Expansion",
    benchmark: {
      actual_ipmh: 0.437,
      target_ipmh: 0.45,
      variance_pct: 3.8,
      total_mh: 56412,
    },
    ss_weld_factors: {
      "3": { mh_per_weld: 4.68, schedule: "10S", material: "SS316" },
      "4": { mh_per_weld: 5.56, schedule: "10S", material: "SS316" },
    },
    bolt_methodology: "field_only",
    small_bore_threshold: 1.5,
    pipe_scope_note: "Branch ISOs only — rack/header pipe tracked separately",
  }
};

function isSmallBoreRollup(item: any): boolean {
  const nps = parseSizeNPSFromString(item.size || "");
  if (nps > 1.5) return false;
  const desc = (item.description || "").toLowerCase();
  const cat = (item.category || "").toLowerCase();
  if (/\b(plug|nipple|coupling|sockolet|weldolet|threadolet)\b/.test(desc)) return true;
  if (cat === "coupling" && nps <= 1.5) return true;
  return false;
}

function detectValveType(description: string): "actuated" | "manual" {
  const d = description.toUpperCase();
  if (/\bAOV\b/.test(d)) return "actuated";
  if (/\bMOV\b/.test(d)) return "actuated";
  if (/\bACTUATED\b/.test(d)) return "actuated";
  if (/\bMOTOR\s*OPERATED\b/.test(d)) return "actuated";
  if (/\bPNEUMATIC\b/.test(d)) return "actuated";
  if (/\bCONTROL\s*VALVE\b/.test(d)) return "actuated";
  if (/\bCV[-\s]/.test(d)) return "actuated";
  return "manual";
}

function detectBranchISOPattern(items: any[]): { detected: boolean; note: string } {
  const pipeItems = items.filter((it: any) => (it.category || "").toLowerCase() === "pipe");
  const reducerItems = items.filter((it: any) => {
    const cat = (it.category || "").toLowerCase();
    return cat === "reducer" || (it.description || "").toLowerCase().includes("reducer");
  });
  
  if (pipeItems.length === 0) return { detected: false, note: "" };
  
  const allSmallPipe = pipeItems.every((it: any) => {
    const nps = parseSizeNPSFromString(it.size || "");
    return nps <= 4;
  });
  
  if (allSmallPipe && reducerItems.length > 0) {
    return {
      detected: true,
      note: "These appear to be branch ISOs. Header/rack pipe may not be included. Add rack pipe quantities separately if needed."
    };
  }
  
  return { detected: false, note: "" };
}

function buildCalibrationSummary(items: any[]): any {
  let totalFieldBoltUps = 0;
  let totalShopBoltUps = 0;
  let totalFieldWelds = 0;
  let totalShopWelds = 0;
  const valveSummary: { manual: Record<string, number>; actuated: Record<string, number> } = { manual: {}, actuated: {} };
  let smallBoreItems = 0;
  
  for (const item of items) {
    const cat = (item.category || "").toLowerCase();
    const loc = item.installLocation || "shop";
    const qty = item.quantity || 0;
    const size = item.size || "unknown";
    
    if (cat === "bolt" || (item.description || "").toLowerCase().includes("bolt")) {
      if (loc === "field") totalFieldBoltUps += qty;
      else totalShopBoltUps += qty;
    }
    
    if (cat === "weld" || (item.description || "").toLowerCase().includes("weld")) {
      if (loc === "field") totalFieldWelds += qty;
      else totalShopWelds += qty;
    }
    
    if (cat === "valve") {
      const vType = item.valveType || "manual";
      const bucket = vType === "actuated" ? valveSummary.actuated : valveSummary.manual;
      bucket[size] = (bucket[size] || 0) + qty;
    }
    
    if (item.smallBoreRollup) smallBoreItems += qty;
  }
  
  const branchISO = detectBranchISOPattern(items);
  
  return {
    totalFieldBoltUps,
    totalShopBoltUps,
    totalFieldWelds,
    totalShopWelds,
    valveSummary,
    smallBoreItems,
    smallBoreRolledUp: smallBoreItems > 0,
    branchISODetected: branchISO.detected,
    rackPipeNote: branchISO.detected ? branchISO.note : undefined,
  };
}

// ============================================================
// MECHANICAL BOM EXTRACTION
// ============================================================

const MECHANICAL_PROMPT = `You are an expert at reading BOM (Bill of Materials) tables from piping isometric drawings. Your job is to produce EXACT, ACCURATE data — this is used for material procurement.

The image is a piping isometric drawing page (or a cropped BOM area from one). The BOM has SHOP and FIELD sub-tables, both with columns: NO. | QTY | SIZE | DESCRIPTION

=== STEP 1: ROW CENSUS (do this FIRST) ===
Before extracting fields, scan the NO. column of every visible BOM sub-table and list every integer you see, in order, for each section (SHOP and FIELD). Your output "items" array must contain ONE entry per visible numbered row. NEVER skip a numbered row — not even if a cell is unreadable, not even if the row appears different (PIPE rows often look different from fitting rows).

Row 1 of the SHOP section is almost always PIPE. If you cannot find a PIPE row but you see fittings (elbow, tee, flange, reducer, valve), look harder — you almost certainly missed it. PIPE rows are commonly missed because:
- The QTY is a length like 38'-1" or 22'-6" instead of an integer
- The DESCRIPTION wraps across multiple visual lines
- The row sits at the very top of the table
Look again. Output the PIPE row.

=== STEP 2: FIELD EXTRACTION (one entry per numbered row) ===
For every NO. value in the row census, output an item with:

- itemNo: the integer from the NO. column
- qty: the value EXACTLY as printed in the QTY cell (see QTY rules below)
- size: the value from the SIZE column
- description: the FULL description text, joining wrapped lines with spaces
- section: "SHOP" or "FIELD"

If any single cell is unreadable or empty, output the row anyway with that field as an empty string "". A row with one empty field is far better than a missing row.

=== QTY RULES ===
The QTY column contains either a length (for pipe) or an integer count (for everything else). Copy the value exactly as printed. Do not convert units, do not do math, do not infer missing units.

- PIPE rows: QTY is a length like 16'-8", 3'-0", 22'-6", or 0'-8". The apostrophe (') means feet, the double-quote (") means inches. Copy it EXACTLY (e.g. qty: "16'-8\"").
- FITTING / VALVE / BOLT / GASKET rows: QTY is a small integer (1, 2, 3, 4, 6, 8, 12, etc.). Distinguish 1 from 4 from 8 carefully.
- If the QTY cell is empty or unreadable, output qty: "".

Note: the SIZE column is BEFORE the QTY column in the table layout. The pipe QTY is the length, not the size. (Post-processing will catch if size and qty get swapped, so don't drop the row out of caution — just output what you read.)

=== SIZE RULES ===
- Valid NPS sizes: 1/2", 3/4", 1", 1-1/4", 1-1/2", 2", 2-1/2", 3", 4", 6", 8", 10", 12", 14", 16", 18", 20", 24", 30", 36", 42", 48". Sizes never exceed 48".
- 1-1/2" is one-and-a-half inches (do not misread as 1" or 11/2").
- Reducers: two sizes like 6"x4" or 2"x1".
- Bolts: diameter x length like 5/8"x4" or 3/4"x4 1/4".

=== DESCRIPTION RULES ===
Copy the FULL text. Join wrapped lines with spaces. Include all specs (ASME B16.xx, ASTM Axxx, CLASS xxxx, SCH xx, etc.).

=== SMALL-BORE COMPLETENESS (1", 3/4", 1/2") ===
Small-bore fittings are the second-most-commonly missed items after PIPE rows. Common patterns:
- Drain/vent assemblies: SOCKOLET + NIPPLE + SW VALVE + CAP, often QTY 4-8 each
- 1" BALL VALVE / PISTON VALVE / SOCKOLET / SW NIPPLE / SW CAP / SW ELBOW / SW TEE

Rules:
- Read the QTY column carefully on small-bore rows — it's often 3-8, not 1.
- Do NOT consolidate similar rows. 8 separate "1" SW ELBOW" rows = 8 separate output items.
- The BOM QTY is the total. Do not multiply by callout count.

=== WHAT TO SKIP ===
- Title block, engineer stamps, revision blocks, drawing border text — these are not BOM rows.
- Header rows / column-name rows that have no NO. value.
- Drawing graphics, callout bubbles, and match-line annotations — the BOM TABLE is the only source. Callout numbers reference BOM rows; do not multiply by their occurrence count.
- If a BOM continues onto another sheet ("CONT'D"), only extract THIS sheet's portion of the table. The fitting AT the continuation boundary appears in both sheets' BOMs — mark it with atContinuation: true.
- If the page has no BOM table at all, return items: [].

=== TITLE BLOCK / DRAWING NUMBER ===
Look in the bottom-right title block for a drawing number like "1\"-150E-080-WW-4805-600" or "6-300B-CS-P-2001" (size, pressure class, material code, line number, sequence). Output as drawingNumber. If absent, set null.

=== CONTINUATION CALLOUTS ===
Look for "CONT'D FROM DWG# ..." or "CONT'D TO DWG# ..." callouts. Add to continuations[] with direction ("from" or "to") and the referenced drawing number. The fitting at the boundary (typically an elbow, tee, reducer, or flange) should also have atContinuation: true.

=== WELD COUNTING ===
On drawings (not BOM crops): count weld symbols visible on the page — filled black dots (BW), open circles (SW), triangles (FW). Output as weldCount: {buttWelds, socketWelds, fieldWelds}. If the image is just a BOM crop with no drawing visible, omit weldCount or set all three to 0.

=== OUTPUT ===
Return ONLY valid JSON (no markdown fences, no commentary):
{"pages": [{"pageNum": PAGE_NUMBER, "drawingNumber": "1\"-150E-080-WW-4805-600", "weldCount": {"buttWelds": 12, "socketWelds": 3, "fieldWelds": 2}, "continuations": [{"direction": "to", "drawing": "P-1001-500", "sheet": 4}], "items": [{"itemNo": 1, "qty": "16'-8\"", "size": "1\"", "description": "PIPE, SMLS, BE OR PE, SCH 80, ASME B36.10, CS ASTM A106, GRD B", "section": "SHOP", "atContinuation": false}]}]}

Final check before returning: count the integers in your row census from Step 1. The items array length must equal that count. If it doesn't, you missed a row — go back and find it.`;

const MECHANICAL_CLOUD_PROMPT = `You are an expert at reading piping isometric drawings AND identifying REVISION CLOUDS. Your job is to produce EXACT, ACCURATE BOM data and flag any items inside revision clouds. Accuracy is critical — this drives material procurement and revision tracking.

For each page, you receive TWO images:
1. The FULL PAGE isometric drawing (piping, fittings, dimensions, AND any revision clouds)
2. The CROPPED BOM TABLE from the same page

=== REVISION CLOUDS ===
Revision clouds are wavy/scalloped bubbles or irregular curved outlines drawn around parts of the drawing to mark changes from the previous revision.

For each BOM item, mark clouded:true if any of these are true:
- The item's piping / fitting / valve on the drawing is enclosed in or touched by a revision cloud
- The item's BOM row itself is enclosed in a revision cloud
- The item's dimension or routing was changed (shown by a cloud)

Provide cloudConfidence (0-100) for each item:
- 90-100: clearly inside or outside a cloud, no ambiguity
- 70-89: near a cloud boundary, likely correct
- 50-69: ambiguous, partially overlapping
- below 50: uncertain, flag for manual review

If there are no revision clouds on the page, set clouded:false and cloudConfidence:100 for every item.

=== STEP 1: ROW CENSUS (do this FIRST for the BOM) ===
Scan the NO. column of every visible BOM sub-table (SHOP and FIELD) and list every integer you see, in order. Your output "items" array must contain ONE entry per visible numbered row. NEVER skip a numbered row — not even if a cell is unreadable.

Row 1 of the SHOP section is almost always PIPE. If you cannot find a PIPE row but you see fittings (elbow, tee, flange, reducer, valve), look harder — you almost certainly missed it. Output the PIPE row.

=== STEP 2: FIELD EXTRACTION ===
For every NO. value in the row census, output an item with itemNo, qty, size, description, section ("SHOP" or "FIELD"), clouded, cloudConfidence, and atContinuation.

If any single cell is unreadable or empty, output the row anyway with that field as an empty string "".

=== QTY RULES ===
Copy the QTY value exactly as printed. Do not convert units, do not infer.
- PIPE rows: a length like 16'-8", 22'-6", 0'-8". The apostrophe (') means feet, double-quote (") means inches.
- FITTING / VALVE / BOLT / GASKET rows: a small integer (1, 2, 3, 4, 6, 8, etc.).
- If the cell is unreadable, output qty: "".

=== SIZE RULES ===
Valid NPS sizes only: 1/2", 3/4", 1", 1-1/4", 1-1/2", 2", 2-1/2", 3", 4", 6", 8", 10", 12", 14", 16", 18", 20", 24", 30", 36", 42", 48". Reducers: "6\"x4\"". Bolts: diameter x length like 5/8"x4".

=== SMALL-BORE COMPLETENESS ===
Small-bore SW fittings (1", 3/4", 1/2") are commonly missed. Read every row. Drain/vent assemblies (sockolet + nipple + SW valve + cap) typically have QTY 4-8 each — read the QTY column carefully. Do not consolidate similar rows.

=== WHAT TO SKIP / AVOID DOUBLE COUNTING ===
- Extract from the BOM table only. Drawing graphics, callout bubbles, and match-line annotations are not BOM rows.
- Callout numbers on the drawing reference BOM rows; do not multiply by their occurrence count.
- For continuation callouts ("CONT'D FROM/TO DWG# ..."), only extract this sheet's portion. The fitting at the boundary appears in both BOMs — mark atContinuation:true so post-processing can dedup.

=== WELD COUNTING ===
Count weld symbols on the FULL PAGE drawing (not the BOM crop): filled black dots = butt welds, open circles = socket welds, small triangles = field welds. Output as weldCount: {buttWelds, socketWelds, fieldWelds}.

=== TITLE BLOCK / DRAWING NUMBER ===
Look in the bottom-right title block for a drawing number like "1\"-150E-080-WW-4805-600". Output as drawingNumber. If absent, set null.

=== OUTPUT ===
Return ONLY valid JSON (no markdown fences):
{"pages": [{"pageNum": PAGE_NUMBER, "drawingNumber": "1\\"-150E-080-WW-4805-600", "weldCount": {"buttWelds": 12, "socketWelds": 3, "fieldWelds": 2}, "continuations": [{"direction": "to", "drawing": "P-1001-500", "sheet": 4}], "items": [{"itemNo": 1, "qty": "16'-8\\"", "size": "1\\"", "description": "PIPE, SMLS, BE, SCH 10S, ASME B36.19, SS ASTM A312", "section": "SHOP", "clouded": false, "cloudConfidence": 95, "atContinuation": false}]}]}

Final check before returning: count the integers in your row census from Step 1. The items array length must equal that count. If it doesn't, you missed a row — go back and find it.`;

// ============================================================
// MECHANICAL VERIFICATION PROMPT (Multi-Pass)
// ============================================================

const MECHANICAL_VERIFY_PROMPT = `You are verifying BOM items extracted from a piping isometric drawing.

You will see:
1. A CROPPED BOM TABLE image from the drawing page
2. A list of previously extracted items from this page

For each item, compare against the image and either CONFIRM it is correct or provide a CORRECTION.

ONLY flag items where you see a clear discrepancy between the extracted data and what is printed on the image.
Do NOT make changes unless you are confident the extracted value is wrong.

Common things to check:
- QTY: Is the quantity correct? Pipe QTY should be feet-inches (e.g., 16'-8"), fittings should be integer counts.
- SIZE: Does the size match what's printed? Watch for 1-1/2" vs 1" confusion.
- DESCRIPTION: Is any important text missing from the description?

Return ONLY valid JSON (no markdown fences):
{"corrections": [{"itemIndex": 0, "field": "qty", "oldValue": "11", "newValue": "11'-0\\"", "reason": "QTY is a pipe length, not a count"}]}

If all items are correct, return: {"corrections": []}`;

async function verifyExtractedItems(
  pageItems: Map<number, { items: any[]; drawingNumber?: string | null; weldCount?: any; continuations?: any[] }>,
  bomImages: Map<number, string>,
): Promise<Map<number, { items: any[]; drawingNumber?: string | null; weldCount?: any; continuations?: any[] }>> {
  const client = getAnthropicClient();
  const result = new Map<number, { items: any[]; drawingNumber?: string | null; weldCount?: any; continuations?: any[] }>();

  // Copy all page data first (preserving drawingNumber, weldCount, continuations)
  for (const [pageNum, pageData] of pageItems) {
    result.set(pageNum, { ...pageData, items: [...pageData.items] });
  }

  // Only verify pages that have items and BOM images
  const pagesToVerify: { pageNum: number; items: any[]; bomPath: string }[] = [];
  for (const [pageNum, pageData] of pageItems) {
    if (pageData.items.length > 0 && bomImages.has(pageNum)) {
      pagesToVerify.push({ pageNum, items: pageData.items, bomPath: bomImages.get(pageNum)! });
    }
  }

  if (pagesToVerify.length === 0) return result;

  // Batch 4-5 pages per verification request
  const VERIFY_BATCH_SIZE = 4;
  for (let i = 0; i < pagesToVerify.length; i += VERIFY_BATCH_SIZE) {
    const batch = pagesToVerify.slice(i, i + VERIFY_BATCH_SIZE);

    const content: Anthropic.MessageCreateParams["messages"][0]["content"] = [];

    for (const page of batch) {
      // Send BOM crop image
      const bomData = fs.readFileSync(page.bomPath);
      content.push({ type: "text" as const, text: `[PAGE ${page.pageNum} - BOM TABLE]` });
      content.push({ type: "image" as const, source: { type: "base64" as const, media_type: "image/png", data: bomData.toString("base64") } });

      // Send extracted items as text
      const itemsList = page.items.map((item: any, idx: number) =>
        `  ${idx}: qty="${item.qty || item.quantity}", size="${item.size}", desc="${item.description}", section="${item.section || "SHOP"}"`
      ).join("\n");
      content.push({ type: "text" as const, text: `Previously extracted items for page ${page.pageNum}:\n${itemsList}` });
    }

    content.push({ type: "text" as const, text: MECHANICAL_VERIFY_PROMPT });

    try {
      const msg = await client.messages.create({
        model: "claude-sonnet-4-5",
        max_tokens: 8192,
        temperature: 0,
        messages: [{ role: "user", content }],
      });

      const responseText = msg.content[0].type === "text" ? msg.content[0].text : "";
      let jsonStr = responseText.trim();
      if (jsonStr.startsWith("```")) {
        jsonStr = jsonStr.replace(/^```(?:json)?\n?/, "").replace(/\n?```$/, "");
      }

      try {
        const parsed = JSON.parse(jsonStr);
        if (parsed.corrections && Array.isArray(parsed.corrections)) {
          for (const correction of parsed.corrections) {
            // In multi-page batches, reject corrections without explicit pageNum to avoid misrouting
            if (!correction.pageNum && batch.length > 1) {
              console.warn(`  Verification correction rejected: no pageNum in multi-page batch (${batch.length} pages)`);
              continue;
            }
            const targetPage = correction.pageNum || batch[0].pageNum;
            const pageData = result.get(targetPage);
            const items = pageData?.items;
            if (items && correction.itemIndex >= 0 && correction.itemIndex < items.length) {
              const item = items[correction.itemIndex];
              if (correction.field === "qty") {
                item.qty = correction.newValue;
              } else if (correction.field === "size") {
                item.size = correction.newValue;
              } else if (correction.field === "description") {
                item.description = correction.newValue;
              }
              item._verified = true;
              item._verifyNote = correction.reason;
              console.log(`  Verification correction page ${targetPage} item ${correction.itemIndex}: ${correction.field} "${correction.oldValue}" → "${correction.newValue}" (${correction.reason})`);
            }
          }
        }
      } catch (parseErr) {
        console.error("  Failed to parse verification response:", parseErr);
      }
    } catch (apiErr: any) {
      console.error("  Verification API error:", apiErr.message || apiErr);
    }
  }

  return result;
}

// ============================================================
// STRUCTURAL BOM EXTRACTION
// ============================================================

const STRUCTURAL_PROMPT = `You are an expert structural engineer and estimator reading structural engineering drawings.
Your job is to extract a COMPLETE structural Bill of Materials (BOM) / takeoff from the drawing.
Accuracy is critical — this is used for material procurement and cost estimating.

WHAT TO EXTRACT:
**STRUCTURAL STEEL:** Wide Flange Beams/Columns (W10x22, W14x30, etc.), HSS/Tube Steel (HSS 6x6x1/4), Angles (L3x3x1/4), Channels, Plates, Base Plates, Bracing, Misc Steel
**CONNECTIONS:** Bolts (A325, A490), Anchor bolts, Welds, Clip angles, Gusset plates
**CONCRETE:** Footings, Grade beams, Walls, Slabs, Columns/Piers (give volume in CY)
**REBAR:** Bars by size (#3 through #11), Wire mesh, Anchor bolts, Dowels

RULES:
1. Read EVERY cell EXACTLY as printed. Do NOT guess, round, or estimate.
2. For each item extract: mark (if any), description, size, quantity, unit, grade/spec, weight.
3. Units: use EA (each), LF (linear feet), SF (square feet), CY (cubic yards), LBS (pounds), TONS.
4. Member marks (B1, C1, etc.) go in the "mark" field.
5. Grade/spec examples: A36, A992, A500-B, 3000 psi, 4000 psi, Grade 60.
6. For concrete volumes, compute CY = (L x W x H in feet) / 27.
7. SKIP: general notes, title blocks, revision clouds, drawing borders.

Return ONLY valid JSON (no markdown, no extra text):
{"pages": [{"pageNum": PAGE_NUMBER, "items": [
  {"mark": "B1", "category": "wide_flange", "description": "BEAM W14x30", "size": "W14x30", "quantity": 4, "unit": "EA", "grade": "A992", "weight": 2400}
]}]}`;

// ============================================================
// CIVIL BOM EXTRACTION
// ============================================================

const CIVIL_PROMPT = `You are an expert civil engineer reading civil construction drawings (site plans, utility plans, grading plans, paving plans).
Your job is to extract a COMPLETE civil takeoff — all quantities needed for bidding and construction.
Accuracy is critical — this is used for material procurement and project bidding.

WHAT TO EXTRACT:
**1. UNDERGROUND UTILITIES (Pipe):** Storm Drain (RCP, HDPE, PVC), Sanitary Sewer, Water Line, Gas Line — material, diameter, length in LF
**2. UNDERGROUND STRUCTURES (EA):** Manholes, Catch Basins, Fire Hydrants, Valves, Fittings, Service Connections
**3. SITEWORK & GRADING (CY, SF):** Cut/Fill Earthwork, Rock Excavation, Trench Excavation, Backfill, Import/Export Material
**4. PAVING & CONCRETE:** Asphalt (HMA), Concrete Paving, Base Course, Curb & Gutter, Sidewalk, Retaining Walls, Pavement Markings
**5. EROSION CONTROL & MISC:** Silt Fence, Inlet Protection, Seeding/Sodding, Fencing, Signage

CIVIL DRAWING READING RULES:
1. Look for pipe schedules, manhole schedules, structure schedules (tables on the drawing)
2. Grading plans show cut/fill volumes in an earthwork summary table
3. Paving plans show typical sections with layer thicknesses AND a paving schedule
4. Utility plans show pipe runs with lengths annotated along the pipe centerline
5. Check ALL tables, notes, legends, and schedules
6. For pipes: sum up ALL annotated lengths on the sheet for each pipe type/size/material
7. Read ALL text on the full page — civil drawings have important info in every corner

Return ONLY valid JSON (no markdown fences):
{"pages": [{"pageNum": PAGE_NUMBER, "items": [
  {"category": "storm_pipe", "description": "18\" RCP STORM DRAIN PIPE CLASS III", "size": "18\"", "quantity": 245, "unit": "LF", "material": "RCP", "depth": "8'"},
  {"category": "manhole", "description": "48\" STORM MANHOLE", "size": "48\"", "quantity": 3, "unit": "EA", "material": "PRECAST CONCRETE", "depth": "8'"}
]}]}`;

// ============================================================
// AI VISION EXTRACTION — Gemini-first, Claude fallback
// ============================================================

/**
 * Try extracting BOM items from page images using Gemini first (free tier),
 * falling back to Claude if Gemini fails or returns no results.
 */
async function extractBatchWithGemini(
  batch: { pageNum: number; imagePath: string }[],
  prompt: string
): Promise<{ pages: { pageNum: number; items: any[] }[] } | null> {
  const geminiKey = getUserGeminiKey();
  if (!geminiKey) return null;

  try {
    // Build Gemini content parts: images + page markers + prompt
    const parts: any[] = [];
    for (const page of batch) {
      const imgData = fs.readFileSync(page.imagePath);
      const b64 = imgData.toString("base64");
      parts.push({ inlineData: { mimeType: "image/png", data: b64 } });
      parts.push({ text: `[PAGE ${page.pageNum}]` });
    }
    parts.push({ text: prompt });

    const resp = await fetch(
      `https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent`,
      {
        method: "POST",
        headers: { "Content-Type": "application/json", "x-goog-api-key": geminiKey },
        body: JSON.stringify({
          contents: [{ parts }],
          generationConfig: { temperature: 0, maxOutputTokens: 16384 },
        }),
      }
    );

    if (!resp.ok) {
      const errText = await resp.text().catch(() => "");
      console.warn(`  Gemini API error: ${resp.status} ${resp.statusText} ${errText.substring(0, 200)}`);
      return null;
    }

    const data = await resp.json();
    const text = data?.candidates?.[0]?.content?.parts?.[0]?.text || "";

    // Parse JSON from response
    let jsonStr = text.trim();
    if (jsonStr.startsWith("```")) {
      jsonStr = jsonStr.replace(/^```(?:json)?\n?/, "").replace(/\n?```$/, "");
    }

    const parsed = JSON.parse(jsonStr);
    if (parsed.pages && Array.isArray(parsed.pages)) {
      // Validate we got items
      const totalItems = parsed.pages.reduce((s: number, p: any) => s + (p.items?.length || 0), 0);
      if (totalItems > 0) {
        console.log(`  Gemini extracted ${totalItems} items from pages ${batch.map(p => p.pageNum).join(", ")}`);
        return parsed;
      }
    }
    console.warn(`  Gemini returned no items for pages ${batch.map(p => p.pageNum).join(", ")}`);
    return null;
  } catch (err: any) {
    console.warn(`  Gemini extraction failed: ${err.message?.substring(0, 100)}`);
    return null;
  }
}

async function extractWithVision(
  pageImages: { pageNum: number; imagePath: string; pdfText?: string }[],
  prompt: string,
  discipline: string
): Promise<{ results: Map<number, { items: any[]; drawingNumber?: string | null; weldCount?: any; continuations?: any[] }>; authFailures: number }> {
  let client = getAnthropicClient();
  const results = new Map<number, { items: any[]; drawingNumber?: string | null; weldCount?: any; continuations?: any[] }>();
  let authFailures = 0;
  // BATCH_SIZE=1: send one page per Claude call. Was 2, but identified as a
  // cross-page contamination bug on May 8 — when two visually-similar BOM
  // images were batched (e.g. consecutive 4" detail spool sheets), Claude
  // collapsed both pages to one BOM, copying the simpler page's quantities
  // onto the more complex one. Single-page-per-call eliminates this entirely
  // at the cost of ~2x calls. Worth it for accuracy.
  const BATCH_SIZE = 1;
  const hasGemini = !!getUserGeminiKey();
  const hasClaude = !!getUserApiKey();

  for (let batchStart = 0; batchStart < pageImages.length; batchStart += BATCH_SIZE) {
    const batch = pageImages.slice(batchStart, batchStart + BATCH_SIZE);
    const pageNums = batch.map(p => p.pageNum).join(", ");

    // === CLAUDE PRIMARY ===
    if (!hasClaude) {
      // No Claude key — try Gemini as fallback
      if (hasGemini) {
        console.log(`  AI Vision [${discipline}] Gemini batch: pages ${pageNums}...`);
        const geminiResult = await extractBatchWithGemini(batch, prompt);
        if (geminiResult && geminiResult.pages) {
          for (const page of geminiResult.pages) {
            if (page.pageNum && Array.isArray(page.items)) {
              results.set(page.pageNum, { items: page.items, drawingNumber: page.drawingNumber || null, weldCount: page.weldCount || null, continuations: page.continuations || [] });
            }
          }
          await new Promise(resolve => setTimeout(resolve, 4500));
          continue;
        }
      }
      console.warn(`  No API keys available for pages ${pageNums}`);
      continue;
    }

    console.log(`  AI Vision [${discipline}] Claude batch: pages ${pageNums}...`);

    const content: Anthropic.MessageCreateParams["messages"][0]["content"] = [];

    for (const page of batch) {
      const imgData = fs.readFileSync(page.imagePath);
      const b64 = imgData.toString("base64");
      content.push({ type: "image" as const, source: { type: "base64" as const, media_type: "image/png", data: b64 } });
      content.push({ type: "text" as const, text: `[PAGE ${page.pageNum}]` });
      // Optional PDF text hint: if the PDF has embedded vector text (CAD-exported
      // ISOs typically do), pass it as evidence — the model treats this as a
      // high-quality OCR result and is much less likely to drop rows. Only add
      // when the text actually contains BOM-shaped content to avoid noise.
      if (page.pdfText && page.pdfText.length > 50) {
        const trimmed = page.pdfText.length > 6000 ? page.pdfText.substring(0, 6000) + "\n[\u2026truncated]" : page.pdfText;
        content.push({
          type: "text" as const,
          text: `[PAGE ${page.pageNum} \u2014 PDF embedded text, treat as evidence not authority; the image is still the source of truth]\n${trimmed}`,
        });
      }
    }

    content.push({ type: "text" as const, text: prompt });

    // Retry logic: 1 retry with exponential backoff on API failure
    // On auth errors: create a fresh client and retry once
    let lastApiErr: any = null;
    for (let attempt = 0; attempt < 2; attempt++) {
      try {
        if (attempt > 0) {
          const backoffMs = 2000 * Math.pow(2, attempt - 1);
          console.log(`  Retrying after ${backoffMs}ms backoff (attempt ${attempt + 1})...`);
          await new Promise(resolve => setTimeout(resolve, backoffMs));
        }

        const msg = await client.messages.create({
          model: "claude-sonnet-4-5",
          max_tokens: 16384,
          temperature: 0,
          messages: [{ role: "user", content }],
        });

        const responseText = msg.content[0].type === "text" ? msg.content[0].text : "";
        let jsonStr = responseText.trim();
        if (jsonStr.startsWith("```")) {
          jsonStr = jsonStr.replace(/^```(?:json)?\n?/, "").replace(/\n?```$/, "");
        }

        let parseSuccess = false;
        try {
          const parsed = JSON.parse(jsonStr);
          if (parsed.pages && Array.isArray(parsed.pages)) {
            for (const page of parsed.pages) {
              if (page.pageNum && Array.isArray(page.items)) {
                results.set(page.pageNum, { items: page.items, drawingNumber: page.drawingNumber || null, weldCount: page.weldCount || null, continuations: page.continuations || [] });
                if (page.items.length === 0) {
                  console.warn(`  Warning: Page ${page.pageNum} returned 0 items`);
                }
              }
            }
            parseSuccess = true;
          } else {
            console.error(`  AI response missing pages array for batch at page ${batch[0].pageNum}`);
          }
        } catch (parseErr) {
          console.error(`  Failed to parse AI response (attempt ${attempt + 1}) for batch at page ${batch[0].pageNum}:`, parseErr);
          console.error(`  Response: ${responseText.substring(0, 500)}`);
        }
        if (parseSuccess) {
          lastApiErr = null;
          break; // Success, exit retry loop
        }
        // Parse failed — treat as retryable error
        if (attempt === 1) {
          console.error(`  Parse failed after all retries for batch at page ${batch[0].pageNum}`);
          for (const page of batch) { results.set(page.pageNum, { items: [] }); }
        } else {
          console.log(`  Parse failed, will retry...`);
        }
        continue; // Retry on parse failure too
      } catch (apiErr: any) {
        lastApiErr = apiErr;
        console.error(`  AI Vision API error (attempt ${attempt + 1}):`, apiErr.message || apiErr);

        // On auth error: refresh client and retry
        if (isAuthError(apiErr) && attempt === 0) {
          console.warn(`  API key may have expired, creating fresh client...`);
          client = getAnthropicClient();
          authFailures++;
          continue;
        }

        if (attempt === 1) {
          // Final attempt failed
          if (isAuthError(apiErr)) authFailures++;
          for (const page of batch) { results.set(page.pageNum, { items: [] }); }
        }
      }
    }
  }

  return { results, authFailures };
}

// Double-pass verification: sends same images + extracted items to Claude for cross-check
async function verifyExtractionPass(
  pageImages: { pageNum: number; imagePath: string }[],
  extractedItems: Map<number, { items: any[]; drawingNumber?: string | null; weldCount?: any; continuations?: any[] }>,
  discipline: string
): Promise<Map<number, { items: any[]; drawingNumber?: string | null; weldCount?: any; continuations?: any[] }>> {
  const client = getAnthropicClient();
  if (!client || !getUserApiKey()) return extractedItems;

  // BATCH_SIZE=1: send one page per Claude call. Was 2, but identified as a
  // cross-page contamination bug on May 8 — when two visually-similar BOM
  // images were batched (e.g. consecutive 4" detail spool sheets), Claude
  // collapsed both pages to one BOM, copying the simpler page's quantities
  // onto the more complex one. Single-page-per-call eliminates this entirely
  // at the cost of ~2x calls. Worth it for accuracy.
  const BATCH_SIZE = 1;
  const verified = new Map(extractedItems);

  for (let batchStart = 0; batchStart < pageImages.length; batchStart += BATCH_SIZE) {
    const batch = pageImages.slice(batchStart, batchStart + BATCH_SIZE);

    const content: Anthropic.MessageCreateParams["messages"][0]["content"] = [];
    const itemsByPage: Record<number, any[]> = {};

    for (const page of batch) {
      const pageData = extractedItems.get(page.pageNum);
      if (!pageData || pageData.items.length === 0) continue;

      const imgData = fs.readFileSync(page.imagePath);
      const b64 = imgData.toString("base64");
      content.push({ type: "image" as const, source: { type: "base64" as const, media_type: "image/png", data: b64 } });
      content.push({ type: "text" as const, text: `[PAGE ${page.pageNum}]` });
      itemsByPage[page.pageNum] = pageData.items;
    }

    if (content.length === 0) continue;

    const itemsSummary = Object.entries(itemsByPage).map(([pageNum, items]) => {
      return `PAGE ${pageNum} extracted items (${items.length}):\n` +
        items.map((item: any, i: number) => `  ${i + 1}. NO:${item.itemNo || "?"} QTY:${item.qty} SIZE:${item.size} DESC:${item.description} SECTION:${item.section || "SHOP"}`).join("\n");
    }).join("\n\n");

    content.push({ type: "text" as const, text: `VERIFICATION PASS — Check the extracted items against the BOM table image.

Here are the items extracted in the first pass:
${itemsSummary}

VERIFY EACH ITEM and most importantly LOOK FOR MISSED ITEMS:

1. *** SMALL-BORE FITTINGS *** — these are the MOST COMMONLY MISSED. Carefully scan the BOM for:
   - 1/2", 3/4", 1", 1-1/4", 2" SOCKET WELD elbows, tees, valves, flanges
   - SOCKOLET, WELDOLET, THREADOLET items
   - SOCKET WELD 3000# class fittings
   Count how many small-bore (≤2") items are in the extracted list above. Now scan the BOM image. If the BOM has MORE small-bore items than the extracted list, ADD the missing ones.

   *** SMALL-BORE QTY VERIFICATION *** — for each small-bore item already extracted, RE-READ its QTY from the BOM image. The QTY column is often misread because the digits are small. Specifically check:
   - 1" BALL VALVE / PISTON VALVE SW: did you read QTY 3 when it says 4? Or QTY 1 when it says 7?
   - 1" SOCKOLET: did you read QTY 2 when it says 8?
   - 1" PIPE NIPPLE: did you read QTY 1 when it says 6?
   - 1" CAP, SW: did you read QTY 1 when it says 4?
   These rows often have QTY > 1 because they are part of repeating drain/vent subassemblies. UPDATE the QTY if you misread it.

   *** SUBASSEMBLY COMPLETENESS CHECK *** — drain/vent subassemblies always come in groups: SOCKOLET + NIPPLE + VALVE + CAP. If the BOM has 4 sockolets, it should also have 4 nipples, 4 valves, and 4 caps (or close to that). If the counts don't roughly match, you missed something. Add the missing rows.

2. ITEM NUMBER CHECK: The BOM has a NO. column with sequential numbers (1, 2, 3...). Find the highest item number in the BOM image. Count items in the extracted list. If extracted count < highest item number, items were missed — find which item numbers are missing and add them.

3. QTY CORRECTNESS: Check feet/inches for pipe. Watch for 8 vs 8", 11 vs 11". A bare "8" in qty almost always means 8 inches.

4. SIZE CORRECTNESS: Watch for 1-1/2" vs 1-1/4", and similar lookalike fractions.

5. DESCRIPTION COMPLETENESS: Full text with specs.

6. SECTION TAG: SHOP vs FIELD assigned correctly.

7. NO DUPLICATES within a single page.

Return the CORRECTED item list including ALL missed items. The output should have AT LEAST as many items as the BOM table actually contains. If the BOM has 25 rows, output 25 items.

Return ONLY valid JSON (no markdown fences):
{"pages": [{"pageNum": NUMBER, "items": [...corrected items + missed items...], "corrections": ["description of each correction or addition"]}]}` });

    try {
      const msg = await client.messages.create({
        model: "claude-sonnet-4-5",
        max_tokens: 16384,
        temperature: 0,
        messages: [{ role: "user", content }],
      });

      const responseText = msg.content[0].type === "text" ? msg.content[0].text : "";
      let jsonStr = responseText.trim();
      if (jsonStr.startsWith("```")) jsonStr = jsonStr.replace(/^```(?:json)?\n?/, "").replace(/\n?```$/, "");

      const parsed = JSON.parse(jsonStr);
      if (parsed.pages && Array.isArray(parsed.pages)) {
        for (const page of parsed.pages) {
          if (page.pageNum && Array.isArray(page.items) && page.items.length > 0) {
            const existing = verified.get(page.pageNum);
            verified.set(page.pageNum, {
              ...existing,
              items: page.items,
            });
            if (page.corrections && page.corrections.length > 0) {
              console.log(`  Verification corrected page ${page.pageNum}: ${page.corrections.join("; ")}`);
            }
          }
        }
      }
    } catch (err: any) {
      console.warn(`  Verification pass failed for pages ${batch.map(p => p.pageNum).join(",")}: ${err.message?.substring(0, 100)}`);
      // Keep original extraction on verification failure
    }
  }

  return verified;
}

// Cloud-aware extraction: sends FULL PAGE + BOM CROP per page
async function extractWithCloudDetection(
  pageImages: { pageNum: number; bomImagePath: string; fullImagePath: string }[],
  prompt: string,
  onPageComplete?: (completedCount: number, totalCount: number) => void
): Promise<{ results: Map<number, { items: any[]; drawingNumber?: string | null; weldCount?: any; continuations?: any[] }>; authFailures: number }> {
  let client = getAnthropicClient();
  const results = new Map<number, { items: any[]; drawingNumber?: string | null; weldCount?: any; continuations?: any[] }>();
  let authFailures = 0;

  async function processOnePage(page: { pageNum: number; bomImagePath: string; fullImagePath: string }): Promise<void> {
    console.log(`  AI Vision [cloud-detect] page ${page.pageNum}...`);

    const content: Anthropic.MessageCreateParams["messages"][0]["content"] = [];

    // Full page image first
    const fullData = fs.readFileSync(page.fullImagePath);
    content.push({ type: "text" as const, text: `[PAGE ${page.pageNum} - FULL DRAWING]` });
    content.push({ type: "image" as const, source: { type: "base64" as const, media_type: "image/png", data: fullData.toString("base64") } });

    // Then BOM crop
    const bomData = fs.readFileSync(page.bomImagePath);
    content.push({ type: "text" as const, text: `[PAGE ${page.pageNum} - BOM TABLE]` });
    content.push({ type: "image" as const, source: { type: "base64" as const, media_type: "image/png", data: bomData.toString("base64") } });

    content.push({ type: "text" as const, text: prompt });

    for (let attempt = 0; attempt < 2; attempt++) {
      try {
        if (attempt > 0) {
          await new Promise(resolve => setTimeout(resolve, 2000 * Math.pow(2, attempt - 1)));
        }

        const msg = await client.messages.create({
          model: "claude-sonnet-4-5",
          max_tokens: 16384,
          temperature: 0,
          messages: [{ role: "user", content }],
        });

        const responseText = msg.content[0].type === "text" ? msg.content[0].text : "";
        let jsonStr = responseText.trim();
        if (jsonStr.startsWith("```")) {
          jsonStr = jsonStr.replace(/^```(?:json)?\n?/, "").replace(/\n?```$/, "");
        }

        let parseSuccess = false;
        try {
          const parsed = JSON.parse(jsonStr);
          if (parsed.pages && Array.isArray(parsed.pages)) {
            // Collect all items from the response, but key them to the KNOWN page number
            // to prevent AI-hallucinated page numbers from losing data
            const allItems: any[] = [];
            let drawingNumber: string | null = null;
            let weldCount: any = null;
            let continuations: any[] = [];
            for (const p of parsed.pages) {
              if (Array.isArray(p.items)) {
                allItems.push(...p.items);
              }
              if (p.drawingNumber && !drawingNumber) drawingNumber = p.drawingNumber;
              if (p.weldCount && !weldCount) weldCount = p.weldCount;
              if (Array.isArray(p.continuations) && p.continuations.length > 0) continuations.push(...p.continuations);
            }
            results.set(page.pageNum, { items: allItems, drawingNumber, weldCount, continuations });
            parseSuccess = true;
          }
        } catch (parseErr) {
          console.error(`  Failed to parse cloud detection response (attempt ${attempt + 1}) page ${page.pageNum}`);
        }
        if (parseSuccess) { break; }
        if (attempt === 1) { results.set(page.pageNum, { items: [] }); }
      } catch (apiErr: any) {
        console.error(`  Cloud detection API error (attempt ${attempt + 1}):`, apiErr.message || apiErr);

        // On auth error: refresh client and retry
        if (isAuthError(apiErr) && attempt === 0) {
          console.warn(`  API key may have expired, creating fresh client...`);
          client = getAnthropicClient();
          authFailures++;
          continue;
        }

        if (attempt === 1) {
          if (isAuthError(apiErr)) authFailures++;
          results.set(page.pageNum, { items: [] });
        }
      }
    }
  }

  // Process 3 pages concurrently in batches
  const CONCURRENT_PAGES = 3;
  let completedPages = 0;
  for (let i = 0; i < pageImages.length; i += CONCURRENT_PAGES) {
    const batch = pageImages.slice(i, i + CONCURRENT_PAGES);
    await Promise.all(batch.map(async (page) => {
      await processOnePage(page);
      completedPages++;
      if (onPageComplete) onPageComplete(completedPages, pageImages.length);
    }));
    // 1-second delay between batches to avoid rate limiting
    if (i + CONCURRENT_PAGES < pageImages.length) {
      await new Promise(r => setTimeout(r, 1000));
    }
  }

  return { results, authFailures };
}

// ============================================================
// ASYNC EXEC HELPERS (unblock event loop)
// ============================================================

function execFileAsync(cmd: string, args: string[], opts: any): Promise<void> {
  return new Promise((resolve, reject) => {
    execFile(cmd, args, opts, (err) => {
      if (err) reject(err);
      else resolve();
    });
  });
}

function execAsync(cmd: string, opts: any): Promise<void> {
  return new Promise((resolve, reject) => {
    exec(cmd, opts, (err) => {
      if (err) reject(err);
      else resolve();
    });
  });
}

// ============================================================
// RENDER FUNCTIONS
// ============================================================

// Safe image processing without shell injection — processes pages using execFileAsync
// Check if tesseract is available (cache result)
let _tesseractAvailable: boolean | null = null;
function isTesseractAvailable(): boolean {
  if (_tesseractAvailable !== null) return _tesseractAvailable;
  try {
    execFileSync("tesseract", ["--version"], { timeout: 5000 });
    _tesseractAvailable = true;
  } catch {
    console.warn("tesseract not available — OCR will be skipped, Claude will read images directly");
    _tesseractAvailable = false;
  }
  return _tesseractAvailable;
}

async function processRenderedPages(jobDir: string, mode: "bom" | "bom+full" | "ocr-only" | "fullpage"): Promise<void> {
  const files = fs.readdirSync(jobDir).filter(f => f.endsWith(".png") && !f.includes("_bom") && !f.includes("_full") && !f.includes("_ocr"));
  // Reduce concurrency to limit peak memory usage (ImageMagick is memory-hungry)
  const CONCURRENCY = 2;
  const hasTesseract = isTesseractAvailable();
  
  for (let i = 0; i < files.length; i += CONCURRENCY) {
    const batch = files.slice(i, i + CONCURRENCY);
    await Promise.all(batch.map(async (imgFile) => {
      const imgPath = path.join(jobDir, imgFile);
      const basename = imgFile.replace(/\.png$/, "");
      try {
        if (mode === "fullpage") {
          // No-crop mode: rename original to *_bom.png so downstream lookups
          // (which always look for `${basename}_bom.png`) just work. This is
          // the recovery / re-extraction mode — the model gets the entire
          // page image and locates the BOM itself. Used when the standard
          // crop has clipped row 1 or another part of the table.
          const bomPath = path.join(jobDir, `${basename}_bom.png`);
          try { fs.renameSync(imgPath, bomPath); } catch (e) { console.warn("Suppressed error:", e); return; }
          // Run OCR on the full page so isTitlePage() filtering still works
          if (hasTesseract) {
            const ocrBase = path.join(jobDir, `${basename}_ocr`);
            try {
              await execFileAsync("tesseract", [bomPath, ocrBase, "--psm", "4", "-l", "eng"], { timeout: 30000 });
            } catch (e) { console.warn("Suppressed error:", e); }
          }
          return;
        }
        if (mode === "bom" || mode === "bom+full") {
          // Get image dimensions
          let dims = "";
          try {
            dims = execFileSync("identify", ["-format", "%wx%h", imgPath], { timeout: 10000 }).toString().trim();
          } catch { return; } // skip unreadable images
          // BUG-7 fix: validate identify output format before parsing
          const dimsMatch = dims.replace(/"/g, "").match(/^(\d+)x(\d+)$/);
          if (!dimsMatch) return;
          const w = parseInt(dimsMatch[1], 10);
          const h = parseInt(dimsMatch[2], 10);
          if (w < 100 || h < 100) return;
          // BOM crop — expanded for row-1 safety after May-8 council review.
          // Old: right-50% × top-65%, starting at +0,+0 (clipped row 1 on ~30%
          // of pages because the BOM header sat against the top edge).
          // New: right-55% × top-80%, still starting at +0,+0 — wider on the
          // left to capture the NO. column when the BOM sits slightly inset,
          // and taller on the bottom so the full SHOP+FIELD tables fit even
          // when the BOM is positioned mid-page.
          const cx = Math.floor(w * 45 / 100);
          const cw = w - cx;
          const ch = Math.floor(h * 80 / 100);
          
          // Crop BOM area
          const bomPath = path.join(jobDir, `${basename}_bom.png`);
          await execFileAsync("convert", [imgPath, "-crop", `${cw}x${ch}+${cx}+0`, "+repage", bomPath], { timeout: 30000 });
          
          if (mode === "bom+full") {
            // Create resized full page image
            const fullPath = path.join(jobDir, `${basename}_full.png`);
            await execFileAsync("convert", [imgPath, "-resize", "50%", fullPath], { timeout: 30000 });
          }
          
          // OCR the BOM crop (skip if tesseract not available)
          if (hasTesseract) {
            const ocrBase = path.join(jobDir, `${basename}_ocr`);
            try {
              await execFileAsync("tesseract", [bomPath, ocrBase, "--psm", "4", "-l", "eng"], { timeout: 30000 });
            } catch (e) { console.warn("Suppressed error:", e); }
          }
          
          // Remove original full-res image (keep _bom and _full)
          try { fs.unlinkSync(imgPath); } catch (e) { console.warn("Suppressed error:", e); }
        } else {
          // OCR-only mode (full page) — skip if tesseract not available
          if (hasTesseract) {
            const ocrBase = path.join(jobDir, `${basename}_ocr`);
            try {
              await execFileAsync("tesseract", [imgPath, ocrBase, "--psm", "4", "-l", "eng"], { timeout: 30000 });
            } catch (e) { console.warn("Suppressed error:", e); }
          }
        }
      } catch (e) { console.warn("Suppressed error processing", imgFile, e); }
    }));
  }
}



async function renderCroppedBomImages(pdfPath: string, pageCount: number): Promise<{
  pageImages: { pageNum: number; imagePath: string; tesseractText: string; pdfText?: string }[];
  jobDir: string;
}> {
  const jobDir = path.join(RENDER_DIR, Date.now().toString());
  fs.mkdirSync(jobDir, { recursive: true });

  // Render page-by-page to avoid OOM on limited-memory servers
  // Use 150 DPI for BOM tables (text-heavy, doesn't need high res)
  // 200 DPI for better small-bore fitting detection. Was 150 but caused
  // missed extractions on dense small-bore (1", 3/4") BOM rows.
  const dpiForBom = 200;
  for (let p = 1; p <= pageCount; p++) {
    try {
      await execFileAsync("pdftoppm", [
        "-r", String(dpiForBom), "-png", "-f", String(p), "-l", String(p),
        pdfPath, path.join(jobDir, "page")
      ], { maxBuffer: 30 * 1024 * 1024, timeout: 45000 });
    } catch (err: any) {
      console.warn(`  pdftoppm page ${p} failed: ${err.message?.substring(0, 100)}`);
    }
  }

    await processRenderedPages(jobDir, "bom");

  // Probe whether the PDF has embedded vector text. If yes, we will pull the
  // per-page text and pass it to the extractor as evidence — this dramatically
  // reduces row-skip errors on CAD-exported ISOs (which is most of them).
  const hasVectorText = pdfHasVectorText(pdfPath, pageCount);
  if (hasVectorText) {
    console.log(`  PDF has embedded vector text \u2014 will pass per-page text to extractor as evidence`);
  }

  const pageImages: { pageNum: number; imagePath: string; tesseractText: string; pdfText?: string }[] = [];
  const padLen = Math.max(2, String(pageCount).length);

  for (let p = 1; p <= pageCount; p++) {
    const padded = String(p).padStart(padLen, "0");
    const bomImg = path.join(jobDir, `page-${padded}_bom.png`);
    const ocrFile = path.join(jobDir, `page-${padded}_ocr.txt`);
    let tesseractText = "";
    if (fs.existsSync(ocrFile)) tesseractText = fs.readFileSync(ocrFile, "utf-8");
    if (fs.existsSync(bomImg) && !isTitlePage(tesseractText)) {
      const pdfText = hasVectorText ? extractPageText(pdfPath, p) : "";
      pageImages.push({ pageNum: p, imagePath: bomImg, tesseractText, pdfText });
    }
  }

  return { pageImages, jobDir };
}

// Renders both full page images AND cropped BOM images for cloud detection
async function renderBomWithFullPages(pdfPath: string, pageCount: number): Promise<{
  pageImages: { pageNum: number; bomImagePath: string; fullImagePath: string; tesseractText: string }[];
  jobDir: string;
}> {
  const jobDir = path.join(RENDER_DIR, Date.now().toString() + "_cloud");
  fs.mkdirSync(jobDir, { recursive: true });

  // Render page-by-page at reduced DPI to save memory
  for (let p = 1; p <= pageCount; p++) {
    try {
      await execFileAsync("pdftoppm", [
        "-r", "200", "-png", "-f", String(p), "-l", String(p),
        pdfPath, path.join(jobDir, "page")
      ], { maxBuffer: 40 * 1024 * 1024, timeout: 50000 });
    } catch (err: any) {
      console.warn(`  pdftoppm page ${p} (cloud) failed: ${err.message?.substring(0, 100)}`);
    }
  }

  // Crop BOM area but KEEP the full page image for cloud detection
    await processRenderedPages(jobDir, "bom+full");

  const pageImages: { pageNum: number; bomImagePath: string; fullImagePath: string; tesseractText: string }[] = [];
  const padLen = Math.max(2, String(pageCount).length);


  for (let p = 1; p <= pageCount; p++) {
    // Try multiple padding lengths since pdftoppm may use 1 or 2+ digits
    let bomImg = "", fullImg = "", ocrFile = "";
    for (const pl of [padLen, 1, 2, 3]) {
      const padded = String(p).padStart(pl, "0");
      const b = path.join(jobDir, `page-${padded}_bom.png`);
      if (fs.existsSync(b)) {
        bomImg = b;
        fullImg = path.join(jobDir, `page-${padded}_full.png`);
        ocrFile = path.join(jobDir, `page-${padded}_ocr.txt`);
        break;
      }
    }
    let tesseractText = "";
    if (ocrFile && fs.existsSync(ocrFile)) tesseractText = fs.readFileSync(ocrFile, "utf-8");
    const bomExists = bomImg && fs.existsSync(bomImg);
    const fullExists = fullImg && fs.existsSync(fullImg);
    const isTitle = isTitlePage(tesseractText);
    if (bomExists && fullExists && !isTitle) {
      pageImages.push({ pageNum: p, bomImagePath: bomImg, fullImagePath: fullImg, tesseractText });
    }
  }

  return { pageImages, jobDir };
}

async function renderFullPageImages(pdfPath: string, pageCount: number, dpi = 300): Promise<{
  pageImages: { pageNum: number; imagePath: string; tesseractText: string }[];
  jobDir: string;
}> {
  const jobDir = path.join(RENDER_DIR, Date.now().toString() + "_" + Math.random().toString(36).slice(2));
  fs.mkdirSync(jobDir, { recursive: true });

  // Render page-by-page at reduced DPI to save memory
  const effectiveDpi = Math.min(dpi, 200); // Cap DPI for cloud deployment
  for (let p = 1; p <= pageCount; p++) {
    try {
      await execFileAsync("pdftoppm", [
        "-r", String(effectiveDpi), "-png", "-f", String(p), "-l", String(p),
        pdfPath, path.join(jobDir, "page")
      ], { maxBuffer: 30 * 1024 * 1024, timeout: 45000 });
    } catch (err: any) {
      console.warn(`  pdftoppm page ${p} (full) failed: ${err.message?.substring(0, 100)}`);
    }
  }

    await processRenderedPages(jobDir, "ocr-only");

  const pageImages: { pageNum: number; imagePath: string; tesseractText: string }[] = [];
  const padLen = Math.max(2, String(pageCount).length);

  for (let p = 1; p <= pageCount; p++) {
    const padded = String(p).padStart(padLen, "0");
    const imgPath = path.join(jobDir, `page-${padded}.png`);
    const ocrFile = path.join(jobDir, `page-${padded}_ocr.txt`);
    let tesseractText = "";
    if (fs.existsSync(ocrFile)) tesseractText = fs.readFileSync(ocrFile, "utf-8");
    if (fs.existsSync(imgPath) && !isTitlePage(tesseractText)) {
      pageImages.push({ pageNum: p, imagePath: imgPath, tesseractText });
    }
  }

  return { pageImages, jobDir };
}

// ============================================================
// CATEGORY DETECTION
// ============================================================

const MECHANICAL_CATEGORY_PATTERNS = [
  { category: "elbow", patterns: [/\bELBOW\b/i, /\bELL\b/i, /\bRETURN\b/i] },
  { category: "tee", patterns: [/\bTEE\b/i] },
  { category: "reducer", patterns: [/\bREDUCER\b/i, /\bSWAGE\b/i] },
  { category: "valve", patterns: [/\bVALVE\b/i] },
  { category: "flange", patterns: [/\bFLANGE\b/i, /\bFLG\b/i] },
  { category: "gasket", patterns: [/\bGASKET\b/i] },
  { category: "bolt", patterns: [/\bSTUD\s*BOLT\b/i, /\bSTUD\b(?!.*PIPE)/i, /\bBOLT\b(?!.*PIPE)/i, /\bHEAVY\s*HEX\s*NUT\b/i] },
  { category: "cap", patterns: [/\bPIPE\s*CAP\b/i, /\bEND\s*CAP\b/i, /\bCAP\b(?!.*SCREW)/i] },
  { category: "coupling", patterns: [/\bCOUPLING\b/i, /\bSOCKOLET\b/i, /\bWELDOLET\b/i, /\bTHREADOLET\b/i, /\bNIPPLE\b/i] },
  { category: "union", patterns: [/\bUNION\b/i] },
  { category: "strainer", patterns: [/\bSTRAINER\b/i] },
  { category: "weld", patterns: [/\bWELD\b/i] },
  { category: "pipe", patterns: [/\bPIPE\b/i, /\bSCH\s*\d+\b/i] },
  { category: "other", patterns: [/\bPLUG\b/i] },
];

function detectMechanicalCategory(text: string): string {
  for (const { category, patterns } of MECHANICAL_CATEGORY_PATTERNS) {
    for (const p of patterns) {
      if (p.test(text)) return category;
    }
  }
  return "other";
}

function extractSpec(text: string): string {
  const m = text.match(/ASTM\s*[A-Z]?\s*\d+/i);
  if (m) return m[0].toUpperCase();
  const m2 = text.match(/ASME\s*B[\d.]+/i);
  if (m2) return m2[0].toUpperCase();
  return "";
}

function extractMaterial(text: string): string {
  const materials = ["CARBON STEEL", "CS", "STAINLESS STEEL", "SS", "304L", "316L", "304", "316", "A106", "A105", "A312", "A182", "A193", "A194", "A234", "A403"];
  const upper = text.toUpperCase();
  for (const mat of materials) { if (upper.includes(mat)) return mat; }
  return "";
}

function extractSchedule(text: string): string {
  const m = text.match(/SCH\s*(\d+[Ss]?|STD|XS|XXS)/i);
  if (m) return `SCH ${m[1].toUpperCase()}`;
  return "";
}

function extractRating(text: string): string {
  const m = text.match(/CLASS\s*(\d+)/i);
  if (m) return `#${m[1]}`;
  const m2 = text.match(/#\s*(\d+)/);
  if (m2) return `#${m2[1]}`;
  return "";
}

function validateSize(rawSize: string, category: string): string {
  if (!rawSize || rawSize === "N/A") return rawSize;
  const numericParts = rawSize.replace(/["''\u201D\u2019]/g, "").split(/x/i);
  for (const part of numericParts) {
    const cleaned = part.trim();
    const mixedMatch = cleaned.match(/^(\d+)[\s-]+(\d+)\/(\d+)$/);
    let value = 0;
    if (mixedMatch) {
      value = parseInt(mixedMatch[1], 10) + parseInt(mixedMatch[2], 10) / parseInt(mixedMatch[3], 10);
    } else {
      const num = parseFloat(cleaned);
      if (!isNaN(num)) value = num;
    }
    if (value > 60 && category !== "bolt") {
      console.warn(`  Suspicious size "${rawSize}" for ${category} — flagging as N/A.`);
      return "N/A";
    }
  }
  return rawSize;
}

// ============================================================
// AUTO-CORRECTION RULES (Spec Item 6)
// ============================================================

function autoCorrectItem(item: any): any {
  const desc = (item.description || "").toUpperCase();
  let category = item.category;
  let size = (item.size || "").trim();
  let unit = item.unit;
  let quantity = item.quantity;

  // Category corrections based on description keywords
  if (/\bNIPPLE\b/.test(desc) && category === "pipe") category = "coupling";
  if (/\bSOCKOLET\b|\bWELDOLET\b|\bTHREADOLET\b/.test(desc) && category !== "fitting" && category !== "coupling") category = "fitting";
  if (/\bGASKET\b/.test(desc) && category !== "gasket") category = "gasket";
  if (/\bSTUD\s*BOLT\b|\bHEX\s*BOLT\b/.test(desc) && category !== "bolt") category = "bolt";
  if (/\bPLUG\b/.test(desc) && category === "other") {
    // Plugs with small sizes are fittings
    const sizeNum = parseFloat(size);
    if (!isNaN(sizeNum) && sizeNum <= 2) category = "fitting";
  }
  if (/\bCAP\b/.test(desc) && category !== "cap" && !/SCREW/.test(desc)) category = "cap";
  if (/\bUNION\b/.test(desc) && category !== "union") category = "union";

  // Auto-detect SS material from description (Calibration Item 4)
  // Set itemMaterial for known SS pipe specs so correct labor rates apply downstream
  let itemMaterial = item.itemMaterial;
  if (!itemMaterial) {
    if (/\b(SS|STAINLESS|TP304|TP316|304L?|316L?|A312|A182|A403)\b/i.test(desc)) {
      // Determine SS304 vs SS316
      if (/\b(316|TP316|316L)\b/i.test(desc)) {
        itemMaterial = "SS";
        item.material = item.material || "SS316";
      } else {
        itemMaterial = "SS";
        item.material = item.material || "SS304";
      }
    }
  }

  // Size corrections
  // Normalize "3 IN" → "3\"", "1/2 IN" → "1/2\""
  size = size.replace(/\s+IN(?:CH(?:ES)?)?$/i, '"');
  // Fix "1 1/2\"" → "1-1/2\"" and "2 1/2\"" → "2-1/2\""
  size = size.replace(/^(\d)\s+(1\/2)/, "$1-$2");
  size = size.replace(/^(\d)\s+(1\/4)/, "$1-$2");
  // Remove leading/trailing spaces
  size = size.trim();
  // Flag if size looks like a length (e.g., "16'-8\"") for non-pipe items
  if (category !== "pipe" && /^\d+['''\u2019]\s*[-–]?\s*\d*/.test(size)) {
    item._sizeWarning = `Size "${size}" looks like a pipe length but category is ${category}`;
  }

  // Unit corrections
  if (category === "pipe" && unit === "EA") unit = "LF";
  if (category !== "pipe" && unit === "LF" && Number.isInteger(quantity) && quantity < 50) unit = "EA";

  return { ...item, category, size, unit, quantity, itemMaterial: itemMaterial || item.itemMaterial };
}

// ============================================================
// POST-PROCESSING VALIDATION (Spec Item 3)
// ============================================================

const VALID_NPS_SIZES = new Set(["1/2\"", "3/4\"", "1\"", "1-1/4\"", "1-1/2\"", "2\"", "2-1/2\"", "3\"", "4\"", "6\"", "8\"", "10\"", "12\"", "14\"", "16\"", "18\"", "20\"", "24\"", "30\"", "36\"", "42\"", "48\""]);
const VALID_NPS_NUMBERS = new Set([0.5, 0.75, 1, 1.25, 1.5, 2, 2.5, 3, 4, 6, 8, 10, 12, 14, 16, 18, 20, 24, 30, 36, 42, 48]);

function isValidNPS(sizeStr: string): boolean {
  if (!sizeStr || sizeStr === "N/A") return false;
  const s = sizeStr.trim();
  if (VALID_NPS_SIZES.has(s)) return true;
  // Check numeric
  const num = parseSizeNPSFromString(s);
  return VALID_NPS_NUMBERS.has(num);
}

function parseSizeNPSFromString(s: string): number {
  const cleaned = s.replace(/["''″\u201D\u2019]/g, "").trim();
  // Mixed fraction: "1-1/2" or "2-1/2"
  const mixedMatch = cleaned.match(/^(\d+)[\s-]+(\d+)\/(\d+)$/);
  if (mixedMatch) return parseInt(mixedMatch[1], 10) + parseInt(mixedMatch[2], 10) / parseInt(mixedMatch[3], 10);
  // Simple fraction: "3/4", "1/2"
  const fracMatch = cleaned.match(/^(\d+)\/(\d+)$/);
  if (fracMatch) return parseInt(fracMatch[1], 10) / parseInt(fracMatch[2], 10);
  const num = parseFloat(cleaned);
  return isNaN(num) ? 0 : num;
}

function isReducerSize(sizeStr: string): boolean {
  return /^\d[\d\-\/]*["\s]*x\s*\d[\d\-\/]*["\s]*$/.test(sizeStr.trim()) && !isBoltSize(sizeStr);
}

function isBoltSize(sizeStr: string): boolean {
  // Pattern: fraction"x fraction" or fraction"x integer"
  return /^\d+\/\d+["''″]?\s*x\s*\d/i.test(sizeStr);
}

function validateExtractedItems(items: any[]): any[] {
  // Group items by sheet number for per-sheet analysis
  const bySheet: Record<string, any[]> = {};
  for (const item of items) {
    const sheetMatch = (item.notes || "").match(/Sheet\s+(\d+)/i);
    const sheet = sheetMatch ? sheetMatch[1] : "0";
    if (!bySheet[sheet]) bySheet[sheet] = [];
    bySheet[sheet].push(item);
  }

  for (const item of items) {
    const notes: string[] = [];
    const sheetMatch = (item.notes || "").match(/Sheet\s+(\d+)/i);
    const sheet = sheetMatch ? sheetMatch[1] : "0";
    const sheetItems = bySheet[sheet] || [];

    // Pipe length sanity
    if (item.category === "pipe") {
      if (item.quantity > 500) {
        notes.push("Pipe qty >500 LF on single sheet — almost certainly wrong");
        item._validationFlag = "low";
      } else if (item.quantity > 100) {
        notes.push("Pipe qty >100 LF on single sheet — unusual");
        if (!item._validationFlag) item._validationFlag = "medium";
      }
      if (item.quantity <= 0) {
        notes.push("Pipe qty is 0 or negative");
        item._validationFlag = "low";
      }
    }

    // Fitting-to-pipe ratio
    const fittingsOnSheet = sheetItems.filter((i: any) => i.category !== "pipe" && i.category !== "bolt" && i.category !== "gasket").length;
    const pipeOnSheet = sheetItems.filter((i: any) => i.category === "pipe").length;
    if (fittingsOnSheet > 20 && pipeOnSheet === 0 && item.category !== "pipe") {
      notes.push("Many fittings but no pipe on this sheet — unusual");
    }

    // Duplicate detection on same sheet
    const dupsOnSheet = sheetItems.filter((i: any) =>
      i !== item && i.category === item.category && i.size === item.size && i.description === item.description
    );
    if (dupsOnSheet.length > 0) {
      notes.push("Potential duplicate — same category/size/description on same sheet");
      if (!item._validationFlag) item._validationFlag = "medium";
    }

    // Size validation
    if (item.size && item.size !== "N/A" && item.category !== "bolt") {
      if (/x/i.test(item.size)) {
        // Reducer — validate both parts
        const parts = item.size.split(/x/i);
        for (const part of parts) {
          const num = parseSizeNPSFromString(part.trim());
          if (num > 0 && !VALID_NPS_NUMBERS.has(num)) {
            notes.push(`Size "${item.size}" contains non-standard NPS`);
            if (!item._validationFlag) item._validationFlag = "medium";
          }
        }
      } else if (!isValidNPS(item.size)) {
        notes.push(`Size "${item.size}" is not standard NPS`);
        if (!item._validationFlag) item._validationFlag = "medium";
      }
    }

    // Bolt size validation
    if (item.category === "bolt" && item.size && item.size !== "N/A") {
      if (!isBoltSize(item.size) && !isValidNPS(item.size)) {
        notes.push(`Bolt size "${item.size}" doesn't match expected pattern`);
      }
    }

    // Spec completeness
    const descUpper = (item.description || "").toUpperCase();
    if (item.category === "pipe" && !/ASTM|ASME/.test(descUpper)) {
      notes.push("Pipe description missing ASME/ASTM spec reference");
      if (!item._validationFlag) item._validationFlag = "medium";
    }
    if (["elbow", "tee", "reducer", "cap", "coupling"].includes(item.category) && !/ASME\s*B16/i.test(descUpper)) {
      notes.push("Fitting missing ASME B16.x reference");
    }
    if (item.category === "bolt" && !/ASTM\s*A193|ASTM\s*A194/i.test(descUpper)) {
      notes.push("Bolt missing ASTM A193/A194 reference");
    }

    // Quantity plausibility
    if (item.category === "bolt" && item.quantity % 4 !== 0 && item.quantity > 1) {
      notes.push("Bolt qty is not a multiple of 4 — verify count");
    }
    // Gasket-flange matching check: compare within same sheet
    if (item.category === "gasket") {
      const matchingFlanges = sheetItems.filter((i: any) => i.category === "flange" && i.size === item.size);
      if (matchingFlanges.length > 0) {
        const flangeTotal = matchingFlanges.reduce((sum: number, f: any) => sum + f.quantity, 0);
        if (flangeTotal > 0 && item.quantity < flangeTotal) {
          notes.push(`Gasket count (${item.quantity}) may not match flange count (${flangeTotal})`);
        }
      }
    }

    if (notes.length > 0) {
      item._validationNotes = notes;
    }
  }

  // Adjacent sheet overlap detection
  const sheetNums = Object.keys(bySheet).map(Number).filter(n => !isNaN(n)).sort((a, b) => a - b);
  for (let i = 0; i < sheetNums.length - 1; i++) {
    const thisSheet = bySheet[String(sheetNums[i])] || [];
    const nextSheet = bySheet[String(sheetNums[i + 1])] || [];
    if (thisSheet.length === 0 || nextSheet.length === 0) continue;

    // Compute overlap: items with matching category+size+description
    const thisKeys = new Set(thisSheet.map((item: any) => `${item.category}|${item.size}|${item.description}`));
    const nextKeys = new Set(nextSheet.map((item: any) => `${item.category}|${item.size}|${item.description}`));
    let overlap = 0;
    for (const key of thisKeys) {
      if (nextKeys.has(key)) overlap++;
    }
    const overlapPct = overlap / Math.max(thisKeys.size, nextKeys.size);
    if (overlapPct > 0.8) {
      // Flag all items on the next sheet as potential continuation overlap
      for (const item of nextSheet) {
        if (!item._validationNotes) item._validationNotes = [];
        item._validationNotes.push(`>80% BOM overlap with Sheet ${sheetNums[i]} — possible continuation page`);
        if (!item._validationFlag) item._validationFlag = "medium";
      }
    }
  }

  return items;
}

// ============================================================
// PDF QUALITY GATE (Council Item 1)
// ============================================================

function classifyPdfQuality(pageImages: { pageNum: number; imagePath: string; tesseractText: string }[]): "vector" | "clean_scan" | "poor_scan" {
  if (pageImages.length === 0) return "clean_scan";

  let totalTextLength = 0;
  let totalFileSize = 0;
  let pageCount = 0;

  for (const page of pageImages) {
    totalTextLength += (page.tesseractText || "").length;
    try {
      const stat = fs.statSync(page.imagePath);
      totalFileSize += stat.size;
    } catch (e) { console.warn("Suppressed error:", e); }
    pageCount++;
  }

  const avgTextLength = pageCount > 0 ? totalTextLength / pageCount : 0;
  const avgFileSize = pageCount > 0 ? totalFileSize / pageCount : 0;

  // Vector PDFs render to large images (>500KB) and have rich OCR text
  if (avgFileSize > 500 * 1024 || avgTextLength > 500) return "vector";

  // When tesseract is not installed, avgTextLength will always be 0.
  // Fall back to file size only: rendered engineering drawings at 150 DPI
  // are typically 100KB+ per page for clean scans.
  if (avgTextLength === 0 && avgFileSize > 50 * 1024) return "clean_scan";

  // Clean scans have moderate OCR text
  if (avgTextLength >= 100) return "clean_scan";

  // Only classify as poor_scan if we actually have tesseract data AND it's sparse
  // Without tesseract, assume clean_scan to avoid false "poor quality" warnings
  if (avgTextLength === 0 && !isTesseractAvailable()) return "clean_scan";

  return "poor_scan";
}

// ============================================================
// PIPING VALIDATION RULES (Council Item 4)
// ============================================================

function validatePipingBom(items: any[]): any[] {
  // Flange-bolt-gasket check
  const flangesBySize: Record<string, any[]> = {};
  const boltsBySize: Record<string, any[]> = {};
  const gasketsBySize: Record<string, any[]> = {};

  for (const item of items) {
    const cat = (item.category || "").toLowerCase();
    const desc = (item.description || "").toLowerCase();
    const size = (item.size || "").toLowerCase().trim();
    if (!size) continue;

    if (cat === "flange" || desc.includes("flange")) {
      if (!flangesBySize[size]) flangesBySize[size] = [];
      flangesBySize[size].push(item);
    } else if (cat === "bolt" || desc.includes("bolt") || desc.includes("stud")) {
      if (!boltsBySize[size]) boltsBySize[size] = [];
      boltsBySize[size].push(item);
    } else if (cat === "gasket" || desc.includes("gasket")) {
      if (!gasketsBySize[size]) gasketsBySize[size] = [];
      gasketsBySize[size].push(item);
    }
  }

  // Check that BOM has ANY bolts/gaskets when flanges exist (bolt sizes differ from flange NPS)
  const hasAnyBolts = Object.keys(boltsBySize).length > 0;
  const hasAnyGaskets = Object.keys(gasketsBySize).length > 0;
  if (Object.keys(flangesBySize).length > 0 && (!hasAnyBolts || !hasAnyGaskets)) {
    const missing: string[] = [];
    if (!hasAnyBolts) missing.push("bolt sets");
    if (!hasAnyGaskets) missing.push("gaskets");
    // Flag all flanges
    for (const flanges of Object.values(flangesBySize)) {
      for (const flange of flanges) {
        flange.notes = (flange.notes || "") + ` | Warning: BOM has flanges but no ${missing.join("/")} found — verify bolt/gasket takeoff`;
      }
    }
  }

  // Size range check
  for (const item of items) {
    const cat = (item.category || "").toLowerCase();
    if (cat === "pipe" || cat === "elbow" || cat === "tee" || cat === "reducer" || cat === "valve" || cat === "flange") {
      const sizeStr = (item.size || "").replace(/[^\d.]/g, "");
      const sizeNum = parseFloat(sizeStr);
      if (sizeNum > 48) {
        item.notes = (item.notes || "") + " | Warning: Pipe size >48\" — suspicious, verify";
        if (!item._validationFlag || item._validationFlag === "high") item._validationFlag = "medium";
      } else if (sizeNum > 0 && sizeNum < 0.5) {
        item.notes = (item.notes || "") + " | Warning: Pipe size <1/2\" — suspicious, verify";
        if (!item._validationFlag || item._validationFlag === "high") item._validationFlag = "medium";
      }
    }
  }

  // Quantity outlier check
  for (const item of items) {
    const cat = (item.category || "").toLowerCase();
    if (cat !== "pipe" && (item.quantity || 0) > 50) {
      item.notes = (item.notes || "") + " | Unusually high quantity \u2014 verify";
      if (!item._validationFlag || item._validationFlag === "high") item._validationFlag = "medium";
    }
  }

  // Spec/material consistency check by line
  // Group by line (from notes like "Sheet X")
  const lineGroups: Record<string, any[]> = {};
  for (const item of items) {
    const lineMatch = (item.notes || "").match(/Sheet\s+(\d+)/i);
    const lineKey = lineMatch ? lineMatch[1] : "unknown";
    if (!lineGroups[lineKey]) lineGroups[lineKey] = [];
    lineGroups[lineKey].push(item);
  }

  for (const [_line, lineItems] of Object.entries(lineGroups)) {
    if (lineItems.length < 3) continue;
    const materialCounts: Record<string, number> = {};
    for (const item of lineItems) {
      const mat = (item.material || "").trim().toUpperCase();
      if (mat) {
        materialCounts[mat] = (materialCounts[mat] || 0) + 1;
      }
    }
    const total = lineItems.filter(i => (i.material || "").trim()).length;
    if (total < 3) continue;
    // Find dominant material
    let dominant = "";
    let dominantCount = 0;
    for (const [mat, count] of Object.entries(materialCounts)) {
      if (count > dominantCount) { dominant = mat; dominantCount = count; }
    }
    if (dominantCount / total >= 0.8) {
      // Flag outliers
      for (const item of lineItems) {
        const mat = (item.material || "").trim().toUpperCase();
        if (mat && mat !== dominant) {
          item.notes = (item.notes || "") + ` | Material spec outlier: ${mat} differs from ${dominant} (${Math.round(dominantCount/total*100)}% of line)`;
          if (!item._validationFlag || item._validationFlag === "high") item._validationFlag = "medium";
        }
      }
    }
  }

  return items;
}

// ============================================================
// PIPING RULES ENGINE — Post-extraction engineering validation
// ============================================================

function validatePipingRules(items: any[]): { warnings: string[]; autoFixes: { itemId: string; field: string; oldValue: any; newValue: any; reason: string }[] } {
  const warnings: string[] = [];
  const autoFixes: { itemId: string; field: string; oldValue: any; newValue: any; reason: string }[] = [];

  // Group items by page for per-sheet validation
  const byPage: Record<number, any[]> = {};
  for (const item of items) {
    const page = item.sourcePage || 0;
    if (!byPage[page]) byPage[page] = [];
    byPage[page].push(item);
  }

  for (const [pageStr, pageItems] of Object.entries(byPage)) {
    const page = parseInt(pageStr);

    // --- Rule 1: Flange-Gasket-Bolt matching ---
    const flanges = pageItems.filter(i => (i.category || "").toLowerCase() === "flange");
    const gaskets = pageItems.filter(i => (i.category || "").toLowerCase() === "gasket");
    const bolts = pageItems.filter(i => (i.category || "").toLowerCase() === "bolt");

    const flangesBySize: Record<string, number> = {};
    for (const f of flanges) {
      const sz = f.size || "?";
      flangesBySize[sz] = (flangesBySize[sz] || 0) + (f.quantity || 0);
    }

    for (const [size, flangeQty] of Object.entries(flangesBySize)) {
      const gasketQty = gaskets.filter(g => g.size === size).reduce((s, g) => s + (g.quantity || 0), 0);
      const expectedGaskets = Math.ceil(flangeQty / 2);

      if (gasketQty === 0 && flangeQty > 0) {
        warnings.push(`Page ${page}: ${flangeQty}x ${size}" flanges found but NO gaskets \u2014 possible missing gaskets`);
      } else if (gasketQty < expectedGaskets - 1) {
        warnings.push(`Page ${page}: ${flangeQty}x ${size}" flanges expect ~${expectedGaskets} gaskets but only ${gasketQty} found`);
      }

      if (bolts.length === 0 && flangeQty > 0) {
        warnings.push(`Page ${page}: ${flangeQty}x ${size}" flanges found but NO bolt sets \u2014 possible missing bolts`);
      }
    }

    // --- Rule 2: Socket weld fittings shouldn't appear on large bore (6"+) ---
    for (const item of pageItems) {
      const desc = (item.description || "").toLowerCase();
      const size = parseFloat(item.size) || 0;
      const isSW = desc.includes("socket weld") || desc.includes(",sw,") || desc.includes(" sw ") || /\\bsw\\b/i.test(desc);

      if (isSW && size >= 6) {
        warnings.push(`Page ${page}: ${item.size}" ${(item.description || "").substring(0, 40)} \u2014 socket weld fitting on ${size}"+ pipe is unusual, verify`);
      }
    }

    // --- Rule 3: Reducer size consistency ---
    for (const item of pageItems) {
      if ((item.category || "").toLowerCase() === "reducer" && item.size) {
        const match = item.size.match(/(\d+(?:[.-]\d+\/\d+)?)\s*[""x\u00d7]\s*(\d+(?:[.-]\d+\/\d+)?)/i);
        if (match) {
          const large = parseFloat(match[1]) || 0;
          const small = parseFloat(match[2]) || 0;
          if (small >= large && large > 0) {
            warnings.push(`Page ${page}: Reducer ${item.size} \u2014 large end should be bigger than small end, verify sizes`);
          }
        }
      }
    }

    // --- Rule 4: Pipe size consistency on a page ---
    const pipeSizes = new Set(
      pageItems.filter(i => (i.category || "").toLowerCase() === "pipe").map(i => i.size).filter(Boolean)
    );
    const fittingSizes = new Set(
      pageItems.filter(i => ["elbow", "tee", "cap", "coupling", "union"].includes((i.category || "").toLowerCase())).map(i => i.size).filter(Boolean)
    );
    for (const fittingSize of fittingSizes) {
      if (pipeSizes.size > 0 && !pipeSizes.has(fittingSize)) {
        const fsNum = parseFloat(fittingSize) || 0;
        const isReducerSize = [...pipeSizes].some(ps => (parseFloat(ps) || 0) > fsNum);
        if (!isReducerSize) {
          warnings.push(`Page ${page}: ${fittingSize}" fitting found but no ${fittingSize}" pipe on this page \u2014 verify size`);
        }
      }
    }
  }

  // --- Global Rule: Total gasket count vs total flange count ---
  const totalFlanges = items.filter(i => (i.category || "").toLowerCase() === "flange").reduce((s, i) => s + (i.quantity || 0), 0);
  const totalGaskets = items.filter(i => (i.category || "").toLowerCase() === "gasket").reduce((s, i) => s + (i.quantity || 0), 0);
  if (totalFlanges > 0 && totalGaskets === 0) {
    warnings.push(`Overall: ${totalFlanges} flanges found across all pages but ZERO gaskets \u2014 gaskets may not have been extracted`);
  }

  return { warnings, autoFixes };
}

// ============================================================
// CONTINUATION GRAPH — Maps drawing connections
// ============================================================

function buildContinuationGraph(items: any[]): {
  graph: Record<string, { page: number; connectsTo: string[]; connectsFrom: string[] }>;
  sharedFittings: { fromPage: number; toPage: number; fitting: string; size: string }[];
} {
  const graph: Record<string, { page: number; connectsTo: string[]; connectsFrom: string[] }> = {};

  for (const item of items) {
    if (item.drawingNumber && item.sourcePage) {
      if (!graph[item.drawingNumber]) {
        graph[item.drawingNumber] = { page: item.sourcePage, connectsTo: [], connectsFrom: [] };
      }
    }
    if (item._continuations && Array.isArray(item._continuations)) {
      const myDrawing = item.drawingNumber || `page_${item.sourcePage}`;
      if (!graph[myDrawing]) graph[myDrawing] = { page: item.sourcePage || 0, connectsTo: [], connectsFrom: [] };

      for (const conn of item._continuations) {
        if (conn.direction === "to" && conn.drawing) {
          if (!graph[myDrawing].connectsTo.includes(conn.drawing)) {
            graph[myDrawing].connectsTo.push(conn.drawing);
          }
        } else if (conn.direction === "from" && conn.drawing) {
          if (!graph[myDrawing].connectsFrom.includes(conn.drawing)) {
            graph[myDrawing].connectsFrom.push(conn.drawing);
          }
        }
      }
    }
  }

  const sharedFittings: { fromPage: number; toPage: number; fitting: string; size: string }[] = [];
  const contItems = items.filter(i => i.atContinuation === true || i.atContinuation === "true");

  for (const [, node] of Object.entries(graph)) {
    for (const connectedDrawing of node.connectsFrom) {
      const connectedNode = graph[connectedDrawing];
      if (!connectedNode) continue;

      const myContItems = contItems.filter(i => i.sourcePage === node.page);
      const theirContItems = contItems.filter(i => i.sourcePage === connectedNode.page);

      for (const myItem of myContItems) {
        for (const theirItem of theirContItems) {
          if (myItem.category === theirItem.category && myItem.size === theirItem.size) {
            sharedFittings.push({
              fromPage: connectedNode.page,
              toPage: node.page,
              fitting: myItem.category,
              size: myItem.size,
            });
          }
        }
      }
    }
  }

  return { graph, sharedFittings };
}

// ============================================================
// HISTORICAL PATTERN APPLICATION — Auto-correct from learned patterns
// ============================================================

// ============================================================
// POTENTIAL DUPLICATE FLAGGING (does NOT remove items)
// ============================================================
//
// User asked: 'I don't want anything to be taken off but i want it flagged
// for review as well and marked low confidence'.
//
// Finds same-page items where the AI extracted multiple identical rows with
// qty=1 each, when the BOM almost certainly had ONE row with qty=N. Common
// case: '2 elbows extracted as 2 separate qty=1 rows instead of one qty=2'.
// We mark them as low-confidence + add a clear note. The estimator decides
// in Review Mode whether to merge / delete / keep.

function flagPotentialDuplicates(items: any[]): { flaggedCount: number; groups: number } {
  // Group items by (sourcePage, normalized desc+size+category)
  const groupKey = (it: any) => {
    const sp = it.sourcePage ?? "none";
    const desc = (it.description || "").toLowerCase().replace(/[\s,]+/g, " ").trim();
    const descSig = desc.split(" ").slice(0, 4).join(" ");
    const size = (it.size || "").toLowerCase().replace(/[\s"\u2018\u2019\u201C\u201D'`]/g, "");
    const cat = (it.category || "").toLowerCase();
    return `${sp}|${cat}|${size}|${descSig}`;
  };

  const groups: Record<string, any[]> = {};
  for (const item of items) {
    const k = groupKey(item);
    if (!groups[k]) groups[k] = [];
    groups[k].push(item);
  }

  let flaggedCount = 0;
  let groupsCount = 0;
  for (const [, group] of Object.entries(groups)) {
    if (group.length < 2) continue;
    // Only flag if MOST entries in the group are qty=1 (the typical AI duplication
    // pattern). If the AI legitimately extracted multiple rows with varying qty,
    // those are probably real distinct items.
    const qty1Count = group.filter(it => Math.round(it.quantity || 0) === 1).length;
    if (qty1Count < 2) continue;

    groupsCount++;
    const totalQty = group.reduce((s, it) => s + (Math.round(it.quantity || 0) || 0), 0);
    const note = `\u26a0 Possible duplicate \u2014 ${group.length} identical rows on same page (combined qty would be ${totalQty}). Verify whether BOM has 1 row of qty=${totalQty} or ${group.length} separate rows.`;
    for (const item of group) {
      item.confidence = "low";
      item.confidenceScore = Math.min(item.confidenceScore || 50, 50);
      // Persist via BOTH _validationNotes (pre-confidence-scoring path) AND
      // confidenceNotes (post-DB-persist path) so the note survives all flows.
      item._validationNotes = item._validationNotes || [];
      if (!item._validationNotes.some((n: string) => n.startsWith("\u26a0 Possible duplicate"))) {
        item._validationNotes.push(note);
      }
      const existingNotes = item.confidenceNotes || "";
      if (!existingNotes.includes("Possible duplicate")) {
        item.confidenceNotes = existingNotes ? `${existingNotes} | ${note}` : note;
      }
      // Mark for review so Review Mode surfaces it
      item.reviewStatus = "unreviewed";
      flaggedCount++;
    }
  }

  if (groupsCount > 0) {
    console.log(`  Flagged ${flaggedCount} item(s) across ${groupsCount} potential duplicate group(s)`);
  }
  return { flaggedCount, groups: groupsCount };
}

// ============================================================
// PIPE QTY BEST-GUESS RETRY (for items with qty=0 / unreadable)
// ============================================================
//
// Pipe rows where the AI couldn't read the QTY cell come back with qty=0 and
// a flag note. User asked for: 'put best guess and then flag with low
// confidence'. This pass sends each unread pipe row back to Claude with the
// FULL ISO drawing + the row's description and asks for a best-guess length.
// Result is filled in as the qty + confidence='low' + clear note.

// ============================================================
// MISSED PIPE-ROW RECOVERY
// ============================================================
//
// User reported: 'I am missing whole lines for pipe, like multiple pages that
// have pipe doesn't even have it on the takeoff'.
//
// The bare-integer guard catches pipe rows where the QTY was misread as a
// number, but cannot help when the AI SKIPS the pipe row entirely. Looking
// at the FP-Isos package, 6 of 22 pages had a pipe row in the BOM that the
// AI never extracted at all.
//
// This pass goes through ALL pages and for each page, sends the rendered BOM
// image to Claude with a tightly focused prompt: 'count and list every PIPE
// row in this BOM table'. Anything Claude finds that we don't already have
// gets added as a new item with low confidence + a clear flag note.
//
// Different from pipeQtyBestGuessRetry (which only fixes existing items with
// qty=0). This recovers ROWS THAT WERE NEVER EXTRACTED.
//
// SUPERSEDED (May 2026 council review): pipeRowRecovery is now a fallback. The
// primary path is detectSuspectPages() + reExtractLowConfidencePages() with
// renderMode="fullpage", which fixes the same row-1 misses upstream by re-
// running the full extraction pipeline at 300 DPI on uncropped pages.
// We keep pipeRowRecovery for safety as a final net but it should rarely fire
// in practice once the BOM crop expansion + re-extract wiring is live.

// Detects pages whose extracted items show signs of an extraction failure:
//   * Total failure: a page-number GAP in the extracted set (e.g., we got
//     items for pages 1,2,3,4,6,7,... — page 5 returned zero items so it's
//     missing from the set entirely). The model occasionally returns an
//     empty items array for a perfectly valid BOM page — these need
//     re-extraction at full-page DPI.
//   * Row-1 miss: fittings present (elbow/tee/flange) but no PIPE row
//   * Gap at top: lowest-numbered itemNo > 1 (rows above were dropped)
//
// Returns the list of suspect global page numbers.
function detectSuspectPages(items: any[]): number[] {
  const byPage: Record<number, any[]> = {};
  for (const it of items) {
    const p = it.sourcePage;
    if (typeof p !== "number") continue;
    if (!byPage[p]) byPage[p] = [];
    byPage[p].push(it);
  }
  const suspect = new Set<number>();

  // Total-failure detection via page-number gaps.
  // Find the min and max source page across all items. Any integer in that
  // range that has zero items is a missing page — almost certainly an empty
  // extraction (the model returned items: []). Title pages won't be in the
  // range because they were never sent to extraction in the first place
  // (filtered upstream by isTitlePage).
  const pagesWithItems = Object.keys(byPage).map(p => parseInt(p)).filter(p => byPage[p].length > 0);
  if (pagesWithItems.length >= 2) {
    const minPage = Math.min(...pagesWithItems);
    const maxPage = Math.max(...pagesWithItems);
    for (let p = minPage + 1; p < maxPage; p++) {
      if (!byPage[p] || byPage[p].length === 0) {
        suspect.add(p);
      }
    }
  }

  for (const [pStr, pageItems] of Object.entries(byPage)) {
    const p = parseInt(pStr);
    if (pageItems.length === 0) continue;
    const cats = pageItems.map(it => (it.category || "").toLowerCase());
    const hasPipe = cats.includes("pipe");
    const hasFittings = cats.some(c => ["elbow", "tee", "flange", "reducer", "cap", "valve"].includes(c));
    // Strong signal: fittings present but no pipe row — row 1 was almost
    // certainly skipped. Real fitting-only pages exist but are rare.
    if (hasFittings && !hasPipe) {
      suspect.add(p);
      continue;
    }
    // Secondary signal: itemNo gap. Most BOMs start at NO. 1. If our lowest
    // extracted itemNo is > 1, the rows above it were dropped. Skip if the
    // AI did not return itemNo at all (older code path).
    const itemNos = pageItems.map(it => it.itemNo).filter((n: any) => typeof n === "number");
    if (itemNos.length > 0) {
      const minItemNo = Math.min(...itemNos);
      if (minItemNo > 1) suspect.add(p);
    }
  }
  return Array.from(suspect).sort((a, b) => a - b);
}

async function pipeRowRecovery(items: any[], pdfPath: string, startPage: number): Promise<{ recoveredCount: number; pagesScanned: number }> {
  if (!fs.existsSync(pdfPath)) return { recoveredCount: 0, pagesScanned: 0 };
  const client = getAnthropicClient();
  if (!client || !getUserApiKey()) return { recoveredCount: 0, pagesScanned: 0 };

  // Group existing pipe items by source page so we know what's already there
  const existingPipesByPage: Record<number, any[]> = {};
  const allPagesWithItems = new Set<number>();
  for (const it of items) {
    const pg = it.sourcePage;
    if (typeof pg !== "number") continue;
    allPagesWithItems.add(pg);
    const cat = (it.category || "").toLowerCase();
    if (cat === "pipe") {
      if (!existingPipesByPage[pg]) existingPipesByPage[pg] = [];
      existingPipesByPage[pg].push(it);
    }
  }

  if (allPagesWithItems.size === 0) return { recoveredCount: 0, pagesScanned: 0 };

  console.log(`  Pipe-row recovery: scanning ${allPagesWithItems.size} page(s) for missed pipe rows...`);

  let recoveredCount = 0;
  let pagesScanned = 0;
  const tmpDir = path.join(RENDER_DIR, `pipe_recovery_${Date.now()}`);
  fs.mkdirSync(tmpDir, { recursive: true });

  try {
    for (const globalPage of allPagesWithItems) {
      const localPage = globalPage - startPage + 1;
      if (localPage < 1) continue;

      // Render the page at 250 DPI (good legibility, manageable size)
      const pageImg = path.join(tmpDir, `p${globalPage}.png`);
      try {
        await execFileAsync("pdftoppm", [
          "-r", "250", "-png", "-f", String(localPage), "-l", String(localPage),
          pdfPath, path.join(tmpDir, `p${globalPage}_raw`),
        ], { maxBuffer: 60 * 1024 * 1024, timeout: 60000 });
        const candidates = fs.readdirSync(tmpDir).filter(f => f.startsWith(`p${globalPage}_raw`) && f.endsWith(".png"));
        if (candidates.length === 0) continue;
        fs.renameSync(path.join(tmpDir, candidates[0]), pageImg);
      } catch (renderErr: any) {
        console.warn(`    Page ${globalPage} render failed: ${renderErr.message?.substring(0, 80)}`);
        continue;
      }

      if (!fs.existsSync(pageImg)) continue;

      // Build a list of existing pipe items so Claude doesn't re-report them
      const existing = existingPipesByPage[globalPage] || [];
      const existingDesc = existing.length > 0
        ? existing.map((p, i) => `[${i}] size=${p.size}, qty=${p.quantity}, desc=${(p.description || "").substring(0, 60)}`).join("\n")
        : "(none extracted yet for this page)";

      const prompt = `You are auditing the BOM table on a piping isometric drawing for missed PIPE rows.

Look at the BOM table on this page. The BOM has columns: NO. | QTY | SIZE | DESCRIPTION.

For EACH pipe row in the BOM (description starts with 'PIPE,' or contains 'PIPE'), list:
  - The size (e.g. "4\"")
  - The QTY exactly as printed (e.g. "38'-1\"", "25'-5\"", "4'", "6.5")
  - The full description

We already have these pipe rows extracted from this page:
${existingDesc}

Report ONLY pipe rows that are in the BOM but NOT in the existing list above. Do not duplicate.

If there are no missed pipe rows, return an empty array.

Return ONLY valid JSON (no markdown):
{
  "missing_pipes": [
    { "size": "4\"", "qty": "38'-1\"", "description": "PIPE, PE, SCH 40, ASTM A53, TYPE E, GRD B", "section": "SHOP" }
  ]
}

Rules:
- Only report pipe rows you can clearly see in the BOM table.
- Always include the unit marker on qty ("38'-1\"", "4'", etc.).
- If you're unsure whether a row is missed, do NOT report it.
- 'section' is "SHOP" or "FIELD" depending on which sub-table the row is in.`;

      try {
        const imgData = fs.readFileSync(pageImg);
        const b64 = imgData.toString("base64");
        const msg = await client.messages.create({
          model: "claude-sonnet-4-5",
          max_tokens: 2048,
          temperature: 0,
          messages: [{
            role: "user",
            content: [
              { type: "image" as const, source: { type: "base64" as const, media_type: "image/png", data: b64 } },
              { type: "text" as const, text: prompt },
            ],
          }],
        });

        const responseText = msg.content[0].type === "text" ? msg.content[0].text : "";
        let jsonStr = responseText.trim();
        if (jsonStr.startsWith("```")) jsonStr = jsonStr.replace(/^```(?:json)?\n?/, "").replace(/\n?```$/, "");

        const parsed = JSON.parse(jsonStr);
        const missing = Array.isArray(parsed.missing_pipes) ? parsed.missing_pipes : [];
        pagesScanned++;

        for (const mp of missing) {
          const size = String(mp.size || "").trim();
          const qtyStr = String(mp.qty || "").trim();
          const desc = String(mp.description || "PIPE").trim();
          if (!size || !qtyStr) continue;

          // Parse the qty as a pipe length
          const parsedLen = parsePipeLength(qtyStr);
          if (parsedLen === null || parsedLen <= 0) continue;

          // Build a takeoff item that looks like the rest
          const drawingNumber = (existing[0] && (existing[0] as any).drawingNumber) || null;
          const newItem: any = {
            id: randomUUID(),
            lineNumber: items.length + 1,
            discipline: "mechanical",
            category: "pipe",
            description: desc.substring(0, 250),
            size,
            quantity: parsedLen,
            unit: "LF",
            spec: extractSpec(desc) || undefined,
            material: extractMaterial(desc) || undefined,
            schedule: extractSchedule(desc) || undefined,
            rating: extractRating(desc) || undefined,
            notes: `Sheet ${globalPage}${drawingNumber ? " | " + drawingNumber : ""} (${(mp.section || "SHOP").toUpperCase()})`,
            sourcePage: globalPage,
            drawingNumber,
            installLocation: (mp.section || "SHOP").toUpperCase() === "FIELD" ? "field" as const : "shop" as const,
            confidence: "low" as const,
            confidenceScore: 45,
            confidenceNotes: `\u26a0 Recovered missed pipe row \u2014 not in initial extraction. Verify against drawing.`,
            reviewStatus: "unreviewed" as const,
          };
          items.push(newItem);
          recoveredCount++;
        }
      } catch (err: any) {
        console.warn(`    Page ${globalPage} pipe recovery failed: ${err.message?.substring(0, 80)}`);
      }
    }
  } finally {
    try { fs.rmSync(tmpDir, { recursive: true, force: true }); } catch {}
  }

  if (recoveredCount > 0) {
    console.log(`  Pipe-row recovery: added ${recoveredCount} missed pipe row(s) (flagged low confidence) across ${pagesScanned} page(s) scanned`);
  } else {
    console.log(`  Pipe-row recovery: no missed rows found across ${pagesScanned} page(s) scanned`);
  }
  return { recoveredCount, pagesScanned };
}

async function pipeQtyBestGuessRetry(items: any[], pdfPath: string, startPage: number): Promise<{ guessedCount: number }> {
  if (!fs.existsSync(pdfPath)) return { guessedCount: 0 };
  const client = getAnthropicClient();
  if (!client || !getUserApiKey()) return { guessedCount: 0 };

  // Find pipe rows that need a best-guess (qty 0 OR explicitly flagged for review)
  const unread = items.filter(it => {
    const cat = (it.category || "").toLowerCase();
    if (cat !== "pipe") return false;
    if (it.quantity > 0) return false;
    return true;
  });
  if (unread.length === 0) return { guessedCount: 0 };

  console.log(`  Pipe-qty best-guess retry: ${unread.length} unread pipe row(s) to estimate from drawings...`);

  // Group by sourcePage so we render each page once
  const byPage: Record<number, any[]> = {};
  for (const it of unread) {
    const p = it.sourcePage ?? 0;
    if (!byPage[p]) byPage[p] = [];
    byPage[p].push(it);
  }

  let guessedCount = 0;
  const tmpDir = path.join(RENDER_DIR, `pipe_retry_${Date.now()}`);
  fs.mkdirSync(tmpDir, { recursive: true });

  try {
    for (const [pageStr, pageItems] of Object.entries(byPage)) {
      const globalPage = parseInt(pageStr);
      const localPage = globalPage - startPage + 1;
      if (localPage < 1) continue;

      // Render the full page at 300 DPI
      const pageImg = path.join(tmpDir, `p${globalPage}.png`);
      try {
        await execFileAsync("pdftoppm", [
          "-r", "300", "-png", "-f", String(localPage), "-l", String(localPage),
          pdfPath, path.join(tmpDir, `p${globalPage}_raw`),
        ], { maxBuffer: 60 * 1024 * 1024, timeout: 60000 });
        // pdftoppm appends a page-number suffix; find the produced file
        const candidates = fs.readdirSync(tmpDir).filter(f => f.startsWith(`p${globalPage}_raw`) && f.endsWith(".png"));
        if (candidates.length === 0) continue;
        fs.renameSync(path.join(tmpDir, candidates[0]), pageImg);
      } catch (renderErr: any) {
        console.warn(`    Page ${globalPage} render failed: ${renderErr.message?.substring(0, 80)}`);
        continue;
      }

      if (!fs.existsSync(pageImg)) continue;

      // Build a single prompt asking for length estimates for all unread pipes on this page
      const pipeList = pageItems.map((it, idx) => {
        return `[${idx}] size=${it.size || "?"}, description=${(it.description || "").substring(0, 80)}`;
      }).join("\n");

      const prompt = `You are estimating pipe lengths from a piping isometric drawing.

I have these PIPE rows from the BOM table where the QTY cell was unreadable:

${pipeList}

Look at the drawing graphic and the dimension callouts (e.g. '5'-3"', '22'-6"', or running dimensions along pipe segments). For each pipe row above, estimate the TOTAL length in feet-inches format.

If you can read the BOM QTY cell after a closer look, use that value. Otherwise estimate from the drawing line work + dimension callouts.

Return ONLY valid JSON (no markdown):
{
  "estimates": [
    { "index": 0, "length": "5'-3\"", "source": "bom" | "drawing" | "unreadable", "confidence": "high" | "medium" | "low" }
  ]
}

Rules:
- 'source": "bom"' means you read the QTY cell directly. Use this if you can read it now.
- 'source": "drawing"' means you estimated from drawing line work / dimension callouts.
- 'source": "unreadable"' means you genuinely cannot determine the length. Set length to "" in that case.
- Always include a unit marker on length (5'-3", 22'-6", 0'-8", etc.) — NEVER a bare integer.
- Be conservative: if you're unsure, mark unreadable.`;

      try {
        const imgData = fs.readFileSync(pageImg);
        const b64 = imgData.toString("base64");
        const msg = await client.messages.create({
          model: "claude-sonnet-4-5",
          max_tokens: 4096,
          temperature: 0,
          messages: [{
            role: "user",
            content: [
              { type: "image" as const, source: { type: "base64" as const, media_type: "image/png", data: b64 } },
              { type: "text" as const, text: prompt },
            ],
          }],
        });

        const responseText = msg.content[0].type === "text" ? msg.content[0].text : "";
        let jsonStr = responseText.trim();
        if (jsonStr.startsWith("```")) jsonStr = jsonStr.replace(/^```(?:json)?\n?/, "").replace(/\n?```$/, "");
        const parsed = JSON.parse(jsonStr);
        const estimates = Array.isArray(parsed.estimates) ? parsed.estimates : [];

        for (const est of estimates) {
          const idx = typeof est.index === "number" ? est.index : -1;
          if (idx < 0 || idx >= pageItems.length) continue;
          const item = pageItems[idx];
          if (est.source === "unreadable" || !est.length) continue;

          const parsedLen = parsePipeLength(String(est.length));
          if (!parsedLen || parsedLen <= 0) continue;

          item.quantity = parsedLen;
          item.unit = "LF";
          item.confidence = "low";
          item.confidenceScore = est.source === "bom" ? 60 : 40;
          // Persist note via BOTH _validationNotes AND confidenceNotes
          // so it survives downstream cleanup and shows up in the UI.
          item._validationNotes = item._validationNotes || [];
          item._validationNotes = item._validationNotes.filter((n: string) =>
            !n.includes("Pipe quantity flagged for review") && !n.includes("matches pipe SIZE"));
          const sourceTxt = est.source === "bom" ? "BOM cell on closer read" : "drawing dimensions / line work estimate";
          const flagNote = `\u26a0 Best-guess pipe length (${parsedLen} LF) from ${sourceTxt}. VERIFY against drawing.`;
          item._validationNotes.push(flagNote);
          // Strip the old 'flagged for review' note from confidenceNotes too
          let existing = item.confidenceNotes || "";
          existing = existing.replace(/[^|]*Pipe quantity flagged for review[^|]*(\|\s*)?/g, "").trim();
          existing = existing.replace(/^\|\s*|\s*\|$/g, "").trim();
          item.confidenceNotes = existing ? `${existing} | ${flagNote}` : flagNote;
          item.reviewStatus = "unreviewed";
          guessedCount++;
        }
      } catch (err: any) {
        console.warn(`    Page ${globalPage} retry failed: ${err.message?.substring(0, 80)}`);
      }
    }
  } finally {
    try { fs.rmSync(tmpDir, { recursive: true, force: true }); } catch {}
  }

  if (guessedCount > 0) {
    console.log(`  Pipe-qty retry: filled ${guessedCount} pipe length(s) with best-guess values (flagged low confidence)`);
  }
  return { guessedCount };
}

function applyHistoricalPatterns(items: any[]): any[] {
  // EMERGENCY DISABLE: Pattern auto-apply was causing ~85%% of items to get
  // qty rewritten to a single value (e.g., qty=5 across 151 of 173 rows on
  // Area 400). The DB clearly has a bad/over-eager pattern. Returning items
  // unchanged until we add a UI to inspect and clear patterns.
  return items;

  // Original logic (kept as reference, currently unreachable):
  // const patterns = storage.getAutoApplyPatterns();
  // if (patterns.length === 0) return items;
  //
  // let correctionCount = 0;
  // for (const item of items) {
  //   for (const pattern of patterns) {
  //     if (pattern.field === "size" && item.size === pattern.original_value) {
  //       const oldSize = item.size;
  //       item.size = pattern.corrected_value;
  //       item.notes = (item.notes || "") + ` | Auto-corrected size: ${oldSize} \u2192 ${pattern.corrected_value} (learned)`;
  //       correctionCount++;
  //     }
  //     if (pattern.field === "quantity" && String(item.quantity) === pattern.original_value) {
  //       item.quantity = parseFloat(pattern.corrected_value) || item.quantity;
  //       correctionCount++;
  //     }
  //   }
  // }
  //
  // if (correctionCount > 0) {
  //   console.log(`  Applied ${correctionCount} historical pattern corrections`);
  // }
  // return items;
}

// ============================================================
// CONTINUATION PAGE DEDUP (Spec Item 5 - Improved)
// ============================================================

function dedupSameDrawingNumber(items: any[]): { items: any[]; dedupCount: number; dupGroups: number } {
  // SAME-DRAWING-NUMBER DEDUP — If two or more pages share the same drawing
  // number (e.g., a multi-revision package with both REV B and REV C of the
  // same spool, or the same drawing appearing twice in the PDF), keep one
  // page's items and mark the others as duplicates.
  //
  // Rule:
  //   1. Group pages by normalized drawing number.
  //   2. For groups with 2+ pages, pick the page with the HIGHEST page number
  //      as the keeper (most recent revision in a typical engineering package
  //      is usually placed later in the binder).
  //   3. Mark all items on the OTHER pages as _dedupCandidate with a clear note.
  //
  // This is conservative: only triggers on EXACT drawing-number match.
  //
  // Returns: items array (unchanged), dedupCount, dupGroups

  if (items.length === 0) return { items, dedupCount: 0, dupGroups: 0 };

  // Build page → drawing number map
  const pageDrawing: Record<number, string> = {};
  for (const item of items) {
    if (item.sourcePage && item.drawingNumber) {
      const dwg = String(item.drawingNumber).replace(/[\s"'\u2019\u201D]/g, "").toLowerCase();
      if (dwg && dwg !== "null" && dwg !== "undefined") {
        pageDrawing[item.sourcePage] = dwg;
      }
    }
  }

  // Group pages by drawing number
  const drawingPages: Record<string, number[]> = {};
  for (const [pageStr, dwg] of Object.entries(pageDrawing)) {
    const page = parseInt(pageStr);
    if (!drawingPages[dwg]) drawingPages[dwg] = [];
    drawingPages[dwg].push(page);
  }

  let dedupCount = 0;
  let dupGroups = 0;

  for (const [dwg, pages] of Object.entries(drawingPages)) {
    if (pages.length < 2) continue;
    dupGroups++;
    pages.sort((a, b) => a - b);
    const keeperPage = pages[pages.length - 1]; // keep the LAST occurrence
    const dropPages = pages.slice(0, -1);
    for (const item of items) {
      if (item.sourcePage && dropPages.includes(item.sourcePage)) {
        item._dedupCandidate = true;
        item.dedupNote = `Same-drawing-number duplicate — drawing "${dwg}" appears on pages ${pages.join(", ")}, items on this page (${item.sourcePage}) are duplicates of the kept page (${keeperPage}).`;
        dedupCount++;
      }
    }
  }

  if (dupGroups > 0) {
    console.log(`  Same-drawing dedup: ${dupGroups} duplicate drawing(s) found, marked ${dedupCount} items as candidates`);
  }
  return { items, dedupCount, dupGroups };
}

function dedupContinuationPages(items: any[]): any[] {
  // CONTINUATION-BASED DEDUP — Only removes items that are explicitly marked
  // as being at a continuation point AND appear on both the "from" and "to" sheets.
  // This uses the AI-extracted continuation references, NOT heuristic matching.
  //
  // Rule: When two connected sheets both have the same fitting at their shared
  // continuation point, keep the item on the "from" sheet and remove it from
  // the "to" sheet.

  // Find items marked as atContinuation
  const continuationItems = items.filter(i => i.atContinuation === true || i.atContinuation === "true");
  if (continuationItems.length === 0) return items;

  // Build a map of continuation connections: page -> [{direction, drawing, sheet}]
  // This info was stored on items during extraction as _continuations on the first item of each page
  const pageConnections: Record<number, { direction: string; drawing: string; sheet: number }[]> = {};
  for (const item of items) {
    if (item._continuations && Array.isArray(item._continuations)) {
      const page = item.sourcePage || 0;
      if (!pageConnections[page]) pageConnections[page] = [];
      pageConnections[page].push(...item._continuations);
    }
  }

  // Build drawing number to page mapping for cross-referencing
  const drawingToPage: Record<string, number> = {};
  for (const item of items) {
    if (item.drawingNumber && item.sourcePage) {
      drawingToPage[item.drawingNumber] = item.sourcePage;
    }
  }

  // For each page with a "from" continuation, find the receiving page and dedup shared items
  let dedupCount = 0;
  const processedPairs = new Set<string>();

  for (const [pageStr, connections] of Object.entries(pageConnections)) {
    const page = parseInt(pageStr);
    for (const conn of connections) {
      // Process "from" connections — this page receives from another drawing
      if (conn.direction !== "from") continue;

      // Find the source page by drawing number reference
      // The conn.drawing is the drawing number of the sheet we're connected FROM
      let sourcePage: number | null = null;

      // Method 1: Match by drawing number in the continuation reference
      if (conn.drawing) {
        sourcePage = drawingToPage[conn.drawing] || null;
        // Also try partial matching (drawing numbers may have slight format differences)
        if (!sourcePage) {
          const connDwgClean = conn.drawing.replace(/["'\s]/g, "").toLowerCase();
          for (const [dwg, pg] of Object.entries(drawingToPage)) {
            if (dwg.replace(/["'\s]/g, "").toLowerCase() === connDwgClean) {
              sourcePage = pg;
              break;
            }
          }
        }
      }

      // Method 2: Look for a page that has a "to" connection referencing back
      if (!sourcePage) {
        const toPageEntry = Object.entries(pageConnections).find(([p, conns]) => {
          return parseInt(p) !== page && conns.some(c => c.direction === "to");
        });
        if (toPageEntry) sourcePage = parseInt(toPageEntry[0]);
      }

      if (!sourcePage || sourcePage === page) continue;

      const pairKey = `${Math.min(page, sourcePage)}-${Math.max(page, sourcePage)}`;
      if (processedPairs.has(pairKey)) continue;
      processedPairs.add(pairKey);

      // This page ("from" direction = receiving continuation) has the duplicate.
      // The source page (the one sending the continuation) keeps the item.
      // Mark continuation items on THIS page as dedup candidates.
      const thisPageContItems = continuationItems.filter(i => i.sourcePage === page);
      const sourcePageContItems = continuationItems.filter(i => i.sourcePage === sourcePage);

      for (const thisItem of thisPageContItems) {
        // Match by category + size (description may vary slightly between sheets)
        const matchKey = `${thisItem.category}|${thisItem.size}`;
        const matchingSourceItem = sourcePageContItems.find(si =>
          `${si.category}|${si.size}` === matchKey
        );

        if (matchingSourceItem) {
          thisItem._dedupCandidate = true;
          thisItem.dedupNote = `Continuation duplicate — this fitting is shared with page ${sourcePage} (${conn.drawing || "connected sheet"}) and counted there.`;
          dedupCount++;
        }
      }
    }
  }

  if (dedupCount > 0) {
    console.log(`  Continuation dedup: marked ${dedupCount} shared connection items (text-referenced, not heuristic)`);
  }
  return items;
}

// ============================================================
// ENHANCED CONFIDENCE SCORING (Spec Item 4)
// ============================================================

function computeConfidenceScore(item: any, pdfQuality?: "vector" | "clean_scan" | "poor_scan"): { confidence: "high" | "medium" | "low"; confidenceScore: number; confidenceNotes: string } {
  let score = 100;
  const notes: string[] = [];

  // Size validation — only penalize truly missing or invalid sizes
  if (!item.size || item.size === "N/A" || item.size.trim() === "") {
    score -= 30;
    notes.push("Missing size");
  } else if (item.category !== "bolt" && !isValidNPS(item.size) && !/x/i.test(item.size)) {
    score -= 10;
    notes.push("Size not standard NPS");
  }

  // Description quality — only penalize very short or empty descriptions
  if (!item.description || item.description.length < 5) {
    score -= 25;
    notes.push("Description too short or missing");
  }
  // ASME/ASTM spec reference is nice to have but not required — many valid
  // BOM entries use abbreviated descriptions without full spec callouts

  // Quantity plausibility — only flag truly implausible values
  const isPipe = item.category === "pipe";
  if (isPipe && item.quantity > 1000) {
    score -= 20;
    notes.push("Pipe qty >1000 LF — verify");
  }
  if (!isPipe && item.quantity > 500) {
    score -= 10;
    notes.push("Large qty — verify count");
  }
  if (item.quantity <= 0) {
    score -= 35;
    notes.push("Qty is 0 or negative");
  }

  // Category consistency
  if (item.category === "other") {
    score -= 10;
    notes.push("Unclassified category");
  }

  // Incorporate validation flags from validateExtractedItems
  if (item._validationFlag === "low") {
    score = Math.min(score, 45);
  } else if (item._validationFlag === "medium") {
    score = Math.min(score, 70);
  }

  // Add validation notes
  if (item._validationNotes && item._validationNotes.length > 0) {
    notes.push(...item._validationNotes);
  }

  // Size warning from autoCorrect
  if (item._sizeWarning) {
    notes.push(item._sizeWarning);
    score -= 10;
  }

  // Inches-vs-feet auto-correction happened — slightly lower confidence
  const itemNotes = item.notes || "";
  if (itemNotes.includes("Auto-corrected") && itemNotes.includes("inches")) {
    score -= 5;
    notes.push("Pipe length auto-corrected (inches assumed)");
  }

  // If item came from a poor_scan page, penalize
  if (pdfQuality === "poor_scan") {
    score -= 20;
    notes.push("Poor scan quality");
  }

  // Clamp score
  score = Math.max(0, Math.min(100, score));

  // Confidence thresholds — most well-extracted items should be "high"
  let confidence: "high" | "medium" | "low";
  if (score >= 75) confidence = "high";
  else if (score >= 50) confidence = "medium";
  else confidence = "low";

  return { confidence, confidenceScore: score, confidenceNotes: notes.length > 0 ? notes.join("; ") : "" };
}

function mapStructuralCategory(cat: string): string {
  const validCategories = ["wide_flange", "hss_tube", "angle", "channel", "plate", "base_plate", "column", "bracing", "embed_plate", "bolt", "weld", "clip_angle", "gusset_plate", "anchor_bolt", "footing", "grade_beam", "concrete_wall", "slab", "concrete_column", "rebar", "wire_mesh", "dowel", "other"];
  const lower = cat.toLowerCase().replace(/[^a-z_]/g, "_");
  if (validCategories.includes(lower)) return lower;
  if (/wide[_\s]?flange|w[\d]+x|beam/i.test(cat)) return "wide_flange";
  if (/hss|tube[_\s]?steel/i.test(cat)) return "hss_tube";
  if (/angle|^l\d/i.test(cat)) return "angle";
  if (/channel|^c\d|^mc\d/i.test(cat)) return "channel";
  if (/^plate|^pl\s/i.test(cat)) return "plate";
  if (/base[_\s]?plate/i.test(cat)) return "base_plate";
  if (/column|pier/i.test(cat)) return "column";
  if (/brac/i.test(cat)) return "bracing";
  if (/bolt(?!_)/i.test(cat)) return "bolt";
  if (/weld/i.test(cat)) return "weld";
  if (/anchor[_\s]?bolt|anchor[_\s]?rod/i.test(cat)) return "anchor_bolt";
  if (/footing|ftg/i.test(cat)) return "footing";
  if (/grade[_\s]?beam/i.test(cat)) return "grade_beam";
  if (/slab|sog/i.test(cat)) return "slab";
  if (/rebar|reinfor/i.test(cat)) return "rebar";
  if (/wire[_\s]?mesh|wwf/i.test(cat)) return "wire_mesh";
  if (/dowel/i.test(cat)) return "dowel";
  return "other";
}

function mapCivilCategory(cat: string): string {
  if (/storm[_\s]?pipe|storm[_\s]?drain|rcp|hdpe[_\s]?pipe/i.test(cat)) return "storm_pipe";
  if (/sanitary[_\s]?sewer|sewer[_\s]?pipe/i.test(cat)) return "sewer_pipe";
  if (/water[_\s]?line|water[_\s]?main|di[_\s]?pipe/i.test(cat)) return "water_pipe";
  if (/gas[_\s]?line|gas[_\s]?pipe/i.test(cat)) return "gas_pipe";
  if (/manhole/i.test(cat)) return "manhole";
  if (/catch[_\s]?basin|inlet/i.test(cat)) return "catch_basin";
  if (/fire[_\s]?hydrant/i.test(cat)) return "fire_hydrant";
  if (/valve/i.test(cat)) return "valve";
  if (/fitting|bend|tee|reducer/i.test(cat)) return "fitting";
  if (/earthwork|cut|fill|excav/i.test(cat)) return "earthwork";
  if (/backfill|import|export/i.test(cat)) return "backfill";
  if (/asphalt|hma|paving/i.test(cat)) return "paving";
  if (/concrete[_\s]?paving|sidewalk|flatwork/i.test(cat)) return "concrete_paving";
  if (/base[_\s]?course|subbase/i.test(cat)) return "base_course";
  if (/curb|gutter/i.test(cat)) return "curb_gutter";
  if (/retaining[_\s]?wall/i.test(cat)) return "retaining_wall";
  if (/silt[_\s]?fence/i.test(cat)) return "silt_fence";
  if (/seeding|sodding|seed|sod/i.test(cat)) return "seeding";
  if (/fencing|fence/i.test(cat)) return "fencing";
  if (/pipe/i.test(cat)) return "storm_pipe";
  return "other";
}

// ============================================================
// PROCESS CHUNKS
// ============================================================

// ============================================================
// TWO-PHASE SPLIT: Render (Phase 1) + Extract (Phase 2)
// ============================================================

// Types for pre-rendered chunk results
type RenderedMechanicalChunk = {
  jobDir: string;
  startPage: number;
  endPage: number;
  hasRevisions: boolean;
  // BOM-only path images (non-revision)
  bomPageImages?: { pageNum: number; imagePath: string; tesseractText: string }[];
  // Cloud detection path images (revision)
  cloudPageImages?: { pageNum: number; bomImagePath: string; fullImagePath: string; tesseractText: string }[];
};

type RenderedStructuralChunk = {
  jobDir: string;
  startPage: number;
  endPage: number;
  pageImages: { pageNum: number; imagePath: string; tesseractText: string }[];
};

type RenderedCivilChunk = {
  jobDir: string;
  startPage: number;
  endPage: number;
  pageImages: { pageNum: number; imagePath: string; tesseractText: string }[];
};

// --- MECHANICAL RENDER (Phase 1) ---
async function renderMechanicalChunk(
  chunkPath: string,
  startPage: number,
  endPage: number,
  hasRevisions: boolean = false
): Promise<RenderedMechanicalChunk> {
  const chunkPageCount = endPage - startPage + 1;

  if (hasRevisions) {
    const renderResult = await renderBomWithFullPages(chunkPath, chunkPageCount);
    return {
      jobDir: renderResult.jobDir,
      startPage,
      endPage,
      hasRevisions: true,
      cloudPageImages: renderResult.pageImages,
    };
  } else {
    const renderResult = await renderCroppedBomImages(chunkPath, chunkPageCount);
    return {
      jobDir: renderResult.jobDir,
      startPage,
      endPage,
      hasRevisions: false,
      bomPageImages: renderResult.pageImages,
    };
  }
}

// --- MECHANICAL EXTRACT (Phase 2) ---
async function extractMechanicalChunk(
  rendered: RenderedMechanicalChunk,
  onPageComplete?: (pagesProcessed: number) => void,
  verifyExtraction: boolean = true
): Promise<{ items: any[]; metadata: any; authFailures: number }> {
  const { startPage, hasRevisions } = rendered;
  const items: any[] = [];
  let metadata: any = {};
  let authFailures = 0;

  // Helper: parse and correct pipe quantity from raw AI output
  function parseMechanicalQty(rawQty: string, category: string, description: string, size?: string): { qty: number; notes: string[] } {
    const rq = rawQty.trim();
    let numericQty: number;
    const notes: string[] = [];
    if (/['''\u2019]/.test(rq)) {
      numericQty = parsePipeLength(rq).feet;
    } else if (/[""\u201D]/.test(rq)) {
      numericQty = parsePipeLength(rq).feet;
    } else {
      numericQty = parseFloat(rq) || 1;
    }
    const isPipe = category === "pipe";
    let qty = isPipe ? Math.round(Math.max(numericQty, 0) * 100) / 100 : Math.max(1, Math.round(numericQty));
    if (isPipe) {
      const correction = correctPipeLengthIfInches(qty, rq, description, { size: size || "" });
      if (correction.wasCorrection) qty = correction.correctedQty;
      if (correction.note) notes.push(correction.note);
    }
    return { qty, notes };
  }

  if (hasRevisions && rendered.cloudPageImages) {
    const pageImages = rendered.cloudPageImages;
    if (pageImages.length === 0) return { items, metadata, authFailures };

    const adjustedPageImages = pageImages.map(pi => ({ ...pi, globalPageNum: pi.pageNum + startPage - 1 }));

    const cloudResult = await extractWithCloudDetection(pageImages, MECHANICAL_CLOUD_PROMPT, (completed, total) => {
      if (onPageComplete) onPageComplete(startPage - 1 + completed);
    });
    const visionResults = cloudResult.results;
    authFailures += cloudResult.authFailures;

    if (onPageComplete) {
      onPageComplete(startPage - 1 + pageImages.length);
    }

    // BOM item count cross-check per page (cloud path)
    for (const [pageNum, pageData] of visionResults) {
      const crossCheckWarnings = crossCheckItemCount(pageData.items, pageNum + startPage - 1);
      for (const w of crossCheckWarnings) console.warn(`  ${w}`);
      if (crossCheckWarnings.length > 0 && pageData.items.length > 0) {
        pageData.items[0]._crossCheckWarning = crossCheckWarnings.join("; ");
      }
    }

    for (const page of adjustedPageImages) {
      const cloudPageData = visionResults.get(page.pageNum) || { items: [] };
      const entries = cloudPageData.items;
      const pageDrawingNumber = cloudPageData.drawingNumber || null;
      const pageWeldCount = cloudPageData.weldCount || null;
      const pageContinuations = cloudPageData.continuations || [];
      for (const entry of entries) {
        if (!entry.description || entry.description.length < 3) continue;
        const rawQty = String(entry.qty || "1").trim();
        const category = detectMechanicalCategory(entry.description);
        const isPipe = category === "pipe";
        const { qty: parsedQty, notes: qtyNotes } = parseMechanicalQty(rawQty, category, entry.description, entry.size);
        const validatedSize = validateSize(String(entry.size || "N/A"), category);
        const extraNotes = qtyNotes.length > 0 ? " | " + qtyNotes.join(" | ") : "";
        let rawItem: any = {
          category,
          description: entry.description.substring(0, 250),
          size: validatedSize,
          quantity: parsedQty,
          unit: isPipe ? "LF" : (/\bLF\b|\bFT\b/i.test(entry.description) ? "LF" : "EA"),
          spec: extractSpec(entry.description) || undefined,
          material: extractMaterial(entry.description) || undefined,
          schedule: extractSchedule(entry.description) || undefined,
          rating: extractRating(entry.description) || undefined,
          notes: `Sheet ${page.globalPageNum}${pageDrawingNumber ? " | " + pageDrawingNumber : ""} (${entry.section || "SHOP"})${extraNotes}`,
          sourcePage: page.globalPageNum,
          itemNo: typeof entry.itemNo === "number" ? entry.itemNo : (parseInt(String(entry.itemNo || "")) || undefined),
          drawingNumber: pageDrawingNumber,
          revisionClouded: entry.clouded === true,
          cloudConfidence: entry.cloudConfidence ?? (entry.clouded === true ? 80 : 100),
        installLocation: (entry.section || "SHOP").toUpperCase() === "FIELD" ? "field" as const : "shop" as const,
        valveType: category === "valve" ? detectValveType(entry.description) : undefined,
        smallBoreRollup: isSmallBoreRollup({ size: validatedSize, description: entry.description, category }),
        atContinuation: entry.atContinuation === true || entry.atContinuation === "true" || false,
        };
        if (qtyNotes.length > 0) {
          rawItem._validationNotes = rawItem._validationNotes || [];
          rawItem._validationNotes.push(...qtyNotes);
        }
        rawItem = autoCorrectItem(rawItem);
        // Store continuation references and weld counts on first item of each page
        const isFirstItemForPage = items.filter(i => i.sourcePage === page.globalPageNum).length === 0;
        if (isFirstItemForPage) {
          if (pageContinuations.length > 0) {
            rawItem._continuations = pageContinuations;
          }
          if (pageWeldCount) {
            rawItem._visualWeldCount = pageWeldCount;
          }
          if (entry._crossCheckWarning) {
            rawItem._validationNotes = rawItem._validationNotes || [];
            rawItem._validationNotes.push(entry._crossCheckWarning);
          }
        }
        items.push(rawItem);
      }
    }

    if (startPage === 1 && pageImages.length > 0) {
      const firstText = pageImages[0].tesseractText || "";
      const lineMatch = firstText.match(/LINE\s*(?:NO\.?|#)?[:\s]*([\w][\w\-\.\/]+)/i);
      if (lineMatch) metadata.lineNumber = lineMatch[1];
      const areaMatch = firstText.match(/(?:AREA|ZONE)[:\s]*([\w][\w\s]*?)(?=\s{2,}|$)/im);
      if (areaMatch) metadata.area = areaMatch[1].trim();
      const revMatch = firstText.match(/(?:REV(?:ISION)?)[:\s]*([\w\d\.]+)/i);
      if (revMatch) metadata.revision = revMatch[1];
      const dateMatch = firstText.match(/(\d{1,2}[\/\-]\d{1,2}[\/\-]\d{2,4})/);
      if (dateMatch) metadata.drawingDate = dateMatch[1];
    }
  } else if (!hasRevisions && rendered.bomPageImages) {
    const pageImages = rendered.bomPageImages;
    if (pageImages.length === 0) return { items, metadata, authFailures };

    const adjustedPageImages = pageImages.map(pi => ({ ...pi, globalPageNum: pi.pageNum + startPage - 1 }));
    const extractResult = await extractWithVision(pageImages, MECHANICAL_PROMPT, "mechanical");
    const visionResults = extractResult.results;
    authFailures += extractResult.authFailures;

    if (onPageComplete) {
      onPageComplete(startPage - 1 + pageImages.length);
    }

    let verifiedResults = visionResults;
    if (verifyExtraction) {
      // Double-pass verification: send images + extracted items for full cross-check
      console.log(`  Running double-pass verification on ${visionResults.size} pages...`);
      verifiedResults = await verifyExtractionPass(pageImages, visionResults, "mechanical");

      // Field-level correction pass: targeted corrections on specific fields
      const bomImageMap = new Map<number, string>();
      for (const pi of pageImages) {
        bomImageMap.set(pi.pageNum, pi.imagePath);
      }
      console.log(`  Running field-level verification on ${verifiedResults.size} pages...`);
      verifiedResults = await verifyExtractedItems(verifiedResults, bomImageMap);
    }

    // BOM item count cross-check per page
    for (const [pageNum, pageData] of verifiedResults) {
      const crossCheckWarnings = crossCheckItemCount(pageData.items, pageNum + startPage - 1);
      for (const w of crossCheckWarnings) console.warn(`  ${w}`);
      if (crossCheckWarnings.length > 0) {
        // Mark first item on this page with cross-check warning
        if (pageData.items.length > 0) {
          pageData.items[0]._crossCheckWarning = crossCheckWarnings.join("; ");
        }
      }
    }

    for (const page of adjustedPageImages) {
      const pageData = verifiedResults.get(page.pageNum) || { items: [] };
      const entries = pageData.items;
      const pageDrawingNumber = pageData.drawingNumber || null;
      const pageWeldCount = pageData.weldCount || null;
      const pageContinuations = pageData.continuations || [];
      for (const entry of entries) {
        if (!entry.description || entry.description.length < 3) continue;
        const rawQty = String(entry.qty || "1").trim();
        const category = detectMechanicalCategory(entry.description);
        const isPipe = category === "pipe";
        const { qty: parsedQty, notes: qtyNotes } = parseMechanicalQty(rawQty, category, entry.description, entry.size);
        const validatedSize = validateSize(String(entry.size || "N/A"), category);
        const extraNotes = qtyNotes.length > 0 ? " | " + qtyNotes.join(" | ") : "";
        let rawItem: any = {
          category,
          description: entry.description.substring(0, 250),
          size: validatedSize,
          quantity: parsedQty,
          unit: isPipe ? "LF" : (/\bLF\b|\bFT\b/i.test(entry.description) ? "LF" : "EA"),
          spec: extractSpec(entry.description) || undefined,
          material: extractMaterial(entry.description) || undefined,
          schedule: extractSchedule(entry.description) || undefined,
          rating: extractRating(entry.description) || undefined,
          notes: `Sheet ${page.globalPageNum}${pageDrawingNumber ? " | " + pageDrawingNumber : ""} (${entry.section || "SHOP"})${extraNotes}`,
          sourcePage: page.globalPageNum,
          itemNo: typeof entry.itemNo === "number" ? entry.itemNo : (parseInt(String(entry.itemNo || "")) || undefined),
          drawingNumber: pageDrawingNumber,
          revisionClouded: entry.clouded === true,
          cloudConfidence: entry.cloudConfidence ?? (entry.clouded === true ? 80 : 100),
        installLocation: (entry.section || "SHOP").toUpperCase() === "FIELD" ? "field" as const : "shop" as const,
        valveType: category === "valve" ? detectValveType(entry.description) : undefined,
        smallBoreRollup: isSmallBoreRollup({ size: validatedSize, description: entry.description, category }),
        atContinuation: entry.atContinuation === true || entry.atContinuation === "true" || false,
        };
        if (qtyNotes.length > 0) {
          rawItem._validationNotes = rawItem._validationNotes || [];
          rawItem._validationNotes.push(...qtyNotes);
        }
        rawItem = autoCorrectItem(rawItem);
        // Store continuation references and weld counts on first item of each page
        const isFirstItemForPage = items.filter(i => i.sourcePage === page.globalPageNum).length === 0;
        if (isFirstItemForPage) {
          if (pageContinuations.length > 0) {
            rawItem._continuations = pageContinuations;
          }
          if (pageWeldCount) {
            rawItem._visualWeldCount = pageWeldCount;
          }
          // Transfer cross-check warnings to item for downstream processing
          if (entry._crossCheckWarning) {
            rawItem._validationNotes = rawItem._validationNotes || [];
            rawItem._validationNotes.push(entry._crossCheckWarning);
          }
        }
        items.push(rawItem);
      }
    }

    if (startPage === 1 && pageImages.length > 0) {
      const firstText = pageImages[0].tesseractText || "";
      const lineMatch = firstText.match(/LINE\s*(?:NO\.?|#)?[:\s]*([\w][\w\-\.\/]+)/i);
      if (lineMatch) metadata.lineNumber = lineMatch[1];
      const areaMatch = firstText.match(/(?:AREA|ZONE)[:\s]*([\w][\w\s]*?)(?=\s{2,}|$)/im);
      if (areaMatch) metadata.area = areaMatch[1].trim();
      const revMatch = firstText.match(/(?:REV(?:ISION)?)[:\s]*([\w\d\.]+)/i);
      if (revMatch) metadata.revision = revMatch[1];
      const dateMatch = firstText.match(/(\d{1,2}[\/\-]\d{1,2}[\/\-]\d{2,4})/);
      if (dateMatch) metadata.drawingDate = dateMatch[1];
    }
  }

  return { items, metadata, authFailures };
}

// --- BOM ITEM COUNT CROSS-CHECK ---
function crossCheckItemCount(items: any[], pageNum: number): string[] {
  const warnings: string[] = [];

  // Find highest item number from itemNo/lineItem fields
  const itemNumbers = items
    .map(i => parseInt(String(i.itemNo || i.lineItem || 0)))
    .filter(n => !isNaN(n) && n > 0);

  if (itemNumbers.length === 0) return warnings;

  const maxItemNo = Math.max(...itemNumbers);
  const extractedCount = items.length;

  if (maxItemNo > extractedCount + 1) {
    const missing = maxItemNo - extractedCount;
    warnings.push(`Page ${pageNum}: BOM has item numbers up to ${maxItemNo} but only ${extractedCount} extracted — ${missing} items may be missing`);
  }

  // Check for gaps in item numbering
  const numberSet = new Set(itemNumbers);
  const gaps: number[] = [];
  for (let i = 1; i <= maxItemNo; i++) {
    if (!numberSet.has(i)) gaps.push(i);
  }
  if (gaps.length > 0 && gaps.length <= 5) {
    warnings.push(`Page ${pageNum}: Missing item numbers: ${gaps.join(", ")} — verify these items exist in the BOM`);
  }

  return warnings;
}

// --- ADAPTIVE DPI RE-EXTRACTION FOR LOW-CONFIDENCE PAGES ---
async function reExtractLowConfidencePages(
  items: any[],
  pdfPath: string,
  discipline: string,
  startPage: number,
  options?: { extraPages?: number[]; renderMode?: "bom" | "fullpage"; pageLimit?: number }
): Promise<{ reExtractedItems: any[]; reExtractedPages: number[]; warnings: string[] }> {
  const warnings: string[] = [];
  const extraPages = options?.extraPages || [];
  const renderMode = options?.renderMode || "bom";
  const pageLimit = options?.pageLimit ?? 12;

  // Group items by page, calculate average confidence per page
  const pageConfidence: Record<number, { total: number; count: number }> = {};
  for (const item of items) {
    const page = item.sourcePage || 0;
    if (!pageConfidence[page]) pageConfidence[page] = { total: 0, count: 0 };
    pageConfidence[page].total += item.confidenceScore || 50;
    pageConfidence[page].count++;
  }

  // Find pages with average confidence below 70
  const lowConfPages = Object.entries(pageConfidence)
    .filter(([, data]) => data.count > 0 && (data.total / data.count) < 70)
    .map(([page]) => parseInt(page))
    .sort((a, b) => a - b);

  // Merge low-confidence pages with caller-supplied suspect pages (e.g. pages
  // where row 1 is missing or non-PIPE). Dedup and sort.
  const candidatePages = Array.from(new Set([...lowConfPages, ...extraPages])).sort((a, b) => a - b);

  if (candidatePages.length === 0) return { reExtractedItems: items, reExtractedPages: [], warnings };

  // Limit how many pages we re-extract (avoid excessive API cost). Default 12,
  // up from the old hard-coded 5 — row-1 misses on Area 300 hit ~30% of pages
  // and a cap of 5 leaves most of them un-recovered.
  const pagesToReExtract = candidatePages.slice(0, pageLimit);

  // Check that PDF still exists for re-rendering
  if (!fs.existsSync(pdfPath)) {
    for (const p of pagesToReExtract) {
      const conf = pageConfidence[p] ? Math.round(pageConfidence[p].total / pageConfidence[p].count) : null;
      const reason = conf !== null ? `Low avg confidence (${conf})` : `Suspect page (row 1 missing or non-PIPE)`;
      warnings.push(`Page ${p}: ${reason} \u2014 re-extraction skipped (PDF no longer available)`);
    }
    return { reExtractedItems: items, reExtractedPages: [], warnings };
  }

  console.log(`  Re-extracting ${pagesToReExtract.length} suspect/low-confidence page(s) at 300 DPI (${renderMode}): ${pagesToReExtract.join(", ")}`);

  const jobDir = path.join(RENDER_DIR, `reextract_${Date.now()}`);
  fs.mkdirSync(jobDir, { recursive: true });

  // Re-render at 300 DPI (double the normal 150 DPI)
  const reRenderedPages: { pageNum: number; imagePath: string }[] = [];
  for (const pageNum of pagesToReExtract) {
    const localPage = pageNum - startPage + 1;
    try {
      await execFileAsync("pdftoppm", [
        "-r", "300", "-png", "-f", String(localPage), "-l", String(localPage),
        pdfPath, path.join(jobDir, `page_${pageNum}`)
      ], { maxBuffer: 100 * 1024 * 1024, timeout: 60000 });

      // Find the rendered file
      const rendered = fs.readdirSync(jobDir).filter(f => f.startsWith(`page_${pageNum}`) && f.endsWith(".png"));
      if (rendered.length > 0) {
        reRenderedPages.push({ pageNum, imagePath: path.join(jobDir, rendered[0]) });
      }
    } catch (err: any) {
      console.warn(`  Re-render page ${pageNum} at 300 DPI failed: ${err.message?.substring(0, 80)}`);
      warnings.push(`Page ${pageNum}: 300 DPI re-render failed`);
    }
  }

  if (reRenderedPages.length === 0) {
    try { cleanupJobDir(jobDir); } catch (e) { /* ignore */ }
    return { reExtractedItems: items, reExtractedPages: [], warnings };
  }

  // Crop BOM tables from the hi-res renders. When renderMode is "fullpage"
  // the model gets the entire page image (no crop) — use this when the
  // standard crop has clipped row 1.
  await processRenderedPages(jobDir, renderMode);

  // Build page images from cropped BOM files
  const pageImages: { pageNum: number; imagePath: string }[] = [];
  for (const rp of reRenderedPages) {
    // processRenderedPages creates _bom.png files from the rendered PNGs
    const base = path.basename(rp.imagePath, ".png");
    const bomPath = path.join(jobDir, base + "_bom.png");
    if (fs.existsSync(bomPath)) {
      pageImages.push({ pageNum: rp.pageNum, imagePath: bomPath });
    } else {
      // If no BOM crop, use the full render
      pageImages.push(rp);
    }
  }

  if (pageImages.length === 0) {
    try { cleanupJobDir(jobDir); } catch (e) { /* ignore */ }
    return { reExtractedItems: items, reExtractedPages: [], warnings };
  }

  // Re-extract with the higher resolution images
  const prompt = discipline === "mechanical" ? MECHANICAL_PROMPT : discipline === "structural" ? STRUCTURAL_PROMPT : CIVIL_PROMPT;
  const reResults = await extractWithVision(pageImages, prompt, discipline);

  // Replace items for re-extracted pages with new extraction results
  const reExtractedPageSet = new Set<number>();
  const updatedItems = items.filter(i => {
    // Keep items that are NOT from re-extracted pages
    if (pagesToReExtract.includes(i.sourcePage || 0)) return false;
    return true;
  });

  for (const rp of pageImages) {
    const pageData = reResults.results.get(rp.pageNum);
    if (!pageData || pageData.items.length === 0) {
      // Re-extraction got no items — keep originals
      const originals = items.filter(i => i.sourcePage === rp.pageNum);
      updatedItems.push(...originals);
      warnings.push(`Page ${rp.pageNum}: 300 DPI re-extraction returned 0 items \u2014 keeping original extraction`);
      continue;
    }

    reExtractedPageSet.add(rp.pageNum);
    const pageDrawingNumber = pageData.drawingNumber || null;
    // Build items from re-extracted data, fully normalized through the same
    // pipeline as the primary extraction (category detect, qty parse, size
    // validation, spec/material/schedule/rating extraction). This is what was
    // missing before — raw entries were being pushed straight to allItems.
    let pageRecoveredCount = 0;
    for (const entry of pageData.items) {
      if (!entry.description || entry.description.length < 3) continue;
      const rawQty = String(entry.qty || entry.quantity || "1").trim();
      const category = discipline === "mechanical" ? detectMechanicalCategory(entry.description) : (entry.category || "other");
      const isPipe = category === "pipe";
      let parsedQty: number;
      const qtyNotes: string[] = [];
      if (isPipe && /['''\u2019""\u201D]/.test(rawQty)) {
        parsedQty = parsePipeLength(rawQty).feet;
      } else if (isPipe) {
        parsedQty = parseFloat(rawQty) || 0;
      } else {
        parsedQty = Math.max(1, Math.round(parseFloat(rawQty) || 1));
      }
      if (isPipe) parsedQty = Math.round(Math.max(parsedQty, 0) * 100) / 100;
      const validatedSize = discipline === "mechanical" ? validateSize(String(entry.size || "N/A"), category) : String(entry.size || "N/A");
      const newItem: any = {
        id: randomUUID(),
        lineNumber: updatedItems.length + 1,
        discipline,
        category,
        description: entry.description.substring(0, 250),
        size: validatedSize,
        quantity: parsedQty,
        unit: isPipe ? "LF" : (/\bLF\b|\bFT\b/i.test(entry.description) ? "LF" : "EA"),
        spec: extractSpec(entry.description) || undefined,
        material: extractMaterial(entry.description) || undefined,
        schedule: extractSchedule(entry.description) || undefined,
        rating: extractRating(entry.description) || undefined,
        notes: `Sheet ${rp.pageNum}${pageDrawingNumber ? " | " + pageDrawingNumber : ""} (${entry.section || "SHOP"}) | Re-extracted at 300 DPI`,
        sourcePage: rp.pageNum,
        itemNo: typeof entry.itemNo === "number" ? entry.itemNo : (parseInt(String(entry.itemNo || "")) || undefined),
        drawingNumber: pageDrawingNumber,
        installLocation: (entry.section || "SHOP").toUpperCase() === "FIELD" ? "field" as const : "shop" as const,
        atContinuation: entry.atContinuation === true || entry.atContinuation === "true" || false,
        _reExtractedAt300DPI: true,
        _reExtractNote: `Re-extracted at 300 DPI (suspect page or low confidence)`,
      };
      updatedItems.push(newItem);
      pageRecoveredCount++;
    }
    console.log(`  Page ${rp.pageNum}: Re-extracted ${pageRecoveredCount} items at 300 DPI (was ${pageConfidence[rp.pageNum]?.count || 0} items)`);
  }

  try { cleanupJobDir(jobDir); } catch (e) { /* ignore */ }

  return {
    reExtractedItems: updatedItems,
    reExtractedPages: Array.from(reExtractedPageSet),
    warnings,
  };
}

// --- STRUCTURAL RENDER (Phase 1) ---
async function renderStructuralChunk(
  chunkPath: string,
  startPage: number,
  endPage: number
): Promise<RenderedStructuralChunk> {
  const chunkPageCount = endPage - startPage + 1;
  const renderResult = await renderFullPageImages(chunkPath, chunkPageCount, 300);
  return {
    jobDir: renderResult.jobDir,
    startPage,
    endPage,
    pageImages: renderResult.pageImages,
  };
}

// --- STRUCTURAL EXTRACT (Phase 2) ---
async function extractStructuralChunk(
  rendered: RenderedStructuralChunk
): Promise<{ items: any[]; metadata: any }> {
  const { startPage, pageImages } = rendered;
  const items: any[] = [];
  let metadata: any = {};

  if (pageImages.length === 0) return { items, metadata };

  const adjustedPageImages = pageImages.map(pi => ({ ...pi, globalPageNum: pi.pageNum + startPage - 1 }));
  const { results: visionResults } = await extractWithVision(pageImages, STRUCTURAL_PROMPT, "structural");

  for (const page of adjustedPageImages) {
    const pageData = visionResults.get(page.pageNum) || { items: [] };
    const entries = pageData.items;
    const pageDrawingNumber = pageData.drawingNumber || null;
    for (const entry of entries) {
      if (!entry.description || entry.description.length < 2) continue;
      const mappedCategory = mapStructuralCategory(entry.category || "other");
      const qty = parseFloat(String(entry.quantity || "1")) || 1;
      const isLfItem = ["LF", "SF", "CY"].includes(String(entry.unit || "EA").toUpperCase());
      items.push({
        category: mappedCategory,
        mark: entry.mark || undefined,
        description: entry.description.substring(0, 300),
        size: String(entry.size || "N/A"),
        quantity: isLfItem ? Math.round(qty * 100) / 100 : Math.max(1, Math.round(qty)),
        unit: String(entry.unit || "EA").toUpperCase(),
        grade: entry.grade || undefined,
        weight: parseFloat(String(entry.weight || "0")) > 0 ? Math.round(parseFloat(String(entry.weight))) : undefined,
        notes: `Sheet ${page.globalPageNum}${pageDrawingNumber ? " | " + pageDrawingNumber : ""}`,
        sourcePage: page.globalPageNum,
        drawingNumber: pageDrawingNumber,
      });
    }
  }

  if (startPage === 1 && pageImages.length > 0) {
    const firstText = pageImages[0].tesseractText || "";
    const revMatch = firstText.match(/(?:REV(?:ISION)?)[:\s]*([A-Z\d\.]+)/i);
    if (revMatch) metadata.revision = revMatch[1];
    const dateMatch = firstText.match(/(\d{1,2}[\/-]\d{1,2}[\/-]\d{2,4})/);
    if (dateMatch) metadata.drawingDate = dateMatch[1];
  }

  return { items, metadata };
}

// --- CIVIL RENDER (Phase 1) ---
async function renderCivilChunk(
  chunkPath: string,
  startPage: number,
  endPage: number
): Promise<RenderedCivilChunk> {
  const chunkPageCount = endPage - startPage + 1;
  const renderResult = await renderFullPageImages(chunkPath, chunkPageCount, 200);
  return {
    jobDir: renderResult.jobDir,
    startPage,
    endPage,
    pageImages: renderResult.pageImages,
  };
}

// --- CIVIL EXTRACT (Phase 2) ---
async function extractCivilChunk(
  rendered: RenderedCivilChunk
): Promise<{ items: any[]; metadata: any }> {
  const { startPage, pageImages } = rendered;
  const items: any[] = [];
  let metadata: any = {};

  if (pageImages.length === 0) return { items, metadata };

  const adjustedPageImages = pageImages.map(pi => ({ ...pi, globalPageNum: pi.pageNum + startPage - 1 }));
  const { results: visionResults } = await extractWithVision(pageImages, CIVIL_PROMPT, "civil");

  for (const page of adjustedPageImages) {
    const pageData = visionResults.get(page.pageNum) || { items: [] };
    const entries = pageData.items;
    const pageDrawingNumber = pageData.drawingNumber || null;
    for (const entry of entries) {
      if (!entry.description || entry.description.length < 2) continue;
      const mappedCategory = mapCivilCategory(entry.category || "other");
      const qty = parseFloat(String(entry.quantity || "1")) || 1;
      const isLfItem = ["LF", "SF", "CY", "SY", "TON", "AC"].includes(String(entry.unit || "EA").toUpperCase());
      items.push({
        category: mappedCategory,
        description: entry.description.substring(0, 300),
        size: String(entry.size || ""),
        quantity: isLfItem ? Math.round(qty * 100) / 100 : Math.max(1, Math.round(qty)),
        unit: String(entry.unit || "EA").toUpperCase(),
        material: entry.material || undefined,
        depth: entry.depth || undefined,
        notes: `Sheet ${page.globalPageNum}${pageDrawingNumber ? " | " + pageDrawingNumber : ""}`,
        sourcePage: page.globalPageNum,
        drawingNumber: pageDrawingNumber,
      });
    }
  }

  if (startPage === 1 && pageImages.length > 0) {
    const firstText = pageImages[0].tesseractText || "";
    const revMatch = firstText.match(/(?:REV(?:ISION)?)[:\s]*([A-Z\d\.]+)/i);
    if (revMatch) metadata.revision = revMatch[1];
    const dateMatch = firstText.match(/(\d{1,2}[\/-]\d{1,2}[\/-]\d{2,4})/);
    if (dateMatch) metadata.drawingDate = dateMatch[1];
  }

  return { items, metadata };
}

// ============================================================
// BACKGROUND PROCESSING
// ============================================================

async function processUploadedPdf(
  jobId: string,
  fileName: string,
  pdfPath: string,
  pageCount: number,
  totalChunks: number,
  discipline: string,
  verifyExtraction: boolean = true,
  hasRevisions: boolean = false,
  dualModel: boolean = false
): Promise<void> {
  console.log(`\n======= Processing [${discipline}]: ${fileName} (${pageCount} pages, jobId=${jobId}, revisions=${hasRevisions}, dualModel=${dualModel}) =======`);

  const allItems: any[] = [];
  let metadata: any = {};
  const warnings: string[] = [];
  let totalAuthFailures = 0;
  let chunksCompleted = 0;
  let chunksFailed = 0;

  const chunks = await splitPdfIntoChunks(pdfPath, pageCount);
  console.log(`Split into ${chunks.length} chunks`);
  const chunkCleanup: string[] = [];

  // PARTIAL RESULT SAVING: Create the project BEFORE chunk processing
  // so that partial results survive crashes
  const project = await storage.createTakeoffProject({
    name: fileName.replace(/\.pdf$/i, ""),
    fileName,
    discipline: discipline as any,
    items: [], // Start with empty items
  });
  console.log(`Created project ${project.id} (will save items incrementally)`);

  const isLargePackage = pageCount > LARGE_PACKAGE_THRESHOLD;
  if (isLargePackage) {
    console.log(`\n=== LARGE PACKAGE MODE: ${pageCount} pages, processing ${chunks.length} chunks sequentially (render→extract→cleanup) ===`);
  }

  // ============================================================
  // PHASE 1: PRE-RENDER ALL CHUNKS (no API calls needed)
  // For large packages, this phase is SKIPPED — chunks are processed end-to-end below
  // ============================================================
  const renderedChunks: { chunkIndex: number; rendered: RenderedMechanicalChunk | RenderedStructuralChunk | RenderedCivilChunk }[] = [];

  if (isLargePackage) {
    // === STREAMING MODE for large packages ===
    // Process each chunk end-to-end: render → extract → save → cleanup → next
    // This keeps memory usage bounded regardless of total page count
    console.log(`\n=== STREAMING MODE: Processing ${chunks.length} chunks end-to-end ===`);
    let pdfQuality: "vector" | "clean_scan" | "poor_scan" = "clean_scan";

    for (let i = 0; i < chunks.length; i++) {
      const chunk = chunks[i];
      const chunkNum = i + 1;
      console.log(`\n--- Chunk ${chunkNum}/${chunks.length}: pages ${chunk.startPage}-${chunk.endPage} (stream) ---`);

      storage.setJobProgress(jobId, {
        jobId, status: "processing", phase: `Section ${chunkNum}/${chunks.length}`,
        chunk: chunkNum, totalChunks: chunks.length,
        pagesProcessed: chunk.startPage - 1, totalPages: pageCount,
        itemsFound: allItems.length,
        warnings: warnings.length > 0 ? warnings : undefined,
      });

      // Step 1: Render this chunk
      let rendered: RenderedMechanicalChunk | RenderedStructuralChunk | RenderedCivilChunk;
      try {
        if (discipline === "mechanical") {
          rendered = await renderMechanicalChunk(chunk.chunkPath, chunk.startPage, chunk.endPage, hasRevisions);
        } else if (discipline === "structural") {
          rendered = await renderStructuralChunk(chunk.chunkPath, chunk.startPage, chunk.endPage);
        } else {
          rendered = await renderCivilChunk(chunk.chunkPath, chunk.startPage, chunk.endPage);
        }
      } catch (renderErr: any) {
        console.error(`  Render chunk ${chunkNum} failed:`, renderErr.message);
        warnings.push(`Section ${chunkNum} render failed: ${renderErr.message?.substring(0, 100)}`);
        chunksFailed++;
        continue; // Skip to next chunk
      }

      // Classify PDF quality on first chunk
      if (i === 0) {
        const firstPageImages: any[] = (rendered as any).bomPageImages || (rendered as any).cloudPageImages || (rendered as any).pageImages || [];
        if (firstPageImages.length > 0) pdfQuality = classifyPdfQuality(firstPageImages);
      }

      // Step 2: Extract from this chunk
      try {
        let chunkResult: { items: any[]; metadata: any; authFailures?: number };
        if (discipline === "mechanical") {
          chunkResult = await extractMechanicalChunk(
            rendered as RenderedMechanicalChunk,
            (pagesProcessed) => {
              storage.setJobProgress(jobId, {
                jobId, status: "processing", phase: `Section ${chunkNum}/${chunks.length}`,
                chunk: chunkNum, totalChunks: chunks.length,
                pagesProcessed, totalPages: pageCount,
                itemsFound: allItems.length,
                warnings: warnings.length > 0 ? warnings : undefined,
              });
            },
            verifyExtraction
          );
        } else if (discipline === "structural") {
          chunkResult = await extractStructuralChunk(rendered as RenderedStructuralChunk);
        } else {
          chunkResult = await extractCivilChunk(rendered as RenderedCivilChunk);
        }

        if (chunkResult.authFailures) totalAuthFailures += chunkResult.authFailures;
        allItems.push(...chunkResult.items);
        if (i === 0 && chunkResult.metadata) metadata = chunkResult.metadata;
        chunksCompleted++;
        console.log(`  Chunk ${chunkNum} complete: ${chunkResult.items.length} items (total: ${allItems.length})`);

        // Save partial results after each chunk — assign lineNumbers first (DB requires NOT NULL)
        const itemsToSave = allItems.map((item: any, idx: number) => ({ ...item, lineNumber: idx + 1 }));
        await storage.updateTakeoffProjectItems(project.id, itemsToSave);
        console.log(`  Saved ${allItems.length} items to project ${project.id}`);
      } catch (extractErr: any) {
        console.error(`  Extract chunk ${chunkNum} failed:`, extractErr.message);
        warnings.push(`Section ${chunkNum} extraction failed: ${extractErr.message?.substring(0, 100)}`);
        chunksFailed++;
      }

      // Step 3: Cleanup rendered images immediately to free memory/disk
      try { cleanupJobDir(rendered.jobDir); } catch (e) { console.warn("Suppressed error:", e); }
      // Clean up chunk PDF if it's not the original
      if (chunk.chunkPath !== pdfPath) {
        try { fs.unlinkSync(chunk.chunkPath); } catch (e) { console.warn("Suppressed error:", e); }
      }
    }

    // Skip to post-processing (bypass the two-phase logic below)
    console.log(`\n=== STREAMING COMPLETE: ${chunksCompleted}/${chunks.length} chunks, ${allItems.length} items ===`);

    // NOTE: Original PDF cleanup deferred until AFTER pipe-qty retry, which
    // needs the PDF to render specific pages at high DPI for best-guess.

    // Run post-processing pipeline on all items
    if (allItems.length > 0) {
      // NOTE: Same-drawing-number auto-dedup is intentionally disabled here.
      // It can be triggered manually via the "Re-run Dedup" button (POST
      // /api/takeoff-projects/:id/redup) when the user reviews the takeoff
      // and wants to apply duplicate-drawing detection. We opted for manual
      // control because automatic dedup occasionally removed legitimate items.
      // Continuation page dedup (only fires on AI-marked atContinuation items)
      const dedupedItems = dedupContinuationPages([...allItems]);
      allItems.length = 0;
      allItems.push(...dedupedItems);

      // Suspect-page re-extraction at 300 DPI / fullpage mode (PRIMARY recovery).
      // Catches pages where row 1 (PIPE) was skipped during the initial extraction
      // because the BOM crop clipped the top of the table, the pipe row's qty
      // format confused the model, or the model otherwise dropped a row. Replaces
      // the earlier pipeRowRecovery() band-aid as the primary recovery path —
      // pipeRowRecovery is kept below as a final safety net.
      if (discipline === "mechanical" && fs.existsSync(pdfPath)) {
        try {
          const suspectPages = detectSuspectPages(allItems);
          if (suspectPages.length > 0) {
            console.log(`  Detected ${suspectPages.length} suspect page(s) (row 1 likely missing): ${suspectPages.slice(0, 10).join(", ")}${suspectPages.length > 10 ? ", ..." : ""}`);
          }
          const reExtractResult = await reExtractLowConfidencePages(allItems, pdfPath, discipline, 1, {
            extraPages: suspectPages,
            renderMode: "fullpage",
            pageLimit: 25,
          });
          if (reExtractResult.reExtractedPages.length > 0) {
            allItems.length = 0;
            allItems.push(...reExtractResult.reExtractedItems);
            warnings.push(`Re-extracted ${reExtractResult.reExtractedPages.length} suspect/low-confidence page(s) at 300 DPI (uncropped).`);
          }
          if (reExtractResult.warnings.length > 0) warnings.push(...reExtractResult.warnings);
        } catch (reExtractErr: any) {
          console.warn(`  Re-extraction failed: ${reExtractErr.message?.substring(0, 100)}`);
        }
      }

      // Missed pipe ROW recovery — SAFETY NET only. The primary fix for skipped
      // pipe rows is the suspect-page re-extraction above. This pass remains as
      // a backstop for cases the suspect detector missed.
      if (discipline === "mechanical" && fs.existsSync(pdfPath)) {
        try {
          const recoveryResult = await pipeRowRecovery(allItems, pdfPath, 1);
          if (recoveryResult.recoveredCount > 0) {
            warnings.push(`Pipe row safety-net: ${recoveryResult.recoveredCount} additional pipe row(s) recovered after re-extraction (low confidence \u2014 verify against drawings).`);
          }
        } catch (recoveryErr: any) {
          console.warn(`  Pipe-row safety-net failed: ${recoveryErr.message?.substring(0, 100)}`);
        }
      }

      // Pipe quantity best-guess retry (must run BEFORE we clean up the PDF)
      if (discipline === "mechanical" && fs.existsSync(pdfPath)) {
        try {
          const retryResult = await pipeQtyBestGuessRetry(allItems, pdfPath, 1);
          if (retryResult.guessedCount > 0) {
            warnings.push(`Pipe quantity best-guess: ${retryResult.guessedCount} pipe length(s) estimated from drawings (flagged low confidence \u2014 verify before bidding).`);
          }
        } catch (retryErr: any) {
          console.warn(`  Pipe-qty retry failed: ${retryErr.message?.substring(0, 100)}`);
        }
      }

      // Flag potential same-page duplicates (does NOT remove anything)
      if (discipline === "mechanical") {
        try {
          const dupResult = flagPotentialDuplicates(allItems);
          if (dupResult.flaggedCount > 0) {
            warnings.push(`Possible duplicates: ${dupResult.flaggedCount} item(s) across ${dupResult.groups} group(s) flagged for review (same description repeated on same page).`);
          }
        } catch (dupErr: any) {
          console.warn(`  Duplicate flagging failed: ${dupErr.message?.substring(0, 100)}`);
        }
      }

      // Now safe to clean up original PDF
      try { if (fs.existsSync(pdfPath)) fs.unlinkSync(pdfPath); } catch (e) { console.warn("Suppressed error:", e); }

      // Apply historical pattern corrections (learned from user edits)
      applyHistoricalPatterns(allItems);

      // Run piping rules engine
      const pipingRulesResult = validatePipingRules(allItems);
      if (pipingRulesResult.warnings.length > 0) {
        console.log(`  Piping rules engine: ${pipingRulesResult.warnings.length} warnings`);
        warnings.push(...pipingRulesResult.warnings);
      }

      // Build continuation graph for metadata
      const continuationGraph = buildContinuationGraph(allItems);

      // Confidence scoring
      for (const item of allItems) {
        const { confidence, confidenceScore, confidenceNotes } = computeConfidenceScore(item, pdfQuality);
        // Preserve any 'low' confidence already set by post-processing (pipe-qty
        // retry, duplicate flagging, etc.) — these are explicit signals from
        // logic we trust more than the heuristic computeConfidenceScore.
        const preservedLow = item.confidence === "low";
        if (item._verifiedCorrect) {
          item.confidence = "high";
          item.confidenceScore = Math.max(confidenceScore, 90);
        } else if (preservedLow) {
          item.confidence = "low";
          item.confidenceScore = Math.min(confidenceScore, item.confidenceScore || 50);
        } else {
          item.confidence = confidence;
          item.confidenceScore = confidenceScore;
        }
        // Merge validation notes (pipe-qty retry, duplicate flag, etc.) into
        // confidenceNotes so they survive the cleanup below.
        const validationNotes: string[] = (item as any)._validationNotes || [];
        if (validationNotes.length > 0) {
          item.confidenceNotes = [confidenceNotes, ...validationNotes].filter(Boolean).join(" | ");
        } else {
          item.confidenceNotes = confidenceNotes;
        }
        delete item._validationFlag;
        delete item._validationNotes;
        delete item._sizeWarning;
        delete item._crossCheckWarning;
      }

      // Assign line numbers and save
      const validItems = allItems.filter((i: any) => i.quantity > 0 || i.category === "pipe");
      validItems.forEach((item: any, idx: number) => { item.lineNumber = idx + 1; });

      await storage.updateTakeoffProjectItems(project.id, validItems);
      if (metadata.lineNumber || metadata.area || metadata.revision || metadata.drawingDate) {
        storage.updateTakeoffProjectMetadata(project.id, metadata);
      }

      // Branch ISO detection
      const branchDetection = detectBranchISOPattern(validItems);
      const calSummary = buildCalibrationSummary(validItems);

      storage.setJobProgress(jobId, {
        jobId, status: "done", phase: "done",
        chunk: chunks.length, totalChunks: chunks.length,
        pagesProcessed: pageCount, totalPages: pageCount,
        itemsFound: validItems.length,
        projectId: project.id,
        warnings: [
          ...warnings,
          ...(branchDetection.detected ? [branchDetection.note] : []),
        ].filter(Boolean),
        continuationGraph: continuationGraph.graph,
        sharedFittings: continuationGraph.sharedFittings,
      });
      console.log(`\n======= DONE: ${validItems.length} items saved to project ${project.id} =======`);
    } else {
      storage.setJobProgress(jobId, {
        jobId, status: "done", phase: "done",
        chunk: chunks.length, totalChunks: chunks.length,
        pagesProcessed: pageCount, totalPages: pageCount,
        itemsFound: 0, projectId: project.id,
        warnings: [...warnings, "No items extracted from any section."],
      });
    }
    return; // Exit early — streaming mode handles everything
  }

  // ============================================================
  // STANDARD MODE (small packages): Two-phase render-then-extract
  // ============================================================
  console.log(`\n=== PHASE 1: Rendering all ${chunks.length} chunks ===`);

  try {
    for (let i = 0; i < chunks.length; i++) {
      const chunk = chunks[i];
      const chunkNum = i + 1;

      console.log(`\n--- Render chunk ${chunkNum}/${chunks.length}: pages ${chunk.startPage}-${chunk.endPage} ---`);
      storage.setJobProgress(jobId, {
        jobId,
        status: "processing",
        phase: "rendering",
        chunk: chunkNum,
        totalChunks: chunks.length,
        pagesProcessed: chunk.startPage - 1,
        totalPages: pageCount,
        itemsFound: 0,
        warnings: warnings.length > 0 ? warnings : undefined,
      });

      try {
        let rendered: RenderedMechanicalChunk | RenderedStructuralChunk | RenderedCivilChunk;
        if (discipline === "mechanical") {
          rendered = await renderMechanicalChunk(chunk.chunkPath, chunk.startPage, chunk.endPage, hasRevisions);
        } else if (discipline === "structural") {
          rendered = await renderStructuralChunk(chunk.chunkPath, chunk.startPage, chunk.endPage);
        } else {
          rendered = await renderCivilChunk(chunk.chunkPath, chunk.startPage, chunk.endPage);
        }
        renderedChunks.push({ chunkIndex: i, rendered });
        console.log(`  Render chunk ${chunkNum} done`);
      } catch (renderErr: any) {
        const errMsg = renderErr.message || String(renderErr);
        const errStderr = renderErr.stderr?.toString().substring(0, 500) || "";
        console.error(`  Render chunk ${chunkNum} failed:`, errMsg);
        if (errStderr) console.error(`  stderr: ${errStderr}`);
        const renderWarning = `Render chunk ${chunkNum} (pages ${chunk.startPage}-${chunk.endPage}) failed: ${errMsg}${errStderr ? " | stderr: " + errStderr.substring(0, 200) : ""}`;
        if (!warnings.includes(renderWarning)) warnings.push(renderWarning);
      }

      storage.setJobProgress(jobId, {
        jobId,
        status: "processing",
        phase: "rendering",
        chunk: chunkNum,
        totalChunks: chunks.length,
        pagesProcessed: chunk.endPage,
        totalPages: pageCount,
        itemsFound: 0,
        warnings: warnings.length > 0 ? warnings : undefined,
      });
    }
  } catch (renderPhaseErr: any) {
    console.error("Render phase failed:", renderPhaseErr.message || renderPhaseErr);
    // Clean up any rendered images on complete render failure
    for (const rc of renderedChunks) {
      try { cleanupJobDir(rc.rendered.jobDir); } catch (e) { console.warn("Suppressed error:", e); }
    }
    storage.setJobProgress(jobId, {
      jobId,
      status: "error",
      phase: "error",
      chunk: 0,
      totalChunks: 0,
      pagesProcessed: 0,
      totalPages: pageCount,
      itemsFound: 0,
      error: `Rendering failed: ${renderPhaseErr.message || "Unknown error"}.`,
      warnings: warnings.length > 0 ? warnings : undefined,
    });
    return;
  }

  if (renderedChunks.length === 0) {
    storage.setJobProgress(jobId, {
      jobId,
      status: "error",
      phase: "error",
      chunk: 0,
      totalChunks: 0,
      pagesProcessed: pageCount,
      totalPages: pageCount,
      itemsFound: 0,
      error: `All ${chunks.length} render chunks failed. No pages could be prepared for extraction.${warnings.length > 0 ? " Details: " + warnings[0] : ""}`,
      warnings: warnings.length > 0 ? warnings : undefined,
    });
    try { if (fs.existsSync(pdfPath)) fs.unlinkSync(pdfPath); } catch (e) { console.warn("Suppressed error:", e); }
    return;
  }

  console.log(`\n=== PHASE 1 COMPLETE: ${renderedChunks.length}/${chunks.length} chunks rendered ===`);

  // ============================================================
  // PDF QUALITY GATE: Analyze first rendered chunk for quality
  // ============================================================
  let pdfQuality: "vector" | "clean_scan" | "poor_scan" = "clean_scan";
  if (renderedChunks.length > 0) {
    const firstRendered = renderedChunks[0].rendered;
    const firstPageImages: { pageNum: number; imagePath: string; tesseractText: string }[] = 
      (firstRendered as any).bomPageImages || (firstRendered as any).cloudPageImages || (firstRendered as any).pageImages || [];
    if (firstPageImages.length > 0) {
      pdfQuality = classifyPdfQuality(firstPageImages);
      console.log(`  PDF Quality Classification: ${pdfQuality}`);
      if (pdfQuality === "poor_scan") {
        const qualityWarning = "Low quality scan detected \u2014 manual review recommended for all items";
        if (!warnings.includes(qualityWarning)) warnings.push(qualityWarning);
      }
    }
  }

  // ============================================================
  // DRAWING TEMPLATE DETECTION (Feature 2)
  // ============================================================
  let detectedTemplate: any = null;
  if (renderedChunks.length > 0) {
    const firstRendered = renderedChunks[0].rendered;
    const firstPageImages: { pageNum: number; imagePath: string; tesseractText: string }[] =
      (firstRendered as any).bomPageImages || (firstRendered as any).cloudPageImages || (firstRendered as any).pageImages || [];
    if (firstPageImages.length > 0) {
      const combinedOcr = firstPageImages.slice(0, 3).map(p => p.tesseractText || "").join("\n");
      detectedTemplate = detectDrawingTemplate(combinedOcr);
      if (detectedTemplate) {
        console.log(`  Matched drawing template: "${detectedTemplate.name}" (id=${detectedTemplate.id})`);
        storage.incrementTemplateUsage(detectedTemplate.id);
      }
    }
  }

  // ============================================================
  // PHASE 2: AI EXTRACTION from pre-rendered images (API calls)
  // ============================================================
  console.log(`\n=== PHASE 2: Extracting from ${renderedChunks.length} rendered chunks ===`);

  try {
    for (let ri = 0; ri < renderedChunks.length; ri++) {
      const { chunkIndex, rendered } = renderedChunks[ri];
      const chunk = chunks[chunkIndex];
      const chunkNum = chunkIndex + 1;

      console.log(`\n--- Extract chunk ${chunkNum}/${chunks.length}: pages ${chunk.startPage}-${chunk.endPage} ---`);
      storage.setJobProgress(jobId, {
        jobId,
        status: "processing",
        phase: "extracting",
        chunk: chunkNum,
        totalChunks: chunks.length,
        pagesProcessed: chunk.startPage - 1,
        totalPages: pageCount,
        itemsFound: allItems.length,
        warnings: warnings.length > 0 ? warnings : undefined,
      });

      try {
        let chunkResult: { items: any[]; metadata: any; authFailures?: number };
        if (discipline === "mechanical") {
          chunkResult = await extractMechanicalChunk(
            rendered as RenderedMechanicalChunk,
            (pagesProcessed) => {
              storage.setJobProgress(jobId, {
                jobId,
                status: "processing",
                phase: "extracting",
                chunk: chunkNum,
                totalChunks: chunks.length,
                pagesProcessed,
                totalPages: pageCount,
                itemsFound: allItems.length,
                warnings: warnings.length > 0 ? warnings : undefined,
              });
            },
            verifyExtraction
          );
        } else if (discipline === "structural") {
          chunkResult = await extractStructuralChunk(rendered as RenderedStructuralChunk);
        } else {
          chunkResult = await extractCivilChunk(rendered as RenderedCivilChunk);
        }

        // Track auth failures for warning generation
        if (chunkResult.authFailures) {
          totalAuthFailures += chunkResult.authFailures;
        }

        allItems.push(...chunkResult.items);
        if (chunkIndex === 0 && chunkResult.metadata) metadata = chunkResult.metadata;

        console.log(`  Extract chunk ${chunkNum} done: ${chunkResult.items.length} items (total: ${allItems.length})`);
        chunksCompleted++;

        // Check if >50% of pages in this chunk had auth failures
        const chunkPages = chunk.endPage - chunk.startPage + 1;
        if (chunkResult.authFailures && chunkResult.authFailures > chunkPages * 0.5) {
          const authWarning = `API authentication issues detected in chunk ${chunkNum}. Some pages may have incomplete results.`;
          if (!warnings.includes(authWarning)) warnings.push(authWarning);
        }

        // DUAL-MODEL: If enabled and Gemini key available, run Gemini on first few pages and compare
        if (dualModel && getUserGeminiKey() && chunkIndex === 0) {
          try {
            const pageImgs: { pageNum: number; imagePath: string; tesseractText: string }[] =
              (rendered as any).bomPageImages || (rendered as any).cloudPageImages || (rendered as any).pageImages || [];
            const samplePages = pageImgs.slice(0, 2); // Compare first 2 pages
            for (const pg of samplePages) {
              if (!pg.imagePath || !fs.existsSync(pg.imagePath)) continue;
              const imgData = fs.readFileSync(pg.imagePath).toString("base64");
              const geminiItems = await extractWithGemini(imgData, `Extract all items from this ${discipline} drawing. Return a JSON array with objects containing: description, size, quantity, unit, category. Only return the JSON array, no other text.`);
              if (geminiItems.length > 0) {
                console.log(`  Gemini found ${geminiItems.length} items on page ${pg.pageNum} (Claude found items in chunk)`);
                // Mark items with dual-model confidence boost if both models agree
                const geminiDescs = new Set(geminiItems.map((g: any) => (g.description || "").toLowerCase().trim().slice(0, 40)));
                for (const item of chunkResult.items) {
                  const itemDesc = (item.description || "").toLowerCase().trim().slice(0, 40);
                  if (geminiDescs.has(itemDesc)) {
                    item._dualModelVerified = true;
                  }
                }
              }
            }
          } catch (geminiErr: any) {
            console.warn(`  Dual-model Gemini comparison failed: ${geminiErr.message}`);
          }
        }

        // PARTIAL RESULT SAVING: Save items to DB after each successful chunk
        const itemsForDb = allItems.map((item: any, idx: number) => ({
          ...item,
          id: item.id || randomUUID(),
          discipline,
          lineNumber: idx + 1,
        }));
        await storage.updateTakeoffProjectItems(project.id, itemsForDb);
        console.log(`  Saved ${itemsForDb.length} items to project ${project.id} after chunk ${chunkNum}`);

      } catch (chunkErr: any) {
        chunksFailed++;
        console.error(`  Extract chunk ${chunkNum} failed:`, chunkErr.message || chunkErr);
        const chunkWarning = `Chunk ${chunkNum} (pages ${chunk.startPage}-${chunk.endPage}) extraction failed: ${chunkErr.message || "Unknown error"}`;
        if (!warnings.includes(chunkWarning)) warnings.push(chunkWarning);
      } finally {
        // Save page thumbnails BEFORE cleanup
        try { savePageThumbnails(project.id, rendered, rendered.startPage); } catch (e) { console.warn("Suppressed error:", e); }
        // Clean up rendered images for this chunk after extraction
        try { cleanupJobDir(rendered.jobDir); } catch (e) { console.warn("Suppressed error:", e); }
      }

      storage.setJobProgress(jobId, {
        jobId,
        status: "processing",
        phase: "extracting",
        chunk: chunkNum,
        totalChunks: chunks.length,
        pagesProcessed: chunk.endPage,
        totalPages: pageCount,
        itemsFound: allItems.length,
        warnings: warnings.length > 0 ? warnings : undefined,
      });

      if (chunk.chunkPath !== pdfPath) chunkCleanup.push(chunk.chunkPath);
    }
  } catch (visionErr: any) {
    console.error("AI Vision pipeline failed:", visionErr.message || visionErr);
    // Clean up remaining rendered images
    for (const rc of renderedChunks) {
      try { cleanupJobDir(rc.rendered.jobDir); } catch (e) { console.warn("Suppressed error:", e); }
    }
    // Project already has partial results saved incrementally
    storage.setJobProgress(jobId, {
      jobId,
      status: allItems.length > 0 ? "done" : "error",
      phase: allItems.length > 0 ? "done" : "error",
      chunk: 0,
      totalChunks: 0,
      pagesProcessed: pageCount,
      totalPages: pageCount,
      itemsFound: allItems.length,
      projectId: allItems.length > 0 ? project.id : undefined,
      error: allItems.length > 0
        ? undefined
        : `AI Vision processing failed: ${visionErr.message || "Unknown error"}.`,
      warnings: allItems.length > 0
        ? [...warnings, `Processing interrupted after ${chunksCompleted} of ${chunks.length} chunks. Partial results saved.`]
        : (warnings.length > 0 ? warnings : undefined),
    });
    return;
  } finally {
    for (const cp of chunkCleanup) { try { fs.unlinkSync(cp); } catch (e) { console.warn("Suppressed error:", e); } }
    if (chunks.length > 1 && chunks[0].chunkPath !== pdfPath) {
      const chunkDir = path.dirname(chunks[0].chunkPath);
      try { fs.rmSync(chunkDir, { recursive: true, force: true }); } catch (e) { console.warn("Suppressed error:", e); }
    }
  }

  // Add overall auth warning if significant failures occurred
  if (totalAuthFailures > 0 && totalAuthFailures > pageCount * 0.5) {
    const overallWarning = "API authentication issues detected. Some pages may have incomplete results.";
    if (!warnings.includes(overallWarning)) warnings.push(overallWarning);
  }

  console.log(`Parsed ${allItems.length} items (auth failures: ${totalAuthFailures})`);

  if (allItems.length === 0) {
    storage.setJobProgress(jobId, {
      jobId,
      status: "error",
      phase: "error",
      chunk: 0,
      totalChunks: 0,
      pagesProcessed: pageCount,
      totalPages: pageCount,
      itemsFound: 0,
      error: `Could not identify ${discipline} components. Processed ${pageCount} pages.`,
      warnings: warnings.length > 0 ? warnings : undefined,
    });
    // Clean up PDF on error too
    try { if (fs.existsSync(pdfPath)) fs.unlinkSync(pdfPath); } catch (e) { console.warn("Suppressed error:", e); }
    return;
  }

  // Pipeline order: extraction (done above) → auto-correction (done in chunk) → verification (done in chunk) → validation → dedup → confidence scoring → final save

  // Step: Post-processing validation
  if (discipline === "mechanical") {
    console.log(`  Running post-processing validation on ${allItems.length} items...`);
    validateExtractedItems(allItems);

    // NOTE: Same-drawing-number auto-dedup is intentionally disabled here.
    // Available manually via the "Re-run Dedup" button on the takeoff page.
    // We opted for manual control because automatic dedup occasionally removed
    // legitimate items.

    // Step: Continuation page dedup (only fires on AI-marked atContinuation items)
    console.log(`  Running continuation page dedup...`);
    const dedupedItems = dedupContinuationPages([...allItems]);
    allItems.length = 0;
    allItems.push(...dedupedItems);

    // Step: Piping validation rules (Council Item 4)
    console.log(`  Running piping validation rules...`);
    validatePipingBom(allItems);
  }

  // Step: Apply historical pattern corrections
  console.log(`  Applying historical pattern corrections...`);
  applyHistoricalPatterns(allItems);

  // Step: Piping rules engine
  if (discipline === "mechanical") {
    console.log(`  Running piping rules engine...`);
    const pipingRulesResult = validatePipingRules(allItems);
    if (pipingRulesResult.warnings.length > 0) {
      console.log(`  Piping rules engine: ${pipingRulesResult.warnings.length} warnings`);
      warnings.push(...pipingRulesResult.warnings);
    }
  }

  // Step: Build continuation graph
  const continuationGraph = buildContinuationGraph(allItems);

  // Step: Enhanced confidence scoring (replaces old simplistic scoring)
  allItems.forEach((item: any) => {
    const { confidence, confidenceScore, confidenceNotes } = computeConfidenceScore(item, pdfQuality);
    // Dual-model verification boost: if both Claude and Gemini agree, bump confidence
    if (item._dualModelVerified && confidence !== "high") {
      item.confidence = "high";
      item.confidenceScore = Math.max(confidenceScore, 90);
      item.confidenceNotes = "Verified by dual-model (Claude + Gemini)";
    } else {
      item.confidence = confidence;
      item.confidenceScore = confidenceScore;
      item.confidenceNotes = confidenceNotes || undefined;
    }
    // Clean up internal flags
    delete item._validationFlag;
    delete item._validationNotes;
    delete item._sizeWarning;
    delete item._verified;
    delete item._verifyNote;
    delete item._dualModelVerified;
    delete item._crossCheckWarning;
    delete item._reExtractedAt300DPI;
    delete item._reExtractNote;
  });

  // Step: Suspect-page re-extraction at 300 DPI / fullpage mode (PRIMARY recovery).
  // Catches pages where row 1 (PIPE) was skipped during the initial extraction
  // because the BOM crop clipped the top of the table, the pipe row's qty
  // format confused the model, or the model otherwise dropped a row.
  // ALSO catches the original "low-confidence pages" use case via the same
  // function. Replaces the earlier pipeRowRecovery() band-aid as the primary
  // recovery path; pipeRowRecovery is kept below as a final safety net.
  if (discipline === "mechanical" && fs.existsSync(pdfPath)) {
    try {
      const startPage = chunks[0]?.startPage || 1;
      const suspectPages = detectSuspectPages(allItems);
      if (suspectPages.length > 0) {
        console.log(`  Detected ${suspectPages.length} suspect page(s) (row 1 likely missing): ${suspectPages.slice(0, 10).join(", ")}${suspectPages.length > 10 ? ", ..." : ""}`);
      }
      const reExtractResult = await reExtractLowConfidencePages(allItems, pdfPath, discipline, startPage, {
        extraPages: suspectPages,
        renderMode: "fullpage",
        pageLimit: 25,
      });
      if (reExtractResult.reExtractedPages.length > 0) {
        // BUG FIX: actually use the result. Old code computed reExtractedItems
        // and discarded it; this step was effectively a logging pass since launch.
        allItems.length = 0;
        allItems.push(...reExtractResult.reExtractedItems);
        warnings.push(`Re-extracted ${reExtractResult.reExtractedPages.length} suspect/low-confidence page(s) at 300 DPI (uncropped).`);
      }
      if (reExtractResult.warnings.length > 0) warnings.push(...reExtractResult.warnings);
    } catch (reExtractErr: any) {
      console.warn(`  Adaptive DPI re-extraction failed: ${reExtractErr.message?.substring(0, 100)}`);
    }
  }

  // Step: Missed pipe ROW recovery -- SAFETY NET only. The primary fix for
  // skipped pipe rows is the suspect-page re-extraction above. This pass
  // remains as a backstop for cases the suspect detector missed.
  if (discipline === "mechanical" && fs.existsSync(pdfPath)) {
    try {
      const startPage = chunks[0]?.startPage || 1;
      const recoveryResult = await pipeRowRecovery(allItems, pdfPath, startPage);
      if (recoveryResult.recoveredCount > 0) {
        warnings.push(`Pipe row safety-net: ${recoveryResult.recoveredCount} additional pipe row(s) recovered after re-extraction (low confidence \u2014 verify against drawings).`);
      }
    } catch (recoveryErr: any) {
      console.warn(`  Pipe-row safety-net failed: ${recoveryErr.message?.substring(0, 100)}`);
    }
  }

  // Step: Pipe quantity best-guess retry for unread pipes
  // Sends the FULL ISO drawing back to Claude for any pipe row with qty=0 and
  // asks for a length estimate from the drawing line work + dimension callouts.
  // Result is filled in with confidence=low + a clear flag note for review.
  if (discipline === "mechanical" && fs.existsSync(pdfPath)) {
    try {
      const startPage = chunks[0]?.startPage || 1;
      const retryResult = await pipeQtyBestGuessRetry(allItems, pdfPath, startPage);
      if (retryResult.guessedCount > 0) {
        warnings.push(`Pipe quantity best-guess: ${retryResult.guessedCount} pipe length(s) estimated from drawings (flagged low confidence \u2014 verify before bidding).`);
      }
    } catch (retryErr: any) {
      console.warn(`  Pipe-qty retry failed: ${retryErr.message?.substring(0, 100)}`);
    }
  }

  // Step: Flag potential same-page duplicate items (does NOT remove anything)
  // User asked: 'I don't want anything to be taken off but i want it flagged
  // for review as well and marked low confidence'.
  if (discipline === "mechanical") {
    try {
      const dupResult = flagPotentialDuplicates(allItems);
      if (dupResult.flaggedCount > 0) {
        warnings.push(`Possible duplicates: ${dupResult.flaggedCount} item(s) across ${dupResult.groups} group(s) flagged for review (same description repeated on same page).`);
      }
    } catch (dupErr: any) {
      console.warn(`  Duplicate flagging failed: ${dupErr.message?.substring(0, 100)}`);
    }
  }

  // Validate items before saving — reject any with missing critical fields
  const validItems = allItems.filter((item: any) => {
    if (!item.description || typeof item.description !== "string" || item.description.trim().length < 2) return false;
    if (!item.category || typeof item.category !== "string") return false;
    if (typeof item.quantity !== "number" || item.quantity < 0) return false;
    return true;
  });
  const rejectedCount = allItems.length - validItems.length;
  if (rejectedCount > 0) {
    console.warn(`  Rejected ${rejectedCount} items with invalid/missing required fields`);
    warnings.push(`${rejectedCount} extracted items rejected due to invalid data (missing description, category, or quantity)`);
  }

  validItems.forEach((item: any, idx: number) => { item.lineNumber = idx + 1; });

  const itemsWithIds = validItems.map((item: any) => ({
    ...item,
    id: item.id || randomUUID(),
    discipline,
  }));

  // Update metadata on the project
  // Note: metadata was collected from the first chunk
  // We do a final update with fully post-processed items
  await storage.updateTakeoffProjectItems(project.id, itemsWithIds);

  // Persist extracted metadata (lineNumber, area, revision, drawingDate) to the project record
  if (metadata.lineNumber || metadata.area || metadata.revision || metadata.drawingDate) {
    storage.updateTakeoffProjectMetadata(project.id, metadata);
  }

  // Auto-create drawing template if no template was detected and extraction was successful
  if (!detectedTemplate && itemsWithIds.length >= 5) {
    try {
      const firstRendered = renderedChunks[0]?.rendered;
      const firstPageImages: { pageNum: number; imagePath: string; tesseractText: string }[] =
        (firstRendered as any)?.bomPageImages || (firstRendered as any)?.cloudPageImages || (firstRendered as any)?.pageImages || [];
      const sampleText = firstPageImages.slice(0, 2).map(p => p.tesseractText || "").join("\n").slice(0, 2000);
      if (sampleText.length > 100) {
        // Extract engineering firm from OCR text
        const firmMatch = sampleText.match(/(?:DESIGNED BY|ENGINEER(?:ED)?(?:\s+BY)?|PREPARED BY|DRAWN BY)[:\s]+([A-Z][A-Za-z\s&,.]+?)(?:\n|$)/i);
        const firmName = firmMatch ? firmMatch[1].trim().slice(0, 100) : undefined;
        storage.createDrawingTemplate({
          name: `Auto: ${fileName.replace(/\.pdf$/i, "").slice(0, 60)}`,
          engineeringFirm: firmName,
          sampleOcrText: sampleText.slice(0, 1500),
          extractionNotes: `Auto-created from successful extraction of ${itemsWithIds.length} items`,
        });
        console.log(`  Auto-created drawing template for "${fileName}"`);
      }
    } catch (e) { console.warn("Suppressed error:", e); }
  }

  const finalWarnings = chunksFailed > 0
    ? [...warnings, `${chunksFailed} of ${chunks.length} chunks had issues. ${chunksCompleted} chunks processed successfully.`]
    : warnings;

  // Calibration: build cross-validation summary
  const calibrationSummary = buildCalibrationSummary(itemsWithIds);
  if (calibrationSummary.branchISODetected) {
    warnings.push(calibrationSummary.rackPipeNote || "Branch ISOs detected - header/rack pipe may not be included");
  }

  console.log(`Done: ${itemsWithIds.length} items from ${pageCount} pages`);

  storage.setJobProgress(jobId, {
    jobId,
    status: "done",
    phase: "done",
    chunk: totalChunks,
    totalChunks,
    pagesProcessed: pageCount,
    totalPages: pageCount,
    itemsFound: itemsWithIds.length,
    projectId: project.id,
    warnings: finalWarnings.length > 0 ? finalWarnings : undefined,
    pdfQuality,
    continuationGraph: continuationGraph.graph,
    sharedFittings: continuationGraph.sharedFittings,
  });

  // Clean up uploaded PDF from /tmp after processing
  try {
    if (fs.existsSync(pdfPath)) fs.unlinkSync(pdfPath);
  } catch (cleanupErr) {
    console.warn(`Failed to clean up uploaded PDF ${pdfPath}:`, cleanupErr);
  }
}

// ============================================================
// ESTIMATOR DATA (Bill's EI method + Justin's factor method)
// ============================================================

let estimatorDataCache: any = null;

function getEstimatorData(): any {
  if (estimatorDataCache) return estimatorDataCache;
  const dataPath = path.join(__dirname, "estimator-data.json");
  try {
    estimatorDataCache = JSON.parse(fs.readFileSync(dataPath, "utf-8"));
  } catch {
    // fallback: try relative to project root
    const altPath = path.join(process.cwd(), "server", "estimator-data.json");
    estimatorDataCache = JSON.parse(fs.readFileSync(altPath, "utf-8"));
  }
  return estimatorDataCache;
}

/** Parse a size string like '4"', '6"', '4 inch', '3/4"' → numeric NPS */
function parseSizeNPS(sizeStr: string): number {
  if (!sizeStr) return 0;
  const s = sizeStr.toString().trim().toLowerCase();

  // Handle reducer compound sizes like "6x4" or "6\"x4\"" — return the LARGER size
  const reducerMatch = s.match(/(\d+(?:[\-\s]?\d+\/\d+)?)[\s"''″]*x[\s"''″]*(\d+(?:[\-\s]?\d+\/\d+)?)/i);
  if (reducerMatch) {
    const a = parseSizeNPS(reducerMatch[1]);
    const b = parseSizeNPS(reducerMatch[2]);
    return Math.max(a, b);
  }

  // Handle fractions — sorted LONGEST-FIRST to prevent substring matching
  const fracMapSorted: [string, number][] = [
    ["2-1/2", 2.5], ["1-1/2", 1.5], ["1-1/4", 1.25],
    ["3/4", 0.75], ["1/2", 0.5], ["3/8", 0.375], ["1/4", 0.25],
  ];
  for (const [k, v] of fracMapSorted) {
    if (s.includes(k)) return v;
  }
  // Extract leading number (handles '6"', '6 inch', '6\'', '6 in', '06')
  const m = s.match(/(\d+\.?\d*)/);
  if (m) return parseFloat(m[1]);
  return 0;
}

/** Find closest numeric key in a record. Returns { key, exact } */
function findClosestKey(table: Record<string, any>, target: number): { key: string; exact: boolean } | null {
  const keys = Object.keys(table).map(Number).filter(k => !isNaN(k));
  if (keys.length === 0) return null;
  let best = keys[0];
  let bestDiff = Math.abs(keys[0] - target);
  for (const k of keys) {
    const diff = Math.abs(k - target);
    if (diff < bestDiff) { bestDiff = diff; best = k; }
  }
  return { key: String(best), exact: bestDiff < 0.01 };
}

/** Normalize schedule: '40' -> '40', 'sch40' -> '40', 'std' -> 'STD', etc. */
function normalizeSchedule(sched: string): string {
  if (!sched) return "STD";
  const s = sched.toString().toUpperCase().replace(/^SCH\s*/i, "").trim();
  if (s === "STANDARD" || s === "STD") return "STD";
  if (s === "XH" || s === "EXTRA HEAVY" || s === "XS") return "XH";
  if (s === "XXH" || s === "160/XXH" || s === "160") return "160/XXH";
  return s;
}

/** Normalize pressure rating string: '150#', '150 LB', '#150', '150' → '150' */
function normalizeRating(rating: string): string {
  if (!rating) return "150";
  const m = rating.toString().match(/(\d+)/);
  return m ? m[1] : "150";
}

/** Determine if item is weld/butt-weld type */
function isWeldItem(item: any): boolean {
  const cat = (item.category || "").toLowerCase();
  const desc = (item.description || "").toLowerCase();
  return cat === "fitting" || desc.includes("weld") || desc.includes("butt") || desc.includes("bw");
}

/** Determine if item is a flanged joint / bolt-up */
function isFlangedItem(item: any): boolean {
  const cat = (item.category || "").toLowerCase();
  const desc = (item.description || "").toLowerCase();
  return (
    desc.includes("flange") || desc.includes("bolt") || desc.includes("bolt-up") ||
    cat === "fitting" && desc.includes("flange")
  );
}

/** Determine if item is a valve */
function isValveItem(item: any): boolean {
  return (item.category || "").toLowerCase() === "valve" || (item.description || "").toLowerCase().includes("valve");
}

/** Determine if item is a support/shoe/other misc */
function isOtherItem(item: any): boolean {
  const desc = (item.description || "").toLowerCase();
  return desc.includes("shoe") || desc.includes("support") || desc.includes("hydro") ||
         desc.includes("grout") || desc.includes("demo");
}

/**
 * Calculate labor hours per unit using Bill's EI method.
 * Returns full result with breakdown string, size match info, and material cost source.
 */
function calculateBillLaborHours(
  item: any,
  material: "CS" | "SS",
  schedule: string,
  billData: any,
  pipeLocation: string,
  elevation: string,
  alloyGroup: string,
  fittingWeldMode: "bundled" | "separate" = "bundled"
): { laborHoursPerUnit: number; materialUnitCostAdjust: number; calcBasis: string; sizeMatchExact: boolean; materialCostSource: string } {
  const nps = parseSizeNPS(item.size);
  const sched = normalizeSchedule(schedule);
  const lr = billData.labor_rates;
  const mhEi = billData.mh_per_ei || { CS: { field: 0.20 }, SS: { field: 0.40 }, material_cost_per_ei: { CS: 0.10, SS: 0.40 } };

  // MH per EI rates
  const fieldMhPerEi = material === "SS" ? (mhEi.SS?.field || 0.40) : (mhEi.CS?.field || 0.20);
  const matCostPerEi = material === "SS" ? (mhEi.material_cost_per_ei?.SS || 0.40) : (mhEi.material_cost_per_ei?.CS || 0.10);

  // Elevation factor
  const elevFactors = billData.elevation_factors || {};
  let elevFactor = 1.0;
  if (elevation === "20-40ft") elevFactor = elevFactors["20-40ft"] || 1.05;
  else if (elevation === "40-80ft") elevFactor = elevFactors["40-80ft"] || 1.10;
  else if (elevation === "80ft+") elevFactor = elevFactors["80ft+"] || 1.20;
  else elevFactor = elevFactors["0-20ft"] || 1.0;

  // Pipe location factor
  const locFactors = billData.pipe_location_factors || {};
  let locationFactor = 1.0;
  if (pipeLocation === "Sleeper Rack") locationFactor = locFactors["Sleeper Rack"] || 0.6;
  else if (pipeLocation === "Underground") locationFactor = locFactors["Underground"] || 0.75;
  else if (pipeLocation === "Open Rack") locationFactor = locFactors["Open Rack"] || 0.8;
  else if (pipeLocation === "Elevated Rack") locationFactor = locFactors["Elevated Rack"] || 1.0;
  else locationFactor = 0.8;

  // FIX 5: Weld-specific location factor (different from pipe handling)
  let weldLocationFactor = 1.0;
  if (pipeLocation === "Sleeper Rack") weldLocationFactor = 0.85;
  else if (pipeLocation === "Underground") weldLocationFactor = 0.9;
  else if (pipeLocation === "Open Rack") weldLocationFactor = 1.0;
  else if (pipeLocation === "Elevated Rack") weldLocationFactor = 1.1;
  else weldLocationFactor = 1.0;

  // Alloy factor (multiplied against welding/fitting labor for non-CS materials)
  let alloyFactor = 1.0;
  if (material === "SS" && billData.alloy_factors?.operations) {
    const group = alloyGroup || "4"; // Default: SS = group 4
    const weldFac = billData.alloy_factors.operations["WELDING"];
    if (weldFac && weldFac[group]) alloyFactor = weldFac[group];
  }

  const catLower = (item.category || "").toLowerCase();
  const descLower = (item.description || "").toLowerCase();

  // Calibration: small-bore items have MH rolled into weld factors
  if (item.smallBoreRollup) {
    return { mh: 0, calcBasis: "Included in weld factor (small-bore rollup)", sizeMatchExact: true };
  }
  let sizeMatchExact = true;

  // Helper for size warnings
  function sizeWarn(found: { key: string; exact: boolean } | null, target: number): string {
    if (!found) return "";
    if (!found.exact) { sizeMatchExact = false; return ` \u26A0 Size ${target}\" not in table, used nearest ${found.key}\"`; }
    return "";
  }

  // --- PIPE: pipe_handling_mh_per_lf × location factor × elevation ---
  if (catLower === "pipe" || (descLower.includes("pipe") && !descLower.includes("support"))) {
    const pipeTable = lr.pipe_handling_mh_per_lf;
    const found = findClosestKey(pipeTable, nps);
    if (found && pipeTable[found.key]) {
      const schedKey = sched in pipeTable[found.key] ? sched : ("STD" in pipeTable[found.key] ? "STD" : Object.keys(pipeTable[found.key])[0]);
      const baseMH = pipeTable[found.key][schedKey] || 0;
      const pipeMatFactor = material === "SS" ? 1.15 : 1.0;
      const mh = baseMH * locationFactor * elevFactor * pipeMatFactor;
      const warn = sizeWarn(found, nps);
      const basis = `Bill's EI: Pipe ${found.key}\" ${schedKey} → base=${baseMH.toFixed(4)} MH/LF × ${locationFactor} (${pipeLocation}) × ${elevFactor} (${elevation})${material === "SS" ? " × 1.15 (SS wt)" : ""} = ${mh.toFixed(4)} MH/LF${warn}`;
      return { laborHoursPerUnit: mh, materialUnitCostAdjust: 0, calcBasis: basis, sizeMatchExact, materialCostSource: "" };
    }
  }

  // --- BUTT WELD: EI × MH/EI × alloy factor × elevation × weldLocationFactor ---
  if (descLower.includes("butt weld") || descLower.includes("butt-weld") || descLower.includes("bw ") || (descLower.includes("weld") && !descLower.includes("socket") && !descLower.includes("fillet") && !descLower.includes("nozzle") && !descLower.includes("miter") && !descLower.includes("slip"))) {
    const weldTable = lr.butt_welds_ei;
    const found = findClosestKey(weldTable, nps);
    if (found && weldTable[found.key]) {
      const schedKey = sched in weldTable[found.key] ? sched : ("STD" in weldTable[found.key] ? "STD" : Object.keys(weldTable[found.key])[0]);
      const ei = weldTable[found.key][schedKey] || 0;
      const mh = ei * fieldMhPerEi * alloyFactor * elevFactor * weldLocationFactor;
      const warn = sizeWarn(found, nps);
      const basis = `Bill's EI: BW ${found.key}\" ${schedKey} → EI=${ei} × ${fieldMhPerEi} MH/EI (${material} field) × ${elevFactor} (${elevation}) × ${alloyFactor} (alloy ${material}) × ${weldLocationFactor} (${pipeLocation} weld) = ${mh.toFixed(4)} MH${warn}`;
      return { laborHoursPerUnit: mh, materialUnitCostAdjust: ei * matCostPerEi, calcBasis: basis, sizeMatchExact, materialCostSource: "allowance" };
    }
  }

  // --- SOCKET WELD: from actual SW EI table lookup ---
  if (descLower.includes("socket") || descLower.includes("sw ") || descLower.includes("s/w")) {
    const swTable = lr.socket_weld_ei;
    if (swTable) {
      const found = findClosestKey(swTable, nps);
      if (found && swTable[found.key]) {
        const entry = swTable[found.key];
        // Determine which SW sub-key to use
        let eiKey = "socket_40_80"; // default for socket welds
        let eiLabel = "socket_40_80";
        if (descLower.includes("coupling") || descLower.includes("half coupling")) {
          eiKey = "coupling_3000"; eiLabel = "coupling_3000";
        } else if (descLower.includes("olet") || descLower.includes("sockolet") || descLower.includes("weldolet") || descLower.includes("threadolet")) {
          eiKey = "olet_3000"; eiLabel = "olet_3000";
        }
        const ei = entry[eiKey] || entry.socket_40_80 || 0;
        const mh = ei * fieldMhPerEi * alloyFactor * elevFactor * weldLocationFactor;
        const warn = sizeWarn(found, nps);
        const basis = `Bill's EI: SW ${found.key}\" → SW EI (${eiLabel}): ${ei} × ${fieldMhPerEi} MH/EI (${material} field) × ${elevFactor} (${elevation}) × ${alloyFactor} (alloy ${material}) × ${weldLocationFactor} (${pipeLocation} weld) = ${mh.toFixed(4)} MH${warn}`;
        return { laborHoursPerUnit: mh, materialUnitCostAdjust: ei * matCostPerEi, calcBasis: basis, sizeMatchExact, materialCostSource: "allowance" };
      }
    }
    // Fallback to old BW × 0.65 method if SW table not available
    const weldTable = lr.butt_welds_ei;
    const found = findClosestKey(weldTable, nps);
    if (found && weldTable[found.key]) {
      const schedKey = sched in weldTable[found.key] ? sched : ("STD" in weldTable[found.key] ? "STD" : Object.keys(weldTable[found.key])[0]);
      const baseEi = weldTable[found.key][schedKey] || 0;
      const ei = baseEi * 0.65;
      const mh = ei * fieldMhPerEi * alloyFactor * elevFactor * weldLocationFactor;
      const warn = sizeWarn(found, nps);
      const basis = `Bill's EI: SW ${found.key}\" ${schedKey} → BW EI=${baseEi} × 0.65 = ${ei.toFixed(1)} EI (fallback) × ${fieldMhPerEi} MH/EI × ${elevFactor} (elev) × ${alloyFactor} (alloy) × ${weldLocationFactor} (${pipeLocation} weld) = ${mh.toFixed(4)} MH${warn}`;
      return { laborHoursPerUnit: mh, materialUnitCostAdjust: ei * matCostPerEi, calcBasis: basis, sizeMatchExact, materialCostSource: "allowance" };
    }
  }

  // --- 90 NOZZLE WELD ---
  if (descLower.includes("nozzle") || descLower.includes("olet") || descLower.includes("sockolet") || descLower.includes("weldolet") || descLower.includes("threadolet")) {
    const nozzleTable = billData.reinforced_90_nozzle_ei || {};
    const found = findClosestKey(nozzleTable, nps);
    if (found && nozzleTable[found.key]) {
      const schedKey = sched in nozzleTable[found.key] ? sched : ("STD" in nozzleTable[found.key] ? "STD" : Object.keys(nozzleTable[found.key])[0]);
      const ei = nozzleTable[found.key][schedKey] || 0;
      const mh = ei * fieldMhPerEi * alloyFactor * elevFactor * weldLocationFactor;
      const warn = sizeWarn(found, nps);
      const basis = `Bill's EI: Nozzle ${found.key}\" ${schedKey} → EI=${ei} × ${fieldMhPerEi} MH/EI × ${elevFactor} (elev) × ${alloyFactor} (alloy) × ${weldLocationFactor} (${pipeLocation} weld) = ${mh.toFixed(4)} MH${warn}`;
      return { laborHoursPerUnit: mh, materialUnitCostAdjust: ei * matCostPerEi, calcBasis: basis, sizeMatchExact, materialCostSource: "allowance" };
    }
  }

  // --- SLIP-ON FLANGE WELD ---
  if (descLower.includes("slip") && descLower.includes("flange")) {
    const weldTable = lr.butt_welds_ei;
    const found = findClosestKey(weldTable, nps);
    if (found && weldTable[found.key]) {
      const schedKey = sched in weldTable[found.key] ? sched : ("STD" in weldTable[found.key] ? "STD" : Object.keys(weldTable[found.key])[0]);
      const baseEi = weldTable[found.key][schedKey] || 0;
      const ei = baseEi * 0.7;
      const mh = ei * fieldMhPerEi * alloyFactor * elevFactor * weldLocationFactor;
      const warn = sizeWarn(found, nps);
      const basis = `Bill's EI: SO Flange ${found.key}\" ${schedKey} → BW EI=${baseEi} × 0.70 = ${ei.toFixed(1)} EI × ${fieldMhPerEi} MH/EI × ${elevFactor} (elev) × ${weldLocationFactor} (${pipeLocation} weld) = ${mh.toFixed(4)} MH${warn}`;
      return { laborHoursPerUnit: mh, materialUnitCostAdjust: ei * matCostPerEi, calcBasis: basis, sizeMatchExact, materialCostSource: "allowance" };
    }
  }

  // --- VALVE: from valve_mh table ---
  if (catLower === "valve" || descLower.includes("valve")) {
    const valveTable = billData.valve_mh || {};
    const found = findClosestKey(valveTable, nps);
    if (found && valveTable[found.key]) {
      const rating = normalizeRating(item.rating || "150");
      let mh = 0;
      let ratingLabel = "";
      if (descLower.includes("screwed") || descLower.includes("threaded") || descLower.includes("thrd")) {
        mh = valveTable[found.key].screwed || valveTable[found.key].flanged_150 || 0;
        ratingLabel = "screwed";
      } else if (Number(rating) >= 600) {
        mh = valveTable[found.key].flanged_600 || valveTable[found.key].flanged_300 || 0;
        ratingLabel = "flanged 600#";
      } else if (Number(rating) >= 300) {
        mh = valveTable[found.key].flanged_300 || valveTable[found.key].flanged_150 || 0;
        ratingLabel = "flanged 300#";
      } else {
        mh = valveTable[found.key].flanged_150 || 0;
        ratingLabel = "flanged 150#";
      }
      const finalMh = mh * elevFactor;
      const warn = sizeWarn(found, nps);
      const basis = `Bill's EI: Valve ${found.key}\" ${ratingLabel} → base=${mh} MH × ${elevFactor} (${elevation}) = ${finalMh.toFixed(4)} MH${warn}`;
      return { laborHoursPerUnit: finalMh, materialUnitCostAdjust: 0, calcBasis: basis, sizeMatchExact, materialCostSource: "" };
    }
  }

  // --- FLANGED JOINT / BOLT-UP ---
  if (descLower.includes("flange") || descLower.includes("bolt") || descLower.includes("stud") || catLower === "bolt" || catLower === "flange") {
    const flangeTable = lr.flanged_joints_mh_per_joint;
    const found = findClosestKey(flangeTable, nps);
    if (found && flangeTable[found.key]) {
      const rating = normalizeRating(item.rating || "150");
      const ratingKey = rating in flangeTable[found.key] ? rating : ("150" in flangeTable[found.key] ? "150" : Object.keys(flangeTable[found.key])[0]);
      const baseMh = flangeTable[found.key][ratingKey] || 0;
      const mh = baseMh * elevFactor;
      const warn = sizeWarn(found, nps);
      const basis = `Bill's EI: Flange/Bolt ${found.key}\" ${ratingKey}# → base=${baseMh} MH × ${elevFactor} (${elevation}) = ${mh.toFixed(4)} MH${warn}`;
      return { laborHoursPerUnit: mh, materialUnitCostAdjust: 0, calcBasis: basis, sizeMatchExact, materialCostSource: "" };
    }
  }

  // --- PIPE SUPPORT ---
  if (catLower === "support" || descLower.includes("support") || descLower.includes("shoe") || descLower.includes("hanger") || descLower.includes("ubolt") || descLower.includes("u-bolt") || descLower.includes("guide") || descLower.includes("dummy")) {
    const supUnits = billData.support_labor_units || {};
    let mh = 3;
    let supType = "generic";
    if (descLower.includes("shoe") && descLower.includes("guide") && descLower.includes("stop")) { mh = supUnits["Shoe, Guide, & Stop"]?.field_mh || 7; supType = "Shoe/Guide/Stop"; }
    else if (descLower.includes("shoe") && descLower.includes("guide")) { mh = supUnits["Shoe & Guide"]?.field_mh || 5; supType = "Shoe & Guide"; }
    else if (descLower.includes("shoe")) { mh = supUnits["Shoe"]?.field_mh || 3; supType = "Shoe"; }
    else if (descLower.includes("ubolt") || descLower.includes("u-bolt") || descLower.includes("u bolt")) { mh = supUnits["Ubolt"]?.field_mh || 1; supType = "U-Bolt"; }
    else if (descLower.includes("hanger") || descLower.includes("beam clamp")) { mh = supUnits["Beam Clamp Hanger"]?.field_mh || 3; supType = "Hanger"; }
    else if (descLower.includes("dummy") || descLower.includes("adjustable")) { mh = supUnits["Adjustable Dummy Leg"]?.field_mh || 3; supType = "Dummy Leg"; }
    const finalMh = mh * elevFactor;
    const basis = `Bill's EI: Support (${supType}) → base=${mh} MH × ${elevFactor} (${elevation}) = ${finalMh.toFixed(4)} MH`;
    return { laborHoursPerUnit: finalMh, materialUnitCostAdjust: 0, calcBasis: basis, sizeMatchExact: true, materialCostSource: "" };
  }

  // --- CUT / BEVEL ---
  if (descLower.includes("cut") || descLower.includes("bevel")) {
    const cbTable = billData.cut_bevel_mh || {};
    const found = findClosestKey(cbTable, nps);
    if (found && cbTable[found.key]) {
      const mh = descLower.includes("bevel") ? (cbTable[found.key].bevel || 0) : (cbTable[found.key].pipe_cut || 0);
      const finalMh = mh * elevFactor;
      const warn = sizeWarn(found, nps);
      const basis = `Bill's EI: Cut/Bevel ${found.key}\" → base=${mh} MH × ${elevFactor} (${elevation}) = ${finalMh.toFixed(4)} MH${warn}`;
      return { laborHoursPerUnit: finalMh, materialUnitCostAdjust: 0, calcBasis: basis, sizeMatchExact, materialCostSource: "" };
    }
  }

  // --- THREADED CONNECTION ---
  if (descLower.includes("thread") || descLower.includes("screwed")) {
    const ctTable = billData.cut_thread_mh || {};
    const found = findClosestKey(ctTable, nps);
    if (found && ctTable[found.key]) {
      const mh = (ctTable[found.key].cut_thread || 0) + (ctTable[found.key].makeup || 0);
      const finalMh = mh * elevFactor;
      const warn = sizeWarn(found, nps);
      const basis = `Bill's EI: Thread ${found.key}\" → base=${mh} MH × ${elevFactor} (${elevation}) = ${finalMh.toFixed(4)} MH${warn}`;
      return { laborHoursPerUnit: finalMh, materialUnitCostAdjust: 0, calcBasis: basis, sizeMatchExact, materialCostSource: "" };
    }
  }

  // --- GASKET ---
  if (catLower === "gasket" || descLower.includes("gasket")) {
    return { laborHoursPerUnit: 0.1, materialUnitCostAdjust: 0, calcBasis: "Bill's EI: Gasket → 0.10 MH (included in bolt-up)", sizeMatchExact: true, materialCostSource: "" };
  }

  // --- STRUCTURAL STEEL ---
  if (catLower === "steel" || descLower.includes("steel") || descLower.includes("beam") || descLower.includes("column") || descLower.includes("brace")) {
    const steelUnits = billData.steel_mh_units || {};
    const mhPerTon = steelUnits.MH_per_ton || 12;
    if ((item.unit || "").toUpperCase() === "TN" || (item.unit || "").toUpperCase() === "TON") {
      const finalMh = mhPerTon * elevFactor;
      return { laborHoursPerUnit: finalMh, materialUnitCostAdjust: 0, calcBasis: `Bill's EI: Steel → ${mhPerTon} MH/TN × ${elevFactor} (${elevation}) = ${finalMh.toFixed(4)} MH`, sizeMatchExact: true, materialCostSource: "" };
    }
    const finalMh = 0.5 * elevFactor;
    return { laborHoursPerUnit: finalMh, materialUnitCostAdjust: 0, calcBasis: `Bill's EI: Steel → 0.5 MH/unit × ${elevFactor} (${elevation}) = ${finalMh.toFixed(4)} MH`, sizeMatchExact: true, materialCostSource: "" };
  }

  // --- CIVIL / CONCRETE ---
  if (catLower === "concrete" || descLower.includes("concrete") || descLower.includes("excavat") || descLower.includes("backfill") || descLower.includes("rebar") || descLower.includes("grout") || descLower.includes("form")) {
    const civilUnits = billData.civil_mh_units || {};
    if (descLower.includes("excavat")) { const mh = civilUnits["Excavation"]?.mh_per_unit || 0.3; return { laborHoursPerUnit: mh, materialUnitCostAdjust: 0, calcBasis: `Bill's EI: Excavation → ${mh} MH/unit`, sizeMatchExact: true, materialCostSource: "" }; }
    if (descLower.includes("backfill")) { const mh = civilUnits["Backfill"]?.mh_per_unit || 0.2; return { laborHoursPerUnit: mh, materialUnitCostAdjust: 0, calcBasis: `Bill's EI: Backfill → ${mh} MH/unit`, sizeMatchExact: true, materialCostSource: "" }; }
    if (descLower.includes("form")) { const mh = civilUnits["Form"]?.mh_per_unit || 0.17; return { laborHoursPerUnit: mh, materialUnitCostAdjust: 0, calcBasis: `Bill's EI: Formwork → ${mh} MH/unit`, sizeMatchExact: true, materialCostSource: "" }; }
    if (descLower.includes("rebar")) { const mh = civilUnits["Rebar"]?.mh_per_unit || 0.015; return { laborHoursPerUnit: mh, materialUnitCostAdjust: 0, calcBasis: `Bill's EI: Rebar → ${mh} MH/unit`, sizeMatchExact: true, materialCostSource: "" }; }
    if (descLower.includes("grout")) { const mh = civilUnits["Grout"]?.mh_per_unit || 4.0; return { laborHoursPerUnit: mh, materialUnitCostAdjust: 0, calcBasis: `Bill's EI: Grout → ${mh} MH/unit`, sizeMatchExact: true, materialCostSource: "" }; }
    if (descLower.includes("anchor")) { const mh = civilUnits["Anchor Bolts"]?.mh_per_unit || 1.0; return { laborHoursPerUnit: mh, materialUnitCostAdjust: 0, calcBasis: `Bill's EI: Anchor Bolts → ${mh} MH/unit`, sizeMatchExact: true, materialCostSource: "" }; }
    if (descLower.includes("concrete")) { const mh = civilUnits["Concrete"]?.mh_per_unit || 1.0; return { laborHoursPerUnit: mh, materialUnitCostAdjust: 0, calcBasis: `Bill's EI: Concrete → ${mh} MH/unit`, sizeMatchExact: true, materialCostSource: "" }; }
    if (descLower.includes("fine grade")) { const mh = civilUnits["Fine Grade"]?.mh_per_unit || 0.03; return { laborHoursPerUnit: mh, materialUnitCostAdjust: 0, calcBasis: `Bill's EI: Fine Grade → ${mh} MH/unit`, sizeMatchExact: true, materialCostSource: "" }; }
    if (descLower.includes("finish")) { const mh = civilUnits["Finish"]?.mh_per_unit || 0.03; return { laborHoursPerUnit: mh, materialUnitCostAdjust: 0, calcBasis: `Bill's EI: Finish → ${mh} MH/unit`, sizeMatchExact: true, materialCostSource: "" }; }
  }

  // --- HYDROTESTING ---
  if (descLower.includes("hydro") || descLower.includes("test")) {
    const hydroFactor = billData.hydrotesting?.field_hydro_factor || 0.1;
    return { laborHoursPerUnit: hydroFactor, materialUnitCostAdjust: 0, calcBasis: `Bill's EI: Hydro → ${hydroFactor} MH/unit`, sizeMatchExact: true, materialCostSource: "" };
  }

  // --- GENERIC FITTING (elbow, tee, reducer, cap, coupling, union) ---
  // Per-fitting weld-end multipliers come from billData.weld_end_multipliers,
  // falling back to Bill's historical hardcoded values when no table exists.
  // In "separate" weld mode the BOM already carries explicit weld rows for
  // each weld end, so the fitting contributes only handling (~0.15 × weld EI);
  // in "bundled" mode the fitting's labor includes its welds.
  if (catLower === "fitting" || catLower === "elbow" || catLower === "tee" || catLower === "reducer" || catLower === "cap" || catLower === "coupling" || catLower === "union") {
    const wem = billData.weld_end_multipliers || {};
    let fittingKey = "fitting";
    let fittingType = "fitting";
    let legacyMult = 0.6;
    if (descLower.includes("90") && descLower.includes("elbow") || (catLower === "elbow" && descLower.includes("90"))) { fittingKey = "elbow_90"; legacyMult = 1.0; fittingType = "90° Elbow"; }
    else if (descLower.includes("45") && descLower.includes("elbow") || (catLower === "elbow" && descLower.includes("45"))) { fittingKey = "elbow_45"; legacyMult = 0.8; fittingType = "45° Elbow"; }
    else if (catLower === "elbow" || descLower.includes("elbow") || descLower.includes("ell") || descLower.includes("return")) { fittingKey = "elbow"; legacyMult = 1.0; fittingType = "Elbow"; }
    else if (catLower === "tee" || descLower.includes("tee")) { fittingKey = "tee"; legacyMult = 1.3; fittingType = "Tee"; }
    else if (catLower === "reducer" || descLower.includes("reducer") || descLower.includes("swage")) { fittingKey = "reducer"; legacyMult = 0.7; fittingType = "Reducer"; }
    else if (catLower === "cap" || descLower.includes("cap")) { fittingKey = "cap"; legacyMult = 0.5; fittingType = "Cap"; }
    else if (catLower === "coupling" || descLower.includes("coupling") || descLower.includes("nipple")) { fittingKey = "coupling"; legacyMult = 0.6; fittingType = "Coupling"; }
    else if (catLower === "union" || descLower.includes("union")) { fittingKey = "union"; legacyMult = 0.4; fittingType = "Union"; }

    const bundledMult = (wem[fittingKey] !== undefined ? wem[fittingKey] : (wem["fitting"] !== undefined ? wem["fitting"] : legacyMult));
    const fittingEiMult = fittingWeldMode === "separate" ? 0.15 : bundledMult;

    const weldTable = lr.butt_welds_ei;
    const found = findClosestKey(weldTable, nps);
    if (found && weldTable[found.key]) {
      const schedKey = sched in weldTable[found.key] ? sched : ("STD" in weldTable[found.key] ? "STD" : Object.keys(weldTable[found.key])[0]);
      const baseEi = weldTable[found.key][schedKey] || 0;
      const ei = baseEi * fittingEiMult;
      const mh = ei * fieldMhPerEi * alloyFactor * elevFactor * weldLocationFactor;
      const warn = sizeWarn(found, nps);
      const modeNote = fittingWeldMode === "separate" ? "separate-welds: handling only" : `bundled ×${bundledMult} weld-ends`;
      const basis = `Bill's EI: ${fittingType} ${found.key}\" ${schedKey} → BW EI=${baseEi} × ${fittingEiMult} (${modeNote}) = ${ei.toFixed(1)} EI × ${fieldMhPerEi} MH/EI × ${elevFactor} (elev) × ${alloyFactor} (alloy) × ${weldLocationFactor} (${pipeLocation} weld) = ${mh.toFixed(4)} MH${warn}`;
      return { laborHoursPerUnit: mh, materialUnitCostAdjust: 0, calcBasis: basis, sizeMatchExact, materialCostSource: "" };
    }
  }

  return { laborHoursPerUnit: 0, materialUnitCostAdjust: 0, calcBasis: "Bill's EI: No matching table entry", sizeMatchExact: true, materialCostSource: "" };
}

/**
 * Calculate labor hours per unit using Justin's factor method.
 * Accounts for: pipe (std/rack/SS), welds (STD/SCH80/SS), bolts (150#/300#),
 * valves (150#/300#), threads (by size), shoes, supports, hydro, grout,
 * demo, exchangers, ID tags, assist, contingency, supervision.
 */
function calculateJustinLaborHours(
  item: any,
  installType: "standard" | "rack",
  material: "CS" | "SS",
  schedule: string,
  justinData: any,
  fittingWeldMode: "bundled" | "separate" = "bundled"
): { mh: number; calcBasis: string; sizeMatchExact: boolean } {
  const nps = parseSizeNPS(item.size);
  const factors = justinData.labor_factors;
  const catLower = (item.category || "").toLowerCase();
  const descLower = (item.description || "").toLowerCase();

  function findBestMatch(table: Record<string, any>): { val: any; matchKey: string; exact: boolean } | null {
    const entries = Object.entries(table);
    let best: any = null;
    let bestKey = "";
    let bestDiff = Infinity;
    for (const [key, val] of entries) {
      const entryNps = parseSizeNPS(key);
      const diff = Math.abs(entryNps - nps);
      if (diff < bestDiff) { bestDiff = diff; best = val; bestKey = key; }
    }
    if (!best) return null;
    return { val: best, matchKey: bestKey, exact: bestDiff < 0.01 };
  }

  // Helper for size warnings in Justin's method
  function jSizeWarn(match: { matchKey: string; exact: boolean } | null): string {
    if (!match || match.exact) return "";
    return ` \u26A0 used nearest ${match.matchKey}`;
  }

  // --- PIPE: 3 columns — Standard, Rack, and SS ---
  if (catLower === "pipe" || (descLower.includes("pipe") && !descLower.includes("support") && !descLower.includes("shoe"))) {
    const match = findBestMatch(factors.pipe || {});
    if (match) {
      let mh = 0;
      let col = "";
      if (material === "SS") { mh = match.val.ss_mh_per_lf || (match.val.standard_mh_per_lf || 0) * 1.8; col = "SS"; }
      else if (installType === "rack") { mh = match.val.rack_mh_per_lf || 0; col = "rack"; }
      else { mh = match.val.standard_mh_per_lf || 0; col = "standard"; }
      const warn = jSizeWarn(match);
      return { mh, calcBasis: `Justin: Pipe ${match.matchKey} (${col}) → ${mh.toFixed(4)} MH/LF${warn}`, sizeMatchExact: match.exact };
    }
  }

  // --- WELD: STD, SCH 80, and SS columns ---
  if (descLower.includes("weld") || descLower.includes("butt") || descLower.includes("bw")) {
    const match = findBestMatch(factors.welds || {});
    if (match) {
      const schedNorm = normalizeSchedule(schedule);
      let mh = 0; let col = "";
      if (material === "SS") {
        // Check calibrated SS weld factors first
        const sizeKey = parseSizeNPS(item.size || "").toString();
        const calibProfile = Object.values(CALIBRATION_DATA)[0];
        const calibFactor = calibProfile?.ss_weld_factors?.[sizeKey];
        if (calibFactor) {
          mh = calibFactor.mh_per_weld;
          col = `SS-CAL(${calibProfile.project.slice(0,15)})`;
        } else {
          mh = match.val.ss_mh_per_weld || 0;
          col = "SS";
        }
      }
      else if (schedNorm === "80" || schedNorm === "XH") { mh = match.val.sch80_mh_per_weld || 0; col = "SCH80"; }
      else { mh = match.val.std_mh_per_weld || 0; col = "STD"; }
      const warn = jSizeWarn(match);
      return { mh, calcBasis: `Justin: Weld ${match.matchKey} (${col}) → ${mh.toFixed(4)} MH/weld${warn}`, sizeMatchExact: match.exact };
    }
  }

  // --- THREADED CONNECTION ---
  if (descLower.includes("thread") || descLower.includes("screwed") || descLower.includes("thrd")) {
    const threads = factors.threads || {};
    let mh = 0; let sz = "";
    if (nps <= 1) { mh = threads['1"']?.mh_per_connection || 0.5; sz = '1"'; }
    else if (nps <= 2) { mh = threads['2"']?.mh_per_connection || 1.3; sz = '2"'; }
    else { mh = threads['4"']?.mh_per_connection || 2.0; sz = '4"'; }
    return { mh, calcBasis: `Justin: Thread ${sz} → ${mh} MH/connection`, sizeMatchExact: true };
  }

  // --- BOLT-UP / FLANGE ---
  if (descLower.includes("bolt") || descLower.includes("stud") || catLower === "bolt" || (descLower.includes("flange") && !descLower.includes("weld"))) {
    const match = findBestMatch(factors.bolts || {});
    if (match) {
      const rating = normalizeRating(item.rating || "150");
      const mh = Number(rating) >= 300 ? (match.val["300_mh_per_set"] || 0) : (match.val["150_mh_per_set"] || 0);
      const warn = jSizeWarn(match);
      // Calibration: shop bolt-ups have MH included in shop fab
      if (item.installLocation === "shop") {
        return { mh: 0, calcBasis: `Justin: Bolt ${match.matchKey} (${rating}#) shop bolt-up (MH included in shop fab)${warn}`, sizeMatchExact: match.exact };
      }
      return { mh, calcBasis: `Justin: Bolt ${match.matchKey} (${rating}#) → ${mh} MH/set${warn}`, sizeMatchExact: match.exact };
    }
  }

  // --- VALVE ---
  if (catLower === "valve" || descLower.includes("valve")) {
    const match = findBestMatch(factors.valves || {});
    if (match) {
      const rating = normalizeRating(item.rating || "150");
      const mh = Number(rating) >= 300 ? (match.val["300_mh_per_valve"] || 0) : (match.val["150_mh_per_valve"] || 0);
      const warn = jSizeWarn(match);
      return { mh, calcBasis: `Justin: Valve ${match.matchKey} (${rating}#) → ${mh} MH/valve${warn}`, sizeMatchExact: match.exact };
    }
  }

  // --- SHOE ---
  if (descLower.includes("shoe") && !descLower.includes("pipe shoe")) {
    const mh = factors.other?.Shoes?.factor || 1.5;
    return { mh, calcBasis: `Justin: Shoe → ${mh} MH`, sizeMatchExact: true };
  }

  // --- SUPPORT ---
  if (catLower === "support" || descLower.includes("support") || descLower.includes("hanger") || descLower.includes("ubolt") || descLower.includes("u-bolt") || descLower.includes("dummy")) {
    const mh = factors.other?.Supports?.factor || 2.0;
    return { mh, calcBasis: `Justin: Support → ${mh} MH`, sizeMatchExact: true };
  }

  // --- HYDROTESTING ---
  if (descLower.includes("hydro") || descLower.includes("test")) { const mh = factors.other?.Hydro?.factor || 20.0; return { mh, calcBasis: `Justin: Hydro → ${mh} MH`, sizeMatchExact: true }; }

  // --- GROUT ---
  if (descLower.includes("grout")) { const mh = factors.other?.Grout?.factor || 4.0; return { mh, calcBasis: `Justin: Grout → ${mh} MH`, sizeMatchExact: true }; }

  // --- DEMO ---
  if (descLower.includes("demo") || descLower.includes("demolit")) { const mh = factors.other?.Demo?.factor || 1.0; return { mh, calcBasis: `Justin: Demo → ${mh} MH`, sizeMatchExact: true }; }

  // --- EXCHANGER ---
  if (descLower.includes("exchanger") || descLower.includes("heat ex")) { const mh = factors.other?.Exchanger?.factor || 1.0; return { mh, calcBasis: `Justin: Exchanger → ${mh} MH`, sizeMatchExact: true }; }

  // --- ID TAGS ---
  if (descLower.includes("tag") || descLower.includes("id tag")) { const mh = factors.other?.["ID Tags"]?.factor || 1.0; return { mh, calcBasis: `Justin: ID Tags → ${mh} MH`, sizeMatchExact: true }; }

  // --- ASSIST ---
  if (descLower.includes("assist")) { const mh = factors.other?.Assist?.factor || 1.0; return { mh, calcBasis: `Justin: Assist → ${mh} MH`, sizeMatchExact: true }; }

  // --- GASKET ---
  if (catLower === "gasket" || descLower.includes("gasket")) return { mh: 0.1, calcBasis: "Justin: Gasket → 0.10 MH (in bolt-up)", sizeMatchExact: true };

  // --- FITTING (tee / elbow / reducer / cap / coupling / union / generic) ---
  // Per-category weld-end multipliers come from justinData.weld_end_multipliers,
  // falling back to 0.5 (the legacy single-multiplier value) if no table exists.
  // In "separate" mode the BOM is assumed to already have full-factor weld rows
  // for every weld end, so the fitting itself contributes only handling labor
  // (~0.15 × weld_factor); in "bundled" mode the fitting's labor includes its
  // welds via the configured weld-end multiplier.
  if (catLower === "fitting" || catLower === "elbow" || catLower === "tee" || catLower === "reducer" || catLower === "cap" || catLower === "coupling" || catLower === "union") {
    const weldMatch = findBestMatch(factors.welds || {});
    if (weldMatch) {
      const schedNorm = normalizeSchedule(schedule);
      const baseMH = material === "SS" ? (weldMatch.val.ss_mh_per_weld || 0) : (schedNorm === "80" ? (weldMatch.val.sch80_mh_per_weld || 0) : (weldMatch.val.std_mh_per_weld || 0));

      // Classify the fitting subtype so we can pick the right multiplier.
      const wem = justinData.weld_end_multipliers || {};
      let fittingKey = "fitting";
      let fittingLabel = "Fitting";
      if (descLower.includes("90") && (catLower === "elbow" || descLower.includes("elbow"))) { fittingKey = "elbow_90"; fittingLabel = "90° Elbow"; }
      else if (descLower.includes("45") && (catLower === "elbow" || descLower.includes("elbow"))) { fittingKey = "elbow_45"; fittingLabel = "45° Elbow"; }
      else if (catLower === "elbow" || descLower.includes("elbow") || descLower.includes("ell") || descLower.includes("return")) { fittingKey = "elbow"; fittingLabel = "Elbow"; }
      else if (catLower === "tee" || descLower.includes("tee")) { fittingKey = "tee"; fittingLabel = "Tee"; }
      else if (catLower === "reducer" || descLower.includes("reducer") || descLower.includes("swage")) { fittingKey = "reducer"; fittingLabel = "Reducer"; }
      else if (catLower === "cap" || descLower.includes("cap")) { fittingKey = "cap"; fittingLabel = "Cap"; }
      else if (catLower === "coupling" || descLower.includes("coupling") || descLower.includes("nipple")) { fittingKey = "coupling"; fittingLabel = "Coupling"; }
      else if (catLower === "union" || descLower.includes("union")) { fittingKey = "union"; fittingLabel = "Union"; }

      const bundledMult = (wem[fittingKey] !== undefined ? wem[fittingKey] : (wem["fitting"] !== undefined ? wem["fitting"] : 0.5));
      const effectiveMult = fittingWeldMode === "separate" ? 0.15 : bundledMult;
      const mh = baseMH * effectiveMult;
      const warn = jSizeWarn(weldMatch);
      const modeNote = fittingWeldMode === "separate" ? "separate-welds: handling only" : `bundled ×${bundledMult} weld-ends`;
      return { mh, calcBasis: `Justin: ${fittingLabel} ${weldMatch.matchKey} → weld base=${baseMH.toFixed(2)} × ${effectiveMult} (${modeNote}) = ${mh.toFixed(4)} MH${warn}`, sizeMatchExact: weldMatch.exact };
    }
  }

  // --- STRUCTURAL STEEL ---
  if (catLower === "steel" || descLower.includes("steel") || descLower.includes("beam") || descLower.includes("column")) {
    return { mh: 0.5, calcBasis: "Justin: Steel → 0.5 MH/unit", sizeMatchExact: true };
  }

  // --- CONCRETE / CIVIL ---
  if (catLower === "concrete" || descLower.includes("concrete") || descLower.includes("excavat") || descLower.includes("backfill")) {
    if (descLower.includes("excavat")) return { mh: 0.3, calcBasis: "Justin: Excavation → 0.3 MH", sizeMatchExact: true };
    if (descLower.includes("backfill")) return { mh: 0.2, calcBasis: "Justin: Backfill → 0.2 MH", sizeMatchExact: true };
    if (descLower.includes("concrete")) return { mh: 1.0, calcBasis: "Justin: Concrete → 1.0 MH", sizeMatchExact: true };
    if (descLower.includes("form")) return { mh: 0.17, calcBasis: "Justin: Formwork → 0.17 MH", sizeMatchExact: true };
    if (descLower.includes("rebar")) return { mh: 0.015, calcBasis: "Justin: Rebar → 0.015 MH", sizeMatchExact: true };
  }

  return { mh: 0, calcBasis: "Justin: No matching factor", sizeMatchExact: true };
}

// ============================================================
// INDUSTRY-STANDARD METHOD (Page's Estimator's Piping Man-Hour Manual, 5e)
// ============================================================
//
// The Industry method's data block in estimator-data.json mirrors Justin's
// shape exactly (pipe / welds / bolts / valves / threads / other) so we can
// reuse Justin's calculator logic without duplication. The only difference
// is the data source: Page's published factors vs Justin's workbook factors.
//
// Per-item calculation differences (Industry vs Justin):
//   - Industry pipe handling MH is slightly lower than Justin's (Page is
//     CS-baseline; Justin includes some project-specific overhead).
//   - Industry weld MH is lower than Justin's (Page CS field weld
//     productivity is ~1.5 MH/in-dia vs Justin's measured higher rates).
//   - SS multiplier ~1.7x CS per Page Section Three alloy factors.
//   - SCH 80 / XH multiplier ~1.2-1.25x STD per Page heavy-wall tables.
//
// Returns { mh, calcBasis, sizeMatchExact } with calcBasis prefixed
// "Industry: ..." so the estimator can see where each number came from.

function calculateIndustryLaborHours(
  item: any,
  installType: "standard" | "rack",
  material: "CS" | "SS",
  schedule: string,
  industryData: any,
  fittingWeldMode: "bundled" | "separate" = "bundled"
): { mh: number; calcBasis: string; sizeMatchExact: boolean } {
  // Defensive: if the Industry data block is missing or malformed, fail with
  // a clear message instead of a cryptic 'Cannot read properties of undefined'.
  // This can happen if estimator-data.json wasn't bundled into dist/, or if
  // a deploy is using an older build that predates the Industry method.
  if (!industryData || typeof industryData !== "object" || !industryData.labor_factors) {
    throw new Error(
      "Industry method data is unavailable on this server. " +
      "This usually means the deployed build is missing server/estimator-data.json " +
      "or predates commit e262014. Try redeploying the latest from main, or fall back " +
      "to Bill's or Justin's method."
    );
  }
  // Reuse the Justin calculator logic against the Industry data block.
  // The shapes match, so this works directly. We post-process calcBasis
  // to rename the prefix so the estimator can see the source clearly.
  const result = calculateJustinLaborHours(item, installType, material, schedule, industryData, fittingWeldMode);
  return {
    ...result,
    calcBasis: result.calcBasis.replace(/^Justin:/, "Industry (Page):"),
  };
}

// ============================================================
// CUSTOM METHOD OVERRIDES
// ============================================================
//
// applyCustomOverrides clones a base method's data tree and applies a
// CustomEstimatorMethod's override map. Override keys are dot-paths into
// the data tree (e.g. "labor_factors.welds.4\"Welds.std_mh_per_weld").
// Returns a fresh data object that can be passed to the base calculator.

// Splits a dot-delimited path while honoring backslash-escaped dots, so that
// keys that literally contain '.' (e.g. Bill's pipe sizes like "0.25") can be
// addressed safely. Example:
//   "labor_rates.butt_welds_ei.0\\.25.STD" -> ["labor_rates", "butt_welds_ei", "0.25", "STD"]
function splitOverridePath(keyPath: string): string[] {
  const parts: string[] = [];
  let buf = "";
  for (let i = 0; i < keyPath.length; i++) {
    const c = keyPath[i];
    if (c === "\\" && i + 1 < keyPath.length && keyPath[i + 1] === ".") {
      buf += ".";
      i++; // skip the escaped dot
    } else if (c === ".") {
      parts.push(buf);
      buf = "";
    } else {
      buf += c;
    }
  }
  parts.push(buf);
  return parts;
}

function applyCustomOverrides(baseData: any, overrides: Record<string, any>): any {
  if (!overrides || Object.keys(overrides).length === 0) return baseData;
  // Deep clone so we don't mutate the cached estimator data.
  const cloned = JSON.parse(JSON.stringify(baseData));
  for (const [keyPath, value] of Object.entries(overrides)) {
    if (typeof keyPath !== "string" || keyPath.length === 0) continue;
    const parts = splitOverridePath(keyPath);
    let cursor: any = cloned;
    for (let i = 0; i < parts.length - 1; i++) {
      const k = parts[i];
      if (cursor[k] === undefined || cursor[k] === null || typeof cursor[k] !== "object") {
        cursor[k] = {};
      }
      cursor = cursor[k];
    }
    cursor[parts[parts.length - 1]] = value;
  }
  return cloned;
}

// ============================================================
// ESTIMATING HELPERS
// ============================================================

function computeEstimateItem(item: any) {
  const materialExtension = (item.quantity || 0) * (item.materialUnitCost || 0);
  const laborExtension = (item.quantity || 0) * (item.laborUnitCost || 0);
  return {
    ...item,
    materialExtension,
    laborExtension,
    totalCost: materialExtension + laborExtension,
  };
}

function mapTakeoffToEstimateCategory(discipline: string, category: string): string {
  const validCategories = ["pipe", "elbow", "tee", "reducer", "valve", "flange", "gasket", "bolt", "cap", "coupling", "union", "weld", "support", "strainer", "trap", "fitting", "steel", "concrete", "rebar", "earthwork", "paving", "electrical", "other"];
  if (discipline === "mechanical") {
    // FIX 4: Preserve original takeoff category as first-class estimate category
    if (validCategories.includes(category)) return category;
    return "other";
  }
  if (discipline === "structural") {
    if (["wide_flange", "hss_tube", "angle", "channel", "plate", "base_plate", "column", "bracing", "embed_plate", "clip_angle", "gusset_plate"].includes(category)) return "steel";
    if (["footing", "grade_beam", "concrete_wall", "slab", "concrete_column"].includes(category)) return "concrete";
    if (["rebar", "wire_mesh", "dowel"].includes(category)) return "rebar";
    if (["bolt", "anchor_bolt", "weld"].includes(category)) return "steel";
    return "other";
  }
  if (discipline === "civil") {
    if (["storm_pipe", "sewer_pipe", "water_pipe", "gas_pipe"].includes(category)) return "pipe";
    if (["manhole", "catch_basin", "fire_hydrant", "valve", "fitting"].includes(category)) return "fitting";
    if (["earthwork", "backfill"].includes(category)) return "earthwork";
    if (["paving", "concrete_paving", "base_course", "curb_gutter"].includes(category)) return "paving";
    return "other";
  }
  return "other";
}

// ============================================================
// WELD INFERENCE FROM FITTINGS (Feature 2)
// ============================================================

function inferWeldsFromFittings(items: any[]): any[] {
  const welds: any[] = [];

  // Detect material and schedule from a source item so inferred welds inherit
  // them. Without this, an SS-pipe fitting would produce a weld row that the
  // calculator runs as CS (default), pulling std_mh_per_weld instead of the
  // correct ss_mh_per_weld. Same idea for SCH 80 / SCH 10S detection.
  function detectItemMaterial(it: any): "CS" | "SS" | undefined {
    if (it.itemMaterial) return it.itemMaterial as "CS" | "SS";
    const blob = `${it.description || ""} ${it.material || ""} ${it.spec || ""}`.toUpperCase();
    if (/\b(SS|STAINLESS|TP304|TP316|304L?|316L?|A312|A182|A403|A240|F304|F316)\b/.test(blob)) return "SS";
    return undefined;
  }
  function detectItemSchedule(it: any): string | undefined {
    if (it.itemSchedule) return it.itemSchedule as string;
    const blob = `${it.description || ""} ${it.schedule || ""} ${it.spec || ""}`.toUpperCase();
    if (/\bSCH\s*80\b|\bXH\b|\bXS\b/.test(blob)) return "80";
    if (/\bSCH\s*160\b|\bXXH\b/.test(blob)) return "160/XXH";
    if (/\bSCH\s*40\b/.test(blob)) return "40";
    if (/\bSCH\s*10S?\b/.test(blob)) return "10";
    if (/\bSTD\b/.test(blob)) return "STD";
    return undefined;
  }
  // Helper to merge detected metadata into the inferred weld so the downstream
  // calculator picks the right column (std / sch80 / ss) per the source's actual
  // material and schedule.
  function withInheritedMeta(source: any, weld: any): any {
    const mat = detectItemMaterial(source);
    const sch = detectItemSchedule(source);
    if (mat) weld.itemMaterial = mat;
    if (sch) weld.itemSchedule = sch;
    return weld;
  }

  // === PIPE LENGTH WELDS ===
  // Pipe is purchased in 40' standard lengths. Every 40' of pipe run requires
  // a field weld where lengths are joined together.
  // Rule: floor(length / 40) field welds per pipe item.
  // Examples:
  //   - 39' run = 0 weld (single length)
  //   - 40' run = 1 weld (one joint between lengths)
  //   - 80' run = 2 welds
  //   - 160' run = 3 welds (welds at 40', 80', 120'; the 160' point connects to a fitting)
  for (const item of items) {
    const cat = (item.category || "").toLowerCase();
    if (cat !== "pipe") continue;
    const lengthLF = item.quantity || 0;
    if (lengthLF < 40) continue;
    const pipeJointWelds = Math.floor(lengthLF / 40);
    if (pipeJointWelds === 0) continue;
    const size = item.size || "";
    welds.push(computeEstimateItem(withInheritedMeta(item, {
      id: randomUUID(), lineNumber: 0, category: "weld" as any,
      description: `FW (Field) for ${size} PIPE joints (40' lengths, auto-inferred)`,
      size, quantity: pipeJointWelds, unit: "EA",
      materialUnitCost: 0, laborUnitCost: 0, laborHoursPerUnit: 0,
      materialExtension: 0, laborExtension: 0, totalCost: 0,
      notes: `Auto-inferred: ${pipeJointWelds} FIELD pipe joint weld(s) for ${lengthLF.toFixed(1)} LF (1 weld per 40 ft of run)`, fromDatabase: false,
      weldAssumption: `${pipeJointWelds} field BW per ${lengthLF.toFixed(1)} LF run (40' standard lengths)`,
      installLocation: "field" as const,
    })));
  }

  // === FITTING WELDS ===
  for (const item of items) {
    const cat = (item.category || "").toLowerCase();
    const desc = (item.description || "").toLowerCase();
    const qty = item.quantity || 0;
    const size = item.size || "";

    // NOTE: Small-bore items (<=1.5") still generate weld counts for connection tracking.
    // The MH rollup for small-bore is handled separately in the manhour calculation.
    // We do NOT skip them here — every fitting has welds that must be counted.

    if (cat === "elbow" || desc.includes("elbow") || desc.includes("ell")) {
      welds.push(computeEstimateItem(withInheritedMeta(item, {
        id: randomUUID(), lineNumber: 0, category: "weld" as any,
        description: `BW for ${size} ELBOW (auto-inferred)`,
        size, quantity: qty * 2, unit: "EA",
        materialUnitCost: 0, laborUnitCost: 0, laborHoursPerUnit: 0,
        materialExtension: 0, laborExtension: 0, totalCost: 0,
        notes: "Auto-inferred: 2 butt welds per elbow", fromDatabase: false,
        weldAssumption: "2 butt welds per elbow (auto-inferred)",
      })));
    } else if (cat === "tee" || desc.includes("tee")) {
      welds.push(computeEstimateItem(withInheritedMeta(item, {
        id: randomUUID(), lineNumber: 0, category: "weld" as any,
        description: `BW for ${size} TEE (auto-inferred)`,
        size, quantity: qty * 3, unit: "EA",
        materialUnitCost: 0, laborUnitCost: 0, laborHoursPerUnit: 0,
        materialExtension: 0, laborExtension: 0, totalCost: 0,
        notes: "Auto-inferred: 3 butt welds per tee", fromDatabase: false,
        weldAssumption: "3 butt welds per tee (auto-inferred)",
      })));
    } else if (cat === "reducer" || desc.includes("reducer") || desc.includes("swage")) {
      welds.push(computeEstimateItem(withInheritedMeta(item, {
        id: randomUUID(), lineNumber: 0, category: "weld" as any,
        description: `BW for ${size} REDUCER (auto-inferred)`,
        size, quantity: qty * 2, unit: "EA",
        materialUnitCost: 0, laborUnitCost: 0, laborHoursPerUnit: 0,
        materialExtension: 0, laborExtension: 0, totalCost: 0,
        notes: "Auto-inferred: 2 butt welds per reducer (larger size)", fromDatabase: false,
        weldAssumption: "2 butt welds per reducer at larger size (auto-inferred)",
      })));
    } else if (cat === "cap" || (desc.includes("cap") && !desc.includes("screw"))) {
      welds.push(computeEstimateItem(withInheritedMeta(item, {
        id: randomUUID(), lineNumber: 0, category: "weld" as any,
        description: `BW for ${size} CAP (auto-inferred)`,
        size, quantity: qty * 1, unit: "EA",
        materialUnitCost: 0, laborUnitCost: 0, laborHoursPerUnit: 0,
        materialExtension: 0, laborExtension: 0, totalCost: 0,
        notes: "Auto-inferred: 1 butt weld per cap", fromDatabase: false,
        weldAssumption: "1 butt weld per cap (auto-inferred)",
      })));
    } else if (cat === "coupling" || desc.includes("coupling")) {
      // Threaded couplings: 0 welds. Socket weld couplings: 2 SW.
      const isThreadedConn = desc.includes("threaded") || desc.includes("npt") || desc.includes("screw");
      if (!isThreadedConn) {
        welds.push(computeEstimateItem(withInheritedMeta(item, {
          id: randomUUID(), lineNumber: 0, category: "weld" as any,
          description: `SW for ${size} COUPLING (auto-inferred)`,
          size, quantity: qty * 2, unit: "EA",
          materialUnitCost: 0, laborUnitCost: 0, laborHoursPerUnit: 0,
          materialExtension: 0, laborExtension: 0, totalCost: 0,
          notes: "Auto-inferred: 2 socket welds per coupling", fromDatabase: false,
          weldAssumption: "2 socket welds per coupling (auto-inferred)",
        })));
      }
    } else if (cat === "flange" || desc.includes("flange")) {
      // Weld-neck (WN) flanges get butt welds, slip-on (SO) flanges get fillet welds
      const isWeldNeck = desc.includes("weld neck") || desc.includes("wn ") || desc.includes(",wn,") || /\bwn\b/i.test(desc);
      welds.push(computeEstimateItem(withInheritedMeta(item, {
        id: randomUUID(), lineNumber: 0, category: "weld" as any,
        description: `${isWeldNeck ? "BW" : "SO weld"} for ${size} FLANGE (auto-inferred)`,
        size, quantity: qty * 1, unit: "EA",
        materialUnitCost: 0, laborUnitCost: 0, laborHoursPerUnit: 0,
        materialExtension: 0, laborExtension: 0, totalCost: 0,
        notes: `Auto-inferred: 1 ${isWeldNeck ? "butt weld" : "slip-on weld"} per flange`, fromDatabase: false,
        weldAssumption: `1 ${isWeldNeck ? "butt weld (WN)" : "slip-on weld (SO)"} per flange (auto-inferred)`,
      })));
      welds.push(computeEstimateItem(withInheritedMeta(item, {
        id: randomUUID(), lineNumber: 0, category: "bolt" as any,
        description: `Bolt-up for ${size} FLANGE (auto-inferred)`,
        size, quantity: qty * 1, unit: "EA",
        materialUnitCost: 0, laborUnitCost: 0, laborHoursPerUnit: 0,
        materialExtension: 0, laborExtension: 0, totalCost: 0,
        notes: "Auto-inferred: 1 bolt-up per flange", fromDatabase: false,
        weldAssumption: "1 bolt-up per flange (auto-inferred)",
      })));
    } else if (cat === "valve" || desc.includes("valve")) {
      // ONLY socket-weld valves generate welds. All other valve types (flanged,
      // threaded, butt-weld end, butterfly, etc.) connect via bolt-up or threads,
      // not welds.
      const isSocketWeldValve = desc.includes("socket") || desc.includes(" sw ") || desc.includes(",sw,") || /\bsw\b/i.test(desc);
      if (isSocketWeldValve) {
        welds.push(computeEstimateItem(withInheritedMeta(item, {
          id: randomUUID(), lineNumber: 0, category: "weld" as any,
          description: `SW for ${size} VALVE (auto-inferred)`,
          size, quantity: qty * 2, unit: "EA",
          materialUnitCost: 0, laborUnitCost: 0, laborHoursPerUnit: 0,
          materialExtension: 0, laborExtension: 0, totalCost: 0,
          notes: "Auto-inferred: 2 socket welds per SW valve", fromDatabase: false,
          weldAssumption: "2 socket welds per socket-weld valve (auto-inferred)",
        })));
      }
      // Flanged, threaded, butterfly, butt-weld end valves: no welds inferred.
      // Their connections come from the flanges/joints around them.
    } else if (desc.includes("sockolet")) {
      welds.push(computeEstimateItem(withInheritedMeta(item, {
        id: randomUUID(), lineNumber: 0, category: "weld" as any,
        description: `SW for ${size} SOCKOLET (auto-inferred)`,
        size, quantity: qty * 2, unit: "EA",
        materialUnitCost: 0, laborUnitCost: 0, laborHoursPerUnit: 0,
        materialExtension: 0, laborExtension: 0, totalCost: 0,
        notes: "Auto-inferred: 2 socket welds per sockolet (header bore + branch)", fromDatabase: false,
        weldAssumption: "2 socket welds per sockolet (auto-inferred)",
      })));
    } else if (desc.includes("weldolet")) {
      welds.push(computeEstimateItem(withInheritedMeta(item, {
        id: randomUUID(), lineNumber: 0, category: "weld" as any,
        description: `BW for ${size} WELDOLET (auto-inferred)`,
        size, quantity: qty * 1, unit: "EA",
        materialUnitCost: 0, laborUnitCost: 0, laborHoursPerUnit: 0,
        materialExtension: 0, laborExtension: 0, totalCost: 0,
        notes: "Auto-inferred: 1 butt weld per weldolet to header", fromDatabase: false,
        weldAssumption: "1 butt weld per weldolet (auto-inferred)",
      })));
    } else if (desc.includes("threadolet")) {
      welds.push(computeEstimateItem(withInheritedMeta(item, {
        id: randomUUID(), lineNumber: 0, category: "weld" as any,
        description: `Weld for ${size} THREADOLET (auto-inferred)`,
        size, quantity: qty * 1, unit: "EA",
        materialUnitCost: 0, laborUnitCost: 0, laborHoursPerUnit: 0,
        materialExtension: 0, laborExtension: 0, totalCost: 0,
        notes: "Auto-inferred: 1 weld per threadolet to header", fromDatabase: false,
        weldAssumption: "1 weld per threadolet (auto-inferred)",
      })));
    }
  }
  return welds;
}

// ============================================================
// ROUTES
// ============================================================

export async function registerRoutes(httpServer: Server, app: Express): Promise<Server> {

  // ===== PUBLIC DIAGNOSTIC ENDPOINT (before auth) =====
  app.get("/api/system-check", (_req, res) => {
    const checks: Record<string, { installed: boolean; version?: string; error?: string }> = {};
    for (const tool of ["pdftoppm", "qpdf", "convert", "identify", "tesseract"]) {
      try {
        const ver = execFileSync(tool, tool === "qpdf" ? ["--version"] : ["--version"], {
          timeout: 5000, encoding: "utf-8",
        }).toString().trim().split("\n")[0];
        checks[tool] = { installed: true, version: ver };
      } catch (err: any) {
        // Some tools return version on stderr or exit non-zero
        const stderr = err.stderr?.toString().trim().split("\n")[0] || "";
        if (stderr && (stderr.includes("version") || stderr.includes("Version") || stderr.includes("ImageMagick") || stderr.includes("poppler") || stderr.includes("tesseract"))) {
          checks[tool] = { installed: true, version: stderr };
        } else if (err.status !== undefined && err.stdout) {
          checks[tool] = { installed: true, version: err.stdout.toString().trim().split("\n")[0] };
        } else {
          checks[tool] = { installed: false, error: err.message?.substring(0, 100) || "not found" };
        }
      }
    }
    const allInstalled = Object.values(checks).every(c => c.installed);
    res.json({
      status: allInstalled ? "ok" : "missing_dependencies",
      runtime: process.env.RENDER ? "render" : "local",
      nodeVersion: process.version,
      tools: checks,
    });
  });

  // Apply auth middleware
  app.use(authMiddleware);

  // ===== AUTH ROUTES =====

  app.post("/api/login", (req, res) => {
    const ip = req.ip || req.socket.remoteAddress || "unknown";
    const now = Date.now();
    const attempt = loginAttempts.get(ip);

    // Check if blocked
    if (attempt) {
      if (now - attempt.firstAttempt > LOGIN_WINDOW_MS) {
        // Window expired, reset
        loginAttempts.delete(ip);
      } else if (attempt.count >= LOGIN_MAX_ATTEMPTS) {
        const remainingSec = Math.ceil((LOGIN_WINDOW_MS - (now - attempt.firstAttempt)) / 1000);
        return res.status(429).json({ message: `Too many login attempts. Try again in ${remainingSec} seconds.` });
      }
    }

    const { username, password } = req.body || {};
    if (!username || !password) {
      return res.status(400).json({ message: "Username and password are required" });
    }
    const result = storage.login(username, password);
    if (!result) {
      // Track failed attempt
      const current = loginAttempts.get(ip);
      if (current) {
        current.count++;
      } else {
        loginAttempts.set(ip, { count: 1, firstAttempt: now, blocked: false });
      }
      return res.status(401).json({ message: "Invalid credentials" });
    }
    // Success — clear attempts for this IP
    loginAttempts.delete(ip);
    res.json(result);
  });

  app.get("/api/auth/validate", (req, res) => {
    const authHeader = req.headers.authorization;
    if (!authHeader || !authHeader.startsWith("Bearer ")) {
      return res.status(401).json({ valid: false });
    }
    const session = storage.validateToken(authHeader.slice(7));
    if (!session) {
      return res.status(401).json({ valid: false });
    }
    res.json({ valid: true, username: session.username });
  });

  app.post("/api/auth/change-password", (req, res) => {
    const authHeader = req.headers.authorization;
    if (!authHeader || !authHeader.startsWith("Bearer ")) {
      return res.status(401).json({ error: "Not authenticated" });
    }
    const session = storage.validateToken(authHeader.slice(7));
    if (!session) return res.status(401).json({ error: "Invalid session" });
    const { currentPassword, newPassword } = req.body || {};
    if (!currentPassword || !newPassword) {
      return res.status(400).json({ error: "currentPassword and newPassword are required" });
    }
    const result = storage.changePassword(session.username, currentPassword, newPassword);
    if (!result.success) return res.status(400).json({ error: result.error });
    res.json({ success: true });
  });

  app.post("/api/logout", (req, res) => {
    const authHeader = req.headers.authorization;
    if (authHeader && authHeader.startsWith("Bearer ")) {
      storage.logout(authHeader.slice(7));
    }
    res.json({ message: "Logged out" });
  });

  // ===== TAKEOFF PROJECTS =====

  app.get("/api/takeoff/projects", (req, res) => {
    const { discipline } = req.query as { discipline?: string };
    const projects = storage.getTakeoffProjectsLite(discipline);
    res.json(projects);
  });

  app.get("/api/takeoff/projects/:id", async (req, res) => {
    const project = await storage.getTakeoffProject(req.params.id);
    if (!project) return res.status(404).json({ error: "Project not found" });
    res.json(project);
  });

  app.patch("/api/takeoff/projects/:id/archive", async (req, res) => {
    const { archived } = req.body as { archived: boolean };
    const updated = await storage.archiveTakeoffProject(req.params.id, archived);
    if (!updated) return res.status(404).json({ error: "Project not found" });
    res.json({ archived });
  });

  app.delete("/api/takeoff/projects/:id", async (req, res) => {
    const deleted = await storage.deleteTakeoffProject(req.params.id);
    if (!deleted) return res.status(404).json({ error: "Project not found" });
    res.json({ deleted: true });
  });

  // Polling endpoint for progress updates
  app.get("/api/progress/:jobId", (req, res) => {
    const { jobId } = req.params;
    const progress = storage.getJobProgress(jobId);
    if (!progress) return res.status(404).json({ error: "Job not found" });
    res.json(progress);
  });

  // Upload takeoff PDF — accepts discipline parameter
  app.post("/api/takeoff/upload", (req, res) => {
    req.setTimeout(3600000);
    if (res.setTimeout) (res as any).setTimeout(3600000);

    upload.single("file")(req, res, async (multerErr: any) => {
      try {
        if (multerErr) {
          if (multerErr.code === "LIMIT_FILE_SIZE") {
            return res.status(400).json({ error: `File too large. Maximum size is ${MAX_FILE_SIZE / (1024 * 1024)}MB.` });
          }
          return res.status(400).json({ error: multerErr.message || "File upload error" });
        }
        if (!req.file) return res.status(400).json({ error: "No file received." });

        const discipline = (req.body.discipline || "mechanical").toLowerCase();
        if (!["mechanical", "structural", "civil"].includes(discipline)) {
          return res.status(400).json({ error: "Invalid discipline. Must be mechanical, structural, or civil." });
        }
        const verifyExtraction = req.body.verifyExtraction !== "false" && req.body.verifyExtraction !== false;
        const hasRevisions = req.body.hasRevisions === "true" || req.body.hasRevisions === true;
        const dualModel = req.body.dualModel === "true" || req.body.dualModel === true;

        const fileName = req.file.originalname;
        const pdfPath = req.file.path;
        const jobId = randomUUID();

        console.log(`\n======= Received: ${fileName} [${discipline}] jobId=${jobId} verify=${verifyExtraction} revisions=${hasRevisions} dualModel=${dualModel} =======`);

        const pageCount = getPdfPageCount(pdfPath);
        const totalChunks = Math.ceil(pageCount / CHUNK_SIZE);

        storage.setJobProgress(jobId, {
          jobId,
          status: "uploading",
          phase: "uploading",
          chunk: 0,
          totalChunks,
          pagesProcessed: 0,
          totalPages: pageCount,
          itemsFound: 0,
        });

        const isLargePackage = pageCount > CHUNK_SIZE;
        res.status(202).json({ 
          jobId, 
          pageCount, 
          totalChunks,
          isLargePackage,
          message: isLargePackage 
            ? `Large package detected (${pageCount} pages). Splitting into ${totalChunks} sections of ~${CHUNK_SIZE} pages each. This will take a while — progress will update as each section completes.`
            : undefined,
        });

        processUploadedPdf(jobId, fileName, pdfPath, pageCount, totalChunks, discipline, verifyExtraction, hasRevisions, dualModel).catch((err) => {
          console.error("Background processing error:", err);
          storage.setJobProgress(jobId, {
            jobId,
            status: "error",
            phase: "error",
            chunk: 0,
            totalChunks: 0,
            pagesProcessed: 0,
            totalPages: pageCount,
            itemsFound: 0,
            error: err.message || "Processing failed.",
          });
        });
      } catch (err: any) {
        console.error("Upload error:", err);
        res.status(422).json({ error: err.message || "Failed to process PDF." });
      }
    });
  });

  // Get completed project by jobId
  app.get("/api/takeoff/by-job/:jobId", async (req, res) => {
    const progress = storage.getJobProgress(req.params.jobId);
    if (!progress) return res.status(404).json({ error: "Job not found" });
    if (progress.status !== "done" || !progress.projectId) {
      return res.status(202).json({ status: progress.status, message: "Still processing" });
    }
    const project = await storage.getTakeoffProject(progress.projectId);
    if (!project) return res.status(404).json({ error: "Project not found" });
    storage.deleteJobProgress(req.params.jobId);
    res.json(project);
  });

  // ---- Corrections feedback endpoints ----

  app.post("/api/takeoff/corrections", (req, res) => {
    const { takeoffProjectId, itemId, fieldName, originalValue, correctedValue, correctedBy } = req.body || {};
    if (!takeoffProjectId || !itemId || !fieldName) {
      return res.status(400).json({ error: "takeoffProjectId, itemId, and fieldName are required" });
    }
    const correction = storage.addCorrection({ takeoffProjectId, itemId, fieldName, originalValue, correctedValue, correctedBy });
    res.json(correction);
  });

  app.get("/api/takeoff/corrections/:projectId", (req, res) => {
    const corrections = storage.getCorrectionsByProject(req.params.projectId);
    res.json(corrections);
  });

  // Import takeoff BOM into an estimate
  app.post("/api/takeoff/import-to-estimate", async (req, res) => {
    const { takeoffProjectId } = req.body || {};
    if (!takeoffProjectId || typeof takeoffProjectId !== "string") {
      return res.status(400).json({ error: "takeoffProjectId is required and must be a string" });
    }
    const takeoffProject = await storage.getTakeoffProject(takeoffProjectId);
    if (!takeoffProject) return res.status(404).json({ error: "Takeoff project not found" });

    const estimateItems = (takeoffProject.items || []).map((item, idx) => {
      const category = mapTakeoffToEstimateCategory(takeoffProject.discipline, item.category) as any;
      const estimateItem = {
        id: randomUUID(),
        lineNumber: idx + 1,
        category,
        description: item.description,
        size: item.size,
        quantity: item.quantity,
        unit: item.unit,
        materialUnitCost: 0,
        laborUnitCost: 0,
        laborHoursPerUnit: 0,
        materialExtension: 0,
        laborExtension: 0,
        totalCost: 0,
        notes: item.notes || "",
        fromDatabase: false,
        revisionClouded: item.revisionClouded || false,
      };

      // Try to match against cost database
      const matches = storage.matchCostEntries([{ description: item.description, size: item.size }]);
      const key = `${item.description.toLowerCase().trim()}|${item.size.toLowerCase().trim()}`;
      const match = matches[key];
      if (match) {
        Object.assign(estimateItem, {
          materialUnitCost: match.materialUnitCost,
          laborUnitCost: match.laborUnitCost,
          laborHoursPerUnit: match.laborHoursPerUnit,
          fromDatabase: true,
          materialCostSource: "database",
        });
      }

      // If no cost DB match, try purchase history for actual prices
      if (!match || match.materialUnitCost === 0) {
        const purchaseMatch = storage.getLatestCostForItem(item.description, item.size);
        if (purchaseMatch && purchaseMatch.unitCost > 0) {
          estimateItem.materialUnitCost = purchaseMatch.unitCost;
          (estimateItem as any).materialCostSource = "purchase_history";
          estimateItem.fromDatabase = true;
        }
      }

      // If still no cost source, set to empty
      if (!(estimateItem as any).materialCostSource) {
        (estimateItem as any).materialCostSource = "";
      }

      return computeEstimateItem(estimateItem);
    });

    // Feature 2: Weld Count Inference from Fittings
    const inferredWelds = inferWeldsFromFittings(estimateItems);
    const allEstimateItems = [...estimateItems, ...inferredWelds].map((item, idx) => ({ ...item, lineNumber: idx + 1 }));

    const estimateProject = storage.createEstimateProject({
      name: `${takeoffProject.name} — Estimate`,
      sourceTakeoffId: takeoffProjectId,
      items: allEstimateItems,
    });

    res.status(201).json(estimateProject);
  });

  // Create an estimate from a takeoff that only includes items inside a
  // revision cloud. Useful for quoting the delta when a drawing comes back
  // with rev N changes — the resulting estimate is its own project that can
  // be priced, exported, and audit-trailed independently.
  app.post("/api/takeoff/import-revision-to-estimate", async (req, res) => {
    const { takeoffProjectId, revisionLabel, inferWelds } = req.body || {};
    if (!takeoffProjectId || typeof takeoffProjectId !== "string") {
      return res.status(400).json({ error: "takeoffProjectId is required and must be a string" });
    }
    // Default: infer welds on revision estimates too. Justin's comparison
    // showed he counts pipe-joint and fitting welds on revisions, so the
    // earlier choice to suppress them was undercounting labor by ~half.
    const shouldInferWelds = inferWelds !== false;
    const takeoffProject = await storage.getTakeoffProject(takeoffProjectId);
    if (!takeoffProject) return res.status(404).json({ error: "Takeoff project not found" });

    const cloudedItems = (takeoffProject.items || []).filter((it: any) => !!it.revisionClouded);
    if (cloudedItems.length === 0) {
      return res.status(400).json({ error: "No revision-clouded items found on this takeoff. Mark items as clouded in the takeoff first." });
    }

    const estimateItems = cloudedItems.map((item, idx) => {
      const category = mapTakeoffToEstimateCategory(takeoffProject.discipline, item.category) as any;
      const estimateItem: any = {
        id: randomUUID(),
        lineNumber: idx + 1,
        category,
        description: item.description,
        size: item.size,
        quantity: item.quantity,
        unit: item.unit,
        materialUnitCost: 0,
        laborUnitCost: 0,
        laborHoursPerUnit: 0,
        materialExtension: 0,
        laborExtension: 0,
        totalCost: 0,
        notes: item.notes || "",
        fromDatabase: false,
        revisionClouded: true,
      };
      const matches = storage.matchCostEntries([{ description: item.description, size: item.size }]);
      const key = `${item.description.toLowerCase().trim()}|${item.size.toLowerCase().trim()}`;
      const match = matches[key];
      if (match) {
        Object.assign(estimateItem, {
          materialUnitCost: match.materialUnitCost,
          laborUnitCost: match.laborUnitCost,
          laborHoursPerUnit: match.laborHoursPerUnit,
          fromDatabase: true,
          materialCostSource: "database",
        });
      }
      if (!match || match.materialUnitCost === 0) {
        const purchaseMatch = storage.getLatestCostForItem(item.description, item.size);
        if (purchaseMatch && purchaseMatch.unitCost > 0) {
          estimateItem.materialUnitCost = purchaseMatch.unitCost;
          estimateItem.materialCostSource = "purchase_history";
          estimateItem.fromDatabase = true;
        }
      }
      if (!estimateItem.materialCostSource) estimateItem.materialCostSource = "";
      return computeEstimateItem(estimateItem);
    });

    // Infer welds (fitting welds + pipe-joint welds for 40' runs) so the
    // revision estimate matches how Justin actually prices change orders.
    // The caller can disable this by passing { inferWelds: false } in the body.
    let finalItems = estimateItems;
    if (shouldInferWelds) {
      const inferredWelds = inferWeldsFromFittings(estimateItems);
      finalItems = [...estimateItems, ...inferredWelds].map((item, idx) => ({ ...item, lineNumber: idx + 1 }));
    }

    const revLabel = (revisionLabel && typeof revisionLabel === "string") ? revisionLabel.trim() : (takeoffProject.revision || "Revision");
    const estimateProject = storage.createEstimateProject({
      name: `${takeoffProject.name} — ${revLabel} (clouded only)`,
      sourceTakeoffId: takeoffProjectId,
      items: finalItems,
    });
    res.status(201).json({ ...estimateProject, revisionItemCount: estimateItems.length, inferredWeldCount: finalItems.length - estimateItems.length });
  });

  // ===== ESTIMATE PROJECTS =====

  app.get("/api/estimates", (_req, res) => {
    res.json(storage.getEstimateProjectsLite());
  });

  app.get("/api/estimates/:id", (req, res) => {
    const project = storage.getEstimateProject(req.params.id);
    if (!project) return res.status(404).json({ message: "Estimate not found" });
    res.json(project);
  });

  app.post("/api/estimates", (req, res) => {
    const parsed = insertEstimateProjectSchema.safeParse(req.body);
    if (!parsed.success) {
      return res.status(400).json({ message: "Validation failed", errors: parsed.error.flatten().fieldErrors });
    }
    const project = storage.createEstimateProject(parsed.data);
    res.status(201).json(project);
  });

  app.patch("/api/estimates/:id", (req, res) => {
    const parsed = patchEstimateSchema.safeParse(req.body);
    if (!parsed.success) {
      return res.status(400).json({ message: "Validation failed", errors: parsed.error.flatten().fieldErrors });
    }
    const updated = storage.updateEstimateProject(req.params.id, parsed.data);
    if (!updated) return res.status(404).json({ message: "Estimate not found" });
    res.json(updated);
  });

  app.put("/api/estimates/:id/items", (req, res) => {
    const project = storage.getEstimateProject(req.params.id);
    if (!project) return res.status(404).json({ message: "Estimate not found" });
    const itemsRaw = req.body.items;
    if (!Array.isArray(itemsRaw)) return res.status(400).json({ message: "items must be an array" });
    const items = itemsRaw.map(computeEstimateItem);
    const updated = storage.updateEstimateProject(req.params.id, { items });
    res.json(updated);
  });

  app.delete("/api/estimates/:id", (req, res) => {
    const deleted = storage.deleteEstimateProject(req.params.id);
    if (!deleted) return res.status(404).json({ message: "Estimate not found" });
    res.status(204).send();
  });

  // Apply database costs to an estimate
  app.post("/api/estimates/:id/apply-database", (req, res) => {
    const project = storage.getEstimateProject(req.params.id);
    if (!project) return res.status(404).json({ message: "Estimate not found" });

    const matchInputs = (project.items || []).map(i => ({ description: i.description, size: i.size }));
    const matches = storage.matchCostEntries(matchInputs);

    const updatedItems = (project.items || []).map(item => {
      const key = `${item.description.toLowerCase().trim()}|${item.size.toLowerCase().trim()}`;
      const match = matches[key];
      if (match) {
        return computeEstimateItem({
          ...item,
          materialUnitCost: match.materialUnitCost,
          laborUnitCost: match.laborUnitCost,
          laborHoursPerUnit: match.laborHoursPerUnit,
          fromDatabase: true,
          materialCostSource: "database" as const,
        });
      }
      return item;
    });

    const updated = storage.updateEstimateProject(req.params.id, { items: updatedItems });
    res.json(updated);
  });

  // Generate RFQ (Request for Quote) email for material procurement
  app.post("/api/estimates/:id/generate-rfq", (req, res) => {
    const project = storage.getEstimateProject(req.params.id);
    if (!project) return res.status(404).json({ message: "Estimate not found" });

    const selectedItemIds: string[] | undefined = req.body.selectedItemIds;
    // RFQ honors the per-row 'includeInBom' flag (default true). When the user
    // explicitly selects items, that selection takes precedence over the flag.
    const items = selectedItemIds
      ? (project.items || []).filter(i => selectedItemIds.includes(i.id))
      : (project.items || []).filter(i => (i as any).includeInBom !== false);

    // Build materials table with spec details
    const materialsTable = items.map(i => ({
      description: i.description,
      size: i.size,
      quantity: i.quantity,
      unit: i.unit,
      category: i.category,
      spec: (i as any).spec || "",
      material: (i as any).itemMaterial || "",
      schedule: (i as any).itemSchedule || "",
      rating: (i as any).rating || "",
      materialCostSource: (i as any).materialCostSource || "",
    }));

    // Find unique cost sources from items
    const costSources = new Set<string>();
    items.forEach(i => {
      if ((i as any).materialCostSource && (i as any).materialCostSource !== "" && (i as any).materialCostSource !== "manual") {
        costSources.add((i as any).materialCostSource);
      }
    });

    // Detect material categories in the estimate
    const categories = new Set(items.map(i => i.category));

    // Build the email text
    const projectName = project.name || "Unnamed Project";
    const projectNumber = project.projectNumber || "N/A";
    const client = project.client || "N/A";
    const location = project.location || "N/A";

    let tableText = "| # | Description | Size | Qty | Unit | Material | Schedule | Rating |\n|---|-------------|------|-----|------|----------|----------|--------|\n";
    materialsTable.forEach((m, idx) => {
      tableText += `| ${idx + 1} | ${m.description} | ${m.size} | ${m.quantity} | ${m.unit} | ${m.material || "-"} | ${m.schedule || "-"} | ${m.rating || "-"} |\n`;
    });

    const emailText = `Subject: Request for Quotation - ${projectName}

Dear Supplier,

We are requesting pricing and lead times for the materials listed below for the following project:

  Project: ${projectName}
  Project Number: ${projectNumber}
  Client: ${client}
  Location: ${location}

MATERIALS REQUIRED:
${tableText}
Total line items: ${materialsTable.length}

Please provide:
1. Unit pricing for each item listed above
2. Lead times for delivery
3. Any minimum order quantities
4. Freight/shipping costs to project location${location !== "N/A" ? ` (${location})` : ""}
5. Payment terms
6. Quotation validity period

Please submit your quotation at your earliest convenience. If you have any questions regarding specifications or quantities, do not hesitate to contact us.

Thank you for your prompt attention to this request.

Best regards,
Picou Group Contractors`;

    // Build supplier suggestions
    const supplierSuggestions: { type: string; suggestion: string; link?: string }[] = [];

    // From cost database sources
    if (costSources.size > 0) {
      supplierSuggestions.push({
        type: "database",
        suggestion: `Items with pricing from cost database: ${Array.from(costSources).join(", ")}`,
      });
    }

    // Location-based
    if (location && location !== "N/A") {
      const categoryTerms = Array.from(categories).slice(0, 3).join("+");
      supplierSuggestions.push({
        type: "location",
        suggestion: `Search for piping/industrial suppliers near ${location}`,
        link: `https://www.google.com/maps/search/industrial+pipe+${categoryTerms}+supplier+near+${encodeURIComponent(location)}`,
      });
    }

    // Category-based suggestions
    const categoryMap: Record<string, string> = {
      pipe: "Pipe distributors (carbon steel, stainless, alloy)",
      valve: "Valve suppliers (gate, globe, check, ball, butterfly)",
      flange: "Flange & fitting suppliers",
      bolt: "Fastener & bolt suppliers (stud bolts, machine bolts)",
      gasket: "Gasket suppliers (spiral wound, ring joint, sheet)",
      elbow: "Pipe fitting suppliers (elbows, tees, reducers)",
      tee: "Pipe fitting suppliers",
      reducer: "Pipe fitting suppliers",
      support: "Pipe support & hanger suppliers",
      steel: "Structural steel suppliers",
    };
    const seen = new Set<string>();
    for (const cat of categories) {
      const mapped = categoryMap[cat];
      if (mapped && !seen.has(mapped)) {
        seen.add(mapped);
        supplierSuggestions.push({ type: "category", suggestion: mapped });
      }
    }

    res.json({
      emailText,
      materialsTable,
      supplierSuggestions,
      projectInfo: { name: projectName, projectNumber, client, location },
    });
  });

  // Auto-calculate labor hours using Bill's EI or Justin's factor method
  app.post("/api/estimates/:id/auto-calculate", (req, res) => {
    const project = storage.getEstimateProject(req.params.id);
    if (!project) return res.status(404).json({ message: "Estimate not found" });

    // Feature 3: Save version snapshot before calculating
    try {
      storage.saveEstimateVersion(req.params.id, `Before auto-calculate (${req.body?.method || "justin"})`);
    } catch (e) { console.warn("Version save failed:", e); }

    const parsed = autoCalculateBodySchema.safeParse(req.body);
    if (!parsed.success) {
      return res.status(400).json({ message: "Validation failed", errors: parsed.error.flatten().fieldErrors });
    }
    const {
      method,
      customMethodId,
      laborRate,
      overtimeRate,
      doubleTimeRate,
      perDiem,
      overtimePercent,
      doubleTimePercent,
      material,
      schedule,
      installType,
      pipeLocation,
      elevation,
      alloyGroup,
      rackFactor,
      fittingWeldMode,
    } = parsed.data;

    // Persist this onto the project so subsequent reads (diagnose, compare,
    // exports) reflect the same mode without the client having to resend it.
    try { storage.updateEstimateProject(req.params.id, { fittingWeldMode } as any); } catch {}

    let estimatorData: any;
    try {
      estimatorData = getEstimatorData();
    } catch (err: any) {
      return res.status(500).json({ message: `Failed to load estimator data: ${err.message}` });
    }

    // If a custom method id is provided, resolve it and layer its overrides
    // onto the base method's data block. The base method (bill/justin/industry)
    // still drives the calculator function; the custom profile only edits factors.
    let customMethod: any = null;
    if (customMethodId) {
      customMethod = storage.getCustomMethod(customMethodId);
      if (!customMethod) {
        return res.status(404).json({ message: `Custom method '${customMethodId}' not found` });
      }
      // Apply the custom overrides to the appropriate base data block.
      const base = customMethod.baseMethod;
      const overridden = applyCustomOverrides(estimatorData[base], customMethod.overrides || {});
      // Replace the data block at that base key so downstream calculators pick it up.
      estimatorData = { ...estimatorData, [base]: overridden };
    }

    // Pre-flight: verify the requested method's data block actually exists in
    // the loaded estimator-data.json. Return a clean error rather than letting
    // the per-item calculator throw an opaque 'undefined' error.
    if (method === "industry" && !estimatorData.industry) {
      return res.status(500).json({
        message: "Industry method data is unavailable on this server. The deployed build is missing 'industry' in estimator-data.json. Redeploy the latest main, or pick Bill's or Justin's method.",
      });
    }
    if (method === "justin" && !estimatorData.justin) {
      return res.status(500).json({ message: "Justin method data missing on the server." });
    }
    if (method === "bill" && !estimatorData.bill) {
      return res.status(500).json({ message: "Bill method data missing on the server." });
    }

    // Compute blended rate from ST/OT/DT percentages
    const stPercent = Math.max(0, (100 - overtimePercent - doubleTimePercent) / 100);
    const otPercent = overtimePercent / 100;
    const dtPercent = doubleTimePercent / 100;
    const blendedRate = (laborRate * stPercent) + (overtimeRate * otPercent) + (doubleTimeRate * dtPercent);
    const perDiemPerHour = perDiem / 10; // assuming 10-hr workdays
    const effectiveRate = blendedRate + perDiemPerHour;

    const updatedItems = (project.items || []).map(item => {
      // Excluded rows are pass-through — keep whatever labor/cost they had,
      // since they don't contribute to totals anyway and re-running the
      // calculator on them would overwrite values the user might want to
      // preserve for if they re-include the row later.
      if ((item as any).includeInEstimate === false) {
        return item;
      }
      // Use per-item overrides if present, otherwise fall back to global settings
      // Auto-detect SS from description if itemMaterial not set (Calibration Item 2)
      let detectedMat: "CS" | "SS" = material;
      if ((item as any).itemMaterial) {
        detectedMat = (item as any).itemMaterial as "CS" | "SS";
      } else {
        const desc = (item.description || "").toUpperCase();
        if (/\b(SS|STAINLESS|TP304|TP316|304L?|316L?|A312|A182|A403)\b/.test(desc)) {
          detectedMat = "SS";
        }
      }
      const itemMat = detectedMat;
      const itemSched = (item as any).itemSchedule || schedule;
      const itemElev = (item as any).itemElevation || elevation;
      const itemPipeLoc = (item as any).itemPipeLocation || pipeLocation;
      const itemAlloy = (item as any).itemAlloyGroup || alloyGroup;
      // Per-line work type: use item's workType if set, otherwise fall back to global installType
      const lineWorkType = (item as any).workType || installType;

      let laborHoursPerUnit = 0;
      let materialUnitCostAdjust = 0;
      let calcBasis = "";
      let sizeMatchExact = true;
      let materialCostSource = (item as any).materialCostSource || "";

      if (method === "bill") {
        const result = calculateBillLaborHours(item, itemMat, itemSched, estimatorData.bill, itemPipeLoc, itemElev, itemAlloy, fittingWeldMode);
        laborHoursPerUnit = result.laborHoursPerUnit;
        materialUnitCostAdjust = result.materialUnitCostAdjust;
        calcBasis = result.calcBasis;
        sizeMatchExact = result.sizeMatchExact;
        if (result.materialCostSource) materialCostSource = result.materialCostSource;
        // Bill's EI method doesn't have rack-specific rates, so apply rackFactor for rack work
        if (lineWorkType === "rack" && rackFactor > 1) {
          laborHoursPerUnit = laborHoursPerUnit * rackFactor;
          calcBasis += ` × ${rackFactor.toFixed(2)} (rack factor) = ${laborHoursPerUnit.toFixed(4)} MH`;
        }
      } else if (method === "industry") {
        // Industry method (Page's Estimator's Piping Man-Hour Manual)
        // Reuses Justin's calculator shape against the Industry data block.
        // Contingency: project's contingencyOverride takes priority over the
        // data-file default (Page guidance: 10%). User-entered as a percent.
        const iResult = calculateIndustryLaborHours(item, lineWorkType as "standard" | "rack", itemMat, itemSched, estimatorData.industry, fittingWeldMode);
        const baseMH = iResult.mh;
        sizeMatchExact = iResult.sizeMatchExact;
        const dataDefault = estimatorData.industry?.cost_params?.contingency_factor ?? 0.10;
        const override = (project as any).contingencyOverride;
        const contingencyFactor = (typeof override === "number" && !Number.isNaN(override)) ? (override / 100) : dataDefault;
        const contingencyMult = 1 + contingencyFactor;
        laborHoursPerUnit = baseMH * contingencyMult;
        const overrideTag = (typeof override === "number" && !Number.isNaN(override)) ? " override" : "";
        calcBasis = `${iResult.calcBasis} × ${contingencyMult.toFixed(2)} (${(contingencyFactor * 100).toFixed(1)}% contingency${overrideTag}) = ${laborHoursPerUnit.toFixed(4)} MH`;
      } else {
        const jResult = calculateJustinLaborHours(item, lineWorkType as "standard" | "rack", itemMat, itemSched, estimatorData.justin, fittingWeldMode);
        const baseMH = jResult.mh;
        sizeMatchExact = jResult.sizeMatchExact;
        // Contingency: project's contingencyOverride takes priority over Justin's
        // data-file default (15%). User-entered as a percent.
        const dataDefault = estimatorData.justin?.cost_params?.contingency_factor || 0.15;
        const override = (project as any).contingencyOverride;
        const contingencyFactor = (typeof override === "number" && !Number.isNaN(override)) ? (override / 100) : dataDefault;
        const contingencyMult = 1 + contingencyFactor;
        laborHoursPerUnit = baseMH * contingencyMult;
        const overrideTag = (typeof override === "number" && !Number.isNaN(override)) ? " override" : "";
        calcBasis = `${jResult.calcBasis} × ${contingencyMult.toFixed(2)} (${(contingencyFactor * 100).toFixed(1)}% contingency${overrideTag}) = ${laborHoursPerUnit.toFixed(4)} MH`;
        // Note: rack vs standard is already handled inside calculateJustinLaborHours
        // by selecting rack_mh_per_lf vs standard columns — no secondary rack factor needed
      }

      const laborUnitCost = laborHoursPerUnit * effectiveRate;
      const updatedItem: any = {
        ...item,
        laborHoursPerUnit,
        laborUnitCost,
        calculationBasis: calcBasis,
        sizeMatchExact,
      };
      // If Bill's method provided a material cost adjust and item has no material cost set, apply it
      if (method === "bill" && materialUnitCostAdjust > 0 && (item.materialUnitCost === 0 || item.materialUnitCost === undefined)) {
        updatedItem.materialUnitCost = materialUnitCostAdjust;
        updatedItem.materialCostSource = "allowance";
      }
      // Preserve or set materialCostSource
      if (materialCostSource) updatedItem.materialCostSource = materialCostSource;
      return computeEstimateItem(updatedItem);
    });

    // Auto-add a supervision line item for the factor-based methods
    // (Justin and Industry). Bill's method already builds supervision into
    // its EI factors so it doesn't need a separate row.
    if (method === "justin" || method === "industry") {
      const hasSupervision = updatedItems.some((i: any) => (i.description || "").toLowerCase().includes("supervision"));
      if (!hasSupervision) {
        const totalHours = updatedItems.reduce((s: number, i: any) => s + (i.quantity || 0) * (i.laborHoursPerUnit || 0), 0);
        const cp = method === "justin" ? estimatorData.justin?.cost_params : estimatorData.industry?.cost_params;
        const supervisionHoursPerWeek = cp?.supervision_hours_per_week || 60;
        const crewSize = 8;
        const hoursPerDay = 10;
        const projectWeeks = Math.max(1, Math.ceil(totalHours / (crewSize * hoursPerDay * 5)));
        const supervisionMH = projectWeeks * supervisionHoursPerWeek;
        const methodLabel = method === "justin" ? "Justin" : "Industry";
        const supervisionItem = computeEstimateItem({
          id: randomUUID(),
          lineNumber: updatedItems.length + 1,
          category: "other" as any,
          description: "Project Supervision",
          size: "",
          quantity: projectWeeks,
          unit: "WK",
          materialUnitCost: 0,
          laborUnitCost: supervisionHoursPerWeek * effectiveRate,
          laborHoursPerUnit: supervisionHoursPerWeek,
          materialExtension: 0,
          laborExtension: 0,
          totalCost: 0,
          notes: `Auto: ${projectWeeks} weeks × ${supervisionHoursPerWeek} hrs/wk (est. from ${totalHours.toFixed(0)} total MH)`,
          fromDatabase: false,
          calculationBasis: `${methodLabel}: Supervision → ${projectWeeks} wk × ${supervisionHoursPerWeek} MH/wk = ${supervisionMH} MH`,
        });
        updatedItems.push(supervisionItem);
      }
    }

    const updated = storage.updateEstimateProject(req.params.id, {
      items: updatedItems,
      laborRate,
      overtimeRate,
      doubleTimeRate,
      perDiem,
      overtimePercent,
      doubleTimePercent,
      estimateMethod: method,
      customMethodId: customMethodId || null,
    } as any);
    res.json(updated);
  });

  // ===== COST DATABASE =====

  app.get("/api/cost-database", (_req, res) => {
    res.json(storage.getCostDatabase());
  });

  app.post("/api/cost-database", (req, res) => {
    const parsed = insertCostDatabaseEntrySchema.safeParse(req.body);
    if (!parsed.success) {
      return res.status(400).json({ message: "Validation failed", errors: parsed.error.flatten().fieldErrors });
    }
    const entry = storage.addCostEntry(parsed.data);
    res.status(201).json(entry);
  });

  app.patch("/api/cost-database/:id", (req, res) => {
    const allowed = insertCostDatabaseEntrySchema.partial().safeParse(req.body);
    if (!allowed.success) {
      return res.status(400).json({ message: "Validation failed", errors: allowed.error.flatten().fieldErrors });
    }
    const updated = storage.updateCostEntry(req.params.id, allowed.data);
    if (!updated) return res.status(404).json({ message: "Entry not found" });
    res.json(updated);
  });

  app.delete("/api/cost-database/:id", (req, res) => {
    const deleted = storage.deleteCostEntry(req.params.id);
    if (!deleted) return res.status(404).json({ message: "Entry not found" });
    res.status(204).send();
  });

  app.post("/api/cost-database/match", (req, res) => {
    const items = req.body.items || [];
    const matches = storage.matchCostEntries(items);

    // Enrich with purchase history where cost DB has no price
    for (const item of items) {
      const key = `${item.description.toLowerCase().trim()}|${(item.size || "").toLowerCase().trim()}`;
      if (!matches[key] || matches[key].materialUnitCost === 0) {
        const purchaseMatch = storage.getLatestCostForItem(item.description, item.size);
        if (purchaseMatch && purchaseMatch.unitCost > 0) {
          if (!matches[key]) {
            matches[key] = {
              id: "", description: item.description, size: item.size,
              category: purchaseMatch.category || "other", unit: purchaseMatch.unit || "EA",
              materialUnitCost: purchaseMatch.unitCost, laborUnitCost: 0, laborHoursPerUnit: 0,
              lastUpdated: purchaseMatch.invoiceDate || new Date().toISOString(),
              materialCostSource: "purchase_history",
            };
          } else {
            matches[key].materialUnitCost = purchaseMatch.unitCost;
            (matches[key] as any).materialCostSource = "purchase_history";
          }
        }
      }
    }

    res.json(matches);
  });

  // ===== ESTIMATOR DATA (Bill / Justin / Industry base factor tables) =====
  // Read-only view of the bundled estimator-data.json. Useful for the UI to
  // render the base factor tables before letting the user clone+edit into a
  // custom method. Each method object includes a `source` field where applicable.
  app.get("/api/estimator-methods", (_req, res) => {
    try {
      const data = getEstimatorData();
      const methods = [
        { key: "bill",     name: "Bill's EI Method",          description: data.bill?.description || "",     source: data.bill?.source || "Picou Group internal" },
        { key: "justin",   name: "Justin's Factor Method",    description: data.justin?.description || "",   source: data.justin?.source || "Picou Group internal" },
        { key: "industry", name: "Industry Standard (Page)",  description: data.industry?.description || "", source: data.industry?.source || "Page's Estimator's Piping Man-Hour Manual" },
      ];
      res.json({ methods, data: { bill: data.bill, justin: data.justin, industry: data.industry } });
    } catch (err: any) {
      res.status(500).json({ message: `Failed to load estimator data: ${err.message}` });
    }
  });

  // ===== CUSTOM ESTIMATOR METHODS (clone-and-edit profiles) =====
  app.get("/api/custom-methods", (_req, res) => {
    res.json(storage.getCustomMethods());
  });

  app.get("/api/custom-methods/:id", (req, res) => {
    const method = storage.getCustomMethod(req.params.id);
    if (!method) return res.status(404).json({ message: "Custom method not found" });
    res.json(method);
  });

  app.post("/api/custom-methods", (req, res) => {
    const parsed = insertCustomEstimatorMethodSchema.safeParse(req.body);
    if (!parsed.success) {
      return res.status(400).json({ message: "Validation failed", errors: parsed.error.flatten().fieldErrors });
    }
    // Enforce unique name (case-insensitive)
    const existing = storage.getCustomMethods().find(m => m.name.toLowerCase() === parsed.data.name.toLowerCase());
    if (existing) {
      return res.status(409).json({ message: `A custom method named '${parsed.data.name}' already exists` });
    }
    const created = storage.createCustomMethod(parsed.data);
    res.status(201).json(created);
  });

  app.patch("/api/custom-methods/:id", (req, res) => {
    const parsed = insertCustomEstimatorMethodSchema.partial().safeParse(req.body);
    if (!parsed.success) {
      return res.status(400).json({ message: "Validation failed", errors: parsed.error.flatten().fieldErrors });
    }
    const updated = storage.updateCustomMethod(req.params.id, parsed.data);
    if (!updated) return res.status(404).json({ message: "Custom method not found" });
    res.json(updated);
  });

  app.delete("/api/custom-methods/:id", (req, res) => {
    const ok = storage.deleteCustomMethod(req.params.id);
    if (!ok) return res.status(404).json({ message: "Custom method not found" });
    res.status(204).end();
  });

  // ===== COMPARE METHODS =====
  // Runs the estimate through Bill, Justin, Industry, and any selected custom
  // methods, returning a summary card + per-line drill-down. The estimate's
  // saved state is NOT modified — this is a pure read-only computation.
  //
  // Body: { customMethodIds?: string[]; laborRate?, overtimeRate?, ... settings }
  // Defaults to comparing bill/justin/industry. If customMethodIds is provided,
  // each one is run with its base method + overrides.
  app.post("/api/estimates/:id/compare-methods", (req, res) => {
    const project = storage.getEstimateProject(req.params.id);
    if (!project) return res.status(404).json({ message: "Estimate not found" });

    const compareBodySchema = z.object({
      customMethodIds: z.array(z.string()).optional().default([]),
      laborRate: z.number().min(0).max(500).default(project.laborRate || 56),
      overtimeRate: z.number().min(0).max(500).default(project.overtimeRate || 79),
      doubleTimeRate: z.number().min(0).max(500).default(project.doubleTimeRate || 100),
      perDiem: z.number().min(0).max(500).default(project.perDiem || 75),
      overtimePercent: z.number().min(0).max(100).default(project.overtimePercent || 15),
      doubleTimePercent: z.number().min(0).max(100).default(project.doubleTimePercent || 2),
      material: z.enum(["CS", "SS"]).default("CS"),
      schedule: z.string().default("40"),
      installType: z.enum(["standard", "rack"]).default("standard"),
      pipeLocation: z.string().default("ground"),
      elevation: z.string().default("ground"),
      alloyGroup: z.string().default("CS"),
      rackFactor: z.number().default(1.3),
    });
    const parsed = compareBodySchema.safeParse(req.body || {});
    if (!parsed.success) {
      return res.status(400).json({ message: "Validation failed", errors: parsed.error.flatten().fieldErrors });
    }
    const settings = parsed.data;

    let estimatorData: any;
    try { estimatorData = getEstimatorData(); } catch (err: any) {
      return res.status(500).json({ message: `Failed to load estimator data: ${err.message}` });
    }

    // Build the list of methods to run: 3 base methods + any requested customs.
    // Methods whose data block is missing from estimator-data.json are skipped
    // gracefully so the comparison still works for methods that ARE available.
    type RunSpec = { key: string; label: string; baseMethod: "bill" | "justin" | "industry"; data: any; customMethodId?: string };
    const runs: RunSpec[] = [];
    const skippedMethods: string[] = [];
    if (estimatorData.bill?.labor_rates)      runs.push({ key: "bill",     label: "Bill",            baseMethod: "bill",     data: estimatorData.bill });
    else                                       skippedMethods.push("Bill");
    if (estimatorData.justin?.labor_factors)  runs.push({ key: "justin",   label: "Justin",          baseMethod: "justin",   data: estimatorData.justin });
    else                                       skippedMethods.push("Justin");
    if (estimatorData.industry?.labor_factors) runs.push({ key: "industry", label: "Industry (Page)", baseMethod: "industry", data: estimatorData.industry });
    else                                       skippedMethods.push("Industry (Page)");
    for (const cmId of settings.customMethodIds) {
      const cm = storage.getCustomMethod(cmId);
      if (!cm) continue;
      const baseBlock = estimatorData[cm.baseMethod];
      if (!baseBlock) { skippedMethods.push(`${cm.name} (base ${cm.baseMethod} missing)`); continue; }
      const overridden = applyCustomOverrides(baseBlock, cm.overrides || {});
      runs.push({ key: `custom:${cm.id}`, label: cm.name, baseMethod: cm.baseMethod, data: overridden, customMethodId: cm.id });
    }
    if (runs.length === 0) {
      return res.status(500).json({ message: `No estimator methods available. Missing: ${skippedMethods.join(", ")}` });
    }

    // Blended labor rate (same logic as auto-calculate).
    const stPercent = Math.max(0, (100 - settings.overtimePercent - settings.doubleTimePercent) / 100);
    const otPercent = settings.overtimePercent / 100;
    const dtPercent = settings.doubleTimePercent / 100;
    const blendedRate = (settings.laborRate * stPercent) + (settings.overtimeRate * otPercent) + (settings.doubleTimeRate * dtPercent);
    const perDiemPerHour = settings.perDiem / 10;
    const effectiveRate = blendedRate + perDiemPerHour;

    // For each item, compute MH/cost under each method. Returns:
    //   summary[]: { key, label, baseMethod, totalMH, totalLaborCost, totalMaterialCost, totalCost, customMethodId? }
    //   lineItems[]: { itemId, description, size, quantity, byMethod: { [key]: { mh, laborCost, calcBasis } } }
    const lineItems: any[] = [];
    const summaryAccum: Record<string, { totalMH: number; totalLaborCost: number; totalMaterialCost: number }> = {};
    for (const r of runs) summaryAccum[r.key] = { totalMH: 0, totalLaborCost: 0, totalMaterialCost: 0 };

    for (const item of (project.items || [])) {
      // Skip rows the user has excluded from this estimate. They stay on the
      // BOM for ordering/takeoff purposes but don't contribute to labor/cost.
      if ((item as any).includeInEstimate === false) continue;
      // Detect material from item or fall back to global
      let detectedMat: "CS" | "SS" = settings.material;
      if ((item as any).itemMaterial) {
        detectedMat = (item as any).itemMaterial as "CS" | "SS";
      } else {
        const desc = (item.description || "").toUpperCase();
        if (/\b(SS|STAINLESS|TP304|TP316|304L?|316L?|A312|A182|A403)\b/.test(desc)) detectedMat = "SS";
      }
      const itemMat = detectedMat;
      const itemSched = (item as any).itemSchedule || settings.schedule;
      const itemElev = (item as any).itemElevation || settings.elevation;
      const itemPipeLoc = (item as any).itemPipeLocation || settings.pipeLocation;
      const itemAlloy = (item as any).itemAlloyGroup || settings.alloyGroup;
      const lineWorkType = (item as any).workType || settings.installType;

      const byMethod: Record<string, any> = {};

      for (const r of runs) {
        let mh = 0;
        let calcBasis = "";
        let materialAdjust = 0;
        if (r.baseMethod === "bill") {
          const result = calculateBillLaborHours(item, itemMat, itemSched, r.data, itemPipeLoc, itemElev, itemAlloy, settings.fittingWeldMode);
          mh = result.laborHoursPerUnit;
          calcBasis = result.calcBasis;
          materialAdjust = result.materialUnitCostAdjust;
          if (lineWorkType === "rack" && settings.rackFactor > 1) {
            mh = mh * settings.rackFactor;
            calcBasis += ` \u00d7 ${settings.rackFactor.toFixed(2)} (rack factor)`;
          }
        } else if (r.baseMethod === "industry") {
          const ir = calculateIndustryLaborHours(item, lineWorkType as "standard" | "rack", itemMat, itemSched, r.data, settings.fittingWeldMode);
          // Project's contingencyOverride takes priority over the data-file default.
          const dataDefault = r.data?.cost_params?.contingency_factor ?? 0.10;
          const override = (project as any).contingencyOverride;
          const contFactor = (typeof override === "number" && !Number.isNaN(override)) ? (override / 100) : dataDefault;
          mh = ir.mh * (1 + contFactor);
          calcBasis = `${ir.calcBasis} \u00d7 ${(1 + contFactor).toFixed(2)} (${(contFactor*100).toFixed(1)}% contingency)`;
        } else {
          const jr = calculateJustinLaborHours(item, lineWorkType as "standard" | "rack", itemMat, itemSched, r.data, settings.fittingWeldMode);
          // Project's contingencyOverride takes priority over Justin's data-file default.
          const dataDefault = r.data?.cost_params?.contingency_factor ?? 0.15;
          const override = (project as any).contingencyOverride;
          const contFactor = (typeof override === "number" && !Number.isNaN(override)) ? (override / 100) : dataDefault;
          mh = jr.mh * (1 + contFactor);
          calcBasis = `${jr.calcBasis} \u00d7 ${(1 + contFactor).toFixed(2)} (${(contFactor*100).toFixed(1)}% contingency)`;
        }
        const laborCost = mh * effectiveRate * (item.quantity || 0);
        const materialUnitCost = (item.materialUnitCost && item.materialUnitCost > 0) ? item.materialUnitCost : materialAdjust;
        const materialCost = materialUnitCost * (item.quantity || 0);
        byMethod[r.key] = {
          mhPerUnit: mh,
          totalMH: mh * (item.quantity || 0),
          laborCost,
          materialCost,
          totalCost: laborCost + materialCost,
          calcBasis,
        };
        summaryAccum[r.key].totalMH += mh * (item.quantity || 0);
        summaryAccum[r.key].totalLaborCost += laborCost;
        summaryAccum[r.key].totalMaterialCost += materialCost;
      }

      lineItems.push({
        itemId: item.id,
        lineNumber: item.lineNumber,
        category: item.category,
        description: item.description,
        size: item.size,
        quantity: item.quantity,
        unit: item.unit,
        byMethod,
      });
    }

    const summary = runs.map(r => ({
      key: r.key,
      label: r.label,
      baseMethod: r.baseMethod,
      customMethodId: r.customMethodId,
      totalMH: summaryAccum[r.key].totalMH,
      totalLaborCost: summaryAccum[r.key].totalLaborCost,
      totalMaterialCost: summaryAccum[r.key].totalMaterialCost,
      totalCost: summaryAccum[r.key].totalLaborCost + summaryAccum[r.key].totalMaterialCost,
    }));

    res.json({
      estimateId: project.id,
      estimateName: project.name,
      itemCount: (project.items || []).length,
      effectiveLaborRate: effectiveRate,
      summary,
      lineItems,
      skippedMethods,
    });
  });

  // ===== EXCEL EXPORT =====

  app.get("/api/estimates/:id/export-bill", async (req, res) => {
    const project = storage.getEstimateProject(req.params.id);
    if (!project) return res.status(404).json({ message: "Estimate not found" });
    try {
      // Strip rows the user has excluded from the estimate so the workbook
      // matches the labor/cost totals shown in the UI.
      const filtered = { ...project, items: (project.items || []).filter((i: any) => i.includeInEstimate !== false) };
      const wb = await generateBillsWorkbook(filtered);
      res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
      res.setHeader("Content-Disposition", `attachment; filename="${project.name.replace(/[^a-zA-Z0-9 _-]/g, "")} - Bills Format.xlsx"`);
      await wb.xlsx.write(res);
      res.end();
    } catch (err: any) {
      console.error("Excel export error:", err);
      res.status(500).json({ message: err.message || "Export failed" });
    }
  });

  app.get("/api/estimates/:id/export-justin", async (req, res) => {
    const project = storage.getEstimateProject(req.params.id);
    if (!project) return res.status(404).json({ message: "Estimate not found" });
    try {
      const filtered = { ...project, items: (project.items || []).filter((i: any) => i.includeInEstimate !== false) };
      const wb = await generateJustinsWorkbook(filtered);
      res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
      res.setHeader("Content-Disposition", `attachment; filename="${project.name.replace(/[^a-zA-Z0-9 _-]/g, "")} - Justins Format.xlsx"`);
      await wb.xlsx.write(res);
      res.end();
    } catch (err: any) {
      console.error("Excel export error:", err);
      res.status(500).json({ message: err.message || "Export failed" });
    }
  });

  // Industry (Page) uses the same calculator output shape as Justin
  // (labor_factors / cost_params), so the Justin workbook layout is
  // a perfect fit — we just relabel the filename. The item rows already
  // carry the Industry man-hour factors because the estimate was run
  // through the Industry calculator.
  app.get("/api/estimates/:id/export-industry", async (req, res) => {
    const project = storage.getEstimateProject(req.params.id);
    if (!project) return res.status(404).json({ message: "Estimate not found" });
    try {
      const filtered = { ...project, items: (project.items || []).filter((i: any) => i.includeInEstimate !== false) };
      const wb = await generateJustinsWorkbook(filtered);
      res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
      res.setHeader("Content-Disposition", `attachment; filename="${project.name.replace(/[^a-zA-Z0-9 _-]/g, "")} - Industry Standard.xlsx"`);
      await wb.xlsx.write(res);
      res.end();
    } catch (err: any) {
      console.error("Excel export error:", err);
      res.status(500).json({ message: err.message || "Export failed" });
    }
  });

  // ===== DIAGNOSE ESTIMATE =====
  // Re-runs the active estimating method over every BOM row and returns a
  // detailed breakdown: which calculator branch matched, what inputs it used,
  // the resulting MH/unit, and a list of project-level warnings (e.g. items
  // with no matching factor, double-counted welds-and-fittings, large nearest-
  // size matches). Pure read-only — the estimate is not mutated.
  //
  // Body: optional settings overrides. Defaults to the project's saved values.
  app.post("/api/estimates/:id/diagnose", (req, res) => {
    const project = storage.getEstimateProject(req.params.id);
    if (!project) return res.status(404).json({ message: "Estimate not found" });

    const diagSchema = z.object({
      method: z.enum(["bill", "justin", "industry"]).default((project as any).estimateMethod && ["bill","justin","industry"].includes((project as any).estimateMethod) ? (project as any).estimateMethod : "justin"),
      customMethodId: z.string().optional().default((project as any).customMethodId || ""),
      material: z.enum(["CS", "SS"]).default("CS"),
      schedule: z.string().default("STD"),
      installType: z.enum(["standard", "rack"]).default("standard"),
      pipeLocation: z.string().default("Open Rack"),
      elevation: z.string().default("0-20ft"),
      alloyGroup: z.string().default("4"),
      rackFactor: z.number().default(1.3),
      fittingWeldMode: z.enum(["bundled", "separate"]).default((project as any).fittingWeldMode || "bundled"),
    });
    const parsed = diagSchema.safeParse(req.body || {});
    if (!parsed.success) {
      return res.status(400).json({ message: "Validation failed", errors: parsed.error.flatten().fieldErrors });
    }
    const cfg = parsed.data;

    let estimatorData: any;
    try { estimatorData = getEstimatorData(); } catch (err: any) {
      return res.status(500).json({ message: `Failed to load estimator data: ${err.message}` });
    }

    // Resolve a custom profile if one is referenced.
    let activeData: any;
    let activeBase = cfg.method as "bill" | "justin" | "industry";
    let customMethodName = "";
    if (cfg.customMethodId) {
      const cm = storage.getCustomMethod(cfg.customMethodId);
      if (cm) {
        activeBase = cm.baseMethod;
        activeData = applyCustomOverrides(estimatorData[cm.baseMethod], cm.overrides || {});
        customMethodName = cm.name;
      } else {
        activeData = estimatorData[activeBase];
      }
    } else {
      activeData = estimatorData[activeBase];
    }
    if (!activeData) {
      return res.status(500).json({ message: `No data available for method '${activeBase}'.` });
    }

    // Per-row breakdown.
    type RowDiag = {
      itemId: string;
      lineNumber: number;
      category: string;
      description: string;
      size: string;
      quantity: number;
      unit: string;
      mhPerUnit: number;
      totalMH: number;
      calcBasis: string;
      sizeMatchExact: boolean;
      warnings: string[];
    };
    const rows: RowDiag[] = [];

    for (const item of (project.items || [])) {
      const itemMat = ((item as any).itemMaterial || cfg.material) as "CS" | "SS";
      const itemSched = (item as any).itemSchedule || cfg.schedule;
      const itemElev = (item as any).itemElevation || cfg.elevation;
      const itemLoc = (item as any).itemPipeLocation || cfg.pipeLocation;
      const itemAlloy = (item as any).itemAlloyGroup || cfg.alloyGroup;
      const lineWorkType = ((item as any).workType || cfg.installType) as "standard" | "rack";

      let mhPerUnit = 0;
      let calcBasis = "";
      let sizeMatchExact = true;
      const warnings: string[] = [];

      try {
        if (activeBase === "bill") {
          const r = calculateBillLaborHours(item, itemMat, itemSched, activeData, itemLoc, itemElev, itemAlloy, cfg.fittingWeldMode);
          mhPerUnit = r.laborHoursPerUnit;
          calcBasis = r.calcBasis;
          sizeMatchExact = r.sizeMatchExact;
        } else if (activeBase === "industry") {
          const r = calculateIndustryLaborHours(item, lineWorkType, itemMat, itemSched, activeData, cfg.fittingWeldMode);
          const dataDefault = activeData.cost_params?.contingency_factor ?? 0.10;
          const override = (project as any).contingencyOverride;
          const cont = (typeof override === "number" && !Number.isNaN(override)) ? (override / 100) : dataDefault;
          mhPerUnit = r.mh * (1 + cont);
          calcBasis = `${r.calcBasis} \u00d7 ${(1 + cont).toFixed(2)} (${(cont*100).toFixed(1)}% contingency)`;
          sizeMatchExact = r.sizeMatchExact;
        } else {
          const r = calculateJustinLaborHours(item, lineWorkType, itemMat, itemSched, activeData, cfg.fittingWeldMode);
          const dataDefault = activeData.cost_params?.contingency_factor ?? 0.15;
          const override = (project as any).contingencyOverride;
          const cont = (typeof override === "number" && !Number.isNaN(override)) ? (override / 100) : dataDefault;
          mhPerUnit = r.mh * (1 + cont);
          calcBasis = `${r.calcBasis} \u00d7 ${(1 + cont).toFixed(2)} (${(cont*100).toFixed(1)}% contingency)`;
          sizeMatchExact = r.sizeMatchExact;
        }
      } catch (err: any) {
        calcBasis = `\u26A0 calculator error: ${err?.message || String(err)}`;
        warnings.push("calculator-error");
      }

      if (mhPerUnit === 0) warnings.push("no-matching-factor");
      if (!sizeMatchExact) warnings.push("nearest-size-used");
      if (calcBasis.includes("\u26A0")) warnings.push("flagged");
      const excluded = (item as any).includeInEstimate === false;
      if (excluded) warnings.push("excluded-from-estimate");

      rows.push({
        itemId: item.id,
        lineNumber: item.lineNumber,
        category: item.category,
        description: item.description,
        size: item.size,
        quantity: item.quantity,
        unit: item.unit,
        mhPerUnit,
        totalMH: excluded ? 0 : mhPerUnit * (item.quantity || 0),
        calcBasis: excluded ? `${calcBasis} \u2014 EXCLUDED FROM ESTIMATE` : calcBasis,
        sizeMatchExact,
        warnings,
      });
    }

    // ---- Project-level warnings ----
    const projectWarnings: { code: string; severity: "info" | "warn" | "error"; title: string; detail: string; affectedItemIds?: string[] }[] = [];

    // 1. Double-count detection: count fitting rows vs explicit weld rows by size.
    //    In bundled mode every fitting carries its own welds, so seeing explicit
    //    weld rows AT THE SAME SIZE as fittings usually means the welds are
    //    double-counted. Flag the affected size buckets.
    if (cfg.fittingWeldMode === "bundled") {
      const fittingsBySize = new Map<string, string[]>();
      const weldsBySize = new Map<string, string[]>();
      for (const it of (project.items || [])) {
        const cat = (it.category || "").toLowerCase();
        const desc = (it.description || "").toLowerCase();
        const isFitting = ["fitting","elbow","tee","reducer","cap","coupling","union"].includes(cat);
        const isWeld = cat === "weld" || desc.includes("butt weld") || /\bbw\b/.test(desc);
        const sizeKey = (it.size || "").trim();
        if (!sizeKey) continue;
        if (isFitting) {
          if (!fittingsBySize.has(sizeKey)) fittingsBySize.set(sizeKey, []);
          fittingsBySize.get(sizeKey)!.push(it.id);
        } else if (isWeld) {
          if (!weldsBySize.has(sizeKey)) weldsBySize.set(sizeKey, []);
          weldsBySize.get(sizeKey)!.push(it.id);
        }
      }
      const overlapSizes: string[] = [];
      const overlapItemIds: string[] = [];
      for (const [sz, fIds] of fittingsBySize) {
        const wIds = weldsBySize.get(sz);
        if (wIds && wIds.length > 0) {
          overlapSizes.push(sz);
          overlapItemIds.push(...fIds, ...wIds);
        }
      }
      if (overlapSizes.length > 0) {
        projectWarnings.push({
          code: "double-count-welds-fittings",
          severity: "warn",
          title: `Possible double-counting at size${overlapSizes.length>1?"s":""} ${overlapSizes.join(", ")}`,
          detail: `The BOM has BOTH fitting rows and explicit weld rows at the same size, while fitting-weld mode is "bundled" (fittings carry their own weld labor). Recommended fix: switch to "Separate weld rows" mode since your BOM already has weld rows — the math will count each weld once at the proper factor. Alternative: click "Strip Auto-Inferred Welds" to remove the inferred rows and keep bundled mode.`,
          affectedItemIds: overlapItemIds,
        });
      }
    } else {
      // In "separate" mode the opposite warning: fittings present but no welds.
      let fittingCount = 0; let weldCount = 0;
      for (const it of (project.items || [])) {
        const cat = (it.category || "").toLowerCase();
        const desc = (it.description || "").toLowerCase();
        if (["fitting","elbow","tee","reducer","cap","coupling","union"].includes(cat)) fittingCount++;
        else if (cat === "weld" || desc.includes("butt weld") || /\bbw\b/.test(desc)) weldCount++;
      }
      if (fittingCount > 0 && weldCount === 0) {
        projectWarnings.push({
          code: "separate-no-welds",
          severity: "warn",
          title: `Mode is "separate" but no weld rows exist`,
          detail: `In "separate" mode fittings only contribute handling labor (×0.15); the BOM needs explicit weld rows to capture the actual weld labor. Either switch back to "bundled" or add weld rows (try the "Infer Welds from Fittings" button).`,
        });
      }
    }

    // 2. Items with no matching factor.
    const noMatch = rows.filter(r => r.warnings.includes("no-matching-factor"));
    if (noMatch.length > 0) {
      projectWarnings.push({
        code: "no-matching-factor",
        severity: "error",
        title: `${noMatch.length} item${noMatch.length>1?"s have":" has"} no matching labor factor`,
        detail: `These rows are contributing zero MH to the estimate. Likely cause: category not recognized, size out of table range, or item type (e.g. specialty) not modeled. Check the calcBasis column to see why.`,
        affectedItemIds: noMatch.map(r => r.itemId),
      });
    }

    // 3. Nearest-size fallback count.
    const nearestUsed = rows.filter(r => r.warnings.includes("nearest-size-used"));
    if (nearestUsed.length > 0) {
      projectWarnings.push({
        code: "nearest-size-fallback",
        severity: "info",
        title: `${nearestUsed.length} item${nearestUsed.length>1?"s":""} used a nearest-size factor`,
        detail: `These rows fell back to the closest available size in the factor table. Verify the factor is appropriate, or add the missing size to the method's factor table.`,
        affectedItemIds: nearestUsed.map(r => r.itemId),
      });
    }

    // ---- Summary ----
    const totalMH = rows.reduce((acc, r) => acc + r.totalMH, 0);
    const itemsWithLabor = rows.filter(r => r.mhPerUnit > 0).length;
    res.json({
      estimateId: project.id,
      estimateName: project.name,
      method: activeBase,
      customMethodId: cfg.customMethodId || undefined,
      customMethodName: customMethodName || undefined,
      fittingWeldMode: cfg.fittingWeldMode,
      totalMH,
      itemCount: rows.length,
      itemsWithLabor,
      rows,
      warnings: projectWarnings,
    });
  });

  // Export a method's full factor tree to Excel. methodKey is one of
  // 'bill' / 'justin' / 'industry' (base method) or 'custom:<id>' (custom profile).
  // For custom profiles we apply overrides on top of the base before exporting,
  // so the workbook reflects the actual effective factors the calculator will use.
  app.get("/api/methods/:methodKey/export", async (req, res) => {
    const key = req.params.methodKey;
    try {
      const estimatorData = getEstimatorData();
      let methodData: any = null;
      let methodName = "";
      let fileLabel = "";
      if (key === "bill" || key === "justin" || key === "industry") {
        methodData = estimatorData[key];
        methodName = key === "bill" ? "Bill's EI Method" : key === "justin" ? "Justin's Factor Method" : "Industry Standard (Page)";
        fileLabel = methodName;
      } else if (key.startsWith("custom:")) {
        const customId = key.substring("custom:".length);
        const cm = storage.getCustomMethod(customId);
        if (!cm) return res.status(404).json({ message: "Custom method not found" });
        const base = estimatorData[cm.baseMethod];
        if (!base) return res.status(500).json({ message: `Base method '${cm.baseMethod}' missing on server\u2014redeploy main` });
        methodData = applyCustomOverrides(base, cm.overrides || {});
        methodName = `${cm.name} (custom \u00b7 base: ${cm.baseMethod})`;
        fileLabel = cm.name;
      } else {
        return res.status(400).json({ message: `Unknown method key '${key}'` });
      }
      if (!methodData) {
        return res.status(500).json({ message: `Method data unavailable for '${key}' \u2014 redeploy main` });
      }
      const wb = generateMethodFactorsWorkbook(key, methodName, methodData);
      const safeName = fileLabel.replace(/[^a-zA-Z0-9 _-]/g, "").trim() || key;
      res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
      res.setHeader("Content-Disposition", `attachment; filename="${safeName} - Factor Table.xlsx"`);
      await wb.xlsx.write(res);
      res.end();
    } catch (err: any) {
      console.error("Method export error:", err);
      res.status(500).json({ message: err.message || "Method export failed" });
    }
  });

  // Export the Compare-All view to Excel. Same body as /compare-methods
  // (customMethodIds, labor rates, etc.) — we re-run the comparison server-side
  // so the export reflects the latest settings, not a stale client-side payload.
  app.post("/api/estimates/:id/compare-methods/export", async (req, res) => {
    // Reuse the compare-methods handler logic by calling it internally. Simplest
    // approach: forge a request to the compare endpoint via a tiny adapter.
    // For now, replicate the logic since it's not too long; if we refactor the
    // compare handler into a reusable function later we can call that here.
    const project = storage.getEstimateProject(req.params.id);
    if (!project) return res.status(404).json({ message: "Estimate not found" });

    const settingsSchema = z.object({
      customMethodIds: z.array(z.string()).optional().default([]),
      laborRate: z.number().min(0).max(500).default(project.laborRate || 56),
      overtimeRate: z.number().min(0).max(500).default(project.overtimeRate || 79),
      doubleTimeRate: z.number().min(0).max(500).default(project.doubleTimeRate || 100),
      perDiem: z.number().min(0).max(500).default(project.perDiem || 75),
      overtimePercent: z.number().min(0).max(100).default(project.overtimePercent || 15),
      doubleTimePercent: z.number().min(0).max(100).default(project.doubleTimePercent || 2),
      material: z.enum(["CS", "SS"]).default("CS"),
      schedule: z.string().default("40"),
      installType: z.enum(["standard", "rack"]).default("standard"),
      pipeLocation: z.string().default("ground"),
      elevation: z.string().default("ground"),
      alloyGroup: z.string().default("CS"),
      rackFactor: z.number().default(1.3),
      fittingWeldMode: z.enum(["bundled", "separate"]).default((project as any).fittingWeldMode || "bundled"),
    });
    const parsed = settingsSchema.safeParse(req.body || {});
    if (!parsed.success) {
      return res.status(400).json({ message: "Validation failed", errors: parsed.error.flatten().fieldErrors });
    }
    const settings = parsed.data;

    let estimatorData: any;
    try { estimatorData = getEstimatorData(); } catch (err: any) {
      return res.status(500).json({ message: `Failed to load estimator data: ${err.message}` });
    }

    type RunSpec = { key: string; label: string; baseMethod: "bill" | "justin" | "industry"; data: any; customMethodId?: string };
    const runs: RunSpec[] = [];
    if (estimatorData.bill?.labor_rates)       runs.push({ key: "bill",     label: "Bill",            baseMethod: "bill",     data: estimatorData.bill });
    if (estimatorData.justin?.labor_factors)   runs.push({ key: "justin",   label: "Justin",          baseMethod: "justin",   data: estimatorData.justin });
    if (estimatorData.industry?.labor_factors) runs.push({ key: "industry", label: "Industry (Page)", baseMethod: "industry", data: estimatorData.industry });
    for (const cmId of settings.customMethodIds) {
      const cm = storage.getCustomMethod(cmId);
      if (!cm) continue;
      const baseBlock = estimatorData[cm.baseMethod];
      if (!baseBlock) continue;
      const overridden = applyCustomOverrides(baseBlock, cm.overrides || {});
      runs.push({ key: `custom:${cm.id}`, label: cm.name, baseMethod: cm.baseMethod, data: overridden, customMethodId: cm.id });
    }
    if (runs.length === 0) {
      return res.status(500).json({ message: "No estimator methods available to compare." });
    }

    const stPercent = Math.max(0, (100 - settings.overtimePercent - settings.doubleTimePercent) / 100);
    const otPercent = settings.overtimePercent / 100;
    const dtPercent = settings.doubleTimePercent / 100;
    const blendedRate = (settings.laborRate * stPercent) + (settings.overtimeRate * otPercent) + (settings.doubleTimeRate * dtPercent);
    const perDiemPerHour = settings.perDiem / 10;
    const effectiveRate = blendedRate + perDiemPerHour;

    const lineItems: any[] = [];
    const summaryAccum: Record<string, { totalMH: number; totalLaborCost: number; totalMaterialCost: number }> = {};
    for (const r of runs) summaryAccum[r.key] = { totalMH: 0, totalLaborCost: 0, totalMaterialCost: 0 };

    for (const item of (project.items || [])) {
      if ((item as any).includeInEstimate === false) continue;
      let detectedMat: "CS" | "SS" = settings.material;
      if ((item as any).itemMaterial) detectedMat = (item as any).itemMaterial as "CS" | "SS";
      else {
        const desc = (item.description || "").toUpperCase();
        if (/\b(SS|STAINLESS|TP304|TP316|304L?|316L?|A312|A182|A403)\b/.test(desc)) detectedMat = "SS";
      }
      const itemMat = detectedMat;
      const itemSched = (item as any).itemSchedule || settings.schedule;
      const itemElev = (item as any).itemElevation || settings.elevation;
      const itemPipeLoc = (item as any).itemPipeLocation || settings.pipeLocation;
      const itemAlloy = (item as any).itemAlloyGroup || settings.alloyGroup;
      const lineWorkType = (item as any).workType || settings.installType;
      const byMethod: Record<string, any> = {};
      for (const r of runs) {
        let mh = 0; let calcBasis = ""; let materialAdjust = 0;
        if (r.baseMethod === "bill") {
          const result = calculateBillLaborHours(item, itemMat, itemSched, r.data, itemPipeLoc, itemElev, itemAlloy, settings.fittingWeldMode);
          mh = result.laborHoursPerUnit;
          calcBasis = result.calcBasis;
          materialAdjust = result.materialUnitCostAdjust;
          if (lineWorkType === "rack" && settings.rackFactor > 1) mh *= settings.rackFactor;
        } else if (r.baseMethod === "industry") {
          const ir = calculateIndustryLaborHours(item, lineWorkType as "standard" | "rack", itemMat, itemSched, r.data, settings.fittingWeldMode);
          const dataDefault = r.data?.cost_params?.contingency_factor ?? 0.10;
          const override = (project as any).contingencyOverride;
          const cf = (typeof override === "number" && !Number.isNaN(override)) ? (override / 100) : dataDefault;
          mh = ir.mh * (1 + cf); calcBasis = ir.calcBasis;
        } else {
          const jr = calculateJustinLaborHours(item, lineWorkType as "standard" | "rack", itemMat, itemSched, r.data, settings.fittingWeldMode);
          const dataDefault = r.data?.cost_params?.contingency_factor ?? 0.15;
          const override = (project as any).contingencyOverride;
          const cf = (typeof override === "number" && !Number.isNaN(override)) ? (override / 100) : dataDefault;
          mh = jr.mh * (1 + cf); calcBasis = jr.calcBasis;
        }
        const laborCost = mh * effectiveRate * (item.quantity || 0);
        const materialUnitCost = (item.materialUnitCost && item.materialUnitCost > 0) ? item.materialUnitCost : materialAdjust;
        const materialCost = materialUnitCost * (item.quantity || 0);
        byMethod[r.key] = { mhPerUnit: mh, totalMH: mh * (item.quantity || 0), laborCost, materialCost, totalCost: laborCost + materialCost, calcBasis };
        summaryAccum[r.key].totalMH += mh * (item.quantity || 0);
        summaryAccum[r.key].totalLaborCost += laborCost;
        summaryAccum[r.key].totalMaterialCost += materialCost;
      }
      lineItems.push({
        itemId: item.id, lineNumber: item.lineNumber, category: item.category, description: item.description,
        size: item.size, quantity: item.quantity, unit: item.unit, byMethod,
      });
    }
    const summary = runs.map(r => ({
      key: r.key, label: r.label, baseMethod: r.baseMethod, customMethodId: r.customMethodId,
      totalMH: summaryAccum[r.key].totalMH,
      totalLaborCost: summaryAccum[r.key].totalLaborCost,
      totalMaterialCost: summaryAccum[r.key].totalMaterialCost,
      totalCost: summaryAccum[r.key].totalLaborCost + summaryAccum[r.key].totalMaterialCost,
    }));
    const compareResult = {
      estimateId: project.id, estimateName: project.name,
      itemCount: (project.items || []).length, effectiveLaborRate: effectiveRate,
      summary, lineItems,
    };
    try {
      const wb = generateCompareWorkbook(compareResult);
      const safeName = project.name.replace(/[^a-zA-Z0-9 _-]/g, "").trim() || "Estimate";
      res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
      res.setHeader("Content-Disposition", `attachment; filename="${safeName} - Method Comparison.xlsx"`);
      await wb.xlsx.write(res);
      res.end();
    } catch (err: any) {
      console.error("Compare export error:", err);
      res.status(500).json({ message: err.message || "Compare export failed" });
    }
  });

  // ===== DATA BACKUP / RESTORE (FIX 8) =====

  app.get("/api/backup", (_req, res) => {
    try {
      const backup = storage.getFullBackup();
      res.setHeader("Content-Type", "application/json");
      res.setHeader("Content-Disposition", `attachment; filename="pg-unified-backup-${new Date().toISOString().slice(0, 10)}.json"`);
      res.json(backup);
    } catch (err: any) {
      res.status(500).json({ message: err.message || "Backup failed" });
    }
  });

  app.post("/api/restore", (req, res) => {
    try {
      const data = req.body;
      if (!data || typeof data !== "object") {
        return res.status(400).json({ message: "Invalid backup data" });
      }
      // Validate backup structure
      const requiredKeys = ["takeoffProjects", "estimateProjects", "costDatabase", "purchaseHistory", "completedProjects"];
      for (const key of requiredKeys) {
        if (data[key] !== undefined && !Array.isArray(data[key])) {
          return res.status(400).json({ message: `Invalid backup: "${key}" must be an array` });
        }
      }
      if (!requiredKeys.some(k => Array.isArray(data[k]))) {
        return res.status(400).json({ message: "Invalid backup: must contain at least one data array (takeoffProjects, estimateProjects, costDatabase, purchaseHistory, or completedProjects)" });
      }
      const result = storage.restoreFromBackup(data);
      res.json({ message: "Restore complete", ...result });
    } catch (err: any) {
      res.status(500).json({ message: err.message || "Restore failed" });
    }
  });

  // ===== ESTIMATE VERSIONING (Feature 3) =====

  app.get("/api/estimates/:id/versions", (req, res) => {
    const project = storage.getEstimateProject(req.params.id);
    if (!project) return res.status(404).json({ message: "Estimate not found" });
    const versions = storage.getEstimateVersions(req.params.id);
    res.json(versions);
  });

  app.post("/api/estimates/:id/restore-version/:versionId", (req, res) => {
    const success = storage.restoreEstimateVersion(req.params.id, req.params.versionId);
    if (!success) return res.status(404).json({ message: "Version not found" });
    const updated = storage.getEstimateProject(req.params.id);
    res.json(updated);
  });

  // ===== CSV/XLSX COST DATABASE IMPORT (Feature 4) =====

  const csvUpload = multer({
    storage: multer.diskStorage({
      destination: UPLOAD_DIR,
      filename: (_req, file, cb) => cb(null, `${Date.now()}_${file.originalname.replace(/[^a-zA-Z0-9._-]/g, "_")}`),
    }),
    limits: { fileSize: 20 * 1024 * 1024 },
  });

  app.post("/api/cost-database/import", (req, res) => {
    csvUpload.single("file")(req, res, async (multerErr: any) => {
      try {
        if (multerErr) return res.status(400).json({ error: multerErr.message });
        if (!req.file) return res.status(400).json({ error: "No file received." });

        const filePath = req.file.path;
        const ext = path.extname(req.file.originalname).toLowerCase();
        let records: any[] = [];

        if (ext === ".csv") {
          const csvContent = fs.readFileSync(filePath, "utf-8");
          records = csvParse(csvContent, { columns: true, skip_empty_lines: true, trim: true });
        } else if (ext === ".xlsx" || ext === ".xls") {
          const ExcelJS = await import("exceljs");
          const wb = new ExcelJS.default.Workbook();
          await wb.xlsx.readFile(filePath);
          const ws = wb.worksheets[0];
          if (!ws) return res.status(400).json({ error: "No worksheet found" });
          const headers: string[] = [];
          ws.getRow(1).eachCell((cell: any, colNumber: number) => {
            headers[colNumber - 1] = String(cell.value || "").trim();
          });
          ws.eachRow((row: any, rowNumber: number) => {
            if (rowNumber === 1) return;
            const obj: any = {};
            row.eachCell((cell: any, colNumber: number) => {
              obj[headers[colNumber - 1] || `col${colNumber}`] = cell.value;
            });
            records.push(obj);
          });
        } else {
          return res.status(400).json({ error: "Unsupported file type. Use CSV or XLSX." });
        }

        // Clean up temp file
        try { fs.unlinkSync(filePath); } catch (e) { console.warn("Suppressed error:", e); }

        // Map columns — collect valid entries, then bulk insert in transaction
        const validEntries: Array<{description: string; size: string; category: string; unit: string; materialUnitCost: number; laborUnitCost: number; laborHoursPerUnit: number}> = [];
        for (const row of records) {
          const description = String(row.Description || row.description || row.DESCRIPTION || "").trim();
          const size = String(row.Size || row.size || row.SIZE || "").trim();
          const category = String(row.Category || row.category || row.CATEGORY || "other").trim();
          const unit = String(row.Unit || row.unit || row.UNIT || "EA").trim();
          const materialUnitCost = parseFloat(row["Material Cost"] || row.materialUnitCost || row.material_cost || 0) || 0;
          const laborUnitCost = parseFloat(row["Labor Cost"] || row.laborUnitCost || row.labor_cost || 0) || 0;
          const laborHoursPerUnit = parseFloat(row["Labor Hours"] || row.laborHoursPerUnit || row.labor_hours || 0) || 0;

          if (!description) continue;
          validEntries.push({ description, size, category, unit, materialUnitCost, laborUnitCost, laborHoursPerUnit });
        }

        const imported = storage.addCostEntriesBulk(validEntries);
        res.json({ message: `Imported ${imported} entries`, imported });
      } catch (err: any) {
        res.status(500).json({ error: err.message || "Import failed" });
      }
    });
  });

  // ===== PURCHASE HISTORY ENDPOINTS =====

  app.get("/api/purchase-history", (_req, res) => {
    const { supplier, category } = _req.query as { supplier?: string; category?: string };
    const records = storage.getPurchaseRecords({ supplier, category });
    res.json(records);
  });

  app.get("/api/purchase-history/suppliers", (_req, res) => {
    res.json(storage.getPurchaseSuppliers());
  });

  app.get("/api/purchase-history/categories", (_req, res) => {
    res.json(storage.getPurchaseCategories());
  });

  app.post("/api/purchase-history", (req, res) => {
    const parsed = insertPurchaseRecordSchema.safeParse(req.body);
    if (!parsed.success) return res.status(400).json({ error: "Invalid data", details: parsed.error.flatten().fieldErrors });
    const record = storage.addPurchaseRecord(parsed.data);
    res.json(record);
  });

  app.post("/api/purchase-history/bulk", (req, res) => {
    const { records } = req.body as { records: any[] };
    if (!records || !Array.isArray(records)) return res.status(400).json({ error: "records array required" });
    // Validate each record has minimum required fields
    const validRecords = records.filter(r => {
      if (!r.description || typeof r.description !== "string" || r.description.trim().length < 1) return false;
      if (r.unitCost != null && typeof r.unitCost !== "number") return false;
      return true;
    });
    const skipped = records.length - validRecords.length;
    const count = storage.addPurchaseRecordsBulk(validRecords);
    res.json({ imported: count, skipped, total: records.length });
  });

  app.delete("/api/purchase-history/:id", (req, res) => {
    const deleted = storage.deletePurchaseRecord(req.params.id);
    if (!deleted) return res.status(404).json({ error: "Record not found" });
    res.json({ deleted: true });
  });

  app.delete("/api/purchase-history", (_req, res) => {
    const count = storage.clearPurchaseHistory();
    res.json({ cleared: count });
  });

  app.post("/api/purchase-history/import", (req, res) => {
    csvUpload.single("file")(req, res, async (multerErr: any) => {
      try {
        if (multerErr) return res.status(400).json({ error: multerErr.message });
        if (!req.file) return res.status(400).json({ error: "No file received." });

        const filePath = req.file.path;
        const ext = path.extname(req.file.originalname).toLowerCase();
        let records: any[] = [];

        if (ext === ".csv") {
          const csvContent = fs.readFileSync(filePath, "utf-8");
          records = csvParse(csvContent, { columns: true, skip_empty_lines: true, trim: true });
        } else if (ext === ".xlsx" || ext === ".xls") {
          const ExcelJS = await import("exceljs");
          const wb = new ExcelJS.default.Workbook();
          await wb.xlsx.readFile(filePath);
          const ws = wb.worksheets[0];
          if (!ws) return res.status(400).json({ error: "No worksheet found" });
          const headers: string[] = [];
          ws.getRow(1).eachCell((cell: any, colNumber: number) => {
            headers[colNumber - 1] = String(cell.value || "").trim();
          });
          ws.eachRow((row: any, rowNumber: number) => {
            if (rowNumber === 1) return;
            const obj: any = {};
            row.eachCell((cell: any, colNumber: number) => {
              obj[headers[colNumber - 1] || `col${colNumber}`] = cell.value;
            });
            records.push(obj);
          });
        } else {
          return res.status(400).json({ error: "Unsupported file type. Use CSV or XLSX." });
        }

        try { fs.unlinkSync(filePath); } catch (e) { console.warn("Suppressed error:", e); }

        // Map CSV columns to purchase records
        const mapped = records.map(row => ({
          description: String(row.description || row.Description || row.DESCRIPTION || "").trim(),
          size: String(row.size || row.Size || row.SIZE || "").trim() || undefined,
          category: String(row.category || row.Category || row.CATEGORY || "other").trim(),
          material: String(row.material || row.Material || "").trim() || undefined,
          schedule: String(row.schedule || row.Schedule || "").trim() || undefined,
          rating: String(row.rating || row.Rating || "").trim() || undefined,
          connectionType: String(row.connectionType || row.ConnectionType || "").trim() || undefined,
          unit: String(row.unit || row.Unit || row.UNIT || "EA").trim(),
          unitCost: parseFloat(row.unitCost || row.materialUnitCost || row.UnitCost || row["Unit Cost"] || 0) || 0,
          quantity: parseFloat(row.quantity || row.Quantity || row.qty || 1) || 1,
          supplier: String(row.supplier || row.Supplier || row.SUPPLIER || "Unknown").trim(),
          invoiceNumber: String(row.invoiceNumber || row.InvoiceNumber || row["Invoice #"] || "").trim() || undefined,
          invoiceDate: String(row.invoiceDate || row.InvoiceDate || row["Invoice Date"] || "").trim() || undefined,
          project: String(row.project || row.Project || "").trim() || undefined,
          poNumber: String(row.poNumber || row.PONumber || row["PO #"] || "").trim() || undefined,
          sourceFile: String(row.sourceFile || row.filename || "").trim() || undefined,
        })).filter(r => r.description);

        const count = storage.addPurchaseRecordsBulk(mapped);
        res.json({ message: `Imported ${count} purchase records`, imported: count });
      } catch (err: any) {
        res.status(500).json({ error: err.message || "Import failed" });
      }
    });
  });

  app.get("/api/purchase-history/latest-cost", (req, res) => {
    const { description, size } = req.query as { description: string; size?: string };
    if (!description) return res.status(400).json({ error: "description required" });
    const match = storage.getLatestCostForItem(description, size);
    res.json(match || null);
  });

  // ===== PROJECT HISTORY ENDPOINTS =====

  app.get("/api/project-history", (_req, res) => {
    const projects = storage.getCompletedProjects();
    res.json(projects);
  });

  app.get("/api/project-history/search", (req, res) => {
    const q = (req.query.q as string) || "";
    if (!q.trim()) {
      return res.json(storage.getCompletedProjects());
    }
    const results = storage.searchCompletedProjects(q.trim());
    res.json(results);
  });

  app.post("/api/project-history", (req, res) => {
    const data = req.body;
    if (!data.name || !data.scopeDescription) {
      return res.status(400).json({ error: "name and scopeDescription are required" });
    }
    const project = storage.addCompletedProject(data);
    res.json(project);
  });

  app.delete("/api/project-history/:id", (req, res) => {
    const deleted = storage.deleteCompletedProject(req.params.id);
    if (!deleted) return res.status(404).json({ error: "Project not found" });
    res.json({ deleted: true });
  });

  // ===== WELD INFERENCE ENDPOINT (Feature 2) =====

  app.post("/api/estimates/:id/infer-welds", (req, res) => {
    const project = storage.getEstimateProject(req.params.id);
    if (!project) return res.status(404).json({ message: "Estimate not found" });

    const inferredWelds = inferWeldsFromFittings(project.items || []);
    if (inferredWelds.length === 0) {
      return res.json({ message: "No fittings found to infer welds from", added: 0 });
    }

    const allItems = [...(project.items || []), ...inferredWelds].map((item, idx) => ({ ...item, lineNumber: idx + 1 }));
    const updated = storage.updateEstimateProject(req.params.id, { items: allItems });
    res.json({ message: `Inferred ${inferredWelds.length} weld/bolt items`, added: inferredWelds.length, project: updated });
  });

  // Remove every BOM row produced by the auto-infer feature. These are tagged
  // in their `notes` field with the string "auto-inferred" so we can identify
  // and delete them safely without touching hand-entered items. Used when the
  // user is in bundled fitting-weld mode and wants to eliminate double-counting
  // by removing the inferred weld rows.
  app.post("/api/estimates/:id/strip-inferred-welds", (req, res) => {
    const project = storage.getEstimateProject(req.params.id);
    if (!project) return res.status(404).json({ message: "Estimate not found" });

    const before = (project.items || []).length;
    const keep = (project.items || []).filter((it: any) => {
      const notes = (it.notes || "").toLowerCase();
      return !notes.includes("auto-inferred");
    });
    const removed = before - keep.length;
    if (removed === 0) {
      return res.json({ message: "No auto-inferred rows to remove", removed: 0 });
    }
    // Renumber lineNumbers so the remaining rows stay sequential.
    const relabeled = keep.map((item: any, idx: number) => ({ ...item, lineNumber: idx + 1 }));
    const updated = storage.updateEstimateProject(req.params.id, { items: relabeled });
    res.json({ message: `Removed ${removed} auto-inferred row${removed === 1 ? "" : "s"}`, removed, project: updated });
  });

  // === API KEY SETTINGS ===

  // Get current API key status
  app.get("/api/settings/api-key", (_req, res) => {
    const key = getUserApiKey();
    if (key) {
      // Return masked version
      const masked = key.substring(0, 7) + "..." + key.substring(key.length - 4);
      res.json({ configured: true, masked, source: "user" });
    } else {
      const envKey = process.env.ANTHROPIC_API_KEY || process.env.LLM_API_KEY || "";
      if (envKey) {
        res.json({ configured: true, masked: "Platform-provided", source: "platform" });
      } else {
        res.json({ configured: false, masked: null, source: null });
      }
    }
  });

  // Set user API key
  app.post("/api/settings/api-key", (req, res) => {
    const { apiKey } = req.body;
    if (!apiKey || typeof apiKey !== "string" || !apiKey.startsWith("sk-ant-")) {
      return res.status(400).json({ error: "Invalid API key. It should start with 'sk-ant-'" });
    }
    setUserApiKey(apiKey.trim());
    // Persist encrypted so it survives restarts
    try {
      storage.saveEncryptedApiKey(apiKey.trim());
    } catch (e) { console.warn("Suppressed error:", e); }
    const masked = apiKey.substring(0, 7) + "..." + apiKey.substring(apiKey.length - 4);
    console.log(`API key configured: ${masked}`);
    res.json({ success: true, masked, source: "user" });
  });

  // Remove user API key (fall back to platform)
  app.delete("/api/settings/api-key", (_req, res) => {
    setUserApiKey(null);
    try {
      const keyFile = path.join(process.cwd(), "data", ".api-key");
      if (fs.existsSync(keyFile)) fs.unlinkSync(keyFile); // Remove any existing file (encrypted or plain)
    } catch (e) { console.warn("Suppressed error:", e); }
    res.json({ success: true });
  });

  // Test API key
  app.post("/api/settings/test-api-key", async (req, res) => {
    const { apiKey } = req.body;
    const testKey = apiKey || getUserApiKey() || process.env.ANTHROPIC_API_KEY || "";
    try {
      // Always use direct Anthropic URL when testing a user key (bypass platform proxy)
      const isUserKey = testKey.startsWith("sk-ant-");
      const client = new Anthropic({ apiKey: testKey, ...(isUserKey ? { baseURL: "https://api.anthropic.com" } : {}) });
      const msg = await client.messages.create({
        model: "claude-sonnet-4-5",
        max_tokens: 20,
        messages: [{ role: "user", content: "Reply with just the word OK" }],
      });
      const text = msg.content[0].type === "text" ? msg.content[0].text : "";
      res.json({ success: true, response: text.trim() });
    } catch (err: any) {
      res.json({ success: false, error: err.message || "API call failed" });
    }
  });

  // Load saved API key on startup (supports both encrypted and legacy plain-text)
  try {
    const loaded = storage.loadEncryptedApiKey();
    if (loaded && loaded.startsWith("sk-ant-")) {
      setUserApiKey(loaded);
      console.log(`Loaded saved API key: ${loaded.substring(0, 7)}...${loaded.substring(loaded.length - 4)}`);
    }
  } catch (e) { console.warn("Suppressed error:", e); }

  // ===== FEATURE 1: Page Thumbnail Serving =====
  app.get("/api/takeoff/projects/:id/page/:pageNum", (req, res) => {
    const { id, pageNum } = req.params;
    // Path traversal protection: validate pageNum is a positive integer and id has no path separators
    if (!/^\d+$/.test(pageNum) || parseInt(pageNum, 10) < 1) {
      return res.status(400).json({ error: "Invalid page number" });
    }
    if (/[\/\\.]/.test(id)) {
      return res.status(400).json({ error: "Invalid project ID" });
    }
    const thumbPath = path.join(THUMBNAIL_DIR, id, `page-${pageNum}.png`);
    if (!fs.existsSync(thumbPath)) {
      return res.status(404).json({ error: "Page thumbnail not found" });
    }
    res.setHeader("Content-Type", "image/png");
    res.setHeader("Cache-Control", "public, max-age=86400");
    fs.createReadStream(thumbPath).pipe(res);
  });

  // List available page thumbnails for a project
  app.get("/api/takeoff/projects/:id/pages", (req, res) => {
    if (/[\/\\.]/.test(req.params.id)) {
      return res.status(400).json({ error: "Invalid project ID" });
    }
    const projDir = path.join(THUMBNAIL_DIR, req.params.id);
    if (!fs.existsSync(projDir)) return res.json({ pages: [] });
    const files = fs.readdirSync(projDir).filter(f => f.endsWith(".png")).sort();
    const pages = files.map(f => {
      const m = f.match(/page-(\d+)\.png/);
      return m ? parseInt(m[1], 10) : 0;
    }).filter(n => n > 0).sort((a, b) => a - b);
    res.json({ pages });
  });

  // ===== FEATURE 4: Scope Gap Detection =====
  app.get("/api/takeoff/projects/:id/scope-gaps", async (req, res) => {
    const project = await storage.getTakeoffProject(req.params.id);
    if (!project) return res.status(404).json({ error: "Project not found" });

    const gaps: { type: string; severity: "warning" | "info"; message: string; affectedItems: string[] }[] = [];
    const items = project.items || [];
    if (items.length === 0) return res.json({ gaps });

    // Group by size
    const flangesBySize: Record<string, any[]> = {};
    const gasketsBySize: Record<string, any[]> = {};
    const boltsBySize: Record<string, any[]> = {};
    const pipeBySize: Record<string, any[]> = {};
    const fittingsBySize: Record<string, any[]> = {};
    const valvesBySize: Record<string, any[]> = {};
    let totalWelds = 0;
    let totalFittings = 0;

    for (const item of items) {
      const cat = (item.category || "").toLowerCase();
      const size = (item.size || "").toLowerCase().trim();
      if (!size || size === "n/a") continue;

      if (cat === "flange") {
        if (!flangesBySize[size]) flangesBySize[size] = [];
        flangesBySize[size].push(item);
      } else if (cat === "gasket") {
        if (!gasketsBySize[size]) gasketsBySize[size] = [];
        gasketsBySize[size].push(item);
      } else if (cat === "bolt") {
        if (!boltsBySize[size]) boltsBySize[size] = [];
        boltsBySize[size].push(item);
      } else if (cat === "pipe") {
        if (!pipeBySize[size]) pipeBySize[size] = [];
        pipeBySize[size].push(item);
      } else if (["elbow", "tee", "reducer", "cap", "coupling"].includes(cat)) {
        if (!fittingsBySize[size]) fittingsBySize[size] = [];
        fittingsBySize[size].push(item);
        totalFittings += item.quantity;
      } else if (cat === "valve") {
        if (!valvesBySize[size]) valvesBySize[size] = [];
        valvesBySize[size].push(item);
      } else if (cat === "weld") {
        totalWelds += item.quantity;
      }
    }

    // Flanges without gaskets
    for (const [size, flanges] of Object.entries(flangesBySize)) {
      if (!gasketsBySize[size]) {
        gaps.push({
          type: "missing_gasket",
          severity: "warning",
          message: `Missing gasket for ${size.toUpperCase()} flanges`,
          affectedItems: flanges.map(f => f.id),
        });
      }
    }

    // Flanges without bolts
    for (const [size, flanges] of Object.entries(flangesBySize)) {
      if (!boltsBySize[size]) {
        gaps.push({
          type: "missing_bolts",
          severity: "warning",
          message: `Missing bolt set for ${size.toUpperCase()} flanges`,
          affectedItems: flanges.map(f => f.id),
        });
      }
    }

    // Pipe with no welds
    const totalPipeLF = Object.values(pipeBySize).flat().reduce((sum, i) => sum + i.quantity, 0);
    if (totalPipeLF > 0 && totalWelds === 0) {
      gaps.push({
        type: "no_welds",
        severity: "info",
        message: `${totalPipeLF.toFixed(0)} LF of pipe with no welds — verify`,
        affectedItems: Object.values(pipeBySize).flat().map(i => i.id),
      });
    }

    // Fittings found for a size but no pipe of that size
    for (const [size, fittings] of Object.entries(fittingsBySize)) {
      if (!pipeBySize[size]) {
        gaps.push({
          type: "fittings_no_pipe",
          severity: "info",
          message: `Fittings found for ${size.toUpperCase()} but no pipe — verify`,
          affectedItems: fittings.map(f => f.id),
        });
      }
    }

    // Very few welds relative to fittings
    if (totalFittings > 5 && totalWelds > 0 && totalWelds < totalFittings * 0.5) {
      gaps.push({
        type: "low_weld_ratio",
        severity: "info",
        message: `Only ${totalWelds} welds for ${totalFittings} fittings — typical ratio is ~2:1`,
        affectedItems: [],
      });
    }

    // Flanged valves without flanges
    for (const [size, valves] of Object.entries(valvesBySize)) {
      const hasFlangedValve = valves.some(v => /flang/i.test(v.description));
      if (hasFlangedValve && !flangesBySize[size]) {
        gaps.push({
          type: "valve_no_flanges",
          severity: "warning",
          message: `Flanged valve found at ${size.toUpperCase()} but no matching flanges`,
          affectedItems: valves.map(v => v.id),
        });
      }
    }

    res.json({ gaps });
  });

  // ===== FEATURE 3: Change Order Generator =====
  app.post("/api/takeoff/projects/:id/change-order", async (req, res) => {
    const project = await storage.getTakeoffProject(req.params.id);
    if (!project) return res.status(404).json({ error: "Project not found" });

    const cloudedItems = (project.items || []).filter(i => i.revisionClouded);
    const nonCloudedItems = (project.items || []).filter(i => !i.revisionClouded);

    if (cloudedItems.length === 0) {
      return res.json({ addedItems: [], removedItems: [], costImpact: { laborHours: 0, laborCost: 0, materialCost: 0, totalImpact: 0 } });
    }

    // Items that are clouded represent changes — these are the "added scope" delta
    const addedItems = cloudedItems.map(item => ({
      id: item.id,
      lineNumber: item.lineNumber,
      category: item.category,
      description: item.description,
      size: item.size,
      quantity: item.quantity,
      unit: item.unit,
      notes: item.notes,
    }));

    // Compute cost impact if an estimate exists
    let costImpact = { laborHours: 0, laborCost: 0, materialCost: 0, totalImpact: 0 };

    // Find estimate linked to this takeoff (targeted query instead of loading all estimates)
    const linkedEstimate = storage.getEstimateBySourceTakeoff(project.id);
    if (linkedEstimate) {
      const cloudedIds = new Set(cloudedItems.map(i => i.id));
      // Match estimate items by line number to clouded takeoff items
      for (const estItem of (linkedEstimate.items || [])) {
        // Check if this estimate item corresponds to a clouded takeoff item
        if (estItem.revisionClouded) {
          costImpact.laborHours += (estItem.quantity || 0) * (estItem.laborHoursPerUnit || 0);
          costImpact.laborCost += estItem.laborExtension || 0;
          costImpact.materialCost += estItem.materialExtension || 0;
        }
      }
      costImpact.totalImpact = costImpact.laborCost + costImpact.materialCost;
    }

    res.json({ addedItems, removedItems: [], costImpact });
  });

  // ===== GEMINI API KEY SETTINGS =====
  app.get("/api/settings/gemini-key", (_req, res) => {
    const key = getUserGeminiKey();
    if (key) {
      const masked = key.substring(0, 6) + "..." + key.substring(key.length - 4);
      res.json({ configured: true, masked });
    } else {
      res.json({ configured: false, masked: null });
    }
  });

  app.post("/api/settings/gemini-key", (req, res) => {
    const { apiKey } = req.body;
    if (!apiKey || typeof apiKey !== "string" || apiKey.length < 10) {
      return res.status(400).json({ error: "Invalid Gemini API key" });
    }
    setUserGeminiKey(apiKey.trim());
    try { storage.saveEncryptedGeminiKey(apiKey.trim()); } catch (e) { console.warn("Suppressed error:", e); }
    const masked = apiKey.substring(0, 6) + "..." + apiKey.substring(apiKey.length - 4);
    console.log(`Gemini API key configured: ${masked}`);
    res.json({ success: true, masked });
  });

  app.delete("/api/settings/gemini-key", (_req, res) => {
    setUserGeminiKey(null);
    try {
      const keyFile = path.join(process.cwd(), "data", ".gemini-key");
      if (fs.existsSync(keyFile)) fs.unlinkSync(keyFile);
    } catch (e) { console.warn("Suppressed error:", e); }
    res.json({ success: true });
  });

  app.post("/api/settings/test-gemini-key", async (req, res) => {
    const { apiKey } = req.body;
    const testKey = apiKey || getUserGeminiKey();
    if (!testKey) return res.json({ success: false, error: "No Gemini key configured" });
    try {
      const resp = await fetch(
        `https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent`,
        {
          method: "POST",
          headers: { "Content-Type": "application/json", "x-goog-api-key": testKey },
          body: JSON.stringify({
            contents: [{ parts: [{ text: "Reply with just the word OK" }] }],
            generationConfig: { temperature: 0, maxOutputTokens: 10 },
          }),
        }
      );
      if (!resp.ok) {
        const errData = await resp.json().catch(() => ({}));
        return res.json({ success: false, error: errData?.error?.message || `HTTP ${resp.status}` });
      }
      const data = await resp.json();
      const text = data?.candidates?.[0]?.content?.parts?.[0]?.text || "";
      res.json({ success: true, response: text.trim() });
    } catch (err: any) {
      res.json({ success: false, error: err.message || "Connection failed" });
    }
  });

  // Load saved Gemini key on startup
  try {
    const loadedGemini = storage.loadEncryptedGeminiKey();
    if (loadedGemini && loadedGemini.length > 10) {
      setUserGeminiKey(loadedGemini);
      console.log(`Loaded saved Gemini key: ${loadedGemini.substring(0, 6)}...${loadedGemini.substring(loadedGemini.length - 4)}`);
    }
  } catch (e) { console.warn("Suppressed error:", e); }

  // ===== DRAWING TEMPLATE ENDPOINTS =====
  app.get("/api/drawing-templates", (_req, res) => {
    res.json(storage.getDrawingTemplates());
  });

  app.post("/api/drawing-templates", (req, res) => {
    const { name, engineeringFirm, bomLayout, columnOrder, commonAbbreviations, sampleOcrText, matchPatterns, extractionNotes } = req.body;
    if (!name) return res.status(400).json({ error: "name is required" });
    const template = storage.createDrawingTemplate({
      name, engineeringFirm, bomLayout, columnOrder, commonAbbreviations,
      sampleOcrText, matchPatterns, extractionNotes,
    });
    res.status(201).json(template);
  });

  app.delete("/api/drawing-templates/:id", (req, res) => {
    const deleted = storage.deleteDrawingTemplate(req.params.id);
    if (!deleted) return res.status(404).json({ error: "Template not found" });
    res.json({ deleted: true });
  });

  // ===== MATERIAL ESCALATION ALERTS =====
  app.get("/api/cost-database/alerts", (_req, res) => {
    const alerts = storage.getMaterialAlerts();
    res.json(alerts);
  });

  // ===== VENDOR QUOTE ENDPOINTS =====
  app.get("/api/vendor-quotes", (_req, res) => {
    res.json(storage.getVendorQuotes());
  });

  app.post("/api/vendor-quotes", (req, res) => {
    const parsed = insertVendorQuoteSchema.safeParse(req.body);
    if (!parsed.success) return res.status(400).json({ error: "Invalid data", details: parsed.error.flatten().fieldErrors });
    const d = parsed.data;
    const quote = storage.createVendorQuote({
      ...d,
      totalPrice: d.totalPrice ?? d.unitPrice * d.quantity,
    });
    res.status(201).json(quote);
  });

  app.delete("/api/vendor-quotes/:id", (req, res) => {
    const deleted = storage.deleteVendorQuote(req.params.id);
    if (!deleted) return res.status(404).json({ error: "Quote not found" });
    res.json({ deleted: true });
  });

  // Vendor quote CSV/XLSX import — use csvUpload (not pdf-only upload)
  app.post("/api/vendor-quotes/import", (req, res) => {
    csvUpload.single("file")(req, res, async () => {
      if (!req.file) return res.status(400).json({ error: "No file uploaded" });
      try {
        const filePath = req.file.path;
        const ext = path.extname(req.file.originalname).toLowerCase();
        let records: any[] = [];

        if (ext === ".csv") {
          const csvContent = fs.readFileSync(filePath, "utf-8");
          records = csvParse(csvContent, { columns: true, skip_empty_lines: true, trim: true });
        } else if (ext === ".xlsx" || ext === ".xls") {
          const ExcelJS = await import("exceljs");
          const wb = new ExcelJS.default.Workbook();
          await wb.xlsx.readFile(filePath);
          const ws = wb.worksheets[0];
          if (!ws) return res.status(400).json({ error: "No worksheet found" });
          const headers: string[] = [];
          ws.getRow(1).eachCell((cell: any, colNumber: number) => {
            headers[colNumber - 1] = String(cell.value || "").trim();
          });
          ws.eachRow((row: any, rowNumber: number) => {
            if (rowNumber === 1) return;
            const obj: any = {};
            row.eachCell((cell: any, colNumber: number) => {
              obj[headers[colNumber - 1] || `col${colNumber}`] = cell.value;
            });
            records.push(obj);
          });
        } else {
          return res.status(400).json({ error: "Unsupported file type. Use CSV or XLSX." });
        }

        try { fs.unlinkSync(filePath); } catch (e) { console.warn("Suppressed error:", e); }

        const mapped = records.map(row => ({
          vendorName: String(row.vendorName || row.Vendor || row.vendor || row.VENDOR || "Unknown").trim(),
          quoteNumber: String(row.quoteNumber || row.QuoteNumber || row["Quote #"] || "").trim() || undefined,
          quoteDate: String(row.quoteDate || row.QuoteDate || row["Quote Date"] || "").trim() || undefined,
          projectName: String(row.projectName || row.Project || row.project || "").trim() || undefined,
          description: String(row.description || row.Description || row.DESCRIPTION || "").trim(),
          size: String(row.size || row.Size || row.SIZE || "").trim() || undefined,
          category: String(row.category || row.Category || "other").trim(),
          unit: String(row.unit || row.Unit || "EA").trim(),
          unitPrice: parseFloat(row.unitPrice || row.UnitPrice || row["Unit Price"] || 0) || 0,
          quantity: parseFloat(row.quantity || row.Quantity || row.qty || 1) || 1,
          totalPrice: parseFloat(row.totalPrice || row.TotalPrice || row["Total Price"] || 0) || 0,
          notes: String(row.notes || row.Notes || "").trim() || undefined,
        })).filter(r => r.description);

        // Calculate totalPrice if not provided
        mapped.forEach(r => {
          if (!r.totalPrice && r.unitPrice > 0) r.totalPrice = r.unitPrice * r.quantity;
        });

        const count = storage.createVendorQuotesBulk(mapped);
        res.json({ message: `Imported ${count} vendor quotes`, imported: count });
      } catch (err: any) {
        res.status(500).json({ error: err.message || "Import failed" });
      }
    });
  });

  // Vendor quote comparison — group by description+size, compare across vendors
  app.get("/api/vendor-quotes/compare", (_req, res) => {
    const quotes = storage.getVendorQuotes();
    const groups: Record<string, any[]> = {};
    for (const q of quotes) {
      const key = `${(q as any).description.toLowerCase().trim()}|${((q as any).size || "").toLowerCase().trim()}`;
      if (!groups[key]) groups[key] = [];
      groups[key].push(q);
    }
    // Only return items with 2+ vendors
    const comparisons = Object.entries(groups)
      .filter(([, items]) => {
        const vendors = new Set(items.map((i: any) => i.vendorName));
        return vendors.size > 1;
      })
      .map(([key, items]) => {
        const [desc, size] = key.split("|");
        const sorted = [...items].sort((a: any, b: any) => a.unitPrice - b.unitPrice);
        const lowest = sorted[0].unitPrice;
        return {
          description: desc,
          size,
          quotes: sorted.map((q: any) => ({
            ...q,
            savings: lowest > 0 ? ((q.unitPrice - lowest) / lowest * 100) : 0,
            isBest: q.unitPrice === lowest,
          })),
          bestVendor: sorted[0].vendorName,
          bestPrice: lowest,
          highestPrice: sorted[sorted.length - 1].unitPrice,
          potentialSavings: sorted[sorted.length - 1].unitPrice - lowest,
        };
      })
      .sort((a, b) => b.potentialSavings - a.potentialSavings);
    res.json(comparisons);
  });

  // ===== QUICK RE-ESTIMATE FROM HISTORY =====
  app.post("/api/project-history/quick-estimate", (req, res) => {
    const { scopeDescription, tags } = req.body;
    if (!scopeDescription) return res.status(400).json({ error: "scopeDescription is required" });

    const projects = storage.getCompletedProjects();
    if (projects.length === 0) return res.json({ matches: [], estimate: null });

    // Score each project by keyword overlap
    const queryWords = scopeDescription.toLowerCase().split(/\s+/).filter((w: string) => w.length > 2);
    const tagWords = (tags || "").toLowerCase().split(/[,\s]+/).filter((w: string) => w.length > 2);
    const allQueryWords = [...queryWords, ...tagWords];

    const scored = projects.map(p => {
      const scopeWords = (p.scopeDescription || "").toLowerCase().split(/\s+/);
      const tagList = (p.tags || "").toLowerCase().split(/[,\s]+/);
      const nameWords = (p.name || "").toLowerCase().split(/\s+/);
      const allWords = [...scopeWords, ...tagList, ...nameWords];

      let score = 0;
      for (const qw of allQueryWords) {
        for (const pw of allWords) {
          if (pw.includes(qw) || qw.includes(pw)) { score += 1; break; }
        }
      }
      return { project: p, score };
    }).filter(s => s.score > 0).sort((a, b) => b.score - a.score).slice(0, 5);

    if (scored.length === 0) return res.json({ matches: [], estimate: null });

    // Compute ranges from matching projects
    const matches = scored.map(s => s.project);
    const totalManhours = matches.map(p => p.totalManhours || 0).filter(v => v > 0);
    const materialCosts = matches.map(p => p.materialCost || 0).filter(v => v > 0);
    const laborCosts = matches.map(p => p.laborCost || 0).filter(v => v > 0);
    const totalCosts = matches.map(p => p.totalCost || 0).filter(v => v > 0);
    const durations = matches.map(p => p.durationDays || 0).filter(v => v > 0);
    const crewSizes = matches.map(p => p.peakCrewSize || 0).filter(v => v > 0);

    const rangeOf = (arr: number[]) => arr.length === 0 ? null : {
      min: Math.min(...arr), max: Math.max(...arr),
      avg: Math.round(arr.reduce((s, v) => s + v, 0) / arr.length),
    };

    res.json({
      matches: scored.map(s => ({ ...s.project, matchScore: s.score })),
      estimate: {
        manhours: rangeOf(totalManhours),
        materialCost: rangeOf(materialCosts),
        laborCost: rangeOf(laborCosts),
        totalCost: rangeOf(totalCosts),
        duration: rangeOf(durations),
        crewSize: rangeOf(crewSizes),
        basedOn: matches.length,
      },
    });
  });

  // ===== FEATURE 2: Bid Tracking CRUD =====
  app.get("/api/bids/stats", (_req, res) => {
    const stats = storage.getBidStats();
    res.json(stats);
  });

  app.get("/api/bids", (_req, res) => {
    const bids = storage.getBids();
    res.json(bids);
  });

  app.post("/api/bids", (req, res) => {
    const parsed = insertBidSchema.safeParse(req.body);
    if (!parsed.success) return res.status(400).json({ error: "Invalid data", details: parsed.error.flatten().fieldErrors });
    const bid = storage.createBid(parsed.data);
    res.status(201).json(bid);
  });

  app.patch("/api/bids/:id", (req, res) => {
    const parsed = insertBidSchema.partial().safeParse(req.body);
    if (!parsed.success) return res.status(400).json({ error: "Invalid data", details: parsed.error.flatten().fieldErrors });
    const updated = storage.updateBid(req.params.id, parsed.data);
    if (!updated) return res.status(404).json({ error: "Bid not found" });
    res.json(updated);
  });

  app.delete("/api/bids/:id", (req, res) => {
    const deleted = storage.deleteBid(req.params.id);
    if (!deleted) return res.status(404).json({ error: "Bid not found" });
    res.json({ deleted: true });
  });


  // ===== CALIBRATION PROFILES =====

  app.get("/api/calibration-profiles", (_req, res) => {
    const profiles = Object.entries(CALIBRATION_DATA).map(([key, data]) => ({
      id: key,
      project: data.project,
      benchmark: data.benchmark,
      ssWeldFactors: data.ss_weld_factors,
      boltMethodology: data.bolt_methodology,
      smallBoreThreshold: data.small_bore_threshold,
      pipeScopeNote: data.pipe_scope_note,
    }));
    res.json(profiles);
  });

  // ===== CHAT ASSISTANT =====

  app.post("/api/chat-assistant", async (req, res) => {
    const { message, context } = req.body || {};
    if (!message || typeof message !== "string") {
      return res.status(400).json({ error: "Message required" });
    }

    const geminiKey = getUserGeminiKey();
    if (!geminiKey) {
      return res.json({
        response: "AI chat requires a Gemini API key. Go to Settings to configure one.",
        mode: "error",
      });
    }

    try {
      const systemPrompt = `You are PG Assistant, the built-in helper for the Picou Group Estimator — an AI-powered takeoff and estimating tool for industrial piping contractors.

You help estimators with:
- How to use the application features
- Piping estimation questions (manhour factors, weld counts, material specs)
- Cost database queries
- Crew planning advice
- General industrial piping knowledge

Key facts about the app:
- Mechanical, Structural, and Civil takeoff from PDF drawings
- Estimating with Bill's Engineering Index or Justin's IPMH method
- Labor rates: ST=$56/hr, OT=$79/hr, DT=$100/hr, Per Diem=$75/day
- Rack factor: 1.3x default
- Cost database with 232+ records
- Stolthaven Phase 6 calibration: IPMH 0.437, within 3.8% of actual manhours
- SS weld factors: 3" = 4.68 MH, 4" = 5.56 MH (calibrated)
- Crew planning with role-specific rates from Phase 6 data

The user is currently on: ${context || "unknown page"}

Be concise, practical, and helpful. Answer in 2-4 sentences when possible. Use specific numbers when relevant.`;

      const resp = await fetch(
        `https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent`,
        {
          method: "POST",
          headers: { "Content-Type": "application/json", "x-goog-api-key": geminiKey },
          body: JSON.stringify({
            contents: [{ parts: [{ text: systemPrompt + "\n\nUser question: " + message }] }],
            generationConfig: { temperature: 0.7, maxOutputTokens: 500 },
          }),
        }
      );

      if (!resp.ok) throw new Error(`Gemini API error: ${resp.status}`);
      const data = await resp.json();
      const text = data?.candidates?.[0]?.content?.parts?.[0]?.text || "Sorry, I couldn't process that.";
      res.json({ response: text, mode: "ai" });
    } catch (err: any) {
      console.error("Chat assistant error:", err.message);
      res.json({ response: "Sorry, I couldn't process that. Please try again.", mode: "error" });
    }
  });

  // ── Inline Edit / Verify Takeoff Items ──

  app.patch("/api/takeoff-items/:itemId", (req, res) => {
    const { itemId } = req.params;
    const updates: Record<string, any> = {};
    const body = req.body || {};

    // Fetch old item for correction tracking
    const oldItem = storage.getTakeoffItemById(itemId);

    const allowedFields = ["size", "quantity", "description", "category", "unit", "material", "schedule", "spec", "rating", "notes"];
    for (const field of allowedFields) {
      if (body[field] !== undefined) {
        updates[field] = field === "quantity" ? (parseFloat(body[field]) || 0) : String(body[field]);
      }
    }

    // Boolean scope flags + revisionClouded — update if the body includes them.
    for (const flag of ["includeInBom", "includeInTakeoff", "includeInEstimate", "revisionClouded"]) {
      if (body[flag] !== undefined) updates[flag] = body[flag] ? 1 : 0;
    }

    if (body.verified === true) {
      updates.confidence = "high";
      updates.confidenceScore = 95;
      updates.confidenceNotes = "Manually verified by estimator";
      updates.manuallyVerified = 1;
    }

    if (Object.keys(updates).length === 0) {
      return res.status(400).json({ error: "No updates provided" });
    }

    const success = storage.updateTakeoffItem(itemId, updates);
    if (!success) return res.status(404).json({ error: "Item not found" });

    // Record correction patterns for historical learning
    if (oldItem) {
      if (updates.size && oldItem.size !== updates.size) {
        storage.recordCorrection(oldItem.size, updates.size, "size");
      }
      if (updates.quantity !== undefined && String(oldItem.quantity) !== String(updates.quantity)) {
        storage.recordCorrection(String(oldItem.quantity), String(updates.quantity), "quantity");
      }
      if (updates.description && oldItem.description !== updates.description) {
        storage.recordCorrection(oldItem.description, updates.description, "description");
      }
    }

    res.json({ success: true });
  });

  // ── Manual Add / Delete Takeoff Items ──
  // For when the PDF extraction missed something the estimator knows belongs
  // on the takeoff. New rows get lineNumber 'M-001', 'M-002', etc. so they
  // sort to the end and are visually distinct from extracted lines.
  app.post("/api/takeoff-projects/:projectId/items", (req, res) => {
    const { projectId } = req.params;
    const body = req.body || {};
    if (!body.description || typeof body.description !== "string") {
      return res.status(400).json({ error: "description (string) is required" });
    }
    const id = storage.addTakeoffItem(projectId, {
      description: body.description,
      category: body.category,
      size: body.size,
      quantity: typeof body.quantity === "number" ? body.quantity : parseFloat(body.quantity) || 1,
      unit: body.unit,
      material: body.material,
      schedule: body.schedule,
      spec: body.spec,
      rating: body.rating,
      notes: body.notes,
      lineNumber: body.lineNumber,
      revisionClouded: !!body.revisionClouded,
    });
    if (!id) return res.status(404).json({ error: "Project not found" });
    res.json({ success: true, id });
  });

  app.delete("/api/takeoff-items/:itemId", (req, res) => {
    const { itemId } = req.params;
    const ok = storage.deleteTakeoffItem(itemId);
    if (!ok) return res.status(404).json({ error: "Item not found" });
    res.json({ success: true });
  });

  // ── Correction Patterns (Historical Learning) ──

  app.get("/api/correction-patterns", (_req, res) => {
    const patterns = storage.getAllCorrectionPatterns();
    res.json(patterns);
  });

  app.patch("/api/correction-patterns/:id", (req, res) => {
    const { id } = req.params;
    const { autoApply } = req.body || {};
    if (typeof autoApply !== "boolean") {
      return res.status(400).json({ error: "autoApply (boolean) required" });
    }
    storage.setCorrectionPatternAutoApply(id, autoApply);
    res.json({ success: true });
  });

  app.delete("/api/correction-patterns/:id", (req, res) => {
    const { id } = req.params;
    storage.deleteCorrectionPattern(id);
    res.json({ success: true });
  });

  // Emergency: clear ALL correction patterns. Used to reset the pattern
  // database when a bad pattern is causing extraction-wide issues.
  app.delete("/api/correction-patterns", (_req, res) => {
    const count = storage.deleteAllCorrectionPatterns();
    res.json({ success: true, deleted: count });
  });

  // ── Project Folders ──

  app.post("/api/folders", (req, res) => {
    const { name, description, color } = req.body;
    if (!name || typeof name !== "string" || !name.trim()) {
      return res.status(400).json({ error: "Folder name is required" });
    }
    const folder = storage.createFolder({ name: name.trim(), description, color });
    res.json(folder);
  });

  app.get("/api/folders", (_req, res) => {
    res.json(storage.getFolders());
  });

  app.get("/api/folders/:id", (req, res) => {
    const folder = storage.getFolder(req.params.id);
    if (!folder) return res.status(404).json({ error: "Folder not found" });
    res.json(folder);
  });

  app.put("/api/folders/:id", (req, res) => {
    const { name, description, color } = req.body;
    const updated = storage.updateFolder(req.params.id, { name, description, color });
    if (!updated) return res.status(404).json({ error: "Folder not found" });
    res.json(updated);
  });

  app.delete("/api/folders/:id", (req, res) => {
    const deleted = storage.deleteFolder(req.params.id);
    if (!deleted) return res.status(404).json({ error: "Folder not found" });
    res.json({ success: true });
  });

  app.post("/api/folders/:id/projects", (req, res) => {
    const { projectId } = req.body;
    if (!projectId) return res.status(400).json({ error: "projectId is required" });
    const folder = storage.getFolder(req.params.id);
    if (!folder) return res.status(404).json({ error: "Folder not found" });
    storage.addProjectToFolder(req.params.id, projectId);
    res.json({ success: true });
  });

  app.delete("/api/folders/:id/projects/:projectId", (req, res) => {
    storage.removeProjectFromFolder(req.params.projectId);
    res.json({ success: true });
  });

  app.get("/api/folders/:id/combined-bom", (req, res) => {
    const folder = storage.getFolder(req.params.id);
    if (!folder) return res.status(404).json({ error: "Folder not found" });
    const projects = storage.getProjectsByFolder(req.params.id);
    const combinedItems = storage.getFolderCombinedItems(req.params.id);
    res.json({ folder, projects, combinedItems });
  });


  // ===== EXCEL EXPORT ENDPOINTS =====

  // Shared Excel styling helper
  function applyExcelStyles(ws: any) {
    const headerFill = { type: "pattern" as const, pattern: "solid" as const, fgColor: { argb: "FF01696F" } };
    const headerFont = { bold: true, color: { argb: "FFFFFFFF" }, size: 11 };
    const headerAlignment = { horizontal: "center" as const, vertical: "middle" as const };
    ws.getRow(1).eachCell((cell: any) => {
      cell.fill = headerFill;
      cell.font = headerFont;
      cell.alignment = headerAlignment;
    });
    ws.columns.forEach((col: any) => {
      let maxLen = 10;
      col.eachCell({ includeEmpty: false }, (cell: any) => {
        const len = cell.value ? String(cell.value).length : 0;
        if (len > maxLen) maxLen = len;
      });
      col.width = Math.min(maxLen + 2, 40);
    });
    for (let r = 2; r <= ws.rowCount; r++) {
      if (r % 2 === 0) {
        ws.getRow(r).eachCell((cell: any) => {
          cell.fill = { type: "pattern" as const, pattern: "solid" as const, fgColor: { argb: "FFF2F2F2" } };
        });
      }
    }
  }

  // 1. Takeoff BOM Export
  app.get("/api/takeoff-projects/:id/export-bom", async (req, res) => {
    const project = await storage.getTakeoffProject(req.params.id);
    if (!project) return res.status(404).json({ message: "Project not found" });
    try {
      const ExcelJS = await import("exceljs");
      const wb = new ExcelJS.default.Workbook();

      // Sheet 1: BOM
      const ws = wb.addWorksheet("BOM");
      ws.columns = [
        { header: "Line #", key: "lineNumber" },
        { header: "Category", key: "category" },
        { header: "Size", key: "size" },
        { header: "Description", key: "description" },
        { header: "Quantity", key: "quantity" },
        { header: "Unit", key: "unit" },
        { header: "Material", key: "material" },
        { header: "Schedule", key: "schedule" },
        { header: "Spec", key: "spec" },
        { header: "Rating", key: "rating" },
        { header: "Confidence", key: "confidence" },
        { header: "Source Page", key: "sourcePage" },
        { header: "Notes", key: "notes" },
      ];
      for (const item of project.items) {
        ws.addRow({
          lineNumber: item.lineNumber || "",
          category: item.category || "",
          size: item.size || "",
          description: item.description || "",
          quantity: item.quantity ?? 0,
          unit: item.unit || "EA",
          material: item.material || "",
          schedule: item.schedule || "",
          spec: item.spec || "",
          rating: item.rating || "",
          confidence: item.confidence ? `${Math.round(item.confidence * 100)}%` : "",
          sourcePage: item.sourcePage ?? "",
          notes: item.notes || "",
        });
      }
      applyExcelStyles(ws);

      // Sheet 2: Summary pivot
      const ws2 = wb.addWorksheet("Summary");
      ws2.columns = [
        { header: "Category", key: "category" },
        { header: "Size", key: "size" },
        { header: "Item Count", key: "count" },
        { header: "Total Qty", key: "totalQty" },
      ];
      const pivot = new Map<string, { count: number; totalQty: number }>();
      for (const item of project.items) {
        const key = `${item.category || "other"}|${item.size || "N/A"}`;
        const existing = pivot.get(key) || { count: 0, totalQty: 0 };
        existing.count++;
        existing.totalQty += item.quantity ?? 0;
        pivot.set(key, existing);
      }
      for (const [key, val] of pivot) {
        const [category, size] = key.split("|");
        ws2.addRow({ category, size, count: val.count, totalQty: val.totalQty });
      }
      // Totals row
      ws2.addRow({ category: "TOTAL", size: "", count: project.items.length, totalQty: project.items.reduce((s, i) => s + (i.quantity ?? 0), 0) });
      const totalRow = ws2.getRow(ws2.rowCount);
      totalRow.font = { bold: true };
      applyExcelStyles(ws2);

      const safeName = project.name.replace(/[^a-zA-Z0-9 _-]/g, "");
      res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
      res.setHeader("Content-Disposition", `attachment; filename="${safeName} - BOM.xlsx"`);
      await wb.xlsx.write(res);
      res.end();
    } catch (err: any) {
      console.error("BOM export error:", err);
      res.status(500).json({ message: err.message || "Export failed" });
    }
  });

  // Re-run dedup on an existing project (no re-extraction needed).
  // Useful after dedup logic improvements: apply the new rules to projects
  // that were extracted before the fix.
  app.post("/api/takeoff-projects/:id/redup", async (req, res) => {
    const project = await storage.getTakeoffProject(req.params.id);
    if (!project) return res.status(404).json({ message: "Project not found" });
    try {
      // Reset all dedup flags first so we start clean
      for (const item of project.items) {
        (item as any)._dedupCandidate = false;
        (item as any).dedupNote = undefined;
      }
      // Run the same-drawing dedup
      const sameResult = dedupSameDrawingNumber(project.items);
      // Run the continuation dedup
      dedupContinuationPages(project.items);
      // Persist all items so the flags stick
      for (const item of project.items) {
        await storage.updateTakeoffItem(item.id, item);
      }
      const dedupTotal = project.items.filter(i => (i as any)._dedupCandidate).length;
      res.json({
        success: true,
        dedupCandidates: dedupTotal,
        sameDrawingDuplicateGroups: sameResult.dupGroups,
        sameDrawingItemsMarked: sameResult.dedupCount,
        message: `Marked ${dedupTotal} items as dedup candidates. ${sameResult.dupGroups} duplicate-drawing group(s) detected.`,
      });
    } catch (err: any) {
      console.error("Re-dedup error:", err);
      res.status(500).json({ message: err.message || "Re-dedup failed" });
    }
  });

  // 2. Connections Export
  app.get("/api/takeoff-projects/:id/export-connections", async (req, res) => {
    const project = await storage.getTakeoffProject(req.params.id);
    if (!project) return res.status(404).json({ message: "Project not found" });
    try {
      const ExcelJS = await import("exceljs");
      const wb = new ExcelJS.default.Workbook();

      // Infer connections from items.
      // We track each connection in BOTH a shop bucket and a field bucket so the
      // estimator can compare directly against fab-shop pivots (which exclude the
      // field welds added for 40-foot pipe-run joints).
      const connectionDetails: Array<{ size: string; fitting: string; qty: number; connectionType: string; connectionCount: number; location: string }> = [];
      type ConnBucket = { buttWelds: number; socketWelds: number; boltUps: number; threaded: number };
      const makeBucket = (): ConnBucket => ({ buttWelds: 0, socketWelds: 0, boltUps: 0, threaded: 0 });
      const sizeMap = new Map<string, { shop: ConnBucket; field: ConnBucket }>();
      const ensureSize = (size: string) => {
        if (!sizeMap.has(size)) sizeMap.set(size, { shop: makeBucket(), field: makeBucket() });
        return sizeMap.get(size)!;
      };

      for (const item of project.items) {
        // Skip dedup candidates (same-drawing-number duplicates, continuation duplicates).
        // These items are kept in the BOM (UI shows them at 50% opacity) but must not
        // be counted in the connections totals.
        if ((item as any)._dedupCandidate) continue;
        const catLower = (item.category || "").toLowerCase();
        const descLower = (item.description || "").toLowerCase();
        const size = item.size || "N/A";
        const qty = item.quantity ?? 0;
        const hasNpt = descLower.includes("threaded") || descLower.includes("screw") || descLower.includes("npt") || descLower.includes("fnpt") || descLower.includes("mnpt");
        const hasSw = descLower.includes("socket weld") || descLower.includes("socket") || /\bsw\b/i.test(descLower) || descLower.includes(",sw,") || descLower.includes(", sw,") || descLower.includes(" sw,");
        // Mixed-end fittings (e.g. "SW x NPT" valve, "SW x THD" coupling) have
        // ONE socket weld end and one threaded end. Detect them and count 1 SW.
        const isMixedSwNpt = hasSw && hasNpt && (
          /sw\s*[xX/-]\s*(npt|thd|threaded|screw)/i.test(item.description || "")
          || /(npt|thd|threaded)\s*[xX/-]\s*sw/i.test(item.description || "")
        );
        // True threaded (both ends or one end without SW): only treat as threaded
        // if there is NO SW indication at all, OR the description explicitly says
        // "threaded x threaded" / has no SW pattern. Mixed SW x NPT is handled separately.
        const isThreaded = hasNpt && !hasSw;
        // True socket weld: SW on both ends, no NPT mention. If both SW and NPT,
        // we treat as mixed (handled above).
        const isSocketWeld = hasSw && !isMixedSwNpt;
        const isFlanged = descLower.includes("flanged") || descLower.includes("flg") || descLower.includes("rf ") || descLower.includes("raised face");
        const location = (item.installLocation || "shop").toLowerCase();
        const isField = location === "field";
        const sm = isField ? ensureSize(size).field : ensureSize(size).shop;

        if (catLower === "elbow" || catLower === "ell") {
          // Mixed SW x NPT elbows are rare but possible — 1 SW + 0 threaded weld
          if (isMixedSwNpt) {
            connectionDetails.push({ size, fitting: item.description || catLower, qty, connectionType: "sw x npt mixed", connectionCount: 1 * qty, location });
            sm.socketWelds += 1 * qty;
            sm.threaded += 1 * qty;
          } else {
            const conns = isThreaded ? 0 : 2;
            const connType = isThreaded ? "threaded" : isSocketWeld ? "socket weld" : "butt weld";
            connectionDetails.push({ size, fitting: item.description || catLower, qty, connectionType: connType, connectionCount: conns * qty, location });
            if (isThreaded) sm.threaded += conns * qty;
            else if (isSocketWeld) sm.socketWelds += conns * qty;
            else sm.buttWelds += conns * qty;
          }
        } else if (catLower === "tee") {
          // Reducing tees (e.g. 8"x6" or 8"X6") have TWO welds at the larger
          // (run) size and ONE weld at the smaller (branch) size, NOT 3 at the
          // larger size. Equal-size tees have 3 welds at the single size.
          const reducerSizeMatch = (size || "").match(/^(\S+?)\s*[xX]\s*(\S+)$/);
          const connType = isThreaded ? "threaded" : isSocketWeld ? "socket weld" : "butt weld";
          if (reducerSizeMatch && !isThreaded) {
            const largerSize = reducerSizeMatch[1].trim();
            const smallerSize = reducerSizeMatch[2].trim();
            const largerBucket = isField ? ensureSize(largerSize).field : ensureSize(largerSize).shop;
            const smallerBucket = isField ? ensureSize(smallerSize).field : ensureSize(smallerSize).shop;
            // 2 run welds at larger size
            if (isSocketWeld) largerBucket.socketWelds += 2 * qty;
            else largerBucket.buttWelds += 2 * qty;
            // 1 branch weld at smaller size
            if (isSocketWeld) smallerBucket.socketWelds += 1 * qty;
            else smallerBucket.buttWelds += 1 * qty;
            connectionDetails.push({ size: largerSize, fitting: item.description || catLower, qty, connectionType: `${connType} (run)`, connectionCount: 2 * qty, location });
            connectionDetails.push({ size: smallerSize, fitting: item.description || catLower, qty, connectionType: `${connType} (branch)`, connectionCount: 1 * qty, location });
          } else {
            const conns = isThreaded ? 0 : 3;
            connectionDetails.push({ size, fitting: item.description || catLower, qty, connectionType: connType, connectionCount: conns * qty, location });
            if (isThreaded) sm.threaded += conns * qty;
            else if (isSocketWeld) sm.socketWelds += conns * qty;
            else sm.buttWelds += conns * qty;
          }
        } else if (catLower === "reducer" || catLower === "reducing") {
          // Reducers have 1 weld at EACH end (one at the larger size, one at the
          // smaller size). Previously we were double-counting both at the larger
          // size, which inflated 6"/8" BW totals.
          const reducerSizeMatch = (size || "").match(/^(\S+?)\s*[xX]\s*(\S+)$/);
          const connType = isSocketWeld ? "socket weld" : "butt weld";
          if (reducerSizeMatch) {
            const largerSize = reducerSizeMatch[1].trim();
            const smallerSize = reducerSizeMatch[2].trim();
            const largerBucket = isField ? ensureSize(largerSize).field : ensureSize(largerSize).shop;
            const smallerBucket = isField ? ensureSize(smallerSize).field : ensureSize(smallerSize).shop;
            if (isSocketWeld) { largerBucket.socketWelds += qty; smallerBucket.socketWelds += qty; }
            else { largerBucket.buttWelds += qty; smallerBucket.buttWelds += qty; }
            connectionDetails.push({ size: largerSize, fitting: item.description || catLower, qty, connectionType: `${connType} (large end)`, connectionCount: qty, location });
            connectionDetails.push({ size: smallerSize, fitting: item.description || catLower, qty, connectionType: `${connType} (small end)`, connectionCount: qty, location });
          } else {
            // Single-size reducer string: fall back to old behavior (2 welds at this size)
            connectionDetails.push({ size, fitting: item.description || catLower, qty, connectionType: connType, connectionCount: 2 * qty, location });
            if (isSocketWeld) sm.socketWelds += 2 * qty;
            else sm.buttWelds += 2 * qty;
          }
        } else if (catLower === "cap") {
          connectionDetails.push({ size, fitting: item.description || catLower, qty, connectionType: "butt weld", connectionCount: 1 * qty, location });
          sm.buttWelds += 1 * qty;
        } else if (descLower.includes("sockolet")) {
          // Sockolet: 2 socket welds (header bore + branch pipe)
          connectionDetails.push({ size, fitting: item.description || "Sockolet", qty, connectionType: "socket weld", connectionCount: 2 * qty, location });
          sm.socketWelds += 2 * qty;
        } else if (descLower.includes("weldolet")) {
          // Weldolet: 1 butt weld to header
          connectionDetails.push({ size, fitting: item.description || "Weldolet", qty, connectionType: "butt weld", connectionCount: 1 * qty, location });
          sm.buttWelds += 1 * qty;
        } else if (descLower.includes("threadolet")) {
          // Threadolet: 1 weld to header + threaded branch (only count the weld)
          connectionDetails.push({ size, fitting: item.description || "Threadolet", qty, connectionType: "butt weld", connectionCount: 1 * qty, location });
          sm.buttWelds += 1 * qty;
        } else if (descLower.includes("olet")) {
          // Generic olet (without sockolet/weldolet/threadolet keyword): treat as olet weld
          connectionDetails.push({ size, fitting: item.description || "Olet", qty, connectionType: "butt weld", connectionCount: 1 * qty, location });
          sm.buttWelds += 1 * qty;
        } else if (catLower === "coupling") {
          // Coupling: 2 SW unless threaded. SW x NPT mixed: 1 SW + 1 threaded.
          if (isMixedSwNpt) {
            connectionDetails.push({ size, fitting: item.description || catLower, qty, connectionType: "sw x npt mixed", connectionCount: 1 * qty, location });
            sm.socketWelds += 1 * qty;
            sm.threaded += 1 * qty;
          } else if (isThreaded) {
            connectionDetails.push({ size, fitting: item.description || catLower, qty, connectionType: "threaded", connectionCount: 2 * qty, location });
            sm.threaded += 2 * qty;
          } else {
            connectionDetails.push({ size, fitting: item.description || catLower, qty, connectionType: "socket weld", connectionCount: 2 * qty, location });
            sm.socketWelds += 2 * qty;
          }
        } else if (catLower === "valve") {
          // Only socket-weld valves generate welds. Flanged/threaded/butterfly: 0 welds.
          // SW x NPT mixed valves have 1 SW end (1 weld) + 1 threaded end (0 welds).
          if (isMixedSwNpt) {
            connectionDetails.push({ size, fitting: item.description || catLower, qty, connectionType: "sw x npt mixed", connectionCount: 1 * qty, location });
            sm.socketWelds += 1 * qty;
            sm.threaded += 1 * qty;
          } else if (isThreaded) {
            connectionDetails.push({ size, fitting: item.description || catLower, qty, connectionType: "threaded", connectionCount: 2 * qty, location });
            sm.threaded += 2 * qty;
          } else if (isSocketWeld) {
            connectionDetails.push({ size, fitting: item.description || catLower, qty, connectionType: "socket weld", connectionCount: 2 * qty, location });
            sm.socketWelds += 2 * qty;
          }
          // Flanged, butterfly, BW valves: 0 welds (connections come from surrounding flanges)
        } else if (catLower === "flange") {
          // SW flanges get socket welds, WN/SO flanges get butt welds
          const flangeWeldType = isSocketWeld ? "socket weld" : "butt weld";
          connectionDetails.push({ size, fitting: item.description || catLower, qty, connectionType: `${flangeWeldType} + bolt-up`, connectionCount: qty, location });
          if (isSocketWeld) sm.socketWelds += qty;
          else sm.buttWelds += qty;
          sm.boltUps += Math.ceil(qty / 2);
        }
      }

      // Add the 40-foot pipe-run joint welds. These are field welds the app
      // infers from each pipe length >40 LF. They are NOT in the fab-shop pivot
      // and must be reported separately so the estimator can compare apples to
      // apples against shop counts.
      const inferredFieldWelds: Array<{ size: string; lengthLF: number; welds: number; pipeDescription: string }> = [];
      for (const item of project.items) {
        if ((item as any)._dedupCandidate) continue;
        const cat = (item.category || "").toLowerCase();
        if (cat !== "pipe") continue;
        const lengthLF = item.quantity || 0;
        if (lengthLF < 40) continue;
        const pipeJointWelds = Math.floor(lengthLF / 40);
        if (pipeJointWelds === 0) continue;
        const size = item.size || "N/A";
        const isPipeSocketWeld = lengthLF > 0 && parseFloat(String(size).replace(/[^0-9.]/g, "")) <= 1.5;
        const fieldBucket = ensureSize(size).field;
        if (isPipeSocketWeld) fieldBucket.socketWelds += pipeJointWelds;
        else fieldBucket.buttWelds += pipeJointWelds;
        inferredFieldWelds.push({
          size,
          lengthLF,
          welds: pipeJointWelds,
          pipeDescription: item.description || "PIPE",
        });
        connectionDetails.push({
          size,
          fitting: `[40' rule] ${item.description || "PIPE"}`,
          qty: 1,
          connectionType: isPipeSocketWeld ? "field socket weld (pipe joint)" : "field butt weld (pipe joint)",
          connectionCount: pipeJointWelds,
          location: "field",
        });
      }

      // Sheet 1: Connections by Size — split into shop vs field columns so the
      // estimator can compare directly against the fab-shop pivot.
      const ws = wb.addWorksheet("Connections by Size");
      ws.columns = [
        { header: "Size", key: "size", width: 10 },
        { header: "Shop BW", key: "shopBW", width: 10 },
        { header: "Shop SW", key: "shopSW", width: 10 },
        { header: "Shop Bolt-Ups", key: "shopBU", width: 14 },
        { header: "Shop Threaded", key: "shopTH", width: 14 },
        { header: "Shop Total", key: "shopTotal", width: 12 },
        { header: "Field BW (40' rule)", key: "fieldBW", width: 18 },
        { header: "Field SW (40' rule)", key: "fieldSW", width: 18 },
        { header: "Field Total", key: "fieldTotal", width: 12 },
        { header: "Grand Total", key: "grandTotal", width: 12 },
      ];
      let totShopBW = 0, totShopSW = 0, totShopBU = 0, totShopTH = 0;
      let totFieldBW = 0, totFieldSW = 0;
      for (const [size, buckets] of sizeMap) {
        const s = buckets.shop;
        const f = buckets.field;
        const shopTotal = s.buttWelds + s.socketWelds + s.boltUps + s.threaded;
        const fieldTotal = f.buttWelds + f.socketWelds;
        ws.addRow({
          size,
          shopBW: s.buttWelds,
          shopSW: s.socketWelds,
          shopBU: s.boltUps,
          shopTH: s.threaded,
          shopTotal,
          fieldBW: f.buttWelds,
          fieldSW: f.socketWelds,
          fieldTotal,
          grandTotal: shopTotal + fieldTotal,
        });
        totShopBW += s.buttWelds; totShopSW += s.socketWelds; totShopBU += s.boltUps; totShopTH += s.threaded;
        totFieldBW += f.buttWelds; totFieldSW += f.socketWelds;
      }
      const shopGrand = totShopBW + totShopSW + totShopBU + totShopTH;
      const fieldGrand = totFieldBW + totFieldSW;
      ws.addRow({
        size: "TOTAL",
        shopBW: totShopBW, shopSW: totShopSW, shopBU: totShopBU, shopTH: totShopTH,
        shopTotal: shopGrand,
        fieldBW: totFieldBW, fieldSW: totFieldSW,
        fieldTotal: fieldGrand,
        grandTotal: shopGrand + fieldGrand,
      });
      ws.getRow(ws.rowCount).font = { bold: true };
      // Add a small legend / explanation block below the table
      ws.addRow({});
      ws.addRow({ size: "Shop Welds", shopBW: "= fittings, flanges, valves, olets (compare to fab-shop pivot)" });
      ws.addRow({ size: "Field Welds", shopBW: "= 40-ft pipe-run joints inferred from pipe LF (NOT in fab-shop pivot)" });
      ws.addRow({ size: "Rule", shopBW: "floor(pipe_LF / 40) field welds added per pipe item; size <=1.5\" => SW, else BW" });
      applyExcelStyles(ws);

      // Sheet 2: Connection Detail
      const ws2 = wb.addWorksheet("Connection Detail");
      ws2.columns = [
        { header: "Size", key: "size" },
        { header: "Fitting", key: "fitting" },
        { header: "Qty", key: "qty" },
        { header: "Connection Type", key: "connectionType" },
        { header: "Connection Count", key: "connectionCount" },
        { header: "Location", key: "location" },
      ];
      for (const d of connectionDetails) ws2.addRow(d);
      applyExcelStyles(ws2);

      // Sheet 3: 40-Foot Field Weld Detail — itemised list of every field weld
      // the app added for pipe joints, so the estimator can audit them.
      const ws3 = wb.addWorksheet("40' Field Welds");
      ws3.columns = [
        { header: "Size", key: "size", width: 10 },
        { header: "Pipe Description", key: "pipeDescription", width: 50 },
        { header: "Length (LF)", key: "lengthLF", width: 14 },
        { header: "Field Welds Added", key: "welds", width: 18 },
        { header: "Weld Type", key: "weldType", width: 12 },
      ];
      let totFieldRowWelds = 0;
      for (const w of inferredFieldWelds) {
        const isSW = parseFloat(String(w.size).replace(/[^0-9.]/g, "")) <= 1.5;
        ws3.addRow({
          size: w.size,
          pipeDescription: w.pipeDescription,
          lengthLF: w.lengthLF,
          welds: w.welds,
          weldType: isSW ? "SW" : "BW",
        });
        totFieldRowWelds += w.welds;
      }
      ws3.addRow({ size: "TOTAL", pipeDescription: "", lengthLF: "", welds: totFieldRowWelds, weldType: "" });
      ws3.getRow(ws3.rowCount).font = { bold: true };
      ws3.addRow({});
      ws3.addRow({ size: "Note", pipeDescription: "Rule: every 40 LF of pipe run requires 1 field weld at the joint between standard 40' lengths." });
      ws3.addRow({ size: "", pipeDescription: "These welds are NOT in fab-shop weld counts — they are installed in the field during erection." });
      applyExcelStyles(ws3);

      const safeName = project.name.replace(/[^a-zA-Z0-9 _-]/g, "");
      res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
      res.setHeader("Content-Disposition", `attachment; filename="${safeName} - Connections.xlsx"`);
      await wb.xlsx.write(res);
      res.end();
    } catch (err: any) {
      console.error("Connections export error:", err);
      res.status(500).json({ message: err.message || "Export failed" });
    }
  });

  // 3. Crew Plan Export
  app.post("/api/export-crew-plan", async (req, res) => {
    const { projectName, scenarios, customCrew, rateCard } = req.body;
    if (!projectName) return res.status(400).json({ message: "projectName required" });
    try {
      const ExcelJS = await import("exceljs");
      const wb = new ExcelJS.default.Workbook();

      // Sheet 1: Scenarios
      if (scenarios && Array.isArray(scenarios)) {
        const ws = wb.addWorksheet("Scenarios");
        ws.columns = [
          { header: "Scenario", key: "scenario", width: 18 },
          { header: "Role", key: "role", width: 25 },
          { header: "Count", key: "count", width: 10 },
          { header: "Rate ($/hr)", key: "rate", width: 14 },
          { header: "Per Diem ($/day)", key: "perDiem", width: 16 },
          { header: "Duration (days)", key: "duration", width: 16 },
          { header: "Daily Burn Rate ($)", key: "dailyBurn", width: 18 },
          { header: "Total Cost ($)", key: "totalCost", width: 16 },
        ];
        for (const scenario of scenarios) {
          if (scenario.roles && Array.isArray(scenario.roles)) {
            for (const role of scenario.roles) {
              ws.addRow({
                scenario: scenario.name || "",
                role: role.name || "",
                count: role.count ?? 0,
                rate: role.rate ?? 0,
                perDiem: role.perDiem ?? 0,
                duration: scenario.duration ?? 0,
                dailyBurn: scenario.dailyBurn ?? 0,
                totalCost: scenario.totalCost ?? 0,
              });
            }
            // Blank row between scenarios
            ws.addRow({});
          }
        }
        applyExcelStyles(ws);
      }

      // Sheet 2: Custom Crew
      if (customCrew && typeof customCrew === "object") {
        const ws2 = wb.addWorksheet("Custom Crew");
        ws2.columns = [
          { header: "Role", key: "role", width: 25 },
          { header: "Count", key: "count", width: 10 },
          { header: "Rate ($/hr)", key: "rate", width: 14 },
          { header: "Per Diem ($/day)", key: "perDiem", width: 16 },
        ];
        if (customCrew.roles && Array.isArray(customCrew.roles)) {
          for (const role of customCrew.roles) {
            ws2.addRow({ role: role.name || "", count: role.count ?? 0, rate: role.rate ?? 0, perDiem: role.perDiem ?? 0 });
          }
        }
        applyExcelStyles(ws2);
      }

      // Sheet 3: Rate Card
      if (rateCard && Array.isArray(rateCard)) {
        const ws3 = wb.addWorksheet("Rate Card");
        ws3.columns = [
          { header: "Role", key: "role", width: 25 },
          { header: "Rate ($/hr)", key: "rate", width: 14 },
          { header: "Per Diem Eligible", key: "perDiem", width: 18 },
        ];
        for (const r of rateCard) {
          ws3.addRow({ role: r.name || "", rate: r.rate ?? 0, perDiem: r.perDiemEligible ? "Yes" : "No" });
        }
        applyExcelStyles(ws3);
      }

      const safeName = (projectName || "Crew Plan").replace(/[^a-zA-Z0-9 _-]/g, "");
      res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
      res.setHeader("Content-Disposition", `attachment; filename="${safeName} - Crew Plan.xlsx"`);
      await wb.xlsx.write(res);
      res.end();
    } catch (err: any) {
      console.error("Crew plan export error:", err);
      res.status(500).json({ message: err.message || "Export failed" });
    }
  });

  // 4. Project Plan Export
  app.post("/api/export-project-plan", async (req, res) => {
    const { projectName, activities, summary } = req.body;
    if (!projectName) return res.status(400).json({ message: "projectName required" });
    try {
      const ExcelJS = await import("exceljs");
      const wb = new ExcelJS.default.Workbook();

      // Sheet 1: Schedule
      const ws = wb.addWorksheet("Schedule");
      ws.columns = [
        { header: "Activity #", key: "id", width: 12 },
        { header: "Activity Name", key: "name", width: 30 },
        { header: "Phase", key: "phase", width: 18 },
        { header: "Duration (days)", key: "duration", width: 15 },
        { header: "Manhours", key: "manhours", width: 12 },
        { header: "Crew", key: "crew", width: 8 },
        { header: "Predecessors", key: "predecessors", width: 18 },
        { header: "Start Date", key: "startDate", width: 14 },
        { header: "End Date", key: "endDate", width: 14 },
        { header: "Float (days)", key: "float", width: 12 },
        { header: "Critical Path", key: "critical", width: 14 },
      ];
      const criticalFill = { type: "pattern" as const, pattern: "solid" as const, fgColor: { argb: "FFFCE4EC" } };
      if (activities && Array.isArray(activities)) {
        for (const a of activities) {
          const row = ws.addRow({
            id: a.id || "",
            name: a.name || "",
            phase: a.phase || "",
            duration: a.duration ?? 0,
            manhours: a.manhours ?? 0,
            crew: a.crew ?? 0,
            predecessors: Array.isArray(a.predecessors) ? a.predecessors.join(", ") : (a.predecessors || ""),
            startDate: a.startDate || "",
            endDate: a.endDate || "",
            float: a.float ?? 0,
            critical: a.critical ? "Yes" : "No",
          });
          if (a.critical) {
            row.eachCell((cell: any) => { cell.fill = criticalFill; });
          }
          if (a.milestone) {
            row.font = { bold: true };
          }
        }
      }
      applyExcelStyles(ws);
      // Re-apply critical path highlighting after styles
      if (activities && Array.isArray(activities)) {
        for (let i = 0; i < activities.length; i++) {
          if (activities[i].critical) {
            ws.getRow(i + 2).eachCell((cell: any) => { cell.fill = criticalFill; });
          }
        }
      }

      // Sheet 2: Summary
      const ws2 = wb.addWorksheet("Summary");
      ws2.columns = [
        { header: "Metric", key: "metric", width: 30 },
        { header: "Value", key: "value", width: 25 },
      ];
      if (summary && typeof summary === "object") {
        ws2.addRow({ metric: "Project Name", value: projectName });
        ws2.addRow({ metric: "Total Duration (days)", value: summary.totalDuration ?? "" });
        ws2.addRow({ metric: "Critical Path Duration (days)", value: summary.criticalPathDuration ?? "" });
        ws2.addRow({ metric: "Total Manhours", value: summary.totalManhours ?? "" });
        ws2.addRow({ metric: "Completion Date", value: summary.completionDate ?? "" });
        ws2.addRow({ metric: "Crew Utilization (%)", value: summary.utilization ?? "" });
      }
      applyExcelStyles(ws2);

      const safeName = (projectName || "Project Plan").replace(/[^a-zA-Z0-9 _-]/g, "");
      res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
      res.setHeader("Content-Disposition", `attachment; filename="${safeName} - Project Plan.xlsx"`);
      await wb.xlsx.write(res);
      res.end();
    } catch (err: any) {
      console.error("Project plan export error:", err);
      res.status(500).json({ message: err.message || "Export failed" });
    }
  });

  // 5. Cost Database Export
  app.get("/api/cost-database/export", async (_req, res) => {
    try {
      const entries = storage.getCostDatabase();
      const ExcelJS = await import("exceljs");
      const wb = new ExcelJS.default.Workbook();

      const ws = wb.addWorksheet("Cost Database");
      ws.columns = [
        { header: "Description", key: "description", width: 35 },
        { header: "Size", key: "size", width: 12 },
        { header: "Category", key: "category", width: 14 },
        { header: "Unit", key: "unit", width: 8 },
        { header: "Material Cost ($)", key: "materialUnitCost", width: 16 },
        { header: "Labor Cost ($)", key: "laborUnitCost", width: 14 },
        { header: "Labor Hours/Unit", key: "laborHoursPerUnit", width: 16 },
        { header: "Last Updated", key: "lastUpdated", width: 16 },
      ];
      for (const e of entries) {
        ws.addRow({
          description: e.description,
          size: e.size || "",
          category: e.category || "",
          unit: e.unit || "EA",
          materialUnitCost: e.materialUnitCost ?? 0,
          laborUnitCost: e.laborUnitCost ?? 0,
          laborHoursPerUnit: e.laborHoursPerUnit ?? 0,
          lastUpdated: e.lastUpdated || "",
        });
      }
      applyExcelStyles(ws);

      // Sheet 2: Summary by category
      const ws2 = wb.addWorksheet("Summary");
      ws2.columns = [
        { header: "Category", key: "category", width: 18 },
        { header: "Count", key: "count", width: 10 },
        { header: "Avg Material Cost ($)", key: "avgMat", width: 22 },
        { header: "Avg Labor Cost ($)", key: "avgLab", width: 20 },
      ];
      const catMap = new Map<string, { count: number; totalMat: number; totalLab: number }>();
      for (const e of entries) {
        const cat = e.category || "other";
        const existing = catMap.get(cat) || { count: 0, totalMat: 0, totalLab: 0 };
        existing.count++;
        existing.totalMat += e.materialUnitCost ?? 0;
        existing.totalLab += e.laborUnitCost ?? 0;
        catMap.set(cat, existing);
      }
      for (const [cat, v] of catMap) {
        ws2.addRow({ category: cat, count: v.count, avgMat: v.count > 0 ? Math.round((v.totalMat / v.count) * 100) / 100 : 0, avgLab: v.count > 0 ? Math.round((v.totalLab / v.count) * 100) / 100 : 0 });
      }
      applyExcelStyles(ws2);

      res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
      res.setHeader("Content-Disposition", `attachment; filename="PG Cost Database.xlsx"`);
      await wb.xlsx.write(res);
      res.end();
    } catch (err: any) {
      console.error("Cost DB export error:", err);
      res.status(500).json({ message: err.message || "Export failed" });
    }
  });

  // 6. Folder Combined BOM Export
  app.get("/api/folders/:id/export-combined", async (req, res) => {
    const folder = storage.getFolder(req.params.id);
    if (!folder) return res.status(404).json({ message: "Folder not found" });
    try {
      const projects = storage.getProjectsByFolder(req.params.id);
      const allItems = storage.getFolderCombinedItems(req.params.id);
      const ExcelJS = await import("exceljs");
      const wb = new ExcelJS.default.Workbook();

      // Build project ID -> name map
      const projMap = new Map<string, string>();
      for (const p of projects) projMap.set(p.id, p.name);

      // Sheet 1: Combined BOM
      const ws = wb.addWorksheet("Combined BOM");
      ws.columns = [
        { header: "Source Project", key: "project" },
        { header: "Line #", key: "lineNumber" },
        { header: "Category", key: "category" },
        { header: "Size", key: "size" },
        { header: "Description", key: "description" },
        { header: "Quantity", key: "quantity" },
        { header: "Unit", key: "unit" },
        { header: "Material", key: "material" },
        { header: "Schedule", key: "schedule" },
        { header: "Spec", key: "spec" },
        { header: "Rating", key: "rating" },
        { header: "Confidence", key: "confidence" },
        { header: "Notes", key: "notes" },
      ];
      for (const item of allItems) {
        ws.addRow({
          project: projMap.get((item as any).projectId || "") || "",
          lineNumber: item.lineNumber || "",
          category: item.category || "",
          size: item.size || "",
          description: item.description || "",
          quantity: item.quantity ?? 0,
          unit: item.unit || "EA",
          material: item.material || "",
          schedule: item.schedule || "",
          spec: item.spec || "",
          rating: item.rating || "",
          confidence: item.confidence ? `${Math.round(item.confidence * 100)}%` : "",
          notes: item.notes || "",
        });
      }
      applyExcelStyles(ws);

      // Sheet 2: By Project
      const ws2 = wb.addWorksheet("By Project");
      ws2.columns = [
        { header: "Project", key: "project" },
        { header: "Category", key: "category" },
        { header: "Size", key: "size" },
        { header: "Description", key: "description" },
        { header: "Quantity", key: "quantity" },
        { header: "Unit", key: "unit" },
      ];
      for (const p of projects) {
        const projItems = allItems.filter((i: any) => i.projectId === p.id);
        for (const item of projItems) {
          ws2.addRow({
            project: p.name,
            category: item.category || "",
            size: item.size || "",
            description: item.description || "",
            quantity: item.quantity ?? 0,
            unit: item.unit || "EA",
          });
        }
        // Subtotal row
        if (projItems.length > 0) {
          const subtotal = ws2.addRow({ project: `${p.name} SUBTOTAL`, category: "", size: "", description: "", quantity: projItems.reduce((s: number, i: any) => s + (i.quantity ?? 0), 0), unit: "" });
          subtotal.font = { bold: true };
        }
      }
      applyExcelStyles(ws2);

      // Sheet 3: Summary pivot
      const ws3 = wb.addWorksheet("Summary");
      ws3.columns = [
        { header: "Category", key: "category", width: 18 },
        { header: "Size", key: "size", width: 12 },
        { header: "Item Count", key: "count", width: 12 },
        { header: "Total Qty", key: "totalQty", width: 12 },
      ];
      const pivot = new Map<string, { count: number; totalQty: number }>();
      for (const item of allItems) {
        const key = `${item.category || "other"}|${item.size || "N/A"}`;
        const existing = pivot.get(key) || { count: 0, totalQty: 0 };
        existing.count++;
        existing.totalQty += item.quantity ?? 0;
        pivot.set(key, existing);
      }
      for (const [key, val] of pivot) {
        const [category, size] = key.split("|");
        ws3.addRow({ category, size, count: val.count, totalQty: val.totalQty });
      }
      ws3.addRow({ category: "TOTAL", size: "", count: allItems.length, totalQty: allItems.reduce((s: number, i: any) => s + (i.quantity ?? 0), 0) });
      ws3.getRow(ws3.rowCount).font = { bold: true };
      applyExcelStyles(ws3);

      const safeName = ((folder as any).name || "Combined BOM").replace(/[^a-zA-Z0-9 _-]/g, "");
      res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
      res.setHeader("Content-Disposition", `attachment; filename="${safeName} - Combined BOM.xlsx"`);
      await wb.xlsx.write(res);
      res.end();
    } catch (err: any) {
      console.error("Folder export error:", err);
      res.status(500).json({ message: err.message || "Export failed" });
    }
  });


  // Scope Split Export
  app.post("/api/export-scope-split", async (req, res) => {
    const { subShop, yourField, yourFull, summary } = req.body;
    try {
      const ExcelJS = await import("exceljs");
      const wb = new ExcelJS.default.Workbook();
      const headerFill = { type: "pattern" as const, pattern: "solid" as const, fgColor: { argb: "FF01696F" } };
      const headerFont = { bold: true, color: { argb: "FFFFFFFF" }, size: 11 };

      function addScopeSheet(name: string, items: any[]) {
        const ws = wb.addWorksheet(name);
        ws.columns = [
          { header: "Category", key: "category", width: 14 },
          { header: "Size", key: "size", width: 10 },
          { header: "Description", key: "description", width: 35 },
          { header: "Quantity", key: "quantity", width: 10 },
          { header: "Unit", key: "unit", width: 8 },
          { header: "Welds", key: "welds", width: 8 },
          { header: "Location", key: "location", width: 10 },
          { header: "Source Page", key: "sourcePage", width: 12 },
        ];
        for (const item of (items || [])) ws.addRow(item);
        ws.getRow(1).eachCell((cell: any) => { cell.fill = headerFill; cell.font = headerFont; cell.alignment = { horizontal: "center" as const }; });
        for (let r = 2; r <= ws.rowCount; r++) {
          if (r % 2 === 0) ws.getRow(r).eachCell((cell: any) => { cell.fill = { type: "pattern" as const, pattern: "solid" as const, fgColor: { argb: "FFF2F2F2" } }; });
        }
      }

      addScopeSheet("Sub Shop Fab", subShop);
      addScopeSheet("Your Field Welds", yourField);
      addScopeSheet("Your Full Scope", yourFull);

      // Summary sheet
      const ws4 = wb.addWorksheet("Summary");
      ws4.columns = [{ header: "Metric", key: "metric", width: 25 }, { header: "Value", key: "value", width: 18 }];
      if (summary) {
        ws4.addRow({ metric: "Total Project Welds", value: summary.totalWelds ?? 0 });
        ws4.addRow({ metric: "Sub\'s Welds", value: summary.subWelds ?? 0 });
        ws4.addRow({ metric: "Your Welds", value: summary.yourWelds ?? 0 });
        ws4.addRow({ metric: "Your Est. Manhours", value: summary.yourMH ?? 0 });
      }
      ws4.getRow(1).eachCell((cell: any) => { cell.fill = headerFill; cell.font = headerFont; cell.alignment = { horizontal: "center" as const }; });

      res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
      res.setHeader("Content-Disposition", `attachment; filename="Scope Split.xlsx"`);
      await wb.xlsx.write(res);
      res.end();
    } catch (err: any) {
      console.error("Scope split export error:", err);
      res.status(500).json({ message: err.message || "Export failed" });
    }
  });

  return httpServer;
}
