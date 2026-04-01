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
import { insertEstimateProjectSchema, insertCostDatabaseEntrySchema, estimateItemSchema, markupsSchema, insertPurchaseRecordSchema, insertBidSchema, insertVendorQuoteSchema } from "@shared/schema";
import { z } from "zod";
import { parse as csvParse } from "csv-parse/sync";

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
  method: z.enum(["bill", "justin"]).default("justin"),
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
  estimateMethod: z.enum(["bill", "justin", "manual"]).optional(),
});
import { generateBillsWorkbook, generateJustinsWorkbook } from "./excelExport";

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

const MAX_FILE_SIZE = 100 * 1024 * 1024; // 100MB
const UPLOAD_DIR = "/tmp/pg-unified-uploads";
const RENDER_DIR = "/tmp/pg-unified-renders";
const CHUNK_SIZE = 40;

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

function splitPdfIntoChunks(pdfPath: string, pageCount: number): { chunkPath: string; startPage: number; endPage: number }[] {
  if (pageCount <= CHUNK_SIZE) {
    return [{ chunkPath: pdfPath, startPage: 1, endPage: pageCount }];
  }
  const chunks: { chunkPath: string; startPage: number; endPage: number }[] = [];
  const chunkDir = path.join(RENDER_DIR, `chunks_${Date.now()}`);
  fs.mkdirSync(chunkDir, { recursive: true });
  for (let start = 1; start <= pageCount; start += CHUNK_SIZE) {
    const end = Math.min(start + CHUNK_SIZE - 1, pageCount);
    const chunkPath = path.join(chunkDir, `chunk_${start}_${end}.pdf`);
    try {
      execFileSync("qpdf", [pdfPath, "--pages", ".", `${start}-${end}`, "--", chunkPath], { timeout: 60000 });
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
 * Detect if a pipe quantity is likely inches misread as feet.
 * On isometrics, short pipe runs (0'-11", 0'-8", etc.) are common.
 * If the AI outputs just "11" without any unit marker, and it's a small integer
 * ≤ 18, it's very likely inches (not 11 feet of pipe on one line).
 * Returns the corrected quantity in feet, or the original if no correction needed.
 */
function correctPipeLengthIfInches(qty: number, rawQty: string, description: string): { correctedQty: number; wasCorrection: boolean; note: string } {
  // If the raw string already had feet/inch markers, parsePipeLength handled it
  if (/[\u2018\u2019\u201C\u201D''""']/.test(rawQty)) {
    return { correctedQty: qty, wasCorrection: false, note: "" };
  }
  // Flag whole numbers 1-18 as ambiguous \u2014 do NOT auto-correct (too aggressive)
  if (Number.isInteger(qty) && qty >= 1 && qty <= 18) {
    return {
      correctedQty: qty,
      wasCorrection: false,
      note: `Pipe qty ${qty}: no unit marker \u2014 could be ${qty}' (feet) or ${qty}" (inches = ${Math.round((qty / 12) * 100) / 100} LF). Verify against drawing.`
    };
  }
  // Values 19-36 are also ambiguous \u2014 flag but don't auto-correct.
  if (Number.isInteger(qty) && qty >= 19 && qty <= 36) {
    return {
      correctedQty: qty,
      wasCorrection: false,
      note: `Pipe qty ${qty}: no unit marker \u2014 could be ${qty}' (feet) or ${qty}" (inches = ${Math.round((qty / 12) * 100) / 100} LF). Verify against drawing.`
    };
  }
  return { correctedQty: qty, wasCorrection: false, note: "" };
}

function isTitlePage(text: string): boolean {
  if (text.trim().length < 50) return true;
  const hasBom = /\b(PIPE|ELBOW|TEE|VALVE|FLANGE|GASKET|BOLT|STUD|REDUCER|W\d+X|HSS|FOOTING|REBAR|STORM|SEWER|ASPHALT)\b/i.test(text);
  if (!hasBom && /\bAREA\s+\d+\b/i.test(text)) return true;
  if (/FOR REFERENCE ONLY/i.test(text) && !hasBom) return true;
  return false;
}

// ============================================================
// MECHANICAL BOM EXTRACTION
// ============================================================

const MECHANICAL_PROMPT = `You are an expert at reading BOM (Bill of Materials) tables from piping isometric drawings.
Your job is to produce EXACT, ACCURATE data. Accuracy is critical — this is used for material procurement.

Each image is a cropped BOM table from a piping isometric drawing page.
The table has SHOP and FIELD sections, each with columns: NO. | QTY | SIZE | DESCRIPTION

RULES — READ CAREFULLY:
1. Read every cell EXACTLY as printed. Do NOT guess, round, estimate, or infer values.
2. QTY column:
   - For PIPE: QTY is a length in feet-inches format like 16'-8" or 3'-0" or 22'-6". Copy it EXACTLY as written (e.g. "16'-8\""). The apostrophe means feet, the double-quote means inches.
   - For fittings/valves/bolts/gaskets: QTY is a simple integer (1, 2, 4, 8, etc.). Read the number carefully — distinguish between 1, 4, 8, etc.
   - DO NOT convert units. DO NOT do math. Just copy what is printed.
3. SIZE column — CRITICAL:
   - Piping sizes are ALWAYS in nominal pipe sizes: 1/2", 3/4", 1", 1-1/2", 2", 3", 4", 6", 8", 10", 12", 14", 16", 18", 20", 24", 30", 36", 42", 48".
   - Reducers have two sizes like 6"x4" or 2"x1".
   - Bolts have sizes like 5/8"x3 3/4" or 3/4"x4 1/4" (bolt diameter x bolt length).
   - NEVER report a pipe size larger than 48". Standard pipe sizes do not exceed 48".
4. DESCRIPTION: Copy the FULL text, joining multi-line text with spaces. Include all specs (ASME B16.xx, ASTM Axxx, CLASS xxxx, SCH xx, etc.)
5. Include ALL rows from BOTH the SHOP table and the FIELD table.
6. SKIP: title block text, engineer stamps, notes, revision blocks, drawing borders.
7. If a page image has NO BOM table or only empty table headers, return an empty items array.

COMMON ERRORS TO AVOID:
- Pipe lengths: Don't confuse 11' (11 feet) with 11" (11 inches = 0.92 feet). The apostrophe (') means FEET, the double-quote (") means INCHES.
- CRITICAL: Short pipe runs are often in INCHES. If you see 11" or 0'-11" that is 11 INCHES (0.92 LF), NOT 11 feet. Always include the " symbol for inches.
- Pipe lengths: 3'-4" means 3 feet 4 inches, NOT 34. 10'-2" means 10 feet 2 inches. 0'-8" means 8 inches = 0.67 feet.
- If a pipe run has no ' or " symbol and the number is small (under 18), it is almost certainly INCHES. Write it as e.g. 11" not 11.
- Size column: 1-1/2" is a valid pipe size (one and a half inches). Don't misread as 1" or 11/2".
- NPS sizes: The ONLY valid NPS sizes are: 1/2", 3/4", 1", 1-1/4", 1-1/2", 2", 2-1/2", 3", 4", 6", 8", 10", 12", 14", 16", 18", 20", 24", 30", 36", 42", 48". If you read something else, you probably misread it.
- Bolt sizes like 5/8"x4" or 3/4"x4 1/4" are bolt diameter x length, NOT pipe sizes.
- QTY for pipe is ALWAYS a length (feet-inches). QTY for everything else is ALWAYS a count (integer).
- Don't confuse item numbers (NO. column) with quantities (QTY column). The NO. column is just a line number.
- SHOP items are fabricated in a shop. FIELD items are installed in the field. Read both tables.
- If a row has empty QTY, SIZE, or DESCRIPTION cells, skip it — it's a header or separator.

VALIDATION — before returning, verify each item:
(a) QTY is either a feet-inches length for pipe OR an integer count for non-pipe.
(b) SIZE is a valid NPS or bolt size from the lists above.
(c) DESCRIPTION starts with a material type keyword (PIPE, ELBOW, TEE, FLANGE, VALVE, GASKET, BOLT, STUD, REDUCER, CAP, etc.)
Double-check your work before returning. Review each item and fix any obvious errors.

Return ONLY valid JSON (no markdown fences, no extra text):
{"pages": [{"pageNum": PAGE_NUMBER, "items": [{"itemNo": 1, "qty": "16'-8\"", "size": "1\"", "description": "PIPE, SMLS, BE OR PE, SCH 80, ASME B36.10, CS ASTM A106, GRD B", "section": "SHOP"}]}]}`;

const MECHANICAL_CLOUD_PROMPT = `You are an expert at reading piping isometric drawings and identifying REVISION CLOUDS.

For each page, you will see TWO images:
1. The FULL PAGE isometric drawing (showing the piping, fittings, and any revision clouds)
2. The CROPPED BOM TABLE from the same page

REVISION CLOUDS are wavy/scalloped bubbles or irregular curved outlines drawn around parts of the drawing to indicate changes from the previous revision. They look like bumpy cloud shapes surrounding modified areas.

Your job:
1. First, extract the BOM exactly as described below.
2. Then, examine the FULL PAGE drawing for any REVISION CLOUDS.
3. For each BOM item, determine whether the corresponding component on the drawing is INSIDE a revision cloud.
4. An item is "clouded" if:
   - Its corresponding piping/fitting/valve on the drawing is enclosed in or touched by a revision cloud
   - Its BOM table row itself is enclosed in a revision cloud
   - The dimension, routing, or connection point it represents was changed (shown by a cloud)
5. If NO revision clouds exist on the page, mark all items as clouded=false.

BOM EXTRACTION RULES:
1. Read every cell EXACTLY as printed. Do NOT guess, round, estimate, or infer values.
2. QTY column:
   - For PIPE: QTY is a length in feet-inches format like 16'-8" or 3'-0" or 22'-6". Copy EXACTLY.
   - For fittings/valves/bolts/gaskets: QTY is a simple integer (1, 2, 4, 8, etc.).
   - DO NOT convert units. Just copy what is printed.
3. SIZE column:
   - Piping sizes: 1/2", 3/4", 1", 1-1/2", 2", 3", 4", 6", 8", 10", 12", etc.
   - Reducers: 6"x4", 2"x1", etc.
   - Bolts: 5/8"x3 3/4", 3/4"x4 1/4", etc.
   - NEVER report pipe sizes larger than 48".
4. DESCRIPTION: Copy the FULL text including all specs.
5. Include ALL rows from BOTH SHOP and FIELD tables.
6. SKIP: title block text, engineer stamps, notes, drawing borders.

COMMON ERRORS TO AVOID:
- Pipe lengths: Don't confuse 11' (11 feet) with 11" (11 inches = 0.92 feet). The apostrophe (') means FEET, the double-quote (") means INCHES.
- CRITICAL: Short pipe runs are often in INCHES. If you see 11" or 0'-11" that is 11 INCHES (0.92 LF), NOT 11 feet. Always include the " symbol for inches.
- Pipe lengths: 3'-4" means 3 feet 4 inches, NOT 34. 10'-2" means 10 feet 2 inches. 0'-8" means 8 inches = 0.67 feet.
- If a pipe run has no ' or " symbol and the number is small (under 18), it is almost certainly INCHES. Write it as e.g. 11" not 11.
- Size column: 1-1/2" is a valid pipe size (one and a half inches). Don't misread as 1" or 11/2".
- NPS sizes: The ONLY valid NPS sizes are: 1/2", 3/4", 1", 1-1/4", 1-1/2", 2", 2-1/2", 3", 4", 6", 8", 10", 12", 14", 16", 18", 20", 24", 30", 36", 42", 48". If you read something else, you probably misread it.
- Bolt sizes like 5/8"x4" or 3/4"x4 1/4" are bolt diameter x length, NOT pipe sizes.
- QTY for pipe is ALWAYS a length (feet-inches). QTY for everything else is ALWAYS a count (integer).
- Don't confuse item numbers (NO. column) with quantities (QTY column). The NO. column is just a line number.
- SHOP items are fabricated in a shop. FIELD items are installed in the field. Read both tables.
- If a row has empty QTY, SIZE, or DESCRIPTION cells, skip it — it's a header or separator.

VALIDATION — before returning, verify each item:
(a) QTY is either a feet-inches length for pipe OR an integer count for non-pipe.
(b) SIZE is a valid NPS or bolt size from the lists above.
(c) DESCRIPTION starts with a material type keyword (PIPE, ELBOW, TEE, FLANGE, VALVE, GASKET, BOLT, STUD, REDUCER, CAP, etc.)
Double-check your work before returning. Review each item and fix any obvious errors.

Return ONLY valid JSON (no markdown fences, no extra text):
{"pages": [{"pageNum": PAGE_NUMBER, "items": [{"itemNo": 1, "qty": "16'-8\"", "size": "1\"", "description": "PIPE, SMLS, BE, SCH 10S, ASME B36.19, SS ASTM A312", "section": "SHOP", "clouded": false}]}]}`;

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
  pageItems: Map<number, any[]>,
  bomImages: Map<number, string>,
): Promise<Map<number, any[]>> {
  const client = getAnthropicClient();
  const result = new Map<number, any[]>();

  // Copy all items first
  for (const [pageNum, items] of pageItems) {
    result.set(pageNum, [...items]);
  }

  // Only verify pages that have items and BOM images
  const pagesToVerify: { pageNum: number; items: any[]; bomPath: string }[] = [];
  for (const [pageNum, items] of pageItems) {
    if (items.length > 0 && bomImages.has(pageNum)) {
      pagesToVerify.push({ pageNum, items, bomPath: bomImages.get(pageNum)! });
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
            const items = result.get(targetPage);
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
// AI VISION EXTRACTION
// ============================================================

async function extractWithVision(
  pageImages: { pageNum: number; imagePath: string }[],
  prompt: string,
  discipline: string
): Promise<{ results: Map<number, any[]>; authFailures: number }> {
  let client = getAnthropicClient();
  const results = new Map<number, any[]>();
  let authFailures = 0;
  const BATCH_SIZE = 2;

  for (let batchStart = 0; batchStart < pageImages.length; batchStart += BATCH_SIZE) {
    const batch = pageImages.slice(batchStart, batchStart + BATCH_SIZE);
    console.log(`  AI Vision [${discipline}] batch: pages ${batch.map(p => p.pageNum).join(", ")}...`);

    const content: Anthropic.MessageCreateParams["messages"][0]["content"] = [];

    for (const page of batch) {
      const imgData = fs.readFileSync(page.imagePath);
      const b64 = imgData.toString("base64");
      content.push({ type: "image" as const, source: { type: "base64" as const, media_type: "image/png", data: b64 } });
      content.push({ type: "text" as const, text: `[PAGE ${page.pageNum}]` });
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
                results.set(page.pageNum, page.items);
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
          for (const page of batch) { results.set(page.pageNum, []); }
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
          for (const page of batch) { results.set(page.pageNum, []); }
        }
      }
    }
  }

  return { results, authFailures };
}

// Cloud-aware extraction: sends FULL PAGE + BOM CROP per page
async function extractWithCloudDetection(
  pageImages: { pageNum: number; bomImagePath: string; fullImagePath: string }[],
  prompt: string,
  onPageComplete?: (completedCount: number, totalCount: number) => void
): Promise<{ results: Map<number, any[]>; authFailures: number }> {
  let client = getAnthropicClient();
  const results = new Map<number, any[]>();
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
            for (const p of parsed.pages) {
              if (Array.isArray(p.items)) {
                allItems.push(...p.items);
              }
            }
            results.set(page.pageNum, allItems);
            parseSuccess = true;
          }
        } catch (parseErr) {
          console.error(`  Failed to parse cloud detection response (attempt ${attempt + 1}) page ${page.pageNum}`);
        }
        if (parseSuccess) { break; }
        if (attempt === 1) { results.set(page.pageNum, []); }
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
          results.set(page.pageNum, []);
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
async function processRenderedPages(jobDir: string, mode: "bom" | "bom+full" | "ocr-only"): Promise<void> {
  const files = fs.readdirSync(jobDir).filter(f => f.endsWith(".png") && !f.includes("_bom") && !f.includes("_full") && !f.includes("_ocr"));
  const CONCURRENCY = 4;
  
  for (let i = 0; i < files.length; i += CONCURRENCY) {
    const batch = files.slice(i, i + CONCURRENCY);
    await Promise.all(batch.map(async (imgFile) => {
      const imgPath = path.join(jobDir, imgFile);
      const basename = imgFile.replace(/\.png$/, "");
      try {
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
          const cx = Math.floor(w * 50 / 100);
          const cw = w - cx;
          const ch = Math.floor(h * 65 / 100);
          
          // Crop BOM area
          const bomPath = path.join(jobDir, `${basename}_bom.png`);
          await execFileAsync("convert", [imgPath, "-crop", `${cw}x${ch}+${cx}+0`, "+repage", bomPath], { timeout: 30000 });
          
          if (mode === "bom+full") {
            // Create resized full page image
            const fullPath = path.join(jobDir, `${basename}_full.png`);
            await execFileAsync("convert", [imgPath, "-resize", "50%", fullPath], { timeout: 30000 });
          }
          
          // OCR the BOM crop
          const ocrBase = path.join(jobDir, `${basename}_ocr`);
          try {
            await execFileAsync("tesseract", [bomPath, ocrBase, "--psm", "4", "-l", "eng"], { timeout: 30000 });
          } catch (e) { console.warn("Suppressed error:", e); }
          
          // Remove original full-res image (keep _bom and _full)
          try { fs.unlinkSync(imgPath); } catch (e) { console.warn("Suppressed error:", e); }
        } else {
          // OCR-only mode (full page)
          const ocrBase = path.join(jobDir, `${basename}_ocr`);
          try {
            await execFileAsync("tesseract", [imgPath, ocrBase, "--psm", "4", "-l", "eng"], { timeout: 30000 });
          } catch (e) { console.warn("Suppressed error:", e); }
        }
      } catch (e) { console.warn("Suppressed error processing", imgFile, e); }
    }));
  }
}



async function renderCroppedBomImages(pdfPath: string, pageCount: number): Promise<{
  pageImages: { pageNum: number; imagePath: string; tesseractText: string }[];
  jobDir: string;
}> {
  const jobDir = path.join(RENDER_DIR, Date.now().toString());
  fs.mkdirSync(jobDir, { recursive: true });

  await execFileAsync("pdftoppm", ["-r", "300", "-png", pdfPath, path.join(jobDir, "page")], {
    maxBuffer: 200 * 1024 * 1024, timeout: 180000,
  });

    await processRenderedPages(jobDir, "bom");

  const pageImages: { pageNum: number; imagePath: string; tesseractText: string }[] = [];
  const padLen = Math.max(2, String(pageCount).length);

  for (let p = 1; p <= pageCount; p++) {
    const padded = String(p).padStart(padLen, "0");
    const bomImg = path.join(jobDir, `page-${padded}_bom.png`);
    const ocrFile = path.join(jobDir, `page-${padded}_ocr.txt`);
    let tesseractText = "";
    if (fs.existsSync(ocrFile)) tesseractText = fs.readFileSync(ocrFile, "utf-8");
    if (fs.existsSync(bomImg) && !isTitlePage(tesseractText)) {
      pageImages.push({ pageNum: p, imagePath: bomImg, tesseractText });
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

  await execFileAsync("pdftoppm", ["-r", "300", "-png", pdfPath, path.join(jobDir, "page")], {
    maxBuffer: 200 * 1024 * 1024, timeout: 180000,
  });

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

  await execFileAsync("pdftoppm", ["-r", String(dpi), "-png", pdfPath, path.join(jobDir, "page")], {
    maxBuffer: 300 * 1024 * 1024, timeout: 240000,
  });

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
  // Clean scans have moderate OCR text
  if (avgTextLength >= 100) return "clean_scan";
  // Poor scans have very little extractable text
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
// CONTINUATION PAGE DEDUP (Spec Item 5 - Improved)
// ============================================================

function dedupContinuationPages(items: any[]): any[] {
  // Group items by sheet number
  const sheetGroups: Record<number, any[]> = {};
  for (const item of items) {
    const sheetMatch = (item.notes || "").match(/Sheet\s+(\d+)/i);
    const sheet = sheetMatch ? parseInt(sheetMatch[1], 10) : 0;
    if (!sheetGroups[sheet]) sheetGroups[sheet] = [];
    sheetGroups[sheet].push(item);
  }

  const sheetNums = Object.keys(sheetGroups).map(Number).sort((a, b) => a - b);

  // Only dedup when the ENTIRE page has >80% identical items to the previous page
  for (let i = 0; i < sheetNums.length - 1; i++) {
    const curSheet = sheetNums[i];
    const nextSheet = sheetNums[i + 1];
    // Only check adjacent sheets
    if (nextSheet - curSheet !== 1) continue;

    const curItems = sheetGroups[curSheet];
    const nextItems = sheetGroups[nextSheet];

    // Build key sets for comparison
    const curKeys = new Set(curItems.map((item: any) => `${item.category}|${item.size}|${item.description}|${item.quantity}`));
    const nextKeys = new Set(nextItems.map((item: any) => `${item.category}|${item.size}|${item.description}|${item.quantity}`));

    // Count how many items on next page match current page
    let matchCount = 0;
    for (const key of nextKeys) {
      if (curKeys.has(key)) matchCount++;
    }

    const overlapPct = nextItems.length > 0 ? matchCount / nextItems.length : 0;

    // Only dedup when >80% of the ENTIRE next page matches
    if (overlapPct <= 0.8) continue;

    // For matching items: mark as dedup candidates (don't hard-delete)
    for (const nextItem of nextItems) {
      const nextKey = `${nextItem.category}|${nextItem.size}|${nextItem.description}|${nextItem.quantity}`;
      // Find matching item on previous page
      const matchingCurItem = curItems.find((ci: any) =>
        `${ci.category}|${ci.size}|${ci.description}|${ci.quantity}` === nextKey
      );

      if (matchingCurItem) {
        // Keep items that match but have different notes or specs
        const sameNotes = (matchingCurItem.spec || "") === (nextItem.spec || "") &&
                         (matchingCurItem.material || "") === (nextItem.material || "") &&
                         (matchingCurItem.schedule || "") === (nextItem.schedule || "");
        if (!sameNotes) continue; // Different specs, keep both

        // Mark as dedup candidate instead of hard-deleting — do NOT add quantity to avoid double-counting
        nextItem._dedupCandidate = true;
        nextItem.dedupNote = `Possible duplicate from Sheet ${curSheet} — review and delete if confirmed`;
      }
    }
  }

  const dedupCount = items.filter((i: any) => i._dedupCandidate).length;
  if (dedupCount > 0) {
    console.log(`  Continuation page dedup: marked ${dedupCount} items as dedup candidates`);
  }
  // Return ALL items (including dedup candidates - UI will dim them)
  return items;
}

// ============================================================
// ENHANCED CONFIDENCE SCORING (Spec Item 4)
// ============================================================

function computeConfidenceScore(item: any, pdfQuality?: "vector" | "clean_scan" | "poor_scan"): { confidence: "high" | "medium" | "low"; confidenceScore: number; confidenceNotes: string } {
  let score = 100;
  const notes: string[] = [];

  // Size validation
  if (!item.size || item.size === "N/A" || item.size.trim() === "") {
    score -= 40;
    notes.push("Missing size");
  } else if (item.category !== "bolt" && !isValidNPS(item.size) && !/x/i.test(item.size)) {
    score -= 20;
    notes.push("Size not standard NPS");
  }

  // Description quality
  if (!item.description || item.description.length < 10) {
    score -= 30;
    notes.push("Description too short");
  } else if (item.description.length < 20) {
    score -= 10;
    notes.push("Description may be incomplete");
  }
  if (item.description && !/ASME|ASTM/.test(item.description.toUpperCase())) {
    score -= 10;
    notes.push("No ASME/ASTM spec reference");
  }

  // Quantity plausibility
  const isPipe = item.category === "pipe";
  if (isPipe && item.quantity > 500) {
    score -= 30;
    notes.push("Pipe qty >500 LF");
  } else if (isPipe && item.quantity > 100) {
    score -= 10;
    notes.push("Pipe qty >100 LF");
  }
  if (!isPipe && item.quantity > 100) {
    score -= 15;
    notes.push("Non-pipe qty >100");
  }
  if (item.quantity <= 0) {
    score -= 40;
    notes.push("Qty is 0 or negative");
  }

  // Category consistency
  if (item.category === "other") {
    score -= 15;
    notes.push("Unclassified category");
  }

  // Incorporate validation flags from validateExtractedItems
  if (item._validationFlag === "low") {
    score = Math.min(score, 50);
  } else if (item._validationFlag === "medium") {
    score = Math.min(score, 75);
  }

  // Add validation notes
  if (item._validationNotes && item._validationNotes.length > 0) {
    notes.push(...item._validationNotes);
  }

  // Size warning from autoCorrect
  if (item._sizeWarning) {
    notes.push(item._sizeWarning);
    score -= 15;
  }

  // If item came from a poor_scan page, subtract 25
  if (pdfQuality === "poor_scan") {
    score -= 25;
    notes.push("Poor scan quality — manual review recommended");
  }

  // Clamp score
  score = Math.max(0, Math.min(100, score));

  let confidence: "high" | "medium" | "low";
  if (score >= 85) confidence = "high";
  else if (score >= 60) confidence = "medium";
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
  function parseMechanicalQty(rawQty: string, category: string, description: string): { qty: number; notes: string[] } {
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
      const correction = correctPipeLengthIfInches(qty, rq, description);
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

    for (const page of adjustedPageImages) {
      const entries = visionResults.get(page.pageNum) || [];
      for (const entry of entries) {
        if (!entry.description || entry.description.length < 3) continue;
        const rawQty = String(entry.qty || "1").trim();
        const category = detectMechanicalCategory(entry.description);
        const isPipe = category === "pipe";
        const { qty: parsedQty, notes: qtyNotes } = parseMechanicalQty(rawQty, category, entry.description);
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
          notes: `Sheet ${page.globalPageNum} (${entry.section || "SHOP"})${extraNotes}`,
          sourcePage: page.globalPageNum,
          revisionClouded: entry.clouded === true,
        };
        if (qtyNotes.length > 0) {
          rawItem._validationNotes = rawItem._validationNotes || [];
          rawItem._validationNotes.push(...qtyNotes);
        }
        rawItem = autoCorrectItem(rawItem);
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
      const bomImageMap = new Map<number, string>();
      for (const pi of pageImages) {
        bomImageMap.set(pi.pageNum, pi.imagePath);
      }
      console.log(`  Running verification pass on ${visionResults.size} pages...`);
      verifiedResults = await verifyExtractedItems(visionResults, bomImageMap);
    }

    for (const page of adjustedPageImages) {
      const entries = verifiedResults.get(page.pageNum) || [];
      for (const entry of entries) {
        if (!entry.description || entry.description.length < 3) continue;
        const rawQty = String(entry.qty || "1").trim();
        const category = detectMechanicalCategory(entry.description);
        const isPipe = category === "pipe";
        const { qty: parsedQty, notes: qtyNotes } = parseMechanicalQty(rawQty, category, entry.description);
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
          notes: `Sheet ${page.globalPageNum} (${entry.section || "SHOP"})${extraNotes}`,
          sourcePage: page.globalPageNum,
          revisionClouded: entry.clouded === true,
        };
        if (qtyNotes.length > 0) {
          rawItem._validationNotes = rawItem._validationNotes || [];
          rawItem._validationNotes.push(...qtyNotes);
        }
        rawItem = autoCorrectItem(rawItem);
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
    const entries = visionResults.get(page.pageNum) || [];
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
        notes: `Sheet ${page.globalPageNum}`,
        sourcePage: page.globalPageNum,
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
    const entries = visionResults.get(page.pageNum) || [];
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
        notes: `Sheet ${page.globalPageNum}`,
        sourcePage: page.globalPageNum,
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

  const chunks = splitPdfIntoChunks(pdfPath, pageCount);
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

  // ============================================================
  // PHASE 1: PRE-RENDER ALL CHUNKS (no API calls needed)
  // ============================================================
  console.log(`\n=== PHASE 1: Rendering all ${chunks.length} chunks ===`);
  const renderedChunks: { chunkIndex: number; rendered: RenderedMechanicalChunk | RenderedStructuralChunk | RenderedCivilChunk }[] = [];

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
        console.error(`  Render chunk ${chunkNum} failed:`, renderErr.message || renderErr);
        const renderWarning = `Render chunk ${chunkNum} (pages ${chunk.startPage}-${chunk.endPage}) failed: ${renderErr.message || "Unknown error"}`;
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
      error: `All ${chunks.length} render chunks failed. No pages could be prepared for extraction.`,
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

    // Step: Continuation page dedup
    console.log(`  Running continuation page dedup...`);
    const dedupedItems = dedupContinuationPages([...allItems]);
    allItems.length = 0;
    allItems.push(...dedupedItems);

    // Step: Piping validation rules (Council Item 4)
    console.log(`  Running piping validation rules...`);
    validatePipingBom(allItems);
  }

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
  });

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
  alloyGroup: string
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
  if (catLower === "fitting" || catLower === "elbow" || catLower === "tee" || catLower === "reducer" || catLower === "cap" || catLower === "coupling" || catLower === "union") {
    let fittingEiMult = 0.6;
    let fittingType = "fitting";
    if (descLower.includes("90") && descLower.includes("elbow") || (catLower === "elbow" && descLower.includes("90"))) { fittingEiMult = 1.0; fittingType = "90° Elbow"; }
    else if (descLower.includes("45") && descLower.includes("elbow") || (catLower === "elbow" && descLower.includes("45"))) { fittingEiMult = 0.8; fittingType = "45° Elbow"; }
    else if (catLower === "elbow" || descLower.includes("elbow") || descLower.includes("ell") || descLower.includes("return")) { fittingEiMult = 1.0; fittingType = "Elbow"; }
    else if (catLower === "tee" || descLower.includes("tee")) { fittingEiMult = 1.3; fittingType = "Tee"; }
    else if (catLower === "reducer" || descLower.includes("reducer") || descLower.includes("swage")) { fittingEiMult = 0.7; fittingType = "Reducer"; }
    else if (catLower === "cap" || descLower.includes("cap")) { fittingEiMult = 0.5; fittingType = "Cap"; }
    else if (catLower === "coupling" || descLower.includes("coupling") || descLower.includes("nipple")) { fittingEiMult = 0.6; fittingType = "Coupling"; }
    else if (catLower === "union" || descLower.includes("union")) { fittingEiMult = 0.4; fittingType = "Union"; }

    const weldTable = lr.butt_welds_ei;
    const found = findClosestKey(weldTable, nps);
    if (found && weldTable[found.key]) {
      const schedKey = sched in weldTable[found.key] ? sched : ("STD" in weldTable[found.key] ? "STD" : Object.keys(weldTable[found.key])[0]);
      const baseEi = weldTable[found.key][schedKey] || 0;
      const ei = baseEi * fittingEiMult;
      const mh = ei * fieldMhPerEi * alloyFactor * elevFactor * weldLocationFactor;
      const warn = sizeWarn(found, nps);
      const basis = `Bill's EI: ${fittingType} ${found.key}\" ${schedKey} → BW EI=${baseEi} × ${fittingEiMult} = ${ei.toFixed(1)} EI × ${fieldMhPerEi} MH/EI × ${elevFactor} (elev) × ${alloyFactor} (alloy) × ${weldLocationFactor} (${pipeLocation} weld) = ${mh.toFixed(4)} MH${warn}`;
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
  justinData: any
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
      if (material === "SS") { mh = match.val.ss_mh_per_weld || 0; col = "SS"; }
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

  // --- FITTING ---
  if (catLower === "fitting" || catLower === "elbow" || catLower === "tee" || catLower === "reducer" || catLower === "cap" || catLower === "coupling") {
    const weldMatch = findBestMatch(factors.welds || {});
    if (weldMatch) {
      const schedNorm = normalizeSchedule(schedule);
      const baseMH = material === "SS" ? (weldMatch.val.ss_mh_per_weld || 0) : (schedNorm === "80" ? (weldMatch.val.sch80_mh_per_weld || 0) : (weldMatch.val.std_mh_per_weld || 0));
      const mh = baseMH * 0.5;
      const warn = jSizeWarn(weldMatch);
      return { mh, calcBasis: `Justin: Fitting ${weldMatch.matchKey} → weld base=${baseMH.toFixed(2)} × 0.50 = ${mh.toFixed(4)} MH${warn}`, sizeMatchExact: weldMatch.exact };
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
  for (const item of items) {
    const cat = (item.category || "").toLowerCase();
    const desc = (item.description || "").toLowerCase();
    const qty = item.quantity || 0;
    const size = item.size || "";

    if (cat === "elbow" || desc.includes("elbow") || desc.includes("ell")) {
      welds.push(computeEstimateItem({
        id: randomUUID(), lineNumber: 0, category: "weld" as any,
        description: `BW for ${size} ELBOW (auto-inferred)`,
        size, quantity: qty * 2, unit: "EA",
        materialUnitCost: 0, laborUnitCost: 0, laborHoursPerUnit: 0,
        materialExtension: 0, laborExtension: 0, totalCost: 0,
        notes: "Auto-inferred: 2 butt welds per elbow", fromDatabase: false,
        weldAssumption: "2 butt welds per elbow (auto-inferred)",
      }));
    } else if (cat === "tee" || desc.includes("tee")) {
      welds.push(computeEstimateItem({
        id: randomUUID(), lineNumber: 0, category: "weld" as any,
        description: `BW for ${size} TEE (auto-inferred)`,
        size, quantity: qty * 3, unit: "EA",
        materialUnitCost: 0, laborUnitCost: 0, laborHoursPerUnit: 0,
        materialExtension: 0, laborExtension: 0, totalCost: 0,
        notes: "Auto-inferred: 3 butt welds per tee", fromDatabase: false,
        weldAssumption: "3 butt welds per tee (auto-inferred)",
      }));
    } else if (cat === "reducer" || desc.includes("reducer") || desc.includes("swage")) {
      welds.push(computeEstimateItem({
        id: randomUUID(), lineNumber: 0, category: "weld" as any,
        description: `BW for ${size} REDUCER (auto-inferred)`,
        size, quantity: qty * 2, unit: "EA",
        materialUnitCost: 0, laborUnitCost: 0, laborHoursPerUnit: 0,
        materialExtension: 0, laborExtension: 0, totalCost: 0,
        notes: "Auto-inferred: 2 butt welds per reducer (larger size)", fromDatabase: false,
        weldAssumption: "2 butt welds per reducer at larger size (auto-inferred)",
      }));
    } else if (cat === "cap" || (desc.includes("cap") && !desc.includes("screw"))) {
      welds.push(computeEstimateItem({
        id: randomUUID(), lineNumber: 0, category: "weld" as any,
        description: `BW for ${size} CAP (auto-inferred)`,
        size, quantity: qty * 1, unit: "EA",
        materialUnitCost: 0, laborUnitCost: 0, laborHoursPerUnit: 0,
        materialExtension: 0, laborExtension: 0, totalCost: 0,
        notes: "Auto-inferred: 1 butt weld per cap", fromDatabase: false,
        weldAssumption: "1 butt weld per cap (auto-inferred)",
      }));
    } else if (cat === "coupling" || desc.includes("coupling")) {
      welds.push(computeEstimateItem({
        id: randomUUID(), lineNumber: 0, category: "weld" as any,
        description: `SW for ${size} COUPLING (auto-inferred)`,
        size, quantity: qty * 2, unit: "EA",
        materialUnitCost: 0, laborUnitCost: 0, laborHoursPerUnit: 0,
        materialExtension: 0, laborExtension: 0, totalCost: 0,
        notes: "Auto-inferred: 2 socket welds per coupling", fromDatabase: false,
        weldAssumption: "2 socket welds per coupling (auto-inferred)",
      }));
    } else if (cat === "flange" || desc.includes("flange")) {
      welds.push(computeEstimateItem({
        id: randomUUID(), lineNumber: 0, category: "weld" as any,
        description: `SO weld for ${size} FLANGE (auto-inferred)`,
        size, quantity: qty * 1, unit: "EA",
        materialUnitCost: 0, laborUnitCost: 0, laborHoursPerUnit: 0,
        materialExtension: 0, laborExtension: 0, totalCost: 0,
        notes: "Auto-inferred: 1 slip-on weld per flange", fromDatabase: false,
        weldAssumption: "1 slip-on weld per flange (auto-inferred)",
      }));
      welds.push(computeEstimateItem({
        id: randomUUID(), lineNumber: 0, category: "bolt" as any,
        description: `Bolt-up for ${size} FLANGE (auto-inferred)`,
        size, quantity: qty * 1, unit: "EA",
        materialUnitCost: 0, laborUnitCost: 0, laborHoursPerUnit: 0,
        materialExtension: 0, laborExtension: 0, totalCost: 0,
        notes: "Auto-inferred: 1 bolt-up per flange", fromDatabase: false,
        weldAssumption: "1 bolt-up per flange (auto-inferred)",
      }));
    }
  }
  return welds;
}

// ============================================================
// ROUTES
// ============================================================

export async function registerRoutes(httpServer: Server, app: Express): Promise<Server> {

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

        res.status(202).json({ jobId, pageCount, totalChunks });

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
    const items = selectedItemIds
      ? (project.items || []).filter(i => selectedItemIds.includes(i.id))
      : project.items;

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
    } = parsed.data;

    let estimatorData: any;
    try {
      estimatorData = getEstimatorData();
    } catch (err: any) {
      return res.status(500).json({ message: `Failed to load estimator data: ${err.message}` });
    }

    // Compute blended rate from ST/OT/DT percentages
    const stPercent = Math.max(0, (100 - overtimePercent - doubleTimePercent) / 100);
    const otPercent = overtimePercent / 100;
    const dtPercent = doubleTimePercent / 100;
    const blendedRate = (laborRate * stPercent) + (overtimeRate * otPercent) + (doubleTimeRate * dtPercent);
    const perDiemPerHour = perDiem / 10; // assuming 10-hr workdays
    const effectiveRate = blendedRate + perDiemPerHour;

    const updatedItems = (project.items || []).map(item => {
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
        const result = calculateBillLaborHours(item, itemMat, itemSched, estimatorData.bill, itemPipeLoc, itemElev, itemAlloy);
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
      } else {
        const jResult = calculateJustinLaborHours(item, lineWorkType as "standard" | "rack", itemMat, itemSched, estimatorData.justin);
        const baseMH = jResult.mh;
        sizeMatchExact = jResult.sizeMatchExact;
        // Apply contingency factor from Justin's data (default 15%)
        const contingencyFactor = estimatorData.justin?.cost_params?.contingency_factor || 0.15;
        const contingencyMult = 1 + contingencyFactor;
        laborHoursPerUnit = baseMH * contingencyMult;
        calcBasis = `${jResult.calcBasis} × ${contingencyMult.toFixed(2)} (${(contingencyFactor * 100).toFixed(0)}% contingency) = ${laborHoursPerUnit.toFixed(4)} MH`;
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

    // Feature 6: Add supervision line item for Justin's method
    if (method === "justin") {
      const hasSupervision = updatedItems.some((i: any) => (i.description || "").toLowerCase().includes("supervision"));
      if (!hasSupervision) {
        const totalHours = updatedItems.reduce((s: number, i: any) => s + (i.quantity || 0) * (i.laborHoursPerUnit || 0), 0);
        const supervisionHoursPerWeek = estimatorData.justin?.cost_params?.supervision_hours_per_week || 60;
        const crewSize = 8;
        const hoursPerDay = 10;
        const projectWeeks = Math.max(1, Math.ceil(totalHours / (crewSize * hoursPerDay * 5)));
        const supervisionMH = projectWeeks * supervisionHoursPerWeek;
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
          calculationBasis: `Justin: Supervision → ${projectWeeks} wk × ${supervisionHoursPerWeek} MH/wk = ${supervisionMH} MH`,
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

  // ===== EXCEL EXPORT =====

  app.get("/api/estimates/:id/export-bill", async (req, res) => {
    const project = storage.getEstimateProject(req.params.id);
    if (!project) return res.status(404).json({ message: "Estimate not found" });
    try {
      const wb = await generateBillsWorkbook(project);
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
      const wb = await generateJustinsWorkbook(project);
      res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
      res.setHeader("Content-Disposition", `attachment; filename="${project.name.replace(/[^a-zA-Z0-9 _-]/g, "")} - Justins Format.xlsx"`);
      await wb.xlsx.write(res);
      res.end();
    } catch (err: any) {
      console.error("Excel export error:", err);
      res.status(500).json({ message: err.message || "Export failed" });
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

  return httpServer;
}
