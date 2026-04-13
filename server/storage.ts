import { randomUUID, createHash, scryptSync, randomBytes, createCipheriv, createDecipheriv } from "crypto";
import Database from "better-sqlite3";
import path from "path";
import fs from "fs";
import type {
  TakeoffProject,
  TakeoffItem,
  EstimateProject,
  EstimateItem,
  CostDatabaseEntry,
  InsertTakeoffProject,
  InsertEstimateProject,
  InsertCostDatabaseEntry,
  JobProgress,
  Markups,
  CompletedProject,
} from "@shared/schema";

// ============================================================
// SQLite persistent storage
// ============================================================

const DB_DIR = path.join(process.cwd(), "data");
if (!fs.existsSync(DB_DIR)) fs.mkdirSync(DB_DIR, { recursive: true });

const DB_PATH = path.join(DB_DIR, "pg-unified.db");
const db = new Database(DB_PATH);

// Enable WAL mode for better concurrent read performance
db.pragma("journal_mode = WAL");
db.pragma("foreign_keys = ON");

// ============================================================
// Create tables
// ============================================================

db.exec(`
  CREATE TABLE IF NOT EXISTS takeoff_projects (
    id TEXT PRIMARY KEY,
    name TEXT NOT NULL,
    fileName TEXT NOT NULL,
    discipline TEXT NOT NULL,
    lineNumber TEXT,
    area TEXT,
    revision TEXT,
    drawingDate TEXT,
    createdAt TEXT NOT NULL,
    rawText TEXT
  );

  CREATE TABLE IF NOT EXISTS takeoff_items (
    id TEXT PRIMARY KEY,
    projectId TEXT NOT NULL,
    lineNumber INTEGER NOT NULL,
    discipline TEXT NOT NULL,
    category TEXT NOT NULL,
    description TEXT NOT NULL,
    size TEXT NOT NULL,
    quantity REAL NOT NULL,
    unit TEXT NOT NULL,
    spec TEXT,
    material TEXT,
    schedule TEXT,
    rating TEXT,
    mark TEXT,
    grade TEXT,
    weight REAL,
    depth TEXT,
    weldType TEXT,
    weldSize TEXT,
    notes TEXT,
    FOREIGN KEY (projectId) REFERENCES takeoff_projects(id) ON DELETE CASCADE
  );

  CREATE TABLE IF NOT EXISTS estimate_projects (
    id TEXT PRIMARY KEY,
    name TEXT NOT NULL,
    projectNumber TEXT DEFAULT '',
    client TEXT DEFAULT '',
    location TEXT DEFAULT '',
    sourceTakeoffId TEXT,
    createdAt TEXT NOT NULL,
    laborRate REAL DEFAULT 56,
    overtimeRate REAL DEFAULT 79,
    doubleTimeRate REAL DEFAULT 100,
    perDiem REAL DEFAULT 75,
    overtimePercent REAL DEFAULT 15,
    doubleTimePercent REAL DEFAULT 2,
    estimateMethod TEXT DEFAULT 'manual',
    markups_json TEXT DEFAULT '{"overhead":10,"profit":10,"tax":8.25,"bond":2}'
  );

  CREATE TABLE IF NOT EXISTS estimate_items (
    id TEXT PRIMARY KEY,
    projectId TEXT NOT NULL,
    lineNumber INTEGER NOT NULL,
    category TEXT NOT NULL,
    description TEXT NOT NULL,
    size TEXT NOT NULL,
    quantity REAL NOT NULL,
    unit TEXT NOT NULL,
    materialUnitCost REAL DEFAULT 0,
    laborUnitCost REAL DEFAULT 0,
    laborHoursPerUnit REAL DEFAULT 0,
    materialExtension REAL DEFAULT 0,
    laborExtension REAL DEFAULT 0,
    totalCost REAL DEFAULT 0,
    notes TEXT DEFAULT '',
    fromDatabase INTEGER DEFAULT 0,
    itemMaterial TEXT,
    itemSchedule TEXT,
    itemElevation TEXT,
    itemPipeLocation TEXT,
    itemAlloyGroup TEXT,
    calculationBasis TEXT,
    sizeMatchExact INTEGER,
    materialCostSource TEXT DEFAULT '',
    FOREIGN KEY (projectId) REFERENCES estimate_projects(id) ON DELETE CASCADE
  );

  CREATE TABLE IF NOT EXISTS cost_database (
    id TEXT PRIMARY KEY,
    description TEXT NOT NULL,
    size TEXT NOT NULL,
    category TEXT NOT NULL,
    unit TEXT NOT NULL,
    materialUnitCost REAL NOT NULL,
    laborUnitCost REAL NOT NULL,
    laborHoursPerUnit REAL NOT NULL,
    lastUpdated TEXT NOT NULL
  );

  CREATE INDEX IF NOT EXISTS idx_takeoff_items_project ON takeoff_items(projectId);
  CREATE INDEX IF NOT EXISTS idx_estimate_items_project ON estimate_items(projectId);

  CREATE TABLE IF NOT EXISTS users (
    id TEXT PRIMARY KEY,
    username TEXT UNIQUE NOT NULL,
    passwordHash TEXT NOT NULL
  );

  CREATE TABLE IF NOT EXISTS estimate_versions (
    id TEXT PRIMARY KEY,
    estimateId TEXT NOT NULL,
    createdAt TEXT NOT NULL,
    items_json TEXT NOT NULL,
    markups_json TEXT NOT NULL,
    laborRate REAL NOT NULL,
    perDiem REAL NOT NULL,
    estimateMethod TEXT NOT NULL,
    notes TEXT DEFAULT '',
    FOREIGN KEY (estimateId) REFERENCES estimate_projects(id) ON DELETE CASCADE
  );

  CREATE INDEX IF NOT EXISTS idx_estimate_versions_estimate ON estimate_versions(estimateId);

  CREATE TABLE IF NOT EXISTS purchase_history (
    id TEXT PRIMARY KEY,
    description TEXT NOT NULL,
    size TEXT,
    category TEXT NOT NULL DEFAULT 'other',
    material TEXT,
    schedule TEXT,
    rating TEXT,
    connectionType TEXT,
    unit TEXT NOT NULL DEFAULT 'EA',
    unitCost REAL NOT NULL DEFAULT 0,
    quantity REAL NOT NULL DEFAULT 1,
    totalCost REAL NOT NULL DEFAULT 0,
    supplier TEXT NOT NULL DEFAULT 'Unknown',
    invoiceNumber TEXT,
    invoiceDate TEXT,
    project TEXT,
    poNumber TEXT,
    notes TEXT,
    sourceFile TEXT,
    createdAt TEXT NOT NULL
  );

  CREATE INDEX IF NOT EXISTS idx_purchase_history_supplier ON purchase_history(supplier);
  CREATE INDEX IF NOT EXISTS idx_purchase_history_category ON purchase_history(category);
  CREATE INDEX IF NOT EXISTS idx_purchase_history_date ON purchase_history(invoiceDate);

  CREATE TABLE IF NOT EXISTS completed_projects (
    id TEXT PRIMARY KEY,
    name TEXT NOT NULL,
    client TEXT,
    location TEXT,
    scopeDescription TEXT NOT NULL,
    startDate TEXT,
    endDate TEXT,
    welderHours REAL DEFAULT 0,
    fitterHours REAL DEFAULT 0,
    helperHours REAL DEFAULT 0,
    foremanHours REAL DEFAULT 0,
    operatorHours REAL DEFAULT 0,
    totalManhours REAL DEFAULT 0,
    materialCost REAL DEFAULT 0,
    laborCost REAL DEFAULT 0,
    totalCost REAL DEFAULT 0,
    peakCrewSize INTEGER,
    durationDays INTEGER,
    tags TEXT,
    notes TEXT,
    createdAt TEXT NOT NULL
  );
`);

// Drawing templates table (Feature 2)
db.exec(`
  CREATE TABLE IF NOT EXISTS drawing_templates (
    id TEXT PRIMARY KEY,
    name TEXT NOT NULL,
    engineeringFirm TEXT,
    bomLayout TEXT,
    columnOrder TEXT,
    commonAbbreviations TEXT,
    sampleOcrText TEXT,
    matchPatterns TEXT,
    extractionNotes TEXT,
    usageCount INTEGER DEFAULT 0,
    createdAt TEXT NOT NULL
  );
`);

// Vendor quotes table (Feature 4)
db.exec(`
  CREATE TABLE IF NOT EXISTS vendor_quotes (
    id TEXT PRIMARY KEY,
    vendorName TEXT NOT NULL,
    quoteNumber TEXT,
    quoteDate TEXT,
    projectName TEXT,
    description TEXT NOT NULL,
    size TEXT,
    category TEXT DEFAULT 'other',
    unit TEXT DEFAULT 'EA',
    unitPrice REAL NOT NULL DEFAULT 0,
    quantity REAL DEFAULT 1,
    totalPrice REAL DEFAULT 0,
    notes TEXT,
    createdAt TEXT NOT NULL
  );
`);

// Bid tracking table
db.exec(`
  CREATE TABLE IF NOT EXISTS bid_tracking (
    id TEXT PRIMARY KEY,
    projectName TEXT NOT NULL,
    client TEXT,
    bidDate TEXT,
    dueDate TEXT,
    bidAmount REAL DEFAULT 0,
    status TEXT DEFAULT 'draft',
    awardAmount REAL,
    competitor TEXT,
    estimateId TEXT,
    notes TEXT,
    createdAt TEXT NOT NULL
  );
`);

// Project folders table
db.exec(`
  CREATE TABLE IF NOT EXISTS project_folders (
    id TEXT PRIMARY KEY,
    name TEXT NOT NULL,
    description TEXT,
    color TEXT DEFAULT '#01696F',
    createdAt TEXT NOT NULL,
    updatedAt TEXT NOT NULL
  );
`);

// Add confidence column to takeoff_items if it doesn't exist
try {
  db.exec(`ALTER TABLE takeoff_items ADD COLUMN confidence TEXT DEFAULT 'high'`);
} catch (e: any) {
  // Column already exists — ignore
}

// Add revisionClouded column to takeoff_items if it doesn't exist
try {
  db.exec(`ALTER TABLE takeoff_items ADD COLUMN revisionClouded INTEGER DEFAULT 0`);
} catch (e: any) {
  // Column already exists — ignore
}

// Add confidenceNotes column to takeoff_items if it doesn't exist
try {
  db.exec(`ALTER TABLE takeoff_items ADD COLUMN confidenceNotes TEXT`);
} catch (e: any) {
  // Column already exists — ignore
}

// Add archived column to takeoff_projects if it doesn't exist
try {
  db.exec(`ALTER TABLE takeoff_projects ADD COLUMN archived INTEGER DEFAULT 0`);
} catch (e: any) {
  // Column already exists — ignore
}

// Add workType column to estimate_items if it doesn't exist
try {
  db.exec(`ALTER TABLE estimate_items ADD COLUMN workType TEXT`);
} catch (e: any) {
  // Column already exists — ignore
}

// Add revisionClouded column to estimate_items if it doesn't exist
try {
  db.exec(`ALTER TABLE estimate_items ADD COLUMN revisionClouded INTEGER DEFAULT 0`);
} catch (e: any) {
  // Column already exists — ignore
}

// Add new labor rate columns to estimate_projects if they don't exist
try {
  db.exec(`ALTER TABLE estimate_projects ADD COLUMN overtimeRate REAL DEFAULT 79`);
} catch (e: any) { /* Column already exists */ }
try {
  db.exec(`ALTER TABLE estimate_projects ADD COLUMN doubleTimeRate REAL DEFAULT 100`);
} catch (e: any) { /* Column already exists */ }
try {
  db.exec(`ALTER TABLE estimate_projects ADD COLUMN overtimePercent REAL DEFAULT 15`);
} catch (e: any) { /* Column already exists */ }
try {
  db.exec(`ALTER TABLE estimate_projects ADD COLUMN doubleTimePercent REAL DEFAULT 2`);
} catch (e: any) { /* Column already exists */ }

// Migrate existing per diem from $/hr to $/day for old projects
// Old default was 7.25/hr, new default is 75/day
try {
  db.exec(`UPDATE estimate_projects SET perDiem = 75 WHERE perDiem = 7.25`);
  db.exec(`UPDATE estimate_projects SET laborRate = 56 WHERE laborRate = 70`);
} catch (e: any) { /* ignore */ }

// Add confidenceScore column to takeoff_items if it doesn't exist
try {
  db.exec(`ALTER TABLE takeoff_items ADD COLUMN confidenceScore INTEGER DEFAULT 100`);
} catch (e: any) {
  // Column already exists — ignore
}

// Add sourcePage column to takeoff_items if it doesn't exist
try {
  db.exec(`ALTER TABLE takeoff_items ADD COLUMN sourcePage INTEGER`);
} catch (e: any) {
  // Column already exists — ignore
}

// Add _dedupCandidate column to takeoff_items if it doesn't exist
try {
  db.exec(`ALTER TABLE takeoff_items ADD COLUMN _dedupCandidate INTEGER DEFAULT 0`);
} catch (e: any) {
  // Column already exists — ignore
}

// Add dedupNote column to takeoff_items if it doesn't exist
try {
  db.exec(`ALTER TABLE takeoff_items ADD COLUMN dedupNote TEXT`);
} catch (e: any) {
  // Column already exists — ignore
}

// Add manuallyVerified column to takeoff_items if it doesn't exist
try {
  db.exec(`ALTER TABLE takeoff_items ADD COLUMN manuallyVerified INTEGER DEFAULT 0`);
} catch (e: any) {
  // Column already exists — ignore
}

// Add weldAssumption column to estimate_items if it doesn't exist
try {
  db.exec(`ALTER TABLE estimate_items ADD COLUMN weldAssumption TEXT`);
} catch (e: any) {
  // Column already exists — ignore
}

// Create corrections table for tracking human edits
db.exec(`
  CREATE TABLE IF NOT EXISTS corrections (
    id TEXT PRIMARY KEY,
    takeoffProjectId TEXT NOT NULL,
    itemId TEXT NOT NULL,
    fieldName TEXT NOT NULL,
    originalValue TEXT,
    correctedValue TEXT,
    correctedBy TEXT DEFAULT 'admin',
    correctedAt TEXT NOT NULL
  );

  CREATE TABLE IF NOT EXISTS auth_sessions (
    token TEXT PRIMARY KEY,
    userId TEXT NOT NULL,
    username TEXT NOT NULL,
    createdAt INTEGER NOT NULL
  );

  CREATE INDEX IF NOT EXISTS idx_auth_sessions_created ON auth_sessions(createdAt);
`);

// Add new columns to estimate_versions for new rate fields
try {
  db.exec(`ALTER TABLE estimate_versions ADD COLUMN overtimeRate REAL DEFAULT 79`);
} catch (e: any) { /* Column already exists */ }
try {
  db.exec(`ALTER TABLE estimate_versions ADD COLUMN doubleTimeRate REAL DEFAULT 100`);
} catch (e: any) { /* Column already exists */ }
try {
  db.exec(`ALTER TABLE estimate_versions ADD COLUMN overtimePercent REAL DEFAULT 15`);
} catch (e: any) { /* Column already exists */ }
try {
  db.exec(`ALTER TABLE estimate_versions ADD COLUMN doubleTimePercent REAL DEFAULT 2`);
} catch (e: any) { /* Column already exists */ }

// Add salt column to users if it doesn't exist
try {
  db.exec(`ALTER TABLE users ADD COLUMN salt TEXT`);
} catch (e: any) {
  // Column already exists — ignore
}

// Add folderId to takeoff_projects for project folders
try {
  db.exec(`ALTER TABLE takeoff_projects ADD COLUMN folderId TEXT REFERENCES project_folders(id) ON DELETE SET NULL`);
} catch (e: any) {
  // Column already exists — ignore
}

// Seed default admin user with scrypt hashing
function hashPassword(pw: string, existingSalt?: string): { hash: string; salt: string } {
  const salt = existingSalt || randomBytes(16).toString("hex");
  const hash = scryptSync(pw, salt, 64).toString("hex");
  return { hash, salt };
}

function verifyPassword(pw: string, storedHash: string, salt: string | null): boolean {
  if (salt) {
    const { hash } = hashPassword(pw, salt);
    return hash === storedHash;
  }
  // Fallback for legacy SHA-256 hashed passwords (migration path)
  const legacyHash = createHash("sha256").update(pw).digest("hex");
  if (legacyHash === storedHash) {
    // Migrate to scrypt on successful legacy login
    const { hash: newHash, salt: newSalt } = hashPassword(pw);
    try {
      db.prepare(`UPDATE users SET passwordHash = ?, salt = ? WHERE passwordHash = ?`).run(newHash, newSalt, storedHash);
    } catch {}
    return true;
  }
  return false;
}

{
  const existingAdmin = db.prepare(`SELECT id, salt FROM users WHERE username = ?`).get("admin") as any;
  if (!existingAdmin) {
    const { hash, salt } = hashPassword("picougroup");
    db.prepare(`INSERT INTO users (id, username, passwordHash, salt) VALUES (?, ?, ?, ?)`).run(
      randomUUID(), "admin", hash, salt
    );
    console.warn("WARNING: Default admin credentials active (admin/picougroup) — change password in Settings immediately!");
  }
}

// Auth sessions — persisted to SQLite (survives server restarts)
const SESSION_EXPIRY_MS = 24 * 60 * 60 * 1000; // 24 hours
const sessionStmts = {
  insert: db.prepare(`INSERT OR REPLACE INTO auth_sessions (token, userId, username, createdAt) VALUES (?, ?, ?, ?)`),
  get: db.prepare(`SELECT * FROM auth_sessions WHERE token = ?`),
  delete: db.prepare(`DELETE FROM auth_sessions WHERE token = ?`),
  cleanup: db.prepare(`DELETE FROM auth_sessions WHERE createdAt < ?`),
  touch: db.prepare(`UPDATE auth_sessions SET createdAt = ? WHERE token = ?`),
};

// API key encryption helpers — generate unique secret per installation if not set via env
function getOrCreateAppSecret(): string {
  if (process.env.APP_SECRET) return process.env.APP_SECRET;
  const secretPath = path.join(process.cwd(), "data", ".app-secret");
  try {
    if (fs.existsSync(secretPath)) {
      return fs.readFileSync(secretPath, "utf-8").trim();
    }
  } catch (e) { console.warn("Could not read app secret file:", e); }
  // Generate a new random secret on first run
  const newSecret = randomBytes(32).toString("hex");
  try {
    fs.writeFileSync(secretPath, newSecret, { mode: 0o600 });
    console.log("Generated new app secret in data/.app-secret");
  } catch (e) { console.warn("Could not write app secret file:", e); }
  return newSecret;
}
const APP_SECRET = getOrCreateAppSecret();
const ENCRYPTION_KEY = scryptSync(APP_SECRET, "pg-unified-salt", 32);
const IV_LENGTH = 16;

function encryptApiKey(plainText: string): string {
  const iv = randomBytes(IV_LENGTH);
  const cipher = createCipheriv("aes-256-cbc", ENCRYPTION_KEY, iv);
  let encrypted = cipher.update(plainText, "utf8", "hex");
  encrypted += cipher.final("hex");
  return iv.toString("hex") + ":" + encrypted;
}

function decryptApiKey(encryptedText: string): string {
  const parts = encryptedText.split(":");
  if (parts.length !== 2) return encryptedText; // fallback for unencrypted keys
  try {
    const iv = Buffer.from(parts[0], "hex");
    const encrypted = parts[1];
    const decipher = createDecipheriv("aes-256-cbc", ENCRYPTION_KEY, iv);
    let decrypted = decipher.update(encrypted, "hex", "utf8");
    decrypted += decipher.final("utf8");
    return decrypted;
  } catch {
    return encryptedText; // fallback for unencrypted keys
  }
}

// Job progress stays in-memory (ephemeral by nature)
const jobProgressMap = new Map<string, JobProgress>();
const jobProgressTimestamps = new Map<string, number>();

// Clean up expired job progress entries every 5 minutes (remove entries older than 1 hour)
setInterval(() => {
  const now = Date.now();
  const ONE_HOUR = 60 * 60 * 1000;
  for (const [jobId, ts] of jobProgressTimestamps) {
    if (now - ts > ONE_HOUR) {
      jobProgressMap.delete(jobId);
      jobProgressTimestamps.delete(jobId);
    }
  }
  // Also clean expired sessions
  try { sessionStmts.cleanup.run(Date.now() - SESSION_EXPIRY_MS); } catch (e) { console.warn("Session cleanup error:", e); }
}, 5 * 60 * 1000);

// Automatic daily backup — alternating between two copies for safety
let backupSlot = 0;
function runAutoBackup(label: string) {
  try {
    backupSlot = (backupSlot + 1) % 2;
    const backupPath = path.join(process.cwd(), "data", `pg-unified-autobackup-${backupSlot + 1}.db`);
    db.backup(backupPath).then(() => {
      console.log(`${label} auto-backup saved to ${backupPath}`);
    }).catch((err: any) => {
      console.warn(`${label} auto-backup failed:`, err?.message || err);
    });
  } catch (e) { console.warn(`${label} auto-backup error:`, e); }
}

setInterval(() => runAutoBackup("Scheduled"), 24 * 60 * 60 * 1000); // Every 24 hours

// Run first backup 5 minutes after startup
setTimeout(() => runAutoBackup("Initial"), 5 * 60 * 1000);

// ============================================================
// Helper: serialize/deserialize
// ============================================================

function serializeTakeoffProject(row: any, items: TakeoffItem[]): TakeoffProject {
  return {
    id: row.id,
    name: row.name,
    fileName: row.fileName,
    discipline: row.discipline,
    lineNumber: row.lineNumber || undefined,
    area: row.area || undefined,
    revision: row.revision || undefined,
    drawingDate: row.drawingDate || undefined,
    createdAt: row.createdAt,
    items,
    archived: row.archived === 1,
  };
}

function rowToTakeoffItem(row: any): TakeoffItem {
  return {
    id: row.id,
    lineNumber: row.lineNumber,
    discipline: row.discipline,
    category: row.category,
    description: row.description,
    size: row.size,
    quantity: row.quantity,
    unit: row.unit,
    spec: row.spec || undefined,
    material: row.material || undefined,
    schedule: row.schedule || undefined,
    rating: row.rating || undefined,
    mark: row.mark || undefined,
    grade: row.grade || undefined,
    weight: row.weight || undefined,
    depth: row.depth || undefined,
    weldType: row.weldType || undefined,
    weldSize: row.weldSize || undefined,
    notes: row.notes || undefined,
    confidence: row.confidence || "high",
    confidenceScore: row.confidenceScore != null ? row.confidenceScore : undefined,
    confidenceNotes: row.confidenceNotes || undefined,
    revisionClouded: row.revisionClouded === 1 || row.revisionClouded === true,
    sourcePage: row.sourcePage != null ? row.sourcePage : undefined,
    _dedupCandidate: row._dedupCandidate === 1 || row._dedupCandidate === true || undefined,
    dedupNote: row.dedupNote || undefined,
    manuallyVerified: row.manuallyVerified === 1 || row.manuallyVerified === true || undefined,
  };
}

function rowToEstimateItem(row: any): EstimateItem {
  return {
    id: row.id,
    lineNumber: row.lineNumber,
    category: row.category,
    description: row.description,
    size: row.size,
    quantity: row.quantity,
    unit: row.unit,
    materialUnitCost: row.materialUnitCost || 0,
    laborUnitCost: row.laborUnitCost || 0,
    laborHoursPerUnit: row.laborHoursPerUnit || 0,
    materialExtension: row.materialExtension || 0,
    laborExtension: row.laborExtension || 0,
    totalCost: row.totalCost || 0,
    notes: row.notes || "",
    fromDatabase: row.fromDatabase === 1,
    itemMaterial: row.itemMaterial || undefined,
    itemSchedule: row.itemSchedule || undefined,
    itemElevation: row.itemElevation || undefined,
    itemPipeLocation: row.itemPipeLocation || undefined,
    itemAlloyGroup: row.itemAlloyGroup || undefined,
    calculationBasis: row.calculationBasis || undefined,
    sizeMatchExact: row.sizeMatchExact != null ? row.sizeMatchExact === 1 : undefined,
    materialCostSource: row.materialCostSource || undefined,
    weldAssumption: row.weldAssumption || undefined,
    workType: row.workType || undefined,
    revisionClouded: row.revisionClouded === 1 || row.revisionClouded === true,
  };
}

function serializeEstimateProject(row: any, items: EstimateItem[]): EstimateProject {
  let markups: Markups = { overhead: 10, profit: 10, tax: 8.25, bond: 2 };
  try {
    if (row.markups_json) markups = JSON.parse(row.markups_json);
  } catch {}
  return {
    id: row.id,
    name: row.name,
    projectNumber: row.projectNumber || "",
    client: row.client || "",
    location: row.location || "",
    sourceTakeoffId: row.sourceTakeoffId || undefined,
    createdAt: row.createdAt,
    items,
    markups,
    laborRate: row.laborRate ?? 56,
    overtimeRate: row.overtimeRate ?? 79,
    doubleTimeRate: row.doubleTimeRate ?? 100,
    perDiem: row.perDiem ?? 75,
    overtimePercent: row.overtimePercent ?? 15,
    doubleTimePercent: row.doubleTimePercent ?? 2,
    estimateMethod: row.estimateMethod || "manual",
  };
}

// ============================================================
// Prepared statements
// ============================================================

const stmts = {
  insertTakeoffProject: db.prepare(`INSERT INTO takeoff_projects (id, name, fileName, discipline, lineNumber, area, revision, drawingDate, createdAt) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)`),
  insertTakeoffItem: db.prepare(`INSERT INTO takeoff_items (id, projectId, lineNumber, discipline, category, description, size, quantity, unit, spec, material, schedule, rating, mark, grade, weight, depth, weldType, weldSize, notes, confidence, revisionClouded, confidenceNotes, confidenceScore, sourcePage, _dedupCandidate, dedupNote) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)`),
  getTakeoffProjects: db.prepare(`SELECT * FROM takeoff_projects ORDER BY createdAt DESC`),
  getTakeoffProjectsByDiscipline: db.prepare(`SELECT * FROM takeoff_projects WHERE discipline = ? ORDER BY createdAt DESC`),
  getTakeoffProjectItemCount: db.prepare(`SELECT projectId, COUNT(*) as itemCount FROM takeoff_items GROUP BY projectId`),
  updateTakeoffProjectMetadata: db.prepare(`UPDATE takeoff_projects SET lineNumber = COALESCE(?, lineNumber), area = COALESCE(?, area), revision = COALESCE(?, revision), drawingDate = COALESCE(?, drawingDate) WHERE id = ?`),
  archiveTakeoffProject: db.prepare(`UPDATE takeoff_projects SET archived = ? WHERE id = ?`),
  getTakeoffProject: db.prepare(`SELECT * FROM takeoff_projects WHERE id = ?`),
  getTakeoffItems: db.prepare(`SELECT * FROM takeoff_items WHERE projectId = ? ORDER BY lineNumber`),
  deleteTakeoffProject: db.prepare(`DELETE FROM takeoff_projects WHERE id = ?`),
  deleteTakeoffItems: db.prepare(`DELETE FROM takeoff_items WHERE projectId = ?`),

  insertEstimateProject: db.prepare(`INSERT INTO estimate_projects (id, name, projectNumber, client, location, sourceTakeoffId, createdAt, laborRate, overtimeRate, doubleTimeRate, perDiem, overtimePercent, doubleTimePercent, estimateMethod, markups_json) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)`),
  insertEstimateItem: db.prepare(`INSERT INTO estimate_items (id, projectId, lineNumber, category, description, size, quantity, unit, materialUnitCost, laborUnitCost, laborHoursPerUnit, materialExtension, laborExtension, totalCost, notes, fromDatabase, itemMaterial, itemSchedule, itemElevation, itemPipeLocation, itemAlloyGroup, calculationBasis, sizeMatchExact, materialCostSource, workType, revisionClouded, weldAssumption) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)`),
  getEstimateProjects: db.prepare(`SELECT * FROM estimate_projects ORDER BY createdAt DESC`),
  getEstimateProject: db.prepare(`SELECT * FROM estimate_projects WHERE id = ?`),
  getEstimateProjectItemCount: db.prepare(`SELECT projectId, COUNT(*) as itemCount FROM estimate_items GROUP BY projectId`),
  getEstimateProjectBySourceTakeoff: db.prepare(`SELECT * FROM estimate_projects WHERE sourceTakeoffId = ? LIMIT 1`),
  getEstimateItems: db.prepare(`SELECT * FROM estimate_items WHERE projectId = ? ORDER BY lineNumber`),
  updateEstimateProject: db.prepare(`UPDATE estimate_projects SET name = ?, projectNumber = ?, client = ?, location = ?, laborRate = ?, overtimeRate = ?, doubleTimeRate = ?, perDiem = ?, overtimePercent = ?, doubleTimePercent = ?, estimateMethod = ?, markups_json = ? WHERE id = ?`),
  deleteEstimateProject: db.prepare(`DELETE FROM estimate_projects WHERE id = ?`),
  deleteEstimateItems: db.prepare(`DELETE FROM estimate_items WHERE projectId = ?`),

  insertCostEntry: db.prepare(`INSERT INTO cost_database (id, description, size, category, unit, materialUnitCost, laborUnitCost, laborHoursPerUnit, lastUpdated) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)`),
  getCostDatabase: db.prepare(`SELECT * FROM cost_database ORDER BY description`),
  getCostEntry: db.prepare(`SELECT * FROM cost_database WHERE id = ?`),
  updateCostEntry: db.prepare(`UPDATE cost_database SET description = ?, size = ?, category = ?, unit = ?, materialUnitCost = ?, laborUnitCost = ?, laborHoursPerUnit = ?, lastUpdated = ? WHERE id = ?`),
  deleteCostEntry: db.prepare(`DELETE FROM cost_database WHERE id = ?`),
  getAllCostEntries: db.prepare(`SELECT * FROM cost_database`),

  // Purchase history
  insertPurchaseRecord: db.prepare(`INSERT INTO purchase_history (id, description, size, category, material, schedule, rating, connectionType, unit, unitCost, quantity, totalCost, supplier, invoiceNumber, invoiceDate, project, poNumber, notes, sourceFile, createdAt) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)`),
  getPurchaseRecords: db.prepare(`SELECT * FROM purchase_history ORDER BY invoiceDate DESC, createdAt DESC`),
  getPurchaseRecordsBySupplier: db.prepare(`SELECT * FROM purchase_history WHERE supplier = ? ORDER BY invoiceDate DESC`),
  getPurchaseRecordsByCategory: db.prepare(`SELECT * FROM purchase_history WHERE category = ? ORDER BY invoiceDate DESC`),
  deletePurchaseRecord: db.prepare(`DELETE FROM purchase_history WHERE id = ?`),
  clearPurchaseHistory: db.prepare(`DELETE FROM purchase_history`),
  getPurchaseSuppliers: db.prepare(`SELECT DISTINCT supplier, COUNT(*) as itemCount, SUM(totalCost) as totalSpend FROM purchase_history GROUP BY supplier ORDER BY totalSpend DESC`),
  getPurchaseCategories: db.prepare(`SELECT DISTINCT category, COUNT(*) as itemCount FROM purchase_history GROUP BY category ORDER BY itemCount DESC`),

  // Corrections
  insertCorrection: db.prepare(`INSERT INTO corrections (id, takeoffProjectId, itemId, fieldName, originalValue, correctedValue, correctedBy, correctedAt) VALUES (?, ?, ?, ?, ?, ?, ?, ?)`),
  getCorrectionsByProject: db.prepare(`SELECT * FROM corrections WHERE takeoffProjectId = ? ORDER BY correctedAt DESC`),

  // Completed Projects (Project History)
  insertCompletedProject: db.prepare(`INSERT INTO completed_projects (id, name, client, location, scopeDescription, startDate, endDate, welderHours, fitterHours, helperHours, foremanHours, operatorHours, totalManhours, materialCost, laborCost, totalCost, peakCrewSize, durationDays, tags, notes, createdAt) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)`),
  getCompletedProjects: db.prepare(`SELECT * FROM completed_projects ORDER BY createdAt DESC`),
  getCompletedProject: db.prepare(`SELECT * FROM completed_projects WHERE id = ?`),
  deleteCompletedProject: db.prepare(`DELETE FROM completed_projects WHERE id = ?`),
  searchCompletedProjects: db.prepare(`SELECT * FROM completed_projects WHERE scopeDescription LIKE ? ESCAPE '\\' OR tags LIKE ? ESCAPE '\\' OR name LIKE ? ESCAPE '\\' OR notes LIKE ? ESCAPE '\\' ORDER BY createdAt DESC`),

  // Targeted cost lookup for auto-pricing
  getLatestCostForItem: db.prepare(`SELECT * FROM purchase_history WHERE LOWER(description) LIKE LOWER(?) AND (? = '' OR LOWER(size) LIKE LOWER(?)) ORDER BY invoiceDate DESC, createdAt DESC LIMIT 1`),

  // Bid tracking
  insertBid: db.prepare(`INSERT INTO bid_tracking (id, projectName, client, bidDate, dueDate, bidAmount, status, awardAmount, competitor, estimateId, notes, createdAt) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)`),
  getBids: db.prepare(`SELECT * FROM bid_tracking ORDER BY createdAt DESC`),
  getBid: db.prepare(`SELECT * FROM bid_tracking WHERE id = ?`),
  updateBid: db.prepare(`UPDATE bid_tracking SET projectName = ?, client = ?, bidDate = ?, dueDate = ?, bidAmount = ?, status = ?, awardAmount = ?, competitor = ?, estimateId = ?, notes = ? WHERE id = ?`),
  deleteBid: db.prepare(`DELETE FROM bid_tracking WHERE id = ?`),
  getBidStats: db.prepare(`SELECT status, COUNT(*) as count, SUM(bidAmount) as totalBid, AVG(bidAmount) as avgBid FROM bid_tracking GROUP BY status`),

  // Drawing templates
  insertDrawingTemplate: db.prepare(`INSERT INTO drawing_templates (id, name, engineeringFirm, bomLayout, columnOrder, commonAbbreviations, sampleOcrText, matchPatterns, extractionNotes, usageCount, createdAt) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)`),
  getDrawingTemplates: db.prepare(`SELECT * FROM drawing_templates ORDER BY usageCount DESC`),
  getDrawingTemplate: db.prepare(`SELECT * FROM drawing_templates WHERE id = ?`),
  deleteDrawingTemplate: db.prepare(`DELETE FROM drawing_templates WHERE id = ?`),
  incrementTemplateUsage: db.prepare(`UPDATE drawing_templates SET usageCount = usageCount + 1 WHERE id = ?`),

  // Vendor quotes
  insertVendorQuote: db.prepare(`INSERT INTO vendor_quotes (id, vendorName, quoteNumber, quoteDate, projectName, description, size, category, unit, unitPrice, quantity, totalPrice, notes, createdAt) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)`),
  getVendorQuotes: db.prepare(`SELECT * FROM vendor_quotes ORDER BY createdAt DESC`),
  getVendorQuote: db.prepare(`SELECT * FROM vendor_quotes WHERE id = ?`),
  deleteVendorQuote: db.prepare(`DELETE FROM vendor_quotes WHERE id = ?`),

  // Project folders
  insertFolder: db.prepare(`INSERT INTO project_folders (id, name, description, color, createdAt, updatedAt) VALUES (?, ?, ?, ?, ?, ?)`),
  getFolders: db.prepare(`SELECT pf.*, COUNT(tp.id) as projectCount FROM project_folders pf LEFT JOIN takeoff_projects tp ON tp.folderId = pf.id GROUP BY pf.id ORDER BY pf.createdAt DESC`),
  getFolder: db.prepare(`SELECT * FROM project_folders WHERE id = ?`),
  updateFolder: db.prepare(`UPDATE project_folders SET name = ?, description = ?, color = ?, updatedAt = ? WHERE id = ?`),
  deleteFolder: db.prepare(`DELETE FROM project_folders WHERE id = ?`),
  addProjectToFolder: db.prepare(`UPDATE takeoff_projects SET folderId = ? WHERE id = ?`),
  removeProjectFromFolder: db.prepare(`UPDATE takeoff_projects SET folderId = NULL WHERE id = ?`),
  getProjectsByFolder: db.prepare(`SELECT * FROM takeoff_projects WHERE folderId = ? ORDER BY createdAt DESC`),
  getFolderCombinedItems: db.prepare(`SELECT ti.* FROM takeoff_items ti INNER JOIN takeoff_projects tp ON ti.projectId = tp.id WHERE tp.folderId = ? ORDER BY tp.name, ti.lineNumber`),
};

// Transaction helpers
const insertTakeoffItemsTransaction = db.transaction((projectId: string, items: TakeoffItem[]) => {
  for (const item of items) {
    stmts.insertTakeoffItem.run(
      item.id || randomUUID(), projectId, item.lineNumber, item.discipline || "", item.category, item.description, item.size, item.quantity, item.unit,
      item.spec || null, item.material || null, item.schedule || null, item.rating || null, item.mark || null, item.grade || null, item.weight || null,
      item.depth || null, item.weldType || null, item.weldSize || null, item.notes || null, (item as any).confidence || "high",
      item.revisionClouded ? 1 : 0,
      (item as any).confidenceNotes || null,
      (item as any).confidenceScore != null ? (item as any).confidenceScore : null,
      (item as any).sourcePage != null ? (item as any).sourcePage : null,
      (item as any)._dedupCandidate ? 1 : 0,
      (item as any).dedupNote || null
    );
  }
});

const insertEstimateItemsTransaction = db.transaction((projectId: string, items: EstimateItem[]) => {
  for (const item of items) {
    stmts.insertEstimateItem.run(
      item.id || randomUUID(), projectId, item.lineNumber, item.category, item.description, item.size, item.quantity, item.unit,
      item.materialUnitCost || 0, item.laborUnitCost || 0, item.laborHoursPerUnit || 0,
      item.materialExtension || 0, item.laborExtension || 0, item.totalCost || 0,
      item.notes || "", item.fromDatabase ? 1 : 0,
      item.itemMaterial || null, item.itemSchedule || null, item.itemElevation || null,
      item.itemPipeLocation || null, item.itemAlloyGroup || null,
      item.calculationBasis || null,
      item.sizeMatchExact != null ? (item.sizeMatchExact ? 1 : 0) : null,
      item.materialCostSource || "",
      item.workType || null,
      item.revisionClouded ? 1 : 0,
      (item as any).weldAssumption || null
    );
  }
});

const replaceEstimateItemsTransaction = db.transaction((projectId: string, items: EstimateItem[]) => {
  stmts.deleteEstimateItems.run(projectId);
  insertEstimateItemsTransaction(projectId, items);
});

const replaceTakeoffItemsTransaction = db.transaction((projectId: string, items: TakeoffItem[]) => {
  stmts.deleteTakeoffItems.run(projectId);
  insertTakeoffItemsTransaction(projectId, items);
});

// ============================================================
// Storage class (same method signatures as before)
// ============================================================

class Storage {
  // ---- Takeoff Projects ----

  async createTakeoffProject(data: InsertTakeoffProject & { items?: TakeoffItem[] }): Promise<TakeoffProject> {
    const id = randomUUID();
    const createdAt = new Date().toISOString();
    const items = (data.items || []).map((item, idx) => ({
      ...item,
      id: item.id || randomUUID(),
      lineNumber: item.lineNumber || idx + 1,
      discipline: data.discipline,
    }));

    stmts.insertTakeoffProject.run(id, data.name, data.fileName, data.discipline, data.lineNumber || null, data.area || null, data.revision || null, data.drawingDate || null, createdAt);
    if (items.length > 0) {
      insertTakeoffItemsTransaction(id, items);
    }

    return {
      id, name: data.name, fileName: data.fileName, discipline: data.discipline,
      lineNumber: data.lineNumber, area: data.area, revision: data.revision,
      drawingDate: data.drawingDate, createdAt, items,
    };
  }

  async getTakeoffProjects(discipline?: string): Promise<TakeoffProject[]> {
    // Delegate to lite version to avoid N+1 query on list views
    return this.getTakeoffProjectsLite(discipline);
  }

  // Lite version: returns projects with item count but no items (avoids N+1 query)
  getTakeoffProjectsLite(discipline?: string): any[] {
    const rows = discipline
      ? stmts.getTakeoffProjectsByDiscipline.all(discipline) as any[]
      : stmts.getTakeoffProjects.all() as any[];
    // Build item count map in a single query
    const countRows = stmts.getTakeoffProjectItemCount.all() as { projectId: string; itemCount: number }[];
    const countMap = new Map(countRows.map(r => [r.projectId, r.itemCount]));
    return rows.map(row => ({
      ...serializeTakeoffProject(row, []),
      itemCount: countMap.get(row.id) || 0,
      items: [], // Empty for lite — use getTakeoffProject(id) for items
    }));
  }

  async getTakeoffProject(id: string): Promise<TakeoffProject | undefined> {
    const row = stmts.getTakeoffProject.get(id) as any;
    if (!row) return undefined;
    const items = (stmts.getTakeoffItems.all(id) as any[]).map(rowToTakeoffItem);
    return serializeTakeoffProject(row, items);
  }

  async updateTakeoffProjectItems(id: string, items: TakeoffItem[]): Promise<TakeoffProject | undefined> {
    const row = stmts.getTakeoffProject.get(id) as any;
    if (!row) return undefined;
    replaceTakeoffItemsTransaction(id, items);
    return serializeTakeoffProject(row, items);
  }

  updateTakeoffProjectMetadata(id: string, metadata: { lineNumber?: string; area?: string; revision?: string; drawingDate?: string }): void {
    stmts.updateTakeoffProjectMetadata.run(metadata.lineNumber || null, metadata.area || null, metadata.revision || null, metadata.drawingDate || null, id);
  }

  async archiveTakeoffProject(id: string, archived: boolean): Promise<boolean> {
    const result = stmts.archiveTakeoffProject.run(archived ? 1 : 0, id);
    return result.changes > 0;
  }

  updateTakeoffItem(itemId: string, updates: Record<string, any>): boolean {
    const allowedFields = ["size", "quantity", "description", "category", "unit", "material", "schedule", "spec", "rating", "notes", "confidence", "confidenceScore", "confidenceNotes", "manuallyVerified"];
    const fields: string[] = [];
    const values: any[] = [];
    for (const field of allowedFields) {
      if (updates[field] !== undefined) {
        fields.push(`${field} = ?`);
        values.push(updates[field]);
      }
    }
    if (fields.length === 0) return false;
    values.push(itemId);
    const result = db.prepare(`UPDATE takeoff_items SET ${fields.join(", ")} WHERE id = ?`).run(...values);
    return result.changes > 0;
  }

  async deleteTakeoffProject(id: string): Promise<boolean> {
    // Items deleted via CASCADE
    const result = stmts.deleteTakeoffProject.run(id);
    return result.changes > 0;
  }

  // ---- Estimate Projects ----

  getEstimateProjects(): EstimateProject[] {
    const rows = stmts.getEstimateProjects.all() as any[];
    return rows.map(row => {
      const items = (stmts.getEstimateItems.all(row.id) as any[]).map(rowToEstimateItem);
      return serializeEstimateProject(row, items);
    });
  }

  // Lite version: returns projects with item count but no items (avoids N+1 query)
  getEstimateProjectsLite(): any[] {
    const rows = stmts.getEstimateProjects.all() as any[];
    const countRows = stmts.getEstimateProjectItemCount.all() as { projectId: string; itemCount: number }[];
    const countMap = new Map(countRows.map(r => [r.projectId, r.itemCount]));
    return rows.map(row => ({
      ...serializeEstimateProject(row, []),
      itemCount: countMap.get(row.id) || 0,
      items: [],
    }));
  }

  getEstimateProject(id: string): EstimateProject | undefined {
    const row = stmts.getEstimateProject.get(id) as any;
    if (!row) return undefined;
    const items = (stmts.getEstimateItems.all(id) as any[]).map(rowToEstimateItem);
    return serializeEstimateProject(row, items);
  }

  getEstimateBySourceTakeoff(takeoffId: string): EstimateProject | undefined {
    const row = stmts.getEstimateProjectBySourceTakeoff.get(takeoffId) as any;
    if (!row) return undefined;
    const items = (stmts.getEstimateItems.all(row.id) as any[]).map(rowToEstimateItem);
    return serializeEstimateProject(row, items);
  }

  createEstimateProject(data: InsertEstimateProject & { items?: EstimateItem[] }): EstimateProject {
    const id = randomUUID();
    const createdAt = new Date().toISOString();
    const markups = { overhead: 10, profit: 10, tax: 8.25, bond: 2 };
    const items = (data.items || []).map((item, idx) => ({
      ...item,
      id: item.id || randomUUID(),
      lineNumber: item.lineNumber || idx + 1,
    }));

    stmts.insertEstimateProject.run(id, data.name, data.projectNumber || "", data.client || "", data.location || "", data.sourceTakeoffId || null, createdAt, 56, 79, 100, 75, 15, 2, "manual", JSON.stringify(markups));
    if (items.length > 0) {
      insertEstimateItemsTransaction(id, items);
    }

    return {
      id, name: data.name, projectNumber: data.projectNumber || "", client: data.client || "",
      location: data.location || "", sourceTakeoffId: data.sourceTakeoffId, createdAt, items,
      markups, laborRate: 56, overtimeRate: 79, doubleTimeRate: 100, perDiem: 75,
      overtimePercent: 15, doubleTimePercent: 2, estimateMethod: "manual",
    };
  }

  updateEstimateProject(id: string, data: Partial<EstimateProject>): EstimateProject | undefined {
    const row = stmts.getEstimateProject.get(id) as any;
    if (!row) return undefined;

    let currentMarkups: Markups = { overhead: 10, profit: 10, tax: 8.25, bond: 2 };
    try { if (row.markups_json) currentMarkups = JSON.parse(row.markups_json); } catch {}

    const name = data.name ?? row.name;
    const projectNumber = data.projectNumber ?? row.projectNumber ?? "";
    const client = data.client ?? row.client ?? "";
    const location = data.location ?? row.location ?? "";
    const laborRate = data.laborRate ?? row.laborRate ?? 56;
    const overtimeRate = data.overtimeRate ?? row.overtimeRate ?? 79;
    const doubleTimeRate = data.doubleTimeRate ?? row.doubleTimeRate ?? 100;
    const perDiem = data.perDiem ?? row.perDiem ?? 75;
    const overtimePercent = data.overtimePercent ?? row.overtimePercent ?? 15;
    const doubleTimePercent = data.doubleTimePercent ?? row.doubleTimePercent ?? 2;
    const estimateMethod = data.estimateMethod ?? row.estimateMethod ?? "manual";
    const markups = data.markups ?? currentMarkups;

    stmts.updateEstimateProject.run(name, projectNumber, client, location, laborRate, overtimeRate, doubleTimeRate, perDiem, overtimePercent, doubleTimePercent, estimateMethod, JSON.stringify(markups), id);

    if (data.items) {
      replaceEstimateItemsTransaction(id, data.items);
    }

    const items = data.items ?? (stmts.getEstimateItems.all(id) as any[]).map(rowToEstimateItem);
    return {
      id, name, projectNumber, client, location,
      sourceTakeoffId: row.sourceTakeoffId || undefined,
      createdAt: row.createdAt, items, markups, laborRate, overtimeRate, doubleTimeRate,
      perDiem, overtimePercent, doubleTimePercent,
      estimateMethod: estimateMethod as "bill" | "justin" | "manual",
    };
  }

  deleteEstimateProject(id: string): boolean {
    const result = stmts.deleteEstimateProject.run(id);
    return result.changes > 0;
  }

  // ---- Cost Database ----

  getCostDatabase(): CostDatabaseEntry[] {
    return (stmts.getCostDatabase.all() as any[]).map(row => ({
      id: row.id,
      description: row.description,
      size: row.size,
      category: row.category,
      unit: row.unit,
      materialUnitCost: row.materialUnitCost,
      laborUnitCost: row.laborUnitCost,
      laborHoursPerUnit: row.laborHoursPerUnit,
      lastUpdated: row.lastUpdated,
    }));
  }

  getCostEntry(id: string): CostDatabaseEntry | undefined {
    const row = stmts.getCostEntry.get(id) as any;
    if (!row) return undefined;
    return {
      id: row.id, description: row.description, size: row.size, category: row.category,
      unit: row.unit, materialUnitCost: row.materialUnitCost, laborUnitCost: row.laborUnitCost,
      laborHoursPerUnit: row.laborHoursPerUnit, lastUpdated: row.lastUpdated,
    };
  }

  addCostEntry(data: InsertCostDatabaseEntry): CostDatabaseEntry {
    const id = randomUUID();
    const lastUpdated = new Date().toISOString();
    stmts.insertCostEntry.run(id, data.description, data.size, data.category, data.unit, data.materialUnitCost, data.laborUnitCost, data.laborHoursPerUnit, lastUpdated);
    return { ...data, id, lastUpdated };
  }

  addCostEntriesBulk(entries: InsertCostDatabaseEntry[]): number {
    const insertMany = db.transaction((items: InsertCostDatabaseEntry[]) => {
      let count = 0;
      for (const data of items) {
        const id = randomUUID();
        const lastUpdated = new Date().toISOString();
        stmts.insertCostEntry.run(id, data.description, data.size, data.category, data.unit, data.materialUnitCost, data.laborUnitCost, data.laborHoursPerUnit, lastUpdated);
        count++;
      }
      return count;
    });
    return insertMany(entries);
  }

  updateCostEntry(id: string, data: Partial<CostDatabaseEntry>): CostDatabaseEntry | undefined {
    const existing = stmts.getCostEntry.get(id) as any;
    if (!existing) return undefined;
    const lastUpdated = new Date().toISOString();
    const merged = {
      description: data.description ?? existing.description,
      size: data.size ?? existing.size,
      category: data.category ?? existing.category,
      unit: data.unit ?? existing.unit,
      materialUnitCost: data.materialUnitCost ?? existing.materialUnitCost,
      laborUnitCost: data.laborUnitCost ?? existing.laborUnitCost,
      laborHoursPerUnit: data.laborHoursPerUnit ?? existing.laborHoursPerUnit,
    };
    stmts.updateCostEntry.run(merged.description, merged.size, merged.category, merged.unit, merged.materialUnitCost, merged.laborUnitCost, merged.laborHoursPerUnit, lastUpdated, id);
    return { id, ...merged, lastUpdated };
  }

  deleteCostEntry(id: string): boolean {
    const result = stmts.deleteCostEntry.run(id);
    return result.changes > 0;
  }

  matchCostEntries(items: { description: string; size: string }[]): Record<string, CostDatabaseEntry> {
    const result: Record<string, CostDatabaseEntry> = {};
    const dbEntries = stmts.getAllCostEntries.all() as any[];

    for (const item of items) {
      const key = `${item.description.toLowerCase().trim()}|${item.size.toLowerCase().trim()}`;
      const match = dbEntries.find(
        (e: any) =>
          e.description.toLowerCase().trim() === item.description.toLowerCase().trim() &&
          e.size.toLowerCase().trim() === item.size.toLowerCase().trim()
      );
      if (match) {
        result[key] = {
          id: match.id, description: match.description, size: match.size,
          category: match.category, unit: match.unit, materialUnitCost: match.materialUnitCost,
          laborUnitCost: match.laborUnitCost, laborHoursPerUnit: match.laborHoursPerUnit,
          lastUpdated: match.lastUpdated,
        };
      }
    }
    return result;
  }

  // ---- Purchase History ----

  getPurchaseRecords(filters?: { supplier?: string; category?: string }): any[] {
    let rows: any[];
    if (filters?.supplier) {
      rows = stmts.getPurchaseRecordsBySupplier.all(filters.supplier) as any[];
    } else if (filters?.category) {
      rows = stmts.getPurchaseRecordsByCategory.all(filters.category) as any[];
    } else {
      rows = stmts.getPurchaseRecords.all() as any[];
    }
    return rows.map(row => ({
      id: row.id,
      description: row.description,
      size: row.size || undefined,
      category: row.category,
      material: row.material || undefined,
      schedule: row.schedule || undefined,
      rating: row.rating || undefined,
      connectionType: row.connectionType || undefined,
      unit: row.unit,
      unitCost: row.unitCost,
      quantity: row.quantity,
      totalCost: row.totalCost,
      supplier: row.supplier,
      invoiceNumber: row.invoiceNumber || undefined,
      invoiceDate: row.invoiceDate || undefined,
      project: row.project || undefined,
      poNumber: row.poNumber || undefined,
      notes: row.notes || undefined,
      sourceFile: row.sourceFile || undefined,
      createdAt: row.createdAt,
    }));
  }

  addPurchaseRecord(data: any): any {
    const id = randomUUID();
    const createdAt = new Date().toISOString();
    const totalCost = (data.unitCost || 0) * (data.quantity || 1);
    stmts.insertPurchaseRecord.run(
      id, data.description, data.size || null, data.category || 'other',
      data.material || null, data.schedule || null, data.rating || null, data.connectionType || null,
      data.unit || 'EA', data.unitCost || 0, data.quantity || 1, totalCost,
      data.supplier || 'Unknown', data.invoiceNumber || null, data.invoiceDate || null,
      data.project || null, data.poNumber || null, data.notes || null, data.sourceFile || null,
      createdAt
    );
    return { id, ...data, totalCost, createdAt };
  }

  addPurchaseRecordsBulk(records: any[]): number {
    let count = 0;
    const txn = db.transaction(() => {
      for (const data of records) {
        const id = randomUUID();
        const createdAt = new Date().toISOString();
        const totalCost = (data.unitCost || 0) * (data.quantity || 1);
        stmts.insertPurchaseRecord.run(
          id, data.description, data.size || null, data.category || 'other',
          data.material || null, data.schedule || null, data.rating || null, data.connectionType || null,
          data.unit || 'EA', data.unitCost || 0, data.quantity || 1, totalCost,
          data.supplier || 'Unknown', data.invoiceNumber || null, data.invoiceDate || null,
          data.project || null, data.poNumber || null, data.notes || null, data.sourceFile || null,
          createdAt
        );
        count++;
      }
    });
    txn();
    return count;
  }

  deletePurchaseRecord(id: string): boolean {
    return stmts.deletePurchaseRecord.run(id).changes > 0;
  }

  clearPurchaseHistory(): number {
    return stmts.clearPurchaseHistory.run().changes;
  }

  getPurchaseSuppliers(): any[] {
    return stmts.getPurchaseSuppliers.all() as any[];
  }

  getPurchaseCategories(): any[] {
    return stmts.getPurchaseCategories.all() as any[];
  }

  getLatestCostForItem(description: string, size?: string): any | undefined {
    // Find the most recent purchase matching description (and optionally size) via SQL
    const descPat = `%${description.trim()}%`;
    const sizePat = size ? `%${size.trim()}%` : '';
    const row = stmts.getLatestCostForItem.get(descPat, sizePat, sizePat) as any;
    return row || undefined;
  }

  // ---- Completed Projects (Project History) ----

  getCompletedProjects(): CompletedProject[] {
    return (stmts.getCompletedProjects.all() as any[]).map(row => ({
      id: row.id,
      name: row.name,
      client: row.client || undefined,
      location: row.location || undefined,
      scopeDescription: row.scopeDescription,
      startDate: row.startDate || undefined,
      endDate: row.endDate || undefined,
      welderHours: row.welderHours || 0,
      fitterHours: row.fitterHours || 0,
      helperHours: row.helperHours || 0,
      foremanHours: row.foremanHours || 0,
      operatorHours: row.operatorHours || 0,
      totalManhours: row.totalManhours || 0,
      materialCost: row.materialCost || 0,
      laborCost: row.laborCost || 0,
      totalCost: row.totalCost || 0,
      peakCrewSize: row.peakCrewSize || undefined,
      durationDays: row.durationDays || undefined,
      tags: row.tags || undefined,
      notes: row.notes || undefined,
      createdAt: row.createdAt,
    }));
  }

  getCompletedProject(id: string): CompletedProject | undefined {
    const row = stmts.getCompletedProject.get(id) as any;
    if (!row) return undefined;
    return {
      id: row.id,
      name: row.name,
      client: row.client || undefined,
      location: row.location || undefined,
      scopeDescription: row.scopeDescription,
      startDate: row.startDate || undefined,
      endDate: row.endDate || undefined,
      welderHours: row.welderHours || 0,
      fitterHours: row.fitterHours || 0,
      helperHours: row.helperHours || 0,
      foremanHours: row.foremanHours || 0,
      operatorHours: row.operatorHours || 0,
      totalManhours: row.totalManhours || 0,
      materialCost: row.materialCost || 0,
      laborCost: row.laborCost || 0,
      totalCost: row.totalCost || 0,
      peakCrewSize: row.peakCrewSize || undefined,
      durationDays: row.durationDays || undefined,
      tags: row.tags || undefined,
      notes: row.notes || undefined,
      createdAt: row.createdAt,
    };
  }

  addCompletedProject(data: any): CompletedProject {
    const id = randomUUID();
    const createdAt = new Date().toISOString();
    const totalManhours = (data.welderHours || 0) + (data.fitterHours || 0) + (data.helperHours || 0) + (data.foremanHours || 0) + (data.operatorHours || 0);
    const totalCost = (data.materialCost || 0) + (data.laborCost || 0);
    stmts.insertCompletedProject.run(
      id, data.name, data.client || null, data.location || null,
      data.scopeDescription, data.startDate || null, data.endDate || null,
      data.welderHours || 0, data.fitterHours || 0, data.helperHours || 0,
      data.foremanHours || 0, data.operatorHours || 0, totalManhours,
      data.materialCost || 0, data.laborCost || 0, totalCost,
      data.peakCrewSize || null, data.durationDays || null,
      data.tags || null, data.notes || null, createdAt
    );
    return {
      id, name: data.name, client: data.client || undefined,
      location: data.location || undefined, scopeDescription: data.scopeDescription,
      startDate: data.startDate || undefined, endDate: data.endDate || undefined,
      welderHours: data.welderHours || 0, fitterHours: data.fitterHours || 0,
      helperHours: data.helperHours || 0, foremanHours: data.foremanHours || 0,
      operatorHours: data.operatorHours || 0, totalManhours,
      materialCost: data.materialCost || 0, laborCost: data.laborCost || 0, totalCost,
      peakCrewSize: data.peakCrewSize || undefined, durationDays: data.durationDays || undefined,
      tags: data.tags || undefined, notes: data.notes || undefined, createdAt,
    };
  }

  deleteCompletedProject(id: string): boolean {
    return stmts.deleteCompletedProject.run(id).changes > 0;
  }

  searchCompletedProjects(query: string): CompletedProject[] {
    const escaped = query.replace(/[%_\\]/g, '\\$&');
    const q = `%${escaped}%`;
    return (stmts.searchCompletedProjects.all(q, q, q, q) as any[]).map(row => ({
      id: row.id,
      name: row.name,
      client: row.client || undefined,
      location: row.location || undefined,
      scopeDescription: row.scopeDescription,
      startDate: row.startDate || undefined,
      endDate: row.endDate || undefined,
      welderHours: row.welderHours || 0,
      fitterHours: row.fitterHours || 0,
      helperHours: row.helperHours || 0,
      foremanHours: row.foremanHours || 0,
      operatorHours: row.operatorHours || 0,
      totalManhours: row.totalManhours || 0,
      materialCost: row.materialCost || 0,
      laborCost: row.laborCost || 0,
      totalCost: row.totalCost || 0,
      peakCrewSize: row.peakCrewSize || undefined,
      durationDays: row.durationDays || undefined,
      tags: row.tags || undefined,
      notes: row.notes || undefined,
      createdAt: row.createdAt,
    }));
  }

  // ---- Job Progress (in-memory, ephemeral) ----

  setJobProgress(jobId: string, progress: JobProgress): void {
    jobProgressMap.set(jobId, progress);
    jobProgressTimestamps.set(jobId, Date.now());
  }

  getJobProgress(jobId: string): JobProgress | undefined {
    return jobProgressMap.get(jobId);
  }

  deleteJobProgress(jobId: string): void {
    jobProgressMap.delete(jobId);
    jobProgressTimestamps.delete(jobId);
  }

  // ---- Auth ----

  login(username: string, password: string): { token: string; username: string } | null {
    const row = db.prepare(`SELECT * FROM users WHERE username = ?`).get(username) as any;
    if (!row) return null;
    if (!verifyPassword(password, row.passwordHash, row.salt || null)) return null;
    const token = randomUUID();
    sessionStmts.insert.run(token, row.id, row.username, Date.now());
    return { token, username: row.username };
  }

  logout(token: string): void {
    sessionStmts.delete.run(token);
  }

  validateToken(token: string): { userId: string; username: string } | null {
    const session = sessionStmts.get.get(token) as any;
    if (!session) return null;
    // Check session expiry (sliding window — refreshed on each validation)
    if (Date.now() - session.createdAt > SESSION_EXPIRY_MS) {
      sessionStmts.delete.run(token);
      return null;
    }
    // Sliding expiry: refresh timestamp so active sessions stay alive
    try { sessionStmts.touch.run(Date.now(), token); } catch {}
    return { userId: session.userId, username: session.username };
  }

  cleanupExpiredSessions(): number {
    const cutoff = Date.now() - SESSION_EXPIRY_MS;
    return sessionStmts.cleanup.run(cutoff).changes;
  }

  changePassword(username: string, currentPassword: string, newPassword: string): { success: boolean; error?: string } {
    const row = db.prepare(`SELECT * FROM users WHERE username = ?`).get(username) as any;
    if (!row) return { success: false, error: "User not found" };
    if (!verifyPassword(currentPassword, row.passwordHash, row.salt || null)) {
      return { success: false, error: "Current password is incorrect" };
    }
    if (newPassword.length < 8) {
      return { success: false, error: "New password must be at least 8 characters" };
    }
    const { hash, salt } = hashPassword(newPassword);
    db.prepare(`UPDATE users SET passwordHash = ?, salt = ? WHERE username = ?`).run(hash, salt, username);
    return { success: true };
  }

  // ---- Corrections ----

  addCorrection(data: { takeoffProjectId: string; itemId: string; fieldName: string; originalValue?: string; correctedValue?: string; correctedBy?: string }): any {
    const id = randomUUID();
    const correctedAt = new Date().toISOString();
    stmts.insertCorrection.run(id, data.takeoffProjectId, data.itemId, data.fieldName, data.originalValue || null, data.correctedValue || null, data.correctedBy || "admin", correctedAt);
    return { id, ...data, correctedAt };
  }

  getCorrectionsByProject(projectId: string): any[] {
    return stmts.getCorrectionsByProject.all(projectId) as any[];
  }

  // ---- API Key Encryption ----

  saveEncryptedApiKey(apiKey: string): void {
    const encrypted = encryptApiKey(apiKey);
    fs.writeFileSync(path.join(process.cwd(), "data", ".api-key"), encrypted, "utf-8");
  }

  loadEncryptedApiKey(): string | null {
    try {
      const keyFile = path.join(process.cwd(), "data", ".api-key");
      if (fs.existsSync(keyFile)) {
        const content = fs.readFileSync(keyFile, "utf-8").trim();
        return decryptApiKey(content);
      }
    } catch {}
    return null;
  }

  saveEncryptedGeminiKey(apiKey: string): void {
    const encrypted = encryptApiKey(apiKey);
    fs.writeFileSync(path.join(process.cwd(), "data", ".gemini-key"), encrypted, "utf-8");
  }

  loadEncryptedGeminiKey(): string | null {
    try {
      const keyFile = path.join(process.cwd(), "data", ".gemini-key");
      if (fs.existsSync(keyFile)) {
        const content = fs.readFileSync(keyFile, "utf-8").trim();
        return decryptApiKey(content);
      }
    } catch {}
    return null;
  }

  // ---- Estimate Versions ----

  saveEstimateVersion(estimateId: string, notes?: string): string {
    const project = this.getEstimateProject(estimateId);
    if (!project) throw new Error("Estimate not found");
    const versionId = randomUUID();
    const createdAt = new Date().toISOString();
    db.prepare(`INSERT INTO estimate_versions (id, estimateId, createdAt, items_json, markups_json, laborRate, perDiem, estimateMethod, notes, overtimeRate, doubleTimeRate, overtimePercent, doubleTimePercent) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)`).run(
      versionId, estimateId, createdAt,
      JSON.stringify(project.items), JSON.stringify(project.markups),
      project.laborRate, project.perDiem, project.estimateMethod,
      notes || "",
      project.overtimeRate ?? 79, project.doubleTimeRate ?? 100,
      project.overtimePercent ?? 15, project.doubleTimePercent ?? 2
    );
    return versionId;
  }

  getEstimateVersions(estimateId: string): any[] {
    const rows = db.prepare(
      `SELECT id, estimateId, createdAt, notes, laborRate, perDiem, estimateMethod, items_json
       FROM estimate_versions WHERE estimateId = ? ORDER BY createdAt DESC`
    ).all(estimateId) as any[];
    return rows.map(row => {
      let itemCount = 0;
      try { itemCount = JSON.parse(row.items_json || "[]").length; } catch {}
      const { items_json, ...rest } = row;
      return { ...rest, itemCount };
    });
  }

  restoreEstimateVersion(estimateId: string, versionId: string): boolean {
    const version = db.prepare(`SELECT * FROM estimate_versions WHERE id = ? AND estimateId = ?`).get(versionId, estimateId) as any;
    if (!version) return false;
    const items = JSON.parse(version.items_json);
    const markups = JSON.parse(version.markups_json);
    this.updateEstimateProject(estimateId, {
      items,
      markups,
      laborRate: version.laborRate,
      overtimeRate: version.overtimeRate ?? 79,
      doubleTimeRate: version.doubleTimeRate ?? 100,
      perDiem: version.perDiem,
      overtimePercent: version.overtimePercent ?? 15,
      doubleTimePercent: version.doubleTimePercent ?? 2,
      estimateMethod: version.estimateMethod,
    });
    return true;
  }

  // ---- Backup / Restore ----

  getFullBackup(): any {
    const takeoffProjects = (stmts.getTakeoffProjects.all() as any[]).map(row => {
      const items = (stmts.getTakeoffItems.all(row.id) as any[]).map(rowToTakeoffItem);
      return serializeTakeoffProject(row, items);
    });
    const estimateProjects = (stmts.getEstimateProjects.all() as any[]).map(row => {
      const items = (stmts.getEstimateItems.all(row.id) as any[]).map(rowToEstimateItem);
      return serializeEstimateProject(row, items);
    });
    const costDatabase = this.getCostDatabase();
    const purchaseHistory = this.getPurchaseRecords();
    const completedProjects = this.getCompletedProjects();
    const bids = this.getBids();
    const vendorQuotes = this.getVendorQuotes();
    const drawingTemplates = this.getDrawingTemplates();
    const corrections = db.prepare(`SELECT * FROM corrections`).all();
    const estimateVersions = db.prepare(`SELECT * FROM estimate_versions`).all();
    return {
      version: 1,
      exportedAt: new Date().toISOString(),
      takeoffProjects,
      estimateProjects,
      costDatabase,
      purchaseHistory,
      completedProjects,
      bids,
      vendorQuotes,
      drawingTemplates,
      corrections,
      estimateVersions,
    };
  }

  restoreFromBackup(data: any): { takeoffs: number; estimates: number; costEntries: number } {
    let takeoffs = 0, estimates = 0, costEntries = 0;

    const restoreTransaction = db.transaction(() => {
      // Restore takeoff projects
      if (data.takeoffProjects && Array.isArray(data.takeoffProjects)) {
        for (const project of data.takeoffProjects) {
          const existing = stmts.getTakeoffProject.get(project.id);
          if (!existing) {
            stmts.insertTakeoffProject.run(project.id, project.name, project.fileName, project.discipline, project.lineNumber || null, project.area || null, project.revision || null, project.drawingDate || null, project.createdAt);
            if (project.items && project.items.length > 0) {
              insertTakeoffItemsTransaction(project.id, project.items);
            }
            takeoffs++;
          }
        }
      }

      // Restore estimate projects
      if (data.estimateProjects && Array.isArray(data.estimateProjects)) {
        for (const project of data.estimateProjects) {
          const existing = stmts.getEstimateProject.get(project.id);
          if (!existing) {
            stmts.insertEstimateProject.run(project.id, project.name, project.projectNumber || "", project.client || "", project.location || "", project.sourceTakeoffId || null, project.createdAt, project.laborRate || 56, project.overtimeRate || 79, project.doubleTimeRate || 100, project.perDiem || 75, project.overtimePercent || 15, project.doubleTimePercent || 2, project.estimateMethod || "manual", JSON.stringify(project.markups || { overhead: 10, profit: 10, tax: 8.25, bond: 2 }));
            if (project.items && project.items.length > 0) {
              insertEstimateItemsTransaction(project.id, project.items);
            }
            estimates++;
          }
        }
      }

      // Restore cost database entries
      if (data.costDatabase && Array.isArray(data.costDatabase)) {
        for (const entry of data.costDatabase) {
          const existing = stmts.getCostEntry.get(entry.id);
          if (!existing) {
            stmts.insertCostEntry.run(entry.id, entry.description, entry.size, entry.category, entry.unit, entry.materialUnitCost, entry.laborUnitCost, entry.laborHoursPerUnit, entry.lastUpdated || new Date().toISOString());
            costEntries++;
          }
        }
      }

      // Restore completed projects
      if (data.completedProjects && Array.isArray(data.completedProjects)) {
        for (const p of data.completedProjects) {
          const existing = stmts.getCompletedProject.get(p.id);
          if (!existing) {
            stmts.insertCompletedProject.run(p.id, p.name, p.client, p.location, p.scopeDescription, p.startDate, p.endDate, p.welderHours, p.fitterHours, p.helperHours, p.foremanHours, p.operatorHours, p.totalManhours, p.materialCost, p.laborCost, p.totalCost, p.peakCrewSize, p.durationDays, p.tags, p.notes, p.createdAt);
          }
        }
      }

      // Restore bids
      if (data.bids && Array.isArray(data.bids)) {
        for (const b of data.bids) {
          const existing = stmts.getBid.get(b.id);
          if (!existing) {
            stmts.insertBid.run(b.id, b.projectName, b.client, b.bidDate, b.dueDate, b.bidAmount, b.status, b.awardAmount, b.competitor, b.estimateId, b.notes, b.createdAt);
          }
        }
      }

      // Restore vendor quotes
      if (data.vendorQuotes && Array.isArray(data.vendorQuotes)) {
        for (const q of data.vendorQuotes) {
          const existing = stmts.getVendorQuote.get(q.id);
          if (!existing) {
            stmts.insertVendorQuote.run(q.id, q.vendorName, q.quoteNumber, q.quoteDate, q.projectName, q.description, q.size, q.category, q.unit, q.unitPrice, q.quantity, q.totalPrice, q.notes, q.createdAt);
          }
        }
      }

      // Restore drawing templates
      if (data.drawingTemplates && Array.isArray(data.drawingTemplates)) {
        for (const t of data.drawingTemplates) {
          const existing = stmts.getDrawingTemplate.get(t.id);
          if (!existing) {
            stmts.insertDrawingTemplate.run(t.id, t.name, t.engineeringFirm, t.bomLayout, t.columnOrder, t.commonAbbreviations, t.sampleOcrText, t.matchPatterns, t.extractionNotes, t.usageCount || 0, t.createdAt);
          }
        }
      }

      // Restore corrections
      if (data.corrections && Array.isArray(data.corrections)) {
        for (const c of data.corrections) {
          try {
            stmts.insertCorrection.run(c.id, c.takeoffProjectId, c.itemId, c.fieldName, c.originalValue, c.correctedValue, c.correctedBy, c.correctedAt);
          } catch (e) { console.warn("Suppressed error restoring correction:", e); }
        }
      }

      // Restore estimate versions
      if (data.estimateVersions && Array.isArray(data.estimateVersions)) {
        const insertVersion = db.prepare(`INSERT OR IGNORE INTO estimate_versions (id, estimateId, createdAt, items_json, markups_json, laborRate, perDiem, estimateMethod, notes, overtimeRate, doubleTimeRate, overtimePercent, doubleTimePercent) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)`);
        for (const v of data.estimateVersions) {
          try {
            insertVersion.run(v.id, v.estimateId, v.createdAt, v.items_json, v.markups_json, v.laborRate, v.perDiem, v.estimateMethod, v.notes || "", v.overtimeRate ?? 79, v.doubleTimeRate ?? 100, v.overtimePercent ?? 15, v.doubleTimePercent ?? 2);
          } catch (e) { console.warn("Suppressed error restoring estimate version:", e); }
        }
      }
    });

    restoreTransaction();
    return { takeoffs, estimates, costEntries };
  }

  // ---- Bid Tracking ----

  getBids(): any[] {
    return stmts.getBids.all() as any[];
  }

  getBid(id: string): any | undefined {
    return stmts.getBid.get(id) as any;
  }

  createBid(data: any): any {
    const id = randomUUID();
    const createdAt = new Date().toISOString();
    stmts.insertBid.run(id, data.projectName, data.client || null, data.bidDate || null, data.dueDate || null, data.bidAmount || 0, data.status || "draft", data.awardAmount || null, data.competitor || null, data.estimateId || null, data.notes || null, createdAt);
    return { id, ...data, createdAt };
  }

  updateBid(id: string, data: any): any | undefined {
    const existing = stmts.getBid.get(id) as any;
    if (!existing) return undefined;
    const merged = {
      projectName: data.projectName ?? existing.projectName,
      client: data.client ?? existing.client,
      bidDate: data.bidDate ?? existing.bidDate,
      dueDate: data.dueDate ?? existing.dueDate,
      bidAmount: data.bidAmount ?? existing.bidAmount,
      status: data.status ?? existing.status,
      awardAmount: data.awardAmount ?? existing.awardAmount,
      competitor: data.competitor ?? existing.competitor,
      estimateId: data.estimateId ?? existing.estimateId,
      notes: data.notes ?? existing.notes,
    };
    stmts.updateBid.run(merged.projectName, merged.client, merged.bidDate, merged.dueDate, merged.bidAmount, merged.status, merged.awardAmount, merged.competitor, merged.estimateId, merged.notes, id);
    return { id, ...merged, createdAt: existing.createdAt };
  }

  deleteBid(id: string): boolean {
    const result = stmts.deleteBid.run(id);
    return result.changes > 0;
  }

  getBidStats(): any[] {
    return stmts.getBidStats.all() as any[];
  }

  // ---- Drawing Templates ----

  getDrawingTemplates(): any[] {
    return stmts.getDrawingTemplates.all() as any[];
  }

  getDrawingTemplate(id: string): any | undefined {
    return stmts.getDrawingTemplate.get(id) as any;
  }

  createDrawingTemplate(data: any): any {
    const id = randomUUID();
    const createdAt = new Date().toISOString();
    stmts.insertDrawingTemplate.run(
      id, data.name, data.engineeringFirm || null, data.bomLayout || null,
      data.columnOrder || null, data.commonAbbreviations || null,
      data.sampleOcrText || null, data.matchPatterns || null,
      data.extractionNotes || null, 0, createdAt
    );
    return { id, ...data, usageCount: 0, createdAt };
  }

  deleteDrawingTemplate(id: string): boolean {
    return stmts.deleteDrawingTemplate.run(id).changes > 0;
  }

  incrementTemplateUsage(id: string): void {
    stmts.incrementTemplateUsage.run(id);
  }

  // ---- Vendor Quotes ----

  getVendorQuotes(): any[] {
    return stmts.getVendorQuotes.all() as any[];
  }

  createVendorQuote(data: any): any {
    const id = randomUUID();
    const createdAt = new Date().toISOString();
    const totalPrice = (data.unitPrice || 0) * (data.quantity || 1);
    stmts.insertVendorQuote.run(
      id, data.vendorName, data.quoteNumber || null, data.quoteDate || null,
      data.projectName || null, data.description, data.size || null,
      data.category || 'other', data.unit || 'EA',
      data.unitPrice || 0, data.quantity || 1, totalPrice,
      data.notes || null, createdAt
    );
    return { id, ...data, totalPrice, createdAt };
  }

  createVendorQuotesBulk(records: any[]): number {
    let count = 0;
    const txn = db.transaction(() => {
      for (const data of records) {
        const id = randomUUID();
        const createdAt = new Date().toISOString();
        const totalPrice = (data.unitPrice || 0) * (data.quantity || 1);
        stmts.insertVendorQuote.run(
          id, data.vendorName || 'Unknown', data.quoteNumber || null, data.quoteDate || null,
          data.projectName || null, data.description || 'Unknown', data.size || null,
          data.category || 'other', data.unit || 'EA',
          data.unitPrice || 0, data.quantity || 1, totalPrice,
          data.notes || null, createdAt
        );
        count++;
      }
    });
    txn();
    return count;
  }

  deleteVendorQuote(id: string): boolean {
    return stmts.deleteVendorQuote.run(id).changes > 0;
  }

  // ---- Material Escalation Alerts (Feature 3) ----

  getMaterialAlerts(): any[] {
    const records = this.getPurchaseRecords();
    if (records.length === 0) return [];

    // Group by category + size
    const groups: Record<string, any[]> = {};
    for (const r of records) {
      const key = `${(r.category || 'other').toLowerCase()}|${(r.size || '').toLowerCase()}`;
      if (!groups[key]) groups[key] = [];
      groups[key].push(r);
    }

    const alerts: any[] = [];
    const now = Date.now();

    for (const [key, items] of Object.entries(groups)) {
      // Sort by invoice date descending
      const sorted = items.sort((a, b) => {
        const da = new Date(a.invoiceDate || a.createdAt).getTime();
        const db2 = new Date(b.invoiceDate || b.createdAt).getTime();
        return db2 - da;
      });

      const latest = sorted[0];
      const latestDate = new Date(latest.invoiceDate || latest.createdAt);
      const daysSince = Math.floor((now - latestDate.getTime()) / (1000 * 60 * 60 * 24));

      // Stale check: >90 days old
      if (daysSince > 90) {
        alerts.push({
          description: latest.description,
          size: latest.size || '',
          category: latest.category,
          lastPurchaseDate: latest.invoiceDate || latest.createdAt,
          daysSinceLastPurchase: daysSince,
          latestPrice: latest.unitCost,
          averagePrice: latest.unitCost,
          pctChange: 0,
          alertType: 'stale',
        });
      }

      // Price shift check: if >=2 records and latest differs >15% from average
      if (sorted.length >= 2) {
        const avg = sorted.reduce((s, r) => s + (r.unitCost || 0), 0) / sorted.length;
        const pctChange = avg > 0 ? Math.abs(((latest.unitCost || 0) - avg) / avg) * 100 : 0;
        if (pctChange > 15) {
          alerts.push({
            description: latest.description,
            size: latest.size || '',
            category: latest.category,
            lastPurchaseDate: latest.invoiceDate || latest.createdAt,
            daysSinceLastPurchase: daysSince,
            latestPrice: latest.unitCost,
            averagePrice: Math.round(avg * 100) / 100,
            pctChange: Math.round(pctChange * 10) / 10,
            alertType: 'price_shift',
          });
        }
      }
    }

    return alerts;
  }

  // ── Project Folders ──

  createFolder(data: { name: string; description?: string; color?: string }): any {
    const id = randomUUID();
    const now = new Date().toISOString();
    stmts.insertFolder.run(id, data.name, data.description || '', data.color || '#01696F', now, now);
    return { id, name: data.name, description: data.description || '', color: data.color || '#01696F', createdAt: now, updatedAt: now, projectCount: 0 };
  }

  getFolders(): any[] {
    return stmts.getFolders.all() as any[];
  }

  getFolder(id: string): any | undefined {
    return stmts.getFolder.get(id) as any | undefined;
  }

  updateFolder(id: string, data: { name?: string; description?: string; color?: string }): any | undefined {
    const existing = this.getFolder(id);
    if (!existing) return undefined;
    const now = new Date().toISOString();
    stmts.updateFolder.run(
      data.name ?? existing.name,
      data.description ?? existing.description,
      data.color ?? existing.color,
      now,
      id
    );
    return { ...existing, ...data, updatedAt: now };
  }

  deleteFolder(id: string): boolean {
    const result = stmts.deleteFolder.run(id);
    return result.changes > 0;
  }

  addProjectToFolder(folderId: string, projectId: string): boolean {
    const result = stmts.addProjectToFolder.run(folderId, projectId);
    return result.changes > 0;
  }

  removeProjectFromFolder(projectId: string): boolean {
    const result = stmts.removeProjectFromFolder.run(projectId);
    return result.changes > 0;
  }

  getProjectsByFolder(folderId: string): any[] {
    const rows = stmts.getProjectsByFolder.all(folderId) as any[];
    return rows.map(row => {
      const items = (stmts.getTakeoffItems.all(row.id) as any[]).map(rowToTakeoffItem);
      return serializeTakeoffProject(row, items);
    });
  }

  getFolderCombinedItems(folderId: string): TakeoffItem[] {
    return (stmts.getFolderCombinedItems.all(folderId) as any[]).map(rowToTakeoffItem);
  }
}

export const storage = new Storage();
