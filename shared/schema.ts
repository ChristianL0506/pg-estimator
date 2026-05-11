import { z } from "zod";

// ============================================================
// TAKEOFF SCHEMAS (Mechanical, Structural, Civil)
// ============================================================

export const takeoffDisciplineSchema = z.enum(["mechanical", "structural", "civil"]);
export type TakeoffDiscipline = z.infer<typeof takeoffDisciplineSchema>;

export const takeoffItemSchema = z.object({
  id: z.string(),
  lineNumber: z.number(),
  discipline: takeoffDisciplineSchema,
  category: z.string(),
  description: z.string(),
  size: z.string(),
  quantity: z.number(),
  unit: z.string(),
  spec: z.string().optional(),
  material: z.string().optional(),
  schedule: z.string().optional(),
  rating: z.string().optional(),
  mark: z.string().optional(),
  grade: z.string().optional(),
  weight: z.number().optional(),
  depth: z.string().optional(),
  weldType: z.string().optional(),
  weldSize: z.string().optional(),
  notes: z.string().optional(),
  confidence: z.enum(["high", "medium", "low"]).optional().default("high"),
  confidenceScore: z.number().optional(),
  confidenceNotes: z.string().optional(),
  revisionClouded: z.boolean().optional().default(false),
  sourcePage: z.number().optional(),
  _dedupCandidate: z.boolean().optional(),
  dedupNote: z.string().optional(),
});
export type TakeoffItem = z.infer<typeof takeoffItemSchema>;

export const takeoffProjectSchema = z.object({
  id: z.string(),
  name: z.string(),
  fileName: z.string(),
  discipline: takeoffDisciplineSchema,
  lineNumber: z.string().optional(),
  area: z.string().optional(),
  revision: z.string().optional(),
  drawingDate: z.string().optional(),
  createdAt: z.string(),
  items: z.array(takeoffItemSchema).default([]),
  summary: z.any().optional(),
  archived: z.boolean().optional().default(false),
});
export type TakeoffProject = z.infer<typeof takeoffProjectSchema>;

export const insertTakeoffProjectSchema = takeoffProjectSchema.omit({ id: true, createdAt: true });
export type InsertTakeoffProject = z.infer<typeof insertTakeoffProjectSchema>;

// ============================================================
// ESTIMATING SCHEMAS
// ============================================================

export const estimateItemSchema = z.object({
  id: z.string(),
  lineNumber: z.number(),
  category: z.enum(["pipe", "elbow", "tee", "reducer", "valve", "flange", "gasket", "bolt", "cap", "coupling", "union", "weld", "support", "strainer", "trap", "fitting", "steel", "concrete", "rebar", "earthwork", "paving", "electrical", "other"]),
  description: z.string(),
  size: z.string(),
  quantity: z.number(),
  unit: z.string(),
  materialUnitCost: z.number().default(0),
  laborUnitCost: z.number().default(0),
  laborHoursPerUnit: z.number().default(0),
  materialExtension: z.number().default(0),
  laborExtension: z.number().default(0),
  totalCost: z.number().default(0),
  notes: z.string().default(""),
  fromDatabase: z.boolean().default(false),
  // Per-line labor assumption overrides (fall back to global when absent)
  itemMaterial: z.enum(["CS", "SS"]).optional(),
  itemSchedule: z.string().optional(),
  itemElevation: z.string().optional(),
  itemPipeLocation: z.string().optional(),
  itemAlloyGroup: z.string().optional(),
  // Calculation breakdown string for tooltip display
  calculationBasis: z.string().optional(),
  // Whether the size lookup was an exact match
  sizeMatchExact: z.boolean().optional(),
  // Material cost provenance: where the material cost came from
  materialCostSource: z.enum(["quoted", "database", "purchase_history", "allowance", "manual", ""]).optional(),
  weldAssumption: z.string().optional(),
  // Per-line work type: "rack" work at elevation takes more labor
  workType: z.enum(["standard", "rack"]).optional(),
  // Whether item was within a revision cloud in the source takeoff
  revisionClouded: z.boolean().optional().default(false),
});
export type EstimateItem = z.infer<typeof estimateItemSchema>;

export const markupsSchema = z.object({
  overhead: z.number().default(10),
  profit: z.number().default(10),
  tax: z.number().default(8.25),
  bond: z.number().default(2),
});
export type Markups = z.infer<typeof markupsSchema>;

export const estimateProjectSchema = z.object({
  id: z.string(),
  name: z.string(),
  projectNumber: z.string().optional().default(""),
  client: z.string().optional().default(""),
  location: z.string().optional().default(""),
  sourceTakeoffId: z.string().optional(),
  createdAt: z.string(),
  items: z.array(estimateItemSchema).default([]),
  markups: markupsSchema.default({ overhead: 10, profit: 10, tax: 8.25, bond: 2 }),
  laborRate: z.number().default(56),
  overtimeRate: z.number().default(79),
  doubleTimeRate: z.number().default(100),
  perDiem: z.number().default(75),
  overtimePercent: z.number().default(15),
  doubleTimePercent: z.number().default(2),
  estimateMethod: z.enum(["bill", "justin", "industry", "manual"]).default("manual"),
  // ID of a saved custom method (CustomEstimatorMethod.id). When set, estimateMethod
  // should be "bill" / "justin" / "industry" indicating the base; customMethodId
  // points at the override profile to layer on top.
  customMethodId: z.string().optional(),
});
export type EstimateProject = z.infer<typeof estimateProjectSchema>;

// ============================================================
// CUSTOM ESTIMATOR METHODS
// User-defined estimator profiles. Each is a clone of an existing base
// method (bill / justin / industry) plus a set of factor overrides that
// layer on top of the base. Stored per-user; selectable on any estimate.
// ============================================================

export const customEstimatorMethodSchema = z.object({
  id: z.string(),
  name: z.string().min(1),
  baseMethod: z.enum(["bill", "justin", "industry"]),
  description: z.string().optional().default(""),
  createdAt: z.string(),
  updatedAt: z.string(),
  // Free-form overrides: nested key path -> replacement value. The
  // override applies to the base method's data tree. Examples:
  //   "labor_factors.welds.4\"Welds.std_mh_per_weld": 3.1
  //   "cost_params.labor_rate_per_hour": 72.0
  // A leaf value can be a number, string, or any JSON value the base supports.
  overrides: z.record(z.string(), z.any()).default({}),
});
export type CustomEstimatorMethod = z.infer<typeof customEstimatorMethodSchema>;

export const insertCustomEstimatorMethodSchema = z.object({
  name: z.string().min(1, "Name is required"),
  baseMethod: z.enum(["bill", "justin", "industry"]),
  description: z.string().optional().default(""),
  overrides: z.record(z.string(), z.any()).optional().default({}),
});
export type InsertCustomEstimatorMethod = z.infer<typeof insertCustomEstimatorMethodSchema>;

export const insertEstimateProjectSchema = z.object({
  name: z.string().min(1, "Project name is required"),
  projectNumber: z.string().optional().default(""),
  client: z.string().optional().default(""),
  location: z.string().optional().default(""),
  sourceTakeoffId: z.string().optional(),
});
export type InsertEstimateProject = z.infer<typeof insertEstimateProjectSchema>;

export const costDatabaseEntrySchema = z.object({
  id: z.string(),
  description: z.string(),
  size: z.string(),
  category: z.string(),
  unit: z.string(),
  materialUnitCost: z.number(),
  laborUnitCost: z.number(),
  laborHoursPerUnit: z.number(),
  lastUpdated: z.string(),
});
export type CostDatabaseEntry = z.infer<typeof costDatabaseEntrySchema>;

export const insertCostDatabaseEntrySchema = costDatabaseEntrySchema.omit({ id: true, lastUpdated: true });
export type InsertCostDatabaseEntry = z.infer<typeof insertCostDatabaseEntrySchema>;

// ============================================================
// PURCHASE HISTORY SCHEMAS
// ============================================================

export const purchaseRecordSchema = z.object({
  id: z.string(),
  description: z.string(),
  size: z.string().optional(),
  category: z.string(),
  material: z.string().optional(),
  schedule: z.string().optional(),
  rating: z.string().optional(),
  connectionType: z.string().optional(),
  unit: z.string(),
  unitCost: z.number(),
  quantity: z.number().default(1),
  totalCost: z.number().default(0),
  supplier: z.string(),
  invoiceNumber: z.string().optional(),
  invoiceDate: z.string().optional(),
  project: z.string().optional(),
  poNumber: z.string().optional(),
  notes: z.string().optional(),
  sourceFile: z.string().optional(),
  createdAt: z.string(),
});
export type PurchaseRecord = z.infer<typeof purchaseRecordSchema>;

export const insertPurchaseRecordSchema = purchaseRecordSchema.omit({ id: true, createdAt: true });
export type InsertPurchaseRecord = z.infer<typeof insertPurchaseRecordSchema>;

// ============================================================
// COMPLETED PROJECT HISTORY SCHEMAS
// ============================================================

export const completedProjectSchema = z.object({
  id: z.string(),
  name: z.string(),
  client: z.string().optional(),
  location: z.string().optional(),
  scopeDescription: z.string(), // e.g. "Install 6\" screw conveyor, run 500' of 8\" SS pipe"
  startDate: z.string().optional(),
  endDate: z.string().optional(),
  // Manhours by trade
  welderHours: z.number().default(0),
  fitterHours: z.number().default(0),
  helperHours: z.number().default(0),
  foremanHours: z.number().default(0),
  operatorHours: z.number().default(0),
  totalManhours: z.number().default(0),
  // Costs
  materialCost: z.number().default(0),
  laborCost: z.number().default(0),
  totalCost: z.number().default(0),
  // Crew info
  peakCrewSize: z.number().optional(),
  durationDays: z.number().optional(),
  // Searchable tags
  tags: z.string().optional(), // comma-separated: "conveyor,stainless,8-inch,tank farm"
  notes: z.string().optional(),
  createdAt: z.string(),
});
export type CompletedProject = z.infer<typeof completedProjectSchema>;

// ============================================================
// BID TRACKING SCHEMAS
// ============================================================

export const bidTrackingSchema = z.object({
  id: z.string(),
  projectName: z.string(),
  client: z.string().optional(),
  bidDate: z.string().optional(),
  dueDate: z.string().optional(),
  bidAmount: z.number().default(0),
  status: z.enum(["draft", "submitted", "won", "lost", "no_bid"]).default("draft"),
  awardAmount: z.number().optional(),
  competitor: z.string().optional(),
  estimateId: z.string().optional(),
  notes: z.string().optional(),
  createdAt: z.string(),
});
export type BidTracking = z.infer<typeof bidTrackingSchema>;

// ============================================================
// INSERT SCHEMAS (for validation on write endpoints)
// ============================================================

export const insertBidSchema = bidTrackingSchema.omit({ id: true, createdAt: true }).partial().required({ projectName: true });
export type InsertBid = z.infer<typeof insertBidSchema>;

export const insertCompletedProjectSchema = completedProjectSchema.omit({ id: true, createdAt: true });
export type InsertCompletedProject = z.infer<typeof insertCompletedProjectSchema>;

export const insertVendorQuoteSchema = z.object({
  vendorName: z.string().min(1),
  quoteNumber: z.string().optional(),
  quoteDate: z.string().optional(),
  projectName: z.string().optional(),
  description: z.string().min(1),
  size: z.string().optional(),
  category: z.string().default("other"),
  unit: z.string().default("EA"),
  unitPrice: z.number().default(0),
  quantity: z.number().default(1),
  totalPrice: z.number().optional(),
  notes: z.string().optional(),
});
export type InsertVendorQuote = z.infer<typeof insertVendorQuoteSchema>;

// Job progress for polling
export const jobProgressSchema = z.object({
  jobId: z.string(),
  status: z.string(),
  phase: z.enum(["uploading", "rendering", "extracting", "verifying", "done", "error"]).optional(),
  chunk: z.number(),
  totalChunks: z.number(),
  pagesProcessed: z.number(),
  totalPages: z.number(),
  itemsFound: z.number(),
  projectId: z.string().optional(),
  error: z.string().optional(),
  warnings: z.array(z.string()).optional(),
  pdfQuality: z.enum(["vector", "clean_scan", "poor_scan"]).optional(),
});
export type JobProgress = z.infer<typeof jobProgressSchema>;
