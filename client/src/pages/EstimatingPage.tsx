import React, { useState, useRef, useMemo, useCallback, useEffect } from "react";
import { useQuery, useMutation } from "@tanstack/react-query";
import { Plus, Trash2, Download, Database, ChevronDown, ChevronRight, ChevronUp, Edit2, Check, X, Search, Calculator, Zap, FileSpreadsheet, Info, Settings2, ArrowUpDown, History, Upload, Wand2, ShoppingCart, AlertCircle, CheckCircle2 } from "lucide-react";
import { Button } from "@/components/ui/button";
import { Input } from "@/components/ui/input";
import { Card, CardContent, CardHeader, CardTitle } from "@/components/ui/card";
import { Badge } from "@/components/ui/badge";
import { Skeleton } from "@/components/ui/skeleton";
import { Tabs, TabsContent, TabsList, TabsTrigger } from "@/components/ui/tabs";
import { Dialog, DialogContent, DialogHeader, DialogTitle, DialogFooter } from "@/components/ui/dialog";
import { Label } from "@/components/ui/label";
import { Select, SelectContent, SelectItem, SelectTrigger, SelectValue } from "@/components/ui/select";
import { Separator } from "@/components/ui/separator";
import { Tooltip, TooltipContent, TooltipProvider, TooltipTrigger } from "@/components/ui/tooltip";
import { Popover, PopoverContent, PopoverTrigger } from "@/components/ui/popover";
import AppLayout from "@/components/AppLayout";
import { useToast } from "@/hooks/use-toast";
import { apiRequest, apiUpload, queryClient } from "@/lib/queryClient";
import { exportEstimatePdf } from "@/lib/pdfExport";
import { generateBidReport } from "@/components/BidReportGenerator";
import { parseQuickEntry } from "@/lib/quickEntryParser";
import RfqModal from "@/components/RfqModal";
import CrewPlanner from "@/components/CrewPlanner";
import ProjectPlanner from "@/components/ProjectPlanner";
import type { EstimateProject, EstimateItem, CostDatabaseEntry } from "@shared/schema";

function fmt$(n: number) { return `$${n.toFixed(2)}`; }

function computeItem(item: EstimateItem): EstimateItem {
  const me = item.quantity * (item.materialUnitCost || 0);
  const le = item.quantity * (item.laborUnitCost || 0);
  return { ...item, materialExtension: me, laborExtension: le, totalCost: me + le };
}

// Compute the verification state of an estimate row so the user can audit at a
// glance whether the calculator handled it correctly. State is derived purely
// from already-persisted fields on the item (calculationBasis, laborHoursPerUnit,
// sizeMatchExact); no extra server roundtrip required.
type RowConfidence = {
  state: "green" | "yellow" | "red" | "stale";
  label: string;
  reasons: string[];
};
function getRowConfidence(item: EstimateItem): RowConfidence {
  const reasons: string[] = [];
  const basis = (item as any).calculationBasis || "";
  const lhpu = item.laborHoursPerUnit || 0;
  const sizeMatchExact = (item as any).sizeMatchExact;

  // Stale (calculator never ran on this row, or labor was zeroed without a basis)
  if (!basis && lhpu === 0) {
    return { state: "stale", label: "Not calculated yet", reasons: ["Click 'Calculate Labor Hours' to populate"] };
  }

  // Red: known problems that mean the labor number can't be trusted
  if (basis.includes("No matching factor") || basis.includes("No matching table entry")) {
    reasons.push("No matching labor factor \u2014 row contributes 0 MH");
  }
  if (basis.includes("calculator error")) reasons.push("Calculator threw an error");
  if (lhpu === 0 && (item.category !== "gasket" && !basis.includes("shop bolt-up"))) {
    reasons.push("Labor hours are zero but item has a calc basis \u2014 worth a manual review");
  }
  if (reasons.length > 0) {
    return { state: "red", label: "Needs review", reasons };
  }

  // Yellow: calculator ran but had to fall back / approximate
  if (sizeMatchExact === false) reasons.push("Used nearest size from the factor table, not an exact match");
  if (basis.includes("\u26A0")) reasons.push("Calculator flagged this row with a warning");
  if (reasons.length > 0) {
    return { state: "yellow", label: "Check", reasons };
  }

  // Green: exact size match, no warnings, non-zero MH
  return { state: "green", label: "OK", reasons: [] };
}

// Project-level mode/BOM mismatch detection. Mirrors the diagnose endpoint's
// logic so we can show a persistent banner without an extra roundtrip.
function detectModeMismatch(
  items: EstimateItem[],
  mode: "bundled" | "separate" | "auto-welds"
): { kind: "double-count" | "missing-welds"; sizes?: string[]; suggestion: "separate" | "bundled" | "auto-welds" } | null {
  const fittingsBySize = new Map<string, number>();
  const weldsBySize = new Map<string, number>();
  for (const it of items) {
    const cat = (it.category || "").toLowerCase();
    const desc = (it.description || "").toLowerCase();
    const isFitting = ["fitting","elbow","tee","reducer","cap","coupling","union"].includes(cat);
    const isWeld = cat === "weld" || desc.includes("butt weld") || /\bbw\b/.test(desc);
    const sz = (it.size || "").trim();
    if (!sz) continue;
    if (isFitting) fittingsBySize.set(sz, (fittingsBySize.get(sz) || 0) + 1);
    else if (isWeld) weldsBySize.set(sz, (weldsBySize.get(sz) || 0) + 1);
  }
  if (mode === "bundled" || mode === "auto-welds") {
    // In both modes the fitting line carries its own weld labor, so any
    // explicit weld rows at the same size are a double-count. Recommend
    // 'separate' if the user actually wants the BOM-driven weld rows to
    // carry the labor.
    const overlap: string[] = [];
    for (const sz of fittingsBySize.keys()) {
      if ((weldsBySize.get(sz) || 0) > 0) overlap.push(sz);
    }
    if (overlap.length > 0) return { kind: "double-count", sizes: overlap, suggestion: "separate" };
    return null;
  }
  // separate mode: any fittings but no welds?
  let fittingCount = 0; let weldCount = 0;
  for (const v of fittingsBySize.values()) fittingCount += v;
  for (const v of weldsBySize.values()) weldCount += v;
  if (fittingCount > 0 && weldCount === 0) return { kind: "missing-welds", suggestion: "auto-welds" };
  return null;
}

const CATEGORIES = ["pipe", "elbow", "tee", "reducer", "valve", "flange", "gasket", "bolt", "cap", "coupling", "union", "weld", "support", "strainer", "trap", "fitting", "steel", "concrete", "rebar", "earthwork", "paving", "electrical", "other"];

export default function EstimatingPage() {
  const { toast } = useToast();
  const [selectedId, setSelectedId] = useState<string | null>(null);
  const [newProjectName, setNewProjectName] = useState("");
  const [showNewProject, setShowNewProject] = useState(false);
  const [mobileSidebarOpen, setMobileSidebarOpen] = useState(false);
  const [quickEntry, setQuickEntry] = useState("");
  const [editingMarkups, setEditingMarkups] = useState(false);
  const [dbSearch, setDbSearch] = useState("");

  // Estimating method state
  const [estMethod, setEstMethod] = useState<"bill" | "justin" | "industry">("justin");
  const [estCustomMethodId, setEstCustomMethodId] = useState<string>("");
  const [showCompareDialog, setShowCompareDialog] = useState(false);
  const [showSaveCustomDialog, setShowSaveCustomDialog] = useState(false);
  const [newCustomMethodName, setNewCustomMethodName] = useState("");
  const [newCustomMethodDescription, setNewCustomMethodDescription] = useState("");
  const [estLaborRate, setEstLaborRate] = useState(56);
  const [estOvertimeRate, setEstOvertimeRate] = useState(79);
  const [estDoubleTimeRate, setEstDoubleTimeRate] = useState(100);
  const [estPerDiem, setEstPerDiem] = useState(75);
  const [estOvertimePercent, setEstOvertimePercent] = useState(15);
  const [estDoubleTimePercent, setEstDoubleTimePercent] = useState(2);
  const [estMaterial, setEstMaterial] = useState<"CS" | "SS">("CS");
  const [estSchedule, setEstSchedule] = useState("STD");
  const [estInstallType, setEstInstallType] = useState<"standard" | "rack">("standard");
  const [estPipeLocation, setEstPipeLocation] = useState("Open Rack");
  const [estElevation, setEstElevation] = useState("0-20ft");
  const [estAlloyGroup, setEstAlloyGroup] = useState("4");
  const [estRackFactor, setEstRackFactor] = useState(1.3);
  // Fitting-weld mode — three options:
  //   "bundled":    fitting MH = weld_factor × weld_end_multiplier (legacy multipliers).
  //   "separate":   fitting MH = weld_factor × 0.15 (handling only); BOM carries weld rows.
  //                 Default because BOM extractor + Infer Welds produce explicit weld rows.
  //   "auto-welds": fitting MH = welds_per_fitting × weld_factor + handling.
  //                 The fitting line itself counts as N welds (elbow=2, tee=3, etc.).
  //                 In this mode you should NOT run Infer Welds — each fitting carries
  //                 its own welds inline. This is what most estimators visualize.
  const [estFittingWeldMode, setEstFittingWeldMode] = useState<"bundled" | "separate" | "auto-welds">("separate");
  const [showDiagnoseDialog, setShowDiagnoseDialog] = useState(false);

  // Version history state
  const [showVersions, setShowVersions] = useState(false);
  // RFQ modal state
  const [showRfq, setShowRfq] = useState(false);

  // Table sorting and filtering state
  type SortField = "lineNumber" | "category" | "size" | "description" | "quantity" | "materialUnitCost" | "laborUnitCost" | "laborHoursPerUnit" | "materialExtension" | "laborExtension" | "totalCost";
  const [sortField, setSortField] = useState<SortField>("lineNumber");
  const [sortDir, setSortDir] = useState<"asc" | "desc">("asc");
  const [itemFilter, setItemFilter] = useState("");

  // Debounced save timer
  const debounceTimerRef = useRef<ReturnType<typeof setTimeout> | null>(null);

  const { data: estimates = [], isLoading } = useQuery<EstimateProject[]>({
    queryKey: ["/api/estimates"],
  });

  const { data: selectedProject } = useQuery<EstimateProject>({
    queryKey: ["/api/estimates", selectedId],
    queryFn: async () => {
      if (!selectedId) throw new Error("No project");
      const res = await apiRequest("GET", `/api/estimates/${selectedId}`);
      return res.json();
    },
    enabled: !!selectedId,
  });

  // When a project loads (or is switched), pick up its saved fittingWeldMode
  // so editing legacy bundled estimates still computes consistently with how
  // they were last saved. Brand-new projects default to "separate" server-side;
  // we never silently flip a saved estimate's mode here.
  useEffect(() => {
    if (selectedProject && (selectedProject as any).fittingWeldMode) {
      setEstFittingWeldMode((selectedProject as any).fittingWeldMode);
    }
  }, [selectedProject?.id]);

  const { data: costDb = [] } = useQuery<CostDatabaseEntry[]>({
    queryKey: ["/api/cost-database"],
  });

  const createMutation = useMutation({
    mutationFn: (name: string) => apiRequest("POST", "/api/estimates", { name }).then(r => r.json()),
    onSuccess: (p: EstimateProject) => {
      queryClient.invalidateQueries({ queryKey: ["/api/estimates"] });
      setSelectedId(p.id);
      setNewProjectName("");
      setShowNewProject(false);
      toast({ title: "Estimate created" });
    },
  });

  const deleteMutation = useMutation({
    mutationFn: (id: string) => apiRequest("DELETE", `/api/estimates/${id}`),
    onSuccess: (_, id) => {
      queryClient.invalidateQueries({ queryKey: ["/api/estimates"] });
      if (selectedId === id) setSelectedId(null);
      toast({ title: "Estimate deleted" });
    },
  });

  const updateItemsMutation = useMutation({
    mutationFn: ({ id, items }: { id: string; items: EstimateItem[] }) =>
      apiRequest("PUT", `/api/estimates/${id}/items`, { items }).then(r => r.json()),
    onSuccess: () => {
      queryClient.invalidateQueries({ queryKey: ["/api/estimates", selectedId] });
    },
  });

  const patchMutation = useMutation({
    mutationFn: ({ id, data }: { id: string; data: any }) =>
      apiRequest("PATCH", `/api/estimates/${id}`, data).then(r => r.json()),
    onSuccess: () => {
      queryClient.invalidateQueries({ queryKey: ["/api/estimates", selectedId] });
      queryClient.invalidateQueries({ queryKey: ["/api/estimates"] });
    },
  });

  const applyDbMutation = useMutation({
    mutationFn: (id: string) => apiRequest("POST", `/api/estimates/${id}/apply-database`).then(r => r.json()),
    onSuccess: () => {
      queryClient.invalidateQueries({ queryKey: ["/api/estimates", selectedId] });
      toast({ title: "Database costs applied" });
    },
  });

  // Custom methods list (saved estimator profiles)
  const customMethodsQuery = useQuery<any[]>({
    queryKey: ["/api/custom-methods"],
    queryFn: () => apiRequest("GET", "/api/custom-methods").then(r => r.json()),
  });
  const customMethods = customMethodsQuery.data || [];

  // Save a new custom method (clone of currently selected base method, no overrides yet)
  const saveCustomMutation = useMutation({
    mutationFn: (data: { name: string; baseMethod: string; description?: string }) =>
      apiRequest("POST", "/api/custom-methods", data).then(r => r.json()),
    onSuccess: (created: any) => {
      queryClient.invalidateQueries({ queryKey: ["/api/custom-methods"] });
      setEstCustomMethodId(created.id);
      setShowSaveCustomDialog(false);
      setNewCustomMethodName("");
      setNewCustomMethodDescription("");
      toast({ title: `Custom method '${created.name}' saved` });
    },
    onError: (err: any) => {
      toast({ title: "Save failed", description: err.message, variant: "destructive" });
    },
  });
  const deleteCustomMutation = useMutation({
    mutationFn: (id: string) => apiRequest("DELETE", `/api/custom-methods/${id}`),
    onSuccess: () => {
      queryClient.invalidateQueries({ queryKey: ["/api/custom-methods"] });
      setEstCustomMethodId("");
      toast({ title: "Custom method deleted" });
    },
  });

  // Compare-methods runner — returns summary + per-line drill-down across all base
  // methods (Bill, Justin, Industry) and any selected custom profiles.
  const compareMutation = useMutation({
    mutationFn: (id: string) =>
      apiRequest("POST", `/api/estimates/${id}/compare-methods`, {
        customMethodIds: customMethods.map(m => m.id),
        laborRate: estLaborRate,
        overtimeRate: estOvertimeRate,
        doubleTimeRate: estDoubleTimeRate,
        perDiem: estPerDiem,
        overtimePercent: estOvertimePercent,
        doubleTimePercent: estDoubleTimePercent,
        material: estMaterial,
        schedule: estSchedule,
        installType: estInstallType,
        pipeLocation: estPipeLocation,
        elevation: estElevation,
        alloyGroup: estAlloyGroup,
        rackFactor: estRackFactor,
        fittingWeldMode: estFittingWeldMode,
      }).then(r => r.json()),
    onError: (err: any) => {
      toast({ title: "Compare failed", description: err.message, variant: "destructive" });
    },
  });

  // Diagnose: re-runs the active method over the BOM and returns per-row
  // formula breakdown + project-level warnings (double-counts, missing factors, etc.)
  const diagnoseMutation = useMutation({
    mutationFn: (id: string) =>
      apiRequest("POST", `/api/estimates/${id}/diagnose`, {
        method: estMethod,
        customMethodId: estCustomMethodId || undefined,
        material: estMaterial,
        schedule: estSchedule,
        installType: estInstallType,
        pipeLocation: estPipeLocation,
        elevation: estElevation,
        alloyGroup: estAlloyGroup,
        rackFactor: estRackFactor,
        fittingWeldMode: estFittingWeldMode,
      }).then(r => r.json()),
    onError: (err: any) => {
      toast({ title: "Diagnose failed", description: err.message, variant: "destructive" });
    },
  });

  const autoCalculateMutation = useMutation({
    mutationFn: (id: string) =>
      apiRequest("POST", `/api/estimates/${id}/auto-calculate`, {
        method: estMethod,
        customMethodId: estCustomMethodId || undefined,
        laborRate: estLaborRate,
        overtimeRate: estOvertimeRate,
        doubleTimeRate: estDoubleTimeRate,
        perDiem: estPerDiem,
        overtimePercent: estOvertimePercent,
        doubleTimePercent: estDoubleTimePercent,
        material: estMaterial,
        schedule: estSchedule,
        installType: estInstallType,
        pipeLocation: estPipeLocation,
        elevation: estElevation,
        alloyGroup: estAlloyGroup,
        rackFactor: estRackFactor,
        fittingWeldMode: estFittingWeldMode,
      }).then(r => r.json()),
    onSuccess: () => {
      queryClient.invalidateQueries({ queryKey: ["/api/estimates", selectedId] });
      const baseLabel = estMethod === "bill" ? "Bill's EI" : estMethod === "industry" ? "Industry Standard (Page)" : "Justin's Factor";
      const customLabel = estCustomMethodId ? ` (custom: ${customMethods.find(m => m.id === estCustomMethodId)?.name || ""})` : "";
      toast({ title: `Labor hours calculated using ${baseLabel}${customLabel} method` });
    },
    onError: (err: any) => {
      toast({ title: "Calculation failed", description: err.message, variant: "destructive" });
    },
  });

  const addDbEntryMutation = useMutation({
    mutationFn: (entry: any) => apiRequest("POST", "/api/cost-database", entry).then(r => r.json()),
    onSuccess: () => { queryClient.invalidateQueries({ queryKey: ["/api/cost-database"] }); toast({ title: "Entry saved to database" }); },
  });

  const deleteDbEntryMutation = useMutation({
    mutationFn: (id: string) => apiRequest("DELETE", `/api/cost-database/${id}`),
    onSuccess: () => { queryClient.invalidateQueries({ queryKey: ["/api/cost-database"] }); },
  });

  // Feature 2: Weld inference mutation
  const inferWeldsMutation = useMutation({
    mutationFn: (id: string) => apiRequest("POST", `/api/estimates/${id}/infer-welds`).then(r => r.json()),
    onSuccess: (data: any) => {
      queryClient.invalidateQueries({ queryKey: ["/api/estimates", selectedId] });
      toast({ title: `Inferred ${data.added || 0} weld/bolt items` });
    },
    onError: (err: any) => { toast({ title: "Weld inference failed", description: err.message, variant: "destructive" }); },
  });

  // Inverse of infer-welds: removes every row whose notes contain "auto-inferred".
  // Lets the user flip a bundled-mode estimate clean in one click after they've
  // run the inferrer earlier (or imported an estimate that had them).
  const stripInferredWeldsMutation = useMutation({
    mutationFn: (id: string) => apiRequest("POST", `/api/estimates/${id}/strip-inferred-welds`).then(r => r.json()),
    onSuccess: (data: any) => {
      queryClient.invalidateQueries({ queryKey: ["/api/estimates", selectedId] });
      queryClient.invalidateQueries({ queryKey: ["/api/estimates"] });
      if (data.removed > 0) toast({ title: `Removed ${data.removed} auto-inferred row${data.removed === 1 ? "" : "s"}` });
      else toast({ title: "No auto-inferred rows to remove" });
    },
    onError: (err: any) => { toast({ title: "Strip failed", description: err.message, variant: "destructive" }); },
  });

  // Feature 3: Version history query
  const { data: versions = [] } = useQuery<any[]>({
    queryKey: ["/api/estimates", selectedId, "versions"],
    queryFn: async () => {
      if (!selectedId) return [];
      const res = await apiRequest("GET", `/api/estimates/${selectedId}/versions`);
      return res.json();
    },
    enabled: !!selectedId && showVersions,
  });

  const restoreVersionMutation = useMutation({
    mutationFn: ({ estId, versionId }: { estId: string; versionId: string }) =>
      apiRequest("POST", `/api/estimates/${estId}/restore-version/${versionId}`).then(r => r.json()),
    onSuccess: () => {
      queryClient.invalidateQueries({ queryKey: ["/api/estimates", selectedId] });
      queryClient.invalidateQueries({ queryKey: ["/api/estimates", selectedId, "versions"] });
      toast({ title: "Version restored" });
    },
  });

  // Feature 4: CSV import mutation
  const csvImportMutation = useMutation({
    mutationFn: async (file: File) => {
      const formData = new FormData();
      formData.append("file", file);
      const res = await apiUpload("/api/cost-database/import", formData);
      return res.json();
    },
    onSuccess: (data: any) => {
      queryClient.invalidateQueries({ queryKey: ["/api/cost-database"] });
      toast({ title: data.message || `Imported ${data.imported} entries` });
    },
    onError: (err: any) => { toast({ title: "Import failed", description: err.message, variant: "destructive" }); },
  });

  const handleQuickEntry = () => {
    if (!quickEntry.trim() || !selectedProject) return;
    const parsed = parseQuickEntry(quickEntry);
    if (!parsed) {
      toast({ title: "Could not parse entry", description: "Try: '3 4\" butt welds' or '100 LF pipe'", variant: "destructive" });
      return;
    }
    const newItem = computeItem({
      id: crypto.randomUUID(),
      lineNumber: (selectedProject.items.length || 0) + 1,
      category: parsed.category as any,
      description: parsed.description,
      size: parsed.size,
      quantity: parsed.qty,
      unit: parsed.unit,
      materialUnitCost: 0,
      laborUnitCost: 0,
      laborHoursPerUnit: 0,
      materialExtension: 0,
      laborExtension: 0,
      totalCost: 0,
      notes: "",
      fromDatabase: false,
    });
    const updatedItems = [...selectedProject.items, newItem];
    updateItemsMutation.mutate({ id: selectedProject.id, items: updatedItems });
    setQuickEntry("");
  };

  const handleDeleteItem = (itemId: string) => {
    if (!selectedProject) return;
    if (!window.confirm("Delete this line item?")) return;
    const items = selectedProject.items.filter(i => i.id !== itemId).map((i, idx) => ({ ...i, lineNumber: idx + 1 }));
    updateItemsMutation.mutate({ id: selectedProject.id, items });
  };

  const handleUpdateItemCost = (itemId: string, field: keyof EstimateItem, value: number) => {
    if (!selectedProject) return;
    const items = selectedProject.items.map(i => i.id === itemId ? computeItem({ ...i, [field]: value }) : i);
    debouncedUpdateItems(selectedProject.id, items);
  };

  const handleUpdateMarkup = (field: string, value: number) => {
    if (!selectedProject) return;
    patchMutation.mutate({ id: selectedProject.id, data: { markups: { ...selectedProject.markups, [field]: value } } });
  };

  // Debounced item update to batch rapid edits
  const updateMutateRef = useRef(updateItemsMutation.mutate);
  updateMutateRef.current = updateItemsMutation.mutate;
  const debouncedUpdateItems = useCallback((id: string, items: EstimateItem[]) => {
    if (debounceTimerRef.current) clearTimeout(debounceTimerRef.current);
    debounceTimerRef.current = setTimeout(() => {
      updateMutateRef.current({ id, items });
    }, 500);
  }, []);

  // Sort toggle handler
  const handleSort = (field: SortField) => {
    if (sortField === field) {
      setSortDir(prev => prev === "asc" ? "desc" : "asc");
    } else {
      setSortField(field);
      setSortDir("asc");
    }
  };

  // Per-line override update
  const handleUpdateItemOverride = (itemId: string, field: string, value: string | boolean | undefined) => {
    if (!selectedProject) return;
    const items = selectedProject.items.map(i => {
      if (i.id !== itemId) return i;
      // For boolean fields, preserve the literal value (including false).
      // For string fields, fall back to undefined when blank.
      const normalized = typeof value === "boolean" ? value : (value || undefined);
      return { ...i, [field]: normalized };
    });
    debouncedUpdateItems(selectedProject.id, items);
  };

  const p = selectedProject;

  // Filtered + sorted items
  const displayItems = useMemo(() => {
    if (!p) return [];
    let items = [...p.items];
    // Filter
    if (itemFilter.trim()) {
      const q = itemFilter.toLowerCase();
      items = items.filter(i =>
        i.description.toLowerCase().includes(q) ||
        i.category.toLowerCase().includes(q) ||
        i.size.toLowerCase().includes(q) ||
        i.notes.toLowerCase().includes(q)
      );
    }
    // Sort
    items.sort((a, b) => {
      let av: any = a[sortField];
      let bv: any = b[sortField];
      if (typeof av === "string") av = av.toLowerCase();
      if (typeof bv === "string") bv = bv.toLowerCase();
      if (av < bv) return sortDir === "asc" ? -1 : 1;
      if (av > bv) return sortDir === "asc" ? 1 : -1;
      return 0;
    });
    return items;
  }, [p?.items, sortField, sortDir, itemFilter]);
  // Totals only include rows flagged includeInEstimate. Rows excluded from the
  // estimate stay visible (greyed out) so the user can verify they intended to
  // exclude them, but they don't contribute to cost or markup math.
  const totalMaterial = p?.items.reduce((s, i) => s + ((i as any).includeInEstimate === false ? 0 : (i.materialExtension || 0)), 0) || 0;
  const bomLabor = p?.items.reduce((s, i) => s + ((i as any).includeInEstimate === false ? 0 : (i.laborExtension || 0)), 0) || 0;
  const bomHours = p?.items.reduce((s, i) => s + ((i as any).includeInEstimate === false ? 0 : i.quantity * (i.laborHoursPerUnit || 0)), 0) || 0;

  // Project-level scope adders — hand-entered hours for hydro, demo, supports,
  // ID tags, supervision etc. that don't show up on the BOM. We compute a
  // blended labor rate (incl. per diem) from the project's stored rate inputs
  // so adders contribute realistic labor cost even before auto-calculate runs.
  const scopeAdders = ((p as any)?.scopeAdders as Array<{ id: string; label: string; mode?: "hours" | "cost"; hours: number; ratePerHour?: number; flatCost?: number; note?: string }> | undefined) || [];
  const stPercent = p ? Math.max(0, (100 - p.overtimePercent - p.doubleTimePercent) / 100) : 0.83;
  const otPercent = p ? p.overtimePercent / 100 : 0.15;
  const dtPercent = p ? p.doubleTimePercent / 100 : 0.02;
  const blendedRate = p ? (p.laborRate * stPercent + p.overtimeRate * otPercent + p.doubleTimeRate * dtPercent) : 56;
  const perDiemPerHour = p ? p.perDiem / 10 : 7.5;
  const effectiveRate = blendedRate + perDiemPerHour;
  // Split adders by mode. "hours" rows feed labor; "cost" rows feed a separate
  // direct-cost line that joins the subtotal (so overhead/profit/bond apply,
  // but tax doesn't — tax is material-only).
  const adderHours = scopeAdders.reduce((s, a) => s + ((a.mode ?? "hours") === "hours" ? (a.hours || 0) : 0), 0);
  const adderLabor = scopeAdders.reduce((s, a) => s + ((a.mode ?? "hours") === "hours" ? (a.hours || 0) * (a.ratePerHour ?? effectiveRate) : 0), 0);
  const adderFlatCost = scopeAdders.reduce((s, a) => s + (a.mode === "cost" ? (a.flatCost || 0) : 0), 0);

  // Totals roll BOM + scope adders together so markups and grand total
  // reflect both. Display tables still separate them for transparency.
  const totalLabor = bomLabor + adderLabor;
  const totalHours = bomHours + adderHours;
  const subtotal = totalMaterial + totalLabor + adderFlatCost;
  const overheadAmt = p ? subtotal * (p.markups.overhead / 100) : 0;
  const profitAmt = p ? (subtotal + overheadAmt) * (p.markups.profit / 100) : 0;
  const taxAmt = p ? totalMaterial * (p.markups.tax / 100) : 0;
  const bondAmt = p ? (subtotal + overheadAmt + profitAmt + taxAmt) * (p.markups.bond / 100) : 0;
  const grandTotal = subtotal + overheadAmt + profitAmt + taxAmt + bondAmt;

  const filteredDb = costDb.filter(e =>
    !dbSearch || e.description.toLowerCase().includes(dbSearch.toLowerCase()) || e.size.toLowerCase().includes(dbSearch.toLowerCase())
  );

  return (
    <AppLayout subtitle="Estimating">
      <div className="flex h-full relative">
        {/* Mobile sidebar toggle FAB */}
        <button
          className="md:hidden fixed bottom-6 right-4 z-50 w-12 h-12 rounded-full bg-primary text-primary-foreground shadow-lg flex items-center justify-center"
          aria-label={mobileSidebarOpen ? "Close estimate list" : "Open estimate list"}
          onClick={() => setMobileSidebarOpen(!mobileSidebarOpen)}
        >
          {mobileSidebarOpen ? <XIcon size={20} /> : <Calculator size={20} />}
        </button>

        {/* Mobile overlay */}
        {mobileSidebarOpen && (
          <div className="fixed inset-0 z-40 bg-black/40 md:hidden" onClick={() => setMobileSidebarOpen(false)} />
        )}

        {/* Left panel */}
        <div className={`${mobileSidebarOpen ? "fixed inset-y-0 left-0 z-50 w-60" : "hidden"} md:relative md:block w-60 shrink-0 border-r border-border flex flex-col bg-card`}>
          <div className="p-3 border-b border-border flex items-center justify-between">
            <h2 className="text-xs font-semibold text-muted-foreground uppercase tracking-wider">Estimates</h2>
            <Button
              size="icon"
              variant="ghost"
              className="h-6 w-6"
              onClick={() => setShowNewProject(true)}
              data-testid="btn-new-estimate"
            >
              <Plus size={13} />
            </Button>
          </div>
          <div className="flex-1 overflow-y-auto">
            {isLoading ? (
              <div className="p-3 space-y-2">{[1, 2].map(i => <Skeleton key={i} className="h-14 rounded" />)}</div>
            ) : estimates.length === 0 ? (
              <div className="p-4 flex flex-col items-center text-center mt-4">
                <div className="w-10 h-10 rounded-full bg-muted flex items-center justify-center mb-2">
                  <Calculator size={16} className="text-muted-foreground" />
                </div>
                <p className="text-xs font-medium text-foreground">No estimates yet</p>
                <button className="text-xs text-primary mt-1 hover:underline" onClick={() => setShowNewProject(true)}>Create one</button>
              </div>
            ) : (
              <div className="p-2 space-y-1">
                {estimates.map(est => {
                  const estTotal = est.items.reduce((s, i) => s + (i.totalCost || 0), 0);
                  return (
                    <div
                      key={est.id}
                      data-testid={`estimate-item-${est.id}`}
                      className={`group flex items-start gap-2 p-2.5 rounded-md cursor-pointer transition-colors
                        ${selectedId === est.id ? "bg-primary/10 border border-primary/20" : "hover:bg-accent"}`}
                      onClick={() => setSelectedId(est.id)}
                    >
                      <div className="min-w-0 flex-1">
                        <p className="text-xs font-medium truncate">{est.name}</p>
                        <p className="text-[10px] text-muted-foreground">{est.items.length} items</p>
                        <p className="text-[10px] text-primary font-medium">{fmt$(estTotal)}</p>
                      </div>
                      <button
                        className="opacity-0 group-hover:opacity-100 text-muted-foreground hover:text-destructive p-0.5 shrink-0"
                        onClick={e => { e.stopPropagation(); if (window.confirm("Delete this estimate? This cannot be undone.")) deleteMutation.mutate(est.id); }}
                        data-testid={`btn-delete-est-${est.id}`}
                      >
                        <Trash2 size={12} />
                      </button>
                    </div>
                  );
                })}
              </div>
            )}
          </div>
        </div>

        {/* Main area */}
        <div className="flex-1 overflow-y-auto">
          <Tabs defaultValue="estimate" className="h-full flex flex-col">
            <div className="border-b border-border px-4 pt-3">
              <TabsList>
                <TabsTrigger value="estimate" className="text-xs" data-testid="tab-estimate">Estimate</TabsTrigger>
                <TabsTrigger value="database" className="text-xs" data-testid="tab-database">Cost Database</TabsTrigger>
              </TabsList>
            </div>

            {/* Estimate Tab */}
            <TabsContent value="estimate" className="flex-1 m-0">
              <div className="p-5 space-y-4">
                {!selectedId ? (
                  <Card className="border-dashed border-card-border shadow-sm">
                    <CardContent className="p-10 flex flex-col items-center justify-center text-center">
                      <div className="w-12 h-12 rounded-full bg-muted flex items-center justify-center mb-3">
                        <Calculator size={20} className="text-muted-foreground" />
                      </div>
                      <p className="text-sm font-medium text-foreground">No estimate selected</p>
                      <p className="text-xs text-muted-foreground mt-1">Select an estimate from the left, or create a new one.</p>
                    </CardContent>
                  </Card>
                ) : !p ? (
                  <Skeleton className="h-64 rounded-lg" />
                ) : (
                  <>
                    {/* Project info */}
                    <div key={p.id} className="grid grid-cols-2 md:grid-cols-4 gap-3">
                      {[
                        { field: "name", label: "Project Name", val: p.name },
                        { field: "projectNumber", label: "Project #", val: p.projectNumber || "" },
                        { field: "client", label: "Client", val: p.client || "" },
                        { field: "location", label: "Location", val: p.location || "" },
                      ].map(({ field, label, val }) => (
                        <div key={field}>
                          <Label className="text-xs text-muted-foreground">{label}</Label>
                          <Input
                            className="mt-1 h-7 text-xs"
                            defaultValue={val}
                            onBlur={e => patchMutation.mutate({ id: p.id, data: { [field]: e.target.value } })}
                            data-testid={`input-proj-${field}`}
                          />
                        </div>
                      ))}
                    </div>

                    {/* Summary Cards */}
                    <div className="grid grid-cols-2 md:grid-cols-4 gap-3">
                      <Card className="border-l-4 border-l-blue-500 shadow-sm">
                        <CardContent className="p-3">
                          <p className="text-[10px] text-muted-foreground uppercase tracking-wider">Material</p>
                          <p className="text-lg font-bold font-mono text-blue-600 dark:text-blue-400">{fmt$(totalMaterial)}</p>
                        </CardContent>
                      </Card>
                      <Card className="border-l-4 border-l-orange-500 shadow-sm">
                        <CardContent className="p-3">
                          <p className="text-[10px] text-muted-foreground uppercase tracking-wider">Labor</p>
                          <p className="text-lg font-bold font-mono text-orange-600 dark:text-orange-400">{fmt$(totalLabor)}</p>
                        </CardContent>
                      </Card>
                      <Card className="border-l-4 border-l-purple-500 shadow-sm">
                        <CardContent className="p-3">
                          <p className="text-[10px] text-muted-foreground uppercase tracking-wider">Manhours</p>
                          <p className="text-lg font-bold font-mono text-purple-600 dark:text-purple-400">{totalHours.toFixed(1)}<span className="text-xs font-normal text-muted-foreground ml-1">hrs</span></p>
                        </CardContent>
                      </Card>
                      <Card className="border-l-4 border-l-teal-500 shadow-sm">
                        <CardContent className="p-3">
                          <p className="text-[10px] text-muted-foreground uppercase tracking-wider">Grand Total</p>
                          <p className="text-lg font-bold font-mono text-teal-600 dark:text-teal-400">{fmt$(grandTotal)}</p>
                        </CardContent>
                      </Card>
                    </div>

                    <Separator />

                    {/* LABOR ASSUMPTIONS — must be set BEFORE estimate */}
                    <Card className="border-primary/30 bg-primary/[0.02]">
                      <CardContent className="p-4">
                        <div className="flex items-center gap-2 mb-3">
                          <Calculator size={14} className="text-primary shrink-0" />
                          <span className="text-xs font-semibold">Labor Assumptions</span>
                          {p.estimateMethod && p.estimateMethod !== "manual" && (
                            <Badge variant="outline" className="ml-auto text-[9px] px-1.5 py-0 text-primary border-primary/30">
                              <Zap size={9} className="mr-1" />
                              Using {p.estimateMethod === "bill" ? "Bill's EI Method" : p.estimateMethod === "industry" ? "Industry Standard (Page)" : "Justin's Factor Method"}
                              {p.customMethodId && customMethods.length > 0 && (() => {
                                const cm = customMethods.find(m => m.id === p.customMethodId);
                                return cm ? ` · ${cm.name}` : "";
                              })()}
                            </Badge>
                          )}
                        </div>

                        {/* Method selector — base methods */}
                        <div className="grid grid-cols-3 gap-2 mb-2">
                          <button
                            onClick={() => setEstMethod("bill")}
                            data-testid="btn-method-bill"
                            className={`text-xs px-2 py-1.5 rounded border transition-colors ${
                              estMethod === "bill"
                                ? "bg-primary text-primary-foreground border-primary"
                                : "border-input hover:bg-accent"
                            }`}
                          >
                            Bill's (EI)
                          </button>
                          <button
                            onClick={() => setEstMethod("justin")}
                            data-testid="btn-method-justin"
                            className={`text-xs px-2 py-1.5 rounded border transition-colors ${
                              estMethod === "justin"
                                ? "bg-primary text-primary-foreground border-primary"
                                : "border-input hover:bg-accent"
                            }`}
                          >
                            Justin's (Factor)
                          </button>
                          <button
                            onClick={() => setEstMethod("industry")}
                            data-testid="btn-method-industry"
                            className={`text-xs px-2 py-1.5 rounded border transition-colors ${
                              estMethod === "industry"
                                ? "bg-primary text-primary-foreground border-primary"
                                : "border-input hover:bg-accent"
                            }`}
                            title="Industry Standard — Page's Estimator's Piping Man-Hour Manual"
                          >
                            Industry (Page)
                          </button>
                        </div>

                        {/* Custom-method profile selector + Save As + Compare buttons */}
                        <div className="flex items-center gap-2 mb-3">
                          <select
                            className="flex-1 text-xs border border-input rounded px-2 py-1 bg-background"
                            value={estCustomMethodId}
                            onChange={e => setEstCustomMethodId(e.target.value)}
                            data-testid="select-custom-method"
                          >
                            <option value="">No custom profile (use base method as-published)</option>
                            {customMethods.map(m => (
                              <option key={m.id} value={m.id}>
                                {m.name} · base: {m.baseMethod}
                              </option>
                            ))}
                          </select>
                          <button
                            onClick={() => { setNewCustomMethodName(""); setNewCustomMethodDescription(""); setShowSaveCustomDialog(true); }}
                            data-testid="btn-save-custom-method"
                            className="text-xs px-2 py-1 rounded border border-input hover:bg-accent"
                            title="Save the currently selected base method as a custom profile you can edit"
                          >
                            Save As…
                          </button>
                          {estCustomMethodId && (
                            <button
                              onClick={() => {
                                if (confirm(`Delete custom method '${customMethods.find(m => m.id === estCustomMethodId)?.name}'?`)) {
                                  deleteCustomMutation.mutate(estCustomMethodId);
                                }
                              }}
                              data-testid="btn-delete-custom-method"
                              className="text-xs px-2 py-1 rounded border border-input hover:bg-destructive hover:text-destructive-foreground"
                            >
                              Delete
                            </button>
                          )}
                          <button
                            onClick={() => { setShowCompareDialog(true); compareMutation.mutate(p.id); }}
                            data-testid="btn-compare-methods"
                            className="text-xs px-2 py-1 rounded border border-primary text-primary hover:bg-primary/10"
                            title="Run this estimate through all methods side-by-side"
                          >
                            Compare All…
                          </button>
                        </div>

                        {/* Labor Rates Grid */}
                        <div className="grid grid-cols-3 gap-2 mb-3">
                          <div>
                            <label className="text-[10px] text-muted-foreground block mb-1">ST Rate ($/hr)</label>
                            <input
                              type="number"
                              step="0.5"
                              className="w-full text-right text-xs border border-input rounded px-2 py-1 bg-background font-mono"
                              value={estLaborRate}
                              onChange={e => setEstLaborRate(parseFloat(e.target.value) || 56)}
                              data-testid="input-labor-rate"
                            />
                          </div>
                          <div>
                            <label className="text-[10px] text-muted-foreground block mb-1">OT Rate ($/hr)</label>
                            <input
                              type="number"
                              step="0.5"
                              className="w-full text-right text-xs border border-input rounded px-2 py-1 bg-background font-mono"
                              value={estOvertimeRate}
                              onChange={e => setEstOvertimeRate(parseFloat(e.target.value) || 79)}
                              data-testid="input-overtime-rate"
                            />
                          </div>
                          <div>
                            <label className="text-[10px] text-muted-foreground block mb-1">DT Rate ($/hr)</label>
                            <input
                              type="number"
                              step="0.5"
                              className="w-full text-right text-xs border border-input rounded px-2 py-1 bg-background font-mono"
                              value={estDoubleTimeRate}
                              onChange={e => setEstDoubleTimeRate(parseFloat(e.target.value) || 100)}
                              data-testid="input-double-time-rate"
                            />
                          </div>
                        </div>

                        {/* Per Diem + OT/DT Percentages */}
                        <div className="grid grid-cols-3 gap-2 mb-3">
                          <div>
                            <label className="text-[10px] text-muted-foreground block mb-1">Per Diem ($/day)</label>
                            <input
                              type="number"
                              step="5"
                              className="w-full text-right text-xs border border-input rounded px-2 py-1 bg-background font-mono"
                              value={estPerDiem}
                              onChange={e => setEstPerDiem(parseFloat(e.target.value) || 0)}
                              data-testid="input-per-diem"
                            />
                            <span className="text-[9px] text-muted-foreground mt-0.5 block">
                              = ${(estPerDiem / 10).toFixed(2)}/hr (10hr day)
                            </span>
                          </div>
                          <div>
                            <label className="text-[10px] text-muted-foreground block mb-1">Expected OT %</label>
                            <input
                              type="number"
                              step="1"
                              min="0"
                              max="100"
                              className="w-full text-right text-xs border border-input rounded px-2 py-1 bg-background font-mono"
                              value={estOvertimePercent}
                              onChange={e => setEstOvertimePercent(parseFloat(e.target.value) || 0)}
                              data-testid="input-overtime-percent"
                            />
                          </div>
                          <div>
                            <label className="text-[10px] text-muted-foreground block mb-1">Expected DT %</label>
                            <input
                              type="number"
                              step="1"
                              min="0"
                              max="100"
                              className="w-full text-right text-xs border border-input rounded px-2 py-1 bg-background font-mono"
                              value={estDoubleTimePercent}
                              onChange={e => setEstDoubleTimePercent(parseFloat(e.target.value) || 0)}
                              data-testid="input-double-time-percent"
                            />
                          </div>
                        </div>

                        {/* Blended / Effective Rate Display */}
                        {(() => {
                          const stPct = (100 - estOvertimePercent - estDoubleTimePercent) / 100;
                          const blended = (estLaborRate * stPct) + (estOvertimeRate * estOvertimePercent / 100) + (estDoubleTimeRate * estDoubleTimePercent / 100);
                          const pdHr = estPerDiem / 10;
                          const effective = blended + pdHr;
                          return (
                            <div className="flex items-center gap-4 px-3 py-2 rounded bg-primary/5 border border-primary/20 mb-3">
                              <div className="text-xs">
                                <span className="text-muted-foreground">Blended Rate:</span>{" "}
                                <span className="font-semibold font-mono text-primary" data-testid="text-blended-rate">${blended.toFixed(2)}/hr</span>
                              </div>
                              <div className="text-xs">
                                <span className="text-muted-foreground">Effective Rate (incl per diem):</span>{" "}
                                <span className="font-semibold font-mono text-primary" data-testid="text-effective-rate">${effective.toFixed(2)}/hr</span>
                              </div>
                            </div>
                          );
                        })()}

                        {/* Method-specific parameters */}
                        <div className="flex flex-wrap items-end gap-3">
                          {/* Bill-specific: Material + Schedule */}
                          {estMethod === "bill" && (
                            <>
                              <div>
                                <label className="text-[10px] text-muted-foreground block mb-1">Material</label>
                                <select
                                  className="text-xs border border-input rounded px-2 py-1 bg-background h-7"
                                  value={estMaterial}
                                  onChange={e => setEstMaterial(e.target.value as "CS" | "SS")}
                                  data-testid="select-material"
                                >
                                  <option value="CS">CS (Carbon Steel)</option>
                                  <option value="SS">SS (Stainless)</option>
                                </select>
                              </div>
                              <div>
                                <label className="text-[10px] text-muted-foreground block mb-1">Schedule</label>
                                <select
                                  className="text-xs border border-input rounded px-2 py-1 bg-background h-7"
                                  value={estSchedule}
                                  onChange={e => setEstSchedule(e.target.value)}
                                  data-testid="select-schedule"
                                >
                                  <option value="STD">STD</option>
                                  <option value="XH">XH</option>
                                  <option value="10">SCH 10</option>
                                  <option value="20">SCH 20</option>
                                  <option value="40">SCH 40</option>
                                  <option value="60">SCH 60</option>
                                  <option value="80">SCH 80</option>
                                  <option value="120">SCH 120</option>
                                  <option value="160/XXH">SCH 160/XXH</option>
                                </select>
                              </div>
                              <div>
                                <label className="text-[10px] text-muted-foreground block mb-1">Pipe Location</label>
                                <select
                                  className="text-xs border border-input rounded px-2 py-1 bg-background h-7"
                                  value={estPipeLocation}
                                  onChange={e => setEstPipeLocation(e.target.value)}
                                  data-testid="select-pipe-location"
                                >
                                  <option value="Sleeper Rack">Sleeper Rack (0.6×)</option>
                                  <option value="Underground">Underground (0.75×)</option>
                                  <option value="Open Rack">Open Rack (0.8×)</option>
                                  <option value="Elevated Rack">Elevated Rack (1.0×)</option>
                                </select>
                              </div>
                              <div>
                                <label className="text-[10px] text-muted-foreground block mb-1">Elevation</label>
                                <select
                                  className="text-xs border border-input rounded px-2 py-1 bg-background h-7"
                                  value={estElevation}
                                  onChange={e => setEstElevation(e.target.value)}
                                  data-testid="select-elevation"
                                >
                                  <option value="0-20ft">0' to 20' (1.0×)</option>
                                  <option value="20-40ft">20' to 40' (1.05×)</option>
                                  <option value="40-80ft">40' to 80' (1.10×)</option>
                                  <option value="80ft+">80'+ (1.20×)</option>
                                </select>
                              </div>
                              {estMaterial === "SS" && (
                                <div>
                                  <label className="text-[10px] text-muted-foreground block mb-1">Alloy Group</label>
                                  <select
                                    className="text-xs border border-input rounded px-2 py-1 bg-background h-7"
                                    value={estAlloyGroup}
                                    onChange={e => setEstAlloyGroup(e.target.value)}
                                    data-testid="select-alloy-group"
                                  >
                                    <option value="4">SS 304/316 (2.0×)</option>
                                    <option value="1">Low Chrome Moly (1.75×)</option>
                                    <option value="2">Med Chrome (2.0×)</option>
                                    <option value="3">High Chrome (2.5×)</option>
                                    <option value="7">Hastelloy (3.0×)</option>
                                    <option value="8">Monel/Inconel/Alloy 20 (2.5×)</option>
                                    <option value="9">Aluminum (3.0×)</option>
                                  </select>
                                </div>
                              )}
                            </>
                          )}

                          {/* Justin-specific: Install Type + Rack Factor */}
                          {estMethod === "justin" && (
                            <>
                              <div>
                                <label className="text-[10px] text-muted-foreground block mb-1">Default Install Type</label>
                                <select
                                  className="text-xs border border-input rounded px-2 py-1 bg-background h-7"
                                  value={estInstallType}
                                  onChange={e => setEstInstallType(e.target.value as "standard" | "rack")}
                                  data-testid="select-install-type"
                                >
                                  <option value="standard">Standard</option>
                                  <option value="rack">Rack</option>
                                </select>
                              </div>
                              <div>
                                <label className="text-[10px] text-muted-foreground block mb-1">Rack Factor</label>
                                <input
                                  type="number"
                                  step="0.05"
                                  min="1"
                                  max="5"
                                  className="w-16 text-right text-xs border border-input rounded px-2 py-1 bg-background font-mono h-7"
                                  value={estRackFactor}
                                  onChange={e => setEstRackFactor(parseFloat(e.target.value) || 1.3)}
                                  data-testid="input-rack-factor"
                                />
                              </div>
                            </>
                          )}

                          {/* Fitting Welds mode — only relevant for Bill's method.
                              Justin and Industry always use: Hrs/Unit = welds_per_fitting × weld_factor,
                              with auto-inferred weld rows automatically zeroed to prevent double-count
                              and pipe-length field welds folded into the per-LF rate. */}
                          <div>
                            <label className="text-[10px] text-muted-foreground block mb-1" title={p.estimateMethod === "bill" ? "Bundled: legacy multiplier on fitting. Separate: weld rows in BOM + fitting handling only. Auto-welds: fitting line counts as its N welds (elbow=2, tee=3, etc.) inline." : "Justin and Industry methods always use welds_per_fitting × factor (handling included in factor). Field welds for pipe runs > 40' are added automatically."}>Fitting Welds</label>
                            {(p.estimateMethod === "justin" || p.estimateMethod === "industry") ? (
                              <div className="text-xs border border-input rounded px-2 py-1 bg-muted/30 h-7 flex items-center text-muted-foreground">
                                Welds in fitting (locked for {p.estimateMethod === "industry" ? "Industry" : "Justin"})
                              </div>
                            ) : (
                              <select
                                className="text-xs border border-input rounded px-2 py-1 bg-background h-7"
                                value={estFittingWeldMode}
                                onChange={e => setEstFittingWeldMode(e.target.value as "bundled" | "separate" | "auto-welds")}
                                data-testid="select-fitting-weld-mode"
                              >
                                <option value="bundled">Bundled (legacy multipliers)</option>
                                <option value="separate">Separate weld rows</option>
                                <option value="auto-welds">Auto-welds (welds in fitting)</option>
                              </select>
                            )}
                          </div>

                          {/* Diagnose button */}
                          <Button
                            size="sm"
                            variant="outline"
                            className="h-7 text-xs"
                            onClick={() => { setSelectedId(p.id); setShowDiagnoseDialog(true); diagnoseMutation.mutate(p.id); }}
                            disabled={p.items.length === 0}
                            data-testid="btn-diagnose"
                            title="Show formula breakdown for every row + double-count warnings"
                          >
                            <Info size={12} className="mr-1.5" />
                            Diagnose
                          </Button>

                          {/* Calculate button */}
                          <Button
                            size="sm"
                            className="h-7 text-xs ml-auto"
                            onClick={() => autoCalculateMutation.mutate(p.id)}
                            disabled={autoCalculateMutation.isPending || p.items.length === 0}
                            data-testid="btn-auto-calculate"
                          >
                            <Calculator size={12} className="mr-1.5" />
                            {autoCalculateMutation.isPending ? "Calculating..." : "Calculate Labor Hours"}
                          </Button>
                        </div>
                      </CardContent>
                    </Card>

                    {/* Quick entry bar */}
                    <div className="flex gap-2">
                      <Input
                        className="h-8 text-xs flex-1"
                        placeholder='Type: 3 4" butt welds, 100 LF 6" pipe, 50 CY concrete...'
                        value={quickEntry}
                        onChange={e => setQuickEntry(e.target.value)}
                        onKeyDown={e => e.key === "Enter" && handleQuickEntry()}
                        data-testid="input-quick-entry"
                      />
                      <Button
                        size="sm"
                        className="h-8 text-xs"
                        onClick={handleQuickEntry}
                        data-testid="btn-quick-add"
                      >
                        <Plus size={13} className="mr-1" />
                        Add
                      </Button>
                    </div>

                    {/* Action buttons */}
                    <div className="flex items-center gap-2 flex-wrap">
                      <Button
                        variant="outline"
                        size="sm"
                        onClick={() => applyDbMutation.mutate(p.id)}
                        disabled={applyDbMutation.isPending}
                        data-testid="btn-apply-db"
                      >
                        <Database size={13} className="mr-1.5" />
                        Apply Database Costs
                      </Button>
                      <Button
                        variant="outline"
                        size="sm"
                        onClick={() => exportEstimatePdf(p)}
                        data-testid="btn-export-estimate"
                      >
                        <Download size={13} className="mr-1.5" />
                        Export PDF
                      </Button>
                      <Button
                        variant="outline"
                        size="sm"
                        className="text-emerald-700 border-emerald-300 hover:bg-emerald-50 dark:text-emerald-400 dark:border-emerald-700 dark:hover:bg-emerald-900/30"
                        onClick={async () => {
                          try {
                            const res = await apiRequest("GET", `/api/estimates/${p.id}/export-bill`);
                            const blob = await res.blob();
                            const url = URL.createObjectURL(blob);
                            const a = document.createElement("a");
                            a.href = url;
                            a.download = `${p.name} - Bills Format.xlsx`;
                            a.click();
                            URL.revokeObjectURL(url);
                            toast({ title: "Downloaded Bill's format workbook" });
                          } catch { toast({ title: "Export failed", variant: "destructive" }); }
                        }}
                        data-testid="btn-export-bill"
                      >
                        <FileSpreadsheet size={13} className="mr-1.5" />
                        Export Bill's Format
                      </Button>
                      <Button
                        variant="outline"
                        size="sm"
                        className="text-emerald-700 border-emerald-300 hover:bg-emerald-50 dark:text-emerald-400 dark:border-emerald-700 dark:hover:bg-emerald-900/30"
                        onClick={async () => {
                          try {
                            const res = await apiRequest("GET", `/api/estimates/${p.id}/export-justin`);
                            const blob = await res.blob();
                            const url = URL.createObjectURL(blob);
                            const a = document.createElement("a");
                            a.href = url;
                            a.download = `${p.name} - Justins Format.xlsx`;
                            a.click();
                            URL.revokeObjectURL(url);
                            toast({ title: "Downloaded Justin's format workbook" });
                          } catch { toast({ title: "Export failed", variant: "destructive" }); }
                        }}
                        data-testid="btn-export-justin"
                      >
                        <FileSpreadsheet size={13} className="mr-1.5" />
                        Export Justin's Format
                      </Button>
                      <Button
                        variant="outline"
                        size="sm"
                        className="text-emerald-700 border-emerald-300 hover:bg-emerald-50 dark:text-emerald-400 dark:border-emerald-700 dark:hover:bg-emerald-900/30"
                        onClick={async () => {
                          try {
                            const res = await apiRequest("GET", `/api/estimates/${p.id}/export-industry`);
                            const blob = await res.blob();
                            const url = URL.createObjectURL(blob);
                            const a = document.createElement("a");
                            a.href = url;
                            a.download = `${p.name} - Industry Standard.xlsx`;
                            a.click();
                            URL.revokeObjectURL(url);
                            toast({ title: "Downloaded Industry Standard workbook" });
                          } catch { toast({ title: "Export failed", variant: "destructive" }); }
                        }}
                        data-testid="btn-export-industry"
                      >
                        <FileSpreadsheet size={13} className="mr-1.5" />
                        Export Industry (Page)
                      </Button>
                      {/* Infer / Strip weld buttons are only relevant for Bill's method.
                          Justin and Industry handle weld math inline on the fitting row
                          (welds_per_fitting × weld_factor), so there's nothing for the
                          user to manually infer or strip. Pipe-field welds are auto-added
                          by ensurePipeFieldWeldRows on import / recalculate. */}
                      {p.estimateMethod !== "justin" && p.estimateMethod !== "industry" && (
                        <>
                          <Button
                            variant="outline"
                            size="sm"
                            onClick={() => inferWeldsMutation.mutate(p.id)}
                            disabled={inferWeldsMutation.isPending || p.items.length === 0}
                            data-testid="btn-infer-welds"
                          >
                            <Wand2 size={13} className="mr-1.5" />
                            {inferWeldsMutation.isPending ? "Inferring..." : "Infer Welds from Fittings"}
                          </Button>
                          <Button
                            variant="outline"
                            size="sm"
                            onClick={() => stripInferredWeldsMutation.mutate(p.id)}
                            disabled={stripInferredWeldsMutation.isPending || p.items.length === 0}
                            data-testid="btn-strip-inferred-welds"
                            title="Remove every row tagged 'auto-inferred' \u2014 useful if you're in Bundled mode and the BOM has been over-counted"
                          >
                            <Trash2 size={13} className="mr-1.5" />
                            {stripInferredWeldsMutation.isPending ? "Stripping..." : "Strip Auto-Inferred Welds"}
                          </Button>
                        </>
                      )}
                      <Button
                        variant="outline"
                        size="sm"
                        onClick={() => { setShowVersions(!showVersions); if (!showVersions) queryClient.invalidateQueries({ queryKey: ["/api/estimates", selectedId, "versions"] }); }}
                        data-testid="btn-version-history"
                      >
                        <History size={13} className="mr-1.5" />
                        Version History
                      </Button>
                      <Button
                        variant="outline"
                        size="sm"
                        onClick={() => setShowRfq(true)}
                        disabled={p.items.length === 0}
                        data-testid="btn-rfq"
                      >
                        <ShoppingCart size={13} className="mr-1.5" />
                        Request Material Quotes
                      </Button>
                      <Button
                        variant="outline"
                        size="sm"
                        className="text-blue-700 border-blue-300 hover:bg-blue-50 dark:text-blue-400 dark:border-blue-700 dark:hover:bg-blue-900/30"
                        onClick={() => generateBidReport(p)}
                        disabled={p.items.length === 0}
                        data-testid="btn-bid-report"
                      >
                        <FileSpreadsheet size={13} className="mr-1.5" />
                        Bid Report PDF
                      </Button>
                    </div>

                    {/* Feature 3: Version History Panel */}
                    {showVersions && (
                      <Card className="border-card-border">
                        <CardHeader className="p-3 pb-1">
                          <CardTitle className="text-xs flex items-center gap-1.5">
                            <History size={12} /> Version History
                          </CardTitle>
                        </CardHeader>
                        <CardContent className="p-3 pt-1">
                          {versions.length === 0 ? (
                            <p className="text-xs text-muted-foreground">No versions saved yet. Versions are auto-saved before each calculation.</p>
                          ) : (
                            <div className="space-y-1.5 max-h-48 overflow-y-auto">
                              {versions.map((v: any) => (
                                <div key={v.id} className="flex items-center justify-between gap-2 p-2 rounded bg-muted/30 border border-border">
                                  <div className="min-w-0">
                                    <p className="text-[10px] font-medium truncate">{v.notes || "Snapshot"}</p>
                                    <p className="text-[10px] text-muted-foreground">
                                      {new Date(v.createdAt).toLocaleString()} · {v.itemCount} items · {v.estimateMethod}
                                    </p>
                                  </div>
                                  <Button
                                    variant="outline"
                                    size="sm"
                                    className="h-6 text-[10px] px-2 shrink-0"
                                    onClick={() => {
                                      if (window.confirm("Restore this version? Current state will be overwritten.")) {
                                        restoreVersionMutation.mutate({ estId: p.id, versionId: v.id });
                                      }
                                    }}
                                    data-testid={`btn-restore-${v.id}`}
                                  >
                                    Restore
                                  </Button>
                                </div>
                              ))}
                            </div>
                          )}
                        </CardContent>
                      </Card>
                    )}

                    {/* Filter input for items table */}
                    {p.items.length > 0 && (
                      <div className="relative max-w-xs">
                        <Search size={14} className="absolute left-2.5 top-1/2 -translate-y-1/2 text-muted-foreground" />
                        <Input
                          className="pl-8 h-7 text-xs"
                          placeholder="Filter items by description, category, size..."
                          value={itemFilter}
                          onChange={e => setItemFilter(e.target.value)}
                          data-testid="input-item-filter"
                        />
                      </div>
                    )}

                    {/* Mode-mismatch banner — catches double-counting (bundled + welds) and missing-welds (separate + no welds) */}
                    {(() => {
                      const mm = detectModeMismatch(p.items, estFittingWeldMode);
                      if (!mm) return null;
                      if (mm.kind === "double-count") {
                        return (
                          <div className="border border-amber-500/40 bg-amber-500/10 rounded p-3 text-xs flex items-start gap-3">
                            <AlertCircle size={16} className="text-amber-600 mt-0.5 shrink-0" />
                            <div className="flex-1">
                              <div className="font-semibold">Likely double-counting at size{mm.sizes && mm.sizes.length > 1 ? "s" : ""} {mm.sizes?.join(", ")}</div>
                              <div className="text-[11px] mt-0.5 leading-relaxed">
                                This estimate is in <strong>Bundled</strong> mode (fittings carry weld labor) but the BOM also has explicit weld rows at the same size. Switch to Separate mode so each weld counts exactly once, or click "Strip Auto-Inferred Welds" to remove the inferred rows.
                              </div>
                            </div>
                            <Button size="sm" variant="outline" className="h-7 text-xs shrink-0" onClick={() => setEstFittingWeldMode("separate")}>
                              Switch to Separate
                            </Button>
                          </div>
                        );
                      }
                      return (
                        <div className="border border-blue-500/40 bg-blue-500/10 rounded p-3 text-xs flex items-start gap-3">
                          <AlertCircle size={16} className="text-blue-600 mt-0.5 shrink-0" />
                          <div className="flex-1">
                            <div className="font-semibold">Mode is "Separate" but no weld rows exist</div>
                            <div className="text-[11px] mt-0.5 leading-relaxed">
                              In Separate mode fittings only contribute handling labor (×0.15) and weld rows carry the welding labor. Your BOM has fittings but no welds, so weld labor isn't being counted at all. Either switch back to Bundled, or click "Infer Welds from Fittings" to add weld rows.
                            </div>
                          </div>
                          <Button size="sm" variant="outline" className="h-7 text-xs shrink-0" onClick={() => setEstFittingWeldMode("bundled")}>
                            Switch to Bundled
                          </Button>
                        </div>
                      );
                    })()}

                    {/* Estimate items table */}
                    <div className="overflow-auto rounded-md border border-border">
                      <table className="w-full text-xs">
                        <thead>
                          <tr className="bg-muted/50 border-b border-border">
                            {([
                              { field: "lineNumber" as SortField, label: "#", cls: "w-8 text-left" },
                              { field: null as any, label: "\u2713", cls: "w-6 text-center" },
                              { field: "category" as SortField, label: "Category", cls: "text-left" },
                              { field: "size" as SortField, label: "Size", cls: "text-left" },
                              { field: "description" as SortField, label: "Description", cls: "text-left min-w-[160px]" },
                              { field: "quantity" as SortField, label: "Qty", cls: "text-right w-16" },
                              { field: null as any, label: "Unit", cls: "text-left w-12" },
                              { field: "materialUnitCost" as SortField, label: "Mat $/Unit", cls: "text-right w-20" },
                              { field: "laborUnitCost" as SortField, label: "Labor $/Unit", cls: "text-right w-20" },
                              { field: "laborHoursPerUnit" as SortField, label: "Hrs/Unit", cls: "text-right w-16" },
                              // Connection count column — shows the qty × per-unit math:
                              // pipe rows show 'N LF', fitting rows '2 welds', flange '1 bolt-up'.
                              { field: null as any, label: "Count", cls: "text-left w-24" },
                              // Factor column — raw, unchanged number from the method's table.
                              // For a 3" SS elbow this is 4.68. Never multiplied by anything.
                              { field: null as any, label: "Factor", cls: "text-left w-28" },
                              { field: "materialExtension" as SortField, label: "Mat Total", cls: "text-right w-24" },
                              { field: "laborExtension" as SortField, label: "Labor Total", cls: "text-right w-24" },
                              { field: "totalCost" as SortField, label: "Total", cls: "text-right w-24" },
                              { field: null as any, label: "", cls: "w-16" },
                            ] as const).map(({ field, label, cls }, idx) => (
                              <th
                                key={idx}
                                className={`px-2 py-2 ${cls} ${field ? "cursor-pointer select-none hover:bg-muted/80" : ""}`}
                                onClick={field ? () => handleSort(field) : undefined}
                              >
                                <span className="inline-flex items-center gap-0.5">
                                  {label}
                                  {field && sortField === field && (
                                    sortDir === "asc" ? <ChevronUp size={10} /> : <ChevronDown size={10} />
                                  )}
                                </span>
                              </th>
                            ))}
                          </tr>
                        </thead>
                        <tbody>
                          {p.items.length === 0 ? (
                            <tr>
                              <td colSpan={16} className="text-center py-8 text-muted-foreground">
                                No items. Use quick entry above or import from takeoff.
                              </td>
                            </tr>
                          ) : displayItems.length === 0 ? (
                            <tr>
                              <td colSpan={16} className="text-center py-6 text-muted-foreground">
                                No items match filter.
                              </td>
                            </tr>
                          ) : (
                            <TooltipProvider delayDuration={300}>
                            {displayItems.map(item => {
                              const conf = getRowConfidence(item);
                              // A row is "excluded somewhere" if any of the three scope flags
                              // is false. We gray the whole row so the user can spot it at a
                              // glance, but still show it (and let them re-include via the
                              // settings popover).
                              const anyExcluded =
                                (item as any).includeInBom === false ||
                                (item as any).includeInTakeoff === false ||
                                (item as any).includeInEstimate === false;
                              const dotColor = conf.state === "green"
                                ? "bg-emerald-500"
                                : conf.state === "yellow"
                                  ? "bg-amber-500"
                                  : conf.state === "red"
                                    ? "bg-rose-500"
                                    : "bg-muted-foreground/40 ring-1 ring-muted-foreground/20";
                              return (
                              <tr key={item.id} className={`border-b border-border hover:bg-muted/20 ${anyExcluded ? "opacity-50 bg-muted/30" : ""}`} data-testid={`est-row-${item.id}`}>
                                <td className="px-2 py-1.5 text-muted-foreground">{item.lineNumber}</td>
                                <td className="px-2 py-1.5 text-center">
                                  <Tooltip>
                                    <TooltipTrigger asChild>
                                      <span className={`inline-block h-2.5 w-2.5 rounded-full ${dotColor} cursor-help`} data-testid={`conf-${conf.state}-${item.id}`} />
                                    </TooltipTrigger>
                                    <TooltipContent className="max-w-sm" side="right">
                                      <div className="text-xs font-semibold mb-1">{conf.label}</div>
                                      {conf.reasons.length > 0 ? (
                                        <ul className="text-[11px] space-y-0.5 list-disc list-inside">
                                          {conf.reasons.map((r, i) => <li key={i}>{r}</li>)}
                                        </ul>
                                      ) : (
                                        <div className="text-[11px] text-muted-foreground">Exact size match, calculator ran cleanly.</div>
                                      )}
                                      {(item as any).calculationBasis && (
                                        <div className="text-[10px] text-muted-foreground mt-2 border-t pt-1">{(item as any).calculationBasis}</div>
                                      )}
                                    </TooltipContent>
                                  </Tooltip>
                                </td>
                                <td className="px-2 py-1.5">
                                  <Badge variant="outline" className="text-[9px] px-1 py-0">{item.category}</Badge>
                                </td>
                                <td className="px-2 py-1.5 font-mono">{item.size}</td>
                                <td className="px-2 py-1.5 max-w-xs">
                                  <span className="line-clamp-1">{item.description}</span>
                                </td>
                                <td className="px-2 py-1.5 text-right font-mono">
                                  {item.unit === "LF" ? item.quantity.toFixed(2) : item.quantity.toLocaleString()}
                                </td>
                                <td className="px-2 py-1.5 text-muted-foreground">{item.unit}</td>
                                {(["materialUnitCost", "laborUnitCost", "laborHoursPerUnit"] as const).map(field => (
                                  <td key={field} className="px-1 py-1">
                                    {(field === "laborHoursPerUnit" || field === "materialUnitCost") && (item as any).calculationBasis ? (
                                      <Tooltip>
                                        <TooltipTrigger asChild>
                                          <div className="relative">
                                            <input
                                              type="number"
                                              step="0.01"
                                              className={`w-full text-right text-xs bg-transparent border-0 border-b border-transparent hover:border-border focus:border-primary outline-none px-1 font-mono ${field === "materialUnitCost" && (item as any).materialCostSource === "allowance" ? "text-amber-600 dark:text-amber-400" : field === "materialUnitCost" && (item as any).materialCostSource === "database" ? "text-green-600 dark:text-green-400" : field === "materialUnitCost" && (item as any).materialCostSource === "purchase_history" ? "text-blue-600 dark:text-blue-400" : ""}`}
                                              defaultValue={item[field]}
                                              onBlur={e => handleUpdateItemCost(item.id, field, parseFloat(e.target.value) || 0)}
                                              data-testid={`input-${field}-${item.id}`}
                                            />
                                            {field === "materialUnitCost" && (item as any).materialCostSource === "allowance" && (
                                              <span className="absolute -top-1 -right-1 text-[7px] bg-amber-100 dark:bg-amber-900/60 text-amber-700 dark:text-amber-300 rounded px-0.5 leading-tight">est</span>
                                            )}
                                            {field === "materialUnitCost" && (item as any).materialCostSource === "database" && (
                                              <span className="absolute -top-1 -right-1 text-[7px] bg-green-100 dark:bg-green-900/60 text-green-700 dark:text-green-300 rounded px-0.5 leading-tight">DB</span>
                                            )}
                                            {field === "materialUnitCost" && (item as any).materialCostSource === "purchase_history" && (
                                              <span className="absolute -top-1 -right-1 text-[7px] bg-blue-100 dark:bg-blue-900/60 text-blue-700 dark:text-blue-300 rounded px-0.5 leading-tight">Hist</span>
                                            )}
                                            {field === "laborHoursPerUnit" && (item as any).sizeMatchExact === false && (
                                              <span className="absolute -top-1 -right-1 text-[7px] bg-orange-100 dark:bg-orange-900/60 text-orange-700 dark:text-orange-300 rounded px-0.5 leading-tight">⚠</span>
                                            )}
                                          </div>
                                        </TooltipTrigger>
                                        <TooltipContent side="top" className="max-w-sm text-[10px] font-mono whitespace-pre-wrap">
                                          {(item as any).calculationBasis}
                                          {field === "materialUnitCost" && (item as any).materialCostSource === "allowance" && "\n⚠ EI-derived allowance, not real pricing"}
                                          {field === "materialUnitCost" && (item as any).materialCostSource === "database" && "\n✓ Priced from cost database"}
                                          {field === "materialUnitCost" && (item as any).materialCostSource === "purchase_history" && "\n✓ Priced from purchase history"}
                                        </TooltipContent>
                                      </Tooltip>
                                    ) : (
                                      <input
                                        type="number"
                                        step="0.01"
                                        className="w-full text-right text-xs bg-transparent border-0 border-b border-transparent hover:border-border focus:border-primary outline-none px-1 font-mono"
                                        defaultValue={item[field]}
                                        onBlur={e => handleUpdateItemCost(item.id, field, parseFloat(e.target.value) || 0)}
                                        data-testid={`input-${field}-${item.id}`}
                                      />
                                    )}
                                  </td>
                                ))}
                                {/* Count column — the connection count for this line.
                                    Pipe rows show 'N LF'. Fitting rows show 'qty × N welds = total'.
                                    Flanges show 'qty bolt-ups'. Threaded shows 'qty threads'.
                                    Hardware (stud bolts, gaskets, supports) shows '—'. */}
                                <td className="px-2 py-1.5 text-[11px] text-muted-foreground">
                                  {(() => {
                                    const cntPerUnit = (item as any).connectionCount;
                                    const typ = (item as any).connectionType as string | undefined;
                                    const calcBasis = (item as any).calculationBasis || "";
                                    const qty = item.quantity || 0;
                                    // Pipe → show LF
                                    if (typ === "pipe") {
                                      return <span className="font-mono" title={calcBasis}>{qty.toLocaleString()} LF</span>;
                                    }
                                    if (typ === undefined || cntPerUnit === undefined) {
                                      return <span className="text-muted-foreground/40" title={calcBasis}>—</span>;
                                    }
                                    if (typ === "none" || cntPerUnit === 0) {
                                      return <span className="text-muted-foreground/40" title={calcBasis}>—</span>;
                                    }
                                    // Explicit weld rows: qty IS the count. Everything else: qty × per-unit.
                                    const cat = (item.category || "").toLowerCase();
                                    const desc = (item.description || "").toLowerCase();
                                    const isExplicitWeldRow = cat === "weld" || desc.startsWith("bw ") || desc.includes(" bw ") || desc.startsWith("fw ") || desc.includes(" fw ") || desc.startsWith("sw ") || desc.includes(" sw ");
                                    const totalCount = isExplicitWeldRow ? qty : (qty * cntPerUnit);
                                    const plural = totalCount !== 1;
                                    const label = typ === "weld" ? (plural ? "welds" : "weld")
                                      : typ === "bolt-up" ? (plural ? "bolt-ups" : "bolt-up")
                                      : typ === "thread" ? (plural ? "threads" : "thread")
                                      : typ === "socket-weld" ? (plural ? "SWs" : "SW")
                                      : typ;
                                    const tooltipDetail = !isExplicitWeldRow && cntPerUnit > 0
                                      ? `${qty} × ${cntPerUnit} ${cntPerUnit === 1 ? label.replace(/s$/, "") : label}/unit = ${totalCount} ${label}\n\n${calcBasis}`
                                      : calcBasis;
                                    return (
                                      <span className="font-mono" title={tooltipDetail}>
                                        {Number.isFinite(totalCount) ? totalCount.toLocaleString() : totalCount} {label}
                                      </span>
                                    );
                                  })()}
                                </td>
                                {/* Factor column — the LITERAL number from the method's table.
                                    For Justin a 3" SS BW is 4.68 MH/weld. Never adjusted. */}
                                <td className="px-2 py-1.5 text-[11px] text-muted-foreground font-mono" title={(item as any).calculationBasis || ""}>
                                  {(() => {
                                    const lbl = (item as any).rawFactorLabel as string | undefined;
                                    const rf = (item as any).rawFactor as number | undefined;
                                    if (lbl) return <span>{lbl}</span>;
                                    if (rf != null) return <span>{rf.toFixed(3)}</span>;
                                    return <span className="text-muted-foreground/40">—</span>;
                                  })()}
                                </td>
                                <td className="px-2 py-1.5 text-right font-mono text-blue-600 dark:text-blue-400">{fmt$(item.materialExtension || 0)}</td>
                                <td className="px-2 py-1.5 text-right font-mono text-orange-600 dark:text-orange-400">{fmt$(item.laborExtension || 0)}</td>
                                <td className="px-2 py-1.5 text-right font-mono font-semibold">{fmt$(item.totalCost || 0)}</td>
                                <td className="px-2 py-1.5 flex items-center gap-0.5">
                                  {/* Per-line work type badge */}
                                  {estMethod === "justin" && (
                                    <button
                                      className={`text-[8px] font-bold w-4 h-4 rounded flex items-center justify-center shrink-0 border transition-colors ${
                                        (item as any).workType === "rack"
                                          ? "bg-orange-100 text-orange-700 border-orange-300 dark:bg-orange-900/40 dark:text-orange-300 dark:border-orange-700"
                                          : "bg-muted/30 text-muted-foreground border-transparent hover:border-border"
                                      }`}
                                      title={`${(item as any).workType === "rack" ? "Rack" : "Standard"} work — click to toggle`}
                                      onClick={() => {
                                        const newType = (item as any).workType === "rack" ? "standard" : "rack";
                                        handleUpdateItemOverride(item.id, "workType", newType);
                                      }}
                                      data-testid={`btn-worktype-${item.id}`}
                                    >
                                      {(item as any).workType === "rack" ? "R" : "S"}
                                    </button>
                                  )}
                                  {/* Per-line overrides popover */}
                                  <Popover>
                                    <PopoverTrigger asChild>
                                      <button
                                        className="text-muted-foreground hover:text-primary"
                                        title="Edit line assumptions"
                                        data-testid={`btn-overrides-${item.id}`}
                                      >
                                        <Settings2 size={11} />
                                      </button>
                                    </PopoverTrigger>
                                    <PopoverContent className="w-64 p-3 space-y-3" side="left">
                                      <div>
                                        <p className="text-[10px] font-semibold text-muted-foreground uppercase mb-1">Include in</p>
                                        <div className="space-y-1">
                                          {([
                                            { field: "includeInBom", label: "BOM / RFQ", hint: "Material orders and vendor RFQs" },
                                            { field: "includeInTakeoff", label: "Takeoff", hint: "Takeoff page tables and PDFs" },
                                            { field: "includeInEstimate", label: "Estimate", hint: "Labor & cost totals" },
                                          ] as const).map(({ field, label, hint }) => {
                                            const checked = (item as any)[field] !== false;
                                            return (
                                              <label key={field} className="flex items-center gap-2 text-[10px] cursor-pointer hover:bg-muted/40 rounded px-1 py-0.5" title={hint}>
                                                <input
                                                  type="checkbox"
                                                  className="h-3 w-3"
                                                  checked={checked}
                                                  onChange={e => handleUpdateItemOverride(item.id, field, e.target.checked)}
                                                  data-testid={`scope-${field}-${item.id}`}
                                                />
                                                <span className="flex-1">{label}</span>
                                                {!checked && <span className="text-[9px] text-amber-600 dark:text-amber-400">excluded</span>}
                                              </label>
                                            );
                                          })}
                                        </div>
                                      </div>
                                      <div className="border-t pt-2">
                                        <p className="text-[10px] font-semibold text-muted-foreground uppercase mb-1">Line Overrides</p>
                                        {[
                                          { field: "workType", label: "Work Type", options: ["standard", "rack"] },
                                          { field: "itemMaterial", label: "Material", options: ["CS", "SS"] },
                                          { field: "itemSchedule", label: "Schedule", options: ["STD", "XH", "10", "20", "40", "80", "160/XXH"] },
                                          { field: "itemElevation", label: "Elevation", options: ["0-20ft", "20-40ft", "40-80ft", "80ft+"] },
                                          { field: "itemPipeLocation", label: "Pipe Location", options: ["Sleeper Rack", "Underground", "Open Rack", "Elevated Rack"] },
                                          { field: "itemAlloyGroup", label: "Alloy Group", options: ["1", "2", "3", "4", "5", "6", "7", "8", "9"] },
                                        ].map(({ field, label, options }) => (
                                          <div key={field} className="flex items-center gap-2 mb-1">
                                            <label className="text-[10px] text-muted-foreground w-20 shrink-0">{label}</label>
                                            <select
                                              className="flex-1 text-[10px] bg-background border border-input rounded px-1.5 py-0.5"
                                              value={(item as any)[field] || ""}
                                              onChange={e => handleUpdateItemOverride(item.id, field, e.target.value || undefined)}
                                            >
                                              <option value="">(global)</option>
                                              {options.map(o => <option key={o} value={o}>{o}</option>)}
                                            </select>
                                          </div>
                                        ))}
                                      </div>
                                    </PopoverContent>
                                  </Popover>
                                  <button
                                    className="text-muted-foreground hover:text-destructive"
                                    onClick={() => handleDeleteItem(item.id)}
                                    data-testid={`btn-delete-item-${item.id}`}
                                  >
                                    <Trash2 size={12} />
                                  </button>
                                </td>
                              </tr>
                              );
                            })}
                            </TooltipProvider>
                          )}
                        </tbody>
                        {p.items.length > 0 && (
                          <tfoot>
                            <tr className="bg-muted/50 font-medium border-t-2 border-border">
                              <td colSpan={10} className="px-2 py-2 text-right text-xs">TOTALS</td>
                              <td className="px-2 py-2 text-right text-xs font-mono text-blue-600 dark:text-blue-400">{fmt$(totalMaterial)}</td>
                              <td className="px-2 py-2 text-right text-xs font-mono text-orange-600 dark:text-orange-400">{fmt$(totalLabor)}</td>
                              <td className="px-2 py-2 text-right text-xs font-mono font-semibold">{fmt$(subtotal)}</td>
                              <td></td>
                            </tr>
                          </tfoot>
                        )}
                      </table>
                    </div>

                    {/* Markup + Totals panel */}
                    <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                      {/* Markups */}
                      <Card className="border-card-border">
                        <CardHeader className="p-4 pb-2">
                          <CardTitle className="text-sm">Markups</CardTitle>
                        </CardHeader>
                        <CardContent className="p-4 pt-0 space-y-2">
                          {[
                            { field: "overhead", label: "Overhead" },
                            { field: "profit", label: "Profit" },
                            { field: "tax", label: "Tax (Material)" },
                            { field: "bond", label: "Bond" },
                          ].map(({ field, label }) => (
                            <div key={field} className="flex items-center gap-3">
                              <label className="text-xs text-muted-foreground w-28 shrink-0">{label}</label>
                              <div className="flex items-center gap-1">
                                <input
                                  type="number"
                                  step="0.1"
                                  className="w-20 text-right text-xs border border-input rounded px-2 py-1 bg-background font-mono"
                                  defaultValue={(p.markups as any)[field]}
                                  onBlur={e => handleUpdateMarkup(field, parseFloat(e.target.value) || 0)}
                                  data-testid={`input-markup-${field}`}
                                />
                                <span className="text-xs text-muted-foreground">%</span>
                              </div>
                            </div>
                          ))}
                          {/* Contingency override — multiplies per-unit MH at calculate time,
                              NOT a post-subtotal markup. Empty = use method default
                              (Industry 10%, Justin 15%, Bill n/a). Re-run Auto-Calculate
                              after changing this value for it to apply to the BOM. */}
                          <Separator className="my-2" />
                          <div className="flex items-center gap-3">
                            <label className="text-xs text-muted-foreground w-28 shrink-0" title="Multiplies per-unit man-hours at calculate time. Industry default 10%, Justin default 15%. Bill's method does not use contingency.">
                              Contingency
                            </label>
                            <div className="flex items-center gap-1">
                              <input
                                type="number"
                                step="0.5"
                                placeholder={p.estimateMethod === "justin" ? "15" : p.estimateMethod === "industry" ? "10" : (p.estimateMethod === "bill" ? "n/a" : "15")}
                                disabled={p.estimateMethod === "bill"}
                                className="w-20 text-right text-xs border border-input rounded px-2 py-1 bg-background font-mono"
                                defaultValue={(p as any).contingencyOverride ?? ""}
                                onBlur={e => {
                                  const raw = e.target.value.trim();
                                  const v = raw === "" ? null : (parseFloat(raw) || 0);
                                  patchMutation.mutate({ id: p.id, data: { contingencyOverride: v } as any });
                                }}
                                data-testid="input-markup-contingency"
                              />
                              <span className="text-xs text-muted-foreground">%</span>
                            </div>
                          </div>
                          <p className="text-[10px] text-muted-foreground/80 leading-tight">
                            {p.estimateMethod === "bill" 
                              ? "Bill's method doesn't apply contingency."
                              : (p as any).contingencyOverride === undefined || (p as any).contingencyOverride === null
                                ? `Empty = use ${p.estimateMethod === "justin" ? "Justin's 15%" : p.estimateMethod === "industry" ? "Industry 10%" : "method"} default. Re-run Auto-Calculate to apply changes.`
                                : `Overriding default (${p.estimateMethod === "justin" ? "15%" : p.estimateMethod === "industry" ? "10%" : "method default"}). Re-run Auto-Calculate to apply.`}
                          </p>
                        </CardContent>
                      </Card>

                      {/* Totals */}
                      <Card className="border-card-border">
                        <CardHeader className="p-4 pb-2">
                          <CardTitle className="text-sm">Summary</CardTitle>
                        </CardHeader>
                        <CardContent className="p-4 pt-0 space-y-1">
                          {([
                            { label: "Material", val: totalMaterial },
                            { label: "Labor (BOM)", val: bomLabor, faint: true },
                            ...(adderLabor > 0 ? [{ label: "Labor (scope adders)", val: adderLabor, faint: true }] : []),
                            { label: "Labor (total)", val: totalLabor },
                            { label: `Labor Hours`, val: null, text: `${totalHours.toFixed(1)} hrs${adderHours > 0 ? ` (BOM ${bomHours.toFixed(1)} + adders ${adderHours.toFixed(1)})` : ""}` },
                            ...(adderFlatCost > 0 ? [{ label: "Other costs (scope adders)", val: adderFlatCost }] : []),
                            { label: "Subtotal", val: subtotal, bold: true },
                            { label: `Overhead (${p.markups.overhead}%)`, val: overheadAmt },
                            { label: `Profit (${p.markups.profit}%)`, val: profitAmt },
                            { label: `Tax (${p.markups.tax}%)`, val: taxAmt },
                            { label: `Bond (${p.markups.bond}%)`, val: bondAmt },
                          ] as Array<{ label: string; val: number | null; text?: string; bold?: boolean; faint?: boolean }>).map(({ label, val, text, bold, faint }) => (
                            <div key={label} className="flex justify-between text-xs">
                              <span className={bold ? "font-semibold" : faint ? "text-muted-foreground/70 pl-2" : "text-muted-foreground"}>{label}</span>
                              <span className={`font-mono ${bold ? "font-semibold" : faint ? "text-muted-foreground/70" : ""}`}>
                                {text !== undefined ? text : fmt$(val || 0)}
                              </span>
                            </div>
                          ))}
                          <Separator className="my-2" />
                          <div className="flex justify-between text-sm font-bold">
                            <span>GRAND TOTAL</span>
                            <span className="font-mono text-primary">{fmt$(grandTotal)}</span>
                          </div>
                        </CardContent>
                      </Card>
                    </div>

                    {/* Reconciliation — estimate qty vs source BOM qty vs connections.
                        Loads on demand. Surfaces any mismatch (e.g. a takeoff line that
                        didn't get into the estimate, or quantity drift). */}
                    <ReconciliationPanel estimateId={p.id} />

                    {/* Scope adders — hand-entered labor for hydro, demo, supports,
                        supervision, etc. Anything that isn't on the BOM but needs to
                        be priced. Editing auto-saves via PATCH /api/estimates/:id. */}
                    <ScopeAddersPanel
                      adders={scopeAdders}
                      effectiveRate={effectiveRate}
                      onChange={next => patchMutation.mutate({ id: p.id, data: { scopeAdders: next } })}
                    />

                    {/* Crew Planner */}
                    <Separator />
                    <div className="flex items-center gap-2 pt-2">
                      <div className="w-8 h-8 rounded-lg bg-indigo-100 dark:bg-indigo-900/40 flex items-center justify-center">
                        <Calculator size={14} className="text-indigo-600 dark:text-indigo-400" />
                      </div>
                      <div>
                        <h3 className="text-sm font-semibold">Crew Planner</h3>
                        <p className="text-[10px] text-muted-foreground">Schedule labor and crew assignments</p>
                      </div>
                    </div>
                    <CrewPlanner
                      totalLaborHours={totalHours}
                      laborRate={p.laborRate}
                      perDiemRate={p.perDiem}
                    />
                    <ProjectPlanner
                      totalManhours={totalHours}
                      projectName={p.name}
                    />
                  </>
                )}
              </div>
            </TabsContent>

            {/* Cost Database Tab */}
            <TabsContent value="database" className="flex-1 m-0">
              <div className="p-5 space-y-4">
                <div className="flex items-center gap-3">
                  <div className="relative flex-1 max-w-sm">
                    <Search size={14} className="absolute left-2.5 top-1/2 -translate-y-1/2 text-muted-foreground" />
                    <Input
                      className="pl-8 h-8 text-xs"
                      placeholder="Search cost database..."
                      value={dbSearch}
                      onChange={e => setDbSearch(e.target.value)}
                      data-testid="input-db-search"
                    />
                  </div>
                  <Button
                    variant="outline"
                    size="sm"
                    onClick={() => {
                      const entry = {
                        description: "New Item",
                        size: "",
                        category: "other",
                        unit: "EA",
                        materialUnitCost: 0,
                        laborUnitCost: 0,
                        laborHoursPerUnit: 0,
                      };
                      addDbEntryMutation.mutate(entry);
                    }}
                    data-testid="btn-add-db-entry"
                  >
                    <Plus size={13} className="mr-1.5" />
                    Add Entry
                  </Button>
                  <Button
                    variant="outline"
                    size="sm"
                    onClick={() => {
                      const input = document.createElement("input");
                      input.type = "file";
                      input.accept = ".csv,.xlsx,.xls";
                      input.onchange = (e: any) => {
                        const file = e.target?.files?.[0];
                        if (file) csvImportMutation.mutate(file);
                      };
                      input.click();
                    }}
                    disabled={csvImportMutation.isPending}
                    data-testid="btn-import-csv"
                  >
                    <Upload size={13} className="mr-1.5" />
                    {csvImportMutation.isPending ? "Importing..." : "Import CSV/XLSX"}
                  </Button>
                </div>

                <div className="rounded-md border border-border overflow-auto">
                  <table className="w-full text-xs">
                    <thead>
                      <tr className="bg-muted/50 border-b border-border">
                        <th className="px-3 py-2 text-left">Description</th>
                        <th className="px-3 py-2 text-left">Size</th>
                        <th className="px-3 py-2 text-left">Category</th>
                        <th className="px-3 py-2 text-left">Unit</th>
                        <th className="px-3 py-2 text-right">Mat $/Unit</th>
                        <th className="px-3 py-2 text-right">Labor $/Unit</th>
                        <th className="px-3 py-2 text-right">Hrs/Unit</th>
                        <th className="px-3 py-2 text-left w-24">Last Updated</th>
                        <th className="px-3 py-2 w-8"></th>
                      </tr>
                    </thead>
                    <tbody>
                      {filteredDb.length === 0 ? (
                        <tr>
                          <td colSpan={9} className="text-center py-8 text-muted-foreground">
                            {costDb.length === 0 ? "No entries in cost database yet." : "No matching entries."}
                          </td>
                        </tr>
                      ) : (
                        filteredDb.map(entry => (
                          <DbRow key={entry.id} entry={entry} onDelete={() => deleteDbEntryMutation.mutate(entry.id)} />
                        ))
                      )}
                    </tbody>
                  </table>
                </div>
              </div>
            </TabsContent>
          </Tabs>
        </div>
      </div>

      {/* RFQ Modal */}
      {p && showRfq && (
        <RfqModal open={showRfq} onOpenChange={setShowRfq} project={p} />
      )}

      {/* New Project Dialog */}
      <Dialog open={showNewProject} onOpenChange={setShowNewProject}>
        <DialogContent className="max-w-sm">
          <DialogHeader>
            <DialogTitle className="text-sm">New Estimate</DialogTitle>
          </DialogHeader>
          <div className="space-y-3 py-2">
            <div>
              <Label className="text-xs">Project Name</Label>
              <Input
                className="mt-1 h-8 text-xs"
                value={newProjectName}
                onChange={e => setNewProjectName(e.target.value)}
                onKeyDown={e => e.key === "Enter" && newProjectName.trim() && createMutation.mutate(newProjectName)}
                placeholder="e.g. Plant 3 - Building A"
                data-testid="input-new-project-name"
                autoFocus
              />
            </div>
          </div>
          <DialogFooter>
            <Button variant="ghost" size="sm" onClick={() => setShowNewProject(false)}>Cancel</Button>
            <Button
              size="sm"
              disabled={!newProjectName.trim() || createMutation.isPending}
              onClick={() => createMutation.mutate(newProjectName)}
              data-testid="btn-create-estimate"
            >
              Create
            </Button>
          </DialogFooter>
        </DialogContent>
      </Dialog>

      {/* Save-as-Custom dialog */}
      <Dialog open={showSaveCustomDialog} onOpenChange={setShowSaveCustomDialog}>
        <DialogContent>
          <DialogHeader>
            <DialogTitle>Save as Custom Method</DialogTitle>
          </DialogHeader>
          <div className="space-y-3 py-2">
            <div>
              <label className="text-xs text-muted-foreground block mb-1">Profile name</label>
              <Input
                value={newCustomMethodName}
                onChange={e => setNewCustomMethodName(e.target.value)}
                placeholder={`My ${estMethod} profile`}
                data-testid="input-custom-method-name"
              />
            </div>
            <div>
              <label className="text-xs text-muted-foreground block mb-1">Description (optional)</label>
              <Input
                value={newCustomMethodDescription}
                onChange={e => setNewCustomMethodDescription(e.target.value)}
                placeholder="e.g. Calibrated to our 2026 industrial crew productivity"
                data-testid="input-custom-method-description"
              />
            </div>
            <div className="text-xs text-muted-foreground bg-muted/40 rounded p-2">
              <div>Base method: <span className="font-semibold">{estMethod === "bill" ? "Bill's EI" : estMethod === "industry" ? "Industry Standard (Page)" : "Justin's Factor"}</span></div>
              <div className="mt-1">The new profile starts as an exact clone of the base method. You can edit any factor later — future overrides will only differ from the base where you explicitly change them.</div>
            </div>
          </div>
          <DialogFooter>
            <Button variant="outline" onClick={() => setShowSaveCustomDialog(false)} data-testid="btn-cancel-custom-save">Cancel</Button>
            <Button
              disabled={!newCustomMethodName.trim() || saveCustomMutation.isPending}
              onClick={() => saveCustomMutation.mutate({
                name: newCustomMethodName.trim(),
                baseMethod: estMethod,
                description: newCustomMethodDescription.trim(),
              })}
              data-testid="btn-confirm-custom-save"
            >
              Save Profile
            </Button>
          </DialogFooter>
        </DialogContent>
      </Dialog>

      {/* Compare-Methods dialog */}
      <Dialog open={showCompareDialog} onOpenChange={setShowCompareDialog}>
        <DialogContent className="max-w-6xl max-h-[90vh] overflow-y-auto">
          <DialogHeader>
            <div className="flex items-center justify-between gap-3 pr-8">
              <DialogTitle>Compare All Estimating Methods</DialogTitle>
              {compareMutation.data && compareMutation.variables && (
                <Button
                  variant="outline"
                  size="sm"
                  className="text-emerald-700 border-emerald-300 hover:bg-emerald-50 dark:text-emerald-400 dark:border-emerald-700 dark:hover:bg-emerald-900/30"
                  onClick={async () => {
                    const id = compareMutation.variables as string;
                    try {
                      const res = await apiRequest("POST", `/api/estimates/${id}/compare-methods/export`, {
                        customMethodIds: customMethods.map(m => m.id),
                        laborRate: estLaborRate,
                        overtimeRate: estOvertimeRate,
                        doubleTimeRate: estDoubleTimeRate,
                        perDiem: estPerDiem,
                        overtimePercent: estOvertimePercent,
                        doubleTimePercent: estDoubleTimePercent,
                        material: estMaterial,
                        schedule: estSchedule,
                        installType: estInstallType,
                        pipeLocation: estPipeLocation,
                        elevation: estElevation,
                        alloyGroup: estAlloyGroup,
                        rackFactor: estRackFactor,
                        fittingWeldMode: estFittingWeldMode,
                      });
                      const blob = await res.blob();
                      const url = URL.createObjectURL(blob);
                      const a = document.createElement("a");
                      a.href = url;
                      const proj = estimates.find(p => p.id === id);
                      a.download = `${proj?.name || "Estimate"} - Compare All Methods.xlsx`;
                      a.click();
                      URL.revokeObjectURL(url);
                      toast({ title: "Downloaded compare workbook" });
                    } catch (err: any) {
                      toast({ title: "Export failed", description: err?.message || String(err), variant: "destructive" });
                    }
                  }}
                  data-testid="btn-export-compare"
                >
                  <FileSpreadsheet size={13} className="mr-1.5" /> Export to Excel
                </Button>
              )}
            </div>
          </DialogHeader>
          {compareMutation.isPending && (
            <div className="text-sm text-muted-foreground py-8 text-center">Running estimate through all methods…</div>
          )}
          {compareMutation.data && (
            <div className="space-y-4">
              {compareMutation.data.skippedMethods && compareMutation.data.skippedMethods.length > 0 && (
                <div className="text-xs bg-amber-500/10 border border-amber-500/30 rounded p-2">
                  <span className="font-semibold">Some methods skipped:</span> {compareMutation.data.skippedMethods.join(", ")}. The data block for these methods is missing from this deploy. Redeploy the latest main to enable them.
                </div>
              )}
              {/* Summary cards */}
              <div className="grid grid-cols-2 md:grid-cols-3 lg:grid-cols-4 gap-3">
                {compareMutation.data.summary.map((s: any) => (
                  <div key={s.key} className="border rounded p-3 bg-card">
                    <div className="text-xs font-semibold text-muted-foreground">{s.label}</div>
                    <div className="text-2xl font-bold tabular-nums mt-1">${Math.round(s.totalCost).toLocaleString()}</div>
                    <div className="text-[10px] text-muted-foreground mt-1 space-y-0.5">
                      <div>{Math.round(s.totalMH).toLocaleString()} MH</div>
                      <div>${Math.round(s.totalLaborCost).toLocaleString()} labor</div>
                      {s.totalMaterialCost > 0 && <div>${Math.round(s.totalMaterialCost).toLocaleString()} material</div>}
                    </div>
                  </div>
                ))}
              </div>

              {/* Per-line drill-down table */}
              <div className="text-xs font-semibold pt-2">Per-line breakdown ({compareMutation.data.itemCount} items)</div>
              <div className="border rounded overflow-x-auto">
                <table className="text-xs w-full">
                  <thead className="bg-muted/40">
                    <tr>
                      <th className="text-left px-2 py-1 sticky left-0 bg-muted/40 z-10">#</th>
                      <th className="text-left px-2 py-1 min-w-[200px]">Description</th>
                      <th className="text-left px-2 py-1">Size</th>
                      <th className="text-right px-2 py-1">Qty</th>
                      {compareMutation.data.summary.map((s: any) => (
                        <th key={s.key} className="text-right px-2 py-1" colSpan={2}>{s.label}</th>
                      ))}
                    </tr>
                    <tr className="bg-muted/20 text-[10px] text-muted-foreground">
                      <th></th><th></th><th></th><th></th>
                      {compareMutation.data.summary.map((s: any) => (
                        <React.Fragment key={s.key}>
                          <th className="text-right px-2 py-1">MH</th>
                          <th className="text-right px-2 py-1">$ Cost</th>
                        </React.Fragment>
                      ))}
                    </tr>
                  </thead>
                  <tbody>
                    {compareMutation.data.lineItems.map((li: any) => (
                      <tr key={li.itemId} className="border-t hover:bg-muted/20">
                        <td className="px-2 py-1 sticky left-0 bg-card">{li.lineNumber}</td>
                        <td className="px-2 py-1">{li.description}</td>
                        <td className="px-2 py-1">{li.size}</td>
                        <td className="px-2 py-1 text-right tabular-nums">{li.quantity}</td>
                        {compareMutation.data.summary.map((s: any) => {
                          const bm = li.byMethod[s.key];
                          return (
                            <React.Fragment key={s.key}>
                              <td className="px-2 py-1 text-right tabular-nums" title={bm?.calcBasis || ""}>
                                {bm ? bm.totalMH.toFixed(2) : "—"}
                              </td>
                              <td className="px-2 py-1 text-right tabular-nums">
                                {bm ? `$${Math.round(bm.totalCost).toLocaleString()}` : "—"}
                              </td>
                            </React.Fragment>
                          );
                        })}
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
              <div className="text-[10px] text-muted-foreground">
                Effective labor rate used: ${compareMutation.data.effectiveLaborRate?.toFixed(2)}/hr (ST/OT/DT blended + per-diem). This view is read-only — changing the selected method here doesn't modify your estimate. Use Auto-Calculate to commit.
              </div>
            </div>
          )}
          <DialogFooter>
            <Button variant="outline" onClick={() => setShowCompareDialog(false)} data-testid="btn-close-compare">Close</Button>
          </DialogFooter>
        </DialogContent>
      </Dialog>

      {/* Diagnose dialog — per-row formula breakdown + project warnings */}
      <Dialog open={showDiagnoseDialog} onOpenChange={setShowDiagnoseDialog}>
        <DialogContent className="max-w-6xl max-h-[90vh] overflow-y-auto">
          <DialogHeader>
            <DialogTitle>Diagnose Estimate</DialogTitle>
          </DialogHeader>
          {diagnoseMutation.isPending && (
            <div className="text-sm text-muted-foreground py-8 text-center">Running diagnostic…</div>
          )}
          {diagnoseMutation.data && (
            <div className="space-y-4">
              {/* Summary header */}
              <div className="flex items-center gap-4 flex-wrap text-xs">
                <span><span className="text-muted-foreground">Method:</span> <strong>{diagnoseMutation.data.method}{diagnoseMutation.data.customMethodName ? ` (${diagnoseMutation.data.customMethodName})` : ""}</strong></span>
                <span><span className="text-muted-foreground">Fitting welds:</span> <strong>{diagnoseMutation.data.fittingWeldMode === "separate" ? "Separate weld rows" : "Bundled in fitting"}</strong></span>
                <span><span className="text-muted-foreground">Total MH:</span> <strong className="tabular-nums">{Math.round(diagnoseMutation.data.totalMH).toLocaleString()}</strong></span>
                <span><span className="text-muted-foreground">Items:</span> <strong>{diagnoseMutation.data.itemsWithLabor} of {diagnoseMutation.data.itemCount} have labor</strong></span>
              </div>

              {/* Project-level warnings */}
              {diagnoseMutation.data.warnings && diagnoseMutation.data.warnings.length > 0 && (
                <div className="space-y-2">
                  {diagnoseMutation.data.warnings.map((w: any) => {
                    const tone = w.severity === "error"
                      ? "bg-destructive/10 border-destructive/40 text-destructive"
                      : w.severity === "warn"
                        ? "bg-amber-500/10 border-amber-500/40"
                        : "bg-blue-500/10 border-blue-500/30";
                    return (
                      <div key={w.code} className={`border rounded p-3 ${tone}`}>
                        <div className="text-xs font-semibold flex items-center gap-2">
                          <AlertCircle size={13} /> {w.title}
                        </div>
                        <div className="text-[11px] mt-1 leading-relaxed">{w.detail}</div>
                        {w.affectedItemIds && w.affectedItemIds.length > 0 && (
                          <div className="text-[10px] text-muted-foreground mt-1">Affects {w.affectedItemIds.length} item{w.affectedItemIds.length === 1 ? "" : "s"}.</div>
                        )}
                      </div>
                    );
                  })}
                </div>
              )}
              {diagnoseMutation.data.warnings && diagnoseMutation.data.warnings.length === 0 && (
                <div className="text-xs bg-emerald-500/10 border border-emerald-500/30 rounded p-2 flex items-center gap-2">
                  <CheckCircle2 size={13} className="text-emerald-600" /> No issues detected. Every row mapped to a labor factor and no double-counting was found.
                </div>
              )}

              {/* Per-row breakdown */}
              <div className="text-xs font-semibold pt-2">Per-row formula breakdown</div>
              <div className="border rounded overflow-x-auto">
                <table className="text-xs w-full">
                  <thead className="bg-muted/40 sticky top-0">
                    <tr>
                      <th className="text-left px-2 py-1">#</th>
                      <th className="text-left px-2 py-1 min-w-[180px]">Description</th>
                      <th className="text-left px-2 py-1">Size</th>
                      <th className="text-right px-2 py-1">Qty</th>
                      <th className="text-right px-2 py-1">MH/unit</th>
                      <th className="text-right px-2 py-1">Total MH</th>
                      <th className="text-left px-2 py-1 min-w-[300px]">Formula</th>
                    </tr>
                  </thead>
                  <tbody>
                    {diagnoseMutation.data.rows.map((r: any) => {
                      const flagged = r.warnings && r.warnings.length > 0;
                      return (
                        <tr key={r.itemId} className={`border-t hover:bg-muted/20 ${flagged ? "bg-amber-500/5" : ""}`}>
                          <td className="px-2 py-1 tabular-nums">{r.lineNumber}</td>
                          <td className="px-2 py-1 truncate max-w-[260px]" title={r.description}>{r.description}</td>
                          <td className="px-2 py-1">{r.size}</td>
                          <td className="px-2 py-1 text-right tabular-nums">{r.quantity} {r.unit}</td>
                          <td className="px-2 py-1 text-right tabular-nums">{r.mhPerUnit.toFixed(3)}</td>
                          <td className="px-2 py-1 text-right tabular-nums font-semibold">{r.totalMH.toFixed(2)}</td>
                          <td className="px-2 py-1 text-[10px] text-muted-foreground" title={r.calcBasis}>{r.calcBasis}</td>
                        </tr>
                      );
                    })}
                  </tbody>
                </table>
              </div>
            </div>
          )}
          <DialogFooter>
            <Button variant="outline" onClick={() => setShowDiagnoseDialog(false)}>Close</Button>
          </DialogFooter>
        </DialogContent>
      </Dialog>
    </AppLayout>
  );
}

// Inline editable DB row
function DbRow({ entry, onDelete }: { entry: CostDatabaseEntry; onDelete: () => void }) {
  const { toast } = useToast();
  const updateMutation = useMutation({
    mutationFn: (data: any) => apiRequest("PATCH", `/api/cost-database/${entry.id}`, data).then(r => r.json()),
    onSuccess: () => { queryClient.invalidateQueries({ queryKey: ["/api/cost-database"] }); },
  });

  return (
    <tr className="border-b border-border hover:bg-muted/20" data-testid={`db-row-${entry.id}`}>
      <td className="px-2 py-1.5">
        <input
          type="text"
          defaultValue={entry.description}
          className="w-full bg-transparent text-xs border-0 border-b border-transparent hover:border-border focus:border-primary outline-none px-0"
          onBlur={e => updateMutation.mutate({ description: e.target.value })}
        />
      </td>
      <td className="px-2 py-1.5">
        <input
          type="text"
          defaultValue={entry.size}
          className="w-full bg-transparent text-xs font-mono border-0 border-b border-transparent hover:border-border focus:border-primary outline-none px-0"
          onBlur={e => updateMutation.mutate({ size: e.target.value })}
        />
      </td>
      <td className="px-2 py-1.5">
        <Badge variant="outline" className="text-[9px] px-1 py-0">{entry.category}</Badge>
      </td>
      <td className="px-2 py-1.5 text-muted-foreground font-mono">{entry.unit}</td>
      {(["materialUnitCost", "laborUnitCost", "laborHoursPerUnit"] as const).map(field => (
        <td key={field} className="px-2 py-1.5 text-right">
          <input
            type="number"
            step="0.01"
            defaultValue={entry[field]}
            className="w-full text-right bg-transparent text-xs font-mono border-0 border-b border-transparent hover:border-border focus:border-primary outline-none"
            onBlur={e => updateMutation.mutate({ [field]: parseFloat(e.target.value) || 0 })}
          />
        </td>
      ))}
      <td className="px-2 py-1.5 text-muted-foreground whitespace-nowrap">
        {new Date(entry.lastUpdated).toLocaleDateString()}
      </td>
      <td className="px-2 py-1.5">
        <button className="text-muted-foreground hover:text-destructive" onClick={() => { if (window.confirm("Delete this cost database entry?")) onDelete(); }} data-testid={`btn-delete-db-${entry.id}`}>
          <Trash2 size={12} />
        </button>
      </td>
    </tr>
  );
}

// ReconciliationPanel — two tables side by side:
//   1) Line-by-line: estimate qty vs BOM qty per (category, size, material).
//   2) Connections: per-size connection counts (shop/field BW/SW, bolt-ups,
//      threaded) summed across the estimate AND the BOM — mirrors the
//      takeoff's ConnectionsSummary view.
// If the estimate has no sourceTakeoffId, the user can pick one from a
// dropdown. The selection is persisted on the project so future reconciles
// pick it up automatically.
function ReconciliationPanel({ estimateId }: { estimateId: string }) {
  // Local BOM override selector. Empty = use whatever the server resolved.
  const [bomOverride, setBomOverride] = useState<string>("");
  // File-mode state: when the user drops in XLSX files, we POST them and
  // store the response here — takes priority over the takeoff-based reconcile.
  const [fileData, setFileData] = useState<any | null>(null);
  const [fileError, setFileError] = useState<string>("");
  const [isUploading, setIsUploading] = useState(false);
  const [bomFile, setBomFile] = useState<File | null>(null);
  const [connFile, setConnFile] = useState<File | null>(null);
  const [isDragging, setIsDragging] = useState(false);

  const queryKey = ["/api/estimates", estimateId, "reconcile", bomOverride];
  const { data: takeoffData, isLoading } = useQuery<any>({
    queryKey,
    queryFn: async () => {
      const qs = bomOverride ? `?bomTakeoffId=${encodeURIComponent(bomOverride)}` : "";
      const res = await apiRequest("GET", `/api/estimates/${estimateId}/reconcile${qs}`);
      return res.json();
    },
    enabled: !fileData,  // skip when using uploaded files
  });

  // The active data source: file upload wins if present, else takeoff-derived.
  const data = fileData || takeoffData;

  // Upload handler — takes whichever files the user has staged.
  const uploadFiles = async (bom: File | null, conn: File | null) => {
    if (!bom && !conn) return;
    setIsUploading(true);
    setFileError("");
    try {
      const formData = new FormData();
      if (bom) formData.append("bom", bom);
      if (conn) formData.append("connections", conn);
      const res = await apiUpload(`/api/estimates/${estimateId}/reconcile-from-files`, formData);
      if (!res.ok) {
        const err = await res.json().catch(() => ({ message: `Upload failed (${res.status})` }));
        throw new Error(err.message || "Upload failed");
      }
      const result = await res.json();
      setFileData(result);
    } catch (err: any) {
      setFileError(err.message || "Failed to parse uploaded file(s)");
    } finally {
      setIsUploading(false);
    }
  };

  // Auto-upload as soon as a file is set so the user doesn't need a separate
  // submit click. Re-runs whenever either staged file changes.
  useEffect(() => {
    if (bomFile || connFile) {
      uploadFiles(bomFile, connFile);
    }
  // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [bomFile, connFile]);

  // Drag-and-drop handler: route files by filename keyword.
  const handleDrop = (e: React.DragEvent) => {
    e.preventDefault();
    setIsDragging(false);
    const dropped = Array.from(e.dataTransfer.files);
    let nextBom: File | null = bomFile;
    let nextConn: File | null = connFile;
    for (const f of dropped) {
      const lower = f.name.toLowerCase();
      if (!/\.xlsx?$/i.test(lower)) continue;
      if (lower.includes("connection")) nextConn = f;
      else if (lower.includes("bom")) nextBom = f;
      else {
        // No keyword match — if user hasn't set BOM yet, default to BOM slot.
        if (!nextBom) nextBom = f;
        else if (!nextConn) nextConn = f;
      }
    }
    setBomFile(nextBom);
    setConnFile(nextConn);
  };

  const clearFiles = () => {
    setBomFile(null);
    setConnFile(null);
    setFileData(null);
    setFileError("");
  };

  // Uses the queryClient singleton imported at the top of this file — same
  // pattern as every other mutation here. (useQueryClient() would also work
  // but isn't imported in this module.)
  const linkBomMutation = useMutation({
    mutationFn: async (takeoffId: string | null) => {
      const res = await apiRequest("PATCH", `/api/estimates/${estimateId}`, { sourceTakeoffId: takeoffId });
      return res.json();
    },
    onSuccess: () => {
      queryClient.invalidateQueries({ queryKey: ["/api/estimates", estimateId] });
      setBomOverride("");
    },
  });

  const rows = (data?.rows as any[]) || [];
  const totals = data?.totals || { estimateQty: 0, bomQty: 0, estimateMh: 0 };
  const connRows = (data?.connections?.rows as any[]) || [];
  const connTotals = data?.connections?.totals || { estimate: { total: 0 }, bom: { total: 0 } };
  const hasBom = !!data?.hasBom;
  const resolvedTakeoffId = data?.resolvedTakeoffId || null;
  const sourceTakeoffId = data?.sourceTakeoffId || null;
  const candidateBoms = (data?.candidateBoms as any[]) || [];
  const lineMismatchCount = rows.filter(r => !r.matches).length;
  const connMismatchCount = connRows.filter(r => !r.matches).length;

  return (
    <div className="space-y-3">
      {/* BOM picker / file drop / link control */}
      <Card className="border-card-border">
        <CardHeader className="p-4 pb-2">
          <div className="flex items-center justify-between gap-3 flex-wrap">
            <div className="min-w-0">
              <CardTitle className="text-sm">BOM Reconciliation</CardTitle>
              <p className="text-[10px] text-muted-foreground mt-0.5">
                {fileData ? <>Using uploaded file{(fileData.bomFromFile && fileData.connectionsFromFile) ? "s" : ""}: {fileData.bomFromFile && <span className="font-mono">{fileData.bomFilename}</span>}{fileData.bomFromFile && fileData.connectionsFromFile && ", "}{fileData.connectionsFromFile && <span className="font-mono">{fileData.connFilename}</span>}. {lineMismatchCount} line mismatch{lineMismatchCount === 1 ? "" : "es"}, {connMismatchCount} connection mismatch{connMismatchCount === 1 ? "" : "es"}.</> :
                  isLoading ? "Loading…" :
                  hasBom ? <>Comparing against takeoff <span className="font-mono">{resolvedTakeoffId?.slice(0, 8) || "?"}</span>. {lineMismatchCount} line mismatch{lineMismatchCount === 1 ? "" : "es"}, {connMismatchCount} connection mismatch{connMismatchCount === 1 ? "" : "es"}.</> :
                  "No source BOM. Pick a takeoff below or drop in BOM/Connections Excel files."}
              </p>
            </div>
            {!fileData && (
              <div className="flex items-center gap-2 shrink-0">
                <select
                  className="text-xs border border-input rounded px-2 py-1 bg-background h-7"
                  value={bomOverride || resolvedTakeoffId || ""}
                  onChange={e => setBomOverride(e.target.value)}
                  data-testid="select-reconcile-bom"
                >
                  <option value="">— Select takeoff —</option>
                  {candidateBoms.map(b => (
                    <option key={b.id} value={b.id}>{b.name} ({b.itemCount} items)</option>
                  ))}
                </select>
                {bomOverride && bomOverride !== sourceTakeoffId && (
                  <Button
                    size="sm"
                    variant="outline"
                    className="h-7 text-xs"
                    onClick={() => linkBomMutation.mutate(bomOverride)}
                    disabled={linkBomMutation.isPending}
                    data-testid="btn-link-bom"
                    title="Save this takeoff as the project's source BOM so it auto-loads next time."
                  >
                    Save as source
                  </Button>
                )}
              </div>
            )}
            {fileData && (
              <Button size="sm" variant="outline" className="h-7 text-xs shrink-0" onClick={clearFiles} data-testid="btn-clear-uploaded-files">
                Clear uploaded files
              </Button>
            )}
          </div>
          {/* Drop zone for BOM / Connections Excel exports. Always shown so the
              user can swap in different files at any time. */}
          <div
            onDrop={handleDrop}
            onDragOver={e => { e.preventDefault(); setIsDragging(true); }}
            onDragLeave={() => setIsDragging(false)}
            className={`mt-3 border-2 border-dashed rounded p-3 text-center text-xs transition-colors ${isDragging ? "border-primary bg-primary/5" : "border-border bg-muted/20"}`}
            data-testid="reconcile-dropzone"
          >
            <div className="flex items-center justify-center gap-3 flex-wrap">
              <div className="flex items-center gap-2">
                <label className="text-[10px] text-muted-foreground">BOM xlsx:</label>
                <input
                  type="file"
                  accept=".xlsx,.xls"
                  onChange={e => { const f = e.target.files?.[0] || null; setBomFile(f); }}
                  className="text-[10px]"
                  data-testid="input-bom-file"
                />
                {bomFile && <span className="text-[10px] font-mono">{bomFile.name}</span>}
              </div>
              <div className="flex items-center gap-2">
                <label className="text-[10px] text-muted-foreground">Connections xlsx:</label>
                <input
                  type="file"
                  accept=".xlsx,.xls"
                  onChange={e => { const f = e.target.files?.[0] || null; setConnFile(f); }}
                  className="text-[10px]"
                  data-testid="input-conn-file"
                />
                {connFile && <span className="text-[10px] font-mono">{connFile.name}</span>}
              </div>
              {isUploading && <span className="text-[10px] text-muted-foreground">Parsing…</span>}
            </div>
            <p className="text-[10px] text-muted-foreground/70 mt-1">
              Or drag &amp; drop your <span className="font-mono">- BOM.xlsx</span> and <span className="font-mono">- Connections.xlsx</span> exports anywhere on this card.
            </p>
            {fileError && <p className="text-[10px] text-destructive mt-1">{fileError}</p>}
          </div>
        </CardHeader>
      </Card>

      {/* Table 1: Line-by-line */}
      <Card className="border-card-border">
        <CardHeader className="p-4 pb-2">
          <div className="flex items-center justify-between gap-2">
            <div>
              <CardTitle className="text-sm">Lines — Estimate vs BOM</CardTitle>
              <p className="text-[10px] text-muted-foreground mt-0.5">Each row groups items by category + size + material. Quantities should match the BOM 1:1.</p>
            </div>
            {hasBom && lineMismatchCount > 0 && (
              <span className="text-[10px] px-2 py-1 rounded bg-amber-100 text-amber-800 dark:bg-amber-900/30 dark:text-amber-300 font-semibold">{lineMismatchCount} mismatch{lineMismatchCount === 1 ? "" : "es"}</span>
            )}
          </div>
        </CardHeader>
        <CardContent className="p-4 pt-0">
          {rows.length === 0 ? (
            <p className="text-xs text-muted-foreground py-2">No items to reconcile. Auto-calculate first.</p>
          ) : (
            <div className="max-h-72 overflow-y-auto">
              <table className="w-full text-xs">
                <thead className="sticky top-0 bg-background">
                  <tr className="text-[10px] text-muted-foreground uppercase tracking-wide border-b border-border">
                    <th className="text-left py-1 pr-2">Category</th>
                    <th className="text-left py-1 px-2">Size</th>
                    <th className="text-left py-1 px-2">Mat</th>
                    <th className="text-right py-1 px-2 w-20">Estimate</th>
                    <th className="text-right py-1 px-2 w-20">BOM</th>
                    <th className="text-right py-1 px-2 w-20">Total MH</th>
                    <th className="text-center py-1 px-2 w-8"></th>
                  </tr>
                </thead>
                <tbody>
                  {rows.map((r, idx) => {
                    const delta = r.estimateQty - r.bomQty;
                    return (
                      <tr key={`${r.category}-${r.size}-${r.material}-${idx}`} className={`border-t border-border ${!r.matches && hasBom && r.bomQty > 0 ? "bg-amber-50/50 dark:bg-amber-950/20" : ""}`}>
                        <td className="py-1.5 pr-2 capitalize">{r.category}</td>
                        <td className="py-1.5 px-2 font-mono">{r.size || "—"}</td>
                        <td className="py-1.5 px-2">{r.material || "—"}</td>
                        <td className="py-1.5 px-2 text-right font-mono">{r.estimateQty.toLocaleString()}</td>
                        <td className="py-1.5 px-2 text-right font-mono">{hasBom ? r.bomQty.toLocaleString() : <span className="text-muted-foreground/40">—</span>}</td>
                        <td className="py-1.5 px-2 text-right font-mono">{r.estimateMh.toFixed(1)}</td>
                        <td className="py-1.5 px-2 text-center">
                          {!hasBom ? null : r.bomQty === 0 ? (
                            <span className="text-[10px] text-muted-foreground/60" title="No matching BOM entry — likely an inferred field weld or manually added row.">—</span>
                          ) : r.matches ? (
                            <span className="text-emerald-600 dark:text-emerald-400" title="Matches BOM">✓</span>
                          ) : (
                            <span className="text-amber-600 dark:text-amber-400" title={`Delta ${delta > 0 ? "+" : ""}${delta}`}>{delta > 0 ? "+" : ""}{delta}</span>
                          )}
                        </td>
                      </tr>
                    );
                  })}
                  <tr className="border-t-2 border-border bg-muted/30 font-semibold">
                    <td className="py-1.5 pr-2" colSpan={3}>Totals</td>
                    <td className="py-1.5 px-2 text-right font-mono">{totals.estimateQty.toLocaleString()}</td>
                    <td className="py-1.5 px-2 text-right font-mono">{hasBom ? totals.bomQty.toLocaleString() : "—"}</td>
                    <td className="py-1.5 px-2 text-right font-mono text-primary">{totals.estimateMh.toFixed(1)}</td>
                    <td></td>
                  </tr>
                </tbody>
              </table>
            </div>
          )}
        </CardContent>
      </Card>

      {/* Table 2: Connections rollup vs BOM connections */}
      <Card className="border-card-border">
        <CardHeader className="p-4 pb-2">
          <div className="flex items-center justify-between gap-2">
            <div>
              <CardTitle className="text-sm">Connections — Estimate vs BOM</CardTitle>
              <p className="text-[10px] text-muted-foreground mt-0.5">Total physical connections per size. Should match the takeoff's Connections Summary view.</p>
            </div>
            {hasBom && connMismatchCount > 0 && (
              <span className="text-[10px] px-2 py-1 rounded bg-amber-100 text-amber-800 dark:bg-amber-900/30 dark:text-amber-300 font-semibold">{connMismatchCount} mismatch{connMismatchCount === 1 ? "" : "es"}</span>
            )}
          </div>
        </CardHeader>
        <CardContent className="p-4 pt-0">
          {connRows.length === 0 ? (
            <p className="text-xs text-muted-foreground py-2">No connections to reconcile.</p>
          ) : (
            <div className="max-h-72 overflow-y-auto">
              <table className="w-full text-xs">
                <thead className="sticky top-0 bg-background">
                  <tr className="text-[10px] text-muted-foreground uppercase tracking-wide border-b border-border">
                    <th className="text-left py-1 pr-2">Size</th>
                    <th className="text-right py-1 px-2">Shop BW</th>
                    <th className="text-right py-1 px-2">Shop SW</th>
                    <th className="text-right py-1 px-2">Field BW</th>
                    <th className="text-right py-1 px-2">Field SW</th>
                    <th className="text-right py-1 px-2">Bolt-Ups</th>
                    <th className="text-right py-1 px-2">Threaded</th>
                    <th className="text-right py-1 px-2 w-20">Est Total</th>
                    <th className="text-right py-1 px-2 w-20">BOM Total</th>
                    <th className="text-center py-1 px-2 w-8"></th>
                  </tr>
                </thead>
                <tbody>
                  {connRows.map((r, idx) => (
                    <tr key={`conn-${r.size}-${idx}`} className={`border-t border-border ${!r.matches && hasBom ? "bg-amber-50/50 dark:bg-amber-950/20" : ""}`}>
                      <td className="py-1.5 pr-2 font-mono">{r.size}</td>
                      <td className="py-1.5 px-2 text-right font-mono">{r.estimate.shopBW.toLocaleString()}</td>
                      <td className="py-1.5 px-2 text-right font-mono">{r.estimate.shopSW.toLocaleString()}</td>
                      <td className="py-1.5 px-2 text-right font-mono">{r.estimate.fieldBW.toLocaleString()}</td>
                      <td className="py-1.5 px-2 text-right font-mono">{r.estimate.fieldSW.toLocaleString()}</td>
                      <td className="py-1.5 px-2 text-right font-mono">{r.estimate.boltUps.toLocaleString()}</td>
                      <td className="py-1.5 px-2 text-right font-mono">{r.estimate.threaded.toLocaleString()}</td>
                      <td className="py-1.5 px-2 text-right font-mono font-semibold">{r.estimate.total.toLocaleString()}</td>
                      <td className="py-1.5 px-2 text-right font-mono">{hasBom ? r.bom.total.toLocaleString() : <span className="text-muted-foreground/40">—</span>}</td>
                      <td className="py-1.5 px-2 text-center">
                        {!hasBom ? null : r.matches ? (
                          <span className="text-emerald-600 dark:text-emerald-400" title="Matches BOM">✓</span>
                        ) : (
                          <span className="text-amber-600 dark:text-amber-400" title={`Delta ${r.delta > 0 ? "+" : ""}${r.delta}`}>{r.delta > 0 ? "+" : ""}{r.delta}</span>
                        )}
                      </td>
                    </tr>
                  ))}
                  <tr className="border-t-2 border-border bg-muted/30 font-semibold">
                    <td className="py-1.5 pr-2">Totals</td>
                    <td className="py-1.5 px-2 text-right font-mono">{connTotals.estimate.shopBW.toLocaleString()}</td>
                    <td className="py-1.5 px-2 text-right font-mono">{connTotals.estimate.shopSW.toLocaleString()}</td>
                    <td className="py-1.5 px-2 text-right font-mono">{connTotals.estimate.fieldBW.toLocaleString()}</td>
                    <td className="py-1.5 px-2 text-right font-mono">{connTotals.estimate.fieldSW.toLocaleString()}</td>
                    <td className="py-1.5 px-2 text-right font-mono">{connTotals.estimate.boltUps.toLocaleString()}</td>
                    <td className="py-1.5 px-2 text-right font-mono">{connTotals.estimate.threaded.toLocaleString()}</td>
                    <td className="py-1.5 px-2 text-right font-mono text-primary">{connTotals.estimate.total.toLocaleString()}</td>
                    <td className="py-1.5 px-2 text-right font-mono">{hasBom ? connTotals.bom.total.toLocaleString() : "—"}</td>
                    <td></td>
                  </tr>
                </tbody>
              </table>
            </div>
          )}
        </CardContent>
      </Card>
    </div>
  );
}

// ScopeAddersPanel — manages project-level hand-entered labor scope rows that
// aren't on the BOM (hydro test, demo, supports, ID tags, supervision, etc.).
// Each row's MH flows into the project total at the effective blended labor
// rate (or per-row override). Add/remove/edit auto-saves through the parent's
// patchMutation via the onChange callback.
type ScopeAdder = { id: string; label: string; mode?: "hours" | "cost"; hours: number; ratePerHour?: number; flatCost?: number; note?: string };
function ScopeAddersPanel({ adders, effectiveRate, onChange }: { adders: ScopeAdder[]; effectiveRate: number; onChange: (next: ScopeAdder[]) => void }) {
  // Two modes per row:
  //   "hours" — hours × rate flows into labor (gets full markup including tax? no, tax is material-only)
  //   "cost"  — flat dollar amount flows directly into subtotal (overhead/profit/bond apply, tax does not)
  // Totals split so the user can see both buckets at a glance.
  const totalHours = adders.reduce((s, a) => s + ((a.mode ?? "hours") === "hours" ? (a.hours || 0) : 0), 0);
  const totalLaborCost = adders.reduce((s, a) => s + ((a.mode ?? "hours") === "hours" ? (a.hours || 0) * (a.ratePerHour ?? effectiveRate) : 0), 0);
  const totalFlatCost = adders.reduce((s, a) => s + (a.mode === "cost" ? (a.flatCost || 0) : 0), 0);
  const totalCost = totalLaborCost + totalFlatCost;

  function update(id: string, patch: Partial<ScopeAdder>) {
    onChange(adders.map(a => a.id === id ? { ...a, ...patch } : a));
  }
  function remove(id: string) {
    onChange(adders.filter(a => a.id !== id));
  }
  function addNew(mode: "hours" | "cost" = "hours") {
    const id = (window.crypto?.randomUUID?.() ?? `adder-${Date.now()}-${Math.random().toString(36).slice(2, 8)}`);
    onChange([...adders, { id, label: mode === "cost" ? "New cost item" : "New scope", mode, hours: 0, flatCost: 0 }]);
  }

  return (
    <Card className="border-card-border">
      <CardHeader className="p-4 pb-2">
        <div className="flex items-center justify-between gap-2">
          <div>
            <CardTitle className="text-sm">Scope Adders</CardTitle>
            <p className="text-[10px] text-muted-foreground mt-0.5">
              Hand-entered scope that isn't on the BOM. Hours rows flow into labor at the blended rate ({fmt$(effectiveRate)}/hr incl. per diem, override per row). Flat $ rows add direct cost (e.g. "MISC Supports $750") — marked up by overhead/profit/bond, not tax.
            </p>
          </div>
          <div className="flex items-center gap-1 shrink-0">
            <Button size="sm" variant="outline" onClick={() => addNew("hours")} className="h-7 text-xs" data-testid="btn-add-adder-hours">
              <Plus size={12} className="mr-1" /> Hours row
            </Button>
            <Button size="sm" variant="outline" onClick={() => addNew("cost")} className="h-7 text-xs" data-testid="btn-add-adder-cost">
              <Plus size={12} className="mr-1" /> $ row
            </Button>
          </div>
        </div>
      </CardHeader>
      <CardContent className="p-4 pt-0">
        {adders.length === 0 ? (
          <p className="text-xs text-muted-foreground py-2">No scope adders. Add an Hours row for hydro/demo/supervision, or a $ row for a flat cost like "MISC Supports $750".</p>
        ) : (
          <table className="w-full text-xs">
            <thead>
              <tr className="text-[10px] text-muted-foreground uppercase tracking-wide">
                <th className="text-left py-1 pr-2">Label</th>
                <th className="text-left py-1 px-2 w-20">Mode</th>
                <th className="text-right py-1 px-2 w-20">Hours</th>
                <th className="text-right py-1 px-2 w-24">Rate $/hr</th>
                <th className="text-right py-1 px-2 w-24">Cost</th>
                <th className="py-1 w-8"></th>
              </tr>
            </thead>
            <tbody>
              {adders.map(a => {
                const mode = a.mode ?? "hours";
                const rate = a.ratePerHour ?? effectiveRate;
                const cost = mode === "hours" ? (a.hours || 0) * rate : (a.flatCost || 0);
                const dimmed = "text-muted-foreground/40";
                return (
                  <tr key={a.id} className="border-t border-border">
                    <td className="py-1 pr-2">
                      <input
                        type="text"
                        defaultValue={a.label}
                        onBlur={e => { if (e.target.value !== a.label) update(a.id, { label: e.target.value }); }}
                        className="w-full bg-transparent border-0 border-b border-transparent hover:border-border focus:border-primary outline-none px-1"
                        placeholder={mode === "cost" ? "MISC Supports, Subcontract…" : "Hydro test, Demo, Supervision…"}
                        data-testid={`input-adder-label-${a.id}`}
                      />
                      {a.note && <p className="text-[9px] text-muted-foreground pl-1 mt-0.5">{a.note}</p>}
                    </td>
                    <td className="py-1 px-2">
                      <select
                        value={mode}
                        onChange={e => update(a.id, { mode: e.target.value as "hours" | "cost" })}
                        className="w-full bg-transparent border border-transparent hover:border-border focus:border-primary outline-none px-1 py-0.5 rounded text-[11px]"
                        data-testid={`select-adder-mode-${a.id}`}
                      >
                        <option value="hours">Hours</option>
                        <option value="cost">Flat $</option>
                      </select>
                    </td>
                    <td className="py-1 px-2 text-right">
                      <input
                        type="number"
                        step="0.5"
                        defaultValue={a.hours}
                        disabled={mode === "cost"}
                        onBlur={e => {
                          const v = parseFloat(e.target.value) || 0;
                          if (v !== a.hours) update(a.id, { hours: v });
                        }}
                        className={`w-full text-right bg-transparent font-mono border-0 border-b border-transparent hover:border-border focus:border-primary outline-none px-1 ${mode === "cost" ? dimmed : ""}`}
                        data-testid={`input-adder-hours-${a.id}`}
                      />
                    </td>
                    <td className="py-1 px-2 text-right">
                      <input
                        type="number"
                        step="0.5"
                        defaultValue={a.ratePerHour ?? ""}
                        placeholder={effectiveRate.toFixed(2)}
                        disabled={mode === "cost"}
                        onBlur={e => {
                          const raw = e.target.value.trim();
                          const v = raw === "" ? undefined : (parseFloat(raw) || 0);
                          if (v !== a.ratePerHour) update(a.id, { ratePerHour: v });
                        }}
                        className={`w-full text-right bg-transparent font-mono border-0 border-b border-transparent hover:border-border focus:border-primary outline-none px-1 ${mode === "cost" ? dimmed : "text-muted-foreground"}`}
                        data-testid={`input-adder-rate-${a.id}`}
                      />
                    </td>
                    <td className="py-1 px-2 text-right">
                      {mode === "cost" ? (
                        <input
                          type="number"
                          step="50"
                          defaultValue={a.flatCost ?? 0}
                          onBlur={e => {
                            const v = parseFloat(e.target.value) || 0;
                            if (v !== a.flatCost) update(a.id, { flatCost: v });
                          }}
                          className="w-full text-right bg-transparent font-mono border-0 border-b border-transparent hover:border-border focus:border-primary outline-none px-1"
                          data-testid={`input-adder-flatcost-${a.id}`}
                        />
                      ) : (
                        <span className="font-mono">{fmt$(cost)}</span>
                      )}
                    </td>
                    <td className="py-1">
                      <button
                        className="text-muted-foreground hover:text-destructive"
                        onClick={() => { if (window.confirm(`Remove '${a.label}'?`)) remove(a.id); }}
                        title="Remove this adder"
                        data-testid={`btn-remove-adder-${a.id}`}
                      >
                        <Trash2 size={12} />
                      </button>
                    </td>
                  </tr>
                );
              })}
              <tr className="border-t border-border bg-muted/30 font-semibold">
                <td className="py-1.5 pr-2">Totals</td>
                <td className="py-1.5 px-2"></td>
                <td className="py-1.5 px-2 text-right font-mono">{totalHours.toFixed(1)}</td>
                <td className="py-1.5 px-2 text-right text-[10px] text-muted-foreground">{totalFlatCost > 0 ? `+${fmt$(totalFlatCost)} flat` : ""}</td>
                <td className="py-1.5 px-2 text-right font-mono text-primary">{fmt$(totalCost)}</td>
                <td></td>
              </tr>
            </tbody>
          </table>
        )}
      </CardContent>
    </Card>
  );
}
