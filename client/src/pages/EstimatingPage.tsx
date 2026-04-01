import { useState, useRef, useMemo, useCallback } from "react";
import { useQuery, useMutation } from "@tanstack/react-query";
import { Plus, Trash2, Download, Database, ChevronDown, ChevronRight, ChevronUp, Edit2, Check, X, Search, Calculator, Zap, FileSpreadsheet, Info, Settings2, ArrowUpDown, History, Upload, Wand2, ShoppingCart } from "lucide-react";
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
import type { EstimateProject, EstimateItem, CostDatabaseEntry } from "@shared/schema";

function fmt$(n: number) { return `$${n.toFixed(2)}`; }

function computeItem(item: EstimateItem): EstimateItem {
  const me = item.quantity * (item.materialUnitCost || 0);
  const le = item.quantity * (item.laborUnitCost || 0);
  return { ...item, materialExtension: me, laborExtension: le, totalCost: me + le };
}

const CATEGORIES = ["pipe", "elbow", "tee", "reducer", "valve", "flange", "gasket", "bolt", "cap", "coupling", "union", "weld", "support", "strainer", "trap", "fitting", "steel", "concrete", "rebar", "earthwork", "paving", "electrical", "other"];

export default function EstimatingPage() {
  const { toast } = useToast();
  const [selectedId, setSelectedId] = useState<string | null>(null);
  const [newProjectName, setNewProjectName] = useState("");
  const [showNewProject, setShowNewProject] = useState(false);
  const [quickEntry, setQuickEntry] = useState("");
  const [editingMarkups, setEditingMarkups] = useState(false);
  const [dbSearch, setDbSearch] = useState("");

  // Estimating method state
  const [estMethod, setEstMethod] = useState<"bill" | "justin">("justin");
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

  const autoCalculateMutation = useMutation({
    mutationFn: (id: string) =>
      apiRequest("POST", `/api/estimates/${id}/auto-calculate`, {
        method: estMethod,
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
      }).then(r => r.json()),
    onSuccess: () => {
      queryClient.invalidateQueries({ queryKey: ["/api/estimates", selectedId] });
      toast({ title: `Labor hours calculated using ${estMethod === "bill" ? "Bill's EI" : "Justin's Factor"} method` });
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
  const handleUpdateItemOverride = (itemId: string, field: string, value: string | undefined) => {
    if (!selectedProject) return;
    const items = selectedProject.items.map(i =>
      i.id === itemId ? { ...i, [field]: value || undefined } : i
    );
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
  const totalMaterial = p?.items.reduce((s, i) => s + (i.materialExtension || 0), 0) || 0;
  const totalLabor = p?.items.reduce((s, i) => s + (i.laborExtension || 0), 0) || 0;
  const totalHours = p?.items.reduce((s, i) => s + i.quantity * (i.laborHoursPerUnit || 0), 0) || 0;
  const subtotal = totalMaterial + totalLabor;
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
      <div className="flex h-full">
        {/* Left panel */}
        <div className="w-60 shrink-0 border-r border-border flex flex-col">
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
                    <div className="grid grid-cols-2 md:grid-cols-4 gap-3">
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
                              Using {p.estimateMethod === "bill" ? "Bill's EI Method" : "Justin's Factor Method"}
                            </Badge>
                          )}
                        </div>

                        {/* Method selector */}
                        <div className="flex gap-2 mb-3">
                          <button
                            onClick={() => setEstMethod("bill")}
                            data-testid="btn-method-bill"
                            className={`flex-1 text-xs px-3 py-1.5 rounded border transition-colors ${
                              estMethod === "bill"
                                ? "bg-primary text-primary-foreground border-primary"
                                : "border-input hover:bg-accent"
                            }`}
                          >
                            Bill's Method (EI-Based)
                          </button>
                          <button
                            onClick={() => setEstMethod("justin")}
                            data-testid="btn-method-justin"
                            className={`flex-1 text-xs px-3 py-1.5 rounded border transition-colors ${
                              estMethod === "justin"
                                ? "bg-primary text-primary-foreground border-primary"
                                : "border-input hover:bg-accent"
                            }`}
                          >
                            Justin's Method (Factor-Based)
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

                    {/* Estimate items table */}
                    <div className="overflow-auto rounded-md border border-border">
                      <table className="w-full text-xs">
                        <thead>
                          <tr className="bg-muted/50 border-b border-border">
                            {([
                              { field: "lineNumber" as SortField, label: "#", cls: "w-8 text-left" },
                              { field: "category" as SortField, label: "Category", cls: "text-left" },
                              { field: "size" as SortField, label: "Size", cls: "text-left" },
                              { field: "description" as SortField, label: "Description", cls: "text-left min-w-[160px]" },
                              { field: "quantity" as SortField, label: "Qty", cls: "text-right w-16" },
                              { field: null as any, label: "Unit", cls: "text-left w-12" },
                              { field: "materialUnitCost" as SortField, label: "Mat $/Unit", cls: "text-right w-20" },
                              { field: "laborUnitCost" as SortField, label: "Labor $/Unit", cls: "text-right w-20" },
                              { field: "laborHoursPerUnit" as SortField, label: "Hrs/Unit", cls: "text-right w-16" },
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
                              <td colSpan={13} className="text-center py-8 text-muted-foreground">
                                No items. Use quick entry above or import from takeoff.
                              </td>
                            </tr>
                          ) : displayItems.length === 0 ? (
                            <tr>
                              <td colSpan={13} className="text-center py-6 text-muted-foreground">
                                No items match filter.
                              </td>
                            </tr>
                          ) : (
                            <TooltipProvider delayDuration={300}>
                            {displayItems.map(item => (
                              <tr key={item.id} className="border-b border-border hover:bg-muted/20" data-testid={`est-row-${item.id}`}>
                                <td className="px-2 py-1.5 text-muted-foreground">{item.lineNumber}</td>
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
                                    <PopoverContent className="w-56 p-3 space-y-2" side="left">
                                      <p className="text-[10px] font-semibold text-muted-foreground uppercase">Line Overrides</p>
                                      {[
                                        { field: "workType", label: "Work Type", options: ["standard", "rack"] },
                                        { field: "itemMaterial", label: "Material", options: ["CS", "SS"] },
                                        { field: "itemSchedule", label: "Schedule", options: ["STD", "XH", "10", "20", "40", "80", "160/XXH"] },
                                        { field: "itemElevation", label: "Elevation", options: ["0-20ft", "20-40ft", "40-80ft", "80ft+"] },
                                        { field: "itemPipeLocation", label: "Pipe Location", options: ["Sleeper Rack", "Underground", "Open Rack", "Elevated Rack"] },
                                        { field: "itemAlloyGroup", label: "Alloy Group", options: ["1", "2", "3", "4", "5", "6", "7", "8", "9"] },
                                      ].map(({ field, label, options }) => (
                                        <div key={field} className="flex items-center gap-2">
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
                            ))}
                            </TooltipProvider>
                          )}
                        </tbody>
                        {p.items.length > 0 && (
                          <tfoot>
                            <tr className="bg-muted/50 font-medium border-t-2 border-border">
                              <td colSpan={9} className="px-2 py-2 text-right text-xs">TOTALS</td>
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
                        </CardContent>
                      </Card>

                      {/* Totals */}
                      <Card className="border-card-border">
                        <CardHeader className="p-4 pb-2">
                          <CardTitle className="text-sm">Summary</CardTitle>
                        </CardHeader>
                        <CardContent className="p-4 pt-0 space-y-1">
                          {[
                            { label: "Material", val: totalMaterial },
                            { label: "Labor", val: totalLabor },
                            { label: `Labor Hours`, val: null, text: `${totalHours.toFixed(1)} hrs` },
                            { label: "Subtotal", val: subtotal, bold: true },
                            { label: `Overhead (${p.markups.overhead}%)`, val: overheadAmt },
                            { label: `Profit (${p.markups.profit}%)`, val: profitAmt },
                            { label: `Tax (${p.markups.tax}%)`, val: taxAmt },
                            { label: `Bond (${p.markups.bond}%)`, val: bondAmt },
                          ].map(({ label, val, text, bold }) => (
                            <div key={label} className="flex justify-between text-xs">
                              <span className={bold ? "font-semibold" : "text-muted-foreground"}>{label}</span>
                              <span className={`font-mono ${bold ? "font-semibold" : ""}`}>
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
